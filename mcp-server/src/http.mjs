// Shared server mode.
//
// One long-running process (started with --http, usually by a launchd agent)
// owns the WebSocket the editor plugins connect to, and serves MCP over
// Streamable HTTP on 127.0.0.1 so that every Claude session uses the same
// documents instead of fighting over the WebSocket port.
//
// Sessions launched with the usual stdio command detect the shared server and
// relay their MCP messages to it (runProxy), so existing configurations keep
// working unchanged.

import { createServer } from 'node:http';
import { randomUUID, timingSafeEqual } from 'node:crypto';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { isInitializeRequest } from '@modelcontextprotocol/sdk/types.js';

const log = (...args) => console.error('[claude-bridge]', ...args);

const MAX_BODY = 8 * 1024 * 1024; // images are sent as data URLs
const SERVER_NAME = 'onlyoffice-claude-bridge';

function sameSecret(a, b) {
  const x = Buffer.from(String(a || ''));
  const y = Buffer.from(String(b || ''));
  return x.length === y.length && timingSafeEqual(x, y);
}

function sendJson(res, status, body) {
  res.writeHead(status, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(body));
}

function rpcError(res, status, message) {
  sendJson(res, status, { jsonrpc: '2.0', error: { code: -32000, message }, id: null });
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > MAX_BODY) { reject(Object.assign(new Error('Request body too large'), { status: 413 })); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => {
      try { resolve(chunks.length ? JSON.parse(Buffer.concat(chunks).toString('utf8')) : undefined); }
      catch { reject(Object.assign(new Error('Invalid JSON'), { status: 400 })); }
    });
    req.on('error', reject);
  });
}

// Serves MCP on http://127.0.0.1:<port>/mcp. `makeServer()` returns a new,
// unconnected McpServer; each MCP session gets its own, all sharing one bridge.
// `getStatus()` feeds GET /status (bearer token required), used by the menu bar app.
export function startHttpServer({ port, token, version, makeServer, getStatus }) {
  const host = '127.0.0.1';
  const allowedHosts = new Set([`127.0.0.1:${port}`, `localhost:${port}`]);
  const sessions = new Map(); // MCP session id -> transport

  const server = createServer(async (req, res) => {
    try {
      const path = (req.url || '').split('?')[0];

      // DNS rebinding protection: only answer to the loopback names.
      if (!allowedHosts.has(String(req.headers.host || '').toLowerCase())) {
        return sendJson(res, 403, { error: 'Host not allowed' });
      }
      // MCP clients do not send an Origin header; browsers do. No web page may call this server.
      if (req.headers.origin) return sendJson(res, 403, { error: 'Browser requests are not allowed' });

      if (path === '/health' && req.method === 'GET') {
        return sendJson(res, 200, { name: SERVER_NAME, version, sessions: sessions.size });
      }
      if (path !== '/mcp' && path !== '/status') return sendJson(res, 404, { error: 'Not found' });

      const auth = String(req.headers.authorization || '');
      if (!auth.startsWith('Bearer ') || !sameSecret(auth.slice(7), token)) {
        return rpcError(res, 401, 'Missing or wrong bearer token');
      }
      if (path === '/status') {
        if (req.method !== 'GET') return sendJson(res, 405, { error: 'Method not allowed' });
        return sendJson(res, 200, { name: SERVER_NAME, version, pid: process.pid, sessions: sessions.size, ...(getStatus ? getStatus() : {}) });
      }

      const sessionId = req.headers['mcp-session-id'];
      const body = req.method === 'POST' ? await readJson(req) : undefined;
      let transport = sessionId ? sessions.get(String(sessionId)) : null;

      if (!transport) {
        if (sessionId) return rpcError(res, 404, 'Unknown MCP session, reconnect');
        if (req.method !== 'POST' || !isInitializeRequest(body)) {
          return rpcError(res, 400, 'No MCP session: send an initialize request first');
        }
        transport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (id) => { sessions.set(id, transport); log(`MCP session opened (${sessions.size} open)`); },
        });
        transport.onclose = () => {
          if (transport.sessionId && sessions.delete(transport.sessionId)) log(`MCP session closed (${sessions.size} open)`);
        };
        await makeServer().connect(transport);
      }
      await transport.handleRequest(req, res, body);
    } catch (err) {
      log('HTTP error:', err.message);
      if (!res.headersSent) rpcError(res, err.status || 500, err.message);
      else res.end();
    }
  });

  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, host, () => {
      log(`shared MCP endpoint: http://${host}:${port}/mcp`);
      resolve(server);
    });
  });
}

// True when a shared server answers on this port.
export async function sharedServerUp(port, timeoutMs = 800) {
  try {
    const r = await fetch(`http://127.0.0.1:${port}/health`, { signal: AbortSignal.timeout(timeoutMs) });
    if (!r.ok) return false;
    const info = await r.json();
    return info?.name === SERVER_NAME;
  } catch {
    return false;
  }
}

// Relays MCP messages between this process's stdio (the Claude session) and
// the shared server, without interpreting them. When the shared server
// restarts (its MCP sessions are lost), the relay opens a new session by
// replaying the session's initialize handshake, then resends the message, so
// the Claude session does not notice.
export async function runProxy({ port, token }) {
  const url = new URL(`http://127.0.0.1:${port}/mcp`);
  const local = new StdioServerTransport();
  const REPLAY_ID = '__claude_bridge_reinitialize__';
  let initRequest = null;  // the session's initialize request
  let initializedNote = null; // and its "initialized" notification
  let protocolVersion = null;
  let upstream = null;
  let reconnecting = null;
  let closing = false;

  const makeUpstream = () => {
    const t = new StreamableHTTPClientTransport(url, {
      requestInit: { headers: { Authorization: `Bearer ${token}` } },
    });
    t.onmessage = (msg) => {
      if (msg?.id === REPLAY_ID) return; // answer to a replayed handshake: the session already has one
      // Later requests must carry the negotiated protocol version.
      if (initRequest && msg?.id === initRequest.id && msg.result?.protocolVersion) {
        protocolVersion = msg.result.protocolVersion;
        t.setProtocolVersion(protocolVersion);
      }
      local.send(msg).catch(() => {});
    };
    t.onerror = (err) => log('shared server:', err.message);
    return t;
  };

  // Opens a new MCP session on the shared server. With `replay`, re-sends the
  // session's handshake (after a restart of the shared server).
  const connect = async ({ replay = true } = {}) => {
    const old = upstream;
    upstream = null;
    if (old) old.terminateSession().catch(() => {}).finally(() => old.close().catch(() => {}));
    const t = makeUpstream();
    await t.start();
    if (replay && initRequest) {
      // The session id is read from the response headers, before send() resolves.
      await t.send({ ...initRequest, id: REPLAY_ID });
      if (protocolVersion) t.setProtocolVersion(protocolVersion);
      if (initializedNote) await t.send(initializedNote);
      log('reconnected to the shared server');
    }
    upstream = t;
  };
  const ensureConnected = () => {
    if (upstream) return Promise.resolve();
    reconnecting ||= connect().finally(() => { reconnecting = null; });
    return reconnecting;
  };
  const sessionLost = (err) => /Unknown MCP session|\b404\b|No MCP session/i.test(String(err?.message || err));

  const forward = async (msg) => {
    if (msg?.method === 'initialize') {
      // A client may initialize more than once over the same stdio (e.g. when it
      // reconnects): each initialize gets a fresh session, since the shared
      // server refuses a second one in the same session.
      if (reconnecting) await reconnecting.catch(() => {});
      reconnecting = connect({ replay: false }).finally(() => { reconnecting = null; });
      await reconnecting;
      await upstream.send(msg);
      return;
    }
    await ensureConnected();
    const t = upstream;
    try {
      await t.send(msg);
    } catch (err) {
      if (!initRequest || msg?.method === 'initialize' || !sessionLost(err)) throw err;
      log('shared server lost this session (restarted?), reconnecting');
      if (upstream === t) upstream = null;
      await ensureConnected();
      await upstream.send(msg);
    }
  };

  const stop = () => {
    if (closing) return;
    closing = true;
    const t = upstream;
    // Ends the MCP session on the shared server (DELETE), then closes both sides.
    (t ? t.terminateSession().catch(() => {}) : Promise.resolve())
      .finally(() => Promise.allSettled([t?.close(), local.close()]))
      .finally(() => process.exit(0));
  };

  local.onmessage = (msg) => {
    if (msg?.method === 'initialize') initRequest = msg;
    else if (msg?.method === 'notifications/initialized') initializedNote = msg;
    forward(msg).catch((err) => {
      if (sessionLost(err) || /fetch failed|ECONNREFUSED/i.test(String(err?.message))) upstream = null; // retry on the next message
      // Answer the request ourselves so the session does not hang.
      if (msg?.id !== undefined && msg.method) {
        local.send({ jsonrpc: '2.0', id: msg.id, error: { code: -32000, message: `Shared OnlyOffice server unreachable: ${err.message}` } }).catch(() => {});
      }
    });
  };
  local.onclose = () => stop();
  process.stdin.on('close', () => stop());

  await connect();
  await local.start();
  log(`relaying this session to the shared server on port ${port}`);
}

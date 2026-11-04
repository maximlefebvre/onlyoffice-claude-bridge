// WebSocket hub between this MCP server and the Claude Bridge plugins running
// inside OnlyOffice editors.

import { WebSocketServer } from 'ws';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const log = (...args) => console.error('[claude-bridge]', ...args); // stdout is reserved for MCP

export const MIN_PLUGIN_VERSION = '0.2.0';

function isCompatible(version) {
  const parse = (v) => String(v || '0').split('.').map((n) => parseInt(n, 10) || 0);
  const [a, b] = [parse(version), parse(MIN_PLUGIN_VERSION)];
  for (let i = 0; i < 3; i++) {
    if ((a[i] || 0) !== (b[i] || 0)) return (a[i] || 0) > (b[i] || 0);
  }
  return true;
}

export function loadToken() {
  const dir = join(homedir(), '.config', 'onlyoffice-claude-bridge');
  const file = join(dir, 'token');
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  if (!existsSync(file)) writeFileSync(file, randomBytes(24).toString('base64url'), { mode: 0o600 });
  return { token: readFileSync(file, 'utf8').trim(), file };
}

export class Bridge {
  constructor({ port, host = '127.0.0.1', allowedOrigins, token, timeoutMs = 20000 }) {
    this.port = port;
    this.host = host;
    this.allowedOrigins = allowedOrigins;
    this.token = token;
    this.timeoutMs = timeoutMs;
    this.clients = new Map(); // ws -> client
    this.pending = new Map(); // request id -> { resolve, reject, timer }
    this.nextId = 1;
    this.listenError = null;
  }

  start() {
    this.wss = new WebSocketServer({
      host: this.host,
      port: this.port,
      verifyClient: ({ origin }, done) => {
        const ok = this.allowedOrigins.includes('*') || this.allowedOrigins.includes(origin);
        if (!ok) log(`connection refused, origin not allowed: ${origin}`);
        done(ok, 403, 'Origin not allowed');
      },
    });
    this.wss.on('listening', () => log(`listening on ws://${this.host}:${this.port}`));
    this.wss.on('error', (err) => {
      this.listenError = err;
      log(err.code === 'EADDRINUSE'
        ? `port ${this.port} is already in use (another instance or the test server is running)`
        : `WebSocket server error: ${err.message}`);
    });
    this.wss.on('connection', (ws, req) => this.#onConnection(ws, req));
  }

  #sameToken(t) {
    const a = Buffer.from(String(t || ''));
    const b = Buffer.from(this.token);
    return a.length === b.length && timingSafeEqual(a, b);
  }

  #onConnection(ws, req) {
    const origin = req.headers.origin;
    const helloTimer = setTimeout(() => ws.close(4001, 'no hello'), 5000);

    ws.on('message', (raw) => {
      let msg;
      try { msg = JSON.parse(raw); } catch { return; }

      if (!this.clients.has(ws)) {
        clearTimeout(helloTimer);
        if (msg.type !== 'hello' || !this.#sameToken(msg.token)) {
          log(`invalid token from ${origin}`);
          ws.close(4001, 'bad token');
          return;
        }
        const info = msg.info || {};
        const client = {
          id: this.nextId++,
          ws,
          origin,
          editor: info.editorType || null,
          title: info.documentTitle || null,
          documentId: info.documentId || null,
          pluginVersion: info.pluginVersion || null,
          connectedAt: new Date(),
          lastUsedAt: new Date(),
        };
        // A reload of the same document replaces its previous connection.
        for (const [otherWs, other] of this.clients) {
          if (client.documentId && other.documentId === client.documentId) {
            this.clients.delete(otherWs);
            try { otherWs.close(4002, 'replaced'); } catch {}
          }
        }
        this.clients.set(ws, client);
        ws.send(JSON.stringify({ type: 'welcome' }));
        log(`document #${client.id} connected: ${client.title} (${client.editor}, plugin ${client.pluginVersion})`);
        return;
      }

      const p = this.pending.get(msg.id);
      if (!p) return;
      this.pending.delete(msg.id);
      clearTimeout(p.timer);
      if (msg.error) p.reject(new Error(msg.error.message || 'Editor error'));
      else p.resolve(msg.result);
    });

    ws.on('close', () => {
      clearTimeout(helloTimer);
      const c = this.clients.get(ws);
      if (c) log(`document #${c.id} disconnected: ${c.title}`);
      this.clients.delete(ws);
    });
  }

  listDocuments() {
    return [...this.clients.values()]
      .sort((a, b) => b.lastUsedAt - a.lastUsedAt)
      .map((c) => ({
        id: c.id,
        title: c.title,
        editor: c.editor,
        pluginVersion: c.pluginVersion,
        connectedAt: c.connectedAt.toISOString(),
      }));
  }

  // Picks the target document: by id, by title substring, or the most recently
  // used one (optionally restricted to an editor type).
  resolve(document, editor) {
    if (this.listenError) {
      throw new Error(`The bridge is not listening on port ${this.port}: ${this.listenError.message}`);
    }
    let candidates = [...this.clients.values()];
    if (!candidates.length) {
      throw new Error('No OnlyOffice document is connected. Open a document, open the Claude Bridge plugin panel and click Connect.');
    }
    if (document !== undefined && document !== null && document !== '') {
      const asNum = Number(document);
      const byId = candidates.find((c) => c.id === asNum);
      if (byId) candidates = [byId];
      else {
        const q = String(document).toLowerCase();
        candidates = candidates.filter((c) => (c.title || '').toLowerCase().includes(q));
        if (!candidates.length) throw new Error(`No connected document matches "${document}". Use list_documents.`);
      }
    }
    if (editor) {
      const matching = candidates.filter((c) => c.editor === editor);
      if (!matching.length) {
        const kinds = { word: 'text document', cell: 'spreadsheet', slide: 'presentation' };
        throw new Error(`This tool needs a ${kinds[editor] || editor}, but no matching document is connected. Use list_documents.`);
      }
      candidates = matching;
    }
    candidates.sort((a, b) => b.lastUsedAt - a.lastUsedAt);
    const target = candidates[0];
    if (!isCompatible(target.pluginVersion)) {
      throw new Error(`"${target.title}" runs Claude Bridge plugin ${target.pluginVersion}, which is too old (need ${MIN_PLUGIN_VERSION}+). Ask the user to update the plugin in the Plugin Manager (or reinstall it) and reconnect.`);
    }
    return target;
  }

  call(client, method, params = {}) {
    client.lastUsedAt = new Date();
    return new Promise((resolve, reject) => {
      const id = randomUUID();
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`The editor did not answer within ${this.timeoutMs / 1000}s`));
      }, this.timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      client.ws.send(JSON.stringify({ id, method, params }));
    });
  }

  // Runs `fn` (a function whose body uses the OnlyOffice `Api` and `Asc.scope`)
  // inside the editor.
  run(client, fn, scope = {}, recalc = true) {
    const code = typeof fn === 'function' ? fn.toString() : `function () {\n${fn}\n}`;
    return this.call(client, 'run_command', { code, scope, recalc });
  }

  method(client, name, args = []) {
    return this.call(client, 'execute_method', { name, args });
  }
}

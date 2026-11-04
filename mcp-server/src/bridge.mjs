// WebSocket hub between this MCP server and the Claude Bridge plugins running
// inside OnlyOffice editors.

import { WebSocketServer } from 'ws';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { execFile } from 'node:child_process';
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

const CONFIG_DIR = join(homedir(), '.config', 'onlyoffice-claude-bridge');

export function loadToken() {
  const file = join(CONFIG_DIR, 'token');
  mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  if (!existsSync(file)) writeFileSync(file, randomBytes(24).toString('base64url'), { mode: 0o600 });
  return { token: readFileSync(file, 'utf8').trim(), file };
}

// Tokens handed out to plugins after the user approved them ("paired devices").
// Only SHA-256 hashes are stored.
class DeviceStore {
  constructor() {
    this.file = join(CONFIG_DIR, 'devices.json');
    try { this.devices = JSON.parse(readFileSync(this.file, 'utf8')); } catch { this.devices = []; }
  }
  static hash(token) { return createHash('sha256').update(String(token)).digest('hex'); }
  has(token) {
    if (!token) return false;
    const h = DeviceStore.hash(token);
    const d = this.devices.find((x) => x.hash === h);
    if (d) { d.lastSeen = new Date().toISOString(); this.#save(); }
    return !!d;
  }
  add(label) {
    const token = randomBytes(24).toString('base64url');
    this.devices.push({ hash: DeviceStore.hash(token), label, created: new Date().toISOString() });
    this.#save();
    return token;
  }
  list() { return this.devices.map(({ label, created, lastSeen }, i) => ({ id: i + 1, label, created, lastSeen })); }
  remove(id) {
    const [d] = this.devices.splice(id - 1, 1);
    this.#save();
    return d || null;
  }
  #save() {
    mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
    writeFileSync(this.file, JSON.stringify(this.devices, null, 2), { mode: 0o600 });
  }
}

// Asks the user on this computer whether to allow a new connection.
// macOS: native dialog. Elsewhere (or if the dialog fails): returns null and the
// request waits for approval through the MCP tools.
function askNative(message) {
  if (process.platform !== 'darwin' || process.env.CLAUDE_BRIDGE_NO_DIALOG) return Promise.resolve(null);
  const esc = (t) => String(t).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  const script = `display dialog "${esc(message)}" with title "Claude Bridge" `
    + 'buttons {"Deny", "Allow"} default button "Allow" cancel button "Deny" '
    + 'with icon caution giving up after 120';
  return new Promise((resolve) => {
    execFile('osascript', ['-e', script], { timeout: 130000 }, (err, stdout) => {
      if (err) {
        // Exit code 1 = "Deny" (user cancelled). Anything else: dialog unavailable.
        resolve(err.code === 1 ? false : null);
        return;
      }
      resolve(/button returned:Allow/.test(stdout) && !/gave up:true/.test(stdout));
    });
  });
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
    this.devices = new DeviceStore();
    this.pairing = new Map(); // pairing id -> { ws, client, resolve }
    this.nextPairingId = 1;
    this.deniedUntil = 0; // after a denial, ignore new requests for a short while
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
        if (ws._pairing) return; // still waiting for the user's decision
        clearTimeout(helloTimer);
        if (msg.type !== 'hello') { ws.close(4001, 'expected hello'); return; }
        const client = this.#makeClient(ws, origin, msg.info || {});
        if (this.#sameToken(msg.token) || this.devices.has(msg.token)) {
          this.#accept(client);
        } else {
          this.#requestPairing(client);
        }
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
      for (const [pid, p] of this.pairing) if (p.client.ws === ws) this.pairing.delete(pid);
      const c = this.clients.get(ws);
      if (c) log(`document #${c.id} disconnected: ${c.title}`);
      this.clients.delete(ws);
    });
  }

  #makeClient(ws, origin, info) {
    return {
      id: null,
      ws,
      origin,
      editor: info.editorType || null,
      title: info.documentTitle || null,
      documentId: info.documentId || null,
      pluginVersion: info.pluginVersion || null,
      connectedAt: new Date(),
      lastUsedAt: new Date(),
    };
  }

  #accept(client, newToken) {
    const { ws } = client;
    if (ws.readyState !== ws.OPEN) return;
    client.id = this.nextId++;
    // Several panels may share a document id (e.g. the same file open twice):
    // keep them all, they are told apart by connection id.
    this.clients.set(ws, client);
    ws.send(JSON.stringify(newToken ? { type: 'welcome', token: newToken } : { type: 'welcome' }));
    log(`document #${client.id} connected: ${client.title} (${client.editor}, plugin ${client.pluginVersion})`);
  }

  async #requestPairing(client) {
    const { ws } = client;
    if (Date.now() < this.deniedUntil) { ws.close(4003, 'denied'); return; }
    const id = this.nextPairingId++;
    ws._pairing = true;
    ws.send(JSON.stringify({ type: 'pending', pairingId: id }));
    log(`pairing request #${id} from ${client.origin}: ${client.title}`);

    const decision = new Promise((resolve) => this.pairing.set(id, { client, resolve }));
    const kinds = { word: 'document', cell: 'spreadsheet', slide: 'presentation' };
    askNative(
      `An OnlyOffice ${kinds[client.editor] || 'document'} wants to connect to Claude:\n\n`
      + `“${client.title || 'Untitled'}”\n\nAllow Claude to read and edit documents opened with this plugin in this browser?`,
    ).then((answer) => {
      if (answer !== null) this.decidePairing(id, answer);
      else log(`pairing request #${id}: no dialog available, waiting for approval via the approve_connection tool`);
    });

    const allowed = await decision;
    ws._pairing = false;
    if (!allowed) {
      this.deniedUntil = Date.now() + 30000;
      log(`pairing request #${id} denied`);
      try { ws.close(4003, 'denied'); } catch {}
      return;
    }
    const token = this.devices.add(`${client.origin} — first used with “${client.title}”`);
    log(`pairing request #${id} approved`);
    this.#accept(client, token);
  }

  // Resolves a pending pairing request (from the native dialog or an MCP tool).
  decidePairing(id, allow) {
    const p = this.pairing.get(id);
    if (!p) return false;
    this.pairing.delete(id);
    p.resolve(!!allow);
    return true;
  }

  listPairingRequests() {
    return [...this.pairing.entries()].map(([id, p]) => ({
      id, title: p.client.title, editor: p.client.editor, origin: p.client.origin,
    }));
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

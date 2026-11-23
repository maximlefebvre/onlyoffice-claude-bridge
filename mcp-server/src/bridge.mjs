// WebSocket hub between this MCP server and the Claude Bridge plugins running
// inside OnlyOffice editors.

import { WebSocket, WebSocketServer } from 'ws';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { execFile } from 'node:child_process';
import { homedir } from 'node:os';
import { join } from 'node:path';

const log = (...args) => console.error('[claude-bridge]', ...args); // stdout is reserved for MCP

export const MIN_PLUGIN_VERSION = '0.2.0';

// ONLYOFFICE Desktop Editors load locally installed plugins from files, so the
// WebSocket handshake carries a non-web origin ("file://" or "null") instead of
// the GitHub Pages one. Accept those only from the Desktop app's own browser
// engine (its user agent contains "AscDesktopEditor"): a web page cannot change
// the user agent of a WebSocket handshake, and a new browser still has to be
// approved by the user before it gets a token.
export function isDesktopEditor(origin, userAgent) {
  if (!/AscDesktopEditor/i.test(userAgent || '')) return false;
  const o = String(origin || 'null').toLowerCase();
  return o === 'null' || o.startsWith('file:');
}

function isCompatible(version) {
  const parse = (v) => String(v || '0').split('.').map((n) => parseInt(n, 10) || 0);
  const [a, b] = [parse(version), parse(MIN_PLUGIN_VERSION)];
  for (let i = 0; i < 3; i++) {
    if ((a[i] || 0) !== (b[i] || 0)) return (a[i] || 0) > (b[i] || 0);
  }
  return true;
}

const CONFIG_DIR = join(homedir(), '.config', 'onlyoffice-claude-bridge');

// Reads a secret from the config folder, creating it on first use.
// "token": manual pairing token for the editor plugin.
// "mcp-token": bearer token of the shared HTTP MCP endpoint.
export function loadToken(name = 'token') {
  const file = join(CONFIG_DIR, name);
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
// Runs a dialog command. Resolves true (allow), false (deny), null when it
// fails or times out, or undefined when the command is not installed.
function runDialog(cmd, args, { allow, deny, parse, env } = {}) {
  return new Promise((resolve) => {
    execFile(cmd, args, { timeout: 130000, env: { ...process.env, ...env }, windowsHide: false }, (err, stdout) => {
      const code = err ? err.code : 0;
      if (code === 'ENOENT') { resolve(undefined); return; }
      if (err && (err.killed || typeof code !== 'number')) { resolve(null); return; }
      if (parse) { resolve(parse(code, String(stdout))); return; }
      resolve(code === allow ? true : code === deny ? false : null);
    });
  });
}

// Asks the user on this computer. Resolves true/false, or null when no dialog
// can be shown (then the request waits for the approve_connection tool).
export async function askNative(message) {
  if (process.env.CLAUDE_BRIDGE_NO_DIALOG) return null;
  return (await showDialog(message)) ?? null;
}

async function showDialog(message) {
  const title = 'Claude Bridge';

  if (process.platform === 'darwin') {
    const esc = (t) => String(t).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    const script = `display dialog "${esc(message)}" with title "${title}" `
      + 'buttons {"Deny", "Allow"} default button "Allow" cancel button "Deny" '
      + 'with icon caution giving up after 120';
    // Exit code 1 = "Deny" (user cancelled).
    return runDialog('osascript', ['-e', script], {
      parse: (code, out) => (code === 0 ? /button returned:Allow/.test(out) && !/gave up:true/.test(out) : code === 1 ? false : null),
    });
  }

  if (process.platform === 'win32') {
    // The message goes through an environment variable: no quoting issues.
    const ps = [
      'Add-Type -AssemblyName System.Windows.Forms',
      '$owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }',
      "$r = [System.Windows.Forms.MessageBox]::Show($owner, $env:CLAUDE_BRIDGE_DIALOG_TEXT, $env:CLAUDE_BRIDGE_DIALOG_TITLE, 'YesNo', 'Warning')",
      'Write-Output $r',
    ].join('; ');
    return runDialog('powershell.exe', ['-NoProfile', '-NonInteractive', '-STA', '-Command', ps], {
      env: { CLAUDE_BRIDGE_DIALOG_TEXT: message, CLAUDE_BRIDGE_DIALOG_TITLE: title },
      parse: (code, out) => (code !== 0 ? null : /\bYes\b/.test(out) ? true : /\bNo\b/.test(out) ? false : null),
    });
  }

  if (process.env.DISPLAY || process.env.WAYLAND_DISPLAY) {
    const markup = String(message).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // zenity: 0 = OK, 1 = Cancel, 5 = timeout.
    const z = await runDialog('zenity', ['--question', `--title=${title}`, `--text=${markup}`,
      '--ok-label=Allow', '--cancel-label=Deny', '--icon-name=dialog-warning', '--timeout=120'], { allow: 0, deny: 1 });
    if (z !== undefined) return z; // only fall back when zenity is not installed
    return runDialog('kdialog', ['--title', title, '--warningyesno', message, '--yes-label', 'Allow', '--no-label', 'Deny'], { allow: 0, deny: 1 });
  }
  return null;
}

// How long to retry binding the port while another process holds it.
const LISTEN_RETRY_MS = 5000;
// After a document disconnects, how long a tool call waits for it (or another
// document) to come back before reporting that nothing is connected. Covers
// plugin reloads and short network hiccups.
const RECONNECT_WINDOW_MS = 60000;

export class Bridge {
  constructor({ port, host = '127.0.0.1', allowedOrigins, token, timeoutMs = 20000, graceMs = 6000, version = '' }) {
    this.version = version;
    this.port = port;
    this.host = host;
    this.allowedOrigins = allowedOrigins;
    this.token = token;
    this.timeoutMs = timeoutMs;
    this.graceMs = graceMs;
    this.clients = new Map(); // ws -> client
    this.pending = new Map(); // request id -> { resolve, reject, timer }
    this.nextId = 1;
    this.knownIds = new Map(); // document key -> id, so a reconnecting document keeps its id
    this.waiters = new Set(); // callbacks run when a document connects
    this.lastDisconnectAt = 0;
    this.listenError = null;
    this.portHolder = null; // what holds the port when it is taken (from a probe)
    this.devices = new DeviceStore();
    this.pairing = new Map(); // pairing id -> { ws, client, resolve }
    this.nextPairingId = 1;
    this.deniedUntil = 0; // after a denial, ignore new requests for a short while
  }

  start() {
    this.#listen();
  }

  // Binds the port. When it is taken (usually by the server of another Claude
  // session), keeps retrying so this session takes over as soon as it is free.
  #listen() {
    const wss = new WebSocketServer({
      host: this.host,
      port: this.port,
      verifyClient: ({ origin, req }, done) => {
        const ua = req.headers['user-agent'] || '';
        const ok = this.allowedOrigins.includes('*') || this.allowedOrigins.includes(origin)
          || isDesktopEditor(origin, ua);
        if (!ok) log(`connection refused, origin not allowed: ${origin} (user agent: ${ua})`);
        done(ok, 403, 'Origin not allowed');
      },
    });
    this.wss = wss;
    wss.on('listening', () => {
      const tookOver = !!this.listenError;
      this.listenError = null;
      this.portHolder = null;
      log(`listening on ws://${this.host}:${this.port}${tookOver ? ' (the port was released by the other process)' : ''}`);
    });
    wss.on('error', (err) => {
      const first = !this.listenError;
      this.listenError = err;
      if (err.code === 'EADDRINUSE') {
        this.#probePortHolder().then((holder) => {
          this.portHolder = holder;
          if (first) log(this.#portTakenMessage() + ` Retrying every ${LISTEN_RETRY_MS / 1000}s.`);
        });
        try { wss.close(); } catch {}
        setTimeout(() => this.#listen(), LISTEN_RETRY_MS).unref();
      } else {
        log(`WebSocket server error: ${err.message}`);
      }
    });
    wss.on('connection', (ws, req) => this.#onConnection(ws, req));
  }

  // Asks whatever holds the port who it is. Resolves { bridge: true, version, pid }
  // for another Claude Bridge server, { bridge: false } otherwise.
  #probePortHolder() {
    return new Promise((resolve) => {
      let done = false;
      const finish = (v) => { if (!done) { done = true; clearTimeout(timer); try { sock.terminate(); } catch {} resolve(v); } };
      const origin = this.allowedOrigins.find((o) => o !== '*') || 'http://localhost';
      const sock = new WebSocket(`ws://${this.host}:${this.port}`, { origin });
      const timer = setTimeout(() => finish({ bridge: false }), 2000);
      sock.on('open', () => sock.send(JSON.stringify({ type: 'probe' })));
      sock.on('message', (raw) => {
        try {
          const m = JSON.parse(raw);
          if (m.type === 'probe_ok') finish({ bridge: true, version: m.version || null, pid: m.pid || null });
        } catch {}
      });
      sock.on('error', () => finish({ bridge: false }));
      sock.on('close', () => finish({ bridge: false }));
    });
  }

  #portTakenMessage() {
    const h = this.portHolder;
    if (h && h.bridge) {
      return `Port ${this.port} is held by another Claude Bridge server`
        + `${h.pid ? ` (pid ${h.pid}` : ' ('}${h.version ? `${h.pid ? ', ' : ''}version ${h.version}` : ''})`
        + ', most likely started by another Claude session. OnlyOffice documents connect to that session only.'
        + ' Close the other session (or stop that process): this session takes over the port automatically within a few seconds.';
    }
    return `Port ${this.port} is used by another program, so OnlyOffice documents cannot connect to this session.`
      + ' Free the port or set CLAUDE_BRIDGE_PORT (and the same port under Advanced in the plugin panel).';
  }

  #assertListening() {
    if (!this.listenError) return;
    if (this.listenError.code === 'EADDRINUSE') throw new Error(this.#portTakenMessage());
    throw new Error(`The bridge is not listening on port ${this.port}: ${this.listenError.message}`);
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
        if (msg.type === 'probe') {
          // Connection check from the setup page: no access to documents.
          ws.send(JSON.stringify({ type: 'probe_ok', server: 'onlyoffice-claude-bridge', version: this.version, pid: process.pid }));
          ws.close(1000, 'probe');
          return;
        }
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
      if (c) {
        log(`document #${c.id} disconnected: ${c.title}`);
        this.lastDisconnectAt = Date.now();
      }
      this.clients.delete(ws);
      // Fail the calls still waiting on this editor instead of letting them time out.
      for (const [id, p] of this.pending) {
        if (p.ws !== ws) continue;
        this.pending.delete(id);
        clearTimeout(p.timer);
        p.reject(new Error('The document disconnected before answering. It reconnects by itself: try again in a moment.'));
      }
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
    // A document that reconnects (plugin reload, server restart of the relay…)
    // keeps its previous id, unless that id is still in use: several plugins
    // may share a document id (e.g. the same file open twice), and they are
    // told apart by connection id.
    const key = client.documentId ? `${client.origin}|${client.documentId}` : null;
    const previous = key ? this.knownIds.get(key) : undefined;
    const inUse = (id) => [...this.clients.values()].some((c) => c.id === id);
    client.id = previous !== undefined && !inUse(previous) ? previous : this.nextId++;
    if (key && previous === undefined) this.knownIds.set(key, client.id);
    this.clients.set(ws, client);
    ws.send(JSON.stringify(newToken ? { type: 'welcome', token: newToken } : { type: 'welcome' }));
    log(`document #${client.id} connected: ${client.title} (${client.editor}, plugin ${client.pluginVersion})`);
    for (const w of [...this.waiters]) w();
  }

  // Resolves once `match()` returns something (checked now and each time a
  // document connects), or with null after `ms`.
  #waitFor(match, ms) {
    const now = match();
    if (now || ms <= 0) return Promise.resolve(now || null);
    return new Promise((resolve) => {
      const check = () => {
        const v = match();
        if (v) { cleanup(); resolve(v); }
      };
      const timer = setTimeout(() => { cleanup(); resolve(null); }, ms);
      const cleanup = () => { clearTimeout(timer); this.waiters.delete(check); };
      this.waiters.add(check);
    });
  }

  // Grace period: wait only when a document went away recently (it is most
  // likely reconnecting); otherwise answer at once.
  #graceMs() {
    return Date.now() - this.lastDisconnectAt < RECONNECT_WINDOW_MS ? this.graceMs : 0;
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

  // Snapshot for the menu bar app: no waiting, never throws.
  status() {
    const h = this.portHolder;
    return {
      port: this.port,
      listening: !this.listenError,
      portHolder: this.listenError ? (h ? { pid: h.pid || null, bridge: !!h.bridge, version: h.version || null } : {}) : null,
      documents: [...this.clients.values()].map((c) => ({ id: c.id, title: c.title, editor: c.editor })),
    };
  }

  async listDocuments() {
    this.#assertListening();
    await this.#waitFor(() => this.clients.size > 0, this.#graceMs());
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
  // used one (optionally restricted to an editor type). When nothing matches
  // right after a disconnection, waits a few seconds for the plugin to
  // reconnect before failing.
  async resolve(document, editor) {
    this.#assertListening();
    let error = null;
    const attempt = () => {
      try { error = null; return this.#pick(document, editor); } catch (e) { error = e; return null; }
    };
    const target = await this.#waitFor(attempt, this.#graceMs());
    if (target) return target;
    throw error || new Error('No OnlyOffice document is connected.');
  }

  #pick(document, editor) {
    let candidates = [...this.clients.values()];
    if (!candidates.length) {
      throw new Error('No OnlyOffice document is connected. Ask the user to open the document in OnlyOffice: the Claude Bridge plugin connects by itself once it is enabled under Plugins → Background plugins. Its panel (Claude tab → Claude Bridge) shows the status and a Connect button.');
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
      if (client.ws.readyState !== client.ws.OPEN) {
        reject(new Error('The document disconnected. It reconnects by itself: try again in a moment.'));
        return;
      }
      const id = randomUUID();
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`The editor did not answer within ${this.timeoutMs / 1000}s`));
      }, this.timeoutMs);
      this.pending.set(id, { resolve, reject, timer, ws: client.ws });
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

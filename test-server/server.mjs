// WebSocket test server for step 1 (no MCP yet).
// Start it, connect the plugin, then type commands: ping, info, sel.
//
//   node server.mjs
//
// Environment variables:
//   PORT             (default 47615)
//   ALLOWED_ORIGINS  comma-separated list (default: GitHub Pages + localhost)

import { WebSocketServer } from 'ws';
import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import readline from 'node:readline';

const PORT = parseInt(process.env.PORT || '47615', 10);
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS ||
  'https://maximlefebvre.github.io,http://localhost:8080,http://127.0.0.1:8080')
  .split(',').map(s => s.trim()).filter(Boolean);

// Persistent token: generated once, pasted once into the plugin.
const dir = join(homedir(), '.config', 'onlyoffice-claude-bridge');
const tokenFile = join(dir, 'token');
mkdirSync(dir, { recursive: true, mode: 0o700 });
if (!existsSync(tokenFile)) writeFileSync(tokenFile, randomBytes(24).toString('base64url'), { mode: 0o600 });
const TOKEN = readFileSync(tokenFile, 'utf8').trim();

const sameToken = (t) => {
  const a = Buffer.from(String(t || '')), b = Buffer.from(TOKEN);
  return a.length === b.length && timingSafeEqual(a, b);
};

const clients = new Map(); // ws -> { id, info }
const pending = new Map(); // requestId -> { resolve, reject, timer }

const wss = new WebSocketServer({
  host: '127.0.0.1',
  port: PORT,
  verifyClient: ({ origin }, done) => {
    const ok = ALLOWED_ORIGINS.includes(origin);
    if (!ok) console.log(`✗ connection refused, origin not allowed: ${origin}`);
    done(ok, 403, 'Origin not allowed');
  },
});

wss.on('connection', (ws, req) => {
  const origin = req.headers.origin;
  const helloTimer = setTimeout(() => ws.close(4001, 'no hello'), 5000);

  ws.on('message', (raw) => {
    let msg;
    try { msg = JSON.parse(raw); } catch { return; }

    if (!clients.has(ws)) {
      clearTimeout(helloTimer);
      if (msg.type !== 'hello' || !sameToken(msg.token)) {
        console.log(`✗ invalid token from ${origin}`);
        ws.close(4001, 'bad token');
        return;
      }
      const c = { id: clients.size + 1, info: msg.info || {} };
      clients.set(ws, c);
      ws.send(JSON.stringify({ type: 'welcome' }));
      console.log(`✓ plugin #${c.id} connected from ${origin}`, c.info);
      return;
    }

    const p = pending.get(msg.id);
    if (!p) return;
    pending.delete(msg.id);
    clearTimeout(p.timer);
    msg.error ? p.reject(new Error(msg.error.message)) : p.resolve(msg.result);
  });

  ws.on('close', () => {
    clearTimeout(helloTimer);
    const c = clients.get(ws);
    if (c) console.log(`– plugin #${c.id} disconnected`);
    clients.delete(ws);
  });
});

function call(ws, method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = randomUUID();
    const timer = setTimeout(() => { pending.delete(id); reject(new Error('timed out')); }, 10000);
    pending.set(id, { resolve, reject, timer });
    ws.send(JSON.stringify({ id, method, params }));
  });
}

wss.on('listening', () => {
  console.log(`Test server on ws://127.0.0.1:${PORT}`);
  console.log(`Allowed origins: ${ALLOWED_ORIGINS.join(', ')}`);
  console.log(`Token to paste into the plugin: ${TOKEN}`);
  console.log('Commands: ping | info | sel | quit');
});

const aliases = { ping: 'ping', info: 'get_info', sel: 'get_selected_text' };
readline.createInterface({ input: process.stdin }).on('line', async (line) => {
  const cmd = line.trim();
  if (!cmd) return;
  if (cmd === 'quit') process.exit(0);
  const method = aliases[cmd];
  if (!method) return console.log('Commands: ping | info | sel | quit');
  if (!clients.size) return console.log('No plugin connected.');
  for (const [ws, c] of clients) {
    try { console.log(`#${c.id} →`, await call(ws, method)); }
    catch (e) { console.log(`#${c.id} error:`, e.message); }
  }
});

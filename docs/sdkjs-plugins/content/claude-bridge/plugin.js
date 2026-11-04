/*
 * Claude Bridge — OnlyOffice plugin.
 *
 * Connects over WebSocket to the local server (127.0.0.1) and runs the commands
 * it sends against the open document.
 *
 * Protocol (JSON):
 *   plugin -> server : { type: "hello", token?, info }
 *   server -> plugin : { type: "pending" }             (unknown browser: the user is asked to approve)
 *   server -> plugin : { type: "welcome", token? }     (token: issued on first approval, stored here)
 *   server closes with 4003 when the user denies the connection.
 *   server -> plugin : { id, method, params }
 *   plugin -> server : { id, result } | { id, error }
 *
 * Methods: ping, get_info, run_command { code, scope, recalc }, execute_method { name, args }
 */
(function (window) {
  'use strict';

  var PLUGIN_VERSION = '0.3.0';
  var STORAGE_KEY = 'claude-bridge-settings';

  var ws = null;
  var wantConnected = false;
  var retryDelay = 1000;
  var retryTimer = null;

  var $ = function (id) { return document.getElementById(id); };

  // ---------- UI ----------

  function setStatus(state, text) {
    $('dot').className = 'dot ' + ({ on: 'dot-on', wait: 'dot-wait' }[state] || 'dot-off');
    $('status').textContent = text;
  }

  function log(msg) {
    var el = $('log');
    var line = document.createElement('div');
    line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
    el.insertBefore(line, el.firstChild);
    while (el.childNodes.length > 50) el.removeChild(el.lastChild);
  }

  function loadSettings() {
    try {
      var s = JSON.parse(window.localStorage.getItem(STORAGE_KEY) || '{}');
      if (s.port) $('port').value = s.port;
      if (s.token) $('token').value = s.token;
      return s;
    } catch (e) { return {}; }
  }

  function saveSettings() {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
        port: $('port').value, token: $('token').value
      }));
    } catch (e) { /* storage unavailable: carry on without it */ }
  }

  // ---------- Document info ----------

  function docInfo() {
    var info = (window.Asc && window.Asc.plugin && window.Asc.plugin.info) || {};
    return {
      editorType: info.editorType || null,
      documentId: info.documentId || null,
      documentTitle: info.documentTitle || document.referrer || null,
      lang: info.lang || null,
      pluginVersion: PLUGIN_VERSION,
      insideEditor: !!(window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod)
    };
  }

  // ---------- Commands ----------

  function executeMethod(name, args) {
    return new Promise(function (resolve, reject) {
      if (!(window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod)) {
        reject(new Error('Not running inside an OnlyOffice editor'));
        return;
      }
      window.Asc.plugin.executeMethod(name, args || [], function (res) { resolve(res); });
    });
  }

  // Editor calls are serialized: OnlyOffice handles one command at a time.
  var queue = Promise.resolve();
  function enqueue(fn) {
    var p = queue.then(fn, fn);
    queue = p.catch(function () {});
    return p;
  }

  // Runs `code` (the source of a JS function) inside the editor with the
  // Document Builder API (`Api`). `scope` is exposed as `Asc.scope`.
  function runCommand(code, scope, recalc) {
    return enqueue(function () {
      return new Promise(function (resolve, reject) {
        if (!(window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand)) {
          reject(new Error('Not running inside an OnlyOffice editor'));
          return;
        }
        window.Asc.scope = scope || {};
        var wrapped =
          'function () {' +
          '  try {' +
          '    var __r = (' + code + ')();' +
          '    return JSON.stringify({ ok: true, value: __r === undefined ? null : __r });' +
          '  } catch (e) {' +
          '    return JSON.stringify({ ok: false, error: String((e && e.message) || e) });' +
          '  }' +
          '}';
        window.Asc.plugin.callCommand(wrapped, false, recalc !== false, function (res) {
          if (typeof res !== 'string') { resolve(res === undefined ? null : res); return; }
          var out;
          try { out = JSON.parse(res); } catch (e) { resolve(res); return; }
          if (out && out.ok === false) reject(new Error(out.error));
          else resolve(out ? out.value : null);
        });
      });
    });
  }

  function executeMethod(name, args) {
    return enqueue(function () {
      return new Promise(function (resolve, reject) {
        if (!(window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod)) {
          reject(new Error('Not running inside an OnlyOffice editor'));
          return;
        }
        window.Asc.plugin.executeMethod(name, args || [], function (res) {
          resolve(res === undefined ? null : res);
        });
      });
    });
  }

  // The plugin stays generic on purpose: every document action is written in
  // the local server and sent here as code, so new actions never require a
  // plugin update.
  var handlers = {
    ping: function () {
      return Promise.resolve({ pong: true, at: Date.now() });
    },
    get_info: function () {
      return Promise.resolve(docInfo());
    },
    run_command: function (params) {
      if (!params.code) return Promise.reject(new Error('Missing "code"'));
      return runCommand(String(params.code), params.scope, params.recalc);
    },
    execute_method: function (params) {
      if (!params.name) return Promise.reject(new Error('Missing "name"'));
      return executeMethod(String(params.name), params.args || []);
    }
  };

  function handleCommand(msg) {
    var h = handlers[msg.method];
    var p = h ? h(msg.params || {}) : Promise.reject(new Error('Unknown method: ' + msg.method));
    log('← ' + msg.method);
    p.then(function (result) {
      send({ id: msg.id, result: result });
    }, function (err) {
      log('✗ ' + msg.method + ': ' + err.message);
      send({ id: msg.id, error: { message: err.message } });
    });
  }

  // ---------- WebSocket ----------

  function send(obj) {
    if (ws && ws.readyState === WebSocket.OPEN) ws.send(JSON.stringify(obj));
  }

  function scheduleRetry() {
    if (!wantConnected) return;
    clearTimeout(retryTimer);
    retryTimer = setTimeout(connect, retryDelay);
    retryDelay = Math.min(retryDelay * 2, 15000);
  }

  function connect() {
    clearTimeout(retryTimer);
    var port = parseInt($('port').value, 10);
    var token = $('token').value.trim();
    if (!port) { setStatus('off', 'Invalid port'); return; }
    wantConnected = true;
    saveSettings();

    if (ws) { try { ws.close(); } catch (e) {} }
    setStatus('wait', 'Connecting…');

    var url = 'ws://127.0.0.1:' + port;
    try {
      ws = new WebSocket(url);
    } catch (e) {
      log('Cannot open ' + url + ': ' + e.message);
      setStatus('off', 'Connection failed');
      scheduleRetry();
      return;
    }

    var welcomed = false;
    var sock = ws;

    sock.onopen = function () {
      send({ type: 'hello', token: token || null, info: docInfo() });
    };

    sock.onmessage = function (ev) {
      var msg;
      try { msg = JSON.parse(ev.data); } catch (e) { return; }
      if (msg.type === 'pending') {
        setStatus('wait', 'Approve the connection on your computer…');
        log('Waiting for approval (a dialog should be open on your computer)');
        return;
      }
      if (msg.type === 'welcome') {
        welcomed = true;
        retryDelay = 1000;
        if (msg.token) {
          $('token').value = msg.token;
          saveSettings();
          log('Approved: this browser is now paired');
        }
        setStatus('on', 'Connected to Claude');
        log('Connected (' + url + ')');
        return;
      }
      if (!welcomed) return;
      if (msg.id !== undefined && msg.method) handleCommand(msg);
    };

    sock.onerror = function () {
      if (!welcomed) setStatus('wait', 'Waiting for Claude…');
      log('Cannot reach the local server (is Claude running with the bridge? is the browser blocking localhost?)');
    };

    sock.onclose = function (ev) {
      if (sock !== ws) return;
      ws = null;
      if (ev.code === 4003) {
        wantConnected = false;
        setStatus('off', 'Connection denied');
        log('The connection was denied on your computer. Click Connect to ask again.');
        return;
      }
      setStatus(wantConnected ? 'wait' : 'off', wantConnected ? 'Reconnecting…' : 'Disconnected');
      scheduleRetry();
    };
  }

  function disconnect() {
    wantConnected = false;
    clearTimeout(retryTimer);
    if (ws) { var s = ws; ws = null; s.close(); }
    setStatus('off', 'Disconnected');
  }

  // ---------- Startup ----------

  function start() {
    $('connect').onclick = connect;
    $('disconnect').onclick = disconnect;
    loadSettings();
    connect(); // connects (or asks for approval) automatically
  }

  var started = false;
  function startOnce() { if (!started) { started = true; start(); } }

  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = startOnce;
    window.Asc.plugin.button = function () { disconnect(); this.executeCommand('close', ''); };
  }
  // Outside the editor (testing in a plain browser): start anyway.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { setTimeout(startOnce, 300); });
  } else {
    setTimeout(startOnce, 300);
  }
})(window);

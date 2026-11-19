/*
 * Claude Bridge — OnlyOffice plugin, background part.
 *
 * This is a "background" plugin variation: the editor keeps it running for as
 * long as the document is open, whether or not any panel is shown. It owns the
 * connection to the local server (127.0.0.1) and runs the commands it sends
 * against the open document. The panel (panel.html) is UI only: it shows the
 * state pushed from here and sends user actions back.
 *
 * Protocol with the server (JSON over WebSocket):
 *   plugin -> server : { type: "hello", token?, info }
 *   server -> plugin : { type: "pending" }             (unknown browser: the user is asked to approve)
 *   server -> plugin : { type: "welcome", token? }     (token: issued on first approval, stored here)
 *   server closes with 4003 when the user denies the connection.
 *   server -> plugin : { id, method, params }
 *   plugin -> server : { id, result } | { id, error }
 *
 * Methods: ping, get_info, run_command { code, scope, recalc }, execute_method { name, args }
 *
 * Protocol with the panel (OnlyOffice plugin window messages):
 *   panel -> here : "onPanelReady" | "onPanelAction" { type: "connect" | "disconnect" | "settings" | "log", ... }
 *   here -> panel : "onState" { status, text, hint, log, port, token, mode, relayKey }
 *
 * Transport: a direct WebSocket when the browser allows it. Browsers with
 * Local Network Access (Firefox 151+, recent Chrome) block it from inside the
 * editor and cannot ask there, so a small relay window (relay.html, same
 * origin, top-level) holds the WebSocket instead. The panel opens it on a
 * click (pop-ups need a user gesture); the relay then finds this frame by
 * itself and talks to it directly, so it survives the panel being hidden.
 *   relay -> here : { cb: "relay-hello", key } then { cb: "open" | "message" | "error" | "closed" }
 *   here -> relay : { cb: "relay-host", key } then { cb: "connect", url, title } | { cb: "send", data } | { cb: "close" }
 */
(function (window) {
  'use strict';

  var PLUGIN_VERSION = '0.5.0';
  var STORAGE_KEY = 'claude-bridge-settings';
  var ORIGIN = window.location.origin;
  var ICONS = 'resources/%theme-type%(light|dark)/icon%scale%(default).svg';

  var settings = { port: '47615', token: '', mode: 'direct' }; // mode: 'direct' or 'relay'
  var ws = null;
  var wantConnected = false;
  var retryDelay = 1000;
  var retryTimer = null;
  var relayWin = null;
  var panelWin = null;

  // State shown by the panel.
  var state = { status: 'off', text: 'Starting…', hint: '', log: [] };

  // ---------- State and panel ----------

  function pushState() {
    if (!panelWin) return;
    try {
      panelWin.command('onState', {
        status: state.status,
        text: state.text,
        hint: state.hint,
        log: state.log,
        port: settings.port,
        token: settings.token,
        mode: settings.mode,
        relayKey: docKey(),
        relayOpen: relayAlive()
      });
    } catch (e) { /* panel is going away */ }
  }

  function setStatus(status, text) {
    state.status = status;
    state.text = text;
    pushState();
  }

  function hint(text) {
    state.hint = text || '';
    pushState();
  }

  function log(msg) {
    state.log.unshift(new Date().toLocaleTimeString() + '  ' + msg);
    if (state.log.length > 50) state.log.length = 50;
    pushState();
  }

  function rawExecuteMethod(name, args, cb) {
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod) {
      window.Asc.plugin.executeMethod(name, args || [], cb);
    }
  }

  // Shows the panel. `activate` expands it (or brings it forward); otherwise it
  // only gets its button in the left bar (OnlyOffice 8.3+ honours isActivated).
  function showPanel(activate) {
    if (!(window.Asc && window.Asc.PluginWindow)) return;
    if (panelWin) {
      if (activate && typeof panelWin.activate === 'function') panelWin.activate();
      return;
    }
    var w = new window.Asc.PluginWindow();
    w.attachEvent('onPanelReady', function () { pushState(); });
    w.attachEvent('onPanelAction', onPanelAction);
    panelWin = w;
    w.show({
      url: 'panel.html',
      description: 'Claude Bridge',
      isVisual: true,
      isModal: false,
      isViewer: false,
      EditorsSupport: ['word', 'cell', 'slide'],
      type: 'panel',
      icons: ICONS,
      buttons: [],
      size: [300, 400],
      isActivated: !!activate
    });
  }

  function onPanelAction(a) {
    a = a || {};
    if (a.type === 'connect') {
      if (a.port) settings.port = String(a.port);
      if (a.token !== undefined) settings.token = String(a.token).trim();
      saveSettings();
      retryDelay = 1000;
      if (a.relayOpening && !relayAlive()) {
        // The panel just opened the relay window: it announces itself shortly.
        wantConnected = true;
        clearTimeout(retryTimer);
        if (ws) { var prev = ws; ws = null; try { prev.close(); } catch (e) {} }
        setStatus('wait', 'Waiting for the relay window…');
        hint('');
        return;
      }
      connect();
    } else if (a.type === 'disconnect') {
      disconnect();
    } else if (a.type === 'settings') {
      if (a.port) settings.port = String(a.port);
      if (a.token !== undefined) settings.token = String(a.token).trim();
      saveSettings();
      pushState();
    } else if (a.type === 'log' && a.text) {
      log(String(a.text));
    }
  }

  function addToolbarButton() {
    if (!(window.Asc && window.Asc.ButtonToolbar && window.Asc.Buttons)) return;
    try {
      var tab = new window.Asc.ButtonToolbar(null);
      tab.text = 'Claude';
      var btn = new window.Asc.ButtonToolbar(tab);
      btn.text = 'Claude Bridge';
      btn.hint = 'Show the Claude Bridge panel (connection status)';
      btn.icons = ICONS;
      btn.attachOnClick(function () { showPanel(true); });
      window.Asc.Buttons.registerToolbarMenu();
    } catch (e) { /* older editor: the panel is still shown at start */ }
  }

  // ---------- Settings ----------

  function loadSettings() {
    try {
      var s = JSON.parse(window.localStorage.getItem(STORAGE_KEY) || '{}');
      if (s.port) settings.port = String(s.port);
      if (s.token) settings.token = String(s.token);
      if (s.mode === 'relay') settings.mode = 'relay';
    } catch (e) { /* storage unavailable: defaults */ }
  }

  function saveSettings() {
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    } catch (e) { /* storage unavailable: carry on without it */ }
  }

  // ---------- Document info ----------

  function docInfo() {
    var info = (window.Asc && window.Asc.plugin && window.Asc.plugin.info) || {};
    return {
      editorType: info.editorType || null,
      documentId: info.documentId || null,
      documentTitle: info.documentTitle || null,
      lang: info.lang || null,
      pluginVersion: PLUGIN_VERSION,
      insideEditor: !!(window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod)
    };
  }

  // Identifies this document to its relay window.
  function docKey() {
    var i = docInfo();
    return String(i.documentId || i.documentTitle || 'default');
  }

  // ---------- Commands ----------

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

  // ---------- Relay window ----------

  function relayAlive() {
    try { return !!(relayWin && !relayWin.closed); } catch (e) { return false; }
  }

  function toRelay(msg) {
    if (relayAlive()) relayWin.postMessage(msg, ORIGIN);
  }

  var relaySocket = null; // the RelaySocket currently in use, if any

  // Messages from the relay window. It announces itself with "relay-hello"
  // and the document key, since it was opened by the panel, not by this frame.
  window.addEventListener('message', function (ev) {
    if (ev.origin !== ORIGIN) return;
    var m = ev.data;
    if (!m || typeof m !== 'object' || typeof m.cb !== 'string') return;
    if (m.cb === 'relay-hello') {
      if (String(m.key) !== docKey()) return;
      var isNew = relayWin !== ev.source;
      relayWin = ev.source;
      toRelay({ cb: 'relay-host', key: docKey() });
      if (isNew) log('Relay window found.');
      // A different relay window than the one carrying the socket: start over through it.
      if (isNew && relaySocket && relaySocket.readyState !== 0) relaySocket._closed(1006, 'relay window replaced');
      else if (relaySocket) relaySocket._connect();
      else if (settings.mode === 'relay' && wantConnected && !ws) connect();
      pushState();
      return;
    }
    if (ev.source !== relayWin || !relaySocket) return;
    relaySocket._onRelay(m);
  });

  // A WebSocket look-alike that goes through the relay window.
  function RelaySocket(url) {
    var self = this;
    this.readyState = 0; // CONNECTING
    this.url = url;
    this._title = docInfo().documentTitle || '';
    relaySocket = this;
    this._watch = setInterval(function () {
      if (!relayAlive()) { relayWin = null; self._closed(4100, 'relay window closed'); }
    }, 1000);
    this._connect();
  }
  RelaySocket.prototype._onRelay = function (m) {
    if (m.cb === 'open') { this.readyState = 1; if (this.onopen) this.onopen({}); return; }
    if (m.cb === 'message') { if (this.onmessage) this.onmessage({ data: m.data }); return; }
    if (m.cb === 'error') { if (this.onerror) this.onerror({}); return; }
    if (m.cb === 'closed') this._closed(m.code || 1006, m.reason || '');
  };
  RelaySocket.prototype._connect = function () {
    if (this.readyState === 0) toRelay({ cb: 'connect', url: this.url, title: this._title });
  };
  RelaySocket.prototype.send = function (data) {
    if (this.readyState === 1) toRelay({ cb: 'send', data: data });
  };
  RelaySocket.prototype.close = function () {
    if (this.readyState === 3) return;
    toRelay({ cb: 'close' });
    this._closed(1000, '');
  };
  RelaySocket.prototype._closed = function (code, reason) {
    if (this.readyState === 3) return;
    this.readyState = 3;
    clearInterval(this._watch);
    if (relaySocket === this) relaySocket = null;
    if (this.onclose) this.onclose({ code: code, reason: reason });
  };

  // Whether this frame is known to be allowed to reach the computer. null when
  // the browser has no such permission (no Local Network Access).
  function localAccessState() {
    if (!navigator.permissions || !navigator.permissions.query) return Promise.resolve(null);
    var names = ['loopback-network', 'local-network-access'];
    function next(i) {
      if (i >= names.length) return Promise.resolve(null);
      return navigator.permissions.query({ name: names[i] })
        .then(function (r) { return r.state; }, function () { return next(i + 1); });
    }
    return next(0);
  }

  // Relay mode without a relay window: wait for the user to click Connect in
  // the panel (opening a pop-up needs a click). The relay window, once open,
  // finds this frame by itself and the connection resumes.
  function needRelayClick() {
    clearTimeout(retryTimer);
    setStatus('wait', 'Click Connect');
    hint('Your browser blocks the direct connection from the editor. Click Connect: a small window opens and relays the connection. Keep it open while you work.');
    showPanel(true);
  }

  // ---------- WebSocket ----------

  function send(obj) {
    if (ws && ws.readyState === 1) ws.send(JSON.stringify(obj));
  }

  function scheduleRetry() {
    if (!wantConnected) return;
    clearTimeout(retryTimer);
    retryTimer = setTimeout(connect, retryDelay);
    retryDelay = Math.min(retryDelay * 2, 15000);
  }

  function connect() {
    clearTimeout(retryTimer);
    var port = parseInt(settings.port, 10);
    var token = settings.token;
    if (!port) { setStatus('off', 'Invalid port'); return; }
    wantConnected = true;

    if (ws) { var prev = ws; ws = null; try { prev.close(); } catch (e) {} }

    var url = 'ws://127.0.0.1:' + port;
    if (settings.mode === 'relay' && !relayAlive()) { needRelayClick(); return; }
    setStatus('wait', 'Connecting…');
    hint(settings.mode === 'relay' ? 'Connected through the relay window: keep it open while you work.' : '');
    try {
      ws = settings.mode === 'relay' ? new RelaySocket(url) : new WebSocket(url);
    } catch (e) {
      log('Cannot open ' + url + ': ' + e.message);
      setStatus('off', 'Connection failed');
      scheduleRetry();
      return;
    }

    var welcomed = false;
    var opened = false;
    var sock = ws;

    sock.onopen = function () {
      opened = true;
      send({ type: 'hello', token: token || null, info: docInfo() });
    };

    sock.onmessage = function (ev) {
      if (sock !== ws) return;
      var msg;
      try { msg = JSON.parse(ev.data); } catch (e) { return; }
      if (msg.type === 'pending') {
        setStatus('wait', 'Approve the connection on your computer…');
        log('Waiting for approval (a dialog should be open on your computer)');
        showPanel(true);
        return;
      }
      if (msg.type === 'welcome') {
        welcomed = true;
        retryDelay = 1000;
        if (msg.token) {
          settings.token = msg.token;
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
      if (sock !== ws) return;
      if (!welcomed) setStatus('wait', 'Waiting for Claude…');
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
      if (ev.code === 4100) { log('The relay window was closed.'); needRelayClick(); return; }
      if (welcomed) log('Disconnected from Claude: reconnecting…');
      else if (retryDelay <= 2000) log('Cannot reach the local server yet (is a Claude session running with the bridge?). Retrying in the background.');
      if (settings.mode === 'direct' && !opened) {
        // Blocked by the browser, or the server is not running? Ask the browser.
        localAccessState().then(function (access) {
          if (access && access !== 'granted' && ws === null && wantConnected && settings.mode === 'direct') {
            settings.mode = 'relay';
            saveSettings();
            log('The browser blocks local connections from the editor: switching to the relay window.');
            needRelayClick();
          }
        });
      }
      setStatus(wantConnected ? 'wait' : 'off', wantConnected ? (welcomed ? 'Reconnecting…' : 'Waiting for Claude…') : 'Disconnected');
      scheduleRetry();
    };
  }

  function disconnect() {
    wantConnected = false;
    clearTimeout(retryTimer);
    if (ws) { var s = ws; ws = null; try { s.close(); } catch (e) {} }
    if (relayAlive()) { try { relayWin.close(); } catch (e) {} }
    relayWin = null;
    hint('');
    setStatus('off', 'Disconnected');
    log('Disconnected. Click Connect to reconnect.');
  }

  // ---------- Startup ----------

  function start() {
    loadSettings();
    addToolbarButton();
    showPanel(false);
    connect(); // connects (or asks for approval) automatically, and keeps reconnecting
  }

  var started = false;
  function startOnce() { if (!started) { started = true; start(); } }

  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = startOnce;
    // Called when a window of this plugin is closed (its id is `windowId`).
    // Closing the panel only hides the UI: the connection stays up here.
    window.Asc.plugin.button = function (id, windowId) {
      if (!windowId) return;
      if (panelWin && panelWin.id === windowId) panelWin = null;
      rawExecuteMethod('CloseWindow', [windowId]);
    };
  }
})(window);

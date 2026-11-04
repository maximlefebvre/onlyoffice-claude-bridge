/*
 * Claude Bridge — OnlyOffice plugin.
 *
 * Connects over WebSocket to the local server (127.0.0.1) and runs the commands
 * it sends against the open document.
 *
 * Protocol (JSON):
 *   plugin -> server : { type: "hello", token, info }
 *   server -> plugin : { type: "welcome" }             (otherwise the connection is closed)
 *   server -> plugin : { id, method, params }
 *   plugin -> server : { id, result } | { id, error }
 */
(function (window) {
  'use strict';

  var PLUGIN_VERSION = '0.1.1';
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

  var handlers = {
    ping: function () {
      return Promise.resolve({ pong: true, at: Date.now() });
    },
    get_info: function () {
      return Promise.resolve(docInfo());
    },
    get_selected_text: function () {
      return executeMethod('GetSelectedText', [{
        Numbering: false, Math: false,
        TableCellSeparator: '\n', ParaSeparator: '\n', TabSymbol: '\t'
      }]).then(function (text) { return { text: text || '' }; });
    },
    // Inserts text at the cursor; each line becomes a paragraph.
    insert_text: function (params) {
      var text = String(params.text || '');
      if (!text) return Promise.reject(new Error('Missing "text"'));
      window.Asc.scope.paragraphs = text.split(/\r?\n/);
      return callCommand(function () {
        var doc = Api.GetDocument();
        var content = Asc.scope.paragraphs.map(function (line) {
          var p = Api.CreateParagraph();
          p.AddText(line);
          return p;
        });
        doc.InsertContent(content);
        return content.length;
      }).then(function (count) { return { inserted: count }; });
    }
  };

  function callCommand(fn) {
    return new Promise(function (resolve, reject) {
      if (!(window.Asc && window.Asc.plugin && window.Asc.plugin.callCommand)) {
        reject(new Error('Not running inside an OnlyOffice editor'));
        return;
      }
      window.Asc.plugin.callCommand(fn, false, true, function (res) { resolve(res); });
    });
  }

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
    if (!port || !token) { setStatus('off', 'Port and token required'); return; }
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
      send({ type: 'hello', token: token, info: docInfo() });
    };

    sock.onmessage = function (ev) {
      var msg;
      try { msg = JSON.parse(ev.data); } catch (e) { return; }
      if (msg.type === 'welcome') {
        welcomed = true;
        retryDelay = 1000;
        setStatus('on', 'Connected to Claude');
        log('Connected (' + url + ')');
        return;
      }
      if (!welcomed) return;
      if (msg.id !== undefined && msg.method) handleCommand(msg);
    };

    sock.onerror = function () {
      log('WebSocket error (is the server running? is the browser blocking localhost?)');
    };

    sock.onclose = function (ev) {
      if (sock !== ws) return;
      ws = null;
      if (ev.code === 4001) {
        wantConnected = false;
        setStatus('off', 'Token rejected');
        log('Token rejected by the server');
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
    var s = loadSettings();
    if (s.token) connect();
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

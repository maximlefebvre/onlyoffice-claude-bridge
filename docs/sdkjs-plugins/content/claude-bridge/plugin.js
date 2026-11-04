/*
 * Claude Bridge — plugin OnlyOffice.
 *
 * Se connecte en WebSocket au serveur local (127.0.0.1) et exécute les commandes
 * qu'il envoie sur le document ouvert.
 *
 * Protocole (JSON) :
 *   plugin -> serveur : { type: "hello", token, info }
 *   serveur -> plugin : { type: "welcome" }            (sinon la connexion est fermée)
 *   serveur -> plugin : { id, method, params }
 *   plugin -> serveur : { id, result } | { id, error }
 */
(function (window) {
  'use strict';

  var PLUGIN_VERSION = '0.1.0';
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
    } catch (e) { /* stockage indisponible : on continue sans */ }
  }

  // ---------- Infos document ----------

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

  // ---------- Commandes ----------

  function executeMethod(name, args) {
    return new Promise(function (resolve, reject) {
      if (!(window.Asc && window.Asc.plugin && window.Asc.plugin.executeMethod)) {
        reject(new Error('Pas dans un éditeur OnlyOffice'));
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
    }
  };

  function handleCommand(msg) {
    var h = handlers[msg.method];
    var p = h ? h(msg.params || {}) : Promise.reject(new Error('Méthode inconnue : ' + msg.method));
    log('← ' + msg.method);
    p.then(function (result) {
      send({ id: msg.id, result: result });
    }, function (err) {
      log('✗ ' + msg.method + ' : ' + err.message);
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
    if (!port || !token) { setStatus('off', 'Port et jeton requis'); return; }
    wantConnected = true;
    saveSettings();

    if (ws) { try { ws.close(); } catch (e) {} }
    setStatus('wait', 'Connexion…');

    var url = 'ws://127.0.0.1:' + port;
    try {
      ws = new WebSocket(url);
    } catch (e) {
      log('Impossible d\'ouvrir ' + url + ' : ' + e.message);
      setStatus('off', 'Échec de connexion');
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
        setStatus('on', 'Connecté à Claude');
        log('Connecté (' + url + ')');
        return;
      }
      if (!welcomed) return;
      if (msg.id !== undefined && msg.method) handleCommand(msg);
    };

    sock.onerror = function () {
      log('Erreur WebSocket (serveur lancé ? navigateur qui bloque localhost ?)');
    };

    sock.onclose = function (ev) {
      if (sock !== ws) return;
      ws = null;
      if (ev.code === 4001) {
        wantConnected = false;
        setStatus('off', 'Jeton refusé');
        log('Jeton refusé par le serveur');
        return;
      }
      setStatus(wantConnected ? 'wait' : 'off', wantConnected ? 'Reconnexion…' : 'Déconnecté');
      scheduleRetry();
    };
  }

  function disconnect() {
    wantConnected = false;
    clearTimeout(retryTimer);
    if (ws) { var s = ws; ws = null; s.close(); }
    setStatus('off', 'Déconnecté');
  }

  // ---------- Démarrage ----------

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
  // Hors éditeur (test dans un navigateur) : démarrer quand même.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { setTimeout(startOnce, 300); });
  } else {
    setTimeout(startOnce, 300);
  }
})(window);

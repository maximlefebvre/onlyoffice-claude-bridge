/*
 * Claude Bridge — panel (UI only).
 *
 * The connection lives in the background part of the plugin (plugin.js), which
 * keeps running while this panel is hidden or closed. This panel shows the
 * state it pushes ("onState") and sends user actions back ("onPanelAction").
 * The one thing done here is opening the relay window, because browsers only
 * allow pop-ups on a click, and clicks happen in this panel.
 */
(function (window) {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var last = null; // last state received from the background part

  function toBackground(name, data) {
    if (window.Asc && window.Asc.plugin && window.Asc.plugin.sendToPlugin) {
      window.Asc.plugin.sendToPlugin(name, data || {});
    }
  }

  function render(s) {
    last = s || {};
    $('dot').className = 'dot ' + ({ on: 'dot-on', wait: 'dot-wait' }[last.status] || 'dot-off');
    $('status').textContent = last.text || '';
    var h = $('hint');
    h.textContent = last.hint || '';
    h.hidden = !last.hint;
    // Do not overwrite a field the user is editing.
    if (document.activeElement !== $('port') && last.port) $('port').value = last.port;
    if (document.activeElement !== $('token')) $('token').value = last.token || '';
    var el = $('log');
    el.textContent = '';
    (last.log || []).forEach(function (line) {
      var div = document.createElement('div');
      div.textContent = line;
      el.appendChild(div);
    });
  }

  function relayName(key) {
    return 'claude-bridge-relay-' + String(key || 'default').replace(/[^\w-]/g, '_');
  }

  // Must run inside the click handler: browsers only allow pop-ups on a user gesture.
  function openRelay() {
    var key = (last && last.relayKey) || 'default';
    var w = null;
    try { w = window.open('relay.html#' + encodeURIComponent(key), relayName(key), 'popup=yes,width=380,height=240'); } catch (e) {}
    return !!w;
  }

  function onConnect() {
    var action = { type: 'connect', port: $('port').value, token: $('token').value };
    if (last && last.mode === 'relay' && !last.relayOpen) {
      if (!openRelay()) {
        toBackground('onPanelAction', { type: 'log', text: 'The relay window was blocked: allow pop-ups for this site, then click Connect.' });
        return;
      }
      action.relayOpening = true;
    }
    toBackground('onPanelAction', action);
  }

  function onSettingsChange() {
    toBackground('onPanelAction', { type: 'settings', port: $('port').value, token: $('token').value });
  }

  function start() {
    $('connect').onclick = onConnect;
    $('disconnect').onclick = function () { toBackground('onPanelAction', { type: 'disconnect' }); };
    $('port').onchange = onSettingsChange;
    $('token').onchange = onSettingsChange;
    window.Asc.plugin.attachEvent('onState', render);
    toBackground('onPanelReady');
  }

  if (window.Asc && window.Asc.plugin) {
    window.Asc.plugin.init = start;
    window.Asc.plugin.button = function () {};
  }
})(window);

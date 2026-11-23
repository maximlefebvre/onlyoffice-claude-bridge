#!/usr/bin/env node
// MCP server for OnlyOffice Claude Bridge.
//
// Claude talks to this server over MCP. The Claude Bridge plugin, running
// inside an OnlyOffice editor, connects to it over a local WebSocket.
//
// Modes:
//   (default)  stdio. If a shared server is running (see --http), this process
//              only relays the session to it. Otherwise it owns the WebSocket
//              itself, as one server per Claude session.
//   --http     shared server: owns the WebSocket and serves MCP over HTTP on
//              127.0.0.1 to every Claude session (bearer token required).
//              Meant to run once, as a login service (scripts/install-service.sh).
//
// Environment variables:
//   CLAUDE_BRIDGE_PORT             WebSocket port (default 47615)
//   CLAUDE_BRIDGE_MCP_PORT         HTTP MCP port of the shared server (default 47616)
//   CLAUDE_BRIDGE_STANDALONE       1 = never relay to a shared server
//   CLAUDE_BRIDGE_ALLOWED_ORIGINS  comma-separated origins allowed to connect
//                                  (default: https://maximlefebvre.github.io;
//                                  ONLYOFFICE Desktop Editors are always allowed)
//   CLAUDE_BRIDGE_TIMEOUT_MS       editor call timeout (default 20000)
//   CLAUDE_BRIDGE_GRACE_MS         how long a tool call waits for a document
//                                  that just disconnected to come back (default 6000)

import { loadToken } from './bridge.mjs';
import { sharedServerUp, runProxy } from './http.mjs';

const VERSION = '0.6.0';
const port = parseInt(process.env.CLAUDE_BRIDGE_PORT || '47615', 10);
const mcpPort = parseInt(process.env.CLAUDE_BRIDGE_MCP_PORT || '47616', 10);
const httpMode = process.argv.includes('--http');
const allowedOrigins = (process.env.CLAUDE_BRIDGE_ALLOWED_ORIGINS || 'https://maximlefebvre.github.io')
  .split(',').map((s) => s.trim()).filter(Boolean);

const INSTRUCTIONS = [
  'These tools edit the documents the user has open in OnlyOffice (for example in kDrive), live, through the Claude Bridge plugin.',
  'Start with list_documents to see what is connected. Tools prefixed word_ work on text documents, cell_ on spreadsheets, slide_ on presentations.',
  'Read before editing (word_read_document, cell_read_range, slide_list) and use the indexes they return.',
  'For anything the dedicated tools do not cover, look up the API with api_reference, then use run_script.',
  'Edits are applied immediately in the user\'s document and can be undone with Ctrl/Cmd+Z in the editor. For substantial rewrites of a text document, consider word_set_track_changes so the user can review.',
  'The plugin keeps each open document connected in the background. If no document is connected, ask the user to check that Claude Bridge is enabled under Plugins → Background plugins, then open its panel (Claude tab → Claude Bridge) and click Connect.',
  'If a tool reports that the port is held by another Claude Bridge server, another Claude session owns the connection: tell the user to close that session (this one takes over automatically), or to install the shared server (scripts/install-service.sh) so that every session shares it.',
  'The first time a browser connects, the user approves it in a dialog on their computer. If list_connection_requests shows a pending request (no dialog available), ask the user and only call approve_connection with their explicit answer.',
].join(' ');

const editorOf = { word_: 'word', cell_: 'cell', slide_: 'slide' };

// The MCP server and the tools (with the large API reference) are only loaded
// when this process serves them itself, not when it just relays to a shared server.
async function loadServerModules() {
  const [{ McpServer }, { StdioServerTransport }, { Bridge }, common, reference, word, cell, slide] = await Promise.all([
    import('@modelcontextprotocol/sdk/server/mcp.js'),
    import('@modelcontextprotocol/sdk/server/stdio.js'),
    import('./bridge.mjs'),
    import('./tools/common.mjs'),
    import('./tools/reference.mjs'),
    import('./tools/word.mjs'),
    import('./tools/cell.mjs'),
    import('./tools/slide.mjs'),
  ]);
  const tools = [...common.commonTools, ...reference.referenceTools, ...word.wordTools, ...cell.cellTools, ...slide.slideTools];
  return { McpServer, StdioServerTransport, Bridge, tools };
}

// A new MCP server with every tool, driving the given bridge.
function makeServer(mods, bridge) {
  const server = new mods.McpServer({ name: 'onlyoffice-claude-bridge', version: VERSION }, { instructions: INSTRUCTIONS });
  for (const tool of mods.tools) {
    const editor = Object.entries(editorOf).find(([p]) => tool.name.startsWith(p))?.[1];
    server.registerTool(
      tool.name,
      {
        title: tool.title,
        description: tool.description,
        inputSchema: tool.schema,
        annotations: {
          title: tool.title,
          readOnlyHint: !!tool.readOnly,
          destructiveHint: !!tool.destructive,
          openWorldHint: false,
        },
      },
      async (args = {}) => {
        try {
          const client = tool.noDocument ? null : await bridge.resolve(args.document, editor);
          const result = await tool.run(bridge, client, args);
          const payload = client
            ? { document: { id: client.id, title: client.title, editor: client.editor }, result }
            : result;
          return { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
        } catch (err) {
          return { isError: true, content: [{ type: 'text', text: `Error: ${err.message}` }] };
        }
      },
    );
  }
  return server;
}

function startBridge(mods) {
  const { token, file: tokenFile } = loadToken();
  const bridge = new mods.Bridge({
    port,
    allowedOrigins,
    token,
    timeoutMs: parseInt(process.env.CLAUDE_BRIDGE_TIMEOUT_MS || '20000', 10),
    graceMs: parseInt(process.env.CLAUDE_BRIDGE_GRACE_MS || '6000', 10),
    version: VERSION,
  });
  bridge.start();
  console.error(`[claude-bridge] token file: ${tokenFile}`);
  const shutdown = () => { try { bridge.wss?.close(); } catch {} process.exit(0); };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
  return { bridge, shutdown };
}

if (httpMode) {
  const mods = await loadServerModules();
  const { startHttpServer } = await import('./http.mjs');
  const { bridge } = startBridge(mods);
  const { token } = loadToken('mcp-token');
  await startHttpServer({
    port: mcpPort,
    token,
    version: VERSION,
    makeServer: () => makeServer(mods, bridge),
    getStatus: () => bridge.status(),
  });
  console.error(`[claude-bridge] shared server ${VERSION} ready`);
} else if (process.env.CLAUDE_BRIDGE_STANDALONE !== '1' && await sharedServerUp(mcpPort)) {
  await runProxy({ port: mcpPort, token: loadToken('mcp-token').token });
} else {
  const mods = await loadServerModules();
  const { bridge, shutdown } = startBridge(mods);
  await makeServer(mods, bridge).connect(new mods.StdioServerTransport());
  console.error(`[claude-bridge] MCP server ${VERSION} ready`);
  process.stdin.on('close', shutdown);
}

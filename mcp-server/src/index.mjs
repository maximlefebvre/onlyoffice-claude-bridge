#!/usr/bin/env node
// MCP server for OnlyOffice Claude Bridge.
//
// Claude talks to this server over stdio (MCP). The Claude Bridge plugin,
// running inside an OnlyOffice editor, connects to it over a local WebSocket.
//
// Environment variables:
//   CLAUDE_BRIDGE_PORT             WebSocket port (default 47615)
//   CLAUDE_BRIDGE_ALLOWED_ORIGINS  comma-separated origins allowed to connect
//                                  (default: https://maximlefebvre.github.io)
//   CLAUDE_BRIDGE_TIMEOUT_MS       editor call timeout (default 20000)

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Bridge, loadToken } from './bridge.mjs';
import { commonTools } from './tools/common.mjs';
import { wordTools } from './tools/word.mjs';
import { cellTools } from './tools/cell.mjs';
import { slideTools } from './tools/slide.mjs';

const VERSION = '0.3.0';
const port = parseInt(process.env.CLAUDE_BRIDGE_PORT || '47615', 10);
const allowedOrigins = (process.env.CLAUDE_BRIDGE_ALLOWED_ORIGINS || 'https://maximlefebvre.github.io')
  .split(',').map((s) => s.trim()).filter(Boolean);
const { token, file: tokenFile } = loadToken();

const bridge = new Bridge({
  port,
  allowedOrigins,
  token,
  timeoutMs: parseInt(process.env.CLAUDE_BRIDGE_TIMEOUT_MS || '20000', 10),
});
bridge.start();
console.error(`[claude-bridge] token file: ${tokenFile}`);

const server = new McpServer(
  { name: 'onlyoffice-claude-bridge', version: VERSION },
  {
    instructions: [
      'These tools edit the documents the user has open in OnlyOffice (for example in kDrive), live, through the Claude Bridge plugin.',
      'Start with list_documents to see what is connected. Tools prefixed word_ work on text documents, cell_ on spreadsheets, slide_ on presentations.',
      'Read before editing (word_read_document, cell_read_range, slide_list) and use the indexes they return.',
      'Edits are applied immediately in the user\'s document and can be undone with Ctrl/Cmd+Z in the editor. For substantial rewrites of a text document, consider word_set_track_changes so the user can review.',
      'If no document is connected, ask the user to open the Claude Bridge panel in the editor and click Connect.',
      'The first time a browser connects, the user approves it in a dialog on their computer. If list_connection_requests shows a pending request (no dialog available), ask the user and only call approve_connection with their explicit answer.',
    ].join(' '),
  },
);

const editorOf = { word_: 'word', cell_: 'cell', slide_: 'slide' };

for (const tool of [...commonTools, ...wordTools, ...cellTools, ...slideTools]) {
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
        const client = tool.noDocument ? null : bridge.resolve(args.document, editor);
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

await server.connect(new StdioServerTransport());
console.error(`[claude-bridge] MCP server ${VERSION} ready`);

const shutdown = () => { try { bridge.wss?.close(); } catch {} process.exit(0); };
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
process.stdin.on('close', shutdown);

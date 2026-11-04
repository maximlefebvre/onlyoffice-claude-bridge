# OnlyOffice Claude Bridge

Let Claude (the Claude app or Claude Code) edit the document open in OnlyOffice, through an OnlyOffice plugin and a local MCP server. Designed for any OnlyOffice instance that accepts a custom plugin store (kDrive, Nextcloud, OnlyOffice Docs, Desktop Editors…).

> Unofficial personal project, not affiliated with Anthropic or Ascensio System SIA (ONLYOFFICE).

```
Claude ──MCP──▶ local server (127.0.0.1) ──WebSocket──▶ OnlyOffice plugin
```

## Contents

- `docs/`: static site published on GitHub Pages.
  - `docs/store/`: the OnlyOffice plugin store app (AGPL-3.0, copied from ONLYOFFICE/onlyoffice.github.io). Plugin list: `docs/store/config.json`.
  - `docs/sdkjs-plugins/content/claude-bridge/`: the plugin. It is deliberately generic: it runs code sent by the local server, so new actions never need a plugin update.
- `mcp-server/`: the MCP server (Node.js ≥ 20). It exposes the tools to Claude and the WebSocket the plugin connects to.

## Setup

### 1. Install the MCP server

```bash
cd mcp-server && npm install
```

Claude Code:

```bash
claude mcp add onlyoffice-claude-bridge -s user -- node /absolute/path/to/mcp-server/src/index.mjs
```

Claude desktop app (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "onlyoffice-claude-bridge": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-server/src/index.mjs"]
    }
  }
}
```

The server listens on `ws://127.0.0.1:47615`. The token the plugin needs is in `~/.config/onlyoffice-claude-bridge/token` (created on first start).

Options (environment variables): `CLAUDE_BRIDGE_PORT`, `CLAUDE_BRIDGE_ALLOWED_ORIGINS` (default `https://maximlefebvre.github.io`), `CLAUDE_BRIDGE_TIMEOUT_MS`.

### 2. Install the plugin

In the OnlyOffice editor: Plugins → Plugin Manager → `</>` (developer mode) → paste
`https://maximlefebvre.github.io/onlyoffice-claude-bridge/store/index.html`, then install **Claude Bridge**. Open its panel, paste the token and click **Connect**.

### 3. Firefox only: allow the local connection

Since Firefox 151, a plugin frame cannot reach `127.0.0.1` by default and no permission prompt is shown. In `about:config`, add `127.0.0.1` to `network.lna.skip-domains`. Note that this lets any site reach services on `127.0.0.1`.

## Tools

| Scope | Tools |
| --- | --- |
| All | `list_documents`, `get_selection`, `replace_selection`, `run_script`, `execute_method` |
| Text documents | `word_read_document`, `word_insert_paragraphs`, `word_replace_paragraph`, `word_delete_elements`, `word_search_replace`, `word_format_text`, `word_set_style`, `word_add_comment`, `word_list_comments`, `word_insert_table`, `word_set_track_changes` |
| Spreadsheets | `cell_list_sheets`, `cell_read_range`, `cell_write_range`, `cell_get_selection`, `cell_format_range`, `cell_clear_range`, `cell_insert_cells`, `cell_delete_cells`, `cell_manage_sheet`, `cell_add_comment` |
| Presentations | `slide_list`, `slide_add`, `slide_manage`, `slide_set_text`, `slide_add_text_box`, `slide_edit_shape`, `slide_set_notes` |

`run_script` runs any [OnlyOffice Office API](https://api.onlyoffice.com/docs/office-api/get-started/overview/) code in the editor, for everything the other tools do not cover.

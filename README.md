# OnlyOffice Claude Bridge

Let Claude edit the documents, spreadsheets and presentations open in OnlyOffice, through an OnlyOffice editor plugin and a local MCP server packaged as a Claude Code plugin. Works with a Claude subscription, no API key. Designed for any OnlyOffice instance that accepts a custom plugin store (kDrive, Nextcloud, OnlyOffice Docs, Desktop Editors…).

> Unofficial personal project, not affiliated with Anthropic or Ascensio System SIA (ONLYOFFICE).

```
Claude ──MCP──▶ local server (127.0.0.1) ──WebSocket──▶ OnlyOffice plugin
```

Step-by-step guide with screenshots: **https://maximlefebvre.github.io/onlyoffice-claude-bridge/install.html**

## Contents

- `.claude-plugin/`, `skills/`: the Claude Code plugin (`onlyoffice-bridge`) and its marketplace. The plugin starts the MCP server and adds an `onlyoffice` skill with the workflow, units and known pitfalls.
- `mcp-server/`: the MCP server (Node.js ≥ 20). It exposes the tools to Claude and the WebSocket the plugin connects to. `dist/server.mjs` is a self-contained bundle (no `npm install` needed), rebuilt with `npm run build`.
- `docs/`: static site published on GitHub Pages: landing page, install guide, browser permission page (`connect.html`).
  - `docs/store/`: the OnlyOffice plugin store app (AGPL-3.0, copied from ONLYOFFICE/onlyoffice.github.io). Plugin list: `docs/store/config.json`.
  - `docs/sdkjs-plugins/content/claude-bridge/`: the editor plugin. It is deliberately generic: it runs code sent by the local server, so new actions never need a plugin update.

## Setup

### 1. Install the Claude Code plugin

```bash
claude plugin marketplace add maximlefebvre/onlyoffice-claude-bridge
claude plugin install onlyoffice-bridge@onlyoffice-bridge
```

Each new Claude Code session then starts the server, which listens on `ws://127.0.0.1:47615`. Only one session at a time can use the port.

#### Other MCP clients

Claude desktop app (Chat tab), in `claude_desktop_config.json`, or any MCP client:

```json
{
  "mcpServers": {
    "onlyoffice": {
      "command": "node",
      "args": ["/absolute/path/to/onlyoffice-claude-bridge/mcp-server/dist/server.mjs"]
    }
  }
}
```

Options (environment variables): `CLAUDE_BRIDGE_PORT`, `CLAUDE_BRIDGE_ALLOWED_ORIGINS` (default `https://maximlefebvre.github.io`), `CLAUDE_BRIDGE_TIMEOUT_MS`, `CLAUDE_BRIDGE_NO_DIALOG` (never show the approval dialog).

### 2. Install the editor plugin

In the OnlyOffice editor: Plugins → Plugin Manager → `</>` (developer mode) → paste
`https://maximlefebvre.github.io/onlyoffice-claude-bridge/store/index.html`, then install **Claude Bridge** and open its panel. It connects by itself.

### 3. Allow the browser

Recent browsers stop websites from reaching `127.0.0.1` unless the user allows it, and Firefox (151+) cannot ask from inside the editor's plugin frame. Open [connect.html](https://maximlefebvre.github.io/onlyoffice-claude-bridge/connect.html) as a normal tab and click **Test connection**: the browser asks once, and the permission is stored for `maximlefebvre.github.io` only (Firefox checks the stored permission of the frame's own origin before anything else). Earlier versions of this guide suggested adding `127.0.0.1` to `network.lna.skip-domains`; that exposes local services to every site and can be removed.

### 4. Approve the first connection

The first time a browser connects, the server asks you to approve it in a dialog on your computer (macOS: AppleScript, Windows: PowerShell, Linux: `zenity` or `kdialog`). Where no dialog can be shown, Claude asks you in the conversation (`list_connection_requests` / `approve_connection`). The plugin then receives its own token and reconnects silently afterwards. Approved browsers are listed in `~/.config/onlyoffice-claude-bridge/devices.json` (hashed tokens) and can be revoked with the `manage_paired_browsers` tool. A manual token (`~/.config/onlyoffice-claude-bridge/token`) can still be pasted under **Advanced** in the panel.

## Development

```bash
cd mcp-server && npm install
npm start          # run from the sources
npm run build      # rebuild dist/server.mjs (commit it: the plugin runs the bundle)
```

To try the plugin from a local checkout: `claude --plugin-dir /path/to/onlyoffice-claude-bridge`.

## Tools

| Scope | Tools |
| --- | --- |
| All | `list_documents`, `get_selection`, `replace_selection`, `list_fonts`, `api_reference`, `run_script`, `execute_method` |
| Connections | `list_connection_requests`, `approve_connection`, `manage_paired_browsers` |
| Text documents | `word_read_document`, `word_insert_paragraphs`, `word_replace_paragraph`, `word_delete_elements`, `word_search_replace`, `word_format_text`, `word_format_paragraphs`, `word_list_styles`, `word_modify_style`, `word_set_style`, `word_add_comment`, `word_list_comments`, `word_insert_table`, `word_insert_image`, `word_set_list`, `word_set_header_footer`, `word_add_footnote`, `word_add_hyperlink`, `word_page_setup`, `word_insert_page_break`, `word_set_track_changes` |
| Spreadsheets | `cell_list_sheets`, `cell_read_range`, `cell_write_range`, `cell_get_selection`, `cell_format_range`, `cell_clear_range`, `cell_insert_cells`, `cell_delete_cells`, `cell_manage_sheet`, `cell_add_comment`, `cell_add_chart`, `cell_add_image` |
| Presentations | `slide_list`, `slide_add`, `slide_manage`, `slide_set_text`, `slide_add_text_box`, `slide_format_text`, `slide_edit_shape`, `slide_set_notes`, `slide_add_image`, `slide_set_background`, `slide_set_bullets` |

Formatting covers fonts, sizes, colors, highlight, caps, alignment, line spacing, space before/after, indents and paragraph backgrounds in text documents (per selection, paragraph, search match, whole document, or by changing a style); fonts, colors, borders, alignment, number formats and sizes in spreadsheets; fonts, colors, alignment and line spacing in presentations.

Images can come from a URL, a file on your computer, or a data URL; the server embeds them, so the document server never has to fetch anything.

### API reference

The OnlyOffice Office API has about 2,600 methods; only the common tasks have dedicated tools. For the rest, Claude searches `api_reference` (generated from the OnlyOffice sources, see `mcp-server/scripts/build-api-reference.mjs`) and writes a `run_script`. To match another OnlyOffice version, regenerate it:

```bash
cd mcp-server && node scripts/build-api-reference.mjs v10.0.0.176
```

`run_script` runs any [OnlyOffice Office API](https://api.onlyoffice.com/docs/office-api/get-started/overview/) code in the editor, for everything the other tools do not cover.

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

Each new Claude Code session then starts the server, which listens on `ws://127.0.0.1:47615`. Only one session at a time can use the port: the others keep retrying, and their tools report which process holds it (another Claude Bridge server, with its pid). When the session holding the port ends, the next one takes it over within a few seconds and the plugin reconnects to it.

#### Shared server (recommended with several Claude sessions)

To let every Claude session use the connected documents at the same time, run one shared server (`dist/server.mjs --http`). On macOS, two ways:

- **Menu bar app**, started and stopped by hand. It bundles the server, shows the connected documents, the Claude sessions and whether the plugin port is free, and opens the log. Launching the app starts the server, quitting it stops it; **Open at Login** is optional. Build it with the Xcode command line tools (no Xcode project):

  ```bash
  macos-app/build.sh install    # builds and copies "Claude Bridge.app" to ~/Applications
  ```

  The app carries its own copy of `dist/server.mjs`: rebuild it after updating the server.

- **Login service** (launchd agent), always running:

  ```bash
  mcp-server/scripts/install-service.sh            # install and start (also: status, restart, uninstall)
  ```

The shared server:

- owns the WebSocket on port 47615, so documents stay connected whichever Claude sessions are open;
- serves MCP over Streamable HTTP on `http://127.0.0.1:47616/mcp` (`CLAUDE_BRIDGE_MCP_PORT`). It only answers to `127.0.0.1` / `localhost` host names, refuses browser requests (any `Origin` header), and requires the bearer token stored in `~/.config/onlyoffice-claude-bridge/mcp-token`.

Nothing else to configure: a session started with the usual stdio command finds the shared server (`/health`) and relays its MCP messages to it instead of opening the port itself. Sessions already open keep their own server until they restart. Set `CLAUDE_BRIDGE_STANDALONE=1` to opt a session out. HTTP-capable clients can also connect directly:

```bash
claude mcp add --scope user --transport http onlyoffice http://127.0.0.1:47616/mcp \
  --header "Authorization: Bearer $(cat ~/.config/onlyoffice-claude-bridge/mcp-token)"
```

The service points at the `dist/server.mjs` of the checkout the script was run from: run the script again after moving the checkout. Log: `~/Library/Logs/onlyoffice-claude-bridge.log`.

When a document disconnects (plugin reload, network hiccup), tool calls wait a few seconds (`CLAUDE_BRIDGE_GRACE_MS`, default 6000) for it to come back before reporting that nothing is connected, and a reconnecting document keeps its id.

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

Options (environment variables): `CLAUDE_BRIDGE_PORT`, `CLAUDE_BRIDGE_MCP_PORT` (shared server, default 47616), `CLAUDE_BRIDGE_STANDALONE`, `CLAUDE_BRIDGE_ALLOWED_ORIGINS` (default `https://maximlefebvre.github.io`; ONLYOFFICE Desktop Editors, which load installed plugins from local files, are always accepted), `CLAUDE_BRIDGE_TIMEOUT_MS`, `CLAUDE_BRIDGE_GRACE_MS`, `CLAUDE_BRIDGE_NO_DIALOG` (never show the approval dialog).

### 2. Install the editor plugin

In the OnlyOffice editor: Plugins → Plugin Manager → `</>` (developer mode) → paste
`https://maximlefebvre.github.io/onlyoffice-claude-bridge/store/index.html`, then install **Claude Bridge**. It is a background plugin: turn it on under **Plugins → Background plugins** if needed (remembered across sessions). It connects by itself and keeps each open document connected while the document is open, whether or not its panel is shown. The panel (left bar, or **Claude** tab → **Claude Bridge**) shows the status, a Connect button and the log.

Requires OnlyOffice 8.1+ (background plugins since 8.0, plugin panels and toolbar buttons since 8.1). Tested target: Desktop Editors 10.0 and kDrive (OnlyOffice 10.0).

How it works: `index.html` / `plugin.js` is the background variation (`"type": "background"` in `config.json`). It owns the WebSocket, runs the commands and reconnects automatically (backoff up to 15 s). It opens `panel.html` / `panel.js` as a plugin panel (`Asc.PluginWindow`), which is UI only: state goes to it with `command("onState")` and user actions come back with `sendToPlugin("onPanelAction")`. Hiding or closing the panel does not touch the connection.

### 3. Browsers with Local Network Access (Firefox 151+, recent Chrome)

These browsers block the plugin frame from reaching `127.0.0.1`: prompts are not allowed in cross-origin frames, and the WebSocket's triggering principal is partitioned under the editor's top-level site, so a permission granted to `maximlefebvre.github.io` in a normal tab does not apply there. The panel detects it (Permissions API) and switches to a **relay window**: on **Connect**, it opens `relay.html` (same origin, top-level), which can ask for the permission, holds the WebSocket and relays frames with `postMessage` to the plugin's background frame. The panel only opens it (a pop-up needs a click); the relay then finds the background frame by itself by walking the editor tab's frames, so it keeps working while the panel is hidden, and finds the frame again after a plugin reload. One relay window per document tab; it closes when the editor tab is closed or the document has been gone for a minute. The `network.lna.skip-domains` workaround is no longer needed (it exposed local services to every site).

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

# OnlyOffice Claude Bridge

Let Claude (the Claude app or Claude Code) edit the document open in OnlyOffice, through an OnlyOffice plugin and a local MCP server. Designed for any OnlyOffice instance that accepts a custom plugin store (kDrive, Nextcloud, OnlyOffice Docs, Desktop Editors…).

> Unofficial personal project, not affiliated with Anthropic or Ascensio System SIA (ONLYOFFICE).

```
Claude ──MCP──▶ local server (127.0.0.1) ──WebSocket──▶ OnlyOffice plugin
```

## Contents

- `docs/`: static site published on GitHub Pages.
  - `docs/store/`: the OnlyOffice plugin store app (AGPL-3.0, copied from ONLYOFFICE/onlyoffice.github.io). Plugin list: `docs/store/config.json`.
  - `docs/sdkjs-plugins/content/claude-bridge/`: the plugin.
- `test-server/`: WebSocket test server (step 1, no MCP yet).

## Step 1: test in an OnlyOffice editor (e.g. kDrive)

1. `cd test-server && npm install && node server.mjs`. The server prints the token (stored in `~/.config/onlyoffice-claude-bridge/token`).
2. Open the editor → Plugins → Plugin Manager → `</>` (developer mode) → paste `https://maximlefebvre.github.io/onlyoffice-claude-bridge/store/index.html`.
3. Install "Claude Bridge", open its panel, paste the token, click "Connect".
4. In the server terminal, type `ping`, `info` or `sel` (after selecting some text).

If the server prints "origin not allowed: X", restart it with `ALLOWED_ORIGINS=X node server.mjs`.

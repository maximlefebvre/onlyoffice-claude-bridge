# kDrive × Claude

Faire éditer les documents OnlyOffice ouverts dans kDrive par Claude (app Claude ou Claude Code), via un plugin OnlyOffice et un serveur MCP local.

```
Claude ──MCP──▶ serveur local (127.0.0.1) ──WebSocket──▶ plugin OnlyOffice (dans kDrive)
```

## Contenu

- `docs/` : site statique à publier sur GitHub Pages.
  - `docs/store/` : application « store » d'OnlyOffice (AGPL-3.0, copiée de ONLYOFFICE/onlyoffice.github.io). Liste des plugins : `docs/store/config.json`.
  - `docs/sdkjs-plugins/content/claude-bridge/` : le plugin.
- `test-server/` : serveur WebSocket de test (étape 1, sans MCP).

## Étape 1 : tester dans kDrive

1. `cd test-server && npm install && node server.mjs`. Le serveur affiche le jeton (enregistré dans `~/.config/kdrive-claude/token`).
2. Publier `docs/` sur GitHub Pages, puis ouvrir l'éditeur kDrive → Plugins → Gestionnaire de plugins → `</>` (mode développeur) → coller `https://<compte>.github.io/<repo>/store/index.html`.
3. Installer « Claude Bridge », ouvrir le panneau, coller le jeton, cliquer sur « Connecter ».
4. Dans le terminal du serveur, taper `ping`, `info` ou `sel` (après avoir sélectionné du texte).

Si le serveur affiche « origine non autorisée : X », relancer avec `ALLOWED_ORIGINS=X node server.mjs`.

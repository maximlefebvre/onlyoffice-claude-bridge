# OnlyOffice Claude Bridge

Permettre à Claude (app Claude ou Claude Code) d'éditer le document OnlyOffice ouvert, via un plugin OnlyOffice et un serveur MCP local. Conçu pour toute instance OnlyOffice qui accepte un store de plugins personnalisé (kDrive, Nextcloud, OnlyOffice Docs, Desktop Editors…).

> Projet personnel non officiel, sans lien avec Anthropic ni avec Ascensio System SIA (ONLYOFFICE).

```
Claude ──MCP──▶ serveur local (127.0.0.1) ──WebSocket──▶ plugin OnlyOffice
```

## Contenu

- `docs/` : site statique à publier sur GitHub Pages.
  - `docs/store/` : application « store » d'OnlyOffice (AGPL-3.0, copiée de ONLYOFFICE/onlyoffice.github.io). Liste des plugins : `docs/store/config.json`.
  - `docs/sdkjs-plugins/content/claude-bridge/` : le plugin.
- `test-server/` : serveur WebSocket de test (étape 1, sans MCP).

## Étape 1 : tester dans un éditeur OnlyOffice (ex. kDrive)

1. `cd test-server && npm install && node server.mjs`. Le serveur affiche le jeton (enregistré dans `~/.config/onlyoffice-claude-bridge/token`).
2. Ouvrir l'éditeur → Plugins → Gestionnaire de plugins → `</>` (mode développeur) → coller `https://maximlefebvre.github.io/onlyoffice-claude-bridge/store/index.html`.
3. Installer « Claude Bridge », ouvrir le panneau, coller le jeton, cliquer sur « Connecter ».
4. Dans le terminal du serveur, taper `ping`, `info` ou `sel` (après avoir sélectionné du texte).

Si le serveur affiche « origine non autorisée : X », relancer avec `ALLOWED_ORIGINS=X node server.mjs`.

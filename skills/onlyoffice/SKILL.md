---
name: onlyoffice
description: Edit the documents, spreadsheets and presentations the user has open in OnlyOffice (kDrive, Nextcloud, ONLYOFFICE DocSpace or any OnlyOffice editor) through the Claude Bridge plugin. Use when the user asks to write, rewrite, format, translate, fill in, chart or restructure "my document", "the sheet", "the slides" or a file they have open in OnlyOffice or kDrive, or when the onlyoffice MCP tools (list_documents, word_*, cell_*, slide_*) are available and the request is about an open office file.
---

# Editing open OnlyOffice files

The `onlyoffice` MCP server talks live to the **Claude Bridge** panel running inside the user's OnlyOffice editor. Every edit appears immediately in their document, and they can undo it there with Cmd/Ctrl+Z. Nothing is uploaded anywhere: the server only listens on 127.0.0.1.

## Workflow

1. **Find the document.** Call `list_documents`. Each entry has an `id`, a `title` and an `editor`: `word` (text), `cell` (spreadsheet) or `slide` (presentation). When several are open, pass `document` (an id or part of the title) to every call. Without it, tools use the most recently used document of the right type.
2. **Read before you write.** Use `word_read_document`, `cell_read_range` / `cell_list_sheets` or `slide_list`. Element, sheet, slide and shape indexes come from these results, and they shift after inserts and deletes, so read again after structural changes.
3. **Use the dedicated tools first.** `word_*`, `cell_*` and `slide_*` cover text, formatting, styles, tables, lists, images, comments, headers/footers, footnotes, charts, borders, number formats, backgrounds, notes and bullets.
4. **Fall back to the API.** For anything else, search `api_reference` (for example `query: "table of contents"`, `class: "ApiParagraph"` or `method: "ApiRange.SetBorders"`), then run the code with `run_script`.
5. **Work one call at a time on a document.** Calls are queued per document, but sending them sequentially keeps indexes predictable.

For large rewrites of a text document, offer `word_set_track_changes` first so the user can review and accept the changes in OnlyOffice.

## When nothing is connected

- **No document listed:** ask the user to open the document, then **Plugins → Claude Bridge**. The panel connects on its own and shows a green dot once connected.
- **Panel stuck on "Waiting for approval":** the first connection from a browser must be approved. On macOS, Windows and Linux desktops a dialog appears. Without a dialog, call `list_connection_requests`, ask the user whether to allow it, and call `approve_connection` **only with their explicit answer**.
- **Panel says the server is unreachable in Firefox:** see the install guide at https://maximlefebvre.github.io/onlyoffice-claude-bridge/install.html.
- **Error "port in use":** another Claude session already runs the bridge. Only one session at a time can serve it.

## run_script essentials

`code` is a function body that runs inside the editor. It can only use the `Api` global and `Asc.scope`, which holds the values you pass in `scope`. Return JSON-serializable data and pass values in through `scope` rather than building code strings.

```js
// word
var doc = Api.GetDocument();
var p = Api.CreateParagraph(); p.AddText(Asc.scope.text); doc.Push(p);
return doc.GetElementsCount();
```

```js
// cell
var ws = Api.GetActiveSheet();            // or Api.GetSheet("Name")
ws.GetRange("A1:B2").SetValue([[1, 2], [3, 4]]);
return ws.GetRange("A1:B2").GetValue();
```

```js
// slide
var pres = Api.GetPresentation(), slide = pres.GetSlideByIndex(0);
return slide.GetAllShapes().map(function (s) { return s.GetDocContent() ? s.GetDocContent().GetText() : null; });
```

### Units

| Value | Unit | Conversion |
|---|---|---|
| Font size in text and slides (`SetFontSize`) | half-points | 12 pt → `24` |
| Font size in spreadsheets (`ApiRange.SetFontSize`) | points | 12 pt → `12` |
| Indents, spacing, page size and margins in text documents | twips | 1 pt = 20, 1 cm = 567 |
| Line spacing (`SetSpacingLine(v, "auto")`) | 240ths of a line | 1.5 lines → `360` |
| Shape and image size and position | EMU | 1 cm = 360000, 1 pt = 12700 |
| Colors | — | `Api.CreateRGBColor(r, g, b)`, or `SetColor(r, g, b)` on text |

### Known pitfalls (OnlyOffice 10)

- `Api.CreateTable(rows, cols)`: older editors expect `(cols, rows)`. Check `GetRowsCount()`.
- `RemoveAllElements()` on a shape, notes, header or footer content leaves one empty paragraph. Remove it after pushing new paragraphs.
- Text read from shapes and table cells ends with `\r\n`, so trim it.
- `Api.GetSelection()` and the cursor are the user's. Avoid moving them unless asked.
- `run_script` cannot reach the network, the DOM or `window`. Load images through `word_insert_image`, `cell_add_image` or `slide_add_image`, which accept URLs, local paths and data URLs.

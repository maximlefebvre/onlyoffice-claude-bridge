// Tools for spreadsheets (OnlyOffice "cell" editor).
// Functions passed to bridge.run() execute inside the editor (see word.mjs).

import { z } from 'zod';
import { documentArg, hexToRgb, parseCell, cellRef } from './shared.mjs';

const sheetArg = {
  sheet: z.string().optional().describe('Sheet name (default: the active sheet).'),
};

const MAX_CELLS = 10000;

export const cellTools = [
  {
    name: 'cell_list_sheets',
    title: 'List sheets',
    description: 'List the sheets of the spreadsheet with their used range, and which one is active.',
    schema: { ...documentArg },
    readOnly: true,
    run: (b, c) => b.run(c, function () {
      var active = Api.GetActiveSheet().GetName();
      return Api.GetSheets().map(function (ws) {
        var used = null;
        try { used = ws.GetUsedRange().GetAddress(false, false); } catch (e) {}
        return { name: ws.GetName(), index: ws.GetIndex(), active: ws.GetName() === active, usedRange: used };
      });
    }, {}, false),
  },

  {
    name: 'cell_read_range',
    title: 'Read cells',
    description: 'Read values (and formulas) of a range such as "A1:D20". Without a range, reads the used range of the sheet.',
    schema: {
      ...documentArg,
      ...sheetArg,
      range: z.string().optional().describe('A1-style range, e.g. "A1:D20" or "B:B".'),
    },
    readOnly: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      var rng = s.range ? ws.GetRange(s.range) : ws.GetUsedRange();
      if (!rng) throw new Error('Invalid range: ' + s.range);
      var count = rng.GetCount();
      if (count > s.max) throw new Error('Range has ' + count + ' cells; read at most ' + s.max + ' at a time');
      var out = { sheet: ws.GetName(), address: rng.GetAddress(false, false) };
      var values = rng.GetValue(), formulas = rng.GetFormula();
      if (count > 1 && !Array.isArray(values)) {
        // Older editors: fall back to a cell-by-cell listing.
        var cells = [];
        rng.ForEach(function (cell) {
          var f = cell.GetFormula(), v = cell.GetValue();
          if (v === '' && !f) return;
          cells.push({ cell: cell.GetAddress(false, false), value: v, formula: f && String(f).charAt(0) === '=' ? f : undefined });
        });
        out.cells = cells;
        return out;
      }
      if (!Array.isArray(values)) { values = [[values]]; formulas = [[formulas]]; }
      out.values = values;
      var hasFormula = false;
      var f2 = (formulas || []).map(function (row) {
        return (Array.isArray(row) ? row : [row]).map(function (f) {
          if (f && String(f).charAt(0) === '=') { hasFormula = true; return f; }
          return null;
        });
      });
      if (hasFormula) out.formulas = f2;
      return out;
    }, { ...a, max: MAX_CELLS }, false),
  },

  {
    name: 'cell_write_range',
    title: 'Write cells',
    description:
      'Write a block of values starting at a cell. `values` is a list of rows. Strings starting with "=" are formulas (e.g. "=SUM(B2:B10)"). null leaves a cell unchanged, "" clears it.',
    schema: {
      ...documentArg,
      ...sheetArg,
      start: z.string().describe('Top-left cell, e.g. "A1".'),
      values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).min(1),
    },
    run: (b, c, a) => {
      const { row, col } = parseCell(a.start);
      const cells = [];
      a.values.forEach((r, i) => r.forEach((v, j) => {
        if (v !== null && v !== undefined) cells.push([cellRef(row + i, col + j), v]);
      }));
      if (cells.length > MAX_CELLS) throw new Error(`Write at most ${MAX_CELLS} cells at a time`);
      return b.run(c, function () {
        var s = Asc.scope;
        var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
        if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
        s.cells.forEach(function (cv) { ws.GetRange(cv[0]).SetValue(cv[1]); });
        return { written: s.cells.length, sheet: ws.GetName() };
      }, { sheet: a.sheet, cells });
    },
  },

  {
    name: 'cell_get_selection',
    title: 'Get selected cells',
    description: 'Return the address and values of the cells currently selected by the user.',
    schema: { ...documentArg },
    readOnly: true,
    run: (b, c) => b.run(c, function () {
      var sel = Api.GetSelection();
      if (!sel) return null;
      var v = sel.GetValue();
      return { sheet: Api.GetActiveSheet().GetName(), address: sel.GetAddress(false, false), values: Array.isArray(v) ? v : [[v]] };
    }, {}, false),
  },

  {
    name: 'cell_format_range',
    title: 'Format cells',
    description: 'Format a range: font (bold, italic, underline, size, color), fill color, number format (e.g. "0.00", "#,##0 €", "dd/mm/yyyy", "0%"), alignment, wrap, merge, column width, autofit.',
    schema: {
      ...documentArg,
      ...sheetArg,
      range: z.string(),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      font_size: z.number().positive().optional().describe('In points.'),
      font_color: z.string().optional().describe('#RRGGBB'),
      fill_color: z.string().optional().describe('#RRGGBB, or "none" to clear.'),
      number_format: z.string().optional(),
      align: z.enum(['left', 'center', 'right', 'justify']).optional(),
      wrap: z.boolean().optional(),
      merge: z.boolean().optional(),
      column_width: z.number().positive().optional().describe('Width in characters for the columns of the range.'),
      autofit: z.boolean().optional().describe('Auto-fit the column widths of the range.'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      var r = ws.GetRange(s.range);
      if (!r) throw new Error('Invalid range: ' + s.range);
      if (s.bold !== undefined) r.SetBold(s.bold);
      if (s.italic !== undefined) r.SetItalic(s.italic);
      if (s.underline !== undefined) r.SetUnderline(s.underline ? 'single' : 'none');
      if (s.font_size) r.SetFontSize(s.font_size);
      if (s.fontRgb) r.SetFontColor(Api.CreateColorFromRGB(s.fontRgb[0], s.fontRgb[1], s.fontRgb[2]));
      if (s.fill_color === 'none') r.SetFillColor(null);
      else if (s.fillRgb) r.SetFillColor(Api.CreateColorFromRGB(s.fillRgb[0], s.fillRgb[1], s.fillRgb[2]));
      if (s.number_format) r.SetNumberFormat(s.number_format);
      if (s.align) r.SetAlignHorizontal(s.align);
      if (s.wrap !== undefined) r.SetWrap(s.wrap);
      if (s.merge === true) r.Merge(false);
      if (s.merge === false) r.UnMerge();
      if (s.column_width) r.SetColumnWidth(s.column_width);
      if (s.autofit) r.AutoFit(false, true);
      return { formatted: r.GetAddress(false, false) };
    }, {
      ...a,
      fontRgb: a.font_color ? hexToRgb(a.font_color) : undefined,
      fillRgb: a.fill_color && a.fill_color !== 'none' ? hexToRgb(a.fill_color) : undefined,
    }),
  },

  {
    name: 'cell_clear_range',
    title: 'Clear cells',
    description: 'Clear the content and formatting of a range.',
    schema: { ...documentArg, ...sheetArg, range: z.string() },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      ws.GetRange(s.range).Clear();
      return { cleared: s.range };
    }, a),
  },

  {
    name: 'cell_insert_cells',
    title: 'Insert rows/columns/cells',
    description: 'Insert empty cells at a range, shifting existing cells. Use a row range like "5:7" to insert whole rows, or a column range like "C:D" for columns.',
    schema: {
      ...documentArg,
      ...sheetArg,
      range: z.string(),
      shift: z.enum(['down', 'right']).optional().describe('Default: down for rows, right for columns.'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      ws.GetRange(s.range).Insert(s.shift);
      return { inserted: s.range };
    }, a),
  },

  {
    name: 'cell_delete_cells',
    title: 'Delete rows/columns/cells',
    description: 'Delete a range, shifting the remaining cells. Use "5:7" for whole rows or "C:D" for whole columns.',
    schema: {
      ...documentArg,
      ...sheetArg,
      range: z.string(),
      shift: z.enum(['up', 'left']).optional(),
    },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      ws.GetRange(s.range).Delete(s.shift);
      return { deleted: s.range };
    }, a),
  },

  {
    name: 'cell_manage_sheet',
    title: 'Add, rename, activate or delete a sheet',
    description: 'Add a sheet, rename a sheet, make a sheet active, or delete a sheet.',
    schema: {
      ...documentArg,
      action: z.enum(['add', 'rename', 'activate', 'delete']),
      sheet: z.string().optional().describe('Target sheet (rename/activate/delete). Default: the active sheet.'),
      name: z.string().optional().describe('New sheet name (add/rename).'),
    },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      if (s.action === 'add') {
        if (!s.name) throw new Error('"name" is required');
        Api.AddSheet(s.name);
        return { added: s.name };
      }
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      if (s.action === 'rename') {
        if (!s.name) throw new Error('"name" is required');
        ws.SetName(s.name);
        return { renamed: s.name };
      }
      if (s.action === 'activate') { ws.SetActive(); return { active: ws.GetName() }; }
      var n = ws.GetName();
      ws.Delete();
      return { deleted: n };
    }, a),
  },

  {
    name: 'cell_add_comment',
    title: 'Comment a cell',
    description: 'Add a comment to a cell.',
    schema: {
      ...documentArg,
      ...sheetArg,
      cell: z.string().describe('e.g. "B3"'),
      text: z.string().min(1),
      author: z.string().optional().describe('Default "Claude".'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope;
      var ws = s.sheet ? Api.GetSheet(s.sheet) : Api.GetActiveSheet();
      if (!ws) throw new Error('Unknown sheet: ' + s.sheet);
      ws.GetRange(s.cell).AddComment(s.text, s.author || 'Claude');
      return { commented: s.cell };
    }, a),
  },
];

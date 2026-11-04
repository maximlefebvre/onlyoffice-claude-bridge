// Tools for text documents (OnlyOffice "word" editor).
//
// Every `function () { ... }` passed to bridge.run() is serialized and executed
// inside the editor: it cannot use variables from this file, only `Api` and
// the values passed through `Asc.scope`.

import { z } from 'zod';
import { hexToRgb, positionSchema, documentArg } from './shared.mjs';

const targetSchema = {
  search: z.string().optional().describe('Apply to every occurrence of this text (takes precedence over paragraph).'),
  match_case: z.boolean().optional().describe('Case-sensitive search (default false).'),
  paragraph: z.number().int().min(0).optional().describe('Apply to the paragraph at this index (from word_read_document).'),
};

export const wordTools = [
  {
    name: 'word_read_document',
    title: 'Read text document',
    description:
      'Read the content of the open text document as a list of top-level elements (paragraphs with their index, style and text; tables as rows of cell text). Use the indexes with the other word_* tools. Paginate long documents with start/limit.',
    schema: {
      ...documentArg,
      start: z.number().int().min(0).optional().describe('First element index (default 0).'),
      limit: z.number().int().min(1).max(1000).optional().describe('Maximum number of elements (default 300).'),
    },
    readOnly: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument();
      var n = doc.GetElementsCount();
      var start = s.start || 0, end = Math.min(n, start + (s.limit || 300));
      var out = [];
      for (var i = start; i < end; i++) {
        var el = doc.GetElement(i), type = el.GetClassType(), item = { index: i, type: type };
        if (type === 'paragraph') {
          item.text = el.GetText().replace(/[\r\n]+$/, '');
          try { var st = el.GetStyle(); if (st) item.style = st.GetName(); } catch (e) {}
        } else if (type === 'table') {
          var rows = [];
          for (var r = 0; r < el.GetRowsCount(); r++) {
            var row = el.GetRow(r), cells = [];
            for (var k = 0; k < row.GetCellsCount(); k++) {
              cells.push(row.GetCell(k).GetContent().GetText().replace(/[\t\r\n]+$/, ''));
            }
            rows.push(cells);
          }
          item.rows = rows;
        } else {
          try { item.text = el.GetText(); } catch (e) {}
        }
        out.push(item);
      }
      return { total: n, start: start, end: end, elements: out };
    }, { start: a.start, limit: a.limit }, false),
  },

  {
    name: 'word_insert_paragraphs',
    title: 'Insert paragraphs',
    description:
      'Insert one or more paragraphs into the text document: at the cursor, at the start or end, or before/after a paragraph index. Each paragraph can have a style name (e.g. "Heading 1", "Heading 2", "Normal", "Title", "List Paragraph").',
    schema: {
      ...documentArg,
      paragraphs: z.array(z.object({
        text: z.string(),
        style: z.string().optional(),
      })).min(1).describe('Paragraphs to insert, in order.'),
      position: positionSchema,
      index: z.number().int().min(0).optional().describe('Reference paragraph index for position "before" or "after".'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), warnings = [];
      var paras = s.paragraphs.map(function (p) {
        var para = Api.CreateParagraph();
        para.AddText(p.text);
        if (p.style) {
          var st = doc.GetStyle(p.style);
          if (st) para.SetStyle(st); else warnings.push('Unknown style: ' + p.style);
        }
        return para;
      });
      var pos = s.position || 'cursor';
      if (pos === 'cursor') {
        doc.InsertContent(paras);
      } else if (pos === 'end') {
        paras.forEach(function (p) { doc.Push(p); });
      } else {
        var at = pos === 'start' ? 0 : (pos === 'before' ? s.index : s.index + 1);
        if ((pos === 'before' || pos === 'after') && (s.index === undefined || s.index === null)) {
          throw new Error('"index" is required for position ' + pos);
        }
        paras.forEach(function (p, k) { doc.AddElement(at + k, p); });
      }
      return { inserted: paras.length, warnings: warnings };
    }, a),
  },

  {
    name: 'word_replace_paragraph',
    title: 'Replace paragraph text',
    description: 'Replace the whole text of the paragraph at an index, keeping its paragraph formatting. Optionally change its style.',
    schema: {
      ...documentArg,
      index: z.number().int().min(0),
      text: z.string(),
      style: z.string().optional(),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), el = doc.GetElement(s.index);
      if (!el || el.GetClassType() !== 'paragraph') throw new Error('No paragraph at index ' + s.index);
      el.RemoveAllElements();
      el.AddText(s.text);
      if (s.style) {
        var st = doc.GetStyle(s.style);
        if (!st) throw new Error('Unknown style: ' + s.style);
        el.SetStyle(st);
      }
      return { replaced: s.index };
    }, a),
  },

  {
    name: 'word_delete_elements',
    title: 'Delete paragraphs or tables',
    description: 'Delete `count` consecutive top-level elements (paragraphs or tables) starting at an index.',
    schema: {
      ...documentArg,
      index: z.number().int().min(0),
      count: z.number().int().min(1).optional().describe('Default 1.'),
    },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), count = s.count || 1;
      var n = doc.GetElementsCount();
      if (s.index >= n) throw new Error('Index out of range (document has ' + n + ' elements)');
      var last = Math.min(n - 1, s.index + count - 1);
      for (var i = last; i >= s.index; i--) doc.RemoveElement(i);
      return { deleted: last - s.index + 1 };
    }, a),
  },

  {
    name: 'word_search_replace',
    title: 'Search and replace',
    description: 'Replace every occurrence of a text in the document. Returns the number of occurrences found.',
    schema: {
      ...documentArg,
      search: z.string().min(1),
      replace: z.string(),
      match_case: z.boolean().optional().describe('Default false.'),
    },
    run: async (b, c, a) => {
      const found = await b.run(c, function () {
        var s = Asc.scope;
        var r = Api.GetDocument().Search(s.search, !!s.match_case);
        return r ? r.length : 0;
      }, a, false);
      if (found) {
        await b.method(c, 'SearchAndReplace', [{ searchString: a.search, replaceString: a.replace, matchCase: !!a.match_case }]);
      }
      return { replaced: found };
    },
  },

  {
    name: 'word_format_text',
    title: 'Format text',
    description:
      'Apply character formatting (bold, italic, underline, strikeout, size, font, color, highlight) and/or paragraph alignment to the current selection (default), to a paragraph index, or to every occurrence of a text.',
    schema: {
      ...documentArg,
      ...targetSchema,
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikeout: z.boolean().optional(),
      font_size: z.number().positive().optional().describe('In points.'),
      font_family: z.string().optional(),
      color: z.string().optional().describe('Text color as #RRGGBB.'),
      highlight: z.enum(['yellow', 'green', 'cyan', 'magenta', 'blue', 'red', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray', 'black', 'none']).optional(),
      align: z.enum(['left', 'center', 'right', 'both']).optional().describe('Paragraph alignment ("both" = justified).'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), ranges;
      if (s.search) {
        ranges = doc.Search(s.search, !!s.match_case) || [];
      } else if (s.paragraph !== undefined && s.paragraph !== null) {
        var el = doc.GetElement(s.paragraph);
        if (!el || el.GetClassType() !== 'paragraph') throw new Error('No paragraph at index ' + s.paragraph);
        ranges = [el.GetRange()];
      } else {
        var sel = doc.GetRangeBySelect();
        if (!sel) throw new Error('Nothing is selected in the document');
        ranges = [sel];
      }
      ranges.forEach(function (r) {
        if (s.bold !== undefined) r.SetBold(s.bold);
        if (s.italic !== undefined) r.SetItalic(s.italic);
        if (s.underline !== undefined) r.SetUnderline(s.underline);
        if (s.strikeout !== undefined) r.SetStrikeout(s.strikeout);
        if (s.font_size) r.SetFontSize(Math.round(s.font_size * 2));
        if (s.font_family) r.SetFontFamily(s.font_family);
        if (s.rgb) r.SetColor(s.rgb[0], s.rgb[1], s.rgb[2]);
        if (s.highlight) r.SetHighlight(s.highlight);
        if (s.align) r.GetAllParagraphs().forEach(function (p) { p.SetJc(s.align); });
      });
      return { formatted: ranges.length };
    }, { ...a, rgb: a.color ? hexToRgb(a.color) : undefined }),
  },

  {
    name: 'word_set_style',
    title: 'Set paragraph style',
    description: 'Apply a paragraph style (e.g. "Heading 1", "Heading 2", "Normal", "Title", "Quote") to one or more paragraphs.',
    schema: {
      ...documentArg,
      indexes: z.array(z.number().int().min(0)).min(1),
      style: z.string(),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), st = doc.GetStyle(s.style);
      if (!st) throw new Error('Unknown style: ' + s.style);
      s.indexes.forEach(function (i) {
        var el = doc.GetElement(i);
        if (!el || el.GetClassType() !== 'paragraph') throw new Error('No paragraph at index ' + i);
        el.SetStyle(st);
      });
      return { styled: s.indexes.length };
    }, a),
  },

  {
    name: 'word_add_comment',
    title: 'Add comment',
    description: 'Add a review comment on the current selection (default), on a paragraph index, or on the first occurrence of a text (all occurrences with all=true).',
    schema: {
      ...documentArg,
      ...targetSchema,
      text: z.string().min(1),
      author: z.string().optional().describe('Default "Claude".'),
      all: z.boolean().optional(),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument(), ranges;
      if (s.search) {
        ranges = doc.Search(s.search, !!s.match_case) || [];
        if (!s.all) ranges = ranges.slice(0, 1);
        if (!ranges.length) throw new Error('Text not found: ' + s.search);
      } else if (s.paragraph !== undefined && s.paragraph !== null) {
        var el = doc.GetElement(s.paragraph);
        if (!el) throw new Error('No element at index ' + s.paragraph);
        ranges = [el.GetRange()];
      } else {
        var sel = doc.GetRangeBySelect();
        if (!sel) throw new Error('Nothing is selected in the document');
        ranges = [sel];
      }
      ranges.forEach(function (r) { r.AddComment(s.text, s.author || 'Claude'); });
      return { comments: ranges.length };
    }, a),
  },

  {
    name: 'word_list_comments',
    title: 'List comments',
    description: 'List the comments of the text document with their author and the text they refer to.',
    schema: { ...documentArg },
    readOnly: true,
    run: (b, c) => b.run(c, function () {
      return (Api.GetDocument().GetAllComments() || []).map(function (cm, i) {
        var o = { index: i, text: cm.GetText(), author: cm.GetAuthorName() };
        try { o.quote = cm.GetQuoteText(); } catch (e) {}
        return o;
      });
    }, {}, false),
  },

  {
    name: 'word_insert_table',
    title: 'Insert table',
    description: 'Insert a table filled with the given rows (first row can be bold as a header), at the cursor, start, end, or before/after a paragraph index.',
    schema: {
      ...documentArg,
      rows: z.array(z.array(z.string())).min(1).describe('Table content, row by row.'),
      header: z.boolean().optional().describe('Make the first row bold (default true).'),
      position: positionSchema,
      index: z.number().int().min(0).optional(),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, doc = Api.GetDocument();
      var nRows = s.rows.length, nCols = Math.max.apply(null, s.rows.map(function (r) { return r.length; }));
      // OnlyOffice 9.4 swapped the arguments from (cols, rows) to (rows, cols)
      // (kDrive runs 10.x). Older editors get the old order as a fallback:
      // create the table, then retry the other way if the shape is wrong.
      var t = Api.CreateTable(nRows, nCols);
      if (t.GetRowsCount() !== nRows) t = Api.CreateTable(nCols, nRows);
      try { var st = doc.GetStyle('Bordered'); if (st) t.SetStyle(st); } catch (e) {}
      try { t.SetWidth('percent', 100); } catch (e) {}
      for (var r = 0; r < nRows; r++) {
        for (var k = 0; k < nCols; k++) {
          var p = t.GetCell(r, k).GetContent().GetElement(0);
          p.AddText(s.rows[r][k] || '');
          if (r === 0 && s.header !== false) p.SetBold(true);
        }
      }
      var pos = s.position || 'cursor';
      if (pos === 'cursor') doc.InsertContent([t]);
      else if (pos === 'end') doc.Push(t);
      else if (pos === 'start') doc.AddElement(0, t);
      else {
        if (s.index === undefined || s.index === null) throw new Error('"index" is required for position ' + pos);
        doc.AddElement(pos === 'before' ? s.index : s.index + 1, t);
      }
      return { rows: nRows, cols: nCols };
    }, a),
  },

  {
    name: 'word_set_track_changes',
    title: 'Track changes',
    description: 'Turn track changes on or off, so the user can review and accept/reject Claude\'s edits.',
    schema: { ...documentArg, enabled: z.boolean() },
    run: (b, c, a) => b.run(c, function () {
      Api.GetDocument().SetTrackRevisions(!!Asc.scope.enabled);
      return { trackChanges: !!Asc.scope.enabled };
    }, a),
  },
];

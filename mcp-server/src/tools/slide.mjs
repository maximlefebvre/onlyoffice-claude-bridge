// Tools for presentations (OnlyOffice "slide" editor).
// Functions passed to bridge.run() execute inside the editor (see word.mjs).
// Positions and sizes are given in percent of the slide size.

import { z } from 'zod';
import { documentArg, hexToRgb } from './shared.mjs';

const boxSchema = {
  x: z.number().min(0).max(100).optional().describe('Left edge, % of slide width.'),
  y: z.number().min(0).max(100).optional().describe('Top edge, % of slide height.'),
  width: z.number().min(1).max(100).optional().describe('% of slide width.'),
  height: z.number().min(1).max(100).optional().describe('% of slide height.'),
};

const textStyleSchema = {
  font_family: z.string().optional().describe('Font name, e.g. "Inter".'),
  font_size: z.number().positive().optional().describe('In points.'),
  bold: z.boolean().optional(),
  color: z.string().optional().describe('#RRGGBB'),
  align: z.enum(['left', 'center', 'right', 'both']).optional(),
  line_spacing: z.number().positive().max(10).optional().describe('Multiple, e.g. 1.2.'),
};

export const slideTools = [
  {
    name: 'slide_list',
    title: 'List slides',
    description: 'List the slides of the presentation with, for each one, its shapes (index, name, placeholder type, text) and the current slide index.',
    schema: { ...documentArg },
    readOnly: true,
    run: (b, c) => b.run(c, function () {
      var pres = Api.GetPresentation(), n = pres.GetSlidesCount(), current = null;
      try { current = pres.GetCurrentSlide().GetSlideIndex(); } catch (e) {}
      var slides = [];
      for (var i = 0; i < n; i++) {
        var slide = pres.GetSlideByIndex(i), shapes = [];
        (slide.GetAllShapes() || []).forEach(function (sh, k) {
          var o = { index: k };
          try { o.name = sh.GetName(); } catch (e) {}
          try { var ph = sh.GetPlaceholder(); if (ph) o.placeholder = ph.GetType(); } catch (e) {}
          try { var dc = sh.GetDocContent(); if (dc) o.text = dc.GetText().replace(/[\r\n]+$/, ''); } catch (e) {}
          shapes.push(o);
        });
        var notes = null;
        try { notes = slide.GetNotesPage().GetBodyShapeText(); } catch (e) {}
        slides.push({ index: i, shapes: shapes, notes: notes || undefined });
      }
      return { count: n, current: current, slides: slides };
    }, {}, false),
  },

  {
    name: 'slide_add',
    title: 'Add slide',
    description: 'Add a slide with an optional title and body text (one bullet/paragraph per line). Inserted at `position` (default: after the last slide).',
    schema: {
      ...documentArg,
      position: z.number().int().min(0).optional().describe('Index of the new slide.'),
      title: z.string().optional(),
      body: z.array(z.string()).optional().describe('Body paragraphs.'),
      title_size: z.number().positive().optional().describe('Points, default 36.'),
      body_size: z.number().positive().optional().describe('Points, default 20.'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, pres = Api.GetPresentation();
      var W = pres.GetWidth(), H = pres.GetHeight();
      var slide = Api.CreateSlide();
      var idx = (s.position === undefined || s.position === null) ? pres.GetSlidesCount() : s.position;
      pres.AddSlide(slide, idx);
      function box(lines, x, y, w, h, size, bold) {
        var sh = Api.CreateShape('rect', w * W, h * H, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));
        sh.SetPosition(x * W, y * H);
        var dc = sh.GetDocContent();
        dc.RemoveAllElements();
        lines.forEach(function (t) {
          var p = Api.CreateParagraph();
          p.AddText(t);
          p.SetFontSize(Math.round(size * 2));
          if (bold) p.SetBold(true);
          dc.Push(p);
        });
        slide.AddObject(sh);
      }
      if (s.title) box([s.title], 0.06, 0.06, 0.88, 0.18, s.title_size || 36, true);
      if (s.body && s.body.length) box(s.body, 0.06, s.title ? 0.27 : 0.08, 0.88, s.title ? 0.65 : 0.84, s.body_size || 20, false);
      return { added: idx };
    }, a),
  },

  {
    name: 'slide_manage',
    title: 'Delete, duplicate, move or show a slide',
    description: 'Delete, duplicate or move a slide, or make it the current slide in the editor.',
    schema: {
      ...documentArg,
      action: z.enum(['delete', 'duplicate', 'move', 'select']),
      slide: z.number().int().min(0),
      to: z.number().int().min(0).optional().describe('Target index for move / duplicate.'),
    },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, pres = Api.GetPresentation(), slide = pres.GetSlideByIndex(s.slide);
      if (!slide) throw new Error('No slide at index ' + s.slide);
      if (s.action === 'delete') slide.Delete();
      else if (s.action === 'duplicate') slide.Duplicate(s.to === undefined ? s.slide + 1 : s.to);
      else if (s.action === 'move') {
        if (s.to === undefined) throw new Error('"to" is required');
        slide.MoveTo(s.to);
      } else slide.Select();
      return { done: s.action, slides: pres.GetSlidesCount() };
    }, a),
  },

  {
    name: 'slide_set_text',
    title: 'Set shape text',
    description: 'Replace the text of a shape (title, body, text box…) on a slide. One paragraph per line. Shape indexes come from slide_list.',
    schema: {
      ...documentArg,
      slide: z.number().int().min(0),
      shape: z.number().int().min(0),
      lines: z.array(z.string()).min(1),
      ...textStyleSchema,
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, slide = Api.GetPresentation().GetSlideByIndex(s.slide);
      if (!slide) throw new Error('No slide at index ' + s.slide);
      var sh = (slide.GetAllShapes() || [])[s.shape];
      if (!sh) throw new Error('No shape at index ' + s.shape + ' on slide ' + s.slide);
      var dc = sh.GetDocContent();
      if (!dc) throw new Error('This shape has no text');
      dc.RemoveAllElements();
      s.lines.forEach(function (t) {
        var p = Api.CreateParagraph();
        p.AddText(t);
        if (s.font_size) p.SetFontSize(Math.round(s.font_size * 2));
        if (s.bold !== undefined) p.SetBold(s.bold);
        if (s.rgb) p.SetColor(s.rgb[0], s.rgb[1], s.rgb[2]);
        if (s.align) p.SetJc(s.align);
        if (s.font_family) p.SetFontFamily(s.font_family);
        if (s.line_spacing) p.SetSpacingLine(Math.round(s.line_spacing * 240), 'auto');
        dc.Push(p);
      });
      return { updated: { slide: s.slide, shape: s.shape } };
    }, { ...a, rgb: a.color ? hexToRgb(a.color) : undefined }),
  },

  {
    name: 'slide_add_text_box',
    title: 'Add text box',
    description: 'Add a text box to a slide. Position and size are in percent of the slide.',
    schema: {
      ...documentArg,
      slide: z.number().int().min(0),
      lines: z.array(z.string()).min(1),
      ...boxSchema,
      ...textStyleSchema,
      fill: z.string().optional().describe('Background color #RRGGBB (default none).'),
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, pres = Api.GetPresentation(), slide = pres.GetSlideByIndex(s.slide);
      if (!slide) throw new Error('No slide at index ' + s.slide);
      var W = pres.GetWidth(), H = pres.GetHeight();
      var x = (s.x === undefined ? 10 : s.x) / 100, y = (s.y === undefined ? 40 : s.y) / 100;
      var w = (s.width === undefined ? 80 : s.width) / 100, h = (s.height === undefined ? 20 : s.height) / 100;
      var fill = s.fillRgb ? Api.CreateSolidFill(Api.CreateRGBColor(s.fillRgb[0], s.fillRgb[1], s.fillRgb[2])) : Api.CreateNoFill();
      var sh = Api.CreateShape('rect', w * W, h * H, fill, Api.CreateStroke(0, Api.CreateNoFill()));
      sh.SetPosition(x * W, y * H);
      var dc = sh.GetDocContent();
      dc.RemoveAllElements();
      s.lines.forEach(function (t) {
        var p = Api.CreateParagraph();
        p.AddText(t);
        p.SetFontSize(Math.round((s.font_size || 18) * 2));
        if (s.bold) p.SetBold(true);
        if (s.rgb) p.SetColor(s.rgb[0], s.rgb[1], s.rgb[2]);
        if (s.align) p.SetJc(s.align);
        if (s.font_family) p.SetFontFamily(s.font_family);
        if (s.line_spacing) p.SetSpacingLine(Math.round(s.line_spacing * 240), 'auto');
        dc.Push(p);
      });
      slide.AddObject(sh);
      return { added: { slide: s.slide, shape: (slide.GetAllShapes() || []).length - 1 } };
    }, {
      ...a,
      rgb: a.color ? hexToRgb(a.color) : undefined,
      fillRgb: a.fill ? hexToRgb(a.fill) : undefined,
    }),
  },

  {
    name: 'slide_format_text',
    title: 'Format text on slides',
    description:
      'Change the text formatting (font, size, color, bold, italic, alignment, line spacing) of one shape, of every shape on a slide, or of every shape in the presentation (all=true). Text content is kept.',
    schema: {
      ...documentArg,
      all: z.boolean().optional().describe('Every shape on every slide.'),
      slide: z.number().int().min(0).optional(),
      shape: z.number().int().min(0).optional().describe('Omit to format every shape of the slide.'),
      italic: z.boolean().optional(),
      ...textStyleSchema,
    },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, pres = Api.GetPresentation(), shapes = [];
      if (s.all) {
        for (var i = 0; i < pres.GetSlidesCount(); i++) shapes = shapes.concat(pres.GetSlideByIndex(i).GetAllShapes() || []);
      } else {
        if (s.slide === undefined || s.slide === null) throw new Error('Pass all=true or a slide index');
        var slide = pres.GetSlideByIndex(s.slide);
        if (!slide) throw new Error('No slide at index ' + s.slide);
        var list = slide.GetAllShapes() || [];
        if (s.shape !== undefined && s.shape !== null) {
          if (!list[s.shape]) throw new Error('No shape at index ' + s.shape + ' on slide ' + s.slide);
          shapes = [list[s.shape]];
        } else shapes = list;
      }
      var count = 0;
      shapes.forEach(function (sh) {
        var dc = null;
        try { dc = sh.GetDocContent(); } catch (e) {}
        if (!dc) return;
        dc.GetAllParagraphs().forEach(function (p) {
          if (s.font_family) p.SetFontFamily(s.font_family);
          if (s.font_size) p.SetFontSize(Math.round(s.font_size * 2));
          if (s.rgb) p.SetColor(s.rgb[0], s.rgb[1], s.rgb[2]);
          if (s.bold !== undefined) p.SetBold(s.bold);
          if (s.italic !== undefined) p.SetItalic(s.italic);
          if (s.align) p.SetJc(s.align);
          if (s.line_spacing) p.SetSpacingLine(Math.round(s.line_spacing * 240), 'auto');
          count++;
        });
      });
      return { shapes: shapes.length, paragraphs: count };
    }, { ...a, rgb: a.color ? hexToRgb(a.color) : undefined }),
  },

  {
    name: 'slide_edit_shape',
    title: 'Move, resize or delete a shape',
    description: 'Move and/or resize a shape (percent of the slide), or delete it.',
    schema: {
      ...documentArg,
      slide: z.number().int().min(0),
      shape: z.number().int().min(0),
      delete: z.boolean().optional(),
      ...boxSchema,
    },
    destructive: true,
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, pres = Api.GetPresentation(), slide = pres.GetSlideByIndex(s.slide);
      if (!slide) throw new Error('No slide at index ' + s.slide);
      var sh = (slide.GetAllShapes() || [])[s.shape];
      if (!sh) throw new Error('No shape at index ' + s.shape + ' on slide ' + s.slide);
      if (s['delete']) { sh.Delete(); return { deleted: s.shape }; }
      var W = pres.GetWidth(), H = pres.GetHeight();
      if (s.width !== undefined || s.height !== undefined) {
        sh.SetSize(s.width !== undefined ? s.width / 100 * W : sh.GetWidth(), s.height !== undefined ? s.height / 100 * H : sh.GetHeight());
      }
      if (s.x !== undefined || s.y !== undefined) {
        sh.SetPosition(s.x !== undefined ? s.x / 100 * W : sh.GetPosX(), s.y !== undefined ? s.y / 100 * H : sh.GetPosY());
      }
      return { updated: s.shape };
    }, a),
  },

  {
    name: 'slide_set_notes',
    title: 'Set speaker notes',
    description: 'Replace the speaker notes of a slide.',
    schema: { ...documentArg, slide: z.number().int().min(0), text: z.string() },
    run: (b, c, a) => b.run(c, function () {
      var s = Asc.scope, slide = Api.GetPresentation().GetSlideByIndex(s.slide);
      if (!slide) throw new Error('No slide at index ' + s.slide);
      var page = slide.GetNotesPage();
      if (!page) throw new Error('This slide has no notes page');
      var body = page.GetBodyShape();
      if (body && body.GetDocContent()) {
        var dc = body.GetDocContent();
        dc.RemoveAllElements();
        s.text.split('\n').forEach(function (t) { var p = Api.CreateParagraph(); p.AddText(t); dc.Push(p); });
      } else {
        page.AddBodyShapeText(s.text);
      }
      return { notes: s.slide };
    }, a),
  },
];

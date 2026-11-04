// api_reference: searchable OnlyOffice Office API documentation, generated from
// the sdkjs sources by scripts/build-api-reference.mjs.

import { readFileSync } from 'node:fs';
import { z } from 'zod';

let ref = null;
function load() {
  if (!ref) {
    ref = JSON.parse(readFileSync(new URL('../../data/api-reference.json', import.meta.url), 'utf8'));
    for (const m of ref.methods) {
      m._name = splitWords(m.method).join(' ');
      m._class = splitWords(m.class.replace(/^Api/, '')).join(' ');
      m._desc = m.description.toLowerCase();
    }
  }
  return ref;
}

function splitWords(s) {
  return String(s).replace(/([a-z0-9])([A-Z])/g, '$1 $2').replace(/_/g, ' ').toLowerCase().split(/\s+/).filter(Boolean);
}

function signature(m) {
  const params = m.params.map((p) => `${p.optional ? '[' : ''}${p.name}: ${p.type}${p.default !== undefined ? ` = ${p.default}` : ''}${p.optional ? ']' : ''}`).join(', ');
  const owner = m.static ? m.class : `${m.class}.prototype`;
  return `${owner}.${m.method}(${params})${m.returns ? ` → ${m.returns.type}` : ''}`;
}

function format(m, detailed) {
  const o = { signature: signature(m), editors: m.editors, description: detailed ? m.description : m.description.slice(0, 220) };
  if (detailed) {
    o.params = m.params;
    if (m.returns?.description) o.returns = m.returns.description;
  }
  if (m.since) o.since = m.since;
  if (m.deprecated) o.deprecated = m.deprecated;
  return o;
}

// Classes whose methods are available on `cls` (itself + ancestors).
function lineage(cls) {
  const out = [cls];
  for (let c = cls; ref.inherits[c]; c = ref.inherits[c]) out.push(ref.inherits[c]);
  return out;
}

export const referenceTools = [
  {
    name: 'api_reference',
    title: 'Search the OnlyOffice API reference',
    description: [
      'Look up the OnlyOffice Office API (the `Api` used by run_script) before writing a script: search methods by English keywords',
      '(e.g. "footnote", "chart", "header footer", "line spacing", "image", "hyperlink", "conditional format"),',
      'list all methods of a class (class: "ApiParagraph", including inherited ones), get full details of one method (method: "ApiRange.SetBorders"),',
      'or the allowed values of a type (type: "LineStyle"). Generated from the OnlyOffice 10.0 sources used by kDrive.',
    ].join(' '),
    schema: {
      query: z.string().optional().describe('English keywords.'),
      class: z.string().optional().describe('List the methods of this class, e.g. "ApiWorksheet".'),
      method: z.string().optional().describe('Full details of a method: "Class.Method" or just "Method".'),
      type: z.string().optional().describe('Show a type or enumeration, e.g. "LineStyle", "ChartType".'),
      editor: z.enum(['word', 'cell', 'slide']).optional().describe('word = text documents, cell = spreadsheets, slide = presentations.'),
      limit: z.number().int().min(1).max(60).optional().describe('Max results for query (default 15).'),
    },
    readOnly: true,
    noDocument: true,
    run: (_b, _c, a) => {
      const r = load();
      const inEditor = (m) => !a.editor || m.editors.includes(a.editor);

      if (a.type) {
        const types = r.types.filter((t) => t.name.toLowerCase() === a.type.toLowerCase() && (!a.editor || t.editor === a.editor));
        if (!types.length) {
          const close = r.types.filter((t) => t.name.toLowerCase().includes(a.type.toLowerCase())).slice(0, 20).map((t) => t.name);
          throw new Error(`Unknown type "${a.type}".${close.length ? ` Similar: ${[...new Set(close)].join(', ')}` : ''}`);
        }
        return types.map(({ editor, name, type, description }) => ({ editor, name, type, description }));
      }

      if (a.method) {
        const [c, m] = a.method.includes('.') ? a.method.split('.') : [null, a.method];
        const found = r.methods.filter((x) => x.method.toLowerCase() === m.toLowerCase()
          && (!c || lineage(c).some((k) => k.toLowerCase() === x.class.toLowerCase()) || x.class.toLowerCase() === c.toLowerCase())
          && inEditor(x));
        if (!found.length) throw new Error(`No method "${a.method}" found. Try query or class.`);
        return found.slice(0, 10).map((x) => format(x, true));
      }

      if (a.class) {
        const chain = lineage(a.class);
        const own = r.methods.filter((x) => chain.includes(x.class) && inEditor(x));
        if (!own.length) {
          const classes = [...new Set(r.methods.filter(inEditor).map((x) => x.class))].sort();
          throw new Error(`Unknown class "${a.class}". Classes: ${classes.join(', ')}`);
        }
        return {
          class: a.class,
          inherits: chain.slice(1),
          methods: own.map((x) => ({ signature: signature(x), from: x.class !== a.class ? x.class : undefined, description: x.description.slice(0, 120) })),
        };
      }

      if (!a.query) throw new Error('Pass query, class, method or type.');
      const words = splitWords(a.query).filter((w) => w.length > 1);
      const scored = [];
      for (const m of r.methods) {
        if (!inEditor(m)) continue;
        let score = 0;
        for (const w of words) {
          if (m._name.split(' ').includes(w)) score += 6;
          else if (m._name.includes(w)) score += 3;
          if (m._class.includes(w)) score += 2;
          if (m._desc.includes(w)) score += 1;
        }
        if (m.deprecated) score -= 2;
        if (score > 0) scored.push([score, m]);
      }
      scored.sort((x, y) => y[0] - x[0]);
      return {
        source: r.source,
        results: scored.slice(0, a.limit || 15).map(([, m]) => format(m, false)),
        hint: 'Use method: "Class.Method" for parameters and details.',
      };
    },
  },
];

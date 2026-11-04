#!/usr/bin/env node
// Builds data/api-reference.json from the JSDoc comments of OnlyOffice's
// Document Builder API (ONLYOFFICE/sdkjs: word/, cell/, slide/apiBuilder.js).
//
//   node scripts/build-api-reference.mjs [sdkjs-tag]      (default: v10.0.0.176)
//
// The tag should match the OnlyOffice version of the editor you use
// (kDrive reports 10.0.2 through executeMethod("GetVersion")).

import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const TAG = process.argv[2] || 'v10.0.0.176';
const EDITORS = { word: 'CDE', cell: 'CSE', slide: 'CPE' };
const here = dirname(fileURLToPath(import.meta.url));

function clean(text) {
  return text
    .replace(/<\/?(em|b|strong|code|i)>/g, '')
    .replace(/<br\s*\/?>/g, ' ')
    .replace(/\{@link [^}]*\}/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Splits a JSDoc block into its description and tags.
function parseDoc(block) {
  const lines = block
    .replace(/^\/\*\*|\*\/$/g, '')
    .split('\n')
    .map((l) => l.replace(/^\s*\*\s?/, ''));
  const doc = { description: [], params: [], returns: null, since: null, deprecated: false, editors: null, memberof: null, see: null, typedef: null };
  let current = null;
  for (const line of lines) {
    const tag = /^@(\w+)\s*(.*)$/.exec(line.trim());
    if (tag) {
      const [, name, rest] = tag;
      current = name;
      if (name === 'param') {
        const m = /^\{(.+?)\}\s+(\[?[\w.$]+(?:=[^\]]*)?\]?)\s*-?\s*(.*)$/.exec(rest);
        if (m) {
          const raw = m[2];
          const optional = raw.startsWith('[');
          const [pname, def] = raw.replace(/^\[|\]$/g, '').split('=');
          doc.params.push({ name: pname, type: m[1], optional, default: def, description: m[3] });
        }
      } else if (name === 'returns' || name === 'return') {
        const m = /^\{(.+?)\}\s*-?\s*(.*)$/.exec(rest);
        doc.returns = m ? { type: m[1], description: m[2] } : { type: rest, description: '' };
        current = 'returns';
      } else if (name === 'since') doc.since = rest.trim();
      else if (name === 'deprecated') doc.deprecated = rest.trim() || true;
      else if (name === 'typeofeditors') { try { doc.editors = JSON.parse(rest.trim()); } catch {} }
      else if (name === 'memberof') doc.memberof = rest.trim();
      else if (name === 'typedef') doc.typedef = rest.trim();
      else if (name === 'description') doc.description.push(rest);
      continue;
    }
    if (current === null || current === 'description') doc.description.push(line);
    else if (current === 'param' && doc.params.length) doc.params.at(-1).description += ' ' + line.trim();
    else if (current === 'returns' && doc.returns) doc.returns.description += ' ' + line.trim();
  }
  doc.description = clean(doc.description.join(' ').replace(/:::[a-z]+\[[^\]]*\]|:::/gi, ' '));
  for (const p of doc.params) p.description = clean(p.description);
  if (doc.returns) doc.returns.description = clean(doc.returns.description);
  return doc;
}

function extract(source, editor) {
  const entries = [];
  const typedefs = [];
  const re = /\/\*\*([\s\S]*?)\*\/\s*\n\s*(?:(Api\w*)\.prototype\.(\w+)|(Api)\.(\w+))\s*=\s*function\s*\(([^)]*)\)/g;
  let m;
  while ((m = re.exec(source))) {
    const doc = parseDoc('/**' + m[1] + '*/');
    const cls = m[2] || m[4];
    const name = m[3] || m[5];
    if (/^private_|^_/.test(name)) continue;
    if (/@undocumented/.test(m[1])) continue;
    const CODES = { CDE: 'word', CSE: 'cell', CPE: 'slide' };
    const editors = doc.editors ? doc.editors.map((x) => CODES[x]).filter(Boolean) : [editor];
    if (!editors.length) continue; // e.g. PDF-only methods
    entries.push({
      editors,
      class: cls,
      method: name,
      static: !m[2],
      description: doc.description,
      params: doc.params.map(({ name: n, type, optional, default: d, description }) => ({ name: n, type, optional: optional || undefined, default: d, description: description || undefined })),
      returns: doc.returns,
      since: doc.since || undefined,
      deprecated: doc.deprecated || undefined,
    });
  }
  // Enumerations / object shapes referenced by parameter types.
  const tre = /\/\*\*((?:(?!\*\/)[\s\S])*?)\*\//g; // one comment block at a time
  while ((m = tre.exec(source))) {
    if (!m[1].includes('@typedef')) continue;
    const body = m[1].replace(/^\s*\*\s?/gm, '');
    const t = /@typedef\s+\{([\s\S]+?)\}\s+(\w+)/.exec(body);
    if (!t) continue;
    const desc = clean(m[1].split('@typedef')[0].replace(/^\s*\*\s?/gm, ''));
    typedefs.push({ editor, name: t[2], type: t[1].replace(/\s+/g, ' ').trim(), description: desc || undefined });
  }
  // Class inheritance (e.g. ApiParagraph -> ApiParaPr).
  const inherits = {};
  const ire = /(Api\w+)\.prototype\s*=\s*Object\.create\((Api\w+)\.prototype\)/g;
  while ((m = ire.exec(source))) inherits[m[1]] = m[2];
  return { entries, typedefs, inherits };
}

// inherits is per editor: the same class name can have a different parent in
// each editor (e.g. ApiTable extends ApiTablePr in word, ApiDrawing in slide).
const out = { source: `ONLYOFFICE/sdkjs@${TAG}`, generated: new Date().toISOString(), inherits: {}, methods: [], types: [] };
for (const editor of Object.keys(EDITORS)) {
  const url = `https://raw.githubusercontent.com/ONLYOFFICE/sdkjs/${TAG}/${editor}/apiBuilder.js`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${url}: HTTP ${res.status}`);
  const { entries, typedefs, inherits } = extract(await res.text(), editor);
  out.inherits[editor] = inherits;
  out.methods.push(...entries);
  out.types.push(...typedefs);
  console.error(`${editor}: ${entries.length} methods, ${typedefs.length} types`);
}

// The same method can be documented in several files: merge them.
const byKey = new Map();
for (const e of out.methods) {
  const key = `${e.class}.${e.method}`;
  const prev = byKey.get(key);
  if (!prev) { byKey.set(key, e); continue; }
  prev.editors = [...new Set([...prev.editors, ...e.editors])];
}
out.methods = [...byKey.values()];
const seenTypes = new Set();
out.types = out.types.filter((t) => { const k = t.editor + t.name; if (seenTypes.has(k)) return false; seenTypes.add(k); return true; });

const file = join(here, '..', 'data', 'api-reference.json');
mkdirSync(dirname(file), { recursive: true });
writeFileSync(file, JSON.stringify(out));
console.error(`wrote ${file}: ${out.methods.length} methods, ${out.types.length} types`);

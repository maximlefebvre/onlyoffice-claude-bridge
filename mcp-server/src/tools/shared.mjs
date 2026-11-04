import { z } from 'zod';

export const documentArg = {
  document: z.union([z.number().int(), z.string()]).optional()
    .describe('Target document: id or part of the title from list_documents. Default: the most recently used connected document of the right type.'),
};

export const positionSchema = z.enum(['cursor', 'start', 'end', 'before', 'after']).optional()
  .describe('Where to insert (default "cursor"). "before"/"after" need "index".');

export function hexToRgb(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex).trim());
  if (!m) throw new Error(`Invalid color "${hex}", expected #RRGGBB`);
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

// ---- A1 notation helpers (spreadsheets) ----

export function colToIndex(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToCol(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

export function parseCell(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(ref).trim());
  if (!m) throw new Error(`Invalid cell reference "${ref}", expected e.g. "B3"`);
  return { col: colToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

export function cellRef(row, col) {
  return `${indexToCol(col)}${row + 1}`;
}

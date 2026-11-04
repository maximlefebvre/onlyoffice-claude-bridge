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

// ---- Images ----

import { readFile } from 'node:fs/promises';
import { homedir as home } from 'node:os';
import { extname } from 'node:path';

export const EMU_PER_CM = 360000;
const MAX_IMAGE_BYTES = 15 * 1024 * 1024;
const MIME = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.bmp': 'image/bmp', '.svg': 'image/svg+xml' };

export const imageSourceSchema = z.string().min(1)
  .describe('Image to insert: an http(s) URL, an absolute path to a file on this computer (e.g. "~/Pictures/logo.png"), or a data: URL.');

function imageSize(buf, mime) {
  try {
    if (mime === 'image/png' && buf.readUInt32BE(12) === 0x49484452) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
    if (mime === 'image/gif') return { w: buf.readUInt16LE(6), h: buf.readUInt16LE(8) };
    if (mime === 'image/bmp') return { w: buf.readInt32LE(18), h: Math.abs(buf.readInt32LE(22)) };
    if (mime === 'image/jpeg') {
      for (let i = 2; i < buf.length - 9;) {
        if (buf[i] !== 0xff) { i++; continue; }
        const marker = buf[i + 1];
        const len = buf.readUInt16BE(i + 2);
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return { w: buf.readUInt16BE(i + 7), h: buf.readUInt16BE(i + 5) };
        i += 2 + len;
      }
    }
    if (mime === 'image/webp' && buf.toString('ascii', 12, 16) === 'VP8X') {
      return { w: 1 + buf.readUIntLE(24, 3), h: 1 + buf.readUIntLE(27, 3) };
    }
  } catch {}
  return null;
}

// Loads an image and returns it as a data URL (the editor accepts URLs and
// base64; embedding avoids depending on the document server fetching the URL).
export async function loadImage(source) {
  let buf, mime;
  const src = String(source).trim();
  if (src.startsWith('data:')) {
    const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(src);
    if (!m) throw new Error('Invalid data URL');
    mime = m[1];
    buf = m[2] ? Buffer.from(m[3], 'base64') : Buffer.from(decodeURIComponent(m[3]));
  } else if (/^https?:\/\//i.test(src)) {
    const res = await fetch(src, { redirect: 'follow' });
    if (!res.ok) throw new Error(`Could not download image: HTTP ${res.status}`);
    buf = Buffer.from(await res.arrayBuffer());
    mime = (res.headers.get('content-type') || '').split(';')[0] || MIME[extname(new URL(src).pathname).toLowerCase()];
  } else {
    const path = src.startsWith('~') ? home() + src.slice(1) : src;
    buf = await readFile(path);
    mime = MIME[extname(path).toLowerCase()];
  }
  if (!mime || !mime.startsWith('image/')) throw new Error('Unsupported image type (use PNG, JPEG, GIF, WebP, BMP or SVG)');
  if (buf.length > MAX_IMAGE_BYTES) throw new Error(`Image too large (${Math.round(buf.length / 1048576)} MB, max 15 MB)`);
  return { dataUrl: `data:${mime};base64,${buf.toString('base64')}`, size: imageSize(buf, mime) };
}

// Width/height in EMU from optional cm values, keeping the aspect ratio.
export function imageExtent(size, widthCm, heightCm, defaultWidthCm = 10) {
  const ratio = size && size.w && size.h ? size.h / size.w : 0.75;
  let w = widthCm, h = heightCm;
  if (!w && !h) w = defaultWidthCm;
  if (w && !h) h = w * ratio;
  if (h && !w) w = h / ratio;
  return { w: Math.round(w * EMU_PER_CM), h: Math.round(h * EMU_PER_CM) };
}

/**
 * Windows metafile (EMF and WMF) to SVG converter.
 *
 * Supports the drawing records used in practice by Office documents:
 * polygons / polylines / Béziers (32 and 16 bit), rectangles, ellipses,
 * arcs, chords, pies, paths, pens, brushes (solid, hatched, null), fonts and
 * text output, bitmaps (DIBs), clipping rectangles, window / viewport
 * mapping and world transforms. EMF+ records embedded in comments are
 * ignored; the EMF fallback records drawn alongside them are rendered.
 */
import { toBase64 } from './base64.js';
import { isSymbolFont, mapSymbolText } from '../pptx/symbols.js';

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const f2 = (n) => (Number.isFinite(n) ? Math.round(n * 100) / 100 : 0);

function rgb(c) {
  const r = c & 0xff, g = (c >> 8) & 0xff, b = (c >> 16) & 0xff;
  return '#' + [r, g, b].map((v) => v.toString(16).padStart(2, '0')).join('');
}

const HATCH = {
  0: 'M0 4H8', 1: 'M4 0V8', 2: 'M0 0L8 8', 3: 'M0 8L8 0', 4: 'M0 4H8M4 0V8', 5: 'M0 0L8 8M0 8L8 0',
};

// LOGFONT charset → WHATWG encoding used for ANSI text records.
const CHARSETS = {
  0: 'windows-1252', 1: 'windows-1252', 128: 'shift_jis', 129: 'euc-kr', 134: 'gbk', 136: 'big5', 161: 'windows-1253', 162: 'windows-1254',
  163: 'windows-1258', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250',
};
const decoders = {};

// Decode single/double-byte text; per-byte advances are merged per character.
// DEFAULT_CHARSET means the writer's ANSI code page; guess it from the face.
const FACE_CODEPAGES = [
  [/yahei|simsun|simhei|nsimsun|kaiti|fangsong|dengxian|宋体|黑体/i, 'gbk'],
  [/mingliu|jhenghei|細明體|新細明體/i, 'big5'],
  [/gothic|mincho|meiryo|ms ui gothic|ｍｓ/i, 'shift_jis'],
  [/malgun|gulim|batang|dotum|gungsuh|굴림|바탕/i, 'euc-kr'],
];

function decodeAnsi(buf, charset, dx = [], face = '') {
  let enc = CHARSETS[charset];
  if (charset === 1) for (const [re, cp] of FACE_CODEPAGES) if (re.test(face)) { enc = cp; break; }
  let dec = null;
  if (enc) {
    try { dec = decoders[enc] || (decoders[enc] = new TextDecoder(enc)); } catch { dec = null; }
  }
  if (!dec) return { str: Array.from(buf, (c) => String.fromCharCode(c)).join(''), dx };
  const str = dec.decode(buf);
  if (!dx.length || str.length === buf.length) return { str, dx };
  const lead = enc === 'shift_jis' ? (b) => (b >= 0x81 && b <= 0x9f) || (b >= 0xe0 && b <= 0xfc) : (b) => b >= 0x81 && b <= 0xfe;
  const merged = [];
  for (let i = 0; i < buf.length; i++) {
    if (lead(buf[i]) && i + 1 < buf.length) { merged.push((dx[i] || 0) + (dx[i + 1] || 0)); i++; } else merged.push(dx[i] || 0);
  }
  return { str, dx: merged.length === str.length ? merged : [] };
}

class DC {
  constructor() {
    this.wOrg = [0, 0];
    this.wExt = [1, 1];
    this.vOrg = [0, 0];
    this.vExt = [1, 1];
    this.mapMode = 1;
    this.xf = [1, 0, 0, 1, 0, 0];
    this.pen = { style: 0, width: 1, color: 0 };
    this.brush = { style: 0, color: 0xffffff };
    this.font = { height: -12, weight: 400, italic: false, underline: false, strike: false, charset: 0, face: 'Arial', escapement: 0 };
    this.textColor = 0;
    this.bkColor = 0xffffff;
    this.bkMode = 2;
    this.textAlign = 0;
    this.fillMode = 1;
    this.pos = [0, 0];
    this.clip = null;
    this.arcDir = 1;
  }

  clone() {
    const d = new DC();
    Object.assign(d, this);
    d.wOrg = [...this.wOrg]; d.wExt = [...this.wExt]; d.vOrg = [...this.vOrg]; d.vExt = [...this.vExt];
    d.xf = [...this.xf]; d.pos = [...this.pos];
    return d;
  }
}

class Renderer {
  constructor() {
    this.dc = new DC();
    this.stack = [];
    this.objects = [];
    this.out = [];
    this.defs = [];
    this.id = 0;
    this.path = null;
    this.clipId = null;
  }

  // Logical -> device coordinates.
  pt(x, y) {
    const d = this.dc;
    const m = d.xf;
    const wx = m[0] * x + m[2] * y + m[4];
    const wy = m[1] * x + m[3] * y + m[5];
    // WMF output space is the logical (window) space itself.
    if (this.identity) return [wx, wy];
    if (d.mapMode === 1) return [wx - d.wOrg[0] + d.vOrg[0], wy - d.wOrg[1] + d.vOrg[1]];
    const sx = d.wExt[0] ? d.vExt[0] / d.wExt[0] : 1;
    const sy = d.wExt[1] ? d.vExt[1] / d.wExt[1] : 1;
    return [(wx - d.wOrg[0]) * sx + d.vOrg[0], (wy - d.wOrg[1]) * sy + d.vOrg[1]];
  }

  scale() {
    const d = this.dc;
    const m = d.xf;
    let sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3]);
    if (this.identity) return [sx, sy];
    if (d.mapMode !== 1) {
      sx *= d.wExt[0] ? Math.abs(d.vExt[0] / d.wExt[0]) : 1;
      sy *= d.wExt[1] ? Math.abs(d.vExt[1] / d.wExt[1]) : 1;
    }
    return [sx, sy];
  }

  strokeAttr() {
    const p = this.dc.pen;
    if (!p || p.style === 5) return 'stroke="none"';
    const w = Math.max(p.width * this.scale()[0], 1);
    let a = `stroke="${rgb(p.color)}" stroke-width="${f2(w)}"`;
    const style = p.style & 0xf;
    if (style === 1) a += ` stroke-dasharray="${f2(w * 3)} ${f2(w)}"`;
    else if (style === 2) a += ` stroke-dasharray="${f2(w)} ${f2(w)}"`;
    else if (style === 3) a += ` stroke-dasharray="${f2(w * 3)} ${f2(w)} ${f2(w)} ${f2(w)}"`;
    else if (style === 4) a += ` stroke-dasharray="${f2(w * 3)} ${f2(w)} ${f2(w)} ${f2(w)} ${f2(w)} ${f2(w)}"`;
    const cap = (p.style >> 8) & 0xf;
    a += cap === 1 ? ' stroke-linecap="square"' : cap === 2 ? ' stroke-linecap="butt"' : ' stroke-linecap="round"';
    const join = (p.style >> 12) & 0xf;
    a += join === 1 ? ' stroke-linejoin="bevel"' : join === 2 ? ' stroke-linejoin="miter"' : ' stroke-linejoin="round"';
    return a;
  }

  fillAttr() {
    const b = this.dc.brush;
    const rule = this.dc.fillMode === 2 ? ' fill-rule="nonzero"' : ' fill-rule="evenodd"';
    if (!b || b.style === 1) return 'fill="none"';
    if (b.style === 2 && HATCH[b.hatch] !== undefined) {
      const id = 'h' + this.id++;
      const bg = this.dc.bkMode === 2 ? `<rect width="8" height="8" fill="${rgb(this.dc.bkColor)}"/>` : '';
      this.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="8" height="8">${bg}<path d="${HATCH[b.hatch]}" stroke="${rgb(b.color)}" stroke-width="1"/></pattern>`);
      return `fill="url(#${id})"${rule}`;
    }
    if (b.style === 3 && b.image) {
      const id = 'b' + this.id++;
      this.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${b.w}" height="${b.h}"><image href="${b.image}" width="${b.w}" height="${b.h}"/></pattern>`);
      return `fill="url(#${id})"${rule}`;
    }
    return `fill="${rgb(b.color)}"${rule}`;
  }

  emit(el) {
    this.out.push(this.clipId ? el.replace(/^<(\w+)/, `<$1 clip-path="url(#${this.clipId})"`) : el);
  }

  shape(d, fill = true, stroke = true) {
    if (!d) return;
    if (this.path) { this.path.push(d); return; }
    this.emit(`<path d="${d}" ${fill ? this.fillAttr() : 'fill="none"'} ${stroke ? this.strokeAttr() : 'stroke="none"'}/>`);
  }

  poly(pts, close, fill, stroke) {
    if (pts.length < 2) return;
    let d = 'M' + pts.map((p) => f2(p[0]) + ' ' + f2(p[1])).join('L');
    if (close) d += 'Z';
    this.shape(d, fill, stroke);
  }

  bezier(start, pts, close, fill, stroke, continueFrom) {
    let d = continueFrom ? '' : `M${f2(start[0])} ${f2(start[1])}`;
    for (let i = 0; i + 2 < pts.length; i += 3) d += `C${f2(pts[i][0])} ${f2(pts[i][1])} ${f2(pts[i + 1][0])} ${f2(pts[i + 1][1])} ${f2(pts[i + 2][0])} ${f2(pts[i + 2][1])}`;
    if (close) d += 'Z';
    return d;
  }

  // Fill a rectangle according to a ternary raster operation; only the
  // operations that do not depend on the destination are representable.
  ropRect(rop, l, t, r, b) {
    const op = (rop >>> 16) & 0xff;
    if (op === 0xaa) return; // D: destination unchanged
    const d = this.rectPath(l, t, r, b);
    if (op === 0x00 || op === 0xff) this.emit(`<path d="${d}" fill="${op ? '#ffffff' : '#000000'}" stroke="none"/>`);
    else if (op === 0xf0 || op === 0xcc) this.shape(d, true, false);
  }

  rectPath(l, t, r, b) {
    const p = [this.pt(l, t), this.pt(r, t), this.pt(r, b), this.pt(l, b)];
    return 'M' + p.map((q) => f2(q[0]) + ' ' + f2(q[1])).join('L') + 'Z';
  }

  ellipsePath(l, t, r, b) {
    const [x0, y0] = this.pt(l, t), [x1, y1] = this.pt(r, b);
    const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2, rx = Math.abs(x1 - x0) / 2, ry = Math.abs(y1 - y0) / 2;
    return `M${f2(cx - rx)} ${f2(cy)}A${f2(rx)} ${f2(ry)} 0 1 0 ${f2(cx + rx)} ${f2(cy)}A${f2(rx)} ${f2(ry)} 0 1 0 ${f2(cx - rx)} ${f2(cy)}Z`;
  }

  arcPath(l, t, r, b, xs, ys, xe, ye, kind) {
    const [x0, y0] = this.pt(l, t), [x1, y1] = this.pt(r, b);
    const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2, rx = Math.abs(x1 - x0) / 2, ry = Math.abs(y1 - y0) / 2;
    const [sx, sy] = this.pt(xs, ys), [ex, ey] = this.pt(xe, ye);
    let a0 = Math.atan2((sy - cy) / (ry || 1), (sx - cx) / (rx || 1));
    let a1 = Math.atan2((ey - cy) / (ry || 1), (ex - cx) / (rx || 1));
    // GDI arcs are drawn counter-clockwise (in device space y down => decreasing angle).
    let sweep = a0 - a1;
    if (sweep <= 0) sweep += Math.PI * 2;
    const P = (a) => [cx + rx * Math.cos(a), cy + ry * Math.sin(a)];
    const [px, py] = P(a0), [qx, qy] = P(a1);
    const large = sweep > Math.PI ? 1 : 0;
    let d = `M${f2(px)} ${f2(py)}A${f2(rx)} ${f2(ry)} 0 ${large} 0 ${f2(qx)} ${f2(qy)}`;
    if (kind === 'chord') d += 'Z';
    else if (kind === 'pie') d += `L${f2(cx)} ${f2(cy)}Z`;
    return { d, end: [qx, qy], start: [px, py] };
  }

  text(x, y, str, opts = {}) {
    if (!str) return;
    const d = this.dc;
    const f = d.font;
    // TA_UPDATECP: draw at, and advance, the current position.
    const cp = !opts.device && (d.textAlign & 1);
    if (cp) {
      [x, y] = d.pos;
      const adv = opts.dx && opts.dx.length ? opts.dx.reduce((a, b) => a + b, 0) : Math.abs(f.height) * 0.5 * str.length;
      d.pos = [x + adv, y];
    }
    let family = esc(f.face);
    if (isSymbolFont(f.face)) {
      str = mapSymbolText(str, f.face);
      family = /^symbol$/i.test(f.face) ? 'serif' : 'sans-serif';
    }
    const [px, py] = opts.device ? [x, y] : this.pt(x, y);
    const sy = this.scale()[1];
    const size = Math.max(1, Math.abs(f.height) * sy * (f.height > 0 ? 0.85 : 1));
    const align = d.textAlign;
    const anchor = (align & 6) === 6 ? 'middle' : (align & 2) ? 'end' : 'start';
    const baseline = (align & 24) === 24 ? 'alphabetic' : (align & 8) ? 'text-after-edge' : 'text-before-edge';
    let attrs = `x="${f2(px)}" y="${f2(py)}" font-family="${family}" font-size="${f2(size)}" fill="${rgb(d.textColor)}" text-anchor="${anchor}" dominant-baseline="${baseline}"`;
    if (f.weight >= 600) attrs += ' font-weight="bold"';
    if (f.italic) attrs += ' font-style="italic"';
    const deco = [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) attrs += ` text-decoration="${deco}"`;
    if (f.escapement) attrs += ` transform="rotate(${f2(-f.escapement / 10)} ${f2(px)} ${f2(py)})"`;
    if (opts.dx && opts.dx.length === [...str].length && opts.dx.some((v) => v)) {
      const xs = [];
      let cx = px;
      for (const v of opts.dx) { xs.push(f2(cx)); cx += v * this.scale()[0]; }
      attrs = attrs.replace(/^x="[^"]*"/, `x="${xs.join(' ')}"`);
    }
    this.emit(`<text ${attrs} xml:space="preserve">${esc(str)}</text>`);
  }

  image(dib, x, y, w, h) {
    const url = dibToDataUrl(dib);
    if (!url) return;
    const [x0, y0] = this.pt(x, y);
    const [x1, y1] = this.pt(x + w, y + h);
    const X = Math.min(x0, x1), Y = Math.min(y0, y1);
    let tr = '';
    if (x1 < x0 || y1 < y0) tr = ` transform="translate(${f2(x1 < x0 ? 2 * X + Math.abs(x1 - x0) : 0)} ${f2(y1 < y0 ? 2 * Y + Math.abs(y1 - y0) : 0)}) scale(${x1 < x0 ? -1 : 1} ${y1 < y0 ? -1 : 1})"`;
    this.emit(`<image href="${url}" x="${f2(X)}" y="${f2(Y)}" width="${f2(Math.abs(x1 - x0))}" height="${f2(Math.abs(y1 - y0))}" preserveAspectRatio="none"${tr}/>`);
  }

  setClipRect(l, t, r, b) {
    const id = 'c' + this.id++;
    this.defs.push(`<clipPath id="${id}"><path d="${this.rectPath(l, t, r, b)}"/></clipPath>`);
    this.clipId = id;
  }
}

/** Convert a packed DIB (BITMAPINFO + bits) into a BMP data URL. */
export function dibToDataUrl(dib) {
  if (!dib || dib.length < 40) return null;
  const v = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const hdr = v.getUint32(0, true);
  if (hdr < 12) return null;
  let bpp, clrUsed = 0, compression = 0;
  if (hdr === 12) bpp = v.getUint16(10, true);
  else { bpp = v.getUint16(14, true); compression = v.getUint32(16, true); clrUsed = v.getUint32(32, true); }
  // Embedded JPEG / PNG.
  if (compression === 4 || compression === 5) {
    const bits = dib.subarray(hdr);
    return `data:${compression === 4 ? 'image/jpeg' : 'image/png'};base64,${toBase64(bits)}`;
  }
  let palette = 0;
  if (bpp <= 8) palette = (clrUsed || 1 << bpp) * (hdr === 12 ? 3 : 4);
  if (compression === 3 && hdr === 40) palette = 12;
  const file = new Uint8Array(14 + dib.length);
  const fv = new DataView(file.buffer);
  file[0] = 0x42; file[1] = 0x4d;
  fv.setUint32(2, file.length, true);
  fv.setUint32(10, 14 + hdr + palette, true);
  file.set(dib, 14);
  return 'data:image/bmp;base64,' + toBase64(file);
}

function parseEmf(bytes) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const R = new Renderer();
  const n = bytes.length;
  let bounds = null;
  let p = 0;
  let ops = 0;
  const pts32 = (off, count) => {
    const out = [];
    for (let i = 0; i < count && off + i * 8 + 8 <= n; i++) out.push(R.pt(v.getInt32(off + i * 8, true), v.getInt32(off + i * 8 + 4, true)));
    return out;
  };
  const pts16 = (off, count) => {
    const out = [];
    for (let i = 0; i < count && off + i * 4 + 4 <= n; i++) out.push(R.pt(v.getInt16(off + i * 4, true), v.getInt16(off + i * 4 + 2, true)));
    return out;
  };
  const utf16 = (off, chars) => {
    let s = '';
    for (let i = 0; i < chars && off + i * 2 + 2 <= n; i++) s += String.fromCharCode(v.getUint16(off + i * 2, true));
    return s;
  };
  while (p + 8 <= n && ops++ < 500000) {
    const type = v.getUint32(p, true);
    const size = v.getUint32(p + 4, true);
    if (size < 8 || p + size > n) break;
    const d = p + 8;
    const dc = R.dc;
    switch (type) {
      case 1: { // HEADER
        const l = v.getInt32(d, true), t = v.getInt32(d + 4, true), r = v.getInt32(d + 8, true), b = v.getInt32(d + 12, true);
        const fl = v.getInt32(d + 16, true), ft = v.getInt32(d + 20, true), fr = v.getInt32(d + 24, true), fb = v.getInt32(d + 28, true);
        const devW = v.getInt32(d + 64, true), devH = v.getInt32(d + 68, true);
        const mmW = v.getInt32(d + 72, true), mmH = v.getInt32(d + 76, true);
        // Prefer the frame (0.01 mm) mapped to device pixels.
        if (mmW > 0 && mmH > 0 && fr > fl && fb > ft) {
          const ppmx = devW / mmW, ppmy = devH / mmH;
          bounds = { x: (fl / 100) * ppmx, y: (ft / 100) * ppmy, w: ((fr - fl) / 100) * ppmx, h: ((fb - ft) / 100) * ppmy };
        } else bounds = { x: l, y: t, w: r - l + 1, h: b - t + 1 };
        break;
      }
      case 14: p = n; continue; // EOF
      case 9: dc.wExt = [v.getInt32(d, true), v.getInt32(d + 4, true)]; break;
      case 10: dc.wOrg = [v.getInt32(d, true), v.getInt32(d + 4, true)]; break;
      case 11: dc.vExt = [v.getInt32(d, true), v.getInt32(d + 4, true)]; break;
      case 12: dc.vOrg = [v.getInt32(d, true), v.getInt32(d + 4, true)]; break;
      case 17: dc.mapMode = v.getUint32(d, true); break;
      case 18: dc.bkMode = v.getUint32(d, true); break;
      case 19: dc.fillMode = v.getUint32(d, true); break;
      case 22: dc.textAlign = v.getUint32(d, true); break;
      case 24: dc.textColor = v.getUint32(d, true); break;
      case 25: dc.bkColor = v.getUint32(d, true); break;
      case 27: dc.pos = [v.getInt32(d, true), v.getInt32(d + 4, true)]; if (R.path) R.path.push(`M${R.pt(...dc.pos).map(f2).join(' ')}`); break;
      case 30: R.setClipRect(v.getInt32(d, true), v.getInt32(d + 4, true), v.getInt32(d + 8, true), v.getInt32(d + 12, true)); break;
      case 33: R.stack.push({ dc: dc.clone(), clip: R.clipId }); break;
      case 34: {
        const rel = v.getInt32(d, true);
        let s = null;
        for (let k = 0; k < Math.max(1, -rel); k++) s = R.stack.pop() || s;
        if (s) { R.dc = s.dc; R.clipId = s.clip; }
        break;
      }
      case 35: dc.xf = [0, 4, 8, 12, 16, 20].map((o) => v.getFloat32(d + o, true)); break;
      case 36: {
        const m = [0, 4, 8, 12, 16, 20].map((o) => v.getFloat32(d + o, true));
        const mode = v.getUint32(d + 24, true);
        const a = dc.xf;
        const mul = (x, y) => [x[0] * y[0] + x[1] * y[2], x[0] * y[1] + x[1] * y[3], x[2] * y[0] + x[3] * y[2], x[2] * y[1] + x[3] * y[3], x[4] * y[0] + x[5] * y[2] + y[4], x[4] * y[1] + x[5] * y[3] + y[5]];
        if (mode === 1) dc.xf = [1, 0, 0, 1, 0, 0];
        else if (mode === 2) dc.xf = mul(m, a);
        else if (mode === 3) dc.xf = mul(a, m);
        else if (mode === 4) dc.xf = m;
        break;
      }
      case 37: { // SELECTOBJECT
        const ih = v.getUint32(d, true);
        if (ih & 0x80000000) {
          const s = ih & 0x7fffffff;
          if (s === 0) dc.brush = { style: 0, color: 0xffffff };
          else if (s === 1) dc.brush = { style: 0, color: 0xc0c0c0 };
          else if (s === 2) dc.brush = { style: 0, color: 0x808080 };
          else if (s === 3) dc.brush = { style: 0, color: 0x404040 };
          else if (s === 4) dc.brush = { style: 0, color: 0 };
          else if (s === 5) dc.brush = { style: 1 };
          else if (s === 6) dc.pen = { style: 0, width: 1, color: 0xffffff };
          else if (s === 7) dc.pen = { style: 0, width: 1, color: 0 };
          else if (s === 8) dc.pen = { style: 5 };
        } else {
          const o = R.objects[ih];
          if (o) {
            if (o.kind === 'pen') dc.pen = o;
            else if (o.kind === 'brush') dc.brush = o;
            else if (o.kind === 'font') dc.font = o;
          }
        }
        break;
      }
      case 38: { // CREATEPEN
        const ih = v.getUint32(d, true);
        R.objects[ih] = { kind: 'pen', style: v.getUint32(d + 4, true), width: v.getInt32(d + 8, true) || 1, color: v.getUint32(d + 16, true) };
        break;
      }
      case 95: { // EXTCREATEPEN
        const ih = v.getUint32(d, true);
        const off = d + 20;
        R.objects[ih] = { kind: 'pen', style: v.getUint32(off, true) & 0xffff | (v.getUint32(off, true) & 0xf000), width: v.getUint32(off + 4, true) || 1, color: v.getUint32(off + 12, true) };
        if (v.getUint32(off + 8, true) === 1) R.objects[ih].style = 5;
        break;
      }
      case 39: { // CREATEBRUSHINDIRECT
        const ih = v.getUint32(d, true);
        R.objects[ih] = { kind: 'brush', style: v.getUint32(d + 4, true), color: v.getUint32(d + 8, true), hatch: v.getUint32(d + 12, true) };
        break;
      }
      case 93: case 94: { // CREATEMONOBRUSH / CREATEDIBPATTERNBRUSHPT
        const ih = v.getUint32(d, true);
        const offBmi = v.getUint32(d + 8, true), cbBmi = v.getUint32(d + 12, true), offBits = v.getUint32(d + 16, true), cbBits = v.getUint32(d + 20, true);
        const dib = new Uint8Array(cbBmi + cbBits);
        dib.set(bytes.subarray(p + offBmi, p + offBmi + cbBmi), 0);
        dib.set(bytes.subarray(p + offBits, p + offBits + cbBits), cbBmi);
        const w = Math.abs(v.getInt32(p + offBmi + 4, true)) || 8, h = Math.abs(v.getInt32(p + offBmi + 8, true)) || 8;
        R.objects[ih] = { kind: 'brush', style: 3, image: dibToDataUrl(dib), w, h, color: 0 };
        break;
      }
      case 82: { // EXTCREATEFONTINDIRECTW
        const ih = v.getUint32(d, true);
        const lf = d + 4;
        R.objects[ih] = {
          kind: 'font', height: v.getInt32(lf, true), escapement: v.getInt32(lf + 8, true), weight: v.getInt32(lf + 16, true),
          italic: bytes[lf + 20] !== 0, underline: bytes[lf + 21] !== 0, strike: bytes[lf + 22] !== 0, charset: bytes[lf + 23],
          face: utf16(lf + 28, 32).replace(/\0.*$/, '') || 'Arial',
        };
        break;
      }
      case 40: delete R.objects[v.getUint32(d, true)]; break;
      case 59: R.path = []; break;
      case 60: break;
      case 61: if (R.path) R.path.push('Z'); break;
      case 68: R.path = null; break;
      case 62: case 63: case 64: { // FILL/STROKEANDFILL/STROKE PATH
        const pathD = R.path ? R.path.join('') : '';
        R.path = null;
        if (pathD) R.shape(pathD, type !== 64, type !== 62);
        break;
      }
      case 67: { // SELECTCLIPPATH
        if (R.path && R.path.length) {
          const id = 'c' + R.id++;
          R.defs.push(`<clipPath id="${id}"><path d="${R.path.join('')}"/></clipPath>`);
          R.clipId = id;
        }
        R.path = null;
        break;
      }
      case 75: { // EXTSELECTCLIPRGN
        const mode = v.getUint32(d + 4, true);
        if (mode === 5 || v.getUint32(d, true) === 0) R.clipId = null;
        break;
      }
      case 2: case 3: case 4: case 5: case 6: case 85: case 86: case 87: case 88: case 89: {
        const is16 = type >= 85;
        const count = v.getUint32(d + 16, true);
        const pts = is16 ? pts16(d + 20, count) : pts32(d + 20, count);
        const kind = is16 ? type - 83 : type;
        if (kind === 3) R.poly(pts, true, true, true);
        else if (kind === 4) R.poly(pts, false, false, true);
        else if (kind === 2) R.shape(R.bezier(pts[0], pts.slice(1), false), false, true);
        else if (kind === 5 || kind === 6) {
          const start = R.pt(...dc.pos);
          let pd;
          if (kind === 5) pd = R.bezier(start, pts, false, false, true, !!R.path);
          else pd = (R.path ? '' : `M${f2(start[0])} ${f2(start[1])}`) + pts.map((q) => `L${f2(q[0])} ${f2(q[1])}`).join('');
          if (R.path) R.path.push(pd); else R.shape(pd, false, true);
          const last = is16 ? [v.getInt16(d + 20 + (count - 1) * 4, true), v.getInt16(d + 22 + (count - 1) * 4, true)] : [v.getInt32(d + 20 + (count - 1) * 8, true), v.getInt32(d + 24 + (count - 1) * 8, true)];
          if (count) dc.pos = last;
        }
        break;
      }
      case 7: case 8: case 90: case 91: { // POLYPOLYLINE / POLYPOLYGON (32/16)
        const is16 = type >= 90;
        const polys = v.getUint32(d + 16, true);
        let off = d + 24 + polys * 4;
        let pd = '';
        for (let k = 0; k < polys; k++) {
          const c = v.getUint32(d + 24 + k * 4, true);
          const pts = is16 ? pts16(off, c) : pts32(off, c);
          off += c * (is16 ? 4 : 8);
          if (pts.length) pd += 'M' + pts.map((q) => f2(q[0]) + ' ' + f2(q[1])).join('L') + (type === 8 || type === 91 ? 'Z' : '');
        }
        const closed = type === 8 || type === 91;
        R.shape(pd, closed, true);
        break;
      }
      case 54: { // LINETO
        const x = v.getInt32(d, true), y = v.getInt32(d + 4, true);
        const a = R.pt(...dc.pos), b = R.pt(x, y);
        if (R.path) R.path.push(`L${f2(b[0])} ${f2(b[1])}`);
        else R.emit(`<path d="M${f2(a[0])} ${f2(a[1])}L${f2(b[0])} ${f2(b[1])}" fill="none" ${R.strokeAttr()}/>`);
        dc.pos = [x, y];
        break;
      }
      case 42: R.shape(R.ellipsePath(v.getInt32(d, true), v.getInt32(d + 4, true), v.getInt32(d + 8, true), v.getInt32(d + 12, true))); break;
      case 43: R.shape(R.rectPath(v.getInt32(d, true), v.getInt32(d + 4, true), v.getInt32(d + 8, true), v.getInt32(d + 12, true))); break;
      case 44: { // ROUNDRECT
        const l = v.getInt32(d, true), t = v.getInt32(d + 4, true), r = v.getInt32(d + 8, true), b = v.getInt32(d + 12, true);
        const cw = v.getInt32(d + 16, true) / 2, ch = v.getInt32(d + 20, true) / 2;
        const [x0, y0] = R.pt(l, t), [x1, y1] = R.pt(r, b);
        const [sx, sy] = R.scale();
        const rx = Math.min(Math.abs(cw * sx), Math.abs(x1 - x0) / 2), ry = Math.min(Math.abs(ch * sy), Math.abs(y1 - y0) / 2);
        const X = Math.min(x0, x1), Y = Math.min(y0, y1), W = Math.abs(x1 - x0), H = Math.abs(y1 - y0);
        R.shape(`M${f2(X + rx)} ${f2(Y)}H${f2(X + W - rx)}A${f2(rx)} ${f2(ry)} 0 0 1 ${f2(X + W)} ${f2(Y + ry)}V${f2(Y + H - ry)}A${f2(rx)} ${f2(ry)} 0 0 1 ${f2(X + W - rx)} ${f2(Y + H)}H${f2(X + rx)}A${f2(rx)} ${f2(ry)} 0 0 1 ${f2(X)} ${f2(Y + H - ry)}V${f2(Y + ry)}A${f2(rx)} ${f2(ry)} 0 0 1 ${f2(X + rx)} ${f2(Y)}Z`);
        break;
      }
      case 45: case 46: case 47: case 55: { // ARC / CHORD / PIE / ARCTO
        const a = [0, 4, 8, 12, 16, 20, 24, 28].map((o) => v.getInt32(d + o, true));
        const kind = type === 46 ? 'chord' : type === 47 ? 'pie' : 'arc';
        const res = R.arcPath(...a, kind);
        if (type === 55) {
          const cur = R.pt(...dc.pos);
          const pd = `M${f2(cur[0])} ${f2(cur[1])}L${f2(res.start[0])} ${f2(res.start[1])}` + res.d.replace(/^M[^A]*/, '');
          if (R.path) R.path.push(pd.replace(/^M[^L]*/, '')); else R.shape(pd, false, true);
        } else R.shape(res.d, kind !== 'arc', true);
        break;
      }
      case 83: case 84: { // EXTTEXTOUTA / W
        const emr = d + 28;
        const x = v.getInt32(emr, true), y = v.getInt32(emr + 4, true);
        const chars = v.getUint32(emr + 8, true);
        const offStr = v.getUint32(emr + 12, true);
        const offDx = v.getUint32(emr + 36, true);
        let dx = [];
        if (offDx) for (let i = 0; i < chars && p + offDx + i * 4 + 4 <= n; i++) dx.push(v.getInt32(p + offDx + i * 4, true));
        let str;
        if (type === 84) str = utf16(p + offStr, chars);
        else ({ str, dx } = decodeAnsi(bytes.subarray(p + offStr, Math.min(p + offStr + chars, n)), dc.font.charset, dx, dc.font.face));
        R.text(x, y, str.replace(/\0+$/, ''), { dx });
        break;
      }
      case 81: case 77: case 76: case 80: case 114: case 116: { // STRETCHDIBITS / STRETCHBLT / BITBLT / SETDIBITSTODEVICE / ALPHABLEND / TRANSPARENTBLT
        let xDest, yDest, cxDest, cyDest, offBmi, cbBmi, offBits, cbBits;
        if (type === 81) {
          xDest = v.getInt32(d + 16, true); yDest = v.getInt32(d + 20, true);
          offBmi = v.getUint32(d + 40, true); cbBmi = v.getUint32(d + 44, true); offBits = v.getUint32(d + 48, true); cbBits = v.getUint32(d + 52, true);
          cxDest = v.getInt32(d + 64, true); cyDest = v.getInt32(d + 68, true);
        } else if (type === 80) {
          xDest = v.getInt32(d + 16, true); yDest = v.getInt32(d + 20, true);
          cxDest = v.getInt32(d + 32, true); cyDest = v.getInt32(d + 36, true);
          offBmi = v.getUint32(d + 40, true); cbBmi = v.getUint32(d + 44, true); offBits = v.getUint32(d + 48, true); cbBits = v.getUint32(d + 52, true);
        } else {
          xDest = v.getInt32(d + 16, true); yDest = v.getInt32(d + 20, true); cxDest = v.getInt32(d + 24, true); cyDest = v.getInt32(d + 28, true);
          // BITBLT, STRETCHBLT, ALPHABLEND, TRANSPARENTBLT share the layout:
          // rop/blend, xSrc, ySrc, xformSrc, bkColorSrc, usageSrc, offBmiSrc...
          const ob = d + 76;
          offBmi = v.getUint32(ob, true); cbBmi = v.getUint32(ob + 4, true); offBits = v.getUint32(ob + 8, true); cbBits = v.getUint32(ob + 12, true);
        }
        const rop = type === 81 ? v.getUint32(d + 60, true) : type === 76 || type === 77 ? v.getUint32(d + 32, true) : 0xcc0020;
        if (((rop >>> 16) & 0xff) === 0xaa) break;
        if (cbBmi && offBmi && p + offBmi + cbBmi <= n && p + offBits + cbBits <= n) {
          const dib = new Uint8Array(cbBmi + cbBits);
          dib.set(bytes.subarray(p + offBmi, p + offBmi + cbBmi), 0);
          dib.set(bytes.subarray(p + offBits, p + offBits + cbBits), cbBmi);
          R.image(dib, xDest, yDest, cxDest, cyDest);
        } else if (type === 76 && cxDest && cyDest) {
          R.ropRect(v.getUint32(d + 32, true), xDest, yDest, xDest + cxDest, yDest + cyDest);
        }
        break;
      }
      case 118: { // GRADIENTFILL (rectangles only)
        const nVer = v.getUint32(d + 16, true), nTri = v.getUint32(d + 20, true), mode = v.getUint32(d + 24, true);
        const verts = [];
        for (let i = 0; i < nVer; i++) {
          const o = d + 28 + i * 16;
          verts.push({ x: v.getInt32(o, true), y: v.getInt32(o + 4, true), c: '#' + [8, 10, 12].map((k) => (v.getUint16(o + k, true) >> 8).toString(16).padStart(2, '0')).join('') });
        }
        if (mode !== 2) {
          for (let i = 0; i < nTri; i++) {
            const o = d + 28 + nVer * 16 + i * 8;
            const a = verts[v.getUint32(o, true)], b = verts[v.getUint32(o + 4, true)];
            if (!a || !b) continue;
            const id = 'g' + R.id++;
            const [x0, y0] = R.pt(a.x, a.y), [x1, y1] = R.pt(b.x, b.y);
            R.defs.push(`<linearGradient id="${id}" x1="0" y1="0" x2="${mode === 0 ? 1 : 0}" y2="${mode === 0 ? 0 : 1}"><stop offset="0" stop-color="${a.c}"/><stop offset="1" stop-color="${b.c}"/></linearGradient>`);
            R.emit(`<rect x="${f2(Math.min(x0, x1))}" y="${f2(Math.min(y0, y1))}" width="${f2(Math.abs(x1 - x0))}" height="${f2(Math.abs(y1 - y0))}" fill="url(#${id})"/>`);
          }
        }
        break;
      }
      default:
        break;
    }
    p += size;
  }
  if (!bounds) return null;
  return { R, bounds };
}

function parseWmf(bytes) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const n = bytes.length;
  let p = 0;
  let bounds = null;
  let inch = 1440;
  if (v.getUint32(0, true) === 0x9ac6cdd7) {
    const l = v.getInt16(6, true), t = v.getInt16(8, true), r = v.getInt16(10, true), b = v.getInt16(12, true);
    inch = v.getUint16(14, true) || 1440;
    bounds = { l, t, r, b };
    p = 22;
  }
  const headerSize = v.getUint16(p + 2, true) * 2;
  p += headerSize;
  const R = new Renderer();
  R.identity = true;
  let firstWindow = null;
  const embedded = { chunks: [], size: 0 };
  const objects = [];
  const addObject = (o) => {
    let i = 0;
    while (objects[i]) i++;
    objects[i] = o;
  };
  const pts16 = (off, count) => {
    const out = [];
    for (let i = 0; i < count && off + i * 4 + 4 <= n; i++) out.push(R.pt(v.getInt16(off + i * 4, true), v.getInt16(off + i * 4 + 2, true)));
    return out;
  };
  let ops = 0;
  while (p + 6 <= n && ops++ < 500000) {
    const size = v.getUint32(p, true) * 2;
    const fn = v.getUint16(p + 4, true);
    if (size < 6 || p + size > n) break;
    const d = p + 6;
    const s16 = (k) => v.getInt16(d + k * 2, true);
    const dc = R.dc;
    if (!firstWindow && R.out.length) firstWindow = { org: [...dc.wOrg], ext: [...dc.wExt] };
    switch (fn) {
      case 0x0000: p = n; continue;
      case 0x061d: { // PATBLT with the current brush
        const h = s16(2), w = s16(3), y = s16(4), x = s16(5);
        R.ropRect(v.getUint32(d, true), x, y, x + w, y + h);
        break;
      }
      case 0x020b: dc.wOrg = [s16(1), s16(0)]; break;
      case 0x020c: dc.wExt = [s16(1), s16(0)]; break;
      case 0x020d: dc.vOrg = [s16(1), s16(0)]; break;
      case 0x020e: dc.vExt = [s16(1), s16(0)]; break;
      case 0x0103: break; // SETMAPMODE (anisotropic mapping assumed)
      case 0x0102: dc.bkMode = s16(0); break;
      case 0x0106: dc.fillMode = s16(0); break;
      case 0x012e: dc.textAlign = v.getUint16(d, true); break;
      case 0x0209: dc.textColor = v.getUint32(d, true); break;
      case 0x0201: dc.bkColor = v.getUint32(d, true); break;
      case 0x001e: R.stack.push({ dc: dc.clone(), clip: R.clipId }); break;
      case 0x0127: { const s = R.stack.pop(); if (s) { R.dc = s.dc; R.clipId = s.clip; } break; }
      case 0x02fa: addObject({ kind: 'pen', style: v.getUint16(d, true), width: s16(1) || 1, color: v.getUint32(d + 6, true) }); break;
      case 0x02fc: addObject({ kind: 'brush', style: v.getUint16(d, true), color: v.getUint32(d + 2, true), hatch: v.getUint16(d + 6, true) }); break;
      case 0x02fb: {
        let end = d + 18;
        while (end < d + 50 && end < p + size && bytes[end]) end++;
        const charset = bytes[d + 13];
        const face = decodeAnsi(bytes.subarray(d + 18, end), charset).str;
        addObject({ kind: 'font', height: s16(0), escapement: s16(2), weight: s16(4), italic: bytes[d + 10] !== 0, underline: bytes[d + 11] !== 0, strike: bytes[d + 12] !== 0, charset, face: face || 'Arial' });
        break;
      }
      case 0x00f7: case 0x0142: case 0x01f9: addObject({ kind: 'other' }); break;
      case 0x012d: { const o = objects[v.getUint16(d, true)]; if (o) { if (o.kind === 'pen') dc.pen = o; else if (o.kind === 'brush') dc.brush = o; else if (o.kind === 'font') dc.font = o; } break; }
      case 0x01f0: objects[v.getUint16(d, true)] = undefined; break;
      case 0x0214: dc.pos = [s16(1), s16(0)]; break;
      case 0x0213: {
        const a = R.pt(...dc.pos), b = R.pt(s16(1), s16(0));
        R.emit(`<path d="M${f2(a[0])} ${f2(a[1])}L${f2(b[0])} ${f2(b[1])}" fill="none" ${R.strokeAttr()}/>`);
        dc.pos = [s16(1), s16(0)];
        break;
      }
      case 0x0324: R.poly(pts16(d + 2, s16(0)), true, true, true); break;
      case 0x0325: R.poly(pts16(d + 2, s16(0)), false, false, true); break;
      case 0x0538: {
        const polys = s16(0);
        let off = d + 2 + polys * 2;
        let pd = '';
        for (let k = 0; k < polys; k++) {
          const c = v.getUint16(d + 2 + k * 2, true);
          const pts = pts16(off, c);
          off += c * 4;
          if (pts.length) pd += 'M' + pts.map((q) => f2(q[0]) + ' ' + f2(q[1])).join('L') + 'Z';
        }
        R.shape(pd);
        break;
      }
      case 0x041b: R.shape(R.rectPath(s16(3), s16(2), s16(1), s16(0))); break;
      case 0x0418: R.shape(R.ellipsePath(s16(3), s16(2), s16(1), s16(0))); break;
      case 0x061c: {
        const [x0, y0] = R.pt(s16(5), s16(4)), [x1, y1] = R.pt(s16(3), s16(2));
        const [sx, sy] = R.scale();
        const rx = Math.abs(s16(1) * sx) / 2, ry = Math.abs(s16(0) * sy) / 2;
        R.emit(`<rect x="${f2(Math.min(x0, x1))}" y="${f2(Math.min(y0, y1))}" width="${f2(Math.abs(x1 - x0))}" height="${f2(Math.abs(y1 - y0))}" rx="${f2(rx)}" ry="${f2(ry)}" ${R.fillAttr()} ${R.strokeAttr()}/>`);
        break;
      }
      case 0x0817: case 0x081a: case 0x0830: {
        const kind = fn === 0x081a ? 'pie' : fn === 0x0830 ? 'chord' : 'arc';
        const res = R.arcPath(s16(7), s16(6), s16(5), s16(4), s16(3), s16(2), s16(1), s16(0), kind);
        R.shape(res.d, kind !== 'arc', true);
        break;
      }
      case 0x0521: { // TEXTOUT
        const len = s16(0);
        const { str } = decodeAnsi(bytes.subarray(d + 2, Math.min(d + 2 + len, p + size)), dc.font.charset, [], dc.font.face);
        const o = d + 2 + len + (len & 1);
        R.text(v.getInt16(o + 2, true), v.getInt16(o, true), str);
        break;
      }
      case 0x0a32: { // EXTTEXTOUT
        const y = s16(0), x = s16(1), len = s16(2), opts = v.getUint16(d + 6, true);
        let o = d + 8 + (opts & 6 ? 8 : 0);
        const raw = bytes.subarray(o, Math.min(o + len, p + size));
        o += len + (len & 1);
        const adv = [];
        for (let i = 0; i < len && o + i * 2 + 2 <= p + size; i++) adv.push(v.getInt16(o + i * 2, true));
        const { str, dx } = decodeAnsi(raw, dc.font.charset, adv, dc.font.face);
        R.text(x, y, str, { dx });
        break;
      }
      case 0x0f43: { // STRETCHDIB
        const dib = bytes.subarray(d + 22, p + size);
        R.image(dib, s16(10), s16(9), s16(8), s16(7));
        break;
      }
      case 0x0b41: { // DIBSTRETCHBLT
        const dib = bytes.subarray(d + 20, p + size);
        R.image(dib, s16(9), s16(8), s16(7), s16(6));
        break;
      }
      case 0x0940: { // DIBBITBLT
        const dib = bytes.subarray(d + 16, p + size);
        R.image(dib, s16(7), s16(6), s16(5), s16(4));
        break;
      }
      case 0x0416: R.setClipRect(s16(3), s16(2), s16(1), s16(0)); break;
      case 0x012c: if (!v.getUint16(d, true)) R.clipId = null; break; // SELECTCLIPREGION(null)
      case 0x0626: { // ESCAPE: MFCOMMENT chunks carrying an embedded EMF ("WMFC")
        if (v.getUint16(d, true) === 15 && size >= 44 && v.getUint32(d + 4, true) === 0x43464d57 && v.getUint32(d + 8, true) === 1) {
          const cur = v.getUint32(d + 26, true);
          embedded.size = v.getUint32(d + 34, true);
          embedded.chunks.push(bytes.subarray(d + 38, Math.min(d + 38 + cur, p + size)));
        }
        break;
      }
      default: break;
    }
    p += size;
  }
  if (embedded.chunks.length) {
    const total = embedded.chunks.reduce((a, c) => a + c.length, 0);
    if (total >= embedded.size && embedded.size > 88) {
      const emf = new Uint8Array(total);
      let o = 0;
      for (const c of embedded.chunks) { emf.set(c, o); o += c.length; }
      const res = parseEmf(emf.subarray(0, embedded.size));
      if (res && res.R.out.length) return res;
    }
  }
  let box;
  if (bounds) {
    const [x0, y0] = R.pt(bounds.l, bounds.t), [x1, y1] = R.pt(bounds.r, bounds.b);
    box = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) };
  } else {
    const win = firstWindow || { org: R.dc.wOrg, ext: R.dc.wExt };
    box = { x: win.org[0], y: win.org[1], w: Math.abs(win.ext[0]) || 100, h: Math.abs(win.ext[1]) || 100 };
  }
  void inch;
  return { R, bounds: box };
}

/**
 * Convert EMF / WMF bytes to an SVG document string, or null.
 * @param {Uint8Array} bytes
 * @param {string} mime image/x-emf or image/x-wmf
 */
export function emfToSvg(bytes, mime) {
  try {
    let res;
    const isEmf = bytes.length > 44 && bytes[0] === 1 && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46;
    if (isEmf || mime === 'image/x-emf') res = parseEmf(bytes);
    else res = parseWmf(bytes);
    if (!res || !res.bounds || !(res.bounds.w > 0) || !(res.bounds.h > 0)) return null;
    const { R, bounds: b } = res;
    if (!R.out.length) return null;
    return `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="${f2(b.x)} ${f2(b.y)} ${f2(b.w)} ${f2(b.h)}" width="${f2(b.w)}" height="${f2(b.h)}" preserveAspectRatio="none">${R.defs.length ? '<defs>' + R.defs.join('') + '</defs>' : ''}${R.out.join('')}</svg>`;
  } catch (e) {
    return null;
  }
}

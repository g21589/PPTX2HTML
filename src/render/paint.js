/**
 * Paint servers: converts fill models into SVG paints (gradients, patterns,
 * pictures, filters) and CSS backgrounds.
 */
import { css } from '../pptx/color.js';
import { r2, r4 } from '../util/units.js';

/** 8x8 bitmaps (one byte per row, MSB = left pixel) for preset patterns. */
const PATTERNS = {
  horz: [0xff, 0, 0, 0, 0, 0, 0, 0],
  vert: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  ltHorz: [0xff, 0, 0, 0, 0xff, 0, 0, 0],
  ltVert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88],
  dkHorz: [0xff, 0xff, 0, 0, 0xff, 0xff, 0, 0],
  dkVert: [0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc, 0xcc],
  narHorz: [0xff, 0, 0xff, 0, 0xff, 0, 0xff, 0],
  narVert: [0xaa, 0xaa, 0xaa, 0xaa, 0xaa, 0xaa, 0xaa, 0xaa],
  dashHorz: [0xf0, 0, 0, 0, 0x0f, 0, 0, 0],
  dashVert: [0x80, 0x80, 0x80, 0x80, 0x08, 0x08, 0x08, 0x08],
  cross: [0xff, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  dnDiag: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01],
  upDiag: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
  ltDnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11],
  ltUpDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
  dkDnDiag: [0xcc, 0x66, 0x33, 0x99, 0xcc, 0x66, 0x33, 0x99],
  dkUpDiag: [0x33, 0x66, 0xcc, 0x99, 0x33, 0x66, 0xcc, 0x99],
  wdDnDiag: [0xc1, 0xe0, 0x70, 0x38, 0x1c, 0x0e, 0x07, 0x83],
  wdUpDiag: [0x83, 0x07, 0x0e, 0x1c, 0x38, 0x70, 0xe0, 0xc1],
  dashDnDiag: [0x88, 0x44, 0x22, 0x11, 0, 0, 0, 0],
  dashUpDiag: [0x11, 0x22, 0x44, 0x88, 0, 0, 0, 0],
  diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81],
  smCheck: [0xcc, 0xcc, 0x33, 0x33, 0xcc, 0xcc, 0x33, 0x33],
  lgCheck: [0xf0, 0xf0, 0xf0, 0xf0, 0x0f, 0x0f, 0x0f, 0x0f],
  smGrid: [0xff, 0x88, 0x88, 0x88, 0xff, 0x88, 0x88, 0x88],
  lgGrid: [0xff, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  dotGrid: [0xaa, 0, 0x80, 0, 0x80, 0, 0x80, 0],
  smConfetti: [0x80, 0x04, 0x40, 0x02, 0x20, 0x01, 0x10, 0x08],
  lgConfetti: [0x8c, 0x0c, 0xc0, 0xc8, 0x31, 0x30, 0x03, 0x13],
  horzBrick: [0xff, 0x80, 0x80, 0x80, 0xff, 0x08, 0x08, 0x08],
  diagBrick: [0x80, 0x40, 0x20, 0x10, 0x18, 0x24, 0x42, 0x81],
  solidDmnd: [0x08, 0x1c, 0x3e, 0x7f, 0x3e, 0x1c, 0x08, 0x00],
  openDmnd: [0x80, 0x41, 0x22, 0x14, 0x08, 0x14, 0x22, 0x41],
  dotDmnd: [0x80, 0x00, 0x22, 0x00, 0x08, 0x00, 0x22, 0x00],
  plaid: [0xaa, 0x55, 0xaa, 0x55, 0xf0, 0xf0, 0xf0, 0xf0],
  sphere: [0x77, 0x89, 0x8f, 0x8f, 0x77, 0x98, 0xf8, 0xf8],
  weave: [0x88, 0x54, 0x22, 0x45, 0x88, 0x14, 0x22, 0x51],
  divot: [0x00, 0x20, 0x10, 0x20, 0x00, 0x04, 0x08, 0x04],
  shingle: [0x03, 0x84, 0x48, 0x30, 0x0c, 0x02, 0x01, 0x01],
  wave: [0x00, 0x18, 0xa4, 0x42, 0x00, 0x18, 0xa4, 0x42],
  trellis: [0xff, 0x66, 0xff, 0x99, 0xff, 0x66, 0xff, 0x99],
  zigZag: [0x81, 0x42, 0x24, 0x18, 0x81, 0x42, 0x24, 0x18],
};

const BAYER = [
  [0, 32, 8, 40, 2, 34, 10, 42], [48, 16, 56, 24, 50, 18, 58, 26], [12, 44, 4, 36, 14, 46, 6, 38], [60, 28, 52, 20, 62, 30, 54, 22],
  [3, 35, 11, 43, 1, 33, 9, 41], [51, 19, 59, 27, 49, 17, 57, 25], [15, 47, 7, 39, 13, 45, 5, 37], [63, 31, 55, 23, 61, 29, 53, 21],
];

/** Bitmap rows for a preset pattern name. */
export function patternBits(prst) {
  if (PATTERNS[prst]) return PATTERNS[prst];
  const m = /^pct(\d+)$/.exec(prst || '');
  const pct = m ? Number(m[1]) : 50;
  const rows = [];
  for (let y = 0; y < 8; y++) {
    let b = 0;
    for (let x = 0; x < 8; x++) if (BAYER[y][x] < (pct / 100) * 64) b |= 0x80 >> x;
    rows.push(b);
  }
  return rows;
}

export function patternSvgMarkup(fill) {
  const bits = patternBits(fill.prst);
  let d = '';
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) if (bits[y] & (0x80 >> x)) d += `M${x} ${y}h1v1h-1z`;
  }
  return `<rect width="8" height="8" fill="${css(fill.bg)}"/>` + (d ? `<path d="${d}" fill="${css(fill.fg)}"/>` : '');
}

function stopsMarkup(stops, reverse = false) {
  const list = reverse ? stops.map((s) => ({ pos: 1 - s.pos, color: s.color })).reverse() : stops;
  return list.map((s) => {
    const a = s.color.a === undefined ? 1 : s.color.a;
    return `<stop offset="${r4(Math.max(0, Math.min(1, s.pos)))}" stop-color="${css({ ...s.color, a: 1 })}"${a < 1 ? ` stop-opacity="${r4(a)}"` : ''}/>`;
  }).join('');
}

/**
 * Linear gradient vector for an angle in a w x h box.
 * scaled: angle defined in the unit square then stretched to the box.
 */
function linearVector(angle, w, h, scaled) {
  const t = (angle * Math.PI) / 180;
  const c = Math.cos(t), s = Math.sin(t);
  if (scaled) {
    const L = (Math.abs(c) + Math.abs(s)) / 2;
    return { x1: (0.5 - c * L) * w, y1: (0.5 - s * L) * h, x2: (0.5 + c * L) * w, y2: (0.5 + s * L) * h };
  }
  const L = (Math.abs(w * c) + Math.abs(h * s)) / 2;
  return { x1: w / 2 - c * L, y1: h / 2 - s * L, x2: w / 2 + c * L, y2: h / 2 + s * L };
}

/** Collected image filter for picture effects. */
function imageFilter(effects, rc) {
  if (!effects || !effects.length) return null;
  const id = rc.uid('if');
  let f = `<filter id="${id}" x="0" y="0" width="100%" height="100%" color-interpolation-filters="sRGB">`;
  let last = 'SourceGraphic';
  let n = 0;
  const step = () => 'r' + (n++);
  for (const e of effects) {
    const out = step();
    switch (e.type) {
      case 'grayscale':
        f += `<feColorMatrix in="${last}" type="saturate" values="0" result="${out}"/>`;
        break;
      case 'biLevel': {
        const g = step();
        f += `<feColorMatrix in="${last}" type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0" result="${g}"/>`;
        const steps = 50;
        const k = Math.round(e.thresh * steps);
        const tv = Array.from({ length: steps }, (_, i) => (i < k ? 0 : 1)).join(' ');
        f += `<feComponentTransfer in="${g}" result="${out}"><feFuncR type="discrete" tableValues="${tv}"/><feFuncG type="discrete" tableValues="${tv}"/><feFuncB type="discrete" tableValues="${tv}"/></feComponentTransfer>`;
        break;
      }
      case 'lum': {
        const c = Math.max(-0.99, Math.min(0.99, e.contrast));
        const slope = c >= 0 ? 1 / (1 - c) : 1 + c;
        const intercept = 0.5 - 0.5 * slope + e.bright;
        const fn = `type="linear" slope="${r4(slope)}" intercept="${r4(intercept)}"`;
        f += `<feComponentTransfer in="${last}" result="${out}"><feFuncR ${fn}/><feFuncG ${fn}/><feFuncB ${fn}/></feComponentTransfer>`;
        break;
      }
      case 'duotone': {
        const g = step();
        f += `<feColorMatrix in="${last}" type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0" result="${g}"/>`;
        const ch = (k) => `${r4(e.c1[k] / 255)} ${r4(e.c2[k] / 255)}`;
        f += `<feComponentTransfer in="${g}" result="${out}"><feFuncR type="table" tableValues="${ch('r')}"/><feFuncG type="table" tableValues="${ch('g')}"/><feFuncB type="table" tableValues="${ch('b')}"/></feComponentTransfer>`;
        break;
      }
      case 'clrChange': {
        // Key out pixels matching the "from" colour (with a small tolerance).
        const tol = 6;
        const table = (v) => Array.from({ length: 256 }, (_, i) => (Math.abs(i - v) <= tol ? 0 : 1)).join(' ');
        const m = step();
        f += `<feComponentTransfer in="${last}" result="${m}"><feFuncR type="discrete" tableValues="${table(Math.round(e.from.r))}"/><feFuncG type="discrete" tableValues="${table(Math.round(e.from.g))}"/><feFuncB type="discrete" tableValues="${table(Math.round(e.from.b))}"/><feFuncA type="identity"/></feComponentTransfer>`;
        const mask = step();
        const keep = e.to && e.to.a !== undefined ? e.to.a : 0;
        f += `<feColorMatrix in="${m}" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 ${1 - keep} ${1 - keep} ${1 - keep} 0 ${keep}" result="${mask}"/>`;
        f += `<feComposite in="${last}" in2="${mask}" operator="in" result="${out}"/>`;
        break;
      }
      case 'clrRepl': {
        const c = e.color;
        f += `<feColorMatrix in="${last}" type="matrix" values="0 0 0 0 ${r4(c.r / 255)} 0 0 0 0 ${r4(c.g / 255)} 0 0 0 0 ${r4(c.b / 255)} 0 0 0 1 0" result="${out}"/>`;
        break;
      }
      case 'hsl': {
        const g = step();
        f += `<feColorMatrix in="${last}" type="hueRotate" values="${r2(e.hue)}" result="${g}"/>`;
        f += `<feColorMatrix in="${g}" type="saturate" values="${r4(Math.max(0, 1 + e.sat))}" result="${out}"/>`;
        break;
      }
      case 'tint': {
        f += `<feColorMatrix in="${last}" type="saturate" values="${r4(Math.max(0, 1 - e.amt))}" result="${out}"/>`;
        break;
      }
      case 'blur':
        f += `<feGaussianBlur in="${last}" stdDeviation="${r2(e.rad / 2)}" result="${out}"/>`;
        break;
      case 'alphaInv':
        f += `<feColorMatrix in="${last}" type="matrix" values="1 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 -1 1" result="${out}"/>`;
        break;
      default:
        n--;
        continue;
    }
    last = out;
  }
  f += '</filter>';
  if (last === 'SourceGraphic') return null;
  return { id, markup: f };
}

/**
 * Picture (blip) placement inside a box.
 * Returns {x, y, w, h} of the full image rectangle, for stretch mode.
 */
export function blipRect(fill, w, h) {
  const fr = fill.stretch || { l: 0, t: 0, r: 0, b: 0 };
  const fx = fr.l * w, fy = fr.t * h;
  const fw = w * (1 - fr.l - fr.r), fh = h * (1 - fr.t - fr.b);
  const sr = fill.srcRect || { l: 0, t: 0, r: 0, b: 0 };
  const vw = 1 - sr.l - sr.r || 1;
  const vh = 1 - sr.t - sr.b || 1;
  const iw = fw / vw, ih = fh / vh;
  return { x: fx - sr.l * iw, y: fy - sr.t * ih, w: iw, h: ih, clip: { x: fx, y: fy, w: fw, h: fh } };
}

/** SVG <image> markup for a picture fill in a box. */
export function blipImageMarkup(fill, w, h, rc, defs) {
  const href = rc.url(fill.image);
  if (!href) return '';
  const filt = imageFilter(fill.effects, rc);
  if (filt) defs.push(filt.markup);
  const op = fill.alpha < 1 ? ` opacity="${r4(fill.alpha)}"` : '';
  const fa = filt ? ` filter="url(#${filt.id})"` : '';
  if (fill.tile && fill.image) {
    return null; // handled by pattern
  }
  const b = blipRect(fill, w, h);
  return `<image href="${href}" x="${r2(b.x)}" y="${r2(b.y)}" width="${r2(b.w)}" height="${r2(b.h)}" preserveAspectRatio="none"${op}${fa}/>`;
}

function tilePattern(fill, w, h, rc, defs) {
  const href = rc.url(fill.image);
  const img = fill.image || {};
  const t = fill.tile;
  const nw = img.width ? (img.width * 96) / (img.dpiX || 96) : 64;
  const nh = img.height ? (img.height * 96) / (img.dpiY || 96) : 64;
  const tw = Math.max(1, nw * t.sx);
  const th = Math.max(1, nh * t.sy);
  // Alignment anchor.
  const al = t.algn || 'tl';
  let ax = 0, ay = 0;
  if (/r$/.test(al)) ax = w - tw;
  else if (al === 't' || al === 'b' || al === 'ctr') ax = (w - tw) / 2;
  if (/^b/.test(al)) ay = h - th;
  else if (al === 'l' || al === 'r' || al === 'ctr') ay = (h - th) / 2;
  const fx = t.flip === 'x' || t.flip === 'xy';
  const fy = t.flip === 'y' || t.flip === 'xy';
  const pw = tw * (fx ? 2 : 1), ph = th * (fy ? 2 : 1);
  const id = rc.uid('pt');
  const filt = imageFilter(fill.effects, rc);
  if (filt) defs.push(filt.markup);
  const op = fill.alpha < 1 ? ` opacity="${r4(fill.alpha)}"` : '';
  const fa = filt ? ` filter="url(#${filt.id})"` : '';
  const im = (tr) => `<image href="${href}" width="${r2(tw)}" height="${r2(th)}" preserveAspectRatio="none"${op}${fa}${tr ? ` transform="${tr}"` : ''}/>`;
  let body = im('');
  if (fx) body += im(`translate(${r2(2 * tw)} 0) scale(-1 1)`);
  if (fy) body += im(`translate(0 ${r2(2 * th)}) scale(1 -1)`);
  if (fx && fy) body += im(`translate(${r2(2 * tw)} ${r2(2 * th)}) scale(-1 -1)`);
  defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" x="${r2(ax + t.tx)}" y="${r2(ay + t.ty)}" width="${r2(pw)}" height="${r2(ph)}">${body}</pattern>`);
  return `url(#${id})`;
}

/**
 * Convert a fill to an SVG paint.
 * @returns {{paint: string, opacity?: number}}
 * @param {object} o {rot, w, h} shape rotation for rotWithShape="0" handling
 */
export function svgPaint(fill, w, h, rc, defs, o = {}) {
  if (!fill || fill.type === 'none') return { paint: 'none' };
  switch (fill.type) {
    case 'solid': {
      const a = fill.color.a === undefined ? 1 : fill.color.a;
      return { paint: css({ ...fill.color, a: 1 }), opacity: a < 1 ? a : undefined };
    }
    case 'grad': {
      const id = rc.uid('g');
      if (fill.path) {
        const ftr = fill.fillToRect || { l: 0, t: 0, r: 0, b: 0 };
        const fx0 = ftr.l * w, fy0 = ftr.t * h, fx1 = (1 - ftr.r) * w, fy1 = (1 - ftr.b) * h;
        const cx = (fx0 + fx1) / 2, cy = (fy0 + fy1) / 2;
        if (fill.path === 'circle') {
          const r = Math.max(Math.hypot(cx, cy), Math.hypot(w - cx, cy), Math.hypot(cx, h - cy), Math.hypot(w - cx, h - cy)) || 1;
          defs.push(`<radialGradient id="${id}" gradientUnits="userSpaceOnUse" cx="${r2(cx)}" cy="${r2(cy)}" r="${r2(r)}">${stopsMarkup(fill.stops)}</radialGradient>`);
        } else {
          // rect / shape: elliptical approximation stretched to the box.
          const rx = Math.max(cx, w - cx) || 1, ry = Math.max(cy, h - cy) || 1;
          defs.push(`<radialGradient id="${id}" gradientUnits="userSpaceOnUse" cx="0" cy="0" r="1" gradientTransform="translate(${r2(cx)} ${r2(cy)}) scale(${r2(rx * 1.2)} ${r2(ry * 1.2)})">${stopsMarkup(fill.stops)}</radialGradient>`);
        }
        return { paint: `url(#${id})` };
      }
      let ang = fill.angle || 0;
      if (!fill.rotWithShape && o.rot) ang -= o.rot;
      const v = linearVector(ang, w, h, fill.scaled);
      defs.push(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${r2(v.x1)}" y1="${r2(v.y1)}" x2="${r2(v.x2)}" y2="${r2(v.y2)}">${stopsMarkup(fill.stops)}</linearGradient>`);
      return { paint: `url(#${id})` };
    }
    case 'patt': {
      const id = rc.uid('p');
      defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="8" height="8">${patternSvgMarkup(fill)}</pattern>`);
      return { paint: `url(#${id})` };
    }
    case 'blip': {
      if (!fill.image) return { paint: 'none' };
      if (fill.tile) return { paint: tilePattern(fill, w, h, rc, defs) };
      const id = rc.uid('b');
      const img = blipImageMarkup(fill, w, h, rc, defs);
      defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" x="0" y="0" width="${r2(Math.max(w, 0.01))}" height="${r2(Math.max(h, 0.01))}">${img}</pattern>`);
      return { paint: `url(#${id})` };
    }
    default:
      return { paint: 'none' };
  }
}

/** CSS background for a fill (used by text runs, table cells, backgrounds). */
export function cssBackground(fill, rc, w = 100, h = 100) {
  if (!fill || fill.type === 'none') return '';
  switch (fill.type) {
    case 'solid':
      return css(fill.color);
    case 'grad': {
      const stops = fill.stops.map((s) => `${css(s.color)} ${r2(s.pos * 100)}%`).join(',');
      if (fill.path) {
        const ftr = fill.fillToRect || { l: 0, t: 0, r: 0, b: 0 };
        const cx = ((ftr.l + 1 - ftr.r) / 2) * 100, cy = ((ftr.t + 1 - ftr.b) / 2) * 100;
        return `radial-gradient(${fill.path === 'circle' ? 'circle' : 'ellipse'} farthest-corner at ${r2(cx)}% ${r2(cy)}%,${stops})`;
      }
      return `linear-gradient(${r2((fill.angle || 0) + 90)}deg,${stops})`;
    }
    case 'patt': {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8">${patternSvgMarkup(fill)}</svg>`;
      return `url('data:image/svg+xml,${encodeURIComponent(svg).replace(/'/g, '%27')}')`;
    }
    case 'blip': {
      const u = rc.url(fill.image);
      if (!u) return '';
      if (fill.tile) return `url('${u}')`;
      return `url('${u}') center/100% 100% no-repeat`;
    }
    default:
      return '';
  }
}

/** A representative solid colour for a fill (for text colour fallbacks). */
export function fillColor(fill) {
  if (!fill) return null;
  if (fill.type === 'solid') return fill.color;
  if (fill.type === 'grad' && fill.stops.length) return fill.stops[0].color;
  if (fill.type === 'patt') return fill.fg;
  return null;
}

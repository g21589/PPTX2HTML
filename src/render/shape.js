/**
 * Shape rendering: geometry (SVG), outlines, arrow heads and effects.
 */
import { segsToSvg, pathEnds } from '../geometry/geometry.js';
import { css } from '../pptx/color.js';
import { r2, r4 } from '../util/units.js';
import { svgPaint, blipImageMarkup } from './paint.js';

const DASHES = {
  solid: null,
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

const CAPS = { flat: 'butt', rnd: 'round', sq: 'square' };

/** stroke-* attributes for a line model (paint must be resolved separately). */
export function strokeAttrs(line, paint) {
  let s = ` stroke="${paint.paint}" stroke-width="${r2(line.w)}"`;
  if (paint.opacity !== undefined) s += ` stroke-opacity="${r4(paint.opacity)}"`;
  s += ` stroke-linecap="${CAPS[line.cap] || 'butt'}" stroke-linejoin="${line.join === 'miter' ? 'miter' : line.join === 'bevel' ? 'bevel' : 'round'}"`;
  if (line.join === 'miter') s += ` stroke-miterlimit="${r2(Math.max(1, line.miterLimit))}"`;
  const pat = line.dash === 'custom' ? (line.custDash || []).flat() : DASHES[line.dash];
  if (pat && pat.length) {
    // Round/square caps extend each dash by half the width on both ends;
    // compensate so that dots stay dots.
    const capped = line.cap === 'rnd' || line.cap === 'sq';
    const arr = pat.map((v, i) => {
      const len = v * line.w;
      if (!capped) return Math.max(0.01, len);
      return i % 2 === 0 ? Math.max(0.01, len - line.w) : len + line.w;
    });
    s += ` stroke-dasharray="${arr.map(r2).join(' ')}"`;
  }
  return s;
}

const END_SIZE = { sm: 2, med: 3, lg: 5 };

/** Geometry of an arrow head with the tip at (0,0) pointing towards +x. */
function arrowShape(end, lw) {
  const base = Math.max(lw, 2.65);
  const W = END_SIZE[end.w] * base;
  const L = END_SIZE[end.len] * base;
  switch (end.type) {
    case 'triangle':
      return { kind: 'fill', d: `M0 0L${r2(-L)} ${r2(-W / 2)}L${r2(-L)} ${r2(W / 2)}Z`, back: L * 0.5 };
    case 'stealth':
      return { kind: 'fill', d: `M0 0L${r2(-L)} ${r2(-W / 2)}L${r2(-L * 0.6)} 0L${r2(-L)} ${r2(W / 2)}Z`, back: L * 0.5 };
    case 'diamond':
      return { kind: 'fill', d: `M${r2(L / 2)} 0L0 ${r2(-W / 2)}L${r2(-L / 2)} 0L0 ${r2(W / 2)}Z`, back: 0 };
    case 'oval':
      return { kind: 'fill', d: `M${r2(L / 2)} 0A${r2(L / 2)} ${r2(W / 2)} 0 1 0 ${r2(-L / 2)} 0A${r2(L / 2)} ${r2(W / 2)} 0 1 0 ${r2(L / 2)} 0Z`, back: 0 };
    case 'arrow':
      return { kind: 'stroke', d: `M${r2(-L)} ${r2(-W / 2)}L0 0L${r2(-L)} ${r2(W / 2)}`, back: Math.min(lw / 2, L * 0.3) };
    default:
      return null;
  }
}

/** Shorten the start or end of a path by `dist` along its tangent. */
function shorten(segs, atEnd, dist) {
  if (dist <= 0 || !segs.length) return segs;
  const out = segs.map((s) => ({ c: s.c, p: s.p.slice() }));
  if (atEnd) {
    let i = out.length - 1;
    while (i > 0 && out[i].c === 'Z') i--;
    const s = out[i];
    if (!s || s.c === 'M') return out;
    const p = s.p;
    const n = p.length;
    const ex = p[n - 2], ey = p[n - 1];
    let px, py;
    if (n >= 4) { px = p[n - 4]; py = p[n - 3]; } else {
      const prev = out[i - 1].p;
      px = prev[prev.length - 2]; py = prev[prev.length - 1];
    }
    const len = Math.hypot(ex - px, ey - py);
    if (len < 1e-6) return out;
    const k = Math.min(dist, len * 0.9) / len;
    const dx = (ex - px) * k, dy = (ey - py) * k;
    p[n - 2] -= dx; p[n - 1] -= dy;
    if (n >= 4 && s.c !== 'L') { p[n - 4] -= dx; p[n - 3] -= dy; }
  } else {
    const m = out[0];
    const s = out[1];
    if (!m || m.c !== 'M' || !s || s.c === 'Z') return out;
    const sx = m.p[0], sy = m.p[1];
    const nx = s.p[0], ny = s.p[1];
    const len = Math.hypot(nx - sx, ny - sy);
    if (len < 1e-6) return out;
    const k = Math.min(dist, len * 0.9) / len;
    const dx = (nx - sx) * k, dy = (ny - sy) * k;
    m.p[0] += dx; m.p[1] += dy;
    if (s.c !== 'L') { s.p[0] += dx; s.p[1] += dy; }
  }
  return out;
}

function arrowMarkup(end, pt, angle, line, paint) {
  const sh = arrowShape(end, line.w);
  if (!sh) return '';
  const tr = `translate(${r2(pt.x)} ${r2(pt.y)}) rotate(${r2((angle * 180) / Math.PI)})`;
  const op = paint.opacity !== undefined ? ` fill-opacity="${r4(paint.opacity)}" stroke-opacity="${r4(paint.opacity)}"` : '';
  if (sh.kind === 'fill') return `<path d="${sh.d}" transform="${tr}" fill="${paint.paint}" stroke="none"${op}/>`;
  return `<path d="${sh.d}" transform="${tr}" fill="none" stroke="${paint.paint}" stroke-width="${r2(line.w)}" stroke-linejoin="miter" stroke-linecap="round"${op}/>`;
}

const DARKEN = { darken: ['#000', 0.4], darkenLess: ['#000', 0.2], lighten: ['#fff', 0.4], lightenLess: ['#fff', 0.2] };

/** Outer shadow / glow / soft edge / blur / inner shadow filters. */
function effectFilters(fx, w, h, rc, defs, o) {
  const res = { before: '', filter: null, innerFilter: null, reflection: null };
  if (!fx) return res;
  const pad = 4 + (o.lineW || 0) + Math.max(
    fx.outerShadow ? fx.outerShadow.blur * 2 + fx.outerShadow.dist : 0,
    fx.glow ? fx.glow.rad * 2 : 0,
    fx.blur ? fx.blur.rad * 2 : 0,
  );
  const region = `filterUnits="userSpaceOnUse" x="${r2(-pad - w * 0.5)}" y="${r2(-pad - h * 0.5)}" width="${r2(w * 2 + 2 * pad)}" height="${r2(h * 2 + 2 * pad)}"`;
  // Main filter on the shape: soft edges, blur, glow, inner shadow.
  let f = '';
  let last = 'SourceGraphic';
  let alpha = 'SourceAlpha';
  if (fx.softEdge && fx.softEdge.rad > 0) {
    const r = fx.softEdge.rad;
    f += `<feMorphology in="SourceAlpha" operator="erode" radius="${r2(r / 2)}" result="se1"/><feGaussianBlur in="se1" stdDeviation="${r2(r / 2.5)}" result="se2"/><feComposite in="SourceGraphic" in2="se2" operator="in" result="se3"/>`;
    last = 'se3';
    alpha = 'se2';
  }
  if (fx.blur && fx.blur.rad > 0) {
    f += `<feGaussianBlur in="${last}" stdDeviation="${r2(fx.blur.rad / 2)}" result="bl"/>`;
    last = 'bl';
  }
  if (fx.innerShadow) {
    const s = fx.innerShadow;
    const off = rotateOffset(s.dx, s.dy, o);
    f += `<feComponentTransfer in="${alpha}" result="is1"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>`;
    f += `<feGaussianBlur in="is1" stdDeviation="${r2(s.blur / 2)}" result="is2"/><feOffset in="is2" dx="${r2(off.dx)}" dy="${r2(off.dy)}" result="is3"/>`;
    f += `<feFlood flood-color="${css({ ...s.color, a: 1 })}" flood-opacity="${r4(s.color.a)}" result="is4"/><feComposite in="is4" in2="is3" operator="in" result="is5"/><feComposite in="is5" in2="${alpha}" operator="in" result="is6"/>`;
    f += `<feMerge result="is7"><feMergeNode in="${last}"/><feMergeNode in="is6"/></feMerge>`;
    last = 'is7';
  }
  if (fx.glow && fx.glow.rad > 0) {
    const g = fx.glow;
    f += `<feMorphology in="${alpha}" operator="dilate" radius="${r2(g.rad / 3)}" result="gl1"/><feGaussianBlur in="gl1" stdDeviation="${r2(g.rad / 2.5)}" result="gl2"/>`;
    f += `<feFlood flood-color="${css({ ...g.color, a: 1 })}" flood-opacity="${r4(g.color.a)}" result="gl3"/><feComposite in="gl3" in2="gl2" operator="in" result="gl4"/>`;
    f += `<feMerge result="gl5"><feMergeNode in="gl4"/><feMergeNode in="${last}"/></feMerge>`;
    last = 'gl5';
  }
  if (f) {
    const id = rc.uid('fx');
    defs.push(`<filter id="${id}" ${region} color-interpolation-filters="sRGB">${f}</filter>`);
    res.filter = id;
  }
  // Outer shadow as a transformed, blurred copy behind the shape.
  if (fx.outerShadow && o.geomId) {
    const s = fx.outerShadow;
    const id = rc.uid('sh');
    const sb = Math.max(0, s.blur / 2);
    defs.push(`<filter id="${id}" ${region} color-interpolation-filters="sRGB"><feGaussianBlur in="SourceAlpha" stdDeviation="${r2(sb)}" result="a"/><feFlood flood-color="${css({ ...s.color, a: 1 })}" flood-opacity="${r4(s.color.a)}"/><feComposite in2="a" operator="in"/></filter>`);
    const off = s.rotWithShape === false ? rotateOffset(s.dx, s.dy, o) : rotateOffset(s.dx, s.dy, o);
    let tr = `translate(${r2(off.dx)} ${r2(off.dy)})`;
    if (s.sx !== 1 || s.sy !== 1 || s.kx || s.ky) {
      const [ax, ay] = anchorPoint(s.algn, w, h);
      const kx = Math.tan((s.kx * Math.PI) / 180);
      const ky = Math.tan((s.ky * Math.PI) / 180);
      tr += ` translate(${r2(ax)} ${r2(ay)}) matrix(${r4(s.sx)} ${r4(ky * s.sy)} ${r4(kx * s.sx)} ${r4(s.sy)} 0 0) translate(${r2(-ax)} ${r2(-ay)})`;
    }
    res.before = `<g class="pp-shadow" filter="url(#${id})" transform="${tr}"><use href="#${o.geomId}"/></g>`;
  }
  if (fx.reflection && o.geomId) res.reflection = fx.reflection;
  return res;
}

function anchorPoint(algn, w, h) {
  switch (algn) {
    case 'tl': return [0, 0];
    case 't': return [w / 2, 0];
    case 'tr': return [w, 0];
    case 'l': return [0, h / 2];
    case 'ctr': return [w / 2, h / 2];
    case 'r': return [w, h / 2];
    case 'bl': return [0, h];
    case 'br': return [w, h];
    case 'b':
    default: return [w / 2, h];
  }
}

/** Shadow offsets are page-relative; undo the shape's rotation / flips. */
function rotateOffset(dx, dy, o) {
  let x = dx, y = dy;
  if (o.rot) {
    const t = (-o.rot * Math.PI) / 180;
    const c = Math.cos(t), s = Math.sin(t);
    [x, y] = [x * c - y * s, x * s + y * c];
  }
  if (o.flipH) x = -x;
  if (o.flipV) y = -y;
  return { dx: x, dy: y };
}

function reflectionMarkup(rf, w, h, geomId, rc, defs) {
  const mid = rc.uid('rm');
  const gid = rc.uid('rg');
  const total = h;
  defs.push(`<linearGradient id="${gid}" gradientUnits="userSpaceOnUse" x1="0" y1="0" x2="0" y2="${r2(total)}"><stop offset="${r4(rf.stPos)}" stop-color="#fff" stop-opacity="${r4(rf.stA)}"/><stop offset="${r4(Math.max(rf.stPos + 0.001, rf.endPos))}" stop-color="#fff" stop-opacity="${r4(rf.endA)}"/></linearGradient>`);
  defs.push(`<mask id="${mid}" maskUnits="userSpaceOnUse" x="${r2(-w)}" y="${r2(-h)}" width="${r2(w * 3)}" height="${r2(h * 3)}"><rect x="${r2(-w)}" y="0" width="${r2(w * 3)}" height="${r2(h)}" fill="url(#${gid})"/></mask>`);
  let blur = '';
  if (rf.blur > 0) {
    const fid = rc.uid('rb');
    defs.push(`<filter id="${fid}" x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation="${r2(rf.blur / 2)}"/></filter>`);
    blur = ` filter="url(#${fid})"`;
  }
  const sy = Math.abs(rf.sy || 1);
  return `<g transform="translate(0 ${r2(h + rf.dist)})"><g mask="url(#${mid})"><g${blur} transform="scale(${r4(rf.sx || 1)} ${r4(-sy)}) translate(0 ${r2(-h)})"><use href="#${geomId}"/></g></g></g>`;
}

/**
 * SVG markup for an element's geometry.
 * @param {object} el element model (shape or picture)
 * @param {object} rc render context
 * @param {object} [o] {blip: picture fill for pictures}
 */
export function geometrySvg(el, rc, o = {}) {
  const w = Math.max(el.w, 0);
  const h = Math.max(el.h, 0);
  const geom = el.geom;
  if (!geom) return '';
  const defs = [];
  const rot = el.rot || 0;
  const fill = o.blip || el.fill;
  const line = el.line;
  const linePaint = line ? svgPaint(line.fill, w, h, rc, defs, { rot }) : null;
  const fillPaint = svgPaint(fill, w, h, rc, defs, { rot });
  const geomId = rc.uid('ge');
  let body = '';
  let strokes = '';
  const hasFill = fillPaint.paint !== 'none';
  for (const p of geom.paths) {
    let segs = p.segs;
    if (!segs.length) continue;
    const ends = line && p.stroke && (line.head || line.tail) ? pathEnds(segs) : null;
    let arrows = '';
    if (ends) {
      if (line.tail) {
        const sh = arrowShape(line.tail, line.w);
        if (sh) { segs = shorten(segs, true, sh.back); arrows += arrowMarkup(line.tail, ends.end, ends.end.angle, line, linePaint); }
      }
      if (line.head) {
        const sh = arrowShape(line.head, line.w);
        if (sh) { segs = shorten(segs, false, sh.back); arrows += arrowMarkup(line.head, ends.start, ends.start.angle, line, linePaint); }
      }
    }
    const d = segsToSvg(segs);
    const dFill = segsToSvg(p.segs);
    if (p.fill !== 'none' && hasFill) {
      const op = fillPaint.opacity !== undefined ? ` fill-opacity="${r4(fillPaint.opacity)}"` : '';
      body += `<path d="${dFill}" fill="${fillPaint.paint}"${op} stroke="none"/>`;
      const dk = DARKEN[p.fill];
      if (dk) body += `<path d="${dFill}" fill="${dk[0]}" fill-opacity="${dk[1]}" stroke="none"/>`;
    }
    if (line && p.stroke) strokes += compoundStroke(d, line, linePaint, rc, defs) + arrows;
  }
  // Pictures with an image but no geometry fill path (should not happen) are drawn as images.
  if (o.blip && !hasFill && o.blip.image) {
    body += blipImageMarkup(o.blip, w, h, rc, defs) || '';
  }
  if (!body && !strokes) return '';
  const fxr = effectFilters(el.effects, w, h, rc, defs, { geomId, rot, flipH: el.flipH, flipV: el.flipV, lineW: line ? line.w : 0 });
  let content = `<g id="${geomId}">${body}${strokes}</g>`;
  if (fxr.filter) content = `<g filter="url(#${fxr.filter})">${content}</g>`;
  let refl = '';
  if (fxr.reflection) refl = reflectionMarkup(fxr.reflection, w, h, geomId, rc, defs);
  const vb = `0 0 ${r2(Math.max(w, 0.01))} ${r2(Math.max(h, 0.01))}`;
  return `<svg class="pp-geom" width="${r2(Math.max(w, 0.01))}" height="${r2(Math.max(h, 0.01))}" viewBox="${vb}" overflow="visible" aria-hidden="true">${defs.length ? '<defs>' + defs.join('') + '</defs>' : ''}${refl}${fxr.before}${content}</svg>`;
}

function compoundStroke(d, line, paint, rc, defs) {
  const base = `<path d="${d}" fill="none"${strokeAttrs(line, paint)}/>`;
  if (!line.cmpd || line.cmpd === 'sng') return base;
  const id = rc.uid('cm');
  const w = line.w;
  let mask;
  const sw = (k) => `stroke-width="${r2(w * k)}" stroke-linecap="${CAPS[line.cap] || 'butt'}" stroke-linejoin="round"`;
  if (line.cmpd === 'tri') {
    mask = `<path d="${d}" fill="none" stroke="#fff" ${sw(1)}/><path d="${d}" fill="none" stroke="#000" ${sw(0.6)}/><path d="${d}" fill="none" stroke="#fff" ${sw(0.2)}/>`;
  } else {
    mask = `<path d="${d}" fill="none" stroke="#fff" ${sw(1)}/><path d="${d}" fill="none" stroke="#000" ${sw(1 / 3)}/>`;
  }
  defs.push(`<mask id="${id}" maskUnits="userSpaceOnUse" x="-10000" y="-10000" width="20000" height="20000">${mask}</mask>`);
  return `<g mask="url(#${id})">${base}</g>`;
}

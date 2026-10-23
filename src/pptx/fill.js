/**
 * Fill, line and effect property parsing.
 *
 * Results are plain objects ("models") with all colours resolved to RGBA and
 * all lengths converted to pixels, ready for rendering.
 */
import { colorFrom, colorChild } from './color.js';
import { emu2px, pct } from '../util/units.js';

export const FILL_NODES = new Set(['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']);

/** Find the fill element inside a properties container (spPr, bgPr, rPr, tcPr...). */
export function fillChild(container) {
  if (!container) return null;
  for (const c of container.children) if (FILL_NODES.has(c.name)) return c;
  return null;
}

function rectAttrs(n, def = 0) {
  if (!n) return null;
  return { l: pct(n.attrs.l, def), t: pct(n.attrs.t, def), r: pct(n.attrs.r, def), b: pct(n.attrs.b, def) };
}

/**
 * Parse a fill element.
 * @param {import('../util/xml.js').XNode} node a fill element (a:solidFill, ...)
 * @param {import('./context.js').Ctx} ctx
 */
export function parseFillNode(node, ctx) {
  if (!node) return null;
  switch (node.name) {
    case 'a:noFill':
      return { type: 'none' };
    case 'a:solidFill': {
      const color = colorFrom(node, ctx);
      return color ? { type: 'solid', color } : { type: 'none' };
    }
    case 'a:gradFill': {
      const stops = [];
      const gsLst = node.child('a:gsLst');
      for (const gs of gsLst ? gsLst.all('a:gs') : []) {
        const color = colorFrom(gs, ctx);
        if (color) stops.push({ pos: pct(gs.attrs.pos, 0), color });
      }
      stops.sort((a, b) => a.pos - b.pos);
      if (!stops.length) return { type: 'none' };
      const lin = node.child('a:lin');
      const path = node.child('a:path');
      const g = {
        type: 'grad',
        stops,
        rotWithShape: node.bool('rotWithShape', true),
        flip: node.attr('flip', 'none'),
        tileRect: rectAttrs(node.child('a:tileRect')),
      };
      if (path) {
        g.path = path.attr('path', 'circle');
        g.fillToRect = rectAttrs(path.child('a:fillToRect')) || { l: 0, t: 0, r: 0, b: 0 };
      } else {
        g.angle = lin ? lin.num('ang', 0) / 60000 : 90;
        g.scaled = lin ? lin.bool('scaled', false) : false;
      }
      return g;
    }
    case 'a:pattFill': {
      const fg = colorFrom(node.child('a:fgClr'), ctx) || { r: 0, g: 0, b: 0, a: 1 };
      const bg = colorFrom(node.child('a:bgClr'), ctx) || { r: 255, g: 255, b: 255, a: 1 };
      return { type: 'patt', prst: node.attr('prst', 'pct5'), fg, bg };
    }
    case 'a:blipFill':
      return parseBlipFill(node, ctx);
    case 'a:grpFill':
      return ctx.groupFill ? ctx.groupFill : { type: 'none' };
    default:
      return null;
  }
}

/** Parse a (a:|p:)blipFill element into a picture fill. */
export function parseBlipFill(node, ctx) {
  const blip = node.child('a:blip');
  const fill = {
    type: 'blip',
    image: null,
    srcRect: rectAttrs(node.child('a:srcRect')),
    rotWithShape: node.bool('rotWithShape', true),
    dpi: node.int('dpi', 0),
    alpha: 1,
    effects: [],
  };
  const stretch = node.child('a:stretch');
  const tile = node.child('a:tile');
  if (tile) {
    fill.tile = {
      tx: emu2px(tile.num('tx', 0)), ty: emu2px(tile.num('ty', 0)),
      sx: pct(tile.attrs.sx, 1), sy: pct(tile.attrs.sy, 1),
      flip: tile.attr('flip', 'none'), algn: tile.attr('algn', 'tl'),
    };
  } else {
    fill.stretch = rectAttrs(stretch && stretch.child('a:fillRect')) || { l: 0, t: 0, r: 0, b: 0 };
  }
  if (blip) {
    // Prefer an SVG version when present.
    let svgId = null;
    for (const e of blip.find('asvg:svgBlip')) svgId = e.attrs['r:embed'] || svgId;
    const embed = blip.attrs['r:embed'];
    const link = blip.attrs['r:link'];
    if (svgId) fill.image = ctx.image(svgId);
    if (!fill.image && embed) fill.image = ctx.image(embed);
    if (!fill.image && link) fill.image = ctx.image(link);
    if (svgId && embed) fill.fallback = ctx.image(embed);
    for (const e of blip.children) {
      switch (e.local) {
        case 'alphaModFix': fill.alpha *= pct(e.attrs.amt, 1); break;
        case 'grayscl': fill.effects.push({ type: 'grayscale' }); break;
        case 'biLevel': fill.effects.push({ type: 'biLevel', thresh: pct(e.attrs.thresh, 0.5) }); break;
        case 'lum': fill.effects.push({ type: 'lum', bright: pct(e.attrs.bright, 0), contrast: pct(e.attrs.contrast, 0) }); break;
        case 'duotone': {
          const cols = e.children.filter((x) => x.name.startsWith('a:') && x.name.endsWith('Clr'));
          if (cols.length >= 2) {
            const c1 = colorFrom({ children: [cols[0]] }, ctx);
            const c2 = colorFrom({ children: [cols[1]] }, ctx);
            fill.effects.push({ type: 'duotone', c1, c2 });
          }
          break;
        }
        case 'clrChange': {
          const from = colorFrom(e.child('a:clrFrom'), ctx);
          const to = colorFrom(e.child('a:clrTo'), ctx);
          if (from && to) fill.effects.push({ type: 'clrChange', from, to });
          break;
        }
        case 'clrRepl': {
          const c = colorFrom(e, ctx);
          if (c) fill.effects.push({ type: 'clrRepl', color: c });
          break;
        }
        case 'hsl': fill.effects.push({ type: 'hsl', hue: e.num('hue', 0) / 60000, sat: pct(e.attrs.sat, 0), lum: pct(e.attrs.lum, 0) }); break;
        case 'tint': fill.effects.push({ type: 'tint', hue: e.num('hue', 0) / 60000, amt: pct(e.attrs.amt, 0) }); break;
        case 'alphaInv': fill.effects.push({ type: 'alphaInv' }); break;
        case 'blur': fill.effects.push({ type: 'blur', rad: emu2px(e.num('rad', 0)) }); break;
        case 'alphaBiLevel': fill.effects.push({ type: 'alphaBiLevel', thresh: pct(e.attrs.thresh, 0.5) }); break;
        case 'alphaFloor': case 'alphaCeiling': fill.effects.push({ type: e.local }); break;
        case 'alphaRepl': fill.alpha = pct(e.attrs.a, 1); break;
        case 'alphaMod': {
          const cont = e.child('a:cont');
          const amt = cont && cont.first('a:alphaModFix');
          if (amt) fill.alpha *= pct(amt.attrs.amt, 1);
          break;
        }
        case 'fillOverlay': {
          const f = parseFillNode(fillChild(e), ctx);
          if (f) fill.effects.push({ type: 'fillOverlay', blend: e.attr('blend', 'over'), fill: f });
          break;
        }
        case 'extLst': {
          // Artistic effects and photo editing are pre-rendered by PowerPoint.
          const art = e.first('a14:imgProps');
          if (art) ctx.note('approx', 'artistic image effects');
          break;
        }
        default: break;
      }
    }
  }
  if (!fill.image) ctx.note('missing', 'picture fill without image data');
  return fill;
}

/** Parse the fill of a properties container; returns undefined when not specified. */
export function parseFill(container, ctx) {
  const n = fillChild(container);
  if (!n) return undefined;
  return parseFillNode(n, ctx);
}

const SIZE = { sm: 'sm', med: 'med', lg: 'lg' };

function parseEnd(n) {
  if (!n) return undefined;
  const type = n.attr('type', 'none');
  return { type, w: SIZE[n.attrs.w] || 'med', len: SIZE[n.attrs.len] || 'med' };
}

/**
 * Parse an a:ln element into a partial line model (only specified fields).
 */
export function parseLine(ln, ctx) {
  if (!ln) return null;
  const out = {};
  if (ln.attrs.w !== undefined) out.w = emu2px(ln.num('w', 0));
  if (ln.attrs.cap) out.cap = ln.attrs.cap;
  if (ln.attrs.cmpd) out.cmpd = ln.attrs.cmpd;
  if (ln.attrs.algn) out.algn = ln.attrs.algn;
  const f = fillChild(ln);
  if (f) out.fill = parseFillNode(f, ctx);
  const pd = ln.child('a:prstDash');
  if (pd) out.dash = pd.attr('val', 'solid');
  const cd = ln.child('a:custDash');
  if (cd) {
    out.dash = 'custom';
    out.custDash = cd.all('a:ds').map((d) => [pct(d.attrs.d, 1), pct(d.attrs.sp, 1)]);
  }
  if (ln.child('a:round')) out.join = 'round';
  else if (ln.child('a:bevel')) out.join = 'bevel';
  else if (ln.child('a:miter')) {
    out.join = 'miter';
    out.miterLimit = ln.child('a:miter').num('lim', 800000) / 100000;
  }
  const he = parseEnd(ln.child('a:headEnd'));
  if (he) out.head = he;
  const te = parseEnd(ln.child('a:tailEnd'));
  if (te) out.tail = te;
  return out;
}

/** Merge partial line models (later arguments override earlier ones). */
export function mergeLines(...parts) {
  const out = {};
  for (const p of parts) if (p) Object.assign(out, p);
  return out;
}

/** Finalise a merged line: returns null when there is no visible line. */
export function finalizeLine(l) {
  if (!l || !l.fill || l.fill.type === 'none') return null;
  return {
    w: l.w === undefined ? 0.75 : Math.max(l.w, 0.5),
    fill: l.fill,
    dash: l.dash || 'solid',
    custDash: l.custDash,
    cap: l.cap || 'flat',
    cmpd: l.cmpd || 'sng',
    algn: l.algn || 'ctr',
    join: l.join || 'round',
    miterLimit: l.miterLimit || 8,
    head: l.head && l.head.type !== 'none' ? l.head : null,
    tail: l.tail && l.tail.type !== 'none' ? l.tail : null,
  };
}

function shadowCommon(n, ctx) {
  const dist = emu2px(n.num('dist', 0));
  const dir = n.num('dir', 0) / 60000;
  return {
    color: colorFrom(n, ctx) || { r: 0, g: 0, b: 0, a: 0.5 },
    blur: emu2px(n.num('blurRad', 0)),
    dx: dist * Math.cos((dir * Math.PI) / 180),
    dy: dist * Math.sin((dir * Math.PI) / 180),
    dist,
    dir,
  };
}

/** Parse an a:effectLst element. */
export function parseEffects(lst, ctx) {
  if (!lst) return null;
  const fx = {};
  for (const e of lst.children) {
    switch (e.name) {
      case 'a:outerShdw': {
        const s = shadowCommon(e, ctx);
        s.sx = pct(e.attrs.sx, 1);
        s.sy = pct(e.attrs.sy, 1);
        s.kx = e.num('kx', 0) / 60000;
        s.ky = e.num('ky', 0) / 60000;
        s.algn = e.attr('algn', 'b');
        s.rotWithShape = e.bool('rotWithShape', true);
        fx.outerShadow = s;
        break;
      }
      case 'a:innerShdw':
        fx.innerShadow = shadowCommon(e, ctx);
        break;
      case 'a:prstShdw': {
        const s = shadowCommon(e, ctx);
        s.prst = e.attr('prst', 'shdw1');
        s.sx = 1; s.sy = 1; s.kx = 0; s.ky = 0;
        fx.outerShadow = s;
        break;
      }
      case 'a:glow':
        fx.glow = { rad: emu2px(e.num('rad', 0)), color: colorFrom(e, ctx) || { r: 255, g: 255, b: 0, a: 0.6 } };
        break;
      case 'a:softEdge':
        fx.softEdge = { rad: emu2px(e.num('rad', 0)) };
        break;
      case 'a:reflection':
        fx.reflection = {
          blur: emu2px(e.num('blurRad', 0)),
          stA: pct(e.attrs.stA, 1), stPos: pct(e.attrs.stPos, 0),
          endA: pct(e.attrs.endA, 0), endPos: pct(e.attrs.endPos, 1),
          dist: emu2px(e.num('dist', 0)), dir: e.num('dir', 5400000) / 60000,
          fadeDir: e.num('fadeDir', 5400000) / 60000,
          sx: pct(e.attrs.sx, 1), sy: pct(e.attrs.sy, -1), algn: e.attr('algn', 'b'),
        };
        break;
      case 'a:blur':
        fx.blur = { rad: emu2px(e.num('rad', 0)), grow: e.bool('grow', true) };
        break;
      case 'a:fillOverlay': {
        const f = parseFillNode(fillChild(e), ctx);
        if (f) fx.fillOverlay = { blend: e.attr('blend', 'over'), fill: f };
        break;
      }
      default:
        break;
    }
  }
  return fx;
}

/** Colour of a style reference (a:fillRef etc.). */
export function refColor(ref, ctx) {
  if (!ref) return null;
  const c = colorChild(ref);
  return c ? colorFrom(ref, ctx) : null;
}

/** Resolve a fill from a theme style reference (<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>). */
export function fillFromRef(ref, ctx) {
  if (!ref) return undefined;
  const idx = ref.int('idx', 0);
  if (idx === 0) return { type: 'none' };
  const phClr = refColor(ref, ctx);
  const theme = ctx.theme;
  let node = null;
  if (idx >= 1000) node = theme.bgFillStyles[idx - 1001];
  else node = theme.fillStyles[idx - 1];
  if (!node) return phClr ? { type: 'solid', color: phClr } : undefined;
  return parseFillNode(node, ctx.with({ phClr }));
}

export function lineFromRef(ref, ctx) {
  if (!ref) return null;
  const idx = ref.int('idx', 0);
  if (idx === 0) return null;
  const phClr = refColor(ref, ctx);
  const node = ctx.theme.lineStyles[idx - 1];
  if (!node) return phClr ? { fill: { type: 'solid', color: phClr }, w: 0.75 } : null;
  return parseLine(node, ctx.with({ phClr }));
}

export function effectsFromRef(ref, ctx) {
  if (!ref) return null;
  const idx = ref.int('idx', 0);
  if (idx === 0) return null;
  const phClr = refColor(ref, ctx);
  const node = ctx.theme.effectStyles[idx - 1];
  if (!node) return null;
  const sub = ctx.with({ phClr });
  const lst = node.child('a:effectLst');
  const fx = lst ? parseEffects(lst, sub) : null;
  if (node.child('a:scene3d') || node.child('a:sp3d')) {
    const sp3d = node.child('a:sp3d');
    if (sp3d && (sp3d.child('a:bevelT') || sp3d.child('a:bevelB'))) ctx.note('approx', '3D bevel from theme');
  }
  return fx;
}

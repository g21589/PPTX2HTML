/**
 * DrawingML geometry engine: evaluates shape guide formulas for preset and
 * custom geometries and produces path segments in pixel space.
 */
import { PRESETS } from './presets.js';
import { TEXT_WARPS } from './textwarps.js';

const DEG = Math.PI / 180;
const ANG = 60000; // 1 degree in DrawingML angle units

function toRad(a) {
  return (a / ANG) * DEG;
}

/** Built-in guide variables for a w x h shape. */
function builtins(w, h) {
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  return {
    w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls,
    wd2: w / 2, wd3: w / 3, wd4: w / 4, wd5: w / 5, wd6: w / 6, wd8: w / 8, wd10: w / 10, wd12: w / 12, wd16: w / 16, wd32: w / 32,
    hd2: h / 2, hd3: h / 3, hd4: h / 4, hd5: h / 5, hd6: h / 6, hd8: h / 8, hd10: h / 10, hd12: h / 12, hd16: h / 16, hd32: h / 32,
    ssd2: ss / 2, ssd4: ss / 4, ssd6: ss / 6, ssd8: ss / 8, ssd16: ss / 16, ssd32: ss / 32,
    cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
  };
}

function val(env, tok) {
  if (tok === undefined) return 0;
  const v = env[tok];
  if (v !== undefined) return v;
  const n = Number(tok);
  return Number.isNaN(n) ? 0 : n;
}

// Evaluate one guide formula string (e.g. "*/ w adj 100000") in an environment.
export function evalFormula(fmla, env) {
  const p = fmla.trim().split(/\s+/);
  const op = p[0];
  const x = val(env, p[1]);
  const y = val(env, p[2]);
  const z = val(env, p[3]);
  switch (op) {
    case 'val': return x;
    case '*/': return z === 0 ? 0 : (x * y) / z;
    case '+-': return x + y - z;
    case '+/': return z === 0 ? 0 : (x + y) / z;
    case '?:': return x > 0 ? y : z;
    case 'abs': return Math.abs(x);
    case 'at2': return (Math.atan2(y, x) / DEG) * ANG;
    case 'cat2': return x * Math.cos(Math.atan2(z, y));
    case 'sat2': return x * Math.sin(Math.atan2(z, y));
    case 'cos': return x * Math.cos(toRad(y));
    case 'sin': return x * Math.sin(toRad(y));
    case 'tan': return x * Math.tan(toRad(y));
    case 'max': return Math.max(x, y);
    case 'min': return Math.min(x, y);
    case 'mod': return Math.sqrt(x * x + y * y + z * z);
    case 'pin': return y < x ? x : y > z ? z : y;
    case 'sqrt': return Math.sqrt(Math.max(0, x));
    default: return 0;
  }
}

function parseGuideList(s) {
  if (!s) return [];
  return s.split(';').map((g) => {
    const i = g.indexOf(' ');
    return [g.slice(0, i), g.slice(i + 1)];
  });
}

const presetCache = new Map();

/** Parse a compact preset definition into a geometry definition object. */
function presetDef(name, table = PRESETS) {
  const key = (table === PRESETS ? 's:' : 'w:') + name;
  if (presetCache.has(key)) return presetCache.get(key);
  const p = Object.prototype.hasOwnProperty.call(table, name) ? table[name] : null;
  if (!p) {
    presetCache.set(key, null);
    return null;
  }
  const def = {
    av: parseGuideList(p.av),
    gd: parseGuideList(p.gd),
    rect: p.rect ? p.rect.split(' ') : null,
    cxn: p.cxn ? p.cxn.split(';').map((c) => c.split(' ')) : [],
    paths: p.path.map((ps) => {
      const parts = ps.split('|');
      const [w, h, fill, stroke, extrusionOk] = parts[0].split(' ');
      return {
        w: Number(w), h: Number(h), fill, stroke: stroke === '1', extrusionOk: extrusionOk === '1',
        cmds: parts.slice(1).map((c) => c.split(' ')),
      };
    }),
  };
  presetCache.set(key, def);
  return def;
}

export function hasPreset(name) {
  return !!PRESETS[name];
}

export function presetNames() {
  return Object.keys(PRESETS);
}

/**
 * Parse a <a:custGeom> element into a geometry definition.
 * @param {import('../util/xml.js').XNode} node
 */
export function customGeometryDef(node) {
  const gl = (n) => (n ? n.all('a:gd').map((g) => [g.attrs.name, g.attrs.fmla || 'val 0']) : []);
  const rectN = node.child('a:rect');
  const def = {
    av: gl(node.child('a:avLst')),
    gd: gl(node.child('a:gdLst')),
    rect: rectN ? [rectN.attrs.l || 'l', rectN.attrs.t || 't', rectN.attrs.r || 'r', rectN.attrs.b || 'b'] : null,
    cxn: [],
    paths: [],
  };
  const pl = node.child('a:pathLst');
  for (const p of pl ? pl.all('a:path') : []) {
    const path = {
      w: p.num('w', 0), h: p.num('h', 0), fill: p.attr('fill', 'norm'),
      stroke: p.bool('stroke', true), extrusionOk: p.bool('extrusionOk', true), cmds: [],
    };
    for (const c of p.children) {
      const pts = [];
      for (const pt of c.all('a:pt')) pts.push(pt.attrs.x || '0', pt.attrs.y || '0');
      switch (c.local) {
        case 'moveTo': path.cmds.push(['M', ...pts]); break;
        case 'lnTo': path.cmds.push(['L', ...pts]); break;
        case 'cubicBezTo': path.cmds.push(['C', ...pts]); break;
        case 'quadBezTo': path.cmds.push(['Q', ...pts]); break;
        case 'arcTo': path.cmds.push(['A', c.attrs.wR || '0', c.attrs.hR || '0', c.attrs.stAng || '0', c.attrs.swAng || '0']); break;
        case 'close': path.cmds.push(['Z']); break;
        default: break;
      }
    }
    def.paths.push(path);
  }
  return def;
}

/** Parse adjust values from <a:avLst> of a prstGeom: {adj: 50000, ...}. */
export function parseAvLst(avLst) {
  const out = {};
  if (!avLst) return out;
  for (const gd of avLst.all('a:gd')) {
    const f = (gd.attrs.fmla || '').trim().split(/\s+/);
    if (f[0] === 'val') out[gd.attrs.name] = Number(f[1]) || 0;
    else out[gd.attrs.name] = f;
  }
  return out;
}

/** Convert an elliptical arc (DrawingML semantics) to cubic Bézier segments. */
function arcToBeziers(x0, y0, wR, hR, stAng, swAng, out) {
  if (wR === 0 || hR === 0 || swAng === 0) {
    return [x0, y0];
  }
  const st = toRad(stAng);
  const sw = toRad(swAng);
  // DrawingML angles are visual angles; convert to parametric ones.
  const param = (a) => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const t0 = param(st);
  let t1 = param(st + sw);
  // Keep the sweep direction and number of turns.
  let dt = t1 - t0;
  const turns = Math.floor(Math.abs(sw) / (2 * Math.PI) + 1e-9);
  if (sw > 0) {
    while (dt < -1e-9) dt += 2 * Math.PI;
    if (Math.abs(dt) < 1e-9 && turns === 0) dt = 0;
  } else {
    while (dt > 1e-9) dt -= 2 * Math.PI;
  }
  dt += Math.sign(sw) * turns * 2 * Math.PI;
  if (Math.abs(sw) >= 2 * Math.PI - 1e-9 && Math.abs(dt) < 1e-6) dt = Math.sign(sw) * 2 * Math.PI;
  const cx = x0 - wR * Math.cos(t0);
  const cy = y0 - hR * Math.sin(t0);
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
  const step = dt / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let a = t0;
  let px = x0;
  let py = y0;
  for (let i = 0; i < n; i++) {
    const b = a + step;
    const cosA = Math.cos(a), sinA = Math.sin(a), cosB = Math.cos(b), sinB = Math.sin(b);
    const ex = cx + wR * cosB;
    const ey = cy + hR * sinB;
    out.push({
      c: 'C',
      p: [px - k * wR * sinA, py + k * hR * cosA, ex + k * wR * sinB, ey - k * hR * cosB, ex, ey],
    });
    px = ex;
    py = ey;
    a = b;
  }
  return [px, py];
}

/**
 * Compute a geometry.
 * @param {object} def geometry definition (preset or custom)
 * @param {number} wEmu shape width in EMU
 * @param {number} hEmu shape height in EMU
 * @param {object} adj adjust value overrides
 * @param {number} scale EMU -> output unit factor
 */
export function computeGeometry(def, wEmu, hEmu, adj = {}, scale = 1) {
  const env = builtins(wEmu, hEmu);
  for (const [name, fmla] of def.av) {
    if (adj[name] !== undefined && !Array.isArray(adj[name])) env[name] = adj[name];
    else if (Array.isArray(adj[name])) env[name] = evalFormula(adj[name].join(' '), env);
    else env[name] = evalFormula(fmla, env);
  }
  // Custom geometries may reference adjust values not listed in avLst.
  for (const k of Object.keys(adj)) if (env[k] === undefined && !Array.isArray(adj[k])) env[k] = adj[k];
  for (const [name, fmla] of def.gd) env[name] = evalFormula(fmla, env);

  const paths = [];
  for (const p of def.paths) {
    const sx = (p.w ? wEmu / p.w : 1) * scale;
    const sy = (p.h ? hEmu / p.h : 1) * scale;
    const segs = [];
    let cx = 0, cy = 0, startX = 0, startY = 0;
    for (const cmd of p.cmds) {
      const op = cmd[0];
      if (op === 'M' || op === 'L') {
        cx = val(env, cmd[1]) * sx;
        cy = val(env, cmd[2]) * sy;
        segs.push({ c: op, p: [cx, cy] });
        if (op === 'M') { startX = cx; startY = cy; }
      } else if (op === 'C') {
        const pts = [];
        for (let i = 1; i < 7; i += 2) pts.push(val(env, cmd[i]) * sx, val(env, cmd[i + 1]) * sy);
        segs.push({ c: 'C', p: pts });
        cx = pts[4]; cy = pts[5];
      } else if (op === 'Q') {
        const pts = [];
        for (let i = 1; i < 5; i += 2) pts.push(val(env, cmd[i]) * sx, val(env, cmd[i + 1]) * sy);
        segs.push({ c: 'Q', p: pts });
        cx = pts[2]; cy = pts[3];
      } else if (op === 'A') {
        const wR = val(env, cmd[1]) * sx;
        const hR = val(env, cmd[2]) * sy;
        if (!segs.length) segs.push({ c: 'M', p: [cx, cy] });
        [cx, cy] = arcToBeziers(cx, cy, wR, hR, val(env, cmd[3]), val(env, cmd[4]), segs);
      } else if (op === 'Z') {
        segs.push({ c: 'Z', p: [] });
        cx = startX; cy = startY;
      }
    }
    paths.push({ segs, fill: p.fill, stroke: p.stroke, extrusionOk: p.extrusionOk });
  }

  let textRect = { l: 0, t: 0, r: wEmu * scale, b: hEmu * scale };
  if (def.rect) {
    textRect = {
      l: val(env, def.rect[0]) * scale,
      t: val(env, def.rect[1]) * scale,
      r: val(env, def.rect[2]) * scale,
      b: val(env, def.rect[3]) * scale,
    };
  }
  const cxn = def.cxn.map(([ang, x, y]) => ({ ang: val(env, ang) / ANG, x: val(env, x) * scale, y: val(env, y) * scale }));
  return { paths, textRect, cxn };
}

/** Preset geometry by name (null when unknown). */
export function presetGeometry(name, wEmu, hEmu, adj, scale) {
  const def = presetDef(name);
  if (!def) return null;
  return computeGeometry(def, wEmu, hEmu, adj, scale);
}

/** Preset text warp geometry (prstTxWarp) by name, null when unknown. */
export function textWarpGeometry(name, wEmu, hEmu, adj, scale) {
  const def = presetDef(name, TEXT_WARPS);
  if (!def) return null;
  return computeGeometry(def, wEmu, hEmu, adj, scale);
}

/** Flatten path segments to a polyline [x0, y0, x1, y1, ...]. */
export function segsToPolyline(segs, steps = 16) {
  const out = [];
  let cx = 0, cy = 0;
  for (const s of segs) {
    const p = s.p;
    if (s.c === 'M' || s.c === 'L') { cx = p[0]; cy = p[1]; out.push(cx, cy); }
    else if (s.c === 'C') {
      for (let i = 1; i <= steps; i++) {
        const t = i / steps, u = 1 - t;
        out.push(u * u * u * cx + 3 * u * u * t * p[0] + 3 * u * t * t * p[2] + t * t * t * p[4], u * u * u * cy + 3 * u * u * t * p[1] + 3 * u * t * t * p[3] + t * t * t * p[5]);
      }
      cx = p[4]; cy = p[5];
    } else if (s.c === 'Q') {
      for (let i = 1; i <= steps; i++) {
        const t = i / steps, u = 1 - t;
        out.push(u * u * cx + 2 * u * t * p[0] + t * t * p[2], u * u * cy + 2 * u * t * p[1] + t * t * p[3]);
      }
      cx = p[2]; cy = p[3];
    }
  }
  return out;
}

export function presetAdjustNames(name) {
  const def = presetDef(name);
  return def ? def.av.map((a) => a[0]) : [];
}

const fmt = (n) => {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

/** Serialize path segments to an SVG path data string. */
export function segsToSvg(segs, dx = 0, dy = 0) {
  let d = '';
  for (const s of segs) {
    const p = s.p;
    switch (s.c) {
      case 'M': d += 'M' + fmt(p[0] + dx) + ' ' + fmt(p[1] + dy); break;
      case 'L': d += 'L' + fmt(p[0] + dx) + ' ' + fmt(p[1] + dy); break;
      case 'C': d += 'C' + fmt(p[0] + dx) + ' ' + fmt(p[1] + dy) + ' ' + fmt(p[2] + dx) + ' ' + fmt(p[3] + dy) + ' ' + fmt(p[4] + dx) + ' ' + fmt(p[5] + dy); break;
      case 'Q': d += 'Q' + fmt(p[0] + dx) + ' ' + fmt(p[1] + dy) + ' ' + fmt(p[2] + dx) + ' ' + fmt(p[3] + dy); break;
      case 'Z': d += 'Z'; break;
      default: break;
    }
  }
  return d;
}

/** Bounding box of path segments (control points included). */
export function segsBounds(segs) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const s of segs) {
    for (let i = 0; i < s.p.length; i += 2) {
      const x = s.p[i], y = s.p[i + 1];
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
  }
  if (x0 === Infinity) return { x: 0, y: 0, w: 0, h: 0 };
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/**
 * Tangent information at the ends of the open sub-paths of a path:
 * [{x, y, angle}] for the start (pointing outwards) and the end.
 */
export function pathEnds(segs) {
  if (!segs.length) return null;
  let start = null;
  let startDir = null;
  let last = null;
  let prev = null;
  let endDir = null;
  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    if (s.c === 'M') {
      if (start === null) start = [s.p[0], s.p[1]];
      prev = [s.p[0], s.p[1]];
      continue;
    }
    if (s.c === 'Z') return null; // closed path: no arrow heads
    const p = s.p;
    const end = [p[p.length - 2], p[p.length - 1]];
    // Direction at the beginning of this segment.
    let fx = p[0], fy = p[1];
    if ((fx === prev[0] && fy === prev[1]) && p.length > 2) { fx = p[2]; fy = p[3]; }
    if ((fx === prev[0] && fy === prev[1]) && p.length > 4) { fx = p[4]; fy = p[5]; }
    if (startDir === null && (fx !== prev[0] || fy !== prev[1])) startDir = [prev[0] - fx, prev[1] - fy];
    // Direction at the end of this segment.
    let bx = p.length > 2 ? p[p.length - 4] : prev[0];
    let by = p.length > 2 ? p[p.length - 3] : prev[1];
    if (bx === end[0] && by === end[1] && p.length > 4) { bx = p[p.length - 6]; by = p[p.length - 5]; }
    if (bx === end[0] && by === end[1]) { bx = prev[0]; by = prev[1]; }
    if (bx !== end[0] || by !== end[1]) endDir = [end[0] - bx, end[1] - by];
    last = end;
    prev = end;
  }
  if (!start || !last) return null;
  return {
    start: { x: start[0], y: start[1], angle: startDir ? Math.atan2(startDir[1], startDir[0]) : Math.PI },
    end: { x: last[0], y: last[1], angle: endDir ? Math.atan2(endDir[1], endDir[0]) : 0 },
  };
}

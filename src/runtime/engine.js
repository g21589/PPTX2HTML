/**
 * Animation engine: evaluates the flattened PresentationML timing model at
 * any instant. Pure logic (no DOM) so it can be unit tested and used for
 * seeking, stepping backwards and rendering final states.
 */
import { evalFormula } from './formula.js';

const VIS_ATTRS = new Set(['style.visibility']);

/** SMIL acceleration / deceleration of a simple progress value. */
export function ease(p, a, d) {
  if (!a && !d) return p;
  if (a + d > 1) { const s = a + d; a /= s; d /= s; }
  const r = 1 / (1 - a / 2 - d / 2);
  if (p < a) return (r * p * p) / (2 * a);
  if (p <= 1 - d) return r * (p - a / 2);
  const q = 1 - p;
  return 1 - (r * q * q) / (2 * d);
}

/** Map step time through the repeat / auto-reverse wrappers of enclosing containers. */
function unwrap(T, wrappers) {
  let t = T;
  for (const w of wrappers || []) {
    const rel = t - w.begin;
    if (rel < 0) return t;
    const single = w.dur / (w.spd || 1);
    if (!(single > 0)) continue;
    const period = single * (w.autoRev ? 2 : 1);
    const total = w.repeat === Infinity ? Infinity : period * (w.repeat || 1);
    let k;
    if (rel >= total) {
      const frac = (w.repeat || 1) % 1;
      if (frac) {
        k = frac * period;
        if (w.autoRev && k > single) k = period - k;
      } else k = w.autoRev ? 0 : single;
    } else {
      k = rel % period;
      if (w.autoRev && k > single) k = period - k;
    }
    let p = k / single;
    p = ease(p, w.accel || 0, w.decel || 0);
    t = w.begin + p * w.dur;
  }
  return t;
}

/** Active duration of a behaviour. */
export function activeDuration(b) {
  if (b.dur === Infinity) return Infinity;
  const single = b.dur / (b.spd || 1);
  if (b.repeatDur !== undefined && b.repeatDur !== null) return b.repeatDur;
  const per = single * (b.autoRev ? 2 : 1);
  return b.repeat === Infinity ? Infinity : per * (b.repeat || 1);
}

/**
 * Progress (0..1) of a behaviour at step time T, or null when it does not
 * contribute (not started yet, or removed after its end).
 * The returned object also reports whether it is still running.
 */
export function behaviourProgress(b, T, extraDelay = 0) {
  const t = unwrap(T, b.wrappers) - (b.begin + extraDelay);
  if (t < 0) return null;
  const single = b.dur === Infinity ? Infinity : b.dur / (b.spd || 1);
  const active = activeDuration(b);
  if (t >= active) {
    if (b.fill === 'remove') return null;
    if (single === Infinity) return { p: 0, running: false };
    const rep = b.repeat === Infinity ? 1 : (b.repeat || 1);
    const frac = rep % 1;
    let p;
    if (frac) {
      p = frac * (b.autoRev ? 2 : 1);
      if (p > 1) p = 2 - p;
    } else p = b.autoRev ? 0 : 1;
    return { p: ease(p, b.accel || 0, b.decel || 0), running: false };
  }
  if (single === Infinity || single <= 0) return { p: single <= 0 ? 1 : 0, running: single > 0 };
  const period = single * (b.autoRev ? 2 : 1);
  let k = t % period;
  if (b.autoRev && k > single) k = period - k;
  return { p: ease(k / single, b.accel || 0, b.decel || 0), running: true };
}

function newState() {
  return {
    visible: undefined, opacity: 1, x: undefined, y: undefined, w: undefined, h: undefined,
    rot: 0, absRot: undefined, dx: 0, dy: 0, sx: 1, sy: 1, skewX: 0, skewY: 0,
    effects: [], fill: undefined, stroke: undefined, color: undefined,
    bold: undefined, italic: undefined, underline: undefined, fontScale: undefined, fillOn: undefined, strokeOn: undefined,
  };
}

function lerp(a, b, t) { return a + (b - a) * t; }

function mixColor(c1, c2, t) {
  return { r: lerp(c1.r, c2.r, t), g: lerp(c1.g, c2.g, t), b: lerp(c1.b, c2.b, t), a: lerp(c1.a === undefined ? 1 : c1.a, c2.a === undefined ? 1 : c2.a, t) };
}

function rgbToHsl({ r, g, b }) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0, s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h *= 60;
  }
  return [h, s, l];
}

function hslToRgb(h, s, l) {
  h = ((h % 360) + 360) % 360;
  s = Math.max(0, Math.min(1, s));
  l = Math.max(0, Math.min(1, l));
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const hk = h / 360;
  return { r: f(hk + 1 / 3) * 255, g: f(hk) * 255, b: f(hk - 1 / 3) * 255, a: 1 };
}

/** Point along a motion path at progress p (paced by arc length). */
export function pathPoint(path, p, aspect = 1) {
  if (!path || !path.length) return [0, 0];
  // Flatten into polyline.
  const pts = [];
  let cx = 0, cy = 0, sx = 0, sy = 0;
  for (const c of path) {
    if (c[0] === 'M') { cx = c[1]; cy = c[2]; sx = cx; sy = cy; pts.push([cx, cy, true]); }
    else if (c[0] === 'L') { cx = c[1]; cy = c[2]; pts.push([cx, cy]); }
    else if (c[0] === 'C') {
      const [, x1, y1, x2, y2, x3, y3] = c;
      for (let i = 1; i <= 16; i++) {
        const t = i / 16, u = 1 - t;
        pts.push([u * u * u * cx + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3, u * u * u * cy + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3]);
      }
      cx = x3; cy = y3;
    } else if (c[0] === 'Z') { cx = sx; cy = sy; pts.push([cx, cy]); }
  }
  if (!pts.length) return [0, 0];
  if (pts.length === 1) return [pts[0][0], pts[0][1]];
  const lens = [0];
  // Pace by on-screen distance: x is in slide widths, y in slide heights.
  for (let i = 1; i < pts.length; i++) lens.push(lens[i - 1] + (pts[i][2] ? 0 : Math.hypot((pts[i][0] - pts[i - 1][0]) * aspect, pts[i][1] - pts[i - 1][1])));
  const total = lens[lens.length - 1];
  if (total === 0) return [pts[pts.length - 1][0], pts[pts.length - 1][1]];
  const target = Math.max(0, Math.min(1, p)) * total;
  let i = 1;
  while (i < lens.length - 1 && lens[i] < target) i++;
  const seg = lens[i] - lens[i - 1];
  const f = seg ? (target - lens[i - 1]) / seg : 0;
  return [lerp(pts[i - 1][0], pts[i][0], f), lerp(pts[i - 1][1], pts[i][1], f)];
}

function variables(geo, st) {
  return {
    '#ppt_x': geo.x, '#ppt_y': geo.y, '#ppt_w': geo.w, '#ppt_h': geo.h, '#ppt_r': geo.rot || 0,
    ppt_x: st.x !== undefined ? st.x : geo.x, ppt_y: st.y !== undefined ? st.y : geo.y,
    ppt_w: st.w !== undefined ? st.w : geo.w, ppt_h: st.h !== undefined ? st.h : geo.h,
    ppt_r: geo.rot || 0, ppt_c: 0,
  };
}

function numeric(v, vars) {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === null || v === undefined) return NaN;
  return evalFormula(String(v), vars);
}

/** Value of an p:anim behaviour at progress p. */
function animValue(b, p, vars, base) {
  const tav = b.tav;
  if (tav && tav.length) {
    let i = 0;
    while (i < tav.length - 1 && tav[i + 1].tm <= p) i++;
    const a = tav[i];
    const nxt = tav[i + 1];
    const evalTav = (t, local) => {
      let v = t.val;
      if (b.valueType === 'str' && typeof v === 'string' && !/[#$]|ppt_|^[-0-9.]/.test(v)) return v;
      if (b.valueType === 'clr' || (v && typeof v === 'object')) return v;
      v = numeric(v, vars);
      if (t.fmla) v = numeric(t.fmla, { ...vars, $: local !== undefined ? local : v });
      return v;
    };
    if (!nxt || b.calcmode === 'discrete' || p <= a.tm) {
      if (b.calcmode === 'fmla' && a.fmla) return numeric(a.fmla, { ...vars, $: p });
      return evalTav(a, p);
    }
    const span = nxt.tm - a.tm;
    const f = span > 0 ? (p - a.tm) / span : 1;
    const va = a.val, vb = nxt.val;
    if (b.calcmode === 'fmla' && a.fmla) {
      return numeric(a.fmla, { ...vars, $: p });
    }
    if (va && typeof va === 'object' && vb && typeof vb === 'object') return mixColor(va, vb, f);
    if (b.valueType === 'str' && typeof va === 'string' && !/[#$]|ppt_|^[-0-9.]/.test(va)) return f < 1 ? va : vb;
    const na = numeric(va, vars), nb = numeric(vb, vars);
    let v = lerp(na, nb, f);
    if (a.fmla) v = numeric(a.fmla, { ...vars, $: v });
    return v;
  }
  // from / to / by.
  const from = b.from !== undefined ? numeric(b.from, vars) : base;
  let to;
  if (b.to !== undefined) to = numeric(b.to, vars);
  else if (b.by !== undefined) to = from + numeric(b.by, vars);
  else to = from;
  if (typeof b.to === 'string' && b.valueType === 'str' && Number.isNaN(to)) return p >= 1 ? b.to : b.from;
  return lerp(from, to, p);
}

const NUMERIC_ATTRS = new Set(['style.opacity', 'ppt_x', 'ppt_y', 'ppt_w', 'ppt_h', 'r', 'ppt_r', 'style.rotation', 'xshear', 'yshear', 'ScaleX', 'scaleX', 'ScaleY', 'scaleY', 'style.fontSize']);

function applyAttr(st, attr, v, b, geo) {
  if (typeof v === 'string' && NUMERIC_ATTRS.has(attr)) {
    const n = parseFloat(v);
    if (Number.isFinite(n)) v = n;
  }
  switch (attr) {
    case 'style.visibility':
      st.visible = v === 'visible' || v === true || v === 1;
      // A later entrance overrides earlier exit filters.
      if (st.visible) st.effects = st.effects.filter((e) => e.transition !== 'out');
      break;
    case 'style.opacity':
      if (Number.isFinite(v)) st.opacity = Math.max(0, Math.min(1, v));
      break;
    case 'ppt_x': if (Number.isFinite(v)) st.x = b.additive === 'sum' ? (st.x !== undefined ? st.x : geo.x) + v : v; break;
    case 'ppt_y': if (Number.isFinite(v)) st.y = b.additive === 'sum' ? (st.y !== undefined ? st.y : geo.y) + v : v; break;
    case 'ppt_w': if (Number.isFinite(v)) st.w = v; break;
    case 'ppt_h': if (Number.isFinite(v)) st.h = v; break;
    case 'r':
    case 'ppt_r':
    case 'style.rotation':
      if (Number.isFinite(v)) st.absRot = v;
      break;
    case 'xshear': if (Number.isFinite(v)) st.skewX = v; break;
    case 'yshear': if (Number.isFinite(v)) st.skewY = v; break;
    case 'ScaleX': case 'scaleX': if (Number.isFinite(v)) st.sx *= v; break;
    case 'ScaleY': case 'scaleY': if (Number.isFinite(v)) st.sy *= v; break;
    case 'fillcolor': case 'fillColor': if (v && typeof v === 'object') st.fill = v; break;
    case 'fill.on': st.fillOn = v === true || v === 'true' || v === 1; break;
    case 'fill.type': break;
    case 'stroke.color': if (v && typeof v === 'object') st.stroke = v; break;
    case 'stroke.on': st.strokeOn = v === true || v === 'true' || v === 1; break;
    case 'style.color': if (v && typeof v === 'object') st.color = v; break;
    case 'style.fontWeight': st.bold = v === 'bold' || v === true; break;
    case 'style.fontStyle': st.italic = v === 'italic'; break;
    case 'style.textDecorationUnderline': st.underline = v === true || v === 'true'; break;
    case 'style.fontSize': if (Number.isFinite(v)) st.fontScale = v; break;
    default: break;
  }
}

/**
 * Apply one behaviour at progress p to the states of its targets.
 */
let aspect = 16 / 9;

function applyBehaviour(b, p, states, geoOf, colorsOf) {
  for (const key of b.keys) {
    let st = states.get(key);
    if (!st) { st = newState(); states.set(key, st); }
    const geo = geoOf(key) || { x: 0.5, y: 0.5, w: 0.1, h: 0.1 };
    switch (b.type) {
      case 'set': {
        for (const a of b.attrs) applyAttr(st, a, b.to, b, geo);
        break;
      }
      case 'anim': {
        const vars = variables(geo, st);
        for (const a of b.attrs) {
          let base = 0;
          if (a === 'ppt_x') base = st.x !== undefined ? st.x : geo.x;
          else if (a === 'ppt_y') base = st.y !== undefined ? st.y : geo.y;
          else if (a === 'ppt_w') base = st.w !== undefined ? st.w : geo.w;
          else if (a === 'ppt_h') base = st.h !== undefined ? st.h : geo.h;
          else if (a === 'style.opacity') base = st.opacity;
          else if (a === 'style.rotation' || a === 'r' || a === 'ppt_r') base = st.absRot !== undefined ? st.absRot : (geo.rot || 0);
          else if (a === 'style.fontSize') base = 1;
          const v = animValue(b, p, vars, base);
          if (b.additive === 'sum' && (a === 'style.rotation' || a === 'r')) applyAttr(st, a, base + v, b, geo);
          else applyAttr(st, a, v, b, geo);
        }
        break;
      }
      case 'effect': {
        const reveal = b.transition === 'out' ? 1 - p : p;
        if (b.transition !== 'none') {
          st.effects = st.effects.filter((e) => e.transition === b.transition);
          st.effects.push({ filter: b.filter, reveal, transition: b.transition });
        }
        break;
      }
      case 'motion': {
        let pt;
        if (b.path && b.path.length) pt = pathPoint(b.path, p, aspect);
        else if (b.to || b.by) {
          const from = b.from || [0, 0];
          const to = b.to || [from[0] + b.by[0], from[1] + b.by[1]];
          pt = [lerp(from[0], to[0], p), lerp(from[1], to[1], p)];
        } else pt = [0, 0];
        st.dx += pt[0];
        st.dy += pt[1];
        break;
      }
      case 'rotate': {
        if (b.by !== undefined) st.rot += b.by * p;
        else if (b.to !== undefined) {
          const from = b.from !== undefined ? b.from : 0;
          st.rot += lerp(from, b.to, p);
        }
        break;
      }
      case 'scale': {
        let fx = 1, fy = 1;
        if (b.by) { fx = lerp(1, b.by[0], p); fy = lerp(1, b.by[1], p); }
        else if (b.to) {
          const from = b.from || [1, 1];
          fx = lerp(from[0], b.to[0], p); fy = lerp(from[1], b.to[1], p);
        }
        st.sx *= fx;
        st.sy *= fy;
        break;
      }
      case 'color': {
        const attr = b.attrs[0] || 'style.color';
        const baseCols = colorsOf ? colorsOf(key) : null;
        const cur = attr === 'fillcolor' || attr === 'fillColor' ? (st.fill || (baseCols && baseCols.fill)) : attr === 'stroke.color' ? (st.stroke || (baseCols && baseCols.stroke)) : (st.color || (baseCols && baseCols.color));
        const from = b.from && b.from.color ? b.from.color : (cur || { r: 0, g: 0, b: 0, a: 1 });
        let c;
        if (b.to && b.to.color) {
          if (b.clrSpc === 'hsl') {
            const [h1, s1, l1] = rgbToHsl(from), [h2, s2, l2] = rgbToHsl(b.to.color);
            let dh = h2 - h1;
            if (b.dir === 'cw' && dh < 0) dh += 360;
            if (b.dir === 'ccw' && dh > 0) dh -= 360;
            c = hslToRgb(h1 + dh * p, lerp(s1, s2, p), lerp(l1, l2, p));
          } else c = mixColor(from, b.to.color, p);
        } else if (b.by && b.by.hsl) {
          const [h1, s1, l1] = rgbToHsl(from);
          const [dh, ds, dl] = b.by.hsl;
          c = hslToRgb(h1 + (b.dir === 'ccw' ? -1 : 1) * dh * p, s1 + ds * p, l1 + dl * p);
        } else if (b.by && b.by.rgb) {
          const [dr, dg, db] = b.by.rgb;
          c = { r: from.r + dr * 2.55 * p, g: from.g + dg * 2.55 * p, b: from.b + db * 2.55 * p, a: 1 };
        } else c = from;
        if (attr === 'fillcolor' || attr === 'fillColor') st.fill = c;
        else if (attr === 'stroke.color') st.stroke = c;
        else st.color = c;
        break;
      }
      default:
        break;
    }
  }
}

/** Iteration (by word / letter) delay for sub-element i of n. */
function iterationDelay(b, i, n) {
  const it = b.iterate;
  if (!it) return 0;
  const idx = it.backwards ? n - 1 - i : i;
  if (it.abs !== undefined) return idx * it.abs;
  if (it.pct !== undefined) {
    const base = it.base !== undefined ? it.base : (b.dur === Infinity ? 0 : b.dur);
    return idx * it.pct * base;
  }
  return 0;
}

/** End time (ms) of a step, taking iteration into account. */
export function stepEnd(step, countOf) {
  let end = 0;
  for (const b of step.behaviours) {
    let e = b.begin + activeDuration(b);
    if (b.iterate && countOf) {
      const n = Math.max(1, ...b.keys.map((k) => countOf(k) || 1));
      e += iterationDelay(b, n - 1, n);
    }
    if (b.wrappers && b.wrappers.length) {
      for (const w of b.wrappers) {
        const per = (w.dur / (w.spd || 1)) * (w.autoRev ? 2 : 1);
        const tot = w.repeat === Infinity ? Infinity : per * (w.repeat || 1);
        e = Math.max(e, w.begin + tot);
      }
    }
    if (e > end) end = e;
  }
  return end;
}

/**
 * Evaluate the timing at a position.
 * @param {object} timing parsed timing ({main, triggers, hidden})
 * @param {object} pos {step: number of completed main steps (current step index), t: ms into the current step,
 *                      triggers: [{index, step, t}]}
 * @param {(key: string) => object} geoOf original geometry {x, y, w, h} (slide fractions) of a target
 * @param {(key: string) => object} [colorsOf] base colours {fill, stroke, color} of a target
 * @param {(key: string) => number} [countOf] number of iteration sub-elements of a target
 * @returns {{states: Map<string, object>, commands: object[], running: boolean}}
 */
export function evaluate(timing, pos, geoOf, colorsOf, countOf, opts = {}) {
  aspect = opts.aspect || 16 / 9;
  const states = new Map();
  const commands = [];
  let running = false;
  if (!timing) return { states, commands, running };
  for (const key of timing.hidden || []) {
    const st = newState();
    st.visible = false;
    states.set(key, st);
  }
  const runStep = (step, T, tag) => {
    for (const b of step.behaviours) {
      if (b.type === 'cmd' || b.type === 'media') {
        const pr = behaviourProgress({ ...b, fill: 'hold' }, T);
        if (pr) commands.push({ b, tag });
        continue;
      }
      if (b.iterate && countOf) {
        // The container itself becomes visible with the first sub-element.
        if (b.type === 'set') {
          const pr0 = behaviourProgress(b, T);
          if (pr0) applyBehaviour(b, pr0.p, states, geoOf, colorsOf);
        }
        for (const key of b.keys) {
          const n = countOf(key) || 0;
          if (!n) {
            const pr = behaviourProgress(b, T);
            if (pr) { applyBehaviour({ ...b, keys: [key] }, pr.p, states, geoOf, colorsOf); if (pr.running) running = true; }
            continue;
          }
          for (let i = 0; i < n; i++) {
            const pr = behaviourProgress(b, T, iterationDelay(b, i, n));
            const sub = key + '#' + i;
            if (!pr) {
              // Sub-elements of an entrance stay hidden until their turn.
              if (b.type === 'set' && b.attrs.some((a) => VIS_ATTRS.has(a)) && b.to === 'visible') {
                let st = states.get(sub);
                if (!st) { st = newState(); states.set(sub, st); }
                st.visible = false;
              }
              continue;
            }
            if (pr.running) running = true;
            applyBehaviour({ ...b, keys: [sub] }, pr.p, states, geoOf, colorsOf);
          }
        }
        continue;
      }
      const pr = behaviourProgress(b, T);
      if (!pr) continue;
      if (pr.running) running = true;
      applyBehaviour(b, pr.p, states, geoOf, colorsOf);
    }
  };
  const main = timing.main || [];
  const cur = Math.min(pos.step, main.length);
  for (let i = 0; i < cur; i++) runStep(main[i], Infinity, 'main:' + i);
  if (cur < main.length && pos.t !== undefined && pos.t !== null && pos.t >= 0) runStep(main[cur], pos.t, 'main:' + cur);
  for (const tr of pos.triggers || []) {
    const seq = timing.triggers[tr.index];
    if (!seq) continue;
    for (let i = 0; i < tr.step && i < seq.steps.length; i++) runStep(seq.steps[i], Infinity, 'trig:' + tr.index + ':' + i);
    if (tr.step < seq.steps.length && tr.t >= 0) runStep(seq.steps[tr.step], tr.t, 'trig:' + tr.index + ':' + tr.step);
  }
  return { states, commands, running };
}

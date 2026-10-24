/**
 * Visual effects for the runtime: animation filters (p:animEffect) and slide
 * transitions, expressed as CSS for a progress value p in [0, 1].
 */

const pc = (v) => Math.round(v * 100000) / 1000 + '%';

/** Deterministic pseudo random permutation ranks. */
function ranks(n, seed) {
  const arr = Array.from({ length: n }, (_, i) => i);
  let s = seed >>> 0 || 1;
  for (let i = n - 1; i > 0; i--) {
    s = (s * 1664525 + 1013904223) >>> 0;
    const j = s % (i + 1);
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  const r = new Array(n);
  arr.forEach((v, i) => { r[v] = i; });
  return r;
}

const RB = ranks(48, 7);

function maskStyle(img) {
  return { maskImage: img, webkitMaskImage: img, maskSize: '100% 100%', webkitMaskSize: '100% 100%', maskRepeat: 'no-repeat', webkitMaskRepeat: 'no-repeat' };
}

/** Random bars as a gradient mask. */
function randomBars(p, vertical, n = 48) {
  const stops = [];
  for (let i = 0; i < n; i++) {
    const on = RB[i % RB.length] < p * n;
    const c = on ? '#000' : 'transparent';
    stops.push(`${c} ${pc(i / n)} ${pc((i + 1) / n)}`);
  }
  return maskStyle(`linear-gradient(${vertical ? 'to right' : 'to bottom'},${stops.join(',')})`);
}

let dissolveCache = new Map();
/** Dissolve: random blocks as an SVG mask image. */
function dissolve(p, cols = 24, rows = 14) {
  const n = cols * rows;
  const k = Math.round(p * n);
  if (dissolveCache.has(k)) return dissolveCache.get(k);
  const r = ranks(n, 99);
  let rects = '';
  for (let i = 0; i < n; i++) if (r[i] < k) rects += `<rect x="${i % cols}" y="${Math.floor(i / cols)}" width="1.02" height="1.02"/>`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${cols} ${rows}" preserveAspectRatio="none">${rects}</svg>`;
  const st = maskStyle(`url("data:image/svg+xml,${encodeURIComponent(svg)}")`);
  if (dissolveCache.size > 400) dissolveCache = new Map();
  dissolveCache.set(k, st);
  return st;
}

/** Polygon (CSS) of the unit square clipped to the half plane a*x + b*y <= c. */
function halfPlane(a, b, c) {
  const sq = [[0, 0], [1, 0], [1, 1], [0, 1]];
  const out = [];
  for (let i = 0; i < 4; i++) {
    const P = sq[i], Q = sq[(i + 1) % 4];
    const fp = a * P[0] + b * P[1] - c;
    const fq = a * Q[0] + b * Q[1] - c;
    if (fp <= 0) out.push(P);
    if ((fp < 0 && fq > 0) || (fp > 0 && fq < 0)) {
      const t = fp / (fp - fq);
      out.push([P[0] + (Q[0] - P[0]) * t, P[1] + (Q[1] - P[1]) * t]);
    }
  }
  if (out.length < 3) return 'polygon(0 0,0 0,0 0)';
  return 'polygon(' + out.map(([x, y]) => pc(x) + ' ' + pc(y)).join(',') + ')';
}

function diagonal(dir, p) {
  // dir: direction the strips move towards.
  switch (dir) {
    case 'downRight': case 'rd': return halfPlane(1, 1, 2 * p);
    case 'downLeft': case 'ld': return halfPlane(-1, 1, 2 * p - 1);
    case 'upRight': case 'ru': return halfPlane(1, -1, 2 * p - 1);
    case 'upLeft': case 'lu': return halfPlane(-1, -1, 2 * p - 2);
    default: return halfPlane(1, 1, 2 * p);
  }
}

function evenOddRect(a) {
  // Outer rect with a centred rectangular hole inset by a on each side.
  const b = 1 - a;
  return `polygon(evenodd,0 0,100% 0,100% 100%,0 100%,0 0,${pc(a)} ${pc(a)},${pc(a)} ${pc(b)},${pc(b)} ${pc(b)},${pc(b)} ${pc(a)},${pc(a)} ${pc(a)})`;
}

function diamondPoly(s) {
  return `polygon(50% ${pc(0.5 - s)},${pc(0.5 + s)} 50%,50% ${pc(0.5 + s)},${pc(0.5 - s)} 50%)`;
}

function plusPoly(p) {
  const a = (1 - p) / 2, b = 1 - a;
  return `polygon(${pc(a)} 0,${pc(b)} 0,${pc(b)} ${pc(a)},100% ${pc(a)},100% ${pc(b)},${pc(b)} ${pc(b)},${pc(b)} 100%,${pc(a)} 100%,${pc(a)} ${pc(b)},0 ${pc(b)},0 ${pc(a)},${pc(a)} ${pc(a)})`;
}

/**
 * CSS for an animation filter at reveal progress p (1 = fully visible).
 * Returns {opacity, clipPath, mask*, transform}.
 */
export function filterStyle(filter, p) {
  p = Math.max(0, Math.min(1, p));
  const m = /^\s*([a-zA-Z]+)\s*(?:\(([^)]*)\))?/.exec(filter || 'fade');
  const name = m ? m[1].toLowerCase() : 'fade';
  const arg = m && m[2] ? m[2].trim() : '';
  if (p >= 1) return {};
  switch (name) {
    case 'fade':
      return { opacity: p };
    case 'wipe':
      switch (arg) {
        case 'up': return { clipPath: `inset(${pc(1 - p)} 0 0 0)` };
        case 'left': return { clipPath: `inset(0 0 0 ${pc(1 - p)})` };
        case 'right': return { clipPath: `inset(0 ${pc(1 - p)} 0 0)` };
        default: return { clipPath: `inset(0 0 ${pc(1 - p)} 0)` };
      }
    case 'barn':
      switch (arg) {
        case 'inVertical': return maskStyle(`linear-gradient(to right,#000 ${pc(p / 2)},transparent ${pc(p / 2)} ${pc(1 - p / 2)},#000 ${pc(1 - p / 2)})`);
        case 'inHorizontal': return maskStyle(`linear-gradient(to bottom,#000 ${pc(p / 2)},transparent ${pc(p / 2)} ${pc(1 - p / 2)},#000 ${pc(1 - p / 2)})`);
        case 'outHorizontal': return { clipPath: `inset(${pc((1 - p) / 2)} 0 ${pc((1 - p) / 2)} 0)` };
        default: return { clipPath: `inset(0 ${pc((1 - p) / 2)} 0 ${pc((1 - p) / 2)})` };
      }
    case 'blinds': {
      const n = 6;
      const s = p / n;
      const dir = arg === 'vertical' ? 'to right' : 'to bottom';
      return maskStyle(`repeating-linear-gradient(${dir},#000 0 ${pc(s)},transparent ${pc(s)} ${pc(1 / n)})`);
    }
    case 'checkerboard': {
      // Two interleaved rows of blinds offset by half a cell.
      const n = 8;
      const dir = arg === 'down' ? 'to bottom' : 'to right';
      const cross = arg === 'down' ? 'to right' : 'to bottom';
      const s = p / n;
      const half = Math.min(1, p * 2) / n / 2;
      const a = `repeating-linear-gradient(${dir},#000 0 ${pc(s)},transparent ${pc(s)} ${pc(1 / n)})`;
      const b = `repeating-linear-gradient(${cross},#000 0 50%,transparent 50% 100%)`;
      void half;
      return { ...maskStyle(`${a},${b}`), maskComposite: 'add', webkitMaskComposite: 'source-over' };
    }
    case 'box':
      if (arg === 'in') return { clipPath: evenOddRect(p / 2) };
      return { clipPath: `inset(${pc((1 - p) / 2)})` };
    case 'circle':
      if (arg === 'in') return maskStyle(`radial-gradient(circle farthest-corner at 50% 50%,transparent ${pc(1 - p)},#000 ${pc(1 - p)})`);
      return { clipPath: `circle(${pc(p * 0.7072)} at 50% 50%)` };
    case 'diamond':
      if (arg === 'in') {
        const s = (1 - p);
        return { clipPath: `polygon(evenodd,0 0,100% 0,100% 100%,0 100%,0 0,50% ${pc(0.5 - s)},${pc(0.5 + s)} 50%,50% ${pc(0.5 + s)},${pc(0.5 - s)} 50%,50% ${pc(0.5 - s)})` };
      }
      return { clipPath: diamondPoly(p) };
    case 'plus':
      if (arg === 'in') {
        const a = p / 2, b = 1 - a;
        // Four shrinking corner squares are hidden: visible = outer frame bands.
        return maskStyle(`linear-gradient(to right,#000 ${pc(a)},transparent ${pc(a)} ${pc(b)},#000 ${pc(b)}),linear-gradient(to bottom,#000 ${pc(a)},transparent ${pc(a)} ${pc(b)},#000 ${pc(b)})`);
      }
      return { clipPath: plusPoly(p) };
    case 'dissolve':
    case 'pixelate':
      return dissolve(p);
    case 'randombar':
      return randomBars(p, arg === 'vertical');
    case 'slide':
      switch (arg) {
        case 'fromTop': return { transform: `translateY(${pc(-(1 - p))})`, clipPath: `inset(${pc(1 - p)} 0 0 0)` };
        case 'fromLeft': return { transform: `translateX(${pc(-(1 - p))})`, clipPath: `inset(0 0 0 ${pc(1 - p)})` };
        case 'fromRight': return { transform: `translateX(${pc(1 - p)})`, clipPath: `inset(0 ${pc(1 - p)} 0 0)` };
        default: return { transform: `translateY(${pc(1 - p)})`, clipPath: `inset(0 0 ${pc(1 - p)} 0)` };
      }
    case 'strips':
      return { clipPath: diagonal(arg, p) };
    case 'wedge': {
      const a = p * 180;
      return maskStyle(`conic-gradient(from 0deg at 50% 50%,#000 0 ${a}deg,transparent ${a}deg ${360 - a}deg,#000 ${360 - a}deg)`);
    }
    case 'wheel': {
      const n = Math.max(1, parseInt(arg, 10) || 1);
      const seg = 360 / n;
      return maskStyle(`repeating-conic-gradient(from 0deg at 50% 50%,#000 0 ${p * seg}deg,transparent ${p * seg}deg ${seg}deg)`);
    }
    default:
      return { opacity: p };
  }
}

const DIRS = {
  l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1],
};

/**
 * Slide transition frame styles.
 * @returns {{inc: object, out: object, outOnTop?: boolean, back?: string}} CSS for incoming / outgoing slides.
 */
export function transitionFrame(tr, p) {
  const e = p < 0 ? 0 : p > 1 ? 1 : p;
  const s = e * e * (3 - 2 * e); // smoothstep for movement
  const dir = tr.dir;
  switch (tr.type) {
    case 'fade':
      if (tr.thruBlk) return e < 0.5 ? { out: { opacity: 1 - e * 2 }, inc: { opacity: 0 }, back: '#000' } : { out: { opacity: 0 }, inc: { opacity: (e - 0.5) * 2 }, back: '#000' };
      return { out: {}, inc: { opacity: e } };
    case 'cut':
      if (tr.thruBlk) return e < 0.5 ? { out: { opacity: 0 }, inc: { opacity: 0 }, back: '#000' } : { out: { opacity: 0 }, inc: {} };
      return { out: {}, inc: {} };
    case 'push': {
      const [dx, dy] = DIRS[dir || 'l'] || [-1, 0];
      return { out: { transform: `translate(${pc(dx * s)},${pc(dy * s)})` }, inc: { transform: `translate(${pc(-dx * (1 - s))},${pc(-dy * (1 - s))})` } };
    }
    case 'cover': {
      const [dx, dy] = DIRS[dir || 'l'] || [-1, 0];
      return { out: {}, inc: { transform: `translate(${pc(-dx * (1 - s))},${pc(-dy * (1 - s))})` } };
    }
    case 'pull': {
      const [dx, dy] = DIRS[dir || 'l'] || [-1, 0];
      return { out: { transform: `translate(${pc(dx * s)},${pc(dy * s)})` }, inc: {}, outOnTop: true };
    }
    case 'wipe': {
      const map = { l: 'left', r: 'right', u: 'up', d: 'down' };
      return { out: {}, inc: filterStyle('wipe(' + (map[dir || 'l'] || 'left') + ')', e) };
    }
    case 'split': {
      const vert = (tr.orient || 'horz') === 'vert';
      const f = (tr.dir || 'out') === 'in' ? (vert ? 'barn(inVertical)' : 'barn(inHorizontal)') : (vert ? 'barn(outVertical)' : 'barn(outHorizontal)');
      return { out: {}, inc: filterStyle(f, e) };
    }
    case 'blinds':
      return { out: {}, inc: filterStyle((tr.dir || 'horz') === 'vert' ? 'blinds(vertical)' : 'blinds(horizontal)', e) };
    case 'checker':
      return { out: {}, inc: filterStyle((tr.dir || 'horz') === 'vert' ? 'checkerboard(down)' : 'checkerboard(across)', e) };
    case 'randomBar':
      return { out: {}, inc: filterStyle((tr.dir || 'horz') === 'vert' ? 'randombar(vertical)' : 'randombar(horizontal)', e) };
    case 'circle':
      return { out: {}, inc: filterStyle('circle(out)', e) };
    case 'diamond':
      return { out: {}, inc: filterStyle('diamond(out)', e) };
    case 'plus':
      return { out: {}, inc: filterStyle('plus(out)', e) };
    case 'wedge':
      return { out: {}, inc: filterStyle('wedge', e) };
    case 'wheel':
      return { out: {}, inc: filterStyle('wheel(' + (tr.spokes || 4) + ')', e) };
    case 'wheelReverse':
      return { out: {}, inc: { ...filterStyle('wheel(' + (tr.spokes || 4) + ')', e), transform: 'scaleX(-1)' }, flipContent: true };
    case 'strips': {
      const map = { lu: 'upLeft', ru: 'upRight', ld: 'downLeft', rd: 'downRight' };
      return { out: {}, inc: filterStyle('strips(' + (map[dir || 'lu'] || 'upLeft') + ')', e) };
    }
    case 'dissolve':
    case 'glitter':
    case 'shred':
      return { out: {}, inc: filterStyle('dissolve', e) };
    case 'zoom':
      if ((dir || 'in') === 'out') return { out: { transform: `scale(${1 + s * 0.5})`, opacity: 1 - e }, inc: { opacity: e }, outOnTop: true };
      return { out: {}, inc: { transform: `scale(${0.3 + 0.7 * s})`, opacity: e } };
    case 'newsflash':
      return { out: {}, inc: { transform: `rotate(${720 * (1 - s)}deg) scale(${Math.max(0.001, s)})` } };
    case 'comb': {
      const vert = (dir || 'horz') === 'vert';
      const n = 8;
      // Alternating bands slide in from opposite sides.
      const off = (1 - s);
      const g = vert ? 'to right' : 'to bottom';
      return {
        out: {},
        inc: { ...maskStyle(`repeating-linear-gradient(${g},#000 0 ${pc(1 / n / 2)},transparent ${pc(1 / n / 2)} ${pc(1 / n)})`), transform: vert ? `translateY(${pc(off)})` : `translateX(${pc(off)})` },
        inc2: { ...maskStyle(`repeating-linear-gradient(${g},transparent 0 ${pc(1 / n / 2)},#000 ${pc(1 / n / 2)} ${pc(1 / n)})`), transform: vert ? `translateY(${pc(-off)})` : `translateX(${pc(-off)})` },
      };
    }
    case 'random':
      return transitionFrame({ ...tr, type: 'fade' }, p);
    case 'flash':
      return e < 0.5 ? { out: { opacity: 1 - e * 2 }, inc: { opacity: 0 }, back: '#fff' } : { out: { opacity: 0 }, inc: { opacity: (e - 0.5) * 2 }, back: '#fff' };
    case 'reveal': {
      const [dx] = DIRS[dir || 'l'] || [-1, 0];
      return e < 0.5 ? { out: { opacity: 1 - e * 2, transform: `translateX(${pc(dx * e * 0.2)})` }, inc: { opacity: 0 }, back: '#000' } : { out: { opacity: 0 }, inc: { opacity: (e - 0.5) * 2, transform: `translateX(${pc(-dx * (1 - e) * 0.2)})` }, back: '#000' };
    }
    case 'ripple':
      return { out: {}, inc: { ...filterStyle('circle(out)', e), opacity: Math.min(1, e * 1.5) } };
    case 'honeycomb':
    case 'flythrough':
    case 'warp':
      return { out: { transform: `scale(${1 + s * 0.6})`, opacity: 1 - e }, inc: { transform: `scale(${0.6 + 0.4 * s})`, opacity: e } };
    case 'vortex':
      return { out: { transform: `rotate(${s * 180}deg) scale(${1 - s * 0.7})`, opacity: 1 - e }, inc: { opacity: e }, outOnTop: true };
    case 'switch':
    case 'flip': {
      const sign = (dir || 'r') === 'l' ? -1 : 1;
      if (e < 0.5) return { out: { transform: `perspective(1600px) rotateY(${sign * s * 180}deg)`, backfaceVisibility: 'hidden' }, inc: { opacity: 0 }, back: '#000' };
      return { out: { opacity: 0 }, inc: { transform: `perspective(1600px) rotateY(${-sign * (1 - s) * 180}deg)`, backfaceVisibility: 'hidden' }, back: '#000' };
    }
    case 'prism':
    case 'ferris':
    case 'conveyor':
    case 'gallery':
    case 'pan':
    case 'window': {
      const sign = (dir || 'l') === 'r' ? 1 : -1;
      return {
        out: { transform: `perspective(1600px) translateX(${pc(sign * s)}) rotateY(${sign * s * 70}deg)`, transformOrigin: sign > 0 ? '0 50%' : '100% 50%' },
        inc: { transform: `perspective(1600px) translateX(${pc(-sign * (1 - s))}) rotateY(${-sign * (1 - s) * 70}deg)`, transformOrigin: sign > 0 ? '100% 50%' : '0 50%' },
        back: '#000',
      };
    }
    case 'doors':
      return {
        out: { ...filterStyle('barn(inVertical)', 1 - e), transform: `perspective(1600px) scale(${1 + s * 0.2})`, opacity: 1 - e * 0.6 },
        inc: { transform: `scale(${0.85 + 0.15 * s})` },
        outOnTop: true,
        back: '#000',
      };
    case 'morph':
      return { out: { opacity: 1 - e }, inc: { opacity: e }, morph: true };
    default:
      if (tr.type && tr.type.startsWith('prst:')) return prstTransition(tr.type.slice(5), e, s, tr);
      return { out: {}, inc: { opacity: e } };
  }
}

function prstTransition(name, e, s, tr) {
  switch (name) {
    case 'curtains':
      return { out: { ...filterStyle('barn(inVertical)', 1 - e), transform: `scaleY(${1 + s * 0.05})` }, inc: {}, outOnTop: true };
    case 'wind': {
      const sign = tr.invX ? 1 : -1;
      return { out: { transform: `translateX(${pc(sign * s * 1.2)}) skewX(${sign * s * 20}deg)`, opacity: 1 - e }, inc: { opacity: Math.min(1, e * 1.4) }, outOnTop: true };
    }
    case 'prestige':
      return { out: { transform: `perspective(1600px) rotateX(${s * 70}deg)`, transformOrigin: '50% 100%', opacity: 1 - e }, inc: { opacity: e } };
    case 'fracture':
    case 'crush':
      return { out: { transform: `scale(${1 - s * 0.9}) rotate(${s * 15}deg)`, opacity: 1 - e }, inc: { opacity: Math.min(1, e * 1.3) }, outOnTop: true };
    case 'peelOff':
    case 'pageCurlSingle':
    case 'pageCurlDouble':
      return { out: { transform: `perspective(1600px) rotateY(${-s * 100}deg)`, transformOrigin: '0 50%', opacity: 1 - e * 0.4 }, inc: {}, outOnTop: true };
    case 'airplane':
    case 'origami':
      return { out: { transform: `perspective(1600px) translate(${pc(s * 0.8)},${pc(-s * 0.8)}) rotateZ(${s * 25}deg) scale(${1 - s * 0.7})`, opacity: 1 - e }, inc: {}, outOnTop: true };
    case 'fallOver':
    default:
      return { out: { transform: `perspective(1600px) rotateX(${-s * 90}deg)`, transformOrigin: '50% 100%', opacity: 1 - e * 0.5 }, inc: {}, outOnTop: true };
  }
}

export const TRANSITION_TYPES = ['cut', 'fade', 'push', 'cover', 'pull', 'wipe', 'split', 'blinds', 'checker', 'randomBar', 'circle', 'diamond', 'plus',
  'wedge', 'wheel', 'wheelReverse', 'strips', 'dissolve', 'glitter', 'shred', 'zoom', 'newsflash', 'comb', 'random', 'flash', 'reveal', 'ripple',
  'honeycomb', 'flythrough', 'warp', 'vortex', 'switch', 'flip', 'prism', 'ferris', 'conveyor', 'gallery', 'pan', 'window', 'doors', 'morph',
  'prst:curtains', 'prst:wind', 'prst:prestige', 'prst:fracture', 'prst:crush', 'prst:peelOff', 'prst:pageCurlSingle', 'prst:pageCurlDouble',
  'prst:airplane', 'prst:origami', 'prst:fallOver'];

export const FILTERS = ['fade', 'wipe(down)', 'wipe(up)', 'wipe(left)', 'wipe(right)', 'barn(inVertical)', 'barn(inHorizontal)', 'barn(outVertical)',
  'barn(outHorizontal)', 'blinds(horizontal)', 'blinds(vertical)', 'checkerboard(across)', 'checkerboard(down)', 'box(in)', 'box(out)',
  'circle(in)', 'circle(out)', 'diamond(in)', 'diamond(out)', 'dissolve', 'plus(in)', 'plus(out)', 'randombar(horizontal)',
  'randombar(vertical)', 'slide(fromTop)', 'slide(fromBottom)', 'slide(fromLeft)', 'slide(fromRight)', 'strips(downLeft)', 'strips(upLeft)',
  'strips(downRight)', 'strips(upRight)', 'wedge', 'wheel(1)', 'wheel(2)', 'wheel(3)', 'wheel(4)', 'wheel(8)'];

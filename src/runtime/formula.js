/**
 * Evaluator for PowerPoint animation formulas (tav values and fmla).
 *
 * Grammar: expressions with + - * / ^ %, unary minus, parentheses,
 * functions (abs acos asin atan ceil cos cosh deg exp floor ln max min rad
 * rand sin sinh sqrt tan tanh), constants (pi, e) and variables
 * ($, ppt_x, ppt_y, ppt_w, ppt_h, ppt_r, ppt_c, #ppt_x ...).
 */

const FUNCS = {
  abs: Math.abs, acos: Math.acos, asin: Math.asin, atan: Math.atan, ceil: Math.ceil, cos: Math.cos,
  cosh: Math.cosh, deg: (x) => (x * 180) / Math.PI, exp: Math.exp, floor: Math.floor, ln: Math.log,
  max: Math.max, min: Math.min, rad: (x) => (x * Math.PI) / 180, rand: (x) => Math.random() * (x === undefined ? 1 : x),
  sin: Math.sin, sinh: Math.sinh, sqrt: Math.sqrt, tan: Math.tan, tanh: Math.tanh,
};

const cache = new Map();

function tokenize(s) {
  const out = [];
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (c === ' ' || c === '\t' || c === '\n') { i++; continue; }
    if (/[0-9.]/.test(c)) {
      const m = /^[0-9]*\.?[0-9]+(?:[eE][-+]?[0-9]+)?|^[0-9]+\.?/.exec(s.slice(i));
      out.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    if (/[#$A-Za-z_]/.test(c)) {
      const m = /^[#$]?[A-Za-z_]*[A-Za-z0-9_.]*/.exec(s.slice(i));
      const name = m[0] || c;
      out.push({ t: 'id', v: name });
      i += name.length || 1;
      continue;
    }
    out.push({ t: 'op', v: c });
    i++;
  }
  return out;
}

/** Compile a formula into a function of a variable map. */
export function compileFormula(src) {
  if (cache.has(src)) return cache.get(src);
  const toks = tokenize(String(src));
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (v) => {
    const t = next();
    if (!t || t.v !== v) throw new Error('formula: expected ' + v + ' in ' + src);
  };
  function primary() {
    const t = next();
    if (!t) throw new Error('formula: unexpected end in ' + src);
    if (t.t === 'num') return () => t.v;
    if (t.t === 'op' && t.v === '(') {
      const e = expr();
      expect(')');
      return e;
    }
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) {
      const e = unary();
      return t.v === '-' ? (v) => -e(v) : e;
    }
    if (t.t === 'id') {
      const name = t.v;
      const lname = name.toLowerCase();
      if (peek() && peek().v === '(' && FUNCS[lname]) {
        next();
        const args = [];
        if (!(peek() && peek().v === ')')) {
          args.push(expr());
          while (peek() && peek().v === ',') { next(); args.push(expr()); }
        }
        expect(')');
        const f = FUNCS[lname];
        return (v) => f(...args.map((a) => a(v)));
      }
      if (lname === 'pi') return () => Math.PI;
      if (lname === 'e') return () => Math.E;
      return (v) => {
        const x = v[name];
        if (x !== undefined) return x;
        const y = v[lname];
        return y === undefined ? 0 : y;
      };
    }
    throw new Error('formula: unexpected token ' + t.v + ' in ' + src);
  }
  function unary() {
    const t = peek();
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) {
      next();
      const e = unary();
      return t.v === '-' ? (v) => -e(v) : e;
    }
    return power();
  }
  function power() {
    const base = primary();
    const t = peek();
    if (t && t.t === 'op' && t.v === '^') {
      next();
      const ex = unary();
      return (v) => Math.pow(base(v), ex(v));
    }
    return base;
  }
  function term() {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || (t.v !== '*' && t.v !== '/' && t.v !== '%')) break;
      next();
      const right = unary();
      const l = left;
      if (t.v === '*') left = (v) => l(v) * right(v);
      else if (t.v === '/') left = (v) => { const d = right(v); return d === 0 ? 0 : l(v) / d; };
      else left = (v) => l(v) % right(v);
    }
    return left;
  }
  function expr() {
    let left = term();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || (t.v !== '+' && t.v !== '-')) break;
      next();
      const right = term();
      const l = left;
      left = t.v === '+' ? (v) => l(v) + right(v) : (v) => l(v) - right(v);
    }
    return left;
  }
  let fn;
  try {
    fn = expr();
    if (p < toks.length) throw new Error('formula: trailing tokens in ' + src);
  } catch (e) {
    fn = () => NaN;
  }
  cache.set(src, fn);
  return fn;
}

export function evalFormula(src, vars) {
  if (typeof src === 'number') return src;
  if (src === null || src === undefined || src === '') return NaN;
  const n = Number(src);
  if (!Number.isNaN(n)) return n;
  return compileFormula(src)(vars);
}

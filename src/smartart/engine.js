/**
 * SmartArt layout engine.
 *
 * PowerPoint 2010 and later store a pre-rendered drawing of every diagram;
 * files written by PowerPoint 2007 and by other producers only carry the
 * data model and the layout definition. This module interprets the layout
 * definition (dgm:layoutDef) against the data model and produces DrawingML
 * shapes (dsp:sp nodes) that the regular shape parser renders.
 *
 * Supported: forEach / choose / if (cnt, pos, depth, maxDepth, var), all
 * data axes, presOf, constraints with references, and the algorithms
 * composite, lin, snake, cycle, pyra, hierRoot / hierChild, tx, conn, sp.
 * Colours and line / fill / effect styles come from the diagram's colours
 * and quick style parts.
 */
import { XNode } from '../util/xml.js';

const EMU = 9525; // per px
const MM = 3.7795; // px per mm
const PT = 4 / 3; // px per pt

// ---- XML helpers ----------------------------------------------------------------

function node(name, attrs = {}, children = []) {
  const n = new XNode(name, attrs, null);
  for (const c of children) if (c) { c.parent = n; n.children.push(c); }
  return n;
}

function clone(n, rename) {
  const c = new XNode(rename ? rename(n.name) : n.name, { ...n.attrs }, null);
  c.text = n.text;
  for (const ch of n.children) { const k = clone(ch, rename); k.parent = c; c.children.push(k); }
  return c;
}

// ---- data model -----------------------------------------------------------------

/** Parse dgm:dataModel into a tree of nodes with their transitions. */
export function parseDataModel(root) {
  const pts = new Map();
  for (const p of root.find('dgm:pt')) {
    const prSet = p.child('dgm:prSet');
    pts.set(p.attrs.modelId, {
      id: p.attrs.modelId, type: p.attr('type', 'node'), pr: prSet ? prSet.attrs : {}, prSet,
      spPr: p.child('dgm:spPr'), t: p.child('dgm:t'),
    });
  }
  const items = new Map();
  for (const pt of pts.values()) {
    if (pt.type === 'doc' || pt.type === 'node' || pt.type === 'asst') items.set(pt.id, { kind: 'node', type: pt.type, pt, children: [], parent: null });
  }
  const pres = new Map();
  for (const pt of pts.values()) {
    if (pt.type === 'pres' && pt.pr.presAssocID !== undefined) pres.set(pt.pr.presAssocID + '|' + (pt.pr.presName || ''), pt);
  }
  const cxns = root.find('dgm:cxn').map((c) => ({
    type: c.attr('type', 'parOf'), src: c.attrs.srcId, dest: c.attrs.destId, srcOrd: c.int('srcOrd', 0),
    parTrans: c.attrs.parTransId, sibTrans: c.attrs.sibTransId,
  })).filter((c) => c.type === 'parOf').sort((a, b) => a.srcOrd - b.srcOrd);
  for (const c of cxns) {
    const p = items.get(c.src), ch = items.get(c.dest);
    if (!p || !ch || ch.parent) continue;
    ch.parent = p;
    p.children.push(ch);
    const tr = (id, type) => (pts.get(id) ? { kind: 'trans', type, pt: pts.get(id), owner: ch } : null);
    ch.parTrans = tr(c.parTrans, 'parTrans');
    ch.sibTrans = tr(c.sibTrans, 'sibTrans');
  }
  let doc = null;
  for (const it of items.values()) if (it.type === 'doc') { doc = it; break; }
  return { doc, pts, pres };
}

// ---- axes and conditions ----------------------------------------------------------

function matchType(it, t) {
  if (!it) return false;
  switch (t || 'all') {
    case 'all': return true;
    case 'node': return it.kind === 'node' && it.type !== 'doc';
    case 'asst': return it.type === 'asst';
    case 'nonAsst': return it.kind === 'node' && it.type === 'node';
    case 'doc': return it.type === 'doc';
    case 'parTrans': case 'sibTrans': return it.type === t;
    case 'norm': return it.type === 'node';
    case 'nonNorm': return it.type !== 'node';
    default: return true;
  }
}

/** Children interleaved with their transitions: p1 n1 s1 p2 n2 s2 ... */
function siblingRow(parent, hideLastTrans) {
  const row = [];
  parent.children.forEach((c, i) => {
    if (c.parTrans) row.push(c.parTrans);
    row.push(c);
    if (c.sibTrans && !(hideLastTrans && i === parent.children.length - 1)) row.push(c.sibTrans);
  });
  return row;
}

function descendants(it, out = []) {
  if (it.kind !== 'node') return out;
  for (const c of it.children) { out.push(c); descendants(c, out); }
  return out;
}

function stepAxis(it, axis, type, hideLast) {
  const owner = it.kind === 'trans' ? it.owner : it;
  let list;
  switch (axis) {
    case 'self': list = [it]; break;
    case 'ch':
      if (it.kind !== 'node') return [];
      if (type === 'sibTrans' || type === 'parTrans' || type === 'all') list = siblingRow(it, hideLast);
      else list = it.children;
      break;
    case 'des': list = descendants(it); break;
    case 'desOrSelf': list = [it, ...descendants(it)]; break;
    case 'par': list = it.kind === 'trans' ? [it.owner.parent] : [it.parent]; break;
    case 'ancst': case 'ancstOrSelf': {
      list = axis === 'ancstOrSelf' ? [it] : [];
      for (let p = owner.parent; p; p = p.parent) list.push(p);
      break;
    }
    case 'root': { let p = owner; while (p.parent) p = p.parent; list = [p]; break; }
    case 'followSib': case 'precedSib': {
      const parent = owner.parent;
      if (!parent) return [];
      const row = siblingRow(parent, hideLast);
      const i = row.indexOf(it);
      list = axis === 'followSib' ? row.slice(i + 1) : row.slice(0, Math.max(0, i));
      break;
    }
    case 'none': return [];
    default: list = [it];
  }
  return list.filter((x) => matchType(x, type));
}

/** Items selected by a forEach / presOf / if: axis and ptType may be lists. */
function select(it, el) {
  if (!it) return [];
  const axes = (el.attrs.axis || 'self').split(/\s+/);
  const types = (el.attrs.ptType || 'all').split(/\s+/);
  const sts = (el.attrs.st || '').split(/\s+/), cnts = (el.attrs.cnt || '').split(/\s+/), steps = (el.attrs.step || '').split(/\s+/);
  const hideLast = el.attrs.hideLastTrans !== '0' && el.attrs.hideLastTrans !== 'false';
  let cur = [it];
  axes.forEach((axis, k) => {
    const next = [];
    for (const c of cur) next.push(...stepAxis(c, axis, types[k] || types[types.length - 1], hideLast));
    let list = next;
    const st = Number(sts[k] || 1), cnt = Number(cnts[k] || 0), step = Number(steps[k] || 1) || 1;
    if (st > 1 || st < 0 || cnt || step !== 1) {
      const out = [];
      let i = st < 0 ? list.length + st : st - 1;
      while (i >= 0 && i < list.length && (!cnt || out.length < cnt)) { out.push(list[i]); i += step; }
      list = out;
    }
    cur = list;
  });
  return cur;
}

function depthOf(it) {
  let d = 0;
  for (let p = it.kind === 'trans' ? it.owner : it; p && p.parent; p = p.parent) d++;
  return d;
}

function maxDepth(it) {
  if (it.kind !== 'node' || !it.children.length) return 0;
  return 1 + Math.max(...it.children.map(maxDepth));
}

function compare(a, op, b) {
  const na = Number(a), nb = Number(b);
  const num = a !== '' && b !== '' && Number.isFinite(na) && Number.isFinite(nb);
  const x = num ? na : String(a), y = num ? nb : String(b);
  switch (op) {
    case 'equ': return x === y;
    case 'neq': return x !== y;
    case 'gt': return x > y;
    case 'lt': return x < y;
    case 'gte': return x >= y;
    case 'lte': return x <= y;
    default: return false;
  }
}

// ---- layout definition ----------------------------------------------------------------

let lnodeSeq = 0;

class LNode {
  constructor(name, data, parent) {
    this.id = ++lnodeSeq;
    this.name = name;
    this.data = data;
    this.parent = parent;
    this.alg = { type: 'composite', params: {} };
    this.shape = null;
    this.presOf = null;
    this.constrs = [];
    this.rules = [];
    this.children = [];
    this.vars = parent ? { ...parent.vars } : {};
    this.styleLbl = '';
    this.box = { x: 0, y: 0, w: 0, h: 0 };
    this.fixed = {};
  }
}

const VAR_DEFAULTS = { dir: 'norm', hierBranch: 'std', bulletEnabled: '0', chMax: '-1', chPref: '-1', orgChart: '0', animOne: 'one', animLvl: 'none', resizeHandles: 'rel' };

class Builder {
  constructor(dm, layoutDef) {
    this.dm = dm;
    this.named = new Map();
    for (const f of layoutDef.find('dgm:forEach')) if (f.attrs.name) this.named.set(f.attrs.name, f);
    this.count = 0;
  }

  vars(el, ln) {
    for (const v of el.children) ln.vars[v.local] = v.attr('val', VAR_DEFAULTS[v.local] || '1');
  }

  presVars(ln) {
    if (!ln.data || !ln.data.pt) return;
    const p = this.dm.pres.get(ln.data.pt.id + '|' + ln.name);
    const lv = p && p.prSet && p.prSet.child('dgm:presLayoutVars');
    if (lv) this.vars(lv, ln);
    // A shape resized by the user (fractions of 100000).
    if (p && (p.pr.custScaleX || p.pr.custScaleY)) ln.custScale = [Number(p.pr.custScaleX || 100000) / 100000, Number(p.pr.custScaleY || 100000) / 100000];
  }

  test(cond, ln) {
    const it = ln.data;
    const f = cond.attrs.func;
    let v;
    switch (f) {
      case 'cnt': v = select(it, cond).length; break;
      case 'pos': case 'revPos': case 'posEven': case 'posOdd': {
        const owner = it && (it.kind === 'trans' ? it.owner : it);
        const sibs = owner && owner.parent ? owner.parent.children : [owner];
        const i = sibs.indexOf(owner) + 1;
        v = f === 'pos' ? i : f === 'revPos' ? sibs.length - i + 1 : f === 'posEven' ? (i % 2 === 0 ? 1 : 0) : (i % 2 === 1 ? 1 : 0);
        break;
      }
      case 'depth': v = it ? depthOf(it) : 0; break;
      case 'maxDepth': v = it ? maxDepth(it) : 0; break;
      case 'var': v = ln.vars[cond.attrs.arg] !== undefined ? ln.vars[cond.attrs.arg] : VAR_DEFAULTS[cond.attrs.arg] || ''; break;
      default: v = 0;
    }
    return compare(v, cond.attrs.op || 'equ', cond.attrs.val !== undefined ? cond.attrs.val : '');
  }

  /** Apply the content of a layoutNode-like element to ln. */
  content(el, ln) {
    for (const c of el.children) {
      switch (c.local) {
        case 'varLst': this.vars(c, ln); this.presVars(ln); break; // user settings win
        case 'alg': ln.alg = { type: c.attr('type', 'composite'), params: Object.fromEntries(c.all('dgm:param').map((p) => [p.attrs.type, p.attrs.val])) }; break;
        case 'shape': ln.shape = {
          type: c.attr('type', ''), rot: c.num('rot', 0), hideGeom: c.bool('hideGeom', false), zOrderOff: c.int('zOrderOff', 0),
          adj: c.child('dgm:adjLst') ? c.child('dgm:adjLst').all('dgm:adj').map((a) => [a.int('idx', 1), a.num('val', 0)]) : [],
        }; break;
        case 'presOf': ln.presOf = c; break;
        case 'constrLst': ln.constrs.push(...c.all('dgm:constr').map((k) => ({
          type: k.attrs.type, for: k.attr('for', 'self'), forName: k.attrs.forName || '', ptType: k.attrs.ptType || '',
          refType: k.attr('refType', 'none'), refFor: k.attr('refFor', 'self'), refForName: k.attrs.refForName || '', refPtType: k.attrs.refPtType || '',
          op: k.attr('op', 'none'), fact: k.num('fact', 1), val: k.num('val', 0),
          // op="equ" without a value or reference only asks for equal values.
          equal: k.attrs.op === 'equ' && k.attrs.val === undefined && !k.attrs.refType,
        }))); break;
        case 'ruleLst': ln.rules.push(...c.all('dgm:rule').map((r) => r.attrs)); break;
        case 'choose': {
          let done = false;
          for (const br of c.children) {
            if (br.local === 'if' && !done && this.test(br, ln)) { this.content(br, ln); done = true; }
            else if (br.local === 'else' && !done) { this.content(br, ln); done = true; }
          }
          break;
        }
        case 'forEach': {
          const def = c.attrs.ref ? this.named.get(c.attrs.ref) || c : c;
          for (const it of select(ln.data, def)) {
            const saved = ln.data;
            ln.data = it;
            this.content(def, ln);
            ln.data = saved;
          }
          break;
        }
        case 'layoutNode': {
          if (++this.count > 2000) break;
          const k = new LNode(c.attr('name', ''), ln.data, ln);
          k.styleLbl = c.attr('styleLbl', '');
          k.chOrder = c.attr('chOrder', 'b');
          this.presVars(k);
          this.content(c, k);
          ln.children.push(k);
          break;
        }
        default: break;
      }
    }
  }

  build(layoutDef) {
    const top = layoutDef.child('dgm:layoutNode');
    const root = new LNode(top ? top.attr('name', '') : '', this.dm.doc, null);
    if (top) {
      this.presVars(root);
      this.content(top, root);
    }
    return root;
  }
}

// ---- constraints ----------------------------------------------------------------------

const DIMENSIONS = new Set(['w', 'h', 'l', 't', 'r', 'b', 'ctrX', 'ctrY', 'wOff', 'hOff', 'sp', 'sibSp', 'secSibSp', 'begPad', 'endPad',
  'lMarg', 'rMarg', 'tMarg', 'bMarg', 'connDist', 'diam', 'bendDist', 'stemThick', 'alignOff', 'hArH', 'wArH', 'userA', 'userB']);

function nodeType(ln) {
  return ln.data ? ln.data.type : '';
}

function applies(c, child) {
  if (c.forName && c.forName !== child.name) return false;
  if (c.ptType && !matchType(child.data, c.ptType)) return false;
  return true;
}

/** Constraints that apply to a layout node: its own, its parent's (for="ch") and ancestors' (for="des"). */
function constraintsFor(ln, type) {
  const out = [];
  for (const c of ln.constrs) if (c.type === type && c.for === 'self') out.push([c, ln]);
  if (ln.parent) for (const c of ln.parent.constrs) if (c.type === type && (c.for === 'ch' || c.for === 'des') && applies(c, ln)) out.push([c, ln.parent]);
  for (let a = ln.parent && ln.parent.parent; a; a = a.parent) {
    for (const c of a.constrs) if (c.type === type && c.for === 'des' && applies(c, ln)) out.push([c, a]);
  }
  return out;
}

function findDes(ln, name, ptType) {
  for (const k of ln.children) {
    if ((!name || k.name === name) && (!ptType || matchType(k.data, ptType))) return k;
    const d = findDes(k, name, ptType);
    if (d) return d;
  }
  return null;
}

/**
 * Value of a constraint type for a layout node. Equality constraints give the
 * value (a zero without reference means "free"); lte / gte ones bound it.
 * Returns undefined when nothing constrains it.
 */
function cval(ln, type, seen = new Set()) {
  if (ln.fixed[type] !== undefined) return ln.fixed[type];
  const key = ln.id + ':' + type;
  if (seen.has(key)) return undefined;
  seen.add(key);
  let eq, lo = -Infinity, hi = Infinity;
  for (const [c, owner] of constraintsFor(ln, type)) {
    if (c.equal) continue;
    const v = evalConstr(c, owner, ln, seen);
    if (v === undefined || !Number.isFinite(v)) continue;
    if (c.op === 'lte') hi = Math.min(hi, v);
    else if (c.op === 'gte') lo = Math.max(lo, v);
    else if (eq === undefined && !(v === 0 && c.refType === 'none' && (type === 'w' || type === 'h'))) eq = v;
  }
  seen.delete(key);
  if (eq === undefined) {
    if (hi < Infinity) eq = hi;
    else if (lo > -Infinity) eq = lo;
    else return undefined;
  }
  return Math.max(lo, Math.min(hi, eq));
}

function evalConstr(c, owner, target, seen) {
  const dim = DIMENSIONS.has(c.type);
  if (c.refType === 'none') return c.val * (dim ? MM : 1);
  let ref;
  let refNode = owner;
  if (c.refFor === 'ch') refNode = owner.children.find((k) => (!c.refForName || k.name === c.refForName) && (!c.refPtType || matchType(k.data, c.refPtType)));
  else if (c.refFor === 'des') refNode = findDes(owner, c.refForName, c.refPtType);
  if (!refNode) return c.val ? c.val * (dim ? MM : 1) : undefined;
  if (c.refType === 'w' || c.refType === 'h') {
    // The constraint owner's own box is known; other nodes use their sizes.
    ref = refNode === owner && refNode !== target && (owner.box.w || owner.box.h) ? owner.box[c.refType] : sizeOf(refNode, c.refType, seen);
  } else {
    ref = cval(refNode, c.refType, seen);
    if (ref === undefined) ref = DEFAULTS[c.refType];
    // Font sizes are in points: dimensions derived from them too.
    if (ref !== undefined && dim && !DIMENSIONS.has(c.refType)) ref *= PT;
  }
  if (ref === undefined || !Number.isFinite(ref)) return undefined;
  return ref * c.fact;
}

const DEFAULTS = { primFontSz: 65, secFontSz: 65 * 0.78 };

/** Preferred width / height of a node (defaults to its parent's box). */
function sizeOf(ln, dim, seen = new Set()) {
  const v = cval(ln, dim, seen);
  if (v !== undefined && Number.isFinite(v)) return v;
  if (ln.box[dim]) return ln.box[dim];
  return ln.parent ? ln.parent.box[dim] : 0;
}

// ---- algorithms ----------------------------------------------------------------------

function visibleKids(ln) {
  return ln.children;
}

function place(ln, x, y, w, h) {
  ln.box = { x, y, w: Math.max(0, w), h: Math.max(0, h) };
}

function layoutComposite(ln) {
  for (const k of visibleKids(ln)) {
    const g = (t) => { const v = cval(k, t); return v !== undefined && Number.isFinite(v) ? v : undefined; };
    let w = g('w'), h = g('h');
    const l = g('l'), t = g('t'), r = g('r'), b = g('b'), cx = g('ctrX'), cy = g('ctrY');
    if (w === undefined) w = l !== undefined && r !== undefined ? r - l : ln.box.w;
    if (h === undefined) h = t !== undefined && b !== undefined ? b - t : ln.box.h;
    const x = l !== undefined ? l : r !== undefined ? r - w : cx !== undefined ? cx - w / 2 : 0;
    const y = t !== undefined ? t : b !== undefined ? b - h : cy !== undefined ? cy - h / 2 : 0;
    k.fixed.w = w; k.fixed.h = h;
    place(k, ln.box.x + x, ln.box.y + y, w, h);
  }
  // Composites larger than their slot are scaled down to fit it.
  fitChildren(ln);
}

function fitChildren(ln) {
  const kids = ln.children;
  if (!kids.length) return;
  const x0 = Math.min(...kids.map((k) => k.box.x)), y0 = Math.min(...kids.map((k) => k.box.y));
  const x1 = Math.max(...kids.map((k) => k.box.x + k.box.w)), y1 = Math.max(...kids.map((k) => k.box.y + k.box.h));
  const s = Math.min(1, ln.box.w / Math.max(1e-6, x1 - Math.min(x0, ln.box.x)), ln.box.h / Math.max(1e-6, y1 - Math.min(y0, ln.box.y)));
  if (s >= 0.999 || !(s > 0)) return;
  for (const k of kids) {
    k.box = { x: ln.box.x + (k.box.x - ln.box.x) * s, y: ln.box.y + (k.box.y - ln.box.y) * s, w: k.box.w * s, h: k.box.h * s };
    k.fixed.w = k.box.w; k.fixed.h = k.box.h;
  }
}

/** Containers without anything to draw take no room in linear layouts. */
function isEmpty(ln) {
  if (ln.emptyCache !== undefined) return ln.emptyCache;
  const drawn = ln.shape && ln.shape.type && ln.shape.type !== 'none' && !ln.shape.hideGeom;
  const text = ln.alg.type === 'tx' && textOf(ln).length > 0;
  ln.emptyCache = ln.alg.type !== 'sp' && !drawn && !text && ln.children.every(isEmpty);
  return ln.emptyCache;
}

function layoutLin(ln) {
  const p = ln.alg.params;
  const dir = p.linDir || 'fromL';
  const horiz = dir === 'fromL' || dir === 'fromR';
  const rev = dir === 'fromR' || dir === 'fromB';
  const kids = visibleKids(ln);
  if (!kids.length) return;
  const along = horiz ? 'w' : 'h', across = horiz ? 'h' : 'w';
  const A = ln.box[along], C = ln.box[across];
  const sp = cval(ln, 'sp') || 0;
  const beg = cval(ln, 'begPad') || 0, end = cval(ln, 'endPad') || 0;
  let sizes = kids.map((k) => (isEmpty(k) ? 0 : sizeOf(k, along)));
  const total = sizes.reduce((a, b) => a + b, 0) + sp * (kids.length - 1) + beg + end;
  const s = total > 0 ? Math.min(1, A / total) : 1;
  sizes = sizes.map((v) => v * s);
  kids.forEach((k, i) => { k.fixed[along] = sizes[i]; });
  const crossSizes = kids.map((k) => Math.min(C, sizeOf(k, across)));
  // Cross sizes derived from the scaled along size keep their proportion.
  const used = sizes.reduce((a, b) => a + b, 0) + (sp * (kids.length - 1) + beg + end) * s;
  let pos = (A - used) / 2 + beg * s;
  const align = horiz ? (p.nodeVertAlign || 'mid') : (p.nodeHorzAlign || 'ctr');
  kids.forEach((k, i) => {
    const a = sizes[i], c = crossSizes[i];
    const off = align === 't' || align === 'l' ? 0 : align === 'b' || align === 'r' ? C - c : (C - c) / 2;
    const at = rev ? A - pos - a : pos;
    if (horiz) place(k, ln.box.x + at, ln.box.y + off, a, c); else place(k, ln.box.x + off, ln.box.y + at, c, a);
    k.fixed[across] = c;
    pos += a + sp * s;
  });
}

function layoutSnake(ln) {
  const p = ln.alg.params;
  const kids = visibleKids(ln);
  const nodes = kids.filter((k) => k.alg.type !== 'sp' && nodeType(k) !== 'sibTrans');
  const n = nodes.length;
  if (!n) return;
  const W = ln.box.w, H = ln.box.h;
  const nw = sizeOf(nodes[0], 'w'), nh = sizeOf(nodes[0], 'h');
  let sp = cval(ln, 'sp');
  if (sp === undefined) sp = 0;
  const vertical = p.flowDir === 'column';
  let best = { s: 0, cols: 1 };
  for (let cols = 1; cols <= n; cols++) {
    const rows = Math.ceil(n / cols);
    const c = vertical ? rows : cols, r = vertical ? cols : rows;
    const s = Math.min(W / (c * nw + (c - 1) * sp), H / (r * nh + (r - 1) * sp));
    if (s > best.s + 1e-9) best = { s, cols };
  }
  const s = best.s;
  const rowsN = Math.ceil(n / best.cols);
  const cw = nw * s, ch = nh * s, gap = sp * s;
  const gridW = (vertical ? rowsN : best.cols) * cw + ((vertical ? rowsN : best.cols) - 1) * gap;
  const gridH = (vertical ? best.cols : rowsN) * ch + ((vertical ? best.cols : rowsN) - 1) * gap;
  const ox = ln.box.x + (W - gridW) / 2, oy = ln.box.y + (H - gridH) / 2;
  const fromRight = (p.grDir || 'tL').endsWith('R'), fromBottom = (p.grDir || 'tL').startsWith('b');
  const snake = p.contDir === 'revDir';
  nodes.forEach((k, i) => {
    let line = Math.floor(i / best.cols), pos = i % best.cols;
    const inLine = Math.min(best.cols, n - line * best.cols);
    if (snake && line % 2 === 1) pos = inLine - 1 - pos;
    // Centre an incomplete last line.
    const shift = p.off === 'off' ? 0 : ((best.cols - inLine) * ((vertical ? ch : cw) + gap)) / 2;
    let col = pos, row = line;
    if (vertical) { col = line; row = pos; }
    let x = vertical ? col * (cw + gap) : col * (cw + gap) + shift;
    let y = vertical ? row * (ch + gap) + shift : row * (ch + gap);
    if (fromRight) x = gridW - x - cw;
    if (fromBottom) y = gridH - y - ch;
    k.fixed.w = cw; k.fixed.h = ch;
    place(k, ox + x, oy + y, cw, ch);
  });
  // Transitions (connector arrows) sit between consecutive nodes.
  const others = kids.filter((k) => !nodes.includes(k));
  others.forEach((k, i) => {
    const a = nodes[i], b = nodes[i + 1];
    if (!a || !b || k.alg.type === 'sp') { place(k, 0, 0, 0, 0); return; }
    connectBetween(k, a.box, b.box);
  });
}

/** Place a connector node between two boxes, pointing from a to b. */
function connectBetween(k, a, b) {
  const ax = a.x + a.w / 2, ay = a.y + a.h / 2, bx = b.x + b.w / 2, by = b.y + b.h / 2;
  const horiz = Math.abs(bx - ax) >= Math.abs(by - ay);
  const gap = horiz ? Math.abs(bx - ax) - (a.w + b.w) / 2 : Math.abs(by - ay) - (a.h + b.h) / 2;
  const len = Math.max(2, gap * 0.6);
  const th = Math.max(2, Math.min(len * 0.8, (horiz ? a.h : a.w) * 0.3));
  k.rot = (Math.atan2(by - ay, bx - ax) * 180) / Math.PI;
  const mx = (ax + bx) / 2 + (horiz ? ((a.w - b.w) / 4) * Math.sign(bx - ax) : 0);
  const my = (ay + by) / 2 + (horiz ? 0 : ((a.h - b.h) / 4) * Math.sign(by - ay));
  k.fixed.w = len; k.fixed.h = th;
  place(k, mx - len / 2, my - th / 2, len, th);
}

function layoutCycle(ln) {
  const p = ln.alg.params;
  const kids = visibleKids(ln);
  let nodes = kids.filter((k) => k.alg.type !== 'conn' && k.alg.type !== 'sp' && nodeType(k) !== 'sibTrans');
  const conns = kids.filter((k) => !nodes.includes(k));
  const W = ln.box.w, H = ln.box.h;
  const cx = ln.box.x + W / 2, cy = ln.box.y + H / 2;
  let center = null;
  if (p.ctrShpMap === 'fNode' && nodes.length) { center = nodes[0]; nodes = nodes.slice(1); }
  const n = nodes.length;
  if (!n && !center) return;
  const st = Number(p.stAng || 0), span = Number(p.spanAng || 360);
  const nw = n ? sizeOf(nodes[0], 'w') : 0, nh = n ? sizeOf(nodes[0], 'h') : 0;
  // Radius so that the ring of nodes fits into the box.
  const step = n > 1 ? (span >= 360 ? 360 / n : span / (n - 1)) : 0;
  let size = Math.max(nw, nh);
  let r = Math.min(W, H) / 2 - size / 2;
  // Shrink nodes when neighbours would overlap.
  if (n > 1) {
    const chord = 2 * r * Math.sin((Math.PI * step) / 360);
    if (chord < size * 1.1) {
      const k = Math.min(W, H) / 2 / (size / 2 + (size * 1.1) / (2 * Math.sin((Math.PI * step) / 360)));
      size *= k;
      r = Math.min(W, H) / 2 - size / 2;
    }
  }
  const sw = nw ? (size * nw) / Math.max(nw, nh) : 0, sh = nh ? (size * nh) / Math.max(nw, nh) : 0;
  const centers = [];
  nodes.forEach((k, i) => {
    const a = ((st + i * step) * Math.PI) / 180;
    const x = cx + r * Math.sin(a), y = cy - r * Math.cos(a);
    centers.push([x, y, a]);
    k.fixed.w = sw; k.fixed.h = sh;
    place(k, x - sw / 2, y - sh / 2, sw, sh);
  });
  if (center) {
    const cs = Math.max(0, r * 2 - size) * 0.9;
    center.fixed.w = cs; center.fixed.h = cs;
    place(center, cx - cs / 2, cy - cs / 2, cs, cs);
  }
  // Connectors between consecutive nodes, along the circle.
  conns.forEach((k, i) => {
    if (i >= centers.length || (i === centers.length - 1 && span < 360)) { place(k, 0, 0, 0, 0); return; }
    const [x0, y0, a0] = centers[i];
    const [x1, y1] = centers[(i + 1) % centers.length];
    void a0;
    const mx = (x0 + x1) / 2, my = (y0 + y1) / 2;
    const dist = Math.hypot(x1 - x0, y1 - y0);
    const len = Math.max(0, dist - size) * 0.6;
    const th = Math.min(len * 0.6, size * 0.25);
    k.rot = (Math.atan2(y1 - y0, x1 - x0) * 180) / Math.PI;
    k.fixed.w = len; k.fixed.h = th;
    place(k, mx - len / 2, my - th / 2, len, th);
  });
}

function layoutPyra(ln) {
  const p = ln.alg.params;
  const kids = visibleKids(ln).filter((k) => k.alg.type !== 'sp' && nodeType(k) !== 'sibTrans');
  const n = kids.length;
  if (!n) return;
  const W = ln.box.w, H = ln.box.h;
  const h = H / n;
  const lvlName = p.pyraLvlNode || 'level';
  const acct = new Set([p.pyraAcctBkgdNode || 'acctBkgd', p.pyraAcctTxNode || 'acctTx']);
  kids.forEach((k, i) => {
    // Level i (from the top): a slice of the triangle spanning the box.
    const y = ln.box.y + i * h;
    const bottom = (W * (i + 1)) / n, mid = (W * (i + 0.5)) / n;
    place(k, ln.box.x, y, W, h);
    k.childrenPlaced = true;
    const shapes = k.children.length ? k.children : [k];
    for (const c of shapes) {
      if (acct.has(c.name)) { place(c, 0, 0, 0, 0); continue; }
      const isLevel = c.name === lvlName || (c.shape && c.shape.type === 'trapezoid');
      const w = isLevel ? bottom : mid;
      place(c, ln.box.x + (W - w) / 2, y, w, h);
      c.fixed.w = w; c.fixed.h = h;
      if (isLevel) c.pyraInset = W / (2 * n);
    }
  });
}

// Hierarchy (org charts): hierRoot = a node above its children, hierChild =
// a row (or hanging column) of hierRoots. Natural sizes come from the
// constraints; the whole tree is then scaled to fit the box.

const isHier = (k) => k.alg.type === 'hierRoot' || k.alg.type === 'hierChild';

function hierParts(ln) {
  const kids = ln.children.filter((k) => k.alg.type !== 'conn' && k.alg.type !== 'sp');
  const comp = kids.find((k) => !isHier(k));
  // Assistants come between the node and its regular children.
  const asst = (k) => k.children.some((c) => c.data && c.data.type === 'asst');
  const subs = kids.filter((k) => k.alg.type === 'hierChild').sort((a, b) => (asst(b) ? 1 : 0) - (asst(a) ? 1 : 0));
  return { comp, subs };
}

function hierNat(ln) {
  if (ln.nat) return ln.nat;
  let nat;
  if (ln.alg.type === 'hierRoot') {
    const { comp, subs } = hierParts(ln);
    const r = comp ? hierNat(comp) : { w: 0, h: 0 };
    let w = r.w, h = r.h;
    const sp = cval(ln, 'sp') ?? r.h * 0.4;
    for (const c of subs) {
      const m = hierNat(c);
      if (!m.w && !m.h) continue;
      const hang = (c.alg.params.linDir || 'fromL') === 'fromT';
      w = Math.max(w, hang ? r.w * 0.3 + m.w : m.w);
      h += sp + m.h;
    }
    nat = { w, h, sp };
  } else if (ln.alg.type === 'hierChild') {
    const items = ln.children.filter((k) => k.alg.type === 'hierRoot').map(hierNat).filter((m) => m.w || m.h);
    const horiz = !['fromT', 'fromB'].includes(ln.alg.params.linDir || 'fromL');
    const sib = cval(ln, 'sibSp') ?? (items.length ? (horiz ? items[0].w : items[0].h) * 0.2 : 0);
    if (!items.length) nat = { w: 0, h: 0, sib };
    else if (horiz) nat = { w: items.reduce((a, m) => a + m.w, 0) + sib * (items.length - 1), h: Math.max(...items.map((m) => m.h)), sib };
    else nat = { w: Math.max(...items.map((m) => m.w)), h: items.reduce((a, m) => a + m.h, 0) + (sib / 2) * (items.length - 1), sib: sib / 2 };
  } else {
    const w = cval(ln, 'w'), h = cval(ln, 'h');
    nat = { w: Number.isFinite(w) && w > 0 ? w : 100, h: Number.isFinite(h) && h > 0 ? h : 60 };
    // User-resized node shapes take their size in the hierarchy.
    const cs = [ln, ...ln.children].find((k) => k.custScale);
    if (cs) nat = { w: nat.w * cs.custScale[0], h: nat.h * cs.custScale[1] };
  }
  ln.nat = nat;
  return nat;
}

function hierPlace(ln, x, y, s) {
  const m = hierNat(ln);
  place(ln, x, y, m.w * s, m.h * s);
  if (ln.alg.type === 'hierRoot') {
    const { comp, subs } = hierParts(ln);
    const r = comp ? hierNat(comp) : { w: 0, h: 0 };
    const align = ln.alg.params.hierAlign || 'tCtrCh';
    let yy = y + (r.h + m.sp) * s;
    if (comp) {
      const cx = align === 'tL' || align === 'lT' ? x : align === 'tR' ? x + (m.w - r.w) * s : x + ((m.w - r.w) * s) / 2;
      place(comp, cx, y, r.w * s, r.h * s);
      comp.fixed.w = r.w * s; comp.fixed.h = r.h * s;
      layout(comp);
    }
    for (const c of subs) {
      const cm = hierNat(c);
      if (!cm.w && !cm.h) { place(c, 0, 0, 0, 0); continue; }
      const hang = (c.alg.params.linDir || 'fromL') === 'fromT';
      const cx = hang ? x + r.w * 0.3 * s : x + ((m.w - cm.w) * s) / 2;
      hierPlace(c, cx, yy, s);
      c.parentComp = comp;
      yy += (cm.h + m.sp) * s;
    }
    return;
  }
  if (ln.alg.type === 'hierChild') {
    const horiz = !['fromT', 'fromB'].includes(ln.alg.params.linDir || 'fromL');
    let pos = 0;
    let lastConn = null;
    for (const k of ln.children) {
      if (k.alg.type === 'conn') { lastConn = k; place(k, 0, 0, 0, 0); continue; }
      if (k.alg.type !== 'hierRoot') { place(k, 0, 0, 0, 0); continue; }
      const km = hierNat(k);
      if (!km.w && !km.h) continue;
      if (horiz) hierPlace(k, x + pos * s, y, s); else hierPlace(k, x, y + pos * s, s);
      pos += (horiz ? km.w : km.h) + m.sib;
      if (lastConn) { lastConn.connTo = k; lastConn.connHang = !horiz; lastConn.connOwner = ln; lastConn = null; }
    }
    return;
  }
  ln.fixed.w = m.w * s; ln.fixed.h = m.h * s;
  layout(ln);
}

/** Elbow connectors from a parent node to each child (org chart lines). */
function hierConnectors(ln) {
  for (const k of ln.children) hierConnectors(k);
  if (!ln.connTo || !ln.connOwner || !ln.connOwner.parentComp) return;
  const p = ln.connOwner.parentComp.box;
  const { comp } = hierParts(ln.connTo);
  if (!comp) return;
  const c = comp.box;
  let pts;
  if (ln.connHang) {
    const vx = p.x + Math.min(p.w * 0.15, (c.x - p.x) / 2);
    pts = [[vx, p.y + p.h], [vx, c.y + c.h / 2], [c.x, c.y + c.h / 2]];
  } else {
    const my = (p.y + p.h + c.y) / 2;
    pts = [[p.x + p.w / 2, p.y + p.h], [p.x + p.w / 2, my], [c.x + c.w / 2, my], [c.x + c.w / 2, c.y]];
  }
  const xs = pts.map((q) => q[0]), ys = pts.map((q) => q[1]);
  const bx = Math.min(...xs), by = Math.min(...ys);
  place(ln, bx, by, Math.max(...xs) - bx, Math.max(...ys) - by);
  ln.connPath = pts.map(([px, py]) => [px - bx, py - by]);
}

function layoutHier(ln) {
  const m = hierNat(ln);
  const s = Math.min(ln.box.w / Math.max(1e-6, m.w), ln.box.h / Math.max(1e-6, m.h));
  const box = { ...ln.box };
  hierPlace(ln, box.x + (box.w - m.w * s) / 2, box.y + (box.h - m.h * s) / 2, s);
  hierConnectors(ln);
}

/** Lay out the children of ln inside ln.box (recursively). */
function layout(ln) {
  if (ln.childrenPlaced) { for (const k of ln.children) layout(k); return; }
  switch (ln.alg.type) {
    case 'lin': layoutLin(ln); break;
    case 'snake': layoutSnake(ln); break;
    case 'cycle': layoutCycle(ln); break;
    case 'pyra': layoutPyra(ln); break;
    case 'hierRoot': case 'hierChild': layoutHier(ln); return; // places the whole subtree
    case 'tx': case 'sp': case 'conn': break;
    default: layoutComposite(ln);
  }
  for (const k of ln.children) if (!(k.alg.type === 'hierRoot' || k.alg.type === 'hierChild') || k.box.w) layout(k);
}

// ---- styles and text ----------------------------------------------------------------

function styleLabel(ln, dm) {
  const p = ln.data && ln.data.pt ? dm.pres.get(ln.data.pt.id + '|' + ln.name) : null;
  return {
    lbl: (p && p.pr.presStyleLbl) || ln.styleLbl || (ln.alg.type === 'conn' ? 'sibTrans2D1' : 'node1'),
    idx: p && p.pr.presStyleIdx !== undefined ? Number(p.pr.presStyleIdx) : null,
    cnt: p && p.pr.presStyleCnt !== undefined ? Number(p.pr.presStyleCnt) : null,
    pres: p,
  };
}

function toHsl({ r, g, b }) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, sat = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, sat, l];
}

function fromHsl(h, sat, l) {
  const k = (n) => (n + h / 30) % 12, a = sat * Math.min(l, 1 - l);
  const f = (n) => l - a * Math.max(-1, Math.min(k(n) - 3, 9 - k(n), 1));
  return [f(0), f(8), f(4)].map((v) => Math.round(v * 255));
}

function pickColor(list, idx, cnt, ctx) {
  if (!list) return null;
  const cols = list.children.filter((c) => /Clr$/.test(c.local));
  if (!cols.length) return null;
  // The default method spreads the colours over the nodes (hue interpolation).
  const meth = list.attr('meth', 'span');
  if (meth === 'span' && cols.length > 1) {
    const a = ctx.color(cols[0]), b = ctx.color(cols[cols.length - 1]);
    if (a && b) {
      const t = cnt > 1 ? Math.max(0, Math.min(1, idx / (cnt - 1))) : 0;
      const [h0, s0, l0] = toHsl(a), [h1, s1, l1] = toHsl(b);
      let dh = h1 - h0;
      if (list.attr('hueDir', 'cw') === 'cw') { if (dh < 0) dh += 360; } else if (dh > 0) dh -= 360;
      const rgb = fromHsl((h0 + dh * t + 360) % 360, s0 + (s1 - s0) * t, l0 + (l1 - l0) * t);
      const hex = rgb.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
      const alpha = (a.a ?? 1) + ((b.a ?? 1) - (a.a ?? 1)) * t;
      return node('a:srgbClr', { val: hex }, alpha < 1 ? [node('a:alpha', { val: String(Math.round(alpha * 100000)) })] : []);
    }
  }
  return clone(cols[(meth === 'span' ? Math.min(idx, cols.length - 1) : idx) % cols.length]);
}

function findLbl(root, name) {
  if (!root) return null;
  for (const l of root.all('dgm:styleLbl')) if (l.attrs.name === name) return l;
  return null;
}

/** Text lines (paragraph nodes with a level) shown by a layout node. */
const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** Add a shape property, replacing one of the same kind. */
function setProp(kids, c) {
  const same = (k) => k.local === c.local || (FILLS.has(k.local) && FILLS.has(c.local));
  const i = kids.findIndex(same);
  if (i >= 0) kids[i] = c;
  else kids.push(c);
}

function textOf(ln) {
  // Only the text algorithm displays text; other nodes use presOf for styling.
  if (!ln.presOf || !ln.data || ln.alg.type !== 'tx') return [];
  const items = select(ln.data, ln.presOf);
  const out = [];
  const base = depthOf(ln.data);
  const first = items.length ? depthOf(items[0]) - base : 0;
  const stBullet = Number(ln.alg.params.stBulletLvl || 0);
  for (const it of items) {
    const t = it.pt && it.pt.t;
    if (!t) continue;
    // Levels count from the first shown item; bullets start at stBulletLvl.
    const lvl = Math.max(0, depthOf(it) - base - first);
    for (const p of t.all('a:p')) out.push({ p, lvl, bullet: (stBullet ? lvl + first >= stBullet : lvl > 0) && ln.vars.bulletEnabled !== '0' });
  }
  return out;
}

/** Rough text fit: largest size (pt) <= max whose wrapped text fits the box. */
function fitFont(lines, w, h, max, min) {
  const words = lines.map(({ p, lvl }) => ({ text: p.textContent(), lvl }));
  if (!words.some((l) => l.text.trim())) return max;
  for (let sz = max; sz >= min; sz -= sz > 20 ? 1 : 0.5) {
    const px = sz * PT;
    let total = 0;
    let ok = true;
    for (const l of words) {
      const fs = l.lvl ? px * 0.78 : px;
      const avail = Math.max(1, w - (l.lvl ? fs * 1.2 : 0));
      let lineW = 0, count = 1;
      for (const word of l.text.split(/\s+/)) {
        const ww = word.length * fs * 0.58;
        if (ww > avail) ok = false;
        if (lineW && lineW + fs * 0.25 + ww > avail) { count++; lineW = ww; } else lineW += (lineW ? fs * 0.25 : 0) + ww;
      }
      total += count * fs * 1.2;
    }
    if (ok && total <= h) return sz;
  }
  return min;
}

// ---- output ----------------------------------------------------------------------------

const SHAPE_ALIAS = { conn: 'rightArrow', none: '', '': '' };

function adjNodes(prst, adj) {
  if (!adj.length) return [];
  const single = adj.length === 1 && ['roundRect', 'snip1Rect', 'round1Rect', 'octagon', 'plaque', 'can', 'cube', 'bevel', 'donut', 'foldedCorner', 'frame', 'hexagon', 'chevron', 'homePlate', 'parallelogram', 'trapezoid', 'triangle', 'plus', 'sun', 'moon', 'smileyFace', 'teardrop', 'flowChartOffpageConnector'].includes(prst);
  return adj.map(([i, v]) => node('a:gd', { name: single && i === 1 ? 'adj' : 'adj' + i, fmla: 'val ' + Math.round(v * 100000) }));
}

/**
 * Lay out a diagram.
 * @param {object} parts {data, layout, style, colors} XML roots
 * @param {number} w frame width (px)
 * @param {number} h frame height (px)
 * @param {object} ctx colour context of the slide
 * @returns {XNode|null} a dsp:spTree with the diagram shapes
 */
/** Debug helper: the layout tree of a diagram. */
export function debugLayout(parts, w, h) {
  const dm = parseDataModel(parts.data);
  const def = parts.layout.name.endsWith('layoutDef') ? parts.layout : parts.layout.first('dgm:layoutDef');
  const root = new Builder(dm, def).build(def);
  place(root, 0, 0, w, h);
  layout(root);
  const out = [];
  const walk = (ln, d) => {
    const txt = ln.data && ln.data.pt && ln.data.pt.t ? ln.data.pt.t.textContent().slice(0, 12) : '';
    out.push(' '.repeat(d) + `${ln.name}${ln.vars.hierBranch ? '(' + ln.vars.hierBranch + ')' : ''} [${ln.alg.type}${Object.keys(ln.alg.params).length ? ' ' + JSON.stringify(ln.alg.params) : ''}] ${ln.shape ? ln.shape.type : '-'} ${Object.values(ln.box).map(Math.round).join(',')} ${ln.data ? ln.data.type : ''} ${txt}`);
    ln.children.forEach((k) => walk(k, d + 1));
  };
  walk(root, 0);
  return out.join('\n');
}

export function layoutDiagram(parts, w, h, ctx) {
  if (!parts.data || !parts.layout) return null;
  const dm = parseDataModel(parts.data);
  if (!dm.doc) return null;
  const def = parts.layout.name.endsWith('layoutDef') ? parts.layout : parts.layout.first('dgm:layoutDef');
  if (!def) return null;
  const root = new Builder(dm, def).build(def);
  place(root, 0, 0, w, h);
  layout(root);

  // Collect drawable nodes in z-order.
  const shapes = [];
  const walk = (ln) => {
    if (ln.box.w > 0.5 || ln.box.h > 0.5 || ln.connPath) {
      const type = ln.shape ? (SHAPE_ALIAS[ln.shape.type] ?? ln.shape.type) : '';
      const lines = textOf(ln);
      if ((type && !ln.shape.hideGeom) || lines.length) shapes.push({ ln, type, lines });
    }
    for (const k of ln.children) walk(k);
  };
  walk(root);

  // Font size per layout node name (nodes of a kind share their size).
  const fonts = new Map();
  for (const s of shapes) {
    if (!s.lines.length) continue;
    const prim = cval(s.ln, 'primFontSz') || 65;
    const rule = s.ln.rules.find((r) => r.type === 'primFontSz');
    const min = rule && Number.isFinite(Number(rule.val)) ? Number(rule.val) : 5;
    const marg = 0.3 * prim * PT * 0.2;
    const size = fitFont(s.lines, s.ln.box.w - 2 * marg, s.ln.box.h - 2 * marg, prim, Math.min(min, prim));
    fonts.set(s.ln.name, Math.min(fonts.has(s.ln.name) ? fonts.get(s.ln.name) : Infinity, size));
  }

  const counts = new Map();
  for (const s of shapes) counts.set(s.ln.name, (counts.get(s.ln.name) || 0) + 1);
  const seen = new Map();
  const tree = node('dsp:spTree', {}, [node('dsp:nvGrpSpPr', {}, [node('dsp:cNvPr', { id: '0', name: '' }), node('dsp:cNvGrpSpPr')]), node('dsp:grpSpPr')]);
  let id = 1;
  for (const s of shapes) {
    const { ln } = s;
    const st = styleLabel(ln, dm);
    const order = seen.get(ln.name) || 0;
    seen.set(ln.name, order + 1);
    const idx = st.idx !== null ? st.idx : order;
    const cnt = st.cnt !== null ? st.cnt : counts.get(ln.name);
    const colors = findLbl(parts.colors, st.lbl);
    const quick = findLbl(parts.style, st.lbl);
    const fillClr = colors && pickColor(colors.child('dgm:fillClrLst'), idx, cnt, ctx);
    const linClr = colors && pickColor(colors.child('dgm:linClrLst'), idx, cnt, ctx);
    const txClr = colors && pickColor(colors.child('dgm:txFillClrLst'), idx, cnt, ctx);
    const qs = quick && quick.child('dgm:style');
    const ref = (name, clr, def) => {
      const q = qs && qs.child('a:' + name);
      const n = node('a:' + name, { idx: q ? q.attr('idx', def) : def });
      const c = clr || (q && q.children[0] ? clone(q.children[0]) : null);
      if (c) { c.parent = n; n.children.push(c); }
      return n;
    };
    const style = node('dsp:style', {}, [ref('lnRef', linClr, '2'), ref('fillRef', fillClr, '1'), ref('effectRef', null, '0'), ref('fontRef', txClr, 'minor')]);
    if (style.children[3]) style.children[3].attrs.idx = (qs && qs.child('a:fontRef') && qs.child('a:fontRef').attrs.idx) || 'minor';
    const b = ln.box;
    const rot = (ln.shape ? ln.shape.rot : 0) + (ln.rot || 0);
    // A shape turned by 90 / 270 degrees fills its layout box when rotated:
    // its own frame has swapped sides; its text stays upright in the box.
    const quarter = !ln.rot && Math.abs(((rot % 180) + 180) % 180 - 90) < 1;
    const fw = quarter ? b.h : b.w, fh = quarter ? b.w : b.h;
    const fx = b.x + (b.w - fw) / 2, fy = b.y + (b.h - fh) / 2;
    const emu = (v) => String(Math.round(v * EMU));
    const xfrm = node('a:xfrm', rot ? { rot: String(Math.round(rot * 60000)) } : {}, [
      node('a:off', { x: emu(fx), y: emu(fy) }),
      node('a:ext', { cx: emu(fw), cy: emu(fh) }),
    ]);
    // Text stays upright (autoTxRot "upr"); txXfrm rotation is relative to the shape.
    const upright = rot && s.lines.length && (ln.alg.params.autoTxRot || 'upr') !== 'none';
    const txXfrm = upright ? node('dsp:txXfrm', { rot: String(Math.round(-rot * 60000)) }, [node('a:off', { x: emu(b.x), y: emu(b.y) }), node('a:ext', { cx: emu(b.w), cy: emu(b.h) })]) : null;
    const spKids = [xfrm];
    let prst = s.type && !ln.shape.hideGeom ? s.type : 'rect';
    if (ln.connPath) prst = '';
    if (ln.pyraInset !== undefined && prst === 'trapezoid') {
      // Pyramid levels: trapezoids whose sides continue the pyramid outline.
      ln.shape.adj = [[1, ln.pyraInset / Math.max(1e-6, Math.min(b.w, b.h))]];
    }
    if (ln.connPath) {
      const w = Math.max(1, Math.round(b.w * EMU)), hh = Math.max(1, Math.round(b.h * EMU));
      const pt = ([px, py]) => node('a:pt', { x: String(Math.round(px * EMU)), y: String(Math.round(py * EMU)) });
      spKids.push(node('a:custGeom', {}, [node('a:avLst'), node('a:gdLst'), node('a:ahLst'), node('a:cxnLst'), node('a:rect', { l: '0', t: '0', r: 'r', b: 'b' }),
        node('a:pathLst', {}, [node('a:path', { w: String(w), h: String(hh), fill: 'none' }, [node('a:moveTo', {}, [pt(ln.connPath[0])]), ...ln.connPath.slice(1).map((q) => node('a:lnTo', {}, [pt(q)]))])])]));
      spKids.push(node('a:noFill'));
    } else spKids.push(node('a:prstGeom', { prst }, [node('a:avLst', {}, ln.shape ? adjNodes(prst, ln.shape.adj) : [])]));
    const hidden = !s.type || ln.shape.hideGeom;
    if (hidden) spKids.push(node('a:noFill'), node('a:ln', {}, [node('a:noFill')]));
    // Formatting set by the user: on the data point (for the shape showing
    // it), then on the presentation point.
    const own = !hidden && ln.presOf && ln.data && ln.data.pt && select(ln.data, ln.presOf).includes(ln.data) ? ln.data.pt.spPr : null;
    for (const custom of [own, st.pres && st.pres.spPr]) {
      if (custom) for (const c of custom.children) if (c.local !== 'xfrm' && !/Geom$/.test(c.local)) setProp(spKids, clone(c));
    }
    const spPr = node('dsp:spPr', {}, spKids);

    let txBody = null;
    if (s.lines.length) {
      const size = fonts.get(ln.name) || 18;
      const marg = Math.round(size * 0.3 * 0.25 * 12700);
      const listed = s.lines.some((l) => l.bullet);
      const para = s.lines.map(({ p, lvl, bullet }) => {
        const c = clone(p);
        let pPr = c.child('a:pPr');
        if (!pPr) { pPr = node('a:pPr'); pPr.parent = c; c.children.unshift(pPr); }
        const ta = ln.alg.params.parTxLTRAlign;
        if (pPr.attrs.algn === undefined) pPr.attrs.algn = listed ? 'l' : ta === 'l' ? 'l' : ta === 'r' ? 'r' : 'ctr';
        if (lvl) pPr.attrs.lvl = String(Math.min(8, lvl));
        if (bullet) {
          pPr.attrs.marL = String(Math.round((lvl + 1) * size * 0.75 * 12700));
          pPr.attrs.indent = String(Math.round(-size * 0.75 * 12700));
          if (!pPr.child('a:buChar')) { const bu = node('a:buChar', { char: '•' }); bu.parent = pPr; pPr.children.push(bu); }
        }
        const fsz = String(Math.round((lvl ? size * 0.78 : size) * 100));
        for (const r of c.all('a:r')) {
          let rPr = r.child('a:rPr');
          if (!rPr) { rPr = node('a:rPr'); rPr.parent = r; r.children.unshift(rPr); }
          if (rPr.attrs.sz === undefined) rPr.attrs.sz = fsz;
        }
        let end = c.child('a:endParaRPr');
        if (!end) { end = node('a:endParaRPr'); end.parent = c; c.children.push(end); }
        if (end.attrs.sz === undefined) end.attrs.sz = fsz;
        return c;
      });
      const anchor = ln.alg.params.txAnchorVert === 't' ? 't' : ln.alg.params.txAnchorVert === 'b' ? 'b' : 'ctr';
      txBody = node('dsp:txBody', {}, [
        node('a:bodyPr', { wrap: 'square', lIns: String(marg), tIns: String(marg), rIns: String(marg), bIns: String(marg), anchor, anchorCtr: '0', spcCol: '0' }, [node('a:noAutofit')]),
        node('a:lstStyle'),
        ...para,
      ]);
    }
    const sp = node('dsp:sp', { modelId: ln.data && ln.data.pt ? ln.data.pt.id : String(id) }, [
      node('dsp:nvSpPr', {}, [node('dsp:cNvPr', { id: String(id++), name: ln.name }), node('dsp:cNvSpPr')]),
      spPr, style, txBody, txXfrm,
    ]);
    tree.children.push(sp);
    sp.parent = tree;
  }
  return tree;
}

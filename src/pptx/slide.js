/**
 * Slide model builder: masters, layouts, slides and the shape tree.
 */
import { Ctx } from './context.js';
import { Theme } from './theme.js';
import { DEFAULT_CLR_MAP } from './color.js';
import { emu2px } from '../util/units.js';
import {
  parseFill, parseFillNode, fillChild, parseLine, mergeLines, finalizeLine, parseEffects,
  fillFromRef, lineFromRef, effectsFromRef, parseBlipFill,
} from './fill.js';
import { customGeometryDef, computeGeometry, presetGeometry, parseAvLst, hasPreset } from '../geometry/geometry.js';
import { listStyle, bodyPrBag, parseTextBody, fontRefInfo, parseHyperlink, hasVisibleText } from './text.js';
import { parseTable } from './table.js';
import { parseChartPart } from '../chart/parse.js';
import { parseTiming } from './timing.js';
import { parseTransition } from './transition.js';
import { layoutDiagram } from '../smartart/engine.js';

const SUPPORTED_MC = new Set(['p14', 'p15', 'p159', 'p1510', 'p1710', 'a14', 'a15', 'a16', 'a1611', 'asvg', 'adec', 'c14', 'c15', 'c16', 'c16r3', 'p188', 'thm15', 'v', 'o', 'cx', 'aink', 'am3d', 'pc', 'ac']);

/** Choose the branch of an mc:AlternateContent element. */
export function chooseAlternate(node, prefer = {}) {
  for (const ch of node.all('mc:Choice')) {
    const req = (ch.attrs.Requires || '').split(/\s+/).filter(Boolean);
    if (!req.every((r) => SUPPORTED_MC.has(r))) continue;
    // Office math and 3D models render better from their fallback picture.
    if (req.includes('a14') && ch.first('a14:m')) continue;
    if (req.includes('am3d') && !prefer.model3d) continue;
    if (req.includes('aink') && !prefer.ink) continue;
    if (req.includes('cx') && prefer.chartex === false) continue;
    return ch.children;
  }
  const fb = node.child('mc:Fallback');
  return fb ? fb.children : [];
}

/** Local-name child lookup accepting p:, dsp: and a: prefixes. */
function pc(node, local) {
  if (!node) return null;
  for (const c of node.children) {
    if (c.local === local && (c.name[0] === 'p' || c.name.startsWith('dsp:') || c.name.startsWith('a:'))) return c;
  }
  return null;
}

function parseXfrm(x) {
  if (!x) return null;
  const off = x.child('a:off');
  const ext = x.child('a:ext');
  const chOff = x.child('a:chOff');
  const chExt = x.child('a:chExt');
  return {
    x: off ? off.num('x', 0) : 0,
    y: off ? off.num('y', 0) : 0,
    cx: ext ? ext.num('cx', 0) : 0,
    cy: ext ? ext.num('cy', 0) : 0,
    chx: chOff ? chOff.num('x', 0) : 0,
    chy: chOff ? chOff.num('y', 0) : 0,
    chcx: chExt ? chExt.num('cx', 0) : 0,
    chcy: chExt ? chExt.num('cy', 0) : 0,
    rot: x.num('rot', 0) / 60000,
    flipH: x.bool('flipH', false),
    flipV: x.bool('flipV', false),
    hasOff: !!off,
  };
}

function nonVisual(node) {
  for (const c of node.children) {
    if (/^nv[A-Za-z]*Pr$/.test(c.local)) return c;
  }
  return null;
}

function cNvPrOf(node) {
  const nv = nonVisual(node);
  return nv ? pc(nv, 'cNvPr') : null;
}

function phOf(node) {
  const nv = nonVisual(node);
  const nvPr = nv && pc(nv, 'nvPr');
  const ph = nvPr && nvPr.child('p:ph');
  if (!ph) return null;
  return {
    type: ph.attr('type', 'obj'),
    idx: ph.attrs.idx !== undefined ? String(ph.attrs.idx) : null,
    hasCustomPrompt: ph.bool('hasCustomPrompt', false),
    orient: ph.attr('orient', 'horz'),
    sz: ph.attr('sz', 'full'),
  };
}

const TITLE_TYPES = new Set(['title', 'ctrTitle']);
const BODY_TYPES = new Set(['body', 'subTitle', 'obj']);
const OTHER_TYPES = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

/** Master placeholder type that a layout/slide placeholder type maps to. */
function masterType(type) {
  if (TITLE_TYPES.has(type)) return 'title';
  if (OTHER_TYPES.has(type)) return type;
  return 'body';
}

/** A slide master or layout. */
class SlideBase {
  constructor(doc, part, kind) {
    this.doc = doc;
    this.part = part;
    this.kind = kind;
    this.root = doc.pkg.xml(part);
    this.cSld = this.root ? this.root.child('p:cSld') : null;
    this.spTree = this.cSld ? this.cSld.child('p:spTree') : null;
    this.name = this.cSld ? this.cSld.attr('name', '') : '';
    this.showMasterSp = this.root ? this.root.bool('showMasterSp', true) : true;
    this.byIdx = new Map();
    this.byType = new Map();
    if (this.spTree) this._index(this.spTree);
  }

  _index(tree) {
    for (const c of tree.children) {
      if (c.name === 'mc:AlternateContent') {
        for (const x of chooseAlternate(c)) this._indexNode(x);
        continue;
      }
      this._indexNode(c);
    }
  }

  _indexNode(c) {
    const ph = phOf(c);
    if (!ph) return;
    if (ph.idx !== null && !this.byIdx.has(ph.idx)) this.byIdx.set(ph.idx, c);
    if (!this.byType.has(ph.type)) this.byType.set(ph.type, c);
  }

  /** Find the placeholder matching a child placeholder (layout lookup). */
  findPlaceholder(ph, isMaster) {
    if (!ph) return null;
    if (isMaster) {
      const t = masterType(ph.type);
      return this.byType.get(t) || (t === 'title' ? this.byType.get('ctrTitle') : null) || null;
    }
    if (ph.idx !== null && ph.idx !== '0' && this.byIdx.has(ph.idx)) {
      const n = this.byIdx.get(ph.idx);
      return n;
    }
    let n = this.byType.get(ph.type);
    if (!n && ph.type === 'title') n = this.byType.get('ctrTitle');
    if (!n && ph.type === 'ctrTitle') n = this.byType.get('title');
    if (!n && ph.type === 'subTitle') n = this.byType.get('body');
    if (!n && ph.type === 'obj') n = this.byType.get('body');
    if (!n && ph.idx !== null && this.byIdx.has(ph.idx)) n = this.byIdx.get(ph.idx);
    return n || null;
  }
}

class Master extends SlideBase {
  constructor(doc, part) {
    super(doc, part, 'master');
    const themeRel = doc.pkg.relByType(part, 'theme');
    this.theme = doc.theme(themeRel ? themeRel.partName : null);
    const cm = this.root && this.root.child('p:clrMap');
    this.clrMap = cm ? { ...DEFAULT_CLR_MAP, ...cm.attrs } : { ...DEFAULT_CLR_MAP };
    const tx = this.root && this.root.child('p:txStyles');
    this.titleStyle = listStyle(tx && tx.child('p:titleStyle'));
    this.bodyStyle = listStyle(tx && tx.child('p:bodyStyle'));
    this.otherStyle = listStyle(tx && tx.child('p:otherStyle'));
    this.hf = this.root && this.root.child('p:hf');
  }

  textStyleFor(phType) {
    if (!phType) return this.otherStyle;
    if (TITLE_TYPES.has(phType)) return this.titleStyle;
    if (BODY_TYPES.has(phType)) return this.bodyStyle;
    if (OTHER_TYPES.has(phType)) return this.otherStyle;
    return this.bodyStyle;
  }
}

class Layout extends SlideBase {
  constructor(doc, part) {
    super(doc, part, 'layout');
    const mr = doc.pkg.relByType(part, 'slideMaster');
    this.master = mr ? doc.master(mr.partName) : null;
    this.type = this.root ? this.root.attr('type', 'cust') : 'cust';
    this.clrMapOvr = this.root && this.root.child('p:clrMapOvr');
  }
}

function clrMapFrom(ovr, base) {
  if (!ovr) return base;
  const o = ovr.child('a:overrideClrMapping');
  if (o) return { ...DEFAULT_CLR_MAP, ...o.attrs };
  return base;
}

/** Per-slide builder. */
export class SlideBuilder {
  constructor(doc, info) {
    this.doc = doc;
    this.info = info;
    this.part = info.part;
    this.root = doc.pkg.xml(info.part);
    const lr = doc.pkg.relByType(info.part, 'slideLayout');
    this.layout = lr ? doc.layout(lr.partName) : null;
    this.master = this.layout ? this.layout.master : null;
    if (!this.master && doc.masters.size) this.master = doc.masters.values().next().value;
    const theme = this.master ? this.master.theme : doc.theme(null);
    let clrMap = this.master ? this.master.clrMap : { ...DEFAULT_CLR_MAP };
    if (this.layout) clrMap = clrMapFrom(this.layout.clrMapOvr, clrMap);
    clrMap = clrMapFrom(this.root && this.root.child('p:clrMapOvr'), clrMap);
    this.baseCtx = new Ctx(doc, info.part).with({ theme, scheme: theme.colorScheme, clrMap });
    this.idSeq = 0;
  }

  ctxFor(part) {
    return this.baseCtx.with({ part });
  }

  build() {
    const doc = this.doc;
    const sld = this.root;
    const slide = {
      index: this.info.index,
      number: this.info.number,
      part: this.part,
      hidden: sld ? sld.attr('show', '1') === '0' : false,
      name: sld && sld.child('p:cSld') ? sld.child('p:cSld').attr('name', '') : '',
      layoutName: this.layout ? this.layout.name : '',
      layoutType: this.layout ? this.layout.type : '',
      background: null,
      elements: [],
      timing: null,
      transition: null,
      notes: null,
    };
    if (!sld) {
      doc.diag.add('error', 'slide part missing: ' + this.part);
      return slide;
    }
    slide.background = this.background();
    const showMasterSp = sld.bool('showMasterSp', true);
    const layoutShowMaster = this.layout ? this.layout.showMasterSp : true;
    if (showMasterSp && layoutShowMaster && this.master && this.master.spTree) {
      slide.elements.push(...this.tree(this.master.spTree, this.ctxFor(this.master.part), { layer: 'master', base: this.master }));
    }
    if (showMasterSp && this.layout && this.layout.spTree) {
      slide.elements.push(...this.tree(this.layout.spTree, this.ctxFor(this.layout.part), { layer: 'layout', base: this.layout }));
    }
    const spTree = sld.path('p:cSld', 'p:spTree');
    if (spTree) slide.elements.push(...this.tree(spTree, this.baseCtx, { layer: 'slide' }));
    const controls = sld.path('p:cSld', 'p:controls');
    if (controls) slide.elements.push(...this.controls(controls, this.baseCtx));

    try {
      slide.transition = parseTransition(sld, this.baseCtx);
    } catch (e) {
      doc.diag.add('error', 'transition: ' + e.message, this.part);
    }
    try {
      slide.timing = parseTiming(sld, this.baseCtx, slide);
    } catch (e) {
      doc.diag.add('error', 'timing: ' + e.message, this.part);
    }
    slide.notes = this.notes();
    return slide;
  }

  background() {
    const chain = [[this.root, this.part], [this.layout && this.layout.root, this.layout && this.layout.part], [this.master && this.master.root, this.master && this.master.part]];
    for (const [root, part] of chain) {
      const bg = root && root.path('p:cSld', 'p:bg');
      if (!bg) continue;
      const ctx = this.ctxFor(part);
      const bgPr = bg.child('p:bgPr');
      if (bgPr) {
        const f = parseFill(bgPr, ctx);
        const fx = parseEffects(bgPr.child('a:effectLst'), ctx);
        return { fill: f || { type: 'none' }, effects: fx };
      }
      const ref = bg.child('p:bgRef');
      if (ref) return { fill: fillFromRef(ref, ctx) || { type: 'none' } };
    }
    return { fill: { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } } };
  }

  notes() {
    const r = this.doc.pkg.relByType(this.part, 'notesSlide');
    if (!r) return null;
    const root = this.doc.pkg.xml(r.partName);
    if (!root) return null;
    const tree = root.path('p:cSld', 'p:spTree');
    if (!tree) return null;
    const out = [];
    for (const sp of tree.all('p:sp')) {
      const ph = phOf(sp);
      if (!ph || ph.type !== 'body') continue;
      const tx = sp.child('p:txBody');
      if (!tx) continue;
      for (const p of tx.all('a:p')) {
        let line = '';
        for (const r2 of p.children) {
          if (r2.name === 'a:r' || r2.name === 'a:fld') line += (r2.child('a:t') || { text: '' }).text;
          else if (r2.name === 'a:br') line += '\n';
        }
        out.push(line);
      }
    }
    const text = out.join('\n').trim();
    return text || null;
  }

  /** Parse a shape tree (spTree / grpSp). */
  tree(tree, ctx, o) {
    const out = [];
    for (const c of tree.children) {
      if (c.name === 'mc:AlternateContent') {
        for (const x of chooseAlternate(c, this.doc.options.alternate || {})) {
          const el = this.element(x, ctx, o);
          if (el) out.push(el);
        }
        continue;
      }
      const el = this.element(c, ctx, o);
      if (el) out.push(el);
    }
    return out;
  }

  element(node, ctx, o) {
    const local = node.local;
    try {
      switch (local) {
        case 'sp': return this.shape(node, ctx, o, false);
        case 'cxnSp': return this.shape(node, ctx, o, true);
        case 'pic': return this.picture(node, ctx, o);
        case 'grpSp': return this.group(node, ctx, o);
        case 'graphicFrame': return this.graphicFrame(node, ctx, o);
        case 'contentPart': return this.contentPart(node, ctx, o);
        case 'nvGrpSpPr': case 'grpSpPr': case 'extLst': case 'txXfrm':
          return null;
        default:
          if (node.name.startsWith('p:') || node.name.startsWith('dsp:')) ctx.note('unsupported', 'element ' + node.name);
          return null;
      }
    } catch (e) {
      ctx.note('error', node.name + ': ' + (e && e.message));
      if (this.doc.options.strict) throw e;
      return null;
    }
  }

  /** Common non-visual info and the placeholder inheritance chain. */
  common(node, ctx, o) {
    const cNvPr = cNvPrOf(node) || { attrs: {}, child: () => null, bool: () => false, attr: (k, d) => d };
    const ph = phOf(node);
    const el = {
      id: cNvPr.attrs.id || '',
      name: cNvPr.attrs.name || '',
      descr: cNvPr.attrs.descr || cNvPr.attrs.title || '',
      hidden: cNvPr.attrs.hidden === '1' || cNvPr.attrs.hidden === 'true',
      layer: o.layer,
      ph,
      domId: 's' + this.info.index + '-' + (o.layer[0]) + '-' + (cNvPr.attrs.id || 'x') + '-' + (this.idSeq++),
    };
    const hl = cNvPr.child && cNvPr.child('a:hlinkClick');
    if (hl) el.link = parseHyperlink(hl, ctx);
    const hov = cNvPr.child && cNvPr.child('a:hlinkHover');
    if (hov) el.hover = parseHyperlink(hov, ctx);
    // Inheritance chain: [slide node, layout node, master node].
    const chain = [{ node, ctx }];
    if (ph && o.layer !== 'master') {
      if (o.layer === 'slide' && this.layout) {
        const ln = this.layout.findPlaceholder(ph, false);
        if (ln) chain.push({ node: ln, ctx: this.ctxFor(this.layout.part), kind: 'layout' });
        const lph = ln ? phOf(ln) : ph;
        const mn = this.master && this.master.findPlaceholder(lph || ph, true);
        if (mn) chain.push({ node: mn, ctx: this.ctxFor(this.master.part), kind: 'master' });
      } else if (o.layer === 'layout' && this.master) {
        const mn = this.master.findPlaceholder(ph, true);
        if (mn) chain.push({ node: mn, ctx: this.ctxFor(this.master.part), kind: 'master' });
      }
    }
    return { el, chain };
  }

  /** Resolve transform from the chain. */
  transform(chain, spPrName = 'spPr') {
    for (const { node } of chain) {
      const spPr = pc(node, spPrName);
      const x = spPr && spPr.child('a:xfrm');
      if (x && x.child('a:ext')) return parseXfrm(x);
    }
    return null;
  }

  applyXfrm(el, xf, o) {
    const map = o.map || ((v) => v);
    if (!xf) {
      el.x = 0; el.y = 0; el.w = 0; el.h = 0; el.rot = 0; el.flipH = false; el.flipV = false;
      return;
    }
    const m = map({ x: xf.x, y: xf.y, cx: xf.cx, cy: xf.cy });
    el.x = emu2px(m.x);
    el.y = emu2px(m.y);
    el.w = emu2px(m.cx);
    el.h = emu2px(m.cy);
    el.wEmu = m.cx;
    el.hEmu = m.cy;
    el.rot = xf.rot;
    el.flipH = xf.flipH;
    el.flipV = xf.flipV;
  }

  geometry(chain, el, spPrName = 'spPr') {
    for (const { node } of chain) {
      const spPr = pc(node, spPrName);
      if (!spPr) continue;
      const prst = spPr.child('a:prstGeom');
      if (prst) {
        const name = prst.attr('prst', 'rect');
        const adj = parseAvLst(prst.child('a:avLst'));
        let g = presetGeometry(name, el.wEmu || 0, el.hEmu || 0, adj, 1 / 9525);
        if (!g) {
          if (!hasPreset(name)) this.doc.diag.add('unsupported', 'preset geometry ' + name, this.part);
          g = presetGeometry('rect', el.wEmu || 0, el.hEmu || 0, {}, 1 / 9525);
        }
        return { name, adj, ...g };
      }
      const cust = spPr.child('a:custGeom');
      if (cust) {
        const def = customGeometryDef(cust);
        return { name: 'custom', ...computeGeometry(def, el.wEmu || 0, el.hEmu || 0, {}, 1 / 9525) };
      }
    }
    return { name: 'rect', ...presetGeometry('rect', el.wEmu || 0, el.hEmu || 0, {}, 1 / 9525) };
  }

  /** Visual properties (fill, line, effects) from spPr chain and style. */
  visual(chain, el, spPrName = 'spPr') {
    // Each inheritance level (shape, layout placeholder, master placeholder)
    // contributes its spPr first and then its style references.
    const levels = chain.map(({ node, ctx }) => ({ spPr: pc(node, spPrName), style: pc(node, 'style'), ctx }));
    const first = levels.find((l) => l.style);
    const style = first ? first.style : null;
    const styleCtx = first ? first.ctx : null;
    // Fill.
    let fill;
    for (const l of levels) {
      const f = l.spPr ? parseFill(l.spPr, l.ctx) : undefined;
      if (f !== undefined) { fill = f; break; }
      if (l.style) { fill = fillFromRef(l.style.child('a:fillRef'), l.ctx); if (fill !== undefined) break; }
    }
    if (fill && fill.type === 'grad' && !fill.stops.length) {
      // A gradient without stops takes them from the style matrix (or from the
      // next level of the inheritance chain).
      let base = style ? fillFromRef(style.child('a:fillRef'), styleCtx) : undefined;
      if (!base || base.type !== 'grad') {
        for (const l of levels.slice(1)) {
          const f = l.spPr ? parseFill(l.spPr, l.ctx) : undefined;
          if (f && f.type === 'grad' && f.stops.length) { base = f; break; }
        }
      }
      if (base && base.type === 'grad') fill = { ...fill, stops: base.stops };
      else if (base && base.type === 'solid') fill = base;
    }
    // Line: merged from the deepest level up.
    const parts = [];
    for (let i = levels.length - 1; i >= 0; i--) {
      const l = levels[i];
      if (l.style) parts.push(lineFromRef(l.style.child('a:lnRef'), l.ctx));
      if (l.spPr && l.spPr.child('a:ln')) parts.push(parseLine(l.spPr.child('a:ln'), l.ctx));
    }
    const line = finalizeLine(mergeLines(...parts));
    // Effects.
    let effects;
    for (const l of levels) {
      const sp = l.spPr;
      if (sp && sp.child('a:effectLst')) { effects = parseEffects(sp.child('a:effectLst'), l.ctx); break; }
      if (sp && sp.child('a:effectDag')) { effects = null; l.ctx.note('approx', 'effectDag'); break; }
      if (l.style) { effects = effectsFromRef(l.style.child('a:effectRef'), l.ctx); if (effects !== undefined) break; }
    }
    // 3D.
    for (const { node, ctx } of chain) {
      const spPr = pc(node, spPrName);
      const sp3d = spPr && spPr.child('a:sp3d');
      if (sp3d) {
        el.sp3d = {
          bevelT: sp3d.child('a:bevelT') ? { w: emu2px(sp3d.child('a:bevelT').num('w', 76200)), h: emu2px(sp3d.child('a:bevelT').num('h', 76200)), prst: sp3d.child('a:bevelT').attr('prst', 'circle') } : null,
          extrusionH: emu2px(sp3d.num('extrusionH', 0)),
          material: sp3d.attr('prstMaterial', 'warmMatte'),
        };
        if (el.sp3d.bevelT || el.sp3d.extrusionH) ctx.note('approx', '3D shape properties');
        break;
      }
    }
    el.fill = fill || { type: 'none' };
    el.line = line;
    el.effects = effects || null;
    el.style = style;
    el.styleCtx = styleCtx;
  }

  /** Text body with the full inheritance chain. */
  text(chain, el, o, ctx) {
    const own = chain[0].node;
    const txBody = pc(own, 'txBody');
    if (!txBody) return null;
    const doc = this.doc;
    const ph = el.ph;
    const lists = [doc.defaultTextStyle];
    const master = this.master;
    // Placeholders use the master's title/body/other styles; other text on
    // slides uses the presentation defaults, on masters and layouts otherStyle.
    if (master && (ph || o.layer !== 'slide')) {
      lists.push(ph ? master.textStyleFor(ph.type) : master.otherStyle);
    }
    const bodyChain = [];
    for (let i = chain.length - 1; i >= 1; i--) {
      const tb = pc(chain[i].node, 'txBody');
      if (!tb) continue;
      lists.push(listStyle(tb.child('a:lstStyle')));
      bodyChain.push(bodyPrBag(tb.child('a:bodyPr')));
    }
    const fontRef = el.style ? fontRefInfo(el.style, el.styleCtx || ctx) : null;
    return parseTextBody(txBody, ctx, {
      chain: lists,
      bodyChain,
      fontRef,
      slideNumber: this.info.number,
      liveDates: doc.options.liveDates !== false,
    });
  }

  shape(node, ctx, o, isConnector) {
    const { el, chain } = this.common(node, ctx, o);
    if (el.hidden && !this.doc.options.includeHidden) return null;
    // Placeholders on masters and layouts are templates only.
    if (el.ph && o.layer !== 'slide') return null;
    el.kind = 'shape';
    el.connector = isConnector;
    const xf = this.transform(chain);
    this.applyXfrm(el, xf, o);
    el.geom = this.geometry(chain, el);
    this.visual(chain, el);
    // useBgFill: the shape shows the slide background behind it.
    if (/^(1|true)$/.test(node.attrs.useBgFill || '')) el.fill = { type: 'bg' };
    if (!isConnector) {
      el.text = this.text(chain, el, o, ctx);
      // dsp:txXfrm (SmartArt drawings) positions the text separately.
      const txXfrm = node.child('dsp:txXfrm');
      if (txXfrm && el.text) {
        const t = parseXfrm(txXfrm);
        const m = (o.map || ((v) => v))({ x: t.x, y: t.y, cx: t.cx, cy: t.cy });
        // Text box centre in the shape's own (unrotated) frame.
        const w = emu2px(m.cx), h = emu2px(m.cy);
        const dx = emu2px(m.x) + w / 2 - (el.x + el.w / 2), dy = emu2px(m.y) + h / 2 - (el.y + el.h / 2);
        const a = (-(el.rot || 0) * Math.PI) / 180;
        let lx = dx * Math.cos(a) - dy * Math.sin(a), ly = dx * Math.sin(a) + dy * Math.cos(a);
        if (el.flipH) lx = -lx;
        if (el.flipV) ly = -ly;
        el.textBox = { x: el.w / 2 + lx - w / 2, y: el.h / 2 + ly - h / 2, w, h, rot: t.rot || 0 };
      }
      if (el.text && !hasVisibleText(el.text) && el.ph && el.fill.type === 'none' && !el.line) {
        // Empty placeholder: nothing is shown in slide show.
        return null;
      }
    }
    return el;
  }

  picture(node, ctx, o) {
    const { el, chain } = this.common(node, ctx, o);
    if (el.hidden && !this.doc.options.includeHidden) return null;
    if (el.ph && o.layer !== 'slide') return null;
    el.kind = 'pic';
    const xf = this.transform(chain);
    this.applyXfrm(el, xf, o);
    el.geom = this.geometry(chain, el);
    this.visual(chain, el);
    const bf = pc(node, 'blipFill');
    el.blip = bf ? parseBlipFill(bf, ctx) : null;
    // Media (audio / video).
    const nv = nonVisual(node);
    const nvPr = nv && pc(nv, 'nvPr');
    if (nvPr) {
      const media = this.mediaInfo(nvPr, ctx);
      if (media) el.media = media;
    }
    if (!el.blip || !el.blip.image) {
      if (!el.media) ctx.note('missing', 'picture without image');
    }
    return el;
  }

  mediaInfo(nvPr, ctx) {
    let kind = null;
    let src = null;
    const vf = nvPr.child('a:videoFile') || nvPr.child('a:quickTimeFile');
    const af = nvPr.child('a:audioFile') || nvPr.child('a:wavAudioFile');
    const cd = nvPr.child('a:audioCd');
    let p14 = null;
    for (const e of nvPr.find('p14:media')) p14 = e;
    if (p14) {
      src = ctx.mediaRel(p14.attrs['r:embed'] || p14.attrs['r:link'], vf ? 'video' : 'audio');
    }
    if (vf) {
      kind = 'video';
      if (!src) src = ctx.mediaRel(vf.attrs['r:link'] || vf.attrs['r:embed'], 'video');
    } else if (af) {
      kind = 'audio';
      if (!src) src = ctx.mediaRel(af.attrs['r:link'] || af.attrs['r:embed'] || af.attrs['r:embed'], 'audio');
    } else if (cd) {
      ctx.note('unsupported', 'audio CD');
      return null;
    }
    if (!kind) return null;
    if (src && src.mime && src.mime.startsWith('audio/') && kind === 'video') kind = 'audio';
    const out = { kind, src };
    if (p14) {
      const trim = p14.child('p14:trim');
      if (trim) out.trim = { st: trim.num('st', 0), end: trim.num('end', 0) };
      const fade = p14.child('p14:fade');
      if (fade) out.fade = { in: fade.num('in', 0), out: fade.num('out', 0) };
    }
    if (!src) ctx.note('missing', kind + ' source');
    return out;
  }

  group(node, ctx, o) {
    const { el } = this.common(node, ctx, o);
    if (el.hidden && !this.doc.options.includeHidden) return null;
    el.kind = 'group';
    const gpr = pc(node, 'grpSpPr');
    const xf = parseXfrm(gpr && gpr.child('a:xfrm'));
    this.applyXfrm(el, xf, o);
    const sub = { ...o };
    if (xf) {
      // Rendered group size (possibly scaled by an outer group) / child extents.
      const sx = xf.chcx ? (el.wEmu || 0) / xf.chcx : (xf.cx ? (el.wEmu || 0) / xf.cx : 1);
      const sy = xf.chcy ? (el.hEmu || 0) / xf.chcy : (xf.cy ? (el.hEmu || 0) / xf.cy : 1);
      // Children are expressed in the child coordinate space; map them to
      // the group's local box (in EMU, relative to the group origin).
      sub.map = (r) => ({
        x: (r.x - xf.chx) * sx,
        y: (r.y - xf.chy) * sy,
        cx: r.cx * sx,
        cy: r.cy * sy,
      });
    } else {
      sub.map = (r) => r;
    }
    const gf = gpr ? parseFill(gpr, ctx) : undefined;
    let gctx = ctx;
    if (gf !== undefined) gctx = ctx.with({ groupFill: gf });
    el.effects = gpr && gpr.child('a:effectLst') ? parseEffects(gpr.child('a:effectLst'), ctx) : null;
    el.children = this.tree(node, gctx, sub);
    return el;
  }

  graphicFrame(node, ctx, o) {
    const { el, chain } = this.common(node, ctx, o);
    if (el.hidden && !this.doc.options.includeHidden) return null;
    if (el.ph && o.layer !== 'slide') return null;
    // graphicFrame transform lives directly under the frame (p:xfrm).
    let xfNode = pc(node, 'xfrm');
    let xf = xfNode ? parseXfrm(xfNode) : null;
    if (!xf || (!xf.cx && !xf.cy && !xf.hasOff)) {
      for (const { node: n } of chain.slice(1)) {
        const sp = pc(n, 'spPr');
        const x2 = (sp && sp.child('a:xfrm')) || pc(n, 'xfrm');
        if (x2) { xf = parseXfrm(x2); break; }
      }
    }
    this.applyXfrm(el, xf, o);
    const gd = node.path('a:graphic', 'a:graphicData');
    if (!gd) return null;
    const uri = gd.attr('uri', '');
    if (uri.endsWith('/table')) {
      const tbl = gd.child('a:tbl');
      if (!tbl) return null;
      el.kind = 'table';
      el.slideNumber = this.info.number;
      el.table = parseTable(tbl, ctx, this.doc, el, [this.doc.defaultTextStyle, this.master ? this.master.otherStyle : null]);
      return el;
    }
    if (uri.endsWith('/chart')) {
      const c = gd.child('c:chart');
      const r = c && ctx.rel(c.attrs['r:id']);
      if (!r || !this.doc.pkg.has(r.partName)) {
        ctx.note('missing', 'chart part');
        return null;
      }
      el.kind = 'chart';
      el.chart = parseChartPart(this.doc, r.partName, this.baseCtx.with({ part: r.partName }), el);
      return el;
    }
    if (uri.indexOf('chartex') >= 0) {
      const c = gd.first('cx:chart');
      const r = c && ctx.rel(c.attrs['r:id']);
      if (!r || !this.doc.pkg.has(r.partName)) return null;
      el.kind = 'chart';
      el.chart = parseChartPart(this.doc, r.partName, this.baseCtx.with({ part: r.partName }), el, true);
      return el;
    }
    if (uri.endsWith('/diagram')) return this.diagram(gd, ctx, o, el);
    if (uri.endsWith('/ole')) return this.ole(gd, ctx, o, el, chain);
    if (uri.endsWith('/lockedCanvas')) {
      const lc = gd.child('lc:lockedCanvas');
      if (!lc) return null;
      return this.lockedCanvas(lc, ctx, o, el);
    }
    // Generic graphic data: look for a picture inside.
    const pic = gd.first('p:pic') || gd.first('pic:pic');
    if (pic) {
      const p = this.picture(pic, ctx, { ...o, map: null });
      if (p) { p.x = el.x; p.y = el.y; p.w = el.w; p.h = el.h; return p; }
    }
    ctx.note('unsupported', 'graphic data ' + uri);
    return null;
  }

  ole(gd, ctx, o, el, chain) {
    let ole = gd.child('p:oleObj');
    if (!ole) {
      const ac = gd.child('mc:AlternateContent');
      if (ac) {
        for (const x of chooseAlternate(ac)) if (x.local === 'oleObj') ole = x;
      }
    }
    if (!ole) return null;
    const pic = ole.child('p:pic');
    let blip = null;
    if (pic) {
      const bf = pic.child('p:blipFill');
      blip = bf ? parseBlipFill(bf, ctx) : null;
    } else if (ole.attrs.spid) {
      const image = this.vmlImage(ole.attrs.spid, ctx);
      if (image) blip = { type: 'blip', image, srcRect: null, rotWithShape: true, dpi: 0, alpha: 1, effects: [] };
      else ctx.note('approx', 'OLE object without picture preview');
    }
    el.kind = 'pic';
    el.ole = { progId: ole.attr('progId', ''), name: ole.attr('name', '') };
    el.geom = { name: 'rect', ...presetGeometry('rect', el.wEmu || 0, el.hEmu || 0, {}, 1 / 9525) };
    el.fill = { type: 'none' };
    el.line = null;
    if (pic) {
      const spPr = pic.child('p:spPr');
      if (spPr && spPr.child('a:ln')) el.line = finalizeLine(parseLine(spPr.child('a:ln'), ctx));
    }
    el.effects = null;
    el.blip = blip;
    void chain;
    return el;
  }

  /** Preview picture of a legacy object: the VML shape with the given spid. */
  vmlImage(spid, ctx) {
    const shape = this.vmlShape(spid, ctx);
    return shape ? shape.image : null;
  }

  /** VML shape by spid: its preview image and CSS-like position in px. */
  vmlShape(spid, ctx) {
    const pkg = this.doc.pkg;
    const ids = /^\d+$/.test(spid) ? ['_x0000_s' + spid, spid] : [spid];
    for (const r of pkg.relsByType(ctx.part, 'vmlDrawing')) {
      if (r.external || !pkg.has(r.partName)) continue;
      const vml = pkg.text(r.partName);
      for (const id of ids) {
        const q = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const m = new RegExp(`<v:shape\\b[^>]*\\b(?:id|o:spid)=["']${q}["'][\\s\\S]*?</v:shape>`).exec(vml);
        if (!m) continue;
        const img = /<v:imagedata\b[^>]*\b(?:o:relid|r:id)=["']([^"']+)["']/.exec(m[0]);
        const style = /\bstyle=(["'])([\s\S]*?)\1/.exec(m[0]);
        const box = {};
        const UNIT = { pt: 4 / 3, px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, emu: 1 / 9525 };
        for (const decl of (style ? style[2] : '').split(';')) {
          const [k, val] = decl.split(':').map((t) => (t || '').trim());
          const mm = /^(-?[\d.]+)([a-z]*)$/.exec(val || '');
          if (mm && ['left', 'top', 'width', 'height', 'margin-left', 'margin-top'].includes(k)) box[k] = Number(mm[1]) * (UNIT[mm[2] || 'emu'] || 1);
        }
        return { image: img ? ctx.with({ part: r.partName }).image(img[1]) : null, box };
      }
    }
    return null;
  }

  /** ActiveX controls: shown with their preview picture. */
  controls(node, ctx) {
    const out = [];
    const list = [];
    for (const c of node.children) {
      if (c.local === 'control') list.push(c);
      else if (c.local === 'AlternateContent') for (const x of chooseAlternate(c)) if (x.local === 'control') list.push(x);
    }
    for (const c of list) {
      const pic = c.child('p:pic');
      if (pic) {
        const el = this.picture(pic, ctx, { layer: 'slide' });
        if (el) out.push(el);
        continue;
      }
      const vml = c.attrs.spid ? this.vmlShape(c.attrs.spid, ctx) : null;
      if (!vml || !vml.image) { ctx.note('approx', 'ActiveX control without picture preview'); continue; }
      const b = vml.box;
      const x = b.left ?? b['margin-left'] ?? 0, y = b.top ?? b['margin-top'] ?? 0;
      const w = b.width ?? emu2px(Number(c.attrs.imgW || 0)), h = b.height ?? emu2px(Number(c.attrs.imgH || 0));
      out.push({
        kind: 'pic', id: c.attrs.spid, name: c.attrs.name || '', descr: '', hidden: false, layer: 'slide', ph: null,
        domId: 's' + this.info.index + '-s-' + c.attrs.spid + '-' + (this.idSeq++),
        x, y, w, h, wEmu: w * 9525, hEmu: h * 9525, rot: 0, flipH: false, flipV: false,
        geom: { name: 'rect', ...presetGeometry('rect', w * 9525, h * 9525, {}, 1 / 9525) },
        fill: { type: 'none' }, line: null, effects: null,
        blip: { type: 'blip', image: vml.image, srcRect: null, rotWithShape: true, dpi: 0, alpha: 1, effects: [] },
        control: { name: c.attrs.name || '' },
      });
    }
    return out;
  }

  diagram(gd, ctx, o, el) {
    const relIds = gd.child('dgm:relIds');
    if (!relIds) return null;
    const dataRel = ctx.rel(relIds.attrs['r:dm']);
    let drawingPart = null;
    if (dataRel && !dataRel.external) {
      const data = this.doc.pkg.xml(dataRel.partName);
      const ext = data && data.first('dsp:dataModelExt');
      if (ext) {
        const dr = ctx.rel(ext.attrs.relId);
        if (dr) drawingPart = dr.partName;
      }
    }
    if (!drawingPart) {
      // Fallback: any diagramDrawing relationship of the slide.
      const all = this.doc.pkg.relsByType(ctx.part, 'diagramDrawing');
      if (all.length === 1) drawingPart = all[0].partName;
    }
    const drawing = drawingPart && this.doc.pkg.xml(drawingPart);
    let tree = drawing && drawing.first('dsp:spTree');
    let dctx = ctx.with({ part: drawingPart });
    if (!tree || !tree.children.some((c) => c.local === 'sp' || c.local === 'grpSp') || this.doc.options.smartArtEngine) {
      // No pre-rendered drawing (PowerPoint 2007, other producers): lay out
      // the diagram from its layout definition.
      const part = (id) => {
        const r = id && ctx.rel(id);
        return r && !r.external ? this.doc.pkg.xml(r.partName) : null;
      };
      try {
        tree = layoutDiagram({
          data: part(relIds.attrs['r:dm']), layout: part(relIds.attrs['r:lo']),
          style: part(relIds.attrs['r:qs']), colors: part(relIds.attrs['r:cs']),
        }, el.w, el.h, ctx);
      } catch (e) {
        ctx.note('error', 'SmartArt layout: ' + e.message);
        tree = null;
      }
      if (!tree || !tree.children.some((c) => c.local === 'sp')) {
        ctx.note('approx', 'SmartArt without pre-rendered drawing (text only)');
        return this.diagramTextFallback(dataRel, ctx, el);
      }
      ctx.note('approx', 'SmartArt laid out from its layout definition');
      dctx = ctx.with({ part: dataRel ? dataRel.partName : ctx.part });
    }
    el.kind = 'group';
    el.smartArt = true;
    el.children = this.tree(tree, dctx, { ...o, map: (r) => r, layer: o.layer });
    return el;
  }

  diagramTextFallback(dataRel, ctx, el) {
    const data = dataRel && this.doc.pkg.xml(dataRel.partName);
    if (!data) return null;
    const lines = [];
    for (const pt of data.find('dgm:pt')) {
      const t = pt.child('dgm:t');
      if (!t) continue;
      const s = t.all('a:p').map((p) => p.textContent()).join(' ').trim();
      if (s) lines.push(s);
    }
    el.kind = 'shape';
    el.geom = { name: 'rect', ...presetGeometry('rect', el.wEmu || 0, el.hEmu || 0, {}, 1 / 9525) };
    el.fill = { type: 'none' };
    el.line = null;
    el.effects = null;
    el.text = {
      body: { lIns: 4, tIns: 4, rIns: 4, bIns: 4, anchor: 'ctr', anchorCtr: false, wrap: 'square', vert: 'horz', rot: 0, upright: false, numCol: 1, spcCol: 0, autofit: { type: 'none', fontScale: 1, lnSpcReduction: 0 }, warp: 'textNoShape' },
      paragraphs: lines.map((t, i) => ({
        index: i, lvl: 0, align: 'center', marL: 0, marR: 0, indent: 0, rtl: false, defTab: 96, lnSpc: null, spcBef: null, spcAft: null, tabs: [], bullet: null,
        runs: [{ kind: 'text', text: t, style: { size: 14, bold: false, italic: false, latin: ctx.theme.minorFont.latin, fill: { type: 'solid', color: ctx.color({ name: 'a:schemeClr', attrs: { val: 'tx1' }, children: [] }) }, cap: 'none', spacing: 0, baseline: 0 } }],
        endStyle: { size: 14 },
      })),
    };
    return el;
  }

  lockedCanvas(lc, ctx, o, el) {
    el.kind = 'group';
    const gpr = lc.child('a:grpSpPr');
    const xf = parseXfrm(gpr && gpr.child('a:xfrm'));
    const sub = { ...o };
    if (xf && xf.chcx && xf.chcy) {
      const sx = (el.wEmu || xf.cx) / xf.chcx;
      const sy = (el.hEmu || xf.cy) / xf.chcy;
      sub.map = (r) => ({ x: (r.x - xf.chx) * sx, y: (r.y - xf.chy) * sy, cx: r.cx * sx, cy: r.cy * sy });
    } else sub.map = (r) => r;
    el.children = [];
    for (const c of lc.children) {
      const local = c.local;
      if (local === 'sp' || local === 'cxnSp') {
        const s = this.lockedShape(c, ctx, sub);
        if (s) el.children.push(s);
      } else if (local === 'pic') {
        const p = this.picture(c, ctx, sub);
        if (p) el.children.push(p);
      } else if (local === 'grpSp') {
        const g = this.group(c, ctx, sub);
        if (g) el.children.push(g);
      }
    }
    return el;
  }

  lockedShape(node, ctx, o) {
    // a:sp inside a locked canvas: text lives in a:txSp/a:txBody.
    const el = this.shape(node, ctx, o, node.local === 'cxnSp');
    if (!el) return null;
    const txSp = node.child('a:txSp');
    if (txSp && txSp.child('a:txBody')) {
      el.text = parseTextBody(txSp.child('a:txBody'), ctx, { chain: [this.doc.defaultTextStyle], slideNumber: this.info.number });
    }
    return el;
  }

  contentPart(node, ctx, o) {
    // Ink: no fallback picture available here.
    void o;
    ctx.note('unsupported', 'ink content part (without fallback)');
    return null;
  }
}

export function makeDocHelpers(doc) {
  doc.masters = new Map();
  doc.layouts = new Map();
  doc.themes = new Map();
  doc.master = (part) => {
    if (!doc.masters.has(part)) doc.masters.set(part, new Master(doc, part));
    return doc.masters.get(part);
  };
  doc.layout = (part) => {
    if (!doc.layouts.has(part)) doc.layouts.set(part, new Layout(doc, part));
    return doc.layouts.get(part);
  };
  doc.theme = (part) => {
    const key = part || '#default';
    if (!doc.themes.has(key)) {
      const root = part ? doc.pkg.xml(part) : null;
      doc.themes.set(key, new Theme(root, part));
    }
    return doc.themes.get(key);
  };
}

export { parseFillNode, fillChild };

/**
 * Text bodies: list-style inheritance, paragraph and run properties.
 *
 * Properties are first collected into "bags" (plain objects keyed by
 * property name, values still raw) so the inheritance chain
 * (presentation default -> master text styles -> master placeholder ->
 * layout placeholder -> shape list style -> paragraph -> run) can be merged
 * property by property. Bags are finally resolved into render models.
 */
import { emu2px, pct } from '../util/units.js';
import { parseFillNode, fillChild, parseLine, parseEffects, FILL_NODES } from './fill.js';
import { resolveColorNode, COLOR_NODES, colorFrom } from './color.js';
import { mapSymbolText, isSymbolFont } from './symbols.js';
import { decodeOoxmlEscapes } from '../util/xml.js';

const PPR_ATTRS = ['marL', 'marR', 'lvl', 'indent', 'algn', 'defTabSz', 'rtl', 'eaLnBrk', 'fontAlgn', 'latinLnBrk', 'hangingPunct'];
const RPR_ATTRS = ['kumimoji', 'lang', 'altLang', 'sz', 'b', 'i', 'u', 'strike', 'kern', 'cap', 'spc', 'normalizeH', 'baseline', 'noProof', 'dirty', 'err', 'bmk'];

/** Collect a run properties element (a:rPr, a:defRPr, a:endParaRPr) into a bag. */
export function rPrBag(node) {
  const bag = {};
  if (!node) return bag;
  for (const a of RPR_ATTRS) if (node.attrs[a] !== undefined) bag[a] = node.attrs[a];
  for (const c of node.children) {
    if (FILL_NODES.has(c.name)) bag.fill = c;
    else switch (c.name) {
      case 'a:ln': bag.ln = c; break;
      case 'a:effectLst': bag.effectLst = c; break;
      case 'a:effectDag': break;
      case 'a:highlight': bag.highlight = c; break;
      case 'a:uLnTx': bag.uLn = 'tx'; break;
      case 'a:uLn': bag.uLn = c; break;
      case 'a:uFillTx': bag.uFill = 'tx'; break;
      case 'a:uFill': bag.uFill = c; break;
      case 'a:latin': bag.latin = c.attrs.typeface; if (c.attrs.pitchFamily) bag.latinPitch = c.attrs.pitchFamily; break;
      case 'a:ea': bag.ea = c.attrs.typeface; break;
      case 'a:cs': bag.cs = c.attrs.typeface; break;
      case 'a:sym': bag.sym = c.attrs.typeface; break;
      case 'a:hlinkClick': bag.hlinkClick = c; break;
      case 'a:hlinkMouseOver': bag.hlinkMouseOver = c; break;
      case 'a:rtl': bag.rtlRun = c.attr('val', '1'); break;
      default: break;
    }
  }
  return bag;
}

/** Collect paragraph properties (a:pPr, a:lvlNpPr, a:defPPr) into a bag. */
export function pPrBag(node) {
  const bag = {};
  if (!node) return bag;
  for (const a of PPR_ATTRS) if (node.attrs[a] !== undefined) bag[a] = node.attrs[a];
  for (const c of node.children) {
    switch (c.name) {
      case 'a:lnSpc': bag.lnSpc = c; break;
      case 'a:spcBef': bag.spcBef = c; break;
      case 'a:spcAft': bag.spcAft = c; break;
      case 'a:buClrTx': bag.buClr = 'tx'; break;
      case 'a:buClr': bag.buClr = c; break;
      case 'a:buSzTx': bag.buSz = 'tx'; break;
      case 'a:buSzPct': bag.buSz = c; break;
      case 'a:buSzPts': bag.buSz = c; break;
      case 'a:buFontTx': bag.buFont = 'tx'; break;
      case 'a:buFont': bag.buFont = c; break;
      case 'a:buNone': case 'a:buAutoNum': case 'a:buChar': case 'a:buBlip': bag.bu = c; break;
      case 'a:tabLst': bag.tabLst = c; break;
      case 'a:defRPr': bag.defRPr = rPrBag(c); break;
      default: break;
    }
  }
  return bag;
}

export function mergePBags(a, b) {
  if (!a) return b ? { ...b } : {};
  if (!b) return { ...a };
  const out = { ...a, ...b };
  if (a.defRPr || b.defRPr) out.defRPr = { ...(a.defRPr || {}), ...(b.defRPr || {}) };
  return out;
}

/** Parse an a:lstStyle (or p:titleStyle, p:bodyStyle, defaultTextStyle) into level bags. */
export function listStyle(node) {
  if (!node) return null;
  const ls = { def: null, lvls: [] };
  for (const c of node.children) {
    if (c.local === 'defPPr') ls.def = pPrBag(c);
    else {
      const m = /^lvl([1-9])pPr$/.exec(c.local);
      if (m) ls.lvls[Number(m[1]) - 1] = pPrBag(c);
    }
  }
  return ls;
}

/** Merge a chain of list styles (low priority first) for a given level (0-8). */
export function levelBag(chain, lvl) {
  let bag = {};
  for (const ls of chain) {
    if (!ls) continue;
    if (ls.def) bag = mergePBags(bag, ls.def);
    if (ls.lvls[lvl]) bag = mergePBags(bag, ls.lvls[lvl]);
  }
  return bag;
}

const BODY_ATTRS = ['rot', 'spcFirstLastPara', 'vertOverflow', 'horzOverflow', 'vert', 'wrap', 'lIns', 'tIns', 'rIns', 'bIns',
  'numCol', 'spcCol', 'rtlCol', 'fromWordArt', 'anchor', 'anchorCtr', 'forceAA', 'upright', 'compatLnSpc'];

export function bodyPrBag(node) {
  const bag = {};
  if (!node) return bag;
  for (const a of BODY_ATTRS) if (node.attrs[a] !== undefined) bag[a] = node.attrs[a];
  for (const c of node.children) {
    switch (c.name) {
      case 'a:noAutofit': case 'a:normAutofit': case 'a:spAutoFit': bag.autofit = c; break;
      case 'a:prstTxWarp': bag.warp = c; break;
      case 'a:scene3d': bag.scene3d = c; break;
      case 'a:sp3d': bag.sp3d = c; break;
      case 'a:flatTx': bag.flatTx = c; break;
      default: break;
    }
  }
  return bag;
}

/** Resolve merged bodyPr bag into a body model. */
export function resolveBodyPr(bag) {
  const n = (k, d) => (bag[k] === undefined ? d : Number(bag[k]));
  const body = {
    lIns: emu2px(n('lIns', 91440)),
    tIns: emu2px(n('tIns', 45720)),
    rIns: emu2px(n('rIns', 91440)),
    bIns: emu2px(n('bIns', 45720)),
    anchor: bag.anchor || 't',
    anchorCtr: bag.anchorCtr === '1' || bag.anchorCtr === 'true',
    wrap: bag.wrap || 'square',
    vert: bag.vert || 'horz',
    rot: n('rot', 0) / 60000,
    upright: bag.upright === '1' || bag.upright === 'true',
    numCol: Math.max(1, n('numCol', 1)),
    spcCol: emu2px(n('spcCol', 0)),
    rtlCol: bag.rtlCol === '1',
    autofit: { type: 'none', fontScale: 1, lnSpcReduction: 0 },
    warp: bag.warp ? bag.warp.attr('prst', 'textNoShape') : 'textNoShape',
    fromWordArt: bag.fromWordArt === '1',
    vertOverflow: bag.vertOverflow || 'overflow',
    horzOverflow: bag.horzOverflow || 'overflow',
  };
  if (bag.warp && body.warp !== 'textNoShape') {
    body.warpAdj = {};
    const av = bag.warp.child('a:avLst');
    if (av) for (const gd of av.all('a:gd')) body.warpAdj[gd.attrs.name] = Number((gd.attrs.fmla || '').split(' ')[1]) || 0;
  }
  if (bag.autofit) {
    const a = bag.autofit;
    if (a.local === 'normAutofit') {
      body.autofit = { type: 'norm', fontScale: pct(a.attrs.fontScale, 1), lnSpcReduction: pct(a.attrs.lnSpcReduction, 0) };
    } else if (a.local === 'spAutoFit') body.autofit.type = 'shape';
  }
  return body;
}

const ALGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', justLow: 'justify', dist: 'distribute', thaiDist: 'distribute' };

function spacing(node) {
  if (!node) return null;
  const p = node.child('a:spcPct');
  if (p) return { pct: pct(p.attrs.val, 1) };
  const s = node.child('a:spcPts');
  if (s) return { pt: s.num('val', 0) / 100 };
  return null;
}

/** Resolve a run property bag into a run style model. */
export function resolveRunProps(bag, ctx, opts = {}) {
  const theme = ctx.theme;
  const lang = bag.lang;
  const fontScale = opts.fontScale || 1;
  const szPt = (bag.sz !== undefined ? Number(bag.sz) / 100 : opts.defaultSize || 18);
  const st = {
    size: szPt * fontScale,
    bold: bag.b === '1' || bag.b === 'true',
    italic: bag.i === '1' || bag.i === 'true',
    underline: bag.u && bag.u !== 'none' ? bag.u : null,
    strike: bag.strike && bag.strike !== 'noStrike' ? bag.strike : null,
    cap: bag.cap || 'none',
    spacing: bag.spc !== undefined ? Number(bag.spc) / 100 : 0,
    baseline: bag.baseline !== undefined ? Number(bag.baseline) / 100000 : 0,
    kern: bag.kern !== undefined ? Number(bag.kern) / 100 : 0,
    lang: lang || null,
    latin: theme ? theme.font(bag.latin, lang) : bag.latin,
    ea: theme ? theme.font(bag.ea, lang) : bag.ea,
    cs: theme ? theme.font(bag.cs, lang) : bag.cs,
    sym: bag.sym ? (theme ? theme.font(bag.sym, lang) : bag.sym) : null,
    fill: null,
    outline: null,
    effects: null,
    highlight: null,
    uLine: null,
    uFill: null,
  };
  if (bag.fill) st.fill = parseFillNode(bag.fill, ctx);
  if (!st.fill) st.fill = { type: 'solid', color: ctx.scheme ? resolveColorNode({ name: 'a:schemeClr', attrs: { val: 'tx1' }, children: [] }, ctx) : { r: 0, g: 0, b: 0, a: 1 } };
  if (bag.ln) {
    const l = parseLine(bag.ln, ctx);
    if (l && l.fill && l.fill.type !== 'none') st.outline = { w: l.w === undefined ? 0.75 : l.w, fill: l.fill };
  }
  if (bag.effectLst) st.effects = parseEffects(bag.effectLst, ctx);
  if (bag.highlight) st.highlight = colorFrom(bag.highlight, ctx);
  if (bag.uLn && bag.uLn !== 'tx') {
    const l = parseLine(bag.uLn, ctx);
    if (l) st.uLine = l;
  }
  if (bag.uFill && bag.uFill !== 'tx') {
    const f = fillChild(bag.uFill);
    if (f) st.uFill = parseFillNode(f, ctx);
  }
  return st;
}

/** Hyperlink model from a:hlinkClick / p:cNvPr hlinkClick. */
export function parseHyperlink(node, ctx) {
  if (!node) return null;
  const id = node.attrs['r:id'];
  const action = node.attrs.action || '';
  const link = { action, tooltip: node.attrs.tooltip || '', url: null, slidePart: null, jump: null, useTextColor: false };
  for (const e of node.find('a:ext')) {
    const hc = e.children.find((c) => c.local === 'hlinkClr');
    if (hc && hc.attrs.val === 'tx') link.useTextColor = true;
  }
  if (action.startsWith('ppaction://hlinkshowjump')) {
    const m = /jump=([a-z]+)/i.exec(action);
    link.jump = m ? m[1] : 'nextslide';
  } else if (action.startsWith('ppaction://hlinksldjump')) {
    const r = ctx.rel(id);
    if (r) link.slidePart = r.partName;
  } else if (action.startsWith('ppaction://noaction') || action.startsWith('ppaction://media')) {
    link.noAction = action.startsWith('ppaction://noaction');
    if (action.startsWith('ppaction://media')) link.media = true;
  } else if (action.startsWith('ppaction://customshow')) {
    const m = /id=(\d+)/.exec(action);
    link.customShow = m ? Number(m[1]) : 0;
    link.jump = 'customshow';
  } else if (action.startsWith('ppaction://program') || action.startsWith('ppaction://macro') || action.startsWith('ppaction://ole')) {
    link.noAction = true;
  } else if (id) {
    const r = ctx.rel(id);
    if (r) {
      if (r.external) link.url = r.target;
      else if (r.type === 'slide') link.slidePart = r.partName;
      else link.url = r.target;
    }
  }
  if (!link.url && !link.slidePart && !link.jump && !link.media && !link.customShow) {
    if (link.noAction) return null;
    if (!id && !action) return null;
  }
  const snd = node.child('a:snd');
  if (snd) link.sound = ctx.mediaRel(snd.attrs['r:embed'], 'audio');
  return link;
}

/** Auto-numbering formatter (ST_TextAutonumberScheme). */
export function autoNumber(scheme, n) {
  const roman = (num, upper) => {
    if (num <= 0 || num >= 4000) return String(num);
    const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let s = '';
    for (const [v, r] of map) while (num >= v) { s += r; num -= v; }
    return upper ? s.toUpperCase() : s;
  };
  const alpha = (num, upper) => {
    let s = '';
    let k = num;
    while (k > 0) { k--; s = String.fromCharCode(97 + (k % 26)) + s; k = Math.floor(k / 26); }
    return upper ? s.toUpperCase() : s;
  };
  const circled = (num) => {
    if (num >= 1 && num <= 20) return String.fromCodePoint(0x2460 + num - 1);
    if (num >= 21 && num <= 35) return String.fromCodePoint(0x3251 + num - 21);
    if (num >= 36 && num <= 50) return String.fromCodePoint(0x32b1 + num - 36);
    return String(num);
  };
  const fullwidth = (s) => s.replace(/[0-9]/g, (d) => String.fromCharCode(0xff10 + Number(d)));
  const cjk = (num) => {
    const digits = '〇一二三四五六七八九';
    if (num < 10) return digits[num];
    if (num < 20) return '十' + (num % 10 ? digits[num % 10] : '');
    if (num < 100) return digits[Math.floor(num / 10)] + '十' + (num % 10 ? digits[num % 10] : '');
    return String(num);
  };
  let base;
  let suffix;
  const s = scheme || 'arabicPeriod';
  if (s.startsWith('alphaLc')) { base = alpha(n, false); suffix = s.slice(7); }
  else if (s.startsWith('alphaUc')) { base = alpha(n, true); suffix = s.slice(7); }
  else if (s.startsWith('romanLc')) { base = roman(n, false); suffix = s.slice(7); }
  else if (s.startsWith('romanUc')) { base = roman(n, true); suffix = s.slice(7); }
  else if (s.startsWith('circleNum')) { base = circled(n); suffix = ''; }
  else if (s.startsWith('arabicDb')) { base = fullwidth(String(n)); suffix = s.slice(8); }
  else if (s.startsWith('arabic')) { base = String(n); suffix = s.slice(6); }
  else if (s.startsWith('ea1') || s.startsWith('thai') || s.startsWith('hindi') || s.startsWith('hebrew')) {
    base = s.startsWith('ea1') ? cjk(n) : String(n);
    suffix = /Period/.test(s) ? 'Period' : /Plain/.test(s) ? 'Plain' : 'Period';
    if (s.startsWith('ea1') && suffix === 'Period') return base + '．';
  } else { base = String(n); suffix = 'Period'; }
  switch (suffix) {
    case 'Period': return base + '.';
    case 'ParenR': return base + ')';
    case 'ParenBoth': return '(' + base + ')';
    case 'Plain': return base;
    case 'Minus': case '1Minus': case '2Minus': return '- ' + base + ' -';
    case 'Paren': return base + ')';
    default: return base + (suffix ? '' : '');
  }
}

/** Format the current date for a datetimeN field. */
export function formatDateField(type, date, lang) {
  const d = date || new Date();
  const pad = (v) => String(v).padStart(2, '0');
  const Y = d.getFullYear(), M = d.getMonth() + 1, D = d.getDate();
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const h = d.getHours(), mi = d.getMinutes(), s = d.getSeconds();
  const h12 = ((h + 11) % 12) + 1;
  const ap = h < 12 ? 'AM' : 'PM';
  void lang;
  switch (type) {
    case 'datetime':
    case 'datetime1': return `${M}/${D}/${Y}`;
    case 'datetime2': return `${days[d.getDay()]}, ${months[M - 1]} ${D}, ${Y}`;
    case 'datetime3': return `${D} ${months[M - 1]} ${Y}`;
    case 'datetime4': return `${months[M - 1]} ${D}, ${Y}`;
    case 'datetime5': return `${D}-${months[M - 1].slice(0, 3)}-${String(Y).slice(2)}`;
    case 'datetime6': return `${months[M - 1]} ${String(Y).slice(2)}`;
    case 'datetime7': return `${months[M - 1].slice(0, 3)}-${String(Y).slice(2)}`;
    case 'datetime8': return `${M}/${D}/${Y} ${h12}:${pad(mi)} ${ap}`;
    case 'datetime9': return `${M}/${D}/${Y} ${h12}:${pad(mi)}:${pad(s)} ${ap}`;
    case 'datetime10': return `${h}:${pad(mi)}`;
    case 'datetime11': return `${h}:${pad(mi)}:${pad(s)}`;
    case 'datetime12': return `${h12}:${pad(mi)} ${ap}`;
    case 'datetime13': return `${h12}:${pad(mi)}:${pad(s)} ${ap}`;
    case 'datetimeFigureOut': return `${M}/${D}/${Y}`;
    default: return `${M}/${D}/${Y}`;
  }
}

/**
 * Parse a text body.
 * @param {import('../util/xml.js').XNode} txBody p:txBody / a:txBody / c:rich
 * @param {import('./context.js').Ctx} ctx
 * @param {object} o
 *   chain: list styles, low priority first (excluding the body's own lstStyle)
 *   bodyChain: bodyPr bags, low priority first (excluding the body's own)
 *   fontRef: {font, color} from shape style
 *   slideNumber: number for slidenum fields
 */
export function parseTextBody(txBody, ctx, o = {}) {
  if (!txBody) return null;
  const ownBody = txBody.child('a:bodyPr');
  let bag = {};
  for (const b of o.bodyChain || []) bag = { ...bag, ...b };
  bag = { ...bag, ...bodyPrBag(ownBody) };
  const body = resolveBodyPr(bag);
  const chain = [...(o.chain || [])];
  const own = listStyle(txBody.child('a:lstStyle'));
  if (o.fontRef) chain.push(o.fontRef.style);
  if (own) chain.push(own);
  const fontScale = body.autofit.type === 'norm' ? body.autofit.fontScale : 1;
  const lnSpcReduction = body.autofit.type === 'norm' ? body.autofit.lnSpcReduction : 0;

  const paragraphs = [];
  const numbering = []; // per level: {scheme, n}
  let paraIndex = 0;
  for (const p of txBody.all('a:p')) {
    const pPr = p.child('a:pPr');
    const own = pPrBag(pPr);
    const lvl = Math.max(0, Math.min(8, Number(own.lvl || 0)));
    const base = levelBag(chain.filter(Boolean), lvl);
    const pb = mergePBags(base, own);
    const defR = pb.defRPr || {};
    const runOpts = { fontScale, defaultSize: 18 };
    const runs = [];
    const items = flattenRuns(p);
    for (const r of items) {
      if (r.name === 'a:br') {
        const rb = { ...defR, ...rPrBag(r.child('a:rPr')) };
        runs.push({ kind: 'br', style: resolveRunProps(rb, ctx, runOpts) });
        continue;
      }
      const rb = { ...defR, ...rPrBag(r.child('a:rPr')) };
      const style = resolveRunProps(rb, ctx, runOpts);
      const tNode = r.child('a:t');
      let text = tNode ? decodeOoxmlEscapes(tNode.text) : '';
      let kind = 'text';
      let field = null;
      if (r.name === 'a:fld') {
        field = r.attr('type', '');
        kind = 'field';
        if (field === 'slidenum' && o.slideNumber !== undefined) text = String(o.slideNumber);
        else if (field.startsWith('datetime') && o.liveDates) text = formatDateField(field, null, style.lang);
      }
      if (style.sym || isSymbolFont(style.latin)) {
        const font = style.sym || style.latin;
        const mapped = mapSymbolText(text, font);
        if (mapped !== text) style.symbolMapped = true;
        text = mapped;
      }
      const link = rb.hlinkClick ? parseHyperlink(rb.hlinkClick, ctx) : null;
      const hover = rb.hlinkMouseOver ? parseHyperlink(rb.hlinkMouseOver, ctx) : null;
      if (link && !link.useTextColor && (link.url || link.slidePart || link.jump)) {
        // PowerPoint shows hyperlinks with the theme hyperlink colour.
        const hl = ctx.scheme ? resolveColorNode({ name: 'a:schemeClr', attrs: { val: 'hlink' }, children: [] }, ctx) : null;
        if (hl) style.fill = { type: 'solid', color: hl };
        if (!style.underline) style.underline = 'sng';
      }
      runs.push({ kind, text, style, field, link, hover });
    }
    const endBag = { ...defR, ...rPrBag(p.child('a:endParaRPr')) };
    const endStyle = resolveRunProps(endBag, ctx, runOpts);

    // Paragraph level properties.
    const para = {
      index: paraIndex++,
      lvl,
      align: ALGN[pb.algn] || 'left',
      marL: emu2px(Number(pb.marL || 0)),
      marR: emu2px(Number(pb.marR || 0)),
      indent: emu2px(Number(pb.indent || 0)),
      rtl: pb.rtl === '1',
      defTab: emu2px(Number(pb.defTabSz || 914400)),
      lnSpc: spacing(pb.lnSpc),
      spcBef: spacing(pb.spcBef),
      spcAft: spacing(pb.spcAft),
      fontAlgn: pb.fontAlgn || 'auto',
      eaLnBrk: pb.eaLnBrk !== '0',
      latinLnBrk: pb.latinLnBrk === '1',
      tabs: [],
      bullet: null,
      runs,
      endStyle,
    };
    if (para.lnSpc && para.lnSpc.pct !== undefined && lnSpcReduction) {
      para.lnSpc = { pct: Math.max(0.1, para.lnSpc.pct - lnSpcReduction) };
    } else if (!para.lnSpc && lnSpcReduction) {
      para.lnSpc = { pct: Math.max(0.1, 1 - lnSpcReduction) };
    }
    if (pb.tabLst) {
      for (const t of pb.tabLst.all('a:tab')) para.tabs.push({ pos: emu2px(t.num('pos', 0)), algn: t.attr('algn', 'l') });
    }
    const hasText = runs.some((r) => r.kind !== 'br' && r.text.length > 0);
    // Bullets.
    const bu = pb.bu;
    if (bu && bu.name !== 'a:buNone' && hasText) {
      const first = runs.find((r) => r.kind !== 'br') || { style: endStyle };
      const bullet = { kind: 'char', char: '•', font: null, color: null, size: 1, sizePt: null };
      if (bu.name === 'a:buChar') {
        bullet.char = bu.attr('char', '•');
      } else if (bu.name === 'a:buAutoNum') {
        bullet.kind = 'num';
        const scheme = bu.attr('type', 'arabicPeriod');
        const startAt = bu.int('startAt', 1);
        const prev = numbering[lvl];
        let n = startAt;
        if (prev && prev.scheme === scheme && prev.startAt === startAt) n = prev.n + 1;
        numbering[lvl] = { scheme, n, startAt };
        bullet.char = autoNumber(scheme, n);
        bullet.scheme = scheme;
        bullet.number = n;
      } else if (bu.name === 'a:buBlip') {
        bullet.kind = 'image';
        const blip = bu.first('a:blip');
        bullet.image = blip ? ctx.image(blip.attrs['r:embed']) : null;
        if (!bullet.image) bullet.kind = 'char';
      }
      // Font.
      if (pb.buFont && pb.buFont !== 'tx') {
        const tf = pb.buFont.attrs.typeface;
        bullet.font = ctx.theme ? ctx.theme.font(tf) : tf;
      } else {
        bullet.font = first.style.latin;
      }
      if (bullet.kind === 'char' && isSymbolFont(bullet.font)) bullet.char = mapSymbolText(bullet.char, bullet.font);
      // Colour.
      if (pb.buClr && pb.buClr !== 'tx') bullet.color = colorFrom(pb.buClr, ctx);
      else bullet.color = first.style.fill && first.style.fill.type === 'solid' ? first.style.fill.color : null;
      bullet.fill = bullet.color ? null : first.style.fill;
      // Size.
      if (pb.buSz && pb.buSz !== 'tx') {
        if (pb.buSz.local === 'buSzPct') bullet.size = pct(pb.buSz.attrs.val, 1);
        else bullet.sizePt = pb.buSz.num('val', 1200) / 100 * fontScale;
      }
      bullet.baseSize = first.style.size;
      bullet.bold = first.style.bold;
      para.bullet = bullet;
    }
    // Numbering restarts after a paragraph of the same or lower level without numbering.
    if (!(para.bullet && para.bullet.kind === 'num')) {
      if (hasText) for (let k = lvl; k < 9; k++) numbering[k] = undefined;
    } else {
      for (let k = lvl + 1; k < 9; k++) numbering[k] = undefined;
    }
    paragraphs.push(para);
  }
  return { body, paragraphs };
}

/** Handle mc:AlternateContent inside paragraphs (equations): use the fallback runs. */
function flattenRuns(p) {
  const out = [];
  for (const c of p.children) {
    if (c.name === 'a:r' || c.name === 'a:br' || c.name === 'a:fld') out.push(c);
    else if (c.name === 'mc:AlternateContent') {
      const fb = c.child('mc:Fallback');
      if (fb) for (const x of fb.children) if (x.name === 'a:r' || x.name === 'a:br' || x.name === 'a:fld') out.push(x);
    } else if (c.name === 'a14:m') {
      // Office math without fallback: show its plain text.
      const t = c.textContent();
      if (t) out.push({ name: 'a:r', child: (n) => (n === 'a:t' ? { text: t } : null), attrs: {}, children: [] });
    }
  }
  return out;
}

/** Whether a text body has any visible text. */
export function hasVisibleText(tb) {
  if (!tb) return false;
  return tb.paragraphs.some((p) => p.runs.some((r) => r.kind !== 'br' && r.text.trim().length));
}

/** Font reference from a shape style (<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>). */
export function fontRefInfo(styleNode, ctx) {
  const fr = styleNode && styleNode.child('a:fontRef');
  if (!fr) return null;
  const idx = fr.attr('idx', 'none');
  const font = idx === 'none' ? null : (idx === 'major' ? '+mj-lt' : '+mn-lt');
  const colorNode = fr.children.find((c) => COLOR_NODES.has(c.name));
  const color = colorNode ? resolveColorNode(colorNode, ctx) : null;
  // Express the font reference as a low priority list style layer so that
  // explicit list styles still win.
  const defRPr = {};
  if (font) {
    defRPr.latin = font;
    defRPr.ea = idx === 'major' ? '+mj-ea' : '+mn-ea';
    defRPr.cs = idx === 'major' ? '+mj-cs' : '+mn-cs';
  }
  if (colorNode) defRPr.fill = { name: 'a:solidFill', local: 'solidFill', attrs: {}, children: [colorNode], child() { return null; }, all() { return []; } };
  return { font, color, style: { def: { defRPr }, lvls: [] } };
}


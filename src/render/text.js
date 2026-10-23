/**
 * Text body rendering (HTML).
 */
import { escapeXml } from '../util/xml.js';
import { css } from '../pptx/color.js';
import { r2, r4, PX_PER_PT } from '../util/units.js';
import { fontStack } from './fonts.js';
import { cssBackground, fillColor } from './paint.js';

const LINE = 1.2; // single line spacing relative to the font size

const UNDERLINE = {
  sng: ['solid', 1], dbl: ['double', 1], heavy: ['solid', 2], dotted: ['dotted', 1], dottedHeavy: ['dotted', 2],
  dash: ['dashed', 1], dashHeavy: ['dashed', 2], dashLong: ['dashed', 1], dashLongHeavy: ['dashed', 2],
  dotDash: ['dashed', 1], dotDashHeavy: ['dashed', 2], dotDotDash: ['dashed', 1], dotDotDashHeavy: ['dashed', 2],
  wavy: ['wavy', 1], wavyHeavy: ['wavy', 2], wavyDbl: ['wavy', 1], words: ['solid', 1],
};

function textShadowCss(fx) {
  if (!fx) return '';
  const parts = [];
  if (fx.outerShadow) {
    const s = fx.outerShadow;
    parts.push(`${r2(s.dx)}px ${r2(s.dy)}px ${r2(s.blur)}px ${css(s.color)}`);
  }
  if (fx.glow && fx.glow.rad > 0) {
    const g = fx.glow;
    parts.push(`0 0 ${r2(g.rad / 2)}px ${css(g.color)}`, `0 0 ${r2(g.rad)}px ${css(g.color)}`);
  }
  return parts.join(',');
}

/** Inline CSS for a run style. */
export function runCss(st, rc, o = {}) {
  let s = '';
  let size = st.size * PX_PER_PT;
  if (st.baseline) size *= 2 / 3;
  s += `font-size:${r2(size)}px;`;
  s += `font-family:${fontStack(st.latin, st.ea, st.cs, { symbol: st.symbolMapped })};`;
  if (st.bold) s += 'font-weight:bold;';
  if (st.italic) s += 'font-style:italic;';
  // Fill.
  const f = st.fill;
  if (f && f.type === 'solid') s += `color:${css(f.color)};`;
  else if (f && f.type === 'none') s += 'color:transparent;';
  else if (f && (f.type === 'grad' || f.type === 'blip' || f.type === 'patt')) {
    const bg = cssBackground(f, rc);
    const c = fillColor(f);
    s += `color:${c ? css(c) : '#000'};background:${bg};-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;`;
  }
  const deco = [];
  if (st.underline) deco.push('underline');
  if (st.strike) deco.push('line-through');
  if (deco.length) {
    s += `text-decoration-line:${deco.join(' ')};`;
    const u = UNDERLINE[st.underline] || ['solid', 1];
    if (st.strike === 'dblStrike' && !st.underline) s += 'text-decoration-style:double;';
    else if (st.underline && u[0] !== 'solid') s += `text-decoration-style:${u[0]};`;
    if (st.underline && u[1] > 1) s += 'text-decoration-thickness:0.12em;';
    if (st.underline === 'words') s += 'text-decoration-skip-ink:auto;';
    const uc = st.uFill ? fillColor(st.uFill) : st.uLine && st.uLine.fill ? fillColor(st.uLine.fill) : null;
    if (uc) s += `text-decoration-color:${css(uc)};`;
  }
  if (st.cap === 'all') s += 'text-transform:uppercase;';
  else if (st.cap === 'small') s += 'font-variant:small-caps;';
  if (st.spacing) s += `letter-spacing:${r2(st.spacing * PX_PER_PT)}px;`;
  if (st.baseline) s += `vertical-align:${r2(st.baseline * st.size * PX_PER_PT)}px;`;
  if (st.highlight) s += `background-color:${css(st.highlight)};`;
  if (st.outline) {
    const oc = fillColor(st.outline.fill);
    if (oc) s += `-webkit-text-stroke:${r2(Math.max(0.3, st.outline.w))}px ${css(oc)};paint-order:stroke fill;`;
  }
  const sh = textShadowCss(st.effects || o.shapeEffects);
  if (sh) s += `text-shadow:${sh};`;
  if (st.effects && st.effects.reflection) s += '-webkit-box-reflect:below 0 linear-gradient(transparent 50%,rgba(255,255,255,.35));';
  if (st.rtlRun === '1') s += 'unicode-bidi:embed;direction:rtl;';
  return s;
}

function linkAttrs(link) {
  if (!link) return null;
  if (link.url) {
    const safe = /^(https?:|mailto:|ftp:|tel:|#|file:)/i.test(link.url) ? link.url : (/^[a-z][a-z0-9+.-]*:/i.test(link.url) ? '#' : link.url);
    return `href="${escapeXml(safe)}" target="_blank" rel="noopener noreferrer"${link.tooltip ? ` title="${escapeXml(link.tooltip)}"` : ''}`;
  }
  if (link.slide !== undefined && link.slide !== null) return `href="#slide-${link.slide + 1}" data-slide="${link.slide}"${link.tooltip ? ` title="${escapeXml(link.tooltip)}"` : ''}`;
  if (link.jump) return `href="#" data-jump="${escapeXml(link.jump)}"${link.customShow !== undefined ? ` data-show="${link.customShow}"` : ''}`;
  return null;
}

function lineHeight(para, firstSize) {
  if (!para.lnSpc) return { css: `line-height:${LINE};` };
  if (para.lnSpc.pct !== undefined) return { css: `line-height:${r4(LINE * para.lnSpc.pct)};` };
  const px = para.lnSpc.pt * PX_PER_PT;
  void firstSize;
  return { css: `line-height:${r2(px)}px;`, px };
}

function spacingPx(sp, size) {
  if (!sp) return 0;
  if (sp.pt !== undefined) return sp.pt * PX_PER_PT;
  return sp.pct * size * PX_PER_PT * LINE;
}

function bulletHtml(para, rc) {
  const b = para.bullet;
  if (!b) return '';
  const base = b.baseSize || (para.runs[0] && para.runs[0].style.size) || 18;
  const sizePt = b.sizePt || base * (b.size || 1);
  const indent = para.indent;
  const minW = indent < 0 ? -indent : 0;
  let s = `font-size:${r2(sizePt * PX_PER_PT)}px;`;
  if (b.kind !== 'image') s += `font-family:${fontStack(b.font, null, null, { symbol: false })};`;
  if (b.color) s += `color:${css(b.color)};`;
  if (b.kind === 'num' && b.bold) s += 'font-weight:bold;';
  if (minW) s += `min-width:${r2(minW)}px;`;
  let content;
  if (b.kind === 'image') {
    const u = rc.url(b.image);
    const sz = r2(sizePt * PX_PER_PT);
    content = `<img src="${u}" alt="" style="width:${sz}px;height:${sz}px;vertical-align:middle">`;
  } else content = escapeXml(b.char);
  return `<span class="pp-bu" style="${s}">${content}</span>${minW ? '' : '<span class="pp-bu-gap"> </span>'}`;
}

/**
 * Render paragraphs of a text body.
 * @param {object} tb text body model
 */
export function paragraphsHtml(tb, rc, o = {}) {
  let html = '';
  const paras = tb.paragraphs;
  const n = paras.length;
  for (let i = 0; i < n; i++) {
    const p = paras[i];
    const firstRun = p.runs.find((r) => r.kind !== 'br');
    const firstSize = firstRun ? firstRun.style.size : p.endStyle.size;
    const lh = lineHeight(p, firstSize);
    let ps = `text-align:${p.align === 'distribute' ? 'justify' : p.align};`;
    if (p.align === 'distribute') ps += 'text-align-last:justify;';
    if (p.rtl) ps += 'direction:rtl;';
    ps += lh.css;
    if (p.marL) ps += p.rtl ? `padding-right:${r2(p.marL)}px;` : `padding-left:${r2(p.marL)}px;`;
    if (p.marR) ps += p.rtl ? `padding-left:${r2(p.marR)}px;` : `padding-right:${r2(p.marR)}px;`;
    if (p.indent) ps += `text-indent:${r2(p.indent)}px;`;
    const before = spacingPx(p.spcBef, firstSize);
    const after = spacingPx(p.spcAft, firstSize);
    if (before) ps += `margin-top:${r2(before)}px;`;
    if (after) ps += `margin-bottom:${r2(after)}px;`;
    ps += `tab-size:${r2(p.defTab)}px;`;
    let inner = '';
    let hasContent = false;
    for (const r of p.runs) {
      if (r.kind === 'br') {
        inner += `<br>`;
        continue;
      }
      if (!r.text) continue;
      hasContent = true;
      const st = runCss(r.style, rc, o);
      const txt = escapeXml(r.text).replace(/\v/g, '<br>');
      const span = `<span style="${st}"${r.field ? ` data-field="${escapeXml(r.field)}"` : ''}>${txt}</span>`;
      const la = linkAttrs(r.link);
      inner += la ? `<a class="pp-a" ${la}>${span}</a>` : span;
    }
    // Trailing line break or empty paragraph: keep the line height using the end run size.
    const endCss = `font-size:${r2(p.endStyle.size * PX_PER_PT)}px;font-family:${fontStack(p.endStyle.latin, p.endStyle.ea, p.endStyle.cs)};`;
    if (!hasContent || (p.runs.length && p.runs[p.runs.length - 1].kind === 'br')) inner += `<span style="${endCss}">&#8203;</span>`;
    const bullet = hasContent ? bulletHtml(p, rc) : '';
    html += `<div class="pp-p" data-para="${p.index}" style="${ps}">${bullet}${inner}</div>`;
  }
  return html;
}

const ANCHOR = { t: 'flex-start', ctr: 'center', b: 'flex-end', just: 'center', dist: 'center' };

/**
 * Render a text body positioned in a box.
 * @param {object} tb text body model
 * @param {{x:number,y:number,w:number,h:number}} box text rectangle (local px)
 * @param {object} o {flipH, flipV, rot, shapeEffects}
 */
export function textBodyHtml(tb, box, rc, o = {}) {
  if (!tb || !tb.paragraphs.length) return '';
  const b = tb.body;
  let x = box.x, y = box.y, w = box.w, h = box.h;
  const vertical = b.vert && b.vert !== 'horz';
  // Insets.
  let s = `left:${r2(x)}px;top:${r2(y)}px;width:${r2(Math.max(0, w))}px;height:${r2(Math.max(0, h))}px;`;
  s += `padding:${r2(b.tIns)}px ${r2(b.rIns)}px ${r2(b.bIns)}px ${r2(b.lIns)}px;`;
  s += `justify-content:${ANCHOR[b.anchor] || 'flex-start'};`;
  const tr = [];
  // Counter flips so text stays readable (flipV results in 180 degree text).
  if (!!o.flipH !== !!o.flipV) tr.push('scaleX(-1)');
  if (b.rot) tr.push(`rotate(${r2(b.rot)}deg)`);
  if (b.upright && o.rot) tr.push(`rotate(${r2(-o.rot)}deg)`);
  if (b.vert === 'vert270') tr.push('rotate(180deg)');
  if (tr.length) s += `transform:${tr.join(' ')};`;
  let cls = 'pp-tx';
  if (vertical) {
    if (b.vert === 'eaVert' || b.vert === 'mongolianVert') s += `writing-mode:${b.vert === 'mongolianVert' ? 'vertical-lr' : 'vertical-rl'};text-orientation:mixed;`;
    else if (b.vert === 'wordArtVert' || b.vert === 'wordArtVertRtl') s += 'writing-mode:vertical-rl;text-orientation:upright;';
    else s += 'writing-mode:vertical-rl;text-orientation:sideways;';
  }
  if (b.wrap === 'none') cls += ' pp-nowrap';
  if (b.anchorCtr) cls += ' pp-actr';
  let inner = paragraphsHtml(tb, rc, o);
  if (b.numCol > 1) inner = `<div class="pp-cols" style="column-count:${b.numCol};column-gap:${r2(b.spcCol)}px;height:100%">${inner}</div>`;
  else if (b.anchorCtr) inner = `<div class="pp-actr-in">${inner}</div>`;
  const fit = b.autofit && b.autofit.type === 'norm' ? ' data-autofit="1"' : '';
  return `<div class="${cls}" style="${s}"${fit}>${inner}</div>`;
}

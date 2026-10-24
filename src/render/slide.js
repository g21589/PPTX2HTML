/**
 * Slide and element rendering.
 */
import { escapeXml } from '../util/xml.js';
import { r2 } from '../util/units.js';
import { css } from '../pptx/color.js';
import { geometrySvg } from './shape.js';
import { textBodyHtml } from './text.js';
import { svgPaint, cssBackground } from './paint.js';
import { tableHtml } from './table.js';
import { chartHtml } from '../chart/render.js';

function transformCss(el) {
  const t = [];
  if (el.rot) t.push(`rotate(${r2(el.rot)}deg)`);
  if (el.flipH || el.flipV) t.push(`scale(${el.flipH ? -1 : 1},${el.flipV ? -1 : 1})`);
  return t.length ? `transform:${t.join(' ')};` : '';
}

function boxCss(el) {
  return `left:${r2(el.x)}px;top:${r2(el.y)}px;width:${r2(Math.max(0, el.w))}px;height:${r2(Math.max(0, el.h))}px;` + transformCss(el);
}

function dataAttrs(el) {
  let a = ` id="${el.domId}" data-spid="${escapeXml(el.id)}" data-layer="${el.layer}"`;
  if (el.name) a += ` data-name="${escapeXml(el.name)}"`;
  const l = el.link;
  if (l) {
    if (l.url) a += ` data-href="${escapeXml(l.url)}"`;
    if (l.slide !== undefined && l.slide !== null) a += ` data-slide="${l.slide}"`;
    if (l.jump) a += ` data-jump="${escapeXml(l.jump)}"`;
    if (l.customShow !== undefined) a += ` data-show="${l.customShow}"`;
  }
  if (el.ph) a += ` data-ph="${escapeXml(el.ph.type)}"`;
  if (el.rot) a += ` data-rot="${r2(el.rot)}"`;
  if (el.flipH) a += ' data-fh="1"';
  if (el.flipV) a += ' data-fv="1"';
  return a;
}

function linkOverlay(el) {
  const l = el.link;
  if (!l) return '';
  if (l.url) {
    const safe = /^(https?:|mailto:|ftp:|tel:|file:)/i.test(l.url) ? l.url : '#';
    return `<a class="pp-hit" href="${escapeXml(safe)}" target="_blank" rel="noopener noreferrer"${l.tooltip ? ` title="${escapeXml(l.tooltip)}"` : ''}></a>`;
  }
  if (l.slide !== undefined && l.slide !== null) return `<a class="pp-hit" href="#slide-${l.slide + 1}" data-slide="${l.slide}"></a>`;
  if (l.jump) return `<a class="pp-hit" href="#" data-jump="${escapeXml(l.jump)}"></a>`;
  return '';
}

function textBox(el) {
  if (el.textBox) return el.textBox;
  const tr = el.geom ? el.geom.textRect : null;
  if (!tr) return { x: 0, y: 0, w: el.w, h: el.h };
  return { x: tr.l, y: tr.t, w: tr.r - tr.l, h: tr.b - tr.t };
}

function shapeHtml(el, rc) {
  const geo = geometrySvg(el, rc);
  let text = '';
  if (el.text) {
    const noVisual = (!el.fill || el.fill.type === 'none') && !el.line;
    const box = textBox(el);
    let tb = textBodyHtml(el.text, box, rc, {
      flipH: el.flipH, flipV: el.flipV, rot: el.rot,
      shapeEffects: noVisual ? el.effects : null,
    });
    if (el.textBox && el.textBox.rot) tb = `<div class="pp-txrot" style="position:absolute;left:0;top:0;width:100%;height:100%;transform-origin:${r2(box.x + box.w / 2)}px ${r2(box.y + box.h / 2)}px;transform:rotate(${r2(el.textBox.rot - (el.rot || 0))}deg)">${tb}</div>`;
    text = tb;
  }
  if (!geo && !text) return '';
  const cls = el.connector ? 'pp-el pp-sp pp-cxn' : 'pp-el pp-sp';
  return `<div class="${cls}"${dataAttrs(el)} style="${boxCss(el)}">${geo}${text}${linkOverlay(el)}</div>`;
}

function mediaHtml(el, rc) {
  const m = el.media;
  const src = m.src ? rc.url(m.src) : '';
  const poster = el.blip && el.blip.image ? rc.url(el.blip.image) : '';
  const trim = m.trim ? ` data-trim-start="${m.trim.st}" data-trim-end="${m.trim.end}"` : '';
  let inner;
  if (m.kind === 'video') {
    inner = `<video class="pp-video" preload="metadata" playsinline${poster ? ` poster="${poster}"` : ''}${trim} src="${src}"></video>`;
  } else {
    inner = `${poster ? `<img class="pp-media-icon" alt="" src="${poster}">` : ''}<audio class="pp-audio" preload="metadata"${trim} src="${src}"></audio>`;
  }
  const geo = el.line || el.effects ? geometrySvg({ ...el, fill: { type: 'none' } }, rc) : '';
  return `<div class="pp-el pp-media pp-${m.kind}"${dataAttrs(el)} data-media="${m.kind}" style="${boxCss(el)}">${inner}${geo}<button class="pp-media-btn" type="button" aria-label="Play">&#9654;</button></div>`;
}

function picHtml(el, rc) {
  if (el.media) return mediaHtml(el, rc);
  const blip = el.blip;
  const geo = geometrySvg(el, rc, { blip: blip && blip.image ? blip : { type: 'none' } });
  if (!geo) return '';
  const alt = el.descr ? ` role="img" aria-label="${escapeXml(el.descr)}"` : '';
  return `<div class="pp-el pp-pic"${dataAttrs(el)}${alt} style="${boxCss(el)}">${geo}${linkOverlay(el)}</div>`;
}

function groupHtml(el, rc) {
  // Children are positioned relative to the group box.
  const origin = rc.origin || { x: 0, y: 0 };
  rc.origin = { x: origin.x + el.x, y: origin.y + el.y };
  let kids;
  try {
    kids = (el.children || []).map((c) => elementHtml(c, rc)).join('');
  } finally {
    rc.origin = origin;
  }
  if (!kids) return '';
  let st = boxCss(el);
  const fx = el.effects;
  if (fx && fx.outerShadow) {
    const s = fx.outerShadow;
    st += `filter:drop-shadow(${r2(s.dx)}px ${r2(s.dy)}px ${r2(s.blur / 2)}px ${css(s.color)});`;
  }
  return `<div class="pp-el pp-grp${el.smartArt ? ' pp-smartart' : ''}"${dataAttrs(el)} style="${st}">${kids}${linkOverlay(el)}</div>`;
}

export function elementHtml(el, rc) {
  try {
    switch (el.kind) {
      case 'shape': return shapeHtml(el, rc);
      case 'pic': return picHtml(el, rc);
      case 'group': return groupHtml(el, rc);
      case 'table': return `<div class="pp-el pp-tbl"${dataAttrs(el)} style="${boxCss(el)}">${tableHtml(el, rc)}</div>`;
      case 'chart': return `<div class="pp-el pp-chart"${dataAttrs(el)} style="${boxCss(el)}">${chartHtml(el, rc)}</div>`;
      default: return '';
    }
  } catch (e) {
    if (rc.opts.strict) throw e;
    rc.pres.diagnostics.add('error', 'render ' + el.kind + ': ' + e.message);
    return '';
  }
}

export function backgroundHtml(slide, rc, width, height) {
  const bg = slide.background;
  if (!bg || !bg.fill || bg.fill.type === 'none') return '<div class="pp-bg" style="background:#fff"></div>';
  const f = bg.fill;
  if (f.type === 'solid') return `<div class="pp-bg" style="background:${css(f.color)}"></div>`;
  const defs = [];
  const p = svgPaint(f, width, height, rc, defs, {});
  if (p.paint === 'none') {
    const cb = cssBackground(f, rc, width, height);
    return `<div class="pp-bg" style="background:${cb || '#fff'}"></div>`;
  }
  return `<div class="pp-bg"><svg width="${r2(width)}" height="${r2(height)}" viewBox="0 0 ${r2(width)} ${r2(height)}" aria-hidden="true"><defs>${defs.join('')}</defs><rect width="${r2(width)}" height="${r2(height)}" fill="${p.paint}"${p.opacity !== undefined ? ` fill-opacity="${p.opacity}"` : ''}/></svg></div>`;
}

/** HTML for one slide (the slide box itself, without the deck wrapper). */
export function slideHtml(slide, rc) {
  const { width, height } = rc.pres;
  rc.slide = slide;
  rc.origin = { x: 0, y: 0 };
  let html = backgroundHtml(slide, rc, width, height);
  for (const el of slide.elements) html += elementHtml(el, rc);
  const attrs = [
    `class="pp-slide"`,
    `id="slide-${slide.index + 1}"`,
    `data-index="${slide.index}"`,
    `data-number="${slide.number}"`,
    slide.hidden ? 'data-hidden="1"' : '',
    `style="width:${r2(width)}px;height:${r2(height)}px"`,
  ].filter(Boolean).join(' ');
  return `<section ${attrs}>${html}</section>`;
}

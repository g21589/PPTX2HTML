/**
 * SVG chart renderer.
 */
import { css, applyTransforms, rgbToHsl, hslToRgb } from '../pptx/color.js';
import { escapeXml } from '../util/xml.js';
import { r2, r4, PX_PER_PT } from '../util/units.js';
import { formatValue, formatGeneral } from './numfmt.js';
import { fontStack } from '../render/fonts.js';
import { svgPaint } from '../render/paint.js';
import { strokeAttrs } from '../render/shape.js';
import { finalizeLine } from '../pptx/fill.js';

const VARIATIONS = [null, [0.6, 0], [0.8, 0.2], [0.8, 0], [0.6, 0.4], [0.5, 0]];

function vary(c, k) {
  const v = VARIATIONS[k % VARIATIONS.length];
  if (!v) return c;
  const hsl = rgbToHsl(c.r, c.g, c.b);
  const l = Math.min(1, hsl.l * v[0] + v[1]);
  return { ...hslToRgb(hsl.h, hsl.s, l), a: c.a };
}

/** Rough text width estimate (no DOM needed). */
export function textWidth(str, sizePx, bold) {
  let w = 0;
  for (const ch of String(str)) {
    const c = ch.codePointAt(0);
    if (c >= 0x2e80) w += 1.0;
    else if (ch === ' ') w += 0.28;
    else if (/[il.,:;'!|]/.test(ch)) w += 0.26;
    else if (/[mwMW]/.test(ch)) w += 0.82;
    else if (/[A-Z]/.test(ch)) w += 0.62;
    else if (/[0-9]/.test(ch)) w += 0.55;
    else w += 0.5;
  }
  return w * sizePx * (bold ? 1.06 : 1);
}

function niceNum(range, round) {
  if (range <= 0 || !Number.isFinite(range)) return 1;
  const exp = Math.floor(Math.log10(range));
  const f = range / Math.pow(10, exp);
  let nf;
  if (round) nf = f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10;
  else nf = f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * Math.pow(10, exp);
}

/** Automatic value axis scaling (Excel-like). */
export function autoScale(dmin, dmax, opts = {}) {
  let min = dmin, max = dmax;
  if (!Number.isFinite(min) || !Number.isFinite(max)) { min = 0; max = 1; }
  if (min === max) {
    if (min === 0) max = 1;
    else if (min > 0) min = 0;
    else max = 0;
  }
  if (opts.min === undefined && min > 0 && (max - min) / max > 1 / 6) min = 0;
  if (opts.max === undefined && max < 0 && (max - min) / -min > 1 / 6) max = 0;
  if (opts.min !== undefined) min = opts.min;
  if (opts.max !== undefined) max = opts.max;
  const target = opts.ticks || 6;
  let unit = opts.majorUnit || niceNum((max - min) / target, true);
  if (!opts.majorUnit) {
    while ((max - min) / unit > 10) unit = niceNum(unit * 2, true);
  }
  let lo = opts.min !== undefined ? min : Math.floor(min / unit + 1e-9) * unit;
  let hi = opts.max !== undefined ? max : Math.ceil((max + (max - min) * 0.05 * (max > 0 ? 1 : 0)) / unit - 1e-9) * unit;
  if (opts.max === undefined && hi < max) hi += unit;
  if (opts.min === undefined && dmin < 0 && lo > dmin) lo -= unit;
  if (hi === lo) hi = lo + unit;
  return { min: lo, max: hi, unit };
}

function logScale(dmin, dmax, base, opts) {
  const lb = Math.log(base);
  const lo = opts.min !== undefined ? opts.min : Math.pow(base, Math.floor(Math.log(Math.max(dmin, 1e-12)) / lb));
  const hi = opts.max !== undefined ? opts.max : Math.pow(base, Math.ceil(Math.log(Math.max(dmax, lo * base)) / lb));
  return { min: lo, max: hi, unit: base, log: base };
}

const MARKERS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dot', 'dash'];

function markerPath(sym, x, y, s) {
  const h = s / 2;
  switch (sym) {
    case 'circle': return `M${r2(x - h)} ${r2(y)}a${r2(h)} ${r2(h)} 0 1 0 ${r2(s)} 0a${r2(h)} ${r2(h)} 0 1 0 ${r2(-s)} 0Z`;
    case 'square': return `M${r2(x - h)} ${r2(y - h)}h${r2(s)}v${r2(s)}h${r2(-s)}Z`;
    case 'diamond': return `M${r2(x)} ${r2(y - h)}L${r2(x + h)} ${r2(y)}L${r2(x)} ${r2(y + h)}L${r2(x - h)} ${r2(y)}Z`;
    case 'triangle': return `M${r2(x)} ${r2(y - h)}L${r2(x + h)} ${r2(y + h)}L${r2(x - h)} ${r2(y + h)}Z`;
    case 'x': return `M${r2(x - h)} ${r2(y - h)}L${r2(x + h)} ${r2(y + h)}M${r2(x + h)} ${r2(y - h)}L${r2(x - h)} ${r2(y + h)}`;
    case 'star': return `M${r2(x - h)} ${r2(y - h)}L${r2(x + h)} ${r2(y + h)}M${r2(x + h)} ${r2(y - h)}L${r2(x - h)} ${r2(y + h)}M${r2(x)} ${r2(y - h)}V${r2(y + h)}`;
    case 'plus': return `M${r2(x - h)} ${r2(y)}H${r2(x + h)}M${r2(x)} ${r2(y - h)}V${r2(y + h)}`;
    case 'dot': return `M${r2(x - h / 2)} ${r2(y - h / 2)}h${r2(h)}v${r2(h)}h${r2(-h)}Z`;
    case 'dash': return `M${r2(x - h)} ${r2(y - s / 10)}h${r2(s)}v${r2(s / 5)}h${r2(-s)}Z`;
    default: return '';
  }
}

/** Catmull-Rom spline through points as cubic Béziers. */
function smoothPath(pts) {
  if (pts.length < 3) return pts.map((p, i) => (i ? 'L' : 'M') + r2(p[0]) + ' ' + r2(p[1])).join('');
  let d = `M${r2(pts[0][0])} ${r2(pts[0][1])}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] || pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] || p2;
    const t = 1 / 6;
    d += `C${r2(p1[0] + (p2[0] - p0[0]) * t)} ${r2(p1[1] + (p2[1] - p0[1]) * t)} ${r2(p2[0] - (p3[0] - p1[0]) * t)} ${r2(p2[1] - (p3[1] - p1[1]) * t)} ${r2(p2[0])} ${r2(p2[1])}`;
  }
  return d;
}

class ChartRenderer {
  constructor(chart, w, h, rc) {
    this.chart = chart;
    this.w = w;
    this.h = h;
    this.rc = rc;
    this.defs = [];
    this.out = [];
    this.modern = !!(chart.text && chart.text.bag && Object.keys(chart.text.bag).length);
  }

  // ---- styles -------------------------------------------------------------
  font(t) {
    const base = this.chart.textResolved || {};
    const st = t || base;
    const size = (st.size || base.size || 10) * PX_PER_PT;
    const color = st.fill && st.fill.type === 'solid' ? st.fill.color : (base.fill && base.fill.type === 'solid' ? base.fill.color : { r: 89, g: 89, b: 89, a: 1 });
    const family = fontStack(st.latin || base.latin || 'Calibri', st.ea || base.ea, st.cs);
    return { size, color, family, bold: !!st.bold, italic: !!st.italic, rot: st.rot };
  }

  text(x, y, str, f, o = {}) {
    if (str === null || str === undefined || str === '') return '';
    const anchor = o.anchor || 'middle';
    const base = o.baseline || 'central';
    const rot = o.rot ? ` transform="rotate(${r2(o.rot)} ${r2(x)} ${r2(y)})"` : '';
    const lines = String(str).split('\n');
    const lh = f.size * 1.2;
    let attrs = `font-size="${r2(f.size)}" font-family="${escapeXml(f.family)}" fill="${css(f.color)}" text-anchor="${anchor}"`;
    if (f.bold) attrs += ' font-weight="bold"';
    if (f.italic) attrs += ' font-style="italic"';
    if (lines.length === 1) return `<text x="${r2(x)}" y="${r2(y)}" dominant-baseline="${base}" ${attrs}${rot}>${escapeXml(str)}</text>`;
    const y0 = y - ((lines.length - 1) * lh) / 2;
    return `<text ${attrs} dominant-baseline="${base}"${rot}>${lines.map((l, i) => `<tspan x="${r2(x)}" y="${r2(y0 + i * lh)}">${escapeXml(l)}</tspan>`).join('')}</text>`;
  }

  seriesColor(ser, i) {
    const pal = this.chart.palette;
    const k = Math.floor(i / pal.length);
    return vary(pal[i % pal.length], k);
  }

  /** Fill for a series (or point) with fallbacks to automatic colours. */
  seriesFill(ser, si, pi) {
    const dp = pi !== undefined && ser.dPts ? ser.dPts[pi] : null;
    if (dp && dp.fill !== undefined) return dp.fill;
    const plot = ser.plot || {};
    if (plot.varyColors && pi !== undefined && this.isPieLike(plot)) {
      return { type: 'solid', color: this.seriesColor(ser, pi) };
    }
    if (plot.varyColors && pi !== undefined && this.chart.series.length === 1) {
      return { type: 'solid', color: this.seriesColor(ser, pi) };
    }
    if (ser.fill !== undefined) return ser.fill;
    return { type: 'solid', color: this.seriesColor(ser, si) };
  }

  seriesLine(ser, si, pi, isLineChart) {
    const dp = pi !== undefined && ser.dPts ? ser.dPts[pi] : null;
    const raw = (dp && dp.line) || ser.line;
    if (raw) {
      const merged = { ...raw };
      if (!merged.fill && isLineChart) merged.fill = { type: 'solid', color: this.seriesColor(ser, si) };
      if (merged.w === undefined && isLineChart) merged.w = 3;
      return finalizeLine(merged);
    }
    if (isLineChart) return finalizeLine({ fill: { type: 'solid', color: this.seriesColor(ser, si) }, w: 3, cap: 'rnd', join: 'round' });
    if (this.isPieLike(ser.plot || {}) && !this.modern) return null;
    if (this.isPieLike(ser.plot || {})) return finalizeLine({ fill: { type: 'solid', color: this.chart.bg1 || { r: 255, g: 255, b: 255, a: 1 } }, w: 1 });
    return null;
  }

  isPieLike(plot) {
    return /^(pie|pie3D|doughnut|ofPie)Chart$/.test(plot.type);
  }

  paintAttr(fill, w, h, o = {}) {
    const p = svgPaint(fill, w, h, this.rc, this.defs, o);
    return `fill="${p.paint}"${p.opacity !== undefined ? ` fill-opacity="${r4(p.opacity)}"` : ''}`;
  }

  strokeAttr(line) {
    if (!line) return 'stroke="none"';
    const p = svgPaint(line.fill, this.w, this.h, this.rc, this.defs, {});
    return strokeAttrs(line, p).trim();
  }

  rect(x, y, w, h, fill, line, extra = '') {
    if (w < 0) { x += w; w = -w; }
    if (h < 0) { y += h; h = -h; }
    const fa = fill ? this.paintAttr(fill, w, h) : 'fill="none"';
    // Paint coordinates for gradients are relative to the rect: use a group translate.
    return `<g transform="translate(${r2(x)} ${r2(y)})"><rect width="${r2(w)}" height="${r2(h)}" ${fa} ${this.strokeAttr(line)}${extra}/></g>`;
  }

  // ---- data helpers ------------------------------------------------------
  categories() {
    let best = null;
    for (const s of this.chart.series) {
      if (s.cat && s.cat.values && (!best || s.cat.values.length > best.values.length)) best = s.cat;
    }
    let n = 0;
    for (const s of this.chart.series) if (s.val) n = Math.max(n, s.val.values.length);
    if (!best) return { labels: Array.from({ length: n }, (_, i) => String(i + 1)), n };
    const labels = best.values.map((v) => {
      if (v === null || v === undefined) return '';
      if (best.numeric) return formatValue(v, best.formatCode || 'General', this.chart.date1904);
      return String(v);
    });
    while (labels.length < n) labels.push(String(labels.length + 1));
    return { labels, n: Math.max(n, labels.length), levels: best.levels };
  }

  axisFormat(ax, series) {
    if (ax && ax.numFmt && !(ax.sourceLinked && series && series.val && series.val.formatCode)) return ax.numFmt;
    if (series && series.val && series.val.formatCode) return series.val.formatCode;
    return ax && ax.numFmt ? ax.numFmt : 'General';
  }

  // ---- layout ------------------------------------------------------------
  titleText() {
    const c = this.chart;
    if (c.title) {
      if (c.title.rich) return c.title.rich.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n');
      if (c.title.text) return c.title.text;
      if (c.series.length === 1 && c.series[0].name) return c.series[0].name;
      return 'Chart Title';
    }
    if (!c.autoTitleDeleted && c.series.length === 1 && c.series[0].name && c.plots.length === 1 && this.isPieLike(c.plots[0])) return c.series[0].name;
    return null;
  }

  render() {
    const c = this.chart;
    const W = this.w, H = this.h;
    // Chart area.
    const caFill = c.chartArea && c.chartArea.fill !== undefined ? c.chartArea.fill : (this.modern ? { type: 'none' } : { type: 'none' });
    const caLine = c.chartArea && c.chartArea.line ? finalizeLine(c.chartArea.line) : null;
    if (caFill && caFill.type !== 'none' || caLine) this.out.push(this.rect(0, 0, W, H, caFill, caLine, c.roundedCorners ? ' rx="8"' : ''));
    if (c.kind === 'chartex') return this.renderChartEx();
    const pad = Math.max(6, Math.min(W, H) * 0.03);
    let top = pad, bottom = H - pad, left = pad, right = W - pad;
    // Title.
    const title = this.titleText();
    if (title) {
      const tst = c.title && c.title.rich && c.title.rich.paras[0] && c.title.rich.paras[0].runs[0] ? c.title.rich.paras[0].runs[0].style : (c.title ? c.title.resolved : null);
      const f = this.font(tst || { ...(c.textResolved || {}), size: ((c.textResolved && c.textResolved.size) || 10) * 1.4 });
      const lines = title.split('\n').length;
      const th = lines * f.size * 1.2;
      let tx = W / 2, ty = top + th / 2;
      if (c.title && c.title.layout && c.title.layout.x !== undefined) {
        tx = c.title.layout.x * W + textWidth(title, f.size, f.bold) / 2;
        ty = c.title.layout.y * H + th / 2;
      }
      this.titleMarkup = this.text(tx, ty, title, f);
      if (!(c.title && c.title.overlay)) top += th + pad * 0.6;
    }
    // Legend.
    let legendBox = null;
    if (c.legend) {
      legendBox = this.layoutLegend(left, top, right, bottom);
      if (legendBox && !c.legend.overlay && !c.legend.layout) {
        switch (legendBox.pos) {
          case 'r': right = legendBox.x - pad; break;
          case 'l': left = legendBox.x + legendBox.w + pad; break;
          case 't': top = legendBox.y + legendBox.h + pad * 0.5; break;
          case 'b': bottom = legendBox.y - pad * 0.5; break;
          case 'tr': right = legendBox.x - pad; break;
          default: break;
        }
      }
    }
    // Plot area.
    let plot = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top), inner: false };
    const pl = c.plotArea && c.plotArea.layout;
    if (pl && pl.x !== undefined && pl.w !== undefined) {
      plot = { x: pl.x * W, y: pl.y * H, w: pl.w * W, h: pl.h * H, inner: pl.target === 'inner' };
    }
    this.plotOuter = plot;
    const parts = [];
    const plots = c.plots;
    const pieLike = plots.length && plots.every((p) => this.isPieLike(p));
    if (pieLike) parts.push(this.renderPies(plot));
    else if (plots.some((p) => p.type === 'radarChart')) parts.push(this.renderRadar(plot));
    else parts.push(this.renderCartesian(plot));
    this.out.push(...parts);
    if (this.titleMarkup) this.out.push(this.titleMarkup);
    if (legendBox) this.out.push(this.renderLegend(legendBox));
    return `<svg class="pp-chart-svg" width="${r2(W)}" height="${r2(H)}" viewBox="0 0 ${r2(W)} ${r2(H)}" overflow="visible">${this.defs.length ? '<defs>' + this.defs.join('') + '</defs>' : ''}${this.out.join('')}</svg>`;
  }

  legendEntries() {
    const c = this.chart;
    const entries = [];
    const del = c.legend ? c.legend.deleted : new Set();
    const single = c.plots.length === 1 ? c.plots[0] : null;
    if (single && (this.isPieLike(single) || (single.varyColors && c.series.length === 1))) {
      const ser = single.series[0];
      if (!ser) return entries;
      const cats = this.categories();
      cats.labels.forEach((l, i) => {
        if (!del.has(i)) entries.push({ label: l, fill: this.seriesFill(ser, 0, i), line: this.seriesLine(ser, 0, i, false), kind: 'box' });
      });
      return entries;
    }
    c.series.forEach((s, i) => {
      if (del.has(s.idx)) return;
      const t = s.plot.type;
      const isLine = /line|scatter|radar/i.test(t) && !(t === 'radarChart' && s.plot.radarStyle === 'filled');
      const marker = isLine ? this.markerFor(s, i) : null;
      entries.push({
        label: s.name || `Series${i + 1}`,
        fill: isLine ? null : this.seriesFill(s, i),
        line: isLine ? this.seriesLine(s, i, undefined, true) : this.seriesLine(s, i, undefined, false),
        marker,
        kind: isLine ? 'line' : 'box',
      });
    });
    if (c.plots.some((p) => p.type === 'barChart' && p.barDir === 'bar' && /stacked/i.test(p.grouping))) entries.reverse();
    return entries;
  }

  layoutLegend(left, top, right, bottom) {
    const c = this.chart;
    const entries = this.legendEntries();
    if (!entries.length) return null;
    const f = this.font(c.legend.resolved);
    const key = f.size * 0.9;
    const gap = f.size * 0.5;
    const pos = c.legend.pos || 'r';
    const W = this.w, H = this.h;
    const itemW = entries.map((e) => key * (e.kind === 'line' ? 2.2 : 1) + gap * 0.6 + textWidth(e.label, f.size, f.bold));
    const lh = f.size * 1.35;
    let box;
    if (pos === 't' || pos === 'b') {
      const total = itemW.reduce((a, b) => a + b + gap * 1.5, 0);
      const rowsNeeded = total > right - left ? Math.ceil(total / (right - left)) : 1;
      const wbox = Math.min(right - left, total);
      const hbox = rowsNeeded * lh + gap * 0.6;
      box = { pos, x: (W - wbox) / 2, y: pos === 't' ? top : bottom - hbox, w: wbox, h: hbox, horizontal: true };
    } else {
      const wbox = Math.min((right - left) * 0.5, Math.max(...itemW) + gap * 1.2);
      const hbox = entries.length * lh + gap * 0.6;
      const y = pos === 'tr' ? top : top + (bottom - top - hbox) / 2;
      box = { pos, x: pos === 'l' ? left : right - wbox, y: Math.max(top, y), w: wbox, h: hbox, horizontal: false };
    }
    const lay = c.legend.layout;
    if (lay && lay.x !== undefined) {
      box.x = lay.x * W; box.y = lay.y * H;
      if (lay.w !== undefined) box.w = lay.w * W;
      if (lay.h !== undefined) box.h = lay.h * H;
    }
    box.entries = entries;
    box.font = f;
    box.key = key;
    box.lh = lh;
    box.gap = gap;
    box.itemW = itemW;
    return box;
  }

  renderLegend(b) {
    let s = '';
    const c = this.chart;
    if (c.legend.fill || c.legend.line) s += this.rect(b.x, b.y, b.w, b.h, c.legend.fill || null, c.legend.line ? finalizeLine(c.legend.line) : null);
    const drawKey = (e, x, y) => {
      if (e.kind === 'line') {
        let k = '';
        if (e.line) k += `<path d="M${r2(x)} ${r2(y)}h${r2(b.key * 2.2)}" ${this.strokeAttr({ ...e.line, w: Math.min(e.line.w, 3), head: null, tail: null })} fill="none"/>`;
        if (e.marker && e.marker.symbol !== 'none') k += this.markerSvg(e.marker, x + b.key * 1.1, y, Math.min(e.marker.size, 7));
        return k;
      }
      return this.rect(x, y - b.key / 2, b.key, b.key, e.fill, e.line ? { ...e.line, w: Math.min(e.line.w, 1) } : null);
    };
    if (b.horizontal) {
      let x = b.x + b.gap * 0.5;
      let y = b.y + b.lh / 2 + b.gap * 0.3;
      b.entries.forEach((e, i) => {
        if (x + b.itemW[i] > b.x + b.w + 1 && x > b.x + b.gap) { x = b.x + b.gap * 0.5; y += b.lh; }
        s += drawKey(e, x, y);
        const kw = b.key * (e.kind === 'line' ? 2.2 : 1) + b.gap * 0.6;
        s += this.text(x + kw, y, e.label, b.font, { anchor: 'start' });
        x += b.itemW[i] + b.gap * 1.5;
      });
    } else {
      b.entries.forEach((e, i) => {
        const y = b.y + b.gap * 0.3 + b.lh * (i + 0.5);
        const x = b.x + b.gap * 0.6;
        s += drawKey(e, x, y);
        const kw = b.key * (e.kind === 'line' ? 2.2 : 1) + b.gap * 0.6;
        s += this.text(x + kw, y, e.label, b.font, { anchor: 'start' });
      });
    }
    return `<g class="pp-chart-legend">${s}</g>`;
  }

  markerFor(ser, si) {
    const plot = ser.plot;
    const m = ser.marker;
    let symbol = m ? m.symbol : 'auto';
    if (plot.type === 'lineChart' && plot.showMarker === false && (!m || symbol === 'auto')) symbol = 'none';
    if (plot.type === 'scatterChart' && /^(line|smooth)$/.test(plot.scatterStyle) && (!m || symbol === 'auto')) symbol = 'none';
    if (plot.type === 'radarChart' && plot.radarStyle !== 'marker' && (!m || symbol === 'auto')) symbol = 'none';
    if (symbol === 'auto') symbol = MARKERS[si % MARKERS.length];
    const color = this.seriesColor(ser, si);
    const fill = m && m.fill !== undefined ? m.fill : { type: 'solid', color };
    const line = m && m.line ? finalizeLine({ fill: { type: 'solid', color }, w: 0.75, ...m.line }) : finalizeLine({ fill: { type: 'solid', color }, w: 0.75 });
    return { symbol, size: m ? m.size : 5, fill, line };
  }

  markerSvg(m, x, y, sizePt) {
    if (!m || m.symbol === 'none') return '';
    const s = (sizePt || m.size || 5) * PX_PER_PT;
    const d = markerPath(m.symbol, x, y, s);
    if (!d) return '';
    const open = m.symbol === 'x' || m.symbol === 'star' || m.symbol === 'plus';
    const fill = open ? 'fill="none"' : this.paintAttr(m.fill, s, s);
    const line = open ? { ...(m.line || {}), fill: m.line ? m.line.fill : m.fill } : m.line;
    return `<path d="${d}" ${fill} ${this.strokeAttr(line ? { ...line, dash: 'solid', head: null, tail: null } : null)}/>`;
  }

  // ---- cartesian -----------------------------------------------------------
  renderCartesian(outer) {
    const c = this.chart;
    const cats = this.categories();
    const plots = c.plots.filter((p) => !this.isPieLike(p));
    if (!plots.length) return '';
    // Axes groups: each plot references [catAx, valAx] (scatter/bubble: [valAx(x), valAx(y)]).
    const axes = c.axes;
    const groups = new Map();
    for (const p of plots) {
      const ids = p.axIds.filter((id) => axes[id]);
      const key = ids.join('|') || 'default';
      if (!groups.has(key)) groups.set(key, { ids, plots: [] });
      groups.get(key).plots.push(p);
    }
    const groupList = [...groups.values()];
    const primary = groupList[0];
    const horizontalBars = plots.some((p) => p.type.startsWith('bar') && p.barDir === 'bar');

    // Compute scales for each group.
    const info = groupList.map((g) => this.scaleForGroup(g, cats, horizontalBars));

    // Space for axis labels.
    let { x, y, w, h } = outer;
    if (!outer.inner) {
      const margins = { l: 0, r: 0, t: 0, b: 0 };
      info.forEach((gi, gIdx) => {
        for (const ax of [gi.catAxis, gi.valAxis]) {
          if (!ax || ax.def.delete || ax.def.tickLblPos === 'none') continue;
          const f = this.font(ax.def.textResolved);
          const side = ax.side;
          if (side === 'l' || side === 'r') {
            const labels = ax.labels.map((l) => l.text);
            const lw = Math.max(0, ...labels.map((t) => textWidth(t, f.size, f.bold))) + f.size * 0.6;
            margins[side] = Math.max(margins[side], Math.min(lw, outer.w * 0.4));
          } else {
            margins[side] = Math.max(margins[side], f.size * 1.5);
            if (ax.kind === 'cat') {
              const slot = outer.w / Math.max(1, ax.labels.length);
              const maxw = Math.max(0, ...ax.labels.map((l) => textWidth(l.text, f.size, f.bold)));
              if (f.rot === undefined && maxw > slot * 0.95 && ax.labels.length > 1) {
                ax.autoRot = maxw > slot * 1.9 ? -45 : 0;
                ax.wrap = maxw <= slot * 1.9;
                const extra = ax.autoRot ? Math.min(maxw * 0.72, outer.h * 0.35) : f.size * 1.25;
                margins[side] = Math.max(margins[side], f.size * 0.6 + extra);
              } else if (f.rot) {
                margins[side] = Math.max(margins[side], f.size * 0.6 + Math.min(maxw * Math.abs(Math.sin((f.rot * Math.PI) / 180)) + f.size, outer.h * 0.4));
              }
            }
          }
          if (ax.def.title) {
            const tf = this.font(ax.def.title.resolved);
            margins[side] += tf.size * 1.6;
          }
        }
        void gIdx;
      });
      x += margins.l; w -= margins.l + margins.r; y += margins.t; h -= margins.t + margins.b;
    }
    const area = { x, y, w: Math.max(4, w), h: Math.max(4, h) };
    this.area = area;
    let s = '';
    // Plot area background.
    const paFill = c.plotArea && c.plotArea.fill;
    const paLine = c.plotArea && c.plotArea.line ? finalizeLine(c.plotArea.line) : null;
    if ((paFill && paFill.type !== 'none') || paLine) s += this.rect(area.x, area.y, area.w, area.h, paFill || null, paLine);
    // Gridlines (primary group only, value + category).
    const gi0 = info[0];
    s += this.gridlines(gi0, area);
    // Plots.
    info.forEach((gi) => {
      for (const p of gi.group.plots) s += this.renderPlot(p, gi, area, cats);
    });
    // Axes.
    info.forEach((gi) => {
      s += this.axisSvg(gi.catAxis, gi, area, cats);
      s += this.axisSvg(gi.valAxis, gi, area, cats);
    });
    void primary;
    return s;
  }

  scaleForGroup(group, cats, horizontalBars) {
    const c = this.chart;
    const axes = c.axes;
    const [id0, id1] = group.ids;
    const a0 = axes[id0];
    const a1 = axes[id1];
    const isXY = group.plots.every((p) => p.type === 'scatterChart' || p.type === 'bubbleChart');
    // Determine which axis is the category axis.
    let catDef = a0 && a0.kind !== 'val' ? a0 : (a1 && a1.kind !== 'val' ? a1 : null);
    let valDef = a0 && a0.kind === 'val' ? a0 : a1;
    if (isXY) { catDef = a0; valDef = a1; }
    if (!valDef) valDef = { kind: 'val', pos: horizontalBars ? 'b' : 'l', orientation: 'minMax', delete: false, tickLblPos: 'nextTo', majorTickMark: 'out', text: {}, textResolved: c.textResolved };
    if (!catDef) catDef = { kind: 'cat', pos: horizontalBars ? 'l' : 'b', orientation: 'minMax', delete: false, tickLblPos: 'nextTo', majorTickMark: 'out', text: {}, textResolved: c.textResolved };
    const series = group.plots.flatMap((p) => p.series);
    const anyBar = group.plots.some((p) => p.type.startsWith('bar'));
    const vertical = !(anyBar && group.plots.find((p) => p.type.startsWith('bar')).barDir === 'bar');
    // Value range.
    let vmin = Infinity, vmax = -Infinity;
    let percent = false;
    for (const p of group.plots) {
      if (/stacked/i.test(p.grouping)) {
        const n = cats.n;
        const pos = new Array(n).fill(0), neg = new Array(n).fill(0);
        for (const ser of p.series) {
          (ser.val ? ser.val.values : []).forEach((v, i) => {
            if (v === null || v === undefined) return;
            if (v >= 0) pos[i] += v; else neg[i] += v;
          });
        }
        if (p.grouping === 'percentStacked') {
          percent = true;
          vmin = Math.min(vmin, neg.some((v) => v < 0) ? -1 : 0);
          vmax = Math.max(vmax, pos.some((v) => v > 0) ? 1 : 0);
        } else {
          vmin = Math.min(vmin, ...neg, ...pos);
          vmax = Math.max(vmax, ...pos, ...neg);
        }
      } else {
        for (const ser of p.series) {
          const vals = (p.type === 'scatterChart' || p.type === 'bubbleChart') ? (ser.y || ser.val) : ser.val;
          for (const v of vals ? vals.values : []) if (v !== null && v !== undefined && Number.isFinite(v)) { vmin = Math.min(vmin, v); vmax = Math.max(vmax, v); }
          if (p.type === 'stockChart') { /* ranges covered */ }
        }
      }
    }
    if (vmin === Infinity) { vmin = 0; vmax = 1; }
    const valScale = valDef.logBase ? logScale(vmin, vmax, valDef.logBase, valDef) : autoScale(vmin, vmax, { min: valDef.min, max: valDef.max, majorUnit: valDef.majorUnit });
    if (percent && valDef.max === undefined) { valScale.max = Math.min(valScale.max, 1); }
    // X scale for XY charts.
    let xScale = null;
    if (isXY) {
      let xmin = Infinity, xmax = -Infinity;
      for (const ser of series) {
        const xs = ser.x && ser.x.values.some((v) => v !== null) ? ser.x.values : (ser.val || ser.y || { values: [] }).values.map((_, i) => i + 1);
        for (const v of xs) if (v !== null && Number.isFinite(v)) { xmin = Math.min(xmin, v); xmax = Math.max(xmax, v); }
      }
      if (xmin === Infinity) { xmin = 0; xmax = 1; }
      xScale = catDef.logBase ? logScale(xmin, xmax, catDef.logBase, catDef) : autoScale(xmin, xmax, { min: catDef.min, max: catDef.max, majorUnit: catDef.majorUnit });
    }
    const firstSer = series[0];
    const valFmt = percent ? (valDef.numFmt && !valDef.sourceLinked ? valDef.numFmt : '0%') : this.axisFormat(valDef, firstSer);
    // Category axis between/midCat.
    const area = group.plots.some((p) => p.type.startsWith('area'));
    let between = valDef.crossBetween ? valDef.crossBetween === 'between' : !area;
    if (group.plots.some((p) => p.type.startsWith('bar') || p.type === 'stockChart')) between = true;
    const valAxis = {
      kind: 'val', def: valDef, scale: valScale, fmt: valFmt,
      side: valDef.pos || (vertical ? 'l' : 'b'),
      labels: [],
    };
    const ticks = [];
    if (valScale.log) {
      for (let v = valScale.min; v <= valScale.max * 1.0001; v *= valScale.log) ticks.push(v);
    } else {
      const n = Math.round((valScale.max - valScale.min) / valScale.unit);
      for (let i = 0; i <= n && i < 500; i++) ticks.push(valScale.min + i * valScale.unit);
    }
    valAxis.ticks = ticks.map((v) => Math.round(v * 1e10) / 1e10);
    valAxis.labels = valAxis.ticks.map((v) => ({ v, text: formatValue(v, valFmt, c.date1904) }));
    let catAxis;
    if (isXY) {
      const xt = [];
      const n = Math.round((xScale.max - xScale.min) / xScale.unit);
      for (let i = 0; i <= n && i < 500; i++) xt.push(Math.round((xScale.min + i * xScale.unit) * 1e10) / 1e10);
      const xfmt = this.axisFormat(catDef, firstSer && firstSer.x ? { val: firstSer.x } : null);
      catAxis = { kind: 'xval', def: catDef, scale: xScale, ticks: xt, labels: xt.map((v) => ({ v, text: formatValue(v, xfmt, c.date1904) })), side: catDef.pos || 'b' };
    } else {
      const skip = catDef.tickLblSkip || 1;
      catAxis = {
        kind: 'cat', def: catDef, n: cats.n, between,
        labels: cats.labels.map((t, i) => ({ i, text: i % skip === 0 ? t : '' })),
        side: catDef.pos || (vertical ? 'b' : 'l'),
      };
    }
    return { group, catAxis, valAxis, vertical, isXY, percent };
  }

  /** Map a value to a pixel coordinate along the value axis. */
  valPos(gi, v, area) {
    const sc = gi.valAxis.scale;
    let t;
    if (sc.log) t = (Math.log(Math.max(v, 1e-300)) - Math.log(sc.min)) / (Math.log(sc.max) - Math.log(sc.min));
    else t = (v - sc.min) / (sc.max - sc.min || 1);
    if (gi.valAxis.def.orientation === 'maxMin') t = 1 - t;
    if (gi.vertical) return area.y + area.h - t * area.h;
    return area.x + t * area.w;
  }

  xPos(gi, v, area) {
    const sc = gi.catAxis.scale;
    let t;
    if (sc.log) t = (Math.log(Math.max(v, 1e-300)) - Math.log(sc.min)) / (Math.log(sc.max) - Math.log(sc.min));
    else t = (v - sc.min) / (sc.max - sc.min || 1);
    if (gi.catAxis.def.orientation === 'maxMin') t = 1 - t;
    return area.x + t * area.w;
  }

  /** Category band: returns {start, size} along the category axis direction. */
  catBand(gi, i, area) {
    const ca = gi.catAxis;
    const n = Math.max(1, ca.n);
    const len = gi.vertical ? area.w : area.h;
    const origin = gi.vertical ? area.x : area.y;
    let idx = i;
    let reversed = ca.def.orientation === 'maxMin';
    if (!gi.vertical) reversed = !reversed; // bar charts list the first category at the bottom
    if (reversed) idx = n - 1 - i;
    if (ca.between) {
      const size = len / n;
      return { start: origin + idx * size, size, center: origin + (idx + 0.5) * size };
    }
    const size = n > 1 ? len / (n - 1) : len;
    const center = n > 1 ? origin + idx * size : origin + len / 2;
    return { start: center - size / 2, size, center };
  }

  crossValue(gi) {
    const vd = gi.valAxis.def;
    const sc = gi.valAxis.scale;
    const cd = gi.catAxis.def;
    if (cd.crossesAt !== undefined) return cd.crossesAt;
    if (cd.crosses === 'max') return sc.max;
    if (cd.crosses === 'min') return sc.min;
    void vd;
    if (sc.log) return sc.min;
    return Math.min(Math.max(0, sc.min), sc.max);
  }

  gridlines(gi, area) {
    let s = '';
    const defaultGrid = finalizeLine({ fill: { type: 'solid', color: this.modern ? { r: 217, g: 217, b: 217, a: 1 } : { r: 134, g: 134, b: 134, a: 1 } }, w: 1 });
    const va = gi.valAxis;
    if (va.def.majorGridlines && !va.def.deleteGrid) {
      const line = va.def.majorGridlines.fill ? finalizeLine(va.def.majorGridlines) : defaultGrid;
      if (line) for (const v of va.ticks) {
        const p = this.valPos(gi, v, area);
        s += gi.vertical ? `<path d="M${r2(area.x)} ${r2(p)}H${r2(area.x + area.w)}" ${this.strokeAttr(line)}/>` : `<path d="M${r2(p)} ${r2(area.y)}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>`;
      }
    }
    if (va.def.minorGridlines) {
      const line = va.def.minorGridlines.fill ? finalizeLine(va.def.minorGridlines) : defaultGrid;
      const sc = va.scale;
      const mu = va.def.minorUnit || sc.unit / 5;
      if (line && !sc.log) for (let v = sc.min; v <= sc.max + 1e-9; v += mu) {
        const p = this.valPos(gi, v, area);
        s += gi.vertical ? `<path d="M${r2(area.x)} ${r2(p)}H${r2(area.x + area.w)}" ${this.strokeAttr(line)}/>` : `<path d="M${r2(p)} ${r2(area.y)}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>`;
      }
    }
    const ca = gi.catAxis;
    if (ca.def.majorGridlines) {
      const line = ca.def.majorGridlines.fill ? finalizeLine(ca.def.majorGridlines) : defaultGrid;
      if (line) {
        if (ca.kind === 'xval') {
          for (const v of ca.ticks) {
            const p = this.xPos(gi, v, area);
            s += `<path d="M${r2(p)} ${r2(area.y)}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>`;
          }
        } else {
          for (let i = 0; i <= ca.n; i++) {
            const b = this.catBand(gi, Math.min(i, ca.n - 1), area);
            const p = ca.between ? (i === ca.n ? b.start + b.size : b.start) : b.center;
            if (!ca.between && i === ca.n) break;
            s += gi.vertical ? `<path d="M${r2(p)} ${r2(area.y)}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>` : `<path d="M${r2(area.x)} ${r2(p)}H${r2(area.x + area.w)}" ${this.strokeAttr(line)}/>`;
          }
        }
      }
    }
    return `<g class="pp-chart-grid">${s}</g>`;
  }

  axisSvg(ax, gi, area) {
    if (!ax || !ax.def || ax.def.delete) return '';
    const def = ax.def;
    const defaultAxisLine = finalizeLine({ fill: { type: 'solid', color: this.modern ? { r: 217, g: 217, b: 217, a: 1 } : { r: 134, g: 134, b: 134, a: 1 } }, w: 1 });
    const line = def.line ? finalizeLine({ w: 0.75, ...def.line }) : defaultAxisLine;
    const f = this.font(def.textResolved);
    let s = '';
    const isValue = ax.kind === 'val';
    // Axis orientation: value axis is vertical when gi.vertical.
    const axisVertical = isValue ? gi.vertical : (ax.kind === 'xval' ? false : !gi.vertical);
    // Position of the axis line.
    let pos;
    if (isValue) {
      // Value axis crosses the category axis at the first category (or max).
      if (gi.isXY) {
        const cv = def.crossesAt !== undefined ? def.crossesAt : 0;
        const sc = gi.catAxis.scale;
        const xv = gi.catAxis.def.crosses === 'max' || def.crosses === 'max' ? sc.max : Math.min(Math.max(cv, sc.min), sc.max);
        pos = this.xPos(gi, xv, area);
      } else {
        const crossMax = gi.catAxis.def.crosses === 'max' || def.crosses === 'max';
        let rev = gi.catAxis.def.orientation === 'maxMin';
        if (!gi.vertical) rev = !rev;
        const atEnd = crossMax !== rev;
        if (gi.vertical) pos = atEnd ? area.x + area.w : area.x;
        else pos = atEnd ? area.y : area.y + area.h;
      }
    } else {
      const cv = this.crossValue(gi);
      pos = this.valPos(gi, cv, area);
    }
    const side = ax.side;
    if (line) {
      s += axisVertical ? `<path d="M${r2(pos)} ${r2(area.y)}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>`
        : `<path d="M${r2(area.x)} ${r2(pos)}H${r2(area.x + area.w)}" ${this.strokeAttr(line)}/>`;
    }
    // Tick marks and labels.
    const tickLen = 4;
    const lblPos = def.tickLblPos || 'nextTo';
    // Labels sit at the plot edge for low/high, else next to the axis line.
    let lp = pos;
    if (lblPos === 'low') lp = axisVertical ? area.x : area.y + area.h;
    if (lblPos === 'high') lp = axisVertical ? area.x + area.w : area.y;
    const leftSide = axisVertical ? (side === 'r' ? false : true) : (side === 't');
    const off = f.size * 0.4 + (def.majorTickMark === 'out' || def.majorTickMark === 'cross' ? tickLen : 0);
    const tick = (p) => {
      if (!line || def.majorTickMark === 'none') return '';
      const a = def.majorTickMark === 'in' ? -tickLen : def.majorTickMark === 'cross' ? tickLen : 0;
      const b = def.majorTickMark === 'in' ? 0 : tickLen;
      const dir = leftSide ? -1 : 1;
      return axisVertical
        ? `<path d="M${r2(pos + dir * -a)} ${r2(p)}H${r2(pos + dir * b)}" ${this.strokeAttr(line)}/>`
        : `<path d="M${r2(p)} ${r2(pos - dir * a)}V${r2(pos + dir * b)}" ${this.strokeAttr(line)}/>`;
    };
    if (isValue || ax.kind === 'xval') {
      for (const lbl of ax.labels) {
        const p = isValue ? this.valPos(gi, lbl.v, area) : this.xPos(gi, lbl.v, area);
        s += tick(p);
        if (lblPos === 'none') continue;
        if (axisVertical) s += this.text(leftSide ? lp - off : lp + off, p, lbl.text, f, { anchor: leftSide ? 'end' : 'start' });
        else s += this.text(p, leftSide ? lp - off - f.size * 0.5 : lp + off + f.size * 0.5, lbl.text, f, f.rot ? { rot: f.rot, anchor: f.rot < 0 ? 'end' : 'start' } : {});
      }
    } else {
      for (let i = 0; i < ax.n; i++) {
        const b = this.catBand(gi, i, area);
        if (ax.between) s += tick(b.start) + (i === ax.n - 1 ? tick(b.start + b.size) : '');
        else s += tick(b.center);
        const lbl = ax.labels[i];
        if (!lbl || lblPos === 'none' || !lbl.text) continue;
        if (axisVertical) {
          s += this.text(leftSide ? lp - off : lp + off, b.center, lbl.text, f, { anchor: leftSide ? 'end' : 'start' });
        } else {
          const rot = f.rot !== undefined ? f.rot : ax.autoRot;
          let txt = lbl.text;
          if (ax.wrap && !rot) txt = wrapText(txt, b.size * 0.98, f.size, f.bold);
          const yy = leftSide ? lp - off - f.size * 0.5 : lp + off + (rot ? 0 : f.size * 0.5 * txt.split('\n').length);
          s += rot ? this.text(b.center, yy, txt, f, { rot, anchor: rot < 0 ? 'end' : 'start', baseline: 'central' }) : this.text(b.center, yy, txt, f);
        }
      }
    }
    // Axis title.
    if (def.title) {
      const tf = this.font(def.title.resolved);
      const t = def.title.rich ? def.title.rich.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n') : (def.title.text || 'Axis Title');
      const o = this.plotOuter;
      if (axisVertical) {
        const xx = leftSide ? o.x + tf.size * 0.7 : o.x + o.w - tf.size * 0.7;
        const rot = def.title.rich && def.title.rich.rot !== undefined ? def.title.rich.rot : -90;
        s += this.text(xx, area.y + area.h / 2, t, tf, { rot });
      } else {
        const yy = leftSide ? o.y + tf.size * 0.7 : o.y + o.h - tf.size * 0.7;
        s += this.text(area.x + area.w / 2, yy, t, tf);
      }
    }
    return `<g class="pp-chart-axis">${s}</g>`;
  }

  renderPlot(p, gi, area, cats) {
    switch (p.type) {
      case 'barChart': case 'bar3DChart': return this.renderBars(p, gi, area, cats);
      case 'lineChart': case 'line3DChart': return this.renderLines(p, gi, area, cats);
      case 'areaChart': case 'area3DChart': return this.renderAreas(p, gi, area, cats);
      case 'scatterChart': return this.renderScatter(p, gi, area);
      case 'bubbleChart': return this.renderBubbles(p, gi, area);
      case 'stockChart': return this.renderStock(p, gi, area, cats);
      case 'surfaceChart': case 'surface3DChart': return this.renderLines({ ...p, type: 'lineChart' }, gi, area, cats);
      default: return '';
    }
  }

  seriesIndex(ser) {
    return this.chart.series.indexOf(ser);
  }

  /** Stacked accumulation helper. */
  stackValues(p, n) {
    const pos = new Array(n).fill(0), neg = new Array(n).fill(0);
    const totals = new Array(n).fill(0);
    if (p.grouping === 'percentStacked') {
      for (const s of p.series) (s.val ? s.val.values : []).forEach((v, i) => { if (v) totals[i] += Math.abs(v); });
    }
    return { pos, neg, totals };
  }

  renderBars(p, gi, area, cats) {
    const n = cats.n;
    const series = p.series;
    const stacked = /stacked/i.test(p.grouping);
    const k = stacked ? 1 : series.length;
    const gap = (p.gapWidth || 0) / 100;
    const ov = stacked ? 1 : Math.max(-1, Math.min(1, (p.overlap || 0) / 100));
    const acc = this.stackValues(p, n);
    let s = '';
    const labels = [];
    series.forEach((ser, sIdx) => {
      const si = this.seriesIndex(ser);
      const vals = ser.val ? ser.val.values : [];
      for (let i = 0; i < n; i++) {
        let v = vals[i];
        if (v === null || v === undefined || !Number.isFinite(v)) continue;
        const band = this.catBand(gi, i, area);
        const bw = band.size / (k - (k - 1) * ov + gap);
        const startOff = band.start + (gap * bw) / 2 + (stacked ? 0 : sIdx * bw * (1 - ov));
        let v0, v1;
        if (stacked) {
          if (p.grouping === 'percentStacked') v = acc.totals[i] ? v / acc.totals[i] : 0;
          if (v >= 0) { v0 = acc.pos[i]; acc.pos[i] += v; v1 = acc.pos[i]; } else { v0 = acc.neg[i]; acc.neg[i] += v; v1 = acc.neg[i]; }
        } else {
          v0 = this.crossValue(gi);
          v1 = v;
        }
        const p0 = this.valPos(gi, v0, area);
        const p1 = this.valPos(gi, v1, area);
        let fill = this.seriesFill(ser, si, i);
        const dp = ser.dPts[i];
        const inv = dp && dp.invertIfNegative !== undefined ? dp.invertIfNegative : ser.invertIfNegative;
        if (inv && v < 0 && fill.type === 'solid' && !(dp && dp.fill)) fill = { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } };
        const line = this.seriesLine(ser, si, i, false);
        let rx, ry, rw, rh;
        if (gi.vertical) { rx = startOff; rw = bw; ry = Math.min(p0, p1); rh = Math.abs(p1 - p0); } else { ry = startOff; rh = bw; rx = Math.min(p0, p1); rw = Math.abs(p1 - p0); }
        s += this.barShape(rx, ry, rw, rh, fill, line, ser, p);
        labels.push({ ser, si, i, v: vals[i], rect: { x: rx, y: ry, w: rw, h: rh }, neg: v < 0, plot: p });
      }
    });
    return `<g class="pp-chart-bars">${s}${this.barLabels(labels, gi, cats)}</g>`;
  }

  barShape(x, y, w, h, fill, line, ser, p) {
    const shape = ser.shape || null;
    void p;
    if (shape === 'cylinder' || shape === 'cone' || shape === 'pyramid') {
      // 3D shapes are drawn flat; pyramids / cones as triangles.
      if (shape !== 'cylinder') {
        const d = `M${r2(x)} ${r2(y + h)}L${r2(x + w / 2)} ${r2(y)}L${r2(x + w)} ${r2(y + h)}Z`;
        return `<path d="${d}" ${this.paintAttr(fill, w, h)} ${this.strokeAttr(line)}/>`;
      }
    }
    return this.rect(x, y, w, h, fill, line);
  }

  labelText(dl, ser, i, v, cats, percent) {
    const pt = dl.points && dl.points[i];
    if (pt && pt.delete) return null;
    if (pt && pt.rich) return pt.rich.paras.map((pp) => pp.runs.map((r) => r.text).join('')).join('\n');
    const o = pt || dl;
    const parts = [];
    if (o.showSerName && ser.name) parts.push(ser.name);
    if (o.showCatName) parts.push(cats.labels[i] || '');
    if (o.showVal) {
      const fmt = dl.numFmt && !dl.sourceLinked ? dl.numFmt : (ser.val && ser.val.formatCode) || dl.numFmt || 'General';
      parts.push(formatValue(v, fmt, this.chart.date1904));
    }
    if (o.showPercent && percent !== undefined) parts.push(formatValue(percent, dl.numFmt && !dl.sourceLinked ? dl.numFmt : '0%'));
    if (!parts.length) return null;
    return parts.join(dl.separator !== null && dl.separator !== undefined ? dl.separator.replace(/\n/g, '\n') : (o.showPercent && o.showCatName && parts.length === 2 ? '\n' : ', '));
  }

  effectiveDLbls(ser) {
    const d = ser.dLbls || ser.plot.dLbls;
    if (!d || d.delete) return null;
    return d;
  }

  barLabels(items, gi, cats) {
    let s = '';
    for (const it of items) {
      const dl = this.effectiveDLbls(it.ser);
      if (!dl) continue;
      const t = this.labelText(dl, it.ser, it.i, it.v, cats);
      if (t === null) continue;
      const f = this.font(dl.resolved);
      const pt = dl.points && dl.points[it.i];
      const pos = (pt && pt.pos) || dl.pos || (/stacked/i.test(it.plot.grouping) ? 'ctr' : 'outEnd');
      const r = it.rect;
      let x = r.x + r.w / 2, y = r.y + r.h / 2;
      const pad = f.size * 0.6;
      if (gi.vertical) {
        const topEnd = it.neg ? r.y + r.h : r.y;
        const dir = it.neg ? 1 : -1;
        if (pos === 'outEnd') y = topEnd + dir * pad;
        else if (pos === 'inEnd') y = topEnd - dir * pad;
        else if (pos === 'inBase') y = (it.neg ? r.y : r.y + r.h) + dir * pad;
      } else {
        const endX = it.neg ? r.x : r.x + r.w;
        const dir = it.neg ? -1 : 1;
        if (pos === 'outEnd') { x = endX + dir * (pad + textWidth(t, f.size) / 2); }
        else if (pos === 'inEnd') x = endX - dir * (pad + textWidth(t, f.size) / 2);
        else if (pos === 'inBase') x = (it.neg ? r.x + r.w : r.x) + dir * (pad + textWidth(t, f.size) / 2);
      }
      s += this.text(x, y, t, f);
    }
    return s;
  }

  linePoints(ser, gi, area, cats, acc, p) {
    const vals = ser.val ? ser.val.values : [];
    const pts = [];
    for (let i = 0; i < cats.n; i++) {
      let v = vals[i];
      if (v === null || v === undefined || !Number.isFinite(v)) {
        if (this.chart.dispBlanksAs === 'zero') v = 0;
        else { pts.push(null); continue; }
      }
      let pv = v;
      if (acc) {
        if (p.grouping === 'percentStacked') pv = acc.totals[i] ? v / acc.totals[i] : 0;
        acc.pos[i] += pv;
        pv = acc.pos[i];
      }
      const b = this.catBand(gi, i, area);
      const vp = this.valPos(gi, pv, area);
      pts.push(gi.vertical ? [b.center, vp, v, i] : [vp, b.center, v, i]);
    }
    return pts;
  }

  renderLines(p, gi, area, cats) {
    const stacked = /stacked/i.test(p.grouping);
    const acc = stacked ? this.stackValues(p, cats.n) : null;
    let s = '';
    let lbl = '';
    const all = [];
    for (const ser of p.series) {
      const si = this.seriesIndex(ser);
      const pts = this.linePoints(ser, gi, area, cats, acc, p);
      all.push(pts);
      const line = this.seriesLine(ser, si, undefined, true);
      // Split on gaps unless spanning.
      const segments = [];
      let cur = [];
      for (const pt of pts) {
        if (!pt) { if (this.chart.dispBlanksAs !== 'span' && cur.length) { segments.push(cur); cur = []; } continue; }
        cur.push(pt);
      }
      if (cur.length) segments.push(cur);
      if (line) for (const seg of segments) {
        const d = ser.smooth ? smoothPath(seg) : seg.map((q, i) => (i ? 'L' : 'M') + r2(q[0]) + ' ' + r2(q[1])).join('');
        s += `<path d="${d}" fill="none" ${this.strokeAttr(line)}/>`;
      }
      const m = this.markerFor(ser, si);
      for (const pt of pts) {
        if (!pt) continue;
        const dp = ser.dPts[pt[3]];
        const mm = dp && dp.marker ? { ...m, ...this.markerFor({ ...ser, marker: dp.marker }, si) } : m;
        s += this.markerSvg(mm, pt[0], pt[1]);
      }
      const dl = this.effectiveDLbls(ser);
      if (dl) {
        const f = this.font(dl.resolved);
        for (const pt of pts) {
          if (!pt) continue;
          const t = this.labelText(dl, ser, pt[3], pt[2], cats);
          if (t === null) continue;
          const pos = dl.pos || 'r';
          let x = pt[0], y = pt[1];
          if (pos === 't') y -= f.size; else if (pos === 'b') y += f.size; else if (pos === 'l') x -= textWidth(t, f.size) / 2 + 6; else if (pos === 'ctr') { /* centred */ } else x += textWidth(t, f.size) / 2 + 6;
          lbl += this.text(x, y, t, f);
        }
      }
      s += this.trendlines(ser, si, gi, area, pts);
    }
    // Drop lines / high-low lines / up-down bars for line charts.
    if (p.hiLowLines) s += this.hiLow(all, p.hiLowLines, gi);
    if (p.upDownBars && all.length >= 2) s += this.upDown(all, p.upDownBars, gi, area, cats);
    if (p.dropLines) {
      const line = finalizeLine({ fill: { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } }, w: 0.75, ...p.dropLines });
      for (const pts of all) for (const pt of pts) if (pt) s += `<path d="M${r2(pt[0])} ${r2(pt[1])}V${r2(area.y + area.h)}" ${this.strokeAttr(line)}/>`;
    }
    return `<g class="pp-chart-lines">${s}${lbl}</g>`;
  }

  hiLow(all, lineDef, gi) {
    const line = finalizeLine({ fill: { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } }, w: 0.75, ...lineDef });
    let s = '';
    const n = Math.max(...all.map((a) => a.length));
    for (let i = 0; i < n; i++) {
      const pts = all.map((a) => a[i]).filter(Boolean);
      if (pts.length < 2) continue;
      if (gi.vertical) {
        const ys = pts.map((q) => q[1]);
        s += `<path d="M${r2(pts[0][0])} ${r2(Math.min(...ys))}V${r2(Math.max(...ys))}" ${this.strokeAttr(line)}/>`;
      } else {
        const xs = pts.map((q) => q[0]);
        s += `<path d="M${r2(Math.min(...xs))} ${r2(pts[0][1])}H${r2(Math.max(...xs))}" ${this.strokeAttr(line)}/>`;
      }
    }
    return s;
  }

  upDown(all, ud, gi, area, cats) {
    let s = '';
    const first = all[0], last = all[all.length - 1];
    for (let i = 0; i < cats.n; i++) {
      const a = first[i], b = last[i];
      if (!a || !b) continue;
      const band = this.catBand(gi, i, area);
      const bw = band.size / (1 + (ud.gapWidth || 150) / 100);
      const up = b[2] >= a[2];
      const sp = up ? ud.up : ud.down;
      const fill = sp && sp.fill !== undefined ? sp.fill : { type: 'solid', color: up ? { r: 255, g: 255, b: 255, a: 1 } : { r: 0, g: 0, b: 0, a: 1 } };
      const line = finalizeLine({ fill: { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } }, w: 0.75, ...((sp && sp.line) || {}) });
      const y0 = Math.min(a[1], b[1]), y1 = Math.max(a[1], b[1]);
      s += this.rect(band.center - bw / 2, y0, bw, Math.max(1, y1 - y0), fill, line);
    }
    return s;
  }

  renderAreas(p, gi, area, cats) {
    const stacked = /stacked/i.test(p.grouping);
    const acc = stacked ? this.stackValues(p, cats.n) : null;
    let s = '';
    let prev = null;
    const base = this.valPos(gi, this.crossValue(gi), area);
    const items = [];
    for (const ser of p.series) {
      const si = this.seriesIndex(ser);
      const pts = this.linePoints(ser, gi, area, cats, acc, p).map((q, i) => q || (gi.vertical ? [this.catBand(gi, i, area).center, base, 0, i] : [base, this.catBand(gi, i, area).center, 0, i]));
      if (!pts.length) continue;
      let d = 'M' + pts.map((q) => r2(q[0]) + ' ' + r2(q[1])).join('L');
      if (stacked && prev) d += 'L' + prev.slice().reverse().map((q) => r2(q[0]) + ' ' + r2(q[1])).join('L');
      else if (gi.vertical) d += `L${r2(pts[pts.length - 1][0])} ${r2(base)}L${r2(pts[0][0])} ${r2(base)}`;
      else d += `L${r2(base)} ${r2(pts[pts.length - 1][1])}L${r2(base)} ${r2(pts[0][1])}`;
      d += 'Z';
      items.push({ d, fill: this.seriesFill(ser, si), line: this.seriesLine(ser, si, undefined, false), ser, pts });
      prev = pts;
    }
    // Overlapping (standard) areas: the first series is drawn in front.
    const order = stacked ? items : items.slice().reverse();
    for (const it of order) {
      s += `<g transform="translate(${r2(area.x)} ${r2(area.y)})"><path transform="translate(${r2(-area.x)} ${r2(-area.y)})" d="${it.d}" ${this.paintAttr(it.fill, area.w, area.h)} ${this.strokeAttr(it.line)}/></g>`;
    }
    for (const it of items) {
      const dl = this.effectiveDLbls(it.ser);
      if (!dl) continue;
      const f = this.font(dl.resolved);
      for (const q of it.pts) {
        const t = this.labelText(dl, it.ser, q[3], q[2], cats);
        if (t !== null) s += this.text(q[0], q[1] + f.size, t, f);
      }
    }
    return `<g class="pp-chart-areas">${s}</g>`;
  }

  xyPoints(ser, gi, area) {
    const ys = (ser.y || ser.val || { values: [] }).values;
    const xs = ser.x && ser.x.values.some((v) => v !== null) ? ser.x.values : ys.map((_, i) => i + 1);
    const pts = [];
    for (let i = 0; i < ys.length; i++) {
      const yv = ys[i], xv = xs[i];
      if (yv === null || yv === undefined || xv === null || xv === undefined || !Number.isFinite(yv) || !Number.isFinite(xv)) { pts.push(null); continue; }
      pts.push([this.xPos(gi, xv, area), this.valPos(gi, yv, area), yv, i, xv]);
    }
    return pts;
  }

  renderScatter(p, gi, area) {
    let s = '';
    const style = p.scatterStyle;
    for (const ser of p.series) {
      const si = this.seriesIndex(ser);
      const pts = this.xyPoints(ser, gi, area);
      const hasLine = ser.line ? ser.line.fill && ser.line.fill.type !== 'none' : /line|smooth/i.test(style) && style !== 'marker';
      const line = hasLine ? this.seriesLine(ser, si, undefined, true) : null;
      if (line && !(ser.line && ser.line.fill && ser.line.fill.type === 'none')) {
        const seg = pts.filter(Boolean);
        const smooth = ser.smooth || /smooth/i.test(style);
        const d = smooth ? smoothPath(seg) : seg.map((q, i) => (i ? 'L' : 'M') + r2(q[0]) + ' ' + r2(q[1])).join('');
        if (seg.length) s += `<path d="${d}" fill="none" ${this.strokeAttr(line)}/>`;
      }
      const m = this.markerFor(ser, si);
      for (const pt of pts) if (pt) s += this.markerSvg(m, pt[0], pt[1]);
      const dl = this.effectiveDLbls(ser);
      if (dl) {
        const f = this.font(dl.resolved);
        for (const pt of pts) {
          if (!pt) continue;
          const t = this.labelText(dl, ser, pt[3], pt[2], { labels: [] });
          if (t !== null) s += this.text(pt[0] + textWidth(t, f.size) / 2 + 6, pt[1], t, f);
        }
      }
      s += this.trendlines(ser, si, gi, area, pts, true);
    }
    return `<g class="pp-chart-scatter">${s}</g>`;
  }

  renderBubbles(p, gi, area) {
    let s = '';
    let maxSize = 0;
    for (const ser of p.series) for (const v of (ser.size ? ser.size.values : [])) if (v !== null && Math.abs(v) > maxSize) maxSize = Math.abs(v);
    const maxR = (Math.min(area.w, area.h) / 2) * 0.25 * ((p.bubbleScale || 100) / 100);
    for (const ser of p.series) {
      const si = this.seriesIndex(ser);
      const pts = this.xyPoints(ser, gi, area);
      const sizes = ser.size ? ser.size.values : [];
      pts.forEach((pt, i) => {
        if (!pt) return;
        const sz = sizes[i];
        if (sz === null || sz === undefined || (sz < 0 && !p.showNegBubbles)) return;
        const r = maxSize ? Math.sqrt(Math.abs(sz) / maxSize) * maxR : maxR / 2;
        const fill = this.seriesFill(ser, si, i);
        const line = this.seriesLine(ser, si, i, false);
        s += `<g transform="translate(${r2(pt[0] - r)} ${r2(pt[1] - r)})"><circle cx="${r2(r)}" cy="${r2(r)}" r="${r2(r)}" ${this.paintAttr(fill, 2 * r, 2 * r)} ${this.strokeAttr(line)}/></g>`;
      });
    }
    return `<g class="pp-chart-bubbles">${s}</g>`;
  }

  renderStock(p, gi, area, cats) {
    const all = p.series.map((ser) => this.linePoints(ser, gi, area, cats, null, p));
    let s = '';
    s += this.hiLow(all, p.hiLowLines || {}, gi);
    if (p.upDownBars && all.length >= 2) s += this.upDown(all, p.upDownBars, gi, area, cats);
    p.series.forEach((ser, k) => {
      const si = this.seriesIndex(ser);
      const m = this.markerFor(ser, si);
      const line = ser.line && ser.line.fill && ser.line.fill.type !== 'none' ? this.seriesLine(ser, si, undefined, true) : null;
      if (line) {
        const seg = all[k].filter(Boolean);
        s += `<path d="${seg.map((q, i) => (i ? 'L' : 'M') + r2(q[0]) + ' ' + r2(q[1])).join('')}" fill="none" ${this.strokeAttr(line)}/>`;
      }
      if (ser.marker) for (const pt of all[k]) if (pt) s += this.markerSvg(m, pt[0], pt[1]);
    });
    return `<g class="pp-chart-stock">${s}</g>`;
  }

  trendlines(ser, si, gi, area, pts, xy) {
    if (!ser.trendlines || !ser.trendlines.length) return '';
    let s = '';
    const data = pts.filter(Boolean).map((q) => [xy ? q[4] : q[3] + 1, q[2]]);
    if (data.length < 2) return '';
    for (const t of ser.trendlines) {
      const line = finalizeLine({ fill: { type: 'solid', color: this.seriesColor(ser, si) }, w: 2.25 * PX_PER_PT / 1.333, dash: 'sysDot', ...(t.line || {}) });
      const fn = fitTrend(t, data);
      if (!fn) continue;
      const x0 = Math.min(...data.map((d) => d[0])) - (t.backward || 0);
      const x1 = Math.max(...data.map((d) => d[0])) + (t.forward || 0);
      const path = [];
      const steps = 40;
      for (let k = 0; k <= steps; k++) {
        const xv = x0 + ((x1 - x0) * k) / steps;
        const yv = fn(xv);
        if (!Number.isFinite(yv)) continue;
        const px = xy ? this.xPos(gi, xv, area) : this.catBand(gi, 0, area).center + (xv - 1) * (this.catBand(gi, 1, area).center - this.catBand(gi, 0, area).center || 0);
        const py = this.valPos(gi, yv, area);
        path.push((path.length ? 'L' : 'M') + r2(gi.vertical ? px : py) + ' ' + r2(gi.vertical ? py : px));
      }
      if (path.length > 1) s += `<path d="${path.join('')}" fill="none" ${this.strokeAttr(line)}/>`;
    }
    return s;
  }

  // ---- pie ---------------------------------------------------------------
  renderPies(outer) {
    const c = this.chart;
    const p = c.plots[0];
    const doughnut = p.type === 'doughnutChart';
    const series = doughnut ? p.series : p.series.slice(0, 1);
    if (!series.length) return '';
    const cats = this.categories();
    // Leave room for outside labels.
    const dl0 = this.effectiveDLbls(series[0]);
    const outside = dl0 && (dl0.pos === 'outEnd' || dl0.pos === 'bestFit');
    const maxExpl = Math.max(0, ...series.flatMap((s) => [s.explosion || 0, ...Object.values(s.dPts).map((d) => d.explosion || 0)])) / 100;
    let R = Math.min(outer.w, outer.h) / 2 * (outside ? 0.78 : 0.92) / (1 + maxExpl);
    const pl = c.plotArea && c.plotArea.layout;
    if (pl && pl.target === 'inner' && pl.w !== undefined) R = Math.min(outer.w, outer.h) / 2 / (1 + maxExpl);
    const cx = outer.x + outer.w / 2;
    let cy = outer.y + outer.h / 2;
    let s = '';
    let labels = '';
    const hole = doughnut ? Math.max(0.1, Math.min(0.9, (p.holeSize || 50) / 100)) : 0;
    // Pseudo 3D pie: squash vertically and draw the visible side walls.
    const is3D = p.type === 'pie3DChart';
    const tilt = is3D ? Math.max(10, Math.min(90, (c.view3D && c.view3D.rotX) || 30)) : 90;
    const sy = Math.sin((tilt * Math.PI) / 180);
    let depth = 0;
    if (is3D) {
      R = Math.min(outer.w / 2, outer.h / (2 * sy + 0.3)) * (outside ? 0.8 : 0.92) / (1 + maxExpl);
      depth = R * 0.22 * ((c.view3D && c.view3D.depth) || 100) / 100 * Math.cos((tilt * Math.PI) / 180) + R * 0.05;
      cy -= depth / 2;
      s += this.pie3DSides(series[0], R, cx, cy, sy, depth, p);
    }
    const ringW = (R * (1 - hole)) / series.length;
    series.forEach((ser, k) => {
      const si = this.seriesIndex(ser);
      const vals = (ser.val ? ser.val.values : []).map((v) => (v && v > 0 ? v : 0));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      let ang = ((p.firstSliceAng || 0) - 90) * (Math.PI / 180);
      const r1 = doughnut ? R - k * ringW : R;
      const r0 = doughnut ? r1 - ringW : 0;
      vals.forEach((v, i) => {
        const sweep = (v / total) * Math.PI * 2;
        const a0 = ang, a1 = ang + sweep;
        ang = a1;
        if (v <= 0) return;
        const dp = ser.dPts[i];
        const ex = ((dp && dp.explosion !== undefined ? dp.explosion : ser.explosion) || 0) / 100;
        const mid = (a0 + a1) / 2;
        const ox = Math.cos(mid) * ex * R, oy = Math.sin(mid) * ex * R * sy;
        const fill = this.seriesFill(ser, si, i);
        const line = this.seriesLine(ser, si, i, false);
        let d;
        const large = sweep > Math.PI ? 1 : 0;
        const P = (r, a) => [cx + ox + Math.cos(a) * r, cy + oy + Math.sin(a) * r * sy];
        if (is3D) {
          const [x0, y0] = P(r1, a0), [x1, y1] = P(r1, a1);
          d = sweep >= Math.PI * 2 - 1e-6
            ? `M${r2(cx + ox + r1)} ${r2(cy + oy)}A${r2(r1)} ${r2(r1 * sy)} 0 1 1 ${r2(cx + ox - r1)} ${r2(cy + oy)}A${r2(r1)} ${r2(r1 * sy)} 0 1 1 ${r2(cx + ox + r1)} ${r2(cy + oy)}Z`
            : `M${r2(cx + ox)} ${r2(cy + oy)}L${r2(x0)} ${r2(y0)}A${r2(r1)} ${r2(r1 * sy)} 0 ${large} 1 ${r2(x1)} ${r2(y1)}Z`;
        } else if (sweep >= Math.PI * 2 - 1e-6) {
          d = `M${r2(cx + ox + r1)} ${r2(cy + oy)}A${r2(r1)} ${r2(r1)} 0 1 1 ${r2(cx + ox - r1)} ${r2(cy + oy)}A${r2(r1)} ${r2(r1)} 0 1 1 ${r2(cx + ox + r1)} ${r2(cy + oy)}Z`;
          if (r0 > 0) d += `M${r2(cx + ox + r0)} ${r2(cy + oy)}A${r2(r0)} ${r2(r0)} 0 1 0 ${r2(cx + ox - r0)} ${r2(cy + oy)}A${r2(r0)} ${r2(r0)} 0 1 0 ${r2(cx + ox + r0)} ${r2(cy + oy)}Z`;
        } else {
          const [x0, y0] = P(r1, a0), [x1, y1] = P(r1, a1);
          if (r0 > 0) {
            const [x2, y2] = P(r0, a1), [x3, y3] = P(r0, a0);
            d = `M${r2(x0)} ${r2(y0)}A${r2(r1)} ${r2(r1)} 0 ${large} 1 ${r2(x1)} ${r2(y1)}L${r2(x2)} ${r2(y2)}A${r2(r0)} ${r2(r0)} 0 ${large} 0 ${r2(x3)} ${r2(y3)}Z`;
          } else {
            d = `M${r2(cx + ox)} ${r2(cy + oy)}L${r2(x0)} ${r2(y0)}A${r2(r1)} ${r2(r1)} 0 ${large} 1 ${r2(x1)} ${r2(y1)}Z`;
          }
        }
        s += `<g transform="translate(${r2(cx - R)} ${r2(cy - R)})"><path transform="translate(${r2(R - cx)} ${r2(R - cy)})" d="${d}" ${this.paintAttr(fill, 2 * R, 2 * R)} ${this.strokeAttr(line)} fill-rule="evenodd"/></g>`;
        const dl = this.effectiveDLbls(ser);
        if (dl) {
          const t = this.labelText(dl, ser, i, ser.val.values[i], cats, v / total);
          if (t !== null) {
            const f = this.font(dl.resolved);
            const pt = dl.points && dl.points[i];
            const pos = (pt && pt.pos) || dl.pos || (doughnut ? 'ctr' : 'bestFit');
            let lr;
            if (doughnut) lr = (r0 + r1) / 2;
            else if (pos === 'outEnd') lr = r1 + f.size * 1.2 + textWidth(t.split('\n')[0], f.size) * 0.35;
            else if (pos === 'inEnd') lr = r1 * 0.78;
            else if (pos === 'ctr') lr = r1 * 0.55;
            else lr = r1 * 0.68;
            const [lx, ly] = P(lr, mid);
            labels += this.text(lx, ly, t, f);
          }
        }
      });
    });
    return `<g class="pp-chart-pie">${s}${labels}</g>`;
  }

  pie3DSides(ser, R, cx, cy, sy, depth, p) {
    const si = this.seriesIndex(ser);
    const vals = (ser.val ? ser.val.values : []).map((v) => (v && v > 0 ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    let ang = ((p.firstSliceAng || 0) - 90) * (Math.PI / 180);
    const parts = [];
    vals.forEach((v, i) => {
      const sweep = (v / total) * Math.PI * 2;
      const a0 = ang, a1 = ang + sweep;
      ang = a1;
      if (v <= 0) return;
      const dp = ser.dPts[i];
      const ex = ((dp && dp.explosion !== undefined ? dp.explosion : ser.explosion) || 0) / 100;
      const mid = (a0 + a1) / 2;
      const ox = Math.cos(mid) * ex * R, oy = Math.sin(mid) * ex * R * sy;
      const fill = this.seriesFill(ser, si, i);
      const base = fill && fill.type === 'solid' ? fill.color : this.seriesColor(ser, i);
      const dark = { type: 'solid', color: { ...applyTransforms(base, [{ local: 'shade', num: () => 70000 }]) } };
      // Visible outer wall: the part of [a0, a1] in the lower half (0..PI).
      for (let k = -1; k <= 1; k++) {
        const lo = Math.max(a0, k * 2 * Math.PI), hi = Math.min(a1, k * 2 * Math.PI + Math.PI);
        if (hi <= lo) continue;
        const P = (a, dy) => [cx + ox + Math.cos(a) * R, cy + oy + Math.sin(a) * R * sy + dy];
        const [x0, y0] = P(lo, 0), [x1, y1] = P(hi, 0), [x2, y2] = P(hi, depth), [x3, y3] = P(lo, depth);
        const large = hi - lo > Math.PI ? 1 : 0;
        const d = `M${r2(x0)} ${r2(y0)}A${r2(R)} ${r2(R * sy)} 0 ${large} 1 ${r2(x1)} ${r2(y1)}L${r2(x2)} ${r2(y2)}A${r2(R)} ${r2(R * sy)} 0 ${large} 0 ${r2(x3)} ${r2(y3)}Z`;
        parts.push({ z: Math.sin((lo + hi) / 2), d: `<path d="${d}" ${this.paintAttr(dark, 2 * R, 2 * R)} stroke="none"/>` });
      }
    });
    parts.sort((a, b) => a.z - b.z);
    return parts.map((q) => q.d).join('');
  }

  // ---- radar -------------------------------------------------------------
  renderRadar(outer) {
    const c = this.chart;
    const p = c.plots.find((x) => x.type === 'radarChart');
    const cats = this.categories();
    const n = Math.max(1, cats.n);
    const ax = Object.values(c.axes).find((a) => a.kind === 'val') || { orientation: 'minMax', textResolved: c.textResolved };
    const catAx = Object.values(c.axes).find((a) => a.kind === 'cat') || { textResolved: c.textResolved };
    let vmin = Infinity, vmax = -Infinity;
    for (const ser of p.series) for (const v of ser.val ? ser.val.values : []) if (v !== null && Number.isFinite(v)) { vmin = Math.min(vmin, v); vmax = Math.max(vmax, v); }
    const sc = autoScale(vmin === Infinity ? 0 : vmin, vmax === -Infinity ? 1 : vmax, { min: ax.min, max: ax.max, majorUnit: ax.majorUnit });
    const f = this.font(catAx.textResolved);
    const R = Math.min(outer.w, outer.h) / 2 - f.size * 2;
    const cx = outer.x + outer.w / 2, cy = outer.y + outer.h / 2;
    const angle = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
    const rad = (v) => ((v - sc.min) / (sc.max - sc.min || 1)) * R;
    let s = '';
    const grid = finalizeLine({ fill: { type: 'solid', color: { r: 217, g: 217, b: 217, a: 1 } }, w: 1, ...(ax.majorGridlines || {}) });
    for (let v = sc.min; v <= sc.max + 1e-9; v += sc.unit) {
      const r = rad(v);
      const d = Array.from({ length: n }, (_, i) => (i ? 'L' : 'M') + r2(cx + Math.cos(angle(i)) * r) + ' ' + r2(cy + Math.sin(angle(i)) * r)).join('') + 'Z';
      s += `<path d="${d}" fill="none" ${this.strokeAttr(grid)}/>`;
      if (!ax.delete && ax.tickLblPos !== 'none') s += this.text(cx - 4, cy - r, formatValue(v, ax.numFmt || 'General'), this.font(ax.textResolved), { anchor: 'end' });
    }
    for (let i = 0; i < n; i++) {
      s += `<path d="M${r2(cx)} ${r2(cy)}L${r2(cx + Math.cos(angle(i)) * R)} ${r2(cy + Math.sin(angle(i)) * R)}" ${this.strokeAttr(grid)}/>`;
      const lx = cx + Math.cos(angle(i)) * (R + f.size), ly = cy + Math.sin(angle(i)) * (R + f.size);
      s += this.text(lx, ly, cats.labels[i], f, { anchor: Math.abs(Math.cos(angle(i))) < 0.2 ? 'middle' : Math.cos(angle(i)) > 0 ? 'start' : 'end' });
    }
    for (const ser of p.series) {
      const si = this.seriesIndex(ser);
      const vals = ser.val ? ser.val.values : [];
      const pts = vals.map((v, i) => [cx + Math.cos(angle(i)) * rad(v || 0), cy + Math.sin(angle(i)) * rad(v || 0)]);
      const d = pts.map((q, i) => (i ? 'L' : 'M') + r2(q[0]) + ' ' + r2(q[1])).join('') + 'Z';
      if (p.radarStyle === 'filled') s += `<path d="${d}" ${this.paintAttr(this.seriesFill(ser, si), this.w, this.h)} ${this.strokeAttr(this.seriesLine(ser, si, undefined, false))}/>`;
      else {
        s += `<path d="${d}" fill="none" ${this.strokeAttr(this.seriesLine(ser, si, undefined, true))}/>`;
        const m = this.markerFor(ser, si);
        for (const q of pts) s += this.markerSvg(m, q[0], q[1]);
      }
    }
    return `<g class="pp-chart-radar">${s}</g>`;
  }

  // ---- chart extensions (waterfall, funnel, treemap, ...) --------------------------
  renderChartEx() {
    const c = this.chart;
    const W = this.w, H = this.h;
    const pad = 8;
    let top = pad;
    if (c.title && c.title.text) {
      const f = this.font({ ...(c.textResolved || {}), size: ((c.textResolved && c.textResolved.size) || 10) * 1.4 });
      this.out.push(this.text(W / 2, top + f.size * 0.7, c.title.text, f));
      top += f.size * 1.6;
    }
    const ser = c.series.find((s) => !s.hidden) || c.series[0];
    if (!ser || !ser.val) return this.finishEx();
    const f = this.font(c.textResolved);
    const vals = ser.val.values.map((v) => (Number.isFinite(v) ? v : 0));
    const labels = ser.cat ? ser.cat.values : vals.map((_, i) => String(i + 1));
    const area = { x: pad + f.size * 3, y: top, w: W - pad * 2 - f.size * 3, h: H - top - pad - f.size * 2 };
    const layout = ser.layoutId;
    const colorAt = (i) => (ser.dPts[i] && ser.dPts[i].fill) || ser.fill || { type: 'solid', color: this.seriesColor(ser, layout === 'waterfall' ? (vals[i] >= 0 ? 0 : 1) : i) };
    let s = '';
    if (layout === 'funnel') {
      const n = vals.length;
      const max = Math.max(...vals.map(Math.abs), 1);
      const bh = area.h / n;
      vals.forEach((v, i) => {
        const w = (Math.abs(v) / max) * area.w;
        s += this.rect(area.x + (area.w - w) / 2, area.y + i * bh + bh * 0.1, w, bh * 0.8, ser.fill || { type: 'solid', color: this.seriesColor(ser, 0) }, null);
        s += this.text(area.x + area.w / 2, area.y + i * bh + bh / 2, formatGeneral(v), { ...f, color: { r: 255, g: 255, b: 255, a: 1 } });
        s += this.text(area.x - 4, area.y + i * bh + bh / 2, labels[i], f, { anchor: 'end' });
      });
    } else if (layout === 'treemap' || layout === 'sunburst') {
      // Squarified-ish slice layout.
      const total = vals.reduce((a, b) => a + Math.abs(b), 0) || 1;
      let x = area.x, y = area.y, w = area.w + f.size * 3, h = area.h + f.size * 2;
      x -= f.size * 3;
      vals.forEach((v, i) => {
        const frac = Math.abs(v) / total;
        const rest = vals.slice(i).reduce((a, b) => a + Math.abs(b), 0) / total || 1;
        const part = frac / rest;
        let rx, ry, rw, rh;
        if (w >= h) { rw = w * part; rh = h; rx = x; ry = y; x += rw; w -= rw; } else { rh = h * part; rw = w; rx = x; ry = y; y += rh; h -= rh; }
        s += this.rect(rx, ry, rw, rh, colorAt(i), finalizeLine({ fill: { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } }, w: 1 }));
        s += this.text(rx + 4, ry + f.size, labels[i], { ...f, color: { r: 255, g: 255, b: 255, a: 1 } }, { anchor: 'start' });
      });
    } else {
      // waterfall, clusteredColumn (histogram/pareto), boxWhisker: columns.
      let data = vals;
      let labs = labels;
      if (layout === 'clusteredColumn' && ser.binning) {
        const bins = histogram(vals, ser.binning);
        data = bins.counts;
        labs = bins.labels;
      }
      let run = 0;
      const tops = [];
      data.forEach((v, i) => {
        if (layout === 'waterfall') {
          if (ser.subtotals && ser.subtotals.includes(i)) { tops.push([0, v]); run = v; } else { tops.push([run, run + v]); run += v; }
        } else tops.push([0, v]);
      });
      const lo = Math.min(0, ...tops.flat()), hi = Math.max(0, ...tops.flat());
      const sc = autoScale(lo, hi, {});
      const gi = { valAxis: { scale: sc, def: { orientation: 'minMax' } }, vertical: true };
      const n = data.length;
      const bw = area.w / n;
      for (let v = sc.min; v <= sc.max + 1e-9; v += sc.unit) {
        const py = this.valPos(gi, v, area);
        s += `<path d="M${r2(area.x)} ${r2(py)}H${r2(area.x + area.w)}" stroke="#d9d9d9" stroke-width="1"/>`;
        s += this.text(area.x - 4, py, formatGeneral(Math.round(v * 1e9) / 1e9), f, { anchor: 'end' });
      }
      tops.forEach(([a, b], i) => {
        const y0 = this.valPos(gi, a, area), y1 = this.valPos(gi, b, area);
        const fill = layout === 'waterfall' && !(ser.dPts[i] && ser.dPts[i].fill) ? { type: 'solid', color: this.seriesColor(ser, ser.subtotals && ser.subtotals.includes(i) ? 2 : (b >= a ? 0 : 1)) } : colorAt(i);
        s += this.rect(area.x + i * bw + bw * 0.15, Math.min(y0, y1), bw * 0.7, Math.abs(y1 - y0), fill, null);
        s += this.text(area.x + i * bw + bw / 2, area.y + area.h + f.size, labs[i], f);
      });
    }
    this.out.push(s);
    return this.finishEx();
  }

  finishEx() {
    return `<svg class="pp-chart-svg" width="${r2(this.w)}" height="${r2(this.h)}" viewBox="0 0 ${r2(this.w)} ${r2(this.h)}" overflow="visible">${this.defs.length ? '<defs>' + this.defs.join('') + '</defs>' : ''}${this.out.join('')}</svg>`;
  }
}

function histogram(vals, binning) {
  const min = Math.min(...vals), max = Math.max(...vals);
  let count = binning.count || 0;
  let size = binning.size || 0;
  if (!count && !size) count = Math.max(1, Math.round(Math.sqrt(vals.length)));
  if (!size) size = (max - min) / count || 1;
  if (!count) count = Math.max(1, Math.ceil((max - min) / size));
  const counts = new Array(count).fill(0);
  for (const v of vals) counts[Math.min(count - 1, Math.floor((v - min) / size))]++;
  const labels = counts.map((_, i) => `[${formatGeneral(+(min + i * size).toFixed(2))}, ${formatGeneral(+(min + (i + 1) * size).toFixed(2))}${i === count - 1 ? ']' : ')'}`);
  return { counts, labels };
}

function wrapText(t, width, size, bold) {
  const words = String(t).split(/(\s+)/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    if (textWidth(cur + w, size, bold) > width && cur.trim()) { lines.push(cur.trim()); cur = w.trimStart(); } else cur += w;
  }
  if (cur.trim()) lines.push(cur.trim());
  return lines.slice(0, 3).join('\n');
}

/** Least squares trend line fitting. Returns f(x) or null. */
export function fitTrend(t, data) {
  const n = data.length;
  if (n < 2) return null;
  const lin = (pts) => {
    let sx = 0, sy = 0, sxx = 0, sxy = 0;
    for (const [x, y] of pts) { sx += x; sy += y; sxx += x * x; sxy += x * y; }
    const den = n * sxx - sx * sx;
    if (!den) return null;
    const b = (n * sxy - sx * sy) / den;
    return { a: (sy - b * sx) / n, b };
  };
  switch (t.type) {
    case 'linear': {
      const r = lin(data);
      return r ? (x) => r.a + r.b * x : null;
    }
    case 'exp': {
      const pts = data.filter(([, y]) => y > 0).map(([x, y]) => [x, Math.log(y)]);
      const r = pts.length >= 2 ? lin(pts) : null;
      return r ? (x) => Math.exp(r.a + r.b * x) : null;
    }
    case 'log': {
      const pts = data.filter(([x]) => x > 0).map(([x, y]) => [Math.log(x), y]);
      const r = pts.length >= 2 ? lin(pts) : null;
      return r ? (x) => r.a + r.b * Math.log(x) : null;
    }
    case 'power': {
      const pts = data.filter(([x, y]) => x > 0 && y > 0).map(([x, y]) => [Math.log(x), Math.log(y)]);
      const r = pts.length >= 2 ? lin(pts) : null;
      return r ? (x) => Math.exp(r.a) * Math.pow(x, r.b) : null;
    }
    case 'poly': {
      const k = Math.max(2, Math.min(6, t.order || 2));
      const coef = polyFit(data, k);
      return coef ? (x) => coef.reduce((acc, c, i) => acc + c * Math.pow(x, i), 0) : null;
    }
    case 'movingAvg': {
      const per = Math.max(2, t.period || 2);
      const sorted = data.slice().sort((a, b) => a[0] - b[0]);
      return (x) => {
        const idx = sorted.findIndex((d) => d[0] >= x);
        if (idx < per - 1) return NaN;
        let s = 0;
        for (let i = idx - per + 1; i <= idx; i++) s += sorted[i][1];
        return s / per;
      };
    }
    default: return null;
  }
}

function polyFit(data, k) {
  const m = k + 1;
  const A = Array.from({ length: m }, () => new Array(m + 1).fill(0));
  for (const [x, y] of data) {
    for (let i = 0; i < m; i++) {
      for (let j = 0; j < m; j++) A[i][j] += Math.pow(x, i + j);
      A[i][m] += y * Math.pow(x, i);
    }
  }
  for (let i = 0; i < m; i++) {
    let p = i;
    for (let r = i + 1; r < m; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r;
    [A[i], A[p]] = [A[p], A[i]];
    if (Math.abs(A[i][i]) < 1e-12) return null;
    for (let r = 0; r < m; r++) {
      if (r === i) continue;
      const f = A[r][i] / A[i][i];
      for (let c2 = i; c2 <= m; c2++) A[r][c2] -= f * A[i][c2];
    }
  }
  return A.map((row, i) => row[m] / A[i][i]);
}

/** Render a chart element to SVG markup. */
export function chartHtml(el, rc) {
  const chart = el.chart;
  if (!chart || chart.kind === 'empty' || chart.kind === 'unsupported') return '';
  try {
    return new ChartRenderer(chart, el.w, el.h, rc).render();
  } catch (e) {
    rc.pres.diagnostics.add('error', 'chart render: ' + e.message);
    if (rc.opts.strict) throw e;
    return '';
  }
}

export { applyTransforms };

/**
 * Chart parsing (DrawingML charts c:chartSpace and chart extensions cx:chartSpace).
 */
import { parseFill, parseLine, finalizeLine, mergeLines, parseEffects } from '../pptx/fill.js';
import { rPrBag, resolveRunProps, resolveBodyPr, bodyPrBag } from '../pptx/text.js';
import { resolveColorNode } from '../pptx/color.js';
import { DEFAULT_CLR_MAP } from '../pptx/color.js';

const PLOT_TYPES = new Set(['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'pieChart', 'pie3DChart', 'doughnutChart',
  'ofPieChart', 'areaChart', 'area3DChart', 'scatterChart', 'bubbleChart', 'radarChart', 'stockChart', 'surfaceChart', 'surface3DChart']);

const val = (n, name, def) => {
  const c = n && n.child('c:' + name);
  return c ? c.attr('val', def) : def;
};
const num = (n, name, def) => {
  const v = val(n, name, undefined);
  if (v === undefined) return def;
  const x = parseFloat(v);
  return Number.isNaN(x) ? def : x;
};
const bool = (n, name, def) => {
  const c = n && n.child('c:' + name);
  if (!c) return def;
  const v = c.attr('val', '1');
  return v === '1' || v === 'true';
};

/** Read a cached reference / literal into an array of values. */
function readData(node, numeric) {
  if (!node) return null;
  let cache = null;
  let formatCode = null;
  for (const c of node.children) {
    if (c.local === 'numRef' || c.local === 'strRef') cache = c.child('c:numCache') || c.child('c:strCache');
    else if (c.local === 'numLit' || c.local === 'strLit') cache = c;
    else if (c.local === 'multiLvlStrRef') {
      const ml = c.child('c:multiLvlStrCache');
      if (ml) {
        const count = num(ml, 'ptCount', 0);
        const levels = ml.all('c:lvl').map((lv) => {
          const arr = new Array(count).fill(null);
          for (const pt of lv.all('c:pt')) {
            const i = pt.int('idx', 0);
            const v = pt.child('c:v');
            if (i < count) arr[i] = v ? v.text : '';
          }
          return arr;
        });
        return { values: levels[0] || [], levels, formatCode: null, count };
      }
    } else if (c.local === 'v') {
      return { values: [numeric ? parseFloat(c.text) : c.text], formatCode: null, count: 1 };
    }
  }
  if (!cache) return null;
  const fc = cache.child('c:formatCode');
  if (fc) formatCode = fc.text;
  const count = num(cache, 'ptCount', 0);
  const pts = cache.all('c:pt');
  const n = Math.max(count, pts.reduce((m, p) => Math.max(m, p.int('idx', 0) + 1), 0));
  const values = new Array(n).fill(null);
  const fmts = [];
  for (const pt of pts) {
    const i = pt.int('idx', 0);
    const v = pt.child('c:v');
    if (!v) continue;
    if (numeric) {
      const x = parseFloat(v.text);
      values[i] = Number.isNaN(x) ? null : x;
    } else values[i] = v.text;
    if (pt.attrs.formatCode) fmts[i] = pt.attrs.formatCode;
  }
  return { values, formatCode, count: n, pointFormats: fmts.length ? fmts : null };
}

function txStyle(txPr, ctx, base) {
  const out = { ...(base || {}) };
  if (!txPr) return out;
  const bp = txPr.child('a:bodyPr');
  if (bp) {
    const body = resolveBodyPr(bodyPrBag(bp));
    // rot="-60000000" means "automatic".
    if (bp.attrs.rot !== undefined && Math.abs(body.rot) <= 360) out.rot = body.rot;
    if (bp.attrs.vert) out.vert = bp.attrs.vert;
  }
  const p = txPr.child('a:p');
  const ppr = p && p.child('a:pPr');
  const def = ppr && ppr.child('a:defRPr');
  if (def) out.bag = { ...(out.bag || {}), ...rPrBag(def) };
  return out;
}

function resolveText(style, ctx, fallbackSize = 10) {
  const bag = (style && style.bag) || {};
  const r = resolveRunProps(bag, ctx, { defaultSize: fallbackSize });
  return { ...r, rot: style ? style.rot : undefined };
}

function richText(tx, ctx, baseStyle) {
  const rich = tx && tx.child('c:rich');
  if (!rich) return null;
  const paras = [];
  const lst = rich.child('a:lstStyle');
  void lst;
  const bp = rich.child('a:bodyPr');
  let rot = bp && bp.attrs.rot !== undefined ? Number(bp.attrs.rot) / 60000 : undefined;
  if (rot !== undefined && Math.abs(rot) > 360) rot = undefined;
  for (const p of rich.all('a:p')) {
    const ppr = p.child('a:pPr');
    const defBag = ppr && ppr.child('a:defRPr') ? rPrBag(ppr.child('a:defRPr')) : {};
    const runs = [];
    for (const r of p.children) {
      if (r.name !== 'a:r' && r.name !== 'a:fld') continue;
      const bag = { ...((baseStyle && baseStyle.bag) || {}), ...defBag, ...rPrBag(r.child('a:rPr')) };
      const t = r.child('a:t');
      runs.push({ text: t ? t.text : '', style: resolveRunProps(bag, ctx, { defaultSize: 10 }) });
    }
    paras.push({ runs, align: ppr ? ppr.attr('algn', 'ctr') : 'ctr', defBag });
  }
  return { paras, rot };
}

function spPr(node, ctx) {
  const sp = node && node.child('c:spPr');
  if (!sp) return { fill: undefined, line: undefined, effects: null };
  const f = parseFill(sp, ctx);
  const ln = sp.child('a:ln');
  const line = ln ? parseLine(ln, ctx) : undefined;
  return { fill: f, line, effects: parseEffects(sp.child('a:effectLst'), ctx) };
}

function layoutOf(n) {
  const ml = n && n.path('c:layout', 'c:manualLayout');
  if (!ml) return null;
  const v = (k) => num(ml, k, undefined);
  return {
    target: val(ml, 'layoutTarget', 'outer'),
    xMode: val(ml, 'xMode', 'factor'), yMode: val(ml, 'yMode', 'factor'),
    wMode: val(ml, 'wMode', 'factor'), hMode: val(ml, 'hMode', 'factor'),
    x: v('x'), y: v('y'), w: v('w'), h: v('h'),
  };
}

function parseMarker(m, ctx) {
  if (!m) return null;
  const sp = spPr(m, ctx);
  return { symbol: val(m, 'symbol', 'auto'), size: num(m, 'size', 5), fill: sp.fill, line: sp.line };
}

function parseDLbls(n, ctx, baseText) {
  const d = n && n.child('c:dLbls');
  if (!d) return null;
  if (bool(d, 'delete', false)) return { delete: true };
  const nf = d.child('c:numFmt');
  const sp = spPr(d, ctx);
  const out = {
    showVal: bool(d, 'showVal', false),
    showPercent: bool(d, 'showPercent', false),
    showCatName: bool(d, 'showCatName', false),
    showSerName: bool(d, 'showSerName', false),
    showLegendKey: bool(d, 'showLegendKey', false),
    showBubbleSize: bool(d, 'showBubbleSize', false),
    separator: d.child('c:separator') ? d.child('c:separator').text : null,
    pos: val(d, 'dLblPos', null),
    numFmt: nf ? nf.attr('formatCode', null) : null,
    sourceLinked: nf ? nf.bool('sourceLinked', false) : true,
    text: txStyle(d.child('c:txPr'), ctx, baseText),
    fill: sp.fill,
    line: sp.line,
    points: {},
  };
  for (const dl of d.all('c:dLbl')) {
    const idx = num(dl, 'idx', 0);
    if (bool(dl, 'delete', false)) { out.points[idx] = { delete: true }; continue; }
    const rich = richText(dl.child('c:tx'), ctx, out.text);
    out.points[idx] = {
      showVal: bool(dl, 'showVal', out.showVal),
      showPercent: bool(dl, 'showPercent', out.showPercent),
      showCatName: bool(dl, 'showCatName', out.showCatName),
      showSerName: bool(dl, 'showSerName', out.showSerName),
      pos: val(dl, 'dLblPos', out.pos),
      rich,
      layout: layoutOf(dl),
    };
  }
  return out;
}

function parseSeries(s, ctx, plotType, baseText) {
  const sp = spPr(s, ctx);
  const txN = s.child('c:tx');
  let name = null;
  if (txN) {
    const d = readData(txN, false);
    if (d && d.values.length) name = d.values.filter((x) => x !== null).join(' ');
  }
  const ser = {
    idx: num(s, 'idx', 0),
    order: num(s, 'order', 0),
    name,
    fill: sp.fill,
    line: sp.line,
    effects: sp.effects,
    cat: readData(s.child('c:cat'), false) || readData(s.child('c:xVal'), false),
    val: readData(s.child('c:val'), true) || readData(s.child('c:yVal'), true),
    x: readData(s.child('c:xVal'), true),
    y: readData(s.child('c:yVal'), true),
    size: readData(s.child('c:bubbleSize'), true),
    marker: parseMarker(s.child('c:marker'), ctx),
    smooth: bool(s, 'smooth', false),
    explosion: num(s, 'explosion', 0),
    invertIfNegative: bool(s, 'invertIfNegative', false),
    dLbls: parseDLbls(s, ctx, baseText),
    dPts: {},
    trendlines: [],
    errBars: [],
    shape: val(s, 'shape', null),
  };
  // Category values may be numeric (numRef): format them later.
  const catNode = s.child('c:cat');
  if (catNode && catNode.child('c:numRef')) {
    const d = readData(catNode, true);
    if (d) ser.cat = { ...d, numeric: true };
  }
  for (const dp of s.all('c:dPt')) {
    const i = num(dp, 'idx', 0);
    const p = spPr(dp, ctx);
    ser.dPts[i] = { fill: p.fill, line: p.line, explosion: num(dp, 'explosion', undefined), marker: parseMarker(dp.child('c:marker'), ctx), invertIfNegative: bool(dp, 'invertIfNegative', undefined) };
  }
  for (const t of s.all('c:trendline')) {
    const p = spPr(t, ctx);
    ser.trendlines.push({
      type: val(t, 'trendlineType', 'linear'), order: num(t, 'order', 2), period: num(t, 'period', 2),
      forward: num(t, 'forward', 0), backward: num(t, 'backward', 0), line: p.line,
      dispRSqr: bool(t, 'dispRSqr', false), dispEq: bool(t, 'dispEq', false), intercept: num(t, 'intercept', undefined),
    });
  }
  for (const e of s.all('c:errBars')) {
    const p = spPr(e, ctx);
    ser.errBars.push({
      dir: val(e, 'errDir', 'y'), type: val(e, 'errBarType', 'both'), valType: val(e, 'errValType', 'fixedVal'),
      val: num(e, 'val', 1), noEndCap: bool(e, 'noEndCap', false), line: p.line,
      plus: readData(e.child('c:plus'), true), minus: readData(e.child('c:minus'), true),
    });
  }
  void plotType;
  return ser;
}

function parseAxis(a, ctx, baseText) {
  const sc = a.child('c:scaling');
  const nf = a.child('c:numFmt');
  const sp = spPr(a, ctx);
  const grid = (name) => {
    const g = a.child('c:' + name);
    if (!g) return null;
    const p = spPr(g, ctx);
    return p.line || {};
  };
  const titleN = a.child('c:title');
  return {
    id: val(a, 'axId', ''),
    kind: a.local.replace('Ax', ''),
    delete: bool(a, 'delete', false),
    pos: val(a, 'axPos', 'b'),
    orientation: sc ? val(sc, 'orientation', 'minMax') : 'minMax',
    min: sc ? num(sc, 'min', undefined) : undefined,
    max: sc ? num(sc, 'max', undefined) : undefined,
    logBase: sc ? num(sc, 'logBase', undefined) : undefined,
    majorUnit: num(a, 'majorUnit', undefined),
    minorUnit: num(a, 'minorUnit', undefined),
    numFmt: nf ? nf.attr('formatCode', 'General') : null,
    sourceLinked: nf ? nf.bool('sourceLinked', false) : true,
    majorGridlines: grid('majorGridlines'),
    minorGridlines: grid('minorGridlines'),
    majorTickMark: val(a, 'majorTickMark', 'out'),
    minorTickMark: val(a, 'minorTickMark', 'none'),
    tickLblPos: val(a, 'tickLblPos', 'nextTo'),
    line: sp.line,
    text: txStyle(a.child('c:txPr'), ctx, baseText),
    title: titleN ? parseTitle(titleN, ctx, baseText) : null,
    crossAx: val(a, 'crossAx', ''),
    crosses: val(a, 'crosses', 'autoZero'),
    crossesAt: num(a, 'crossesAt', undefined),
    crossBetween: val(a, 'crossBetween', null),
    lblOffset: num(a, 'lblOffset', 100),
    tickLblSkip: num(a, 'tickLblSkip', undefined),
    tickMarkSkip: num(a, 'tickMarkSkip', undefined),
    baseTimeUnit: val(a, 'baseTimeUnit', null),
    dispUnits: a.child('c:dispUnits') ? (val(a.child('c:dispUnits'), 'builtInUnit', null)) : null,
  };
}

function parseTitle(t, ctx, baseText) {
  const tx = t.child('c:tx');
  const style = txStyle(t.child('c:txPr'), ctx, baseText);
  let rich = richText(tx, ctx, style);
  let text = null;
  if (!rich && tx) {
    const d = readData(tx, false);
    if (d) text = d.values.join(' ');
  }
  const sp = spPr(t, ctx);
  return { rich, text, style, overlay: bool(t, 'overlay', false), layout: layoutOf(t), fill: sp.fill, line: sp.line };
}

function parseChartEx(root, ctx, doc, part) {
  // Chart extensions (waterfall, funnel, treemap, sunburst, histogram, box & whisker...).
  const chart = { kind: 'chartex', plots: [], series: [], axes: {}, title: null, legend: null, style: {} };
  const data = new Map();
  for (const d of root.find('cx:data')) {
    const id = d.attr('id', '0');
    const entry = { str: [], num: [] };
    for (const sd of d.all('cx:strDim')) {
      const lv = sd.child('cx:lvl');
      entry.str.push({ type: sd.attr('type', 'cat'), levels: sd.all('cx:lvl').map((l) => l.all('cx:pt').map((p) => p.text)), values: lv ? lv.all('cx:pt').map((p) => p.text) : [] });
    }
    for (const nd of d.all('cx:numDim')) {
      const lv = nd.child('cx:lvl');
      const vals = [];
      if (lv) for (const p of lv.all('cx:pt')) vals[p.int('idx', vals.length)] = parseFloat(p.text);
      entry.num.push({ type: nd.attr('type', 'val'), values: vals, formatCode: lv ? lv.attr('formatCode', null) : null });
    }
    data.set(id, entry);
  }
  const baseText = txStyle(root.child('cx:txPr'), ctx, { bag: {} });
  chart.text = baseText;
  const ch = root.child('cx:chart');
  const titleN = ch && ch.child('cx:title');
  if (titleN) {
    const t = titleN.first('cx:txData');
    const rich = titleN.first('cx:rich');
    let text = t ? t.textContent().trim() : null;
    if (rich) text = rich.all('a:p').map((p) => p.textContent()).join('\n');
    // A title without text shows Office's default caption.
    chart.title = { text: text || 'Chart Title', rich: null, style: baseText, overlay: titleN.bool('overlay', false), pos: titleN.attr('pos', 't') };
  }
  const plotArea = ch && ch.child('cx:plotArea');
  const region = plotArea && plotArea.child('cx:plotAreaRegion');
  for (const s of region ? region.all('cx:series') : []) {
    const layoutId = s.attr('layoutId', 'clusteredColumn');
    const dId = s.child('cx:dataId') ? s.child('cx:dataId').attr('val', '0') : '0';
    const d = data.get(dId) || { str: [], num: [] };
    const txN = s.child('cx:tx');
    const name = txN ? txN.textContent().trim() : null;
    const sp = s.child('cx:spPr');
    const lp = s.child('cx:layoutPr');
    const subtotals = lp && lp.child('cx:subtotals') ? lp.child('cx:subtotals').all('cx:idx').map((i) => i.int('val', 0)) : [];
    const dPts = {};
    for (const dp of s.all('cx:dataPt')) {
      const spp = dp.child('cx:spPr');
      dPts[dp.int('idx', 0)] = { fill: spp ? parseFill(spp, ctx) : undefined };
    }
    chart.series.push({
      layoutId, name, hidden: s.bool('hidden', false),
      cat: d.str[0] ? { values: d.str[0].values, levels: d.str[0].levels } : null,
      val: d.num[0] ? { values: d.num[0].values, formatCode: d.num[0].formatCode } : null,
      fill: sp ? parseFill(sp, ctx) : undefined,
      line: sp && sp.child('a:ln') ? parseLine(sp.child('a:ln'), ctx) : undefined,
      dPts,
      subtotals,
      dLbls: s.child('cx:dataLabels') ? { showVal: !!s.child('cx:dataLabels').first('cx:visibility') ? s.child('cx:dataLabels').first('cx:visibility').bool('value', true) : true, pos: s.child('cx:dataLabels').attr('pos', null) } : null,
      binning: lp && lp.child('cx:binning') ? { count: lp.child('cx:binning').first('cx:binCount') ? lp.child('cx:binning').first('cx:binCount').int('val', 0) : 0, size: lp.child('cx:binning').first('cx:binSize') ? lp.child('cx:binning').first('cx:binSize').num('val', 0) : 0 } : null,
      ownerIdx: s.attr('ownerIdx', null),
    });
  }
  chart.plots.push({ type: chart.series.length ? chart.series[0].layoutId : 'clusteredColumn', series: chart.series });
  const leg = ch && ch.child('cx:legend');
  chart.legend = leg ? { pos: leg.attr('pos', 'b'), overlay: leg.bool('overlay', false), style: baseText, deleted: new Set() } : null;
  chart.axes = {};
  for (const ax of plotArea ? plotArea.all('cx:axis') : []) {
    const id = ax.attr('id', '0');
    const hidden = ax.bool('hidden', false);
    chart.axes[id] = { id, kind: ax.child('cx:catScaling') ? 'cat' : 'val', delete: hidden, majorGridlines: ax.child('cx:majorGridlines') ? {} : null, text: baseText, tickLblPos: ax.child('cx:tickLabels') ? 'nextTo' : 'none' };
  }
  const sp = root.child('cx:spPr');
  chart.chartArea = { fill: sp ? parseFill(sp, ctx) : undefined, line: sp && sp.child('a:ln') ? parseLine(sp.child('a:ln'), ctx) : undefined };
  chart.palette = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((a) => resolveColorNode({ name: 'a:schemeClr', attrs: { val: a }, children: [] }, ctx));
  void doc; void part;
  return chart;
}

/**
 * Parse a chart part into a chart model.
 * @param {object} doc document state
 * @param {string} part chart part name
 * @param {import('../pptx/context.js').Ctx} ctx context bound to the chart part
 */
export function parseChartPart(doc, part, ctx, el, isChartEx = false) {
  const root = doc.pkg.xml(part);
  if (!root) return { kind: 'empty' };
  // Charts may override the colour mapping.
  const cmo = root.child('c:clrMapOvr');
  if (cmo) ctx = ctx.with({ clrMap: { ...DEFAULT_CLR_MAP, ...cmo.attrs } });
  // Theme override part.
  const to = doc.pkg.relByType(part, 'themeOverride');
  if (to) {
    const t = doc.pkg.xml(to.partName);
    const cs = t && t.child('a:clrScheme');
    if (cs) {
      const scheme = { ...ctx.scheme };
      for (const c of cs.children) {
        const cn = c.children[0];
        if (cn) scheme[c.local] = resolveColorNode(cn, { scheme: {}, clrMap: {} });
      }
      ctx = ctx.with({ scheme });
    }
  }
  if (isChartEx || root.name.startsWith('cx:')) return parseChartEx(root, ctx, doc, part);

  const chartN = root.child('c:chart');
  if (!chartN) return { kind: 'empty' };
  const baseText = txStyle(root.child('c:txPr'), ctx, { bag: {} });
  const chart = {
    kind: 'chart',
    style: num(root, 'style', 2),
    date1904: bool(root, 'date1904', false),
    roundedCorners: bool(root, 'roundedCorners', false),
    text: baseText,
    title: null,
    autoTitleDeleted: bool(chartN, 'autoTitleDeleted', false),
    plots: [],
    series: [],
    axes: {},
    legend: null,
    plotArea: null,
    chartArea: null,
    dispBlanksAs: val(chartN, 'dispBlanksAs', 'gap'),
    view3D: null,
  };
  // c14:style inside AlternateContent.
  for (const ac of root.all('mc:AlternateContent')) {
    const s = ac.first('c14:style');
    if (s) chart.style = Math.max(1, s.int('val', 102) - 100);
  }
  const t = chartN.child('c:title');
  if (t) chart.title = parseTitle(t, ctx, { bag: { b: '0', ...baseText.bag, sz: baseText.bag && baseText.bag.sz ? String(Math.round(Number(baseText.bag.sz) * 1.2)) : '1400' } });
  const v3 = chartN.child('c:view3D');
  if (v3) chart.view3D = { rotX: num(v3, 'rotX', 0), rotY: num(v3, 'rotY', 0), rAngAx: bool(v3, 'rAngAx', false), depth: num(v3, 'depthPercent', 100) };
  const pa = chartN.child('c:plotArea');
  if (pa) {
    const sp = spPr(pa, ctx);
    chart.plotArea = { layout: layoutOf(pa), fill: sp.fill, line: sp.line };
    for (const c of pa.children) {
      if (PLOT_TYPES.has(c.local)) {
        const plot = {
          type: c.local,
          barDir: val(c, 'barDir', 'col'),
          grouping: val(c, 'grouping', c.local.startsWith('bar') ? 'clustered' : 'standard'),
          varyColors: bool(c, 'varyColors', false),
          gapWidth: num(c, 'gapWidth', 150),
          overlap: num(c, 'overlap', c.child('c:grouping') && /stacked/i.test(val(c, 'grouping', '')) ? 100 : 0),
          holeSize: num(c, 'holeSize', 50),
          firstSliceAng: num(c, 'firstSliceAng', 0),
          scatterStyle: val(c, 'scatterStyle', 'marker'),
          radarStyle: val(c, 'radarStyle', 'marker'),
          showMarker: bool(c, 'marker', true),
          ofPieType: val(c, 'ofPieType', 'pie'),
          secondPieSize: num(c, 'secondPieSize', 75),
          splitType: val(c, 'splitType', 'auto'),
          splitPos: num(c, 'splitPos', undefined),
          bubbleScale: num(c, 'bubbleScale', 100),
          showNegBubbles: bool(c, 'showNegBubbles', false),
          axIds: c.all('c:axId').map((a) => a.attr('val', '')),
          dLbls: parseDLbls(c, ctx, baseText),
          hiLowLines: c.child('c:hiLowLines') ? (spPr(c.child('c:hiLowLines'), ctx).line || {}) : null,
          dropLines: c.child('c:dropLines') ? (spPr(c.child('c:dropLines'), ctx).line || {}) : null,
          serLines: c.child('c:serLines') ? (spPr(c.child('c:serLines'), ctx).line || {}) : null,
          upDownBars: null,
          series: [],
        };
        const udb = c.child('c:upDownBars');
        if (udb) {
          plot.upDownBars = {
            gapWidth: num(udb, 'gapWidth', 150),
            up: spPr(udb.child('c:upBars'), ctx),
            down: spPr(udb.child('c:downBars'), ctx),
          };
        }
        for (const s of c.all('c:ser')) {
          const ser = parseSeries(s, ctx, c.local, plot.dLbls ? plot.dLbls.text : baseText);
          ser.plot = plot;
          plot.series.push(ser);
          chart.series.push(ser);
        }
        plot.series.sort((a, b) => a.order - b.order);
        chart.plots.push(plot);
        if (/3D/.test(c.local) && c.local !== 'bar3DChart' && c.local !== 'pie3DChart') ctx.note('approx', '3D chart rendered in 2D');
        if (/surface/.test(c.local)) ctx.note('approx', 'surface chart');
      } else if (/Ax$/.test(c.local)) {
        const ax = parseAxis(c, ctx, baseText);
        chart.axes[ax.id] = ax;
      } else if (c.local === 'dTable') {
        chart.dataTable = { horz: bool(c, 'showHorzBorder', true), vert: bool(c, 'showVertBorder', true), outline: bool(c, 'showOutline', true), keys: bool(c, 'showKeys', true) };
      }
    }
  }
  const lg = chartN.child('c:legend');
  if (lg) {
    const sp = spPr(lg, ctx);
    const deleted = new Set();
    for (const le of lg.all('c:legendEntry')) if (bool(le, 'delete', false)) deleted.add(num(le, 'idx', 0));
    chart.legend = {
      pos: val(lg, 'legendPos', 'r'),
      overlay: bool(lg, 'overlay', false),
      layout: layoutOf(lg),
      style: txStyle(lg.child('c:txPr'), ctx, baseText),
      fill: sp.fill,
      line: sp.line,
      deleted,
    };
  }
  const sp = spPr(root, ctx);
  chart.chartArea = { fill: sp.fill, line: sp.line, effects: sp.effects };
  // Resolve text styles now (they need the context).
  chart.resolveText = null;
  chart.ctx = null;
  attachResolved(chart, ctx);
  void el; void finalizeLine; void mergeLines;
  return chart;
}

/** Pre-resolve text styles and colours that need the parse context. */
function attachResolved(chart, ctx) {
  const res = (st, size) => resolveText(st, ctx, size);
  chart.textResolved = res(chart.text, 10);
  if (chart.title) {
    chart.title.resolved = res(chart.title.style, 14);
  }
  if (chart.legend) chart.legend.resolved = res(chart.legend.style, 10);
  for (const ax of Object.values(chart.axes)) {
    ax.textResolved = res(ax.text, 10);
    if (ax.title) ax.title.resolved = res(ax.title.style, 10);
  }
  for (const s of chart.series) {
    if (s.dLbls && !s.dLbls.delete) s.dLbls.resolved = res(s.dLbls.text, 10);
  }
  for (const p of chart.plots) if (p.dLbls && !p.dLbls.delete) p.dLbls.resolved = res(p.dLbls.text, 10);
  // Automatic series colours.
  const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((a) => resolveColorNode({ name: 'a:schemeClr', attrs: { val: a }, children: [] }, ctx));
  chart.palette = accents;
  chart.tx1 = resolveColorNode({ name: 'a:schemeClr', attrs: { val: 'tx1' }, children: [] }, ctx);
  chart.bg1 = resolveColorNode({ name: 'a:schemeClr', attrs: { val: 'bg1' }, children: [] }, ctx);
}

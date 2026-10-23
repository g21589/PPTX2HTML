/**
 * Fixture catalog: pictures, groups, tables, charts, SmartArt, media, backgrounds and masters.
 */
import { shape, label, solid, ln, png, wav, emf, wmf, picture, group, graphicFrame, textBox, px, esc } from '../lib.mjs';

export const name = 'objects';
export const title = 'Pictures, groups, tables, charts, media, backgrounds';

const C_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function photo() {
  return png(160, 120, (x, y) => {
    const sky = y < 70;
    if (sky) return [120 + y, 170 + y / 2, 235, 255];
    const g = 90 + ((x * 7 + y * 3) % 40);
    return [40, g, 50, 255];
  });
}

function cache(values, numeric, fmt) {
  const pts = values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`)).join('');
  return numeric ? `<c:numCache><c:formatCode>${fmt || 'General'}</c:formatCode><c:ptCount val="${values.length}"/>${pts}</c:numCache>` : `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache>`;
}

function series(i, name, cats, vals, extra = '') {
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$1</c:f>${cache([name], false)}</c:strRef></c:tx>${extra}<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${cats.length + 1}</c:f>${cache(cats, false)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$${String.fromCharCode(66 + i)}$2</c:f>${cache(vals, true)}</c:numRef></c:val></c:ser>`;
}

function xySeries(i, name, xs, ys, extra = '', sizes) {
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${esc(name)}</c:v></c:tx>${extra}<c:xVal><c:numRef><c:f>x</c:f>${cache(xs, true)}</c:numRef></c:xVal><c:yVal><c:numRef><c:f>y</c:f>${cache(ys, true)}</c:numRef></c:yVal>${sizes ? `<c:bubbleSize><c:numRef><c:f>s</c:f>${cache(sizes, true)}</c:numRef></c:bubbleSize>` : ''}</c:ser>`;
}

const axes = (catPos = 'b', valPos = 'l', o = {}) => `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/>${o.max !== undefined ? `<c:max val="${o.max}"/>` : ''}</c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${o.grid === false ? '' : '<c:majorGridlines/>'}${o.title ? `<c:title><c:tx><c:rich><a:bodyPr rot="-5400000"/><a:p><a:r><a:t>${o.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : ''}<c:numFmt formatCode="${o.fmt || 'General'}" sourceLinked="${o.fmt ? 0 : 1}"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;

const valAxes = () => `<c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;

function chartSpace(title, plot, o = {}) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${C_NS}><c:roundedCorners val="0"/><c:chart>${title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400"/><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>'}<c:plotArea><c:layout/>${plot}</c:plotArea>${o.legend === false ? '' : `<c:legend><c:legendPos val="${o.legend || 'b'}"/><c:overlay val="0"/></c:legend>`}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>`;
}

export function build(deck) {
  const out = [];
  // --- Pictures -------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Pictures' });
    label(s, 'Pictures: crop, shape clip, rotation, flip, outline, effects, SVG');
    const rid = s.image(photo(), 'png');
    const add = (o) => s.add(picture({ id: s.id(), rid, ...o }));
    add({ name: 'plain', x: 40, y: 70, w: 240, h: 180 });
    add({ name: 'crop', x: 320, y: 70, w: 240, h: 180, srcRect: '<a:srcRect l="25000" t="0" r="25000" b="40000"/>' });
    add({ name: 'ellipse', x: 600, y: 70, w: 240, h: 180, prst: 'ellipse' });
    add({ name: 'rot', x: 880, y: 70, w: 240, h: 180, rot: 20, line: ln(4, solid('C00000')) });
    add({ name: 'flipH', x: 40, y: 290, w: 240, h: 180, flipH: true });
    add({ name: 'gray', x: 320, y: 290, w: 240, h: 180, blipInner: '<a:grayscl/>' });
    add({ name: 'duotone', x: 600, y: 290, w: 240, h: 180, blipInner: '<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>' });
    add({ name: 'bright', x: 880, y: 290, w: 240, h: 180, blipInner: '<a:lum bright="40000" contrast="-20000"/>' });
    add({ name: 'shadow', x: 40, y: 510, w: 240, h: 160, effects: '<a:effectLst><a:outerShdw blurRad="76200" dist="50800" dir="2700000"><a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:outerShdw></a:effectLst>' });
    add({ name: 'soft', x: 320, y: 510, w: 240, h: 160, effects: '<a:effectLst><a:softEdge rad="254000"/></a:effectLst>' });
    const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100" viewBox="0 0 100 100"><circle cx="50" cy="50" r="45" fill="#7030A0"/><path d="M30 50h40M50 30v40" stroke="#fff" stroke-width="8"/></svg>');
    const svgRid = s.image(svg, 'svg');
    const pngRid = s.image(png(10, 10, () => [255, 0, 0]), 'png');
    s.add(picture({ id: s.id(), name: 'svg', rid: pngRid, x: 640, y: 510, w: 160, h: 160, blipInner: `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svgRid}"/></a:ext></a:extLst>` }));
    const keyed = png(80, 60, (x, y) => (x > 20 && x < 60 && y > 15 && y < 45 ? [0, 112, 192, 255] : [255, 255, 255, 255]));
    const kRid = s.image(keyed, 'png');
    s.add(textBox({ id: s.id(), name: 'behind', x: 860, y: 520, w: 240, h: 140, text: [{ runs: [{ t: 'Transparent colour', rPr: 'sz="2400" b="1"', inner: solid('FF0000') }] }] }));
    s.add(picture({ id: s.id(), name: 'clrChange', rid: kRid, x: 860, y: 510, w: 240, h: 160, blipInner: '<a:clrChange><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom><a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>' }));
    out.push({
      slide: s.index, feature: 'picture', compareLO: true,
      checks: [
        { kind: 'pixel', x: 160, y: 80, color: '#7EAAEB', tol: 40 },
        { kind: 'pixel', x: 604, y: 74, color: '#FFFFFF', tol: 10 },
        { kind: 'pixel', x: 680, y: 560, color: '#7030A0', tol: 30 },
        { kind: 'dom', selector: 'feColorMatrix', min: 2 },
      ],
    });
  }

  // --- Metafiles ---------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Metafiles' });
    label(s, 'Metafile pictures: EMF and WMF converted to SVG (brushes, pens, shapes, hatches)');
    // COLORREF is 0x00bbggrr.
    const emfRid = s.image(emf(1000, 750, [
      [39, [1, 0, 0x0050b000, 0]], [37, [1]], // green brush
      [38, [2, 0, 6, 0, 0]], [37, [2]], // black pen
      [43, [100, 100, 900, 650]], // RECTANGLE
      [39, [3, 0, 0x000000ff, 0]], [37, [3]], // red brush
      [42, [300, 200, 700, 550]], // ELLIPSE
      [38, [4, 0, 10, 0, 0x00c07000]], [37, [4]], // blue pen
      [4, [150, 600, 850, 700, 3, 150, 700, 500, 600, 850, 700]], // POLYLINE
    ]), 'emf');
    s.add(picture({ id: s.id(), name: 'emf', rid: emfRid, x: 40, y: 110, w: 500, h: 375 }));
    const wmfRid = s.image(wmf(1000, 750, [
      [0x020b, [0, 0]], [0x020c, [750, 1000]], // window origin / extent
      [0x02fc, [0, 0x7000, 0x00c0, 0]], [0x012d, [0]], // blue brush
      [0x02fa, [0, 6, 0, 0, 0]], [0x012d, [1]], // black pen
      [0x041b, [650, 900, 100, 100]], // RECTANGLE
      [0x02fc, [0, 0xc0ff, 0, 0]], [0x012d, [2]], // yellow brush
      [0x0418, [550, 700, 200, 300]], // ELLIPSE
      [0x02fc, [2, 0x00ff, 0, 5]], [0x012d, [3]], // red diagonal-cross hatch
      [0x0324, [3, 200, 740, 500, 660, 800, 740]], // POLYGON
    ]), 'wmf');
    s.add(picture({ id: s.id(), name: 'wmf', rid: wmfRid, x: 600, y: 110, w: 500, h: 375 }));
    out.push({
      slide: s.index, feature: 'metafile', compareLO: true,
      checks: [
        { kind: 'pixel', x: 290, y: 297, color: '#FF0000', tol: 20 },
        { kind: 'pixel', x: 115, y: 185, color: '#00B050', tol: 20 },
        { kind: 'pixel', x: 65, y: 135, color: '#FFFFFF', tol: 10 },
        { kind: 'pixel', x: 850, y: 297, color: '#FFC000', tol: 20 },
        { kind: 'pixel', x: 675, y: 185, color: '#0070C0', tol: 20 },
        { kind: 'pixel', x: 625, y: 135, color: '#FFFFFF', tol: 10 },
        { kind: 'dom', selector: 'image[href^="data:image/svg+xml"]', min: 2 },
      ],
    });
  }

  // --- Groups ------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Groups' });
    label(s, 'Groups: nested, rotated, flipped, scaled (chExt != ext), group fill');
    const kid = (x, y, c, t) => shape({ id: s.id(), name: t, x, y, w: 100, h: 60, fill: solid(c), text: [{ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="1200"' }] }] });
    s.add(group({ id: s.id(), name: 'g1', x: 40, y: 80, w: 220, h: 140 }, [kid(40, 80, 'C00000', 'A'), kid(160, 160, '0070C0', 'B')]));
    s.add(group({ id: s.id(), name: 'rot', x: 320, y: 80, w: 220, h: 140, rot: 30 }, [kid(320, 80, 'C00000', 'A'), kid(440, 160, '0070C0', 'B')]));
    s.add(group({ id: s.id(), name: 'flip', x: 600, y: 80, w: 220, h: 140, flipH: true }, [kid(600, 80, 'C00000', 'A'), kid(720, 160, '0070C0', 'B')]));
    s.add(group({ id: s.id(), name: 'scaled', x: 880, y: 80, w: 330, h: 210, chOff: [0, 0], chExt: [220, 140] }, [kid(0, 0, 'C00000', 'A'), kid(120, 80, '0070C0', 'B')]));
    s.add(group({ id: s.id(), name: 'nested', x: 40, y: 320, w: 500, h: 300 }, [
      kid(40, 320, '70AD47', 'outer'),
      group({ id: s.id(), name: 'inner', x: 200, y: 400, w: 340, h: 220, chOff: [0, 0], chExt: [170, 110] }, [kid(0, 0, 'FFC000', 'in1'), kid(70, 50, '7030A0', 'in2')]),
    ]));
    s.add(group({ id: s.id(), name: 'grpFill', x: 640, y: 360, w: 300, h: 200, fill: solid('ED7D31') }, [
      shape({ id: s.id(), name: 'uses group fill', x: 640, y: 360, w: 140, h: 200, fill: '<a:grpFill/>' }),
      shape({ id: s.id(), name: 'own fill', x: 800, y: 360, w: 140, h: 200, fill: solid('5B9BD5') }),
    ]));
    out.push({
      slide: s.index, feature: 'group', compareLO: true,
      checks: [
        { kind: 'pixel', x: 50, y: 88, color: '#C00000', tol: 25 },
        { kind: 'pixel', x: 1180, y: 260, color: '#0070C0', tol: 25 },
        { kind: 'pixel', x: 710, y: 520, color: '#ED7D31', tol: 25 },
        { kind: 'pixel', x: 812, y: 88, color: '#C00000', tol: 25 },
        { kind: 'dom', selector: '.pp-grp .pp-grp', min: 1 },
      ],
    });
  }

  // --- Tables -------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Tables' });
    label(s, 'Tables: built-in styles, banding, merged cells, borders, cell fill/margins/anchors');
    const tbl = (o) => {
      const cols = o.cols || [140, 140, 140];
      const rows = o.rows;
      const grid = cols.map((w) => `<a:gridCol w="${px(w)}"/>`).join('');
      const trs = rows.map((r) => `<a:tr h="${px(r.h || 36)}">${r.cells.map((c) => {
        if (c.merge) return `<a:tc ${c.merge}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>`;
        return `<a:tc${c.span ? ` ${c.span}` : ''}><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${c.pPr ? `<a:pPr ${c.pPr}/>` : ''}<a:r><a:rPr lang="en-US" sz="1200" dirty="0"/><a:t>${esc(c.t)}</a:t></a:r></a:p></a:txBody><a:tcPr ${c.tcPrA || ''}>${c.tcPr || ''}</a:tcPr></a:tc>`;
      }).join('')}</a:tr>`).join('');
      return graphicFrame(o, `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr ${o.flags || 'firstRow="1" bandRow="1"'}><a:tableStyleId>${o.style}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData>`);
    };
    const simple = (n) => [{ cells: ['Header 1', 'Header 2', 'Header 3'].map((t) => ({ t })) }, ...Array.from({ length: n }, (_, r) => ({ cells: [0, 1, 2].map((c) => ({ t: `R${r + 1}C${c + 1}` })) }))];
    s.add(tbl({ id: s.id(), name: 'Medium Style 2 Accent 1', x: 40, y: 70, w: 420, h: 180, style: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', rows: simple(4) }));
    s.add(tbl({ id: s.id(), name: 'Light Style 1 Accent 2', x: 500, y: 70, w: 420, h: 180, style: '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}', rows: simple(4) }));
    s.add(tbl({ id: s.id(), name: 'Dark Style 1', x: 960, y: 70, w: 280, h: 180, cols: [90, 90, 100], style: '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}', flags: 'firstRow="1" firstCol="1" lastRow="1" bandRow="1"', rows: simple(4) }));
    s.add(tbl({
      id: s.id(), name: 'Merged and formatted', x: 40, y: 300, w: 600, h: 300, cols: [150, 150, 150, 150], style: '{2D5ABB26-0587-4C30-8999-92F81FD0307C}', flags: 'bandRow="1"',
      rows: [
        { h: 50, cells: [{ t: 'Merged across 2', span: 'gridSpan="2"', tcPr: solid('FFC000'), pPr: 'algn="ctr"' }, { merge: 'hMerge="1"' }, { t: 'Diagonal', tcPr: `<a:lnTlToBr w="12700">${solid('FF0000')}</a:lnTlToBr>` }, { t: 'Thick border', tcPr: `<a:lnL w="50800">${solid('0070C0')}</a:lnL><a:lnR w="50800">${solid('0070C0')}</a:lnR><a:lnT w="50800">${solid('0070C0')}</a:lnT><a:lnB w="50800">${solid('0070C0')}</a:lnB>` }] },
        { h: 100, cells: [{ t: 'Row span 2', span: 'rowSpan="2"', tcPr: solid('A9D18E'), tcPrA: 'anchor="ctr"' }, { t: 'bottom', tcPrA: 'anchor="b"' }, { t: 'centre', tcPrA: 'anchor="ctr"', pPr: 'algn="ctr"' }, { t: 'margins', tcPrA: 'marL="365760" marT="182880"' }] },
        { h: 60, cells: [{ merge: 'vMerge="1"' }, { t: 'gradient', tcPr: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFF00"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>' }, { t: 'vertical', tcPrA: 'vert="vert270"' }, { t: 'dash', tcPr: `<a:lnB w="25400">${solid('000000')}<a:prstDash val="dash"/></a:lnB>` }] },
      ],
    }));
    out.push({
      slide: s.index, feature: 'table', compareLO: true,
      checks: [
        { kind: 'pixel', x: 150, y: 80, color: '#4472C4', tol: 25 },
        { kind: 'pixel', x: 150, y: 125, color: '#CFD5EA', tol: 25 },
        { kind: 'pixel', x: 150, y: 160, color: '#E9EBF5', tol: 25 },
        { kind: 'pixel', x: 400, y: 380, color: '#FFFFFF', tol: 10 },
        { kind: 'pixel', x: 200, y: 310, color: '#FFC000', tol: 25 },
        { kind: 'dom', selector: 'td[colspan="2"]', min: 1 },
        { kind: 'dom', selector: 'td[rowspan="2"]', min: 1 },
        { kind: 'dom', selector: '.pp-cell-diag', min: 1 },
      ],
    });
  }

  // --- Charts --------------------------------------------------------------------
  {
    const cats = ['Q1', 'Q2', 'Q3', 'Q4'];
    const charts = [
      ['Clustered column', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, 'North', cats, [4.3, 2.5, 3.5, 4.5])}${series(1, 'South', cats, [2.4, 4.4, 1.8, 2.8])}<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes()}`],
      ['Stacked bar (horizontal)', `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${series(0, 'A', cats, [4, 3, 2, 5])}${series(1, 'B', cats, [1, 2, 3, 1])}<c:gapWidth val="80"/><c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes('l', 'b')}`],
      ['100% stacked column', `<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>${series(0, 'A', cats, [4, 3, 2, 5])}${series(1, 'B', cats, [1, 2, 3, 1])}<c:overlap val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes('b', 'l', { fmt: '0%' })}`],
      ['Line with markers', `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series(0, 'Sales', cats, [10, 14, 9, 17], '<c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>')}${series(1, 'Cost', cats, [7, 8, 11, 9], '<c:marker><c:symbol val="square"/><c:size val="6"/></c:marker><c:smooth val="1"/>')}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes()}`],
      ['Area', `<c:areaChart><c:grouping val="stacked"/><c:varyColors val="0"/>${series(0, 'A', cats, [3, 4, 2, 5])}${series(1, 'B', cats, [2, 1, 3, 2])}<c:axId val="1"/><c:axId val="2"/></c:areaChart>${axes()}`],
      ['Pie with labels', `<c:pieChart><c:varyColors val="1"/>${series(0, 'Share', ['Alpha', 'Beta', 'Gamma', 'Delta'], [45, 25, 20, 10], '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:explosion val="12"/></c:dPt><c:dLbls><c:numFmt formatCode="0%" sourceLinked="0"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>')}<c:firstSliceAng val="0"/></c:pieChart>`],
      ['Doughnut', `<c:doughnutChart><c:varyColors val="1"/>${series(0, 'D', cats, [1, 2, 3, 4])}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`],
      ['Scatter', `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${xySeries(0, 'pts', [1, 2, 3.5, 5, 7], [2, 4.5, 3, 6, 7.2], '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>')}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>${valAxes()}`],
      ['Bubble', `<c:bubbleChart><c:varyColors val="0"/>${xySeries(0, 'b', [1, 2, 3, 4], [3, 1, 4, 2], '', [10, 40, 25, 60])}<c:bubbleScale val="100"/><c:showNegBubbles val="0"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart>${valAxes()}`],
      ['Radar', `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${series(0, 'r1', ['a', 'b', 'c', 'd', 'e'], [4, 3, 5, 2, 4])}${series(1, 'r2', ['a', 'b', 'c', 'd', 'e'], [2, 5, 3, 4, 1])}<c:axId val="1"/><c:axId val="2"/></c:radarChart>${axes('b', 'l', { grid: true })}`],
      ['Stock (HLC)', `<c:stockChart>${series(0, 'High', cats, [12, 14, 13, 15])}${series(1, 'Low', cats, [8, 9, 10, 11])}${series(2, 'Close', cats, [10, 12, 11, 14], '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="dash"/><c:size val="8"/></c:marker>')}<c:hiLowLines/><c:axId val="1"/><c:axId val="2"/></c:stockChart>${axes()}`],
      ['Combo + data labels', `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series(0, 'Units', cats, [120, 150, 90, 180], '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series(1, 'Trend', cats, [100, 130, 110, 170])}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes('b', 'l', { title: 'Units' })}`],
    ];
    for (let k = 0; k < charts.length; k += 6) {
      const s = deck.slide({ name: 'Charts ' + (k / 6 + 1) });
      label(s, 'Charts ' + (k / 6 + 1));
      const sub = charts.slice(k, k + 6);
      sub.forEach(([t, plot], i) => {
        const col = i % 3, row = Math.floor(i / 3);
        const rid = s.chart(chartSpace(t, plot, { legend: t.startsWith('Pie') ? 'r' : 'b' }));
        s.add(graphicFrame({ id: s.id(), name: t, x: 20 + col * 420, y: 60 + row * 330, w: 400, h: 310 }, `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData>`));
      });
      out.push({
        slide: s.index, feature: 'chart/' + (k / 6 + 1), compareLO: true,
        checks: [{ kind: 'dom', selector: '.pp-chart svg', min: sub.length }, { kind: 'text', contains: sub[0][0] }, { kind: 'model', desc: 'chart series parsed', fn: (sl) => sl.elements.filter((e) => e.kind === 'chart' && e.chart.series.length).length === sub.length }],
      });
    }
  }

  // --- Backgrounds & master -----------------------------------------------------------
  {
    const bgs = [
      ['Solid background', `<p:bg><p:bgPr>${solid('DEEBF7')}<a:effectLst/></p:bgPr></p:bg>`, [[640, 400, '#DEEBF7']]],
      ['Gradient background', '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="1F4E79"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>', [[640, 4, '#FFFFFF'], [640, 716, '#1F4E79']]],
      ['Theme background (bgRef 1003)', '<p:bg><p:bgRef idx="1003"><a:schemeClr val="accent1"/></p:bgRef></p:bg>', []],
      ['Pattern background', '<p:bg><p:bgPr><a:pattFill prst="smGrid"><a:fgClr><a:srgbClr val="BDD7EE"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill><a:effectLst/></p:bgPr></p:bg>', []],
    ];
    for (const [t, bg, pts] of bgs) {
      const s = deck.slide({ name: t, bg });
      label(s, t);
      out.push({ slide: s.index, feature: 'background/' + t.split(' ')[0].toLowerCase(), compareLO: true, checks: pts.map(([x, y, c]) => ({ kind: 'pixel', x, y, color: c, tol: 25 })) });
    }
    const s = deck.slide({ name: 'Picture background', layout: 'blank' });
    const rid = s.image(photo(), 'png');
    s.bg = `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
    label(s, 'Picture background');
    out.push({ slide: s.index, feature: 'background/picture', compareLO: true, checks: [{ kind: 'pixel', x: 640, y: 680, color: '#2E7032', tol: 70 }] });
  }

  // --- Placeholders / layouts -------------------------------------------------------
  {
    const s = deck.slide({ name: 'Title layout', layout: 'title' });
    s.add(shape({ id: s.id(), name: 'Title 1', ph: 'type="ctrTitle"', noXfrm: true, style: '', text: [{ runs: [{ t: 'Title Slide Layout' }] }], bodyPr: '<a:bodyPr/>' }));
    s.add(shape({ id: s.id(), name: 'Subtitle 2', ph: 'type="subTitle" idx="1"', noXfrm: true, style: '', text: [{ runs: [{ t: 'Placeholder inheritance from layout and master' }] }], bodyPr: '<a:bodyPr/>' }));
    s.add(shape({ id: s.id(), name: 'Slide Number', ph: 'type="sldNum" sz="quarter" idx="12"', noXfrm: true, style: '', text: [{ runs: [{ fld: 'slidenum', t: '1' }] }], bodyPr: '<a:bodyPr/>' }));
    out.push({
      slide: s.index, feature: 'placeholder/title', compareLO: true,
      checks: [
        { kind: 'text', contains: 'Title Slide Layout' },
        { kind: 'model', desc: 'title inherits 60pt centred', fn: (sl) => { const t = sl.elements.find((e) => e.name === 'Title 1'); return t && t.text.paragraphs[0].runs[0].style.size === 60 && t.text.paragraphs[0].align === 'center'; } },
        { kind: 'model', desc: 'title inherits layout position', fn: (sl) => { const t = sl.elements.find((e) => e.name === 'Title 1'); return t && Math.round(t.x) === 160; } },
        { kind: 'text', contains: String(s.index) },
      ],
    });
    const s2 = deck.slide({ name: 'Content layout', layout: 'content' });
    s2.add(shape({ id: s2.id(), name: 'Title 1', ph: 'type="title"', noXfrm: true, style: '', text: [{ runs: [{ t: 'Title and Content' }] }], bodyPr: '<a:bodyPr/>' }));
    s2.add(shape({
      id: s2.id(), name: 'Content 2', ph: 'idx="1"', noXfrm: true, style: '', bodyPr: '<a:bodyPr/>',
      text: [{ runs: [{ t: 'Level 1 bullet from master body style' }] }, { pPr: 'lvl="1"', runs: [{ t: 'Level 2' }] }, { pPr: 'lvl="2"', runs: [{ t: 'Level 3' }] }, { runs: [{ t: 'Another level 1' }] }],
    }));
    out.push({
      slide: s2.index, feature: 'placeholder/content', compareLO: true,
      checks: [
        { kind: 'text', contains: '•Level 1 bullet' },
        { kind: 'model', desc: 'body levels sizes 28/24/20', fn: (sl) => { const b = sl.elements.find((e) => e.name === 'Content 2'); return b && b.text.paragraphs.map((p) => p.runs[0].style.size).join(',') === '28,24,20,28'; } },
      ],
    });
    const s3 = deck.slide({ name: 'Colour map override', layout: 'blank', clrMapOvr: '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' });
    s3.bg = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
    s3.add(textBox({ id: s3.id(), name: 'light text', x: 100, y: 300, w: 1000, h: 100, text: [{ runs: [{ t: 'White text on black via clrMapOvr', rPr: 'sz="4000"' }] }] }));
    out.push({ slide: s3.index, feature: 'placeholder/clrMapOvr', compareLO: true, checks: [{ kind: 'pixel', x: 640, y: 600, color: '#000000', tol: 10 }, { kind: 'model', desc: 'text colour white', fn: (sl) => { const e = sl.elements.find((x) => x.name === 'light text'); const c = e.text.paragraphs[0].runs[0].style.fill.color; return c.r > 250 && c.g > 250 && c.b > 250; } }] });
    const s4 = deck.slide({ name: 'Hidden slide', hidden: true, notes: 'These are speaker notes.' });
    label(s4, 'Hidden slide with speaker notes');
    out.push({ slide: s4.index, feature: 'slide/hidden-notes', compareLO: false, checks: [{ kind: 'model', desc: 'hidden flag + notes', fn: (sl) => sl.hidden === true && sl.notes === 'These are speaker notes.' }] });
  }

  // --- Media ------------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Media' });
    label(s, 'Audio and video with poster frames');
    const a = s.mediaFile(wav(), 'wav', 'audio');
    const icon = s.image(png(64, 64, (x, y) => (Math.hypot(x - 32, y - 32) < 28 ? [68, 114, 196] : [255, 255, 255, 0])), 'png');
    s.add(`<p:pic><p:nvPicPr><p:cNvPr id="${s.id()}" name="audio.wav"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="${a.link}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE67E2D6D2C}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${a.embed}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${icon}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${px(100)}" y="${px(200)}"/><a:ext cx="${px(120)}" cy="${px(120)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
    out.push({ slide: s.index, feature: 'media/audio', compareLO: false, checks: [{ kind: 'dom', selector: '.pp-media audio[src^="data:audio/wav"]', min: 1 }] });
  }

  // --- SmartArt (pre-rendered drawing) ------------------------------------------------
  {
    const s = deck.slide({ name: 'SmartArt' });
    label(s, 'SmartArt (diagram drawing part)');
    const DSP = 'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';
    const boxes = [0, 1, 2].map((i) => `<dsp:sp modelId="{0000000${i}-0000-0000-0000-000000000000}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="${px(i * 260)}" y="${px(40)}"/><a:ext cx="${px(220)}" cy="${px(140)}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent${i + 1}"/></a:solidFill></dsp:spPr><dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2000"/><a:t>Step ${i + 1}</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="${px(i * 260)}" y="${px(40)}"/><a:ext cx="${px(220)}" cy="${px(140)}"/></dsp:txXfrm></dsp:sp>`).join('');
    const drawingRid = s.part('diagrams/drawing1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dsp:drawing ${DSP}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${boxes}</dsp:spTree></dsp:drawing>`, 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', 'application/vnd.ms-office.drawingml.diagramDrawing+xml');
    const dataRid = s.part('diagrams/data1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel ${DSP}><dgm:ptLst><dgm:pt modelId="{A}" type="doc"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="${drawingRid}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData', 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml');
    s.add(graphicFrame({ id: s.id(), name: 'Diagram', x: 200, y: 200, w: 760, h: 220 }, `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="${dataRid}" r:lo="${dataRid}" r:qs="${dataRid}" r:cs="${dataRid}"/></a:graphicData>`));
    out.push({ slide: s.index, feature: 'smartart', compareLO: true, checks: [{ kind: 'text', contains: 'Step 2' }, { kind: 'pixel', x: 310, y: 250, color: '#4472C4', tol: 25 }] });
  }

  // --- AlternateContent fallback (equation-like) --------------------------------------
  {
    const s = deck.slide({ name: 'AlternateContent' });
    label(s, 'mc:AlternateContent: unsupported Choice falls back');
    s.add(`<mc:AlternateContent><mc:Choice Requires="unknownExt">${shape({ id: s.id(), name: 'choice', x: 100, y: 200, w: 300, h: 100, fill: solid('FF0000') })}</mc:Choice><mc:Fallback>${shape({ id: s.id(), name: 'fallback', x: 100, y: 200, w: 300, h: 100, fill: solid('00B050'), text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Fallback used' }] }] })}</mc:Fallback></mc:AlternateContent>`);
    s.add(`<mc:AlternateContent><mc:Choice Requires="p14">${shape({ id: s.id(), name: 'p14 choice', x: 500, y: 200, w: 300, h: 100, fill: solid('0070C0'), text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Choice used' }] }] })}</mc:Choice><mc:Fallback>${shape({ id: s.id(), name: 'fb2', x: 500, y: 200, w: 300, h: 100, fill: solid('FF0000') })}</mc:Fallback></mc:AlternateContent>`);
    out.push({ slide: s.index, feature: 'alternate-content', compareLO: false, checks: [{ kind: 'pixel', x: 250, y: 220, color: '#00B050', tol: 20 }, { kind: 'pixel', x: 650, y: 220, color: '#0070C0', tol: 20 }] });
  }
  return out;
}

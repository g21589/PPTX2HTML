/**
 * Fixture catalog: geometry, fills, outlines and effects.
 */
import { shape, label, solid, scheme, ln, png, picture } from '../lib.mjs';
import { presetNames } from '../../../src/geometry/geometry.js';

export const name = 'drawing';
export const title = 'Shapes, fills, lines and effects';

const PATTERNS = ['pct5', 'pct10', 'pct20', 'pct25', 'pct30', 'pct40', 'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90', 'horz', 'vert',
  'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert', 'dashHorz', 'dashVert', 'cross', 'dnDiag', 'upDiag', 'ltDnDiag', 'ltUpDiag',
  'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag', 'dashUpDiag', 'diagCross', 'smCheck', 'lgCheck', 'smGrid', 'lgGrid', 'dotGrid',
  'smConfetti', 'lgConfetti', 'horzBrick', 'diagBrick', 'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot', 'shingle', 'wave',
  'trellis', 'zigZag'];

export function build(deck) {
  const out = [];
  // --- All preset geometries -------------------------------------------------
  const names = presetNames();
  const perSlide = 40;
  for (let p = 0; p * perSlide < names.length; p++) {
    const s = deck.slide({ name: `Preset shapes ${p + 1}` });
    label(s, `Preset geometries ${p * perSlide + 1}-${Math.min(names.length, (p + 1) * perSlide)}`);
    const chunk = names.slice(p * perSlide, (p + 1) * perSlide);
    const checks = [];
    chunk.forEach((n, i) => {
      const col = i % 8, row = Math.floor(i / 8);
      const x = 30 + col * 155, y = 70 + row * 128;
      const lineLike = /^(line|straightConnector1|bentConnector[2-5]|curvedConnector[2-5]|arc|leftBracket|rightBracket|leftBrace|rightBrace|bracketPair|bracePair)$/.test(n);
      s.add(shape({ id: s.id(), name: n, x, y, w: 110, h: 80, prst: n, line: lineLike ? ln(2, solid('C00000')) : undefined }));
      s.add(shape({ id: s.id(), name: 'cap ' + n, x: x - 20, y: y + 84, w: 150, h: 22, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t: n, rPr: 'sz="900"' }] }], bodyPr: '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>' }));
    });
    checks.push({ kind: 'dom', selector: '.pp-sp svg.pp-geom path', min: chunk.length });
    checks.push({ kind: 'model', desc: 'all presets known', fn: (sl) => sl.elements.filter((e) => e.geom && e.geom.paths.some((pp) => pp.segs.length)).length >= chunk.length });
    out.push({ slide: s.index, feature: `geometry/presets-${p + 1}`, checks, compareLO: true });
  }

  // --- Custom geometry ----------------------------------------------------------
  {
    const s = deck.slide({ name: 'Custom geometry' });
    label(s, 'Custom geometry (moveTo, lnTo, arcTo, quadBezTo, cubicBezTo, close, guides)');
    const cust = `<a:custGeom><a:avLst/><a:gdLst><a:gd name="hx" fmla="*/ w 1 2"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="200" h="200"><a:moveTo><a:pt x="0" y="200"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="200" y="200"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`;
    s.add(shape({ id: s.id(), name: 'tri', x: 60, y: 100, w: 240, h: 240, geom: cust, fill: solid('2E75B6') }));
    const curve = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="400" h="200"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:cubicBezTo><a:pt x="100" y="-50"/><a:pt x="300" y="250"/><a:pt x="400" y="100"/></a:cubicBezTo><a:lnTo><a:pt x="400" y="200"/></a:lnTo><a:quadBezTo><a:pt x="200" y="120"/><a:pt x="0" y="200"/></a:quadBezTo><a:close/></a:path></a:pathLst></a:custGeom>`;
    s.add(shape({ id: s.id(), name: 'curve', x: 360, y: 120, w: 400, h: 200, geom: curve, fill: solid('70AD47') }));
    const arc = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="200" h="200"><a:moveTo><a:pt x="100" y="0"/></a:moveTo><a:arcTo wR="100" hR="100" stAng="16200000" swAng="10800000"/><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`;
    s.add(shape({ id: s.id(), name: 'arc', x: 820, y: 100, w: 240, h: 240, geom: arc, fill: solid('FFC000') }));
    const holes = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path><a:path w="100" h="100" fill="none"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo></a:path><a:path w="100" h="100" fill="darken" stroke="0"><a:moveTo><a:pt x="50" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`;
    s.add(shape({ id: s.id(), name: 'multipath', x: 100, y: 420, w: 220, h: 220, geom: holes, fill: solid('5B9BD5'), line: ln(2) }));
    out.push({
      slide: s.index, feature: 'geometry/custom', compareLO: true,
      checks: [
        { kind: 'pixel', x: 180, y: 300, color: '#2E75B6' },
        { kind: 'pixel', x: 1000, y: 220, color: '#FFC000' },
        { kind: 'pixel', x: 860, y: 180, color: '#FFFFFF' },
        { kind: 'pixel', x: 300, y: 432, color: darken('#5B9BD5', 0.6), tol: 40 },
      ],
    });
  }

  // --- Solid fills and colour models ----------------------------------------------
  {
    const s = deck.slide({ name: 'Solid fills' });
    label(s, 'Solid fills: srgb, scheme (+transforms), sys, prst, hsl, scrgb, alpha');
    const items = [
      ['srgb FF0000', solid('FF0000'), '#FF0000'],
      ['scheme accent1', scheme('accent1'), '#4472C4'],
      ['accent2 lumMod60 lumOff40', scheme('accent2', '<a:lumMod val="60000"/><a:lumOff val="40000"/>'), '#F4B183'],
      ['accent1 lumMod75', scheme('accent1', '<a:lumMod val="75000"/>'), '#2F5597'],
      ['accent6 shade 50%', scheme('accent6', '<a:shade val="50000"/>'), null],
      ['accent4 tint 50%', scheme('accent4', '<a:tint val="50000"/>'), null],
      ['sysClr windowText', '<a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill>', '#000000'],
      ['prstClr orange', '<a:solidFill><a:prstClr val="orange"/></a:solidFill>', '#FFA500'],
      ['hslClr 120/100/50', '<a:solidFill><a:hslClr hue="7200000" sat="100000" lum="50000"/></a:solidFill>', '#00FF00'],
      ['scrgbClr blue', '<a:solidFill><a:scrgbClr r="0" g="0" b="100000"/></a:solidFill>', '#0000FF'],
      ['alpha 50% on white', solid('0000FF', '<a:alpha val="50000"/>'), '#7F7FFF'],
      ['inv of 00FFFF', solid('00FFFF', '<a:inv/>'), '#FF0000'],
      ['comp of red', solid('FF0000', '<a:comp/>'), '#00FFFF'],
      ['gray of green', solid('00FF00', '<a:gray/>'), null],
      ['hueOff +120', solid('FF0000', '<a:hueOff val="7200000"/>'), '#00FF00'],
      ['satMod 0', solid('FF0000', '<a:satMod val="0"/>'), '#808080'],
      ['bg2 (lt2)', scheme('bg2'), '#E7E6E6'],
      ['tx2 (dk2)', scheme('tx2'), '#44546A'],
    ];
    const checks = [];
    items.forEach(([t, fill, hex], i) => {
      const col = i % 6, row = Math.floor(i / 6);
      const x = 30 + col * 205, y = 80 + row * 210;
      s.add(shape({ id: s.id(), name: t, x, y, w: 180, h: 140, fill, line: '<a:ln><a:noFill/></a:ln>' }));
      s.add(shape({ id: s.id(), name: 'cap', x, y: y + 145, w: 180, h: 30, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="1000"' }] }] }));
      if (hex) checks.push({ kind: 'pixel', x: x + 90, y: y + 70, color: hex, tol: 12 });
    });
    out.push({ slide: s.index, feature: 'fill/solid-colors', checks, compareLO: true });
  }

  // --- Gradient fills ------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Gradients' });
    label(s, 'Gradient fills: linear angles (scaled/unscaled), radial, rect, path, multi-stop, alpha');
    const gs = (stops) => `<a:gsLst>${stops.map(([p, c, a]) => `<a:gs pos="${p}"><a:srgbClr val="${c}">${a !== undefined ? `<a:alpha val="${a}"/>` : ''}</a:srgbClr></a:gs>`).join('')}</a:gsLst>`;
    const g = [
      ['lin 0', `<a:gradFill>${gs([[0, 'FF0000'], [100000, '0000FF']])}<a:lin ang="0" scaled="0"/></a:gradFill>`, [[15, 70, '#FF0000'], [165, 70, '#0000FF']]],
      ['lin 90', `<a:gradFill>${gs([[0, 'FF0000'], [100000, '0000FF']])}<a:lin ang="5400000" scaled="0"/></a:gradFill>`, [[90, 8, '#FF0000'], [90, 132, '#0000FF']]],
      ['lin 45 scaled', `<a:gradFill>${gs([[0, 'FFFF00'], [100000, '00B050']])}<a:lin ang="2700000" scaled="1"/></a:gradFill>`, [[10, 10, '#FFFF00'], [170, 130, '#00B050']]],
      ['lin 135', `<a:gradFill>${gs([[0, '7030A0'], [100000, 'FFC000']])}<a:lin ang="8100000" scaled="0"/></a:gradFill>`, []],
      ['3 stops', `<a:gradFill>${gs([[0, 'FF0000'], [50000, '00FF00'], [100000, '0000FF']])}<a:lin ang="0" scaled="0"/></a:gradFill>`, [[90, 70, '#00FF00']]],
      ['radial circle', `<a:gradFill>${gs([[0, 'FFFFFF'], [100000, '002060']])}<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`, [[90, 70, '#FFFFFF']]],
      ['radial corner', `<a:gradFill>${gs([[0, 'FFFFFF'], [100000, 'C00000']])}<a:path path="circle"><a:fillToRect l="0" t="0" r="100000" b="100000"/></a:path></a:gradFill>`, [[3, 3, '#FFFFFF']]],
      ['rect path', `<a:gradFill>${gs([[0, 'FFFF00'], [100000, 'FF0000']])}<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`, [[90, 70, '#FFFF00'], [90, 4, '#FF1000'], [45, 35, '#FF8000']]],
      ['shape path', `<a:gradFill>${gs([[0, 'FFFFFF'], [100000, '00B0F0']])}<a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`, [[45, 35, '#80D8F8']]],
      ['alpha stops', `<a:gradFill>${gs([[0, '000000', 0], [100000, '000000', 100000]])}<a:lin ang="0" scaled="0"/></a:gradFill>`, [[8, 70, '#FFFFFF']]],
      ['theme style 2', null, []],
      ['rotWithShape rotated', `<a:gradFill rotWithShape="1">${gs([[0, 'FF0000'], [100000, '0000FF']])}<a:lin ang="0" scaled="0"/></a:gradFill>`, []],
    ];
    const checks = [];
    g.forEach(([t, fill, pts], i) => {
      const col = i % 6, row = Math.floor(i / 6);
      const x = 30 + col * 205, y = 80 + row * 300;
      const style = fill === null ? '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>' : undefined;
      s.add(shape({ id: s.id(), name: t, x, y, w: 180, h: 140, fill: fill || '', line: fill ? '<a:ln><a:noFill/></a:ln>' : '', style, rot: t.includes('rotated') ? 30 : 0 }));
      s.add(shape({ id: s.id(), name: 'cap', x, y: y + 150, w: 180, h: 30, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="1000"' }] }] }));
      for (const [dx, dy, c] of pts) checks.push({ kind: 'pixel', x: x + dx, y: y + dy, color: c, tol: 45 });
    });
    out.push({ slide: s.index, feature: 'fill/gradient', checks, compareLO: true });
  }

  // --- Pattern fills ---------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Patterns' });
    label(s, 'Pattern fills (all ' + PATTERNS.length + ' presets)');
    PATTERNS.forEach((pt, i) => {
      const col = i % 9, row = Math.floor(i / 9);
      const x = 30 + col * 137, y = 60 + row * 108;
      s.add(shape({ id: s.id(), name: pt, x, y, w: 120, h: 80, fill: `<a:pattFill prst="${pt}"><a:fgClr><a:srgbClr val="1F4E79"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`, line: ln(0.75) }));
      s.add(shape({ id: s.id(), name: 'cap', x, y: y + 80, w: 120, h: 20, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t: pt, rPr: 'sz="900"' }] }], bodyPr: '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>' }));
    });
    out.push({ slide: s.index, feature: 'fill/pattern', checks: [{ kind: 'dom', selector: 'pattern', min: PATTERNS.length }], compareLO: true });
  }

  // --- Picture fills ---------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Picture fills' });
    label(s, 'Picture fills: stretch, tile (scale/align/flip), crop, in shapes, transparency');
    const img = png(64, 64, (x, y) => (x < 32) === (y < 32) ? [220, 40, 40] : [40, 90, 220]);
    const rid = s.image(img, 'png');
    const items = [
      ['stretch', `<a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, 'rect', [[20, 20, '#DC2828'], [150, 20, '#285ADC']]],
      ['tile 50%', `<a:blipFill><a:blip r:embed="${rid}"/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="tl"/></a:blipFill>`, 'rect', [[8, 8, '#DC2828'], [24, 8, '#285ADC']]],
      ['tile flip xy', `<a:blipFill><a:blip r:embed="${rid}"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="xy" algn="tl"/></a:blipFill>`, 'rect', []],
      ['ellipse', `<a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, 'ellipse', [[4, 4, '#FFFFFF']]],
      ['crop left half', `<a:blipFill><a:blip r:embed="${rid}"/><a:srcRect r="50000"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, 'rect', [[20, 20, '#DC2828'], [160, 20, '#DC2828'], [160, 120, '#285ADC']]],
      ['alpha 40%', `<a:blipFill><a:blip r:embed="${rid}"><a:alphaModFix amt="40000"/></a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, 'rect', [[20, 20, '#F1A9A9']]],
      ['star', `<a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, 'star5', []],
      ['fillRect inset', `<a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect l="25000" t="25000" r="25000" b="25000"/></a:stretch></a:blipFill>`, 'rect', [[8, 8, '#FFFFFF']]],
    ];
    const checks = [];
    items.forEach(([t, fill, prst, pts], i) => {
      const col = i % 4, row = Math.floor(i / 4);
      const x = 40 + col * 305, y = 80 + row * 310;
      s.add(shape({ id: s.id(), name: t, x, y, w: 180, h: 140, prst, fill, line: '<a:ln><a:noFill/></a:ln>' }));
      s.add(shape({ id: s.id(), name: 'cap', x, y: y + 150, w: 180, h: 30, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="1000"' }] }] }));
      for (const [dx, dy, c] of pts) checks.push({ kind: 'pixel', x: x + dx, y: y + dy, color: c, tol: 40 });
    });
    out.push({ slide: s.index, feature: 'fill/picture', checks, compareLO: true });
  }

  // --- Lines ---------------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Lines' });
    label(s, 'Outlines: widths, dashes, caps, joins, compound lines');
    const dashes = ['solid', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot'];
    dashes.forEach((d, i) => {
      const y = 70 + i * 34;
      s.add(shape({ id: s.id(), name: 'dash ' + d, x: 40, y, w: 380, h: 0, prst: 'line', style: '', line: `<a:ln w="38100" cap="flat">${solid('2F5597')}<a:prstDash val="${d}"/></a:ln>` }));
      s.add(shape({ id: s.id(), name: 'cap', x: 430, y: y - 12, w: 150, h: 24, style: '', fill: '<a:noFill/>', text: [{ runs: [{ t: d, rPr: 'sz="1000"' }] }], bodyPr: '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/>' }));
    });
    [0.25, 1, 3, 6, 12].forEach((w, i) => s.add(shape({ id: s.id(), name: 'w' + w, x: 40, y: 470 + i * 40, w: 380, h: 0, prst: 'line', style: '', line: ln(w, solid('C00000')) })));
    ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'].forEach((c, i) => {
      s.add(shape({ id: s.id(), name: 'cmpd ' + c, x: 640, y: 70 + i * 60, w: 300, h: 40, style: '', fill: solid('FFFFFF'), line: `<a:ln w="101600" cmpd="${c}">${solid('000000')}</a:ln>` }));
    });
    ['flat', 'rnd', 'sq'].forEach((c, i) => s.add(shape({ id: s.id(), name: 'cap ' + c, x: 680, y: 400 + i * 40, w: 220, h: 0, prst: 'line', style: '', line: `<a:ln w="190500" cap="${c}">${solid('548235')}</a:ln>` })));
    [['round', '<a:round/>'], ['bevel', '<a:bevel/>'], ['miter', '<a:miter lim="800000"/>']].forEach(([n, j], i) => s.add(shape({ id: s.id(), name: 'join ' + n, x: 1000 + (i % 2) * 120, y: 400 + Math.floor(i / 2) * 140, w: 90, h: 90, prst: 'triangle', style: '', fill: '<a:noFill/>', line: `<a:ln w="152400">${solid('7030A0')}${j}</a:ln>` })));
    s.add(shape({ id: s.id(), name: 'custDash', x: 40, y: 680, w: 380, h: 0, prst: 'line', style: '', line: `<a:ln w="38100">${solid('000000')}<a:custDash><a:ds d="400000" sp="100000"/><a:ds d="100000" sp="100000"/></a:custDash></a:ln>` }));
    out.push({ slide: s.index, feature: 'line/styles', compareLO: true, checks: [{ kind: 'pixel', x: 41, y: 70, color: '#2F5597', tol: 30 }, { kind: 'pixel', x: 790, y: 90, color: '#FFFFFF', tol: 30 }] });
  }

  // --- Arrow heads ------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Arrows' });
    label(s, 'Line ends: types and sizes, connectors');
    const types = ['triangle', 'stealth', 'diamond', 'oval', 'arrow'];
    const sizes = ['sm', 'med', 'lg'];
    types.forEach((t, i) => sizes.forEach((sz, j) => {
      s.add(shape({ id: s.id(), name: `${t} ${sz}`, x: 60 + j * 220, y: 80 + i * 70, w: 180, h: 0, prst: 'line', style: '', line: `<a:ln w="28575">${solid('000000')}<a:headEnd type="${t}" w="${sz}" len="${sz}"/><a:tailEnd type="${t}" w="${sz}" len="${sz}"/></a:ln>` }));
    }));
    const cx = (prst, x, y, w, h, extra = '') => s.add(shape({ cxn: true, id: s.id(), name: prst, x, y, w, h, prst, style: '', line: `<a:ln w="25400">${solid('C00000')}<a:tailEnd type="triangle"/>${extra}</a:ln>` }));
    cx('straightConnector1', 760, 80, 200, 120);
    cx('bentConnector3', 760, 240, 200, 120);
    cx('curvedConnector3', 760, 400, 200, 120);
    cx('bentConnector2', 1000, 80, 200, 120);
    cx('curvedConnector4', 1000, 240, 200, 120);
    s.add(shape({ cxn: true, id: s.id(), name: 'flipped', x: 1000, y: 420, w: 200, h: 120, prst: 'bentConnector3', flipH: true, style: '', line: `<a:ln w="25400">${solid('00B050')}<a:headEnd type="oval"/><a:tailEnd type="stealth"/></a:ln>` }));
    out.push({ slide: s.index, feature: 'line/arrows', compareLO: true, checks: [{ kind: 'pixel', x: 236, y: 80, color: '#000000', tol: 60 }] });
  }

  // --- Effects -----------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Effects' });
    label(s, 'Effects: outer/inner/perspective shadow, glow, soft edges, reflection, blur, theme effect');
    const fx = [
      ['outer shadow', '<a:effectLst><a:outerShdw blurRad="50800" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="60000"/></a:prstClr></a:outerShdw></a:effectLst>'],
      ['outer shadow big blur', '<a:effectLst><a:outerShdw blurRad="254000" dist="0" dir="0" algn="ctr" rotWithShape="0"><a:srgbClr val="C00000"><a:alpha val="80000"/></a:srgbClr></a:outerShdw></a:effectLst>'],
      ['perspective shadow', '<a:effectLst><a:outerShdw blurRad="76200" dist="12700" dir="18900000" sy="23000" kx="-1200000" algn="bl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>'],
      ['inner shadow', '<a:effectLst><a:innerShdw blurRad="114300" dist="38100" dir="2700000"><a:prstClr val="black"><a:alpha val="70000"/></a:prstClr></a:innerShdw></a:effectLst>'],
      ['glow', '<a:effectLst><a:glow rad="228600"><a:srgbClr val="FFC000"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst>'],
      ['soft edges', '<a:effectLst><a:softEdge rad="127000"/></a:effectLst>'],
      ['reflection', '<a:effectLst><a:reflection blurRad="6350" stA="52000" endA="300" endPos="35000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>'],
      ['blur', '<a:effectLst><a:blur rad="76200"/></a:effectLst>'],
      ['preset shadow', '<a:effectLst><a:prstShdw prst="shdw13" dist="53882" dir="2700000"><a:prstClr val="black"/></a:prstShdw></a:effectLst>'],
      ['rotated shadow', '<a:effectLst><a:outerShdw blurRad="38100" dist="76200" dir="5400000" algn="t" rotWithShape="0"><a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:outerShdw></a:effectLst>'],
    ];
    fx.forEach(([t, e], i) => {
      const col = i % 5, row = Math.floor(i / 5);
      const x = 60 + col * 240, y = 110 + row * 300;
      s.add(shape({ id: s.id(), name: t, x, y, w: 160, h: 110, prst: 'roundRect', fill: solid('4472C4'), effects: e, rot: t.startsWith('rotated') ? 45 : 0 }));
      s.add(shape({ id: s.id(), name: 'cap', x: x - 20, y: y + 160, w: 200, h: 30, style: '', fill: '<a:noFill/>', text: [{ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="1100"' }] }] }));
    });
    s.add(shape({ id: s.id(), name: 'theme effect', x: 1100, y: 600, w: 120, h: 60, style: '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent6"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>' }));
    out.push({
      slide: s.index, feature: 'effects', compareLO: true,
      checks: [{ kind: 'pixel', x: 140, y: 165, color: '#4472C4', tol: 20 }, { kind: 'lumBelow', x: 222, y: 165, max: 235 }, { kind: 'dom', selector: 'filter', min: 6 }],
    });
  }
  return out;
}

function darken(hex, k) {
  const n = parseInt(hex.slice(1), 16);
  const f = (v) => Math.round(v * k).toString(16).padStart(2, '0');
  return '#' + f((n >> 16) & 255) + f((n >> 8) & 255) + f(n & 255);
}

export { picture };

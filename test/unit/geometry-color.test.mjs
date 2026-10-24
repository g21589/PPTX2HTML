import test from 'node:test';
import assert from 'node:assert/strict';
import { presetNames, presetGeometry, evalFormula, segsToSvg, customGeometryDef, computeGeometry } from '../../src/geometry/geometry.js';
import { parseXml } from '../../src/util/xml.js';
import { resolveColorNode, hexToRgb } from '../../src/pptx/color.js';

test('all 187 preset geometries produce finite paths', () => {
  const names = presetNames();
  assert.equal(names.length, 187);
  for (const n of names) {
    for (const [w, h] of [[914400, 457200], [0, 914400], [1, 1]]) {
      const g = presetGeometry(n, w, h, {}, 1 / 9525);
      const d = g.paths.map((p) => segsToSvg(p.segs)).join('');
      assert.ok(d.length > 0, n);
      assert.ok(!/NaN|Infinity/.test(d), n + ' ' + d.slice(0, 80));
    }
  }
});

test('guide formula operators', () => {
  const env = { w: 100, h: 50, a: 3 };
  assert.equal(evalFormula('*/ w 3 4', env), 75);
  assert.equal(evalFormula('+- w h 10', env), 140);
  assert.equal(evalFormula('+/ w h 3', env), 50);
  assert.equal(evalFormula('?: a w h', env), 100);
  assert.equal(evalFormula('pin 0 120 100', env), 100);
  assert.equal(evalFormula('max w h', env), 100);
  assert.equal(Math.round(evalFormula('at2 1 1', env)), 2700000);
  assert.equal(Math.round(evalFormula('cos 100 5400000', env)), 0);
  assert.equal(evalFormula('mod 3 4 0', env), 5);
  assert.equal(evalFormula('sqrt 16', env), 4);
});

test('ellipse arcs follow DrawingML visual angles', () => {
  const g = presetGeometry('ellipse', 914400, 457200, {}, 1 / 9525);
  const d = segsToSvg(g.paths[0].segs);
  assert.match(d, /^M0 24C/);
  assert.match(d, /96 24/);
});

test('adjust values override preset defaults', () => {
  const a = presetGeometry('roundRect', 952500, 952500, { adj: 50000 }, 1 / 9525);
  assert.equal(segsToSvg(a.paths[0].segs).slice(0, 7), 'M0 50C0');
});

test('custom geometry with path coordinate space', () => {
  const node = parseXml('<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pathLst><a:path w="2" h="2"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="2" y="1"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>');
  const g = computeGeometry(customGeometryDef(node), 952500, 952500, {}, 1 / 9525);
  assert.equal(segsToSvg(g.paths[0].segs), 'M0 0L100 50Z');
});

test('colour transforms', () => {
  const env = { scheme: { accent1: hexToRgb('4472C4'), accent2: hexToRgb('ED7D31'), lt1: hexToRgb('FFFFFF'), dk1: hexToRgb('000000') }, clrMap: { bg1: 'lt1', tx1: 'dk1', accent1: 'accent1', accent2: 'accent2' } };
  const c = (xml) => {
    const r = resolveColorNode(parseXml(xml.replace('<a:', '<a:').replace(/^<a:(\w+)/, '<a:$1 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')), env);
    return '#' + [r.r, r.g, r.b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase() + (r.a < 1 ? '@' + r.a : '');
  };
  assert.equal(c('<a:schemeClr val="accent2"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'), '#F4B183');
  assert.equal(c('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'), '#2F5597');
  assert.equal(c('<a:schemeClr val="tx1"/>'), '#000000');
  assert.equal(c('<a:srgbClr val="FF0000"><a:alpha val="40000"/></a:srgbClr>'), '#FF0000@0.4');
  assert.equal(c('<a:srgbClr val="00FFFF"><a:inv/></a:srgbClr>'), '#FF0000');
  assert.equal(c('<a:hslClr hue="7200000" sat="100000" lum="50000"/>'), '#00FF00');
  assert.equal(c('<a:prstClr val="orange"/>'), '#FFA500');
  assert.equal(c('<a:sysClr val="window" lastClr="FFFFFF"/>'), '#FFFFFF');
  assert.equal(c('<a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr>'), '#BCBCBC');
  assert.equal(c('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>'), '#BCBCBC');
});

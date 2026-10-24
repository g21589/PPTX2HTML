import test from 'node:test';
import assert from 'node:assert/strict';
import { parseXml } from '../../src/util/xml.js';
import { layoutDiagram, parseDataModel } from '../../src/smartart/engine.js';

const DGM = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

function data(items) {
  // doc -> items (each may have children)
  const pts = ['<dgm:pt modelId="doc" type="doc"/>'];
  const cxns = [];
  let n = 0;
  const add = (parent, list) => list.forEach((it, i) => {
    const id = 'n' + ++n;
    pts.push(`<dgm:pt modelId="${id}"><dgm:prSet/><dgm:t><a:bodyPr/><a:p><a:r><a:t>${it.t}</a:t></a:r></a:p></dgm:t></dgm:pt>`);
    pts.push(`<dgm:pt modelId="${id}p" type="parTrans"/><dgm:pt modelId="${id}s" type="sibTrans"/>`);
    cxns.push(`<dgm:cxn modelId="c${id}" srcId="${parent}" destId="${id}" srcOrd="${i}" destOrd="0" parTransId="${id}p" sibTransId="${id}s"/>`);
    if (it.ch) add(id, it.ch);
  });
  add('doc', items);
  return parseXml(`<dgm:dataModel ${DGM}><dgm:ptLst>${pts.join('')}</dgm:ptLst><dgm:cxnLst>${cxns.join('')}</dgm:cxnLst></dgm:dataModel>`);
}

// A block list: snake grid of rectangles, spacing from the sibling transitions.
const BLOCKS = parseXml(`<dgm:layoutDef ${DGM}><dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/></dgm:varLst>
<dgm:choose><dgm:if func="var" arg="dir" op="equ" val="norm"><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/></dgm:alg></dgm:if>
<dgm:else><dgm:alg type="snake"><dgm:param type="grDir" val="tR"/></dgm:alg></dgm:else></dgm:choose>
<dgm:shape/><dgm:presOf/>
<dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.5"/>
<dgm:constr type="w" for="ch" forName="sib" refType="w" refFor="ch" refForName="node" fact="0.1"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="sib"/><dgm:constr type="primFontSz" for="ch" forName="node" val="40"/></dgm:constrLst>
<dgm:forEach axis="ch" ptType="node"><dgm:layoutNode name="node" styleLbl="node1"><dgm:alg type="tx"/><dgm:shape type="rect"/><dgm:presOf axis="desOrSelf" ptType="node"/></dgm:layoutNode>
<dgm:forEach axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="sib"><dgm:alg type="sp"/><dgm:shape/><dgm:presOf/></dgm:layoutNode></dgm:forEach></dgm:forEach>
</dgm:layoutNode></dgm:layoutDef>`);

const COLORS = parseXml(`<dgm:colorsDef ${DGM}><dgm:styleLbl name="node1"><dgm:fillClrLst><a:srgbClr val="FF0000"/><a:srgbClr val="0000FF"/></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:srgbClr val="FFFFFF"/></dgm:linClrLst><dgm:txFillClrLst meth="repeat"><a:srgbClr val="FFFFFF"/></dgm:txFillClrLst></dgm:styleLbl></dgm:colorsDef>`);

const ctx = { color: (n) => { const v = n.attrs.val; return { r: parseInt(v.slice(0, 2), 16), g: parseInt(v.slice(2, 4), 16), b: parseInt(v.slice(4, 6), 16), a: 1 }; } };
const shapes = (tree) => tree.children.filter((c) => c.local === 'sp');
const box = (sp) => { const x = sp.first('a:xfrm'); return ['x', 'y'].map((k) => Number(x.child('a:off').attrs[k]) / 9525).concat(['cx', 'cy'].map((k) => Number(x.child('a:ext').attrs[k]) / 9525)); };

test('data model tree with transitions', () => {
  const dm = parseDataModel(data([{ t: 'a', ch: [{ t: 'a1' }] }, { t: 'b' }]));
  assert.equal(dm.doc.children.length, 2);
  assert.equal(dm.doc.children[0].children[0].pt.t.textContent(), 'a1');
  assert.equal(dm.doc.children[1].sibTrans.type, 'sibTrans');
});

test('snake layout: grid, spacing, text, spanned colours', () => {
  const tree = layoutDiagram({ data: data([{ t: 'a' }, { t: 'b' }, { t: 'c' }, { t: 'd' }]), layout: BLOCKS, colors: COLORS }, 600, 400, ctx);
  const sp = shapes(tree);
  assert.equal(sp.length, 4);
  const b = sp.map(box);
  // Two columns of 2:1 boxes separated by 10% of their width.
  assert.ok(Math.abs(b[0][2] - 2 * b[0][3]) < 1);
  assert.ok(Math.abs(b[1][0] - (b[0][0] + b[0][2] * 1.1)) < 1, 'column gap');
  assert.ok(Math.abs(b[2][1] - (b[0][1] + b[0][3] + b[0][2] * 0.1)) < 1, 'row gap');
  assert.equal(sp[0].first('a:t').text, 'a');
  // Colours span red -> blue over the hue wheel; first and last are exact.
  const fill = (s) => s.first('dsp:style').child('a:fillRef').child('a:srgbClr').attrs.val;
  assert.equal(fill(sp[0]), 'FF0000');
  assert.equal(fill(sp[3]), '0000FF');
  assert.notEqual(fill(sp[1]), fill(sp[2]));
});

test('choose / var: reversed direction mirrors the grid', () => {
  const dm = data([{ t: 'a' }, { t: 'b' }]);
  // presLayoutVars on the diagram's presentation point switch dir to rev.
  const ptLst = dm.first('dgm:ptLst');
  const pres = parseXml(`<dgm:pt ${DGM} modelId="p0" type="pres"><dgm:prSet presAssocID="doc" presName="diagram"><dgm:presLayoutVars><dgm:dir val="rev"/></dgm:presLayoutVars></dgm:prSet></dgm:pt>`);
  pres.parent = ptLst;
  ptLst.children.push(pres);
  const sp = shapes(layoutDiagram({ data: dm, layout: BLOCKS, colors: COLORS }, 900, 200, ctx));
  assert.ok(box(sp[0])[0] > box(sp[1])[0], 'first node on the right');
});

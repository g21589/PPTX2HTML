/**
 * Tables (a:tbl) and table styles.
 */
import { emu2px } from '../util/units.js';
import { parseFill, parseFillNode, parseLine, finalizeLine, mergeLines, refColor } from './fill.js';
import { colorChild } from './color.js';
import { parseTextBody } from './text.js';
import { builtinTableStyle } from './tablestyles.js';
import { parseXml } from '../util/xml.js';

const PARTS = ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'seCell', 'swCell', 'firstRow', 'neCell', 'nwCell'];

/** Find the table style definition node for an id. */
function findStyle(doc, id, tblPr) {
  const inline = tblPr && tblPr.child('a:tableStyle');
  if (inline) return inline;
  if (!id) return null;
  if (!doc.tableStyles) {
    doc.tableStyles = new Map();
    const pres = doc.mainPart;
    const r = doc.pkg.relByType(pres, 'tableStyles');
    const root = r ? doc.pkg.xml(r.partName) : null;
    if (root) {
      doc.tableStylesDef = root.attr('def', '');
      for (const s of root.all('a:tblStyle')) doc.tableStyles.set(s.attr('styleId', ''), s);
    }
  }
  if (doc.tableStyles.has(id)) return doc.tableStyles.get(id);
  const xml = builtinTableStyle(id);
  if (xml) {
    const node = parseXml(xml);
    doc.tableStyles.set(id, node);
    return node;
  }
  return null;
}

/** Parse a style part (a:wholeTbl etc.) into fill/border/text info. */
function stylePart(node, ctx) {
  if (!node) return null;
  const out = { fill: undefined, borders: {}, text: null };
  const tcStyle = node.child('a:tcStyle');
  if (tcStyle) {
    const fillN = tcStyle.child('a:fill');
    if (fillN) out.fill = parseFill(fillN, ctx);
    const fillRef = tcStyle.child('a:fillRef');
    if (fillRef && out.fill === undefined) {
      const idx = fillRef.int('idx', 0);
      const phClr = refColor(fillRef, ctx);
      const n = idx >= 1000 ? ctx.theme.bgFillStyles[idx - 1001] : ctx.theme.fillStyles[idx - 1];
      out.fill = n ? parseFillNode(n, ctx.with({ phClr })) : (phClr ? { type: 'solid', color: phClr } : undefined);
    }
    const bdr = tcStyle.child('a:tcBdr');
    if (bdr) {
      for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV', 'tl2br', 'tr2bl']) {
        const b = bdr.child('a:' + side);
        if (!b) continue;
        const ln = b.child('a:ln');
        if (ln) out.borders[side] = parseLine(ln, ctx);
        else {
          const lr = b.child('a:lnRef');
          if (lr) {
            const idx = lr.int('idx', 0);
            const phClr = refColor(lr, ctx);
            const n = ctx.theme.lineStyles[idx - 1];
            out.borders[side] = n ? parseLine(n, ctx.with({ phClr })) : null;
          }
        }
      }
    }
  }
  const tx = node.child('a:tcTxStyle');
  if (tx) {
    const t = {};
    if (tx.attrs.b === 'on') t.b = '1';
    if (tx.attrs.b === 'off') t.b = '0';
    if (tx.attrs.i === 'on') t.i = '1';
    if (tx.attrs.i === 'off') t.i = '0';
    const font = tx.child('a:font');
    if (font) {
      const l = font.child('a:latin');
      if (l) t.latin = l.attrs.typeface;
      const ea = font.child('a:ea');
      if (ea) t.ea = ea.attrs.typeface;
    }
    const fr = tx.child('a:fontRef');
    if (fr) {
      const idx = fr.attr('idx', 'minor');
      t.latin = idx === 'major' ? '+mj-lt' : '+mn-lt';
      t.ea = idx === 'major' ? '+mj-ea' : '+mn-ea';
    }
    const cn = colorChild(tx) || (fr && colorChild(fr));
    if (cn) t.fill = { name: 'a:solidFill', local: 'solidFill', attrs: {}, children: [cn], child() { return null; }, all() { return []; } };
    out.text = t;
  }
  return out;
}

function cellMargins(tcPr) {
  const g = (k, d) => emu2px(tcPr ? tcPr.num(k, d) : d);
  return { l: g('marL', 91440), r: g('marR', 91440), t: g('marT', 45720), b: g('marB', 45720) };
}

/**
 * @param {import('../util/xml.js').XNode} tbl
 * @param {import('./context.js').Ctx} ctx
 */
export function parseTable(tbl, ctx, doc, el, baseChain = []) {
  const tblPr = tbl.child('a:tblPr');
  const flags = {
    firstRow: tblPr ? tblPr.bool('firstRow', false) : false,
    firstCol: tblPr ? tblPr.bool('firstCol', false) : false,
    lastRow: tblPr ? tblPr.bool('lastRow', false) : false,
    lastCol: tblPr ? tblPr.bool('lastCol', false) : false,
    bandRow: tblPr ? tblPr.bool('bandRow', false) : false,
    bandCol: tblPr ? tblPr.bool('bandCol', false) : false,
    rtl: tblPr ? tblPr.bool('rtl', false) : false,
  };
  const styleIdNode = tblPr && tblPr.child('a:tableStyleId');
  const styleId = styleIdNode ? styleIdNode.text.trim() : null;
  const styleNode = findStyle(doc, styleId, tblPr);
  const parts = {};
  if (styleNode) for (const p of PARTS) parts[p] = stylePart(styleNode.child('a:' + p), ctx);
  else if (styleId) ctx.note('missing', 'table style ' + styleId);
  // Table-level fill (behind cells).
  const tableFill = tblPr ? parseFill(tblPr, ctx) : undefined;
  const bgNode = styleNode && styleNode.child('a:tblBg');
  let tableBg = tableFill;
  if (tableBg === undefined && bgNode) {
    tableBg = parseFill(bgNode, ctx);
    const fr = bgNode.child('a:fillRef');
    if (tableBg === undefined && fr) {
      const phClr = refColor(fr, ctx);
      const n = ctx.theme.fillStyles[fr.int('idx', 0) - 1];
      tableBg = n ? parseFillNode(n, ctx.with({ phClr })) : undefined;
    }
  }

  const grid = tbl.child('a:tblGrid');
  const cols = grid ? grid.all('a:gridCol').map((c) => emu2px(c.num('w', 0))) : [];
  const trs = tbl.all('a:tr');
  const nRows = trs.length;
  const nCols = cols.length;

  const rows = [];
  for (let r = 0; r < nRows; r++) {
    const tr = trs[r];
    const row = { h: emu2px(tr.num('h', 0)), cells: [] };
    const tcs = tr.all('a:tc');
    for (let c = 0; c < tcs.length; c++) {
      const tc = tcs[c];
      const tcPr = tc.child('a:tcPr');
      const cell = {
        row: r,
        col: c,
        gridSpan: tc.int('gridSpan', 1),
        rowSpan: tc.int('rowSpan', 1),
        hMerge: tc.bool('hMerge', false),
        vMerge: tc.bool('vMerge', false),
        margins: cellMargins(tcPr),
        anchor: tcPr ? tcPr.attr('anchor', 't') : 't',
        anchorCtr: tcPr ? tcPr.bool('anchorCtr', false) : false,
        vert: tcPr ? tcPr.attr('vert', 'horz') : 'horz',
        fill: undefined,
        borders: {},
        text: null,
      };
      // Which style parts apply (low -> high priority).
      const applied = [];
      const isFirstRow = flags.firstRow && r === 0;
      const isLastRow = flags.lastRow && r === nRows - 1;
      const isFirstCol = flags.firstCol && c === 0;
      const isLastCol = flags.lastCol && c === nCols - 1;
      applied.push('wholeTbl');
      if (flags.bandRow && !isFirstRow && !isLastRow) {
        const idx = r - (flags.firstRow ? 1 : 0);
        applied.push(idx % 2 === 0 ? 'band1H' : 'band2H');
      }
      if (flags.bandCol && !isFirstCol && !isLastCol) {
        const idx = c - (flags.firstCol ? 1 : 0);
        applied.push(idx % 2 === 0 ? 'band1V' : 'band2V');
      }
      if (isLastCol) applied.push('lastCol');
      if (isFirstCol) applied.push('firstCol');
      if (isLastRow) applied.push('lastRow');
      if (isFirstRow) applied.push('firstRow');
      if (isLastRow && isLastCol) applied.push('seCell');
      if (isLastRow && isFirstCol) applied.push('swCell');
      if (isFirstRow && isLastCol) applied.push('neCell');
      if (isFirstRow && isFirstCol) applied.push('nwCell');

      const lastR = r + cell.rowSpan - 1 >= nRows - 1;
      const lastC = c + cell.gridSpan - 1 >= nCols - 1;
      let txBag = {};
      for (const name of applied) {
        const p = parts[name];
        if (!p) continue;
        if (p.fill !== undefined) cell.fill = p.fill;
        const b = p.borders;
        // Outer borders of the style region and inside borders between cells.
        const regionTop = name === 'lastRow' ? true : r === 0;
        const regionBottom = name === 'firstRow' ? true : lastR;
        const regionLeft = name === 'lastCol' ? true : c === 0;
        const regionRight = name === 'firstCol' ? true : lastC;
        const pick = (outer, inner, isOuter) => (isOuter ? b[outer] : b[inner]);
        const top = pick('top', 'insideH', regionTop);
        const bottom = pick('bottom', 'insideH', regionBottom);
        const left = pick('left', 'insideV', regionLeft);
        const right = pick('right', 'insideV', regionRight);
        if (top !== undefined) cell.borders.t = top;
        if (bottom !== undefined) cell.borders.b = bottom;
        if (left !== undefined) cell.borders.l = left;
        if (right !== undefined) cell.borders.r = right;
        if (b.tl2br !== undefined) cell.borders.tl2br = b.tl2br;
        if (b.tr2bl !== undefined) cell.borders.bl2tr = b.tr2bl;
        if (p.text) txBag = { ...txBag, ...p.text };
      }
      // Cell overrides.
      if (tcPr) {
        const f = parseFill(tcPr, ctx);
        if (f !== undefined) cell.fill = f;
        const map = { lnL: 'l', lnR: 'r', lnT: 't', lnB: 'b', lnTlToBr: 'tl2br', lnBlToTr: 'bl2tr' };
        for (const [k, side] of Object.entries(map)) {
          const ln = tcPr.child('a:' + k);
          if (ln) cell.borders[side] = mergeLines(cell.borders[side], parseLine(ln, ctx));
        }
      }
      for (const k of Object.keys(cell.borders)) cell.borders[k] = finalizeLine(cell.borders[k]);
      if (cell.fill === undefined) cell.fill = { type: 'none' };
      const txBody = tc.child('a:txBody');
      if (txBody) {
        const styleLayer = { def: { defRPr: txBag }, lvls: [] };
        cell.text = parseTextBody(txBody, ctx, {
          chain: [...baseChain, styleLayer],
          bodyChain: [],
          slideNumber: el ? el.slideNumber : undefined,
        });
        // Cell properties define insets/anchor instead of bodyPr.
        Object.assign(cell.text.body, {
          lIns: cell.margins.l, rIns: cell.margins.r, tIns: cell.margins.t, bIns: cell.margins.b,
          anchor: cell.anchor, anchorCtr: cell.anchorCtr, vert: cell.vert, wrap: 'square',
        });
      }
      row.cells.push(cell);
    }
    rows.push(row);
  }
  return { cols, rows, flags, background: tableBg || null, styleId };
}

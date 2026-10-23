/**
 * Table rendering (HTML table).
 */
import { r2 } from '../util/units.js';
import { css } from '../pptx/color.js';
import { cssBackground, fillColor } from './paint.js';
import { paragraphsHtml } from './text.js';

const DASH_CSS = { solid: 'solid', dot: 'dotted', sysDot: 'dotted', dash: 'dashed', sysDash: 'dashed', lgDash: 'dashed', dashDot: 'dashed', lgDashDot: 'dashed', lgDashDotDot: 'dashed', sysDashDot: 'dashed', sysDashDotDot: 'dashed' };

function borderCss(line) {
  if (!line) return 'none';
  const c = fillColor(line.fill);
  if (!c) return 'none';
  const style = line.cmpd === 'dbl' ? 'double' : DASH_CSS[line.dash] || 'solid';
  const w = line.cmpd === 'dbl' ? Math.max(3, line.w) : line.w;
  return `${r2(w)}px ${style} ${css(c)}`;
}

const VALIGN = { t: 'top', ctr: 'middle', b: 'bottom', just: 'middle', dist: 'middle' };

export function tableHtml(el, rc) {
  const t = el.table;
  const total = t.cols.reduce((a, b) => a + b, 0);
  let html = `<table class="pp-table" style="width:${r2(total)}px${t.flags.rtl ? ';direction:rtl' : ''}"><colgroup>`;
  for (const w of t.cols) html += `<col style="width:${r2(w)}px">`;
  html += '</colgroup><tbody>';
  if (t.background) {
    const bg = cssBackground(t.background, rc);
    if (bg) html = html.replace('class="pp-table" style="', `class="pp-table" style="background:${bg};`);
  }
  for (const row of t.rows) {
    html += `<tr style="height:${r2(row.h)}px">`;
    for (const cell of row.cells) {
      if (cell.hMerge || cell.vMerge) continue;
      let s = '';
      const bg = cssBackground(cell.fill, rc);
      if (bg) s += `background:${bg};`;
      s += `border-left:${borderCss(cell.borders.l)};border-right:${borderCss(cell.borders.r)};border-top:${borderCss(cell.borders.t)};border-bottom:${borderCss(cell.borders.b)};`;
      s += `padding:${r2(cell.margins.t)}px ${r2(cell.margins.r)}px ${r2(cell.margins.b)}px ${r2(cell.margins.l)}px;`;
      s += `vertical-align:${VALIGN[cell.anchor] || 'top'};`;
      if (cell.vert && cell.vert !== 'horz') s += `writing-mode:vertical-rl;${cell.vert === 'vert270' ? 'transform:rotate(180deg);' : ''}`;
      const span = (cell.gridSpan > 1 ? ` colspan="${cell.gridSpan}"` : '') + (cell.rowSpan > 1 ? ` rowspan="${cell.rowSpan}"` : '');
      let inner = cell.text ? paragraphsHtml(cell.text, rc) : '';
      const diag = [];
      if (cell.borders.tl2br) diag.push(['0', '0', '100%', '100%', cell.borders.tl2br]);
      if (cell.borders.bl2tr) diag.push(['0', '100%', '100%', '0', cell.borders.bl2tr]);
      if (diag.length) {
        const lines = diag.map(([x1, y1, x2, y2, ln]) => `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${css(fillColor(ln.fill))}" stroke-width="${r2(ln.w)}"/>`).join('');
        inner = `<svg class="pp-cell-diag" aria-hidden="true">${lines}</svg>` + inner;
        s += 'position:relative;';
      }
      html += `<td${span} style="${s}">${inner}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table>';
  return html;
}

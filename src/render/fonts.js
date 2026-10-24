/**
 * CSS font stacks with metric compatible substitutes for common Office fonts.
 */

const SUBSTITUTES = {
  'calibri': ['Carlito'],
  'calibri light': ['Calibri', 'Carlito'],
  'cambria': ['Caladea'],
  'cambria math': ['Cambria Math', 'STIX Two Math', 'Caladea'],
  'arial': ['Liberation Sans', 'Arimo', 'Helvetica'],
  'helvetica': ['Arial', 'Liberation Sans'],
  'arial narrow': ['Liberation Sans Narrow'],
  'arial black': ['Arial', 'Liberation Sans'],
  'times new roman': ['Liberation Serif', 'Tinos', 'Times'],
  'times': ['Times New Roman', 'Liberation Serif'],
  'courier new': ['Liberation Mono', 'Cousine', 'Courier'],
  'courier': ['Courier New', 'Liberation Mono'],
  'verdana': ['DejaVu Sans'],
  'tahoma': ['DejaVu Sans'],
  'segoe ui': ['Segoe UI', 'system-ui', 'Noto Sans', 'DejaVu Sans'],
  'segoe ui light': ['Segoe UI', 'system-ui'],
  'segoe ui semibold': ['Segoe UI', 'system-ui'],
  'century gothic': ['URW Gothic', 'Avant Garde'],
  'georgia': ['Gelasio'],
  'garamond': ['EB Garamond'],
  'consolas': ['Inconsolata', 'DejaVu Sans Mono'],
  'lucida console': ['DejaVu Sans Mono'],
  'gill sans mt': ['Gill Sans'],
  'trebuchet ms': ['DejaVu Sans'],
  'franklin gothic book': ['Libre Franklin'],
  'aptos': ['Aptos', 'Carlito'],
  'aptos display': ['Aptos', 'Carlito'],
  'symbol': ['OpenSymbol'],
  'wingdings': ['OpenSymbol'],
  // CJK.
  '微軟正黑體': ['Microsoft JhengHei', 'PingFang TC', 'Noto Sans CJK TC', 'WenQuanYi Zen Hei'],
  'microsoft jhenghei': ['微軟正黑體', 'PingFang TC', 'Noto Sans CJK TC', 'WenQuanYi Zen Hei'],
  '微软雅黑': ['Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'WenQuanYi Zen Hei'],
  'microsoft yahei': ['微软雅黑', 'PingFang SC', 'Noto Sans CJK SC', 'WenQuanYi Zen Hei'],
  '新細明體': ['PMingLiU', 'MingLiU', 'Noto Serif CJK TC', 'WenQuanYi Zen Hei'],
  'pmingliu': ['新細明體', 'Noto Serif CJK TC', 'WenQuanYi Zen Hei'],
  '標楷體': ['DFKai-SB', 'BiauKai', 'Noto Serif CJK TC'],
  '宋体': ['SimSun', 'Songti SC', 'Noto Serif CJK SC'],
  'simsun': ['宋体', 'Songti SC', 'Noto Serif CJK SC'],
  '黑体': ['SimHei', 'Heiti SC', 'Noto Sans CJK SC'],
  'ｍｓ ゴシック': ['MS Gothic', 'IPAGothic', 'Noto Sans CJK JP'],
  'ms gothic': ['ＭＳ ゴシック', 'IPAGothic', 'Noto Sans CJK JP'],
  'ｍｓ ｐゴシック': ['MS PGothic', 'IPAPGothic', 'Noto Sans CJK JP'],
  'ms pgothic': ['ＭＳ Ｐゴシック', 'IPAPGothic', 'Noto Sans CJK JP'],
  'メイリオ': ['Meiryo', 'Hiragino Sans', 'Noto Sans CJK JP'],
  'meiryo': ['メイリオ', 'Hiragino Sans', 'Noto Sans CJK JP'],
  'yu gothic': ['游ゴシック', 'YuGothic', 'Noto Sans CJK JP'],
  '游ゴシック': ['Yu Gothic', 'YuGothic', 'Noto Sans CJK JP'],
  '맑은 고딕': ['Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans CJK KR'],
  'malgun gothic': ['맑은 고딕', 'Apple SD Gothic Neo', 'Noto Sans CJK KR'],
};

const SERIF = /serif|times|roman|cambria|georgia|garamond|book|mincho|明朝|明體|宋|song|batang|caladea|palatino|bodoni|didot/i;
const MONO = /mono|courier|consolas|console|code/i;
const CJK_GENERIC = ['PingFang TC', 'Noto Sans CJK TC', 'Microsoft JhengHei', 'WenQuanYi Zen Hei', 'IPAPGothic'];

function quote(name) {
  // Single quotes: stacks are embedded in double-quoted style attributes.
  return /^[a-zA-Z-]+$/.test(name) && !['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'].includes(name) ? name : "'" + name.replace(/['"\\<>&]/g, '') + "'";
}

const cache = new Map();

/** Build a CSS font-family value from latin / east asian / complex script typefaces. */
export function fontStack(latin, ea, cs, opts = {}) {
  const key = latin + '|' + ea + '|' + cs + '|' + (opts.symbol ? 1 : 0);
  if (cache.has(key)) return cache.get(key);
  const names = [];
  const add = (n) => {
    if (!n) return;
    const t = n.trim();
    if (!t || names.some((x) => x.toLowerCase() === t.toLowerCase())) return;
    names.push(t);
  };
  const withSubs = (n) => {
    if (!n) return;
    add(n);
    for (const s of SUBSTITUTES[n.trim().toLowerCase()] || []) add(s);
  };
  if (!opts.symbol) withSubs(latin);
  withSubs(ea);
  withSubs(cs);
  const generic = MONO.test(latin || '') ? 'monospace' : SERIF.test(latin || '') ? 'serif' : 'sans-serif';
  if (generic === 'sans-serif') { add('Carlito'); add('Liberation Sans'); }
  for (const n of CJK_GENERIC) add(n);
  const out = names.map(quote).join(',') + (names.length ? ',' : '') + generic;
  cache.set(key, out);
  return out;
}

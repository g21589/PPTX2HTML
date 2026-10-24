/**
 * Theme part: colour scheme, font scheme and formatting scheme.
 */
import { resolveColorNode, COLOR_NODES, hexToRgb } from './color.js';

const SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

export const DEFAULT_SCHEME = {
  dk1: hexToRgb('000000'), lt1: hexToRgb('FFFFFF'), dk2: hexToRgb('44546A'), lt2: hexToRgb('E7E6E6'),
  accent1: hexToRgb('4472C4'), accent2: hexToRgb('ED7D31'), accent3: hexToRgb('A5A5A5'), accent4: hexToRgb('FFC000'),
  accent5: hexToRgb('5B9BD5'), accent6: hexToRgb('70AD47'), hlink: hexToRgb('0563C1'), folHlink: hexToRgb('954F72'),
};

function parseFontCollection(node) {
  const out = { latin: '', ea: '', cs: '', scripts: {} };
  if (!node) return out;
  for (const c of node.children) {
    if (c.name === 'a:latin') out.latin = c.attrs.typeface || '';
    else if (c.name === 'a:ea') out.ea = c.attrs.typeface || '';
    else if (c.name === 'a:cs') out.cs = c.attrs.typeface || '';
    else if (c.name === 'a:font' && c.attrs.script) out.scripts[c.attrs.script] = c.attrs.typeface || '';
  }
  return out;
}

export function parseColorScheme(node, base = DEFAULT_SCHEME) {
  const scheme = { ...base };
  if (!node) return scheme;
  for (const slot of SLOTS) {
    const n = node.child('a:' + slot);
    if (!n) continue;
    const c = n.children.find((x) => COLOR_NODES.has(x.name));
    if (c) scheme[slot] = resolveColorNode(c, { scheme: {}, clrMap: {} });
  }
  return scheme;
}

export class Theme {
  /** @param {import('../util/xml.js').XNode|null} root a:theme element */
  constructor(root, partName) {
    this.partName = partName || null;
    this.name = root ? root.attr('name', '') : '';
    const el = root ? root.child('a:themeElements') : null;
    this.colorScheme = parseColorScheme(el && el.child('a:clrScheme'));
    const fs = el && el.child('a:fontScheme');
    this.majorFont = parseFontCollection(fs && fs.child('a:majorFont'));
    this.minorFont = parseFontCollection(fs && fs.child('a:minorFont'));
    if (!this.majorFont.latin) this.majorFont.latin = 'Calibri Light';
    if (!this.minorFont.latin) this.minorFont.latin = 'Calibri';
    const fmt = el && el.child('a:fmtScheme');
    const list = (name) => {
      const n = fmt && fmt.child(name);
      return n ? n.children.slice() : [];
    };
    this.fillStyles = list('a:fillStyleLst');
    this.lineStyles = list('a:lnStyleLst');
    this.effectStyles = list('a:effectStyleLst');
    this.bgFillStyles = list('a:bgFillStyleLst');
    this.objectDefaults = root ? root.child('a:objectDefaults') : null;
  }

  /** Resolve a theme font reference like "+mj-lt" or "+mn-ea". */
  font(ref, lang) {
    if (!ref || ref[0] !== '+') return ref;
    const coll = ref.startsWith('+mj') ? this.majorFont : this.minorFont;
    const kind = ref.slice(4);
    let name = kind === 'ea' ? coll.ea : kind === 'cs' ? coll.cs : coll.latin;
    if (!name && lang) {
      const script = scriptForLang(lang);
      if (script && coll.scripts[script]) name = coll.scripts[script];
    }
    return name || coll.latin;
  }

  /** Font for a theme font collection index ("major" / "minor"). */
  fontByIdx(idx, kind = 'latin') {
    const coll = idx === 'major' ? this.majorFont : idx === 'minor' ? this.minorFont : null;
    if (!coll) return null;
    return coll[kind] || coll.latin;
  }
}

function scriptForLang(lang) {
  const l = lang.toLowerCase();
  if (l.startsWith('ja')) return 'Jpan';
  if (l.startsWith('ko')) return 'Hang';
  if (l === 'zh-cn' || l === 'zh-sg' || l.startsWith('zh-hans')) return 'Hans';
  if (l.startsWith('zh')) return 'Hant';
  if (l.startsWith('ar')) return 'Arab';
  if (l.startsWith('he')) return 'Hebr';
  if (l.startsWith('th')) return 'Thai';
  if (l.startsWith('ru') || l.startsWith('uk') || l.startsWith('bg')) return 'Cyrl';
  if (l.startsWith('el')) return 'Grek';
  if (l.startsWith('hi')) return 'Deva';
  return null;
}

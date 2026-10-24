/**
 * Small, fast, dependency free XML parser producing a light DOM.
 *
 * Element and attribute names are canonicalised: whatever prefix a document
 * uses for a namespace, the node name is rewritten with the canonical prefix
 * registered in NS (e.g. "p:sp", "a:solidFill").  This lets the rest of the
 * code base address nodes by their well known names even when a producer
 * chose unusual prefixes or a default namespace.
 */

export const NS = {
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
  'http://purl.oclc.org/ooxml/presentationml/main': 'p',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://purl.oclc.org/ooxml/drawingml/main': 'a',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://purl.oclc.org/ooxml/drawingml/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas': 'lc',
  'http://schemas.openxmlformats.org/drawingml/2006/compatibility': 'com',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',
  'http://schemas.openxmlformats.org/package/2006/content-types': 'ct',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing': 'cdr',
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main': 'x',
  'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
  'http://schemas.microsoft.com/office/powerpoint/2012/main': 'p15',
  'http://schemas.microsoft.com/office/powerpoint/2015/09/main': 'p159',
  'http://schemas.microsoft.com/office/powerpoint/2015/10/main': 'p1510',
  'http://schemas.microsoft.com/office/powerpoint/2017/10/main': 'p1710',
  'http://schemas.microsoft.com/office/powerpoint/2018/8/main': 'p188',
  'http://schemas.microsoft.com/office/powerpoint/2019/12/main': 'p1912',
  'http://schemas.microsoft.com/office/powerpoint/2013/main/command': 'pc',
  'http://schemas.microsoft.com/office/powerpoint/2006/main': 'p2006',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.microsoft.com/office/drawing/2012/main': 'a15',
  'http://schemas.microsoft.com/office/drawing/2014/main': 'a16',
  'http://schemas.microsoft.com/office/drawing/2016/11/main': 'a1611',
  'http://schemas.microsoft.com/office/drawing/2017/decorative': 'adec',
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main': 'asvg',
  'http://schemas.microsoft.com/office/drawing/2017/model3d': 'am3d',
  'http://schemas.microsoft.com/office/drawing/2016/ink': 'aink',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2007/8/2/chart': 'c14',
  'http://schemas.microsoft.com/office/drawing/2012/chart': 'c15',
  'http://schemas.microsoft.com/office/drawing/2014/chart': 'c16',
  'http://schemas.microsoft.com/office/drawing/2017/03/chart': 'c16r3',
  'http://schemas.microsoft.com/office/thememl/2012/main': 'thm15',
  'http://schemas.microsoft.com/office/drawing/2013/main/command': 'ac',
  'http://www.w3.org/XML/1998/namespace': 'xml',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties': 'ep',
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
  'http://purl.org/dc/elements/1.1/': 'dc',
  'http://purl.org/dc/terms/': 'dcterms',
  'http://www.w3.org/2001/XMLSchema-instance': 'xsi',
  'http://www.w3.org/2003/InkML': 'inkml',
};

/** Prefixes that are always bound even if a document forgets to declare them. */
const IMPLICIT = { xml: 'http://www.w3.org/XML/1998/namespace' };

export class XNode {
  constructor(name, attrs, parent) {
    this.name = name;
    this.attrs = attrs;
    this.parent = parent;
    this.children = [];
    this.text = '';
  }

  get local() {
    const i = this.name.indexOf(':');
    return i < 0 ? this.name : this.name.slice(i + 1);
  }

  /** First direct child with the given name (or first child if no name). */
  child(name) {
    const ch = this.children;
    if (name === undefined) return ch[0] || null;
    for (let i = 0; i < ch.length; i++) if (ch[i].name === name) return ch[i];
    return null;
  }

  /** All direct children with the given name (or all children). */
  all(name) {
    if (name === undefined) return this.children;
    const out = [];
    const ch = this.children;
    for (let i = 0; i < ch.length; i++) if (ch[i].name === name) out.push(ch[i]);
    return out;
  }

  /** Walk a path of child names, returns null when any step is missing. */
  path(...names) {
    let n = this;
    for (const name of names) {
      n = n.child(name);
      if (!n) return null;
    }
    return n;
  }

  attr(name, def) {
    const v = this.attrs[name];
    return v === undefined ? def : v;
  }

  int(name, def) {
    const v = this.attrs[name];
    if (v === undefined || v === '') return def;
    const n = parseInt(v, 10);
    return Number.isNaN(n) ? def : n;
  }

  num(name, def) {
    const v = this.attrs[name];
    if (v === undefined || v === '') return def;
    const n = parseFloat(v);
    return Number.isNaN(n) ? def : n;
  }

  bool(name, def) {
    const v = this.attrs[name];
    if (v === undefined) return def;
    return v === '1' || v === 'true' || v === 'on';
  }

  /** Depth-first search for descendants with the given name. */
  find(name, out = []) {
    for (const c of this.children) {
      if (c.name === name) out.push(c);
      c.find(name, out);
    }
    return out;
  }

  /** First descendant with the given name. */
  first(name) {
    for (const c of this.children) {
      if (c.name === name) return c;
      const r = c.first(name);
      if (r) return r;
    }
    return null;
  }

  /** Concatenated text of the node and all descendants. */
  textContent() {
    let s = this.text;
    for (const c of this.children) s += c.textContent();
    return s;
  }
}

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeEntities(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, e) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m;
      return String.fromCodePoint(code);
    }
    const v = ENTITIES[e];
    return v === undefined ? m : v;
  });
}

/** OOXML escapes control characters as _xHHHH_ in some string values. */
export function decodeOoxmlEscapes(s) {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
}

class Scope {
  constructor(parent) {
    this.parent = parent;
    this.map = null; // prefix -> uri
  }
  lookup(prefix) {
    for (let s = this; s; s = s.parent) {
      if (s.map && prefix in s.map) return s.map[prefix];
    }
    return IMPLICIT[prefix];
  }
}

function canon(prefix, uri) {
  if (uri !== undefined) {
    const c = NS[uri];
    if (c !== undefined) return c;
  }
  return prefix;
}

/**
 * Parse an XML string. Returns the document root element (an XNode).
 * The parser is tolerant: malformed input yields a best-effort tree.
 */
export function parseXml(src) {
  if (src.charCodeAt(0) === 0xfeff) src = src.slice(1);
  const docNode = new XNode('#document', {}, null);
  let cur = docNode;
  let scope = new Scope(null);
  const len = src.length;
  let i = 0;

  while (i < len) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      if (cur !== docNode) cur.text += decodeEntities(src.slice(i));
      break;
    }
    if (lt > i && cur !== docNode) cur.text += decodeEntities(src.slice(i, lt));
    const c1 = src.charCodeAt(lt + 1);
    if (c1 === 33 /* ! */) {
      if (src.startsWith('<!--', lt)) {
        const e = src.indexOf('-->', lt + 4);
        i = e < 0 ? len : e + 3;
      } else if (src.startsWith('<![CDATA[', lt)) {
        const e = src.indexOf(']]>', lt + 9);
        const end = e < 0 ? len : e;
        if (cur !== docNode) cur.text += src.slice(lt + 9, end);
        i = e < 0 ? len : e + 3;
      } else {
        // DOCTYPE or other declaration - skip, honouring an internal subset.
        let depth = 0;
        let j = lt + 2;
        for (; j < len; j++) {
          const ch = src.charCodeAt(j);
          if (ch === 91) depth++;
          else if (ch === 93) depth--;
          else if (ch === 62 && depth <= 0) break;
        }
        i = j + 1;
      }
      continue;
    }
    if (c1 === 63 /* ? */) {
      const e = src.indexOf('?>', lt + 2);
      i = e < 0 ? len : e + 2;
      continue;
    }
    if (c1 === 47 /* / */) {
      const e = src.indexOf('>', lt + 2);
      i = e < 0 ? len : e + 1;
      if (cur.parent) {
        cur = cur.parent;
        scope = scope.parent || scope;
      }
      continue;
    }

    // Start tag.
    let j = lt + 1;
    while (j < len) {
      const ch = src.charCodeAt(j);
      if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 47 || ch === 62) break;
      j++;
    }
    const rawName = src.slice(lt + 1, j);
    const rawAttrs = [];
    let selfClose = false;
    // Attributes.
    while (j < len) {
      let ch = src.charCodeAt(j);
      while (ch === 32 || ch === 9 || ch === 10 || ch === 13) ch = src.charCodeAt(++j);
      if (ch === 62) { j++; break; }
      if (ch === 47) {
        selfClose = true;
        const e = src.indexOf('>', j);
        j = e < 0 ? len : e + 1;
        break;
      }
      const nameStart = j;
      while (j < len) {
        ch = src.charCodeAt(j);
        if (ch === 61 || ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 62 || ch === 47) break;
        j++;
      }
      const aname = src.slice(nameStart, j);
      while (ch === 32 || ch === 9 || ch === 10 || ch === 13) ch = src.charCodeAt(++j);
      if (ch !== 61) {
        if (aname) rawAttrs.push(aname, '');
        if (j === nameStart) j++;
        continue;
      }
      ch = src.charCodeAt(++j);
      while (ch === 32 || ch === 9 || ch === 10 || ch === 13) ch = src.charCodeAt(++j);
      let value;
      if (ch === 34 || ch === 39) {
        const q = ch === 34 ? '"' : "'";
        const e = src.indexOf(q, j + 1);
        const end = e < 0 ? len : e;
        value = src.slice(j + 1, end);
        j = end + 1;
      } else {
        const vs = j;
        while (j < len) {
          ch = src.charCodeAt(j);
          if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 62) break;
          j++;
        }
        value = src.slice(vs, j);
      }
      rawAttrs.push(aname, value);
    }
    i = j;

    // Namespace declarations.
    const sc = new Scope(scope);
    for (let k = 0; k < rawAttrs.length; k += 2) {
      const an = rawAttrs[k];
      if (an === 'xmlns') {
        (sc.map || (sc.map = {}))[''] = decodeEntities(rawAttrs[k + 1]);
      } else if (an.startsWith('xmlns:')) {
        (sc.map || (sc.map = {}))[an.slice(6)] = decodeEntities(rawAttrs[k + 1]);
      }
    }
    const nsScope = sc.map ? sc : scope;

    const colon = rawName.indexOf(':');
    let name;
    if (colon < 0) {
      const uri = nsScope.lookup('');
      const p = uri !== undefined ? NS[uri] : undefined;
      name = p ? p + ':' + rawName : rawName;
    } else {
      const prefix = rawName.slice(0, colon);
      name = canon(prefix, nsScope.lookup(prefix)) + rawName.slice(colon);
    }

    const attrs = {};
    for (let k = 0; k < rawAttrs.length; k += 2) {
      const an = rawAttrs[k];
      if (an === 'xmlns' || an.startsWith('xmlns:')) continue;
      const ac = an.indexOf(':');
      let key = an;
      if (ac > 0) {
        const prefix = an.slice(0, ac);
        key = canon(prefix, nsScope.lookup(prefix)) + an.slice(ac);
      }
      attrs[key] = decodeEntities(rawAttrs[k + 1]);
    }

    // mc:Choice@Requires lists namespace prefixes: canonicalise them too.
    if (name === 'mc:Choice' && attrs.Requires) {
      attrs.Requires = attrs.Requires.split(/\s+/).filter(Boolean)
        .map((p) => canon(p, nsScope.lookup(p))).join(' ');
    }

    const node = new XNode(name, attrs, cur);
    cur.children.push(node);
    if (!selfClose) {
      cur = node;
      scope = nsScope === scope ? new Scope(scope) : nsScope;
    }
  }

  return docNode.children.find((n) => n.name) || docNode;
}

/** Escape text for inclusion in XML/HTML text content. */
export function escapeXml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : c === '"' ? '&quot;' : '&#39;'
  ));
}

/**
 * Open Packaging Conventions: content types, relationships and part access.
 */
import { ZipArchive } from '../util/zip.js';
import { parseXml } from '../util/xml.js';

export const REL = {
  officeDocument: 'officeDocument',
  slide: 'slide',
  slideLayout: 'slideLayout',
  slideMaster: 'slideMaster',
  notesSlide: 'notesSlide',
  notesMaster: 'notesMaster',
  handoutMaster: 'handoutMaster',
  theme: 'theme',
  themeOverride: 'themeOverride',
  image: 'image',
  chart: 'chart',
  chartEx: 'chartEx',
  hyperlink: 'hyperlink',
  media: 'media',
  video: 'video',
  audio: 'audio',
  oleObject: 'oleObject',
  package: 'package',
  diagramData: 'diagramData',
  diagramLayout: 'diagramLayout',
  diagramQuickStyle: 'diagramQuickStyle',
  diagramColors: 'diagramColors',
  diagramDrawing: 'diagramDrawing',
  tableStyles: 'tableStyles',
  presProps: 'presProps',
  viewProps: 'viewProps',
  font: 'font',
  comments: 'comments',
  commentAuthors: 'commentAuthors',
  tags: 'tags',
  vmlDrawing: 'vmlDrawing',
  customXml: 'customXml',
  chartStyle: 'chartStyle',
  chartColorStyle: 'chartColorStyle',
  chartUserShapes: 'chartUserShapes',
  coreProperties: 'core-properties',
  extendedProperties: 'extended-properties',
  thumbnail: 'thumbnail',
};

/** Reduce a relationship type URI to its short name ("slideLayout", "image"...). */
export function relTypeName(type) {
  if (!type) return '';
  const i = type.lastIndexOf('/');
  return i < 0 ? type : type.slice(i + 1);
}

/** Directory part of a part name ("/ppt/slides/slide1.xml" -> "/ppt/slides/"). */
export function partDir(partName) {
  const i = partName.lastIndexOf('/');
  return i < 0 ? '/' : partName.slice(0, i + 1);
}

/** Resolve a relationship target relative to the source part. */
export function resolvePartName(base, target) {
  if (!target) return '';
  let t = target;
  try { t = decodeURIComponent(t); } catch (e) { /* keep raw */ }
  t = t.replace(/\\/g, '/');
  const hash = t.indexOf('#');
  if (hash >= 0) t = t.slice(0, hash);
  let parts;
  if (t.startsWith('/')) parts = t.split('/');
  else parts = (partDir(base) + t).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '' || p === '.') continue;
    if (p === '..') out.pop();
    else out.push(p);
  }
  return '/' + out.join('/');
}

export const MIME_BY_EXT = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', dib: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml',
  webp: 'image/webp', ico: 'image/x-icon', emf: 'image/x-emf', wmf: 'image/x-wmf', wdp: 'image/vnd.ms-photo',
  jxr: 'image/vnd.ms-photo', heic: 'image/heic', avif: 'image/avif',
  mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', ogv: 'video/ogg',
  avi: 'video/x-msvideo', wmv: 'video/x-ms-wmv', mpg: 'video/mpeg', mpeg: 'video/mpeg', mkv: 'video/x-matroska',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', wma: 'audio/x-ms-wma', ogg: 'audio/ogg',
  oga: 'audio/ogg', aac: 'audio/aac', mid: 'audio/midi', midi: 'audio/midi', aiff: 'audio/aiff', aif: 'audio/aiff',
  flac: 'audio/flac',
  xml: 'application/xml', rels: 'application/vnd.openxmlformats-package.relationships+xml',
  ttf: 'font/ttf', otf: 'font/otf', woff: 'font/woff', woff2: 'font/woff2', fntdata: 'application/x-fontdata',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bin: 'application/octet-stream',
};

export function extOf(name) {
  const i = name.lastIndexOf('.');
  return i < 0 ? '' : name.slice(i + 1).toLowerCase();
}

/** Sniff a media type from the first bytes. */
export function sniffMime(bytes) {
  if (!bytes || bytes.length < 4) return null;
  const b = bytes;
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8) return 'image/jpeg';
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'image/gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'image/bmp';
  if ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a) || (b[0] === 0x4d && b[1] === 0x4d && b[3] === 0x2a)) return 'image/tiff';
  if (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x00 && b[3] === 0x00 && b.length > 44 && b[40] === 0x20 && b[41] === 0x45 && b[42] === 0x4d && b[43] === 0x46) return 'image/x-emf';
  if ((b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) || (b[0] === 0x01 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00) || (b[0] === 0x02 && b[1] === 0x00 && b[2] === 0x09 && b[3] === 0x00)) return 'image/x-wmf';
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b.length > 12) {
    const t = String.fromCharCode(b[8], b[9], b[10], b[11]);
    if (t === 'WEBP') return 'image/webp';
    if (t === 'WAVE') return 'audio/wav';
    if (t === 'AVI ') return 'video/x-msvideo';
  }
  if (b.length > 12 && b[4] === 0x66 && b[5] === 0x74 && b[6] === 0x79 && b[7] === 0x70) {
    const brand = String.fromCharCode(b[8], b[9], b[10], b[11]);
    if (brand === 'qt  ') return 'video/quicktime';
    if (brand.startsWith('M4A')) return 'audio/mp4';
    if (brand === 'avif') return 'image/avif';
    if (brand === 'heic' || brand === 'mif1') return 'image/heic';
    return 'video/mp4';
  }
  if (b[0] === 0x1a && b[1] === 0x45 && b[2] === 0xdf && b[3] === 0xa3) return 'video/webm';
  if (b[0] === 0x49 && b[1] === 0x44 && b[2] === 0x33) return 'audio/mpeg';
  if (b[0] === 0xff && (b[1] & 0xe0) === 0xe0) return 'audio/mpeg';
  if (b[0] === 0x4f && b[1] === 0x67 && b[2] === 0x67 && b[3] === 0x53) return 'audio/ogg';
  if (b[0] === 0x30 && b[1] === 0x26 && b[2] === 0xb2 && b[3] === 0x75) return 'video/x-ms-wmv';
  if (b[0] === 0x49 && b[1] === 0x49 && b[2] === 0xbc) return 'image/vnd.ms-photo';
  // SVG / XML text.
  let s = '';
  for (let i = 0; i < Math.min(b.length, 256); i++) s += String.fromCharCode(b[i]);
  if (/<svg[\s>]/i.test(s) || (/^\s*<\?xml/.test(s) && /<svg/i.test(s))) return 'image/svg+xml';
  return null;
}

export class OpcPackage {
  /** @param {Uint8Array|ArrayBuffer} data */
  constructor(data) {
    this.zip = new ZipArchive(data);
    this._xml = new Map();
    this._rels = new Map();
    this.overrides = new Map();
    this.defaults = new Map();
    const ct = this.xml('/[Content_Types].xml');
    if (ct) {
      for (const n of ct.children) {
        if (n.local === 'Override') this.overrides.set((n.attrs.PartName || '').toLowerCase(), n.attrs.ContentType);
        else if (n.local === 'Default') this.defaults.set((n.attrs.Extension || '').toLowerCase(), n.attrs.ContentType);
      }
    }
  }

  has(partName) {
    return this.zip.has(partName);
  }

  bytes(partName) {
    return this.zip.read(partName);
  }

  text(partName) {
    return this.zip.text(partName);
  }

  /** Parsed XML of a part (cached), or null. */
  xml(partName) {
    if (this._xml.has(partName)) return this._xml.get(partName);
    const t = this.zip.text(partName);
    const doc = t === null ? null : parseXml(t);
    this._xml.set(partName, doc);
    return doc;
  }

  contentType(partName) {
    const o = this.overrides.get(partName.toLowerCase());
    if (o) return o;
    return this.defaults.get(extOf(partName)) || null;
  }

  /** Media type of a binary part, using content types, extension and magic. */
  mimeOf(partName) {
    const bytes = this.bytes(partName);
    const sniffed = sniffMime(bytes);
    if (sniffed) return sniffed;
    const ct = this.contentType(partName);
    if (ct && !/octet-stream/.test(ct)) return ct === 'image/jpg' ? 'image/jpeg' : ct;
    return MIME_BY_EXT[extOf(partName)] || 'application/octet-stream';
  }

  /** Relationship map for a part: id -> {id, type, target, external, partName}. */
  rels(partName) {
    if (this._rels.has(partName)) return this._rels.get(partName);
    const dir = partDir(partName);
    const file = partName.slice(dir.length);
    const relsName = dir + '_rels/' + file + '.rels';
    const doc = this.xml(relsName);
    const map = new Map();
    if (doc) {
      for (const n of doc.children) {
        if (n.local !== 'Relationship') continue;
        const external = n.attrs.TargetMode === 'External';
        const target = n.attrs.Target || '';
        map.set(n.attrs.Id, {
          id: n.attrs.Id,
          type: relTypeName(n.attrs.Type),
          typeUri: n.attrs.Type,
          target,
          external,
          partName: external ? null : resolvePartName(partName, target),
        });
      }
    }
    this._rels.set(partName, map);
    return map;
  }

  rel(partName, id) {
    return this.rels(partName).get(id) || null;
  }

  /** First relationship of a given short type. */
  relByType(partName, type) {
    for (const r of this.rels(partName).values()) if (r.type === type) return r;
    return null;
  }

  relsByType(partName, type) {
    const out = [];
    for (const r of this.rels(partName).values()) if (r.type === type) out.push(r);
    return out;
  }

  /** Main document part ("/ppt/presentation.xml"). */
  mainPart() {
    const r = this.relByType('/', REL.officeDocument);
    if (r && this.has(r.partName)) return r.partName;
    for (const [name, ct] of this.overrides) {
      if (/presentationml\.(presentation|slideshow|template)\.main\+xml|presentationml\.presentation\.macroEnabled\.main/.test(ct)) {
        return this.zip._entry(name) ? '/' + this.zip._entry(name).name : name;
      }
    }
    if (this.has('/ppt/presentation.xml')) return '/ppt/presentation.xml';
    return null;
  }
}

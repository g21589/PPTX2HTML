/**
 * Parsing context shared by the model builders.
 */
import { resolveColorNode } from './color.js';
import { imageSize } from '../util/image.js';

export class Diagnostics {
  constructor() {
    /** @type {Map<string, {kind: string, message: string, count: number, parts: Set<string>}>} */
    this.items = new Map();
  }

  add(kind, message, part) {
    const key = kind + '|' + message;
    let it = this.items.get(key);
    if (!it) {
      it = { kind, message, count: 0, parts: new Set() };
      this.items.set(key, it);
    }
    it.count++;
    if (part) it.parts.add(part);
  }

  list() {
    return [...this.items.values()].map((i) => ({ kind: i.kind, message: i.message, count: i.count, parts: [...i.parts] }));
  }

  /** Items of a kind ("unsupported", "approx", "error", "missing"). */
  ofKind(kind) {
    return this.list().filter((i) => i.kind === kind);
  }
}

/** Media registry: every binary part referenced by the deck. */
export class MediaRegistry {
  constructor(pkg) {
    this.pkg = pkg;
    /** @type {Map<string, object>} */
    this.items = new Map();
    this.externals = new Map();
  }

  /** Register an internal part, returns a media descriptor. */
  part(partName) {
    let m = this.items.get(partName);
    if (m) return m;
    if (!this.pkg.has(partName)) return null;
    const bytes = this.pkg.bytes(partName);
    const mime = this.pkg.mimeOf(partName);
    m = { id: partName, partName, mime, size: bytes.length, external: false };
    if (mime.startsWith('image/')) {
      const sz = imageSize(bytes, mime);
      if (sz) {
        m.width = sz.width;
        m.height = sz.height;
        m.dpiX = sz.dpiX || 96;
        m.dpiY = sz.dpiY || 96;
      }
    }
    this.items.set(partName, m);
    return m;
  }

  external(url, kind) {
    let m = this.externals.get(url);
    if (!m) {
      m = { id: url, url, external: true, mime: kind || null };
      this.externals.set(url, m);
    }
    return m;
  }

  bytes(id) {
    const m = this.items.get(id);
    return m ? this.pkg.bytes(m.partName) : null;
  }
}

export class Ctx {
  /**
   * @param {object} doc document level state ({pkg, media, diag, options, ...})
   * @param {string} part current part name (used to resolve relationships)
   */
  constructor(doc, part) {
    this.doc = doc;
    this.part = part;
    this.theme = null;
    this.scheme = null;
    this.clrMap = null;
    this.phClr = null;
    this.groupFill = null;
  }

  /** Shallow copy with overrides. */
  with(o) {
    const c = Object.create(Ctx.prototype);
    Object.assign(c, this, o);
    return c;
  }

  get pkg() {
    return this.doc.pkg;
  }

  color(node) {
    return resolveColorNode(node, this);
  }

  rel(id) {
    return id ? this.pkg.rel(this.part, id) : null;
  }

  /** Resolve an image relationship to a media descriptor. */
  image(id) {
    const r = this.rel(id);
    if (!r) {
      if (id) this.note('missing', 'image relationship ' + id + ' not found');
      return null;
    }
    if (r.external) return this.doc.media.external(r.target, 'image');
    const m = this.doc.media.part(r.partName);
    if (!m) this.note('missing', 'image part ' + r.partName + ' not found');
    return m;
  }

  /** Resolve any binary relationship (media, audio, video). */
  mediaRel(id, kind) {
    const r = this.rel(id);
    if (!r) return null;
    if (r.external) return this.doc.media.external(r.target, kind);
    return this.doc.media.part(r.partName);
  }

  note(kind, message) {
    this.doc.diag.add(kind, message, this.part);
  }
}

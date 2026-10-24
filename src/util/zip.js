/**
 * Minimal ZIP archive reader (store + deflate, ZIP64 aware).
 */
import { inflateRaw } from './inflate.js';

const CP437 = '\u0000☺☻♥♦♣♠•◘○◙♂♀♪♫☼►◄↕‼¶§▬↨↑↓→←∟↔▲▼ !"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~⌂ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';

const utf8 = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : null;

export function decodeUtf8(bytes) {
  if (utf8) return utf8.decode(bytes);
  // Fallback (very old environments).
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return decodeURIComponent(escape(s));
}

export function decodeText(bytes) {
  // Detect BOMs: UTF-16 LE/BE.
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  }
  return decodeUtf8(bytes);
}

function decodeName(bytes, isUtf8) {
  if (isUtf8) return decodeUtf8(bytes);
  let ascii = true;
  for (let i = 0; i < bytes.length; i++) if (bytes[i] > 127) { ascii = false; break; }
  if (ascii) return String.fromCharCode.apply(null, bytes);
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += CP437[bytes[i]];
  return s;
}

export class ZipArchive {
  /** @param {Uint8Array|ArrayBuffer} data */
  constructor(data) {
    this.data = data instanceof Uint8Array ? data : new Uint8Array(data);
    this.view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
    /** @type {Map<string, object>} */
    this.entries = new Map();
    this.lower = new Map();
    this._readDirectory();
  }

  _readDirectory() {
    const d = this.data;
    const v = this.view;
    const n = d.length;
    if (n < 22) throw new Error('Not a ZIP file (too small)');
    let eocd = -1;
    const stop = Math.max(0, n - 22 - 65535);
    for (let i = n - 22; i >= stop; i--) {
      if (d[i] === 0x50 && d[i + 1] === 0x4b && d[i + 2] === 0x05 && d[i + 3] === 0x06) { eocd = i; break; }
    }
    if (eocd < 0) {
      if (d[0] === 0xd0 && d[1] === 0xcf && d[2] === 0x11 && d[3] === 0xe0) {
        throw new Error('This is a legacy binary (OLE2) or encrypted Office file, not an OOXML package');
      }
      if (d[0] === 0x50 && d[1] === 0x4b && d[2] === 0x03 && d[3] === 0x04) {
        this._recoverLocalHeaders();
        if (this.entries.size) return;
      }
      throw new Error('Not a ZIP file (end of central directory not found)');
    }
    let count = v.getUint16(eocd + 10, true);
    let cdSize = v.getUint32(eocd + 12, true);
    let cdOff = v.getUint32(eocd + 16, true);
    // ZIP64 locator.
    if ((cdOff === 0xffffffff || count === 0xffff) && eocd >= 20 && v.getUint32(eocd - 20, true) === 0x07064b50) {
      const z64 = Number(v.getBigUint64(eocd - 12, true));
      if (v.getUint32(z64, true) === 0x06064b50) {
        count = Number(v.getBigUint64(z64 + 32, true));
        cdSize = Number(v.getBigUint64(z64 + 40, true));
        cdOff = Number(v.getBigUint64(z64 + 48, true));
      }
    }
    // Some writers produce archives with a prefix (self extracting etc.).
    let shift = 0;
    if (v.getUint32(cdOff, true) !== 0x02014b50) {
      const real = eocd - cdSize;
      if (real >= 0 && v.getUint32(real, true) === 0x02014b50) shift = real - cdOff;
    }
    let p = cdOff + shift;
    for (let k = 0; k < count && p + 46 <= n; k++) {
      if (v.getUint32(p, true) !== 0x02014b50) break;
      const flags = v.getUint16(p + 8, true);
      const method = v.getUint16(p + 10, true);
      let csize = v.getUint32(p + 20, true);
      let usize = v.getUint32(p + 24, true);
      const nameLen = v.getUint16(p + 28, true);
      const extraLen = v.getUint16(p + 30, true);
      const commentLen = v.getUint16(p + 32, true);
      let local = v.getUint32(p + 42, true);
      const name = decodeName(d.subarray(p + 46, p + 46 + nameLen), (flags & 0x800) !== 0);
      // ZIP64 extra field.
      let e = p + 46 + nameLen;
      const eEnd = e + extraLen;
      while (e + 4 <= eEnd) {
        const id = v.getUint16(e, true);
        const sz = v.getUint16(e + 2, true);
        if (id === 0x0001) {
          let q = e + 4;
          if (usize === 0xffffffff) { usize = Number(v.getBigUint64(q, true)); q += 8; }
          if (csize === 0xffffffff) { csize = Number(v.getBigUint64(q, true)); q += 8; }
          if (local === 0xffffffff) { local = Number(v.getBigUint64(q, true)); q += 8; }
        }
        e += 4 + sz;
      }
      const entry = { name, method, csize, usize, local: local + shift, flags, data: null };
      if (!name.endsWith('/')) {
        this.entries.set(name, entry);
        this.lower.set(name.toLowerCase(), entry);
      }
      p += 46 + nameLen + extraLen + commentLen;
    }
  }

  /** Rebuild the directory of a truncated archive by scanning local headers. */
  _recoverLocalHeaders() {
    const d = this.data;
    const v = this.view;
    const n = d.length;
    let p = 0;
    while (p + 30 <= n && v.getUint32(p, true) === 0x04034b50) {
      const flags = v.getUint16(p + 6, true);
      const method = v.getUint16(p + 8, true);
      let csize = v.getUint32(p + 18, true);
      const usize = v.getUint32(p + 22, true);
      const nameLen = v.getUint16(p + 26, true);
      const extraLen = v.getUint16(p + 28, true);
      const start = p + 30 + nameLen + extraLen;
      if (start > n) break;
      const name = decodeName(d.subarray(p + 30, p + 30 + nameLen), (flags & 0x800) !== 0);
      if ((flags & 8) || csize === 0 || start + csize > n) {
        // Unknown size: the next local header (or the end) bounds the data.
        let q = start;
        while (q + 4 <= n && !(d[q] === 0x50 && d[q + 1] === 0x4b && d[q + 2] === 0x03 && d[q + 3] === 0x04)) q++;
        csize = Math.min(q, n) - start;
      }
      const entry = { name, method, csize, usize, local: p, flags, data: null };
      if (!name.endsWith('/')) {
        this.entries.set(name, entry);
        this.lower.set(name.toLowerCase(), entry);
      }
      p = start + csize;
      if (flags & 8) {
        // Skip an optional data descriptor.
        if (p + 4 <= n && v.getUint32(p, true) === 0x08074b50) p += 16;
        else if (p + 12 <= n && v.getUint32(p + 12, true) !== 0x04034b50) p += 12;
      }
    }
  }

  _entry(name) {
    if (name.startsWith('/')) name = name.slice(1);
    return this.entries.get(name) || this.lower.get(name.toLowerCase()) || null;
  }

  has(name) {
    return this._entry(name) !== null;
  }

  names() {
    return [...this.entries.keys()];
  }

  /** @returns {Uint8Array|null} */
  read(name) {
    const e = this._entry(name);
    if (!e) return null;
    if (e.data) return e.data;
    const v = this.view;
    if (v.getUint32(e.local, true) !== 0x04034b50) throw new Error('Corrupt ZIP local header for ' + e.name);
    if (e.flags & 1) throw new Error('Encrypted ZIP entries are not supported: ' + e.name);
    const nameLen = v.getUint16(e.local + 26, true);
    const extraLen = v.getUint16(e.local + 28, true);
    const start = e.local + 30 + nameLen + extraLen;
    const raw = this.data.subarray(start, start + e.csize);
    let out;
    if (e.method === 0) out = raw;
    else if (e.method === 8) out = inflateRaw(raw, e.usize);
    else throw new Error('Unsupported ZIP compression method ' + e.method + ' for ' + e.name);
    e.data = out;
    return out;
  }

  text(name) {
    const b = this.read(name);
    return b ? decodeText(b) : null;
  }
}

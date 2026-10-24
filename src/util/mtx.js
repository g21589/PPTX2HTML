/**
 * MicroType Express (MTX) font decompression, the format of compressed
 * Embedded OpenType data.
 *
 * Implemented from the W3C MicroType Express submission: three LZCOMP
 * streams (LZ77 with adaptive Huffman coded symbols) carry a TrueType font in
 * Compact Table Format (CTF), its glyph push data and its glyph programs.
 * The CTF font is turned back into a regular TrueType font: 'cvt ' and
 * 'glyf' are decoded, 'loca' is rebuilt and the device metric tables
 * ('hdmx', 'VDMX'), which browsers do not need, are dropped.
 */

class BitReader {
  constructor(bytes) {
    this.b = bytes;
    this.p = 0;
    this.buf = 0;
    this.n = 0;
  }

  bit() {
    if (this.n === 0) {
      if (this.p >= this.b.length) throw new Error('MTX: unexpected end of data');
      this.buf = this.b[this.p++];
      this.n = 8;
    }
    this.n--;
    return (this.buf >> this.n) & 1;
  }

  read(bits) {
    let v = 0;
    for (let i = 0; i < bits; i++) v = v * 2 + this.bit();
    return v;
  }
}

/** Adaptive Huffman model over the symbols [0, range) (shared with the test encoder). */
export class AHuff {
  constructor(bits, range) {
    const n = 2 * range;
    this.bits = bits;
    this.up = new Int32Array(n);
    this.left = new Int32Array(n);
    this.right = new Int32Array(n);
    this.code = new Int32Array(n);
    this.weight = new Int32Array(n);
    this.index = new Int32Array(range);
    for (let i = 2; i < n; i++) {
      this.up[i] = i >> 1;
      this.weight[i] = 1;
    }
    for (let i = 1; i < range; i++) {
      this.left[i] = 2 * i;
      this.right[i] = 2 * i + 1;
    }
    for (let i = 0; i < range; i++) {
      this.code[i] = -1;
      this.code[range + i] = i;
      this.left[range + i] = this.right[range + i] = -1;
      this.index[i] = range + i;
    }
    for (let i = range - 1; i >= 1; i--) this.weight[i] = this.weight[2 * i] + this.weight[2 * i + 1];
    // Initial statistics favour the copy commands of the LZCOMP symbol set.
    if (range > 256 && range < 512) {
      this.update(this.index[256]);
      this.update(this.index[257]);
      for (let i = 0; i < 12; i++) this.update(this.index[range - 3]);
      for (let i = 0; i < 6; i++) this.update(this.index[range - 2]);
    } else {
      for (let j = 0; j < 2; j++) for (let i = 0; i < range; i++) this.update(this.index[i]);
    }
  }

  swap(a, b) {
    const { left, right, code, weight } = this;
    let t = left[a]; left[a] = left[b]; left[b] = t;
    t = right[a]; right[a] = right[b]; right[b] = t;
    t = code[a]; code[a] = code[b]; code[b] = t;
    t = weight[a]; weight[a] = weight[b]; weight[b] = t;
    for (const x of [a, b]) {
      if (code[x] < 0) this.up[left[x]] = this.up[right[x]] = x;
      else this.index[code[x]] = x;
    }
  }

  update(a) {
    const { up, weight } = this;
    for (; a !== 1; a = up[a]) {
      const w = weight[a];
      let b = a - 1;
      if (weight[b] === w) {
        do b--; while (weight[b] === w);
        b++;
        if (b > 1) {
          this.swap(a, b);
          a = b;
        }
      }
      weight[a] = w + 1;
    }
    weight[1]++;
  }

  read() {
    let a = 1;
    do a = this.bits.bit() ? this.right[a] : this.left[a];
    while (this.code[a] < 0);
    const sym = this.code[a];
    this.update(a);
    return sym;
  }
}

const PRELOAD = 2 * 32 * 96 + 4 * 256;

/**
 * Decompress one LZCOMP block.
 * @param {Uint8Array} bytes
 * @param {number} [version] MTX version (1 has no run-length flag)
 */
export function lzcompDecode(bytes, version = 3) {
  if (!bytes.length) return new Uint8Array(0);
  const bits = new BitReader(bytes);
  const runLength = version !== 1 && bits.bit() === 1;
  const dist = new AHuff(bits, 8);
  const len = new AHuff(bits, 8);
  const outLen = bits.read(24);
  let ranges = 1;
  while (2 ** (3 * ranges) < outLen) ranges++;
  const DUP2 = 256 + 8 * ranges;
  const sym = new AHuff(bits, DUP2 + 3);
  // The history starts with a fixed dictionary so early bytes can be copies.
  const buf = new Uint8Array(PRELOAD + outLen);
  let pos = 0;
  for (let k = 0; k < 32; k++) for (let j = 0; j < 96; j++) { buf[pos++] = k; buf[pos++] = j; }
  for (let j = 0; j < 256; j++) for (let r = 0; r < 4; r++) buf[pos++] = j;
  const end = buf.length;
  while (pos < end) {
    const s = sym.read();
    if (s < 256) buf[pos++] = s;
    else if (s >= DUP2) {
      buf[pos] = buf[pos - 2 * (s - DUP2 + 1)];
      pos++;
    } else {
      let b = s - 256;
      const nDist = (b >> 3) + 1;
      b &= 7;
      let length = 0;
      for (;;) {
        length = length * 4 + (b & 3);
        if (!(b & 4)) break;
        b = len.read();
      }
      length += 2;
      let d = 0;
      for (let i = 0; i < nDist; i++) d = d * 8 + dist.read();
      d += 1;
      if (d >= 512) length++;
      const start = pos - d - length + 1;
      if (start < 0 || pos + length > end) throw new Error('MTX: copy out of range');
      for (let j = 0; j < length; j++) buf[pos++] = buf[start + j];
    }
  }
  const out = buf.subarray(PRELOAD);
  return runLength ? unRunLength(out) : out;
}

/** Undo the escape-byte run-length coding applied before LZCOMP. */
function unRunLength(data) {
  if (!data.length) return data;
  const esc = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = data[i];
    if (v !== esc) { out.push(v); continue; }
    const count = data[++i];
    if (count === 0) out.push(esc);
    else {
      const byte = data[++i];
      for (let k = 0; k < count; k++) out.push(byte);
    }
  }
  return Uint8Array.from(out);
}

class Reader {
  constructor(bytes, pos = 0) {
    this.b = bytes;
    this.p = pos;
  }

  u8() {
    if (this.p >= this.b.length) throw new Error('MTX: truncated font data');
    return this.b[this.p++];
  }

  u16() {
    return (this.u8() << 8) | this.u8();
  }

  i16() {
    const v = this.u16();
    return v & 0x8000 ? v - 0x10000 : v;
  }

  bytes(n) {
    if (this.p + n > this.b.length) throw new Error('MTX: truncated font data');
    const out = this.b.subarray(this.p, this.p + n);
    this.p += n;
    return out;
  }

  u255() {
    const c = this.u8();
    if (c === 253) return this.u16();
    if (c === 255) return 253 + this.u8();
    if (c === 254) return 506 + this.u8();
    return c;
  }

  s255() {
    let c = this.u8();
    if (c === 253) return this.i16();
    let sign = 1;
    if (c === 250) {
      sign = -1;
      c = this.u8();
    }
    const v = c === 255 ? 250 + this.u8() : c === 254 ? 500 + this.u8() : c;
    return v * sign;
  }
}

class Writer {
  constructor() {
    this.a = [];
  }

  u8(v) {
    this.a.push(v & 0xff);
  }

  u16(v) {
    this.a.push((v >> 8) & 0xff, v & 0xff);
  }

  u32(v) {
    this.a.push((v >>> 24) & 0xff, (v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff);
  }

  bytes(b) {
    for (const v of b) this.a.push(v);
  }

  pad(n) {
    while (this.a.length % n) this.a.push(0);
  }

  get length() {
    return this.a.length;
  }
}

/** cvt values: deltas against the previous entry, one to three bytes each. */
function decodeCvt(r) {
  const n = r.u16();
  const w = new Writer();
  let prev = 0;
  for (let i = 0; i < n; i++) {
    const c = r.u8();
    let v;
    if (c < 238) v = c;
    else if (c >= 248) v = r.u8() + (c - 247) * 238;
    else if (c >= 239) v = -(r.u8() + (c - 239) * 238);
    else v = r.i16();
    prev = (prev + v) & 0xffff;
    w.u16(prev);
  }
  return w.a;
}

const sign = (flag, v) => (flag & 1 ? v : -v);

/** Point triplet: flag byte index (low 7 bits) plus 1-4 data bytes. */
function triplet(f, r) {
  if (f < 10) return [0, sign(f, ((f & 14) << 7) + r.u8())];
  if (f < 20) return [sign(f, (((f - 10) & 14) << 7) + r.u8()), 0];
  if (f < 84) {
    const b = f - 20, c = r.u8();
    return [sign(f, 1 + (b & 0x30) + (c >> 4)), sign(f >> 1, 1 + ((b & 0x0c) << 2) + (c & 0x0f))];
  }
  if (f < 120) {
    const b = f - 84;
    const x = r.u8(), y = r.u8();
    return [sign(f, 1 + (Math.floor(b / 12) << 8) + x), sign(f >> 1, 1 + (((b % 12) >> 2) << 8) + y)];
  }
  if (f < 124) {
    const b0 = r.u8(), b1 = r.u8(), b2 = r.u8();
    return [sign(f, (b0 << 4) + (b1 >> 4)), sign(f >> 1, ((b1 & 0x0f) << 8) + b2)];
  }
  const x = r.u16(), y = r.u16();
  return [sign(f, x), sign(f >> 1, y)];
}

/** Decode the push values of one glyph (with the Hop3 / Hop4 shorthands). */
function pushValues(r, count) {
  const out = [];
  while (out.length < count) {
    const c = r.b[r.p];
    if (c === 251 || c === 252) {
      r.p++;
      const a = out[out.length - 2] || 0;
      out.push(a, r.s255(), a);
      if (c === 252) out.push(r.s255(), a);
    } else out.push(r.s255());
  }
  return out;
}

/**
 * The shortest TrueType PUSHB / PUSHW / NPUSHB / NPUSHW sequence for the
 * values (a run of bytes or of words per instruction, at most 255 values).
 */
function pushInstructions(values) {
  const n = values.length;
  const cost = new Float64Array(n + 1).fill(Infinity);
  const from = new Int32Array(n + 1), words = new Uint8Array(n + 1);
  cost[0] = 0;
  for (let i = 1; i <= n; i++) {
    let bytes = true;
    for (let j = i - 1; j >= 0 && i - j <= 255; j--) {
      bytes = bytes && values[j] >= 0 && values[j] <= 255;
      const k = i - j, head = k <= 8 ? 1 : 2;
      const c = cost[j] + head + (bytes ? k : 2 * k);
      if (c < cost[i]) {
        cost[i] = c;
        from[i] = j;
        words[i] = bytes ? 0 : 1;
      }
    }
  }
  const runs = [];
  for (let i = n; i > 0; i = from[i]) runs.push([from[i], i, words[i]]);
  const out = [];
  for (const [j, i, w] of runs.reverse()) {
    const k = i - j;
    if (k <= 8) out.push((w ? 0xb8 : 0xb0) + k - 1);
    else out.push(w ? 0x41 : 0x40, k);
    for (let m = j; m < i; m++) {
      if (w) out.push((values[m] >> 8) & 0xff, values[m] & 0xff);
      else out.push(values[m]);
    }
  }
  return out;
}

function instructions(g, push, code) {
  const pushCount = g.u255();
  const codeSize = g.u255();
  return pushInstructions(pushValues(push, pushCount)).concat(Array.from(code.bytes(codeSize)));
}

function simpleGlyph(nc, g, push, code, w) {
  let box = null;
  if (nc === 0x7fff) {
    nc = g.i16();
    box = [g.i16(), g.i16(), g.i16(), g.i16()];
  }
  const ends = [];
  for (let i = 0; i < nc; i++) ends.push(i ? ends[i - 1] + g.u255() : g.u255());
  const n = nc ? ends[nc - 1] + 1 : 0;
  const flags = g.bytes(n);
  const dx = new Int32Array(n), dy = new Int32Array(n);
  let x = 0, y = 0, xMin = Infinity, yMin = Infinity, xMax = -Infinity, yMax = -Infinity;
  for (let i = 0; i < n; i++) {
    [dx[i], dy[i]] = triplet(flags[i] & 0x7f, g);
    x += dx[i];
    y += dy[i];
    xMin = Math.min(xMin, x); xMax = Math.max(xMax, x);
    yMin = Math.min(yMin, y); yMax = Math.max(yMax, y);
  }
  const prog = instructions(g, push, code);
  if (!box) box = n ? [xMin, yMin, xMax, yMax] : [0, 0, 0, 0];
  w.u16(nc);
  for (const v of box) w.u16(v);
  for (const e of ends) w.u16(e);
  w.u16(prog.length);
  w.bytes(prog);
  // Flags (with repeats), then x and y deltas in their shortest forms.
  const tf = new Uint8Array(n);
  for (let i = 0; i < n; i++) {
    let f = flags[i] & 0x80 ? 0 : 1;
    if (dx[i] === 0) f |= 0x10;
    else if (Math.abs(dx[i]) < 256) f |= 0x02 | (dx[i] > 0 ? 0x10 : 0);
    if (dy[i] === 0) f |= 0x20;
    else if (Math.abs(dy[i]) < 256) f |= 0x04 | (dy[i] > 0 ? 0x20 : 0);
    tf[i] = f;
  }
  for (let i = 0; i < n;) {
    let run = 0;
    while (i + run + 1 < n && tf[i + run + 1] === tf[i] && run < 255) run++;
    if (run) {
      w.u8(tf[i] | 0x08);
      w.u8(run);
    } else w.u8(tf[i]);
    i += run + 1;
  }
  for (const [d, short] of [[dx, 0x02], [dy, 0x04]]) {
    for (let i = 0; i < n; i++) {
      if (tf[i] & short) w.u8(Math.abs(d[i]));
      else if (d[i] !== 0) w.u16(d[i]);
    }
  }
  return prog.length;
}

function compositeGlyph(g, push, code, w) {
  w.u16(0xffff);
  for (let i = 0; i < 4; i++) w.u16(g.i16());
  let flags;
  do {
    flags = g.u16();
    w.u16(flags);
    w.u16(g.u16());
    let n = flags & 0x1 ? 4 : 2;
    if (flags & 0x8) n += 2;
    else if (flags & 0x40) n += 4;
    else if (flags & 0x80) n += 8;
    w.bytes(g.bytes(n));
  } while (flags & 0x20);
  if (!(flags & 0x100)) return 0;
  const prog = instructions(g, push, code);
  w.u16(prog.length);
  w.bytes(prog);
  return prog.length;
}

function decodeGlyf(r, end, numGlyphs, push, code) {
  const w = new Writer();
  const loca = [0];
  let maxInstr = 0;
  for (let gid = 0; gid < numGlyphs && r.p < end; gid++) {
    const nc = r.i16();
    if (nc < 0) maxInstr = Math.max(maxInstr, compositeGlyph(r, push, code, w));
    else if (nc > 0) maxInstr = Math.max(maxInstr, simpleGlyph(nc, r, push, code, w));
    w.pad(4);
    loca.push(w.length);
  }
  while (loca.length <= numGlyphs) loca.push(w.length);
  return { glyf: w.a, loca, maxInstr };
}

function checksum(bytes) {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) {
    sum = (sum + (((bytes[i] << 24) | ((bytes[i + 1] || 0) << 16) | ((bytes[i + 2] || 0) << 8) | (bytes[i + 3] || 0)) >>> 0)) >>> 0;
  }
  return sum;
}

/** Rebuild a TrueType font from the three decompressed CTF blocks. */
export function ctfToTtf(font, pushData, codeData) {
  const r = new Reader(font);
  const sfntVersion = (r.u16() << 16) | r.u16();
  const numTables = r.u16();
  r.p += 6;
  const dir = new Map();
  for (let i = 0; i < numTables; i++) {
    const tag = String.fromCharCode(...r.bytes(4));
    r.p += 4;
    const offset = (r.u16() << 16) | r.u16();
    const length = (r.u16() << 16) | r.u16();
    dir.set(tag, { offset: offset >>> 0, length: length >>> 0 });
  }
  const raw = (tag) => {
    const e = dir.get(tag);
    return e ? font.subarray(e.offset, e.offset + e.length) : null;
  };
  const tables = new Map();
  for (const [tag] of dir) {
    if (['loca', 'glyf', 'cvt ', 'hdmx', 'VDMX', 'DSIG'].includes(tag)) continue;
    tables.set(tag, Array.from(raw(tag)));
  }
  if (dir.has('cvt ')) tables.set('cvt ', decodeCvt(new Reader(font, dir.get('cvt ').offset)));
  const maxp = tables.get('maxp');
  const head = tables.get('head');
  if (dir.has('glyf') && maxp && head) {
    const e = dir.get('glyf');
    const numGlyphs = (maxp[4] << 8) | maxp[5];
    const { glyf, loca, maxInstr } = decodeGlyf(new Reader(font, e.offset), e.offset + e.length, numGlyphs, new Reader(pushData), new Reader(codeData));
    tables.set('glyf', glyf);
    const lw = new Writer();
    for (const o of loca) lw.u32(o);
    tables.set('loca', lw.a);
    head[50] = 0;
    head[51] = 1; // long loca offsets
    if (maxp.length >= 28 && maxInstr > ((maxp[26] << 8) | maxp[27])) {
      maxp[26] = maxInstr >> 8;
      maxp[27] = maxInstr & 0xff;
    }
  }
  if (tables.has('cmap')) fixCmap(tables.get('cmap'));
  return writeSfnt(sfntVersion >>> 0, tables);
}

/**
 * Embedded font subsets may leave junk in the language field of format
 * 12 / 13 cmap subtables; it is only meaningful for Macintosh subtables and
 * browsers reject fonts where it is not zero.
 */
function fixCmap(cmap) {
  const u16 = (p) => (cmap[p] << 8) | cmap[p + 1];
  const n = u16(2);
  for (let i = 0; i < n && 4 + i * 8 + 8 <= cmap.length; i++) {
    const rec = 4 + i * 8;
    const off = ((cmap[rec + 4] << 24) | (cmap[rec + 5] << 16) | (cmap[rec + 6] << 8) | cmap[rec + 7]) >>> 0;
    if (u16(rec) === 1 || off + 12 > cmap.length) continue;
    const format = u16(off);
    if (format === 12 || format === 13) cmap.fill(0, off + 8, off + 12);
  }
}

function writeSfnt(version, tables) {
  const tags = [...tables.keys()].sort();
  const n = tags.length;
  let pow = 1, log = 0;
  while (pow * 2 <= n) { pow *= 2; log++; }
  const w = new Writer();
  w.u32(version);
  w.u16(n);
  w.u16(pow * 16);
  w.u16(log);
  w.u16(n * 16 - pow * 16);
  let offset = 12 + 16 * n;
  const head = tables.get('head');
  if (head) head.fill(0, 8, 12);
  for (const tag of tags) {
    const data = tables.get(tag);
    for (let i = 0; i < 4; i++) w.u8(tag.charCodeAt(i));
    w.u32(checksum(data));
    w.u32(offset);
    w.u32(data.length);
    offset += (data.length + 3) & ~3;
  }
  for (const tag of tags) {
    w.bytes(tables.get(tag));
    w.pad(4);
  }
  const out = Uint8Array.from(w.a);
  if (head) {
    const at = 12 + 16 * n + tags.slice(0, tags.indexOf('head')).reduce((s, t) => s + ((tables.get(t).length + 3) & ~3), 0);
    const adj = (0xb1b0afba - checksum(out)) >>> 0;
    new DataView(out.buffer).setUint32(at + 8, adj);
  }
  return out;
}

/**
 * Decompress MTX data (the font data of a compressed EOT file) to TrueType.
 * @param {Uint8Array} bytes
 */
export function decodeMtx(bytes) {
  if (bytes.length < 10) throw new Error('MTX: data too short');
  const u24 = (p) => (bytes[p] << 16) | (bytes[p + 1] << 8) | bytes[p + 2];
  const version = bytes[0];
  const off2 = u24(4), off3 = u24(7);
  if (off2 < 10 || off3 < off2 || off3 > bytes.length) throw new Error('MTX: bad block offsets');
  const font = lzcompDecode(bytes.subarray(10, off2), version);
  const push = lzcompDecode(bytes.subarray(off2, off3), version);
  const code = lzcompDecode(bytes.subarray(off3), version);
  return ctfToTtf(font, push, code);
}

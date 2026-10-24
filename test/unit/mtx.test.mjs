import test from 'node:test';
import assert from 'node:assert/strict';
import { AHuff, lzcompDecode, decodeMtx } from '../../src/util/mtx.js';
import { fontFaceCss } from '../../src/render/embedded-fonts.js';

// ---- A small MTX encoder (CTF + LZCOMP) used to exercise the decoder ----

class BitWriter {
  constructor() { this.out = []; this.acc = 0; this.n = 0; }
  bit(b) {
    this.acc = (this.acc << 1) | (b ? 1 : 0);
    if (++this.n === 8) { this.out.push(this.acc); this.acc = 0; this.n = 0; }
  }
  value(v, bits) { for (let i = bits - 1; i >= 0; i--) this.bit((v >> i) & 1); }
  bytes() { return Uint8Array.from(this.n ? [...this.out, this.acc << (8 - this.n)] : this.out); }
}

function writeSymbol(model, w, sym) {
  const leaf = model.index[sym];
  const path = [];
  for (let a = leaf; a !== 1; a = model.up[a]) path.push(model.right[model.up[a]] === a ? 1 : 0);
  for (let i = path.length - 1; i >= 0; i--) w.bit(path[i]);
  model.update(leaf);
}

const PRELOAD = 2 * 32 * 96 + 4 * 256;
const bitsUsed = (v) => Math.max(1, 32 - Math.clz32(v));

function runLengthPack(data) {
  const count = new Array(256).fill(0);
  for (const b of data) count[b]++;
  const esc = count.indexOf(Math.min(...count));
  const out = [esc];
  for (let i = 0; i < data.length;) {
    let run = 1;
    while (i + run < data.length && data[i + run] === data[i] && run < 255) run++;
    if (run > 3) { out.push(esc, run, data[i]); i += run; continue; }
    if (data[i] === esc) out.push(esc, 0);
    else out.push(data[i]);
    i++;
  }
  return Uint8Array.from(out);
}

function lzcompEncode(input, runLength = false) {
  const data = runLength ? runLengthPack(input) : input;
  const w = new BitWriter();
  w.bit(runLength);
  const dist = new AHuff(null, 8), len = new AHuff(null, 8);
  w.value(data.length, 24);
  let ranges = 1;
  while (2 ** (3 * ranges) < data.length) ranges++;
  const DUP2 = 256 + 8 * ranges;
  const sym = new AHuff(null, DUP2 + 3);
  const buf = new Uint8Array(PRELOAD + data.length);
  let p = 0;
  for (let k = 0; k < 32; k++) for (let j = 0; j < 96; j++) { buf[p++] = k; buf[p++] = j; }
  for (let j = 0; j < 256; j++) for (let r = 0; r < 4; r++) buf[p++] = j;
  buf.set(data, PRELOAD);
  const maxDist = 2 ** (3 * ranges);
  for (let i = PRELOAD; i < buf.length;) {
    let best = 0, bestD = 0;
    for (let s = Math.max(0, i - 3000); s < i; s++) {
      let l = 0;
      while (i + l < buf.length && s + l < i && buf[s + l] === buf[i + l]) l++;
      const d = i - (s + l - 1);
      if (l >= 4 && d <= maxDist && l > best) { best = l; bestD = d; }
    }
    if (best) {
      const v = best - (bestD >= 512 ? 3 : 2);
      const groups = Math.ceil(bitsUsed(v) / 2);
      const nDist = Math.ceil(bitsUsed(bestD - 1) / 3);
      for (let g = groups - 1; g >= 0; g--) {
        const s = ((v >> (2 * g)) & 3) | (g ? 4 : 0);
        if (g === groups - 1) writeSymbol(sym, w, 256 + (nDist - 1) * 8 + s);
        else writeSymbol(len, w, s);
      }
      for (let g = nDist - 1; g >= 0; g--) writeSymbol(dist, w, ((bestD - 1) >> (3 * g)) & 7);
      i += best;
      continue;
    }
    const dup = [2, 4, 6].findIndex((k) => buf[i - k] === buf[i]);
    writeSymbol(sym, w, dup >= 0 && buf[i] !== buf[i - 1] ? DUP2 + dup : buf[i]);
    i++;
  }
  return w.bytes();
}

class Out {
  constructor() { this.a = []; }
  u8(...v) { for (const x of v) this.a.push(x & 0xff); return this; }
  u16(...v) { for (const x of v) this.a.push((x >> 8) & 0xff, x & 0xff); return this; }
  u32(v) { return this.u16(v >>> 16, v & 0xffff); }
  u255(v) { return v < 253 ? this.u8(v) : this.u8(253).u16(v); }
  s255(v) { return v >= 0 && v < 250 ? this.u8(v) : v < 0 && v > -250 ? this.u8(250, -v) : this.u8(253).u16(v); }
}

/** CTF triplet for a point delta. */
function triplet(dx, dy, flags, coords) {
  const ax = Math.abs(dx), ay = Math.abs(dy), sx = dx > 0 ? 1 : 0, sy = dy > 0 ? 1 : 0;
  let f;
  if (dx === 0 && ay < 1280) { f = 2 * (ay >> 8) + sy; coords.u8(ay); }
  else if (dy === 0 && ax < 1280) { f = 10 + 2 * (ax >> 8) + sx; coords.u8(ax); }
  else if (ax && ay && ax <= 64 && ay <= 64) { f = 20 + ((ax - 1) & 0x30) + (((ay - 1) >> 4) << 2) + 2 * sy + sx; coords.u8((((ax - 1) & 15) << 4) | ((ay - 1) & 15)); }
  else if (ax && ay && ax <= 768 && ay <= 768) { f = 84 + 12 * ((ax - 1) >> 8) + (((ay - 1) >> 8) << 2) + 2 * sy + sx; coords.u8(ax - 1, ay - 1); }
  else if (ax < 4096 && ay < 4096) { f = 120 + 2 * sy + sx; coords.u8(ax >> 4, ((ax & 15) << 4) | (ay >> 8), ay); }
  else { f = 124 + 2 * sy + sx; coords.u16(ax, ay); }
  flags.push(f);
}

function simple(g, contours, push, code, box) {
  if (box) g.u16(0x7fff, contours.length, ...box);
  else g.u16(contours.length);
  contours.forEach((c, i) => g.u255(i ? c.length : c.length - 1));
  const flags = [], coords = new Out();
  let x = 0, y = 0;
  for (const [px, py, on] of contours.flat()) {
    const n = flags.length;
    triplet(px - x, py - y, flags, coords);
    if (!on) flags[n] |= 0x80;
    x = px; y = py;
  }
  g.u8(...flags);
  g.a.push(...coords.a);
  g.u255(push.n).u255(code.length);
}

// Glyph 1: two contours whose deltas hit every triplet class.
const OUTLINE = [
  [[0, -5, 1], [300, -5, 1], [303, -12, 0], [503, -112, 1], [2503, -1612, 1]],
  [[-27497, 18388, 1], [-27497, 18388, 0], [-27000, 18000, 1]],
];

function buildFont() {
  const glyf = new Out(), push = new Out(), code = [];
  glyf.u16(0); // glyph 0: empty
  // Glyph 1 pushes 1, 2, 1, 3, 1 (with a Hop3 code) and runs two instructions.
  push.u8(1, 2, 251, 3);
  code.push(0x2b, 0x2b);
  simple(glyf, OUTLINE, { n: 5 }, [0x2b, 0x2b]);
  // Glyph 2: explicit bounding box, no instructions.
  simple(glyf, [[[10, 10, 1], [20, 10, 1], [20, 20, 1]]], { n: 0 }, [], [0, 0, 100, 100]);
  // Glyph 3: composite of glyph 1 pushing 300, -5, 300, 7, 300, 8, 300 (Hop4).
  glyf.u16(0xffff, 0, 0, 50, 50);
  glyf.u16(0x0002 | 0x0100, 1).u8(5, 6);
  glyf.u255(7).u255(1);
  push.s255(300).s255(-5).u8(252).s255(7).s255(8);
  code.push(0x2c);

  const cvtValues = [100, 600, -400, 30000, 29990, 0];
  const cvt = new Out().u16(cvtValues.length);
  let prev = 0;
  for (const v of cvtValues) {
    const d = ((v - prev + 0x8000) & 0xffff) - 0x8000;
    prev = v;
    const a = Math.abs(d), idx = Math.floor(a / 238);
    if (idx > 8) cvt.u8(238).u16(d);
    else if (d < 0) cvt.u8(239 + idx, a - idx * 238);
    else if (idx) cvt.u8(247 + idx, a - idx * 238);
    else cvt.u8(a);
  }

  const cmap = new Out().u16(0, 1, 3, 10).u32(12).u16(12, 0).u32(28).u32(0xdeadbeef).u32(1).u32(0x41).u32(0x41).u32(1);
  const head = new Out().u32(0x10000).u32(0x10000).u32(0).u32(0x5f0f3cf5).u16(0, 1000).u32(0).u32(0).u32(0).u32(0).u16(0, 0, 100, 100, 0, 8, 2, 0, 0);
  const maxp = new Out().u32(0x10000).u16(4, 8, 2, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0);
  const tables = [['cmap', cmap.a], ['cvt ', cvt.a], ['glyf', glyf.a], ['hdmx', [1, 2, 3, 4]], ['head', head.a], ['loca', []], ['maxp', maxp.a]];
  const font = new Out().u32(0x10000).u16(tables.length, 0, 0, 0);
  let off = 12 + 16 * tables.length;
  for (const [tag, data] of tables) {
    font.u8(...[...tag].map((c) => c.charCodeAt(0))).u32(0).u32(data.length ? off : 0).u32(data.length);
    off += data.length;
  }
  for (const [, data] of tables) font.a.push(...data);
  return { font: Uint8Array.from(font.a), push: Uint8Array.from(push.a), code: Uint8Array.from(code), cvtValues };
}

function mtx(font, push, code) {
  const blocks = [lzcompEncode(font, true), lzcompEncode(push), lzcompEncode(code)];
  const off2 = 10 + blocks[0].length, off3 = off2 + blocks[1].length;
  const head = [3, 0, 0, 0, off2 >> 16, (off2 >> 8) & 0xff, off2 & 0xff, off3 >> 16, (off3 >> 8) & 0xff, off3 & 0xff];
  return Uint8Array.from([...head, ...blocks[0], ...blocks[1], ...blocks[2]]);
}

// ---- TrueType reading helpers ----

function sfnt(bytes) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const tables = {};
  for (let i = 0; i < v.getUint16(4); i++) {
    const p = 12 + i * 16;
    const tag = String.fromCharCode(...bytes.subarray(p, p + 4));
    tables[tag] = new DataView(bytes.buffer, bytes.byteOffset + v.getUint32(p + 8), v.getUint32(p + 12));
  }
  return tables;
}

function glyph(t, gid) {
  const start = t.loca.getUint32(gid * 4), end = t.loca.getUint32(gid * 4 + 4);
  if (start === end) return null;
  const g = t.glyf;
  let p = start;
  const nc = g.getInt16(p);
  const box = [2, 4, 6, 8].map((k) => g.getInt16(p + k));
  p += 10;
  if (nc < 0) return { composite: true, box, flags: g.getUint16(p), gid: g.getUint16(p + 2), instr: [...new Uint8Array(g.buffer, g.byteOffset + p + 8, g.getUint16(p + 6))] };
  const ends = [];
  for (let i = 0; i < nc; i++, p += 2) ends.push(g.getUint16(p));
  const ilen = g.getUint16(p);
  const instr = [...new Uint8Array(g.buffer, g.byteOffset + p + 2, ilen)];
  p += 2 + ilen;
  const n = ends[nc - 1] + 1, flags = [];
  while (flags.length < n) {
    const f = g.getUint8(p++);
    flags.push(f);
    if (f & 8) for (let r = g.getUint8(p++); r > 0; r--) flags.push(f);
  }
  const coord = (short, same) => {
    let v = 0;
    return flags.map((f) => {
      if (f & short) v += f & same ? g.getUint8(p++) : -g.getUint8(p++);
      else if (!(f & same)) { v += g.getInt16(p); p += 2; }
      return v;
    });
  };
  const xs = coord(2, 16), ys = coord(4, 32);
  return { ends, instr, box, points: xs.map((x, i) => [x, ys[i], flags[i] & 1]) };
}

test('LZCOMP: literals, single-byte and long copies, run-length coding', () => {
  const data = new Uint8Array(4000);
  for (let i = 0; i < data.length; i++) data[i] = (i * 7 + (i >> 5)) % 251;
  data.set(data.subarray(100, 300), 1500); // far copy (distance > 512)
  data.set(data.subarray(1500, 1520), 1530); // near copy
  data.fill(9, 2000, 2400); // run
  for (let i = 3000; i < 3100; i += 2) data[i] = 77; // DUP2 pattern
  for (const rl of [false, true]) assert.deepEqual(lzcompDecode(lzcompEncode(data, rl)), data);
  assert.deepEqual(lzcompDecode(new Uint8Array(0)), new Uint8Array(0));
});

test('MTX: CTF font is rebuilt as TrueType', () => {
  const { font, push, code, cvtValues } = buildFont();
  const ttf = decodeMtx(mtx(font, push, code));
  const t = sfnt(ttf);
  assert.deepEqual(Object.keys(t).sort(), ['cmap', 'cvt ', 'glyf', 'head', 'loca', 'maxp']);
  // Whole-font checksum and long loca format.
  let sum = 0;
  for (let i = 0; i < ttf.length; i += 4) sum = (sum + new DataView(ttf.buffer).getUint32(i)) >>> 0;
  assert.equal(sum, 0xb1b0afba);
  assert.equal(t.head.getInt16(50), 1);
  assert.deepEqual(cvtValues.map((_, i) => t['cvt '].getInt16(i * 2)), cvtValues);
  assert.equal(t.cmap.getUint32(12 + 8), 0, 'cmap format 12 language reset');

  assert.equal(glyph(t, 0), null);
  const g1 = glyph(t, 1);
  assert.deepEqual(g1.ends, [4, 7]);
  assert.deepEqual(g1.points, OUTLINE.flat());
  assert.deepEqual(g1.box, [-27497, -1612, 2503, 18388]);
  assert.deepEqual(g1.instr, [0xb4, 1, 2, 1, 3, 1, 0x2b, 0x2b]);
  assert.deepEqual(glyph(t, 2).box, [0, 0, 100, 100]);
  const g3 = glyph(t, 3);
  assert.ok(g3.composite);
  assert.equal(g3.gid, 1);
  // 300 needs words, -5 is negative: everything is pushed with PUSHW.
  assert.deepEqual(g3.instr, [0xbe, 1, 44, 0xff, 0xfb, 1, 44, 0, 7, 1, 44, 0, 8, 1, 44, 0x2c]);
  assert.equal(t.maxp.getUint16(26), 16, 'maxSizeOfInstructions');
});

test('compressed EOT fonts become @font-face rules', () => {
  const { font, push, code } = buildFont();
  const data = mtx(font, push, code).map((b) => b ^ 0x50);
  const eot = new Uint8Array(82 + 4 * 4 + data.length);
  const v = new DataView(eot.buffer);
  v.setUint32(0, eot.length, true);
  v.setUint32(4, data.length, true);
  v.setUint32(8, 0x00020001, true);
  v.setUint32(12, 0x4 | 0x10000000, true);
  v.setUint16(34, 0x504c, true);
  eot.set(data, eot.length - data.length);
  const diags = [];
  const pres = { fonts: [{ typeface: 'Test Font', variants: { regular: 'f' } }], media: { pkg: { bytes: () => eot } }, diagnostics: { add: (...a) => diags.push(a) } };
  const css = fontFaceCss(pres, { opts: {} });
  assert.match(css, /font-family:"Test Font";src:url\(data:font\/ttf;base64,AAEAAA/);
  assert.deepEqual(diags, []);
});

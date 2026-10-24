import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { PNG } from 'pngjs';
import { deflateRaw } from '../../src/util/deflate.js';
import { encodePng } from '../../src/util/png.js';
import { decodeTiff } from '../../src/util/tiff.js';
import { toWebImage } from '../../src/util/webimage.js';

test('deflate round-trips through zlib', () => {
  const text = new TextEncoder().encode('PPTX2HTML '.repeat(500) + 'tail');
  const noise = Uint8Array.from({ length: 5000 }, (_, i) => (i * 7919) % 251);
  for (const data of [new Uint8Array(0), Uint8Array.of(42), text, noise]) {
    assert.deepEqual(new Uint8Array(zlib.inflateRawSync(deflateRaw(data))), data);
  }
  assert.ok(deflateRaw(text).length < 100);
});

test('PNG encoder output decodes with pngjs', () => {
  const w = 37, h = 11;
  const rgba = Uint8Array.from({ length: w * h * 4 }, (_, i) => (i * 31 + (i >> 5)) & 0xff);
  const png = PNG.sync.read(Buffer.from(encodePng(w, h, rgba)));
  assert.equal(png.width, w);
  assert.deepEqual(new Uint8Array(png.data), rgba);
  const opaque = rgba.map((v, i) => (i % 4 === 3 ? 255 : v));
  const bytes = encodePng(w, h, opaque);
  assert.equal(bytes[25], 2); // colour type RGB
  assert.deepEqual(new Uint8Array(PNG.sync.read(Buffer.from(bytes)).data), opaque);
});

/** Build a single-strip TIFF. tags: [tag, type, values] */
function tiff(le, tags, data) {
  const entries = tags.length;
  const ifd = 8, extra = ifd + 2 + entries * 12 + 4;
  const blobs = [];
  let off = extra;
  const size = { 3: 2, 4: 4 };
  const enc = tags.map(([tag, type, vals]) => {
    const bytes = vals.length * size[type];
    if (bytes <= 4) return { tag, type, vals, inline: true };
    const e = { tag, type, vals, off };
    blobs.push(e);
    off += bytes;
    return e;
  });
  const dataOff = off;
  const buf = new ArrayBuffer(dataOff + data.length);
  const v = new DataView(buf);
  new Uint8Array(buf).set(le ? [0x49, 0x49] : [0x4d, 0x4d]);
  v.setUint16(2, 42, le);
  v.setUint32(4, ifd, le);
  v.setUint16(ifd, entries, le);
  const put = (p, type, val) => (type === 3 ? v.setUint16(p, val, le) : v.setUint32(p, val, le));
  enc.forEach((e, i) => {
    const p = ifd + 2 + i * 12;
    v.setUint16(p, e.tag, le);
    v.setUint16(p + 2, e.type, le);
    v.setUint32(p + 4, e.vals.length, le);
    const vals = e.tag === 273 ? [dataOff] : e.vals;
    if (e.inline) vals.forEach((x, k) => put(p + 8 + k * size[e.type], e.type, x));
    else v.setUint32(p + 8, e.off, le);
  });
  for (const e of blobs) e.vals.forEach((x, k) => put(e.off + k * size[e.type], e.type, x));
  new Uint8Array(buf).set(data, dataOff);
  return new Uint8Array(buf);
}

function lzwEncode(data) {
  // TIFF LZW (MSB first, early change).
  const out = [];
  let acc = 0, bits = 0, codeLen = 9;
  const emit = (c) => { acc = (acc << codeLen) | c; bits += codeLen; while (bits >= 8) { out.push((acc >>> (bits - 8)) & 0xff); bits -= 8; } };
  let dict = new Map(), next = 258;
  emit(256);
  let w = String.fromCharCode(data[0]);
  for (let i = 1; i < data.length; i++) {
    const c = String.fromCharCode(data[i]);
    if (dict.has(w + c) || (w + c).length === 1) { w += c; continue; }
    emit(w.length === 1 ? w.charCodeAt(0) : dict.get(w));
    dict.set(w + c, next++);
    if (next + 1 >= 1 << codeLen) codeLen++;
    w = c;
  }
  emit(w.length === 1 ? w.charCodeAt(0) : dict.get(w));
  emit(257);
  if (bits) out.push((acc << (8 - bits)) & 0xff);
  return Uint8Array.from(out);
}

test('TIFF: uncompressed RGB, PackBits gray, palette and LZW with predictor', () => {
  const rgb = Uint8Array.of(255, 0, 0, 0, 255, 0, 0, 0, 255, 10, 20, 30);
  let img = decodeTiff(tiff(true, [[256, 3, [2]], [257, 3, [2]], [258, 3, [8, 8, 8]], [259, 3, [1]], [262, 3, [2]], [273, 4, [0]], [277, 3, [3]], [278, 3, [2]], [279, 4, [12]]], rgb));
  assert.deepEqual([...img.rgba.slice(0, 8)], [255, 0, 0, 255, 0, 255, 0, 255]);
  assert.deepEqual([...img.rgba.slice(12)], [10, 20, 30, 255]);

  // PackBits: 3 x 0x80 then a literal 0x10, big-endian, white-is-zero.
  const pb = Uint8Array.of(0xfe, 0x80, 0x00, 0x10);
  img = decodeTiff(tiff(false, [[256, 3, [4]], [257, 3, [1]], [258, 3, [8]], [259, 3, [32773]], [262, 3, [0]], [273, 4, [0]], [278, 3, [1]], [279, 4, [4]]], pb));
  assert.deepEqual([...img.rgba.filter((_, i) => i % 4 === 0)], [127, 127, 127, 239]);

  // 4-bit palette: colour map holds 16 entries per channel (16-bit values).
  const cmap = [];
  for (const ch of [0, 1, 2]) for (let i = 0; i < 16; i++) cmap.push(i === 1 ? [0xff00, 0x8000, 0x0000][ch] : 0);
  img = decodeTiff(tiff(true, [[256, 3, [2]], [257, 3, [1]], [258, 3, [4]], [259, 3, [1]], [262, 3, [3]], [273, 4, [0]], [278, 3, [1]], [279, 4, [1]], [320, 3, cmap]], Uint8Array.of(0x10)));
  assert.deepEqual([...img.rgba], [255, 128, 0, 255, 0, 0, 0, 255]);

  // LZW + horizontal differencing (RGBA).
  const W = 16, H = 4, px = [];
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) px.push(x * 10, y * 40, 200, 255);
  const diff = [];
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) for (let s = 0; s < 4; s++) {
    const i = (y * W + x) * 4 + s;
    diff.push(x ? (px[i] - px[i - 4]) & 0xff : px[i]);
  }
  const lz = lzwEncode(Uint8Array.from(diff));
  img = decodeTiff(tiff(true, [[256, 3, [W]], [257, 3, [H]], [258, 3, [8, 8, 8, 8]], [259, 3, [5]], [262, 3, [2]], [273, 4, [0]], [277, 3, [4]], [278, 3, [H]], [279, 4, [lz.length]], [317, 3, [2]], [338, 3, [2]]], lz));
  assert.deepEqual([...img.rgba], px);
});

test('toWebImage converts TIFF to PNG and leaves PNG alone', () => {
  const t = tiff(true, [[256, 3, [1]], [257, 3, [1]], [258, 3, [8]], [259, 3, [1]], [262, 3, [1]], [273, 4, [0]], [278, 3, [1]], [279, 4, [1]]], Uint8Array.of(200));
  const web = toWebImage(t, 'image/tiff');
  assert.equal(web.mime, 'image/png');
  assert.deepEqual([...PNG.sync.read(Buffer.from(web.bytes)).data], [200, 200, 200, 255]);
  assert.equal(toWebImage(web.bytes, 'image/png'), null);
});

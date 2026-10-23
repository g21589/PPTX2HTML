import test from 'node:test';
import assert from 'node:assert/strict';
import { emfToSvg } from '../../src/util/emf.js';

// Little-endian record writers.
function emfRecord(type, ints) {
  const b = new DataView(new ArrayBuffer(8 + ints.length * 4));
  b.setUint32(0, type, true);
  b.setUint32(4, b.byteLength, true);
  ints.forEach((v, i) => b.setInt32(8 + i * 4, v, true));
  return new Uint8Array(b.buffer);
}

function emf(records) {
  // EMR_HEADER: bounds, frame (0.01 mm), signature, version, size, count, handles, ..., device px, device mm.
  const header = emfRecord(1, [0, 0, 99, 99, 0, 0, 10000, 10000, 0x464d4520, 0x10000, 0, records.length + 2, 4, 0, 0, 0, 1000, 1000, 100, 100]);
  const all = [header, ...records, emfRecord(14, [0, 16, 20])];
  const out = new Uint8Array(all.reduce((a, r) => a + r.length, 0));
  let o = 0;
  for (const r of all) { out.set(r, o); o += r.length; }
  return out;
}

function wmf(records) {
  const words = [1, 9, 0x300, 0, 0, 0, 0, 0, 0]; // META_HEADER (sizes are not checked)
  for (const [fn, args] of records) words.push(3 + args.length, 0, fn, ...args);
  words.push(3, 0, 0);
  const b = new DataView(new ArrayBuffer(words.length * 2));
  words.forEach((w, i) => b.setUint16(i * 2, w & 0xffff, true));
  return new Uint8Array(b.buffer);
}

test('EMF shapes use the selected brush and raster no-ops are skipped', () => {
  const svg = emfToSvg(emf([
    emfRecord(39, [1, 0, 0x0000ff, 0]), // CREATEBRUSHINDIRECT red (COLORREF is 0x00bbggrr)
    emfRecord(37, [1]), // SELECTOBJECT
    emfRecord(43, [10, 10, 50, 40]), // RECTANGLE
    // BITBLT with ROP "D" (destination unchanged) and no bitmap.
    emfRecord(76, [0, 0, 99, 99, 0, 0, 100, 100, 0x00aa0029, 0, 0, 0x3f800000, 0, 0, 0x3f800000, 0, 0, 0, 0, 0, 0, 0, 0]),
  ]), 'image/x-emf');
  assert.ok(svg.startsWith('<svg'));
  assert.match(svg, /<path d="M10 10L50 10L50 40L10 40Z" fill="#ff0000"/);
  assert.equal((svg.match(/<path/g) || []).length, 1);
});

test('WMF text honours font charset, current position and symbol fonts', () => {
  const font = (height, charset, face) => {
    const bytes = [...Buffer.from(face, 'latin1'), 0];
    if (bytes.length & 1) bytes.push(0);
    const words = [height, 0, 0, 0, 400, 0, charset << 8, 0, 0];
    for (let i = 0; i < bytes.length; i += 2) words.push(bytes[i] | (bytes[i + 1] << 8));
    return [0x02fb, words];
  };
  const textOut = (str, x, y) => {
    const words = [str.length];
    for (let i = 0; i < str.length; i += 2) words.push(str[i] | ((str[i + 1] || 0) << 8));
    words.push(y, x);
    return [0x0521, words];
  };
  const svg = emfToSvg(wmf([
    [0x020b, [0, 0]], [0x020c, [100, 200]], // window org / ext
    font(-20, 134, 'SimSun'), [0x012d, [0]],
    textOut([0xd0, 0xc2, 0xbd, 0xa8], 10, 20), // "新建" in GBK
    font(-20, 2, 'Symbol'), [0x012d, [1]],
    [0x012e, [25]], // TA_UPDATECP | TA_BASELINE
    [0x0214, [50, 60]], // MOVETO y=50 x=60
    [0x0a32, [0, 0, 1, 0, 0x70, 30]], // EXTTEXTOUT "p" with a 30 unit advance
    [0x0a32, [0, 0, 1, 0, 0x70, 30]],
  ]), 'image/x-wmf');
  assert.match(svg, />新建<\/text>/);
  const pis = [...svg.matchAll(/<text x="([\d.]+)" y="([\d.]+)"[^>]*font-family="serif"[^>]*>π<\/text>/g)];
  assert.equal(pis.length, 2);
  assert.deepEqual(pis.map((m) => [+m[1], +m[2]]), [[60, 50], [90, 50]]);
});

test('empty metafiles yield null', () => {
  assert.equal(emfToSvg(emf([]), 'image/x-emf'), null);
  assert.equal(emfToSvg(new Uint8Array(10), 'image/x-wmf'), null);
});

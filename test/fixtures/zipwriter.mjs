/** Minimal ZIP writer (deflate) used to generate PPTX fixtures. */
import zlib from 'node:zlib';

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * @param {Array<{name: string, data: Buffer|string}>} files
 * @returns {Buffer}
 */
export function zip(files) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8');
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, 'utf8');
    const crc = crc32(data);
    const comp = zlib.deflateRawSync(data, { level: 6 });
    const useStore = comp.length >= data.length;
    const body = useStore ? data : comp;
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);
    lh.writeUInt16LE(0x800, 6);
    lh.writeUInt16LE(useStore ? 0 : 8, 8);
    lh.writeUInt16LE(0, 10);
    lh.writeUInt16LE(0x21, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(body.length, 18);
    lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(name.length, 26);
    lh.writeUInt16LE(0, 28);
    locals.push(lh, name, body);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);
    ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x800, 8);
    ch.writeUInt16LE(useStore ? 0 : 8, 10);
    ch.writeUInt16LE(0, 12);
    ch.writeUInt16LE(0x21, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(body.length, 20);
    ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(name.length, 28);
    ch.writeUInt32LE(offset, 42);
    central.push(ch, name);
    offset += 30 + name.length + body.length;
  }
  const cd = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

/** Encode an RGBA image as PNG. */
export function png(width, height, pixel) {
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[y * (width * 4 + 1)] = 0;
    for (let x = 0; x < width; x++) {
      const [r, g, b, a] = pixel(x, y);
      const o = y * (width * 4 + 1) + 1 + x * 4;
      raw[o] = r; raw[o + 1] = g; raw[o + 2] = b; raw[o + 3] = a === undefined ? 255 : a;
    }
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type, 'ascii'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** A short mono 16-bit WAV tone. */
export function wav(ms = 400, freq = 440, rate = 8000) {
  const n = Math.round((rate * ms) / 1000);
  const data = Buffer.alloc(n * 2);
  for (let i = 0; i < n; i++) data.writeInt16LE(Math.round(Math.sin((2 * Math.PI * freq * i) / rate) * 8000), i * 2);
  const h = Buffer.alloc(44);
  h.write('RIFF', 0); h.writeUInt32LE(36 + data.length, 4); h.write('WAVE', 8); h.write('fmt ', 12);
  h.writeUInt32LE(16, 16); h.writeUInt16LE(1, 20); h.writeUInt16LE(1, 22); h.writeUInt32LE(rate, 24);
  h.writeUInt32LE(rate * 2, 28); h.writeUInt16LE(2, 32); h.writeUInt16LE(16, 34); h.write('data', 36); h.writeUInt32LE(data.length, 40);
  return Buffer.concat([h, data]);
}

/**
 * A minimal EMF (MM_TEXT, 1 logical unit = 1 device pixel).
 * `records` are [type, int32[]] pairs written after the header.
 */
export function emf(width, height, records) {
  const rec = (type, ints) => {
    const b = Buffer.alloc(8 + ints.length * 4);
    b.writeUInt32LE(type, 0); b.writeUInt32LE(b.length, 4);
    ints.forEach((v, i) => b.writeInt32LE(v, 8 + i * 4));
    return b;
  };
  const body = records.map(([t, ints]) => rec(t, ints));
  const eof = rec(14, [0, 16, 20]);
  const mm = [Math.round(width / 4), Math.round(height / 4)]; // 4 px per mm
  const header = rec(1, [0, 0, width - 1, height - 1, 0, 0, mm[0] * 100, mm[1] * 100, 0x464d4520, 0x10000, 0, body.length + 2, 8, 0, 0, 0, width, height, mm[0], mm[1]]);
  const out = Buffer.concat([header, ...body, eof]);
  out.writeUInt32LE(out.length, 48);
  return out;
}

/** A placeable WMF; `records` are [function, int16[]] pairs. */
export function wmf(width, height, records) {
  const words = [];
  for (const [fn, args] of records) words.push(3 + args.length, 0, fn, ...args);
  words.push(3, 0, 0);
  const total = 9 + words.length;
  const h = Buffer.alloc(22 + 18);
  h.writeUInt32LE(0x9ac6cdd7, 0);
  h.writeInt16LE(0, 6); h.writeInt16LE(0, 8); h.writeInt16LE(width, 10); h.writeInt16LE(height, 12); h.writeUInt16LE(1440, 14);
  let sum = 0;
  for (let i = 0; i < 10; i++) sum ^= h.readUInt16LE(i * 2);
  h.writeUInt16LE(sum, 20);
  [1, 9, 0x300, total & 0xffff, total >>> 16, 8, 64, 0, 0].forEach((w, i) => h.writeUInt16LE(w, 22 + i * 2));
  const b = Buffer.alloc(words.length * 2);
  words.forEach((w, i) => b.writeUInt16LE(w & 0xffff, i * 2));
  return Buffer.concat([h, b]);
}

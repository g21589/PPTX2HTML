/**
 * Baseline TIFF decoder (first image of the file) producing RGBA pixels.
 *
 * Supports strips and tiles; compression none, LZW, PackBits and Deflate;
 * horizontal predictor; bilevel / grayscale (1-16 bit), palette, RGB(A) and
 * CMYK images. JPEG-compressed TIFFs are not decoded (returns null).
 */
import { inflateRaw } from './inflate.js';

function lzw(src, expected) {
  const out = new Uint8Array(expected);
  let o = 0;
  let pos = 0, bits = 0, buf = 0;
  let codeLen = 9;
  let dict = [];
  const reset = () => { dict = []; for (let i = 0; i < 256; i++) dict.push([i]); dict.push(null, null); codeLen = 9; };
  reset();
  let prev = null;
  for (;;) {
    while (bits < codeLen) {
      if (pos >= src.length) return out.subarray(0, o);
      buf = (buf << 8) | src[pos++];
      bits += 8;
    }
    const code = (buf >>> (bits - codeLen)) & ((1 << codeLen) - 1);
    bits -= codeLen;
    if (code === 257) break;
    if (code === 256) { reset(); prev = null; continue; }
    let entry;
    if (code < dict.length && dict[code]) entry = dict[code];
    else if (prev) entry = prev.concat(prev[0]);
    else break;
    for (let i = 0; i < entry.length && o < expected; i++) out[o++] = entry[i];
    if (prev) dict.push(prev.concat(entry[0]));
    prev = entry;
    // TIFF LZW switches code width one code early.
    if (dict.length + 1 >= (1 << codeLen) && codeLen < 12) codeLen++;
    if (o >= expected) break;
  }
  return out.subarray(0, o);
}

function packBits(src, expected) {
  const out = new Uint8Array(expected);
  let i = 0, o = 0;
  while (i < src.length && o < expected) {
    const n = (src[i++] << 24) >> 24;
    if (n >= 0) {
      for (let k = 0; k <= n && i < src.length && o < expected; k++) out[o++] = src[i++];
    } else if (n !== -128) {
      const v = src[i++];
      for (let k = 0; k < 1 - n && o < expected; k++) out[o++] = v;
    }
  }
  return out;
}

/**
 * Decode a TIFF file.
 * @param {Uint8Array} b
 * @returns {{width: number, height: number, rgba: Uint8Array} | null}
 */
export function decodeTiff(b) {
  if (b.length < 8) return null;
  const le = b[0] === 0x49 && b[1] === 0x49;
  if (!le && !(b[0] === 0x4d && b[1] === 0x4d)) return null;
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const u16 = (o) => v.getUint16(o, le), u32 = (o) => v.getUint32(o, le);
  if (u16(2) !== 42) return null;
  const ifd = u32(4);
  if (ifd + 2 > b.length) return null;
  const SIZES = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 16: 8 };
  const tags = {};
  const n = u16(ifd);
  for (let i = 0; i < n; i++) {
    const e = ifd + 2 + i * 12;
    if (e + 12 > b.length) break;
    const tag = u16(e), type = u16(e + 2), count = u32(e + 4);
    const size = (SIZES[type] || 1) * count;
    const off = size <= 4 ? e + 8 : u32(e + 8);
    const vals = [];
    for (let k = 0; k < Math.min(count, 1 << 20); k++) {
      const p = off + k * (SIZES[type] || 1);
      if (p + (SIZES[type] || 1) > b.length) break;
      vals.push(type === 3 ? u16(p) : type === 4 || type === 16 ? u32(p) : type === 5 ? u32(p) / (u32(p + 4) || 1) : b[p]);
    }
    tags[tag] = vals;
  }
  const one = (t, d) => (tags[t] && tags[t].length ? tags[t][0] : d);
  const W = one(256, 0), H = one(257, 0);
  if (!W || !H || W * H > 64e6) return null;
  const bps = tags[258] || [1];
  const bits = bps[0];
  const comp = one(259, 1);
  const photo = one(262, 1);
  const spp = one(277, 1);
  const planar = one(284, 1);
  const predictor = one(317, 1);
  const extra = tags[338] || [];
  if (![1, 5, 8, 32773, 32946].includes(comp)) return null;
  if (planar !== 1 && spp > 1) return null;

  const tiled = !!tags[322];
  const tw = tiled ? one(322, W) : W;
  const th = tiled ? one(323, H) : one(278, H);
  const offsets = tiled ? tags[324] : tags[273];
  const counts = tiled ? tags[325] : tags[279];
  if (!offsets || !counts) return null;
  const rowBytes = Math.ceil((tw * spp * bits) / 8);
  const chunkBytes = rowBytes * th;
  const across = tiled ? Math.ceil(W / tw) : 1;

  // Samples of the whole image, 8 or 16 bits per sample unpacked to numbers.
  const samples = new Uint16Array(W * H * spp);
  const maxv = (1 << bits) - 1;
  for (let c = 0; c < offsets.length; c++) {
    const raw = b.subarray(offsets[c], Math.min(b.length, offsets[c] + counts[c]));
    let data;
    if (comp === 1) data = raw;
    else if (comp === 5) data = lzw(raw, chunkBytes);
    else if (comp === 32773) data = packBits(raw, chunkBytes);
    else {
      try { data = inflateRaw(raw.subarray(2)); } catch (e) { return null; }
    }
    const x0 = tiled ? (c % across) * tw : 0;
    const y0 = tiled ? Math.floor(c / across) * th : c * th;
    for (let y = 0; y < th && y0 + y < H; y++) {
      const row = y * rowBytes;
      const vals = new Uint16Array(tw * spp);
      for (let i = 0; i < tw * spp; i++) {
        if (bits === 8) vals[i] = data[row + i] || 0;
        else if (bits === 16) vals[i] = le ? (data[row + i * 2] | (data[row + i * 2 + 1] << 8)) : ((data[row + i * 2] << 8) | data[row + i * 2 + 1]);
        else {
          const bit = i * bits;
          const byte = data[row + (bit >> 3)] || 0;
          vals[i] = (byte >> (8 - bits - (bit & 7))) & maxv;
        }
      }
      if (predictor === 2) for (let i = spp; i < tw * spp; i++) vals[i] = (vals[i] + vals[i - spp]) & maxv;
      for (let x = 0; x < tw && x0 + x < W; x++) {
        for (let s = 0; s < spp; s++) samples[((y0 + y) * W + x0 + x) * spp + s] = vals[x * spp + s];
      }
    }
  }

  const rgba = new Uint8Array(W * H * 4);
  const to8 = (val) => (bits === 8 ? val : bits === 16 ? val >> 8 : Math.round((val * 255) / maxv));
  const cmap = tags[320];
  const alphaIdx = spp > (photo === 2 ? 3 : photo === 5 ? 4 : 1) ? (photo === 2 ? 3 : photo === 5 ? 4 : 1) : -1;
  const premultiplied = extra[0] === 1;
  for (let p = 0; p < W * H; p++) {
    const s = p * spp, o = p * 4;
    let r, g, bl;
    if (photo === 2) { r = to8(samples[s]); g = to8(samples[s + 1]); bl = to8(samples[s + 2]); }
    else if (photo === 3 && cmap) {
      const n3 = cmap.length / 3, i = samples[s];
      r = cmap[i] >> 8; g = cmap[n3 + i] >> 8; bl = cmap[2 * n3 + i] >> 8;
    } else if (photo === 5) {
      const k = 255 - to8(samples[s + 3]);
      r = ((255 - to8(samples[s])) * k) / 255; g = ((255 - to8(samples[s + 1])) * k) / 255; bl = ((255 - to8(samples[s + 2])) * k) / 255;
    } else {
      const gray = to8(samples[s]);
      r = g = bl = photo === 0 ? 255 - gray : gray;
    }
    let a = alphaIdx >= 0 ? to8(samples[s + alphaIdx]) : 255;
    if (premultiplied && a > 0 && a < 255) { r = (r * 255) / a; g = (g * 255) / a; bl = (bl * 255) / a; }
    rgba[o] = Math.min(255, r); rgba[o + 1] = Math.min(255, g); rgba[o + 2] = Math.min(255, bl); rgba[o + 3] = a;
  }
  return { width: W, height: H, rgba };
}

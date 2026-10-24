/**
 * Image header parsing: intrinsic size and resolution of common formats.
 */

function u16be(b, i) { return (b[i] << 8) | b[i + 1]; }
function u32be(b, i) { return ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3]; }
function u16le(b, i) { return b[i] | (b[i + 1] << 8); }
function i32le(b, i) { return b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24); }

function png(b) {
  if (b.length < 24) return null;
  const out = { width: u32be(b, 16), height: u32be(b, 20) };
  // Look for pHYs.
  let p = 8;
  while (p + 12 <= b.length) {
    const len = u32be(b, p);
    const type = String.fromCharCode(b[p + 4], b[p + 5], b[p + 6], b[p + 7]);
    if (type === 'pHYs' && len >= 9) {
      const px = u32be(b, p + 8);
      const py = u32be(b, p + 12);
      if (b[p + 16] === 1 && px > 0 && py > 0) {
        out.dpiX = px * 0.0254;
        out.dpiY = py * 0.0254;
      }
      break;
    }
    if (type === 'IDAT' || type === 'IEND') break;
    p += 12 + len;
  }
  return out;
}

function jpeg(b) {
  let p = 2;
  let dpiX = 0, dpiY = 0;
  while (p + 9 < b.length) {
    if (b[p] !== 0xff) { p++; continue; }
    const m = b[p + 1];
    if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { p += 2; continue; }
    const len = u16be(b, p + 2);
    if (m === 0xe0 && b[p + 4] === 0x4a && b[p + 5] === 0x46 && b[p + 6] === 0x49 && b[p + 7] === 0x46) {
      const unit = b[p + 11];
      const dx = u16be(b, p + 12);
      const dy = u16be(b, p + 14);
      if (unit === 1) { dpiX = dx; dpiY = dy; } else if (unit === 2) { dpiX = dx * 2.54; dpiY = dy * 2.54; }
    }
    if ((m >= 0xc0 && m <= 0xcf) && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
      return { height: u16be(b, p + 5), width: u16be(b, p + 7), dpiX: dpiX || undefined, dpiY: dpiY || undefined };
    }
    p += 2 + len;
  }
  return null;
}

function gif(b) {
  return b.length >= 10 ? { width: u16le(b, 6), height: u16le(b, 8) } : null;
}

function bmp(b) {
  if (b.length < 26) return null;
  const hs = i32le(b, 14);
  if (hs === 12) return { width: u16le(b, 18), height: u16le(b, 20) };
  const out = { width: Math.abs(i32le(b, 18)), height: Math.abs(i32le(b, 22)) };
  if (hs >= 40 && b.length >= 46) {
    const ppmX = i32le(b, 38), ppmY = i32le(b, 42);
    if (ppmX > 0 && ppmY > 0) { out.dpiX = ppmX * 0.0254; out.dpiY = ppmY * 0.0254; }
  }
  return out;
}

function webp(b) {
  if (b.length < 30) return null;
  const fmt = String.fromCharCode(b[12], b[13], b[14], b[15]);
  if (fmt === 'VP8 ') return { width: u16le(b, 26) & 0x3fff, height: u16le(b, 28) & 0x3fff };
  if (fmt === 'VP8L') {
    const bits = b[21] | (b[22] << 8) | (b[23] << 16) | (b[24] << 24);
    return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
  }
  if (fmt === 'VP8X') {
    return { width: 1 + (b[24] | (b[25] << 8) | (b[26] << 16)), height: 1 + (b[27] | (b[28] << 8) | (b[29] << 16)) };
  }
  return null;
}

function emf(b) {
  if (b.length < 88) return null;
  // rclFrame in 0.01 mm units.
  const l = i32le(b, 24), t = i32le(b, 28), r = i32le(b, 32), bt = i32le(b, 36);
  const w = ((r - l) / 100 / 25.4) * 96;
  const h = ((bt - t) / 100 / 25.4) * 96;
  if (w > 0 && h > 0) return { width: Math.round(w), height: Math.round(h) };
  const bl = i32le(b, 8), bt2 = i32le(b, 12), br = i32le(b, 16), bb = i32le(b, 20);
  return { width: br - bl, height: bb - bt2 };
}

function wmf(b) {
  if (b.length >= 22 && b[0] === 0xd7 && b[1] === 0xcd) {
    const l = (b[6] | (b[7] << 8)) << 16 >> 16;
    const t = (b[8] | (b[9] << 8)) << 16 >> 16;
    const r = (b[10] | (b[11] << 8)) << 16 >> 16;
    const bt = (b[12] | (b[13] << 8)) << 16 >> 16;
    const inch = u16le(b, 14) || 1440;
    return { width: Math.round(((r - l) / inch) * 96), height: Math.round(((bt - t) / inch) * 96) };
  }
  return null;
}

function svg(b) {
  let s = '';
  const n = Math.min(b.length, 4096);
  for (let i = 0; i < n; i++) s += String.fromCharCode(b[i]);
  const m = /<svg[^>]*>/i.exec(s);
  if (!m) return null;
  const tag = m[0];
  const num = (name) => {
    const r = new RegExp('\\s' + name + '\\s*=\\s*["\']\\s*([0-9.]+)\\s*(px|pt|in|cm|mm)?', 'i').exec(tag);
    if (!r) return 0;
    const v = parseFloat(r[1]);
    const u = (r[2] || 'px').toLowerCase();
    return u === 'pt' ? v * 96 / 72 : u === 'in' ? v * 96 : u === 'cm' ? v * 96 / 2.54 : u === 'mm' ? v * 96 / 25.4 : v;
  };
  let w = num('width'), h = num('height');
  if (!w || !h) {
    const vb = /viewBox\s*=\s*["']\s*([-0-9.]+)[\s,]+([-0-9.]+)[\s,]+([0-9.]+)[\s,]+([0-9.]+)/i.exec(tag);
    if (vb) { w = w || parseFloat(vb[3]); h = h || parseFloat(vb[4]); }
  }
  return w && h ? { width: w, height: h } : null;
}

function tiff(b) {
  const le = b[0] === 0x49;
  const r16 = (i) => (le ? b[i] | (b[i + 1] << 8) : (b[i] << 8) | b[i + 1]);
  const r32 = (i) => (le ? (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0 : u32be(b, i));
  if (b.length < 8) return null;
  const ifd = r32(4);
  if (ifd + 2 > b.length) return null;
  const n = r16(ifd);
  let w = 0, h = 0;
  for (let k = 0; k < n; k++) {
    const e = ifd + 2 + k * 12;
    if (e + 12 > b.length) break;
    const tag = r16(e);
    const type = r16(e + 2);
    const v = type === 3 ? r16(e + 8) : r32(e + 8);
    if (tag === 256) w = v;
    if (tag === 257) h = v;
  }
  return w && h ? { width: w, height: h } : null;
}

/** Return {width, height, dpiX?, dpiY?} for an image or null. */
export function imageSize(bytes, mime) {
  if (!bytes) return null;
  try {
    switch (mime) {
      case 'image/png': return png(bytes);
      case 'image/jpeg': return jpeg(bytes);
      case 'image/gif': return gif(bytes);
      case 'image/bmp': return bmp(bytes);
      case 'image/webp': return webp(bytes);
      case 'image/x-emf': return emf(bytes);
      case 'image/x-wmf': return wmf(bytes);
      case 'image/svg+xml': return svg(bytes);
      case 'image/tiff': return tiff(bytes);
      default: return null;
    }
  } catch (e) {
    return null;
  }
}

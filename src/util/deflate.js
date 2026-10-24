/**
 * Small DEFLATE (RFC 1951) compressor: LZ77 with hash chains and dynamic
 * Huffman codes per block, plus a zlib (RFC 1950) wrapper. Used to encode
 * PNG images (e.g. converted TIFF pictures).
 */

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

// Code lookup for lengths 3..258 and distances 1..32768.
const LEN_CODE = new Uint8Array(259);
for (let c = 0; c < 29; c++) for (let l = LEN_BASE[c]; l < (c < 28 ? LEN_BASE[c + 1] : 259); l++) LEN_CODE[l] = c;
function distCode(d) {
  let lo = 0, hi = 29;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (DIST_BASE[mid] <= d) lo = mid; else hi = mid - 1;
  }
  return lo;
}

class BitWriter {
  constructor(size) {
    this.buf = new Uint8Array(Math.max(1024, size));
    this.pos = 0;
    this.bit = 0;
    this.acc = 0;
  }

  ensure(n) {
    if (this.pos + n < this.buf.length) return;
    const nb = new Uint8Array(Math.max(this.buf.length * 2, this.pos + n + 1024));
    nb.set(this.buf.subarray(0, this.pos));
    this.buf = nb;
  }

  /** Write `n` bits of `v`, least significant first. */
  bits(v, n) {
    this.acc |= v << this.bit;
    this.bit += n;
    while (this.bit >= 8) {
      this.ensure(1);
      this.buf[this.pos++] = this.acc & 0xff;
      this.acc >>>= 8;
      this.bit -= 8;
    }
  }

  /** Write a Huffman code (stored MSB first). */
  code(c, n) {
    let r = 0;
    for (let i = 0; i < n; i++) { r = (r << 1) | (c & 1); c >>= 1; }
    this.bits(r, n);
  }

  align() {
    if (this.bit > 0) this.bits(0, 8 - this.bit);
  }

  bytes() {
    this.align();
    return this.buf.subarray(0, this.pos);
  }
}

/** Length-limited Huffman code lengths (package-merge free heuristic). */
function codeLengths(freq, maxBits) {
  const n = freq.length;
  const lens = new Uint8Array(n);
  const used = [];
  for (let i = 0; i < n; i++) if (freq[i]) used.push(i);
  if (used.length === 0) return lens;
  // A single symbol still needs a complete code: pair it with an unused one.
  if (used.length === 1) { lens[used[0]] = 1; lens[used[0] === 0 ? 1 : 0] = 1; return lens; }
  // Standard Huffman via a sorted merge of leaves and internal nodes.
  let nodes = used.map((s) => ({ f: freq[s], s, l: null, r: null }));
  nodes.sort((a, b) => a.f - b.f);
  const queue = [];
  let i = 0;
  const take = () => (i < nodes.length && (!queue.length || nodes[i].f <= queue[0].f) ? nodes[i++] : queue.shift());
  while (nodes.length - i + queue.length > 1) {
    const a = take(), b = take();
    queue.push({ f: a.f + b.f, s: -1, l: a, r: b });
  }
  const root = queue[0] || nodes[i];
  const walk = (nd, d) => {
    if (nd.s >= 0) { lens[nd.s] = Math.max(1, d); return; }
    walk(nd.l, d + 1);
    walk(nd.r, d + 1);
  };
  walk(root, 0);
  // Limit lengths: flatten overly long codes while keeping the Kraft sum <= 1.
  let over = false;
  for (const s of used) if (lens[s] > maxBits) { lens[s] = maxBits; over = true; }
  if (over) {
    const kraft = () => used.reduce((a, s) => a + 2 ** (maxBits - lens[s]), 0);
    let k = kraft();
    const limit = 2 ** maxBits;
    // Lengthen the shortest codes of the least frequent symbols until it fits.
    const order = used.slice().sort((a, b) => freq[a] - freq[b]);
    while (k > limit) {
      for (const s of order) {
        if (lens[s] < maxBits) { k -= 2 ** (maxBits - lens[s] - 1); lens[s]++; if (k <= limit) break; }
      }
    }
    // Decoders reject incomplete codes: give spare space to frequent symbols.
    while (k < limit) {
      let changed = false;
      for (let j = order.length - 1; j >= 0 && k < limit; j--) {
        const s = order[j];
        const gain = 2 ** (maxBits - lens[s]);
        if (lens[s] > 1 && k + gain <= limit) { lens[s]--; k += gain; changed = true; }
      }
      if (!changed) break;
    }
  }
  return lens;
}

/** Canonical codes for code lengths. */
function canonical(lens) {
  const max = Math.max(0, ...lens);
  const blCount = new Uint16Array(max + 1);
  for (const l of lens) if (l) blCount[l]++;
  const next = new Uint16Array(max + 2);
  let code = 0;
  for (let b = 1; b <= max; b++) { code = (code + blCount[b - 1]) << 1; next[b] = code; }
  const codes = new Uint16Array(lens.length);
  for (let s = 0; s < lens.length; s++) if (lens[s]) codes[s] = next[lens[s]]++;
  return codes;
}

/** Run-length encode code lengths with symbols 16/17/18. */
function rleLengths(lens) {
  const out = [];
  for (let i = 0; i < lens.length;) {
    const v = lens[i];
    let run = 1;
    while (i + run < lens.length && lens[i + run] === v) run++;
    let left = run;
    if (v === 0) {
      while (left >= 11) { const k = Math.min(138, left); out.push([18, k - 11, 7]); left -= k; }
      if (left >= 3) { out.push([17, left - 3, 3]); left = 0; }
    } else {
      out.push([v, 0, 0]);
      left--;
      while (left >= 3) { const k = Math.min(6, left); out.push([16, k - 3, 2]); left -= k; }
    }
    while (left-- > 0) out.push([v, 0, 0]);
    i += run;
  }
  return out;
}

function writeBlock(w, syms, final) {
  const lf = new Uint32Array(286), df = new Uint32Array(30);
  for (const s of syms) {
    if (s < 256) lf[s]++;
    else { lf[257 + LEN_CODE[s >>> 16 & 0x1ff]]++; df[distCode(s & 0xffff)]++; }
  }
  lf[256] = 1;
  const ll = codeLengths(lf, 15), dl = codeLengths(df, 15);
  if (!dl.some((x) => x)) dl[0] = 1;
  let hlit = 286; while (hlit > 257 && !ll[hlit - 1]) hlit--;
  let hdist = 30; while (hdist > 1 && !dl[hdist - 1]) hdist--;
  const all = new Uint8Array(hlit + hdist);
  all.set(ll.subarray(0, hlit), 0);
  all.set(dl.subarray(0, hdist), hlit);
  const rle = rleLengths(all);
  const cf = new Uint32Array(19);
  for (const [s] of rle) cf[s]++;
  const cl = codeLengths(cf, 7);
  let hclen = 19; while (hclen > 4 && !cl[CL_ORDER[hclen - 1]]) hclen--;
  const lc = canonical(ll), dc = canonical(dl), cc = canonical(cl);
  w.bits(final ? 1 : 0, 1);
  w.bits(2, 2);
  w.bits(hlit - 257, 5);
  w.bits(hdist - 1, 5);
  w.bits(hclen - 4, 4);
  for (let i = 0; i < hclen; i++) w.bits(cl[CL_ORDER[i]], 3);
  for (const [s, extra, n] of rle) {
    w.code(cc[s], cl[s]);
    if (n) w.bits(extra, n);
  }
  for (const s of syms) {
    if (s < 256) { w.code(lc[s], ll[s]); continue; }
    const len = s >>> 16 & 0x1ff, dist = s & 0xffff;
    const c = LEN_CODE[len];
    w.code(lc[257 + c], ll[257 + c]);
    if (LEN_EXTRA[c]) w.bits(len - LEN_BASE[c], LEN_EXTRA[c]);
    const d = distCode(dist);
    w.code(dc[d], dl[d]);
    if (DIST_EXTRA[d]) w.bits(dist - DIST_BASE[d], DIST_EXTRA[d]);
  }
  w.code(lc[256], ll[256]);
}

/** Raw DEFLATE compression. */
export function deflateRaw(data) {
  const n = data.length;
  const w = new BitWriter((n >> 1) + 64);
  const HBITS = 15, HSIZE = 1 << HBITS, WMASK = 32767;
  const head = new Int32Array(HSIZE).fill(-1);
  const prev = new Int32Array(32768);
  const hash = (i) => ((data[i] << 10) ^ (data[i + 1] << 5) ^ data[i + 2]) & (HSIZE - 1);
  let syms = [];
  const BLOCK = 1 << 16;
  let i = 0;
  // Symbols: literal byte, or (length << 16 | distance) with length >= 3.
  while (i < n) {
    let bestLen = 0, bestDist = 0;
    if (i + 2 < n) {
      const h = hash(i);
      let j = head[h], chain = 48;
      const maxLen = Math.min(258, n - i);
      while (j >= 0 && i - j <= 32768 && chain-- > 0) {
        if (data[j + bestLen] === data[i + bestLen] && data[j] === data[i]) {
          let l = 0;
          while (l < maxLen && data[j + l] === data[i + l]) l++;
          if (l > bestLen) { bestLen = l; bestDist = i - j; if (l === maxLen) break; }
        }
        const pj = prev[j & WMASK];
        if (pj >= j) break;
        j = pj;
      }
      prev[i & WMASK] = head[h];
      head[h] = i;
    }
    if (bestLen >= 3) {
      syms.push((bestLen << 16) | bestDist);
      // Index the skipped positions.
      for (let k = 1; k < bestLen; k++) {
        const p = i + k;
        if (p + 2 < n) { const h = hash(p); prev[p & WMASK] = head[h]; head[h] = p; }
      }
      i += bestLen;
    } else {
      syms.push(data[i]);
      i++;
    }
    if (syms.length >= BLOCK) { writeBlock(w, syms, false); syms = []; }
  }
  writeBlock(w, syms, true);
  return w.bytes();
}

function adler32(data) {
  let a = 1, b = 0;
  for (let i = 0; i < data.length;) {
    const end = Math.min(data.length, i + 3800);
    for (; i < end; i++) { a += data[i]; b += a; }
    a %= 65521; b %= 65521;
  }
  return ((b << 16) | a) >>> 0;
}

/** zlib stream (CMF/FLG header, DEFLATE data, Adler-32). */
export function zlibCompress(data) {
  const body = deflateRaw(data);
  const out = new Uint8Array(body.length + 6);
  out[0] = 0x78; out[1] = 0x9c;
  out.set(body, 2);
  const a = adler32(data);
  out[out.length - 4] = a >>> 24; out[out.length - 3] = (a >>> 16) & 0xff; out[out.length - 2] = (a >>> 8) & 0xff; out[out.length - 1] = a & 0xff;
  return out;
}

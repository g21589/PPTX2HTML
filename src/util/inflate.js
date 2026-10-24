/**
 * Raw DEFLATE (RFC 1951) decoder.
 *
 * Table driven: every Huffman code is looked up with a single read of
 * `maxBits` bits using a table indexed by the bit-reversed code.
 */

const LEN_BASE = new Uint16Array([3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258]);
const LEN_EXTRA = new Uint8Array([0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0]);
const DIST_BASE = new Uint16Array([1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577]);
const DIST_EXTRA = new Uint8Array([0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13]);
const CL_ORDER = new Uint8Array([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]);

const REV = new Uint16Array(32768);
for (let i = 0; i < 32768; i++) {
  let x = i;
  x = ((x & 0xaaaa) >>> 1) | ((x & 0x5555) << 1);
  x = ((x & 0xcccc) >>> 2) | ((x & 0x3333) << 2);
  x = ((x & 0xf0f0) >>> 4) | ((x & 0x0f0f) << 4);
  x = ((x & 0xff00) >>> 8) | ((x & 0x00ff) << 8);
  REV[i] = x >>> 1;
}

/**
 * Build a lookup table: index = next `maxBits` input bits (LSB first),
 * value = (symbol << 4) | codeLength.
 */
function buildTable(lengths, n) {
  let maxBits = 0;
  for (let i = 0; i < n; i++) if (lengths[i] > maxBits) maxBits = lengths[i];
  if (maxBits === 0) return { table: new Uint16Array(1), bits: 1 };
  const blCount = new Uint16Array(16);
  for (let i = 0; i < n; i++) blCount[lengths[i]]++;
  blCount[0] = 0;
  const nextCode = new Uint16Array(16);
  let code = 0;
  for (let b = 1; b <= 15; b++) {
    code = (code + blCount[b - 1]) << 1;
    nextCode[b] = code;
  }
  const size = 1 << maxBits;
  const table = new Uint16Array(size);
  for (let sym = 0; sym < n; sym++) {
    const len = lengths[sym];
    if (!len) continue;
    const c = nextCode[len]++;
    // Reverse the `len` bits of c (Huffman codes are stored MSB first).
    const r = REV[c << (15 - len)];
    const entry = (sym << 4) | len;
    const step = 1 << len;
    for (let k = r; k < size; k += step) table[k] = entry;
  }
  return { table, bits: maxBits };
}

let FIXED_LIT = null;
let FIXED_DIST = null;
function fixedTables() {
  if (!FIXED_LIT) {
    const l = new Uint8Array(288);
    for (let i = 0; i < 144; i++) l[i] = 8;
    for (let i = 144; i < 256; i++) l[i] = 9;
    for (let i = 256; i < 280; i++) l[i] = 7;
    for (let i = 280; i < 288; i++) l[i] = 8;
    FIXED_LIT = buildTable(l, 288);
    const d = new Uint8Array(30).fill(5);
    FIXED_DIST = buildTable(d, 30);
  }
  return [FIXED_LIT, FIXED_DIST];
}

/**
 * Inflate raw deflate data.
 * @param {Uint8Array} input
 * @param {number} [sizeHint] expected output size
 * @returns {Uint8Array}
 */
export function inflateRaw(input, sizeHint = 0) {
  let out = new Uint8Array(sizeHint > 0 ? sizeHint : Math.max(1024, input.length * 4));
  let op = 0;
  let pos = 0; // byte position
  let bitBuf = 0;
  let bitCnt = 0;
  const inLen = input.length;

  function need(n) {
    while (bitCnt < n) {
      bitBuf |= (pos < inLen ? input[pos] : 0) << bitCnt;
      pos++;
      bitCnt += 8;
    }
  }
  function bits(n) {
    need(n);
    const v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n;
    bitCnt -= n;
    return v;
  }
  function ensure(extra) {
    if (op + extra <= out.length) return;
    let ns = out.length * 2;
    while (ns < op + extra) ns *= 2;
    const n = new Uint8Array(ns);
    n.set(out.subarray(0, op));
    out = n;
  }
  function decodeSym(t) {
    need(t.bits);
    const e = t.table[bitBuf & ((1 << t.bits) - 1)];
    const l = e & 15;
    if (l === 0) throw new Error('inflate: invalid Huffman code');
    bitBuf >>>= l;
    bitCnt -= l;
    return e >>> 4;
  }

  let final = 0;
  while (!final) {
    if (pos > inLen + 4) throw new Error('inflate: unexpected end of data');
    final = bits(1);
    const type = bits(2);
    if (type === 0) {
      // Stored block: discard remaining bits of current byte.
      bitBuf = 0;
      bitCnt = 0;
      if (pos + 4 > inLen) throw new Error('inflate: truncated stored block');
      const len = input[pos] | (input[pos + 1] << 8);
      pos += 4;
      ensure(len);
      out.set(input.subarray(pos, pos + len), op);
      op += len;
      pos += len;
      continue;
    }
    let lit, dist;
    if (type === 1) {
      [lit, dist] = fixedTables();
    } else if (type === 2) {
      const hlit = bits(5) + 257;
      const hdist = bits(5) + 1;
      const hclen = bits(4) + 4;
      const clLens = new Uint8Array(19);
      for (let i = 0; i < hclen; i++) clLens[CL_ORDER[i]] = bits(3);
      const cl = buildTable(clLens, 19);
      const lens = new Uint8Array(hlit + hdist);
      let i = 0;
      while (i < hlit + hdist) {
        const sym = decodeSym(cl);
        if (sym < 16) {
          lens[i++] = sym;
        } else if (sym === 16) {
          if (i === 0) throw new Error('inflate: bad repeat');
          const prev = lens[i - 1];
          let r = 3 + bits(2);
          while (r--) lens[i++] = prev;
        } else if (sym === 17) {
          let r = 3 + bits(3);
          while (r--) lens[i++] = 0;
        } else {
          let r = 11 + bits(7);
          while (r--) lens[i++] = 0;
        }
      }
      lit = buildTable(lens.subarray(0, hlit), hlit);
      dist = buildTable(lens.subarray(hlit), hdist);
    } else {
      throw new Error('inflate: invalid block type');
    }

    const litT = lit.table, litB = lit.bits, litM = (1 << litB) - 1;
    const distT = dist.table, distB = dist.bits, distM = (1 << distB) - 1;
    for (;;) {
      // Inline literal/length decode.
      while (bitCnt < litB) {
        bitBuf |= (pos < inLen ? input[pos] : 0) << bitCnt;
        pos++;
        bitCnt += 8;
      }
      if (pos > inLen + 8) throw new Error('inflate: unexpected end of data');
      const e = litT[bitBuf & litM];
      const l = e & 15;
      if (l === 0) throw new Error('inflate: invalid literal code');
      bitBuf >>>= l;
      bitCnt -= l;
      const sym = e >>> 4;
      if (sym < 256) {
        if (op >= out.length) ensure(1);
        out[op++] = sym;
        continue;
      }
      if (sym === 256) break;
      const li = sym - 257;
      if (li >= 29) throw new Error('inflate: bad length symbol');
      const len = LEN_BASE[li] + (LEN_EXTRA[li] ? bits(LEN_EXTRA[li]) : 0);
      while (bitCnt < distB) {
        bitBuf |= (pos < inLen ? input[pos] : 0) << bitCnt;
        pos++;
        bitCnt += 8;
      }
      const de = distT[bitBuf & distM];
      const dl = de & 15;
      if (dl === 0) throw new Error('inflate: invalid distance code');
      bitBuf >>>= dl;
      bitCnt -= dl;
      const ds = de >>> 4;
      const d = DIST_BASE[ds] + (DIST_EXTRA[ds] ? bits(DIST_EXTRA[ds]) : 0);
      if (d > op) throw new Error('inflate: distance too far back');
      ensure(len);
      let from = op - d;
      if (d >= len) {
        out.copyWithin(op, from, from + len);
        op += len;
      } else {
        for (let k = 0; k < len; k++) out[op++] = out[from++];
      }
    }
    if (pos > inLen + 4) throw new Error('inflate: unexpected end of data');
  }
  return op === out.length ? out : out.subarray(0, op);
}

const CHARS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/** Encode bytes as base64 (works in browsers and Node). */
export function toBase64(bytes) {
  if (typeof Buffer !== 'undefined' && Buffer.from) {
    return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString('base64');
  }
  const n = bytes.length;
  const parts = [];
  const CHUNK = 3 * 4096;
  for (let start = 0; start < n; start += CHUNK) {
    const end = Math.min(n, start + CHUNK);
    let s = '';
    let i = start;
    for (; i + 2 < end; i += 3) {
      const v = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
      s += CHARS[v >> 18] + CHARS[(v >> 12) & 63] + CHARS[(v >> 6) & 63] + CHARS[v & 63];
    }
    if (i < end) {
      const rem = end - i;
      const v = (bytes[i] << 16) | (rem > 1 ? bytes[i + 1] << 8 : 0);
      s += CHARS[v >> 18] + CHARS[(v >> 12) & 63] + (rem > 1 ? CHARS[(v >> 6) & 63] : '=') + '=';
    }
    parts.push(s);
  }
  return parts.join('');
}

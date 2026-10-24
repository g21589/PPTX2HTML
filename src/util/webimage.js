/**
 * Convert picture formats browsers cannot display: EMF/WMF become SVG,
 * TIFF becomes PNG. Returns null when the bytes are kept as they are.
 */
import { emfToSvg } from './emf.js';
import { decodeTiff } from './tiff.js';
import { encodePng } from './png.js';

/**
 * @param {Uint8Array} bytes
 * @param {string} mime
 * @returns {{bytes: Uint8Array, mime: string, ext: string} | null}
 */
export function toWebImage(bytes, mime) {
  if (!bytes) return null;
  if (mime === 'image/x-emf' || mime === 'image/x-wmf') {
    const svg = emfToSvg(bytes, mime);
    return svg ? { bytes: new TextEncoder().encode(svg), mime: 'image/svg+xml', ext: 'svg' } : null;
  }
  if (mime === 'image/tiff') {
    let img = null;
    try { img = decodeTiff(bytes); } catch (e) { img = null; }
    return img ? { bytes: encodePng(img.width, img.height, img.rgba), mime: 'image/png', ext: 'png' } : null;
  }
  return null;
}

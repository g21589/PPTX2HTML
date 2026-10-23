/**
 * Embedded fonts (p:embeddedFontLst) to @font-face rules.
 *
 * PowerPoint stores embedded fonts as Embedded OpenType (EOT) data. EOT
 * files that are neither MTX-compressed nor XOR-obfuscated simply wrap a
 * TrueType font, which is extracted here. Obfuscated data is de-obfuscated.
 */
import { toBase64 } from '../util/base64.js';

function readEot(bytes) {
  if (bytes.length < 82) return null;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eotSize = v.getUint32(0, true);
  const fontDataSize = v.getUint32(4, true);
  const version = v.getUint32(8, true);
  const flags = v.getUint32(12, true);
  const magic = v.getUint16(34, true);
  if (magic !== 0x504c || eotSize > bytes.length + 16) return null;
  if (version !== 0x00010000 && version !== 0x00020001 && version !== 0x00020002) return null;
  if (flags & 0x4) return { compressed: true };
  const data = bytes.subarray(eotSize - fontDataSize, eotSize);
  let out = data;
  if (flags & 0x10000000) {
    out = new Uint8Array(data);
    for (let i = 0; i < out.length; i++) out[i] ^= 0x50;
  }
  return { font: out };
}

function sniffFont(bytes) {
  if (bytes.length < 4) return null;
  const tag = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
  if (tag === 'OTTO') return 'font/otf';
  if (tag === 'wOFF') return 'font/woff';
  if (tag === 'wOF2') return 'font/woff2';
  if ((bytes[0] === 0 && bytes[1] === 1 && bytes[2] === 0 && bytes[3] === 0) || tag === 'true') return 'font/ttf';
  return null;
}

export function fontFaceCss(pres, rc) {
  if (!pres.fonts || !pres.fonts.length || rc.opts.embedFonts === false) return '';
  let out = '';
  const pkg = pres.media.pkg;
  for (const f of pres.fonts) {
    for (const [variant, part] of Object.entries(f.variants)) {
      const bytes = pkg.bytes(part);
      if (!bytes) continue;
      let font = null;
      let mime = sniffFont(bytes);
      if (mime) font = bytes;
      else {
        const eot = readEot(bytes);
        if (eot && eot.font) {
          font = eot.font;
          mime = sniffFont(font) || 'font/ttf';
        } else {
          pres.diagnostics.add('approx', 'embedded font ' + f.typeface + ' uses compressed EOT data');
        }
      }
      if (!font) continue;
      const weight = variant === 'bold' || variant === 'boldItalic' ? 'bold' : 'normal';
      const style = variant === 'italic' || variant === 'boldItalic' ? 'italic' : 'normal';
      out += `@font-face{font-family:"${f.typeface.replace(/"/g, '')}";src:url(data:${mime};base64,${toBase64(font)});font-weight:${weight};font-style:${style};}`;
    }
  }
  return out;
}

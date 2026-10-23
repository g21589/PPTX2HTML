/**
 * Render context: unique ids, media URLs and options shared by renderers.
 */
import { toBase64 } from '../util/base64.js';
import { emfToSvg } from '../util/emf.js';
import { escapeXml } from '../util/xml.js';

/** Only allow harmless URL schemes for linked media. */
function safeUrl(u) {
  const s = String(u || '').trim();
  if (/^(https?:|data:image\/|file:|blob:)/i.test(s)) return s;
  if (/^[a-z][a-z0-9+.-]*:/i.test(s)) return '';
  return s.replace(/\\/g, '/');
}

export class RenderContext {
  /**
   * @param {object} pres presentation model
   * @param {object} opts
   *   mediaUrl(media) => string  custom media URL resolver
   *   idPrefix: prefix for generated ids (default "pp")
   */
  constructor(pres, opts = {}) {
    this.pres = pres;
    this.opts = opts;
    this.prefix = opts.idPrefix || 'pp';
    this.seq = 0;
    this.urls = new Map();
    this.slide = null;
  }

  uid(kind = 'd') {
    return this.prefix + '-' + kind + (this.seq++).toString(36);
  }

  /** URL for a media descriptor (data URI by default), escaped for use in attributes. */
  url(media) {
    if (!media) return '';
    if (media.external) return escapeXml(safeUrl(media.url));
    if (this.urls.has(media.id)) return this.urls.get(media.id);
    let url;
    if (this.opts.mediaUrl) url = this.opts.mediaUrl(media, this.pres.media);
    if (!url) {
      let bytes = this.pres.media.bytes(media.id);
      let mime = media.mime;
      if ((mime === 'image/x-emf' || mime === 'image/x-wmf') && bytes && this.opts.convertMetafiles !== false) {
        const svg = emfToSvg(bytes, mime);
        if (svg) {
          bytes = new TextEncoder().encode(svg);
          mime = 'image/svg+xml';
        }
      }
      url = bytes ? 'data:' + mime + ';base64,' + toBase64(bytes) : '';
    } else url = escapeXml(url);
    this.urls.set(media.id, url);
    return url;
  }
}

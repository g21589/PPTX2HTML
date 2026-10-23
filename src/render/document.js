/**
 * Complete HTML document assembly.
 */
import { escapeXml } from '../util/xml.js';
import { RenderContext } from './context.js';
import { slideHtml } from './slide.js';
import { BASE_CSS } from './styles.js';
import { fontFaceCss } from './embedded-fonts.js';

/** Data consumed by the runtime player (timing, transitions, notes...). */
export function playerData(pres) {
  return {
    width: pres.width,
    height: pres.height,
    show: pres.show,
    customShows: pres.customShows,
    slides: pres.slides.map((s) => ({
      index: s.index,
      number: s.number,
      hidden: !!s.hidden,
      name: s.name || '',
      transition: s.transition || null,
      timing: s.timing || null,
      notes: s.notes || null,
    })),
  };
}

function jsonScript(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

/**
 * Render all slides.
 * @returns {{slides: string[], css: string, data: object, rc: RenderContext}}
 */
export function renderSlides(pres, opts = {}) {
  const rc = new RenderContext(pres, opts);
  const slides = pres.slides.map((s) => slideHtml(s, rc));
  const css = BASE_CSS + fontFaceCss(pres, rc);
  return { slides, css, data: playerData(pres), rc };
}

/**
 * Render a standalone HTML document.
 * @param {object} pres presentation model
 * @param {object} opts
 *   title: document title
 *   player: runtime script source (string) to embed; omit for a static page
 *   mode: 'deck' (presentation) or 'list' (all slides stacked)
 */
export function renderDocument(pres, opts = {}) {
  const r = renderSlides(pres, opts);
  const title = opts.title || pres.props.title || 'Presentation';
  const mode = opts.mode || (opts.player ? 'deck' : 'list');
  const listCss = `
body.pp-body{margin:0;background:#3a3a3a;font-family:system-ui,sans-serif}
.pp-list .pp-slide{margin:24px auto;box-shadow:0 2px 12px rgba(0,0,0,.4)}
`;
  const deck = `<div class="pp-deck pp-${mode}" data-width="${pres.width}" data-height="${pres.height}">${r.slides.join('\n')}</div>`;
  const data = `<script type="application/json" id="pp-data">${jsonScript(r.data)}</script>`;
  const player = opts.player ? `<script>${opts.player}\n</script>` : '';
  return `<!DOCTYPE html>
<html lang="${escapeXml(opts.lang || 'en')}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="PPTX2HTML">
<title>${escapeXml(title)}</title>
<style>${r.css}${listCss}${opts.extraCss || ''}</style>
</head>
<body class="pp-body">
${deck}
${data}
${player}
</body>
</html>
`;
}

/**
 * PPTX2HTML - convert PowerPoint presentations to HTML.
 *
 *   import { convert } from 'pptx2html';
 *   const { html } = convert(bytes);
 */
import { loadPresentation } from './pptx/presentation.js';
import { renderDocument, renderSlides, playerData } from './render/document.js';
import { PLAYER_SOURCE } from './generated/player-source.js';

export { loadPresentation, renderDocument, renderSlides, playerData };
export { BASE_CSS } from './render/styles.js';

/**
 * Convert a PPTX file to a standalone HTML document.
 * @param {ArrayBuffer|Uint8Array} data PPTX bytes
 * @param {object} [options]
 * @param {'deck'|'list'} [options.mode] presentation player (default) or a static list of slides
 * @param {boolean} [options.player] embed the runtime player (default true)
 * @param {string} [options.title] document title
 * @param {(media: object) => string} [options.mediaUrl] custom URL for images/media (default: data URIs)
 * @returns {{html: string, presentation: object, diagnostics: object[]}}
 */
export function convert(data, options = {}) {
  const pres = loadPresentation(data, options);
  const withPlayer = options.player !== false;
  const html = renderDocument(pres, {
    ...options,
    player: withPlayer ? PLAYER_SOURCE : null,
    mode: options.mode || (withPlayer ? 'deck' : 'list'),
  });
  return { html, presentation: pres, diagnostics: pres.diagnostics.list() };
}

export const version = '1.0.0';

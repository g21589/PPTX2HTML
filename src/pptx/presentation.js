/**
 * Presentation loader: builds the complete document model of a PPTX file.
 */
import { OpcPackage } from '../opc/package.js';
import { Diagnostics, MediaRegistry } from './context.js';
import { listStyle } from './text.js';
import { SlideBuilder, makeDocHelpers } from './slide.js';
import { emu2px } from '../util/units.js';

/**
 * @typedef {object} LoadOptions
 * @property {boolean} [includeHidden] render hidden shapes
 * @property {boolean} [liveDates] update date fields to the current date (default true)
 * @property {boolean} [strict] rethrow parse errors
 * @property {number[]} [slides] only build these (0-based) slide indices
 */

/**
 * Parse a PPTX file into a presentation model.
 * @param {ArrayBuffer|Uint8Array} data
 * @param {LoadOptions} [options]
 */
export function loadPresentation(data, options = {}) {
  const t0 = Date.now();
  const pkg = new OpcPackage(data);
  const diag = new Diagnostics();
  const doc = { pkg, diag, options, media: new MediaRegistry(pkg) };
  makeDocHelpers(doc);
  const main = pkg.mainPart();
  if (!main) throw new Error('Not a PowerPoint presentation (presentation part not found)');
  doc.mainPart = main;
  const pres = pkg.xml(main);
  if (!pres) throw new Error('Presentation part is not readable');

  const sldSz = pres.child('p:sldSz');
  const width = emu2px(sldSz ? sldSz.num('cx', 9144000) : 9144000);
  const height = emu2px(sldSz ? sldSz.num('cy', 6858000) : 6858000);
  doc.slideSize = { width, height };
  doc.defaultTextStyle = listStyle(pres.child('p:defaultTextStyle'));

  // Masters in declaration order (also loads their themes).
  const masterIds = pres.child('p:sldMasterIdLst');
  for (const m of masterIds ? masterIds.all('p:sldMasterId') : []) {
    const r = pkg.rel(main, m.attrs['r:id']);
    if (r && pkg.has(r.partName)) doc.master(r.partName);
  }

  // Slide list.
  const slideInfos = [];
  const sldIdLst = pres.child('p:sldIdLst');
  let number = pres.int('firstSlideNum', 1);
  for (const s of sldIdLst ? sldIdLst.all('p:sldId') : []) {
    const r = pkg.rel(main, s.attrs['r:id']);
    if (!r || !pkg.has(r.partName)) {
      diag.add('missing', 'slide part ' + (r ? r.partName : s.attrs['r:id']));
      continue;
    }
    slideInfos.push({ part: r.partName, index: slideInfos.length, number: number++, sldId: s.attrs.id });
  }

  // Sections (PowerPoint 2010+).
  const sections = [];
  for (const sec of pres.find('p14:section')) {
    const ids = sec.find('p14:sldId').map((x) => x.attrs.id);
    sections.push({ name: sec.attr('name', ''), slides: slideInfos.filter((s) => ids.includes(s.sldId)).map((s) => s.index) });
  }

  // Custom shows.
  const customShows = [];
  const csl = pres.child('p:custShowLst');
  for (const cs of csl ? csl.all('p:custShow') : []) {
    const parts = [];
    for (const sld of cs.find('p:sld')) {
      const r = pkg.rel(main, sld.attrs['r:id']);
      if (r) {
        const info = slideInfos.find((s) => s.part === r.partName);
        if (info) parts.push(info.index);
      }
    }
    customShows.push({ id: cs.int('id', 0), name: cs.attr('name', ''), slides: parts });
  }

  // Show settings.
  const show = { loop: false, kiosk: false, useTimings: true, showNarration: true, showAnimation: true };
  const pp = pkg.relByType(main, 'presProps');
  const presProps = pp ? pkg.xml(pp.partName) : null;
  const showPr = presProps && presProps.child('p:showPr');
  if (showPr) {
    show.loop = showPr.bool('loop', false);
    show.useTimings = showPr.bool('useTimings', true);
    show.showNarration = showPr.bool('showNarration', true);
    show.showAnimation = showPr.bool('showAnimation', true);
    show.kiosk = !!showPr.child('p:kiosk');
    const kiosk = showPr.child('p:kiosk');
    if (kiosk) show.restart = kiosk.int('restart', 300000);
  }

  // Embedded fonts.
  const fonts = [];
  const efl = pres.child('p:embeddedFontLst');
  for (const ef of efl ? efl.all('p:embeddedFont') : []) {
    const f = ef.child('p:font');
    if (!f) continue;
    const face = { typeface: f.attr('typeface', ''), variants: {} };
    for (const v of ['regular', 'bold', 'italic', 'boldItalic']) {
      const n = ef.child('p:' + v);
      if (!n) continue;
      const r = pkg.rel(main, n.attrs['r:id']);
      if (r && pkg.has(r.partName)) face.variants[v] = r.partName;
    }
    fonts.push(face);
  }

  // Document properties.
  const props = {};
  const coreRel = pkg.relByType('/', 'core-properties');
  const core = coreRel ? pkg.xml(coreRel.partName) : null;
  if (core) {
    for (const k of ['dc:title', 'dc:creator', 'dc:subject', 'cp:keywords', 'dc:description', 'cp:lastModifiedBy', 'dcterms:created', 'dcterms:modified']) {
      const n = core.child(k);
      if (n && n.text.trim()) props[k.split(':')[1]] = n.text.trim();
    }
  }
  const thumbRel = pkg.relByType('/', 'thumbnail');

  const only = options.slides ? new Set(options.slides) : null;
  const slides = [];
  for (const info of slideInfos) {
    if (only && !only.has(info.index)) continue;
    try {
      slides.push(new SlideBuilder(doc, info).build());
    } catch (e) {
      diag.add('error', 'slide ' + (info.index + 1) + ': ' + e.message, info.part);
      if (options.strict) throw e;
      slides.push({ index: info.index, number: info.number, part: info.part, elements: [], background: null, error: e.message });
    }
  }
  // Resolve slide jump hyperlinks to slide indices.
  const partToIndex = new Map(slideInfos.map((s) => [s.part, s.index]));
  const resolveLinks = (els) => {
    for (const el of els) {
      for (const l of [el.link, el.hover]) if (l && l.slidePart) l.slide = partToIndex.has(l.slidePart) ? partToIndex.get(l.slidePart) : null;
      if (el.text) for (const p of el.text.paragraphs) for (const r of p.runs) {
        for (const l of [r.link, r.hover]) if (l && l.slidePart) l.slide = partToIndex.has(l.slidePart) ? partToIndex.get(l.slidePart) : null;
      }
      if (el.children) resolveLinks(el.children);
      if (el.table) for (const row of el.table.rows) for (const c of row.cells) {
        if (c.text) for (const p of c.text.paragraphs) for (const r of p.runs) if (r.link && r.link.slidePart) r.link.slide = partToIndex.get(r.link.slidePart);
      }
    }
  };
  for (const s of slides) resolveLinks(s.elements || []);

  return {
    width,
    height,
    slides,
    slideCount: slideInfos.length,
    sections,
    customShows,
    show,
    fonts,
    props,
    thumbnail: thumbRel && pkg.has(thumbRel.partName) ? doc.media.part(thumbRel.partName) : null,
    media: doc.media,
    diagnostics: diag,
    parseTime: Date.now() - t0,
  };
}

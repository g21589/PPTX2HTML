/**
 * PPTX deck builder for test fixtures. Writes raw PresentationML so that
 * every feature (including animations and transitions) can be exercised.
 */
import { zip } from './zipwriter.mjs';
import { NS, NS_EXT, themeXml, masterXml, layoutXml } from './templates.mjs';

export const EMU = 9525; // per px
export const px = (v) => Math.round(v * EMU);
export const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const LAYOUTS = ['title', 'content', 'titleOnly', 'blank', 'twoContent'];

const MIME = {
  png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp',
  wav: 'audio/wav', mp3: 'audio/mpeg', mp4: 'video/mp4', webm: 'video/webm', emf: 'image/x-emf', wmf: 'image/x-wmf',
  xml: 'application/xml', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

class SlideBuilder {
  constructor(deck, index, o) {
    this.deck = deck;
    this.index = index;
    this.layout = o.layout || 'blank';
    this.bg = o.bg || '';
    this.items = [];
    this.rels = [];
    this.timing = o.timing || '';
    this.transition = o.transition || '';
    this.hidden = !!o.hidden;
    this.notes = o.notes || null;
    this.nextId = 2;
    this.clrMapOvr = o.clrMapOvr || '';
    this.showMasterSp = o.showMasterSp;
    this.name = o.name || '';
  }

  id() {
    return this.nextId++;
  }

  add(xml) {
    this.items.push(xml);
    return this;
  }

  rel(type, target, external = false) {
    const id = 'rId' + (this.rels.length + 2);
    this.rels.push({ id, type, target, external });
    return id;
  }

  image(bytes, ext) {
    const name = this.deck.media(bytes, ext);
    return this.rel(REL + '/image', '../media/' + name);
  }

  mediaFile(bytes, ext, kind) {
    const name = this.deck.media(bytes, ext);
    const target = '../media/' + name;
    const r1 = this.rel(REL + '/' + (kind === 'video' ? 'video' : 'audio'), target);
    const r2 = this.rel('http://schemas.microsoft.com/office/2007/relationships/media', target);
    return { link: r1, embed: r2 };
  }

  hyperlink(url) {
    return this.rel(REL + '/hyperlink', url, true);
  }

  slideLink(n) {
    return this.rel(REL + '/slide', `slide${n}.xml`);
  }

  chart(xml) {
    const n = this.deck.charts.push(xml);
    return this.rel(REL + '/chart', `../charts/chart${n}.xml`);
  }

  part(path, xml, type, contentType) {
    this.deck.extraParts.push({ path: 'ppt/' + path, xml, contentType });
    return this.rel(type, '../' + path);
  }

  xml() {
    const show = this.showMasterSp === false ? ' showMasterSp="0"' : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_EXT}${this.hidden ? ' show="0"' : ''}${show}><p:cSld${this.name ? ` name="${esc(this.name)}"` : ''}>${this.bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${this.items.join('')}</p:spTree></p:cSld><p:clrMapOvr>${this.clrMapOvr || '<a:masterClrMapping/>'}</p:clrMapOvr>${this.transition}${this.timing}</p:sld>`;
  }
}

export class Deck {
  constructor(o = {}) {
    this.W = o.width || 12192000;
    this.H = o.height || 6858000;
    this.theme = o.theme || {};
    this.slides = [];
    this.mediaFiles = [];
    this.charts = [];
    this.extraParts = [];
    this.masterShapes = o.masterShapes || '';
    this.title = o.title || 'Fixture';
    this.tableStyles = o.tableStyles || '';
    this.defaultTextStyle = o.defaultTextStyle || '';
  }

  slide(o = {}) {
    const s = new SlideBuilder(this, this.slides.length + 1, o);
    this.slides.push(s);
    return s;
  }

  media(bytes, ext) {
    const name = `media${this.mediaFiles.length + 1}.${ext}`;
    this.mediaFiles.push({ name, bytes, ext });
    return name;
  }

  build() {
    const files = [];
    const ct = new Map([
      ['/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
      ['/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
      ['/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
      ['/ppt/presProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'],
      ['/ppt/viewProps.xml', 'application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml'],
      ['/ppt/tableStyles.xml', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'],
      ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
      ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
    ]);
    LAYOUTS.forEach((l, i) => ct.set(`/ppt/slideLayouts/slideLayout${i + 1}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'));
    const add = (name, data) => files.push({ name, data });

    add('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
    add('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(this.title)}</dc:title><dc:creator>PPTX2HTML fixtures</dc:creator></cp:coreProperties>`);
    add('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>${this.slides.length}</Slides></Properties>`);

    const presRels = [];
    presRels.push(`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`);
    this.slides.forEach((s, i) => presRels.push(`<Relationship Id="rId${i + 10}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`));
    presRels.push(`<Relationship Id="rId2" Type="${REL}/theme" Target="theme/theme1.xml"/>`);
    presRels.push(`<Relationship Id="rId3" Type="${REL}/presProps" Target="presProps.xml"/>`);
    presRels.push(`<Relationship Id="rId4" Type="${REL}/viewProps" Target="viewProps.xml"/>`);
    presRels.push(`<Relationship Id="rId5" Type="${REL}/tableStyles" Target="tableStyles.xml"/>`);
    add('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presRels.join('')}</Relationships>`);
    const dts = this.defaultTextStyle || `<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => `<a:lvl${i}pPr marL="${(i - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i}pPr>`).join('')}`;
    add('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${this.slides.map((s, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 10}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${this.W}" cy="${this.H}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle>${dts}</p:defaultTextStyle></p:presentation>`);
    add('ppt/presProps.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr ${NS}/>`);
    add('ppt/viewProps.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr ${NS}/>`);
    add('ppt/tableStyles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${this.tableStyles}</a:tblStyleLst>`);
    add('ppt/theme/theme1.xml', themeXml(this.theme));
    add('ppt/slideMasters/slideMaster1.xml', masterXml(this.W, this.H, {
      layoutIds: LAYOUTS.map((l, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join(''),
      masterShapes: this.masterShapes,
    }));
    add('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${LAYOUTS.map((l, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`).join('')}<Relationship Id="rId${LAYOUTS.length + 1}" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>`);
    LAYOUTS.forEach((l, i) => {
      add(`ppt/slideLayouts/slideLayout${i + 1}.xml`, layoutXml(l, this.W, this.H));
      add(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
    });
    this.slides.forEach((s, i) => {
      const n = i + 1;
      ct.set(`/ppt/slides/slide${n}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
      const xml = s.xml();
      add(`ppt/slides/slide${n}.xml`, xml);
      const li = Math.max(0, LAYOUTS.indexOf(s.layout)) + 1;
      const rels = [`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout${li}.xml"/>`];
      for (const r of s.rels) rels.push(`<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`);
      if (s.notes) {
        ct.set(`/ppt/notesSlides/notesSlide${n}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
        rels.push(`<Relationship Id="rIdN" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide${n}.xml"/>`);
        add(`ppt/notesSlides/notesSlide${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${esc(s.notes)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`);
        add(`ppt/notesSlides/_rels/notesSlide${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slide" Target="../slides/slide${n}.xml"/></Relationships>`);
      }
      add(`ppt/slides/_rels/slide${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
    });
    this.charts.forEach((xml, i) => {
      ct.set(`/ppt/charts/chart${i + 1}.xml`, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      add(`ppt/charts/chart${i + 1}.xml`, xml);
    });
    for (const p of this.extraParts) {
      if (p.contentType) ct.set('/' + p.path, p.contentType);
      add(p.path, p.xml);
    }
    for (const m of this.mediaFiles) add('ppt/media/' + m.name, m.bytes);
    const exts = new Set(this.mediaFiles.map((m) => m.ext));
    const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...exts].filter((e) => e !== 'xml').map((e) => `<Default Extension="${e}" ContentType="${MIME[e] || 'application/octet-stream'}"/>`).join('')}${[...ct.entries()].map(([k, v]) => `<Override PartName="${k}" ContentType="${v}"/>`).join('')}</Types>`;
    files.unshift({ name: '[Content_Types].xml', data: ctXml });
    return zip(files);
  }
}

// ---------------------------------------------------------------------------
// Shape helpers (positions in px).

export function xfrm(x, y, w, h, o = {}) {
  return `<a:xfrm${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="${px(x)}" y="${px(y)}"/><a:ext cx="${px(w)}" cy="${px(h)}"/></a:xfrm>`;
}

export const solid = (hex, mods = '') => `<a:solidFill><a:srgbClr val="${hex}">${mods}</a:srgbClr></a:solidFill>`;
export const scheme = (name, mods = '') => `<a:solidFill><a:schemeClr val="${name}">${mods}</a:schemeClr></a:solidFill>`;
export const ln = (w, fill = solid('000000'), extra = '') => `<a:ln w="${Math.round(w * 12700)}">${fill}${extra}</a:ln>`;

let fieldSeq = 0;

/** Text body from a simple description. */
export function txBody(paras, o = {}) {
  const body = o.bodyPr || '<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>';
  const ps = (Array.isArray(paras) ? paras : [paras]).map((p) => {
    if (typeof p === 'string') p = { runs: [{ t: p }] };
    const runs = (p.runs || []).map((r) => {
      if (r.br) return `<a:br>${r.rPr ? `<a:rPr ${r.rPr}/>` : ''}</a:br>`;
      if (r.fld) return `<a:fld id="{B6F15528-21DE-4FAA-801E-${(++fieldSeq).toString(16).toUpperCase().padStart(12, '0')}}" type="${r.fld}"><a:rPr lang="en-US" ${r.rPr || ''}>${r.inner || ''}</a:rPr><a:t>${esc(r.t || '')}</a:t></a:fld>`;
      return `<a:r><a:rPr lang="${r.lang || 'en-US'}" dirty="0" ${r.rPr || ''}>${r.inner || ''}</a:rPr><a:t>${esc(r.t)}</a:t></a:r>`;
    }).join('');
    return `<a:p>${p.pPr !== undefined ? `<a:pPr ${p.pPr}>${p.pInner || ''}</a:pPr>` : (p.pInner ? `<a:pPr>${p.pInner}</a:pPr>` : '')}${runs}<a:endParaRPr lang="en-US" dirty="0" ${p.endRPr || ''}/></a:p>`;
  }).join('');
  return `<p:txBody>${body}<a:lstStyle>${o.lstStyle || ''}</a:lstStyle>${ps}</p:txBody>`;
}

/**
 * Generic shape.
 * o: {id, name, x, y, w, h, prst, av, geom (custom xml), fill, line, effects, text, style, rot, flipH, flipV, ph, extraSpPr, hlink, descr, useBgFill}
 */
export function shape(o) {
  const geom = o.geom || `<a:prstGeom prst="${o.prst || 'rect'}"><a:avLst>${o.av || ''}</a:avLst></a:prstGeom>`;
  const style = o.style === undefined ? '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>' : o.style;
  const nvPr = o.ph ? `<p:nvPr><p:ph ${o.ph}/></p:nvPr>` : '<p:nvPr/>';
  const tag = o.cxn ? 'cxnSp' : 'sp';
  const nv = o.cxn
    ? `<p:nvCxnSpPr><p:cNvPr id="${o.id}" name="${esc(o.name || 'Connector ' + o.id)}"/><p:cNvCxnSpPr/>${nvPr}</p:nvCxnSpPr>`
    : `<p:nvSpPr><p:cNvPr id="${o.id}" name="${esc(o.name || 'Shape ' + o.id)}"${o.descr ? ` descr="${esc(o.descr)}"` : ''}${o.hidden ? ' hidden="1"' : ''}>${o.hlink || ''}</p:cNvPr><p:cNvSpPr${o.txBox ? ' txBox="1"' : ''}/>${nvPr}</p:nvSpPr>`;
  const spPr = o.ph && o.noXfrm ? `<p:spPr>${o.fill || ''}${o.line || ''}</p:spPr>` : `<p:spPr>${xfrm(o.x, o.y, o.w, o.h, o)}${geom}${o.fill || ''}${o.line || ''}${o.effects || ''}${o.extraSpPr || ''}</p:spPr>`;
  const text = o.text === undefined ? '' : (typeof o.text === 'string' && o.text.startsWith('<p:txBody') ? o.text : txBody(o.text, o));
  return `<p:${tag}${o.useBgFill ? ' useBgFill="1"' : ''}>${nv}${spPr}${o.cxn ? '' : style}${o.cxn ? (o.style || '') : ''}${text}</p:${tag}>`;
}

/** Text box (no fill / line, black text). */
export function textBox(o) {
  return shape({ style: '', txBox: true, fill: '<a:noFill/>', ...o, text: o.text, bodyPr: o.bodyPr || '<a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>' });
}

/** Slide title text box at the top (used to label fixture slides). */
export function label(s, text, o = {}) {
  return s.add(textBox({ id: s.id(), name: 'Label', x: 20, y: 8, w: 1240, h: 40, text: [{ runs: [{ t: text, rPr: `sz="${o.sz || 2000}" b="1"` }] }] }));
}

export function picture(o) {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${o.id}" name="${esc(o.name || 'Picture ' + o.id)}"${o.descr ? ` descr="${esc(o.descr)}"` : ''}>${o.hlink || ''}</p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr>${o.nvPrInner || ''}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${o.rid}"${o.link ? ` r:link="${o.link}"` : ''}>${o.blipInner || ''}</a:blip>${o.srcRect || ''}${o.tile || '<a:stretch><a:fillRect/></a:stretch>'}</p:blipFill><p:spPr>${xfrm(o.x, o.y, o.w, o.h, o)}<a:prstGeom prst="${o.prst || 'rect'}"><a:avLst/></a:prstGeom>${o.line || ''}${o.effects || ''}</p:spPr></p:pic>`;
}

export function group(o, children) {
  const off = o.chOff || [o.x, o.y];
  const ext = o.chExt || [o.w, o.h];
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${o.id}" name="${esc(o.name || 'Group ' + o.id)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="${px(o.x)}" y="${px(o.y)}"/><a:ext cx="${px(o.w)}" cy="${px(o.h)}"/><a:chOff x="${px(off[0])}" y="${px(off[1])}"/><a:chExt cx="${px(ext[0])}" cy="${px(ext[1])}"/></a:xfrm>${o.fill || ''}</p:grpSpPr>${children.join('')}</p:grpSp>`;
}

export function graphicFrame(o, graphic) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${o.id}" name="${esc(o.name || 'Frame ' + o.id)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${px(o.x)}" y="${px(o.y)}"/><a:ext cx="${px(o.w)}" cy="${px(o.h)}"/></p:xfrm><a:graphic>${graphic}</a:graphic></p:graphicFrame>`;
}

export { NS, NS_EXT };

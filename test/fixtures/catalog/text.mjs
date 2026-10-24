/**
 * Fixture catalog: text.
 */
import { shape, textBox, label, solid, ln } from '../lib.mjs';

export const name = 'text';
export const title = 'Text formatting and layout';

const box = (s, o) => s.add(shape({ style: '', fill: solid('F2F2F2'), line: ln(0.75, solid('A6A6A6')), ...o, id: s.id() }));

export function build(deck) {
  const out = [];
  // --- Run formatting -----------------------------------------------------------
  {
    const s = deck.slide({ name: 'Run formatting' });
    label(s, 'Character formatting');
    const runs = [
      { t: 'Bold ', rPr: 'b="1"' }, { t: 'Italic ', rPr: 'i="1"' }, { t: 'Underline ', rPr: 'u="sng"' },
      { t: 'Double ', rPr: 'u="dbl"' }, { t: 'Wavy ', rPr: 'u="wavy"' }, { t: 'Dotted ', rPr: 'u="dotted"' },
      { t: 'Strike ', rPr: 'strike="sngStrike"' }, { t: 'DblStrike ', rPr: 'strike="dblStrike"' },
    ];
    box(s, { name: 'decor', x: 40, y: 70, w: 1200, h: 60, text: [{ runs: runs.map((r) => ({ ...r, rPr: r.rPr + ' sz="2400"' })) }], bodyPr: '<a:bodyPr anchor="ctr"/>' });
    box(s, { name: 'sizes', x: 40, y: 140, w: 1200, h: 90, text: [{ runs: [8, 12, 18, 24, 36, 54].map((z) => ({ t: z + 'pt ', rPr: `sz="${z * 100}"` })) }], bodyPr: '<a:bodyPr anchor="ctr"/>' });
    box(s, {
      name: 'colors', x: 40, y: 240, w: 1200, h: 60, bodyPr: '<a:bodyPr anchor="ctr"/>',
      text: [{ runs: [
        { t: 'Red ', rPr: 'sz="2800"', inner: solid('FF0000') },
        { t: 'Accent1 ', rPr: 'sz="2800"', inner: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' },
        { t: 'Gradient ', rPr: 'sz="2800" b="1"', inner: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>' },
        { t: 'Outline ', rPr: 'sz="2800" b="1"', inner: '<a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln><a:noFill/>' },
        { t: 'Highlight ', rPr: 'sz="2800"', inner: '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>' },
        { t: 'Shadow', rPr: 'sz="2800" b="1"', inner: `${solid('4472C4')}<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>` },
      ] }],
    });
    box(s, {
      name: 'misc', x: 40, y: 310, w: 1200, h: 60, bodyPr: '<a:bodyPr anchor="ctr"/>',
      text: [{ runs: [
        { t: 'H', rPr: 'sz="2400"' }, { t: '2', rPr: 'sz="2400" baseline="-25000"' }, { t: 'O and E=mc', rPr: 'sz="2400"' }, { t: '2', rPr: 'sz="2400" baseline="30000"' },
        { t: '   all caps   ', rPr: 'sz="2400" cap="all"' }, { t: 'Small Caps', rPr: 'sz="2400" cap="small"' },
        { t: '   S p a c e d', rPr: 'sz="2400" spc="600"' }, { t: '   Condensed', rPr: 'sz="2400" spc="-300"' },
      ] }],
    });
    box(s, {
      name: 'fonts', x: 40, y: 380, w: 1200, h: 70, bodyPr: '<a:bodyPr anchor="ctr"/>',
      text: [{ runs: [
        { t: 'Calibri ', rPr: 'sz="2400"', inner: '<a:latin typeface="Calibri"/>' },
        { t: 'Times New Roman ', rPr: 'sz="2400"', inner: '<a:latin typeface="Times New Roman"/>' },
        { t: 'Courier New ', rPr: 'sz="2400"', inner: '<a:latin typeface="Courier New"/>' },
        { t: 'Arial ', rPr: 'sz="2400"', inner: '<a:latin typeface="Arial"/>' },
        { t: 'Major font ', rPr: 'sz="2400"', inner: '<a:latin typeface="+mj-lt"/>' },
        { t: '中文字型 日本語 한국어', rPr: 'sz="2400" altLang="en-US"', lang: 'zh-TW', inner: '<a:ea typeface="+mn-ea"/>' },
      ] }],
    });
    box(s, {
      name: 'symbols', x: 40, y: 460, w: 1200, h: 70, bodyPr: '<a:bodyPr anchor="ctr"/>',
      text: [{ runs: [
        { t: 'Wingdings: ', rPr: 'sz="2400"' }, { t: 'Øü§vq', rPr: 'sz="2400"', inner: '<a:sym typeface="Wingdings"/><a:latin typeface="Wingdings"/>' },
        { t: '   Symbol: ', rPr: 'sz="2400"' }, { t: 'abgdp', rPr: 'sz="2400"', inner: '<a:latin typeface="Symbol"/><a:sym typeface="Symbol"/>' },
        { t: '   Emoji: 😀🚀', rPr: 'sz="2400"' },
      ] }],
    });
    box(s, {
      name: 'fields', x: 40, y: 540, w: 1200, h: 60, bodyPr: '<a:bodyPr anchor="ctr"/>',
      text: [{ runs: [{ t: 'Slide number: ', rPr: 'sz="2000"' }, { fld: 'slidenum', t: '‹#›', rPr: 'sz="2000"' }, { t: '   Date: ', rPr: 'sz="2000"' }, { fld: 'datetime1', t: '1/1/2020', rPr: 'sz="2000"' }] }],
    });
    out.push({
      slide: s.index, feature: 'text/run-formatting', compareLO: true,
      checks: [
        { kind: 'text', contains: 'Bold' },
        { kind: 'dom', selector: '.pp-p span[style*="font-weight:bold"]', min: 1 },
        { kind: 'dom', selector: '.pp-p span[style*="font-style:italic"]', min: 1 },
        { kind: 'dom', selector: '.pp-p span[style*="text-decoration-line:underline"]', min: 3 },
        { kind: 'dom', selector: '.pp-p span[style*="line-through"]', min: 2 },
        { kind: 'dom', selector: '.pp-p span[style*="vertical-align"]', min: 2 },
        { kind: 'dom', selector: '.pp-p span[style*="background-clip:text"]', min: 1 },
        { kind: 'dom', selector: '.pp-p span[style*="-webkit-text-stroke"]', min: 1 },
        { kind: 'dom', selector: '.pp-p span[style*="background-color"]', min: 1 },
        { kind: 'dom', selector: '.pp-p span[style*="text-shadow"]', min: 1 },
        { kind: 'text', contains: '➢✔▪❖❑' },
        { kind: 'text', contains: 'αβγδπ' },
        { kind: 'text', contains: `Slide number: ${s.index}` },
        { kind: 'model', desc: 'superscript baseline', fn: (sl) => sl.elements.some((e) => e.text && e.text.paragraphs.some((p) => p.runs.some((r) => r.style && r.style.baseline > 0.29))) },
      ],
    });
  }

  // --- Alignment & anchoring ---------------------------------------------------
  {
    const s = deck.slide({ name: 'Alignment' });
    label(s, 'Paragraph alignment and vertical anchoring, insets, autofit, wrap');
    const al = ['l', 'ctr', 'r', 'just', 'dist'];
    al.forEach((a, i) => box(s, { name: 'algn ' + a, x: 40, y: 60 + i * 80, w: 360, h: 70, text: [{ pPr: `algn="${a}"`, runs: [{ t: `Align ${a}: the quick brown fox jumps over the lazy dog`, rPr: 'sz="1400"' }] }], bodyPr: '<a:bodyPr anchor="ctr"/>' }));
    ['t', 'ctr', 'b'].forEach((a, i) => box(s, { name: 'anchor ' + a, x: 440 + i * 140, y: 60, w: 120, h: 240, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'anchor ' + a, rPr: 'sz="1400"' }] }], bodyPr: `<a:bodyPr anchor="${a}"/>` }));
    box(s, { name: 'insets', x: 440, y: 320, w: 400, h: 100, text: [{ runs: [{ t: 'Large insets (36pt left, 18pt top)', rPr: 'sz="1400"' }] }], bodyPr: '<a:bodyPr lIns="457200" tIns="228600" rIns="0" bIns="0" anchor="t"/>' });
    box(s, { name: 'nowrap', x: 440, y: 440, w: 160, h: 50, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'wrap none: this line keeps going past the box', rPr: 'sz="1400"' }] }], bodyPr: '<a:bodyPr wrap="none" anchor="ctr"/>' });
    box(s, { name: 'shrink', x: 440, y: 520, w: 300, h: 80, text: [{ runs: [{ t: 'Shrink text on overflow: fontScale 62.5% with line spacing reduction applied to this long paragraph of text', rPr: 'sz="2400"' }] }], bodyPr: '<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>' });
    box(s, { name: 'anchorCtr', x: 900, y: 60, w: 340, h: 120, text: [{ pPr: 'algn="l"', runs: [{ t: 'anchorCtr: block centred', rPr: 'sz="1400"' }] }, { pPr: 'algn="l"', runs: [{ t: 'second', rPr: 'sz="1400"' }] }], bodyPr: '<a:bodyPr anchor="ctr" anchorCtr="1"/>' });
    box(s, { name: 'columns', x: 900, y: 200, w: 340, h: 200, text: [{ runs: [{ t: 'Two columns of text. '.repeat(10), rPr: 'sz="1200"' }] }], bodyPr: '<a:bodyPr numCol="2" spcCol="182880"/>' });
    box(s, { name: 'rtl', x: 900, y: 420, w: 340, h: 60, text: [{ pPr: 'algn="r" rtl="1"', runs: [{ t: 'مرحبا بالعالم', rPr: 'sz="2000"', lang: 'ar-SA' }] }], bodyPr: '<a:bodyPr anchor="ctr"/>' });
    out.push({
      slide: s.index, feature: 'text/alignment', compareLO: true,
      checks: [
        { kind: 'dom', selector: '.pp-p[style*="text-align:center"]', min: 2 },
        { kind: 'dom', selector: '.pp-p[style*="text-align:right"]', min: 2 },
        { kind: 'dom', selector: '.pp-p[style*="text-align:justify"]', min: 2 },
        { kind: 'dom', selector: '.pp-tx[style*="justify-content:flex-end"]', min: 1 },
        { kind: 'dom', selector: '.pp-cols', min: 1 },
        { kind: 'dom', selector: '.pp-nowrap', min: 1 },
        { kind: 'model', desc: 'autofit fontScale applied', fn: (sl) => sl.elements.some((e) => e.name === 'shrink' && Math.abs(e.text.paragraphs[0].runs[0].style.size - 15) < 0.01) },
      ],
    });
  }

  // --- Bullets and numbering -----------------------------------------------------
  {
    const s = deck.slide({ name: 'Bullets' });
    label(s, 'Bullets, picture-less symbols, auto numbering schemes, levels');
    const bul = (char, font, extra = '') => `marL="342900" indent="-342900"`;
    void bul;
    const paras = [
      { pPr: 'marL="342900" indent="-342900"', pInner: '<a:buFont typeface="Arial"/><a:buChar char="•"/>', runs: [{ t: 'Arial bullet', rPr: 'sz="1800"' }] },
      { pPr: 'marL="342900" indent="-342900"', pInner: '<a:buFont typeface="Wingdings" pitchFamily="2" charset="2"/><a:buChar char="§"/>', runs: [{ t: 'Wingdings square', rPr: 'sz="1800"' }] },
      { pPr: 'marL="342900" indent="-342900"', pInner: '<a:buFont typeface="Wingdings" pitchFamily="2" charset="2"/><a:buChar char="Ø"/>', runs: [{ t: 'Wingdings arrow', rPr: 'sz="1800"' }] },
      { pPr: 'marL="342900" indent="-342900"', pInner: '<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPct val="150000"/><a:buFont typeface="Arial"/><a:buChar char="■"/>', runs: [{ t: 'Coloured, 150%', rPr: 'sz="1800"' }] },
      { pPr: 'marL="800100" lvl="1" indent="-342900"', pInner: '<a:buFont typeface="Courier New"/><a:buChar char="o"/>', runs: [{ t: 'Level 2', rPr: 'sz="1600"' }] },
      { pPr: 'marL="1257300" lvl="2" indent="-342900"', pInner: '<a:buFont typeface="Arial"/><a:buChar char="–"/>', runs: [{ t: 'Level 3', rPr: 'sz="1400"' }] },
      { pPr: 'marL="0" indent="0"', pInner: '<a:buNone/>', runs: [{ t: 'No bullet', rPr: 'sz="1800"' }] },
    ];
    box(s, { name: 'bullets', x: 40, y: 60, w: 560, h: 330, text: paras, bodyPr: '<a:bodyPr anchor="t"/>' });
    const schemes = ['arabicPeriod', 'arabicParenR', 'arabicParenBoth', 'romanUcPeriod', 'romanLcParenR', 'alphaUcPeriod', 'alphaLcParenBoth', 'circleNumDbPlain', 'arabicPlain', 'ea1ChsPeriod'];
    schemes.forEach((sc, k) => {
      const col = k % 2, row = Math.floor(k / 2);
      box(s, {
        name: 'num ' + sc, x: 640 + col * 300, y: 60 + row * 130, w: 280, h: 120, bodyPr: '<a:bodyPr anchor="t"/>',
        text: [0, 1, 2].map((n) => ({ pPr: 'marL="457200" indent="-457200"', pInner: `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${sc}"${n === 0 && sc === 'arabicPeriod' ? '' : ''}/>`, runs: [{ t: n === 0 ? sc : 'item', rPr: 'sz="1400"' }] })),
      });
    });
    box(s, {
      name: 'startAt', x: 40, y: 420, w: 560, h: 120, bodyPr: '<a:bodyPr anchor="t"/>',
      text: [0, 1, 2].map(() => ({ pPr: 'marL="457200" indent="-457200"', pInner: '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod" startAt="5"/>', runs: [{ t: 'starts at 5', rPr: 'sz="1600"' }] })),
    });
    out.push({
      slide: s.index, feature: 'text/bullets', compareLO: true,
      checks: [
        { kind: 'text', contains: '•Arial bullet' },
        { kind: 'text', contains: '▪Wingdings square' },
        { kind: 'text', contains: 'iii)' },
        { kind: 'text', contains: '(c)' },
        { kind: 'text', contains: '③' },
        { kind: 'text', contains: '7.starts at 5' },
        { kind: 'dom', selector: '.pp-bu', min: 30 },
      ],
    });
  }

  // --- Spacing ------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Spacing' });
    label(s, 'Line spacing (percent / points), space before / after, tabs, line breaks');
    const lorem = 'Line spacing sample text that wraps onto multiple lines to show the distance between lines.';
    [['single', 'spcPct val="100000"'], ['1.5 lines', 'spcPct val="150000"'], ['double', 'spcPct val="200000"'], ['exact 30pt', 'spcPts val="3000"']].forEach(([t, v], i) => {
      box(s, { name: 'lnSpc ' + t, x: 40 + i * 305, y: 60, w: 290, h: 220, text: [{ pInner: `<a:lnSpc><a:${v}/></a:lnSpc>`, runs: [{ t: t + ': ' + lorem, rPr: 'sz="1400"' }] }], bodyPr: '<a:bodyPr anchor="t"/>' });
    });
    box(s, {
      name: 'before/after', x: 40, y: 300, w: 600, h: 260, bodyPr: '<a:bodyPr anchor="t"/>',
      text: [
        { runs: [{ t: 'Paragraph one', rPr: 'sz="1600"' }], pInner: '<a:spcAft><a:spcPts val="2400"/></a:spcAft>' },
        { runs: [{ t: 'Paragraph two (24pt after previous)', rPr: 'sz="1600"' }], pInner: '<a:spcBef><a:spcPts val="0"/></a:spcBef>' },
        { runs: [{ t: 'Paragraph three (spcBef 200%)', rPr: 'sz="1600"' }], pInner: '<a:spcBef><a:spcPct val="200000"/></a:spcBef>' },
        { runs: [{ t: 'Soft', rPr: 'sz="1600"' }, { br: true }, { t: 'line break', rPr: 'sz="1600"' }] },
        { runs: [] },
        { runs: [{ t: 'After an empty paragraph', rPr: 'sz="1600"' }] },
      ],
    });
    box(s, { name: 'tabs', x: 680, y: 300, w: 560, h: 120, bodyPr: '<a:bodyPr anchor="t"/>', text: [{ pPr: 'defTabSz="914400"', runs: [{ t: 'A\tB\tC\tD', rPr: 'sz="1600"' }] }] });
    // Custom stops: left at 1.5in, centre at 3in, right at 5in.
    box(s, {
      name: 'tab stops', x: 680, y: 440, w: 560, h: 120, bodyPr: '<a:bodyPr anchor="t"/>',
      text: [{ pPr: '', pInner: '<a:tabLst><a:tab pos="1371600" algn="l"/><a:tab pos="2743200" algn="ctr"/><a:tab pos="4572000" algn="r"/></a:tabLst>', runs: [{ t: 'Left\tStop\tCentre\tRight', rPr: 'sz="1600"' }] }],
    });
    // Auto-fit: more text than fits at 24pt; the runtime shrinks it into the box.
    box(s, { name: 'autofit', x: 1000, y: 580, w: 240, h: 100, bodyPr: '<a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr>', text: [{ runs: [{ t: 'Shrink this text so that it fits inside its box', rPr: 'sz="2400"' }] }] });
    const fitCheck = (sec) => {
      const tx = sec.querySelector('.pp-tx[data-fit="norm"]');
      if (!tx) return 'no auto-fit body';
      const k = parseFloat(tx.style.getPropertyValue('--pp-fs'));
      const p = tx.querySelector('.pp-p');
      const cs = getComputedStyle(tx);
      const avail = tx.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
      return (k > 0.4 && k < 1 && p.offsetHeight <= avail + 1) || `scale ${k}, text ${p.offsetHeight} / ${avail}`;
    };
    const pitchCheck = (sec) => {
      const pitch = (prefix) => {
        const span = [...sec.querySelectorAll('.pp-p > span')].find((e) => e.textContent.startsWith(prefix));
        if (!span) return NaN;
        const r = document.createRange();
        r.selectNodeContents(span);
        const tops = [...new Set([...r.getClientRects()].map((q) => Math.round(q.bottom * 10) / 10))];
        return (tops[tops.length - 1] - tops[0]) / (tops.length - 1);
      };
      const got = [pitch('exact 30pt'), pitch('double')];
      return (Math.abs(got[0] - 40) <= 1 && Math.abs(got[1] - 44.8) <= 1) || 'pitch ' + got.map((v) => v.toFixed(1)).join(', ');
    };
    const tabCheck = (sec) => {
      const p = sec.querySelector('.pp-p[data-tabs^="144:l"]');
      if (!p) return 'paragraph not found';
      const o = p.getBoundingClientRect().left;
      const t = [...p.querySelectorAll('.pp-tab')].map((e) => e.getBoundingClientRect());
      const end = p.querySelector('.pp-tab-end').getBoundingClientRect().left;
      const got = [t[0].right - o, (t[1].right + t[2].left) / 2 - o, end - o];
      const exp = [144, 288, 480];
      return got.every((v, i) => Math.abs(v - exp[i]) <= 2) || 'stops at ' + got.map((v) => v.toFixed(1)).join(', ');
    };
    out.push({ slide: s.index, feature: 'text/spacing', compareLO: true, checks: [{ kind: 'dom', selector: '.pp-p[style*="margin-bottom:32px"]', min: 1 }, { kind: 'dom', selector: '.pp-p br', min: 1 }, { kind: 'eval', desc: 'line pitch: exact 30pt = 40px, double 14pt = 44.8px', fn: pitchCheck }, { kind: 'eval', desc: 'left / centre / right tab stops', fn: tabCheck }, { kind: 'eval', desc: 'over-full auto-fit text shrinks into its box', fn: fitCheck }] });
  }

  // --- Vertical text, rotation ------------------------------------------------------
  {
    const s = deck.slide({ name: 'Vertical text' });
    label(s, 'Vertical text (vert, vert270, eaVert, wordArtVert), rotated text, flipped shapes');
    const v = ['vert', 'vert270', 'eaVert', 'wordArtVert', 'mongolianVert'];
    v.forEach((m, i) => box(s, { name: m, x: 40 + i * 150, y: 60, w: 120, h: 300, text: [{ runs: [{ t: m + ' 直書き', rPr: 'sz="1800"' }] }], bodyPr: `<a:bodyPr vert="${m}" anchor="ctr"/>` }));
    s.add(shape({ id: s.id(), name: 'rot45', x: 820, y: 120, w: 300, h: 100, rot: 45, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Rotated 45°', rPr: 'sz="2000"' }] }] }));
    s.add(shape({ id: s.id(), name: 'flipH', x: 820, y: 400, w: 160, h: 100, prst: 'rightArrow', flipH: true, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'flipH', rPr: 'sz="1600"' }] }] }));
    s.add(shape({ id: s.id(), name: 'flipV', x: 1020, y: 400, w: 160, h: 100, prst: 'rightArrow', flipV: true, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'flipV', rPr: 'sz="1600"' }] }] }));
    s.add(shape({ id: s.id(), name: 'bodyRot', x: 820, y: 560, w: 300, h: 100, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'bodyPr rot=90', rPr: 'sz="1600"' }] }], bodyPr: '<a:bodyPr rot="5400000" anchor="ctr"/>' }));
    out.push({ slide: s.index, feature: 'text/vertical', compareLO: true, checks: [{ kind: 'dom', selector: '.pp-tx[style*="writing-mode:vertical-rl"]', min: 3 }, { kind: 'dom', selector: '.pp-tx[style*="scaleX(-1)"]', min: 2 }] });
  }

  // --- Hyperlinks ---------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Links' });
    label(s, 'Hyperlinks: URL, slide jump, next/previous/first/last, shape click');
    const url = s.hyperlink('https://github.com/g21589/PPTX2HTML');
    const sl = s.slideLink(1);
    box(s, {
      name: 'links', x: 40, y: 80, w: 700, h: 300, bodyPr: '<a:bodyPr anchor="t"/>',
      text: [
        { runs: [{ t: 'External link', rPr: 'sz="2400"', inner: `<a:hlinkClick r:id="${url}" tooltip="Project page"/>` }] },
        { runs: [{ t: 'Jump to slide 1', rPr: 'sz="2400"', inner: `<a:hlinkClick r:id="${sl}" action="ppaction://hlinksldjump"/>` }] },
        { runs: [{ t: 'Next slide', rPr: 'sz="2400"', inner: '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/>' }] },
        { runs: [{ t: 'Last slide', rPr: 'sz="2400"', inner: '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=lastslide"/>' }] },
      ],
    });
    s.add(shape({ id: s.id(), name: 'button', x: 800, y: 120, w: 300, h: 100, prst: 'actionButtonForwardNext', hlink: '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide" highlightClick="1"/>' }));
    out.push({
      slide: s.index, feature: 'text/hyperlinks', compareLO: false,
      checks: [
        { kind: 'dom', selector: 'a.pp-a[href="https://github.com/g21589/PPTX2HTML"]', min: 1 },
        { kind: 'dom', selector: 'a.pp-a[data-slide="0"]', min: 1 },
        { kind: 'dom', selector: '[data-jump="nextslide"]', min: 2 },
        { kind: 'dom', selector: '[data-jump="lastslide"]', min: 1 },
      ],
    });
  }

  // --- Text effects / WordArt --------------------------------------------------------
  {
    const s = deck.slide({ name: 'WordArt' });
    label(s, 'WordArt styles: outline, gradient, glow, reflection, pattern text');
    s.add(textBox({ id: s.id(), name: 'glow', x: 40, y: 80, w: 1200, h: 100, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Glowing text', rPr: 'sz="6000" b="1"', inner: `${solid('FFFFFF')}<a:effectLst><a:glow rad="139700"><a:srgbClr val="4472C4"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst>` }] }] }));
    s.add(textBox({ id: s.id(), name: 'reflect', x: 40, y: 220, w: 1200, h: 100, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Reflected text', rPr: 'sz="5400" b="1"', inner: `${solid('ED7D31')}<a:effectLst><a:reflection blurRad="6350" stA="55000" endA="300" endPos="45500" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>` }] }] }));
    s.add(textBox({ id: s.id(), name: 'pattern', x: 40, y: 400, w: 1200, h: 100, text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Pattern filled', rPr: 'sz="6600" b="1"', inner: '<a:pattFill prst="wdDnDiag"><a:fgClr><a:srgbClr val="7030A0"/></a:fgClr><a:bgClr><a:srgbClr val="FFC000"/></a:bgClr></a:pattFill>' }] }] }));
    s.add(textBox({ id: s.id(), name: 'warp', x: 40, y: 540, w: 1200, h: 140, bodyPr: '<a:bodyPr wrap="none"><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp><a:spAutoFit/></a:bodyPr>', text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Arch up WordArt', rPr: 'sz="4800" b="1"', inner: solid('00B050') }] }] }));
    out.push({ slide: s.index, feature: 'text/wordart', compareLO: true, checks: [{ kind: 'text', contains: 'Glowing text' }, { kind: 'dom', selector: '.pp-p span[style*="text-shadow"]', min: 1 }] });
  }

  // --- WordArt transforms (prstTxWarp) --------------------------------------------------
  {
    const s = deck.slide({ name: 'WordArt warps' });
    label(s, 'WordArt transforms: arch, circle, button, wave, inflate, deflate, slant, triangle, chevron, can, fade, cascade, ring, plain');
    const warps = ['textArchUp', 'textArchDown', 'textCircle', 'textButton', 'textWave1', 'textInflate', 'textDeflate', 'textSlantUp',
      'textTriangle', 'textChevron', 'textCanUp', 'textFadeRight', 'textCascadeUp', 'textRingInside', 'textDeflateInflate', 'textPlain'];
    warps.forEach((w, i) => {
      const x = 30 + (i % 4) * 310, y = 70 + Math.floor(i / 4) * 160;
      const lines = w === 'textButton' || w === 'textDeflateInflate' ? ['Word', 'Art', 'Text'].slice(0, w === 'textButton' ? 3 : 2) : ['WordArt'];
      s.add(shape({
        id: s.id(), name: w, x, y, w: 280, h: 130, fill: '<a:noFill/>', line: '<a:ln><a:noFill/></a:ln>', style: '',
        bodyPr: `<a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:prstTxWarp prst="${w}"><a:avLst/></a:prstTxWarp><a:noAutofit/></a:bodyPr>`,
        text: lines.map((t) => ({ pPr: 'algn="ctr"', runs: [{ t, rPr: 'sz="3200" b="1"', inner: solid(['C00000', '0070C0', '00B050', '7030A0'][i % 4]) }] })),
      }));
    });
    // The glyphs of an arch are rotated along the path; plain text is stretched over the box.
    const warpCheck = (sec) => {
      const tx = [...sec.querySelectorAll('.pp-tx[data-warp]')];
      if (tx.length !== 16) return tx.length + ' warped text bodies';
      const arch = tx[0].querySelectorAll('.pp-wg');
      if (arch.length !== 7) return arch.length + ' glyphs in the arch';
      const rot = (g) => { const m = new DOMMatrix(getComputedStyle(g).transform); return Math.atan2(m.b, m.a) * 180 / Math.PI; };
      const r0 = rot(arch[0]), r6 = rot(arch[6]);
      if (!(r0 < -8 && r6 > 8)) return `arch rotations ${r0.toFixed(1)} / ${r6.toFixed(1)}`;
      const plain = tx[15].getBoundingClientRect(), glyphs = [...tx[15].querySelectorAll('.pp-wg')].map((g) => g.getBoundingClientRect());
      const left = Math.min(...glyphs.map((r) => r.left)), right = Math.max(...glyphs.map((r) => r.right));
      const bottom = Math.max(...glyphs.map((r) => r.bottom)), top = Math.min(...glyphs.map((r) => r.top));
      // Ink fills the box: glyph boxes span its width and (with their leading) cover its height.
      return (Math.abs(left - plain.left) < 3 && Math.abs(right - plain.right) < 3 && top <= plain.top && bottom >= plain.bottom) || `plain text box ${[left, top, right, bottom].map((v) => v.toFixed(0))} vs ${[plain.left, plain.top, plain.right, plain.bottom].map((v) => v.toFixed(0))}`;
    };
    out.push({ slide: s.index, feature: 'text/warp', compareLO: true, checks: [{ kind: 'eval', desc: 'glyphs follow arch and fill plain warp', fn: warpCheck }] });
  }
  return out;
}

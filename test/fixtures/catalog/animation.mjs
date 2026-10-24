/**
 * Fixture catalog: animations (p:timing) and slide transitions.
 */
import { shape, label, solid, txBody, wav } from '../lib.mjs';

export const name = 'animation';
export const title = 'Animations and transitions';

let ctn = 100;
const id = () => ctn++;

const set = (spid, vis = 'visible', delay = 0, tgt = null) => `<p:set><p:cBhvr><p:cTn id="${id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn><p:tgtEl>${tgt || `<p:spTgt spid="${spid}"/>`}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${vis}"/></p:to></p:set>`;
const effect = (spid, filter, transition, dur, tgt) => `<p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${id()}" dur="${dur}"/><p:tgtEl>${tgt || `<p:spTgt spid="${spid}"/>`}</p:tgtEl></p:cBhvr></p:animEffect>`;
const anim = (spid, attr, tav, dur, o = {}) => `<p:anim calcmode="${o.calcmode || 'lin'}" valueType="${o.valueType || 'num'}"><p:cBhvr additive="${o.additive || 'base'}"><p:cTn id="${id()}" dur="${dur}" fill="hold"${o.cTn || ''}/><p:tgtEl>${o.tgt || `<p:spTgt spid="${spid}"/>`}</p:tgtEl><p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>${tav.map(([tm, v, fmla]) => `<p:tav tm="${tm}"${fmla ? ` fmla="${fmla}"` : ''}><p:val>${typeof v === 'number' ? `<p:fltVal val="${v}"/>` : `<p:strVal val="${v}"/>`}</p:val></p:tav>`).join('')}</p:tavLst></p:anim>`;

/** Effect container (preset node). */
function eff(o, children) {
  return `<p:par><p:cTn id="${id()}" presetID="${o.preset || 10}" presetClass="${o.cls}" presetSubtype="${o.sub || 0}"${o.repeat ? ` repeatCount="${o.repeat}"` : ''}${o.autoRev ? ' autoRev="1"' : ''} fill="hold"${o.grpId !== undefined ? ` grpId="${o.grpId}"` : ' grpId="0"'} nodeType="${o.node || 'clickEffect'}"><p:stCondLst><p:cond delay="${o.delay || 0}"/></p:stCondLst>${o.iterate || ''}<p:childTnLst>${children}</p:childTnLst></p:cTn></p:par>`;
}

/** Click group: list of sub groups (each sub group starts after `delay`). */
function click(subs, o = {}) {
  const cond = o.auto ? '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : '<p:cond delay="indefinite"/>';
  return `<p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${subs.map(([delay, effects]) => `<p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${effects.join('')}</p:childTnLst></p:cTn></p:par>`).join('')}</p:childTnLst></p:cTn></p:par>`;
}

function timing(clicks, o = {}) {
  ctn = 100;
  const main = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${clicks.join('')}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${clicks.length ? main : ''}${o.extraSeq || ''}</p:childTnLst></p:cTn></p:par></p:tnLst>${o.bld ? `<p:bldLst>${o.bld}</p:bldLst>` : ''}</p:timing>`;
}

// Preset behaviour bundles -------------------------------------------------------
const P = {
  appear: (sp, o = {}) => eff({ preset: 1, cls: 'entr', ...o }, set(sp)),
  fade: (sp, o = {}) => eff({ preset: 10, cls: 'entr', ...o }, set(sp) + effect(sp, 'fade', 'in', o.dur || 500)),
  filterIn: (sp, filter, o = {}) => eff({ preset: o.preset || 22, cls: 'entr', ...o }, set(sp) + effect(sp, filter, 'in', o.dur || 500)),
  flyInBottom: (sp, o = {}) => eff({ preset: 2, sub: 4, cls: 'entr', ...o }, set(sp) + anim(sp, 'ppt_x', [[0, '#ppt_x'], [100000, '#ppt_x']], o.dur || 500) + anim(sp, 'ppt_y', [[0, '1+#ppt_h/2'], [100000, '#ppt_y']], o.dur || 500)),
  flyInLeft: (sp, o = {}) => eff({ preset: 2, sub: 8, cls: 'entr', ...o }, set(sp) + anim(sp, 'ppt_x', [[0, '0-#ppt_w/2'], [100000, '#ppt_x']], o.dur || 500) + anim(sp, 'ppt_y', [[0, '#ppt_y'], [100000, '#ppt_y']], o.dur || 500)),
  zoomIn: (sp, o = {}) => eff({ preset: 53, sub: 16, cls: 'entr', ...o }, set(sp) + anim(sp, 'ppt_w', [[0, 0], [100000, '#ppt_w']], o.dur || 500) + anim(sp, 'ppt_h', [[0, 0], [100000, '#ppt_h']], o.dur || 500) + effect(sp, 'fade', 'in', o.dur || 500)),
  spin: (sp, o = {}) => eff({ preset: 8, cls: 'emph', ...o }, `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${id()}" dur="${o.dur || 2000}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`),
  grow: (sp, o = {}) => eff({ preset: 6, cls: 'emph', ...o }, `<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="${o.dur || 1000}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl></p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`),
  pulse: (sp, o = {}) => eff({ preset: 26, cls: 'emph', ...o }, `<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="250" autoRev="1" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl></p:cBhvr><p:by x="110000" y="110000"/></p:animScale>`),
  transparency: (sp, o = {}) => eff({ preset: 9, cls: 'emph', ...o }, `<p:set><p:cBhvr><p:cTn id="${id()}" dur="${o.dur || 500}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl><p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="0.5"/></p:to></p:set>`),
  fillColor: (sp, hex, o = {}) => eff({ preset: 1, cls: 'emph', ...o }, `<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr><p:cTn id="${id()}" dur="${o.dur || 500}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="${hex}"/></p:to></p:animClr>`),
  fontColor: (sp, hex, o = {}) => eff({ preset: 3, cls: 'emph', ...o }, `<p:animClr clrSpc="hsl" dir="cw"><p:cBhvr override="childStyle"><p:cTn id="${id()}" dur="${o.dur || 500}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl><p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="${hex}"/></p:to></p:animClr>`),
  teeter: (sp, o = {}) => eff({ preset: 32, cls: 'emph', ...o }, anim(sp, 'r', [[0, 0], [25000, 5], [50000, -5], [75000, 5], [100000, 0]], o.dur || 1000, { additive: 'sum' })),
  fadeOut: (sp, o = {}) => eff({ preset: 10, cls: 'exit', ...o }, effect(sp, 'fade', 'out', o.dur || 500) + set(sp, 'hidden', (o.dur || 500) - 1)),
  flyOutRight: (sp, o = {}) => eff({ preset: 2, sub: 2, cls: 'exit', ...o }, anim(sp, 'ppt_x', [[0, 'ppt_x'], [100000, '1+ppt_w/2']], o.dur || 500) + anim(sp, 'ppt_y', [[0, 'ppt_y'], [100000, 'ppt_y']], o.dur || 500) + set(sp, 'hidden', (o.dur || 500) - 1)),
  path: (sp, path, o = {}) => eff({ preset: 0, cls: 'path', ...o }, `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${id()}" dur="${o.dur || 2000}" fill="hold"/><p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`),
};

const box = (s, name, x, y, c = '4472C4', text, prst = 'rect') => {
  const i = s.id();
  s.add(shape({ id: i, name, x, y, w: 160, h: 100, prst, fill: solid(c), text: [{ pPr: 'algn="ctr"', runs: [{ t: text || name, rPr: 'sz="1400"' }] }] }));
  return i;
};

const TRANSITIONS = [
  ['fade', '<p:fade/>'], ['fade thruBlk', '<p:fade thruBlk="1"/>'], ['push', '<p:push dir="u"/>'], ['wipe', '<p:wipe dir="r"/>'],
  ['split', '<p:split orient="vert" dir="out"/>'], ['cover', '<p:cover dir="l"/>'], ['pull', '<p:pull dir="r"/>'], ['cut', '<p:cut/>'],
  ['blinds', '<p:blinds dir="vert"/>'], ['checker', '<p:checker/>'], ['circle', '<p:circle/>'], ['comb', '<p:comb/>'], ['diamond', '<p:diamond/>'],
  ['dissolve', '<p:dissolve/>'], ['newsflash', '<p:newsflash/>'], ['plus', '<p:plus/>'], ['random', '<p:random/>'], ['randomBar', '<p:randomBar dir="vert"/>'],
  ['strips', '<p:strips dir="rd"/>'], ['wedge', '<p:wedge/>'], ['wheel', '<p:wheel spokes="4"/>'], ['zoom', '<p:zoom/>'],
  ['vortex', '<p14:vortex dir="r"/>', 'p14'], ['switch', '<p14:switch dir="r"/>', 'p14'], ['flip', '<p14:flip dir="r"/>', 'p14'], ['ripple', '<p14:ripple/>', 'p14'],
  ['honeycomb', '<p14:honeycomb/>', 'p14'], ['prism', '<p14:prism/>', 'p14'], ['doors', '<p14:doors dir="vert"/>', 'p14'], ['window', '<p14:window dir="vert"/>', 'p14'],
  ['ferris', '<p14:ferris dir="l"/>', 'p14'], ['gallery', '<p14:gallery dir="l"/>', 'p14'], ['conveyor', '<p14:conveyor dir="l"/>', 'p14'], ['pan', '<p14:pan dir="d"/>', 'p14'],
  ['glitter', '<p14:glitter dir="l" pattern="hexagon"/>', 'p14'], ['warp', '<p14:warp dir="in"/>', 'p14'], ['flythrough', '<p14:flythrough/>', 'p14'], ['flash', '<p14:flash/>', 'p14'],
  ['shred', '<p14:shred/>', 'p14'], ['reveal', '<p14:reveal/>', 'p14'], ['wheelReverse', '<p14:wheelReverse spokes="2"/>', 'p14'],
  ['curtains', '<p15:prstTrans prst="curtains"/>', 'p15'], ['wind', '<p15:prstTrans prst="wind"/>', 'p15'], ['fracture', '<p15:prstTrans prst="fracture"/>', 'p15'],
  ['pageCurlDouble', '<p15:prstTrans prst="pageCurlDouble"/>', 'p15'], ['airplane', '<p15:prstTrans prst="airplane"/>', 'p15'], ['fallOver', '<p15:prstTrans prst="fallOver"/>', 'p15'],
  ['morph', '<p159:morph option="byObject"/>', 'p159'],
];

export function build(deck) {
  const out = [];
  // --- Entrance effects ------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Entrance effects' });
    label(s, 'Entrance effects (one per click)');
    const filters = [['Fade', null], ['Wipe up', 'wipe(up)'], ['Split', 'barn(outVertical)'], ['Blinds', 'blinds(horizontal)'], ['Box', 'box(out)'], ['Checker', 'checkerboard(across)'],
      ['Circle', 'circle(out)'], ['Diamond', 'diamond(out)'], ['Dissolve', 'dissolve'], ['Plus', 'plus(out)'], ['Random bars', 'randombar(horizontal)'], ['Strips', 'strips(downRight)'],
      ['Wedge', 'wedge'], ['Wheel', 'wheel(2)'], ['Slide', 'slide(fromBottom)']];
    const ids = filters.map(([n], i) => box(s, n, 40 + (i % 6) * 205, 80 + Math.floor(i / 6) * 150, ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'][i % 6]));
    const appear = box(s, 'Appear', 40, 560, '7030A0');
    const fly = box(s, 'Fly in', 245, 560, 'C00000');
    const fly2 = box(s, 'Fly left', 450, 560, '00B050');
    const zoom = box(s, 'Zoom', 655, 560, '002060');
    const clicks = filters.map(([n, f], i) => click([[0, [f ? P.filterIn(ids[i], f) : P.fade(ids[i])]]]));
    clicks.push(click([[0, [P.appear(appear)]]]), click([[0, [P.flyInBottom(fly)]]]), click([[0, [P.flyInLeft(fly2)]]]), click([[0, [P.zoomIn(zoom)]]]));
    s.timing = timing(clicks);
    const checks = [
      { kind: 'anim', step: 0, spid: ids[0], expect: { visible: false } },
      { kind: 'anim', step: 0, t: 250, spid: ids[0], expect: { visible: true, opacity: [0.35, 0.65] } },
      { kind: 'anim', step: 1, spid: ids[0], expect: { visible: true, opacity: [0.99, 1] } },
      { kind: 'anim', step: 1, t: 250, spid: ids[1], expect: { visible: true, clip: true } },
      { kind: 'anim', step: 2, spid: ids[1], expect: { visible: true, clip: false } },
      { kind: 'anim', step: 6, t: 250, spid: ids[6], expect: { clip: true } },
      { kind: 'anim', step: 8, t: 250, spid: ids[8], expect: { mask: true } },
      { kind: 'anim', step: 15, t: 1, spid: appear, expect: { visible: true } },
      { kind: 'anim', step: 16, t: 1, spid: fly, expect: { visible: true, ty: [150, 170] } },
      { kind: 'anim', step: 16, t: 250, spid: fly, expect: { ty: [70, 90] } },
      { kind: 'anim', step: 17, spid: fly, expect: { ty: [-1, 1] } },
      { kind: 'anim', step: 17, t: 1, spid: fly2, expect: { tx: [-800, -300] } },
      { kind: 'anim', step: 18, t: 250, spid: zoom, expect: { scale: [0.3, 0.7] } },
      { kind: 'model', desc: '19 click steps, all entrance targets hidden', fn: (sl) => sl.timing.main.length === 19 && sl.timing.hidden.length === 19 },
    ];
    out.push({ slide: s.index, feature: 'animation/entrance', checks });
  }

  // --- Emphasis, exit, motion path ----------------------------------------------------
  {
    const s = deck.slide({ name: 'Emphasis exit path' });
    label(s, 'Emphasis, exit and motion path effects');
    const spin = box(s, 'Spin', 60, 100, 'ED7D31');
    const grow = box(s, 'Grow', 300, 100, '70AD47');
    const pulse = box(s, 'Pulse', 540, 100, '5B9BD5');
    const transp = box(s, 'Transparency', 780, 100, 'FFC000');
    const fillc = box(s, 'Fill colour', 1020, 100, '4472C4');
    const fontc = box(s, 'Font colour', 60, 300, 'FFFFFF');
    const teeter = box(s, 'Teeter', 300, 300, 'A5A5A5');
    const fadeOut = box(s, 'Fade out', 540, 300, '7030A0');
    const flyOut = box(s, 'Fly out', 780, 300, 'C00000');
    const mover = box(s, 'Motion path', 60, 520, '00B050', 'Path', 'ellipse');
    s.timing = timing([
      click([[0, [P.spin(spin)]]]),
      click([[0, [P.grow(grow)]]]),
      click([[0, [P.pulse(pulse)]]]),
      click([[0, [P.transparency(transp)]]]),
      click([[0, [P.fillColor(fillc, 'FF0000')]]]),
      click([[0, [P.fontColor(fontc, 'FF0000')]]]),
      click([[0, [P.teeter(teeter)]]]),
      click([[0, [P.fadeOut(fadeOut)]]]),
      click([[0, [P.flyOutRight(flyOut)]]]),
      click([[0, [P.path(mover, 'M 0 0 L 0.5 0 L 0.5 -0.25 E')]]]),
    ]);
    out.push({
      slide: s.index, feature: 'animation/emphasis-exit-path',
      checks: [
        { kind: 'anim', step: 0, t: 500, spid: spin, expect: { rot: [80, 100] } },
        { kind: 'anim', step: 1, spid: spin, expect: { rot: [-1, 1] } },
        { kind: 'anim', step: 1, t: 1000, spid: grow, expect: { scale: [1.45, 1.55] } },
        { kind: 'anim', step: 2, spid: grow, expect: { scale: [1.45, 1.55] } },
        { kind: 'anim', step: 2, t: 250, spid: pulse, expect: { scale: [1.08, 1.12] } },
        { kind: 'anim', step: 3, spid: pulse, expect: { scale: [0.99, 1.01] } },
        { kind: 'anim', step: 4, spid: transp, expect: { opacity: [0.45, 0.55] } },
        { kind: 'anim', step: 5, spid: fillc, expect: { fill: '#FF0000' } },
        { kind: 'anim', step: 6, spid: fontc, expect: { textColor: '#FF0000' } },
        { kind: 'anim', step: 6, t: 250, spid: teeter, expect: { rot: [4, 6] } },
        { kind: 'anim', step: 7, t: 250, spid: fadeOut, expect: { visible: true, opacity: [0.35, 0.65] } },
        { kind: 'anim', step: 8, spid: fadeOut, expect: { visible: false } },
        { kind: 'anim', step: 8, t: 250, spid: flyOut, expect: { tx: [100, 600] } },
        { kind: 'anim', step: 9, spid: flyOut, expect: { visible: false } },
        { kind: 'anim', step: 9, t: 1000, spid: mover, expect: { tx: [400, 420], ty: [-1, 1] } },
        { kind: 'anim', step: 10, spid: mover, expect: { tx: [630, 650], ty: [-190, -170] } },
      ],
    });
  }

  // --- Sequencing ------------------------------------------------------------------
  {
    const s = deck.slide({ name: 'Sequencing' });
    label(s, 'Timing: auto start, with previous, after previous, delays, repeat, auto-reverse');
    const a = box(s, 'A auto', 60, 100);
    const b = box(s, 'B with A', 300, 100);
    const c = box(s, 'C after (delay 500)', 540, 100);
    const d = box(s, 'D click', 780, 100);
    const e = box(s, 'E repeat x3', 60, 320, '70AD47');
    const f = box(s, 'F autoRev', 300, 320, 'ED7D31');
    s.timing = timing([
      click([[0, [P.fade(a, { node: 'afterEffect' }), P.fade(b, { node: 'withEffect' })]], [1000, [P.fade(c, { node: 'afterEffect', delay: 500 })]]], { auto: true }),
      click([[0, [P.flyInLeft(d)]]]),
      click([[0, [P.spin(e, { repeat: 3000, dur: 400 })]]]),
      click([[0, [P.grow(f, { autoRev: true, dur: 500 })]]]),
    ]);
    out.push({
      slide: s.index, feature: 'animation/sequencing',
      checks: [
        { kind: 'model', desc: 'first step auto', fn: (sl) => sl.timing.main[0].auto === true && sl.timing.main.length === 4 },
        { kind: 'model', desc: 'after-previous begins at 1500', fn: (sl) => sl.timing.main[0].behaviours.some((x) => x.keys[0] === String(c) && x.begin === 1500) },
        { kind: 'anim', step: 0, t: 250, spid: b, expect: { opacity: [0.35, 0.65] } },
        { kind: 'anim', step: 0, t: 1200, spid: c, expect: { visible: false } },
        { kind: 'anim', step: 0, t: 1750, spid: c, expect: { visible: true, opacity: [0.35, 0.65] } },
        { kind: 'anim', step: 2, t: 600, spid: e, expect: { rot: [170, 190] } },
        { kind: 'anim', step: 3, spid: e, expect: { rot: [-1, 1] } },
        { kind: 'anim', step: 3, t: 500, spid: f, expect: { scale: [1.45, 1.55] } },
        { kind: 'anim', step: 3, t: 900, spid: f, expect: { scale: [1, 1.15] } },
        { kind: 'anim', step: 4, spid: f, expect: { scale: [0.99, 1.01] } },
      ],
    });
  }

  // --- Build by paragraph / by letter / triggers / media ----------------------------------
  {
    const s = deck.slide({ name: 'Text builds and triggers' });
    label(s, 'Build by paragraph, by letter, triggers, media command');
    const tid = s.id();
    s.add(shape({
      id: tid, name: 'Bulleted', x: 60, y: 80, w: 500, h: 260, fill: solid('DEEBF7'), style: '',
      text: txBody(['First point', 'Second point', 'Third point'].map((t) => ({ pPr: 'marL="285750" indent="-285750"', pInner: '<a:buFont typeface="Arial"/><a:buChar char="•"/>', runs: [{ t, rPr: 'sz="2400"' }] })), { bodyPr: '<a:bodyPr anchor="t"/>' }),
    }));
    const ptgt = (i) => `<p:spTgt spid="${tid}"><p:txEl><p:pRg st="${i}" end="${i}"/></p:txEl></p:spTgt>`;
    const paraFade = (i) => eff({ preset: 10, cls: 'entr', grpId: 0 }, set(tid, 'visible', 0, ptgt(i)) + effect(tid, 'fade', 'in', 500, ptgt(i)));
    const lid = s.id();
    s.add(shape({ id: lid, name: 'Letters', x: 640, y: 80, w: 560, h: 100, style: '', fill: '<a:noFill/>', text: [{ runs: [{ t: 'By letter', rPr: 'sz="4000" b="1"', inner: solid('C00000') }] }] }));
    const letters = eff({ preset: 10, cls: 'entr', iterate: '<p:iterate type="lt"><p:tmPct val="10000"/></p:iterate>' }, set(lid) + effect(lid, 'fade', 'in', 500));
    const btn = s.id();
    s.add(shape({ id: btn, name: 'Trigger button', x: 640, y: 300, w: 200, h: 80, prst: 'roundRect', fill: solid('ED7D31'), text: [{ pPr: 'algn="ctr"', runs: [{ t: 'Click me', rPr: 'sz="1800"' }] }] }));
    const target = box(s, 'Triggered', 900, 300, '7030A0');
    const trig = `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${id()}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${btn}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${P.fade(target)}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${btn}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
    const media = s.mediaFile(wav(300, 660), 'wav', 'audio');
    const icon = s.image(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><circle cx="5" cy="5" r="5" fill="#4472C4"/></svg>'), 'svg');
    const aud = s.id();
    s.add(`<p:pic><p:nvPicPr><p:cNvPr id="${aud}" name="sound.wav"/><p:cNvPicPr/><p:nvPr><a:audioFile r:link="${media.link}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE67E2D6D2C}"><p14:media r:embed="${media.embed}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${icon}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${60 * 9525}" y="${560 * 9525}"/><a:ext cx="${80 * 9525}" cy="${80 * 9525}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
    const play = eff({ preset: 1, cls: 'mediacall' }, `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${id()}" dur="300" fill="hold"/><p:tgtEl><p:spTgt spid="${aud}"/></p:tgtEl></p:cBhvr></p:cmd>`);
    s.timing = timing([click([[0, [paraFade(0)]]]), click([[0, [paraFade(1)]]]), click([[0, [paraFade(2)]]]), click([[0, [letters]]]), click([[0, [play]]])], { extraSeq: trig, bld: `<p:bldP spid="${tid}" grpId="0" build="p"/><p:bldP spid="${lid}" grpId="0"/>` });
    out.push({
      slide: s.index, feature: 'animation/builds-triggers',
      checks: [
        { kind: 'anim', step: 0, spid: tid, para: 0, expect: { visible: false } },
        { kind: 'anim', step: 0, spid: tid, expect: { visible: true } },
        { kind: 'anim', step: 1, spid: tid, para: 0, expect: { visible: true } },
        { kind: 'anim', step: 1, spid: tid, para: 1, expect: { visible: false } },
        { kind: 'anim', step: 3, spid: tid, para: 2, expect: { visible: true } },
        { kind: 'anim', step: 3, t: 300, spid: lid, letter: 0, expect: { visible: true } },
        { kind: 'anim', step: 3, t: 300, spid: lid, letter: 7, expect: { visible: false } },
        { kind: 'anim', step: 4, spid: target, expect: { visible: false } },
        { kind: 'trigger', spid: btn, target, expect: { visible: true } },
        { kind: 'model', desc: 'trigger sequence parsed', fn: (sl) => sl.timing.triggers.length === 1 && sl.timing.triggers[0].spid === String(btn) },
        { kind: 'model', desc: 'media command', fn: (sl) => sl.timing.main[4].behaviours.some((x) => x.type === 'cmd' && x.cmd === 'playFrom(0.0)') },
      ],
    });
  }

  // --- Transitions ------------------------------------------------------------------
  const colors = ['DEEBF7', 'FBE5D6', 'E2F0D9', 'FFF2CC', 'EDE7F6', 'FCE4EC'];
  TRANSITIONS.forEach(([n, xml, ns], i) => {
    const tr = ns ? `<mc:AlternateContent><mc:Choice Requires="${ns}"><p:transition spd="slow" p14:dur="1200">${xml}</p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>` : `<p:transition spd="slow">${xml}</p:transition>`;
    const s = deck.slide({ name: 'Transition ' + n, transition: tr, bg: `<p:bg><p:bgPr>${solid(colors[i % colors.length])}<a:effectLst/></p:bgPr></p:bg>` });
    label(s, 'Transition: ' + n, { sz: 3200 });
    s.add(shape({ id: s.id(), name: 'Box', x: 200 + (i % 3) * 300, y: 250, w: 300, h: 200, prst: ['ellipse', 'rect', 'triangle'][i % 3], fill: solid(['C00000', '0070C0', '00B050'][i % 3]) }));
    out.push({
      slide: s.index, feature: 'transition/' + n,
      checks: [
        { kind: 'model', desc: 'transition type parsed', fn: (sl) => sl.transition && (sl.transition.type === n || sl.transition.type === 'prst:' + n || (n === 'fade thruBlk' && sl.transition.thruBlk)) },
        { kind: 'transition', mid: !['cut', 'random'].includes(n) },
      ],
    });
  });

  // Auto advance.
  {
    const s = deck.slide({ name: 'Auto advance', transition: '<p:transition spd="fast" advClick="0" advTm="1500"><p:fade/></p:transition>' });
    label(s, 'Auto advance after 1.5 s, no click advance');
    out.push({ slide: s.index, feature: 'transition/auto-advance', checks: [{ kind: 'model', desc: 'advTm 1500, advClick false', fn: (sl) => sl.transition.advTm === 1500 && sl.transition.advClick === false }] });
    const end = deck.slide({ name: 'End' });
    label(end, 'End of show');
  }
  return out;
}

export { TRANSITIONS };

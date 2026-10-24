/**
 * Slide transitions (p:transition, including PowerPoint 2010+ extensions
 * delivered through mc:AlternateContent).
 */

const SPEED = { fast: 500, med: 750, slow: 1000 };

const KNOWN = new Set([
  // ECMA-376
  'blinds', 'checker', 'circle', 'comb', 'cover', 'cut', 'diamond', 'dissolve', 'fade', 'newsflash', 'plus', 'pull',
  'push', 'random', 'randomBar', 'split', 'strips', 'wedge', 'wheel', 'wipe', 'zoom',
  // PowerPoint 2010 (p14)
  'vortex', 'switch', 'flip', 'ripple', 'honeycomb', 'prism', 'doors', 'window', 'ferris', 'gallery', 'conveyor', 'pan',
  'glitter', 'warp', 'flythrough', 'flash', 'shred', 'reveal', 'wheelReverse',
  // PowerPoint 2013 (p15)
  'prstTrans',
  // PowerPoint 2016+
  'morph',
]);

function fromNode(tr, ctx) {
  const out = {
    type: 'none',
    dur: SPEED[tr.attr('spd', 'fast')] || 500,
    advClick: tr.attr('advClick', '1') !== '0' && tr.attr('advClick', 'true') !== 'false',
    advTm: tr.attrs.advTm !== undefined ? Number(tr.attrs.advTm) : null,
    sound: null,
    loopSound: false,
    stopSound: false,
  };
  const p14dur = tr.attrs['p14:dur'];
  if (p14dur !== undefined) out.dur = Number(p14dur);
  for (const c of tr.children) {
    if (c.local === 'sndAc') {
      const st = c.child('p:stSnd');
      const snd = st && st.child('p:snd');
      if (snd) {
        out.sound = ctx.mediaRel(snd.attrs['r:embed'], 'audio');
        out.loopSound = st.bool('loop', false);
      }
      if (c.child('p:endSnd')) out.stopSound = true;
      continue;
    }
    if (c.local === 'extLst') continue;
    const name = c.local;
    if (!KNOWN.has(name)) {
      ctx.note('unsupported', 'transition ' + c.name);
      out.type = 'fade';
      continue;
    }
    out.type = name;
    const a = c.attrs;
    if (a.dir !== undefined) out.dir = a.dir;
    if (a.orient !== undefined) out.orient = a.orient;
    if (a.thruBlk !== undefined) out.thruBlk = a.thruBlk === '1' || a.thruBlk === 'true';
    if (a.spokes !== undefined) out.spokes = Number(a.spokes);
    if (a.pattern !== undefined) out.pattern = a.pattern;
    if (a.isContent !== undefined) out.isContent = a.isContent === '1' || a.isContent === 'true';
    if (a.isInverted !== undefined) out.isInverted = a.isInverted === '1' || a.isInverted === 'true';
    if (a.prst !== undefined) out.prst = a.prst;
    if (a.invX !== undefined) out.invX = a.invX === '1' || a.invX === 'true';
    if (a.invY !== undefined) out.invY = a.invY === '1' || a.invY === 'true';
    if (a.option !== undefined) out.option = a.option;
    if (name === 'prstTrans') out.type = 'prst:' + (a.prst || 'fallOver');
  }
  if (!tr.children.some((c) => c.local !== 'sndAc' && c.local !== 'extLst')) {
    // A transition element without an effect only carries timing/sound.
    out.type = 'none';
  }
  return out;
}

/**
 * @param {import('../util/xml.js').XNode} sld p:sld
 */
export function parseTransition(sld, ctx) {
  let tr = sld.child('p:transition');
  // PowerPoint 2010+ wraps richer transitions in AlternateContent.
  for (const ac of sld.all('mc:AlternateContent')) {
    let chosen = null;
    for (const ch of ac.all('mc:Choice')) {
      const req = (ch.attrs.Requires || '').split(/\s+/).filter(Boolean);
      if (req.every((r) => ['p14', 'p15', 'p159', 'p1510', 'p1710', 'a14'].includes(r))) {
        const t = ch.child('p:transition');
        if (t) { chosen = t; break; }
      }
    }
    if (!chosen) {
      const fb = ac.child('mc:Fallback');
      if (fb) chosen = fb.child('p:transition');
    }
    if (chosen) tr = chosen;
  }
  if (!tr) return null;
  return fromNode(tr, ctx);
}

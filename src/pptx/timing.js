/**
 * Slide timing (animations) parser.
 *
 * The PresentationML time tree (p:par / p:seq / p:excl containers holding
 * behaviours such as p:set, p:anim, p:animEffect ...) is flattened into
 * "steps" (one per click of the main sequence or of a trigger sequence),
 * each holding behaviour instances with begin times relative to the step.
 * Repeat / auto-reverse settings of enclosing containers are preserved as
 * "wrappers" so the runtime can evaluate any instant exactly.
 */
import { resolveColorNode, COLOR_NODES } from './color.js';

const INDEF = 'indefinite';

function parseTime(v) {
  if (v === undefined || v === null || v === '') return undefined;
  if (v === INDEF) return Infinity;
  const n = Number(v);
  return Number.isNaN(n) ? undefined : n;
}

/** Common time node (p:cTn). */
function cTnInfo(cTn) {
  if (!cTn) return { dur: undefined, repeat: 1, delay: 0, conds: [] };
  const rc = cTn.attrs.repeatCount;
  const info = {
    id: cTn.attrs.id,
    dur: parseTime(cTn.attrs.dur),
    repeat: rc === INDEF ? Infinity : rc !== undefined ? Number(rc) / 1000 : 1,
    repeatDur: parseTime(cTn.attrs.repeatDur),
    autoRev: cTn.bool('autoRev', false),
    accel: cTn.num('accel', 0) / 100000,
    decel: cTn.num('decel', 0) / 100000,
    spd: cTn.num('spd', 100000) / 100000,
    fill: cTn.attr('fill', null),
    presetClass: cTn.attr('presetClass', null),
    presetID: cTn.int('presetID', null),
    presetSubtype: cTn.int('presetSubtype', null),
    nodeType: cTn.attr('nodeType', null),
    grpId: cTn.attr('grpId', null),
    restart: cTn.attr('restart', null),
    conds: [],
    endConds: [],
    iterate: null,
  };
  const st = cTn.child('p:stCondLst');
  for (const c of st ? st.all('p:cond') : []) info.conds.push(cond(c));
  const en = cTn.child('p:endCondLst');
  for (const c of en ? en.all('p:cond') : []) info.endConds.push(cond(c));
  const it = cTn.child('p:iterate');
  if (it) {
    const abs = it.child('p:tmAbs');
    const pc = it.child('p:tmPct');
    info.iterate = {
      type: it.attr('type', 'el'),
      backwards: it.bool('backwards', false),
      abs: abs ? parseTime(abs.attrs.val) : undefined,
      pct: pc ? pc.num('val', 0) / 100000 : undefined,
    };
  }
  return info;
}

function cond(c) {
  const out = { delay: parseTime(c.attrs.delay) ?? 0, evt: c.attr('evt', null), tn: null, spid: null, rtn: null };
  const tn = c.child('p:tn');
  if (tn) out.tn = tn.attr('val', null);
  const rtn = c.child('p:rtn');
  if (rtn) out.rtn = rtn.attr('val', null);
  const tgt = c.child('p:tgtEl');
  const sp = tgt && tgt.child('p:spTgt');
  if (sp) out.spid = sp.attr('spid', null);
  if (tgt && tgt.child('p:sldTgt')) out.slide = true;
  return out;
}

/** Target element keys: "spid", "spid:bg", "spid:pN". */
function targetOf(tgtEl, ctx, para) {
  if (!tgtEl) return null;
  const sp = tgtEl.child('p:spTgt');
  if (sp) {
    const spid = sp.attr('spid', '');
    if (sp.child('p:bg')) return { spid, keys: [spid + ':bg'], kind: 'bg' };
    const tx = sp.child('p:txEl');
    if (tx) {
      const pRg = tx.child('p:pRg');
      if (pRg) {
        const st = pRg.int('st', 0), end = pRg.int('end', st);
        const keys = [];
        for (let i = st; i <= end; i++) keys.push(spid + ':p' + i);
        return { spid, keys, kind: 'para', paras: [st, end] };
      }
      const cRg = tx.child('p:charRg');
      if (cRg) return { spid, keys: [spid], kind: 'chars', chars: [cRg.int('st', 0), cRg.int('end', 0)] };
    }
    const sub = sp.child('p:subSp');
    if (sub) return { spid, keys: [spid], kind: 'sub', sub: sub.attr('spid', '') };
    const ole = sp.child('p:oleChartEl');
    if (ole) return { spid, keys: [spid], kind: 'chartEl', chartEl: { type: ole.attr('type', ''), lvl: ole.int('lvl', 0) } };
    const ge = sp.child('p:graphicEl');
    if (ge) return { spid, keys: [spid], kind: 'graphicEl' };
    return { spid, keys: [spid], kind: 'shape' };
  }
  const snd = tgtEl.child('p:sndTgt');
  if (snd) {
    const media = ctx.mediaRel(snd.attrs['r:embed'], 'audio');
    return { kind: 'sound', keys: [], sound: media, name: snd.attr('name', '') };
  }
  if (tgtEl.child('p:sldTgt')) return { kind: 'slide', keys: ['slide'] };
  void para;
  return null;
}

function attrNames(cBhvr) {
  const l = cBhvr && cBhvr.child('p:attrNameLst');
  return l ? l.all('p:attrName').map((a) => a.text.trim()) : [];
}

function tavValue(valNode, ctx) {
  if (!valNode) return null;
  for (const c of valNode.children) {
    switch (c.local) {
      case 'strVal': return c.attr('val', '');
      case 'fltVal': return c.num('val', 0);
      case 'intVal': return c.int('val', 0);
      case 'boolVal': return c.attr('val', 'false') === 'true' || c.attr('val', '0') === '1';
      case 'clrVal': {
        const cn = c.children.find((x) => COLOR_NODES.has(x.name));
        return cn ? colorObj(resolveColorNode(cn, ctx)) : null;
      }
      default: break;
    }
  }
  return null;
}

function colorObj(c) {
  return c ? { r: Math.round(c.r), g: Math.round(c.g), b: Math.round(c.b), a: c.a } : null;
}

function clrSpec(n, ctx) {
  if (!n) return null;
  const cn = n.children.find((x) => COLOR_NODES.has(x.name));
  if (cn) return { color: colorObj(resolveColorNode(cn, ctx)) };
  const rgb = n.child('p:rgb');
  if (rgb) return { rgb: [rgb.num('r', 0) / 1000, rgb.num('g', 0) / 1000, rgb.num('b', 0) / 1000] };
  const hsl = n.child('p:hsl');
  if (hsl) return { hsl: [hsl.num('h', 0) / 60000, hsl.num('s', 0) / 1000, hsl.num('l', 0) / 1000] };
  return null;
}

/** Parse a motion path string ("M 0 0 L 0.25 0.33 C ... E") into commands. */
export function parseMotionPath(s) {
  const toks = String(s || '').replace(/,/g, ' ').match(/[MLCZEmlcze]|-?[0-9.]+(?:e[-+]?\d+)?/g) || [];
  const out = [];
  let i = 0;
  let cmd = null;
  while (i < toks.length) {
    const t = toks[i];
    if (/^[A-Za-z]$/.test(t)) {
      cmd = t.toUpperCase();
      i++;
      if (cmd === 'Z' || cmd === 'E') { out.push([cmd === 'Z' ? 'Z' : 'E']); continue; }
    }
    if (!cmd) { i++; continue; }
    const need = cmd === 'C' ? 6 : 2;
    const nums = toks.slice(i, i + need).map(Number);
    if (nums.length < need || nums.some((x) => Number.isNaN(x))) break;
    out.push([cmd, ...nums]);
    i += need;
    if (cmd === 'M') cmd = 'L';
  }
  return out.filter((c) => c[0] !== 'E');
}

class TimingParser {
  constructor(ctx, slide) {
    this.ctx = ctx;
    this.slide = slide;
    this.effectSeq = 0;
  }

  /**
   * Flatten a time node into behaviour instances.
   * @param node p:par/p:seq/p:excl/behaviour element
   * @param t0 begin time of the parent (ms, relative to the step)
   * @param wrappers repeat / autoRev wrappers of enclosing containers
   * @param effect enclosing preset effect info
   * @returns {number} end time (ms, may be Infinity)
   */
  walk(node, t0, wrappers, effect, out) {
    const local = node.local;
    if (local === 'par' || local === 'seq' || local === 'excl') {
      const cTn = node.child('p:cTn');
      const info = cTnInfo(cTn);
      const delay = this.delayOf(info);
      const begin = t0 + delay;
      let eff = effect;
      if (info.presetClass || (info.nodeType && /Effect$/.test(info.nodeType))) {
        eff = {
          id: this.effectSeq++,
          cls: info.presetClass || 'custom',
          preset: info.presetID,
          subtype: info.presetSubtype,
          nodeType: info.nodeType,
          grpId: info.grpId,
        };
      }
      const kids = cTn && cTn.child('p:childTnLst') ? cTn.child('p:childTnLst').children : [];
      let wr = wrappers;
      if (this.needsWrap(info)) {
        let single = info.dur;
        if (single === undefined) {
          // Natural duration of the container: dry run over the children.
          const saved = this.effectSeq;
          single = this.walkChildren(local, kids, begin, wrappers, eff, []) - begin;
          this.effectSeq = saved;
        }
        if (single !== Infinity && single > 0) {
          wr = [...wrappers, { begin, dur: single, repeat: info.repeat, autoRev: info.autoRev, accel: info.accel, decel: info.decel, spd: info.spd }];
        }
      }
      const startIdx = out.length;
      const end = this.walkChildren(local, kids, begin, wr, eff, out);
      if (info.iterate && kids.length) this.applyIterate(info.iterate, out.slice(startIdx), eff);
      const single = info.dur !== undefined ? info.dur : end - begin;
      return begin + this.activeDur(single, info);
    }
    // Behaviours.
    const b = this.behaviour(node, effect);
    if (!b) return t0;
    const info = b.info;
    delete b.info;
    b.begin = t0 + this.delayOf(info);
    b.wrappers = wrappers;
    out.push(b);
    if (info.iterate) this.applyIterate(info.iterate, [b], effect);
    return b.begin + this.activeDur(b.dur, b);
  }

  delayOf(info) {
    let d = 0;
    for (const c of info.conds) {
      if (c.delay !== Infinity && !c.evt) d = Math.max(d, c.delay);
      else if (c.evt === 'onBegin' || c.evt === 'begin') d = Math.max(d, c.delay === Infinity ? 0 : c.delay);
    }
    return d;
  }

  activeDur(dur, info) {
    if (dur === undefined) return 0;
    if (dur === Infinity) return Infinity;
    let d = dur / (info.spd || 1);
    if (info.autoRev) d *= 2;
    if (info.repeatDur !== undefined) return info.repeatDur;
    return d * (info.repeat === undefined ? 1 : info.repeat);
  }

  needsWrap(info) {
    return (info.repeat !== undefined && info.repeat !== 1) || info.autoRev || info.accel || info.decel || (info.spd && info.spd !== 1);
  }

  walkChildren(local, kids, begin, wrappers, eff, out) {
    if (local === 'seq') {
      let t = begin;
      for (const k of kids) {
        t = this.walk(k, t, wrappers, eff, out);
        if (t === Infinity) break;
      }
      return t;
    }
    let max = begin;
    for (const k of kids) {
      const e = this.walk(k, begin, wrappers, eff, out);
      if (e > max) max = e;
    }
    return max;
  }

  applyIterate(it, items, eff) {
    // Text animated by word / letter: the runtime splits the text; record
    // the iteration so that each sub-element gets a staggered start. The
    // percentage delay is relative to the duration of the whole effect.
    const sel = items.filter((b) => !eff || b.effect === eff.id);
    let base = 0;
    for (const b of sel) if (b.dur !== Infinity) base = Math.max(base, b.dur);
    for (const b of sel) b.iterate = { ...it, base };
  }

  behaviour(node, effect) {
    const ctx = this.ctx;
    const cBhvr = node.child('p:cBhvr');
    const cTn = cBhvr ? cBhvr.child('p:cTn') : node.first('p:cTn');
    const info = cTnInfo(cTn);
    const tgt = targetOf(cBhvr ? cBhvr.child('p:tgtEl') : node.first('p:tgtEl'), ctx);
    const b = {
      type: null,
      info,
      dur: info.dur !== undefined ? info.dur : (node.local === 'set' ? 1 : 500),
      repeat: info.repeat,
      repeatDur: info.repeatDur,
      autoRev: info.autoRev,
      accel: info.accel,
      decel: info.decel,
      spd: info.spd,
      // SMIL: elements with a duration default to fill="remove"; sets are held.
      fill: info.fill || (node.local === 'set' ? 'hold' : 'remove'),
      effect: effect ? effect.id : null,
      cls: effect ? effect.cls : null,
      target: tgt ? tgt.kind : null,
      spid: tgt ? tgt.spid : null,
      keys: tgt ? tgt.keys : [],
      additive: cBhvr ? cBhvr.attr('additive', null) : null,
      attrs: attrNames(cBhvr),
    };
    if (tgt && tgt.chartEl) b.chartEl = tgt.chartEl;
    if (tgt && tgt.chars) b.chars = tgt.chars;
    switch (node.local) {
      case 'set': {
        b.type = 'set';
        b.to = tavValue(node.child('p:to'), ctx);
        return b;
      }
      case 'anim': {
        b.type = 'anim';
        b.calcmode = node.attr('calcmode', 'lin');
        b.valueType = node.attr('valueType', 'num');
        b.by = node.attrs.by;
        b.from = node.attrs.from;
        b.to = node.attrs.to;
        const tl = node.child('p:tavLst');
        b.tav = tl ? tl.all('p:tav').map((t) => ({
          tm: t.attrs.tm === INDEF || t.attrs.tm === undefined ? null : t.num('tm', 0) / 100000,
          val: tavValue(t.child('p:val'), ctx),
          fmla: t.attrs.fmla || null,
        })) : [];
        // Fill missing key times evenly.
        if (b.tav.length && b.tav.some((t) => t.tm === null)) {
          const n = b.tav.length;
          b.tav.forEach((t, i) => { if (t.tm === null) t.tm = n > 1 ? i / (n - 1) : 0; });
        }
        return b;
      }
      case 'animEffect': {
        b.type = 'effect';
        b.transition = node.attr('transition', 'in');
        b.filter = node.attr('filter', '');
        const pr = node.child('p:progress');
        if (pr) b.progress = tavValue(pr, ctx);
        return b;
      }
      case 'animMotion': {
        b.type = 'motion';
        b.origin = node.attr('origin', 'layout');
        b.pathEditMode = node.attr('pathEditMode', 'relative');
        b.path = parseMotionPath(node.attr('path', ''));
        const pt = (n) => (n ? [n.num('x', 0) / 100000, n.num('y', 0) / 100000] : null);
        b.by = pt(node.child('p:by'));
        b.from = pt(node.child('p:from'));
        b.to = pt(node.child('p:to'));
        b.rAng = node.num('rAng', 0) / 60000;
        return b;
      }
      case 'animRot': {
        b.type = 'rotate';
        b.by = node.attrs.by !== undefined ? node.num('by', 0) / 60000 : undefined;
        b.from = node.attrs.from !== undefined ? node.num('from', 0) / 60000 : undefined;
        b.to = node.attrs.to !== undefined ? node.num('to', 0) / 60000 : undefined;
        return b;
      }
      case 'animScale': {
        b.type = 'scale';
        const pt = (n) => (n ? [n.num('x', 100000) / 100000, n.num('y', 100000) / 100000] : null);
        b.by = pt(node.child('p:by'));
        b.from = pt(node.child('p:from'));
        b.to = pt(node.child('p:to'));
        b.zoomContents = node.bool('zoomContents', false);
        return b;
      }
      case 'animClr': {
        b.type = 'color';
        b.clrSpc = node.attr('clrSpc', 'rgb');
        b.dir = node.attr('dir', 'cw');
        b.by = clrSpec(node.child('p:by'), ctx);
        b.from = clrSpec(node.child('p:from'), ctx);
        b.to = clrSpec(node.child('p:to'), ctx);
        return b;
      }
      case 'cmd': {
        b.type = 'cmd';
        b.cmdType = node.attr('type', 'call');
        b.cmd = node.attr('cmd', '');
        return b;
      }
      case 'audio':
      case 'video': {
        const cm = node.child('p:cMediaNode');
        const ci = cTnInfo(cm && cm.child('p:cTn'));
        const t2 = targetOf(cm && cm.child('p:tgtEl'), ctx);
        return {
          type: 'media',
          info: ci,
          media: node.local,
          dur: ci.dur !== undefined ? ci.dur : 0,
          repeat: 1, autoRev: false, accel: 0, decel: 0, spd: 1,
          fill: 'hold',
          effect: effect ? effect.id : null,
          cls: effect ? effect.cls : null,
          target: t2 ? t2.kind : null,
          spid: t2 ? t2.spid : null,
          keys: t2 ? t2.keys : [],
          sound: t2 && t2.sound ? t2.sound : null,
          vol: cm ? cm.num('vol', 80000) / 100000 : 0.8,
          mute: cm ? cm.bool('mute', false) : false,
          numSld: cm ? cm.int('numSld', 1) : 1,
          showWhenStopped: cm ? cm.bool('showWhenStopped', true) : true,
          fullScrn: node.local === 'video' ? node.bool('fullScrn', false) : false,
        };
      }
      default:
        if (node.local !== 'extLst') ctx.note('unsupported', 'animation node ' + node.name);
        return null;
    }
  }

  /** Build steps from a sequence's children (click groups). */
  steps(seqCTn, kind) {
    const kids = seqCTn.child('p:childTnLst') ? seqCTn.child('p:childTnLst').children : [];
    const steps = [];
    for (const par of kids) {
      const cTn = par.child('p:cTn');
      const info = cTnInfo(cTn);
      const auto = info.conds.some((c) => (c.evt === 'onBegin' || c.evt === 'begin') && c.delay !== Infinity) ||
        (info.conds.length > 0 && info.conds.every((c) => c.delay !== Infinity && !c.evt));
      const delay = info.conds.reduce((d, c) => (c.delay !== Infinity ? Math.max(d, c.delay) : d), 0);
      const behaviours = [];
      // The click group itself: walk its children at t=0.
      const childKids = cTn && cTn.child('p:childTnLst') ? cTn.child('p:childTnLst').children : [];
      let end = 0;
      for (const k of childKids) {
        const e = this.walk(k, 0, [], null, behaviours);
        if (e > end) end = e;
      }
      steps.push({ auto: auto && kind === 'main', delay: auto ? delay : 0, behaviours, dur: end });
    }
    return steps;
  }
}

/**
 * Parse the timing of a slide.
 * @returns {object|null} {main: steps[], triggers: [{spid, steps}], hidden: keys[], builds}
 */
export function parseTiming(sld, ctx, slide) {
  const timing = sld.child('p:timing');
  if (!timing) {
    const ac = sld.child('mc:AlternateContent');
    if (!ac) return null;
  }
  const t = timing || null;
  if (!t) return null;
  const parser = new TimingParser(ctx, slide);
  const root = t.path('p:tnLst', 'p:par', 'p:cTn');
  const result = { main: [], triggers: [], hidden: [], builds: {}, textBuilds: {} };
  if (root) {
    const kids = root.child('p:childTnLst') ? root.child('p:childTnLst').children : [];
    for (const seq of kids) {
      if (seq.local !== 'seq') continue;
      const cTn = seq.child('p:cTn');
      if (!cTn) continue;
      const nt = cTn.attr('nodeType', '');
      if (nt === 'mainSeq') {
        result.main = parser.steps(cTn, 'main');
      } else if (nt === 'interactiveSeq') {
        const info = cTnInfo(cTn);
        const trig = info.conds.find((c) => c.spid);
        const bm = info.conds.find((c) => c.evt === 'onMediaBookmark');
        result.triggers.push({
          spid: trig ? trig.spid : null,
          evt: trig ? trig.evt : (bm ? 'bookmark' : null),
          steps: parser.steps(cTn, 'trigger'),
        });
      }
    }
  }
  // Build list.
  const bl = t.child('p:bldLst');
  if (bl) {
    for (const b of bl.children) {
      const spid = b.attr('spid', '');
      if (b.local === 'bldP') result.builds[spid] = { type: 'p', build: b.attr('build', 'whole'), animBg: b.bool('animBg', false), grpId: b.attr('grpId', '0'), rev: b.bool('rev', false) };
      else if (b.local === 'bldGraphic') result.builds[spid] = { type: 'graphic', asOne: !!b.child('p:bldAsOne') };
      else if (b.local === 'bldOleChart' || b.local === 'bldDgm') result.builds[spid] = { type: b.local, bld: b.attr('bld', 'allAtOnce') };
    }
  }
  // Initially hidden targets: first effect of a target is an entrance.
  const seen = new Set();
  const hidden = new Set();
  const visit = (steps) => {
    for (const st of steps) {
      // Order by begin time within the step.
      const byEffect = new Map();
      for (const b of st.behaviours) {
        if (b.effect === null || b.effect === undefined) continue;
        if (!byEffect.has(b.effect)) byEffect.set(b.effect, { begin: b.begin, cls: b.cls, keys: new Set() });
        const e = byEffect.get(b.effect);
        e.begin = Math.min(e.begin, b.begin);
        for (const k of b.keys) e.keys.add(k);
      }
      const effects = [...byEffect.values()].sort((a, b) => a.begin - b.begin);
      for (const e of effects) {
        for (const k of e.keys) {
          const spid = k.split(':')[0];
          const isSub = k.includes(':');
          if (seen.has(k) || (isSub && seen.has(spid))) continue;
          seen.add(k);
          if (e.cls === 'entr') hidden.add(k);
        }
      }
    }
  };
  visit(result.main);
  for (const tr of result.triggers) visit(tr.steps);
  result.hidden = [...hidden];
  // Paragraph builds: list the paragraph indices animated per shape.
  for (const st of [...result.main, ...result.triggers.flatMap((x) => x.steps)]) {
    for (const b of st.behaviours) {
      for (const k of b.keys) {
        const m = /^(.+):p(\d+)$/.exec(k);
        if (m) (result.textBuilds[m[1]] || (result.textBuilds[m[1]] = new Set())).add(Number(m[2]));
      }
    }
  }
  for (const k of Object.keys(result.textBuilds)) result.textBuilds[k] = [...result.textBuilds[k]].sort((a, b) => a - b);
  if (!result.main.length && !result.triggers.length) return null;
  void slide;
  return result;
}

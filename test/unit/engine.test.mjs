import test from 'node:test';
import assert from 'node:assert/strict';
import { evalFormula } from '../../src/runtime/formula.js';
import { ease, behaviourProgress, evaluate, pathPoint } from '../../src/runtime/engine.js';
import { parseMotionPath } from '../../src/pptx/timing.js';
import { filterStyle, transitionFrame, TRANSITION_TYPES, FILTERS } from '../../src/runtime/effects.js';

test('animation formulas', () => {
  const v = { '#ppt_x': 0.25, '#ppt_w': 0.1, '#ppt_h': 0.2, $: 0.5 };
  assert.equal(evalFormula('#ppt_x', v), 0.25);
  assert.ok(Math.abs(evalFormula('0-#ppt_w/2', v) + 0.05) < 1e-12);
  assert.ok(Math.abs(evalFormula('1+#ppt_h/2', v) - 1.1) < 1e-12);
  assert.ok(Math.abs(evalFormula('sin(pi*$)', v) - 1) < 1e-12);
  assert.equal(evalFormula('2^3+max(1,4)', v), 12);
  assert.ok(Number.isNaN(evalFormula('(', v)));
});

test('ease (accel/decel) is monotonic and preserves endpoints', () => {
  for (const [a, d] of [[0, 0], [0.5, 0.5], [0.2, 0.3], [1, 0]]) {
    assert.equal(ease(0, a, d), 0);
    assert.ok(Math.abs(ease(1, a, d) - 1) < 1e-9);
    let prev = 0;
    for (let p = 0; p <= 1; p += 0.05) { const e = ease(p, a, d); assert.ok(e >= prev - 1e-9); prev = e; }
  }
});

test('behaviour progress: repeat, auto reverse, fill', () => {
  const b = { begin: 100, dur: 1000, repeat: 2, autoRev: true, fill: 'hold', wrappers: [] };
  assert.equal(behaviourProgress(b, 50), null);
  assert.equal(behaviourProgress(b, 600).p, 0.5);
  assert.ok(Math.abs(behaviourProgress(b, 1600).p - 0.5) < 1e-9);
  assert.equal(behaviourProgress(b, 100000).p, 0);
  assert.equal(behaviourProgress({ ...b, fill: 'remove' }, 100000), null);
  // Container wrapper repeats the child 3 times.
  const c = { begin: 0, dur: 100, fill: 'hold', wrappers: [{ begin: 0, dur: 100, repeat: 3 }] };
  assert.equal(behaviourProgress(c, 250).p, 0.5);
});

test('motion paths', () => {
  const p = parseMotionPath('M 0 0 L 0.5 0 L 0.5 -0.25 E');
  assert.deepEqual(p, [['M', 0, 0], ['L', 0.5, 0], ['L', 0.5, -0.25]]);
  assert.deepEqual(pathPoint(p, 1), [0.5, -0.25]);
  const mid = pathPoint(p, 0.5, 1);
  assert.ok(Math.abs(mid[0] - 0.375) < 1e-9);
});

test('evaluate: initial hidden state, set + fade, completed steps', () => {
  const timing = {
    hidden: ['3'],
    main: [{ auto: false, behaviours: [
      { type: 'set', keys: ['3'], attrs: ['style.visibility'], to: 'visible', begin: 0, dur: 1, fill: 'hold', wrappers: [] },
      { type: 'effect', keys: ['3'], filter: 'fade', transition: 'in', begin: 0, dur: 500, fill: 'remove', wrappers: [] },
    ] }],
    triggers: [],
  };
  const geo = () => ({ x: 0.5, y: 0.5, w: 0.1, h: 0.1 });
  assert.equal(evaluate(timing, { step: 0 }, geo).states.get('3').visible, false);
  const mid = evaluate(timing, { step: 0, t: 250 }, geo).states.get('3');
  assert.equal(mid.visible, true);
  assert.equal(mid.effects[0].reveal, 0.5);
  const done = evaluate(timing, { step: 1 }, geo).states.get('3');
  assert.equal(done.visible, true);
  assert.equal(done.effects.length, 0);
});

test('effect styles exist for every filter and transition', () => {
  for (const f of FILTERS) {
    const st = filterStyle(f, 0.5);
    assert.ok(Object.keys(st).length > 0, f);
    assert.deepEqual(filterStyle(f, 1), {});
  }
  for (const t of TRANSITION_TYPES) {
    const fr = transitionFrame({ type: t, dir: 'l' }, 0.5);
    assert.ok(fr.inc && fr.out, t);
  }
});

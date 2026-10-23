import test from 'node:test';
import assert from 'node:assert/strict';
import { formatValue } from '../../src/chart/numfmt.js';
import { autoScale, fitTrend } from '../../src/chart/render.js';

test('number formats', () => {
  const cases = [[1234.5, '#,##0.00', '1,234.50'], [0.256, '0%', '26%'], [0.256, '0.0%', '25.6%'], [-12.3, '0.0;(0.0)', '(12.3)'], [42, '"$"#,##0', '$42'],
    [45000, 'm/d/yyyy', '3/15/2023'], [45000, 'mmm-yy', 'Mar-23'], [12345.678, '0.00E+00', '1.23E+04'], [0, '0.0;-0.0;"zero"', 'zero'], [1500000, '#,##0,,"M"', '2M'],
    [1 / 3, 'General', '0.3333333333'], [5, '0', '5'], [44927.5, 'h:mm AM/PM', '12:00 PM']];
  for (const [v, f, exp] of cases) assert.equal(formatValue(v, f), exp, f);
});

test('axis auto scaling', () => {
  assert.deepEqual(autoScale(0, 4.5, {}), { min: 0, max: 5, unit: 1 });
  const s = autoScale(-3, 17, {});
  assert.ok(s.min <= -3 && s.max >= 17);
  assert.equal(autoScale(0, 1, { max: 2, majorUnit: 0.5 }).max, 2);
});

test('trend line fitting', () => {
  const f = fitTrend({ type: 'linear' }, [[1, 3], [2, 5], [3, 7]]);
  assert.ok(Math.abs(f(4) - 9) < 1e-9);
  const p = fitTrend({ type: 'poly', order: 2 }, [[0, 1], [1, 2], [2, 5], [3, 10]]);
  assert.ok(Math.abs(p(4) - 17) < 1e-6);
});

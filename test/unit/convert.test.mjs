import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { convert, loadPresentation } from '../../src/index.js';
import { buildFixtures } from '../fixtures/build-fixtures.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const decks = buildFixtures(path.join(here, '../fixtures/out'));

test('fixture decks convert without errors or unsupported elements', () => {
  for (const d of decks) {
    const bytes = fs.readFileSync(d.file);
    const { html, presentation, diagnostics } = convert(bytes);
    assert.match(html, /^<!DOCTYPE html>/);
    assert.equal(presentation.slides.length, d.slideCount);
    const bad = diagnostics.filter((x) => x.kind === 'error' || x.kind === 'unsupported');
    assert.deepEqual(bad, [], d.name);
    assert.equal((html.match(/<section class="pp-slide"/g) || []).length, d.slideCount);
  }
});

test('model checks of the feature catalog', () => {
  for (const d of decks) {
    const pres = loadPresentation(fs.readFileSync(d.file));
    for (const f of d.features) {
      for (const c of f.checks) {
        if (c.kind !== 'model') continue;
        assert.equal(c.fn(pres.slides[f.slide - 1], pres), true, `${f.feature}: ${c.desc}`);
      }
    }
  }
});

test('legacy sample deck converts', () => {
  const file = path.join(here, '../../samples/sample.pptx');
  if (!fs.existsSync(file)) return;
  const { presentation } = convert(fs.readFileSync(file));
  assert.ok(presentation.slides.length > 0);
});

test('non-zip input fails with a clear message', () => {
  assert.throws(() => convert(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23])), /ZIP/);
});

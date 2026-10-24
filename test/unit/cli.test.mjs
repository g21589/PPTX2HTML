import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { version } from '../../src/index.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const bin = path.join(here, '../../bin/pptx2html.js');
const sample = path.join(here, '../../samples/sample.pptx');
const run = (...args) => execFileSync(process.execPath, [bin, ...args], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

test('cli prints version and rejects unknown options', () => {
  assert.equal(run('--version').trim(), version);
  const r = spawnSync(process.execPath, [bin, '--nope'], { encoding: 'utf8' });
  assert.equal(r.status, 2);
  assert.match(r.stderr, /unknown option --nope/);
});

test('cli converts with slide selection and external assets', { skip: !fs.existsSync(sample) }, () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx2html-cli-'));
  try {
    const out = path.join(dir, 'deck.html');
    run(sample, '-o', out, '--slides', '1-2', '--assets', '-q');
    const html = fs.readFileSync(out, 'utf8');
    assert.equal((html.match(/<section class="pp-slide/g) || []).length, 2);
    assert.match(html, /PPTXPlayer/);
    const stdout = run(sample, '-o', '-', '--static', '-q');
    assert.ok(stdout.startsWith('<!DOCTYPE html>'));
    assert.doesNotMatch(stdout, /PPTXPlayer/);
    // Every slide with a picture now references a file in deck_files/.
    run(sample, '-o', out, '--assets', '-q');
    const full = fs.readFileSync(out, 'utf8');
    const refs = [...full.matchAll(/(?:href|src)="(deck_files\/[^"]+)"/g)].map((m) => m[1]);
    assert.ok(refs.length > 0);
    for (const r of refs) assert.ok(fs.existsSync(path.join(dir, decodeURIComponent(r))), r);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

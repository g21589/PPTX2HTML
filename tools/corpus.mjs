#!/usr/bin/env node
/** Convert every .pptx under the given directories; report errors and diagnostics. */
import fs from 'node:fs';
import path from 'node:path';
import { loadPresentation, renderDocument } from '../src/index.js';

const dirs = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const verbose = process.argv.includes('--verbose');
const files = [];
for (const d of dirs) {
  const walk = (p) => {
    for (const f of fs.readdirSync(p)) {
      const full = path.join(p, f);
      if (fs.statSync(full).isDirectory()) walk(full);
      else if (/\.(pptx|pptm|ppsx|potx)$/i.test(f)) files.push(full);
    }
  };
  walk(d);
}
const agg = new Map();
let ok = 0, failed = 0, slides = 0;
const t0 = Date.now();
for (const f of files) {
  if (process.argv.includes("--trace")) console.log("..", f);
  try {
    const pres = loadPresentation(fs.readFileSync(f));
    renderDocument(pres, { mode: 'list' });
    slides += pres.slides.length;
    ok++;
    for (const d of pres.diagnostics.list()) {
      const k = d.kind + ': ' + d.message.replace(/rId\d+|\/ppt\/[^ ]+|\{[0-9A-F-]+\}/g, '*');
      const a = agg.get(k) || { n: 0, files: [] };
      a.n++;
      if (a.files.length < 3) a.files.push(path.basename(f));
      agg.set(k, a);
    }
  } catch (e) {
    failed++;
    console.log('FAIL', path.basename(f), '-', e.message.split('\n')[0]);
    if (verbose) console.log(e.stack);
  }
}
console.log(`\n${ok} ok, ${failed} failed, ${slides} slides, ${Date.now() - t0} ms`);
const sorted = [...agg.entries()].sort((a, b) => b[1].n - a[1].n);
for (const [k, v] of sorted) console.log(String(v.n).padStart(4), k, '  e.g.', v.files.join(', '));

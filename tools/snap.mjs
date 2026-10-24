#!/usr/bin/env node
/**
 * Render a PPTX with PPTX2HTML and screenshot every slide with Chromium.
 *   node tools/snap.mjs deck.pptx outDir [--slides 1,2,3] [--lo]
 * --lo also renders reference PNGs with LibreOffice (outDir/lo-N.png).
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { chromium } from 'playwright';
import { loadPresentation, renderDocument } from '../src/index.js';
import { PLAYER_SOURCE } from '../src/generated/player-source.js';

const args = process.argv.slice(2);
const file = args[0];
const out = args[1] || 'out';
const slidesArg = args.indexOf('--slides') >= 0 ? args[args.indexOf('--slides') + 1].split(',').map((n) => Number(n) - 1) : null;
fs.mkdirSync(out, { recursive: true });
const pres = loadPresentation(fs.readFileSync(file), slidesArg ? { slides: slidesArg } : {});
const html = renderDocument(pres, { mode: 'list', player: PLAYER_SOURCE, extraCss: '.pp-list .pp-slide{margin:0;box-shadow:none} body.pp-body{background:#fff}' });
const htmlPath = path.join(out, 'index.html');
fs.writeFileSync(htmlPath, html);
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: Math.ceil(pres.width), height: Math.ceil(pres.height) } });
page.on('console', (m) => { if (m.type() === 'error') console.log('console:', m.text()); });
page.on('pageerror', (e) => console.log('pageerror:', e.message));
await page.goto('file://' + path.resolve(htmlPath));
await page.evaluate(() => document.fonts.ready);
const sections = await page.$$('section.pp-slide');
for (const s of sections) {
  const n = await s.getAttribute('data-number');
  await s.screenshot({ path: path.join(out, `pp-${n}.png`) });
}
await browser.close();
console.log('rendered', sections.length, 'slides', JSON.stringify(pres.diagnostics.list()));
if (args.includes('--lo')) {
  const tmp = path.join(out, 'lo');
  fs.mkdirSync(tmp, { recursive: true });
  execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', tmp, path.resolve(file)], { stdio: 'ignore', timeout: 180000 });
  const pdf = path.join(tmp, path.basename(file).replace(/\.[^.]+$/, '.pdf'));
  execFileSync('pdftoppm', ['-png', '-r', '72', '-scale-to-x', String(Math.ceil(pres.width)), '-scale-to-y', '-1', pdf, path.join(out, 'lo')]);
  console.log('libreoffice reference rendered');
}

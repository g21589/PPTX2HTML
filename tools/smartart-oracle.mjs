#!/usr/bin/env node
/**
 * Checks the SmartArt layout engine against PowerPoint: decks saved by
 * PowerPoint 2010+ contain a pre-rendered drawing of every diagram. Each
 * diagram slide is rendered twice (with the drawing, and laid out by the
 * engine from the layout definition) and the screenshots are compared.
 *
 *   node tools/smartart-oracle.mjs outDir deck.pptx ...
 *
 * Writes outDir/<deck>-ppt.png, -eng.png and results.json.
 */
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';
import { PNG } from 'pngjs';
import { loadPresentation, renderDocument } from '../src/index.js';
import { PLAYER_SOURCE } from '../src/generated/player-source.js';

const [out, ...files] = process.argv.slice(2);
if (!out || !files.length) {
  console.error('usage: node tools/smartart-oracle.mjs outDir deck.pptx ...');
  process.exit(2);
}
fs.mkdirSync(out, { recursive: true });
const hasDiagram = (e) => e.smartArt || (e.children || []).some(hasDiagram);
const browser = await chromium.launch();
const results = [];
for (const file of files) {
  const name = path.basename(file).replace(/\.pptx$/i, '');
  const bytes = fs.readFileSync(file);
  const idx = loadPresentation(bytes).slides.findIndex((s) => s.elements.some(hasDiagram));
  if (idx < 0) continue;
  const shots = [];
  for (const [tag, opts] of [['ppt', {}], ['eng', { smartArtEngine: true }]]) {
    const pres = loadPresentation(bytes, { slides: [idx], ...opts });
    const html = path.join(out, `${name}-${tag}.html`);
    fs.writeFileSync(html, renderDocument(pres, { mode: 'list', player: PLAYER_SOURCE, extraCss: '.pp-list .pp-slide{margin:0;box-shadow:none} body.pp-body{background:#fff}' }));
    const page = await browser.newPage({ viewport: { width: Math.ceil(pres.width), height: Math.ceil(pres.height) } });
    await page.goto('file://' + path.resolve(html));
    await page.evaluate(() => document.fonts.ready);
    const png = path.join(out, `${name}-${tag}.png`);
    await (await page.$('section.pp-slide')).screenshot({ path: png });
    await page.close();
    shots.push(PNG.sync.read(fs.readFileSync(png)));
  }
  // Share of sampled pixels that match (ignoring small colour differences).
  const [a, b] = shots;
  let same = 0, n = 0;
  for (let i = 0; i < Math.min(a.data.length, b.data.length); i += 16) {
    const d = Math.abs(a.data[i] - b.data[i]) + Math.abs(a.data[i + 1] - b.data[i + 1]) + Math.abs(a.data[i + 2] - b.data[i + 2]);
    if (d <= 60) same++;
    n++;
  }
  results.push({ deck: name, slide: idx + 1, similarity: same / n });
  console.log(`${(same / n).toFixed(3)}  ${name} (slide ${idx + 1})`);
}
await browser.close();
fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(results, null, 2));
if (results.length) console.log('mean', (results.reduce((s, r) => s + r.similarity, 0) / results.length).toFixed(3));

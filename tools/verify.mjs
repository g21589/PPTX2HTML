#!/usr/bin/env node
/**
 * Self-verification of PPTX2HTML.
 *
 *   node tools/verify.mjs [--no-lo] [--only feature-substring] [--corpus dir ...] [--out verify-report]
 *
 * 1. Generates the feature fixture decks (test/fixtures) covering every
 *    supported PPTX feature.
 * 2. Converts them, runs model checks (Node), DOM / text / pixel checks on
 *    Chromium screenshots, and animation / transition checks through the
 *    runtime player.
 * 3. Renders the same decks with LibreOffice and scores visual similarity.
 * 4. Writes an HTML + JSON report. Exit code 1 when a check fails.
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { PNG } from 'pngjs';
import { chromium } from 'playwright';
import { loadPresentation, renderDocument } from '../src/index.js';
import { PLAYER_SOURCE } from '../src/generated/player-source.js';
import { buildFixtures } from '../test/fixtures/build-fixtures.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const argv = process.argv.slice(2);
const opt = (name, def) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : def; };
const outDir = path.resolve(opt('--out', path.join(root, 'verify-report')));
const only = opt('--only', null);
const useLO = !argv.includes('--no-lo') && hasCmd('soffice') && hasCmd('pdftoppm');
const corpusDirs = [];
argv.forEach((a, i) => { if (a === '--corpus') corpusDirs.push(argv[i + 1]); });

function hasCmd(c) {
  try { execFileSync('which', [c], { stdio: 'ignore' }); return true; } catch (e) { return false; }
}

// Start from a clean report but keep LibreOffice references (re-rendered when stale).
fs.mkdirSync(outDir, { recursive: true });
for (const f of fs.readdirSync(outDir)) if (f !== 'lo') fs.rmSync(path.join(outDir, f), { recursive: true, force: true });
fs.mkdirSync(path.join(outDir, 'img'), { recursive: true });

// ---- image helpers ----------------------------------------------------------------
function readPng(file) {
  return PNG.sync.read(fs.readFileSync(file));
}

function samplePx(img, x, y, r = 1) {
  let R = 0, G = 0, B = 0, n = 0;
  for (let dy = -r; dy <= r; dy++) for (let dx = -r; dx <= r; dx++) {
    const xx = Math.round(x) + dx, yy = Math.round(y) + dy;
    if (xx < 0 || yy < 0 || xx >= img.width || yy >= img.height) continue;
    const o = (yy * img.width + xx) * 4;
    R += img.data[o]; G += img.data[o + 1]; B += img.data[o + 2]; n++;
  }
  return n ? [R / n, G / n, B / n] : [0, 0, 0];
}

const hex2rgb = (h) => { const n = parseInt(h.replace('#', ''), 16); return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
const rgb2hex = (c) => '#' + c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();

/** Visual similarity of two images (0..1) on a downscaled grid. */
function similarity(a, b, gw = 160, gh = 90) {
  const cell = (img, gx, gy) => {
    const x0 = Math.floor((gx * img.width) / gw), x1 = Math.max(x0 + 1, Math.floor(((gx + 1) * img.width) / gw));
    const y0 = Math.floor((gy * img.height) / gh), y1 = Math.max(y0 + 1, Math.floor(((gy + 1) * img.height) / gh));
    let r = 0, g = 0, bl = 0, n = 0;
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
      const o = (y * img.width + x) * 4;
      r += img.data[o]; g += img.data[o + 1]; bl += img.data[o + 2]; n++;
    }
    return [r / n, g / n, bl / n];
  };
  let err = 0;
  for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
    const p = cell(a, gx, gy), q = cell(b, gx, gy);
    err += (Math.abs(p[0] - q[0]) + Math.abs(p[1] - q[1]) + Math.abs(p[2] - q[2])) / 3;
  }
  return 1 - err / (gw * gh * 255);
}

function diffFraction(a, b) {
  let n = 0;
  const len = Math.min(a.data.length, b.data.length);
  for (let i = 0; i < len; i += 4) {
    if (Math.abs(a.data[i] - b.data[i]) + Math.abs(a.data[i + 1] - b.data[i + 1]) + Math.abs(a.data[i + 2] - b.data[i + 2]) > 12) n++;
  }
  return n / (len / 4);
}

// ---- LibreOffice reference ---------------------------------------------------------
function libreOfficeRender(file, W, H, dir) {
  fs.mkdirSync(dir, { recursive: true });
  const pdf = path.join(dir, path.basename(file).replace(/\.[^.]+$/, '.pdf'));
  // Reuse a reference rendered from the same input (LibreOffice is slow).
  const cached = fs.existsSync(pdf) && fs.statSync(pdf).mtimeMs >= fs.statSync(file).mtimeMs;
  if (!cached) {
    const filter = 'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}';
    try {
      execFileSync('soffice', ['--headless', '--convert-to', filter, '--outdir', dir, file], { stdio: 'ignore', timeout: 300000 });
    } catch (e) {
      return null;
    }
    if (!fs.existsSync(pdf)) return null;
  }
  for (const f of fs.readdirSync(dir)) if (/^p-\d+\.png$/.test(f)) fs.rmSync(path.join(dir, f));
  execFileSync('pdftoppm', ['-png', '-scale-to-x', String(Math.round(W)), '-scale-to-y', String(Math.round(H)), pdf, path.join(dir, 'p')], { stdio: 'ignore' });
  const pages = fs.readdirSync(dir).filter((f) => /^p-\d+\.png$/.test(f)).sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
  return pages.map((p) => path.join(dir, p));
}

// ---- browser checks ----------------------------------------------------------------
async function screenshotSlides(page, htmlFile, pres) {
  await page.goto('file://' + htmlFile);
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(150);
  const shots = [];
  for (const s of pres.slides) {
    const el = await page.$(`section#slide-${s.index + 1}`);
    const file = path.join(outDir, 'img', `${pres.__name}-${s.index + 1}.png`);
    await el.screenshot({ path: file });
    shots.push(file);
  }
  return shots;
}

async function domCheck(page, slideNo, c) {
  return page.evaluate(({ slideNo, c }) => {
    const sec = document.querySelector('section#slide-' + slideNo);
    if (!sec) return { ok: false, msg: 'slide missing' };
    if (c.kind === 'text') {
      const t = sec.textContent.replace(/​/g, '');
      return { ok: t.includes(c.contains), msg: t.includes(c.contains) ? '' : 'text not found: ' + c.contains };
    }
    if (c.kind === 'eval') {
      // eslint-disable-next-line no-eval
      const v = (0, eval)('(' + c.fn + ')')(sec);
      return { ok: v === true, msg: v === true ? c.desc : `${c.desc}: ${v}` };
    }
    const n = sec.querySelectorAll(c.selector).length;
    const ok = n >= (c.min || 1) && (c.max === undefined || n <= c.max);
    return { ok, msg: `${n} × ${c.selector}` };
  }, { slideNo, c: { ...c, fn: c.fn ? String(c.fn) : undefined } });
}

async function animState(page, c, slideIdx) {
  return page.evaluate(({ c, slideIdx }) => {
    const p = window.pptxPlayer;
    p.seek(slideIdx, c.step, c.t);
    const sec = document.querySelector('section#slide-' + (slideIdx + 1));
    let el = sec.querySelector(`.pp-el[data-layer="slide"][data-spid="${c.spid}"]`);
    if (!el) return { error: 'element not found' };
    if (c.para !== undefined) el = el.querySelector(`.pp-p[data-para="${c.para}"]`);
    if (c.letter !== undefined) el = el && el.querySelectorAll('.pp-it')[c.letter];
    if (!el) return { error: 'sub element not found' };
    let visible = true, opacity = 1, node = el;
    while (node && node !== sec) {
      const cs = getComputedStyle(node);
      if (cs.visibility === 'hidden' || cs.display === 'none') visible = false;
      opacity *= parseFloat(cs.opacity);
      node = node.parentElement;
    }
    const cs = getComputedStyle(el);
    const m = new DOMMatrix(cs.transform === 'none' ? undefined : cs.transform);
    const toHex = (s) => {
      const mm = /rgba?\(([^)]+)\)/.exec(s || '');
      if (!mm) return null;
      return '#' + mm[1].split(',').slice(0, 3).map((v) => Math.round(parseFloat(v)).toString(16).padStart(2, '0')).join('').toUpperCase();
    };
    const path0 = el.querySelector && el.querySelector('.pp-geom path');
    const span0 = el.querySelector && el.querySelector('.pp-p span');
    return {
      visible, opacity,
      tx: m.e, ty: m.f,
      rot: (Math.atan2(m.b, m.a) * 180) / Math.PI,
      scale: Math.hypot(m.a, m.b),
      clip: cs.clipPath !== 'none' && cs.clipPath !== '',
      mask: (cs.maskImage && cs.maskImage !== 'none') || (cs.webkitMaskImage && cs.webkitMaskImage !== 'none'),
      fill: path0 ? toHex(getComputedStyle(path0).fill) : null,
      textColor: span0 ? toHex(getComputedStyle(span0).color) : null,
    };
  }, { c, slideIdx });
}

function judgeAnim(st, expect) {
  if (st.error) return { ok: false, msg: st.error };
  const bad = [];
  for (const [k, v] of Object.entries(expect)) {
    const got = st[k];
    if (Array.isArray(v)) { if (!(got >= v[0] && got <= v[1])) bad.push(`${k}=${typeof got === 'number' ? got.toFixed(3) : got} ∉ [${v}]`); }
    else if (typeof v === 'string') { if (String(got).toUpperCase() !== v.toUpperCase()) bad.push(`${k}=${got} ≠ ${v}`); }
    else if (got !== v) bad.push(`${k}=${got} ≠ ${v}`);
  }
  return { ok: !bad.length, msg: bad.join('; ') || JSON.stringify(Object.fromEntries(Object.keys(expect).map((k) => [k, typeof st[k] === 'number' ? Math.round(st[k] * 1000) / 1000 : st[k]]))) };
}

// ---- main --------------------------------------------------------------------------
const LIST_CSS = '.pp-list .pp-slide{margin:0;box-shadow:none} body.pp-body{background:#fff}';
const results = [];
const t0 = Date.now();
const browser = await chromium.launch();
const decks = buildFixtures(path.join(outDir, 'fixtures'));
let failed = 0, passed = 0;

for (const deck of decks) {
  const features = deck.features.filter((f) => !only || f.feature.includes(only));
  if (!features.length) continue;
  const bytes = fs.readFileSync(deck.file);
  const pres = loadPresentation(bytes);
  pres.__name = deck.name;
  const W = Math.round(pres.width), H = Math.round(pres.height);
  const listHtml = path.join(outDir, deck.name + '-list.html');
  fs.writeFileSync(listHtml, renderDocument(pres, { mode: 'list', player: PLAYER_SOURCE, extraCss: LIST_CSS }));
  const deckHtml = path.join(outDir, deck.name + '.html');
  fs.writeFileSync(deckHtml, renderDocument(pres, { mode: 'deck', player: PLAYER_SOURCE }));
  const page = await browser.newPage({ viewport: { width: W, height: H } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  const shots = await screenshotSlides(page, listHtml, pres);
  let loShots = null;
  if (useLO) loShots = libreOfficeRender(deck.file, W, H, path.join(outDir, 'lo', deck.name));
  // Deck page for animation checks.
  const dpage = await browser.newPage({ viewport: { width: W, height: H } });
  dpage.on('pageerror', (e) => errors.push('deck: ' + e.message));
  await dpage.goto('file://' + deckHtml + '?notransitions');
  await dpage.waitForFunction(() => !!window.pptxPlayer);
  for (const f of features) {
    const idx = f.slide - 1;
    const slideModel = pres.slides[idx];
    const res = { deck: deck.name, slide: f.slide, feature: f.feature, checks: [], image: path.relative(outDir, shots[idx]) };
    for (const c of f.checks) {
      let r;
      try {
        switch (c.kind) {
          case 'model': {
            const v = c.fn(slideModel, pres);
            r = { ok: v === true, msg: v === true ? c.desc : `${c.desc}: ${v}` };
            break;
          }
          case 'pixel': {
            const img = readPng(shots[idx]);
            const got = samplePx(img, c.x, c.y);
            const exp = hex2rgb(c.color);
            const d = Math.max(...got.map((v, i) => Math.abs(v - exp[i])));
            r = { ok: d <= (c.tol || 20), msg: `(${c.x},${c.y}) ${rgb2hex(got)} vs ${c.color} Δ${Math.round(d)}` };
            break;
          }
          case 'lumBelow': {
            const got = samplePx(readPng(shots[idx]), c.x, c.y);
            const lum = 0.2126 * got[0] + 0.7152 * got[1] + 0.0722 * got[2];
            r = { ok: lum <= c.max, msg: `(${c.x},${c.y}) luminance ${Math.round(lum)} ≤ ${c.max}` };
            break;
          }
          case 'text':
          case 'dom':
          case 'eval':
            r = await domCheck(page, f.slide, c);
            break;
          case 'anim': {
            const st = await animState(dpage, c, idx);
            r = judgeAnim(st, c.expect);
            r.msg = `step ${c.step}${c.t !== undefined ? ' t=' + c.t : ''} #${c.spid}${c.para !== undefined ? ' p' + c.para : ''}${c.letter !== undefined ? ' l' + c.letter : ''}: ` + r.msg;
            break;
          }
          case 'trigger': {
            const st = await dpage.evaluate(async ({ c, idx }) => {
              const p = window.pptxPlayer;
              p.seek(idx, 0);
              const sec = document.querySelector('section#slide-' + (idx + 1));
              const btn = sec.querySelector(`.pp-el[data-spid="${c.spid}"]`);
              const r = btn.getBoundingClientRect();
              btn.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: r.left + 5, clientY: r.top + 5, button: 0 }));
              await new Promise((res) => setTimeout(res, 800));
              const t = sec.querySelector(`.pp-el[data-spid="${c.target}"]`);
              return { visible: getComputedStyle(t).visibility !== 'hidden', opacity: parseFloat(getComputedStyle(t).opacity) };
            }, { c, idx });
            r = judgeAnim(st, c.expect);
            break;
          }
          case 'transition': {
            const pad = (n) => path.join(outDir, 'img', `${deck.name}-${f.slide}-tr-${n}.png`);
            const frames = [];
            for (const p of [0, 0.5, 1]) {
              await dpage.evaluate(({ idx, p }) => {
                const pl = window.pptxPlayer;
                pl.goto(idx - 1, { instant: true, atEnd: true });
                if (p === 0) return;
                pl.previewTransition(idx, p);
              }, { idx, p });
              await dpage.screenshot({ path: pad(p * 100) });
              frames.push(readPng(pad(p * 100)));
            }
            const dA = diffFraction(frames[0], frames[1]);
            const dB = diffFraction(frames[1], frames[2]);
            const end = similarity(frames[2], readPng(shots[idx]));
            const midOk = !c.mid || (dA > 0.01 && dB > 0.01);
            r = { ok: midOk && end > 0.97, msg: `mid differs from start ${(dA * 100).toFixed(1)}% / end ${(dB * 100).toFixed(1)}%, final similarity ${(end * 100).toFixed(1)}%` };
            res.frames = [0, 50, 100].map((n) => path.relative(outDir, pad(n)));
            break;
          }
          default:
            r = { ok: false, msg: 'unknown check ' + c.kind };
        }
      } catch (e) {
        r = { ok: false, msg: 'exception: ' + e.message };
      }
      res.checks.push({ kind: c.kind, ...r });
      if (r.ok) passed++; else failed++;
    }
    if (loShots && loShots[idx]) {
      res.lo = path.relative(outDir, loShots[idx]);
      try { res.similarity = similarity(readPng(shots[idx]), readPng(loShots[idx])); } catch (e) { res.similarity = null; }
    }
    results.push(res);
  }
  const diag = pres.diagnostics.list().filter((d) => d.kind === 'error' || d.kind === 'unsupported');
  results.push({ deck: deck.name, feature: `${deck.name}/diagnostics`, checks: [
    { kind: 'diagnostics', ok: diag.length === 0, msg: diag.length ? diag.map((d) => `${d.kind}: ${d.message}`).join('; ') : 'no errors or unsupported elements' },
    { kind: 'console', ok: errors.length === 0, msg: errors.length ? errors.slice(0, 5).join(' | ') : 'no page errors' },
  ] });
  if (diag.length) failed++; else passed++;
  if (errors.length) failed++; else passed++;
  await page.close();
  await dpage.close();
}

// ---- optional corpus comparison ------------------------------------------------------
const corpus = [];
for (const dir of corpusDirs) {
  const files = fs.readdirSync(dir).filter((f) => /\.pptx$/i.test(f)).slice(0, Number(opt('--corpus-limit', 400)));
  for (const f of files) {
    const file = path.join(dir, f);
    let pres;
    try { pres = loadPresentation(fs.readFileSync(file)); } catch (e) { corpus.push({ file: f, error: e.message }); continue; }
    if (!pres.slides.length) continue;
    pres.__name = 'corpus-' + f.replace(/\W+/g, '_');
    const W = Math.round(pres.width), H = Math.round(pres.height);
    if (W * H > 4000 * 4000) continue;
    const html = path.join(outDir, pres.__name + '.html');
    fs.writeFileSync(html, renderDocument(pres, { mode: 'list', player: PLAYER_SOURCE, extraCss: LIST_CSS }));
    const page = await browser.newPage({ viewport: { width: Math.min(W, 4000), height: Math.min(H, 4000) } });
    let shots = [];
    try { shots = await screenshotSlides(page, html, pres); } catch (e) { corpus.push({ file: f, error: 'render: ' + e.message }); await page.close(); continue; }
    await page.close();
    const lo = useLO ? libreOfficeRender(file, W, H, path.join(outDir, 'lo', pres.__name)) : null;
    if (argv.includes('--progress')) console.log(`corpus ${corpus.length} ${f}`);
    pres.slides.forEach((s, i) => {
      const entry = { file: f, slide: i + 1, image: path.relative(outDir, shots[i]) };
      if (lo && lo[i]) {
        entry.lo = path.relative(outDir, lo[i]);
        try { entry.similarity = similarity(readPng(shots[i]), readPng(lo[i])); } catch (e) { entry.similarity = null; }
      }
      corpus.push(entry);
    });
  }
}
await browser.close();

// ---- report ------------------------------------------------------------------------
const summary = {
  date: new Date().toISOString(),
  durationMs: Date.now() - t0,
  passed, failed,
  features: results.length,
  libreOffice: useLO,
  meanSimilarity: avg(results.map((r) => r.similarity).filter((v) => typeof v === 'number')),
  corpusSlides: corpus.length,
  corpusMeanSimilarity: avg(corpus.map((r) => r.similarity).filter((v) => typeof v === 'number')),
};
function avg(a) { return a.length ? a.reduce((x, y) => x + y, 0) / a.length : null; }
fs.writeFileSync(path.join(outDir, 'summary.json'), JSON.stringify({ summary, results, corpus }, null, 2));
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const pct = (v) => (typeof v === 'number' ? (v * 100).toFixed(1) + '%' : '–');
let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>PPTX2HTML self-verification</title><style>
body{font:14px system-ui,sans-serif;margin:24px;color:#222;background:#fafafa}h1{margin:0 0 8px}table{border-collapse:collapse;width:100%;margin:16px 0;background:#fff}
td,th{border:1px solid #ddd;padding:6px 8px;vertical-align:top;text-align:left}th{background:#f0f0f0}.ok{color:#17803d}.bad{color:#c62828;font-weight:600}
img{max-width:320px;border:1px solid #ccc;display:block}.imgs{display:flex;gap:8px;flex-wrap:wrap}.sum span{display:inline-block;margin-right:18px;font-size:16px}
ul{margin:0;padding-left:18px}</style></head><body><h1>PPTX2HTML self-verification</h1>
<div class="sum"><span>Checks passed: <b class="ok">${passed}</b></span><span>failed: <b class="${failed ? 'bad' : 'ok'}">${failed}</b></span><span>Features: ${results.length}</span><span>LibreOffice similarity (mean): ${pct(summary.meanSimilarity)}</span><span>${summary.date}</span></div>
<table><tr><th>Feature</th><th>Checks</th><th>PPTX2HTML</th><th>LibreOffice</th></tr>`;
for (const r of results) {
  const ok = r.checks.every((c) => c.ok);
  html += `<tr><td><b class="${ok ? 'ok' : 'bad'}">${ok ? '✔' : '✘'}</b> ${esc(r.feature)}${r.slide ? `<br><small>${esc(r.deck)} slide ${r.slide}</small>` : ''}${typeof r.similarity === 'number' ? `<br><small>similarity ${pct(r.similarity)}</small>` : ''}</td><td><ul>${r.checks.map((c) => `<li class="${c.ok ? 'ok' : 'bad'}">${esc(c.kind)}: ${esc(c.msg || '')}</li>`).join('')}</ul></td>`;
  html += `<td><div class="imgs">${r.image ? `<img src="${r.image}">` : ''}${(r.frames || []).map((f) => `<img src="${f}" style="max-width:150px">`).join('')}</div></td><td>${r.lo ? `<img src="${r.lo}">` : ''}</td></tr>`;
}
html += '</table>';
if (corpus.length) {
  html += `<h2>Corpus (${corpus.length} slides, mean similarity ${pct(summary.corpusMeanSimilarity)})</h2><table><tr><th>File</th><th>Similarity</th><th>PPTX2HTML</th><th>LibreOffice</th></tr>`;
  for (const c of corpus.slice().sort((a, b) => (a.similarity ?? 1) - (b.similarity ?? 1))) {
    html += `<tr><td>${esc(c.file)} ${c.slide ? '#' + c.slide : ''}${c.error ? `<br><span class="bad">${esc(c.error)}</span>` : ''}</td><td>${pct(c.similarity)}</td><td>${c.image ? `<img src="${c.image}">` : ''}</td><td>${c.lo ? `<img src="${c.lo}">` : ''}</td></tr>`;
  }
  html += '</table>';
}
html += '</body></html>';
fs.writeFileSync(path.join(outDir, 'index.html'), html);
console.log(`\nSelf-verification: ${passed} checks passed, ${failed} failed (${results.length} features) in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
if (summary.meanSimilarity !== null) console.log(`Mean visual similarity to LibreOffice: ${pct(summary.meanSimilarity)}`);
if (corpus.length) console.log(`Corpus: ${corpus.length} slides, mean similarity ${pct(summary.corpusMeanSimilarity)}`);
for (const r of results) for (const c of r.checks) if (!c.ok) console.log(`  ✘ ${r.feature} [${c.kind}] ${c.msg}`);
console.log('Report: ' + path.join(outDir, 'index.html'));
process.exit(failed ? 1 : 0);

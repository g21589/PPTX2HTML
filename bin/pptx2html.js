#!/usr/bin/env node
/**
 * Command line interface.
 *
 *   pptx2html deck.pptx                 -> deck.html (presentation player)
 *   pptx2html deck.pptx -o out.html --static
 *   pptx2html a.pptx b.pptx --out-dir site --assets
 */
import fs from 'node:fs';
import path from 'node:path';
import { convert, version } from '../src/index.js';
import { emfToSvg } from '../src/util/emf.js';

const HELP = `pptx2html ${version} - convert PowerPoint (.pptx) presentations to HTML

Usage: pptx2html [options] <input.pptx...>

Options:
  -o, --output <file>   output HTML file (single input only; "-" for stdout)
  -d, --out-dir <dir>   directory for the generated HTML files
  -s, --slides <list>   only convert these slides, e.g. "1,3-5" (1-based)
      --static          static page listing every slide (no player script)
      --mode <mode>     "deck" (presentation player, default) or "list"
      --title <text>    document title (default: presentation title)
      --assets [dir]    write media to files instead of embedding data URIs
                        (default dir: <output name>_files)
      --lang <code>     value of the <html lang> attribute
      --strict          fail on the first slide that cannot be rendered
  -q, --quiet           do not print diagnostics
  -h, --help            show this help
  -v, --version         show the version
`;

function parseArgs(argv) {
  const o = { inputs: [], quiet: false };
  const value = (i, name) => {
    if (i + 1 >= argv.length) fail(`missing value for ${name}`);
    return argv[i + 1];
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    switch (a) {
      case '-h': case '--help': o.help = true; break;
      case '-v': case '--version': o.version = true; break;
      case '-o': case '--output': o.output = value(i, a); i++; break;
      case '-d': case '--out-dir': o.outDir = value(i, a); i++; break;
      case '-s': case '--slides': o.slides = parseSlides(value(i, a)); i++; break;
      case '--static': o.static = true; break;
      case '--mode': o.mode = value(i, a); i++; if (!['deck', 'list'].includes(o.mode)) fail('--mode must be "deck" or "list"'); break;
      case '--title': o.title = value(i, a); i++; break;
      case '--lang': o.lang = value(i, a); i++; break;
      case '--strict': o.strict = true; break;
      case '-q': case '--quiet': o.quiet = true; break;
      case '--assets':
        o.assets = true;
        if (argv[i + 1] && !argv[i + 1].startsWith('-') && !/\.pptx$/i.test(argv[i + 1])) o.assetsDir = argv[++i];
        break;
      default:
        if (a.startsWith('-') && a !== '-') fail(`unknown option ${a}`);
        o.inputs.push(a);
    }
  }
  return o;
}

function parseSlides(spec) {
  const out = new Set();
  for (const part of spec.split(',')) {
    const m = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part);
    if (!m) fail(`invalid slide list "${spec}"`);
    const a = +m[1], b = m[2] ? +m[2] : a;
    for (let n = Math.min(a, b); n <= Math.max(a, b); n++) if (n >= 1) out.add(n - 1);
  }
  return [...out].sort((x, y) => x - y);
}

function fail(msg) {
  process.stderr.write(`pptx2html: ${msg}\n`);
  process.exit(2);
}

/** mediaUrl option writing every media part into `dir`, referenced relative to `base`. */
function assetWriter(dir, base) {
  const used = new Set();
  return (media, store) => {
    let bytes = store.bytes(media.id);
    if (!bytes) return '';
    let name = path.posix.basename(media.partName);
    if (media.mime === 'image/x-emf' || media.mime === 'image/x-wmf') {
      const svg = emfToSvg(bytes, media.mime);
      if (svg) { bytes = Buffer.from(svg); name = name.replace(/\.[^.]*$/, '') + '.svg'; }
    }
    // Media from different folders may share a file name.
    let unique = name, k = 1;
    while (used.has(unique)) unique = name.replace(/(\.[^.]*)?$/, `-${k++}$1`);
    used.add(unique);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, unique), bytes);
    return path.relative(base, path.join(dir, unique)).split(path.sep).map(encodeURIComponent).join('/');
  };
}

function main() {
  const o = parseArgs(process.argv.slice(2));
  if (o.help) { process.stdout.write(HELP); return; }
  if (o.version) { process.stdout.write(version + '\n'); return; }
  if (!o.inputs.length) { process.stderr.write(HELP); process.exit(2); }
  if (o.output && o.inputs.length > 1) fail('--output needs a single input; use --out-dir');
  if (o.output === '-' && o.assets) fail('--assets cannot be used when writing to stdout');
  let failed = 0;
  for (const input of o.inputs) {
    let data;
    try {
      data = fs.readFileSync(input);
    } catch (e) {
      process.stderr.write(`pptx2html: cannot read ${input}: ${e.message}\n`);
      failed++;
      continue;
    }
    const stem = path.basename(input).replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, '');
    const output = o.output || path.join(o.outDir || path.dirname(input), stem + '.html');
    const options = { title: o.title, lang: o.lang, strict: o.strict, slides: o.slides };
    if (o.static) { options.player = false; options.mode = 'list'; }
    if (o.mode) options.mode = o.mode;
    if (o.assets) {
      const base = path.dirname(path.resolve(output));
      const dir = path.resolve(o.assetsDir || path.join(base, path.basename(output).replace(/\.html?$/i, '') + '_files'));
      options.mediaUrl = assetWriter(dir, base);
    }
    let result;
    try {
      result = convert(data, options);
    } catch (e) {
      process.stderr.write(`pptx2html: ${input}: ${e.message}\n`);
      failed++;
      continue;
    }
    if (output === '-') process.stdout.write(result.html);
    else {
      fs.mkdirSync(path.dirname(path.resolve(output)), { recursive: true });
      fs.writeFileSync(output, result.html);
    }
    if (!o.quiet) {
      const n = result.presentation.slides.length;
      const msg = output === '-' ? '' : `${input} -> ${output} (${n} slide${n === 1 ? '' : 's'})\n`;
      process.stderr.write(msg);
      for (const d of result.diagnostics) process.stderr.write(`  ${d.kind}: ${d.message}${d.count > 1 ? ` (x${d.count})` : ''}\n`);
    }
  }
  if (failed) process.exit(1);
}

main();

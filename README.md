PPTX2HTML
=========
[![MIT License][license-image]][license-url]

Convert PowerPoint presentations (`.pptx`, `.pptm`, `.ppsx`, `.potx`) to a
single, self-contained HTML file that looks like the original slide show and
plays its animations and slide transitions.

* Pure JavaScript, **zero runtime dependencies**. Runs in the browser and in
  Node.js 18+.
* Renders every element type PowerPoint writes: shapes (all 187 preset
  geometries and custom geometry), text, pictures, tables, charts, SmartArt,
  media, groups, OLE previews, metafiles, backgrounds, masters and layouts.
* Plays the **PresentationML timing tree** (entrance, emphasis, exit, motion
  path, sequencing, text builds, triggers, sounds, media commands) and **all
  slide transitions**, including the PowerPoint 2010+ effects and Morph.
* **Self-verifying**: a generated catalogue of fixture decks exercises every
  supported feature; each one is converted, screenshotted in Chromium and
  checked (model, DOM, pixels, animation state, transitions), and compared
  against LibreOffice's rendering. See [Self-verification](#self-verification).

Try the [online demo] — the file never leaves your browser.

繁體中文簡介：PPTX2HTML 以純 JavaScript 將 PowerPoint 簡報轉換成單一 HTML 檔，
支援所有圖形、文字、表格、圖表、SmartArt、多媒體、動畫與轉場效果，並內建自動
驗證流程（產生涵蓋所有功能的測試簡報、在 Chromium 截圖檢查並與 LibreOffice
比對）。

Quick start
-----------

### Command line

```sh
npm install && npm link                    # or run node bin/pptx2html.js directly
pptx2html deck.pptx                        # -> deck.html (slide show player)
pptx2html deck.pptx -o out.html --static
pptx2html *.pptx --out-dir site --assets   # media written next to the HTML
```

| Option | Meaning |
| --- | --- |
| `-o, --output <file>` | Output file (`-` for stdout). Default: input name with `.html`. |
| `-d, --out-dir <dir>` | Directory for the generated files (several inputs). |
| `-s, --slides <list>` | Only convert some slides, e.g. `1,3-5`. |
| `--static` | Plain page listing every slide, without the player script. |
| `--mode deck\|list` | Slide show (default) or all slides stacked, both with the player. |
| `--assets [dir]` | Write pictures and media to files instead of data URIs. |
| `--title`, `--lang` | Document title and `<html lang>`. |
| `--strict` | Stop at the first slide that fails instead of rendering a placeholder. |
| `-q, --quiet` | Do not print diagnostics. |

### Node.js

```js
import fs from 'node:fs';
import { convert } from 'pptx2html';

const { html, presentation, diagnostics } = convert(fs.readFileSync('deck.pptx'));
fs.writeFileSync('deck.html', html);
console.log(presentation.slides.length, 'slides', diagnostics);
```

### Browser

```html
<script type="module">
  import { convert } from './src/index.js'; // or dist/pptx2html.js
  const file = document.querySelector('input[type=file]').files[0];
  const { html } = convert(new Uint8Array(await file.arrayBuffer()));
  document.querySelector('iframe').src = URL.createObjectURL(new Blob([html], { type: 'text/html' }));
</script>
```

`npm run build` also produces `dist/pptx2html.min.js`, an IIFE bundle that
exposes a global `PPTX2HTML`.

API
---

### `convert(data, options?) → { html, presentation, diagnostics }`

`data` is an `ArrayBuffer` or `Uint8Array` with the package bytes.

| Option | Default | |
| --- | --- | --- |
| `mode` | `'deck'` | `'deck'` (slide show) or `'list'` (all slides). |
| `player` | `true` | Embed the runtime player. `false` gives a static page. |
| `slides` | all | 0-based indices of the slides to build. |
| `title`, `lang` | from the deck | Document title and language. |
| `mediaUrl(media, store)` | data URIs | Return a URL for a picture or media part (`store.bytes(media.id)` gives its bytes). |
| `convertMetafiles` | `true` | Convert EMF/WMF pictures to SVG. |
| `liveDates` | `false` | Render date fields with the current date. |
| `strict` | `false` | Throw on the first slide that cannot be built. |
| `idPrefix` | `'pp'` | Prefix of generated element ids. |

`diagnostics` lists everything that was approximated or skipped, e.g.
`{ kind: 'approx', message: '3D chart rendered in 2D', count: 1, parts: [...] }`.

Lower-level building blocks are exported as well:

* `loadPresentation(data, options)` → presentation model (slides, elements,
  timing, transitions, notes, theme data, media registry).
* `renderSlides(presentation, options)` → `{ slides: string[], css, data }`
  to embed slides into your own page.
* `renderDocument(presentation, options)` → complete HTML document.
* `playerData(presentation)` → the JSON consumed by the runtime player.
* `BASE_CSS`, `version`.

The player can be used on its own (`pptx2html/player`, or
`dist/pptx2html-player.js`, which defines `window.PPTXPlayer`).

### Slide show

Keys: <kbd>→</kbd> <kbd>Space</kbd> <kbd>Enter</kbd> next,
<kbd>←</kbd> <kbd>Backspace</kbd> previous, <kbd>Home</kbd>/<kbd>End</kbd>,
<kbd>F</kbd> fullscreen, <kbd>S</kbd> speaker notes, <kbd>B</kbd>/<kbd>W</kbd>
black/white screen. Click, mouse wheel and swipe gestures also navigate.

URL parameters: `#slide-3` opens a slide, `?mode=list` shows every slide,
`?hidden` includes hidden slides, `?notransitions` disables transitions.

The running player is available as `window.pptxPlayer` with `next()`,
`prev()`, `goto(index)`, `state()`, and the deterministic testing helpers
`seek(slide, step, ms)` and `previewTransition(slide, progress)`.

What is rendered
----------------

**Package**: ZIP (stored/deflate, data descriptors, damaged central
directories), relationships, content types, `mc:AlternateContent`, strict
OOXML namespaces, slide sections and custom shows, hidden slides, notes.

**Theme and inheritance**: theme colours, colour maps and overrides, all
colour transforms (tint, shade, lumMod/Off, satMod, hue, alpha, gray,
inverse, gamma…), the placeholder chain slide → layout → master, text style
inheritance (`defaultTextStyle`, master text styles, list styles), style
matrix references (`fillRef`, `lnRef`, `effectRef`, `fontRef`).

**Shapes**: the 187 preset geometries computed from their guide formulas,
custom geometry paths, connectors, solid / gradient (linear, radial,
rectangular, path) / pattern (all 48 presets) / picture fills with tiling and
stretching, group fills, line dashes, caps, joins, compound lines and arrow
heads, shadows, glows, reflections, soft edges, blur, 3D rotation, bevel and
extrusion approximations, rotation and flipping.

**Text**: fonts (Latin / East Asian / complex script, theme fonts, embedded
fonts), sizes, bold, italic, underline styles, strike, caps, baseline,
spacing, kerning, highlight, outline, gradient text, shadows, bullets
(characters, pictures, auto-numbering), indents, line and paragraph spacing,
alignment, vertical text, columns, insets, autofit (font scale and line
spacing reduction), fields (slide number, dates), hyperlinks and actions,
symbol fonts mapped to Unicode.

**Objects**: pictures (crop, duotone, grayscale, brightness/contrast, colour
change, alpha, SVG blips, EMF/WMF converted to SVG), tables (table styles,
banding, merged cells, cell borders and fills), charts (bar, column, line,
area, pie, doughnut, scatter, bubble, radar, stock, combo, secondary axes,
trend lines, data labels, number formats; 3D charts drawn in 2D),
SmartArt (pre-rendered drawing), audio and video with bookmarks and
trimming, OLE objects (preview picture), slide backgrounds.

**Animations**: entrance / emphasis / exit / motion path effects with their
filters (wipe, split, blinds, box, checkerboard, circle, diamond, dissolve,
random bars, strips, wedge, wheel, plus…), `anim`/`set`/`animEffect`/
`animMotion`/`animRot`/`animScale`/`animClr`/`cmd`/media behaviours, formula
evaluation (`#ppt_x`, `$`, `sin`…), repeat, auto-reverse, acceleration,
by-paragraph and by-letter builds, after-previous / with-previous timing,
interactive triggers, after-effects (dim / hide), sounds.

**Transitions**: all PowerPoint 2007 transitions, the 2010+ "exciting" set
(vortex, switch, flip, ripple, honeycomb, prism, doors, window, ferris,
gallery, conveyor, pan, glitter, warp, flythrough, flash, shred, reveal,
curtains, wind, fracture, page curl, airplane, origami/fall over…), Morph,
speed / duration, sounds, and advance on click or after time.

Limitations: 3D models, ink without a fallback picture, and EMF+-only
metafiles are not drawn; 3D effects and 3D charts are approximations;
compressed EOT fonts fall back to installed fonts. Each case is reported in
`diagnostics`.

Self-verification
-----------------

```sh
npm install
npm test               # unit tests (parser, geometry, colours, charts, engine, metafiles…)
npm run verify:quick   # feature catalogue in Chromium
npm run verify         # … plus LibreOffice reference renders and similarity scores
```

`tools/verify.mjs`:

1. **Builds fixture decks** from raw OOXML (`test/fixtures/catalog/*.mjs`)
   covering drawing, text, objects and animation features — one feature per
   slide, each with its own list of checks.
2. **Converts** them and runs the checks: model checks in Node, DOM / text /
   pixel / luminance checks on Chromium screenshots, animation checks that seek
   the player to given steps and times, trigger checks, and transition checks
   frozen at given progress.
3. **Renders the decks with LibreOffice** (when `soffice` and `pdftoppm` are
   installed) and scores the visual similarity of every slide.
4. Writes `verify-report/index.html` (side-by-side screenshots, check
   results) and `verify-report/summary.json`; exits with 1 when a check fails.

Options: `--no-lo`, `--only <feature>`, `--out <dir>`, and
`--corpus <dir>` (repeatable) to also convert and compare a folder of real
presentations. The generated decks are also in [`samples/features`](samples/features)
(`npm run samples` regenerates them).

Project layout
--------------

```
src/util       XML parser, ZIP + inflate, base64, image sizes, EMF/WMF → SVG
src/opc        Open Packaging Conventions (parts, relationships, content types)
src/geometry   preset geometry definitions and the guide formula evaluator
src/pptx       presentation model: theme, colours, fills, text, slides, tables, timing, transitions
src/chart      chart parser, number formats and SVG renderer
src/render     HTML/SVG renderers and document assembly
src/runtime    animation engine, effect filters, transitions and the player
bin            command line interface
tools          build, verification, snapshot and corpus tools
test           unit tests and the fixture catalogue
samples        sample decks used by the demo page
```

`npm run build` regenerates `src/generated/player-source.js` (the minified
player embedded in every document) and the `dist/` bundles; run it after
changing `src/runtime`. CI checks that the generated file is up to date.

License
-------

MIT

[license-image]: http://img.shields.io/badge/license-MIT-blue.svg?style=flat
[license-url]: LICENSE
[online demo]: http://g21589.github.io/PPTX2HTML

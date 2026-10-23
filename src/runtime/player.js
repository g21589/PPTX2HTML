/**
 * PPTX2HTML runtime player.
 *
 * Turns the rendered slides into a presentation: slide transitions,
 * animations (main sequence, triggers), media, hyperlinks, auto advance,
 * keyboard / mouse / touch navigation. Exposes a small API on
 * `window.pptxPlayer` for embedding and automated testing.
 */
import { evaluate, stepEnd } from './engine.js';
import { filterStyle, transitionFrame } from './effects.js';

const PLAYER_CSS = `
html.pp-deck-mode,html.pp-deck-mode body{margin:0;height:100%;overflow:hidden;background:#000}
.pp-deck.pp-deck-active{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:#000;overflow:hidden;user-select:none;-webkit-user-select:none}
.pp-deck-active .pp-stage{position:relative;flex:none;transform-origin:50% 50%}
.pp-deck-active .pp-slide{position:absolute;left:0;top:0;display:none;margin:0;box-shadow:none}
.pp-deck-active .pp-slide.pp-current,.pp-deck-active .pp-slide.pp-prev{display:block}
.pp-deck-active .pp-slide.pp-prev{z-index:1}
.pp-deck-active .pp-slide.pp-current{z-index:2}
.pp-deck-active .pp-slide.pp-prev.pp-on-top{z-index:3}
.pp-stage-back{position:absolute;inset:0;z-index:0}
.pp-ui{position:fixed;right:12px;bottom:12px;display:flex;gap:6px;align-items:center;opacity:0;transition:opacity .3s;z-index:10;font:13px system-ui,sans-serif;color:#fff}
.pp-ui.pp-show{opacity:1}
.pp-ui button{background:rgba(40,40,40,.75);color:#fff;border:0;border-radius:4px;min-width:32px;height:30px;cursor:pointer;font-size:14px}
.pp-ui button:hover{background:rgba(80,80,80,.9)}
.pp-ui .pp-count{background:rgba(40,40,40,.75);padding:6px 10px;border-radius:4px;min-width:64px;text-align:center}
.pp-blank{position:fixed;inset:0;z-index:20;display:none}
.pp-notes{position:fixed;left:0;right:0;bottom:0;max-height:30%;overflow:auto;background:rgba(0,0,0,.8);color:#fff;font:15px/1.4 system-ui,sans-serif;padding:12px 18px;white-space:pre-wrap;z-index:15;display:none}
.pp-progress{position:fixed;left:0;bottom:0;height:3px;background:#3b82f6;z-index:11;transition:width .3s}
.pp-it{display:inline-block;white-space:pre}
`;

function css(c) {
  if (!c) return '';
  const a = c.a === undefined ? 1 : c.a;
  return `rgba(${Math.round(c.r)},${Math.round(c.g)},${Math.round(c.b)},${Math.round(a * 1000) / 1000})`;
}

function parseColor(s) {
  if (!s) return null;
  let m = /^#([0-9a-f]{6})$/i.exec(s.trim());
  if (m) {
    const n = parseInt(m[1], 16);
    return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
  }
  m = /rgba?\(([^)]+)\)/.exec(s);
  if (m) {
    const p = m[1].split(',').map((x) => parseFloat(x));
    return { r: p[0], g: p[1], b: p[2], a: p[3] === undefined ? 1 : p[3] };
  }
  return null;
}

/** Per-slide animation targets. */
class SlideTargets {
  constructor(section, timing, W, H) {
    this.section = section;
    this.timing = timing;
    this.W = W;
    this.H = H;
    this.map = new Map();
    this.originals = new Map();
    this.counts = new Map();
    if (timing) this.collect();
  }

  keys() {
    const t = this.timing;
    const set = new Set(t.hidden || []);
    const add = (steps) => { for (const st of steps) for (const b of st.behaviours) for (const k of b.keys) set.add(k); };
    add(t.main || []);
    for (const tr of t.triggers || []) add(tr.steps);
    return set;
  }

  shapeEl(spid) {
    const esc = String(spid).replace(/["\\]/g, '');
    return this.section.querySelector(`:scope > .pp-el[data-layer="slide"][data-spid="${esc}"]`) || this.section.querySelector(`.pp-el[data-layer="slide"][data-spid="${esc}"]`);
  }

  collect() {
    const iterKeys = new Map();
    const scan = (steps) => {
      for (const st of steps) for (const b of st.behaviours) if (b.iterate) for (const k of b.keys) iterKeys.set(k, b.iterate.type);
    };
    scan(this.timing.main || []);
    for (const tr of this.timing.triggers || []) scan(tr.steps);
    for (const key of this.keys()) {
      const [spid, sub] = key.split(':');
      const el = this.shapeEl(spid);
      if (!el) continue;
      let els = [el];
      if (sub === 'bg') {
        const g = el.querySelector(':scope > .pp-geom');
        els = g ? [g] : [];
      } else if (sub && sub[0] === 'p') {
        const n = sub.slice(1);
        els = [...el.querySelectorAll(`.pp-p[data-para="${n}"]`)];
      }
      if (!els.length) continue;
      this.map.set(key, { els, kind: sub === 'bg' ? 'bg' : sub ? 'para' : 'shape', shape: el });
      for (const e of els) if (!this.originals.has(e)) this.originals.set(e, e.getAttribute('style') || '');
      if (iterKeys.has(key)) this.split(key, els, iterKeys.get(key));
    }
  }

  /** Split text into words / letters for "animate text by word / letter". */
  split(key, els, type) {
    const parts = [];
    for (const root of els) {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (walker.nextNode()) nodes.push(walker.currentNode);
      for (const tn of nodes) {
        if (!tn.parentElement || tn.parentElement.closest('.pp-bu') || !tn.nodeValue) continue;
        const text = tn.nodeValue;
        const pieces = type === 'lt' ? [...text] : text.split(/(\s+)/).filter((x) => x.length);
        const frag = document.createDocumentFragment();
        for (const piece of pieces) {
          if (/^\s+$/.test(piece) && type !== 'lt') { frag.appendChild(document.createTextNode(piece)); continue; }
          const span = document.createElement('span');
          span.className = 'pp-it';
          span.textContent = piece;
          frag.appendChild(span);
          parts.push(span);
        }
        tn.parentNode.replaceChild(frag, tn);
      }
    }
    parts.forEach((p, i) => {
      const k = key + '#' + i;
      this.map.set(k, { els: [p], kind: 'iter' });
      this.originals.set(p, p.getAttribute('style') || '');
    });
    this.counts.set(key, parts.length);
  }

  /** Geometry (slide fractions) of a key, measured from layout. */
  geo(key) {
    const entry = this.map.get(key);
    if (!entry) return null;
    if (entry.geo) return entry.geo;
    const el = entry.els[0];
    let g;
    if (entry.kind === 'shape' || entry.kind === 'bg') {
      const s = entry.shape.style;
      const x = parseFloat(s.left) || 0, y = parseFloat(s.top) || 0, w = parseFloat(s.width) || 0, h = parseFloat(s.height) || 0;
      g = { x: (x + w / 2) / this.W, y: (y + h / 2) / this.H, w: w / this.W, h: h / this.H, rot: parseFloat(entry.shape.dataset.rot || '0') };
    } else {
      const sr = this.section.getBoundingClientRect();
      const scale = sr.width / this.W || 1;
      const r = el.getBoundingClientRect();
      g = { x: ((r.left + r.right) / 2 - sr.left) / scale / this.W, y: ((r.top + r.bottom) / 2 - sr.top) / scale / this.H, w: r.width / scale / this.W, h: r.height / scale / this.H, rot: 0 };
    }
    entry.geo = g;
    return g;
  }

  colors(key) {
    const entry = this.map.get(key);
    if (!entry) return null;
    const el = entry.els[0];
    const path = el.querySelector && el.querySelector('.pp-geom path[fill]:not([fill="none"])');
    const span = el.querySelector && el.querySelector('.pp-p span');
    return {
      fill: path ? parseColor(path.getAttribute('fill')) : null,
      stroke: (() => { const p = el.querySelector && el.querySelector('.pp-geom path[stroke]:not([stroke="none"])'); return p ? parseColor(p.getAttribute('stroke')) : null; })(),
      color: span ? parseColor(getComputedStyle(span).color) : null,
    };
  }

  apply(states) {
    for (const [key, entry] of this.map) this.applyOne(key, entry, states.get(key));
  }

  applyOne(key, entry, st) {
    for (const el of entry.els) {
      const orig = this.originals.get(el) || '';
      if (!st) {
        if ((el.getAttribute('style') || '') !== orig) el.setAttribute('style', orig);
        this.resetColors(el, entry);
        continue;
      }
      el.setAttribute('style', orig);
      const s = el.style;
      if (st.visible === false) s.visibility = 'hidden';
      let op = st.opacity;
      let extra = '';
      for (const ef of st.effects) {
        const f = filterStyle(ef.filter, ef.reveal);
        if (f.opacity !== undefined) op *= f.opacity;
        if (f.clipPath) { s.clipPath = f.clipPath; s.webkitClipPath = f.clipPath; }
        if (f.maskImage) {
          for (const k of ['maskImage', 'webkitMaskImage', 'maskSize', 'webkitMaskSize', 'maskRepeat', 'webkitMaskRepeat', 'maskComposite', 'webkitMaskComposite']) if (f[k]) s[k] = f[k];
        }
        if (f.transform) extra += ' ' + f.transform;
      }
      if (op < 1) s.opacity = String(Math.max(0, op));
      // Transform.
      const g = this.geo(key) || { x: 0, y: 0, w: 1, h: 1, rot: 0 };
      const tx = ((st.x !== undefined ? st.x - g.x : 0) + st.dx) * this.W;
      const ty = ((st.y !== undefined ? st.y - g.y : 0) + st.dy) * this.H;
      const isShape = entry.kind === 'shape';
      const baseRot = isShape ? parseFloat(entry.shape.dataset.rot || '0') : 0;
      const fh = isShape && entry.shape.dataset.fh === '1' ? -1 : 1;
      const fv = isShape && entry.shape.dataset.fv === '1' ? -1 : 1;
      const rot = baseRot + st.rot + (st.absRot !== undefined ? st.absRot : 0);
      let sx = st.sx * (st.w !== undefined && g.w ? st.w / g.w : 1);
      let sy = st.sy * (st.h !== undefined && g.h ? st.h / g.h : 1);
      if (!Number.isFinite(sx)) sx = 1;
      if (!Number.isFinite(sy)) sy = 1;
      const changed = tx || ty || rot !== baseRot || sx !== 1 || sy !== 1 || st.skewX || st.skewY || extra;
      if (changed) {
        let t = '';
        if (tx || ty) t += `translate(${tx.toFixed(2)}px,${ty.toFixed(2)}px) `;
        if (rot) t += `rotate(${rot.toFixed(3)}deg) `;
        if (sx * fh !== 1 || sy * fv !== 1) t += `scale(${(sx * fh).toFixed(4)},${(sy * fv).toFixed(4)}) `;
        if (st.skewX || st.skewY) t += `skew(${(st.skewX || 0) * 45}deg,${(st.skewY || 0) * 45}deg) `;
        s.transform = (t + extra).trim() || 'none';
      }
      this.applyColors(el, entry, st);
    }
  }

  applyColors(el, entry, st) {
    const geom = entry.kind === 'bg' ? [el] : [...el.querySelectorAll(':scope > .pp-geom')];
    for (const g of geom) {
      for (const p of g.querySelectorAll('path')) {
        if (st.fill && p.getAttribute('fill') && p.getAttribute('fill') !== 'none') p.style.fill = css(st.fill);
        else if (st.fillOn === false && p.getAttribute('fill') !== 'none') p.style.fill = 'none';
        else p.style.fill = '';
        if (st.stroke && p.getAttribute('stroke') && p.getAttribute('stroke') !== 'none') p.style.stroke = css(st.stroke);
        else if (st.strokeOn === false) p.style.stroke = 'none';
        else p.style.stroke = '';
      }
    }
    if (entry.kind !== 'bg') {
      for (const sp of el.querySelectorAll('.pp-p span')) {
        if (sp.classList.contains('pp-bu-gap')) continue;
        sp.style.color = st.color ? css(st.color) : '';
        sp.style.webkitTextFillColor = st.color ? css(st.color) : '';
        sp.style.fontWeight = st.bold === undefined ? '' : (st.bold ? 'bold' : 'normal');
        sp.style.fontStyle = st.italic === undefined ? '' : (st.italic ? 'italic' : 'normal');
        sp.style.textDecorationLine = st.underline ? 'underline' : '';
      }
      if (st.fontScale !== undefined && st.fontScale !== 1) {
        const tx = el.querySelector('.pp-tx') || el;
        tx.style.zoom = String(st.fontScale);
      }
    }
  }

  resetColors(el, entry) {
    if (!el.querySelectorAll) return;
    if (entry.kind !== 'iter') {
      for (const p of el.querySelectorAll('.pp-geom path')) { p.style.fill = ''; p.style.stroke = ''; }
      for (const sp of el.querySelectorAll('.pp-p span')) {
        if (sp.style.color || sp.style.fontWeight || sp.style.fontStyle || sp.style.textDecorationLine) {
          sp.style.color = ''; sp.style.webkitTextFillColor = ''; sp.style.fontWeight = ''; sp.style.fontStyle = ''; sp.style.textDecorationLine = '';
        }
      }
      const tx = el.querySelector('.pp-tx');
      if (tx && tx.style.zoom) tx.style.zoom = '';
    }
  }

  media(spid) {
    const el = this.shapeEl(spid);
    return el ? el.querySelector('video,audio') : null;
  }
}

export class Player {
  /**
   * @param {HTMLElement} deck .pp-deck element
   * @param {object} data player data (timing, transitions...)
   * @param {object} [opts] {mode: 'deck'|'list', start: slide number, showHidden}
   */
  constructor(deck, data, opts = {}) {
    this.deck = deck;
    this.data = data;
    this.opts = opts;
    this.W = data.width;
    this.H = data.height;
    this.sections = [...deck.querySelectorAll(':scope > .pp-slide, :scope > .pp-stage > .pp-slide')];
    this.cur = -1;
    this.step = 0;
    this.stepStart = null;
    this.triggers = [];
    this.targets = new Map();
    this.done = new Set();
    this.raf = 0;
    this.transition = null;
    this.advTimer = 0;
    this.listeners = [];
    this.sounds = [];
    this.history = [];
  }

  get slideData() {
    return this.data.slides[this.cur] || { timing: null, transition: null };
  }

  order() {
    return this.data.slides.map((s, i) => i).filter((i) => this.opts.showHidden || !this.data.slides[i].hidden);
  }

  init() {
    const style = document.createElement('style');
    style.textContent = PLAYER_CSS;
    document.head.appendChild(style);
    if (this.opts.mode === 'list') {
      this.initMedia(this.deck);
      this.initLinks(this.deck, true);
      return this;
    }
    document.documentElement.classList.add('pp-deck-mode');
    this.deck.classList.remove('pp-list');
    this.deck.classList.add('pp-deck-active');
    this.stage = document.createElement('div');
    this.stage.className = 'pp-stage';
    this.stage.style.width = this.W + 'px';
    this.stage.style.height = this.H + 'px';
    this.back = document.createElement('div');
    this.back.className = 'pp-stage-back';
    this.stage.appendChild(this.back);
    for (const s of this.sections) this.stage.appendChild(s);
    this.deck.appendChild(this.stage);
    this.buildUi();
    this.fit();
    this.on(window, 'resize', () => this.fit());
    this.on(document, 'keydown', (e) => this.onKey(e));
    this.on(this.deck, 'click', (e) => this.onClick(e));
    this.on(this.deck, 'contextmenu', (e) => { e.preventDefault(); this.prev(); });
    this.on(this.deck, 'wheel', (e) => { if (Math.abs(e.deltaY) > 20) { if (e.deltaY > 0) this.next(); else this.prev(); } }, { passive: true });
    let tx = null, ty = null;
    this.on(this.deck, 'touchstart', (e) => { tx = e.touches[0].clientX; ty = e.touches[0].clientY; }, { passive: true });
    this.on(this.deck, 'touchend', (e) => {
      if (tx === null) return;
      const dx = e.changedTouches[0].clientX - tx, dy = e.changedTouches[0].clientY - ty;
      if (Math.abs(dx) > 40 && Math.abs(dx) > Math.abs(dy)) { if (dx < 0) this.next(); else this.prev(); }
      tx = null;
    }, { passive: true });
    this.on(document, 'mousemove', () => this.showUi());
    this.on(window, 'hashchange', () => {
      const n = this.hashSlide();
      if (n !== null && n !== this.cur) this.goto(n, { instant: true });
    });
    this.initMedia(this.stage);
    this.initLinks(this.stage, false);
    let start = this.hashSlide();
    if (start === null) start = this.opts.start ? this.opts.start - 1 : this.order()[0] || 0;
    this.goto(start, { instant: true });
    return this;
  }

  on(target, evt, fn, o) {
    target.addEventListener(evt, fn, o);
    this.listeners.push([target, evt, fn, o]);
  }

  destroy() {
    for (const [t, e, f, o] of this.listeners) t.removeEventListener(e, f, o);
    cancelAnimationFrame(this.raf);
    clearTimeout(this.advTimer);
  }

  hashSlide() {
    const m = /#slide-(\d+)/.exec(location.hash || '');
    if (!m) return null;
    const n = Number(m[1]) - 1;
    return n >= 0 && n < this.sections.length ? n : null;
  }

  buildUi() {
    const ui = document.createElement('div');
    ui.className = 'pp-ui';
    ui.innerHTML = '<button type="button" data-a="prev" aria-label="Previous">&#9664;</button><span class="pp-count"></span><button type="button" data-a="next" aria-label="Next">&#9654;</button><button type="button" data-a="notes" aria-label="Notes">&#9776;</button><button type="button" data-a="fs" aria-label="Fullscreen">&#9974;</button>';
    ui.addEventListener('click', (e) => {
      e.stopPropagation();
      const a = e.target.closest('button');
      if (!a) return;
      const act = a.dataset.a;
      if (act === 'prev') this.prev();
      else if (act === 'next') this.next();
      else if (act === 'fs') this.toggleFullscreen();
      else if (act === 'notes') this.toggleNotes();
    });
    this.deck.appendChild(ui);
    this.ui = ui;
    this.count = ui.querySelector('.pp-count');
    this.blank = document.createElement('div');
    this.blank.className = 'pp-blank';
    this.deck.appendChild(this.blank);
    this.notesEl = document.createElement('div');
    this.notesEl.className = 'pp-notes';
    this.deck.appendChild(this.notesEl);
    this.progress = document.createElement('div');
    this.progress.className = 'pp-progress';
    this.deck.appendChild(this.progress);
  }

  showUi() {
    if (!this.ui) return;
    this.ui.classList.add('pp-show');
    clearTimeout(this.uiTimer);
    this.uiTimer = setTimeout(() => this.ui.classList.remove('pp-show'), 2500);
  }

  fit() {
    const vw = this.deck.clientWidth || window.innerWidth;
    const vh = this.deck.clientHeight || window.innerHeight;
    const s = Math.min(vw / this.W, vh / this.H);
    this.scale = s;
    this.stage.style.transform = `scale(${s})`;
    for (const t of this.targets.values()) for (const e of t.map.values()) if (e.kind !== 'shape' && e.kind !== 'bg') e.geo = null;
  }

  toggleFullscreen() {
    const d = document;
    if (d.fullscreenElement) d.exitFullscreen();
    else if (this.deck.requestFullscreen) this.deck.requestFullscreen().catch(() => {});
  }

  toggleNotes() {
    const n = this.notesEl;
    n.style.display = n.style.display === 'block' ? 'none' : 'block';
    this.updateUi();
  }

  updateUi() {
    if (!this.count) return;
    const order = this.order();
    const pos = order.indexOf(this.cur);
    this.count.textContent = `${pos + 1} / ${order.length}`;
    this.progress.style.width = ((pos + 1) / order.length) * 100 + '%';
    if (this.notesEl.style.display === 'block') this.notesEl.textContent = this.slideData.notes || '';
  }

  onKey(e) {
    if (e.target && /input|textarea|select/i.test(e.target.tagName)) return;
    const k = e.key;
    if (this.blank.style.display === 'block' && !['b', 'B', 'w', 'W', '.', ','].includes(k)) {
      this.blank.style.display = 'none';
      return;
    }
    switch (k) {
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter': case 'n': case 'N':
        if (k === 'Enter' && this.typed) { this.goto(Number(this.typed) - 1); this.typed = ''; break; }
        this.next(); break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace': case 'p': case 'P':
        this.prev(); break;
      case 'Home': this.goto(this.order()[0], { instant: false }); break;
      case 'End': this.goto(this.order()[this.order().length - 1], {}); break;
      case 'b': case 'B': case '.': this.blankScreen('#000'); break;
      case 'w': case 'W': case ',': this.blankScreen('#fff'); break;
      case 'f': case 'F': this.toggleFullscreen(); break;
      case 's': case 'S': this.toggleNotes(); break;
      default:
        if (/^[0-9]$/.test(k)) { this.typed = (this.typed || '') + k; return; }
        return;
    }
    e.preventDefault();
  }

  blankScreen(color) {
    const b = this.blank;
    if (b.style.display === 'block') { b.style.display = 'none'; return; }
    b.style.background = color;
    b.style.display = 'block';
  }

  onClick(e) {
    if (e.button !== 0) return;
    const t = e.target;
    if (t.closest('.pp-ui')) return;
    if (this.blank.style.display === 'block') { this.blank.style.display = 'none'; return; }
    const link = t.closest('[data-slide],[data-jump],a[href]');
    if (link && this.handleLink(link, e)) return;
    if (t.closest('.pp-media')) return;
    // Interactive triggers.
    const el = t.closest('.pp-el[data-layer="slide"]');
    const timing = this.slideData.timing;
    if (el && timing && timing.triggers && timing.triggers.length) {
      let top = el;
      while (top.parentElement && top.parentElement.closest('.pp-el[data-layer="slide"]')) top = top.parentElement.closest('.pp-el[data-layer="slide"]');
      const spid = top.dataset.spid;
      const idx = timing.triggers.findIndex((tr) => tr.spid === spid);
      if (idx >= 0) {
        this.fireTrigger(idx);
        return;
      }
    }
    const tr = this.slideData.transition;
    if (tr && tr.advClick === false && this.step >= (timing && timing.main ? timing.main.length : 0)) return;
    this.next();
  }

  handleLink(link, e) {
    if (link.dataset.slide !== undefined) {
      e.preventDefault();
      this.goto(Number(link.dataset.slide));
      return true;
    }
    if (link.dataset.jump) {
      e.preventDefault();
      const order = this.order();
      const pos = order.indexOf(this.cur);
      switch (link.dataset.jump) {
        case 'nextslide': this.goto(order[Math.min(order.length - 1, pos + 1)]); break;
        case 'previousslide': this.goto(order[Math.max(0, pos - 1)]); break;
        case 'firstslide': this.goto(order[0]); break;
        case 'lastslide': this.goto(order[order.length - 1]); break;
        case 'lastslideviewed': if (this.history.length > 1) this.goto(this.history[this.history.length - 2]); break;
        case 'endshow': this.end(); break;
        default: break;
      }
      return true;
    }
    if (link.tagName === 'A' && link.getAttribute('href')) {
      e.stopPropagation();
      return true;
    }
    return false;
  }

  end() {
    this.blankScreen('#000');
  }

  initMedia(root) {
    for (const m of root.querySelectorAll('.pp-media')) {
      const media = m.querySelector('video,audio');
      if (!media) continue;
      const btn = m.querySelector('.pp-media-btn');
      const toggle = (e) => {
        if (e) { e.stopPropagation(); e.preventDefault(); }
        if (media.paused) media.play().catch(() => {});
        else media.pause();
      };
      m.addEventListener('click', toggle);
      if (btn) btn.addEventListener('click', toggle);
      media.addEventListener('play', () => { m.classList.add('pp-playing'); m.classList.remove('pp-paused'); });
      media.addEventListener('pause', () => { m.classList.remove('pp-playing'); m.classList.add('pp-paused'); });
      const st = Number(media.dataset.trimStart || 0) / 1000;
      if (st) media.addEventListener('loadedmetadata', () => { media.currentTime = st; }, { once: true });
      const end = Number(media.dataset.trimEnd || 0) / 1000;
      if (end) media.addEventListener('timeupdate', () => { if (media.duration && media.currentTime >= media.duration - end) media.pause(); });
    }
  }

  initLinks(root, list) {
    if (!list) return;
    root.addEventListener('click', (e) => {
      const link = e.target.closest('[data-slide]');
      if (link) {
        e.preventDefault();
        const s = root.querySelector('#slide-' + (Number(link.dataset.slide) + 1));
        if (s) s.scrollIntoView({ behavior: 'smooth' });
      }
    });
  }

  targetsFor(i) {
    if (!this.targets.has(i)) this.targets.set(i, new SlideTargets(this.sections[i], this.data.slides[i] && this.data.slides[i].timing, this.W, this.H));
    return this.targets.get(i);
  }

  // ---- animation state ---------------------------------------------------
  position(now) {
    return {
      step: this.step,
      t: this.stepStart !== null ? now - this.stepStart : undefined,
      triggers: this.triggers.map((tr) => ({ index: tr.index, step: tr.step, t: tr.start !== null ? now - tr.start : undefined })),
    };
  }

  render(now) {
    const i = this.cur;
    const timing = this.slideData.timing;
    if (!timing) return false;
    const tg = this.targetsFor(i);
    const res = evaluate(timing, this.position(now), (k) => tg.geo(k), (k) => tg.colors(k), (k) => tg.counts.get(k) || 0, { aspect: this.W / this.H });
    tg.apply(res.states);
    this.runCommands(res.commands, tg);
    // Step completion.
    let running = res.running;
    if (this.stepStart !== null) {
      const st = timing.main[this.step];
      const end = st ? stepEnd(st, (k) => tg.counts.get(k) || 0) : 0;
      if (!st || now - this.stepStart >= end) {
        this.step++;
        this.stepStart = null;
        this.onStepComplete(now);
      } else running = true;
    }
    for (const tr of this.triggers) {
      if (tr.start === null) continue;
      const seq = timing.triggers[tr.index];
      const st = seq && seq.steps[tr.step];
      const end = st ? stepEnd(st, (k) => tg.counts.get(k) || 0) : 0;
      if (!st || now - tr.start >= end) { tr.step++; tr.start = null; } else running = true;
    }
    return running;
  }

  runCommands(cmds, tg) {
    for (const { b, tag } of cmds) {
      const id = this.cur + '|' + tag + '|' + b.begin + '|' + (b.spid || '') + '|' + (b.cmd || b.media || '');
      if (this.done.has(id)) continue;
      this.done.add(id);
      if (this.silent) continue;
      if (b.type === 'media' && b.sound && b.sound.url) {
        try {
          const a = new Audio(b.sound.url);
          a.volume = Math.max(0, Math.min(1, b.vol === undefined ? 1 : b.vol));
          a.play().catch(() => {});
          this.sounds.push(a);
        } catch (e) { /* ignore */ }
        continue;
      }
      const media = b.spid ? tg.media(b.spid) : null;
      if (!media) continue;
      const cmd = b.cmd || (b.type === 'media' ? 'play' : '');
      const m = /^(\w+)(?:\(([^)]*)\))?/.exec(cmd);
      const name = m ? m[1] : cmd;
      try {
        if (name === 'playFrom') { media.currentTime = parseFloat(m[2]) || 0; media.play().catch(() => {}); }
        else if (name === 'play' || name === 'resume') media.play().catch(() => {});
        else if (name === 'pause') media.pause();
        else if (name === 'stop') { media.pause(); media.currentTime = 0; }
        else if (name === 'togglePause') { if (media.paused) media.play().catch(() => {}); else media.pause(); }
        else if (name === 'seek') media.currentTime = parseFloat(m[2]) || 0;
      } catch (e) { /* ignore */ }
    }
  }

  loop() {
    cancelAnimationFrame(this.raf);
    const frame = (now) => {
      let running = false;
      if (this.transition) running = this.renderTransition(now) || running;
      if (!this.transition) running = this.render(now) || running;
      if (running) this.raf = requestAnimationFrame(frame);
      else this.idle();
    };
    this.raf = requestAnimationFrame(frame);
  }

  idle() {
    this.scheduleAdvance();
    if (this.onidle) this.onidle();
  }

  onStepComplete(now) {
    const timing = this.slideData.timing;
    const st = timing && timing.main[this.step];
    if (st && st.auto) this.startStep(now + (st.delay || 0));
  }

  startStep(at) {
    this.stepStart = at !== undefined ? at : performance.now();
    this.loop();
  }

  fireTrigger(index) {
    const timing = this.slideData.timing;
    const seq = timing.triggers[index];
    let tr = this.triggers.find((x) => x.index === index);
    if (!tr) { tr = { index, step: 0, start: null }; this.triggers.push(tr); }
    if (tr.start !== null) { tr.step++; tr.start = null; }
    if (tr.step >= seq.steps.length) tr.step = 0;
    tr.start = performance.now();
    this.loop();
  }

  scheduleAdvance() {
    clearTimeout(this.advTimer);
    const tr = this.slideData.transition;
    if (!tr || tr.advTm === null || tr.advTm === undefined || this.transition || this.stepStart !== null) return;
    if (this.data.show && this.data.show.useTimings === false) return;
    const elapsed = performance.now() - (this.slideStart || 0);
    const wait = Math.max(0, tr.advTm - elapsed);
    this.advTimer = setTimeout(() => this.next(), wait);
  }

  // ---- navigation ------------------------------------------------------------
  next() {
    if (this.transition) { this.finishTransition(); return; }
    const timing = this.slideData.timing;
    const steps = timing ? timing.main : [];
    const now = performance.now();
    if (this.stepStart !== null) {
      // Finish the running step immediately, then continue with the next one.
      this.step++;
      this.stepStart = null;
      this.render(now + 1e9);
    }
    if (this.step < steps.length) {
      this.startStep(now);
      return;
    }
    const order = this.order();
    const pos = order.indexOf(this.cur);
    if (pos + 1 < order.length) this.goto(order[pos + 1]);
    else if (this.data.show && this.data.show.loop) this.goto(order[0]);
    else this.end();
  }

  prev() {
    if (this.transition) this.finishTransition();
    const timing = this.slideData.timing;
    if (this.stepStart !== null) { this.stepStart = null; }
    if (timing && this.step > 0) {
      this.step--;
      this.stepStart = null;
      this.triggers = [];
      this.stopSounds();
      this.render(performance.now());
      return;
    }
    const order = this.order();
    const pos = order.indexOf(this.cur);
    if (pos > 0) this.goto(order[pos - 1], { instant: true, atEnd: true });
  }

  stopSounds() {
    for (const a of this.sounds) { try { a.pause(); } catch (e) { /* ignore */ } }
    this.sounds = [];
  }

  /**
   * Show slide i.
   * @param {object} o {instant: skip transition, atEnd: show with all animations completed, step: completed steps}
   */
  goto(i, o = {}) {
    if (i === undefined || i === null || i < 0 || i >= this.sections.length) return;
    clearTimeout(this.advTimer);
    if (this.transition) this.finishTransition();
    const prev = this.cur;
    const prevSec = prev >= 0 ? this.sections[prev] : null;
    this.stopSounds();
    if (prevSec) {
      for (const m of prevSec.querySelectorAll('video,audio')) { try { m.pause(); } catch (e) { /* ignore */ } }
    }
    this.cur = i;
    this.history.push(i);
    const timing = this.slideData.timing;
    this.step = o.atEnd && timing ? timing.main.length : (o.step || 0);
    this.stepStart = null;
    this.triggers = [];
    this.done = new Set();
    for (const s of this.sections) s.classList.remove('pp-current', 'pp-prev', 'pp-on-top');
    const sec = this.sections[i];
    sec.classList.add('pp-current');
    try { if (location.hash !== '#slide-' + (i + 1)) history.replaceState(null, '', '#slide-' + (i + 1)); } catch (e) { /* ignore */ }
    this.updateUi();
    // Initial animation state (silently, without running media commands).
    this.silent = true;
    this.render(performance.now());
    this.silent = false;
    const tr = this.slideData.transition;
    if (!o.instant && prevSec && prevSec !== sec && tr && tr.type && tr.type !== 'none' && !this.opts.noTransitions) {
      prevSec.classList.add('pp-prev');
      this.transition = { tr, from: prevSec, to: sec, start: performance.now(), dur: Math.max(1, tr.dur || 500) };
      if (tr.sound && tr.sound.url) {
        try { const a = new Audio(tr.sound.url); a.loop = !!tr.loopSound; a.play().catch(() => {}); this.sounds.push(a); } catch (e) { /* ignore */ }
      }
      if (tr.type === 'morph') this.prepareMorph(prevSec, sec);
      this.loop();
    } else {
      this.afterEnter();
    }
  }

  afterEnter() {
    this.slideStart = performance.now();
    const timing = this.slideData.timing;
    const st = timing && timing.main[this.step];
    if (st && st.auto) this.startStep(performance.now() + (st.delay || 0));
    else this.loop();
  }

  renderTransition(now) {
    const T = this.transition;
    const p = Math.min(1, (now - T.start) / T.dur);
    const f = transitionFrame(T.tr, p);
    const apply = (el, st) => {
      el.style.opacity = st.opacity !== undefined ? String(st.opacity) : '';
      el.style.transform = st.transform || '';
      el.style.transformOrigin = st.transformOrigin || '';
      el.style.clipPath = st.clipPath || '';
      el.style.webkitClipPath = st.clipPath || '';
      el.style.maskImage = st.maskImage || '';
      el.style.webkitMaskImage = st.webkitMaskImage || '';
      el.style.maskSize = st.maskSize || '';
      el.style.webkitMaskSize = st.webkitMaskSize || '';
      el.style.maskRepeat = st.maskRepeat || '';
      el.style.webkitMaskRepeat = st.webkitMaskRepeat || '';
      el.style.backfaceVisibility = st.backfaceVisibility || '';
    };
    apply(T.from, f.out || {});
    apply(T.to, f.inc || {});
    T.from.classList.toggle('pp-on-top', !!f.outOnTop);
    this.back.style.background = f.back || '';
    if (T.morph) this.renderMorph(p);
    if (p >= 1) {
      this.finishTransition();
      return false;
    }
    return true;
  }

  finishTransition() {
    const T = this.transition;
    if (!T) return;
    this.transition = null;
    for (const el of [T.from, T.to]) {
      for (const k of ['opacity', 'transform', 'transformOrigin', 'clipPath', 'webkitClipPath', 'maskImage', 'webkitMaskImage', 'maskSize', 'webkitMaskSize', 'maskRepeat', 'webkitMaskRepeat', 'backfaceVisibility']) el.style[k] = '';
    }
    T.from.classList.remove('pp-prev', 'pp-on-top');
    this.back.style.background = '';
    if (T.morph) this.endMorph(T.morph);
    this.afterEnter();
  }

  // ---- morph ---------------------------------------------------------------
  prepareMorph(from, to) {
    const els = (sec) => [...sec.querySelectorAll(':scope > .pp-el')];
    const kind = (e) => (e.className.match(/pp-(sp|pic|grp|tbl|chart|media)\b/) || [''])[0];
    const key = (e) => {
      const n = e.dataset.name || '';
      return n.startsWith('!!') ? n : n + '|' + kind(e);
    };
    const fromMap = new Map();
    for (const e of els(from)) if (!fromMap.has(key(e))) fromMap.set(key(e), e);
    const pairs = [];
    for (const e of els(to)) {
      const k = key(e);
      const a = fromMap.get(k);
      if (a && (e.dataset.name || '')) {
        pairs.push({ a, b: e, ra: this.box(a), rb: this.box(e), rotA: parseFloat(a.dataset.rot || '0'), rotB: parseFloat(e.dataset.rot || '0') });
        fromMap.delete(k);
      }
    }
    const saved = new Map();
    for (const sec of [from, to]) {
      for (const e of sec.querySelectorAll(':scope > .pp-el, :scope > .pp-bg')) saved.set(e, e.getAttribute('style') || '');
    }
    this.transition.morph = { pairs, saved };
  }

  box(e) {
    return { x: parseFloat(e.style.left) || 0, y: parseFloat(e.style.top) || 0, w: parseFloat(e.style.width) || 1, h: parseFloat(e.style.height) || 1 };
  }

  renderMorph(p) {
    const M = this.transition.morph;
    const s = p * p * (3 - 2 * p);
    const paired = new Set();
    for (const pr of M.pairs) {
      const { ra, rb } = pr;
      paired.add(pr.b);
      const cx = ra.x + ra.w / 2 + (rb.x + rb.w / 2 - ra.x - ra.w / 2) * s;
      const cy = ra.y + ra.h / 2 + (rb.y + rb.h / 2 - ra.y - ra.h / 2) * s;
      const sx = (ra.w + (rb.w - ra.w) * s) / rb.w;
      const sy = (ra.h + (rb.h - ra.h) * s) / rb.h;
      const rot = pr.rotA + (pr.rotB - pr.rotA) * s;
      const tx = cx - (rb.x + rb.w / 2), ty = cy - (rb.y + rb.h / 2);
      const fh = pr.b.dataset.fh === '1' ? -1 : 1, fv = pr.b.dataset.fv === '1' ? -1 : 1;
      pr.b.style.transform = `translate(${tx}px,${ty}px) rotate(${rot}deg) scale(${sx * fh},${sy * fv})`;
      pr.b.style.opacity = String(Math.min(1, s * 2));
      pr.a.style.opacity = String(1 - s);
    }
    // Paired shapes move; everything else cross-fades.
    this.transition.to.style.opacity = '1';
    for (const e of this.transition.to.querySelectorAll(':scope > .pp-el, :scope > .pp-bg')) {
      if (!paired.has(e)) e.style.opacity = String(s);
    }
  }

  endMorph(M) {
    for (const [e, st] of M.saved) e.setAttribute('style', st);
    this.render(performance.now());
  }

  // ---- testing API -----------------------------------------------------------
  /** Jump to (slide, completed steps, time into the next step) without playing. */
  seek(slide, step = 0, t = undefined) {
    this.goto(slide, { instant: true, step });
    cancelAnimationFrame(this.raf);
    clearTimeout(this.advTimer);
    if (t !== undefined) {
      const now = performance.now();
      this.stepStart = now - t;
      this.render(now);
      this.stepStart = null;
      this._frozen = { slide, step, t };
    }
    return this.state();
  }

  /**
   * Render the transition into slide i frozen at progress p (0..1), starting
   * from the slide currently shown. Used by the self-verification suite.
   */
  previewTransition(i, p) {
    const from = this.sections[this.cur];
    this.goto(i, { instant: true });
    cancelAnimationFrame(this.raf);
    clearTimeout(this.advTimer);
    const to = this.sections[i];
    const tr = this.slideData.transition;
    if (!tr || !from || from === to) return false;
    from.classList.add('pp-prev');
    const dur = Math.max(1, tr.dur || 500);
    const start = performance.now() - p * dur;
    this.transition = { tr, from, to, start, dur };
    if (tr.type === 'morph') this.prepareMorph(from, to);
    if (p >= 1) { this.finishTransition(); cancelAnimationFrame(this.raf); return true; }
    this.renderTransition(start + p * dur);
    return true;
  }

  state() {
    const timing = this.slideData.timing;
    return {
      slide: this.cur,
      step: this.step,
      steps: timing ? timing.main.length : 0,
      running: this.stepStart !== null || !!this.transition,
      transition: this.transition ? this.transition.tr.type : null,
    };
  }
}

/** Auto-start when loaded as a standalone document. */
export function autoStart() {
  const deck = document.querySelector('.pp-deck');
  const dataEl = document.getElementById('pp-data');
  if (!deck || !dataEl) return null;
  const data = JSON.parse(dataEl.textContent);
  const params = new URLSearchParams(location.search);
  const mode = params.get('mode') || (deck.classList.contains('pp-list') ? 'list' : 'deck');
  const player = new Player(deck, data, { mode, showHidden: params.has('hidden'), noTransitions: params.has('notransitions') }).init();
  window.pptxPlayer = player;
  return player;
}

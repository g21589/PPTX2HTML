/**
 * Text layout that needs measuring in the browser: custom tab stops and
 * WordArt warps.
 *
 * Tab characters are rendered as <span class="pp-tab"> inside paragraphs
 * carrying data-tabs ("pos:algn,..." in px from the text box edge) and
 * data-deftab. Each tab gets the width that moves the following text to its
 * stop. Measurements use offsetLeft, which ignores the transforms used to
 * scale slides. Without this script tabs fall back to CSS tab-size.
 */

function stopsOf(p) {
  const list = (p.getAttribute('data-tabs') || '').split(',').filter(Boolean).map((t) => {
    const [pos, algn] = t.split(':');
    return { pos: Number(pos), algn: algn || 'l' };
  });
  return list.sort((a, b) => a.pos - b.pos);
}

/** Lay out the tabs of every paragraph under root (idempotent). */
export function layoutTabs(root) {
  for (const p of root.querySelectorAll('.pp-p[data-deftab]')) {
    const tabs = [...p.querySelectorAll('.pp-tab')];
    if (!tabs.length || p.offsetParent === null) continue;
    const stops = stopsOf(p);
    const def = Number(p.getAttribute('data-deftab')) || 96;
    const end = p.querySelector('.pp-tab-end');
    for (const t of tabs) {
      t.textContent = '';
      t.style.display = 'inline-block';
      t.style.width = '0px';
    }
    const left = (el) => el.offsetLeft - p.offsetLeft;
    tabs.forEach((t, i) => {
      const x = left(t);
      let stop = stops.find((s) => s.pos > x + 0.5);
      if (!stop) stop = { pos: (Math.floor((x + 0.5) / def) + 1) * def, algn: 'l' };
      let w = stop.pos - x;
      if (stop.algn !== 'l') {
        // Width of the text up to the next tab (or the end of the line).
        const next = tabs[i + 1] || end;
        const seg = next && next.offsetTop === t.offsetTop ? Math.max(0, left(next) - x) : 0;
        w -= stop.algn === 'ctr' ? seg / 2 : seg;
      }
      t.style.width = Math.max(0, w) + 'px';
    });
  }
}

// ---- WordArt warps ----------------------------------------------------------

/** Polyline with arc-length parameterisation. */
function track(flat) {
  const pts = [];
  for (let i = 0; i + 1 < flat.length; i += 2) {
    const p = [flat[i], flat[i + 1]];
    const q = pts[pts.length - 1];
    if (!q || Math.hypot(p[0] - q[0], p[1] - q[1]) > 1e-6) pts.push(p);
  }
  const len = [0];
  for (let i = 1; i < pts.length; i++) len.push(len[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
  const total = len[len.length - 1] || 1;
  /** Point and unit tangent at fraction u of the length. */
  const at = (u) => {
    if (pts.length < 2) return { x: pts[0] ? pts[0][0] : 0, y: pts[0] ? pts[0][1] : 0, tx: 1, ty: 0 };
    const d = Math.max(0, Math.min(1, u)) * total;
    let i = 1;
    while (i < len.length - 1 && len[i] < d) i++;
    const seg = len[i] - len[i - 1] || 1;
    const t = (d - len[i - 1]) / seg;
    const [x0, y0] = pts[i - 1], [x1, y1] = pts[i];
    const l = Math.hypot(x1 - x0, y1 - y0) || 1;
    return { x: x0 + (x1 - x0) * t, y: y0 + (y1 - y0) * t, tx: (x1 - x0) / l, ty: (y1 - y0) / l };
  };
  at.len = total;
  return at;
}

const SPACE = /^\s+$/;

/** Wrap every character of the text body in an inline-block glyph span. */
function splitGlyphs(tx) {
  const walker = document.createTreeWalker(tx, NodeFilter.SHOW_TEXT);
  const nodes = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.nodeValue && !n.parentNode.closest('.pp-bu')) nodes.push(n);
  for (const n of nodes) {
    const frag = document.createDocumentFragment();
    for (const part of n.nodeValue.match(/\s+|\S+/g) || []) {
      const chars = Array.from(part);
      if (SPACE.test(part)) {
        for (const c of chars) frag.appendChild(glyph(c));
        continue;
      }
      const word = document.createElement('span');
      word.className = 'pp-ww';
      word.style.whiteSpace = 'nowrap';
      for (const c of chars) word.appendChild(glyph(c));
      frag.appendChild(word);
    }
    n.parentNode.replaceChild(frag, n);
  }
}

function glyph(c) {
  const g = document.createElement('span');
  g.className = 'pp-wg';
  g.style.display = 'inline-block';
  g.style.whiteSpace = 'pre';
  g.style.transformOrigin = '0 0';
  g.textContent = c;
  return g;
}

const ASCENDERS = /[A-Z0-9bdfhiklt!?#$%&'"()[\]{}\/|@\u00c0-\u024f\u0370-\u03ff\u0400-\u04ff\u2e80-\uffff]/;
const DESCENDERS = /[gjpqyQ,;()[\]{}|@_\u2e80-\uffff]/;

/** Map the glyphs of warped text bodies onto their warp paths (idempotent). */
export function layoutWarps(root) {
  for (const tx of root.querySelectorAll('.pp-tx[data-warp]')) {
    if (tx.offsetParent === null) continue;
    let warp;
    try { warp = JSON.parse(tx.getAttribute('data-warp')); } catch (e) { continue; }
    if (!tx.ppWarpSplit) { splitGlyphs(tx); tx.ppWarpSplit = true; }
    const glyphs = [...tx.querySelectorAll('.pp-wg, .pp-bu')];
    for (const g of glyphs) g.style.transform = '';
    // Measure (offsets are relative to the text box and ignore transforms).
    const rel = (el) => {
      let x = 0, y = 0;
      for (let e = el; e && e !== tx; e = e.offsetParent) { x += e.offsetLeft; y += e.offsetTop; }
      return [x, y];
    };
    const marker = document.createElement('span');
    marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
    const lines = [];
    let cur = null, lastY = -Infinity, lastH = 0, lastPara = null;
    for (const g of glyphs) {
      const [x, y] = rel(g);
      const para = g.closest('.pp-p');
      g.appendChild(marker);
      const m = { g, x, y, w: g.offsetWidth, h: g.offsetHeight, base: rel(marker)[1] - y, size: parseFloat(getComputedStyle(g).fontSize) || 16, text: g.textContent, space: SPACE.test(g.textContent) };
      marker.remove();
      // A wrapped line starts at least half a glyph lower (offsets are whole
      // pixels, so horizontal positions are too coarse to tell).
      if (!cur || para !== lastPara || y >= lastY + lastH / 2) { cur = []; lines.push(cur); }
      cur.push(m);
      lastY = y;
      lastH = m.h;
      lastPara = para;
    }
    const metrics = [];
    for (const ln of lines) {
      const ink = ln.filter((m) => !m.space);
      if (!ink.length) continue;
      // Ink box: cap (or x-) height above the baseline, descenders below.
      const baseline = Math.max(...ink.map((m) => m.y + m.base));
      const size = Math.max(...ink.map((m) => m.size));
      const text = ink.map((m) => m.text).join('');
      metrics.push({
        glyphs: ln,
        x0: Math.min(...ink.map((m) => m.x)),
        x1: Math.max(...ink.map((m) => m.x + m.w)),
        baseline,
        top: baseline - size * (ASCENDERS.test(text) ? 0.72 : 0.52),
        bottom: baseline + (DESCENDERS.test(text) ? size * 0.22 : 0),
        pitch: Math.max(...ln.map((m) => m.h)),
        algn: getComputedStyle(ln[0].g.closest('.pp-p')).textAlign,
      });
    }
    if (!metrics.length) continue;
    const paths = warp.p.map(track);
    const follow = paths.length % 2 === 1; // arches, circle and button follow paths
    const groups = follow ? paths.length : paths.length / 2;
    const L = metrics.length;
    const first = follow && L < groups ? Math.floor((groups - L) / 2) : 0;
    metrics.forEach((ln, i) => {
      ln.group = L <= groups ? first + i : Math.min(groups - 1, Math.floor((i * groups) / L));
    });
    const f = (v) => Math.round(v * 10000) / 10000;
    for (let k = 0; k < groups; k++) {
      const mine = metrics.filter((l) => l.group === k);
      if (!mine.length) continue;
      if (follow) {
        // Text runs along the path on its inner side (towards the centre of
        // the box): it hangs below an arch, stands on a smile, lines stack
        // inwards. Glyphs keep their size; the line is aligned on the path.
        const P = paths[k];
        const len = P.len;
        const mid = P(0.5);
        const cx = tx.offsetWidth / 2, cy = tx.offsetHeight / 2;
        // Up normal of the path is (ty, -tx); inward when it points to the
        // centre. A path through the centre carries vertically centred text.
        const side = mid.ty * (cx - mid.x) - mid.tx * (cy - mid.y);
        const centred = Math.abs(side) < 0.1 * Math.min(cx, cy) * 2;
        const upInward = side > 0;
        let depth = 0;
        for (const ln of mine) {
          const width = ln.x1 - ln.x0 || 1;
          const squeeze = width > len ? len / width : 1;
          const start = ln.algn === 'left' || ln.algn === 'start' ? 0 : ln.algn === 'right' || ln.algn === 'end' ? len - width * squeeze : (len - width * squeeze) / 2;
          const size = Math.max(...ln.glyphs.map((m) => m.size));
          // Signed distance of the baseline above the path.
          const lift = centred ? -size * 0.33 - depth : upInward ? size * 0.25 + depth : -(size * 0.9 + depth);
          for (const m of ln.glyphs) {
            const d = start + (m.x + m.w / 2 - ln.x0) * squeeze;
            const a = P(d / len);
            const cos = a.tx, sin = a.ty;
            const px = a.x + sin * lift, py = a.y - cos * lift;
            const hw = (m.w * squeeze) / 2;
            m.g.style.transform = `matrix(${f(cos * squeeze)},${f(sin * squeeze)},${f(-sin)},${f(cos)},${f(px - (cos * hw - sin * m.base) - m.x)},${f(py - (sin * hw + cos * m.base) - m.y)})`;
          }
          depth += ln.pitch;
        }
        continue;
      }
      const T = paths[2 * k], B = paths[2 * k + 1];
      const H = mine.reduce((a, l) => a + (l.bottom - l.top), 0) || 1;
      let acc = 0;
      for (const ln of mine) {
        const ih = ln.bottom - ln.top || 1;
        const s0 = acc / H, s1 = (acc + ih) / H;
        acc += ih;
        const width = ln.x1 - ln.x0 || 1;
        const corner = (x, y) => {
          const u = (x - ln.x0) / width, v = s0 + ((y - ln.top) / ih) * (s1 - s0);
          const a = T(u), b = B(u);
          return [a.x + (b.x - a.x) * v, a.y + (b.y - a.y) * v];
        };
        for (const m of ln.glyphs) {
          const tl = corner(m.x, m.y), tr = corner(m.x + m.w, m.y), bl = corner(m.x, m.y + m.h), br = corner(m.x + m.w, m.y + m.h);
          const w = m.w || 1, h = m.h || 1;
          const ax = (tr[0] - tl[0] + br[0] - bl[0]) / (2 * w), ay = (tr[1] - tl[1] + br[1] - bl[1]) / (2 * w);
          const cx = (bl[0] - tl[0] + br[0] - tr[0]) / (2 * h), cy = (bl[1] - tl[1] + br[1] - tr[1]) / (2 * h);
          const mid = [(tl[0] + tr[0] + bl[0] + br[0]) / 4, (tl[1] + tr[1] + bl[1] + br[1]) / 4];
          const ox = mid[0] - (ax * w + cx * h) / 2, oy = mid[1] - (ay * w + cy * h) / 2;
          m.g.style.transform = `matrix(${f(ax)},${f(ay)},${f(cx)},${f(cy)},${f(ox - m.x)},${f(oy - m.y)})`;
        }
      }
    }
  }
}

// ---- autofit ------------------------------------------------------------------

/** Height of the paragraphs of a text body (offsets ignore slide scaling). */
function contentHeight(tx) {
  const ps = tx.querySelectorAll(':scope > .pp-p, :scope > div > .pp-p');
  if (!ps.length) return 0;
  const first = ps[0], last = ps[ps.length - 1];
  const cs = getComputedStyle(first), ce = getComputedStyle(last);
  return last.offsetTop + last.offsetHeight + parseFloat(ce.marginBottom) - (first.offsetTop - parseFloat(cs.marginTop));
}

const visible = (n) => {
  if (n.nodeType === 3) return /[^\s\u200b]/.test(n.nodeValue);
  if (n.nodeType !== 1 || n.classList.contains('pp-bu') || n.classList.contains('pp-bu-gap')) return false;
  return [...n.childNodes].some(visible);
};

/**
 * Hide what follows the last visible text (empty paragraphs, trailing line
 * breaks): it takes space but does not show, so it does not count as
 * overflow. Returns a function restoring it.
 */
function hideTrailing(tx) {
  const hidden = [];
  const hide = (el) => { hidden.push([el, el.style.display]); el.style.display = 'none'; };
  const ps = [...tx.querySelectorAll(':scope > .pp-p, :scope > div > .pp-p')];
  let k = ps.length - 1;
  while (k > 0 && !visible(ps[k])) hide(ps[k--]);
  const kids = ps[k] ? [...ps[k].childNodes] : [];
  for (let i = kids.length - 1; i >= 0 && !visible(kids[i]); i--) {
    if (kids[i].nodeType === 1 && !kids[i].classList.contains('pp-tab-end')) hide(kids[i]);
  }
  return () => { for (const [el, d] of hidden) el.style.display = d; };
}

/**
 * Auto-fitting text bodies (normAutofit / spAutoFit) were fitted by Office
 * with the original fonts; substituted fonts can be larger. Shrink the text
 * (through --pp-fs) until it fits its box again. Where Office stored its own
 * fit (data-fit-min), only a small correction is made: needing more means
 * the text overflows in Office too.
 */
export function layoutFit(root) {
  for (const tx of root.querySelectorAll('.pp-tx[data-fit]')) {
    if (tx.offsetParent === null) continue;
    tx.style.removeProperty('--pp-fs');
    const cs = getComputedStyle(tx);
    const avail = tx.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    const nowrap = tx.classList.contains('pp-nowrap');
    const availW = tx.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const over = () => contentHeight(tx) > avail + 1 || (nowrap && tx.scrollWidth > tx.clientWidth + 1 && availW > 0);
    if (avail <= 0) continue;
    const restore = hideTrailing(tx);
    const min = parseFloat(tx.dataset.fitMin) || 0.4;
    let fs = 1;
    if (over()) {
      tx.style.setProperty('--pp-fs', String(min));
      if (!over() || min <= 0.4) {
        let lo = min, hi = 1;
        for (let i = 0; i < 7; i++) {
          const mid = (lo + hi) / 2;
          tx.style.setProperty('--pp-fs', String(mid));
          if (over()) hi = mid; else lo = mid;
        }
        fs = Math.round(lo * 1000) / 1000;
      }
    }
    restore();
    if (fs < 1) tx.style.setProperty('--pp-fs', String(fs));
    else tx.style.removeProperty('--pp-fs');
  }
}

/** All measured text layout for a slide or deck. */
export function layoutText(root) {
  layoutFit(root);
  layoutTabs(root);
  layoutWarps(root);
}

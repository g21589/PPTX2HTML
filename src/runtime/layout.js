/**
 * Text layout that needs measuring in the browser: custom tab stops.
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

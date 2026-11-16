// Previews that use the real output files: the PDF that would be created
// (rendered with pdf.js) and the sheets of the .xlsx that would be written.
import { openPdf } from './pdf-reader.js';
import { el } from './util.js';

// ------------------------------------------------------------------ PDF
/**
 * Render the pages of a PDF blob lazily (only pages near the viewport are drawn).
 * @returns {{ pages:number, goto(i), destroy() }}
 */
export async function showPdf(blob, { pagesEl, thumbsEl, scroller, scale = 1.25, onPage = () => {} }) {
  const pdf = await openPdf(new Uint8Array(await blob.arrayBuffer()));
  const n = pdf.numPages;
  const first = await pdf.getPage(1);
  const vp1 = first.getViewport({ scale: 1 });
  const dpr = window.devicePixelRatio || 1;
  const THUMB_W = 92;
  const pageDivs = [], thumbDivs = [];
  pagesEl.replaceChildren();
  thumbsEl.replaceChildren();
  for (let i = 1; i <= n; i++) {
    const d = el('div', { class: 'pdfpage', 'data-p': String(i) });
    d.style.width = Math.round(vp1.width * scale) + 'px';
    d.style.height = Math.round(vp1.height * scale) + 'px';
    pagesEl.append(d);
    pageDivs.push(d);
    const box = el('div', { class: 'thumb-box' });
    box.style.width = THUMB_W + 'px';
    box.style.height = Math.round(THUMB_W * vp1.height / vp1.width) + 'px';
    const t = el('button', { class: 'thumb-btn', title: `Page ${i}`, onclick: () => d.scrollIntoView({ block: 'start' }) }, box, el('span', {}, String(i)));
    thumbsEl.append(t);
    thumbDivs.push(t);
  }
  const drawn = new Set(), thumbDrawn = new Set();
  let dead = false;

  async function draw(i, target, width, set) {
    if (set.has(i)) return;
    set.add(i);
    try {
      const page = await pdf.getPage(i);
      if (dead) return;
      const sc = (width / vp1.width) * dpr;
      const vp = page.getViewport({ scale: sc });
      const c = document.createElement('canvas');
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      c.style.width = '100%'; c.style.height = '100%'; c.style.display = 'block';
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      if (!dead) target.append(c);
      page.cleanup();
    } catch { /* the page could not be drawn; the placeholder stays */ }
  }
  const io = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) draw(+e.target.dataset.p, e.target, Math.round(vp1.width * scale), drawn);
  }, { root: scroller, rootMargin: '700px 0px' });
  pageDivs.forEach(d => io.observe(d));
  const io2 = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) draw(+e.target.dataset.p, e.target.querySelector('.thumb-box'), THUMB_W, thumbDrawn);
  }, { rootMargin: '300px 0px' });
  thumbDivs.forEach((t, i) => { t.dataset.p = String(i + 1); io2.observe(t); });

  let last = 0;
  const update = () => {
    const mid = scroller.getBoundingClientRect().top + scroller.clientHeight * 0.3;
    let cur = 1;
    for (const d of pageDivs) { if (d.getBoundingClientRect().top <= mid) cur = +d.dataset.p; else break; }
    if (cur !== last) {
      last = cur;
      thumbDivs.forEach((t, i) => t.classList.toggle('on', i + 1 === cur));
      onPage(cur, n);
    }
  };
  let raf = 0;
  const onScroll = () => { cancelAnimationFrame(raf); raf = requestAnimationFrame(update); };
  scroller.addEventListener('scroll', onScroll);
  update();
  return {
    pages: n,
    goto: (i) => pageDivs[Math.max(0, Math.min(n - 1, i - 1))].scrollIntoView({ block: 'start' }),
    destroy() { dead = true; io.disconnect(); io2.disconnect(); scroller.removeEventListener('scroll', onScroll); pdf.destroy(); },
  };
}

// ------------------------------------------------------------------ Excel
/** Write the .xlsx exactly as the app would, then read the sheets back for display. */
export async function buildExcelSheets(model) {
  const { writeXlsx } = await import('./xlsx-writer.js');
  const { loadWorkbook } = await import('./xlsx-reader.js');
  const blob = await writeXlsx(model);
  const wb = await loadWorkbook(await blob.arrayBuffer());
  return wb.sheets.map(sh => {
    let maxR = -1, maxC = -1;
    for (const c of sh.cells.values()) { maxR = Math.max(maxR, c.r); maxC = Math.max(maxC, c.c); }
    for (const m of sh.merges) { maxR = Math.max(maxR, m.r1); maxC = Math.max(maxC, m.c1); }
    return { name: sh.name, cells: sh.cells, merges: sh.merges, maxR, maxC, isText: sh.name === 'Document text' };
  });
}

/** HTML of one sheet (limit = number of rows shown). */
export async function sheetHtml(sh, limit = 500) {
  const { encodeCol } = await import('./xlsx-reader.js');
  const cols = Math.max(sh.maxC + 1, sh.isText ? 1 : 6);
  const total = sh.maxR + 1;
  const rows = Math.min(Math.max(total, sh.isText ? 1 : 18), limit);
  const covered = new Set(), anchors = new Map();
  for (const m of sh.merges) {
    anchors.set(m.r0 * 20000 + m.c0, m);
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(r * 20000 + c);
  }
  const widths = [];
  for (let c = 0; c < cols; c++) {
    let w = 0;
    for (let r = 0; r < Math.min(total, 200); r++) { const cell = sh.cells.get(r * 20000 + c); if (cell) w = Math.max(w, Math.min(48, Math.max(...cell.display.split('\n').map(x => x.length)))); }
    widths.push(sh.isText ? 760 : Math.max(64, Math.round(w * 7.2 + 16)));
  }
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  let h = `<table><colgroup><col width="44">${widths.map(w => `<col width="${w}">`).join('')}</colgroup><thead><tr><th></th>${Array.from({ length: cols }, (_, c) => `<th>${encodeCol(c)}</th>`).join('')}</tr></thead><tbody>`;
  for (let r = 0; r < rows; r++) {
    h += `<tr><th class="rn">${r + 1}</th>`;
    for (let c = 0; c < cols; c++) {
      const key = r * 20000 + c;
      if (covered.has(key)) continue;
      const cell = sh.cells.get(key), m = anchors.get(key);
      const span = m ? `${m.r1 > m.r0 ? ` rowspan="${m.r1 - m.r0 + 1}"` : ''}${m.c1 > m.c0 ? ` colspan="${m.c1 - m.c0 + 1}"` : ''}` : '';
      const cls = [sh.isText ? 'tx' : (r === 0 ? 'hd' : 'dt'), cell && typeof cell.value === 'number' ? 'num' : '', m ? 'mg' : ''].filter(Boolean).join(' ');
      h += `<td class="${cls}"${span}>${cell ? esc(cell.display).replace(/\n/g, '<br>') : ''}</td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  return { html: h, shown: rows, total };
}

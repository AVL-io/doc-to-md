// Markdown reader: read, search, check and edit a Markdown file, and see how it
// would look as Word, PDF, Excel or a web page before any document is created.
import { el, download, esc } from './lib/util.js';
import { analyzeMarkdown } from './lib/md-checks.js';
import { readMarkdown } from './lib/md-reader.js';
import { modelToParts } from './lib/model-html.js';
import { markdownToBlocksHtml } from './lib/html-writer.js';

const $ = (id) => document.getElementById(id);
const VIEW_LABEL = { md: 'Markdown', word: 'Word', pdf: 'PDF', excel: 'Excel', web: 'Web page' };
const CREATE_AS = { word: ['docx', 'Word (.docx)'], pdf: ['pdf', 'PDF (.pdf)'], excel: ['xlsx', 'Excel (.xlsx)'], web: ['html', 'web page (.html)'] };
const PREF_KEY = 'doc-to-md:reader';
const BIG = 400000;            // above this size, syntax colours are switched off in the editor
const MAX_MARKS = 2000;

const S = {
  opts: null, text: '', original: '', view: 'md', split: false, scale: 1, theme: 'auto',
  q: '', qi: 0, marks: [], analysis: null, model: null, blocks: [], caret: { line: 1, col: 1 },
  editor: null, preview: null, checksOpen: true, rt: 0, lock: false, ready: false, timers: {},
};

const LH = () => Math.round(21 * S.scale);

// ------------------------------------------------------------------ preferences (per viewer)
function loadPrefs() { try { return JSON.parse(localStorage.getItem(PREF_KEY) || '{}'); } catch { return {}; } }
function savePrefs() { try { localStorage.setItem(PREF_KEY, JSON.stringify({ theme: S.theme, scale: S.scale })); } catch { /* storage unavailable */ } }

// ------------------------------------------------------------------ public entry point
/**
 * @param {object} o  { name, text, original?, view?, nav?:{index,total,go(delta)},
 *                      onEdit(text, reverted), create(formats, text), formats(), notify(kind,title,text) }
 */
export function openReader(o) {
  if (!S.ready) init();
  flushEdit();
  S.opts = o;
  S.text = o.text;
  S.original = o.original ?? o.text;
  S.view = o.view || 'md';
  S.split = false;
  S.q = '';
  $('rd-q').value = '';
  S.checksOpen = true;
  const dlg = $('reader');
  applyTheme();
  $('rd-name').textContent = o.name;
  const nav = o.nav;
  $('rd-nav').hidden = !nav || nav.total < 2;
  if (nav) $('rd-pos').textContent = ` ${nav.index + 1} / ${nav.total} `;
  analyse();
  refreshChrome();
  renderPanes();
  if (!dlg.open) {
    // notices must stay visible above the reader window
    dlg.append($('notices'));
    dlg.showModal();
  }
}

export function closeReader() { const d = $('reader'); if (d.open) d.close(); }

// ------------------------------------------------------------------ one-time setup
function init() {
  S.ready = true;
  const prefs = loadPrefs();
  S.theme = prefs.theme || 'auto';
  S.scale = prefs.scale || 1;
  const dlg = $('reader');

  $('rd-close').addEventListener('click', closeReader);
  $('rd-back').addEventListener('click', closeReader);
  dlg.addEventListener('close', () => {
    flushEdit();
    disposePanes();
    document.body.append($('notices'));
    $('rd-pop').hidden = true;
  });
  dlg.addEventListener('click', (e) => { if (!e.target.closest('.create-wrap')) $('rd-pop').hidden = true; });
  $('rd-prev').addEventListener('click', () => S.opts.nav && S.opts.nav.go(-1));
  $('rd-next').addEventListener('click', () => S.opts.nav && S.opts.nav.go(1));

  $('rd-views').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-view]');
    if (!b) return;
    setView(b.dataset.view);
  });
  $('rd-split').addEventListener('change', (e) => {
    S.split = e.target.checked;
    if (S.split && S.view === 'md') S.view = 'word';
    S.checksOpen = S.view === 'md';
    renderPanes();
    refreshChrome();
  });
  $('rd-fs-').addEventListener('click', () => setScale(S.scale - 0.15));
  $('rd-fs+').addEventListener('click', () => setScale(S.scale + 0.15));
  $('rd-theme').addEventListener('click', () => {
    const dark = S.theme === 'dark' || (S.theme === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
    S.theme = dark ? 'light' : 'dark';
    applyTheme();
    savePrefs();
  });

  $('rd-q').addEventListener('input', () => { S.q = $('rd-q').value; S.qi = 0; runSearch(); setActive(true); });
  $('rd-q').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); stepMatch(e.shiftKey ? -1 : 1); }
  });
  $('rd-qn').addEventListener('click', () => stepMatch(1));
  $('rd-qp').addEventListener('click', () => stepMatch(-1));

  $('rd-save').addEventListener('click', save);
  $('rd-revert').addEventListener('click', revert);
  $('rd-create').addEventListener('click', (e) => {
    e.stopPropagation();
    if (S.view === 'md') {
      const wanted = new Set((S.opts.formats && S.opts.formats()) || ['docx']);
      document.querySelectorAll('input[name="rdfmt"]').forEach(i => { i.checked = wanted.has(i.value); });
      $('rd-pop').hidden = !$('rd-pop').hidden;
    } else {
      createFormats([CREATE_AS[S.view][0]]);
    }
  });
  $('rd-pop-go').addEventListener('click', () => {
    const f = [...document.querySelectorAll('input[name="rdfmt"]:checked')].map(i => i.value);
    if (!f.length) { S.opts.notify && S.opts.notify('warn', 'No output format selected', 'Select at least one format.'); return; }
    $('rd-pop').hidden = true;
    createFormats(f);
  });
  $('rd-status-r').addEventListener('click', () => { S.checksOpen = !S.checksOpen; renderChecks(); });

  dlg.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); $('rd-q').focus(); $('rd-q').select(); }
    else if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); save(); }
  });
}

// ------------------------------------------------------------------ state helpers
function analyse() {
  S.analysis = analyzeMarkdown(S.text);
  try { S.model = readMarkdown(S.text); } catch { S.model = { blocks: [], footnotes: [] }; }
  S.blocks = S.model.blocks.filter(b => b.line);
}

const baseName = () => S.opts.name.replace(/\.[^.]+$/, '');
const dirty = () => S.text !== S.original;

function applyTheme() {
  const dlg = $('reader');
  if (S.theme === 'auto') dlg.removeAttribute('data-theme'); else dlg.dataset.theme = S.theme;
  dlg.style.setProperty('--rd-scale', String(S.scale));
  dlg.style.setProperty('--lh', LH() + 'px');
  dlg.style.setProperty('--rd-fs', (13 * S.scale).toFixed(1) + 'px');
}

function setScale(v) {
  S.scale = Math.max(0.85, Math.min(1.75, Math.round(v * 100) / 100));
  applyTheme();
  savePrefs();
  if (S.preview && S.preview.view === 'pdf') renderPreview();
}

function setView(v) {
  if (v === 'md') { S.split = false; S.view = 'md'; }
  else S.view = v;
  S.checksOpen = S.view === 'md';
  $('rd-split').checked = S.split;
  renderPanes();
  refreshChrome();
}

function refreshChrome() {
  const st = S.analysis.stats;
  const merged = st.tables.reduce((s, t) => s + (t.merged || 0), 0);
  $('rd-meta').textContent = [
    `${st.lines} line${st.lines === 1 ? '' : 's'}`, `${st.words.toLocaleString()} words`,
    `${st.headings.length} heading${st.headings.length === 1 ? '' : 's'}`,
    `${st.tables.length} table${st.tables.length === 1 ? '' : 's'}${merged ? ` (${merged} merged cell${merged === 1 ? '' : 's'})` : ''}`,
    `${st.footnotes.length} footnote${st.footnotes.length === 1 ? '' : 's'}`,
  ].join(' · ');
  const d = dirty();
  $('rd-edited').hidden = !d;
  $('rd-revert').hidden = !d;
  $('rd-split').checked = S.split;
  document.querySelectorAll('#rd-views button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.view === S.view)));
  const lbl = S.view === 'md' ? 'Create documents ▾' : `Create ${CREATE_AS[S.view][1]}`;
  $('rd-create').textContent = lbl;
  const bad = S.analysis.items.filter(i => i.status === 'bad').length, warn = S.analysis.items.length - bad;
  const r = $('rd-status-r');
  r.replaceChildren(
    ...(bad ? [el('span', { class: 'chip2 bad' }, `🔴 ${bad}`)] : []),
    ...(warn ? [el('span', { class: 'chip2 warn' }, `🟠 ${warn}`)] : []),
    document.createTextNode(S.analysis.items.length ? ' checks' : '🟢 No issues found'));
  renderChecks();
}

function setNote(t) { $('rd-status-note').textContent = t || ''; }

function setStatusLeft() {
  const t = $('rd-status-l');
  if (S.editor) t.textContent = `Ln ${S.caret.line}, Col ${S.caret.col} · UTF-8 · Markdown${dirty() ? ' · edited' : ''}`;
  else if (S.preview && S.preview.status) t.textContent = S.preview.status();
  else t.textContent = '';
}

// ------------------------------------------------------------------ edits
function scheduleEditNotify() {
  clearTimeout(S.timers.edit);
  S.timers.edit = setTimeout(flushEdit, 500);
}
function flushEdit() {
  clearTimeout(S.timers.edit);
  if (!S.opts || !S.opts.onEdit) return;
  if (S.opts._sent === S.text) return;
  S.opts._sent = S.text;
  S.opts.onEdit(S.text, S.text === S.original);
}

function afterTextChange() {
  analyse();
  refreshChrome();
  buildSide();
  if (S.editor) { flagsRefresh(S.editor); onCaret(true); }
  if (S.preview) renderPreview();
  runSearch(true);
}
function scheduleRefresh() {
  clearTimeout(S.timers.refresh);
  S.timers.refresh = setTimeout(afterTextChange, 250);
}

function revert() {
  if (!dirty()) return;
  S.text = S.original;
  if (S.editor) { S.editor.ta.value = S.text; editorRefresh(S.editor); }
  afterTextChange();
  flushEdit();
}

function save() {
  const name = dirty() ? `${baseName()} (edited).md` : `${baseName()}.md`;
  download(new Blob([S.text], { type: 'text/markdown;charset=utf-8' }), name);
  S.opts.notify && S.opts.notify('info', 'Markdown saved', name);
}

async function createFormats(f) {
  flushEdit();
  const b = $('rd-create');
  b.disabled = true;
  try { await S.opts.create(f, S.text); } finally { b.disabled = false; }
}

// ------------------------------------------------------------------ panes
function disposePanes() {
  S.rt++;
  if (S.preview && S.preview.dispose) S.preview.dispose();
  S.preview = null; S.editor = null;
  $('rd-main').replaceChildren();
}

function paneBox(label, kind, content) {
  return el('div', { class: `pane ${kind}` }, el('div', { class: 'pane-h' }, label), content);
}

function renderPanes() {
  disposePanes();
  const wrap = el('div', { class: 'panes' });
  if (S.view === 'md' || S.split) {
    S.editor = buildEditor();
    wrap.append(paneBox('Markdown source (editable)', 'src', S.editor.root));
  }
  if (S.view !== 'md') {
    S.preview = buildPreview(S.view);
    wrap.append(paneBox(S.split ? `${VIEW_LABEL[S.view]} view` : '', 'prev', S.preview.root));
  }
  $('rd-main').append(wrap);
  buildSide();
  if (S.editor) { editorRefresh(S.editor); flagsRefresh(S.editor); onCaret(true); }
  if (S.preview) renderPreview();
  setStatusLeft();
  runSearch();
}

// ------------------------------------------------------------------ editor
const TAG_RE = /(&lt;\/?[a-z][a-z0-9]*[^&]*?&gt;)/gi;
function hlLine(line) {
  if (line.startsWith('<!--')) return `<span class="t-c">${esc(line)}</span>`;
  if (/^ {0,3}#{1,6}\s/.test(line)) return `<span class="t-h">${esc(line)}</span>`;
  let s = esc(line).replace(/&quot;/g, '"');
  s = s.replace(TAG_RE, '<span class="t-tag">$1</span>');
  s = s.replace(/\*\*(.+?)\*\*/g, '<span class="t-b">**$1**</span>');
  s = s.replace(/^(\s*(?:[-*+]|\d+[.)]) )/, '<span class="t-l">$1</span>');
  s = s.replace(/(\[\^[^\]\s]+\]:?)/g, '<span class="t-l">$1</span>');
  s = s.replace(/(!\[[^\]]*\]\([^)]*\))/g, '<span class="t-img">$1</span>');
  return s;
}

function buildEditor() {
  const gutter = el('pre', { class: 'ed-gutter' });
  const flags = el('div', { class: 'ed-flags' });
  const pre = el('pre', { class: 'ed-hl', 'aria-hidden': 'true' });
  const ta = el('textarea', { class: 'ed-ta', spellcheck: 'false', autocapitalize: 'off', autocomplete: 'off', wrap: 'off', 'aria-label': 'Markdown source' });
  ta.value = S.text;
  ta.setSelectionRange(0, 0);
  const root = el('div', { class: 'ed scroller' }, el('div', { class: 'ed-gcol' }, gutter, flags), el('div', { class: 'ed-code' }, pre, ta));
  const ed = { root, scroller: root, ta, pre, gutter, flags, curDivs: [] };
  ta.addEventListener('input', () => {
    S.text = ta.value;
    editorRefresh(ed);
    clearTimeout(S.timers.marks);
    S.timers.marks = setTimeout(() => runSearch(true), 150);
    scheduleRefresh();
    scheduleEditNotify();
    refreshChrome();
    onCaret(false);
  });
  ['keyup', 'click', 'focus'].forEach(ev => ta.addEventListener(ev, () => onCaret(true)));
  ta.addEventListener('keydown', (e) => {
    if (e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey) {
      // keep keyboard users able to leave the editor with Esc; Tab inserts two spaces
      e.preventDefault();
      document.execCommand('insertText', false, '  ');
    }
  });
  root.addEventListener('scroll', () => { if (S.split && S.preview) syncFromEditor(); });
  return ed;
}

function editorRefresh(ed) {
  const text = ed.ta.value;
  const big = text.length > BIG;
  ed.root.classList.toggle('plain', big);
  if (big) ed.pre.textContent = text + '\n';
  else ed.pre.innerHTML = text.split('\n').map(l => `<div class="l">${hlLine(l) || ''}</div>`).join('');
  const n = text.split('\n').length;
  ed.gutter.textContent = Array.from({ length: n }, (_, i) => i + 1).join('\n');
  ed.curDivs = [];
  flagsRefresh(ed);
}

function flagsRefresh(ed) {
  ed.flags.replaceChildren(...S.analysis.items.slice(0, 400).map(i => el('i', {
    class: `dot ${i.status}`, title: i.text, style: `top:${8 + (i.line - 1) * LH()}px`,
  }, '●')));
  // wavy underline on the flagged lines
  if (!ed.root.classList.contains('plain')) {
    ed.pre.querySelectorAll('.l.warn,.l.bad').forEach(d => d.classList.remove('warn', 'bad'));
    for (const i of S.analysis.items) {
      const d = ed.pre.children[i.line - 1];
      if (d && !d.classList.contains('bad')) d.classList.add(i.status);
    }
  }
}

function caretInfo(ta) {
  const pos = ta.selectionStart;
  const before = ta.value.slice(0, pos);
  const nl = before.lastIndexOf('\n');
  return { line: before.split('\n').length, col: pos - nl };
}

function blockAt(line) {
  let hit = null;
  for (const b of S.blocks) { if (b.line <= line && line <= b.endLine) { hit = b; break; } if (b.line > line) break; }
  return hit;
}

function setCurLines(a, b) {
  const ed = S.editor;
  if (!ed || ed.root.classList.contains('plain')) return;
  ed.curDivs.forEach(d => d.classList.remove('cur'));
  ed.curDivs = [];
  if (!a) return;
  for (let i = a; i <= b; i++) { const d = ed.pre.children[i - 1]; if (d) { d.classList.add('cur'); ed.curDivs.push(d); } }
}

function onCaret(link) {
  const ed = S.editor;
  if (!ed) return;
  S.caret = caretInfo(ed.ta);
  const b = blockAt(S.caret.line);
  setCurLines(b ? b.line : 0, b ? b.endLine : 0);
  setStatusLeft();
  if (link && S.preview && S.preview.blks) markPreviewBlock(b, true);
}

function lineOffset(line) {
  const lines = S.text.split('\n');
  let off = 0;
  for (let i = 0; i < Math.min(line - 1, lines.length); i++) off += lines[i].length + 1;
  return { off, len: (lines[line - 1] || '').length };
}

export function goToLine(line, select = true) {
  if (!S.editor) { S.view = 'md'; S.split = false; renderPanes(); refreshChrome(); }
  const ed = S.editor;
  const { off, len } = lineOffset(line);
  ed.ta.focus({ preventScroll: true });
  ed.ta.setSelectionRange(off, select ? off + len : off);
  ed.scroller.scrollTop = Math.max(0, (line - 1) * LH() - ed.scroller.clientHeight / 3);
  onCaret(true);
}

// ------------------------------------------------------------------ previews
function buildPreview(view) {
  const builders = { word: buildWord, web: buildWeb, pdf: buildPdf, excel: buildExcel };
  const pv = builders[view]();
  pv.view = view;
  return pv;
}

async function fillChunks(container, parts, foot, rt) {
  container.replaceChildren();
  const wrap = (p) => `<div class="blk" data-ls="${p.ls}" data-le="${p.le}">${p.html}</div>`;
  const all = [...parts, ...(foot && foot.length ? [{ ls: foot[0].ls, le: foot[0].le, html: '<div class="fn-rule"></div>' }, ...foot] : [])];
  let i = 0;
  const add = (to) => { container.insertAdjacentHTML('beforeend', all.slice(i, to).map(wrap).join('')); i = to; };
  add(Math.min(all.length, 120));
  while (i < all.length) {
    if (rt !== S.rt) return false;
    setNote(`Rendering… ${Math.round(100 * i / all.length)}%`);
    await new Promise(r => setTimeout(r, 0));
    add(Math.min(all.length, i + 150));
  }
  setNote('');
  return true;
}

function collectBlocks(pv) {
  pv.blks = [...pv.content.querySelectorAll('.blk')].filter(b => +b.dataset.ls > 0);
}

function buildWord() {
  const paper = el('div', { class: 'paper word' });
  const scroller = el('div', { class: 'backdrop scroller' }, paper);
  const pv = {
    content: paper, scroller, searchRoot: paper,
    root: el('div', { class: 'pv' }, el('div', { class: 'fid approx' }, '⚠ Approximate preview: built from the same content that is written to the .docx file. Word may break pages and apply fonts slightly differently.'), scroller),
    async render(rt) {
      const { parts, foot } = modelToParts(S.model);
      if (await fillChunks(paper, parts, foot, rt)) { collectBlocks(pv); runSearch(true); onCaret(true); }
    },
  };
  wirePreviewClicks(pv);
  scroller.addEventListener('scroll', () => { if (S.split && S.editor) syncFromPreview(); });
  return pv;
}

function buildWeb() {
  const body = el('div', { class: 'web' });
  const scroller = el('div', { class: 'scroller browser-scroll' }, el('div', { class: 'browser' }, el('div', { class: 'bchrome' }, el('i'), el('i'), el('i'), el('span', { class: 'url' }, `${baseName()}.html`)), body));
  const pv = {
    content: body, scroller, searchRoot: body,
    root: el('div', { class: 'pv' }, el('div', { class: 'fid ok' }, '✓ Exact preview: this is the content of the web page file that would be created.'), scroller),
    async render(rt) {
      const parts = markdownToBlocksHtml(S.text);
      if (await fillChunks(body, parts, null, rt)) { collectBlocks(pv); runSearch(true); onCaret(true); }
    },
  };
  wirePreviewClicks(pv);
  scroller.addEventListener('scroll', () => { if (S.split && S.editor) syncFromPreview(); });
  return pv;
}

function buildPdf() {
  const pagesEl = el('div', { class: 'pdfpages' });
  const scroller = el('div', { class: 'backdrop scroller' }, pagesEl);
  const thumbs = el('div', { class: 'thumbs' });
  let view = null, pageNo = 1, total = 0;
  const pv = {
    scroller, searchRoot: null,
    sideEl: el('div', {}, el('div', { class: 'side-h' }, 'Pages'), thumbs),
    root: el('div', { class: 'pv' }, el('div', { class: 'fid ok' }, '✓ Exact preview: this is the PDF file that would be created.'), scroller),
    status: () => (total ? `Page ${pageNo} of ${total} · ${Math.round(S.scale * 100)}%` : 'Creating the PDF preview…'),
    async render(rt) {
      setNote('Creating the PDF preview…');
      const keep = scroller.scrollTop;
      try {
        const { writePdf } = await import('./lib/pdf-writer.js');
        const { showPdf } = await import('./lib/preview.js');
        const blob = await writePdf(S.model, { title: baseName() });
        if (rt !== S.rt) return;
        if (view) view.destroy();
        view = await showPdf(blob, { pagesEl, thumbsEl: thumbs, scroller, scale: 1.25 * S.scale, onPage: (p, n) => { pageNo = p; total = n; setStatusLeft(); } });
        total = view.pages;
        scroller.scrollTop = keep;
      } catch (e) {
        pagesEl.replaceChildren(el('p', { class: 'muted' }, 'The PDF preview could not be created: ' + (e && e.message ? e.message : e)));
      }
      setNote('');
      setStatusLeft();
    },
    dispose() { if (view) view.destroy(); },
  };
  scroller.addEventListener('scroll', () => { if (S.split && S.editor) syncProportion(scroller, S.editor.scroller); });
  return pv;
}

function buildExcel() {
  const grid = el('div', { class: 'xl scroller' });
  const tabs = el('div', { class: 'xtabs' });
  const sideList = el('ul', { class: 'sl' });
  let sheets = [], cur = 0, limit = 500;
  const pv = {
    scroller: grid, searchRoot: grid,
    sideEl: el('div', {}, el('div', { class: 'side-h' }, 'Sheets'), sideList),
    root: el('div', { class: 'pv' }, el('div', { class: 'fid ok' }, '✓ Structure preview: sheets, values and merged cells as written to the .xlsx file. Fonts and column widths may differ in Excel.'), grid, tabs),
    status: () => (sheets.length ? `Sheet ${cur + 1} of ${sheets.length} · ${sheets[cur].name}` : 'Creating the Excel preview…'),
    async show(i, rt = S.rt) {
      cur = i;
      limit = 500;
      await draw(rt);
    },
    async render(rt) {
      setNote('Creating the Excel preview…');
      try {
        const { buildExcelSheets } = await import('./lib/preview.js');
        const s = await buildExcelSheets(S.model);
        if (rt !== S.rt) return;
        sheets = s;
        cur = Math.min(cur, sheets.length - 1);
        await draw(rt);
      } catch (e) {
        grid.replaceChildren(el('p', { class: 'muted', style: 'padding:14px' }, 'The Excel preview could not be created: ' + (e && e.message ? e.message : e)));
      }
      setNote('');
    },
  };
  async function draw(rt) {
    const { sheetHtml } = await import('./lib/preview.js');
    if (rt !== S.rt || !sheets[cur]) return;
    const { html, shown, total } = await sheetHtml(sheets[cur], limit);
    grid.innerHTML = html;
    if (shown < total) {
      const more = el('button', { class: 'btn small', style: 'margin:10px' }, `Show more rows (${(total - shown).toLocaleString()} not shown)`);
      more.addEventListener('click', () => { limit += 1000; draw(rt); });
      grid.append(more);
    }
    tabs.replaceChildren(...sheets.map((s, i) => el('span', { class: i === cur ? 'on' : '', role: 'button', tabindex: '0', onclick: () => pv.show(i) }, s.name)));
    sideList.replaceChildren(...sheets.map((s, i) => el('li', { class: i === cur ? 'on' : '', onclick: () => pv.show(i) }, s.name)));
    setStatusLeft();
    runSearch(true);
  }
  return pv;
}

let _renderTimer = 0;
function renderPreview() {
  const pv = S.preview;
  if (!pv) return;
  const rt = ++S.rt;
  clearTimeout(_renderTimer);
  _renderTimer = setTimeout(() => pv.render(rt), 0);
}

// ------------------------------------------------------------------ linking source <-> preview
function lock() { S.lock = true; clearTimeout(S.timers.lock); S.timers.lock = setTimeout(() => { S.lock = false; }, 140); }

function blkAtLine(blks, line) {
  let lo = 0, hi = blks.length - 1, ans = -1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; if (+blks[mid].dataset.ls <= line) { ans = mid; lo = mid + 1; } else hi = mid - 1; }
  return ans < 0 ? null : blks[ans];
}
const topOf = (el2, scroller) => el2.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop;

function syncFromEditor() {
  if (S.lock || !S.preview) return;
  const pv = S.preview, ed = S.editor;
  if (!pv.blks) { if (pv.view === 'pdf') { lock(); syncProportion(ed.scroller, pv.scroller); } return; }
  const line = Math.floor(ed.scroller.scrollTop / LH()) + 1;
  const b = blkAtLine(pv.blks, line);
  if (!b) return;
  lock();
  pv.scroller.scrollTop = Math.max(0, topOf(b, pv.scroller) - 8);
}

function syncFromPreview() {
  if (S.lock || !S.editor || !S.preview || !S.preview.blks) return;
  const pv = S.preview;
  const top = pv.scroller.scrollTop + 12;
  let lo = 0, hi = pv.blks.length - 1, ans = 0;
  while (lo <= hi) { const mid = (lo + hi) >> 1; if (topOf(pv.blks[mid], pv.scroller) <= top) { ans = mid; lo = mid + 1; } else hi = mid - 1; }
  const b = pv.blks[ans];
  if (!b) return;
  lock();
  S.editor.scroller.scrollTop = Math.max(0, (+b.dataset.ls - 1) * LH());
}

function syncProportion(from, to) {
  if (S.lock) return;
  lock();
  const r = from.scrollTop / Math.max(1, from.scrollHeight - from.clientHeight);
  to.scrollTop = r * Math.max(0, to.scrollHeight - to.clientHeight);
}

function markPreviewBlock(b, scroll) {
  const pv = S.preview;
  if (!pv || !pv.blks) return;
  pv.blks.forEach(x => x.classList.remove('curbox'));
  if (!b || !S.split) return;
  const target = blkAtLine(pv.blks, b.line);
  if (!target) return;
  target.classList.add('curbox');
  if (scroll) {
    const r = target.getBoundingClientRect(), sr = pv.scroller.getBoundingClientRect();
    if (r.top < sr.top || r.bottom > sr.bottom) { lock(); pv.scroller.scrollTop = Math.max(0, topOf(target, pv.scroller) - 40); }
  }
}

function wirePreviewClicks(pv) {
  pv.root.addEventListener('click', (e) => {
    const b = e.target.closest('.blk');
    if (!b || !S.editor || !S.split) return;
    const ls = +b.dataset.ls;
    if (!ls) return;
    const { off } = lineOffset(ls);
    S.editor.ta.setSelectionRange(off, off);
    lock();
    S.editor.scroller.scrollTop = Math.max(0, (ls - 1) * LH() - S.editor.scroller.clientHeight / 3);
    onCaret(true);
  });
}

// ------------------------------------------------------------------ sidebar
function buildSide() {
  const side = $('rd-side');
  if (S.preview && S.preview.sideEl && (S.view === 'pdf' || S.view === 'excel')) { side.replaceChildren(S.preview.sideEl); return; }
  const st = S.analysis.stats;
  const go = (line) => { if (S.editor) goToLine(line); else scrollPreviewTo(line); };
  const mk = (cls, text, line, extra) => el('li', { class: cls, tabindex: '0', onclick: () => go(line), onkeydown: (e) => { if (e.key === 'Enter') go(line); } }, text, extra || null);
  const kids = [];
  kids.push(el('div', { class: 'side-h' }, 'Outline'));
  const minLevel = Math.min(6, ...st.headings.map(h => h.level));
  kids.push(st.headings.length
    ? el('ul', { class: 'outline' }, st.headings.slice(0, 600).map(h => mk(`l${Math.min(4, h.level - minLevel + 1)}`, h.text, h.line)))
    : el('p', { class: 'muted small' }, 'No headings'));
  kids.push(el('div', { class: 'side-h' }, `Tables (${st.tables.length})`));
  kids.push(el('ul', { class: 'sl' }, st.tables.slice(0, 200).map((t, i) => mk('', `Table ${i + 1} · ${t.rows} × ${t.cols}${t.merged ? ` · ${t.merged} merged` : ''}`, t.line))));
  if (st.footnotes.length) {
    kids.push(el('div', { class: 'side-h' }, `Footnotes (${st.footnotes.length})`));
    kids.push(el('ul', { class: 'sl' }, st.footnotes.slice(0, 100).map(id => mk('', `[^${id}]`, footLine(id)))));
  }
  if (st.images.length) {
    kids.push(el('div', { class: 'side-h' }, `Images (${st.images.length})`));
    kids.push(el('ul', { class: 'sl' }, st.images.slice(0, 100).map(im => mk('', `${im.alt || im.src || 'image'} `, im.line, el('span', { class: 'st-warn' }, '🟠')))));
  }
  side.replaceChildren(...kids);
  // mark the heading that contains the caret
  const cur = [...st.headings].reverse().find(h => h.line <= S.caret.line);
  if (cur) { const i = st.headings.indexOf(cur); const li = side.querySelectorAll('.outline li')[i]; if (li) li.classList.add('on'); }
}

function footLine(id) {
  const i = S.text.split('\n').findIndex(l => l.startsWith(`[^${id}]:`));
  return i < 0 ? 1 : i + 1;
}

function scrollPreviewTo(line) {
  const pv = S.preview;
  if (!pv || !pv.blks) return;
  const b = blkAtLine(pv.blks, line);
  if (b) pv.scroller.scrollTop = Math.max(0, topOf(b, pv.scroller) - 8);
}

// ------------------------------------------------------------------ checks drawer
function renderChecks() {
  const box = $('rd-checks');
  const items = S.analysis.items;
  box.hidden = !items.length || !S.checksOpen;
  if (box.hidden) return;
  const bad = items.filter(i => i.status === 'bad').length, warn = items.length - bad;
  box.replaceChildren(
    el('h4', {}, 'Checks before converting ',
      bad ? el('span', { class: 'chip2 bad' }, `🔴 ${bad}`) : null, warn ? el('span', { class: 'chip2 warn' }, `🟠 ${warn}`) : null,
      el('span', { class: 'sp' }), el('button', { class: 'icon-btn', onclick: () => { S.checksOpen = false; renderChecks(); } }, '▾ hide')),
    el('div', { class: 'rows' }, items.slice(0, 200).map(i => el('div', { class: `row ${i.status}` },
      el('span', {}, i.status === 'bad' ? '🔴' : '🟠'),
      el('span', { class: 't' }, `Line ${i.line}: ${i.text}`),
      el('button', { class: 'btn small', onclick: () => goToLine(i.line) }, 'Go to line')))));
}

// ------------------------------------------------------------------ search
function clearMarks(root) {
  if (!root) return;
  root.querySelectorAll('mark.sr').forEach(m => m.replaceWith(document.createTextNode(m.textContent)));
  root.normalize();
}

function markText(root, q, out) {
  const ql = q.toLowerCase();
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (w.nextNode()) nodes.push(w.currentNode);
  for (const n of nodes) {
    if (out.length >= MAX_MARKS) break;
    const text = n.nodeValue;
    if (!text || (n.parentElement && n.parentElement.closest('textarea,.ed-gutter,.ed-flags'))) continue;
    const lower = text.toLowerCase();
    let idx = lower.indexOf(ql);
    if (idx < 0) continue;
    const frag = document.createDocumentFragment();
    let last = 0;
    while (idx >= 0 && out.length < MAX_MARKS) {
      frag.append(text.slice(last, idx));
      const m = el('mark', { class: 'sr' }, text.slice(idx, idx + q.length));
      frag.append(m); out.push(m);
      last = idx + q.length;
      idx = lower.indexOf(ql, last);
    }
    frag.append(text.slice(last));
    n.replaceWith(frag);
  }
}

function runSearch(keep = false) {
  const roots = [S.editor && S.editor.pre, S.preview && S.preview.searchRoot].filter(Boolean);
  roots.forEach(clearMarks);
  S.marks = [];
  const q = S.q.trim();
  if (q) roots.forEach(r => markText(r, q, S.marks));
  if (!keep) S.qi = 0;
  S.qi = Math.max(0, Math.min(S.qi, S.marks.length - 1));
  const c = $('rd-qc');
  if (!q) c.textContent = '';
  else if (!S.marks.length) c.textContent = S.view === 'pdf' && !S.editor ? 'PDF pages are not searched' : 'No matches';
  else c.textContent = `${S.qi + 1} / ${S.marks.length >= MAX_MARKS ? MAX_MARKS + '+' : S.marks.length}`;
  setActive(false);
}

function setActive(scroll) {
  S.marks.forEach(m => m.classList.remove('cur'));
  const m = S.marks[S.qi];
  if (!m) return;
  m.classList.add('cur');
  if (scroll) m.scrollIntoView({ block: 'center', inline: 'nearest' });
}

function stepMatch(d) {
  if (!S.marks.length) return;
  S.qi = (S.qi + d + S.marks.length) % S.marks.length;
  $('rd-qc').textContent = `${S.qi + 1} / ${S.marks.length >= MAX_MARKS ? MAX_MARKS + '+' : S.marks.length}`;
  setActive(true);
}

// test hook
export const _state = S;

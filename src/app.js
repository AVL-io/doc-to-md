// User interface.
import { convertToMarkdown, convertFromMarkdown, DEFAULT_OPTIONS, TO_MD_EXT, FROM_MD_EXT, extOf, mdName } from './lib/convert.js';
import { setPdfWorkerSrc } from './lib/pdf-reader.js';

const VERSION = typeof __APP_VERSION__ !== 'undefined' ? __APP_VERSION__ : 'dev';
const $ = (id) => document.getElementById(id);
const state = { mode: 'to', files: [], results: [] };

// The single-file offline build provides the PDF worker as a blob URL.
setPdfWorkerSrc(window.__PDF_WORKER_URL__ || './pdf.worker.min.mjs');

// ------------------------------------------------------------ storage (per-viewer preferences only)
const PREF_KEY = 'doc-to-md:options';
function loadPrefs() {
  try { return { ...DEFAULT_OPTIONS, autoDownload: true, downloadReports: false, ...JSON.parse(localStorage.getItem(PREF_KEY) || '{}') }; }
  catch { return { ...DEFAULT_OPTIONS, autoDownload: true, downloadReports: false }; }
}
function savePrefs(p) { try { localStorage.setItem(PREF_KEY, JSON.stringify(p)); } catch { /* storage unavailable */ } }

// ------------------------------------------------------------ helpers
function fmtSize(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(0) + ' KB';
  return (n / 1048576).toFixed(1) + ' MB';
}
function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v === true) e.setAttribute(k, '');
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  for (const c of children.flat()) if (c != null) e.append(c);
  return e;
}
function download(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = el('a', { href: url, download: name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
const mdBlob = (text) => new Blob([text], { type: 'text/markdown;charset=utf-8' });
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ------------------------------------------------------------ mode / tabs
function setMode(mode) {
  state.mode = mode;
  state.files = [];
  $('tab-to').setAttribute('aria-selected', String(mode === 'to'));
  $('tab-from').setAttribute('aria-selected', String(mode === 'from'));
  $('dz-sub').textContent = mode === 'to' ? 'or click to choose · .pdf, .docx, .xlsx' : 'or click to choose · .md files';
  $('file-input').accept = (mode === 'to' ? TO_MD_EXT : FROM_MD_EXT).map(e => '.' + e).join(',');
  $('formats').hidden = mode !== 'from';
  $('convert-btn').textContent = mode === 'to' ? 'Convert to Markdown' : 'Convert Markdown';
  renderFiles();
}

// ------------------------------------------------------------ file intake
function addFiles(list) {
  const allowed = state.mode === 'to' ? TO_MD_EXT : FROM_MD_EXT;
  const rejected = [];
  for (const f of list) {
    if (!allowed.includes(extOf(f.name))) { rejected.push(f.name); continue; }
    if (!state.files.some(x => x.name === f.name && x.size === f.size)) state.files.push(f);
  }
  if (rejected.length) alert(`These files were not added (unsupported type):\n\n${rejected.join('\n')}\n\nSupported: ${allowed.map(e => '.' + e).join(', ')}`);
  renderFiles();
}

function renderFiles() {
  const ul = $('file-list');
  ul.replaceChildren(...state.files.map((f, i) => el('li', {},
    el('span', { class: 'type-tag' }, extOf(f.name)),
    el('span', { class: 'name', title: f.name }, f.name),
    el('span', { class: 'meta' }, fmtSize(f.size)),
    el('button', { class: 'icon-btn', title: 'Remove', 'aria-label': `Remove ${f.name}`, onclick: () => { state.files.splice(i, 1); renderFiles(); } }, '✕'),
  )));
  $('convert-btn').disabled = !state.files.length;
  $('clear-btn').hidden = !state.files.length;
}

function setupDropzone() {
  const dz = $('dropzone');
  const input = $('file-input');
  input.addEventListener('change', () => { addFiles([...input.files]); input.value = ''; });
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  ['dragenter', 'dragover'].forEach(t => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', (e) => addFiles([...(e.dataTransfer?.files || [])]));
  // prevent the browser from opening files dropped outside the zone
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
}

// ------------------------------------------------------------ options dialog
function askOptions() {
  const dlg = $('options-dialog');
  const form = $('options-form');
  const prefs = loadPrefs();
  const kinds = new Set(state.files.map(f => { const x = extOf(f.name); return x.startsWith('xls') ? 'xlsx' : x.startsWith('doc') ? 'docx' : x; }));
  for (const elx of form.querySelectorAll('[data-only]')) elx.hidden = !kinds.has(elx.dataset.only);
  for (const input of form.elements) {
    if (!input.name) continue;
    if (input.type === 'checkbox') input.checked = !!prefs[input.name];
    if (input.type === 'radio') input.checked = prefs[input.name] === input.value;
  }
  return new Promise((resolve) => {
    dlg.addEventListener('close', function onClose() {
      dlg.removeEventListener('close', onClose);
      if (dlg.returnValue !== 'ok') return resolve(null);
      const opts = { ...prefs };
      for (const input of form.elements) {
        if (!input.name) continue;
        if (input.type === 'checkbox') opts[input.name] = input.checked;
        if (input.type === 'radio' && input.checked) opts[input.name] = input.value;
      }
      savePrefs(opts);
      resolve(opts);
    });
    dlg.returnValue = '';
    dlg.showModal();
    $('proceed-btn').focus();
  });
}

// ------------------------------------------------------------ conversion
function setProgress(fraction, text) {
  $('progress').hidden = false;
  $('progress-bar').style.width = Math.round(Math.max(0, Math.min(1, fraction)) * 100) + '%';
  $('progress-text').textContent = text;
}

async function runToMarkdown() {
  const opts = await askOptions();
  if (!opts) return;
  const files = [...state.files];
  $('convert-btn').disabled = true;
  const n = files.length;
  for (let i = 0; i < n; i++) {
    const f = files[i];
    const base = i / n;
    setProgress(base, `Converting ${f.name} (${i + 1} of ${n})…`);
    let result;
    try {
      const r = await convertToMarkdown(f, opts, (p) => {
        if (p.stage === 'page') setProgress(base + (p.page / p.pages) * 0.8 / n, `Reading ${f.name}: page ${p.page} of ${p.pages}`);
        else if (p.stage === 'verifying') setProgress(base + 0.9 / n, `Verifying ${f.name}…`);
      });
      result = { file: f, ok: true, ...r, outName: mdName(f.name) };
      if (opts.autoDownload) { download(mdBlob(r.md), result.outName); await sleep(350); }
      if (opts.downloadReports) { download(mdBlob(r.reportMd), result.outName.replace(/\.md$/, '.verification.md')); await sleep(350); }
    } catch (e) {
      console.error(e);
      result = { file: f, ok: false, error: e && e.message ? e.message : String(e) };
    }
    state.results.unshift(result);
    renderResults();
  }
  setProgress(1, `Finished: ${n} file${n > 1 ? 's' : ''} processed.`);
  state.files = [];
  renderFiles();
}

async function runFromMarkdown() {
  const formats = [...document.querySelectorAll('input[name="fmt"]:checked')].map(i => i.value);
  if (!formats.length) { alert('Select at least one output format.'); return; }
  const files = [...state.files];
  $('convert-btn').disabled = true;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    setProgress(i / files.length, `Converting ${f.name} (${i + 1} of ${files.length})…`);
    let result;
    try {
      const outputs = await convertFromMarkdown(f, formats);
      for (const o of outputs) { download(o.blob, o.name); await sleep(350); }
      result = { file: f, ok: true, outputs, kind: 'md' };
    } catch (e) {
      console.error(e);
      result = { file: f, ok: false, error: e && e.message ? e.message : String(e) };
    }
    state.results.unshift(result);
    renderResults();
  }
  setProgress(1, 'Finished.');
  state.files = [];
  renderFiles();
}

// ------------------------------------------------------------ results
const VERDICT_LABEL = { complete: 'Complete', review: 'Review recommended', incomplete: 'Incomplete' };

function renderResults() {
  $('results').hidden = !state.results.length;
  $('result-list').replaceChildren(...state.results.map(resultCard));
}

function resultCard(r) {
  if (!r.ok) {
    return el('li', { class: 'result' },
      el('div', { class: 'result-top' }, el('span', { class: 'name' }, r.file.name), el('span', { class: 'verdict error' }, 'Failed')),
      el('p', { class: 'error-text' }, r.error));
  }
  if (r.kind === 'md') {
    return el('li', { class: 'result' },
      el('div', { class: 'result-top' }, el('span', { class: 'name' }, r.file.name), el('span', { class: 'verdict complete' }, 'Converted')),
      el('div', { class: 'result-actions' }, r.outputs.map(o => el('button', { class: 'btn small', onclick: () => download(o.blob, o.name) }, `Download ${o.name}`))));
  }
  const rep = r.report;
  return el('li', { class: 'result' },
    el('div', { class: 'result-top' },
      el('span', { class: 'name' }, `${r.file.name} → ${r.outName}`),
      el('span', { class: `verdict ${rep.verdict}` }, VERDICT_LABEL[rep.verdict])),
    el('div', { class: 'stats' },
      el('span', {}, 'Text coverage: ', el('b', {}, rep.words.coverage.toFixed(2) + '%')),
      el('span', {}, 'Numbers missing: ', el('b', {}, String(rep.numbers.missing.length))),
      el('span', {}, 'Words not found: ', el('b', {}, String(rep.words.missing))),
      el('span', {}, 'Removed by options: ', el('b', {}, rep.words.excluded + ' words')),
      r.kind === 'pdf' ? el('span', {}, 'Tables: ', el('b', {}, String(r.log.stats.tables))) : null),
    el('div', { class: 'result-actions' },
      el('button', { class: 'btn small primary', onclick: () => download(mdBlob(r.md), r.outName) }, 'Download .md'),
      el('button', { class: 'btn small', onclick: () => openViewer(r, 'report') }, 'Verification report'),
      el('button', { class: 'btn small', onclick: () => openViewer(r, 'review') }, 'Side-by-side review'),
      el('button', { class: 'btn small', onclick: () => download(mdBlob(r.reportMd), r.outName.replace(/\.md$/, '.verification.md')) }, 'Download report')));
}

async function downloadZip() {
  const { default: JSZip } = await import('jszip');
  const zip = new JSZip();
  const used = new Set();
  const unique = (n) => { let x = n, k = 2; while (used.has(x)) x = n.replace(/(\.[^.]+)$/, ` (${k++})$1`); used.add(x); return x; };
  for (const r of state.results) {
    if (!r.ok) continue;
    if (r.kind === 'md') for (const o of r.outputs) zip.file(unique(o.name), o.blob);
    else { zip.file(unique(r.outName), r.md); zip.file(unique('verification/' + r.outName.replace(/\.md$/, '.verification.md')), r.reportMd); }
  }
  download(await zip.generateAsync({ type: 'blob' }), `doc-to-md-${new Date().toISOString().slice(0, 10)}.zip`);
}

// ------------------------------------------------------------ viewer
async function openViewer(r, tab) {
  const dlg = $('viewer');
  $('viewer-title').textContent = r.file.name;
  const tabs = [
    ['review', 'Side-by-side review'],
    ['report', 'Verification report'],
    ['source', 'Markdown source'],
  ];
  const nav = $('viewer-tabs');
  const show = async (t) => {
    nav.querySelectorAll('button').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
    const body = $('viewer-body');
    body.replaceChildren(el('p', { class: 'muted' }, 'Loading…'));
    const { markdownToSafeHtml } = await import('./lib/html-writer.js');
    if (t === 'report') {
      const div = el('div', { class: 'md-body' });
      div.innerHTML = markdownToSafeHtml(r.reportMd);
      body.replaceChildren(div);
    } else if (t === 'source') {
      body.replaceChildren(el('pre', { class: 'source' }, r.md));
    } else {
      body.replaceChildren(await reviewPane(r, markdownToSafeHtml));
    }
  };
  nav.replaceChildren(...tabs.map(([k, label]) => el('button', { 'data-tab': k, onclick: () => show(k) }, label)));
  dlg.showModal();
  await show(tab);
}

async function reviewPane(r, toHtml) {
  const right = el('div', { class: 'md-body' });
  if (r.kind !== 'pdf') {
    right.innerHTML = toHtml(r.md);
    const note = el('div', {}, el('p', { class: 'muted' }, r.kind === 'docx'
      ? 'Word documents have no fixed pages, so the original cannot be displayed here. Compare the rendered output on the right with the document open in Word. Table positions are also compared cell by cell in the verification report.'
      : 'Excel cells are compared value by value in the verification report. The rendered output is shown on the right.'));
    return el('div', { class: 'split' }, note, right);
  }
  // PDF: original page on the left, the matching part of the output on the right
  const { renderPdfPage } = await import('./lib/pdf-reader.js');
  const canvas = el('canvas', { class: 'page' });
  const input = el('input', { type: 'number', min: '1', value: '1', 'aria-label': 'Page number' });
  const total = el('span', { class: 'muted' });
  const left = el('div', {}, el('div', { class: 'page-nav' },
    el('button', { class: 'btn small', onclick: () => go(+input.value - 1) }, '◀'), 'Page ', input, total,
    el('button', { class: 'btn small', onclick: () => go(+input.value + 1) }, '▶')), canvas);
  const sections = splitByPage(r.md);
  let pages = 1;
  async function go(p) {
    p = Math.max(1, Math.min(pages, p || 1));
    input.value = p;
    pages = await renderPdfPage(r.buf, p, canvas, Math.min(900, left.clientWidth - 24 || 700));
    total.textContent = ` of ${pages}`;
    input.max = String(pages);
    right.innerHTML = toHtml(sections.get(p) || (sections.size ? '*(No converted content was marked for this page. The table or paragraph may continue from the previous page.)*' : r.md));
  }
  input.addEventListener('change', () => go(+input.value));
  setTimeout(() => go(1), 0);
  return el('div', { class: 'split' }, left, right);
}

function splitByPage(md) {
  const map = new Map();
  const re = /<!-- Page (\d+) -->/g;
  let m, last = null, lastIdx = 0;
  while ((m = re.exec(md))) {
    if (last != null) map.set(last, md.slice(lastIdx, m.index));
    last = +m[1];
    lastIdx = m.index + m[0].length;
  }
  if (last != null) map.set(last, md.slice(lastIdx));
  return map;
}

// ------------------------------------------------------------ PWA: install, offline, updates
function setupPwa() {
  let deferred = null;
  window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferred = e; $('install-btn').hidden = false; });
  $('install-btn').addEventListener('click', async () => { if (!deferred) return; deferred.prompt(); await deferred.userChoice; deferred = null; $('install-btn').hidden = true; });
  window.addEventListener('appinstalled', () => { $('install-btn').hidden = true; });

  const status = $('offline-status');
  if (location.protocol === 'file:') {
    status.hidden = false; status.textContent = 'Offline copy';
    $('offline-link-wrap').hidden = true;
    return;
  }
  if (!('serviceWorker' in navigator)) return;
  // an update can only exist if this page was already controlled by an installed version
  const hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.register('sw.js').then((reg) => {
    const ready = () => { status.hidden = false; status.textContent = '✓ Works offline'; };
    if (reg.active) ready();
    navigator.serviceWorker.ready.then(ready);
    const showUpdate = (w) => {
      if (!hadController) { w.postMessage('skipWaiting'); return; }
      $('update-banner').hidden = false;
      $('update-btn').onclick = () => w.postMessage('skipWaiting');
    };
    if (reg.waiting && navigator.serviceWorker.controller) showUpdate(reg.waiting);
    reg.addEventListener('updatefound', () => {
      const w = reg.installing;
      w && w.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) showUpdate(w); });
    });
    setInterval(() => reg.update().catch(() => {}), 60 * 60 * 1000);
  }).catch(() => {});
  // reload only when an updated version takes over (not on the very first install)
  let reloading = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (hadController && !reloading) { reloading = true; location.reload(); } });
}

// ------------------------------------------------------------ init
function init() {
  $('version').textContent = 'Version ' + VERSION;
  $('tab-to').addEventListener('click', () => setMode('to'));
  $('tab-from').addEventListener('click', () => setMode('from'));
  $('convert-btn').addEventListener('click', () => (state.mode === 'to' ? runToMarkdown() : runFromMarkdown()));
  $('clear-btn').addEventListener('click', () => { state.files = []; renderFiles(); });
  $('zip-btn').addEventListener('click', downloadZip);
  $('viewer-close').addEventListener('click', () => $('viewer').close());
  setupDropzone();
  setMode('to');
  setupPwa();
  // "Open with" from the operating system once the app is installed
  if ('launchQueue' in window) {
    window.launchQueue.setConsumer(async (params) => {
      if (!params.files || !params.files.length) return;
      const files = await Promise.all(params.files.map(h => h.getFile()));
      setMode(files.every(f => FROM_MD_EXT.includes(extOf(f.name))) ? 'from' : 'to');
      addFiles(files);
    });
  }
  // test hook (used by the automated browser tests)
  window.__docToMd = { convertToMarkdown, convertFromMarkdown, state };
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();

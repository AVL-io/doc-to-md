// Local OCR (English) with Tesseract.js. The engine files are part of the app
// (served from the same address, folder "ocr/"), so no document and no request
// ever leaves the computer. The engine runs in Web Workers.
import Tesseract from 'tesseract.js/dist/tesseract.esm.min.js';

const { createWorker } = Tesseract;
const LSTM_ONLY = 1;
const IDLE_MS = 45000;

const url = (p) => new URL(p, document.baseURI).href;
const PATHS = () => ({ workerPath: url('ocr/worker.min.js'), corePath: url('ocr/core'), langPath: url('ocr/lang') });

let engine = null, starting = null, idleTimer = 0, users = 0;

export { ocrSupported as ocrAvailable } from './scan.js';

function workerCount() {
  const c = navigator.hardwareConcurrency || 2;
  return Math.max(1, Math.min(2, Math.floor(c / 2)));
}

async function start(dpi, onStatus) {
  const n = workerCount();
  const workers = [];
  let progress = 0;
  const logger = (m) => {
    if (!onStatus || !m || typeof m.progress !== 'number') return;
    if (/loading|initializing/.test(m.status || '')) { progress = Math.max(progress, m.progress); onStatus(`Preparing the OCR engine (${m.status})…`); }
  };
  const created = await Promise.all(Array.from({ length: n }, () => createWorker('eng', LSTM_ONLY, {
    ...PATHS(), workerBlobURL: false, cacheMethod: 'none', gzip: true, logger,
  }, { tessedit_pageseg_mode: '3', preserve_interword_spaces: '1', user_defined_dpi: String(dpi) })));
  workers.push(...created);
  const idle = [...workers];
  const waiters = [];
  const stoppedError = () => { const e = new Error('The OCR engine was stopped.'); e.name = 'AbortError'; return e; };
  const eng = {
    dpi, workers, stopped: false,
    async setDpi(d) {
      if (d === this.dpi) return;
      if (this.stopped) throw stoppedError();
      this.dpi = d;
      await Promise.all(workers.map(w => w.setParameters({ user_defined_dpi: String(d) })));
    },
    /** Run fn(worker) on the next free worker (at most one job per worker at a time). */
    async run(fn) {
      if (this.stopped) throw stoppedError();
      const w = idle.pop() || await new Promise((resolve, reject) => waiters.push({ resolve, reject }));
      if (this.stopped) throw stoppedError();
      try { return await fn(w); }
      finally { const next = waiters.shift(); if (next) next.resolve(w); else idle.push(w); }
    },
    async stop() {
      this.stopped = true;
      for (const x of waiters.splice(0)) x.reject(stoppedError());
      await Promise.all(workers.map(w => w.terminate().catch(() => {})));
    },
  };
  return eng;
}

/** Get the shared engine (started on first use, stopped after a period of inactivity). */
export async function acquireOcr({ dpi = 200, onStatus } = {}) {
  clearTimeout(idleTimer);
  users++;
  try {
    if (engine && engine.stopped) engine = null;
    if (!engine) {
      if (!starting) starting = start(dpi, onStatus).finally(() => { starting = null; });
      engine = await starting;
    }
    await engine.setDpi(dpi);
    return engine;
  } catch (e) {
    users--;
    throw new Error('The OCR engine could not be started: ' + (e && e.message ? e.message : e));
  }
}

export function releaseOcr() {
  users = Math.max(0, users - 1);
  if (users === 0) {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(async () => { const e = engine; engine = null; if (e) await e.stop(); }, IDLE_MS);
  }
}

/** Stop the engine at once (used when a conversion is cancelled). */
export async function abortOcr() {
  clearTimeout(idleTimer);
  const e = engine;
  engine = null; users = 0;
  if (e) await e.stop();
}

/** Download and cache the OCR files now (so that OCR also works offline later). */
export async function warmUpOcr(onStatus) {
  const e = await acquireOcr({ dpi: 200, onStatus });
  releaseOcr();
  return !!e;
}

/** Recognise one page image. Returns lines of words with boxes in pixels. */
export async function recogniseCanvas(worker, canvas) {
  const res = await worker.recognize(canvas, {}, { blocks: true });
  const data = res.data || {};
  const lines = [];
  for (const b of data.blocks || []) {
    for (const p of b.paragraphs || []) {
      for (const l of p.lines || []) {
        const words = (l.words || []).filter(w => w.text && w.text.trim()).map(w => ({ text: w.text.trim(), conf: w.confidence, x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1 }));
        if (words.length) lines.push({ words, bbox: l.bbox, baseline: l.baseline, conf: l.confidence, row: l.rowAttributes });
      }
    }
  }
  return { lines, text: data.text || '', confidence: data.confidence };
}

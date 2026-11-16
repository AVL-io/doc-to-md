// Image analysis for scanned pages: finds the ruling lines of tables in the page
// image (so that merged cells can be rebuilt as for digital PDFs) and measures how
// heavy the ink of each word is (to tell bold text from regular text).

/** OCR needs web workers and WebAssembly, and the engine files served next to the app (not a file:// copy). */
export const ocrSupported = () => typeof Worker !== 'undefined' && typeof WebAssembly !== 'undefined' && location.protocol !== 'file:';

function otsu(hist, total) {
  let sum = 0;
  for (let i = 0; i < 256; i++) sum += i * hist[i];
  let sumB = 0, wB = 0, max = 0, thr = 128;
  for (let t = 0; t < 256; t++) {
    wB += hist[t];
    if (!wB) continue;
    const wF = total - wB;
    if (!wF) break;
    sumB += t * hist[t];
    const mB = sumB / wB, mF = (sum - sumB) / wF;
    const v = wB * wF * (mB - mF) * (mB - mF);
    if (v > max) { max = v; thr = t; }
  }
  return thr;
}

/**
 * @param canvas  the rendered page
 * @param scale   pixels per PDF unit
 * @returns {{ rules: Array, ink(x0,y0,x1,y1): number }}  rules in PDF units: {o:'h',y,a,b} / {o:'v',x,a,b}
 */
export function analyseScan(canvas, scale) {
  const w = canvas.width, h = canvas.height;
  const img = canvas.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
  const hist = new Uint32Array(256);
  const lum = new Uint8Array(w * h);
  for (let i = 0, p = 0; i < img.length; i += 4, p++) {
    const v = (img[i] * 299 + img[i + 1] * 587 + img[i + 2] * 114) / 1000 | 0;
    lum[p] = v; hist[v]++;
  }
  const thr = Math.max(80, Math.min(200, otsu(hist, w * h)));
  const dark = new Uint8Array(w * h);
  for (let p = 0; p < dark.length; p++) dark[p] = lum[p] <= thr ? 1 : 0;

  // runs of ink with small gaps: long horizontal and vertical runs are table borders
  const minH = Math.round(w * 0.035), minV = Math.round(h * 0.02);
  const gapTol = Math.max(2, Math.round(scale * 0.8));
  const hSegs = [], vSegs = [];
  const vRun = new Int32Array(w), vGap = new Int32Array(w), vStart = new Int32Array(w);
  for (let y = 0; y < h; y++) {
    let run = 0, gap = 0, start = 0;
    const row = y * w;
    for (let x = 0; x < w; x++) {
      if (dark[row + x]) {
        if (run === 0) start = x;
        run += gap + 1; gap = 0;
        if (vRun[x] === 0) vStart[x] = y;
        vRun[x] += vGap[x] + 1; vGap[x] = 0;
      } else {
        if (run > 0 && ++gap > gapTol) { if (run >= minH) hSegs.push([y, start, start + run]); run = 0; gap = 0; }
        if (vRun[x] > 0 && ++vGap[x] > gapTol) { if (vRun[x] >= minV) vSegs.push([x, vStart[x], vStart[x] + vRun[x]]); vRun[x] = 0; vGap[x] = 0; }
      }
    }
    if (run >= minH) hSegs.push([y, start, start + run]);
  }
  for (let x = 0; x < w; x++) if (vRun[x] >= minV) vSegs.push([x, vStart[x], vStart[x] + vRun[x]]);

  const edge = 0.012;
  const rules = [];
  for (const [y, a, b] of hSegs) {
    if (y < h * edge || y > h * (1 - edge)) continue;
    rules.push({ o: 'h', y: (y + 0.5) / scale, a: a / scale, b: b / scale });
  }
  for (const [x, a, b] of vSegs) {
    if (x < w * edge || x > w * (1 - edge)) continue;
    rules.push({ o: 'v', x: (x + 0.5) / scale, a: a / scale, b: b / scale });
  }

  return {
    rules,
    /** typical stroke width (pixels): median length of the horizontal ink runs inside a box */
    stroke(x0, y0, x1, y1) {
      x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
      x1 = Math.min(w, Math.ceil(x1)); y1 = Math.min(h, Math.ceil(y1));
      const counts = new Uint32Array(41);
      let total = 0;
      for (let y = y0; y < y1; y++) {
        const row = y * w;
        let run = 0;
        for (let x = x0; x <= x1; x++) {
          if (x < x1 && dark[row + x]) run++;
          else if (run) { counts[Math.min(40, run)]++; total++; run = 0; }
        }
      }
      if (!total) return 0;
      let acc = 0;
      for (let k = 1; k <= 40; k++) { acc += counts[k]; if (acc >= total / 2) return k; }
      return 40;
    },
    /** share of ink pixels in a box (pixels) */
    ink(x0, y0, x1, y1) {
      x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
      x1 = Math.min(w, Math.ceil(x1)); y1 = Math.min(h, Math.ceil(y1));
      let n = 0;
      for (let y = y0; y < y1; y++) { const row = y * w; for (let x = x0; x < x1; x++) n += dark[row + x]; }
      const area = Math.max(1, (x1 - x0) * (y1 - y0));
      return n / area;
    },
  };
}

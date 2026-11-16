// Browser tests for scanned pages and local OCR.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from '../tools/playwright.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const dist = path.join(root, 'dist');
const fx = (n) => path.join(root, 'tests/fixtures', n);
const outDir = path.join(root, 'test-output');
fs.mkdirSync(outDir, { recursive: true });
const types = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.webmanifest': 'application/manifest+json', '.gz': 'application/gzip' };
let requests = [];
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  requests.push(p);
  if (p.endsWith('/')) p += 'index.html';
  const f = path.join(dist, p);
  if (!f.startsWith(dist) || !fs.existsSync(f)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': types[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise(r => server.listen(0, r));
const base = `http://localhost:${server.address().port}/`;

let failures = 0;
const ok = (c, m) => { if (c) console.log('  ✓ ' + m); else { failures++; console.log('  ✗ FAIL: ' + m); } };

const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1400, height: 900 } });
const page = await ctx.newPage();
const errors = [];
const external = [];
let step = 'start';
page.on('pageerror', e => errors.push(`[${step}] ${e.message} ${(e.stack || '').split('\n').slice(1, 4).join(' | ')}`));
page.on('console', m => { if (m.type() === 'error' && !/Detected \d+ diacritics/.test(m.text())) errors.push(m.text()); });
page.on('request', r => { const u = new URL(r.url()); if (u.origin !== base.replace(/\/$/, '') && !/^(data|blob):/.test(r.url())) external.push(r.url()); });
await page.goto(base);
await page.waitForFunction(() => window.__docToMd);

const convert = (file, opts = {}) => page.evaluate(async ({ b64, name, opts }) => {
  const f = new File([Uint8Array.from(atob(b64), c => c.charCodeAt(0))], name);
  const t0 = performance.now();
  const prog = [];
  try {
    const r = await window.__docToMd.convertToMarkdown(f, opts, (p) => { if (p.ocr) prog.push(p.ocr.done + '/' + p.ocr.total); });
    return { ok: true, md: r.md, report: r.report, reportMd: r.reportMd, ocr: r.log.ocr, warnings: r.log.warnings, ms: performance.now() - t0, prog };
  } catch (e) { return { ok: false, error: String(e && e.message || e), name: e && e.name }; }
}, { b64: fs.readFileSync(file).toString('base64'), name: path.basename(file), opts });

console.log('Scanned document (clean)'); step = 'Scanned document (clean)';
const a = await convert(fx('synthetic-scan.pdf'));
ok(a.ok, 'conversion succeeds' + (a.ok ? '' : ': ' + a.error));
fs.writeFileSync(path.join(outDir, 'scan.md'), a.md);
ok(a.ocr.pages.length === 2, 'both pages were recognised by OCR');
ok(/# Environmental Impact Assessment/.test(a.md) && /# 1 Introduction/.test(a.md), 'title and numbered heading found');
ok(/## 1\.1 Scope/.test(a.md), 'sub-heading level recovered from numbering');
ok(/55\.5 dB\(A\)/.test(a.md), 'text and numbers read (55.5 dB(A))');
ok(/<td>48\.2<\/td>/.test(a.md) && /<td>61\.7<\/td>/.test(a.md) && /<td>21\.4<\/td>/.test(a.md), 'table values read correctly');
ok(/<td rowspan="2">NO2<\/td>/.test(a.md) && /<td rowspan="2">200<\/td>/.test(a.md), 'merged cells recovered from the page image');
ok(/<th colspan="3">Section A/.test(a.md), 'column span recovered');
ok(/^- Air quality monitoring/m.test(a.md) && /^1\. Baseline survey/m.test(a.md), 'bullet and numbered lists recovered');
ok(!/TEST REPORT HEADER/.test(a.md) && !/Footer text/.test(a.md), 'headers and footers removed');
ok(/Total cost was 1,250\.75 AED/.test(a.md), 'page 2 read');
ok(a.report.verdict === 'review', `verdict is “review”, never “complete” (got ${a.report.verdict})`);
ok(a.report.reviewItems.some(i => /Scanned page recognised by OCR/.test(i.item)), 'checklist lists the OCR pages');
ok(a.report.reviewItems.some(i => /Table recognised from a scanned page/.test(i.reason)), 'checklist flags the OCR tables');
ok(a.ocr.pages.every(p => p.conf > 70), 'confidence reported for each page');
ok(/cannot detect misread/.test(a.reportMd), 'report explains the limits of the text check');
console.log(`  ${(a.ms / 1000).toFixed(1)} s for 2 pages; confidence ${a.ocr.pages.map(p => Math.round(p.conf) + '%').join(', ')}`);

console.log('Mixed document (digital page + scanned page)'); step = 'Mixed document (digital page + scanned page)';
const m = await convert(fx('synthetic-mixed.pdf'));
ok(m.ok && m.ocr.pages.length === 1 && m.ocr.pages[0].page === 2, 'only page 2 is recognised by OCR');
ok(/Environmental Impact Assessment/.test(m.md) && /Total cost was 1,250\.75 AED/.test(m.md), 'digital page 1 and scanned page 2 are both converted');
ok(/<!-- Page 2 -->/.test(m.md), 'page markers kept');

console.log('Scanner-like document (rotated, noisy, blurred)'); step = 'Scanner-like document (rotated, noisy, blurred)';
const s = await convert(fx('synthetic-scan-skewed.pdf'));
ok(s.ok && /Introduction/.test(s.md) && /55\.5 dB\(A\)/.test(s.md), 'text is still read (' + (s.ok ? s.ocr.pages.map(p => Math.round(p.conf) + '%').join(', ') : s.error) + ')');
ok(s.ok && s.report.verdict === 'review', 'still reported as “review”');
fs.writeFileSync(path.join(outDir, 'scan-skewed.md'), s.md || '');

console.log('OCR switched off'); step = 'OCR switched off';
const off = await convert(fx('synthetic-scan.pdf'), { ocr: false });
ok(off.ok && !off.ocr.pages.length, 'no OCR was run');
ok(off.ok && off.report.verdict === 'incomplete', 'document is reported as incomplete');
ok(off.ok && off.warnings.some(w => /OCR is switched off/.test(w)), 'warning explains that OCR is switched off');

console.log('Digital PDFs are not affected'); step = 'Digital PDFs are not affected';
const d = await convert(fx('synthetic-report.pdf'));
ok(d.ok && d.ocr.pages.length === 0 && d.report.verdict === 'complete', 'digital PDF: no OCR, verdict “complete”');

console.log('Cancel'); step = 'Cancel';
const c = await page.evaluate(async (b64) => {
  const f = new File([Uint8Array.from(atob(b64), x => x.charCodeAt(0))], 'x.pdf');
  const ac = new AbortController();
  setTimeout(() => ac.abort(), 600);
  try { await window.__docToMd.convertToMarkdown(f, { signal: ac.signal }, () => {}); return 'finished'; } catch (e) { return e.name; }
}, fs.readFileSync(fx('synthetic-scan.pdf')).toString('base64'));
ok(c === 'AbortError', `cancelling stops the conversion (${c})`);
const after = await convert(fx('synthetic-scan.pdf'));
ok(after.ok && after.ocr.pages.length === 2, 'conversion works again after a cancel');

console.log('Privacy'); step = 'Privacy';
ok(!external.length, 'no request left this address' + (external.length ? ': ' + external.join(', ') : ''));
ok(requests.some(r => r.includes('/ocr/lang/eng.traineddata.gz')), 'the English model was loaded from this app');

console.log('User interface'); step = 'User interface';
await page.click('#tab-to');
await page.setInputFiles('#file-input', fx('synthetic-mixed.pdf'));
await page.click('#convert-btn');
await page.waitForSelector('#options-dialog[open]');
ok(await page.isVisible('#opt-ocr') && await page.isChecked('#opt-ocr'), 'options dialog offers OCR (on by default)');
ok(await page.isChecked('input[name="ocrDpi"][value="300"]'), 'high quality (300 dpi) is the default');
await page.screenshot({ path: path.join(outDir, 'ocr-options.png') });
await page.click('#proceed-btn');
await page.waitForSelector('.result', { timeout: 60000 });
ok(/Scanned pages \(OCR\)/.test(await page.textContent('.result .stats')), 'result card shows the scanned-page summary');
ok(/review/i.test(await page.textContent('.result .verdict')), 'result card says “Review recommended”');
await page.click('.result button:has-text("Side-by-side review")');
await page.waitForSelector('canvas.page');
await page.fill('input[type=number]', '2');
await page.press('input[type=number]', 'Enter');
await page.waitForTimeout(1500);
await page.screenshot({ path: path.join(outDir, 'ocr-review.png') });
await page.click('#viewer-close');
// checklist entry with outlined low-confidence words
await page.click('.result button:has-text("Review checklist")');
await page.waitForSelector('table.checklist');
ok(await page.locator('table.checklist tr', { hasText: 'Scanned page recognised by OCR' }).count() >= 1, 'checklist shows the OCR page');
await page.click('#viewer-close');

await page.setInputFiles('#file-input', fx('synthetic-scan.pdf'));
await page.click('#convert-btn');
await page.waitForSelector('#options-dialog[open]');
await page.click('#ocr-prep');
await page.waitForFunction(() => /Ready/.test(document.getElementById('ocr-prep-status').textContent), null, { timeout: 60000 });
ok(true, '“Download OCR files now” prepares the engine');
await page.click('#options-form button[value="cancel"]');

console.log('Offline use after the first visit (service worker)'); step = 'Offline use after the first visit (service worker)';
const pw = await ctx.newPage();
const pwErrors = [];
pw.on('pageerror', e => pwErrors.push(e.message));
await pw.goto(base);
await pw.waitForFunction(async () => (await navigator.serviceWorker.ready) && !!navigator.serviceWorker.controller || (await caches.keys()).length > 0, null, { timeout: 30000 });
await pw.reload();
await pw.waitForFunction(() => window.__docToMd && navigator.serviceWorker.controller);
const run = (pg) => pg.evaluate(async (b64) => {
  const f = new File([Uint8Array.from(atob(b64), c => c.charCodeAt(0))], 's.pdf');
  try { const r = await window.__docToMd.convertToMarkdown(f, {}, () => {}); return { ok: true, n: r.log.ocr.pages.length, has: /Introduction/.test(r.md) }; } catch (e) { return { ok: false, error: String(e && e.message || e) }; }
}, fs.readFileSync(fx('synthetic-scan.pdf')).toString('base64'));
const online = await run(pw);
ok(online.ok && online.n === 2, 'OCR works online through the service worker');
const cached = await pw.evaluate(async () => { const c = await caches.open('doc-to-md-ocr-eng-1'); return (await c.keys()).map(r => new URL(r.url).pathname.split('/ocr/')[1]); });
ok(cached.length >= 3 && cached.some(x => x.startsWith('worker')) && cached.some(x => x.includes('traineddata')), `OCR files were cached on first use (${cached.join(', ')})`);
await ctx.setOffline(true);
await pw.reload();
await pw.waitForFunction(() => window.__docToMd, null, { timeout: 20000 });
const offline = await run(pw);
ok(offline.ok && offline.n === 2 && offline.has, 'OCR also works with no network' + (offline.ok ? '' : ': ' + offline.error));
await ctx.setOffline(false);
ok(!pwErrors.length, 'no page errors while offline' + (pwErrors.length ? ': ' + pwErrors.join(' | ') : ''));

console.log('Single-file offline copy'); step = 'Single-file offline copy';
const sf = await ctx.newPage();
await sf.goto('file://' + path.join(dist, 'doc-to-md-offline.html'));
await sf.waitForFunction(() => window.__docToMd, null, { timeout: 30000 });
await sf.setInputFiles('#file-input', fx('synthetic-scan.pdf'));
await sf.click('#convert-btn');
await sf.waitForSelector('#options-dialog[open]');
ok(await sf.isDisabled('#opt-ocr'), 'OCR option is disabled in the single-file copy');
ok(/not available in the single-file/.test(await sf.textContent('#ocr-note')), 'the dialog explains why');
await sf.click('#proceed-btn');
await sf.waitForSelector('.result', { timeout: 60000 });
ok(/not available in this copy/.test(await sf.textContent('.result')) || /incomplete/i.test(await sf.textContent('.result .verdict')), 'scanned pages are reported as not converted');
await sf.close();

ok(!errors.length, 'no page errors' + (errors.length ? ': ' + errors.join(' | ') : ''));
await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) failed` : '\nAll OCR checks passed');
process.exit(failures ? 1 : 0);

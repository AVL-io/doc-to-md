// Browser tests: serves dist/, loads the app in headless Chromium (with its real
// Content Security Policy) and converts every file in tests/fixtures (committed,
// synthetic) and tests/private (git-ignored, for real documents).
//
//   npm test                 build + run
//   node tests/run-tests.mjs run only (after npm run build)
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from '../tools/playwright.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const dist = path.join(root, 'dist');
const outDir = path.join(root, 'test-output');
fs.mkdirSync(outDir, { recursive: true });
const expectations = JSON.parse(fs.readFileSync(path.join(root, 'tests/expectations.json'), 'utf8'));

const types = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.webmanifest': 'application/manifest+json' };
const server = http.createServer((req, res) => {
  let p = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (p.endsWith('/')) p += 'index.html';
  const f = path.join(dist, p);
  if (!f.startsWith(dist) || !fs.existsSync(f)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': types[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise(r => server.listen(0, r));
const base = `http://localhost:${server.address().port}/`;

const browser = await chromium.launch();
const page = await browser.newPage();
const problems = [];
page.on('pageerror', e => problems.push('page error: ' + e.message));
page.on('console', m => { if (m.type() === 'error') problems.push('console: ' + m.text()); });
await page.goto(base);
await page.waitForFunction(() => window.__docToMd);

const dirs = ['tests/fixtures', 'tests/private'].map(d => path.join(root, d)).filter(d => fs.existsSync(d));
const inputs = dirs.flatMap(d => fs.readdirSync(d).filter(f => /\.(docx|xlsx|xlsm|pdf)$/i.test(f)).map(f => path.join(d, f)));

let failures = 0;
const check = (cond, msg) => { if (!cond) { failures++; console.log('    FAIL: ' + msg); } };

for (const file of inputs) {
  const name = path.basename(file);
  const b64 = fs.readFileSync(file).toString('base64');
  for (const tableMode of ['html', 'flatten']) {
    const t0 = Date.now();
    const res = await page.evaluate(async ({ name, b64, tableMode }) => {
      const bin = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
      const f = new File([bin], name);
      try {
        const r = await window.__docToMd.convertToMarkdown(f, { tableMode });
        // round trip: Markdown -> Word / Excel / PDF / HTML
        const md = new File([r.md], name + '.md');
        const outs = await window.__docToMd.convertFromMarkdown(md, ['docx', 'xlsx', 'pdf', 'html']);
        const sizes = Object.fromEntries(outs.map(o => [o.name.split('.').pop(), o.blob.size]));
        // re-read the generated Word file and compare its tables with the first conversion
        const docx = outs.find(o => o.name.endsWith('.docx'));
        const again = await window.__docToMd.convertToMarkdown(new File([docx.blob], 'roundtrip.docx'), { tableMode });
        return { ok: true, md: r.md, reportMd: r.reportMd, report: r.report, sizes, roundTripMd: again.md };
      } catch (e) {
        return { ok: false, error: String(e && e.stack || e) };
      }
    }, { name, b64, tableMode });
    const ms = Date.now() - t0;
    const stem = `${name}.${tableMode}`;
    if (!res.ok) { failures++; console.log(`✗ ${stem}: ${res.error}`); continue; }
    fs.writeFileSync(path.join(outDir, `${stem}.md`), res.md);
    fs.writeFileSync(path.join(outDir, `${stem}.verification.md`), res.reportMd);
    fs.writeFileSync(path.join(outDir, `${stem}.roundtrip.md`), res.roundTripMd);
    const r = res.report;
    console.log(`✓ ${stem}: ${r.verdict}, coverage ${r.words.coverage.toFixed(2)}%, words missing ${r.words.missing}, numbers missing ${r.numbers.missing.length}, ${ms} ms; outputs ${JSON.stringify(res.sizes)}`);
    for (const s of r.structure) if (!s.ok) console.log(`    structure: ${s.label} source=${s.source} output=${s.output}`);
    for (const t of r.tableChecks) if (!t.ok) console.log(`    table ${t.index}: ${t.summary}`);
    for (const p of r.passages.slice(0, 5)) console.log(`    missing @ ${p.loc}: ${p.text.slice(0, 100)}`);

    const exp = expectations[name];
    if (exp) {
      check(r.words.missing === 0, `${stem}: expected no missing words`);
      check(r.numbers.missing.length === 0, `${stem}: expected no missing numbers`);
      for (const s of exp.contains?.[tableMode] || []) check(res.md.includes(s), `${stem}: output should contain ${JSON.stringify(s)}`);
      for (const s of exp.contains?.any || []) check(res.md.includes(s), `${stem}: output should contain ${JSON.stringify(s)}`);
      for (const s of exp.notContains || []) check(!res.md.includes(s), `${stem}: output should not contain ${JSON.stringify(s)}`);
      if (exp.verdict) check(r.verdict === exp.verdict, `${stem}: verdict ${r.verdict}, expected ${exp.verdict}`);
    }
    for (const [k, v] of Object.entries(res.sizes)) check(v > 200, `${stem}: ${k} output is empty`);
  }
}

// offline single-file build: must load from file:// and convert
const off = await browser.newPage();
const offProblems = [];
off.on('pageerror', e => offProblems.push(e.message));
off.on('console', m => { if (m.type() === 'error') offProblems.push(m.text()); });
await off.goto('file://' + path.join(dist, 'doc-to-md-offline.html'));
await off.waitForFunction(() => window.__docToMd, null, { timeout: 20000 });
const pdfFixture = inputs.find(f => f.endsWith('.pdf'));
if (pdfFixture) {
  const ok = await off.evaluate(async (b64) => {
    const f = new File([Uint8Array.from(atob(b64), c => c.charCodeAt(0))], 'x.pdf');
    const r = await window.__docToMd.convertToMarkdown(f, {});
    return r.md.length > 50;
  }, fs.readFileSync(pdfFixture).toString('base64'));
  check(ok, 'offline single-file build converts a PDF');
  console.log(`✓ offline single-file build loads from file:// and converts a PDF`);
}
check(!offProblems.length, 'offline build errors: ' + offProblems.join(' | '));
check(!problems.length, 'page errors: ' + problems.join(' | '));

await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) failed` : '\nAll checks passed');
process.exit(failures ? 1 : 0);

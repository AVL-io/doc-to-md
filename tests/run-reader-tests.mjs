// Browser tests for the Markdown reader (opens the built app in headless Chromium).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from '../tools/playwright.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const dist = path.join(root, 'dist');
const sample = path.join(root, 'tests/fixtures/reader-sample.md');
const out = path.join(root, 'test-output');
fs.mkdirSync(out, { recursive: true });
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

let failures = 0;
const ok = (cond, msg) => { if (cond) console.log('  ✓ ' + msg); else { failures++; console.log('  ✗ FAIL: ' + msg); } };

const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1500, height: 900 } });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', e => errors.push(e.message));
page.on('console', m => m.type() === 'error' && errors.push(m.text()));
let nativeDialogs = 0;
page.on('dialog', d => { nativeDialogs++; d.dismiss(); });
await page.goto(base);
await page.waitForFunction(() => window.__docToMd);

// ---------------------------------------------------------------- checks (analysis)
console.log('Checks');
const text = fs.readFileSync(sample, 'utf8');
const a = await page.evaluate((t) => window.__docToMd.analyze(t), text);
ok(a.items.length === 2, `sample has 2 checks (got ${a.items.length})`);
ok(a.items[0].status === 'bad' && /\[\^2\]/.test(a.items[0].text), 'undefined footnote is serious');
ok(a.items[1].status === 'warn' && /Image/.test(a.items[1].text), 'image is a note');
ok(a.stats.tables.length === 1 && a.stats.tables[0].merged === 3 && a.stats.tables[0].cols === 4, 'table: 4 columns, 3 merged cells');
ok(a.stats.headings.length === 4 && a.stats.footnotes.length === 1, '4 headings, 1 footnote');
const edge = await page.evaluate((t) => window.__docToMd.analyze(t), [
  '# One', '### Three', '', '<table>', '<tr><td>x</td></tr>', '',
  '| a | b |', '| --- | --- |', '| 1 |', '', '[link](other.md)', '<div>text</div>', '[^9]: unused', '',
  '```', 'code never closed',
].join('\n'));
const msgs = edge.items.map(i => i.text).join(' | ');
ok(/jumps from 1 to 3/.test(msgs), 'skipped heading level found');
ok(/<table> is never closed/.test(msgs), 'unclosed <table> found');
ok(/never closed/.test(msgs), 'unclosed code block found');
ok(/has 1 cell but the header has 2/.test(msgs), 'short table row found');
ok(/other\.md/.test(msgs), 'relative link found');
ok(/<div>/.test(msgs), 'unsupported HTML tag found');
ok(/\[\^9\].*never used/.test(msgs), 'unused footnote found');
const empty = await page.evaluate((t) => window.__docToMd.analyze(t), '');
ok(empty.items.length === 1 && /empty/.test(empty.items[0].text), 'empty file is reported');

// ---------------------------------------------------------------- open from the queue
console.log('Queue and reader');
await page.click('#tab-from');
await page.setInputFiles('#file-input', sample);
ok(await page.locator('.file-list .read-btn').count() === 1, 'queue row has a Read button');
await page.click('.file-list .read-btn');
await page.waitForSelector('#reader[open]');
ok((await page.textContent('#rd-name')) === 'reader-sample.md', 'reader shows the file name');
ok(/29 lines/.test(await page.textContent('#rd-meta')) && /1 table \(3 merged cells\)/.test(await page.textContent('#rd-meta')), 'meta line: 29 lines, 1 table (3 merged cells)');
ok((await page.inputValue('.ed-ta')) === text, 'Markdown view shows the file text');
ok(await page.locator('#rd-side .outline li').count() === 4, 'outline lists 4 headings');
ok(await page.locator('#rd-checks .row').count() === 2, 'checks drawer lists 2 items');
ok(await page.locator('.ed-hl .t-tag').count() > 5, 'syntax colours are applied');
await page.screenshot({ path: path.join(out, 'reader-md.png') });

// ---------------------------------------------------------------- views
console.log('Views');
await page.click('#rd-views [data-view="word"]');
await page.waitForSelector('.paper.word table');
ok(await page.locator('.paper.word td[rowspan="2"]').count() === 3, 'Word view: 3 merged cells');
ok(/Image: Site plan/.test(await page.textContent('.paper.word')), 'Word view: image placeholder');
ok(await page.locator('.paper.word sup').count() >= 2, 'Word view: footnote markers');
ok(/Approximate/.test(await page.textContent('.fid')), 'Word view is labelled approximate');
ok(/Create Word/.test(await page.textContent('#rd-create')), 'Create button names the format');
await page.screenshot({ path: path.join(out, 'reader-word.png') });

await page.click('#rd-views [data-view="pdf"]');
await page.waitForSelector('.pdfpage canvas', { timeout: 30000 });
ok(await page.locator('.pdfpage').count() >= 1, 'PDF view: pages created');
ok(await page.locator('.thumb-btn').count() === await page.locator('.pdfpage').count(), 'PDF view: one thumbnail per page');
ok(/Page 1 of/.test(await page.textContent('#rd-status-l')), 'PDF view: page indicator');
await page.screenshot({ path: path.join(out, 'reader-pdf.png') });

await page.click('#rd-views [data-view="excel"]');
await page.waitForSelector('.xl td.tx', { timeout: 30000 });
ok(/Ambient Air Quality/.test(await page.textContent('.xl')), 'Excel view: text sheet first');
ok(await page.locator('.xtabs span').count() === 2, 'Excel view: 2 sheets (text + table)');
await page.locator('.xtabs span').nth(1).click();
await page.waitForSelector('.xl td[rowspan]');
ok(await page.locator('.xl td[rowspan]').count() === 3, 'Excel view: 3 merged cells on the table sheet');
await page.screenshot({ path: path.join(out, 'reader-excel.png') });

await page.click('#rd-views [data-view="web"]');
await page.waitForSelector('.web table');
ok(await page.locator('.web td[rowspan="2"]').count() === 3, 'Web view: 3 merged cells');
await page.screenshot({ path: path.join(out, 'reader-web.png') });

// ---------------------------------------------------------------- search
console.log('Search');
await page.click('#rd-views [data-view="word"]');
await page.waitForSelector('.paper.word table');
await page.fill('#rd-q', 'PM10');
await page.waitForTimeout(200);
ok(await page.locator('.paper.word mark.sr').count() === 2, 'search: 2 matches in the Word view');
ok(/1 \/ 2/.test(await page.textContent('#rd-qc')), 'search: counter shows 1 / 2');
await page.press('#rd-q', 'Enter');
ok(/2 \/ 2/.test(await page.textContent('#rd-qc')), 'search: Enter moves to match 2');
await page.fill('#rd-q', '');

// ---------------------------------------------------------------- side by side and linking
console.log('Side by side');
await page.check('#rd-split');
await page.waitForSelector('.pane.src .ed-ta');
await page.waitForSelector('.pane.prev .paper.word table');
ok(await page.locator('.pane').count() === 2, 'two panes');
await page.click('.pane.prev table');
await page.waitForTimeout(200);
const caret = await page.evaluate(() => { const t = document.querySelector('.ed-ta'); return t.value.slice(0, t.selectionStart).split('\n').length; });
ok(caret >= 17 && caret <= 23, `click in the preview moves the editor caret to the table (line ${caret})`);
ok(await page.locator('.ed-hl .l.cur').count() >= 7, 'the table lines are highlighted in the source');
await page.evaluate(() => { document.querySelector('.ed-ta').focus(); });
await page.screenshot({ path: path.join(out, 'reader-split.png') });
await page.uncheck('#rd-split');

// ---------------------------------------------------------------- editing
console.log('Editing');
await page.click('#rd-views [data-view="md"]');
await page.waitForSelector('.ed-ta');
const edited = text.replace('[^2]', '').replace('## 2.3 Conclusion', '## 2.3 Discussion\n\n## 2.4 Added section');
await page.fill('.ed-ta', edited);
await page.waitForTimeout(700);
ok(!(await page.isHidden('#rd-edited')), '“Edited” pill appears');
ok(await page.locator('#rd-side .outline li', { hasText: '2.4 Added section' }).count() === 1, 'outline follows the edit');
ok(await page.locator('#rd-checks .row').count() === 1, 'the undefined-footnote check disappears after the fix');
ok(/5 headings/.test(await page.textContent('#rd-meta')), 'meta line updated (5 headings)');
await page.click('#rd-views [data-view="word"]');
await page.waitForSelector('.paper.word h2');
ok(/2\.4 Added section/.test(await page.textContent('.paper.word')), 'Word view shows the edited text');
const dl = page.waitForEvent('download');
await page.click('#rd-save');
const d = await dl;
ok(/\(edited\)\.md$/.test(d.suggestedFilename()), `Save names the file “${d.suggestedFilename()}”`);
ok(fs.readFileSync(await d.path(), 'utf8') === edited, 'saved file contains the edited text');
const dl2 = page.waitForEvent('download');
await page.click('#rd-create');
const d2 = await dl2;
ok(d2.suggestedFilename().endsWith('.docx'), `Create Word gives “${d2.suggestedFilename()}”`);
ok(fs.statSync(await d2.path()).size > 2000, 'the .docx has content');
await page.click('#rd-theme');
ok(['dark', 'light'].includes(await page.getAttribute('#reader', 'data-theme')), 'theme toggle sets an explicit theme');
await page.click('#rd-fs\\+');
ok(parseFloat(await page.evaluate(() => getComputedStyle(document.getElementById('reader')).getPropertyValue('--rd-scale'))) > 1, 'A+ enlarges the text');

// ---------------------------------------------------------------- back to the queue
console.log('Queue after editing');
await page.click('#rd-close');
await page.waitForTimeout(700);
ok(await page.locator('.file-list .pill-n.edited').count() === 1, 'queue row shows “Edited”');
const q = await page.evaluate(async () => { const f = window.__docToMd.state.queues.from[0]; return { edited: f.edited, text: await f.text(), name: f.name }; });
ok(q.edited && q.text.includes('2.4 Added section') && q.name === 'reader-sample.md', 'the queued file now holds the edited text');
await page.click('.file-list .read-btn');
await page.waitForSelector('#reader[open]');
await page.click('#rd-revert');
await page.waitForTimeout(700);
ok((await page.inputValue('.ed-ta')) === text, 'Revert restores the original text');
await page.click('#rd-close');
await page.waitForTimeout(300);
ok(await page.locator('.file-list .pill-n.edited').count() === 0, '“Edited” pill is gone after Revert');

// ---------------------------------------------------------------- converted Markdown
console.log('Converted Markdown');
await page.click('#tab-to');
await page.setInputFiles('#file-input', path.join(root, 'tests/fixtures/synthetic-report.docx'));
await page.click('#convert-btn'); await page.waitForSelector('#options-dialog[open]'); await page.click('#proceed-btn');
await page.waitForSelector('.result');
await page.click('.result .read-btn');
await page.waitForSelector('#reader[open]');
ok((await page.inputValue('.ed-ta')).startsWith('<!-- Source: synthetic-report.docx'), 'converted Markdown opens in the reader');
await page.click('#rd-views [data-view="excel"]');
await page.waitForSelector('.xtabs span', { timeout: 30000 });
await page.locator('.xtabs span').nth(1).click();
await page.waitForSelector('.xl td[rowspan]');
ok(await page.locator('.xl td[rowspan], .xl td[colspan]').count() >= 3, 'converted tables keep merged cells in the Excel view');
await page.click('#rd-views [data-view="md"]');
await page.fill('.ed-ta', '# Changed\n\ntext');
await page.waitForTimeout(700);
await page.click('#rd-close');
await page.waitForTimeout(300);
ok(await page.locator('.result .pill-n.edited').count() === 1, 'result card shows “Markdown edited”');

// ---------------------------------------------------------------- large file
console.log('Large file');
await page.click('#tab-from');
const big = Array.from({ length: 9000 }, (_, i) => `## Section ${i}\n\nParagraph ${i} with some words to fill the line.\n`).join('\n');
fs.writeFileSync(path.join(out, 'big.md'), big);
await page.setInputFiles('#file-input', path.join(out, 'big.md'));
await page.locator('.file-list li', { hasText: 'big.md' }).locator('.read-btn').click();
await page.waitForSelector('#reader[open]');
ok((await page.inputValue('.ed-ta')).length === big.length, `large file (${(big.length / 1024).toFixed(0)} KB) opens`);
ok(await page.locator('.ed.plain').count() === 1, 'syntax colours are switched off for a large file');
await page.click('#rd-views [data-view="web"]');
await page.waitForSelector('.web .blk');
await page.waitForFunction(() => document.querySelectorAll('.web .blk').length >= 18000, null, { timeout: 60000 });
ok(true, 'Web view renders all 18,000 blocks in chunks');
await page.click('#rd-close');

ok(nativeDialogs === 0, 'no browser pop-ups');
ok(!errors.length, 'no page errors' + (errors.length ? ': ' + errors.join(' | ') : ''));
await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) failed` : '\nAll reader checks passed');
process.exit(failures ? 1 : 0);

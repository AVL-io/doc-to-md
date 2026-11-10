// Conversion verification: compares an independent inventory of the source
// file with the Markdown output and produces a report.
import { openPackage, mainPartPath, desc, kid, kids, attr, wVal, is } from './ooxml.js';
import { loadWorkbook, sheetIncluded, encodeCol } from './xlsx-reader.js';

// ------------------------------------------------------------ tokens

const TOKEN_RE = /[\p{L}\p{N}]+(?:[.,'’][\p{L}\p{N}]+)*/gu;

export function normalizeText(s) {
  return String(s || '').normalize('NFKC').toLowerCase();
}
export function tokens(s) {
  return normalizeText(s).match(TOKEN_RE) || [];
}
const isNumberToken = (t) => /^\p{N}/u.test(t) && /\d/.test(t);
const squash = (s) => normalizeText(s).replace(/[^\p{L}\p{N}]+/gu, '');

/** Strip Markdown / embedded HTML syntax so only the visible text remains. */
export function markdownToPlain(md) {
  return md
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<\/?(b|i|s|u|em|strong|sup|sub|code|span|a)(\s[^>]*)?>/gi, '')
    .replace(/<\/?[a-z][^>]*>/gi, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
    .replace(/\\([\\`*_{}\[\]()#+\-.!|>~<])/g, '$1');
}

// ------------------------------------------------------------ inventories

/** Every piece of text in a .docx, read straight from the XML. */
async function docxInventory(buf) {
  const pkg = await openPackage(buf);
  const mainPath = await mainPartPath(pkg, 'word/document.xml');
  const rels = await pkg.rels(mainPath);
  const styles = {};
  const stPath = Object.values(rels).find(r => r.type === 'styles')?.target;
  const sd = stPath && await pkg.xml(stPath);
  if (sd) for (const s of kids(sd.documentElement, 'w', 'style')) {
    styles[attr(s, 'w', 'styleId')] = {
      name: (wVal(kid(s, 'w', 'name')) || '').toLowerCase(),
      outline: wVal(kid(kid(s, 'w', 'pPr'), 'w', 'outlineLvl')),
      num: !!kid(kid(s, 'w', 'pPr'), 'w', 'numPr'),
    };
  }
  const pieces = [];   // { text, loc, part }
  const structure = { headings: 0, listItems: 0, tables: 0, cells: 0 };
  const sourceTables = [];

  const underFallback = (el) => { for (let p = el.parentNode; p; p = p.parentNode) if (p.localName === 'Fallback') return true; return false; };
  const closestP = (el) => { for (let p = el.parentNode; p; p = p.parentNode) if (p.localName === 'p' && p.namespaceURI === el.namespaceURI) return p; return null; };

  const readPart = async (path, part) => {
    const d = await pkg.xml(path);
    if (!d) return;
    let loc = part === 'Body' ? 'Start of document' : part;
    let curP = null, buf2 = '';
    const flush = () => { if (buf2.trim()) pieces.push({ text: buf2, loc, part }); buf2 = ''; };
    const all = d.getElementsByTagName('*');
    for (const el of all) {
      const ln = el.localName;
      if (!(ln === 't' || ln === 'tab' || ln === 'fldChar')) continue;
      const ns = el.namespaceURI || '';
      const isW = /wordprocessingml/.test(ns), isM = /officeDocument\/2006\/math|ooxml\/officeDocument\/math/.test(ns);
      if (!isW && !(isM && ln === 't')) continue;
      if (underFallback(el)) continue;
      const p = closestP(el) || el.parentNode;
      if (p !== curP) {
        flush();
        curP = p;
        if (part === 'Body' && p && p.localName === 'p') {
          const sid = wVal(kid(kid(p, 'w', 'pPr'), 'w', 'pStyle'));
          const st = sid && styles[sid];
          if (st && (/^heading \d$|^title$/.test(st.name) || (st.outline != null && +st.outline < 9))) loc = 'Section: ' + (p.textContent || '').trim().slice(0, 80);
        }
      }
      buf2 += ln === 'tab' || ln === 'fldChar' ? ' ' : el.textContent;
    }
    flush();
    if (part === 'Body') {
      const singleColumn = (tbl) => desc(tbl, 'w', 'tr').filter(tr => closestTbl(tr) === tbl).every(tr => {
        let n = parseInt(wVal(kid(kid(tr, 'w', 'trPr'), 'w', 'gridBefore')) || '0', 10);
        for (const tc of desc(tr, 'w', 'tc')) if (closestTbl(tc) === tbl) n += Math.max(1, parseInt(wVal(kid(kid(tc, 'w', 'tcPr'), 'w', 'gridSpan')) || '1', 10));
        return n <= 1;
      });
      for (const p of desc(d.documentElement, 'w', 'p')) {
        if (underFallback(p) || !p.textContent.trim()) continue;
        { let inTable = false; for (let a = p.parentNode; a; a = a.parentNode) if (a.localName === 'tbl' && !singleColumn(a)) { inTable = true; break; } if (inTable) continue; }
        const pPr = kid(p, 'w', 'pPr');
        const sid = wVal(kid(pPr, 'w', 'pStyle'));
        const st = sid && styles[sid];
        const ol = wVal(kid(pPr, 'w', 'outlineLvl')) ?? st?.outline;
        if (st && /^(toc \d|toc heading|table of figures)$/.test(st.name)) continue;
        if ((st && /^heading \d$|^title$/.test(st.name)) || (ol != null && +ol < 9)) structure.headings++;
        else {
          const np = kid(pPr, 'w', 'numPr');
          const numId = np ? wVal(kid(np, 'w', 'numId')) : null;
          if ((numId && numId !== '0') || (!np && st?.num)) structure.listItems++;
        }
      }
      for (const tbl of desc(d.documentElement, 'w', 'tbl')) {
        if (underFallback(tbl) || singleColumn(tbl)) continue;
        structure.tables++;
        let nested = false;
        for (let p = tbl.parentNode; p; p = p.parentNode) if (p.localName === 'tbl' && !singleColumn(p)) { nested = true; break; }
        const grid = [];
        let r = 0;
        const anchorAt = [];
        for (const tr of desc(tbl, 'w', 'tr')) {
          if (closestTbl(tr) !== tbl) continue;
          grid[r] = [];
          let c = parseInt(wVal(kid(kid(tr, 'w', 'trPr'), 'w', 'gridBefore')) || '0', 10);
          for (let k = 0; k < c; k++) grid[r][k] = { text: '' };
          for (const tc of desc(tr, 'w', 'tc')) {
            if (closestTbl(tc) !== tbl) continue;
            const tcPr = kid(tc, 'w', 'tcPr');
            const span = Math.max(1, parseInt(wVal(kid(tcPr, 'w', 'gridSpan')) || '1', 10));
            const vm = kid(tcPr, 'w', 'vMerge');
            const cont = vm && (wVal(vm) || 'continue') === 'continue';
            const hm = kid(tcPr, 'w', 'hMerge');
            const hcont = hm && (wVal(hm) || 'continue') === 'continue';
            let cell;
            if (cont && anchorAt[c]) cell = anchorAt[c];
            else if (hcont && c > 0 && grid[r][c - 1]) cell = grid[r][c - 1];
            else { cell = { text: textOf(tc) }; structure.cells++; }
            if (cont && anchorAt[c]) { const t = textOf(tc); if (t.trim()) cell.text += ' ' + t; }
            for (let k = c; k < c + span; k++) { grid[r][k] = cell; anchorAt[k] = cell; }
            c += span;
          }
          r++;
        }
        if (!nested) sourceTables.push(grid);
      }
    }
  };
  // visible text of an element: skips values of ASK/SET fields, which Word does not display
  const textOf = (el) => {
    const out = [];
    const stack = [];
    for (const e of el.getElementsByTagName('*')) {
      if (!/wordprocessingml/.test(e.namespaceURI || '') || underFallback(e)) continue;
      if (e.localName === 'fldChar') {
        const t = attr(e, 'w', 'fldCharType');
        if (t === 'begin') stack.push({ instr: '', phase: 'code' });
        else if (t === 'separate' && stack.length) stack[stack.length - 1].phase = 'result';
        else if (t === 'end') stack.pop();
        out.push(' ');
      } else if (e.localName === 'instrText' && stack.length) stack[stack.length - 1].instr += e.textContent;
      else if (e.localName === 't') {
        const hidden = stack.some(f => f.phase === 'result' && /^\s*(ASK|SET)\b/i.test(f.instr));
        let inSimpleAsk = false;
        for (let p = e.parentNode; p && p !== el; p = p.parentNode) if (p.localName === 'fldSimple' && /^\s*(ASK|SET)\b/i.test(attr(p, 'w', 'instr') || '')) inSimpleAsk = true;
        if (!hidden && !inSimpleAsk) out.push(e.textContent);
      }
    }
    return out.join('');
  };
  const closestTbl = (el) => { for (let p = el.parentNode; p; p = p.parentNode) if (p.localName === 'tbl') return p; return null; };

  await readPart(mainPath, 'Body');
  for (const r of Object.values(rels)) {
    if (r.external) continue;
    if (r.type === 'footnotes' || r.type === 'endnotes') await readPart(r.target, 'Footnotes');
    else if (r.type === 'comments') await readPart(r.target, 'Comments');
    else if (r.type === 'header' || r.type === 'footer') await readPart(r.target, 'Header/Footer');
  }
  return { pieces, structure, sourceTables };
}

async function xlsxInventory(buf, opts) {
  const wb = await loadWorkbook(buf);
  const pieces = [];
  const values = [];
  let merges = 0, sheets = 0;
  for (const sh of wb.sheets) {
    const included = sheetIncluded(sh, opts);
    if (included) sheets++;
    const cells = [...sh.cells.values()].sort((a, b) => a.r - b.r || a.c - b.c);
    for (const c of cells) {
      if (!c.display) continue;
      pieces.push({ text: c.display, loc: `${sh.name}!${encodeCol(c.c)}${c.r + 1}`, part: 'Cells' });
      const hidden = !opts.includeHiddenRowsCols && (sh.hiddenRows.has(c.r) || sh.hiddenCols.has(c.c));
      if (included && !hidden && squash(c.display)) values.push(squash(c.display));
    }
    pieces.push({ text: sh.name, loc: `${sh.name} (sheet name)`, part: 'Sheet names', label: true });
    for (const cm of sh.comments) pieces.push({ text: cm.text, loc: `${sh.name}!${cm.ref} (comment)`, part: 'Comments' });
    if (included) for (const m of sh.merges) {
      const a = sh.cells.get(m.r0 * 20000 + m.c0);
      if (a && a.display) merges++;
    }
  }
  return { pieces, values, merges, sheets };
}

// ------------------------------------------------------------ output parsing

/** Find the tables in generated Markdown: [{ grid: string[][], html }] */
export function parseMarkdownTables(md) {
  const lines = md.split('\n');
  const tables = [];
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i].trim();
    if (l === '<table>') {
      let j = i;
      while (j < lines.length && lines[j].trim() !== '</table>') j++;
      const html = lines.slice(i, j + 1).join('\n');
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const t = doc.querySelector('table');
      if (t) tables.push({ grid: htmlTableGrid(t), html: true });
      i = j;
      continue;
    }
    if (l.startsWith('|') && i + 1 < lines.length && /^\|(\s*:?-{3,}:?\s*\|)+\s*$/.test(lines[i + 1].trim())) {
      const rows = [splitPipe(l)];
      let j = i + 2;
      while (j < lines.length && lines[j].trim().startsWith('|')) rows.push(splitPipe(lines[j].trim())), j++;
      tables.push({ grid: rows, html: false });
      i = j - 1;
    }
  }
  return tables;
}

function splitPipe(line) {
  const inner = line.replace(/^\|/, '').replace(/\|\s*$/, '');
  const cells = [];
  let cur = '';
  for (let i = 0; i < inner.length; i++) {
    if (inner[i] === '\\' && inner[i + 1] === '|') { cur += '|'; i++; continue; }
    if (inner[i] === '|') { cells.push(cur); cur = ''; continue; }
    cur += inner[i];
  }
  cells.push(cur);
  return cells.map(c => markdownToPlain(c).trim());
}

function htmlTableGrid(table) {
  const grid = [];
  const rows = [...table.rows].filter(r => r.closest('table') === table);
  rows.forEach((tr, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    for (const td of tr.cells) {
      while (grid[r][c] != null) c++;
      const rs = Math.max(1, td.rowSpan || 1), cs = Math.max(1, td.colSpan || 1);
      const text = td.textContent;
      for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) {
        grid[r + i] = grid[r + i] || [];
        grid[r + i][c + j] = text;
      }
      c += cs;
    }
  });
  return grid.slice(0, rows.length).map(r => Array.from(r, v => v ?? ''));
}

// ------------------------------------------------------------ main

/**
 * @param {object} p { kind: 'docx'|'xlsx'|'pdf', buf, md, log, pages, opts }
 */
export async function verifyConversion({ kind, buf, md, log, pages, opts = {}, checkMd }) {
  let inv;
  const structure = [];
  const tableChecks = [];
  const warnings = [...(log.warnings || [])];

  if (kind === 'docx') inv = await docxInventory(buf);
  else if (kind === 'xlsx') inv = await xlsxInventory(buf, opts);
  else if (kind === 'pdf') inv = { pieces: pages.map(p => ({ text: p.rawText, loc: `Page ${p.num}`, part: 'Page' })) };
  else throw new Error('Unknown file type for verification');

  // ---- word-level comparison ----
  const plain = markdownToPlain(md);
  const outCount = countMap(tokens(plain));
  const exclCount = countMap(log.excluded.flatMap(e => tokens(e.text)));
  const outSquash = squash(plain);
  let total = 0, matched = 0, excluded = 0;
  const num = { total: 0, matched: 0, excluded: 0, missing: [] };
  const missingTokens = [];   // { t, loc, idx }
  let idx = 0;
  const perLoc = new Map();
  for (const piece of inv.pieces) {
    for (const t of tokens(piece.text)) {
      total++;
      const isNum = isNumberToken(t);
      if (isNum) num.total++;
      const loc = piece.loc;
      if (!perLoc.has(loc)) perLoc.set(loc, { total: 0, missing: 0 });
      perLoc.get(loc).total++;
      if (take(outCount, t)) { matched++; if (isNum) num.matched++; }
      else if (take(exclCount, t)) { excluded++; if (isNum) num.excluded++; }
      else if (!isNum && t.length >= 3 && outSquash.includes(squash(t))) { matched++; } // spacing differs only
      else {
        missingTokens.push({ t, loc, idx, part: piece.part });
        perLoc.get(loc).missing++;
        if (isNum) num.missing.push({ value: t, loc });
      }
      idx++;
    }
  }
  const considered = total - excluded;
  const coverage = considered > 0 ? (matched / considered) * 100 : 100;

  // group missing tokens into passages (consecutive in source order)
  const passages = [];
  for (const m of missingTokens) {
    const last = passages[passages.length - 1];
    if (last && last.loc === m.loc && m.idx - last.lastIdx <= 3) { last.words.push(m.t); last.lastIdx = m.idx; }
    else passages.push({ loc: m.loc, words: [m.t], lastIdx: m.idx, part: m.part });
  }

  // ---- duplication check: words that occur more often in the output than in the source ----
  const srcAll = countMap(inv.pieces.flatMap(p => tokens(p.text)));
  const GENERATED = new Set(['sheet', 'range', 'image', 'chart', 'comment', 'note', 'header', 'footer', 'page', 'table', 'see', 'this', 'is', 'empty', 'the']);
  let extra = 0;
  const extraExamples = [];
  const dupText = (checkMd ? markdownToPlain(checkMd) : plain).replace(/^\*Range [A-Z]+\d+:[A-Z]+\d+\*$/gm, ' ');
  for (const [t, n] of countMap(tokens(dupText))) {
    if (isNumberToken(t) || GENERATED.has(t) || t.length < 2) continue;
    const d = n - (srcAll.get(t) || 0);
    if (d > 0) { extra += d; if (extraExamples.length < 8) extraExamples.push(t); }
  }
  const duplication = extra > Math.max(5, total * 0.005);
  if (duplication) warnings.push(`${extra} word(s) appear more often in the output than in the source (for example: ${extraExamples.join(', ')}). Some content may be duplicated.`);

  // ---- structure ----
  const outTables = parseMarkdownTables(md);
  const mdNoTables = md.split('\n').filter(l => !/^\s*(\||<t)/.test(l)).join('\n');
  const outHeadings = (mdNoTables.match(/^#{1,6} /gm) || []).length;
  const outListItems = (mdNoTables.match(/^\s*([-*+]|\d+\.)\s+\S/gm) || []).length;

  if (kind === 'docx') {
    const s = inv.structure;
    structure.push(row('Headings', s.headings, outHeadings, 'Headings inside tables are not counted.'));
    structure.push(row('List items', s.listItems, outListItems, opts.removeToc === false ? 'Output also counts table-of-contents entries.' : ''));
    structure.push(row('Tables', s.tables, countTablesDeep(md), opts.tableMode === 'flatten' ? 'Nested tables are flattened to text in pure-Markdown mode.' : ''));
    // cell-by-cell table comparison
    const src = inv.sourceTables;
    const n = Math.max(src.length, outTables.length);
    for (let i = 0; i < n; i++) tableChecks.push(compareGrid(i + 1, src[i], outTables[i]));
  } else if (kind === 'xlsx') {
    structure.push(row('Sheets converted', inv.sheets, log.stats.sheets));
    const outVals = countMap(outTables.flatMap(t => t.grid.flat()).map(squash).filter(Boolean));
    let ok = 0;
    const miss = [];
    for (const v of inv.values) { if (take(outVals, v)) ok++; else miss.push(v); }
    structure.push(row('Non-empty cells', inv.values.length, ok, miss.length ? `${miss.length} cell value(s) not found in the output tables (e.g. ${miss.slice(0, 5).join(', ')}).` : ''));
    if (opts.tableMode !== 'flatten') structure.push(row('Merged cell regions', inv.merges, log.stats.merges));
  } else {
    const markers = (md.match(/<!-- Page \d+ -->/g) || []).length;
    const joined = log.stats.pageMarkersJoined || 0;
    structure.push(row('Pages', pages.length, opts.pageMarkers === false ? pages.length : markers + joined,
      opts.pageMarkers === false ? 'Page markers were switched off.' : joined ? `${joined} page marker(s) omitted where a table continues across pages.` : ''));
    structure.push({ label: 'Tables detected', source: '—', output: log.stats.tables, ok: true, note: 'PDF files do not store tables explicitly; tables are rebuilt from border lines and tags.' });
    structure.push({ label: 'Merged cells detected', source: '—', output: log.stats.mergedCells, ok: true, note: '' });
    const inferred = [...new Set(log.stats.inferredTables || [])];
    const low = [...new Set(log.stats.lowConfidenceTables || [])];
    if (inferred.length) warnings.push(`Tables without border lines on page(s) ${inferred.join(', ')} were rebuilt from the PDF tags; merged cells in these tables are inferred. Visual review is recommended.`);
    if (low.length) warnings.push(`Tables with irregular borders on page(s) ${low.join(', ')} could not be fully resolved. Visual review is recommended.`);
  }

  // ---- exclusions summary ----
  const exMap = new Map();
  for (const e of log.excluded) {
    const g = exMap.get(e.category) || { category: e.category, count: 0, words: 0, examples: [] };
    g.count++;
    g.words += tokens(e.text).length;
    if (e.text && g.examples.length < 3 && !g.examples.includes(e.text.slice(0, 80))) g.examples.push(e.text.slice(0, 80));
    exMap.set(e.category, g);
  }

  const structureOk = structure.every(s => s.ok) && tableChecks.every(t => t.ok);
  const missing = missingTokens.length;
  let verdict;
  if (missing === 0 && structureOk) verdict = 'complete';
  else if (num.missing.length === 0 && coverage >= 99.5) verdict = 'review';
  else verdict = 'incomplete';
  if (verdict === 'complete' && ((kind === 'pdf' && warnings.some(w => /review/i.test(w))) || duplication)) verdict = 'review';

  const reviewItems = buildReviewItems({ kind, log, passages, structure, tableChecks, warnings, duplication });

  const pagesWithLoss = kind === 'pdf' ? [...perLoc.entries()].filter(([, v]) => v.missing > 0).map(([k, v]) => ({ loc: k, missing: v.missing, total: v.total })) : [];

  return {
    kind, verdict,
    words: { total, matched, excluded, missing, coverage, extra },
    numbers: num,
    passages: passages.slice(0, 200).map(p => ({ loc: p.loc, part: p.part, text: p.words.join(' ') })),
    passageCount: passages.length,
    structure, tableChecks,
    exclusions: [...exMap.values()],
    warnings,
    pagesWithLoss,
    reviewItems,
  };
}

/**
 * The checklist of items a person should look at, most serious first.
 * status: 'bad' (missing / mismatched) or 'warn' (review recommended).
 * highlight: where to show the item in the side-by-side review.
 */
function buildReviewItems({ kind, log, passages, structure, tableChecks, warnings, duplication }) {
  const items = [];
  const pageOf = (loc) => { const m = /^Page (\d+)$/.exec(loc || ''); return m ? +m[1] : null; };
  for (const p of passages.slice(0, 300)) {
    const words = p.words;
    const text = words.join(' ');
    const hasNumber = words.some(isNumberToken);
    const page = pageOf(p.loc);
    items.push({
      status: 'bad', page, location: p.loc,
      item: `"${text.length > 80 ? text.slice(0, 80) + '…' : text}"`,
      reason: hasNumber ? 'Number(s) not found in the output.' : 'Text not found in the output.',
      highlight: page ? { kind: 'text', page, words } : null,
    });
  }
  for (const t of tableChecks) {
    if (t.ok) continue;
    const m = t.mismatches[0];
    items.push({
      status: 'bad', location: `Table ${t.index}`, item: `Table ${t.index}`,
      reason: t.summary + (m ? ` First difference: row ${m.row}, column ${m.col}.` : ''),
      highlight: { kind: 'table', docIndex: t.index },
    });
  }
  for (const s of structure) {
    if (s.ok) continue;
    items.push({ status: kind === 'xlsx' ? 'bad' : 'warn', location: 'Whole document', item: s.label, reason: `Source ${s.source}, output ${s.output}. ${s.note || ''}`.trim(), highlight: null });
  }
  for (const t of log.tables || []) {
    if (!t.inferred && !t.lowConfidence) continue;
    const pages = [...new Set(t.boxes.map(b => b.page))];
    items.push({
      status: 'warn', page: t.page,
      location: pages.length > 1 ? `Pages ${pages[0]}–${pages[pages.length - 1]}` : `Page ${t.page}`,
      item: `Table ${t.docIndex} (${t.rows} rows × ${t.cols} columns${t.merged ? `, ${t.merged} merged cell${t.merged > 1 ? 's' : ''}` : ''})`,
      reason: t.lowConfidence
        ? 'Irregular border lines: some merged cells could not be resolved. Check the cell layout.'
        : 'No border lines: structure rebuilt from the PDF tags. Check row/column alignment and merged cells.',
      highlight: { kind: 'table', page: t.page, indexOnPage: t.indexOnPage, docIndex: t.docIndex, boxes: t.boxes },
    });
  }
  for (const w of warnings) {
    if (/were rebuilt from the PDF tags|irregular borders/.test(w)) continue; // listed per table above
    const pm = /^Page (\d+) contains no extractable text/.exec(w);
    items.push({
      status: pm ? 'bad' : 'warn', page: pm ? +pm[1] : null, location: pm ? `Page ${pm[1]}` : 'Whole document',
      item: duplication && /appear more often/.test(w) ? 'Possible duplication' : 'Warning', reason: w,
      highlight: pm ? { kind: 'page', page: +pm[1] } : null,
    });
  }
  const rank = { bad: 0, warn: 1 };
  items.sort((a, b) => rank[a.status] - rank[b.status] || (a.page ?? 1e9) - (b.page ?? 1e9));
  items.forEach((it, i) => { it.n = i + 1; it.id = `${it.status}|${it.location}|${it.item}`; });
  return items;
}

function row(label, source, output, note = '') {
  return { label, source, output, ok: source === output, note };
}
function countTablesDeep(md) {
  return (md.match(/<table>/g) || []).length + parseMarkdownTables(md).filter(t => !t.html).length;
}
function countMap(arr) {
  const m = new Map();
  for (const t of arr) m.set(t, (m.get(t) || 0) + 1);
  return m;
}
function take(m, t) {
  const v = m.get(t);
  if (v > 0) { m.set(t, v - 1); return true; }
  return false;
}
function isSubsequence(a, b) {
  let i = 0;
  for (let j = 0; j < b.length && i < a.length; j++) if (a[i] === b[j]) i++;
  return i === a.length;
}

function compareGrid(n, src, out) {
  if (!src) return { index: n, ok: false, summary: 'Extra table in the output (not found in the source).', mismatches: [] };
  if (!out) return { index: n, ok: false, summary: 'Table missing from the output.', mismatches: [] };
  const sr = src.length, sc = Math.max(0, ...src.map(r => r.length));
  const or = out.grid.length, oc = Math.max(0, ...out.grid.map(r => r.length));
  const mismatches = [];
  let positions = 0, okCount = 0;
  for (let r = 0; r < Math.max(sr, or); r++) for (let c = 0; c < Math.max(sc, oc); c++) {
    const s = squash(src[r]?.[c]?.text ?? '');
    const o = squash(out.grid[r]?.[c] ?? '');
    positions++;
    if (s === o || (s && isSubsequence(s, o) && o.length <= s.length * 1.15 + 8)) okCount++;
    else if (mismatches.length < 5) mismatches.push({ row: r + 1, col: c + 1, source: (src[r]?.[c]?.text ?? '').slice(0, 60), output: (out.grid[r]?.[c] ?? '').slice(0, 60) });
  }
  const dimsOk = sr === or && sc === oc;
  const ok = dimsOk && okCount === positions;
  return {
    index: n, ok,
    summary: `${sr}×${sc} grid in source, ${or}×${oc} in output; ${okCount}/${positions} cell positions match${dimsOk ? '' : ' (grid size differs)'}.`,
    mismatches,
  };
}

// ------------------------------------------------------------ report rendering

export const MARK = { ok: '🟢', warn: '🟠', bad: '🔴' };
const VERDICT_MARK = { complete: MARK.ok, review: MARK.warn, incomplete: MARK.bad };

const VERDICT_TEXT = {
  complete: 'COMPLETE: all source text and numbers were found in the output, and the structure checks agree.',
  review: 'REVIEW RECOMMENDED: no numbers are missing and coverage is at least 99.5%, but some items need checking (see below).',
  incomplete: 'INCOMPLETE: some text or numbers from the source were not found in the output (see below).',
};

export function reviewChecklistMarkdown(rep) {
  const items = rep.reviewItems || [];
  if (!items.length) return `${MARK.ok} Nothing requires review.`;
  const esc = (s) => String(s).replace(/\|/g, '/').replace(/\n/g, ' ');
  return `${items.filter(i => i.status === 'bad').length} item(s) marked ${MARK.bad} and ${items.filter(i => i.status === 'warn').length} item(s) marked ${MARK.warn}.\n\n` +
    `| # | Status | Location | Item | Reason | Reviewed |\n| --- | --- | --- | --- | --- | --- |\n` +
    items.map(i => `| ${i.n} | ${MARK[i.status]} | ${esc(i.location)} | ${esc(i.item)} | ${esc(i.reason)} | ☐ |`).join('\n');
}

export function reportToMarkdown(rep, fileName, { checklist = true } = {}) {
  const L = [];
  const pct = rep.words.coverage.toFixed(2);
  L.push(`# Conversion Verification Report`);
  L.push(`**Source file:** ${fileName}  \n**Checked:** ${new Date().toISOString().slice(0, 16).replace('T', ' ')} (local processing; no data left this computer)`);
  L.push(`## Result\n\n${VERDICT_MARK[rep.verdict]} **${VERDICT_TEXT[rep.verdict]}**\n\nKey: ${MARK.ok} OK · ${MARK.warn} review recommended · ${MARK.bad} missing or mismatched`);
  if (checklist) L.push(`## Items to review\n\n` + reviewChecklistMarkdown(rep));
  L.push(`## 1. Text inventory check\n\n| Measure | Value |\n| --- | --- |\n| Words in source | ${rep.words.total} |\n| Found in output | ${rep.words.matched} |\n| Deliberately excluded (options) | ${rep.words.excluded} |\n| Not found | ${rep.words.missing ? MARK.bad + ' ' : MARK.ok + ' '}${rep.words.missing} |\n| Coverage of retained content | ${pct}% |`);
  L.push(`## 2. Numerical data check\n\n| Measure | Value |\n| --- | --- |\n| Numbers in source | ${rep.numbers.total} |\n| Found in output | ${rep.numbers.matched} |\n| Deliberately excluded | ${rep.numbers.excluded} |\n| Not found | ${rep.numbers.missing.length ? MARK.bad + ' ' : MARK.ok + ' '}${rep.numbers.missing.length} |`);
  if (rep.numbers.missing.length) L.push(rep.numbers.missing.slice(0, 100).map(m => `- ${MARK.bad} \`${m.value}\` (${m.loc})`).join('\n'));
  L.push(`## 3. Structural check\n\n| Element | Source | Output | Status | Note |\n| --- | --- | --- | --- | --- |\n` + rep.structure.map(s => `| ${s.label} | ${s.source} | ${s.output} | ${s.ok ? MARK.ok + ' OK' : MARK.bad + ' CHECK'} | ${s.note || ''} |`).join('\n'));
  if (rep.tableChecks.length) {
    L.push(`### Table-by-table comparison\n\n` + rep.tableChecks.map(t => `- ${t.ok ? MARK.ok : MARK.bad} **Table ${t.index}:** ${t.ok ? 'OK' : 'CHECK'}: ${t.summary}` + t.mismatches.map(m => `\n  - Row ${m.row}, column ${m.col}: source "${m.source}" / output "${m.output}"`).join('')).join('\n'));
  }
  L.push(`## 4. Deliberate exclusions\n\n` + (rep.exclusions.length ? `| Category | Items | Words | Examples |\n| --- | --- | --- | --- |\n` + rep.exclusions.map(e => `| ${e.category} | ${e.count} | ${e.words} | ${e.examples.map(x => x.replace(/\|/g, '/')).join('; ')} |`).join('\n') : 'None.'));
  L.push(`## 5. Text not found in output\n\n` + (rep.passages.length ? rep.passages.map(p => `- ${MARK.bad} **${p.loc}:** ${p.text}`).join('\n') + (rep.passageCount > rep.passages.length ? `\n- … ${rep.passageCount - rep.passages.length} more` : '') : 'None.'));
  if (rep.pagesWithLoss.length) L.push(`### Pages with missing text\n\n` + rep.pagesWithLoss.map(p => `- ${MARK.bad} ${p.loc}: ${p.missing} of ${p.total} words`).join('\n'));
  L.push(`## 6. Warnings\n\n` + (rep.warnings.length ? rep.warnings.map(w => `- ${MARK.warn} ${w}`).join('\n') : 'None.'));
  L.push(`## Limitations\n\nThe text check confirms that every word and number is *present* in the output. It does not fully prove that each value sits in the correct table cell. ${rep.kind === 'pdf' ? 'For PDF files, table structure is reconstructed from the page layout, so critical tables should be compared visually with the original (use the side-by-side review).' : 'For Word and Excel files, table positions are compared cell by cell against the source (section 3).'}`);
  return L.join('\n\n') + '\n';
}

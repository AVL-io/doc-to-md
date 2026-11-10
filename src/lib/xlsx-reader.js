// Excel (.xlsx / .xlsm) -> document model.
import { openPackage, mainPartPath, kids, kid, desc, attr } from './ooxml.js';
import { BUILTIN, formatValue } from './numfmt.js';

export const DEFAULT_OPTIONS = {
  includeHiddenSheets: false,
  includeHiddenRowsCols: false,
  removeComments: true,
  splitOnBlankRows: true,
};

/** Load the workbook into a plain structure (shared by the reader and the verifier). */
export async function loadWorkbook(buf) {
  const pkg = await openPackage(buf);
  const wbPath = await mainPartPath(pkg, 'xl/workbook.xml');
  const wb = await pkg.xml(wbPath);
  if (!wb) throw new Error('This file does not look like an Excel workbook (xl/workbook.xml is missing).');
  const rels = await pkg.rels(wbPath);
  const relOf = (type) => Object.values(rels).filter(r => r.type === type).map(r => r.target);
  const date1904 = /^(1|true)$/i.test(attr(kid(wb.documentElement, 's', 'workbookPr'), 's', 'date1904') || '');

  // shared strings
  const shared = [];
  const ssPath = relOf('sharedStrings')[0];
  const ss = ssPath && await pkg.xml(ssPath);
  if (ss) for (const si of kids(ss.documentElement, 's', 'si')) shared.push(stringItem(si));

  // number formats per style index
  const fmts = [];
  const stPath = relOf('styles')[0];
  const sx = stPath && await pkg.xml(stPath);
  if (sx) {
    const custom = {};
    for (const nf of desc(kid(sx.documentElement, 's', 'numFmts'), 's', 'numFmt')) custom[nf.getAttribute('numFmtId')] = nf.getAttribute('formatCode');
    for (const xf of kids(kid(sx.documentElement, 's', 'cellXfs'), 's', 'xf')) {
      const id = xf.getAttribute('numFmtId') || '0';
      fmts.push(custom[id] ?? BUILTIN[id] ?? 'General');
    }
  }

  const sheets = [];
  for (const sh of desc(kid(wb.documentElement, 's', 'sheets'), 's', 'sheet')) {
    const rid = attr(sh, 'r', 'id');
    const target = rels[rid]?.target;
    const state = sh.getAttribute('state') || 'visible';
    const sheet = { name: sh.getAttribute('name'), state, cells: new Map(), merges: [], hiddenRows: new Set(), hiddenCols: new Set(), comments: [], drawings: 0, path: target, kind: rels[rid]?.type };
    sheets.push(sheet);
    if (!target || rels[rid]?.type !== 'worksheet') continue;
    const d = await pkg.xml(target);
    if (!d) continue;
    const root = d.documentElement;
    for (const col of desc(kid(root, 's', 'cols'), 's', 'col')) {
      if (/^(1|true)$/i.test(col.getAttribute('hidden') || '')) {
        for (let c = +col.getAttribute('min'); c <= +col.getAttribute('max'); c++) sheet.hiddenCols.add(c - 1);
      }
    }
    for (const row of kids(kid(root, 's', 'sheetData'), 's', 'row')) {
      const r = parseInt(row.getAttribute('r'), 10) - 1;
      if (/^(1|true)$/i.test(row.getAttribute('hidden') || '')) sheet.hiddenRows.add(r);
      let nextCol = 0;
      for (const c of kids(row, 's', 'c')) {
        const ref = c.getAttribute('r');
        const pos = ref ? decodeRef(ref) : { r, c: nextCol };
        nextCol = pos.c + 1;
        const t = c.getAttribute('t') || 'n';
        const vEl = kid(c, 's', 'v');
        const fEl = kid(c, 's', 'f');
        let value = null, display = '';
        if (t === 's') { value = shared[parseInt(vEl?.textContent, 10)] ?? ''; display = value; }
        else if (t === 'inlineStr') { const is = kid(c, 's', 'is'); value = is ? stringItem(is) : ''; display = value; }
        else if (t === 'str') { value = vEl?.textContent ?? ''; display = value; }
        else if (t === 'b') { value = vEl ? vEl.textContent === '1' : null; display = value == null ? '' : value ? 'TRUE' : 'FALSE'; }
        else if (t === 'e') { value = vEl?.textContent ?? ''; display = value; }
        else if (t === 'd') { value = vEl?.textContent ?? ''; display = value; }
        else if (vEl && vEl.textContent !== '') {
          value = parseFloat(vEl.textContent);
          const fmt = fmts[parseInt(c.getAttribute('s') || '0', 10)] || 'General';
          display = formatValue(value, fmt, date1904);
        }
        const formula = fEl ? (fEl.textContent || '(shared formula)') : null;
        if (display === '' && formula == null) continue;
        sheet.cells.set(pos.r * 20000 + pos.c, { r: pos.r, c: pos.c, value, display: String(display), formula, uncalculated: formula != null && (!vEl || vEl.textContent === '') });
      }
    }
    for (const m of desc(kid(root, 's', 'mergeCells'), 's', 'mergeCell')) {
      const [a, b] = (m.getAttribute('ref') || '').split(':');
      if (!a) continue;
      const s = decodeRef(a), e = decodeRef(b || a);
      sheet.merges.push({ r0: Math.min(s.r, e.r), c0: Math.min(s.c, e.c), r1: Math.max(s.r, e.r), c1: Math.max(s.c, e.c) });
    }
    const srels = await pkg.rels(target);
    for (const rel of Object.values(srels)) {
      if (rel.type === 'drawing') {
        const dd = await pkg.xml(rel.target);
        if (dd) sheet.drawings += dd.documentElement.children.length;
      }
      if (rel.type === 'comments') {
        const cd = await pkg.xml(rel.target);
        if (!cd) continue;
        const authors = desc(cd.documentElement, 's', 'author').map(a => a.textContent);
        for (const cm of desc(cd.documentElement, 's', 'comment')) {
          const t = kid(cm, 's', 'text');
          sheet.comments.push({ ref: cm.getAttribute('ref'), author: authors[+cm.getAttribute('authorId')] || '', text: t ? stringItem(t) : '' });
        }
      }
    }
  }
  return { sheets, date1904 };
}

function stringItem(si) {
  // concatenate <t> in the item itself and in rich-text runs; skip phonetic runs
  let s = '';
  for (const c of si.children) {
    if (c.localName === 't') s += c.textContent;
    else if (c.localName === 'r') for (const t of c.children) if (t.localName === 't') s += t.textContent;
  }
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
}

export function decodeRef(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim());
  if (!m) return { r: 0, c: 0 };
  let c = 0;
  for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}

export function encodeCol(c) {
  let s = '';
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

/** Is a sheet included under the given options? */
export function sheetIncluded(sheet, opts) {
  if (sheet.kind && sheet.kind !== 'worksheet') return false;
  return sheet.state === 'visible' || opts.includeHiddenSheets;
}

export async function readXlsx(buf, options = {}) {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const wb = await loadWorkbook(buf);
  const log = { excluded: [], warnings: [], stats: { sheets: 0, tables: 0, cells: 0, merges: 0, drawings: 0, comments: 0 } };
  const exclude = (category, text) => log.excluded.push({ category, text: (text || '').trim() });
  const blocks = [];
  const footnotes = [];

  for (const sheet of wb.sheets) {
    if (sheet.kind && sheet.kind !== 'worksheet') { log.warnings.push(`Sheet "${sheet.name}" is a chart sheet or macro sheet and was not converted.`); continue; }
    if (!sheetIncluded(sheet, opts)) {
      exclude('Hidden sheet', `${sheet.name} ${[...sheet.cells.values()].map(c => c.display).join(' ')}`);
      continue;
    }
    log.stats.sheets++;
    log.stats.drawings += sheet.drawings;
    if (sheet.drawings) exclude('Chart / image', `${sheet.drawings} drawing object(s) on sheet "${sheet.name}"`);
    blocks.push({ type: 'heading', level: 2, inlines: [{ type: 'text', text: `Sheet: ${sheet.name}` }] });

    const hiddenR = opts.includeHiddenRowsCols ? new Set() : sheet.hiddenRows;
    const hiddenC = opts.includeHiddenRowsCols ? new Set() : sheet.hiddenCols;
    const uncalculated = [];
    for (const cell of sheet.cells.values()) {
      if (hiddenR.has(cell.r) || hiddenC.has(cell.c)) { if (cell.display) exclude('Hidden row/column', cell.display); continue; }
      if (cell.uncalculated) uncalculated.push(encodeCol(cell.c) + (cell.r + 1));
    }
    if (uncalculated.length) log.warnings.push(`Sheet "${sheet.name}": ${uncalculated.length} formula cell(s) have no saved result (e.g. ${uncalculated.slice(0, 5).join(', ')}). Open and save the file in Excel to store the results.`);

    const tables = sheetTables(sheet, hiddenR, hiddenC, opts.splitOnBlankRows);
    if (!tables.length) blocks.push({ type: 'paragraph', inlines: [{ type: 'text', text: '(This sheet is empty.)', italic: true }] });
    for (const t of tables) {
      log.stats.tables++;
      log.stats.cells += t.nonEmpty;
      log.stats.merges += t.merges;
      if (tables.length > 1) blocks.push({ type: 'paragraph', inlines: [{ type: 'text', text: `Range ${t.range}`, italic: true }] });
      blocks.push(t.table);
    }

    if (sheet.comments.length) {
      log.stats.comments += sheet.comments.length;
      if (opts.removeComments) for (const c of sheet.comments) exclude('Comment', c.text);
      else blocks.push({
        type: 'list', ordered: false,
        items: sheet.comments.map(c => ({ blocks: [{ type: 'paragraph', inlines: [{ type: 'text', text: `Comment on ${c.ref}${c.author ? ' (' + c.author + ')' : ''}: `, bold: true }, { type: 'text', text: c.text.replace(/\s+/g, ' ') }] }] })),
      });
    }
  }
  return { doc: { blocks, footnotes }, log };
}

/**
 * Build one or more tables from the visible cells of a sheet. Fully empty rows
 * split the sheet into separate tables; fully empty columns are dropped.
 */
function sheetTables(sheet, hiddenR, hiddenC, split) {
  const merges = sheet.merges.filter(m => !hiddenR.has(m.r0) && !hiddenC.has(m.c0));
  const mergeAt = new Map();
  const covered = new Set();
  for (const m of merges) {
    mergeAt.set(m.r0 * 20000 + m.c0, m);
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(r * 20000 + c);
  }
  const cellAt = (r, c) => sheet.cells.get(r * 20000 + c);
  let maxR = -1, maxC = -1;
  const rowHas = new Set();
  for (const cell of sheet.cells.values()) {
    if (hiddenR.has(cell.r) || hiddenC.has(cell.c) || covered.has(cell.r * 20000 + cell.c) || cell.display === '') continue;
    maxR = Math.max(maxR, cell.r); maxC = Math.max(maxC, cell.c);
    rowHas.add(cell.r);
  }
  if (maxR < 0) return [];
  // rows spanned by a merge that has content are never treated as separators
  for (const m of merges) {
    const a = cellAt(m.r0, m.c0);
    if (a && a.display !== '') for (let r = m.r0; r <= m.r1; r++) rowHas.add(r);
  }
  // group rows into bands separated by empty rows
  const bands = [];
  let cur = null;
  for (let r = 0; r <= maxR; r++) {
    if (hiddenR.has(r)) continue;
    if (rowHas.has(r)) { if (!cur) { cur = []; bands.push(cur); } cur.push(r); }
    else if (split) cur = null;
  }
  const out = [];
  for (const rows of bands) {
    const rowSet = new Set(rows);
    const colHas = new Set();
    for (const r of rows) for (let c = 0; c <= maxC; c++) {
      if (hiddenC.has(c)) continue;
      const cell = cellAt(r, c);
      if (cell && cell.display !== '' && !covered.has(r * 20000 + c)) {
        colHas.add(c);
        const m = mergeAt.get(r * 20000 + c);
        if (m) for (let k = m.c0; k <= m.c1; k++) if (!hiddenC.has(k)) colHas.add(k);
      }
    }
    const cols = [...colHas].sort((a, b) => a - b);
    const colIdx = new Map(cols.map((c, i) => [c, i]));
    const table = { type: 'table', rows: [] };
    let nonEmpty = 0, mergeCount = 0;
    for (const r of rows) {
      const row = [];
      for (const c of cols) {
        const key = r * 20000 + c;
        if (covered.has(key)) continue;
        const cell = cellAt(r, c);
        const m = mergeAt.get(key);
        let rowspan = 1, colspan = 1;
        if (m) {
          rowspan = 0; for (let k = m.r0; k <= m.r1; k++) if (rowSet.has(k)) rowspan++;
          colspan = 0; for (let k = m.c0; k <= m.c1; k++) if (colIdx.has(k)) colspan++;
          rowspan = Math.max(1, rowspan); colspan = Math.max(1, colspan);
          if (rowspan > 1 || colspan > 1) mergeCount++;
        }
        const text = cell ? cell.display : '';
        if (text !== '') nonEmpty++;
        row.push({ blocks: text === '' ? [] : [{ type: 'paragraph', inlines: textInlines(text) }], rowspan, colspan, header: false });
      }
      table.rows.push(row);
    }
    // a merge covering cells in a removed (empty) column may leave covered cells
    // that fall in a column we kept but whose anchor column we dropped: handled
    // because colHas always includes the full width of merges with content.
    const range = `${encodeCol(cols[0])}${rows[0] + 1}:${encodeCol(cols[cols.length - 1])}${rows[rows.length - 1] + 1}`;
    out.push({ table, nonEmpty, merges: mergeCount, range });
  }
  return out;
}

function textInlines(text) {
  const parts = String(text).split(/\r?\n/);
  const out = [];
  parts.forEach((p, i) => { if (i) out.push({ type: 'break' }); if (p) out.push({ type: 'text', text: p }); });
  return out;
}

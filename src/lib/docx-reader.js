// Word (.docx) -> document model.
import { openPackage, mainPartPath, is, inNs, kids, kid, desc, attr, wVal, onOff, symbolChar } from './ooxml.js';
import { buildLists, trimInlines, inlinesText, blocksText } from './ast.js';

export const DEFAULT_OPTIONS = {
  removeImages: true,
  removeHeadersFooters: true,
  removePageNumbers: true,
  removeFootnotes: false,
  removeComments: true,
  acceptChanges: true,
  removeToc: true,
  removeCover: false,
};

export async function readDocx(buf, options = {}) {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const pkg = await openPackage(buf);
  const mainPath = await mainPartPath(pkg, 'word/document.xml');
  const main = await pkg.xml(mainPath);
  if (!main) throw new Error('This file does not look like a Word document (word/document.xml is missing).');
  const rels = await pkg.rels(mainPath);
  const partOf = (type) => Object.values(rels).filter(r => r.type === type && !r.external).map(r => r.target);

  const log = { excluded: [], warnings: [], stats: { images: 0, charts: 0, textBoxes: 0, footnotes: 0, comments: 0, tables: 0, headings: 0, listItems: 0 } };
  const exclude = (category, text) => log.excluded.push({ category, text: (text || '').trim() });

  const styles = await loadStyles(pkg, partOf('styles')[0]);
  const numbering = await loadNumbering(pkg, partOf('numbering')[0]);

  const st = {
    opts, rels, styles, numbering, log, exclude,
    counters: {},
    fields: [],        // open complex fields (span runs and paragraphs)
    noteNumbers: {},   // footnote id -> display number
    noteOrder: [],
    commentOrder: [],
    endnoteOrder: [],
  };

  const body = kid(main.documentElement, 'w', 'body');
  const ctx = { rels, toc: false, cover: false, hf: false };
  const blocks = parseContainer(body, ctx, st);

  // ---- headers & footers ----
  const hfBlocks = { header: [], footer: [] };
  for (const kind of ['header', 'footer']) {
    const seen = new Set();
    for (const p of partOf(kind)) {
      const d = await pkg.xml(p);
      if (!d) continue;
      const hctx = { rels: await pkg.rels(p), toc: false, cover: false, hf: kind };
      const parsed = parseContainer(d.documentElement, hctx, st);
      const t = blocksText(parsed).replace(/\s+/g, ' ').trim();
      if (!t) continue;
      if (opts.removeHeadersFooters) { exclude(kind === 'header' ? 'Header' : 'Footer', t); continue; }
      if (seen.has(t)) { exclude(kind === 'header' ? 'Header (repeated)' : 'Footer (repeated)', t); continue; }
      seen.add(t);
      hfBlocks[kind].push({ type: 'paragraph', style: 'quote', inlines: [{ type: 'text', text: `${kind === 'header' ? 'Header' : 'Footer'}: `, bold: true }, { type: 'text', text: t }] });
    }
  }

  // ---- footnotes, endnotes, comments ----
  const footnotes = [];
  const loadNotes = async (type, tag, order, prefix, category) => {
    const path = partOf(type)[0];
    const d = path && await pkg.xml(path);
    if (!d) return;
    const nctx = { rels: await pkg.rels(path), toc: false, cover: false, hf: false, note: true };
    const byId = {};
    for (const n of kids(d.documentElement, 'w', tag)) {
      const t = attr(n, 'w', 'type');
      if (t && t !== 'normal') continue;
      byId[attr(n, 'w', 'id')] = n;
    }
    if (opts[category === 'Comment' ? 'removeComments' : 'removeFootnotes']) {
      for (const n of Object.values(byId)) exclude(category, blocksText(parseContainer(n, nctx, st)));
      return;
    }
    for (const [id, num] of order) {
      const n = byId[id];
      if (!n) continue;
      let nb = parseContainer(n, nctx, st);
      if (category === 'Comment') {
        const author = attr(n, 'w', 'author');
        nb = [{ type: 'paragraph', inlines: [{ type: 'text', text: `Comment${author ? ' (' + author + ')' : ''}: ` }, ...(nb.flatMap(b => b.inlines ? [...b.inlines, { type: 'text', text: ' ' }] : []))] }];
      }
      footnotes.push({ id: prefix + num, blocks: nb });
      delete byId[id];
    }
    // notes that exist but are never referenced in the body
    for (const n of Object.values(byId)) {
      const t = blocksText(parseContainer(n, nctx, st));
      if (t.trim()) log.warnings.push(`An unreferenced ${category.toLowerCase()} was found and kept at the end: "${t.slice(0, 60)}"`), footnotes.push({ id: prefix + 'x' + (footnotes.length + 1), blocks: parseContainer(n, nctx, st) });
    }
  };
  await loadNotes('footnotes', 'footnote', st.noteOrder, '', 'Footnote');
  await loadNotes('endnotes', 'endnote', st.endnoteOrder, 'e', 'Footnote');
  await loadNotes('comments', 'comment', st.commentOrder, 'c', 'Comment');
  log.stats.footnotes = st.noteOrder.length + st.endnoteOrder.length;
  log.stats.comments = st.commentOrder.length;

  if (rels && Object.values(rels).some(r => r.type === 'aFChunk')) log.warnings.push('The document contains embedded alternative-format content (altChunk) that cannot be converted.');

  return {
    doc: { blocks: [...hfBlocks.header, ...blocks, ...hfBlocks.footer], footnotes },
    log,
  };
}

// ------------------------------------------------------------------ styles

async function loadStyles(pkg, path) {
  const map = {};
  const d = path && await pkg.xml(path);
  if (!d) return { get: () => null, chain: () => [] };
  for (const s of kids(d.documentElement, 'w', 'style')) {
    const id = attr(s, 'w', 'styleId');
    const pPr = kid(s, 'w', 'pPr');
    const rPr = kid(s, 'w', 'rPr');
    const numPr = kid(pPr, 'w', 'numPr');
    map[id] = {
      id,
      type: attr(s, 'w', 'type'),
      name: (wVal(kid(s, 'w', 'name')) || id || '').toLowerCase(),
      basedOn: wVal(kid(s, 'w', 'basedOn')),
      numId: numPr ? wVal(kid(numPr, 'w', 'numId')) : null,
      ilvl: numPr ? wVal(kid(numPr, 'w', 'ilvl')) : null,
      outline: wVal(kid(pPr, 'w', 'outlineLvl')),
      bold: kid(rPr, 'w', 'b') ? onOff(kid(rPr, 'w', 'b')) : null,
      italic: kid(rPr, 'w', 'i') ? onOff(kid(rPr, 'w', 'i')) : null,
    };
  }
  const chain = (id) => {
    const out = [];
    const seen = new Set();
    while (id && map[id] && !seen.has(id)) { seen.add(id); out.push(map[id]); id = map[id].basedOn; }
    return out;
  };
  return { get: (id) => map[id] || null, chain };
}

function styleProp(styles, id, prop) {
  for (const s of styles.chain(id)) if (s[prop] != null) return s[prop];
  return null;
}

// --------------------------------------------------------------- numbering

async function loadNumbering(pkg, path) {
  const abstracts = {}, nums = {};
  const d = path && await pkg.xml(path);
  if (d) {
    for (const a of kids(d.documentElement, 'w', 'abstractNum')) {
      const levels = {};
      for (const l of kids(a, 'w', 'lvl')) {
        levels[attr(l, 'w', 'ilvl')] = {
          fmt: wVal(kid(l, 'w', 'numFmt')) || 'decimal',
          text: wVal(kid(l, 'w', 'lvlText')) ?? '',
          start: parseInt(wVal(kid(l, 'w', 'start')) ?? '1', 10),
        };
      }
      abstracts[attr(a, 'w', 'abstractNumId')] = { levels };
    }
    for (const n of kids(d.documentElement, 'w', 'num')) {
      const overrides = {};
      for (const o of kids(n, 'w', 'lvlOverride')) {
        const so = kid(o, 'w', 'startOverride');
        if (so) overrides[attr(o, 'w', 'ilvl')] = parseInt(wVal(so), 10);
      }
      nums[attr(n, 'w', 'numId')] = { abstractId: wVal(kid(n, 'w', 'abstractNumId')), overrides };
    }
  }
  return { abstracts, nums };
}

function nextNumber(st, numId, ilvl) {
  const num = st.numbering.nums[numId];
  if (!num) return null;
  const abs = st.numbering.abstracts[num.abstractId];
  const lvl = abs?.levels[ilvl];
  if (!lvl) return null;
  const hasOverride = Object.keys(num.overrides).length > 0;
  const key = hasOverride ? 'n' + numId : 'a' + num.abstractId;
  const vals = st.counters[key] || (st.counters[key] = []);
  const startOf = (k) => num.overrides[k] ?? abs.levels[k]?.start ?? 1;
  vals[ilvl] = vals[ilvl] == null ? startOf(ilvl) : vals[ilvl] + 1;
  for (let k = +ilvl + 1; k < 10; k++) vals[k] = undefined;
  for (let k = 0; k < ilvl; k++) if (vals[k] == null) vals[k] = startOf(k);
  if (lvl.fmt === 'none') return { none: true, ordered: false, label: '' };
  if (lvl.fmt === 'bullet') return { ordered: false, label: '•' };
  const label = lvl.text.replace(/%(\d)/g, (m, d) => formatNumber(vals[d - 1] ?? startOf(d - 1), abs.levels[d - 1]?.fmt || 'decimal'));
  return { ordered: true, label };
}

export function formatNumber(n, fmt) {
  switch (fmt) {
    case 'lowerLetter': return letters(n).toLowerCase();
    case 'upperLetter': return letters(n);
    case 'lowerRoman': return roman(n).toLowerCase();
    case 'upperRoman': return roman(n);
    case 'decimalZero': return n < 10 ? '0' + n : String(n);
    case 'bullet': return '•';
    case 'none': return '';
    default: return String(n);
  }
}
function letters(n) {
  if (n < 1) return String(n);
  const ch = String.fromCharCode(65 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}
function roman(n) {
  if (n < 1 || n > 3999) return String(n);
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let s = '';
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return s;
}

// --------------------------------------------------------------- containers

function parseContainer(el, ctx, st) {
  const blocks = [];
  let pending = [];
  const flush = () => { if (pending.length) { blocks.push(...buildLists(pending)); pending = []; } };

  const visit = (child, c) => {
    if (!inNs(child, 'w')) {
      if (is(child, 'mc', 'AlternateContent')) {
        const choice = child.firstElementChild;
        if (choice) for (const g of choice.children) visit(g, c);
      }
      return;
    }
    switch (child.localName) {
      case 'p': {
        const res = parseParagraph(child, c, st);
        if (res.listItem) pending.push(res.listItem);
        else if (res.block) { flush(); blocks.push(res.block); }
        if (res.extra.length) { flush(); blocks.push(...res.extra); }
        break;
      }
      case 'tbl': {
        flush();
        const t = parseTable(child, c, st);
        if (t && isLayoutTable(t)) {
          // single-column tables are layout boxes (cover blocks, call-outs): keep their content as text
          st.log.stats.layoutTables = (st.log.stats.layoutTables || 0) + 1;
          st.log.stats.tables--;
          for (const row of t.rows) for (const cell of row) blocks.push(...cell.blocks);
        } else if (t) blocks.push(t);
        break;
      }
      case 'sdt': {
        const gallery = desc(kid(child, 'w', 'sdtPr'), 'w', 'docPartGallery').map(wVal).join(' ');
        const c2 = { ...c };
        if (/table of contents/i.test(gallery)) c2.toc = true;
        if (/cover page/i.test(gallery)) c2.cover = true;
        for (const g of kid(child, 'w', 'sdtContent')?.children || []) visit(g, c2);
        break;
      }
      case 'customXml': case 'ins': case 'moveTo': case 'smartTag':
        for (const g of child.children) visit(g, c);
        break;
      case 'del': case 'moveFrom':
        if (!st.opts.acceptChanges) for (const g of child.children) visit(g, c);
        break;
      case 'altChunk':
        st.log.warnings.push('Embedded alternative-format content (altChunk) was found and could not be converted.');
        break;
      default:
        break;
    }
  };
  for (const child of el.children) visit(child, ctx);
  flush();
  return blocks;
}

// --------------------------------------------------------------- paragraphs

function parseParagraph(p, ctx, st) {
  const { styles, opts } = st;
  const pPr = kid(p, 'w', 'pPr');
  const styleId = wVal(kid(pPr, 'w', 'pStyle')) || defaultParagraphStyle(styles);
  const styleName = styles.get(styleId)?.name || '';

  const tocBefore = st.fields.some(f => f.kind === 'TOC');
  const acc = { inlines: [], extra: [], tocField: tocBefore };
  parseInlineChildren(p, ctx, st, acc, {});
  const isToc = ctx.toc || acc.tocField || /^(toc \d|toc heading|table of figures)$/.test(styleName);

  let inlines = trimInlines(acc.inlines);
  const plain = inlinesText(inlines).replace(/\s+/g, ' ').trim();

  // numbering
  const numPr = kid(pPr, 'w', 'numPr');
  let numId = numPr ? wVal(kid(numPr, 'w', 'numId')) : null;
  let ilvl = numPr ? wVal(kid(numPr, 'w', 'ilvl')) : null;
  if (numId == null) numId = styleProp(styles, styleId, 'numId');
  if (ilvl == null) ilvl = styleProp(styles, styleId, 'ilvl') ?? '0';
  const num = numId && numId !== '0' ? nextNumber(st, numId, ilvl) : null;

  const res = { block: null, listItem: null, extra: acc.extra };

  if (ctx.cover && opts.removeCover) {
    if (plain) st.exclude('Cover page', plain);
    for (const b of acc.extra) st.exclude('Cover page', blocksText([b]));
    res.extra = [];
    return res;
  }
  if (isToc) {
    if (opts.removeToc) { if (plain) st.exclude('Table of contents', plain); return res; }
    if (!plain) return res;
    const m = /^toc (\d)$/.exec(styleName);
    res.listItem = { level: m ? +m[1] - 1 : 0, ordered: false, blocks: [{ type: 'paragraph', inlines: [{ type: 'text', text: plain }] }] };
    return res;
  }
  if (!plain) return res;

  // heading level
  let level = null;
  const hm = /^heading (\d)$/.exec(styleName);
  if (hm) level = +hm[1];
  else if (styleName === 'title') level = 1;
  else {
    const ol = wVal(kid(pPr, 'w', 'outlineLvl')) ?? styleProp(styles, styleId, 'outline');
    if (ol != null && +ol < 9) level = +ol + 1;
  }

  if (level) {
    st.log.stats.headings++;
    if (num && !num.none && num.label) inlines = [{ type: 'text', text: num.label.trim() + ' ' }, ...inlines];
    res.block = { type: 'heading', level: Math.min(level, 6), inlines };
    return res;
  }
  if (num && !num.none) {
    st.log.stats.listItems++;
    res.listItem = { level: +ilvl, ordered: num.ordered, label: num.ordered ? num.label : undefined, blocks: [{ type: 'paragraph', inlines }] };
    return res;
  }
  const block = { type: 'paragraph', inlines };
  if (styleName === 'caption') block.style = 'caption';
  else if (/quote/.test(styleName)) block.style = 'quote';
  res.block = block;
  return res;
}

function defaultParagraphStyle(styles) {
  return styles.get('Normal') ? 'Normal' : null;
}

function runFormat(rPr, st) {
  const fmt = {};
  if (!rPr) return fmt;
  const rStyle = wVal(kid(rPr, 'w', 'rStyle'));
  if (rStyle) {
    const b = styleProp(st.styles, rStyle, 'bold'); if (b) fmt.bold = true;
    const i = styleProp(st.styles, rStyle, 'italic'); if (i) fmt.italic = true;
  }
  const b = kid(rPr, 'w', 'b'); if (b) fmt.bold = onOff(b);
  const i = kid(rPr, 'w', 'i'); if (i) fmt.italic = onOff(i);
  if (onOff(kid(rPr, 'w', 'strike')) && kid(rPr, 'w', 'strike')) fmt.strike = true;
  const va = wVal(kid(rPr, 'w', 'vertAlign'));
  if (va === 'superscript') fmt.sup = true;
  if (va === 'subscript') fmt.sub = true;
  if (kid(rPr, 'w', 'vanish') && onOff(kid(rPr, 'w', 'vanish'))) fmt.hidden = true;
  if (!fmt.bold) delete fmt.bold;
  if (!fmt.italic) delete fmt.italic;
  return fmt;
}

function fieldKind(instr) {
  const m = /^\s*([A-Z]+)/i.exec(instr || '');
  return m ? m[1].toUpperCase() : '';
}

/** Walk the inline content of a paragraph (runs, hyperlinks, fields, revisions). */
function parseInlineChildren(el, ctx, st, acc, fmt) {
  for (const child of el.children) {
    if (is(child, 'mc', 'AlternateContent')) {
      const choice = child.firstElementChild;
      if (choice) parseInlineChildren(choice, ctx, st, acc, fmt);
      continue;
    }
    if (inNs(child, 'm')) {
      if (child.localName === 'oMath' || child.localName === 'oMathPara') {
        const t = desc(child, 'm', 't').map(e => e.textContent).join('');
        if (t) acc.inlines.push({ type: 'text', text: t, ...fmt });
      }
      continue;
    }
    if (!inNs(child, 'w')) continue;
    switch (child.localName) {
      case 'r': parseRun(child, ctx, st, acc, fmt); break;
      case 'hyperlink': {
        const id = attr(child, 'r', 'id');
        const rel = id && ctx.rels[id];
        parseInlineChildren(child, ctx, st, acc, rel && rel.external ? { ...fmt, href: rel.target } : fmt);
        break;
      }
      case 'fldSimple': {
        const kind = fieldKind(attr(child, 'w', 'instr'));
        if (kind === 'TOC') acc.tocField = true;
        parseInlineChildren(child, ctx, st, acc, isPageField(kind) ? { ...fmt, pageNo: true } : fmt);
        break;
      }
      case 'ins': case 'moveTo': case 'smartTag': case 'customXml': case 'dir': case 'bdo':
        parseInlineChildren(child, ctx, st, acc, fmt);
        break;
      case 'sdt':
        parseInlineChildren(kid(child, 'w', 'sdtContent') || child, ctx, st, acc, fmt);
        break;
      case 'del': case 'moveFrom':
        if (!st.opts.acceptChanges) parseInlineChildren(child, ctx, st, acc, { ...fmt, strike: true, deleted: true });
        break;
      default: break;
    }
  }
}

const isPageField = (k) => k === 'PAGE' || k === 'NUMPAGES' || k === 'SECTIONPAGES';

function parseRun(r, ctx, st, acc, outerFmt) {
  const fmt = { ...outerFmt, ...runFormat(kid(r, 'w', 'rPr'), st) };
  const push = (text) => {
    const top = st.fields[st.fields.length - 1];
    if (top && top.phase === 'code') return;           // field instructions are never content
    if (st.fields.some(f => (f.kind === 'ASK' || f.kind === 'SET') && f.phase === 'result')) { st.exclude('Hidden field value (ASK/SET)', text); return; }
    const pageNo = fmt.pageNo || st.fields.some(f => isPageField(f.kind) && f.phase === 'result');
    if (pageNo && ctx.hf && st.opts.removePageNumbers) { st.exclude('Page number', text); return; }
    if (fmt.hidden) { st.exclude('Hidden text', text); return; }
    const { hidden, pageNo: _p, deleted, ...f } = fmt;
    acc.inlines.push({ type: 'text', text, ...f });
  };
  for (const c of r.children) {
    if (is(c, 'mc', 'AlternateContent')) {
      const choice = c.firstElementChild;
      if (choice) for (const g of choice.children) if (is(g, 'w', 'drawing') || is(g, 'w', 'pict')) handleDrawing(g, ctx, st, acc);
      continue;
    }
    if (!inNs(c, 'w')) continue;
    switch (c.localName) {
      case 't': push(c.textContent); break;
      case 'delText': if (!st.opts.acceptChanges) push(c.textContent); break;
      case 'instrText': {
        const top = st.fields[st.fields.length - 1];
        if (top && top.phase === 'code') {
          top.instr += c.textContent;
          top.kind = fieldKind(top.instr);
          if (top.kind === 'TOC') acc.tocField = true;
        }
        break;
      }
      case 'fldChar': {
        const t = attr(c, 'w', 'fldCharType');
        if (t === 'begin') st.fields.push({ instr: '', kind: '', phase: 'code' });
        else if (t === 'separate') { const top = st.fields[st.fields.length - 1]; if (top) top.phase = 'result'; }
        else if (t === 'end') st.fields.pop();
        break;
      }
      case 'tab': case 'ptab': push('\t'); break;
      case 'br': {
        const type = attr(c, 'w', 'type');
        if (type === 'page' || type === 'column') push(' ');
        else acc.inlines.push({ type: 'break' });
        break;
      }
      case 'cr': acc.inlines.push({ type: 'break' }); break;
      case 'noBreakHyphen': push('-'); break;
      case 'softHyphen': break;
      case 'sym': {
        const code = parseInt(attr(c, 'w', 'char') || '', 16);
        if (!isNaN(code)) { const ch = symbolChar(code); if (ch) push(ch); }
        break;
      }
      case 'drawing': case 'pict': handleDrawing(c, ctx, st, acc); break;
      case 'object': {
        st.log.stats.images++;
        st.exclude('Embedded object', '');
        break;
      }
      case 'footnoteReference': case 'endnoteReference': {
        const id = attr(c, 'w', 'id');
        if (st.opts.removeFootnotes) break;
        const order = c.localName === 'footnoteReference' ? st.noteOrder : st.endnoteOrder;
        let entry = order.find(e => e[0] === id);
        if (!entry) { entry = [id, order.length + 1]; order.push(entry); }
        acc.inlines.push({ type: 'footref', id: (c.localName === 'endnoteReference' ? 'e' : '') + entry[1] });
        break;
      }
      case 'commentReference': {
        if (st.opts.removeComments) break;
        const id = attr(c, 'w', 'id');
        let entry = st.commentOrder.find(e => e[0] === id);
        if (!entry) { entry = [id, st.commentOrder.length + 1]; st.commentOrder.push(entry); }
        acc.inlines.push({ type: 'footref', id: 'c' + entry[1] });
        break;
      }
      default: break;
    }
  }
}

function handleDrawing(el, ctx, st, acc) {
  const boxes = desc(el, 'w', 'txbxContent');
  // keep only outermost text boxes (a text box may itself contain shapes)
  const outer = boxes.filter(b => !boxes.some(o => o !== b && o.contains(b)));
  if (outer.length) {
    for (const tb of outer) {
      st.log.stats.textBoxes++;
      acc.extra.push(...parseContainer(tb, ctx, st));
    }
    return;
  }
  const isChart = el.getElementsByTagNameNS('*', 'chart').length > 0;
  const docPr = desc(el, 'wp', 'docPr')[0];
  const alt = (docPr && (docPr.getAttribute('descr') || docPr.getAttribute('title'))) || '';
  if (isChart) st.log.stats.charts++; else st.log.stats.images++;
  const kind = isChart ? 'Chart' : 'Image';
  if (st.opts.removeImages) { st.exclude(kind, alt); return; }
  acc.extra.push({ type: 'image', alt: (isChart ? 'Chart' : '') + (alt ? (isChart ? ': ' : '') + alt : '') });
}

// --------------------------------------------------------------- tables

export function isLayoutTable(t) {
  return t.rows.every(r => r.reduce((s, c) => s + (c.colspan || 1), 0) <= 1);
}

function parseTable(tbl, ctx, st) {
  st.log.stats.tables++;
  const rows = [];
  const anchorAt = [];
  const trs = [];
  const collectRows = (el) => {
    for (const c of el.children) {
      if (is(c, 'w', 'tr')) trs.push(c);
      else if (is(c, 'w', 'sdt')) collectRows(kid(c, 'w', 'sdtContent') || c);
      else if (is(c, 'w', 'customXml') || is(c, 'w', 'ins')) collectRows(c);
    }
  };
  collectRows(tbl);

  for (const tr of trs) {
    const trPr = kid(tr, 'w', 'trPr');
    const header = !!kid(trPr, 'w', 'tblHeader') && onOff(kid(trPr, 'w', 'tblHeader'));
    const gridBefore = parseInt(wVal(kid(trPr, 'w', 'gridBefore')) || '0', 10);
    const row = [];
    let col = 0;
    if (gridBefore > 0) {
      row.push({ blocks: [], rowspan: 1, colspan: gridBefore, header: false, filler: true });
      for (let k = 0; k < gridBefore; k++) anchorAt[k] = null;
      col = gridBefore;
    }
    const tcs = [];
    const collectCells = (el) => {
      for (const c of el.children) {
        if (is(c, 'w', 'tc')) tcs.push(c);
        else if (is(c, 'w', 'sdt')) collectCells(kid(c, 'w', 'sdtContent') || c);
        else if (is(c, 'w', 'customXml')) collectCells(c);
      }
    };
    collectCells(tr);
    for (const tc of tcs) {
      const tcPr = kid(tc, 'w', 'tcPr');
      const span = Math.max(1, parseInt(wVal(kid(tcPr, 'w', 'gridSpan')) || '1', 10));
      const vm = kid(tcPr, 'w', 'vMerge');
      const vmVal = vm ? (wVal(vm) || 'continue') : null;
      const hm = kid(tcPr, 'w', 'hMerge');
      if (vmVal === 'continue' && anchorAt[col]) {
        anchorAt[col].rowspan = (anchorAt[col].rowspan || 1) + 1;
        // any content in a continued cell is appended to the merged cell
        const extra = parseContainer(tc, ctx, st);
        if (blocksText(extra).trim()) anchorAt[col].blocks.push(...extra);
        col += span;
        continue;
      }
      if (hm && (wVal(hm) || 'continue') === 'continue' && row.length) {
        row[row.length - 1].colspan += span;
        col += span;
        continue;
      }
      const cell = { blocks: parseContainer(tc, ctx, st), rowspan: 1, colspan: span, header };
      for (let k = col; k < col + span; k++) anchorAt[k] = cell;
      row.push(cell);
      col += span;
    }
    for (let k = col; k < anchorAt.length; k++) anchorAt[k] = null;
    rows.push(row);
  }
  if (ctx.cover && st.opts.removeCover) {
    st.exclude('Cover page', rows.map(r => r.map(c => blocksText(c.blocks)).join(' ')).join(' '));
    return null;
  }
  if (!rows.length) return null;
  return { type: 'table', rows };
}

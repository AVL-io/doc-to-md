// Document model -> Markdown text.
import { tableGrid, normalizeTable, hasMergedCells, mergeInlines, trimInlines, inlinesText } from './ast.js';

/**
 * @param {object} doc   document model
 * @param {object} opts  { tableMode: 'html' | 'flatten', sourceName }
 */
export function toMarkdown(doc, opts = {}) {
  const tableMode = opts.tableMode || 'html';
  const out = [];
  if (opts.sourceName) out.push(`<!-- Source: ${opts.sourceName.replace(/--/g, '-')} | Converted by doc-to-md -->`);
  writeBlocks(doc.blocks, out, '', tableMode);
  if (doc.footnotes && doc.footnotes.length) {
    const notes = doc.footnotes.map(fn => `[^${fn.id}]: ${blocksInlineMd(fn.blocks)}`);
    out.push(notes.join('\n'));
  }
  return out.join('\n\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

function writeBlocks(blocks, out, indent, tableMode) {
  for (const b of blocks) {
    const s = blockToMd(b, indent, tableMode);
    if (s != null && s.trim() !== '') out.push(s);
  }
}

function blockToMd(b, indent, tableMode) {
  switch (b.type) {
    case 'heading': {
      const t = escapeMd(inlinesText(trimInlines(b.inlines)).replace(/\s+/g, ' ').trim(), true);
      return t ? `${indent}${'#'.repeat(Math.min(6, Math.max(1, b.level)))} ${t}` : null;
    }
    case 'paragraph': {
      let t = inlinesToMd(trimInlines(b.inlines));
      if (!t.trim()) return null;
      t = escapeLineStarts(t);
      if (b.style === 'caption') t = `*${t}*`;
      const lines = t.split('\n');
      if (b.style === 'quote') return lines.map(l => `${indent}> ${l}`).join('\n');
      return lines.map(l => indent + l).join('\n');
    }
    case 'list':
      return listToMd(b, indent, tableMode);
    case 'table':
      return tableToMd(b, indent, tableMode);
    case 'code':
      return `${indent}\`\`\`\n${b.text.split('\n').map(l => indent + l).join('\n')}\n${indent}\`\`\``;
    case 'hr':
      return `${indent}---`;
    case 'image':
      return `${indent}*(Image${b.alt ? ': ' + escapeMd(b.alt.replace(/\s+/g, ' ').trim()) : ''})*`;
    case 'pagebreak':
      return `${indent}<!-- Page ${b.page} -->`;
    default:
      return null;
  }
}

function listToMd(list, indent, tableMode) {
  const lines = [];
  let n = list.start || 1;
  for (const item of list.items) {
    let marker;
    if (list.ordered) {
      const lab = (item.label || '').trim();
      const m = /^(\d+)[.)]?$/.exec(lab);
      if (m) { marker = `${m[1]}.`; n = +m[1] + 1; }
      else if (lab) marker = `- ${escapeMd(lab)}`;
      else marker = `${n++}.`;
    } else marker = '-';
    const childIndent = indent + ' '.repeat(marker.startsWith('- ') ? 2 : marker.length + 1);
    const blocks = item.blocks || [];
    let first = true;
    for (const b of blocks) {
      if (first && b.type === 'paragraph') {
        const t = inlinesToMd(trimInlines(b.inlines)).split('\n');
        lines.push(`${indent}${marker} ${t[0]}`);
        for (const l of t.slice(1)) lines.push(childIndent + l);
        first = false;
        continue;
      }
      if (first) { lines.push(`${indent}${marker}`); first = false; }
      const s = blockToMd(b, childIndent, tableMode);
      if (s) lines.push(b.type === 'list' ? s : '\n' + s);
    }
    if (first) lines.push(`${indent}${marker}`);
  }
  return lines.join('\n');
}

// ---------- tables ----------

function tableToMd(table, indent, tableMode) {
  table = normalizeTable(table);
  if (!table.rows.length) return null;
  const simple = !hasMergedCells(table) && table.rows.every(r => r.every(c => isSimpleCell(c)));
  if (tableMode === 'flatten' || simple) return pipeTable(table, indent);
  return htmlTable(table, indent);
}

function isSimpleCell(cell) {
  const bl = cell.blocks || [];
  return bl.length <= 1 && bl.every(b => b.type === 'paragraph' || b.type === 'heading');
}

function pipeTable(table, indent) {
  const { grid, nRows, nCols } = tableGrid(table);
  if (!nCols) return null;
  const rows = [];
  for (let r = 0; r < nRows; r++) {
    const cells = [];
    for (let c = 0; c < nCols; c++) {
      const cell = grid[r][c].cell;
      cells.push(cellToPipe(r === 0 || cell.header ? stripEmphasis(cell) : cell));
    }
    rows.push(cells);
  }
  const line = cells => `${indent}| ${cells.join(' | ')} |`;
  const out = [line(rows[0]), `${indent}|${' --- |'.repeat(nCols)}`];
  for (const r of rows.slice(1)) out.push(line(r));
  return out.join('\n');
}

function cellToPipe(cell) {
  const parts = cellParts(cell, 'md');
  return parts.join('<br>').replace(/\|/g, '\\|').replace(/\n/g, '<br>').trim();
}

function htmlTable(table, indent) {
  const out = [`${indent}<table>`];
  for (const row of table.rows) {
    const cells = row.map(c => {
      const tag = c.header ? 'th' : 'td';
      const attrs = (c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : '') + (c.colspan > 1 ? ` colspan="${c.colspan}"` : '');
      const content = cellParts(c.header ? stripEmphasis(c) : c, 'html').join('<br>');
      return `<${tag}${attrs}>${content}</${tag}>`;
    });
    out.push(`${indent}<tr>${cells.join('')}</tr>`);
  }
  out.push(`${indent}</table>`);
  return out.join('\n');
}

/** Header cells are already emphasised by <th>; drop bold/italic inside them. */
function stripEmphasis(cell) {
  const strip = (blocks) => blocks.map(b => b.inlines ? { ...b, inlines: b.inlines.map(r => r.type === 'text' ? { ...r, bold: false, italic: false } : r) }
    : b.type === 'list' ? { ...b, items: b.items.map(it => ({ ...it, blocks: strip(it.blocks) })) } : b);
  return { ...cell, blocks: strip(cell.blocks || []) };
}

/** Flatten cell blocks into display lines, in Markdown or HTML inline syntax. */
function cellParts(cell, mode, depth = 0) {
  const parts = [];
  const inl = mode === 'html' ? inlinesToHtml : inlinesToMd;
  for (const b of cell.blocks || []) {
    if (b.type === 'paragraph' || b.type === 'heading') {
      const t = inl(trimInlines(b.inlines)).replace(/\n/g, '<br>');
      if (t.trim()) parts.push(t);
    } else if (b.type === 'list') {
      let n = 1;
      for (const it of b.items) {
        const pad = mode === 'html' ? '&nbsp;&nbsp;'.repeat(depth) : '  '.repeat(depth);
        const lab = b.ordered ? (it.label || `${n}.`) : (depth ? '◦' : '•');
        n++;
        const sub = cellParts({ blocks: it.blocks }, mode, depth + 1);
        if (sub.length) {
          parts.push(`${pad}${mode === 'html' ? escapeHtml(lab) : lab} ${sub[0]}`);
          parts.push(...sub.slice(1));
        }
      }
    } else if (b.type === 'table') {
      if (mode === 'html') parts.push(htmlTable(normalizeTable(b), '').replace(/\n/g, ''));
      else {
        const { grid, nRows, nCols } = tableGrid(b);
        for (let r = 0; r < nRows; r++) {
          const row = [];
          for (let c = 0; c < nCols; c++) row.push(cellParts(grid[r][c].cell, mode, depth).join(' '));
          parts.push(row.join(' ; '));
        }
      }
    } else if (b.type === 'image') {
      parts.push(`(Image${b.alt ? ': ' + (mode === 'html' ? escapeHtml(b.alt) : escapeMd(b.alt)) : ''})`);
    } else if (b.type === 'code') {
      parts.push(mode === 'html' ? `<code>${escapeHtml(b.text)}</code>` : '`' + b.text + '`');
    }
  }
  return parts;
}

// ---------- inline rendering ----------

export function inlinesToMd(inlines) {
  let s = '';
  for (const r of mergeInlines(inlines)) {
    if (r.type === 'break') { s += '  \n'; continue; }
    if (r.type === 'footref') { s += `[^${r.id}]`; continue; }
    if (r.type !== 'text') continue;
    if (r.code) { s += codeSpan(r.text); continue; }
    let t = escapeMd(r.text.replace(/\n/g, ' '));
    if (r.sup) t = `<sup>${t}</sup>`;
    if (r.sub) t = `<sub>${t}</sub>`;
    t = wrapMarker(t, r.bold && r.italic ? '***' : r.bold ? '**' : r.italic ? '*' : '');
    if (r.strike) t = wrapMarker(t, '~~');
    if (r.href && /^(https?:|mailto:)/i.test(r.href)) t = `[${t}](${r.href.replace(/[()\s]/g, encodeURIComponent)})`;
    s += t;
  }
  return s;
}

function wrapMarker(t, m) {
  if (!m || !t.trim()) return t;
  const lead = t.match(/^\s*/)[0], trail = t.match(/\s*$/)[0];
  return `${lead}${m}${t.trim()}${m}${trail}`;
}

function codeSpan(t) {
  const ticks = (t.match(/`+/g) || []).reduce((a, b) => Math.max(a, b.length), 0);
  const f = '`'.repeat(ticks + 1);
  return ticks ? `${f} ${t} ${f}` : `${f}${t}${f}`;
}

export function inlinesToHtml(inlines) {
  let s = '';
  for (const r of mergeInlines(inlines)) {
    if (r.type === 'break') { s += '<br>'; continue; }
    if (r.type === 'footref') { s += `[^${escapeHtml(String(r.id))}]`; continue; }
    if (r.type !== 'text') continue;
    let t = escapeHtml(r.text.replace(/\n/g, ' '));
    if (r.code) t = `<code>${t}</code>`;
    if (r.sup) t = `<sup>${t}</sup>`;
    if (r.sub) t = `<sub>${t}</sub>`;
    if (r.bold && t.trim()) t = `<b>${t}</b>`;
    if (r.italic && t.trim()) t = `<i>${t}</i>`;
    if (r.strike && t.trim()) t = `<s>${t}</s>`;
    if (r.href && /^(https?:|mailto:)/i.test(r.href)) t = `<a href="${escapeHtml(r.href)}">${t}</a>`;
    s += t;
  }
  return s;
}

function blocksInlineMd(blocks) {
  return (blocks || []).map(b => b.inlines ? inlinesToMd(trimInlines(b.inlines)) : '').filter(Boolean).join(' ').replace(/\n/g, ' ');
}

// ---------- escaping ----------

export function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const WORD = /[\p{L}\p{N}]/u;
export function escapeMd(s, heading = false) {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\\' || ch === '*' || ch === '`' || ch === '[' || ch === ']') out += '\\' + ch;
    else if (ch === '_') {
      const a = s[i - 1], b = s[i + 1];
      out += (a && WORD.test(a) && b && WORD.test(b)) ? '_' : '\\_';
    } else if (ch === '<' && /[A-Za-z/!?]/.test(s[i + 1] || '')) out += '&lt;';
    else if (heading && ch === '#' && i === s.length - 1) out += '\\#';
    else out += ch;
  }
  return out;
}

/** Prevent paragraph lines from being read as headings, lists, quotes or tables. */
function escapeLineStarts(t) {
  return t.split('\n').map(line =>
    line.replace(/^(\s*)(#{1,6}(?=\s|$)|>|[-+](?=\s)|=+\s*$|-{3,}\s*$)/, (m, sp, tok) => sp + '\\' + tok)
        .replace(/^(\s*)(\d{1,9})([.)])(?=\s)/, '$1$2\\$3')
        .replace(/^(\s*)\|/, '$1\\|')
  ).join('\n');
}

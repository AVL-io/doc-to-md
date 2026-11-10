// Shared document model used by every reader and writer.
//
// Document: { blocks: Block[], footnotes: {id, blocks}[], meta: {...} }
//
// Block types:
//   heading   { level 1-6, inlines }
//   paragraph { inlines, style?: 'caption' | 'quote' }
//   list      { ordered, items: [{ label?, blocks }] }
//   table     { rows: Cell[][], meta? }   rows hold anchor cells only (HTML model)
//   code      { text }
//   hr        {}
//   image     { alt }
//   pagebreak { page }        page marker (PDF only)
//
// Cell:   { blocks, rowspan, colspan, header }
// Inline: { type:'text', text, bold?, italic?, strike?, code?, sup?, sub?, href? }
//         { type:'break' }
//         { type:'footref', id }

export const text = (t, fmt = {}) => ({ type: 'text', text: t, ...fmt });

/** Merge adjacent text runs that share the same formatting. */
export function mergeInlines(inlines) {
  const out = [];
  for (const r of inlines) {
    if (r.type === 'text' && r.text === '') continue;
    const prev = out[out.length - 1];
    if (prev && prev.type === 'text' && r.type === 'text' && sameFormat(prev, r)) {
      prev.text += r.text;
    } else {
      out.push({ ...r });
    }
  }
  return out;
}

function sameFormat(a, b) {
  return !!a.bold === !!b.bold && !!a.italic === !!b.italic && !!a.strike === !!b.strike &&
    !!a.code === !!b.code && !!a.sup === !!b.sup && !!a.sub === !!b.sub && (a.href || '') === (b.href || '');
}

/** Trim leading/trailing whitespace of an inline sequence. */
export function trimInlines(inlines) {
  const out = mergeInlines(inlines);
  while (out.length && out[0].type === 'break') out.shift();
  while (out.length && out[out.length - 1].type === 'break') out.pop();
  if (out.length && out[0].type === 'text') {
    out[0].text = out[0].text.replace(/^\s+/, '');
    if (!out[0].text) out.shift();
  }
  const last = out[out.length - 1];
  if (last && last.type === 'text') {
    last.text = last.text.replace(/\s+$/, '');
    if (!last.text) out.pop();
  }
  return out;
}

export function inlinesText(inlines) {
  return (inlines || []).map(r => r.type === 'text' ? r.text : r.type === 'break' ? '\n' : '').join('');
}

export function blocksText(blocks) {
  const parts = [];
  for (const b of blocks || []) {
    if (b.inlines) parts.push(inlinesText(b.inlines));
    else if (b.type === 'list') b.items.forEach(it => parts.push(blocksText(it.blocks)));
    else if (b.type === 'table') b.rows.forEach(r => r.forEach(c => parts.push(blocksText(c.blocks))));
    else if (b.type === 'code') parts.push(b.text);
    else if (b.type === 'image') parts.push(b.alt || '');
  }
  return parts.join('\n');
}

/**
 * Turn a flat sequence of list paragraphs into nested list blocks.
 * items: [{ level, ordered, label?, blocks }]
 * Returns an array: a bulleted list directly followed by a numbered list (or
 * the reverse) at the outer level becomes two separate lists.
 */
export function buildLists(flat) {
  const roots = [];
  const minLevel = Math.min(...flat.map(it => it.level));
  let root = null;
  let stack = [];
  for (const it of flat) {
    if (!root || (it.level <= minLevel && it.ordered !== root.ordered)) {
      root = { type: 'list', ordered: it.ordered, items: [] };
      roots.push(root);
      stack = [{ list: root, level: minLevel }];
    }
    let top = stack[stack.length - 1];
    while (it.level < top.level && stack.length > 1) { stack.pop(); top = stack[stack.length - 1]; }
    if (it.level > top.level && top.list.items.length) {
      const parentItem = top.list.items[top.list.items.length - 1];
      const sub = { type: 'list', ordered: it.ordered, items: [] };
      parentItem.blocks.push(sub);
      stack.push({ list: sub, level: it.level });
      top = stack[stack.length - 1];
    }
    top.list.items.push({ label: it.label, blocks: it.blocks });
  }
  return roots;
}

/**
 * Compute the full grid of a table. Holes are filled with empty cells so every
 * writer can rely on a rectangular grid.
 * Returns { grid: (cellRef)[][], nRows, nCols } where cellRef = { cell, r, c, anchor }.
 */
export function tableGrid(table) {
  const occ = [];
  let nCols = 0;
  table.rows.forEach((row, r) => {
    occ[r] = occ[r] || [];
    let c = 0;
    for (const cell of row) {
      while (occ[r][c]) c++;
      const rs = Math.max(1, cell.rowspan || 1), cs = Math.max(1, cell.colspan || 1);
      for (let i = 0; i < rs; i++) {
        occ[r + i] = occ[r + i] || [];
        for (let j = 0; j < cs; j++) occ[r + i][c + j] = { cell, r, c, anchor: i === 0 && j === 0 };
      }
      c += cs;
      nCols = Math.max(nCols, c);
    }
  });
  // Rows that exist only because of a row-span overhang are clipped.
  const nRows = table.rows.length;
  for (let r = 0; r < nRows; r++) {
    for (let c = 0; c < nCols; c++) {
      if (!occ[r][c]) {
        const filler = { blocks: [], rowspan: 1, colspan: 1, header: false, filler: true };
        occ[r][c] = { cell: filler, r, c, anchor: true };
      }
    }
  }
  return { grid: occ.slice(0, nRows), nRows, nCols };
}

/**
 * Return a copy of the table in which spans are clipped to the table and holes
 * are filled, so the anchor-cell representation is always consistent.
 */
export function normalizeTable(table) {
  const { grid, nRows, nCols } = tableGrid(table);
  const rows = [];
  for (let r = 0; r < nRows; r++) {
    rows[r] = [];
    for (let c = 0; c < nCols; c++) {
      const ref = grid[r][c];
      if (!ref.anchor || ref.r !== r || ref.c !== c) continue;
      const cell = { ...ref.cell };
      cell.rowspan = Math.min(cell.rowspan || 1, nRows - r);
      cell.colspan = Math.min(cell.colspan || 1, nCols - c);
      rows[r].push(cell);
    }
  }
  return { ...table, rows };
}

export function hasMergedCells(table) {
  return table.rows.some(r => r.some(c => (c.rowspan || 1) > 1 || (c.colspan || 1) > 1));
}

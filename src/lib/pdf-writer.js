// Document model -> PDF, using pdfmake (runs entirely in the browser).
import pdfMake from 'pdfmake/build/pdfmake';
import pdfFonts from 'pdfmake/build/vfs_fonts';
import { tableGrid, normalizeTable } from './ast.js';

const vfs = (pdfFonts && (pdfFonts.pdfMake?.vfs || pdfFonts.vfs)) || pdfFonts;
pdfMake.vfs = vfs;

const SIZES = [20, 16, 14, 12.5, 11.5, 11];

export function writePdf(doc, { title = 'Document' } = {}) {
  const inl = (inlines, extra = {}) => (inlines || []).map(r => {
    if (r.type === 'break') return { text: '\n' };
    if (r.type === 'footref') return { text: `[${r.id}]`, sup: true };
    if (r.type !== 'text') return { text: '' };
    const o = { text: r.text };
    if (r.bold || extra.bold) o.bold = true;
    if (r.italic || extra.italic) o.italics = true;
    if (r.strike) o.decoration = 'lineThrough';
    if (r.sup) o.sup = true;
    if (r.sub) o.sub = true;
    if (r.code) o.background = '#f2f2f2';
    if (r.href && /^(https?:|mailto:)/i.test(r.href)) { o.link = r.href; o.color = '#1a5fb4'; }
    return o;
  });

  const conv = (blocks) => {
    const out = [];
    for (const b of blocks) {
      switch (b.type) {
        case 'heading':
          out.push({ text: inl(b.inlines), fontSize: SIZES[Math.min(5, b.level - 1)], bold: true, margin: [0, b.level <= 2 ? 12 : 8, 0, 4], headlineLevel: b.level });
          break;
        case 'paragraph':
          out.push({ text: inl(b.inlines, b.style === 'caption' ? { italic: true } : {}), margin: [b.style === 'quote' ? 18 : 0, 0, 0, 6] });
          break;
        case 'list': out.push(list(b)); break;
        case 'table': out.push(table(b)); break;
        case 'code': out.push({ text: b.text, background: '#f2f2f2', fontSize: 9, margin: [0, 0, 0, 6], preserveLeadingSpaces: true }); break;
        case 'hr': out.push({ canvas: [{ type: 'line', x1: 0, y1: 0, x2: 515, y2: 0, lineWidth: 0.5, lineColor: '#999999' }], margin: [0, 6, 0, 6] }); break;
        case 'image': out.push({ text: `(Image${b.alt ? ': ' + b.alt : ''})`, italics: true, color: '#666666', margin: [0, 0, 0, 6] }); break;
        default: break;
      }
    }
    return out;
  };

  const list = (b) => {
    const items = b.items.map(it => {
      const parts = conv(it.blocks);
      if (b.ordered && it.label && !/^\d+[.)]?$/.test(it.label) && parts[0] && parts[0].text) parts[0] = { ...parts[0], text: [{ text: it.label + ' ' }, ...[].concat(parts[0].text)] };
      parts.forEach(p => { if (p.margin) p.margin = [p.margin[0], 0, 0, 2]; });
      return parts.length === 1 ? parts[0] : { stack: parts };
    });
    return b.ordered ? { ol: items, start: b.start || 1, margin: [0, 0, 0, 6] } : { ul: items, margin: [0, 0, 0, 6] };
  };

  const table = (t) => {
    const norm = normalizeTable(t);
    const { grid, nRows, nCols } = tableGrid(norm);
    const body = [];
    for (let r = 0; r < nRows; r++) {
      const row = [];
      for (let c = 0; c < nCols; c++) {
        const ref = grid[r][c];
        if (ref.r !== r || ref.c !== c) { row.push({}); continue; }
        const content = conv(ref.cell.blocks || []);
        content.forEach(p => { if (p.margin) p.margin = [0, 0, 0, 2]; });
        const cell = content.length === 1 ? { ...content[0] } : { stack: content.length ? content : [{ text: '' }] };
        if ((ref.cell.rowspan || 1) > 1) cell.rowSpan = ref.cell.rowspan;
        if ((ref.cell.colspan || 1) > 1) cell.colSpan = ref.cell.colspan;
        if (ref.cell.header) { cell.fillColor = '#e7ecf2'; cell.bold = true; }
        row.push(cell);
      }
      body.push(row);
    }
    let headerRows = 0;
    while (headerRows < norm.rows.length && norm.rows[headerRows].length && norm.rows[headerRows].every(c => c.header)) headerRows++;
    if (!body.length) return { text: '' };
    return { table: { headerRows, widths: Array(nCols).fill('*'), body, dontBreakRows: false }, layout: { hLineColor: () => '#999999', vLineColor: () => '#999999', hLineWidth: () => 0.5, vLineWidth: () => 0.5 }, fontSize: 9, margin: [0, 2, 0, 10] };
  };

  const content = conv(doc.blocks.filter(b => b.type !== 'pagebreak'));
  if (doc.footnotes && doc.footnotes.length) {
    content.push({ canvas: [{ type: 'line', x1: 0, y1: 0, x2: 150, y2: 0, lineWidth: 0.5 }], margin: [0, 12, 0, 4] });
    for (const fn of doc.footnotes) content.push({ text: [{ text: `[${fn.id}] `, bold: true }, ...fn.blocks.flatMap(b => inl(b.inlines))], fontSize: 8.5, margin: [0, 0, 0, 2] });
  }
  const dd = {
    info: { title, creator: 'doc-to-md' },
    pageSize: 'A4',
    pageMargins: [40, 48, 40, 48],
    defaultStyle: { fontSize: 10.5, lineHeight: 1.2 },
    content,
    footer: (page, pages) => ({ text: `${page} / ${pages}`, alignment: 'center', fontSize: 8, color: '#888888', margin: [0, 16, 0, 0] }),
  };
  return new Promise((resolve, reject) => {
    try { pdfMake.createPdf(dd).getBlob(resolve); } catch (e) { reject(e); }
  });
}

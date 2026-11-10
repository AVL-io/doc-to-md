// Document model -> Word (.docx), using the "docx" library.
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  ExternalHyperlink, AlignmentType, LevelFormat, BorderStyle, ShadingType, FootnoteReferenceRun,
} from 'docx';
import { normalizeTable, inlinesText } from './ast.js';

const HEADINGS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];

export async function writeDocx(doc, { title = 'Document' } = {}) {
  let listInstance = 0;
  const footnoteIds = new Map();
  const footnotes = {};
  (doc.footnotes || []).forEach((fn, i) => {
    footnoteIds.set(String(fn.id), i + 1);
    footnotes[i + 1] = { children: fn.blocks.filter(b => b.inlines).map(b => new Paragraph({ children: runs(b.inlines) })) };
  });

  function runs(inlines, extra = {}) {
    const out = [];
    for (const r of inlines || []) {
      if (r.type === 'break') { out.push(new TextRun({ text: '', break: 1 })); continue; }
      if (r.type === 'footref') {
        const id = footnoteIds.get(String(r.id));
        out.push(id ? new FootnoteReferenceRun(id) : new TextRun({ text: `[${r.id}]`, superScript: true }));
        continue;
      }
      if (r.type !== 'text') continue;
      const opts = {
        text: r.text, bold: r.bold || extra.bold, italics: r.italic || extra.italic, strike: r.strike,
        superScript: r.sup, subScript: r.sub, font: r.code ? 'Consolas' : undefined,
      };
      if (r.href && /^(https?:|mailto:)/i.test(r.href)) out.push(new ExternalHyperlink({ link: r.href, children: [new TextRun({ ...opts, style: 'Hyperlink' })] }));
      else out.push(new TextRun(opts));
    }
    return out;
  }

  function blocks(list, depth = 0) {
    const out = [];
    for (const b of list) {
      switch (b.type) {
        case 'heading':
          out.push(new Paragraph({ heading: HEADINGS[Math.min(5, b.level - 1)], children: runs(b.inlines) }));
          break;
        case 'paragraph':
          out.push(new Paragraph({
            children: runs(b.inlines, b.style === 'caption' ? { italic: true } : {}),
            indent: b.style === 'quote' ? { left: 720 } : undefined,
            style: b.style === 'caption' ? 'Caption' : undefined,
          }));
          break;
        case 'list': out.push(...listBlocks(b, depth)); break;
        case 'table': out.push(tableBlock(b), new Paragraph({ children: [] })); break;
        case 'code':
          for (const line of b.text.split('\n')) out.push(new Paragraph({ children: [new TextRun({ text: line, font: 'Consolas', size: 19 })], shading: { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' }, spacing: { after: 0 } }));
          break;
        case 'hr':
          out.push(new Paragraph({ children: [], border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } } }));
          break;
        case 'image':
          out.push(new Paragraph({ children: [new TextRun({ text: `(Image${b.alt ? ': ' + b.alt : ''})`, italics: true, color: '666666' })] }));
          break;
        default: break;
      }
    }
    return out;
  }

  function listBlocks(list, depth) {
    const out = [];
    const instance = ++listInstance;
    for (const item of list.items) {
      let first = true;
      for (const b of item.blocks) {
        if (b.type === 'list') { out.push(...listBlocks(b, depth + 1)); continue; }
        if (first && (b.type === 'paragraph' || b.type === 'heading')) {
          const label = list.ordered && item.label && !/^\d+[.)]?$/.test(item.label) ? [new TextRun({ text: item.label + ' ' })] : [];
          out.push(new Paragraph({
            numbering: { reference: list.ordered ? 'numbers' : 'bullets', level: Math.min(depth, 8), instance },
            children: [...label, ...runs(b.inlines)],
          }));
          first = false;
        } else {
          for (const p of blocks([b], depth)) out.push(p);
        }
      }
    }
    return out;
  }

  function tableBlock(t) {
    const table = normalizeTable(t);
    const nCols = Math.max(1, ...table.rows.map(r => r.reduce((s, c) => s + (c.colspan || 1), 0)));
    const headerRows = [];
    for (const r of table.rows) { if (r.length && r.every(c => c.header)) headerRows.push(r); else break; }
    const rows = table.rows.map((r, ri) => new TableRow({
      tableHeader: ri < headerRows.length,
      children: r.map(c => {
        const content = blocks(c.blocks || []).filter(x => x instanceof Paragraph || x instanceof Table);
        return new TableCell({
          children: content.length ? content : [new Paragraph({ children: [] })],
          rowSpan: c.rowspan > 1 ? c.rowspan : undefined,
          columnSpan: c.colspan > 1 ? c.colspan : undefined,
          shading: c.header ? { type: ShadingType.CLEAR, fill: 'E7ECF2', color: 'auto' } : undefined,
        });
      }),
    }));
    return new Table({
      rows: rows.length ? rows : [new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [] })] })] })],
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: Array(nCols).fill(Math.floor(9000 / nCols)),
    });
  }

  const levels = (fmt) => Array.from({ length: 9 }, (_, i) => ({
    level: i,
    format: fmt === 'bullet' ? LevelFormat.BULLET : [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN][i % 3],
    text: fmt === 'bullet' ? ['•', '◦', '▪'][i % 3] : `%${i + 1}.`,
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 360 * (i + 1), hanging: 360 } } },
  }));

  const d = new Document({
    title,
    creator: 'doc-to-md',
    footnotes,
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    numbering: { config: [{ reference: 'bullets', levels: levels('bullet') }, { reference: 'numbers', levels: levels('number') }] },
    sections: [{ children: blocks(doc.blocks.filter(b => b.type !== 'pagebreak')) }],
  });
  void inlinesText;
  return Packer.toBlob(d);
}

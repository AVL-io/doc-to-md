// Document model -> Excel (.xlsx). Each table becomes a worksheet (merged cells
// preserved); text outside tables goes to a "Document text" sheet.
import JSZip from 'jszip';
import { normalizeTable, tableGrid, blocksText, inlinesText } from './ast.js';
import { encodeCol } from './xlsx-reader.js';

const xmlEsc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export async function writeXlsx(doc) {
  const sheets = [];
  const textRows = [];
  const used = new Set();
  let lastHeading = '';
  let tableNo = 0;

  const sheetName = (base) => {
    let n = (base || 'Table').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 28) || 'Table';
    let name = n, k = 2;
    while (used.has(name.toLowerCase())) name = `${n.slice(0, 26)} ${k++}`;
    used.add(name.toLowerCase());
    return name;
  };
  used.add('document text');

  const walk = (blocks, depth = 0) => {
    for (const b of blocks) {
      if (b.type === 'heading') { lastHeading = inlinesText(b.inlines).trim(); textRows.push({ text: lastHeading, bold: true }); }
      else if (b.type === 'paragraph') { const t = inlinesText(b.inlines).trim(); if (t) textRows.push({ text: t }); }
      else if (b.type === 'list') b.items.forEach((it, i) => {
        const lab = b.ordered ? (it.label || `${(b.start || 1) + i}.`) : '•';
        const [first, ...rest] = it.blocks;
        if (first && first.inlines) textRows.push({ text: `${'  '.repeat(depth)}${lab} ${inlinesText(first.inlines).trim()}` });
        walk(first && first.inlines ? rest : it.blocks, depth + 1);
      });
      else if (b.type === 'code') textRows.push({ text: b.text });
      else if (b.type === 'table') {
        tableNo++;
        const name = sheetName(lastHeading ? `${tableNo} ${lastHeading}` : `Table ${tableNo}`);
        textRows.push({ text: `[Table ${tableNo}: see sheet "${name}"]`, italic: true });
        sheets.push({ name, table: normalizeTable(b) });
      }
    }
  };
  walk(doc.blocks);

  const all = [];
  if (textRows.length) all.push({ name: 'Document text', text: textRows });
  all.push(...sheets);
  if (!all.length) all.push({ name: 'Document text', text: [{ text: '' }] });

  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${all.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${all.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${all.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${all.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  // styles: 0 default, 1 bold+wrap (header), 2 wrap, 3 italic
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><i/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE7ECF2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf/><xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf borderId="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf fontId="2" applyFont="1"/></cellXfs></styleSheet>`);

  all.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, s.table ? tableSheet(s.table) : textSheet(s.text)));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function cellXml(ref, text, style) {
  const s = style ? ` s="${style}"` : '';
  if (text === '' || text == null) return style ? `<c r="${ref}"${s}/>` : '';
  if (/^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(text)) return `<c r="${ref}"${s}><v>${text}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(text)}</t></is></c>`;
}

function tableSheet(table) {
  const { grid, nRows, nCols } = tableGrid(table);
  const widths = Array(nCols).fill(8);
  const rows = [];
  const merges = [];
  for (let r = 0; r < nRows; r++) {
    let cells = '';
    for (let c = 0; c < nCols; c++) {
      const ref = grid[r][c];
      const isAnchor = ref.r === r && ref.c === c;
      const text = isAnchor ? blocksText(ref.cell.blocks).replace(/\n+/g, '\n').trim() : '';
      if (isAnchor && ((ref.cell.rowspan || 1) > 1 || (ref.cell.colspan || 1) > 1)) {
        merges.push(`${encodeCol(c)}${r + 1}:${encodeCol(c + (ref.cell.colspan || 1) - 1)}${r + 1 + (ref.cell.rowspan || 1) - 1}`);
      }
      if (isAnchor && (ref.cell.colspan || 1) === 1) widths[c] = Math.max(widths[c], Math.min(60, Math.max(...text.split('\n').map(l => l.length)) + 2));
      cells += cellXml(`${encodeCol(c)}${r + 1}`, text, ref.cell.header ? 1 : 2);
    }
    rows.push(`<row r="${r + 1}">${cells}</row>`);
  }
  const cols = widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rows.join('')}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}</worksheet>`;
}

function textSheet(lines) {
  const rows = lines.map((l, i) => `<row r="${i + 1}">${l.text ? `<c r="A${i + 1}"${l.bold ? ' s="1"' : l.italic ? ' s="3"' : ''} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(l.text)}</t></is></c>` : ''}</row>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="120" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData></worksheet>`;
}

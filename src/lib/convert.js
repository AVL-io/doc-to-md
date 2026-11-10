// Orchestration: file in -> converted output + verification report.

export const TO_MD_EXT = ['docx', 'docm', 'xlsx', 'xlsm', 'pdf'];
export const FROM_MD_EXT = ['md', 'markdown', 'txt'];

export const DEFAULT_OPTIONS = {
  tableMode: 'html',
  removeImages: true,
  removeHeadersFooters: true,
  removePageNumbers: true,
  removeFootnotes: false,
  removeComments: true,
  acceptChanges: true,
  removeToc: true,
  removeCover: false,
  pageMarkers: true,
  includeHiddenSheets: false,
  includeHiddenRowsCols: false,
};

export function extOf(name) {
  const m = /\.([^.]+)$/.exec(name || '');
  return m ? m[1].toLowerCase() : '';
}

/** Identify the file type from its content (falls back to the extension). */
export function detectKind(name, buf) {
  const b = new Uint8Array(buf.slice(0, 8));
  const ext = extOf(name);
  if (b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46) return 'pdf';
  if (b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0) {
    throw new Error(ext === 'xls' || ext === 'xlsx' ? 'This is an old-format Excel file (.xls) or a password-protected workbook. Save it as .xlsx in Excel and try again.'
      : 'This is an old-format Word file (.doc) or a password-protected document. Save it as .docx in Word and try again.');
  }
  if (b[0] === 0x50 && b[1] === 0x4B) {
    if (ext === 'xlsx' || ext === 'xlsm') return 'xlsx';
    if (ext === 'docx' || ext === 'docm') return 'docx';
  }
  if (ext === 'pdf') throw new Error('The file has a .pdf extension but is not a valid PDF.');
  throw new Error('Unsupported file type. Supported: .pdf, .docx, .xlsx.');
}

export function mdName(name) {
  return name.replace(/\.[^.]+$/, '') + '.md';
}

export async function convertToMarkdown(file, options, onProgress = () => {}) {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const buf = await file.arrayBuffer();
  const kind = detectKind(file.name, buf);
  let res;
  onProgress({ stage: 'reading' });
  if (kind === 'docx') res = await (await import('./docx-reader.js')).readDocx(buf, opts);
  else if (kind === 'xlsx') res = await (await import('./xlsx-reader.js')).readXlsx(buf, opts);
  else res = await (await import('./pdf-reader.js')).readPdf(buf, opts, (i, n) => onProgress({ stage: 'page', page: i, pages: n }));
  onProgress({ stage: 'writing' });
  const { toMarkdown } = await import('./md-writer.js');
  const md = toMarkdown(res.doc, { tableMode: opts.tableMode, sourceName: file.name });
  onProgress({ stage: 'verifying' });
  const { verifyConversion, reportToMarkdown } = await import('./verify.js');
  // in pure-Markdown mode merged values are repeated on purpose; the duplication check uses the merged form
  const checkMd = opts.tableMode === 'flatten' ? toMarkdown(res.doc, { tableMode: 'html' }) : undefined;
  const report = await verifyConversion({ kind, buf, md, log: res.log, pages: res.pages, opts, checkMd });
  return { kind, md, report, reportMd: reportToMarkdown(report, file.name), log: res.log, buf };
}

export const OUTPUT_FORMATS = {
  docx: { label: 'Word (.docx)', ext: 'docx' },
  xlsx: { label: 'Excel (.xlsx)', ext: 'xlsx' },
  pdf: { label: 'PDF (.pdf)', ext: 'pdf' },
  html: { label: 'Web page (.html)', ext: 'html' },
};

export async function convertFromMarkdown(file, formats) {
  const md = await file.text();
  const base = file.name.replace(/\.[^.]+$/, '');
  const { readMarkdown } = await import('./md-reader.js');
  const doc = readMarkdown(md);
  const out = [];
  for (const f of formats) {
    let blob;
    if (f === 'docx') blob = await (await import('./docx-writer.js')).writeDocx(doc, { title: base });
    else if (f === 'xlsx') blob = await (await import('./xlsx-writer.js')).writeXlsx(doc);
    else if (f === 'pdf') blob = await (await import('./pdf-writer.js')).writePdf(doc, { title: base });
    else if (f === 'html') blob = new Blob([(await import('./html-writer.js')).standaloneHtml(md, base)], { type: 'text/html' });
    out.push({ name: `${base}.${OUTPUT_FORMATS[f].ext}`, blob });
  }
  return out;
}

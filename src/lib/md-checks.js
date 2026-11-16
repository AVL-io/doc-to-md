// Checks on Markdown source before conversion, plus the outline of the document.
// items: [{ status: 'bad' | 'warn', line, text }]  (bad = serious, warn = note)

const OK_TAGS = new Set(['table', 'thead', 'tbody', 'tfoot', 'tr', 'th', 'td', 'caption', 'br', 'sup', 'sub', 'b', 'i', 'strong', 'em', 's', 'del', 'u', 'code', 'a', 'p', 'hr']);
const SEP_RE = /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$/;
const MAX_ITEMS = 500;

function splitRow(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|') && !s.endsWith('\\|')) s = s.slice(0, -1);
  const cells = [];
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    if (s[i] === '\\' && s[i + 1] === '|') { cur += '|'; i++; continue; }
    if (s[i] === '|') { cells.push(cur); cur = ''; continue; }
    cur += s[i];
  }
  cells.push(cur);
  return cells;
}

function htmlTableDims(text) {
  try {
    const doc = new DOMParser().parseFromString(text, 'text/html');
    const t = doc.querySelector('table');
    if (!t) return { rows: 0, cols: 0, merged: 0 };
    const rows = [...t.rows].filter(r => r.closest('table') === t);
    let cols = 0, merged = 0;
    const carry = [];   // columns still occupied by a row-span
    rows.forEach((r, ri) => {
      let c = 0;
      for (const cell of r.cells) {
        while (carry[c] > 0) c++;
        const rs = Math.max(1, cell.rowSpan || 1), cs = Math.max(1, cell.colSpan || 1);
        if (rs > 1 || cs > 1) merged++;
        for (let k = 0; k < cs; k++) carry[c + k] = rs;
        c += cs;
      }
      cols = Math.max(cols, c);
      for (let k = 0; k < carry.length; k++) if (carry[k] > 0) carry[k]--;
      void ri;
    });
    return { rows: rows.length, cols, merged };
  } catch {
    return { rows: 0, cols: 0, merged: 0 };
  }
}

export function analyzeMarkdown(text) {
  const lines = text.split(/\r?\n/);
  const items = [];
  const add = (status, line, msg) => { if (items.length < MAX_ITEMS) items.push({ status, line, text: msg }); };
  const headings = [], images = [], tables = [], refs = [];
  const defs = new Map();
  const tagSeen = new Set();
  let fence = null;
  let prevLevel = 0;
  const tableStack = [];

  if (!text.trim()) add('warn', 1, 'The file is empty.');

  for (let i = 0; i < lines.length; i++) {
    const n = i + 1, l = lines[i];
    const fm = /^\s{0,3}(```+|~~~+)/.exec(l);
    if (fm) {
      if (!fence) fence = { line: n, ch: fm[1][0] };
      else if (fm[1][0] === fence.ch) fence = null;
      continue;
    }
    if (fence) continue;
    const scan = l.replace(/`[^`]*`/g, m => ' '.repeat(m.length));

    const hm = /^ {0,3}(#{1,6})\s+(.+?)\s*#*\s*$/.exec(l);
    if (hm) {
      const level = hm[1].length;
      headings.push({ level, text: hm[2].replace(/[*_`]/g, '').replace(/<[^>]+>/g, ''), line: n });
      if (prevLevel && level > prevLevel + 1) add('warn', n, `Heading level jumps from ${prevLevel} to ${level}. Styles in Word and PDF may look uneven.`);
      prevLevel = level;
    }

    const dm = /^\[\^([^\]\s]+)\]:/.exec(l);
    if (dm) {
      if (defs.has(dm[1])) add('warn', n, `Footnote [^${dm[1]}] is defined more than once.`);
      else defs.set(dm[1], n);
    }
    for (const m of scan.matchAll(/\[\^([^\]\s]+)\](?!:)/g)) refs.push({ id: m[1], line: n });

    for (const m of scan.matchAll(/!\[([^\]]*)\]\(([^)\s]*)[^)]*\)/g)) images.push({ alt: m[1], src: m[2], line: n });
    for (const m of scan.matchAll(/<img\b[^>]*>/gi)) images.push({ alt: (/alt="([^"]*)"/i.exec(m[0]) || [])[1] || '', src: '', line: n });

    for (const m of scan.matchAll(/(?<!!)\[([^\]]+)\]\(([^)\s]+)[^)]*\)/g)) {
      if (!/^(https?:|mailto:|#)/i.test(m[2])) add('warn', n, `Link “${m[1].slice(0, 40)}” points to “${m[2].slice(0, 40)}”, which cannot be followed inside a document. Only the text is kept.`);
    }

    for (const m of scan.matchAll(/<\/?([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>/g)) {
      if (/^<(https?|mailto|ftp):/i.test(m[0]) || /@/.test(m[0])) continue;
      const tag = m[1].toLowerCase();
      if (tag === 'table') {
        if (!m[0].startsWith('</')) tableStack.push(n);
        else if (!tableStack.length) add('bad', n, '</table> has no matching <table>.');
        else { const s = tableStack.pop(); if (!tableStack.length) tables.push({ kind: 'html', line: s, endLine: n }); }
      } else if (tag !== 'img' && !OK_TAGS.has(tag) && !tagSeen.has(tag)) {
        tagSeen.add(tag);
        add('warn', n, `HTML tag <${tag}> is not supported. The tag is ignored and its text is kept.`);
      }
    }

    // pipe table: header row followed by a separator row
    if (l.includes('|') && i + 1 < lines.length && SEP_RE.test(lines[i + 1]) && lines[i + 1].includes('-')) {
      const cols = splitRow(l).length;
      let j = i + 2, bad = 0;
      while (j < lines.length && lines[j].trim() && lines[j].includes('|')) {
        const c = splitRow(lines[j]).length;
        if (c !== cols && bad < 5) { bad++; add('warn', j + 1, `Table row has ${c} cell${c === 1 ? '' : 's'} but the header has ${cols}. Missing cells are left empty.`); }
        j++;
      }
      tables.push({ kind: 'pipe', line: n, endLine: j, rows: j - i - 1, cols, merged: 0 });
      i = j - 1;
    }
  }

  if (fence) add('bad', fence.line, 'This code block is never closed. Everything after it is read as code.');
  for (const s of tableStack) add('bad', s, '<table> is never closed.');

  const seen = new Set();
  for (const r of refs) {
    const key = r.id + '@' + r.line;
    if (!defs.has(r.id) && !seen.has(key)) { seen.add(key); add('bad', r.line, `Footnote [^${r.id}] is used but never defined. It will appear as plain “[${r.id}]”.`); }
  }
  const used = new Set(refs.map(r => r.id));
  for (const [id, line] of defs) if (!used.has(id)) add('warn', line, `Footnote [^${id}] is defined but never used.`);
  for (const im of images) add('warn', im.line, `Image “${im.alt || im.src || 'untitled'}”. Images are not included in the created documents; a placeholder is written instead.`);

  for (const t of tables) {
    if (t.kind === 'html') Object.assign(t, htmlTableDims(lines.slice(t.line - 1, t.endLine).join('\n')));
  }

  const plain = text.replace(/<!--[\s\S]*?-->/g, ' ').replace(/<[^>]+>/g, ' ');
  const words = (plain.match(/[\p{L}\p{N}]+(?:['’.\-][\p{L}\p{N}]+)*/gu) || []).length;

  items.sort((a, b) => (a.status === b.status ? a.line - b.line : a.status === 'bad' ? -1 : 1));
  return {
    items,
    stats: { lines: text ? (text.endsWith('\n') ? lines.length - 1 : lines.length) : 0, words, headings, tables, footnotes: [...defs.keys()], images, refs: refs.length },
  };
}

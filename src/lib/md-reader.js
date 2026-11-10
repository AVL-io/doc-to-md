// Markdown text -> document model (for Markdown -> Word / Excel / PDF).
import { marked } from 'marked';
import { mergeInlines, trimInlines } from './ast.js';

export function readMarkdown(md) {
  const tokens = marked.lexer(md, { gfm: true });
  const footnotes = [];
  const blocks = convertBlocks(tokens, footnotes);
  return { blocks, footnotes };
}

function convertBlocks(tokens, footnotes) {
  const out = [];
  for (const t of tokens) {
    switch (t.type) {
      case 'heading':
        out.push({ type: 'heading', level: t.depth, inlines: trimInlines(convertInlines(t.tokens)) });
        break;
      case 'paragraph': {
        const fn = /^\[\^([^\]]+)\]:\s*([\s\S]*)$/.exec(t.text);
        if (fn) { footnotes.push({ id: fn[1], blocks: [{ type: 'paragraph', inlines: [{ type: 'text', text: fn[2].replace(/\s+/g, ' ') }] }] }); break; }
        out.push(...splitImages({ type: 'paragraph', inlines: trimInlines(convertInlines(t.tokens)) }));
        break;
      }
      case 'text':
        out.push({ type: 'paragraph', inlines: trimInlines(t.tokens ? convertInlines(t.tokens) : [{ type: 'text', text: t.text }]) });
        break;
      case 'list':
        out.push({
          type: 'list', ordered: t.ordered, start: typeof t.start === 'number' ? t.start : 1,
          items: t.items.map(it => ({ blocks: convertBlocks(it.tokens, footnotes) })),
        });
        break;
      case 'table': {
        const rows = [t.header.map(c => ({ blocks: cellBlocks(c.tokens), rowspan: 1, colspan: 1, header: true }))];
        for (const r of t.rows) rows.push(r.map(c => ({ blocks: cellBlocks(c.tokens), rowspan: 1, colspan: 1, header: false })));
        out.push({ type: 'table', rows });
        break;
      }
      case 'code':
        out.push({ type: 'code', text: t.text });
        break;
      case 'blockquote':
        for (const b of convertBlocks(t.tokens, footnotes)) out.push(b.type === 'paragraph' ? { ...b, style: 'quote' } : b);
        break;
      case 'hr':
        out.push({ type: 'hr' });
        break;
      case 'html':
        out.push(...htmlBlocks(t.text));
        break;
      default:
        break;
    }
  }
  return out;
}

function cellBlocks(tokens) {
  const inl = convertInlines(tokens);
  // split on <br> into separate paragraphs
  const paras = [[]];
  for (const r of inl) {
    if (r.type === 'break') paras.push([]);
    else paras[paras.length - 1].push(r);
  }
  return paras.map(p => ({ type: 'paragraph', inlines: trimInlines(p) })).filter(p => p.inlines.length);
}

function splitImages(block) {
  const imgs = block.inlines.filter(r => r.type === 'image');
  if (!imgs.length) return [block];
  const rest = block.inlines.filter(r => r.type !== 'image');
  const out = imgs.map(i => ({ type: 'image', alt: i.alt }));
  if (rest.some(r => r.type !== 'text' || r.text.trim())) out.unshift({ ...block, inlines: trimInlines(rest) });
  return out;
}

function convertInlines(tokens, baseFmt = {}) {
  const out = [];
  let fmt = { ...baseFmt };   // inline HTML tags such as <sup> ... </sup> change the format of what follows
  for (const t of tokens || []) {
    switch (t.type) {
      case 'text':
        if (t.tokens && t.tokens.length) out.push(...convertInlines(t.tokens, fmt));
        else out.push({ type: 'text', text: decodeEntities(t.text), ...fmt });
        break;
      case 'escape': out.push({ type: 'text', text: t.text, ...fmt }); break;
      case 'strong': out.push(...convertInlines(t.tokens, { ...fmt, bold: true })); break;
      case 'em': out.push(...convertInlines(t.tokens, { ...fmt, italic: true })); break;
      case 'del': out.push(...convertInlines(t.tokens, { ...fmt, strike: true })); break;
      case 'codespan': out.push({ type: 'text', text: decodeEntities(t.text), ...fmt, code: true }); break;
      case 'br': out.push({ type: 'break' }); break;
      case 'link': out.push(...convertInlines(t.tokens, { ...fmt, href: t.href })); break;
      case 'image': out.push({ type: 'image', alt: t.text || '' }); break;
      case 'html': {
        const tag = /^<\s*(\/?)\s*([a-z0-9]+)/i.exec(t.text.trim());
        if (!tag) break;
        const name = tag[2].toLowerCase();
        if (name === 'br') { out.push({ type: 'break' }); break; }
        const prop = { sup: 'sup', sub: 'sub', b: 'bold', strong: 'bold', i: 'italic', em: 'italic', s: 'strike', del: 'strike', code: 'code', u: null }[name];
        if (prop) fmt = { ...fmt, [prop]: !tag[1] };
        break;
      }
      default:
        if (t.raw) out.push({ type: 'text', text: t.raw, ...fmt });
    }
  }
  return mergeInlines(out.map(r => {
    if (r.type !== 'text') return r;
    const c = { ...r };
    for (const k of ['bold', 'italic', 'sup', 'sub', 'strike', 'code']) if (!c[k]) delete c[k];
    return c;
  }));
}

// ---------- embedded HTML ----------

function domInlines(node, fmt) {
  const out = [];
  for (const n of node.childNodes) {
    if (n.nodeType === 3) { out.push({ type: 'text', text: n.textContent.replace(/\s+/g, ' '), ...fmt }); continue; }
    if (n.nodeType !== 1) continue;
    const tag = n.tagName.toLowerCase();
    if (tag === 'br') out.push({ type: 'break' });
    else if (tag === 'b' || tag === 'strong') out.push(...domInlines(n, { ...fmt, bold: true }));
    else if (tag === 'i' || tag === 'em') out.push(...domInlines(n, { ...fmt, italic: true }));
    else if (tag === 's' || tag === 'del' || tag === 'strike') out.push(...domInlines(n, { ...fmt, strike: true }));
    else if (tag === 'sup') out.push(...domInlines(n, { ...fmt, sup: true }));
    else if (tag === 'sub') out.push(...domInlines(n, { ...fmt, sub: true }));
    else if (tag === 'code') out.push(...domInlines(n, { ...fmt, code: true }));
    else if (tag === 'a') out.push(...domInlines(n, { ...fmt, href: n.getAttribute('href') || undefined }));
    else if (tag === 'table') continue;
    else out.push(...domInlines(n, fmt));
  }
  return out;
}

function htmlTable(table) {
  const rows = [];
  const trs = [...table.querySelectorAll('tr')].filter(tr => tr.closest('table') === table);
  for (const tr of trs) {
    const row = [];
    for (const td of tr.children) {
      if (!/^t[dh]$/i.test(td.tagName)) continue;
      const blocks = [];
      const inl = domInlines(td, {});
      const paras = [[]];
      for (const r of inl) { if (r.type === 'break') paras.push([]); else paras[paras.length - 1].push(r); }
      for (const p of paras) { const t = trimInlines(mergeInlines(p)); if (t.length) blocks.push({ type: 'paragraph', inlines: t }); }
      for (const nested of td.querySelectorAll('table')) if (nested.parentElement.closest('table') === table) blocks.push(htmlTable(nested));
      row.push({ blocks, rowspan: Math.max(1, +td.getAttribute('rowspan') || 1), colspan: Math.max(1, +td.getAttribute('colspan') || 1), header: td.tagName.toLowerCase() === 'th' });
    }
    rows.push(row);
  }
  return { type: 'table', rows };
}

function htmlBlocks(html) {
  const trimmed = html.trim();
  if (!trimmed || /^<!--[\s\S]*-->$/.test(trimmed)) return [];
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const out = [];
  for (const el of doc.body.children) {
    if (el.tagName === 'TABLE') out.push(htmlTable(el));
    else {
      const inl = trimInlines(mergeInlines(domInlines(el, {})));
      if (inl.length) out.push({ type: 'paragraph', inlines: inl });
    }
  }
  if (!doc.body.children.length && doc.body.textContent.trim()) out.push({ type: 'paragraph', inlines: [{ type: 'text', text: doc.body.textContent.trim() }] });
  return out;
}

function decodeEntities(s) {
  if (!/&[#a-z0-9]+;/i.test(s)) return s;
  const ta = document.createElement('textarea');
  ta.innerHTML = s;
  return ta.value;
}


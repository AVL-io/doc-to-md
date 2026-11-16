// Document model -> HTML for the "Word" preview. Every piece of text is escaped,
// so a Markdown file cannot inject markup into the page.
import { mergeInlines, trimInlines, normalizeTable } from './ast.js';
import { esc } from './util.js';

function inl(inlines, fnIdx) {
  let s = '';
  for (const r of mergeInlines(inlines || [])) {
    if (r.type === 'break') { s += '<br>'; continue; }
    if (r.type === 'footref') { const n = fnIdx.get(String(r.id)); s += `<sup>${n ? n : '[' + esc(r.id) + ']'}</sup>`; continue; }
    if (r.type !== 'text') continue;
    let t = esc(r.text.replace(/\n/g, ' '));
    if (r.code) t = `<code>${t}</code>`;
    if (r.sup) t = `<sup>${t}</sup>`;
    if (r.sub) t = `<sub>${t}</sub>`;
    if (r.bold && t.trim()) t = `<b>${t}</b>`;
    if (r.italic && t.trim()) t = `<i>${t}</i>`;
    if (r.strike && t.trim()) t = `<s>${t}</s>`;
    if (r.href && /^(https?:|mailto:)/i.test(r.href)) t = `<a href="${esc(r.href)}" target="_blank" rel="noopener noreferrer">${t}</a>`;
    s += t;
  }
  return s;
}

function list(b, fnIdx) {
  const tag = b.ordered ? 'ol' : 'ul';
  const start = b.ordered && b.start > 1 ? ` start="${b.start}"` : '';
  return `<${tag}${start}>` + b.items.map(it => `<li>${it.blocks.map(x => block(x, fnIdx, true)).join('')}</li>`).join('') + `</${tag}>`;
}

function table(b, fnIdx) {
  const t = normalizeTable(b);
  return '<table class="tb">' + t.rows.map(r => '<tr>' + r.map(c => {
    const tag = c.header ? 'th' : 'td';
    const attrs = (c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : '') + (c.colspan > 1 ? ` colspan="${c.colspan}"` : '');
    const content = (c.blocks || []).map(x => x.type === 'table' ? table(x, fnIdx) : x.type === 'list' ? list(x, fnIdx) : x.inlines ? inl(trimInlines(x.inlines), fnIdx) : '').filter(Boolean).join('<br>');
    return `<${tag}${attrs}>${content}</${tag}>`;
  }).join('') + '</tr>').join('') + '</table>';
}

function block(b, fnIdx, inner = false) {
  switch (b.type) {
    case 'heading': return `<h${Math.min(6, b.level)}>${inl(trimInlines(b.inlines), fnIdx)}</h${Math.min(6, b.level)}>`;
    case 'paragraph': {
      const t = inl(trimInlines(b.inlines), fnIdx);
      if (!t) return '';
      if (b.style === 'quote') return `<blockquote>${t}</blockquote>`;
      if (b.style === 'caption') return `<p class="caption"><i>${t}</i></p>`;
      return inner ? `<span class="li-text">${t}</span>` : `<p>${t}</p>`;
    }
    case 'list': return list(b, fnIdx);
    case 'table': return table(b, fnIdx);
    case 'code': return `<pre class="code">${esc(b.text)}</pre>`;
    case 'hr': return '<hr>';
    case 'image': return `<p class="img">(Image${b.alt ? ': ' + esc(b.alt) : ''})</p>`;
    default: return '';
  }
}

/** parts: one entry per top-level block with its source line range; foot: footnotes block. */
export function modelToParts(doc) {
  const fnIdx = new Map((doc.footnotes || []).map((f, i) => [String(f.id), i + 1]));
  const parts = [];
  for (const b of doc.blocks) {
    if (b.type === 'pagebreak') continue;
    const html = block(b, fnIdx);
    if (html) parts.push({ ls: b.line || 0, le: b.endLine || b.line || 0, html });
  }
  const foot = (doc.footnotes || []).map((f, i) => ({
    ls: f.line || 0, le: f.line || 0,
    html: `<p class="fn-item"><sup>${i + 1}</sup> ${f.blocks.map(x => x.inlines ? inl(trimInlines(x.inlines), fnIdx) : '').join(' ')}</p>`,
  }));
  return { parts, foot };
}

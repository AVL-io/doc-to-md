// Markdown -> safe HTML (preview and standalone .html export).
import { marked } from 'marked';

const ALLOWED = new Set(['A', 'B', 'STRONG', 'I', 'EM', 'S', 'DEL', 'U', 'SUP', 'SUB', 'CODE', 'PRE', 'P', 'BR', 'HR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
  'UL', 'OL', 'LI', 'BLOCKQUOTE', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TH', 'TD', 'CAPTION', 'SPAN', 'DIV', 'IMG', 'INPUT', 'SECTION']);
const ATTRS = { A: ['href'], TD: ['rowspan', 'colspan', 'align'], TH: ['rowspan', 'colspan', 'align'], OL: ['start'], IMG: ['alt'], INPUT: ['type', 'checked', 'disabled'] };

/** Render Markdown to HTML and remove anything that is not plain document markup. */
export function markdownToSafeHtml(md) {
  const raw = marked.parse(md, { gfm: true, async: false });
  const doc = new DOMParser().parseFromString(`<body>${raw}</body>`, 'text/html');
  const clean = (node) => {
    for (const child of [...node.childNodes]) {
      if (child.nodeType === 8) { child.remove(); continue; }
      if (child.nodeType !== 1) continue;
      if (!ALLOWED.has(child.tagName)) {
        if (/^(SCRIPT|STYLE|IFRAME|OBJECT|EMBED|LINK|META|FORM|SVG|MATH|TEMPLATE)$/.test(child.tagName)) { child.remove(); continue; }
        clean(child);
        child.replaceWith(...child.childNodes);
        continue;
      }
      for (const a of [...child.attributes]) {
        const ok = (ATTRS[child.tagName] || []).includes(a.name.toLowerCase());
        if (!ok) child.removeAttribute(a.name);
      }
      if (child.tagName === 'A') {
        const href = child.getAttribute('href') || '';
        if (!/^(https?:|mailto:|#)/i.test(href)) child.removeAttribute('href');
        else { child.setAttribute('rel', 'noopener noreferrer'); child.setAttribute('target', '_blank'); }
      }
      if (child.tagName === 'IMG') { child.replaceWith(doc.createTextNode(`(Image${child.getAttribute('alt') ? ': ' + child.getAttribute('alt') : ''})`)); continue; }
      clean(child);
    }
  };
  clean(doc.body);
  return doc.body.innerHTML;
}

export function standaloneHtml(md, title) {
  const body = markdownToSafeHtml(md);
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
body{font-family:Calibri,Segoe UI,Arial,sans-serif;max-width:60rem;margin:2rem auto;padding:0 1rem;line-height:1.5;color:#1d2433}
table{border-collapse:collapse;margin:1rem 0;width:100%;font-size:.92em}th,td{border:1px solid #9aa4b2;padding:.35rem .5rem;vertical-align:top;text-align:left}th{background:#e7ecf2}
blockquote{border-left:3px solid #9aa4b2;margin:0;padding-left:1rem;color:#4a5568}code{background:#f2f2f2;padding:0 .2rem}pre{background:#f2f2f2;padding:.75rem;overflow:auto}
</style></head><body>
${body}
</body></html>
`;
}

// PDF -> document model. Works on digitally generated PDFs (text is extracted,
// not recognised). Tables are rebuilt from the drawn border lines, which lets
// merged cells (row-spans / column-spans) be recovered. Tagged PDFs (e.g. Word
// "Save as PDF") additionally provide headings, lists and header/footer marks.
import * as pdfjs from 'pdfjs-dist/build/pdf.mjs';
import { buildLists, mergeInlines, trimInlines, inlinesText, blocksText } from './ast.js';

export const DEFAULT_OPTIONS = {
  removeImages: true,
  removeHeadersFooters: true,
  removePageNumbers: true,
  removeFootnotes: false,
  removeToc: true,
  pageMarkers: true,
};

export function setPdfWorkerSrc(src) {
  pdfjs.GlobalWorkerOptions.workerSrc = src;
}

export async function openPdf(buf) {
  const data = buf instanceof Uint8Array ? buf.slice() : new Uint8Array(buf.slice(0));
  try {
    return await pdfjs.getDocument({ data, isEvalSupported: false, disableFontFace: true, useSystemFonts: false, verbosity: 0 }).promise;
  } catch (e) {
    if (e && e.name === 'PasswordException') throw new Error('The PDF is password-protected. Remove the password and try again.');
    throw new Error('The PDF could not be opened: ' + (e && e.message ? e.message : e));
  }
}

const BLOCK_ROLES = new Set(['P', 'H', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'Title', 'Caption', 'LI', 'TOCI', 'Note', 'BlockQuote', 'Quote', 'Code', 'Formula', 'Figure', 'Lbl', 'LBody']);
const BULLET_RE = /^([•●○◦▪▫■□►▸‣⁃∙·➢➤✓✔❖\-–—*]|[\uE000-\uF8FF])(\s|$)/;
const ORDER_RE = /^(\(?(\d{1,3}|[a-zA-Z]|[ivxIVX]{1,5})[.)])(\s|$)/;
const PAGE_NO_RE = /^(page\s*)?(\d{1,4}|[ivxlcdm]{1,7})(\s*(of|\/)\s*\d{1,4})?$/i;
const TOC_LINE_RE = /^(.{2,}?)\s*(?:[.·…_]{3,}|\s{3,}|\t)\s*(\d{1,4}|[ivxlc]{1,6})$/i;

/**
 * @returns {{doc, log, pages}}  pages: per-page info for verification
 */
export async function readPdf(buf, options = {}, onProgress = () => {}) {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const pdf = await openPdf(buf);
  const n = pdf.numPages;
  const log = { excluded: [], warnings: [], stats: { pages: n, tables: 0, cells: 0, mergedCells: 0, images: 0, tagged: false, lowConfidenceTables: [] } };
  const exclude = (category, text, page) => log.excluded.push({ category, text: (text || '').trim(), page });

  // ---------- pass 1: extract every page ----------
  const pages = [];
  for (let i = 1; i <= n; i++) {
    onProgress(i, n);
    const page = await pdf.getPage(i);
    pages.push(await extractPage(page, i, log));
    page.cleanup();
  }
  await pdf.destroy();
  log.stats.tagged = pages.some(p => p.tagged);

  // ---------- pass 2: document-wide statistics ----------
  const sizeHist = new Map();
  for (const p of pages) for (const l of p.lines) if (l.frags && !l.artifact) {
    const k = Math.round(l.size * 2) / 2;
    sizeHist.set(k, (sizeHist.get(k) || 0) + l.text.length);
  }
  let bodySize = 10, best = -1;
  for (const [k, v] of sizeHist) if (v > best) { best = v; bodySize = k; }
  const repeated = findRepeatedMargins(pages);
  const hasHeadingTags = pages.some(p => p.lines.some(l => /^H\d?$/.test(l.role || '')));

  // heading sizes (untagged documents)
  const headingSizes = [...sizeHist.keys()].filter(s => s >= bodySize * 1.15).sort((a, b) => b - a).slice(0, 5);

  // ---------- pass 3: build blocks ----------
  const blocks = [];
  for (const p of pages) {
    if (opts.pageMarkers) blocks.push({ type: 'pagebreak', page: p.num });
    if (!p.lines.length && !p.tables.length) {
      if (!p.hasText) log.warnings.push(`Page ${p.num} contains no extractable text. It may be a scanned image; scanned pages are not supported.`);
    }
    const units = groupParagraphs(p, { bodySize, repeated, opts, log, exclude, hasHeadingTags, headingSizes });
    blocks.push(...units);
  }

  const doc = { blocks: postProcess(blocks, opts, log, exclude), footnotes: [] };
  return { doc, log, pages: pages.map(p => ({ num: p.num, rawText: p.rawText, width: p.width, height: p.height })) };
}

// ======================================================================
// Pass 1: page extraction
// ======================================================================

async function extractPage(page, num, log) {
  const vp = page.getViewport({ scale: 1 });
  const [tc, ops, st] = await Promise.all([
    page.getTextContent({ includeMarkedContent: true }),
    page.getOperatorList(),
    page.getStructTree().catch(() => null),
  ]);
  const width = vp.width, height = vp.height;

  // ---- structure tree: mcid -> ancestor chain ----
  const mcInfo = new Map();
  let nid = 0;
  const walk = (node, chain) => {
    if (!node) return;
    if (node.type === 'content' && node.id) { mcInfo.set(node.id, chain); return; }
    if (!node.role) return;
    const me = { role: normRole(node.role), nid: ++nid, alt: node.alt || '' };
    const next = node.role === 'Root' ? chain : [...chain, me];
    for (const k of node.children || []) walk(k, next);
  };
  walk(st, []);
  const tagged = mcInfo.size > 0;

  // ---- fonts ----
  const fontCache = new Map();
  const fontStyle = (name) => {
    if (fontCache.has(name)) return fontCache.get(name);
    let bold = false, italic = false, mono = false;
    try {
      if (page.commonObjs.has(name)) {
        const f = page.commonObjs.get(name);
        const fn = (f && (f.name || f.loadedName) || '') + ' ' + (tc.styles[name]?.fontFamily || '');
        bold = !!(f && f.bold) || /bold|black|heavy|semibold|demi/i.test(fn);
        italic = !!(f && f.italic) || /italic|oblique/i.test(fn);
        mono = /courier|mono|consol/i.test(fn);
      }
    } catch { /* font not available */ }
    const res = { bold, italic, mono };
    fontCache.set(name, res);
    return res;
  };

  // ---- text fragments ----
  const frags = [];
  const mcStack = [];
  let rawText = '';
  let figures = [];
  for (const it of tc.items) {
    if (it.type === 'beginMarkedContent' || it.type === 'beginMarkedContentProps') {
      mcStack.push({ tag: it.tag, id: it.id || null });
      const chain = it.id && mcInfo.get(it.id);
      const fig = chain && [...chain].reverse().find(a => a.role === 'Figure');
      if (fig && !figures.includes(fig.nid)) {
        figures.push(fig.nid);
        frags.push({ figure: true, alt: fig.alt, chain, order: frags.length });
      }
      continue;
    }
    if (it.type === 'endMarkedContent') { mcStack.pop(); continue; }
    if (it.str == null) continue;
    if (it.hasEOL) rawText += it.str + '\n'; else rawText += it.str;
    if (!it.str) { if (it.hasEOL && frags.length) frags[frags.length - 1].eol = true; continue; }
    const [a, b, c, d, e, f] = it.transform;
    const size = Math.hypot(c, d) || Math.hypot(a, b) || it.height || 10;
    const [x, y] = pdfjs.Util.applyTransform([e, f], vp.transform);
    const rotated = Math.abs(b) > Math.abs(a) * 0.1 || a < 0;
    const artifact = mcStack.some(m => m.tag === 'Artifact');
    let mcid = null;
    for (let k = mcStack.length - 1; k >= 0; k--) if (mcStack[k].id) { mcid = mcStack[k].id; break; }
    const chain = mcid ? (mcInfo.get(mcid) || null) : null;
    if (!it.str.trim()) {
      if (frags.length && !frags[frags.length - 1].figure) frags[frags.length - 1].spaceAfter = true;
      continue;
    }
    const fs = fontStyle(it.fontName);
    frags.push({
      str: it.str.replace(/[\uE000-\uF8FF]/g, ch => PUA_BULLETS[ch] ?? '\u2022'),
      x, y, w: it.width, size, rotated,
      bold: fs.bold, italic: fs.italic, mono: fs.mono,
      artifact, chain, eol: it.hasEOL, order: frags.length,
    });
  }

  // ---- ruling lines ----
  const rules = extractRules(ops, vp.transform);
  let images = 0;
  for (const fn of ops.fnArray) if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintInlineImageXObject || fn === pdfjs.OPS.paintImageMaskXObject) images++;
  log.stats.images += Math.max(images, figures.length);

  // ---- tables from ruling lines ----
  const tables = detectTables(rules, frags.filter(f => !f.figure && !f.artifact), num, log);
  tables.push(...structTables(frags.filter(f => !f.figure && !f.artifact && f.tableId == null && chainHas(f.chain, 'Table')), num, log));
  // tables become pseudo-fragments so they keep their reading-order position
  for (const t of tables) frags.push({ tableBlock: t.block, order: t.order, x: 0, y: t.block.meta.top, size: 0, str: '' });

  // ---- lines (non-table text) ----
  const lines = buildLines(frags.filter(f => f.figure || f.tableId == null));

  return { num, width, height, lines, tables, tagged, rawText, hasText: frags.some(f => !f.figure) };
}

const PUA_BULLETS = { '\uF0B7': '•', '\uF0A7': '▪', '\uF0D8': '➢', '\uF076': '❖', '\uF0FC': '✓', '\uF06E': '■', '\uF0A8': '□', '\uF0E0': '→' };

function normRole(r) {
  const s = String(r);
  if (/^H[1-6]$/i.test(s)) return s.toUpperCase();
  const map = { 'Heading': 'H', 'Paragraph': 'P', 'ListItem': 'LI', 'Table': 'Table' };
  return map[s] || s;
}

function chainHas(chain, role) {
  return !!chain && chain.some(a => a.role === role);
}
function chainFind(chain, pred) {
  if (!chain) return null;
  for (let i = chain.length - 1; i >= 0; i--) if (pred(chain[i])) return chain[i];
  return null;
}

// ---------------- ruling-line extraction ----------------

function mul(m, n) {
  return [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
}
function apply(m, x, y) {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function isWhite(c) {
  if (!c) return false;
  if (typeof c === 'string') {
    const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(c);
    return !!m && [m[1], m[2], m[3]].every(h => parseInt(h, 16) >= 245);
  }
  if (c.length >= 3) return c[0] >= 245 && c[1] >= 245 && c[2] >= 245;
  return false;
}

function extractRules(ops, vpT) {
  const OPS = pdfjs.OPS;
  const out = [];
  let ctm = [1, 0, 0, 1, 0, 0];
  const stack = [];
  let path = null; // { rects:[], polys:[] } in device space
  let fillColor = null, strokeColor = null;
  const colorStack = [];
  const toDev = (x, y) => { const [u, v] = apply(ctm, x, y); return apply(vpT, u, v); };

  const addRect = (x0, y0, x1, y1, how) => {
    const w = Math.abs(x1 - x0), h = Math.abs(y1 - y0);
    const X0 = Math.min(x0, x1), X1 = Math.max(x0, x1), Y0 = Math.min(y0, y1), Y1 = Math.max(y0, y1);
    if (how === 'fill') {
      if (h <= 3 && w >= 2 && w > h) out.push({ o: 'h', y: (Y0 + Y1) / 2, a: X0, b: X1 });
      else if (w <= 3 && h >= 2) out.push({ o: 'v', x: (X0 + X1) / 2, a: Y0, b: Y1 });
    } else {
      // stroked rectangle: four edges
      if (w >= 2) { out.push({ o: 'h', y: Y0, a: X0, b: X1 }); out.push({ o: 'h', y: Y1, a: X0, b: X1 }); }
      if (h >= 2) { out.push({ o: 'v', x: X0, a: Y0, b: Y1 }); out.push({ o: 'v', x: X1, a: Y0, b: Y1 }); }
    }
  };

  const paint = (how) => {
    if (!path) return;
    if ((how === 'fill' && isWhite(fillColor)) || (how === 'stroke' && isWhite(strokeColor))) { path = null; return; }
    for (const r of path.rects) addRect(r[0], r[1], r[2], r[3], how);
    for (const poly of path.polys) {
      if (how === 'fill') {
        // a closed axis-aligned 4-corner polygon is a filled rectangle
        const xs = poly.map(p => p[0]), ys = poly.map(p => p[1]);
        const axis = poly.every((p, i) => { const q = poly[(i + 1) % poly.length]; return Math.abs(p[0] - q[0]) < 0.5 || Math.abs(p[1] - q[1]) < 0.5; });
        if (axis && poly.length >= 4) addRect(Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys), 'fill');
      } else {
        for (let i = 0; i + 1 < poly.length; i++) {
          const [x0, y0] = poly[i], [x1, y1] = poly[i + 1];
          if (Math.abs(y0 - y1) < 0.8 && Math.abs(x1 - x0) >= 2) out.push({ o: 'h', y: (y0 + y1) / 2, a: Math.min(x0, x1), b: Math.max(x0, x1) });
          else if (Math.abs(x0 - x1) < 0.8 && Math.abs(y1 - y0) >= 2) out.push({ o: 'v', x: (x0 + x1) / 2, a: Math.min(y0, y1), b: Math.max(y0, y1) });
        }
      }
    }
    path = null;
  };

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i], args = ops.argsArray[i];
    switch (fn) {
      case OPS.save: stack.push(ctm); colorStack.push([fillColor, strokeColor]); break;
      case OPS.restore: ctm = stack.pop() || ctm; { const c = colorStack.pop(); if (c) [fillColor, strokeColor] = c; } break;
      case OPS.transform: ctm = mul(ctm, args); break;
      case OPS.paintFormXObjectBegin: stack.push(ctm); if (args && args[0]) ctm = mul(ctm, args[0]); break;
      case OPS.paintFormXObjectEnd: ctm = stack.pop() || ctm; break;
      case OPS.setFillRGBColor: fillColor = args && (args[0] ?? args); break;
      case OPS.setStrokeRGBColor: strokeColor = args && (args[0] ?? args); break;
      case OPS.constructPath: {
        const [sub, coords] = args;
        path = path || { rects: [], polys: [] };
        let j = 0, cur = null;
        for (const op of sub) {
          if (op === OPS.rectangle) {
            const x = coords[j++], y = coords[j++], w = coords[j++], h = coords[j++];
            const p0 = toDev(x, y), p1 = toDev(x + w, y + h);
            const q0 = toDev(x + w, y), q1 = toDev(x, y + h);
            // only axis-aligned rectangles are of interest
            if (Math.abs(p0[1] - q0[1]) < 0.5 || Math.abs(p0[0] - q0[0]) < 0.5) path.rects.push([p0[0], p0[1], p1[0], p1[1]]);
            else void q1;
          } else if (op === OPS.moveTo) {
            cur = [toDev(coords[j++], coords[j++])];
            path.polys.push(cur);
          } else if (op === OPS.lineTo) {
            const p = toDev(coords[j++], coords[j++]);
            if (!cur) { cur = [p]; path.polys.push(cur); } else cur.push(p);
          } else if (op === OPS.curveTo) { j += 6; cur = null; }
          else if (op === OPS.curveTo2 || op === OPS.curveTo3) { j += 4; cur = null; }
          else if (op === OPS.closePath) { if (cur && cur.length > 1) cur.push(cur[0]); cur = null; }
        }
        break;
      }
      case OPS.stroke: case OPS.closeStroke: paint('stroke'); break;
      case OPS.fill: case OPS.eoFill: paint('fill'); break;
      case OPS.fillStroke: case OPS.eoFillStroke: case OPS.closeFillStroke: case OPS.closeEOFillStroke: {
        const saved = path; paint('fill'); path = saved; paint('stroke'); break;
      }
      case OPS.endPath: path = null; break;
      default: break;
    }
  }
  return mergeRules(out);
}

function mergeRules(rules) {
  const merged = [];
  for (const o of ['h', 'v']) {
    const list = rules.filter(r => r.o === o).map(r => ({ ...r, c: o === 'h' ? r.y : r.x }));
    list.sort((p, q) => p.c - q.c || p.a - q.a);
    const groups = [];
    for (const r of list) {
      const g = groups[groups.length - 1];
      if (g && Math.abs(g.c - r.c) <= 1.2) g.items.push(r); else groups.push({ c: r.c, items: [r] });
    }
    for (const g of groups) {
      const c = g.items.reduce((s, r) => s + r.c, 0) / g.items.length;
      g.items.sort((p, q) => p.a - q.a);
      let cur = null;
      for (const r of g.items) {
        if (cur && r.a <= cur.b + 1.5) cur.b = Math.max(cur.b, r.b);
        else { if (cur) merged.push(cur); cur = { o, c, a: r.a, b: r.b }; }
      }
      if (cur) merged.push(cur);
    }
  }
  return merged.filter(r => r.b - r.a >= 4);
}

// ---------------- table detection ----------------

function cluster(values, tol) {
  const s = [...values].sort((a, b) => a - b);
  const out = [];
  for (const v of s) {
    const last = out[out.length - 1];
    if (last && v - last.max <= tol) { last.sum += v; last.n++; last.max = v; }
    else out.push({ sum: v, n: 1, max: v });
  }
  return out.map(g => g.sum / g.n);
}

function detectTables(rules, frags, pageNum, log) {
  const H = rules.filter(r => r.o === 'h'), V = rules.filter(r => r.o === 'v');
  if (H.length < 2 || V.length < 2) return [];
  const tol = 2.5;
  // connected components of intersecting rules
  const all = [...H, ...V];
  const parent = all.map((_, i) => i);
  const find = (i) => (parent[i] === i ? i : (parent[i] = find(parent[i])));
  const unite = (i, j) => { parent[find(i)] = find(j); };
  for (let i = 0; i < H.length; i++) for (let j = 0; j < V.length; j++) {
    const h = H[i], v = V[j];
    if (v.c >= h.a - tol && v.c <= h.b + tol && h.c >= v.a - tol && h.c <= v.b + tol) unite(i, H.length + j);
  }
  const comps = new Map();
  all.forEach((r, i) => { const k = find(i); if (!comps.has(k)) comps.set(k, []); comps.get(k).push(r); });

  const tables = [];
  let tIndex = 0;
  for (const rs of comps.values()) {
    const hs = rs.filter(r => r.o === 'h'), vs = rs.filter(r => r.o === 'v');
    if (hs.length < 2 || vs.length < 2) continue;
    const x0 = Math.min(...hs.map(h => h.a), ...vs.map(v => v.c)), x1 = Math.max(...hs.map(h => h.b), ...vs.map(v => v.c));
    const y0 = Math.min(...vs.map(v => v.a), ...hs.map(h => h.c)), y1 = Math.max(...vs.map(v => v.b), ...hs.map(h => h.c));
    if (x1 - x0 < 20 || y1 - y0 < 8) continue;
    let xs = cluster([...vs.map(v => v.c), x0, x1], tol);
    let ys = cluster([...hs.map(h => h.c), y0, y1], tol);
    const nR = ys.length - 1, nC = xs.length - 1;
    if (nR < 1 || nC < 1) continue;
    if (nR * nC === 1) continue; // a single box (frame / text box) is not a table

    const hasV = (xi, r) => {
      const a = ys[r], b = ys[r + 1], need = (b - a) * 0.5;
      return vs.some(v => Math.abs(v.c - xs[xi]) <= tol && Math.min(v.b, b) - Math.max(v.a, a) >= need);
    };
    const hasH = (yi, c) => {
      const a = xs[c], b = xs[c + 1], need = (b - a) * 0.5;
      return hs.some(h => Math.abs(h.c - ys[yi]) <= tol && Math.min(h.b, b) - Math.max(h.a, a) >= need);
    };
    const id = (r, c) => r * nC + c;
    const par = Array.from({ length: nR * nC }, (_, i) => i);
    const f2 = (i) => (par[i] === i ? i : (par[i] = f2(par[i])));
    for (let r = 0; r < nR; r++) for (let c = 0; c < nC; c++) {
      if (c < nC - 1 && !hasV(c + 1, r)) par[f2(id(r, c))] = f2(id(r, c + 1));
      if (r < nR - 1 && !hasH(r + 1, c)) par[f2(id(r, c))] = f2(id(r + 1, c));
    }
    let regions = new Map();
    for (let r = 0; r < nR; r++) for (let c = 0; c < nC; c++) {
      const k = f2(id(r, c));
      const g = regions.get(k) || { r0: r, r1: r, c0: c, c1: c, n: 0, frags: [] };
      g.r0 = Math.min(g.r0, r); g.r1 = Math.max(g.r1, r); g.c0 = Math.min(g.c0, c); g.c1 = Math.max(g.c1, c); g.n++;
      regions.set(k, g);
    }
    let lowConfidence = false;
    // non-rectangular regions are split back into single grid units
    const fixed = [];
    for (const g of regions.values()) {
      if (g.n === (g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1)) fixed.push(g);
      else {
        lowConfidence = true;
        for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) if (regions.get(f2(id(r, c))) === g) fixed.push({ r0: r, r1: r, c0: c, c1: c, n: 1, frags: [] });
      }
    }
    // assign text by fragment centre
    const unitOf = new Map();
    for (const g of fixed) for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) unitOf.set(id(r, c), g);
    let count = 0;
    for (const f of frags) {
      if (f.tableId != null) continue;
      const cx = f.x + Math.min(f.w, 4000) / 2, cy = f.y - f.size * 0.3;
      if (cx < xs[0] - 1 || cx > xs[nC] + 1 || cy < ys[0] - 1 || cy > ys[nR] + 1) continue;
      // fragment spanning several columns: use its start instead of its centre
      const px = f.w > (xs[nC] - xs[0]) * 0.6 ? f.x + 1 : cx;
      let r = 0; while (r < nR - 1 && cy > ys[r + 1]) r++;
      let c = 0; while (c < nC - 1 && px > xs[c + 1]) c++;
      const g = unitOf.get(id(r, c));
      if (!g) continue;
      g.frags.push(f);
      f.tableId = `${pageNum}-${tIndex}`;
      count++;
    }
    if (!count) { continue; }
    // A frame around ordinary text (one column, text-only) is not a table.
    if (nC === 1 && fixed.length <= 2) { for (const f of frags) if (f.tableId === `${pageNum}-${tIndex}`) f.tableId = null; continue; }

    // remove grid lines that no cell starts at (caused by slightly misaligned borders)
    const colUsed = new Array(nC).fill(false), rowUsed = new Array(nR).fill(false);
    colUsed[0] = rowUsed[0] = true;
    for (const g of fixed) { colUsed[g.c0] = true; rowUsed[g.r0] = true; }
    const colMap = [], rowMap = [];
    let k = -1; for (let c = 0; c < nC; c++) { if (colUsed[c]) k++; colMap[c] = k; }
    const nCols = k + 1;
    k = -1; for (let r = 0; r < nR; r++) { if (rowUsed[r]) k++; rowMap[r] = k; }
    const nRows = k + 1;
    const keptXs = xs.filter((x, i) => i === nC || colUsed[i]);

    const rows = Array.from({ length: nRows }, () => []);
    fixed.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
    for (const g of fixed) {
      const R0 = rowMap[g.r0], R1 = rowMap[g.r1], C0 = colMap[g.c0], C1 = colMap[g.c1];
      rows[R0].push({ c: C0, cell: { blocks: cellBlocks(g.frags), rowspan: R1 - R0 + 1, colspan: C1 - C0 + 1, header: false, _frags: g.frags } });
    }
    const tableRows = rows.map(r => r.sort((a, b) => a.c - b.c).map(x => x.cell));
    markHeader(tableRows);
    for (const r of tableRows) for (const c of r) delete c._frags;
    const firstOrder = Math.min(...frags.filter(f => f.tableId === `${pageNum}-${tIndex}`).map(f => f.order));
    const merged = tableRows.reduce((s, r) => s + r.filter(c => c.rowspan > 1 || c.colspan > 1).length, 0);
    log.stats.tables++;
    log.stats.cells += tableRows.reduce((s, r) => s + r.length, 0);
    log.stats.mergedCells += merged;
    if (lowConfidence) log.stats.lowConfidenceTables.push(pageNum);
    tables.push({
      order: firstOrder,
      block: { type: 'table', rows: tableRows, meta: { page: pageNum, xs: keptXs, top: ys[0], bottom: ys[nR], nCols, lowConfidence } },
    });
    tIndex++;
  }
  return tables;
}

/**
 * Tables without border lines: rebuild rows/cells from the PDF tags (TR/TD/TH).
 * Column spans are inferred from cell positions and row spans from gaps at the
 * start of a row, so the result is marked as inferred.
 */
function structTables(frags, pageNum, log) {
  const byTable = new Map();
  for (const f of frags) {
    const t = chainFind(f.chain, a => a.role === 'Table');
    if (!t) continue;
    if (!byTable.has(t.nid)) byTable.set(t.nid, []);
    byTable.get(t.nid).push(f);
  }
  const out = [];
  let idx = 0;
  for (const [tid, tf] of byTable) {
    const rowsMap = new Map();
    for (const f of tf) {
      const tr = chainFind(f.chain, a => a.role === 'TR');
      const td = chainFind(f.chain, a => a.role === 'TD' || a.role === 'TH');
      if (!tr || !td) continue;
      if (!rowsMap.has(tr.nid)) rowsMap.set(tr.nid, { nid: tr.nid, cells: new Map(), thead: chainHas(f.chain, 'THead') });
      const row = rowsMap.get(tr.nid);
      if (!row.cells.has(td.nid)) row.cells.set(td.nid, { nid: td.nid, th: td.role === 'TH', frags: [] });
      row.cells.get(td.nid).frags.push(f);
    }
    const rows = [...rowsMap.values()].sort((a, b) => a.nid - b.nid).map(r => ({ ...r, cells: [...r.cells.values()].sort((a, b) => a.nid - b.nid) }));
    if (!rows.length) continue;
    for (const r of rows) for (const c of r.cells) c.x = Math.min(...c.frags.map(f => f.x));
    const nCols = Math.max(...rows.map(r => r.cells.length));
    if (nCols < 2) continue; // single-column layout boxes are kept as ordinary text
    const full = rows.filter(r => r.cells.length === nCols);
    const colStarts = [];
    for (let k = 0; k < nCols; k++) {
      const v = full.map(r => r.cells[k].x).sort((a, b) => a - b);
      colStarts.push(v[Math.floor(v.length / 2)]);
    }
    const colOf = (x) => { let k = 0; while (k < nCols - 1 && x >= colStarts[k + 1] - 3) k++; return k; };
    const occ = [];   // occ[r][c] = cell
    const tableRows = [];
    rows.forEach((r, ri) => {
      occ[ri] = [];
      const starts = r.cells.map(c => colOf(c.x));
      // keep starts strictly increasing, else fall back to sequential placement
      const ok = starts.every((s, i) => i === 0 || s > starts[i - 1]);
      const st = ok ? starts : r.cells.map((_, i) => i);
      const row = [];
      r.cells.forEach((c, i) => {
        const end = i + 1 < st.length ? st[i + 1] : nCols;
        const cell = { blocks: cellBlocks(c.frags), rowspan: 1, colspan: Math.max(1, end - st[i]), header: r.thead, _frags: c.frags };
        for (let k = st[i]; k < st[i] + cell.colspan; k++) occ[ri][k] = cell;
        row.push(cell);
      });
      // holes before the first cell: a cell above spans down into this row
      for (let k = 0; k < (st[0] || 0); k++) {
        const above = ri > 0 ? occ[ri - 1][k] : null;
        if (above && !occ[ri][k]) {
          if (k === 0 || occ[ri - 1][k - 1] !== above) above.rowspan++;
          occ[ri][k] = above;
        }
      }
      tableRows.push(row);
    });
    if (!tableRows.some(r => r.length)) continue;
    if (!rows[0].thead) markHeader(tableRows);
    for (const r of tableRows) for (const c of r) delete c._frags;
    for (const f of tf) f.tableId = `${pageNum}-s${idx}`;
    log.stats.tables++;
    log.stats.cells += tableRows.reduce((s, r) => s + r.length, 0);
    log.stats.mergedCells += tableRows.reduce((s, r) => s + r.filter(c => c.rowspan > 1 || c.colspan > 1).length, 0);
    log.stats.inferredTables = (log.stats.inferredTables || []);
    log.stats.inferredTables.push(pageNum);
    const ys = tf.map(f => f.y);
    out.push({
      order: Math.min(...tf.map(f => f.order)),
      block: { type: 'table', rows: tableRows, meta: { page: pageNum, xs: [...colStarts, Math.max(...tf.map(f => f.x + f.w))], top: Math.min(...ys), bottom: Math.max(...ys), nCols, inferred: true } },
    });
    idx++;
    void tid;
  }
  return out;
}

function markHeader(rows) {
  if (rows.length < 2) return;
  const first = rows[0];
  const thead = first.every(c => c._frags.length && c._frags.every(f => chainHas(f.chain, 'THead')));
  const allBold = first.every(c => !c._frags.length || c._frags.every(f => f.bold)) && first.some(c => c._frags.length);
  const bodyBold = rows.slice(1).every(r => r.every(c => c._frags.length && c._frags.every(f => f.bold)));
  if (thead || (allBold && !bodyBold)) for (const c of first) c.header = true;
}

// ---------------- lines & inline text ----------------

function sameLine(a, b) {
  const size = Math.max(a.size, b.size);
  return Math.abs(a.y - b.y) < size * 0.55 && b.x >= a.x + a.w - size * 1.5;
}

/** Group fragments (in content-stream order) into visual lines. */
function buildLines(frags) {
  frags.sort((a, b) => a.order - b.order);
  const lines = [];
  let cur = null;
  for (const f of frags) {
    if (f.tableBlock || f.figure) { cur = null; lines.push(f.tableBlock ? { table: f.tableBlock, order: f.order } : { figure: true, alt: f.alt, chain: f.chain, order: f.order }); continue; }
    // a list item (label + body, possibly several paragraphs) is one block
    const block = chainFind(f.chain, a => a.role === 'LI' || a.role === 'TD' || a.role === 'TH')?.role === 'LI'
      ? chainFind(f.chain, a => a.role === 'LI')
      : chainFind(f.chain, a => BLOCK_ROLES.has(a.role) && a.role !== 'Lbl' && a.role !== 'LBody');
    const bid = block ? block.nid : null;
    if (cur && !cur.table && !cur.figure && sameLine(cur.frags[cur.frags.length - 1], f) && cur.bid === bid && cur.artifact === f.artifact && !f.rotated) {
      cur.frags.push(f);
    } else {
      cur = { frags: [f], bid, role: block ? block.role : null, chain: f.chain, artifact: f.artifact, order: f.order };
      lines.push(cur);
    }
  }
  for (const l of lines) if (l.frags) finishLine(l);
  return lines;
}

function finishLine(l) {
  const main = l.frags.reduce((m, f) => (f.size > m ? f.size : m), 0);
  l.size = main;
  const base = l.frags.find(f => f.size >= main * 0.9) || l.frags[0];
  l.y = base.y;
  l.x0 = Math.min(...l.frags.map(f => f.x));
  l.x1 = Math.max(...l.frags.map(f => f.x + f.w));
  l.bold = l.frags.every(f => f.bold || !/[\p{L}\p{N}]/u.test(f.str));
  l.inlines = fragsToInlines(l.frags, main);
  l.text = inlinesText(l.inlines);
  const lbl = l.frags.filter(f => chainHas(f.chain, 'Lbl'));
  l.label = lbl.length ? lbl.map(f => f.str).join('').trim() : null;
  l.listDepth = l.chain ? l.chain.filter(a => a.role === 'L').length : 0;
}

function fragsToInlines(frags, lineSize) {
  const out = [];
  let prev = null;
  for (const f of frags) {
    let t = f.str;
    if (prev) {
      const gap = f.x - (prev.x + prev.w);
      const raised = f.size < lineSize * 0.8 && f.y < prev.y - lineSize * 0.15;
      const needSpace = !raised && (prev.spaceAfter || gap > Math.max(prev.size, f.size) * 0.18) && !/\s$/.test(prev.str) && !/^\s/.test(t);
      if (needSpace) {
        // attach the space to the previous run so formatting runs stay contiguous
        if (out.length && out[out.length - 1].type === 'text') out[out.length - 1].text += ' ';
        else t = ' ' + t;
      }
    }
    out.push({ type: 'text', text: t, bold: f.bold || undefined, italic: f.italic || undefined });
    prev = f;
  }
  return mergeInlines(out);
}

/** Join several lines into one inline sequence (paragraph reflow). */
function joinLines(lines) {
  const out = [];
  lines.forEach((l, i) => {
    if (i) {
      const prevText = inlinesText(out);
      const next = l.text;
      if (/\S-$/u.test(prevText) && /^[\p{L}\p{N}]/u.test(next)) { /* hyphenated word or range: join */ }
      else if (!/\s$/.test(prevText)) {
        const last = out[out.length - 1];
        if (last && last.type === 'text') last.text += ' '; else out.push({ type: 'text', text: ' ' });
      }
    }
    out.push(...l.inlines);
  });
  return mergeInlines(out);
}

function cellBlocks(frags) {
  if (!frags.length) return [];
  const lines = buildLines(frags.map(f => ({ ...f })));
  const paras = [];
  let cur = null;
  for (const l of lines) {
    if (!l.frags) continue;
    const startsItem = l.label != null || (l.bid == null && BULLET_RE.test(l.text));
    const prev = cur && cur[cur.length - 1];
    const gap = prev ? l.y - prev.y : 0;
    const newPara = !cur || startsItem || (l.bid != null && l.bid !== prev.bid) || (l.bid == null && (gap > Math.max(l.size, prev.size) * 1.75 || gap < -2));
    if (newPara) { cur = [l]; paras.push(cur); } else cur.push(l);
  }
  return paras.map(p => ({ type: 'paragraph', inlines: trimInlines(joinLines(p)) })).filter(b => b.inlines.length);
}

// ======================================================================
// Pass 2: repeated header/footer detection (untagged PDFs)
// ======================================================================

function marginKey(text) {
  return text.toLowerCase().replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();
}

function findRepeatedMargins(pages) {
  const counts = new Map();
  for (const p of pages) {
    const seen = new Set();
    for (const l of p.lines) {
      if (!l.frags || l.artifact) continue;
      const band = l.y < p.height * 0.1 ? 'top' : l.y > p.height * 0.9 ? 'bottom' : null;
      if (!band) continue;
      const k = band + '|' + marginKey(l.text);
      if (!seen.has(k)) { seen.add(k); counts.set(k, (counts.get(k) || 0) + 1); }
    }
  }
  const min = Math.max(2, Math.ceil(pages.length * 0.4));
  const rep = new Set();
  if (pages.length >= 2) for (const [k, v] of counts) if (v >= min) rep.add(k);
  return rep;
}

// ======================================================================
// Pass 3: paragraphs and block classification
// ======================================================================

function groupParagraphs(page, env) {
  const { opts, exclude, repeated } = env;
  const out = [];
  let cur = null;   // current paragraph { lines, kind }
  const units = [];

  const marginOf = (l) => {
    const band = l.y < page.height * 0.12 ? 'top' : l.y > page.height * 0.88 ? 'bottom' : null;
    if (l.artifact) {
      if (PAGE_NO_RE.test(l.text.trim())) return 'Page number';
      return band === 'top' ? 'Header' : band === 'bottom' ? 'Footer' : 'Artifact';
    }
    if (!band) return null;
    if (PAGE_NO_RE.test(l.text.trim()) && (l.y < page.height * 0.08 || l.y > page.height * 0.9)) return 'Page number';
    if (repeated.has(band + '|' + marginKey(l.text))) return band === 'top' ? 'Header' : 'Footer';
    return null;
  };

  for (const l of page.lines) {
    if (l.table) { cur = null; units.push({ table: l.table }); continue; }
    if (l.figure) {
      cur = null;
      if (opts.removeImages) exclude('Image', l.alt, page.num);
      else units.push({ image: l.alt });
      continue;
    }
    const m = marginOf(l);
    if (m) {
      const remove = (m === 'Page number' && opts.removePageNumbers) || ((m === 'Header' || m === 'Footer') && opts.removeHeadersFooters) || (m === 'Artifact' && opts.removeHeadersFooters);
      if (remove) { exclude(m === 'Artifact' ? 'Header/footer (artifact)' : m, l.text, page.num); continue; }
    }
    const prev = cur && cur.lines[cur.lines.length - 1];
    const listStart = l.label != null || (l.bid == null && (BULLET_RE.test(l.text) || (ORDER_RE.test(l.text) && l.text.length > 3)));
    let cont = false;
    if (cur && !cur.margin && !m) {
      if (l.bid != null || prev.bid != null) cont = l.bid === prev.bid && l.label == null;
      else {
        const gap = l.y - prev.y;
        const size = Math.max(l.size, prev.size);
        cont = !listStart && gap > 0 && gap <= size * 1.6 &&
          Math.abs(l.size - prev.size) <= Math.max(0.6, size * 0.08) &&
          l.bold === prev.bold &&
          (cur.list ? l.x0 >= cur.textX - 3 : Math.abs(l.x0 - prev.x0) <= size * 1.2 || l.x0 < prev.x0) &&
          !(/[.:;!?]$/.test(prev.text.trim()) && prev.x1 < cur.maxX1 - size * 4);
      }
    }
    if (cont) { cur.lines.push(l); cur.maxX1 = Math.max(cur.maxX1, l.x1); }
    else {
      cur = { lines: [l], list: listStart, textX: textStartX(l), maxX1: l.x1, margin: !!m };
      units.push(cur);
    }
  }

  // classify
  for (const u of units) {
    if (u.table) { out.push(u.table); continue; }
    if (u.image != null) { out.push({ type: 'image', alt: u.image }); continue; }
    const first = u.lines[0];
    const role = first.role;
    const chain = first.chain;
    let inlines = trimInlines(joinLines(u.lines));
    const text = inlinesText(inlines).replace(/\s+/g, ' ').trim();
    if (!text) continue;

    // table of contents
    const isToc = role === 'TOCI' || chainHas(chain, 'TOC') || (u.lines.every(l => TOC_LINE_RE.test(l.text.trim())) && !first.label);
    if (isToc) { out.push({ type: '_toc', text, inlines, page: page.num, depth: chain ? chain.filter(a => a.role === 'TOC').length - 1 : 0 }); continue; }

    if (role === 'Note') {
      if (opts.removeFootnotes) { exclude('Footnote', text, page.num); continue; }
      out.push({ type: 'paragraph', inlines: [{ type: 'text', text: 'Note: ', italic: true }, ...inlines] });
      continue;
    }
    if (role === 'Caption') { out.push({ type: 'paragraph', style: 'caption', inlines: plain(inlines) }); continue; }
    if (role === 'BlockQuote' || role === 'Quote') { out.push({ type: 'paragraph', style: 'quote', inlines }); continue; }

    // headings
    let level = null;
    if (/^H[1-6]$/.test(role || '')) level = +role[1];
    else if (role === 'Title') level = 1;
    else if (role === 'H') level = headingLevelBySize(first.size, env, text);
    else if (role === 'P' && first.size >= env.bodySize * 1.5 && u.lines.length <= 3 && text.length <= 150) level = 1; // large title tagged as a paragraph
    else if (!env.hasHeadingTags && (!u.list || first.bold || first.size >= env.bodySize * 1.15)) level = guessHeading(u, text, env);
    if (level) { out.push({ type: 'heading', level, inlines: plain(inlines) }); continue; }

    // list items
    if (first.label != null || u.list) {
      let label = first.label;
      let body = inlines;
      if (label == null) {
        const m = BULLET_RE.exec(text) || ORDER_RE.exec(text);
        label = m ? m[1] : '•';
        body = stripPrefix(inlines, m ? m[0].length : 0);
      } else {
        body = stripPrefix(inlines, first.label.length);
      }
      const ordered = ORDER_RE.test(label + ' ') && !BULLET_RE.test(label + ' ');
      const depth = first.listDepth ? first.listDepth - 1 : null;
      out.push({ type: '_li', ordered, label: ordered ? label : undefined, depth, x: first.x0, inlines: trimInlines(body) });
      continue;
    }
    out.push({ type: 'paragraph', inlines });
  }
  return out;
}

function textStartX(l) {
  // x where the text of a list item starts (after the bullet)
  if (l.frags.length > 1 && (BULLET_RE.test(l.frags[0].str + ' ') || ORDER_RE.test(l.frags[0].str + ' '))) return l.frags[1].x;
  return l.x0;
}

function stripPrefix(inlines, n) {
  const out = [];
  let skip = n;
  for (const r of inlines) {
    if (skip > 0 && r.type === 'text') {
      if (r.text.length <= skip) { skip -= r.text.length; continue; }
      out.push({ ...r, text: r.text.slice(skip) });
      skip = 0;
    } else out.push(r);
  }
  return trimInlines(out);
}

function plain(inlines) {
  return [{ type: 'text', text: inlinesText(inlines).replace(/\s+/g, ' ').trim() }];
}

function headingLevelBySize(size, env, text) {
  const num = /^(\d+(?:\.\d+)*)\.?\s/.exec(text);
  if (num) return Math.min(6, num[1].split('.').length);
  const i = env.headingSizes.findIndex(s => Math.abs(s - size) < 0.6);
  return i >= 0 ? Math.min(6, i + 1) : Math.min(6, env.headingSizes.length + 1);
}

function guessHeading(u, text, env) {
  if (u.lines.length > 3 || text.length > 160 || /[,;]$/.test(text)) return null;
  const first = u.lines[0];
  const big = first.size >= env.bodySize * 1.15;
  const bold = u.lines.every(l => l.bold);
  const numbered = /^(\d+(\.\d+)*\.?|[A-Z]\.|(appendix|annex|section|chapter|part)\s+[\w.]+)\s+\S/i.exec(text);
  if (!big && !(bold && numbered && !/[.:]$/.test(text) && text.length <= 120)) return null;
  if (numbered && /^\d/.test(text)) return Math.min(6, text.split(/\s/)[0].replace(/\.$/, '').split('.').length);
  if (big) {
    const i = env.headingSizes.findIndex(s => Math.abs(s - first.size) < 0.6);
    return Math.min(6, (i >= 0 ? i : env.headingSizes.length) + 1);
  }
  return Math.min(6, env.headingSizes.length + 1);
}

// ======================================================================
// Post-processing: lists, TOC, multi-page tables
// ======================================================================

function postProcess(blocks, opts, log, exclude) {
  // 1. table of contents
  const out1 = [];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.type === '_toc') {
      if (opts.removeToc) { exclude('Table of contents', b.text, b.page); continue; }
      out1.push({ type: '_li', ordered: false, depth: Math.max(0, b.depth || 0), inlines: b.inlines });
      continue;
    }
    out1.push(b);
  }
  // remove a "Contents" / "List of Tables" title left directly before a removed TOC
  if (opts.removeToc) {
    for (let i = out1.length - 1; i >= 0; i--) {
      const b = out1[i];
      if ((b.type === 'heading' || b.type === 'paragraph') && /^(table of )?contents$|^list of (tables|figures|plates|appendices|annexes|drawings)$/i.test(inlinesText(b.inlines).trim())) {
        const idx = blocks.indexOf(b);
        const next = blocks.slice(idx + 1).find(x => x.type !== 'pagebreak');
        if (next && next.type === '_toc') { exclude('Table of contents', inlinesText(b.inlines)); out1.splice(i, 1); }
      }
    }
  }

  // 1b. a paragraph or list item that continues on the next page
  for (let i = 0; i < out1.length; i++) {
    const b = out1[i];
    if (b.type !== 'pagebreak') continue;
    let k = i - 1;
    while (k >= 0 && out1[k].type === 'pagebreak') k--;
    let j = i + 1;
    while (j < out1.length && out1[j].type === 'pagebreak') j++;
    const a = out1[k], c = out1[j];
    if (!a || !c || c.type !== 'paragraph' || c.style) continue;
    if (a.type !== 'paragraph' && a.type !== '_li') continue;
    const at = inlinesText(a.inlines).trim(), ct = inlinesText(c.inlines).trim();
    if (!/^[\p{Ll}]/u.test(ct) || /[.!?:]$/.test(at)) continue;
    a.inlines = mergeInlines([...a.inlines, { type: 'text', text: ' ' }, ...c.inlines]);
    out1.splice(j, 1);
  }

  // 2. lists: consecutive _li blocks (page markers between items are moved after the list)
  const out2 = [];
  for (let i = 0; i < out1.length; i++) {
    if (out1[i].type !== '_li') { out2.push(out1[i]); continue; }
    const items = [];
    const deferred = [];
    let j = i;
    for (; j < out1.length; j++) {
      const b = out1[j];
      if (b.type === '_li') items.push(b);
      else if (b.type === 'pagebreak' && out1.slice(j + 1).find(x => x.type !== 'pagebreak')?.type === '_li') deferred.push(b);
      else break;
    }
    // depth from tags, else from indentation
    const xs = cluster(items.filter(it => it.depth == null).map(it => it.x), 6);
    const flat = items.map(it => ({
      level: it.depth != null ? it.depth : Math.max(0, xs.findIndex(x => Math.abs(x - it.x) <= 6)),
      ordered: it.ordered, label: it.label,
      blocks: [{ type: 'paragraph', inlines: it.inlines }],
    }));
    out2.push(...buildLists(flat), ...deferred);
    i = j - 1;
  }

  // 3. tables continued across a page break
  const out3 = [];
  for (const b of out2) {
    if (b.type === 'table') {
      let k = out3.length - 1;
      const between = [];
      while (k >= 0 && out3[k].type === 'pagebreak') { between.push(out3[k]); k--; }
      const prev = out3[k];
      if (prev && prev.type === 'table' && between.length && prev.meta && b.meta && b.meta.page === prev.meta.page + between.length &&
          prev.meta.nCols === b.meta.nCols && prev.meta.xs.length === b.meta.xs.length &&
          prev.meta.xs.every((x, i) => Math.abs(x - b.meta.xs[i]) <= 3)) {
        let rows = b.rows;
        const sig = (r) => r.map(c => blocksText(c.blocks).trim()).join('|');
        if (rows.length && prev.rows.length && sig(rows[0]) === sig(prev.rows[0]) && sig(rows[0]).replace(/\|/g, '').trim()) {
          exclude('Repeated table header row', sig(rows[0]).replace(/\|/g, ' '), b.meta.page);
          rows = rows.slice(1);
        }
        prev.rows.push(...rows);
        prev.meta.bottom = b.meta.bottom;
        prev.meta.continued = (prev.meta.continued || 1) + 1;
        log.stats.tables--;
        log.stats.pageMarkersJoined = (log.stats.pageMarkersJoined || 0) + (out3.length - (k + 1));
        out3.splice(k + 1); // drop the page markers inside the table
        continue;
      }
    }
    out3.push(b);
  }
  return out3;
}

/** Render a page to a canvas (used by the side-by-side review). */
export async function renderPdfPage(buf, pageNum, canvas, maxWidth = 900) {
  const pdf = await openPdf(buf);
  try {
    const page = await pdf.getPage(pageNum);
    const vp0 = page.getViewport({ scale: 1 });
    const scale = Math.min(2, maxWidth / vp0.width) * (window.devicePixelRatio || 1);
    const vp = page.getViewport({ scale });
    canvas.width = vp.width; canvas.height = vp.height;
    canvas.style.width = (vp.width / (window.devicePixelRatio || 1)) + 'px';
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    return pdf.numPages;
  } finally {
    await pdf.destroy();
  }
}

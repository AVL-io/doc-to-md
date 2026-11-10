// Small helpers for reading Office Open XML packages (docx / xlsx).
import JSZip from 'jszip';

export const NS = {
  w: ['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main'],
  r: ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'],
  m: ['http://schemas.openxmlformats.org/officeDocument/2006/math', 'http://purl.oclc.org/ooxml/officeDocument/math'],
  mc: ['http://schemas.openxmlformats.org/markup-compatibility/2006'],
  wp: ['http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing'],
  s: ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'],
};

export function is(el, ns, name) {
  return !!el && el.nodeType === 1 && el.localName === name && NS[ns].includes(el.namespaceURI);
}
export function inNs(el, ns) {
  return !!el && el.nodeType === 1 && NS[ns].includes(el.namespaceURI);
}
export function kids(el, ns, name) {
  const out = [];
  if (!el) return out;
  for (const c of el.children) if (is(c, ns, name)) out.push(c);
  return out;
}
export function kid(el, ns, name) {
  if (!el) return null;
  for (const c of el.children) if (is(c, ns, name)) return c;
  return null;
}
/** Descendant search by namespace + local name (document order). */
export function desc(el, ns, name) {
  const out = [];
  if (!el) return out;
  for (const uri of NS[ns]) for (const e of el.getElementsByTagNameNS(uri, name)) out.push(e);
  if (NS[ns].length > 1) out.sort((a, b) => (a.compareDocumentPosition(b) & 4 ? -1 : 1));
  return out;
}
/** Attribute in a namespace (falls back to any attribute with that local name). */
export function attr(el, ns, name) {
  if (!el) return null;
  for (const uri of NS[ns]) {
    const v = el.getAttributeNS(uri, name);
    if (v != null && v !== '') return v;
  }
  for (const a of el.attributes) if (a.localName === name) return a.value;
  return null;
}
export const wVal = (el) => attr(el, 'w', 'val');

/** OOXML on/off property: present without val, or val true/1/on. */
export function onOff(el) {
  if (!el) return false;
  const v = wVal(el);
  return v == null || !/^(0|false|off|none)$/i.test(v);
}

export function parseXml(str) {
  const doc = new DOMParser().parseFromString(str, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('The file contains malformed XML and could not be read.');
  return doc;
}

export async function openPackage(buf) {
  let zip;
  try {
    zip = await JSZip.loadAsync(buf);
  } catch (e) {
    throw new Error('The file is not a valid Office document (it could not be opened as a ZIP package). It may be password-protected or saved in the older .doc/.xls format.');
  }
  if (zip.file('EncryptionInfo') || zip.file('EncryptedPackage')) {
    throw new Error('The file is password-protected. Remove the password in Word/Excel and try again.');
  }
  const cache = new Map();
  const pkg = {
    zip,
    has: (p) => !!zip.file(p),
    async text(p) {
      const f = zip.file(p);
      return f ? f.async('string') : null;
    },
    async xml(p) {
      if (cache.has(p)) return cache.get(p);
      const t = await pkg.text(p);
      const d = t == null ? null : parseXml(t);
      cache.set(p, d);
      return d;
    },
    /** Relationships of a part: id -> { type, target (resolved path), external } */
    async rels(partPath) {
      const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/') + 1) : '';
      const relPath = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
      const d = await pkg.xml(relPath);
      const map = {};
      if (!d) return map;
      for (const r of d.getElementsByTagName('Relationship')) {
        const external = r.getAttribute('TargetMode') === 'External';
        const target = r.getAttribute('Target') || '';
        map[r.getAttribute('Id')] = {
          type: (r.getAttribute('Type') || '').split('/').pop(),
          target: external ? target : resolvePath(dir, target),
          external,
        };
      }
      return map;
    },
  };
  return pkg;
}

export function resolvePath(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dir + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export async function mainPartPath(pkg, fallback) {
  const rels = await pkg.rels('');
  for (const r of Object.values(rels)) if (r.type === 'officeDocument') return r.target;
  return fallback;
}

/** Map Wingdings/Symbol private-use bullet code points to readable characters. */
export function symbolChar(code) {
  const map = { 0xF0B7: '•', 0xF0A7: '▪', 0xF0D8: '➢', 0xF076: '❖', 0xF0FC: '✓', 0xF0A8: '□', 0xF06E: '■', 0xF0E0: '→', 0xF0E8: '➔', 0xF0D7: '▸', 0xF02D: '-', 0xF0B0: '°', 0xF0B1: '±', 0xF06D: 'µ', 0xF0A3: '≤', 0xF0B3: '≥' };
  if (map[code]) return map[code];
  if (code >= 0xF020 && code <= 0xF0FF) return '';
  return String.fromCodePoint(code);
}

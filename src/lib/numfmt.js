// Minimal Excel number-format renderer: shows cell values the way Excel displays them
// for the formats that occur in practice (general, fixed decimals, thousands
// separators, percentages, scientific, currency literals, dates and times).

export const BUILTIN = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  12: '# ?/?', 13: '# ??/??', 14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy-mm-dd h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
};

export function formatGeneral(v) {
  if (!isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const a = Math.abs(v);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) {
    const [m, e] = v.toExponential(5).split('e');
    const exp = parseInt(e, 10);
    return `${m.replace(/\.?0+$/, '')}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  let s = v.toPrecision(11);
  if (s.includes('e')) s = String(Number(s));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

function splitSections(fmt) {
  const out = [];
  let cur = '', q = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') q = !q;
    if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

/** Strip colour / locale / condition brackets but keep currency symbols. */
function cleanSection(s) {
  return s.replace(/\[\$([^\]-]*)(-[^\]]*)?\]/g, (m, sym) => sym ? `"${sym}"` : '')
          .replace(/\[(?!h\]|hh\]|m\]|mm\]|s\]|ss\])[^\]]*\]/gi, '');
}

function isDateFormat(s) {
  const stripped = s.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, m => /^\[(h+|m+|s+)\]$/i.test(m) ? 'h' : '');
  return /[ymdhs]/i.test(stripped) && !/[0#?]/.test(stripped.replace(/[ymdhs.:\/\- ]+/gi, ''));
}

export function formatValue(v, fmt, date1904 = false) {
  if (typeof v !== 'number') return String(v);
  if (!fmt || /^general$/i.test(fmt)) return formatGeneral(v);
  if (fmt === '@') return formatGeneral(v);
  const sections = splitSections(fmt).map(cleanSection);
  let sec = sections[0];
  let val = v;
  let negHandled = false;
  if (v < 0 && sections.length > 1 && sections[1] !== '') { sec = sections[1]; val = -v; negHandled = true; }
  else if (v === 0 && sections.length > 2 && sections[2] !== '') { sec = sections[2]; }
  if (isDateFormat(sec)) return formatDate(v, sec, date1904);
  try {
    let s = formatNumberSection(Math.abs(val), sec);
    if (val < 0 && !negHandled && /\d/.test(s) && /[1-9]/.test(s)) s = '-' + s;
    return s;
  } catch {
    return formatGeneral(v);
  }
}

function formatNumberSection(a, sec) {
  // literal pieces
  let lit = '';
  const toks = [];
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); lit += sec.slice(i + 1, j < 0 ? undefined : j); i = j < 0 ? sec.length : j; continue; }
    if (ch === '\\') { lit += sec[i + 1] || ''; i++; continue; }
    if (ch === '_') { lit += ' '; i++; continue; }
    if (ch === '*') { i++; continue; }
    const isPh = '0#?.,%/'.includes(ch) || ((ch === 'E' || ch === 'e') && /[+-]/.test(sec[i + 1] || '')) ||
      ((ch === '+' || ch === '-') && /[Ee]/.test(sec[i - 1] || ''));
    if (isPh) {
      if (lit) { toks.push({ lit }); lit = ''; }
      toks.push({ ph: ch });
      continue;
    }
    if (ch === '@') { lit += ''; continue; }
    lit += ch;
  }
  if (lit) toks.push({ lit });
  const pattern = toks.filter(t => t.ph).map(t => t.ph).join('');
  if (!pattern) return toks.map(t => t.lit).join('');
  const percent = (pattern.match(/%/g) || []).length;
  let x = a * Math.pow(100, percent);
  let core;
  if (/[Ee]/.test(pattern)) {
    const mant = pattern.split(/[Ee]/)[0];
    const dec = (mant.split('.')[1] || '').replace(/[^0#?]/g, '').length;
    const [m, e] = x.toExponential(dec).split('e');
    const en = parseInt(e, 10);
    core = `${m}E${en < 0 ? '-' : '+'}${String(Math.abs(en)).padStart(2, '0')}`;
  } else if (pattern.includes('/')) {
    core = formatGeneral(x);
  } else {
    const num = pattern.replace(/%/g, '');
    const [intPat, decPat = ''] = num.split('.');
    const trailingCommas = (intPat.match(/,+$/) || [''])[0].length;
    x = x / Math.pow(1000, trailingCommas);
    const minDec = (decPat.match(/0/g) || []).length;
    const maxDec = (decPat.match(/[0#?]/g) || []).length;
    let s = x.toFixed(maxDec);
    if (maxDec > minDec) {
      let [ip, dp = ''] = s.split('.');
      while (dp.length > minDec && dp.endsWith('0')) dp = dp.slice(0, -1);
      s = dp ? `${ip}.${dp}` : ip;
    }
    let [ip, dp] = s.split('.');
    const minInt = (intPat.match(/0/g) || []).length;
    if (ip === '0' && minInt === 0) ip = '';
    ip = ip.padStart(minInt, '0');
    if (/,/.test(intPat.replace(/,+$/, ''))) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    core = dp != null ? `${ip}.${dp}` : ip;
    if (core === '' ) core = '0';
  }
  // place literals: prefix = literals before first placeholder, suffix = after last
  const firstPh = toks.findIndex(t => t.ph);
  let lastPh = -1;
  toks.forEach((t, i) => { if (t.ph) lastPh = i; });
  const pre = toks.slice(0, firstPh).map(t => t.lit).join('');
  const mid = toks.slice(firstPh, lastPh + 1).filter(t => t.lit).map(t => t.lit).join('');
  const post = toks.slice(lastPh + 1).map(t => t.lit).join('');
  return (pre + core + mid + '%'.repeat(percent) + post).replace(/\s+$/, '');
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

export function serialToDate(v, date1904 = false) {
  let days = Math.floor(v);
  const frac = v - days;
  if (date1904) days += 1462;
  // Excel's 1900 leap-year bug: serial 60 is 29-Feb-1900, which does not exist.
  const base = Date.UTC(1899, 11, 31);
  let ms = base + (days > 59 && !date1904 ? days - 1 : days) * 86400000;
  ms += Math.round(frac * 86400) * 1000;
  return new Date(ms);
}

function formatDate(v, sec, date1904) {
  const d = serialToDate(v, date1904);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), wd = d.getUTCDay();
  let h = d.getUTCHours();
  const mi = d.getUTCMinutes(), s = d.getUTCSeconds();
  const ampm = /AM\/PM|A\/P/i.test(sec);
  const totalHours = Math.floor(v * 24);
  let out = '';
  const pad = (n, l = 2) => String(n).padStart(l, '0');
  let lastWasHour = false;
  for (let i = 0; i < sec.length;) {
    const rest = sec.slice(i);
    let m;
    if (rest[0] === '"') { const j = sec.indexOf('"', i + 1); out += sec.slice(i + 1, j < 0 ? undefined : j); i = j < 0 ? sec.length : j + 1; continue; }
    if (rest[0] === '\\') { out += rest[1] || ''; i += 2; continue; }
    if ((m = /^\[h+\]/i.exec(rest))) { out += totalHours; i += m[0].length; lastWasHour = true; continue; }
    if ((m = /^\[m+\]/i.exec(rest))) { out += Math.floor(v * 1440); i += m[0].length; continue; }
    if ((m = /^(AM\/PM|A\/P)/i.exec(rest))) { out += h < 12 ? 'AM' : 'PM'; i += m[0].length; continue; }
    if ((m = /^y+/i.exec(rest))) { out += m[0].length <= 2 ? pad(Y % 100) : Y; i += m[0].length; continue; }
    if ((m = /^d+/i.exec(rest))) {
      const n = m[0].length;
      out += n === 1 ? D : n === 2 ? pad(D) : n === 3 ? DAYS[wd].slice(0, 3) : DAYS[wd];
      i += n; continue;
    }
    if ((m = /^h+/i.exec(rest))) {
      let hh = h;
      if (ampm) { hh = h % 12; if (hh === 0) hh = 12; }
      out += m[0].length >= 2 ? pad(hh) : hh;
      i += m[0].length; lastWasHour = true; continue;
    }
    if ((m = /^m+/i.exec(rest))) {
      const n = m[0].length;
      const minuteCtx = lastWasHour || /^m+[^a-z]*s/i.test(rest);
      if (minuteCtx && n <= 2) out += n === 2 ? pad(mi) : mi;
      else out += n === 1 ? M + 1 : n === 2 ? pad(M + 1) : n === 3 ? MONTHS[M].slice(0, 3) : n === 5 ? MONTHS[M][0] : MONTHS[M];
      i += n; lastWasHour = false; continue;
    }
    if ((m = /^s+(\.0+)?/i.exec(rest))) {
      out += pad(s);
      if (m[1]) out += '.' + '0'.repeat(m[1].length - 1);
      i += m[0].length; continue;
    }
    if (rest[0] === '_' ) { out += ' '; i += 2; continue; }
    if (rest[0] === '*') { i += 2; continue; }
    out += rest[0];
    i++;
  }
  return out.trim();
}

/* Formly browser engine — replaces FastAPI + openpyxl + LibreOffice.
 * XLSX is edited at the XML level (JSZip + DOMParser) so original styles, merges,
 * column widths and print settings survive. PDF is rendered from an HTML replica
 * of the sheet (html2canvas + jsPDF), so Korean text renders without embedded fonts. */
'use strict';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_XML = 'http://www.w3.org/XML/1998/namespace';

/* ---------- address helpers ---------- */
function colLetters(c) { let s = ''; while (c) { c--; s = String.fromCharCode(65 + (c % 26)) + s; c = Math.floor(c / 26); } return s; }
function colIndex(l) { let n = 0; for (const ch of l.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(a || '').trim());
  if (!m) return null;
  const c = colIndex(m[1]), r = +m[2];
  if (r < 1 || r > 1048576 || c < 1 || c > 16384) return null;
  return { r, c };
}
const addrOf = (r, c) => colLetters(c) + r;
function normAddr(a) { const p = parseAddr(a); return p ? addrOf(p.r, p.c) : null; }
function parseRange(ref) {
  const [a, b] = String(ref || '').split(':'); const p = parseAddr(a), q = parseAddr(b || a);
  return p && q ? { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) } : null;
}

/* ---------- XML helpers ---------- */
function parseXML(s) {
  const d = new DOMParser().parseFromString(String(s).replace(/^\uFEFF/, '').trimStart(), 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw Error('XML을 해석할 수 없습니다.');
  return d;
}
const kids = (el, n) => el ? [...el.children].filter(c => c.localName === n) : [];
const kid = (el, n) => kids(el, n)[0] || null;
const desc = (el, n) => el ? [...el.getElementsByTagNameNS('*', n)] : [];
const at = (el, n) => el ? el.getAttribute(n) : null;
const truthy = v => v === '1' || v === 'true';
function mk(doc, ref, name) { return doc.createElementNS(NS_MAIN, ref.prefix ? ref.prefix + ':' + name : name); }
function serialize(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}
function resolvePart(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base ? base.split('/') : ['']; parts.pop();
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
  return parts.join('/');
}
async function readRels(zip, part) {
  const relPath = part ? part.replace(/([^/]*)$/, '_rels/$1.rels') : '_rels/.rels';
  const f = zip.file(relPath); if (!f) return [];
  return desc(parseXML(await f.async('string')), 'Relationship').map(r => ({
    id: at(r, 'Id'), type: at(r, 'Type') || '', external: at(r, 'TargetMode') === 'External',
    target: resolvePart(part, at(r, 'Target') || ''),
  }));
}
const richText = el => desc(el, 't').filter(t => t.parentNode.localName !== 'rPh').map(t => t.textContent).join('');

/* ---------- colors ---------- */
const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
function applyTint(hex, tint) {
  if (!tint) return hex;
  let [r, g, b] = [1, 3, 5].map(i => parseInt(hex.substr(i, 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b); let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) { const d = max - min; s = l > .5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (!s) r = g = b = l; else { const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return '#' + [r, g, b].map(v => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
}
function colorOf(el, theme) {
  if (!el || truthy(at(el, 'auto'))) return null;
  let hex = null;
  const rgb = at(el, 'rgb');
  if (rgb) hex = '#' + rgb.slice(-6);
  else if (at(el, 'theme') != null) hex = theme[+at(el, 'theme')] || null;
  else if (at(el, 'indexed') != null) { const v = INDEXED[+at(el, 'indexed')]; hex = v ? '#' + v : null; }
  return hex ? applyTint(hex, +(at(el, 'tint') || 0)) : null;
}
function parseTheme(doc) {
  const scheme = desc(doc, 'clrScheme')[0]; if (!scheme) return [];
  const list = [...scheme.children].map(ch => { const c = ch.firstElementChild; return c ? '#' + (at(c, 'lastClr') || at(c, 'val') || '000000').slice(-6) : '#000000'; });
  return [list[1], list[0], list[3], list[2], ...list.slice(4)];
}

/* ---------- styles & number formats ---------- */
const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy-mm-dd h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 49: '@' };
function isDateFmt(id, code) {
  if ((id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58)) return true;
  if (!code || /general/i.test(code)) return false;
  const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
  return /[ymdhs]/i.test(c) && !/[#0?]/.test(c.replace(/\.0+/, ''));
}
function parseStyles(doc, theme) {
  const empty = { fonts: [{ sz: 11 }], fills: [{}], borders: [{}], xfs: [{}], numFmts: {} };
  if (!doc) return empty;
  const root = doc.documentElement, color = el => colorOf(el, theme);
  const numFmts = {};
  kids(kid(root, 'numFmts'), 'numFmt').forEach(n => { numFmts[at(n, 'numFmtId')] = at(n, 'formatCode'); });
  const flag = (f, n) => { const e = kid(f, n); return !!e && at(e, 'val') !== '0' && at(e, 'val') !== 'false'; };
  const fonts = kids(kid(root, 'fonts'), 'font').map(f => ({
    b: flag(f, 'b'), i: flag(f, 'i'), strike: flag(f, 'strike'), u: !!kid(f, 'u') && at(kid(f, 'u'), 'val') !== 'none',
    sz: +at(kid(f, 'sz'), 'val') || 11, color: color(kid(f, 'color')), name: at(kid(f, 'name'), 'val') || '',
  }));
  const fills = kids(kid(root, 'fills'), 'fill').map(f => {
    const p = kid(f, 'patternFill');
    if (p) { const t = at(p, 'patternType'); if (!t || t === 'none' || t === 'gray125') return {};
      return { bg: color(kid(p, 'fgColor')) || color(kid(p, 'bgColor')) }; }
    const g = kid(f, 'gradientFill'); if (g) { const st = desc(g, 'stop')[0]; return { bg: color(kid(st, 'color')) }; }
    return {};
  });
  const borders = kids(kid(root, 'borders'), 'border').map(b => {
    const o = {};
    for (const side of ['left', 'right', 'top', 'bottom']) { const e = kid(b, side), st = at(e, 'style');
      if (st && st !== 'none') o[side] = { style: st, color: color(kid(e, 'color')) || '#000000' }; }
    return o;
  });
  const xfs = kids(kid(root, 'cellXfs'), 'xf').map(x => { const a = kid(x, 'alignment'); return {
    numFmtId: +at(x, 'numFmtId') || 0, fontId: +at(x, 'fontId') || 0, fillId: +at(x, 'fillId') || 0, borderId: +at(x, 'borderId') || 0,
    h: at(a, 'horizontal'), v: at(a, 'vertical'), wrap: truthy(at(a, 'wrapText')), indent: +at(a, 'indent') || 0, rot: +at(a, 'textRotation') || 0,
  }; });
  return { fonts: fonts.length ? fonts : empty.fonts, fills, borders, xfs: xfs.length ? xfs : empty.xfs, numFmts };
}
function fmtCode(book, s) { const id = (book.styles.xfs[s] || {}).numFmtId || 0; return { id, code: book.styles.numFmts[id] || BUILTIN_FMT[id] || 'General' }; }
function splitSections(code) {
  const out = []; let cur = '', q = false;
  for (let i = 0; i < code.length; i++) { const ch = code[i];
    if (ch === '"') q = !q; if (ch === '\\' && !q) { cur += ch + (code[++i] || ''); continue; }
    if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; }
  out.push(cur); return out;
}
function generalNum(v) { return Number.isInteger(v) ? String(v) : String(+v.toPrecision(11)); }
const WEEK_KO = '일월화수목금토', WEEK_EN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], MON_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function formatDate(v, sec) {
  const d = new Date(Math.round((v - 25569) * 86400000)); const p2 = n => String(n).padStart(2, '0');
  const toks = sec.replace(/\[[^\]]*\]/g, '').match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p|[eE]|_.|\*.|[\s\S]/g) || [];
  const isH = t => /^h/i.test(t), isS = t => /^s/i.test(t), ampm = toks.some(t => /^(am\/pm|a\/p)$/i.test(t));
  let out = '';
  toks.forEach((t, i) => {
    const prev = toks.slice(0, i).reverse().find(x => /^[a-z]/i.test(x) && !/^(am\/pm|a\/p)$/i.test(x));
    const next = toks.slice(i + 1).find(x => /^[a-z]/i.test(x) && !/^(am\/pm|a\/p)$/i.test(x));
    const lt = t.toLowerCase(); const H = d.getUTCHours(), h12 = H % 12 || 12;
    if (t[0] === '"') out += t.slice(1, -1); else if (t[0] === '\\') out += t[1]; else if (t[0] === '_') out += ' '; else if (t[0] === '*') out += '';
    else if (lt === 'yyyy' || lt === 'e') out += d.getUTCFullYear(); else if (lt === 'yy') out += p2(d.getUTCFullYear() % 100);
    else if (lt === 'mmmmm') out += MON_EN[d.getUTCMonth()][0]; else if (lt === 'mmmm' || lt === 'mmm') out += MON_EN[d.getUTCMonth()];
    else if (lt === 'mm' || lt === 'm') { const minute = (prev && isH(prev)) || (next && isS(next)); out += minute ? (lt === 'mm' ? p2(d.getUTCMinutes()) : d.getUTCMinutes()) : (lt === 'mm' ? p2(d.getUTCMonth() + 1) : d.getUTCMonth() + 1); }
    else if (lt === 'dddd' || lt === 'ddd') out += WEEK_EN[d.getUTCDay()]; else if (lt === 'dd') out += p2(d.getUTCDate()); else if (lt === 'd') out += d.getUTCDate();
    else if (lt === 'aaaa') out += WEEK_KO[d.getUTCDay()] + '요일'; else if (lt === 'aaa') out += WEEK_KO[d.getUTCDay()];
    else if (lt === 'hh') out += p2(ampm ? h12 : H); else if (lt === 'h') out += ampm ? h12 : H;
    else if (lt === 'ss') out += p2(d.getUTCSeconds()); else if (lt === 's') out += d.getUTCSeconds();
    else if (lt === 'am/pm') out += H < 12 ? 'AM' : 'PM'; else if (lt === 'a/p') out += H < 12 ? 'A' : 'P';
    else out += t;
  });
  return out;
}
function formatNumber(v, sec, negSection) {
  sec = sec.replace(/\[[^\]]*\]/g, '');
  if (/^\s*general\s*$/i.test(sec)) return (v < 0 && !negSection ? '-' : '') + generalNum(Math.abs(v));
  const toks = sec.match(/"[^"]*"|\\.|_.|\*.|[#0?][#0?,]*(?:\.[#0?]*)?|\.[#0?]+|[\s\S]/g) || [];
  let lit = (t) => t[0] === '"' ? t.slice(1, -1) : t[0] === '\\' ? t[1] : t[0] === '_' ? ' ' : t[0] === '*' ? '' : t;
  const pct = toks.some(t => t === '%'); let x = Math.abs(v) * (pct ? 100 : 1);
  const sci = toks.findIndex(t => /^E[+-]?$/i.test(t));
  let out = '', done = false;
  for (const t of toks) {
    if (!done && /^[#0?.]/.test(t) && /[#0?]/.test(t)) {
      done = true;
      const [ip, fp = ''] = t.split('.'); const dec = (fp.match(/[0#?]/g) || []).length, minDec = (fp.match(/0/g) || []).length;
      let str = x.toFixed(dec);
      if (dec > minDec) str = str.replace(new RegExp(`0{1,${dec - minDec}}$`), '').replace(/\.$/, '');
      let [i, f] = str.split('.');
      if (/,/.test(ip.replace(/,+$/, ''))) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      if (i === '0' && !/0/.test(ip)) i = '';
      out += i + (f != null ? '.' + f : (dec && minDec ? '' : ''));
    } else if (sci >= 0 && /^E/i.test(t)) { /* handled below */ } else out += lit(t);
  }
  if (sci >= 0) return (v < 0 ? '-' : '') + x.toExponential(2).toUpperCase();
  if (!done) out += '';
  return (v < 0 && !negSection && x !== 0 && /\d/.test(out) ? '-' : '') + out;
}
function formatValue(book, cell) {
  const v = cell.value; if (v == null || v === '') return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  let { id, code } = fmtCode(book, cell.s);
  if (/\[\$-F800\]/i.test(code)) code = 'yyyy"년" m"월" d"일" aaaa';
  else if (/\[\$-F400\]/i.test(code)) code = 'AM/PM h:mm:ss';
  if (typeof v !== 'number') {
    const secs = splitSections(code); const tsec = secs.length >= 4 ? secs[3] : (secs.length === 1 && /@/.test(secs[0]) ? secs[0] : null);
    if (!tsec) return String(v);
    return (tsec.match(/"[^"]*"|\\.|_.|\*.|@|[\s\S]/g) || []).map(t => t === '@' ? String(v) : t[0] === '"' ? t.slice(1, -1) : t[0] === '\\' ? t[1] : t[0] === '_' ? ' ' : t[0] === '*' ? '' : t).join('').trim();
  }
  try {
    const secs = splitSections(code);
    let sec = secs[0], neg = false;
    if (v < 0 && secs.length > 1 && secs[1] !== '') { sec = secs[1]; neg = true; } else if (v === 0 && secs.length > 2) sec = secs[2];
    if (isDateFmt(id, sec)) return formatDate(v, BUILTIN_FMT[id] && !book.styles.numFmts[id] ? 'yyyy-mm-dd' + (/h/i.test(BUILTIN_FMT[id]) ? ' hh:mm' : '') : sec);
    if (/^\s*@\s*$/.test(sec)) return generalNum(v);
    return formatNumber(v, sec, neg);
  } catch { return generalNum(v); }
}

/* ---------- workbook ---------- */
async function openBook(bytes) {
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { throw Error('열 수 없는 Excel 파일입니다. 올바른 .xlsx 파일을 업로드하세요.'); }
  const rootRels = await readRels(zip, '');
  const wbPath = (rootRels.find(r => /\/officeDocument$/.test(r.type)) || { target: 'xl/workbook.xml' }).target;
  if (!zip.file(wbPath)) throw Error('열 수 없는 Excel 파일입니다. (workbook.xml 없음)');
  const wbDoc = parseXML(await zip.file(wbPath).async('string'));
  const rels = await readRels(zip, wbPath);
  const byId = Object.fromEntries(rels.map(r => [r.id, r]));
  const rawSheets = desc(wbDoc, 'sheet').map((s, index) => {
    const rel = byId[s.getAttributeNS(NS_R, 'id') || at(s, 'r:id')];
    return { index, name: at(s, 'name'), path: rel && rel.target, state: at(s, 'state') || 'visible', worksheet: !!rel && /\/worksheet$/.test(rel.type) };
  });
  const sheets = rawSheets.filter(s => s.worksheet && s.path && zip.file(s.path));
  if (!sheets.length) throw Error('워크시트가 없는 파일입니다.');
  const part = re => rels.find(r => re.test(r.type) && zip.file(r.target));
  const themeRel = part(/\/theme$/), stylesRel = part(/\/styles$/), sstRel = part(/\/sharedStrings$/);
  const theme = themeRel ? parseTheme(parseXML(await zip.file(themeRel.target).async('string'))) : [];
  const stylesDoc = stylesRel ? parseXML(await zip.file(stylesRel.target).async('string')) : null;
  const sst = sstRel ? kids(parseXML(await zip.file(sstRel.target).async('string')).documentElement, 'si').map(richText) : [];
  const printAreas = {};
  desc(wbDoc, 'definedName').forEach(d => {
    if (at(d, 'name') !== '_xlnm.Print_Area' || at(d, 'localSheetId') == null) return;
    const sheet = rawSheets[+at(d, 'localSheetId')]; if (!sheet) return;
    const first = d.textContent.split(',')[0]; const rng = parseRange(first.split('!').pop().replace(/\$/g, ''));
    if (rng) printAreas[sheet.name] = rng;
  });
  const book = {
    zip, wbPath, wbDoc, sheets, sst, theme, stylesDoc, stylesPath: stylesRel && stylesRel.target,
    styles: parseStyles(stylesDoc, theme), printAreas, cache: new Map(), dirty: new Set(),
  };
  return book;
}
async function loadSheet(book, name) {
  if (book.cache.has(name)) return book.cache.get(name);
  const meta = book.sheets.find(s => s.name === name);
  if (!meta) throw Error('존재하지 않는 시트: ' + name);
  const doc = parseXML(await book.zip.file(meta.path).async('string'));
  const entry = { meta, doc, model: sheetModel(book, doc) };
  book.cache.set(name, entry);
  return entry;
}
/* ---------- formula text helpers ---------- */
// Split a formula into string literals / quoted sheet names (left untouched) and code segments.
function formulaParts(f) { return f.match(/"(?:[^"]|"")*"|'(?:[^']|'')*'|[^"']+/g) || []; }
function shiftFormula(f, dr, dc) {
  if (!dr && !dc) return f;
  const re = /(^|[^A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])|(^|[^A-Za-z0-9_.$])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![\d(A-Za-z_])/g;
  return formulaParts(f).map(p => /^["']/.test(p) ? p : p.replace(re, (m, pre, ca, col, ra, row, pre2, ca1, col1, ca2, col2) => {
    if (col !== undefined) {
      const c = ca ? colIndex(col) : colIndex(col) + dc, r = ra ? +row : +row + dr;
      return c < 1 || r < 1 ? m : pre + ca + colLetters(c) + ra + r;
    }
    const a = ca1 ? colIndex(col1) : colIndex(col1) + dc, b = ca2 ? colIndex(col2) : colIndex(col2) + dc;
    return pre2 + ca1 + colLetters(a) + ':' + ca2 + colLetters(b);
  })).join('');
}
// Split "NAME(args)" calls so individual arguments can be rewritten.
function rewriteCalls(f, names, fn) {
  let out = '', i = 0; const re = new RegExp('\\b(' + names.join('|') + ')\\(', 'gi'); let m;
  while ((m = re.exec(f))) {
    let depth = 1, j = m.index + m[0].length, q = null, args = [], start = j;
    for (; j < f.length && depth; j++) { const ch = f[j];
      if (q) { if (ch === q) q = null; continue; }
      if (ch === '"' || ch === "'") q = ch; else if (ch === '(') depth++; else if (ch === ')') { depth--; if (!depth) break; }
      else if (ch === ',' && depth === 1) { args.push(f.slice(start, j)); start = j + 1; } }
    if (depth) break;
    args.push(f.slice(start, j));
    const inner = args.map(a => rewriteCalls(a, names, fn));
    out += f.slice(i, m.index) + fn(m[1].toUpperCase(), inner); i = j + 1; re.lastIndex = i;
  }
  return out + f.slice(i);
}
const SINGLE_REF = /^\s*(?:'(?:[^']|'')+'!|[^\s,()!'"]+!)?\$?[A-Z]{1,3}\$?\d+\s*$/;
const RANGE_REF = /^\s*((?:'(?:[^']|'')+'!|[^\s,()!'"]+!)?)(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(\$?)(\d+)\s*$/;
// Excel treats a blank criterion cell as 0 and silently resizes SUMIF's sum_range; HyperFormula errors on both.
function excelCompat(f) {
  return rewriteCalls(f, ['COUNTIF', 'SUMIF', 'AVERAGEIF'], (name, a) => {
    if (a[1] && SINGLE_REF.test(a[1])) a[1] = `IF(ISBLANK(${a[1].trim()}),0,${a[1].trim()})`;
    if ((name === 'SUMIF' || name === 'AVERAGEIF') && a[2]) {
      const c = RANGE_REF.exec(a[0]), sr = RANGE_REF.exec(a[2]);
      if (c && sr) { const h = +c[9] - +c[5], w = colIndex(c[7]) - colIndex(c[3]);
        a[2] = `${sr[1]}${sr[2]}${sr[3]}${sr[4]}${sr[5]}:${sr[6]}${colLetters(colIndex(sr[3]) + w)}${sr[8]}${+sr[5] + h}`; }
    }
    return `${name}(${a.join(',')})`;
  });
}
// HyperFormula needs quoted sheet names when they are not plain ASCII identifiers, and no _xlfn. prefixes.
function hfFormula(f) {
  const quoted = formulaParts(f).map(p => /^["']/.test(p) ? p : p.replace(/_xlfn\.|_xlws\./g, '')
    .replace(/(^|[,(=+\-*/&^<>:; ])([^\s,()=+\-*/&^<>:;!'"$\d][^\s,()=+\-*/&^<>:;!'"]*)!/g, (m, pre, name) => `${pre}'${name}'!`)).join('');
  return '=' + excelCompat(quoted);
}

function visibleStyle(book, s) {
  if (!s) return false; const xf = book.styles.xfs[s] || {};
  return !!((book.styles.fills[xf.fillId] || {}).bg || Object.keys(book.styles.borders[xf.borderId] || {}).length);
}
function sheetModel(book, doc) {
  const root = doc.documentElement, fmt = kid(root, 'sheetFormatPr');
  const defRowHt = +at(fmt, 'defaultRowHeight') || 15;
  const defColW = at(fmt, 'defaultColWidth') != null ? +at(fmt, 'defaultColWidth') : (+(at(fmt, 'baseColWidth') || 8)) + 0.71;
  const cols = kids(kid(root, 'cols'), 'col').map(c => ({ min: +at(c, 'min'), max: +at(c, 'max'), width: at(c, 'width') != null ? +at(c, 'width') : null, hidden: truthy(at(c, 'hidden')), s: +at(c, 'style') || 0 }));
  const rows = new Map(), cells = new Map(), shared = new Map(); let maxR = 0, maxC = 0, fullR = 0, fullC = 0, ri = 0;
  for (const row of kids(kid(root, 'sheetData'), 'row')) {
    ri = at(row, 'r') ? +at(row, 'r') : ri + 1;
    rows.set(ri, { ht: at(row, 'ht') != null ? +at(row, 'ht') : null, hidden: truthy(at(row, 'hidden')), s: truthy(at(row, 'customFormat')) ? +at(row, 's') || 0 : null });
    let ci = 0;
    for (const c of kids(row, 'c')) {
      const p = at(c, 'r') ? parseAddr(at(c, 'r')) : null; ci = p ? p.c : ci + 1;
      const t = at(c, 't') || 'n', s = +at(c, 's') || 0, v = kid(c, 'v'), raw = v ? v.textContent : null;
      let value = null;
      if (t === 's') value = raw != null ? (book.sst[+raw] ?? '') : null;
      else if (t === 'inlineStr') value = richText(kid(c, 'is'));
      else if (t === 'b') value = raw == null ? null : raw === '1';
      else if (t === 'n') value = raw == null || raw === '' ? null : +raw;
      else value = raw;
      const f = kid(c, 'f'); let ftext = null;
      if (f) {
        ftext = f.textContent || null;
        if (at(f, 't') === 'shared') { const si = at(f, 'si');
          if (ftext) shared.set(si, { text: ftext, r: ri, c: ci }); else { const m = shared.get(si); ftext = m ? shiftFormula(m.text, ri - m.r, ci - m.c) : null; } }
      }
      cells.set(addrOf(ri, ci), { r: ri, c: ci, t, s, value, formula: !!f, ftext });
      fullR = Math.max(fullR, ri); fullC = Math.max(fullC, ci);
      if ((value != null && value !== '') || ftext || visibleStyle(book, s)) { maxR = Math.max(maxR, ri); maxC = Math.max(maxC, ci); }
    }
  }
  const merges = kids(kid(root, 'mergeCells'), 'mergeCell').map(m => parseRange(at(m, 'ref'))).filter(Boolean);
  merges.forEach(m => { const c = cells.get(addrOf(m.r1, m.c1));
    if (c && ((c.value != null && c.value !== '') || c.ftext || visibleStyle(book, c.s))) { maxR = Math.max(maxR, m.r2); maxC = Math.max(maxC, m.c2); } });
  const pm = kid(root, 'pageMargins'), ps = kid(root, 'pageSetup'), po = kid(root, 'printOptions');
  const fitToPage = truthy(at(kid(kid(root, 'sheetPr'), 'pageSetUpPr'), 'fitToPage'));
  const validations = [];
  desc(root, 'dataValidation').forEach(dv => {
    if (at(dv, 'type') !== 'list') return;
    const f1 = kid(dv, 'formula1'); if (!f1) return;
    const formula = (kid(f1, 'f') || f1).textContent.trim();
    const sq = at(dv, 'sqref') || (kid(dv, 'sqref') || {}).textContent || '';
    const ranges = sq.trim().split(/\s+/).map(parseRange).filter(Boolean);
    if (formula && ranges.length) validations.push({ formula, ranges });
  });
  const inch = (n, d) => (pm && at(pm, n) != null ? +at(pm, n) : d) * 25.4;
  return {
    defRowHt, defColW, cols, rows, cells, merges, maxR, maxC, fullR: Math.max(fullR, maxR), fullC: Math.max(fullC, maxC), validations,
    scale: at(ps, 'scale') ? +at(ps, 'scale') / 100 : 1,
    fit: fitToPage ? { w: at(ps, 'fitToWidth') != null ? +at(ps, 'fitToWidth') : 1, h: at(ps, 'fitToHeight') != null ? +at(ps, 'fitToHeight') : 1 } : null,
    margins: { l: inch('left', .7), r: inch('right', .7), t: inch('top', .75), b: inch('bottom', .75) },
    landscape: at(ps, 'orientation') === 'landscape', centerH: truthy(at(po, 'horizontalCentered')),
  };
}
function colDef(model, c) { return model.cols.find(d => c >= d.min && c <= d.max); }
function colPx(model, c) {
  const d = colDef(model, c); if (d && d.hidden) return 0;
  const w = d && d.width != null ? d.width : model.defColW;
  return Math.max(0, Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7));
}
function rowPx(model, r) { const d = model.rows.get(r); if (d && d.hidden) return 0; return Math.round(((d && d.ht != null) ? d.ht : model.defRowHt) * 4 / 3); }
function mergeAt(model, addr) { const p = parseAddr(addr); return p && model.merges.find(m => p.r >= m.r1 && p.r <= m.r2 && p.c >= m.c1 && p.c <= m.c2); }

/* ---------- HTML replica of a sheet ---------- */
const BORDER_CSS = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dotted',
  medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double' };
const CJK_FONT = /gothic|고딕|굴림|돋움|바탕|명조|맑은|malgun|gulim|dotum|batang|nanum|나눔|noto|cjk|hy|한컴|함초롬|궁서/i;
function fontStack(name) {
  const safe = String(name || '').replace(/['";{}<>\\]/g, '');
  return (safe && !CJK_FONT.test(safe) ? `'${safe}', ` : '') + `'Noto Sans KR', 'Malgun Gothic', sans-serif`;
}
function styleOf(model, r, c) {
  const cell = model.cells.get(addrOf(r, c)); if (cell) return cell.s;
  const rowS = (model.rows.get(r) || {}).s; return rowS != null ? rowS : (colDef(model, c) || {}).s || 0;
}
function cellCss(book, model, r, c, merge, isNum) {
  const s = styleOf(model, r, c);
  const st = book.styles, xf = st.xfs[s] || st.xfs[0] || {};
  const font = st.fonts[xf.fontId] || st.fonts[0] || {}, fill = st.fills[xf.fillId] || {};
  const css = [`font-size:${((font.sz || 11) * 4 / 3).toFixed(2)}px`, `font-family:${fontStack(font.name)}`];
  if (font.b) css.push('font-weight:700'); if (font.i) css.push('font-style:italic');
  const deco = [font.u && 'underline', font.strike && 'line-through'].filter(Boolean); if (deco.length) css.push('text-decoration:' + deco.join(' '));
  if (font.color) css.push('color:' + font.color);
  if (fill.bg) css.push('background:' + fill.bg);
  // borders: for merges use the outer edge cells
  const edge = (side) => {
    let rr = r, cc = c;
    if (merge) { if (side === 'right') cc = merge.c2; if (side === 'bottom') rr = merge.r2; }
    const ec = model.cells.get(addrOf(rr, cc)); const es = ec ? ec.s : s;
    const ex = st.xfs[es] || {}; return (st.borders[ex.borderId] || {})[side];
  };
  for (const side of ['left', 'right', 'top', 'bottom']) { const b = edge(side); if (b) css.push(`border-${side}:${BORDER_CSS[b.style] || '1px solid'} ${b.color}`); }
  const h = xf.h === 'centerContinuous' ? 'center' : xf.h === 'fill' || xf.h === 'distributed' ? 'justify' : xf.h;
  css.push('text-align:' + (h && h !== 'general' ? h : isNum ? 'right' : 'left'));
  css.push('vertical-align:' + (xf.v === 'center' ? 'middle' : xf.v === 'top' ? 'top' : 'bottom'));
  if (xf.indent) css.push(`padding-left:${3 + xf.indent * 9}px`);
  return { css: css.join(';'), wrap: !!xf.wrap };
}
/** opts: {range, grid:bool, overrides:Map, mapped:Map(addr->n), active, escape} */
function sheetHTML(book, model, opts) {
  const { r1, c1, r2, c2 } = opts.range, esc = opts.escape;
  const tops = new Map(), covered = new Set();
  for (const m of model.merges) {
    if (m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2) continue;
    tops.set(addrOf(m.r1, m.c1), m);
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(addrOf(r, c));
  }
  const widths = []; for (let c = c1; c <= c2; c++) widths.push(colPx(model, c));
  const total = widths.reduce((a, b) => a + b, 0) + (opts.grid ? 38 : 0);
  let html = `<table class="xl ${opts.grid ? 'grid' : 'print'}" style="width:${total}px"><colgroup>${opts.grid ? '<col style="width:38px">' : ''}${widths.map(w => `<col style="width:${w}px">`).join('')}</colgroup>`;
  if (opts.grid) html += `<thead><tr><th class="xl-corner"></th>${widths.map((w, i) => `<th>${w ? colLetters(c1 + i) : ''}</th>`).join('')}</tr></thead>`;
  html += '<tbody>';
  for (let r = r1; r <= r2; r++) {
    const h = rowPx(model, r);
    html += `<tr style="height:${h}px"${h ? '' : ' class="xl-hidden"'}>` + (opts.grid ? `<th class="xl-rh">${r}</th>` : '');
    for (let c = c1; c <= c2; c++) {
      const a = addrOf(r, c); if (covered.has(a)) continue;
      const m = tops.get(a), cell = model.cells.get(a);
      const val = opts.values && opts.values.has(a) ? opts.values.get(a) : cell ? cell.value : null;
      const text = formatValue(book, { s: styleOf(model, r, c), value: val });
      const isNum = typeof val === 'number';
      const hiddenCol = !colPx(model, c) && !m;
      let { css, wrap } = cellCss(book, model, r, c, m, isNum);
      if (hiddenCol) css += ';padding:0;border:none';
      const override = opts.highlight && opts.highlight.has(a);
      let spill = 0;
      const contCenter = ((book.styles.xfs[styleOf(model, r, c)] || {}).h) === 'centerContinuous';
      if (!m && !wrap && text && (contCenter || (!isNum && !/text-align:(center|right|justify)/.test(css)))) {
        let w = 0;
        for (let cc = c + 1; cc <= c2; cc++) {
          const na = addrOf(r, cc); if (covered.has(na) || tops.has(na)) break;
          const nc = model.cells.get(na), nv = opts.values && opts.values.has(na) ? opts.values.get(na) : nc ? nc.value : null;
          if (nv != null && nv !== '') break; w += colPx(model, cc);
        }
        if (w) spill = colPx(model, c) + w - 6;
        if (contCenter && !w) spill = 0;
      }
      const span = m ? `${m.r2 > m.r1 ? ` rowspan="${Math.min(m.r2, r2) - r + 1}"` : ''}${m.c2 > m.c1 ? ` colspan="${Math.min(m.c2, c2) - c + 1}"` : ''}` : '';
      let height = 0; if (m) for (let rr = r; rr <= Math.min(m.r2, r2); rr++) height += rowPx(model, rr); else height = h;
      const mi = opts.mapped && opts.mapped.get(a);
      const cls = [mi != null && 'mapped', opts.active === a && 'active', override && text !== '' && 'filled'].filter(Boolean).join(' ');
      const tdCls = [cls, spill && 'ov'].filter(Boolean).join(' ');
      html += `<td${span}${tdCls ? ` class="${tdCls}"` : ''}${opts.grid ? ` data-act="cell" data-cell="${a}" title="${esc(a)}"` : ''} style="${css}"><div class="xc${wrap ? ' wrap' : ''}${spill ? ' spill' : ''}" style="max-height:${height}px${spill ? `;width:${spill}px` : ''}">${hiddenCol ? '' : esc(text)}</div>${mi != null ? `<span class="cell-dot">${mi}</span>` : ''}</td>`;
    }
    html += '</tr>';
  }
  return html + '</tbody></table>';
}
function printRange(book, name, model) {
  return book.printAreas[name] || { r1: 1, c1: 1, r2: Math.max(1, model.maxR), c2: Math.max(1, model.maxC) };
}

/* ---------- writing values ---------- */
function ensureDateStyle(book, s) {
  book.dateXf = book.dateXf || new Map();
  if (book.dateXf.has(s)) return book.dateXf.get(s);
  const { id, code } = fmtCode(book, s);
  if (isDateFmt(id, code)) { book.dateXf.set(s, s); return s; }
  const doc = book.stylesDoc; const root = doc && doc.documentElement; const cellXfs = kid(root, 'cellXfs');
  if (!cellXfs) { book.dateXf.set(s, null); return null; }
  if (!book.dateFmtId) {
    let nf = kid(root, 'numFmts');
    if (!nf) { nf = mk(doc, root, 'numFmts'); root.insertBefore(nf, root.firstElementChild); }
    const ids = kids(nf, 'numFmt').map(n => +at(n, 'numFmtId'));
    book.dateFmtId = Math.max(163, ...ids) + 1;
    const n = mk(doc, root, 'numFmt'); n.setAttribute('numFmtId', book.dateFmtId); n.setAttribute('formatCode', 'yyyy-mm-dd');
    nf.appendChild(n); nf.setAttribute('count', kids(nf, 'numFmt').length);
  }
  const list = kids(cellXfs, 'xf'), base = list[s] || list[0];
  const clone = base.cloneNode(true); clone.setAttribute('numFmtId', book.dateFmtId); clone.setAttribute('applyNumberFormat', '1');
  cellXfs.appendChild(clone); cellXfs.setAttribute('count', list.length + 1);
  book.stylesDirty = true; book.dateXf.set(s, list.length);
  return list.length;
}
/** val: {kind:'s', text} | {kind:'n', num} | {kind:'d', num} */
function writeCell(book, entry, addr, val) {
  const doc = entry.doc, root = doc.documentElement; let sd = kid(root, 'sheetData');
  if (!sd) { sd = mk(doc, root, 'sheetData'); root.appendChild(sd); }
  const { r, c } = parseAddr(addr);
  if (val.kind === 'clear') { // blank optional field: empty the cell but keep its style
    const ce = desc(sd, 'c').find(x => at(x, 'r') === addrOf(r, c));
    if (ce) { while (ce.firstChild) ce.removeChild(ce.firstChild); ['t', 'cm', 'vm'].forEach(a => ce.removeAttribute(a)); book.dirty.add(entry.meta.name); }
    return;
  }
  let row = null, before = null, idx = 0;
  for (const rw of kids(sd, 'row')) {
    idx = at(rw, 'r') ? +at(rw, 'r') : idx + 1; if (!at(rw, 'r')) rw.setAttribute('r', idx);
    if (idx === r) { row = rw; break; } if (idx > r) { before = rw; break; }
  }
  if (!row) { row = mk(doc, sd, 'row'); row.setAttribute('r', r); sd.insertBefore(row, before); }
  let cell = null; before = null; idx = 0;
  for (const ce of kids(row, 'c')) {
    const p = at(ce, 'r') ? parseAddr(at(ce, 'r')) : null; idx = p ? p.c : idx + 1;
    if (!p) ce.setAttribute('r', addrOf(r, idx));
    if (idx === c) { cell = ce; break; } if (idx > c) { before = ce; break; }
  }
  if (!cell) {
    cell = mk(doc, row, 'c'); cell.setAttribute('r', addrOf(r, c));
    const inherit = truthy(at(row, 'customFormat')) ? at(row, 's') : (colDef(entry.model, c) || {}).s;
    if (inherit && +inherit) cell.setAttribute('s', inherit);
    row.insertBefore(cell, before);
  }
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  ['t', 'cm', 'vm'].forEach(a => cell.removeAttribute(a));
  row.removeAttribute('spans');
  if (val.kind === 'n' || val.kind === 'd') {
    if (val.kind === 'd') { const ns = ensureDateStyle(book, +at(cell, 's') || 0);
      if (ns == null) return writeCell(book, entry, addr, { kind: 's', text: val.text }); cell.setAttribute('s', ns); }
    const v = mk(doc, cell, 'v'); v.textContent = String(val.num); cell.appendChild(v);
  } else {
    cell.setAttribute('t', 'inlineStr');
    const is = mk(doc, cell, 'is'), t = mk(doc, cell, 't');
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve'); t.textContent = val.text;
    is.appendChild(t); cell.appendChild(is);
  }
  book.dirty.add(entry.meta.name);
}
function forceRecalc(book) {
  const doc = book.wbDoc, root = doc.documentElement; let calc = kid(root, 'calcPr');
  if (!calc) {
    calc = mk(doc, root, 'calcPr');
    const after = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
    root.insertBefore(calc, [...root.children].find(ch => after.includes(ch.localName)) || null);
  }
  calc.setAttribute('fullCalcOnLoad', '1');
}
async function saveBook(book) {
  for (const name of book.dirty) { const e = book.cache.get(name); book.zip.file(e.meta.path, serialize(e.doc)); }
  if (book.stylesDirty && book.stylesPath) book.zip.file(book.stylesPath, serialize(book.stylesDoc));
  forceRecalc(book); book.zip.file(book.wbPath, serialize(book.wbDoc));
  return book.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ---------- formula recalculation (HyperFormula) ---------- */
const isErr = v => v && typeof v === 'object' && 'value' in v && typeof v.value === 'string' && v.value[0] === '#';
function hfInput(cell) {
  if (cell.ftext) return hfFormula(cell.ftext);
  const v = cell.value;
  if (v == null) return null;
  if (typeof v === 'string') return v === '' ? null : "'" + v;
  return v;
}
async function buildCalc(book) {
  if (!window.HyperFormula) throw Error('수식 엔진(HyperFormula)을 불러오지 못했습니다. 네트워크를 확인하세요.');
  const entries = []; for (const s of book.sheets) entries.push(await loadSheet(book, s.name));
  const data = {};
  for (const e of entries) {
    const m = e.model, rows = Array.from({ length: Math.max(1, m.fullR) }, () => new Array(Math.max(1, m.fullC)).fill(null));
    for (const cell of m.cells.values()) rows[cell.r - 1][cell.c - 1] = hfInput(cell);
    data[e.meta.name] = rows;
  }
  const names = desc(book.wbDoc, 'definedName').filter(d => !/^_xlnm\./.test(at(d, 'name')) && at(d, 'localSheetId') == null && !/#REF!/.test(d.textContent))
    .map(d => ({ name: at(d, 'name'), expression: hfFormula(d.textContent) }));
  let hf;
  try { hf = HyperFormula.buildFromSheets(data, { licenseKey: 'gpl-v3', precisionRounding: 10 }, names); }
  catch { hf = HyperFormula.buildFromSheets(data, { licenseKey: 'gpl-v3', precisionRounding: 10 }); }
  const ids = Object.fromEntries(entries.map(e => [e.meta.name, hf.getSheetId(e.meta.name)]));
  const calc = {
    hf,
    /** inputs: [{sheet, addr, value}] where value is a number or string */
    set(inputs) {
      hf.batch(() => { for (const x of inputs) { const p = parseAddr(x.addr); if (p && ids[x.sheet] != null)
        hf.setCellContents({ sheet: ids[x.sheet], row: p.r - 1, col: p.c - 1 }, [[x.value == null || x.value === '' ? null : typeof x.value === 'string' ? "'" + x.value : x.value]]); } });
    },
    /** computed value for one cell; falls back to the file's cached value where the engine lacks a function */
    value(sheet, cell) {
      const v = hf.getCellValue({ sheet: ids[sheet], row: cell.r - 1, col: cell.c - 1 });
      if (isErr(v)) return /NAME|ERROR/.test(v.type) && cell.formula ? cell.value : v.value;
      if (v == null && cell.formula) return 0;
      return v;
    },
    values(sheet) {
      const e = book.cache.get(sheet), out = new Map(); if (!e) return out;
      for (const cell of e.model.cells.values()) if (cell.formula) out.set(addrOf(cell.r, cell.c), this.value(sheet, cell));
      return out;
    },
    /** write computed results into each formula cell's cached <v>, so every viewer shows the new values */
    writeCache() {
      for (const e of entries) {
        for (const c of desc(e.doc.documentElement, 'c')) {
          const f = kid(c, 'f'); if (!f) continue;
          const p = parseAddr(at(c, 'r')); if (!p) continue;
          const cell = e.model.cells.get(addrOf(p.r, p.c)); if (!cell) continue;
          const raw = hf.getCellValue({ sheet: ids[e.meta.name], row: p.r - 1, col: p.c - 1 });
          if (isErr(raw) && /NAME|ERROR/.test(raw.type)) continue; // unsupported here: keep Excel's cached value
          const v = isErr(raw) ? raw.value : raw == null ? 0 : raw;
          kids(c, 'v').forEach(x => c.removeChild(x)); kids(c, 'is').forEach(x => c.removeChild(x));
          const ve = mk(e.doc, c, 'v');
          if (isErr(raw)) { c.setAttribute('t', 'e'); ve.textContent = v; }
          else if (typeof v === 'number') { c.removeAttribute('t'); ve.textContent = String(v); }
          else if (typeof v === 'boolean') { c.setAttribute('t', 'b'); ve.textContent = v ? '1' : '0'; }
          else { c.setAttribute('t', 'str'); ve.textContent = String(v); }
          f.after(ve);
          book.dirty.add(e.meta.name);
        }
      }
    },
  };
  return calc;
}

/* ---------- data-validation dropdown lists ---------- */
async function listOptions(book, sheetName, addr) {
  const { model } = await loadSheet(book, sheetName); const p = parseAddr(addr); if (!p) return null;
  const dv = model.validations.find(v => v.ranges.some(g => p.r >= g.r1 && p.r <= g.r2 && p.c >= g.c1 && p.c <= g.c2));
  if (!dv) return null;
  let f = dv.formula.replace(/^=/, '');
  if (/^".*"$/.test(f)) return [...new Set(f.slice(1, -1).split(',').map(x => x.trim()).filter(Boolean))];
  const named = desc(book.wbDoc, 'definedName').find(d => at(d, 'name') === f); if (named) f = named.textContent;
  return rangeValues(book, f, sheetName);
}
async function rangeValues(book, f, sheetName) {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(f);
  const target = m ? (m[1] || m[2]).replace(/''/g, "'") : sheetName, rng = parseRange((m ? m[3] : f).replace(/\$/g, ''));
  if (!rng || !book.sheets.some(s => s.name === target)) return null;
  const { model: tm } = await loadSheet(book, target); const out = [];
  for (let r = rng.r1; r <= Math.min(rng.r2, rng.r1 + 500); r++) for (let c = rng.c1; c <= rng.c2; c++) {
    const cell = tm.cells.get(addrOf(r, c)); const t = cell ? formatValue(book, cell).trim() : ''; if (t) out.push(t);
  }
  return [...new Set(out)];
}

/* ---------- PDF ---------- */
async function renderPdf(book, sheetNames, escape) {
  const { jsPDF } = window.jspdf;
  const host = document.createElement('div'); host.className = 'print-host'; document.body.appendChild(host);
  const PX_MM = 25.4 / 96, RES = 2; let pdf = null; const pages = [];
  try {
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    for (const name of sheetNames) {
      const { model } = await loadSheet(book, name);
      if (!model.maxR && !book.printAreas[name]) continue;
      host.innerHTML = sheetHTML(book, model, { range: printRange(book, name, model), grid: false, escape });
      const table = host.querySelector('table');
      const [pw, ph] = model.landscape ? [297, 210] : [210, 297], mg = model.margins;
      const availW = pw - mg.l - mg.r, availH = ph - mg.t - mg.b;
      const tw = table.offsetWidth, th = table.offsetHeight;
      let scale = Math.min(model.scale || 1, availW / (tw * PX_MM));
      if (model.fit) { const w = availW / (tw * PX_MM), h = availH / (th * PX_MM);
        scale = model.fit.h === 1 && model.fit.w <= 1 ? Math.min(w, h) : model.fit.w === 0 ? Math.min(1, h) : w; }
      const canvas = await html2canvas(table, { scale: RES, backgroundColor: '#ffffff', logging: false, useCORS: false });
      const k = canvas.width / tw;
      const bounds = [...table.rows].map(tr => tr.offsetTop + tr.offsetHeight);
      const pageHpx = availH / (PX_MM * scale);
      let start = 0;
      while (start < th - 0.5) {
        const limit = start + pageHpx;
        let end = Math.max(...bounds.filter(b => b > start + 1 && b <= limit + 0.5), 0);
        if (!end) end = Math.min(limit, th);
        if (th - end < 8) end = th; // trailing border pixels belong to the last page
        const pxPerMm = RES / PX_MM;
        const page = document.createElement('canvas');
        page.width = Math.round(pw * pxPerMm); page.height = Math.round(ph * pxPerMm);
        const g = page.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, page.width, page.height);
        const drawW = tw * PX_MM * scale, x = model.centerH ? (pw - drawW) / 2 : mg.l;
        g.drawImage(canvas, 0, Math.round(start * k), canvas.width, Math.round((end - start) * k),
          x * pxPerMm, mg.t * pxPerMm, drawW * pxPerMm, (end - start) * PX_MM * scale * pxPerMm);
        const img = page.toDataURL('image/jpeg', 0.94);
        const orient = model.landscape ? 'l' : 'p';
        if (!pdf) pdf = new jsPDF({ orientation: orient, unit: 'mm', format: 'a4', compress: true });
        else pdf.addPage('a4', orient);
        pdf.addImage(img, 'JPEG', 0, 0, pw, ph);
        pages.push({ img, landscape: model.landscape, sheet: name });
        start = end;
      }
    }
  } finally { host.remove(); }
  if (!pdf) throw Error('PDF로 출력할 내용이 없습니다.');
  return { bytes: new Uint8Array(pdf.output('arraybuffer')), pages };
}

window.XL = { buildCalc, listOptions, rangeValues, colLetters, parseAddr, normAddr, addrOf, openBook, loadSheet, sheetHTML, printRange, mergeAt, writeCell, saveBook, renderPdf, formatValue };

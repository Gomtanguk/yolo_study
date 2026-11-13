/* Formly UI — ported from web/app.js; the FastAPI calls now run against the in-browser engine + artifact storage. */
'use strict';
const MAX_UPLOAD = 10 * 1024 * 1024;
const CHUNK = 180000; // base64 chars per stored document (db doc limit: 256 KiB)
const TYPES = ['text', 'textarea', 'number', 'date', 'select', 'checkbox'];

/* ---------- utils ---------- */
const esc = x => String(x ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const nowIso = () => new Date().toISOString();
const clone = o => JSON.parse(JSON.stringify(o));
function toB64(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); }
function fromB64(b64) { const s = atob(b64), out = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i); return out; }
class UserError extends Error {}
const fail = msg => { throw new UserError(msg); };

/* ---------- storage: artifact db (per-person, private) with in-memory fallback ---------- */
const store = {
  mode: 'connecting', db: null, uid: null, mem: new Map(), files: new Map(),
  async init() {
    try {
      if (window.claude && typeof window.claude.use === 'function') {
        const [db, user] = await Promise.all([window.claude.use('db'), window.claude.use('user')]);
        const id = user ? await user.id() : null;
        if (db && id) { this.db = db; this.uid = id; this.mode = 'db'; return; }
      }
    } catch (e) { console.warn('db unavailable', e); }
    this.mode = 'memory';
  },
  profile() { return this.db.doc(`data/users/${this.uid}/profile`); },
  tcol() { return this.profile().collection('templates'); },
  fcol() { return this.profile().collection('files'); },
  async getProfile() { if (this.mode !== 'db') return this.memProfile || {}; const s = await this.profile().get(); return s.exists ? s.data() : {}; },
  async setProfile(p) { if (this.mode !== 'db') { this.memProfile = p; return; } await this.profile().set(p); },
  async list() {
    if (this.mode !== 'db') return [...this.mem.values()].map(clone);
    const snap = await this.tcol().get();
    return snap.docs.filter(d => d.exists).map(d => ({ ...clone(d.data()), id: d.id }));
  },
  async put(t) { const body = clone(t); delete body.id; if (this.mode !== 'db') { this.mem.set(t.id, clone(t)); return; } await this.tcol().doc(t.id).set(body); },
  async putFile(id, bytes) {
    this.files.set(id, bytes);
    if (this.mode !== 'db') return 1;
    const b64 = toB64(bytes), n = Math.max(1, Math.ceil(b64.length / CHUNK));
    for (let i = 0; i < n; i++) await this.fcol().doc(`${id}~${i}`).set({ b64: b64.slice(i * CHUNK, (i + 1) * CHUNK), i, n });
    return n;
  },
  async getFile(t) {
    if (this.files.has(t.id)) return this.files.get(t.id);
    if (this.mode !== 'db') fail('원본 파일을 찾을 수 없습니다.');
    const parts = await Promise.all(Array.from({ length: t.chunks || 1 }, (_, i) => this.fcol().doc(`${t.id}~${i}`).get()));
    if (parts.some(p => !p.exists)) fail('저장된 원본 파일이 손상됐습니다. 양식을 다시 등록하세요.');
    const bytes = fromB64(parts.map(p => p.data().b64).join(''));
    this.files.set(t.id, bytes); return bytes;
  },
  async remove(t) {
    this.files.delete(t.id); this.mem.delete(t.id);
    if (this.mode !== 'db') return;
    await this.tcol().doc(t.id).delete();
    for (let i = 0; i < (t.chunks || 1); i++) await this.fcol().doc(`${t.id}~${i}`).delete();
  },
};
const books = new Map(); // template id -> parsed (read-only) workbook
async function bookFor(t) { if (!books.has(t.id)) books.set(t.id, await XL.openBook(await store.getFile(t))); return books.get(t.id); }

/* ---------- capabilities ---------- */
let downloads = null;
async function initDownloads() { try { if (window.claude && window.claude.use) downloads = await window.claude.use('downloads'); } catch { downloads = null; } render(); }
async function saveFile(filename, data) {
  if (!downloads) fail('이 화면에서는 파일 저장을 사용할 수 없습니다. claude.ai에서 아티팩트를 열어주세요.');
  try { await downloads.save({ filename, data: new Blob([data]) }); state.success = `'${filename}' 저장을 요청했습니다.`; }
  catch (e) {
    if (e && e.code === 'declined') return;
    if (e && e.code === 'rate_limited') fail('저장 창이 이미 열려 있습니다. 잠시 후 다시 시도하세요.');
    fail('파일을 저장할 수 없습니다. (' + (e && (e.code || e.message) || e) + ')');
  }
}

/* ---------- state ---------- */
const state = {
  page: 'loading', templates: [], active: null, draft: null, book: null, sheet: '', selected: 0,
  result: null, values: {}, uploadOpen: false, uploadName: '', uploadFile: null,
  error: '', success: '', busy: false, busyText: '', dirty: false, search: '', confirm: null,
};
const icons = { home: '▦', settings: '⚙', file: '▤', pin: '⌖', info: 'ⓘ' };

/* ---------- domain logic (ported from backend/app/main.py) ---------- */
function summary(t) { return { ...t, field_count: t.fields.length, attachment_count: t.attachments.length }; }
async function refresh() { state.templates = (await store.list()).sort((a, b) => (b.updated_at || '').localeCompare(a.updated_at || '')); }

async function createTemplate(name, filename, bytes, fields = [], attachments = []) {
  const book = await XL.openBook(bytes);
  const id = uid();
  const chunks = await store.putFile(id, bytes);
  const t = { id, name, filename, fields, attachments, size: bytes.length, chunks, created_at: nowIso(), updated_at: nowIso() };
  await store.put(t); books.set(id, book);
  return t;
}
const SAMPLE_VERSION = 2;
const SAMPLE_FIELDS = [
  // [key, label, type, required, cell, placeholder, optionsFrom]
  ['traveler', '출장자 성명', 'select', true, 'D6', '', "'사번'!$A$2:$A$34"],
  ['project', '과제명 (비고)', 'select', true, 'P6', '', 'dv'],
  ['trip_name', '출장명', 'text', true, 'F8', '예: OO 전시회 참관 및 미팅'],
  ['place1', '방문처 1', 'text', true, 'E11', '예: 코엑스'],
  ['region1', '출장지 1 (지역/주소)', 'text', true, 'H11', '예: 서울특별시 강남구 영동대로 513'],
  ['region2', '출장지 2 (지역/주소)', 'text', false, 'H12', '경유지가 있으면 입력'],
  ['start_date', '출장 시작일', 'date', true, 'D15'],
  ['start_time', '시작 시간', 'select', true, 'I15', '', 'dv'],
  ['end_date', '출장 종료일', 'date', true, 'L15'],
  ['end_time', '종료 시간', 'select', true, 'Q15', '', 'dv'],
  ['meals_provided', '식사 제공 횟수 (회의비/타 기관)', 'number', false, 'Q16', '0'],
  ['perdiem_region', '일비 산정 지역', 'select', true, 'W9', '', 'dv'],
  ['perdiem_end', '일비 산정 종료일', 'date', true, 'AC9', '보통 출장 종료일과 같음'],
  ['lodging_region', '숙박 지역', 'select', false, 'W19', '', "'출장지역 및 비용한도'!$I$5:$I$16"],
  ['lodging_amount', '숙박 결제금액 (원)', 'number', false, 'AA19', '60000'],
  ['lodging_nights', '숙박 일수', 'number', false, 'AM19', '1'],
  ['lodging_pay', '숙박 결제방법', 'select', false, 'AK19', '', 'dv'],
  ['lodging_note', '숙박 비고 (호텔명)', 'text', false, 'AN19', '예: OO호텔'],
  ['etc1_desc', '기타경비 1 내용', 'text', false, 'W28', '예: KTX요금 (동대구 → 서울)'],
  ['etc1_account', '기타경비 1 계정', 'select', false, 'AB28', '', 'dv'],
  ['etc1_amount', '기타경비 1 금액 (원)', 'number', false, 'AE28', '36000'],
  ['etc1_pay', '기타경비 1 결제방법', 'select', false, 'AK28', '', 'dv'],
  ['etc2_desc', '기타경비 2 내용', 'text', false, 'W29', '예: KTX요금 (서울 → 동대구)'],
  ['etc2_account', '기타경비 2 계정', 'select', false, 'AB29', '', 'dv'],
  ['etc2_amount', '기타경비 2 금액 (원)', 'number', false, 'AE29', '36000'],
  ['etc2_pay', '기타경비 2 결제방법', 'select', false, 'AK29', '', 'dv'],
  ['remarks', '특이사항', 'textarea', false, 'H42', '특이사항이 있으면 입력'],
];
async function seedSample() {
  const profile = await store.getProfile();
  if ((profile.sampleVersion || (profile.seeded ? 1 : 0)) >= SAMPLE_VERSION) return;
  // retire the old built-in sample (구매 신청서) if it is still there
  for (const t of await store.list()) if (t.sample || (t.filename === 'sample_request.xlsx' && t.name === '샘플 구매 신청서')) await store.remove(t);
  const bytes = fromB64(SAMPLE_XLSX_B64), book = await XL.openBook(bytes), sheet = '여비산출(국내)';
  const fields = [];
  for (const [key, label, type, required, cell, placeholder = '', from] of SAMPLE_FIELDS) {
    let options = [];
    if (from === 'dv') options = await XL.listOptions(book, sheet, cell) || [];
    else if (from) options = await XL.rangeValues(book, from, sheet) || [];
    fields.push({ key, label, type: type === 'select' && !options.length ? 'text' : type, required, sheet, cell, placeholder, options });
  }
  const t = await createTemplate('출장비 정산서 (국내)', '출장비_정산서_샘플.xlsx', bytes, fields, [
    { title: '교통비 영수증 (KTX·항공권 등)', note: '결제 내역이 보이는 영수증 또는 e-티켓' },
    { title: '숙박 영수증', note: '숙박한 경우 호텔 결제 영수증' },
    { title: '출장 결과 보고서', note: '회사 양식으로 작성' },
  ]);
  t.sample = SAMPLE_VERSION; t.printSheets = ['지결 양식(국내)', '여비산출(국내)'];
  await store.put(t);
  await store.setProfile({ ...profile, seeded: true, sampleVersion: SAMPLE_VERSION });
}

async function validateMapping(fields, book) {
  const keys = new Set(), cells = new Set();
  for (const f of fields) {
    if (!f.label.trim()) fail('항목명을 입력하세요.');
    if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(f.key) || f.key.length > 48) fail(`'${f.label}' 키는 영문으로 시작하고 영문·숫자·_만 사용할 수 있습니다.`);
    if (keys.has(f.key)) fail('중복 키: ' + f.key); keys.add(f.key);
    if (!book.sheets.some(s => s.name === f.sheet)) fail('존재하지 않는 시트: ' + f.sheet);
    const pos = XL.normAddr(f.cell); if (!pos) fail(`'${f.label}' 셀 주소가 올바르지 않습니다: ${f.cell || '(미지정)'}`);
    const where = f.sheet + '!' + pos; if (cells.has(where)) fail('동일한 셀에 여러 항목을 연결할 수 없습니다: ' + where); cells.add(where);
    const { model } = await XL.loadSheet(book, f.sheet);
    const m = XL.mergeAt(model, pos);
    if (m && XL.addrOf(m.r1, m.c1) !== pos) fail('병합 셀은 왼쪽 위 셀을 선택해야 합니다: ' + where);
    const cell = model.cells.get(pos); if (cell && cell.formula) fail('수식을 덮어쓰는 위치는 사용할 수 없습니다: ' + where);
    if (f.type === 'select' && !f.options.length) fail(`선택 항목 '${f.label}'에 옵션을 추가하세요.`);
    f.cell = pos;
  }
}
function convertValue(f, value) {
  if (value == null || value === '' || (f.type === 'checkbox' && value === false && !f.required)) {
    if (f.required && f.type !== 'checkbox') fail(`'${f.label}' 항목은 필수입니다.`);
    if (f.type !== 'checkbox') return { kind: 'clear' };
  }
  if (f.type === 'checkbox') { if (f.required && !value) fail(`'${f.label}' 항목을 확인해야 합니다.`); return { kind: 's', text: value ? '예' : '아니오' }; }
  if (f.type === 'date') {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value)); const d = m && new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    if (!d || d.getUTCDate() !== +m[3]) fail(`'${f.label}' 날짜가 올바르지 않습니다.`);
    return { kind: 'd', num: Math.round((d - Date.UTC(1899, 11, 30)) / 86400000), text: String(value) };
  }
  if (f.type === 'number') { const n = Number(String(value).trim()); if (!isFinite(n) || String(value).trim() === '') fail(`'${f.label}' 숫자가 올바르지 않습니다.`); return { kind: 'n', num: n }; }
  const text = String(value); if (text.length > 5000) fail(`'${f.label}' 내용이 너무 깁니다.`);
  if (f.type === 'select' && !f.options.includes(text)) fail(`'${f.label}' 값이 선택지에 없습니다.`);
  return { kind: 's', text };
}
function printSheetsOf(t, book) {
  const visible = book.sheets.filter(s => s.state === 'visible').map(s => s.name);
  const chosen = (t.printSheets || []).filter(n => visible.includes(n));
  return chosen.length ? chosen : visible;
}
const hfValue = v => v.kind === 'clear' ? null : v.kind === 's' ? v.text : v.num;
async function generate(t, values) {
  if (!t.fields.length) fail('먼저 관리자 화면에서 입력 항목을 설정하세요.');
  const converted = t.fields.map(f => [f, convertValue(f, values[f.key])]);
  state.busyText = '엑셀에 값을 기록하고 수식을 다시 계산하는 중...'; render();
  const book = await XL.openBook(await store.getFile(t)); // fresh, mutable copy
  const calc = await XL.buildCalc(book);
  for (const [f, v] of converted) XL.writeCell(book, await XL.loadSheet(book, f.sheet), XL.normAddr(f.cell), v);
  calc.set(converted.map(([f, v]) => ({ sheet: f.sheet, addr: XL.normAddr(f.cell), value: hfValue(v) })));
  calc.writeCache();
  const out = await XL.saveBook(book);
  state.busyText = 'PDF를 만드는 중...'; render();
  const filled = await XL.openBook(out);
  const pdf = await XL.renderPdf(filled, printSheetsOf(t, filled), esc);
  return { template_name: t.name, attachments: t.attachments, xlsx: out, pdf: pdf.bytes, pages: pdf.pages };
}
/** lenient conversion for the live preview (never throws) */
function previewInputs(t, values) {
  return t.fields.map(f => {
    const raw = values[f.key]; let value = null;
    if (f.type === 'checkbox') value = raw ? '예' : '아니오';
    else if (raw == null || raw === '') value = null;
    else if (f.type === 'number') { const n = Number(raw); value = isFinite(n) ? n : String(raw); }
    else if (f.type === 'date') { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw); value = m ? Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000) : String(raw); }
    else value = String(raw);
    return { sheet: f.sheet, addr: XL.normAddr(f.cell), value, filled: raw != null && raw !== '' && raw !== false };
  }).filter(x => x.addr);
}

/* ---------- views ---------- */
function layout(inner) {
  const ok = store.mode === 'db';
  const crumb = { loading: '불러오는 중', dashboard: '양식 라이브러리', editor: '양식 매핑 설정', fill: '서류 작성', result: '작성 완료' }[state.page];
  return `<div class="app-shell"><aside class="sidebar">
 <div class="brand" data-act="dashboard"><div class="brand-icon">F</div><div><strong>formly<span>.</span></strong><small>DOCUMENT WORKSPACE</small></div></div>
 <div class="nav-group-label">WORKSPACE</div>
 <button class="side-link ${state.page === 'dashboard' ? 'active' : ''}" data-act="dashboard">${icons.home}　양식 라이브러리</button>
 <button class="side-link ${state.page === 'editor' ? 'active' : ''}" data-act="editor-current">${icons.settings}　양식 매핑 관리</button>
 <div class="nav-group-label second">QUICK START</div>
 <div class="workflow-step"><span>01</span><div><b>양식 등록</b><small>Excel 파일 업로드</small></div></div>
 <div class="workflow-step"><span>02</span><div><b>필드 매핑</b><small>입력 항목 ↔ 셀 주소</small></div></div>
 <div class="workflow-step"><span>03</span><div><b>PDF 생성</b><small>입력 후 다운로드</small></div></div>
 <div class="side-bottom"><div class="secure-note"><span>●</span><div><strong>BROWSER WORKSPACE</strong><span>파일은 브라우저에서 처리되고 ${ok ? '내 계정 전용 저장소에만 보관돼요' : '이번 방문 동안만 유지돼요'}</span></div></div><small>v0.1.0 · No AI, no server</small></div>
 </aside><main class="main-panel"><header class="topbar"><div class="breadcrumb"><span>Workspace</span><span class="crumb-slash">/</span><b>${esc(crumb)}</b></div><div class="top-actions"><span class="status-dot ${ok ? 'green' : 'yellow'}"></span><span class="server-status">${store.mode === 'connecting' ? '저장소 연결 중' : ok ? '계정 저장소 연결됨 · PDF 엔진 준비됨' : '임시 모드 · 새로고침 시 사라짐'}</span><div class="avatar">MY</div></div></header>
 <div class="content">${state.error ? `<div class="alert error">ⓘ <span>${esc(state.error)}</span><button data-act="dismiss">×</button></div>` : ''}${state.success ? `<div class="alert success">✓ <span>${esc(state.success)}</span><button data-act="dismiss">×</button></div>` : ''}${inner}</div></main></div>${state.uploadOpen ? uploadModal() : ''}${state.confirm ? confirmModal() : ''}${state.busy && state.busyText ? `<div class="busy-toast"><span class="spin">◌</span> ${esc(state.busyText)}</div>` : ''}`;
}
function heading(eyebrow, title, desc, right = '') {
  return `<div class="eyebrow">${eyebrow}</div><div class="page-heading"><div><h1>${esc(title)}<span class="title-period">.</span></h1><p>${esc(desc)}</p></div>${right}</div>`;
}
function dashboard() {
  const list = state.templates.map(summary);
  const cards = list.filter(t => (t.name + ' ' + t.filename).toLowerCase().includes(state.search.toLowerCase())).map(t => `
 <article class="template-card"><div class="template-card-top"><div class="doc-icon">${icons.file}</div><span class="pill">XLSX</span></div>
 <h3>${esc(t.name)}</h3><p class="filename">${esc(t.filename)}</p><div class="template-meta"><span>${icons.pin} ${t.field_count}개 필드</span><span>${icons.file} 첨부 ${t.attachment_count}개</span></div>
 <div class="template-card-actions"><button class="btn small soft" data-act="edit" data-id="${t.id}">⚙ 매핑 설정</button><button class="btn small dark" data-act="fill" data-id="${t.id}" ${!t.field_count ? 'disabled' : ''}>서류 작성 →</button></div><button class="delete-card" title="삭제" data-act="delete" data-id="${t.id}">×</button></article>`).join('');
  return heading('DOCUMENT WORKSPACE / OVERVIEW', '양식 라이브러리', '양식을 등록하고 입력 항목을 연결한 뒤 문서를 완성하세요.', `<button class="btn primary" data-act="upload-open">+ 새 양식 등록</button>`) + `
<div class="metrics"><div class="metric"><div class="metric-top"><span>등록된 양식</span>▤</div><strong>${list.length}<small>개</small></strong><p>사용 가능한 문서 템플릿</p></div>
<div class="metric"><div class="metric-top"><span>매핑된 입력 항목</span>⌖</div><strong>${list.reduce((a, t) => a + t.field_count, 0)}<small>개</small></strong><p>전체 양식의 입력 필드</p></div>
<div class="metric"><div class="metric-top"><span>시스템 상태</span>✓</div><strong class="metric-text">${store.mode === 'db' ? '준비 완료' : '임시 모드'}</strong><p>${store.mode === 'db' ? 'Excel → PDF 변환 가능 · 계정에 저장' : 'Excel → PDF 변환 가능 · 저장되지 않음'}</p></div></div>
<div class="section-heading"><div><h2>내 양식</h2><p>양식을 선택해 작성하거나 매핑 설정을 편집하세요.</p></div><div class="search-box">⌕ <input data-input="search" placeholder="양식 검색..." value="${esc(state.search)}"></div></div>
<div class="template-grid">${cards}<button class="add-template-card" data-act="upload-open"><span class="add-card-icon">+</span><b>새 양식 추가</b><span>Excel 파일을 가져와 입력 항목을 설정해요</span></button></div>
${list.length ? `<div class="info-banner"><div class="info-banner-icon">ⓘ</div><div><strong>샘플(출장비 정산서)로 먼저 테스트해보세요</strong><p>필요한 항목만 입력하면 여비 산출 내역서의 수식이 다시 계산되고, 지출결의 양식까지 채워진 PDF가 만들어집니다. 모든 처리는 이 브라우저 안에서 이뤄집니다.</p></div><button data-act="edit" data-id="${list[list.length - 1].id}">시작하기 →</button></div>` : ''}`;
}
function fieldOption(type) { return TYPES.map((t, i) => `<option value="${t}" ${type === t ? 'selected' : ''}>${['한 줄 텍스트', '긴 텍스트', '숫자', '날짜', '선택 목록', '체크박스'][i]}</option>`).join(''); }
function gridView(d, f) {
  const entry = state.book && state.book.cache.get(state.sheet); if (!entry) return '<div class="empty-state">시트를 불러오는 중...</div>';
  const m = entry.model;
  const range = { r1: 1, c1: 1, r2: Math.min(Math.max(m.maxR + 3, 24), 300), c2: Math.min(Math.max(m.maxC + 2, 10), 60) };
  const mapped = new Map(); d.fields.forEach((x, i) => { const a = XL.normAddr(x.cell); if (a && x.sheet === state.sheet) mapped.set(a, i + 1); });
  const active = f && f.sheet === state.sheet ? XL.normAddr(f.cell) : null;
  return { html: `<div class="sheet-scroll">${XL.sheetHTML(state.book, m, { range, grid: true, mapped, active, escape: esc })}</div>`, range };
}
function editor() {
  const d = state.draft; if (!d) return heading('TEMPLATE SETTINGS', '양식 매핑 설정', '양식 라이브러리에서 편집할 양식을 선택하세요.');
  const f = d.fields[state.selected];
  const cards = d.fields.map((x, i) => `<button class="field-item ${i === state.selected ? 'selected' : ''}" data-act="select-field" data-index="${i}"><span class="field-number">${String(i + 1).padStart(2, '0')}</span><span class="field-details"><b>${esc(x.label)}</b><small>${esc(x.sheet)}!${esc(x.cell || '셀 미지정')}</small></span><span class="mapping-tag ${x.cell ? 'complete' : ''}">${x.cell ? '✓' : '?'}</span></button>`).join('');
  const grid = gridView(d, f);
  const sheets = state.book.sheets.map(s => s.name);
  return heading('TEMPLATE SETTINGS / FIELD MAPPING', '양식 매핑 설정', '입력 항목을 만들고 오른쪽 시트의 셀을 클릭하여 연결하세요.', `<button class="btn primary" data-act="save">✓ 설정 저장</button>`) + `
<div class="editor-grid"><div class="editor-left">
<section class="panel editor-general"><div class="panel-title"><span class="section-icon">▤</span><div><h3>기본 정보</h3><p>등록한 양식의 이름과 형식</p></div></div><label class="input-label">양식 이름</label><input class="form-input" data-prop="name" value="${esc(d.name)}" maxlength="150"><label class="input-label">원본 파일</label><div class="file-row">▤ ${esc(d.filename)} · ${(d.size / 1024).toFixed(1)}KB</div></section>
<section class="panel fields-panel"><div class="panel-title spaced"><span class="section-icon">⌖</span><div><h3>입력 항목 <span class="count-badge">${d.fields.length}</span></h3><p>선택 후 시트 미리보기에서 연결</p></div><button class="icon-btn" data-act="add-field">+</button></div><div class="field-list">${cards}</div><button class="add-field" data-act="add-field">+ 입력 항목 추가</button></section>
<section class="panel"><div style="padding:23px"><div class="panel-title spaced"><span class="section-icon">⎙</span><div><h3>PDF 출력 시트</h3><p>선택한 시트만 순서대로 PDF에 포함됩니다.</p></div></div>${state.book.sheets.filter(s => s.state === 'visible').map(s => `<label class="check-toggle sheet-check"><input type="checkbox" data-print="${esc(s.name)}" ${(d.printSheets && d.printSheets.length ? d.printSheets.includes(s.name) : true) ? 'checked' : ''}> ${esc(s.name)}</label>`).join('')}</div></section>
<section class="panel attachments-panel"><div class="panel-title spaced"><span class="section-icon">▤</span><div><h3>첨부서류 안내 <span class="count-badge">${d.attachments.length}</span></h3><p>최종 페이지에 목록만 표시</p></div></div>${d.attachments.map((a, i) => `<div class="attach-edit"><input data-attachment="${i}" data-part="title" value="${esc(a.title)}" placeholder="첨부서류명" maxlength="150"><input data-attachment="${i}" data-part="note" value="${esc(a.note)}" placeholder="안내 문구 (선택)"><button data-act="delete-attachment" data-index="${i}">×</button></div>`).join('')}<button class="add-field" data-act="add-attachment">+ 첨부서류 추가</button></section>
</div><div class="editor-right"><section class="panel mapping-panel"><div class="panel-title spaced"><span class="section-icon teal">▤</span><div><h3>원본 시트 미리보기</h3><p>셀을 클릭하면 선택된 입력 항목에 연결됩니다. 병합 셀은 자동으로 왼쪽 위 셀이 선택돼요.</p></div><div class="sheet-selector"><select data-change="sheet">${sheets.map(s => `<option ${s === state.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div></div><div class="mapping-hint">⌖ 현재 연결할 항목: <b>${esc(f ? f.label : '입력 항목을 추가하세요')}</b> ${f && f.cell ? `<span class="current-cell">${esc(f.sheet)}!${esc(f.cell)}</span>` : ''}</div>${grid.html || grid}<div class="sheet-footnote">ⓘ 원본 서식으로 표시 · ${grid.range ? `${grid.range.r2}행 × ${XL.colLetters(grid.range.c2)}열까지` : ''} · 더 먼 셀은 아래에서 직접 주소 입력</div></section>
${f ? `<section class="panel inspector"><div class="panel-title spaced"><span class="section-icon">⚙</span><div><h3>선택한 항목 설정</h3><p>입력 형식과 연결 위치를 조정하세요.</p></div><button class="icon-btn danger" data-act="remove-field" title="항목 삭제">×</button></div><div class="inspector-form"><div class="form-group"><label>사용자에게 보일 항목명</label><input class="form-input" data-field="label" value="${esc(f.label)}" maxlength="100"></div><div class="form-group"><label>데이터 키 (영문, 중복 불가)</label><input class="form-input mono" data-field="key" value="${esc(f.key)}" maxlength="48"></div><div class="form-group"><label>입력 형식</label><select class="form-input" data-field="type">${fieldOption(f.type)}</select></div><div class="form-group"><label>시트</label><select class="form-input" data-field="sheet">${sheets.map(s => `<option ${s === f.sheet ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></div><div class="form-group"><label>셀 주소</label><input class="form-input mono" data-field="cell" value="${esc(f.cell)}" placeholder="예: B7"></div><div class="form-group"><label>입력 도움말</label><input class="form-input" data-field="placeholder" value="${esc(f.placeholder)}"></div>${f.type === 'select' ? `<div class="form-group full"><label>선택 옵션 (쉼표로 구분)</label><input class="form-input" data-field="options" value="${esc(f.options.join(', '))}"></div>` : ''}<label class="check-toggle"><input type="checkbox" data-field="required" ${f.required ? 'checked' : ''}> 필수 입력 항목</label></div></section>` : ''}
</div></div><div class="bottom-action"><div>✓ ${state.dirty ? '<b style="color:#c9853a">저장하지 않은 변경사항이 있어요.</b>' : '먼저 매핑 설정을 저장한 다음 PDF를 생성해보세요.'}</div><div><button class="btn soft" data-act="save">설정 저장</button><button class="btn dark" data-act="test">작성 테스트 →</button></div></div>`;
}
function fillPreview() {
  const t = state.active; if (!t || !state.book) return '';
  const inputs = previewInputs(t, state.values);
  if (state.calc) { try { state.calc.set(inputs); } catch (e) { console.warn(e); } }
  return printSheetsOf(t, state.book).map(name => {
    const entry = state.book.cache.get(name); if (!entry) return '';
    const values = state.calc ? state.calc.values(name) : new Map();
    const highlight = new Set();
    for (const x of inputs) if (x.sheet === name) { values.set(x.addr, x.value); if (x.filled) highlight.add(x.addr); }
    return `<div class="paper-label">${esc(name)}</div><div class="paper"><div class="paper-fit">${XL.sheetHTML(state.book, entry.model, { range: XL.printRange(state.book, name, entry.model), grid: false, values, highlight, escape: esc })}</div></div>`;
  }).join('');
}
function fitPapers() {
  document.querySelectorAll('.paper-fit').forEach(box => {
    const table = box.firstElementChild; if (!table) return; box.style.zoom = '';
    const w = box.parentElement.clientWidth - 40, k = Math.min(1, w / table.offsetWidth); if (k < 1) box.style.zoom = k.toFixed(3);
  });
}
function fill() {
  const t = state.active; if (!t) return '';
  const done = t.fields.filter(f => f.type === 'checkbox' ? state.values[f.key] : state.values[f.key] != null && state.values[f.key] !== '').length;
  const fields = t.fields.map(f => `<div class="user-field ${f.type === 'textarea' ? 'wide' : ''}"><label for="f_${esc(f.key)}">${esc(f.label)} ${f.required ? '<em>*</em>' : ''}</label>${f.type === 'textarea' ? `<textarea name="${esc(f.key)}" id="f_${esc(f.key)}" placeholder="${esc(f.placeholder || f.label + ' 입력')}" ${f.required ? 'required' : ''} maxlength="5000">${esc(state.values[f.key] || '')}</textarea>` : f.type === 'select' ? `<select name="${esc(f.key)}" id="f_${esc(f.key)}" ${f.required ? 'required' : ''}><option value="">선택하세요</option>${f.options.map(s => `<option ${state.values[f.key] === s ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select>` : f.type === 'checkbox' ? `<label class="checkbox-user"><input name="${esc(f.key)}" id="f_${esc(f.key)}" type="checkbox" ${f.required ? 'required' : ''} ${state.values[f.key] ? 'checked' : ''}> 예, 확인했습니다</label>` : `<input name="${esc(f.key)}" id="f_${esc(f.key)}" type="${f.type === 'date' ? 'date' : f.type === 'number' ? 'number' : 'text'}" ${f.type === 'number' ? 'step="any"' : ''} ${f.required ? 'required' : ''} placeholder="${esc(f.placeholder || f.label + ' 입력')}" value="${esc(state.values[f.key] || '')}">`}</div>`).join('');
  return heading('DOCUMENT GENERATOR / FORM ENTRY', '서류 작성', '필요한 정보만 입력하면 지정한 엑셀 셀에 자동 반영됩니다.', `<span class="step-pill">STEP 02 / 03</span>`) + `
<div class="fill-layout"><div><div class="panel fill-panel"><div class="fill-head"><div class="doc-icon">▤</div><div><h2>${esc(t.name)}</h2><span>${esc(t.filename)}</span></div></div><div class="progress-bar"><span style="width:${Math.round(done / Math.max(1, t.fields.length) * 100)}%"></span></div><form id="form-generate" novalidate><div class="form-section-title"><span>01</span><div><h3>기본 정보 입력</h3><p>별표(*) 표시 항목은 필수입니다.</p></div></div><div class="user-fields">${fields}</div><div class="fill-footer"><div>✓ 입력 정보는 이 브라우저 안에서만 처리되며 저장되지 않습니다.</div><button type="submit" class="btn primary" ${state.busy ? 'disabled' : ''}>${state.busy ? '생성 중...' : 'PDF 생성하기 →'}</button></div></form></div>
<section class="panel live-preview"><div class="panel-title"><span class="section-icon teal">▤</span><div><h3>실시간 미리보기</h3><p>입력한 내용이 원본 양식의 어느 위치에 들어가는지 바로 확인하세요.</p></div></div><div id="live-preview">${fillPreview()}</div></section></div>
<aside class="helper-panel"><div class="helper-icon">▤</div><h3>작성 안내</h3><p>원본 서식을 그대로 유지하면서 관리자가 연결한 셀에만 내용이 반영됩니다.</p><div class="helper-stat"><span>입력 항목</span><b>${t.fields.length}개</b></div><div class="helper-stat"><span>필수 항목</span><b>${t.fields.filter(f => f.required).length}개</b></div><div class="helper-stat"><span>작성 완료</span><b>${done}개</b></div><div class="helper-stat"><span>첨부서류 안내</span><b>${t.attachments.length}개</b></div><div class="helper-tip">ⓘ 마지막 화면에서는 첨부서류를 업로드하지 않고 목록만 확인합니다.</div></aside></div>`;
}
function result() {
  const r = state.result; if (!r) return '';
  const canSave = !!downloads;
  return heading('DOCUMENT GENERATOR / COMPLETE', '문서 작성 완료', '완성된 PDF를 확인하고 필요한 첨부서류를 준비하세요.', `<span class="complete-tag">✓ 생성 완료</span>`) + `<div class="result-layout"><div class="panel result-preview"><div class="result-top"><div><h3>▤ ${esc(r.template_name)}.pdf</h3><p>입력한 정보가 원본 양식에 적용됐습니다 · ${r.pages.length}쪽 · A4</p></div><button class="btn primary" data-act="save-pdf" ${canSave ? '' : 'disabled'}>↓ PDF 다운로드</button></div><div class="pages">${r.pages.map((p, i) => `<img src="${p.img}" alt="${i + 1}쪽 미리보기" class="${p.landscape ? 'land' : ''}">`).join('')}</div><div class="result-foot"><span>${canSave ? '다운로드 시 저장 확인 창이 표시됩니다.' : '이 화면에서는 파일 저장을 사용할 수 없습니다.'}</span><button class="link-btn" data-act="save-xlsx" ${canSave ? '' : 'disabled'}>▤ Excel 결과 받기</button></div></div><aside class="panel result-attachments"><div class="result-side-heading"><div class="helper-icon light">▤</div><h3>필요한 첨부서류</h3><p>PDF와 함께 필요한 문서를 준비하세요.</p></div><div class="attachment-list">${r.attachments.length ? r.attachments.map((a, i) => `<div class="attachment-row"><span class="attachment-number">${String(i + 1).padStart(2, '0')}</span><div><b>${esc(a.title)}</b>${a.note ? `<p>${esc(a.note)}</p>` : ''}</div></div>`).join('') : '<div class="no-attachments">별도 등록된 첨부서류가 없습니다.</div>'}</div><div class="readonly-note">ⓘ 안내 전용 화면입니다. 첨부파일 업로드 기능은 없습니다.</div><button class="btn soft full-width" data-act="refill">↺ 새로운 내용으로 작성</button></aside></div>`;
}
function uploadModal() {
  return `<div class="modal-backdrop" data-backdrop="upload"><div class="modal" role="dialog"><div class="modal-head"><div class="modal-icon">⇧</div><button class="icon-btn" data-act="upload-close">×</button></div><h2>새 양식 등록</h2><p>Excel 문서를 업로드하고 입력 항목을 직접 설정하세요. 파일은 ${store.mode === 'db' ? '내 계정 전용 저장소에' : '이번 방문 동안만'} 보관됩니다.</p><label class="input-label">양식 이름</label><input class="form-input" data-input="upload-name" placeholder="예: 구매 신청서" value="${esc(state.uploadName)}" maxlength="150"><label class="input-label">원본 Excel 파일</label><input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" id="upload-file-input" data-input="upload-file" style="display:none"><button class="upload-drop" data-act="choose-file" data-drop="1"><span style="font-size:26px">▤</span><b>${esc(state.uploadFile ? state.uploadFile.name : '파일을 선택하거나 끌어다 놓으세요')}</b><small>.xlsx · 최대 10MB</small></button><div class="modal-actions"><button class="btn soft" data-act="upload-close">취소</button><button class="btn primary" data-act="upload-submit" ${!state.uploadFile || state.busy ? 'disabled' : ''}>양식 등록 →</button></div></div></div>`;
}
function confirmModal() {
  return `<div class="modal-backdrop" data-backdrop="confirm"><div class="modal" role="alertdialog"><div class="modal-head"><div class="modal-icon" style="background:#fff1f2;color:#be5a6a">×</div></div><h2>${esc(state.confirm.title)}</h2><p>${esc(state.confirm.message)}</p><div class="modal-actions"><button class="btn soft" data-act="confirm-no">취소</button><button class="btn dark" data-act="confirm-yes">${esc(state.confirm.yes || '확인')}</button></div></div></div>`;
}
function loadingView() { return `<div class="center-loader"><span class="spin" style="font-size:26px">◌</span>양식을 불러오는 중...</div>`; }
function render() {
  const views = { loading: loadingView, dashboard, editor, fill, result };
  const scroll = document.querySelector('.sheet-scroll'); const pos = scroll ? [scroll.scrollLeft, scroll.scrollTop] : null;
  document.getElementById('app').innerHTML = layout(views[state.page]());
  const next = document.querySelector('.sheet-scroll'); if (pos && next) [next.scrollLeft, next.scrollTop] = pos;
  fitPapers();
  XL.fitText(document.getElementById('app'));
}

/* ---------- navigation ---------- */
const findT = id => { const t = state.templates.find(x => x.id === id); if (!t) fail('양식을 찾을 수 없습니다.'); return t; };
async function openEditor(id) {
  const t = findT(id); const book = await bookFor(t);
  state.active = t; state.draft = clone(t); state.book = book; state.selected = 0;
  const first = t.fields[0] && book.sheets.some(s => s.name === t.fields[0].sheet) ? t.fields[0].sheet : book.sheets[0].name;
  state.sheet = first; await XL.loadSheet(book, first);
  state.page = 'editor'; state.dirty = false;
}
async function openFill(id) {
  const t = findT(id);
  if (!t.fields.length) { await openEditor(id); fail('입력 항목이 없습니다. 매핑을 먼저 등록하세요.'); }
  const book = await bookFor(t);
  for (const s of new Set([...t.fields.map(f => f.sheet), ...printSheetsOf(t, book)])) if (book.sheets.some(x => x.name === s)) await XL.loadSheet(book, s);
  state.calc = null;
  try { state.calc = await XL.buildCalc(book); } catch (e) { console.warn('calc unavailable', e); }
  state.active = t; state.book = book; state.values = {}; state.page = 'fill';
}
async function save() {
  const d = state.draft; const name = d.name.trim(); if (!name) fail('양식 이름을 입력하세요.');
  const fields = clone(d.fields); await validateMapping(fields, state.book);
  if (d.printSheets && !d.printSheets.length) fail('PDF에 포함할 시트를 1개 이상 선택하세요.');
  const updated = { ...clone(state.active), name, fields, printSheets: d.printSheets || null, attachments: d.attachments.filter(x => x.title.trim()).map(x => ({ title: x.title.trim().slice(0, 150), note: x.note.trim() })), updated_at: nowIso() };
  await store.put(updated); await refresh();
  state.active = findT(updated.id); state.draft = clone(state.active); state.dirty = false;
  state.success = '설정 저장 완료! 작성 테스트를 진행할 수 있어요.';
}
async function run(task) {
  if (state.busy) return; state.busy = true; state.error = ''; state.success = ''; render();
  try { await task(); }
  catch (e) { console.error(e); state.error = e instanceof UserError ? e.message : (e && e.code === 'quota_exceeded') ? '저장 공간이 가득 찼습니다. 사용하지 않는 양식을 삭제하세요.' : (e && e.message) || String(e); }
  finally { state.busy = false; state.busyText = ''; render(); }
}
async function submitUpload() {
  const file = state.uploadFile; if (!file) fail('엑셀 파일을 선택하세요.');
  if (!/\.xlsx$/i.test(file.name)) fail('현재 MVP는 .xlsx 형식만 지원합니다.');
  if (file.size > MAX_UPLOAD) fail('파일 크기는 10MB 이하여야 합니다.');
  state.busyText = '양식을 저장하는 중...'; render();
  const bytes = new Uint8Array(await file.arrayBuffer());
  const t = await createTemplate((state.uploadName.trim() || file.name.replace(/\.xlsx$/i, '')).slice(0, 150), file.name, bytes);
  state.uploadOpen = false; state.uploadFile = null; state.uploadName = '';
  await refresh(); await openEditor(t.id);
  state.success = '양식을 등록했어요. 입력 항목을 추가해보세요.';
}

/* ---------- events ---------- */
async function onAction(act, el) {
  switch (act) {
    case 'dashboard': state.page = 'dashboard'; state.result = null; state.error = ''; render(); break;
    case 'editor-current': if (state.active) await run(() => openEditor(state.active.id)); else { state.page = 'dashboard'; state.error = '먼저 편집할 양식을 선택하세요.'; render(); } break;
    case 'upload-open': state.uploadOpen = true; state.error = ''; render(); break;
    case 'upload-close': state.uploadOpen = false; render(); break;
    case 'choose-file': document.getElementById('upload-file-input').click(); break;
    case 'upload-submit': await run(submitUpload); break;
    case 'edit': await run(() => openEditor(el.dataset.id)); break;
    case 'fill': await run(() => openFill(el.dataset.id)); break;
    case 'delete': { const t = findT(el.dataset.id); state.confirm = { title: '양식 삭제', message: `'${t.name}' 양식과 매핑 설정을 삭제할까요? 되돌릴 수 없습니다.`, yes: '삭제', go: async () => { await store.remove(t); books.delete(t.id); if (state.active && state.active.id === t.id) { state.active = null; state.draft = null; } await refresh(); state.page = 'dashboard'; state.success = '양식을 삭제했습니다.'; } }; render(); break; }
    case 'confirm-no': state.confirm = null; render(); break;
    case 'confirm-yes': { const go = state.confirm.go; state.confirm = null; await run(go); break; }
    case 'dismiss': state.error = ''; state.success = ''; render(); break;
    case 'add-field': { const d = state.draft; let i = d.fields.length + 1; while (d.fields.some(f => f.key === `field_${i}`)) i++;
      d.fields.push({ key: `field_${i}`, label: `새 항목 ${i}`, type: 'text', required: true, sheet: state.sheet, cell: '', placeholder: '', options: [] });
      state.selected = d.fields.length - 1; state.dirty = true; render(); break; }
    case 'remove-field': state.draft.fields.splice(state.selected, 1); state.selected = Math.max(0, state.selected - 1); state.dirty = true; render(); break;
    case 'select-field': { state.selected = +el.dataset.index; const f = state.draft.fields[state.selected];
      if (f.sheet && f.sheet !== state.sheet && state.book.sheets.some(s => s.name === f.sheet)) await run(async () => { await XL.loadSheet(state.book, f.sheet); state.sheet = f.sheet; });
      else render(); break; }
    case 'cell': { const f = state.draft.fields[state.selected]; if (!f) { state.error = '먼저 입력 항목을 추가하세요.'; render(); return; }
      f.sheet = state.sheet; f.cell = el.dataset.cell; state.dirty = true; state.error = ''; state.success = '';
      try {
        const opts = await XL.listOptions(state.book, f.sheet, f.cell);
        if (opts && opts.length && (f.type === 'text' || f.type === 'select')) { f.type = 'select'; f.options = opts; state.success = `이 셀의 드롭다운 목록(${opts.length}개)을 선택 옵션으로 가져왔어요.`; }
      } catch (e) { console.warn(e); }
      render(); break; }
    case 'save': await run(save); break;
    case 'test': if (state.dirty) { state.error = '저장하지 않은 변경사항이 있습니다. 먼저 설정을 저장하세요.'; state.success = ''; render(); } else await run(() => openFill(state.draft.id)); break;
    case 'add-attachment': state.draft.attachments.push({ title: '', note: '' }); state.dirty = true; render(); break;
    case 'delete-attachment': state.draft.attachments.splice(+el.dataset.index, 1); state.dirty = true; render(); break;
    case 'refill': await run(() => openFill(state.active.id)); break;
    case 'save-pdf': await run(() => saveFile(`${state.result.template_name}_작성완료.pdf`, state.result.pdf)); break;
    case 'save-xlsx': await run(() => saveFile(`${state.result.template_name}_작성완료.xlsx`, state.result.xlsx)); break;
  }
}
document.addEventListener('click', async e => {
  const bd = e.target.closest('[data-backdrop]');
  if (bd && e.target === bd) { if (bd.dataset.backdrop === 'upload') state.uploadOpen = false; else state.confirm = null; render(); return; }
  const el = e.target.closest('[data-act]'); if (!el || el.disabled) return;
  await onAction(el.dataset.act, el);
});
let previewTimer = 0;
function readForm(form) {
  const values = {};
  for (const f of state.active.fields) { const input = form.elements.namedItem(f.key); if (input) values[f.key] = f.type === 'checkbox' ? input.checked : input.value; }
  return values;
}
function updateFillLive(form) {
  state.values = readForm(form);
  clearTimeout(previewTimer);
  previewTimer = setTimeout(() => {
    const box = document.getElementById('live-preview'); if (box) { box.innerHTML = fillPreview(); XL.fitText(box); fitPapers(); }
    const t = state.active, done = t.fields.filter(f => f.type === 'checkbox' ? state.values[f.key] : state.values[f.key]).length;
    const bar = document.querySelector('.progress-bar span'); if (bar) bar.style.width = Math.round(done / Math.max(1, t.fields.length) * 100) + '%';
    const stat = document.querySelectorAll('.helper-stat b')[2]; if (stat) stat.textContent = done + '개';
  }, 180);
}
document.addEventListener('input', e => {
  const el = e.target;
  if (el.dataset.input === 'search') { state.search = el.value; const p = el.selectionStart; render(); const n = document.querySelector('[data-input="search"]'); if (n) { n.focus(); n.setSelectionRange(p, p); } return; }
  if (el.dataset.input === 'upload-name') { state.uploadName = el.value; return; }
  if (el.form && el.form.id === 'form-generate') { updateFillLive(el.form); return; }
  if (el.dataset.prop === 'name' && state.draft) { state.draft.name = el.value; state.dirty = true; return; }
  if (el.dataset.field && state.draft && state.draft.fields[state.selected]) {
    const f = state.draft.fields[state.selected], p = el.dataset.field;
    if (p === 'required') f.required = el.checked; else if (p === 'options') f.options = el.value.split(',').map(s => s.trim()).filter(Boolean); else f[p] = el.value;
    state.dirty = true; return;
  }
  if (el.dataset.print != null && state.draft) {
    const all = state.book.sheets.filter(s => s.state === 'visible').map(s => s.name);
    const cur = new Set(state.draft.printSheets && state.draft.printSheets.length ? state.draft.printSheets : all);
    if (el.checked) cur.add(el.dataset.print); else cur.delete(el.dataset.print);
    state.draft.printSheets = all.filter(n => cur.has(n)); state.dirty = true; return;
  }
  if (el.dataset.attachment != null && state.draft) { state.draft.attachments[+el.dataset.attachment][el.dataset.part] = el.value; state.dirty = true; }
});
document.addEventListener('change', async e => {
  const el = e.target;
  if (el.dataset.input === 'upload-file') { state.uploadFile = el.files && el.files[0] || null; render(); return; }
  if (el.form && el.form.id === 'form-generate') { updateFillLive(el.form); return; }
  if (el.dataset.change === 'sheet') { const name = el.value; await run(async () => { await XL.loadSheet(state.book, name); state.sheet = name; }); return; }
  if (el.dataset.field === 'sheet') { const name = el.value; await run(async () => { await XL.loadSheet(state.book, name); state.sheet = name; }); return; }
  // A blur-triggered re-render would replace the button being clicked; the action renders anyway.
  if ((el.dataset.field === 'cell' || el.dataset.field === 'type' || el.dataset.field === 'label') && !pressingAction) render();
});
let pressingAction = false;
document.addEventListener('pointerdown', e => { pressingAction = !!e.target.closest('[data-act]'); }, true);
document.addEventListener('pointerup', () => setTimeout(() => { pressingAction = false; }, 0), true);
document.addEventListener('dragover', e => { if (e.target.closest('[data-drop]')) e.preventDefault(); });
document.addEventListener('drop', e => { const z = e.target.closest('[data-drop]'); if (!z) return; e.preventDefault(); const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) { state.uploadFile = f; render(); } });
document.addEventListener('submit', async e => {
  if (e.target.id !== 'form-generate') return; e.preventDefault();
  state.values = readForm(e.target);
  const t = state.active;
  await run(async () => { state.result = await generate(t, state.values); state.page = 'result'; });
});

/* ---------- boot ---------- */
(async () => {
  render();
  await store.init();
  try { await seedSample(); await refresh(); state.page = 'dashboard'; }
  catch (e) { console.error(e); state.page = 'dashboard'; state.error = '저장소를 불러오지 못했습니다: ' + (e.message || e.code || e); }
  render();
  initDownloads();
})();

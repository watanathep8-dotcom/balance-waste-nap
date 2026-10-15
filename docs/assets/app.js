/* ================= Balance Waste — ระบบบริหารจัดการของเสีย ================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const view = $('#view');
const S = { types: [], balance: {}, containers: {}, company: {}, limits: {}, charts: [] };
const store = (fn, k, v) => { try { return fn === 'get' ? localStorage.getItem(k) : fn === 'del' ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch { return null; } };
/* ผู้ใช้งาน: ชื่อที่ตั้งไว้มุมขวาบน (เก็บในเบราว์เซอร์นี้) — ส่งไปกับทุกการบันทึก เป็น “ผู้บันทึก” ในประวัติ / Log */
S.user = store('get', 'wasteUser') || '';

const KIND = {
  store:   { name: 'บันทึกจัดเก็บ', short: 'จัดเก็บ', icon: 'fa-box-archive', c: '#10B981', cls: 'in', sign: '+' },
  ship:    { name: 'ส่งกำจัด / นำออก', short: 'ส่งกำจัด', icon: 'fa-truck-arrow-right', c: '#F97316', cls: 'out', sign: '−' },
  opening: { name: 'ยอดยกมา', short: 'ยอดยกมา', icon: 'fa-flag-checkered', c: '#6366F1', cls: 'open', sign: '+' },
  treat:   { name: 'บำบัดเองภายใน', short: 'บำบัดเอง', icon: 'fa-flask-vial', c: '#7C3AED', cls: 'treat', sign: '−' },
};
// Movement kinds the current view may show (treatment only while unlocked).
const kinds = () => Object.entries(KIND).filter(([k]) => k !== 'treat' || treatOn());
const CONT = {
  drum200: { name: 'ถัง 200 ลิตร', sub: 'Drum 200 L' },
  ibc:     { name: 'IBC', sub: 'IBC Tank 1,000 L' },
  bigbag:  { name: 'Bigbag', sub: 'Jumbo bag' },
};
const CONT_KEYS = Object.keys(CONT);
const PAGES = {
  dashboard: ['แดชบอร์ด', 'ภาพรวมของเสียในโรงงาน · โควตาตามหนังสืออนุญาต สก.2'],
  store: ['บันทึกจัดเก็บ', 'ลงข้อมูลของเสียที่จัดเก็บในโรงงาน — เลือกประเภท กรอกน้ำหนัก เลือกภาชนะ'],
  ship: ['ส่งกำจัด / นำออก', 'ส่งของเสียให้ผู้รับดำเนินการ — ไม่เกินยอดคงเหลือ'],
  treat: ['บำบัดเองภายใน', 'บำบัดของเสียด้วยระบบของโรงงานเอง — ตัดยอดคงเหลือ ไม่นับโควตา สก.2'],
  balance: ['คงเหลือ & โควตา', 'ยอดคงเหลือในโรงงาน ภาชนะ และการใช้โควตาปีนี้'],
  history: ['ประวัติรายการ', 'ทุกรายการจัดเก็บ / ส่งกำจัด · ค้นหา · ส่งออก Excel'],
  logs: ['บันทึกประวัติ (Log)', 'ทุกการกระทำในระบบ — ใคร ทำอะไร เมื่อไร'],
  settings: ['ตั้งค่าประเภทของเสีย', 'รหัสของเสีย ปริมาณที่ได้รับอนุญาต ผู้รับดำเนินการ'],
};

/* ---------- API (Google Apps Script web app) ---------- */
// อ่านข้อมูล = GET ?action=...&_=ts  /  บันทึก = POST text/plain (JSON) — ไม่เกิด CORS preflight
const API_URL = String((window.WASTE_CONFIG || {}).apiUrl || '').trim();
const NO_API_MSG = 'ยังไม่ได้ตั้งค่า apiUrl — ใส่ URL ของ Google Apps Script (/exec) ในไฟล์ docs/config.js';
async function gasFetch(url, opt) {
  let j;
  try {
    const res = await fetch(url, { mode: 'cors', credentials: 'omit', cache: 'no-store', ...opt });
    j = await res.json();
  } catch { throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ (ตรวจสอบอินเทอร์เน็ต / Google Apps Script)'); }
  if (!j || !j.ok) {
    const e = new Error((j && j.error) || 'เกิดข้อผิดพลาด');
    e.passwordError = !!(j && j.passwordError);
    e.code = (j && j.code) || '';
    throw e;
  }
  return { ok: true, ...(j.data || {}) };
}
function gasGet(action, params = '') {
  if (!API_URL) return Promise.reject(new Error(NO_API_MSG));
  const u = new URL(API_URL, location.href);
  new URLSearchParams(params.replace(/^&/, '')).forEach((v, k) => u.searchParams.set(k, v));
  u.searchParams.set('action', action);
  u.searchParams.set('_', Date.now());
  return gasFetch(u.toString(), { method: 'GET' });
}
// Actions that change data: the user's name (top-right) must be set first.
const NAME_REQUIRED = ['store', 'ship', 'opening', 'treat', 'movement_update', 'movement_delete', 'type_save', 'reset_data'];
async function gasPost(payload) {
  if (!API_URL) throw new Error(NO_API_MSG);
  if (NAME_REQUIRED.includes(payload.action) && !S.user) {
    await askUserName(true);
    if (!S.user) throw new Error('กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนบันทึก');
  }
  const body = { actor: S.user || '', ...payload };
  if (S.tt) body.treatToken = S.tt; // in the POST body only, never in a URL
  try {
    return await gasFetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify(body) });
  } catch (e) {
    if (e.code === 'TREAT_LOCKED') dropTreat();
    throw e;
  }
}
// Reads: GET normally; POST while the treatment view is unlocked (token in the body).
function readApi(action, params = '') {
  if (!S.tt) return gasGet(action, params);
  return gasPost({ ...Object.fromEntries(new URLSearchParams(params.replace(/^&/, ''))), action });
}
async function api(action, body, params = '') {
  if (!body) return readApi(action, params);
  return gasPost({ ...body, action });
}

/* ---------- บำบัดเองภายใน: ซ่อนไว้โดยค่าเริ่มต้น ----------
   ปลดล็อกด้วยรหัสผ่านผู้ดูแล → เซิร์ฟเวอร์ออก token ชั่วคราว (เก็บใน sessionStorage ของแท็บนี้เท่านั้น)
   ระหว่างปลดล็อก การอ่านจะส่งแบบ POST พร้อม token และไม่เก็บแคชใน localStorage · ล็อกเมื่อไร token ถูกลบทันที */
const TT_KEY = 'wasteTT';
function ttLoad() {
  try {
    const v = JSON.parse(sessionStorage.getItem(TT_KEY) || 'null');
    if (v && /^[0-9a-f]{64}$/.test(v.t) && v.exp > Date.now()) return v.t;
    sessionStorage.removeItem(TT_KEY);
  } catch {}
  return '';
}
S.tt = ttLoad();
const treatOn = () => !!S.tt;
function paintTreat() {
  const on = treatOn();
  // the menu is always there; while locked it carries a lock badge and asks for the admin password on click
  $('#treatNav').classList.toggle('locked', !on);
  $('#treatNav').title = on ? '' : 'ต้องใส่รหัสผ่านผู้ดูแลเพื่อเปิด';
  const b = $('#treatLock');
  b.classList.toggle('on', on);
  b.innerHTML = `<i class="fa-solid ${on ? 'fa-lock-open' : 'fa-lock'}"></i>`;
  b.title = on ? 'กำลังแสดงข้อมูลบำบัดเองภายใน — คลิกเพื่อซ่อน (ล็อก)' : 'แสดงข้อมูลบำบัดเองภายใน (ต้องใช้รหัสผ่านผู้ดูแล)';
  document.body.classList.toggle('treat-on', on);
}
// Token gone (locked here, expired or refused by the server): back to the hidden view.
function dropTreat(msg) {
  if (!S.tt) return;
  S.tt = '';
  try { sessionStorage.removeItem(TT_KEY); } catch {}
  rcClear();
  paintTreat();
  if (msg) toast(msg, 'info');
}
function unlockTreat() {
  if (S.tt) return Promise.resolve(true);
  return new Promise(resolve => {
    const box = $('#pwModal'), card = $('#pwCard');
    card.innerHTML = `<div class="treat-ic sm" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-flask-vial"></i></span></div>
      <h3 style="justify-content:center">แสดงข้อมูลบำบัดเองภายใน</h3>
      <p class="hint" style="text-align:center">เมนูบำบัดเอง ยอดบำบัด และรายการบำบัดในประวัติ จะแสดงเฉพาะแท็บนี้หลังใส่รหัสผ่านผู้ดูแล (สูงสุด 6 ชั่วโมง หรือจนกว่าจะกดล็อก / ปิดแท็บ)</p>
      <div class="field mt"><label>รหัสผ่านผู้ดูแล</label><input type="password" id="ttInput" autocomplete="off"></div>
      <div id="ttErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <div class="modal-acts"><button class="btn btn-ghost" id="ttCancel">ยกเลิก</button><button class="btn btn-treat" id="ttOk"><i class="fa-solid fa-lock-open"></i>ปลดล็อก</button></div>`;
    box.hidden = false;
    const inp = $('#ttInput', card), ok = $('#ttOk', card), err = $('#ttErr', card);
    const done = v => { box.hidden = true; box.onclick = null; card.innerHTML = ''; resolve(v); };
    const submit = async () => {
      const pw = inp.value;
      if (!pw) { err.textContent = 'กรุณาใส่รหัสผ่าน'; inp.focus(); return; }
      ok.disabled = true; ok.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังตรวจสอบ...'; err.textContent = '';
      try {
        const r = await gasPost({ action: 'treat_unlock', resetPassword: pw });
        S.tt = r.token;
        try { sessionStorage.setItem(TT_KEY, JSON.stringify({ t: r.token, exp: r.expires_at })); } catch {}
        rcClear(); paintTreat(); done(true);
        toast('แสดงข้อมูลบำบัดเองภายในแล้ว (แท็บนี้)', 'info');
      } catch (e) {
        err.textContent = e.message; inp.value = ''; inp.focus();
        ok.disabled = false; ok.innerHTML = '<i class="fa-solid fa-lock-open"></i>ปลดล็อก';
      }
    };
    ok.onclick = submit;
    $('#ttCancel', card).onclick = () => done(false);
    box.onclick = e => { if (e.target === box) done(false); };
    inp.onkeydown = e => { e.stopPropagation(); if (e.key === 'Enter') submit(); if (e.key === 'Escape') done(false); };
    inp.focus();
  });
}
async function lockTreat() {
  const tok = S.tt;
  dropTreat('ซ่อนข้อมูลบำบัดเองภายในแล้ว');
  if (tok) gasFetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ action: 'treat_lock', treatToken: tok }) }).catch(() => {});
  if (currentPage() === 'treat') location.hash = 'dashboard'; else route();
}
/* แก้ไข / ลบรายการ, Log, ตั้งค่าประเภท: ใช้รหัสผ่านผู้ดูแล (WASTE_RESET_PASSWORD)
   เก็บในหน่วยความจำเท่านั้นหลังตรวจผ่าน (หายเมื่อรีเฟรช/ปิดหน้า) และล้างทิ้งเมื่อเซิร์ฟเวอร์แจ้งรหัสผิด/ถูกล็อก */
let recordPassword = '';
async function recordApi(action, body, reads = routeReads()) {
  const pw = await askPassword();
  try {
    const r = await gasPost({ ...body, ...readsField(reads), action, resetPassword: pw });
    afterWrite(r, reads);
    return r;
  } catch (e) {
    if (e.passwordError) recordPassword = '';
    throw e;
  }
}
/* ถามรหัสครั้งเดียวต่อการเปิดหน้า แล้วตรวจกับเซิร์ฟเวอร์ (verifyResetPassword) */
function askPassword() {
  if (recordPassword) return Promise.resolve(recordPassword);
  if (S.pwPending) return S.pwPending;
  S.pwPending = new Promise((resolve, reject) => {
    const box = $('#pwModal'), card = $('#pwCard');
    card.innerHTML = `<h3><i class="fa-solid fa-user-shield" style="color:var(--brand)"></i> รหัสผ่านผู้ดูแล</h3>
      <p class="hint">การแก้ไข / ลบรายการ การดูบันทึกประวัติ (Log) การตั้งค่าประเภทของเสีย และการรีเซ็ตข้อมูล ต้องใช้รหัสผ่าน (ตั้งโดยผู้ดูแลระบบ · ถามครั้งเดียวต่อการเปิดหน้านี้)</p>
      <div class="field mt"><label>รหัสผ่านผู้ดูแล</label><input type="password" id="pwInput" autocomplete="off"></div>
      <div id="pwErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <div class="modal-acts"><button class="btn btn-ghost" id="pwCancel">ยกเลิก</button><button class="btn btn-primary" id="pwOk"><i class="fa-solid fa-unlock"></i>ยืนยัน</button></div>`;
    box.hidden = false;
    const inp = $('#pwInput', card), ok = $('#pwOk', card), err = $('#pwErr', card);
    const done = () => { box.hidden = true; box.onclick = null; card.innerHTML = ''; };
    const cancel = () => { done(); reject(new Error('ยกเลิก — ต้องใส่รหัสผ่านผู้ดูแลก่อน')); };
    const submit = async () => {
      const pw = inp.value;
      if (!pw) { err.textContent = 'กรุณาใส่รหัสผ่าน'; inp.focus(); return; }
      ok.disabled = true; ok.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังตรวจสอบ...'; err.textContent = '';
      try {
        await gasPost({ action: 'verifyResetPassword', resetPassword: pw });
        recordPassword = pw; done(); resolve(pw);
      } catch (e) {
        err.textContent = e.message; inp.value = ''; inp.focus();
        ok.disabled = false; ok.innerHTML = '<i class="fa-solid fa-unlock"></i>ยืนยัน';
      }
    };
    ok.onclick = submit;
    $('#pwCancel', card).onclick = cancel;
    box.onclick = e => { if (e.target === box) cancel(); };
    inp.onkeydown = e => { e.stopPropagation(); if (e.key === 'Enter') submit(); if (e.key === 'Escape') cancel(); };
    inp.focus();
  }).finally(() => { S.pwPending = null; });
  return S.pwPending;
}
/* บันทึกประวัติ (Log): ส่งแบบ POST พร้อมรหัส — รหัสไม่อยู่ใน URL */
async function logsApi(action, params) {
  const pw = await askPassword();
  try {
    return await gasPost({ ...Object.fromEntries(new URLSearchParams(params.replace(/^&/, ''))), action, resetPassword: pw });
  } catch (e) {
    if (e.passwordError) recordPassword = '';
    throw e;
  }
}
/* ส่งออก Excel (CSV, UTF-8 BOM ให้ Excel ภาษาไทยอ่านถูก) */
async function downloadCsv(action, params) {
  try {
    const r = action === 'logs_export' ? await logsApi(action, params) : await readApi(action, params);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['\uFEFF' + r.csv], { type: 'text/csv;charset=utf-8' }));
    a.download = r.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) { toast(e.message, 'err'); }
}
if (!API_URL) $('#cfgBanner').hidden = false;
function applyBoot(j) {
  Object.assign(S, { types: j.types, balance: j.balance, containers: j.containers || {}, company: j.company || {}, limits: j.limits || {}, now: j.now });
  if (S.tt && !j.treat) dropTreat('หมดเวลาแสดงข้อมูลบำบัดเองภายใน — ซ่อนแล้ว (ปลดล็อกใหม่ได้ที่ไอคอนกุญแจด้านล่าง)');
}

/* ---------- อ่านหลายรายการในคำขอเดียว (batch) ---------- */
const readsSpec = reads => reads.map(([a, p = '']) => ({ action: a, ...Object.fromEntries(new URLSearchParams(p.replace(/^&/, ''))) }));
const readsField = reads => reads.length ? { reads: readsSpec(reads) } : {};
async function readMany(reads) {
  const r = S.tt ? await gasPost({ action: 'batch', reads: readsSpec(reads) })
    : await gasGet('batch', '&reads=' + encodeURIComponent(JSON.stringify(readsSpec(reads))));
  return r.results.map(d => ({ ok: true, ...d }));
}

/* ---------- แคชผลการอ่านล่าสุดในเบราว์เซอร์ (แสดงทันที แล้วโหลดใหม่เบื้องหลัง) ----------
   ใช้เพื่อแสดงผลเท่านั้น — การบันทึกตรวจยอดที่เซิร์ฟเวอร์เสมอ · Log ไม่เก็บ · ล้างทั้งหมดหลังบันทึกทุกครั้ง
   ระหว่างแสดงข้อมูลบำบัดเอง (S.tt) ไม่อ่าน/เขียนแคชนี้เลย */
const RC_PREFIX = 'wasteRC:';
const RC_FRESH_MS = 5000;
const rcKey = ([a, p = '']) => RC_PREFIX + a + p;
function rcGet(r) {
  if (S.tt) return null;
  try { const v = JSON.parse(localStorage.getItem(rcKey(r)) || 'null'); return v && v.d ? v : null; } catch { return null; }
}
function rcPut(r, d) {
  if (!d || S.tt) return;
  try { const s = JSON.stringify({ t: Date.now(), d }); if (s.length < 500000) localStorage.setItem(rcKey(r), s); } catch {}
}
function rcClear() {
  try { Object.keys(localStorage).filter(k => k.startsWith(RC_PREFIX)).forEach(k => localStorage.removeItem(k)); } catch {}
}
function afterWrite(r, reads) {
  rcClear();
  if (Array.isArray(r?.reads) && r.reads.length === reads.length) r.reads.forEach((d, i) => rcPut(reads[i], { ok: true, ...d }));
}
async function writeApi(action, body, reads = routeReads()) {
  const r = await api(action, { ...body, ...readsField(reads) });
  afterWrite(r, reads);
  return r;
}

/* ---------- helpers ---------- */
const fmt = (n, d = 0) => Number(n || 0).toLocaleString('th-TH', { maximumFractionDigits: d });
const fmtKg = n => fmt(n, 2);
const fmtT = kg => fmt(Number(kg || 0) / 1000, Math.abs(kg) < 10000 ? 3 : 2);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const typeById = id => S.types.find(t => +t.id === +id);
const balOf = id => S.balance[id] || { balance_kg: 0, containers: { drum200: 0, ibc: 0, bigbag: 0 }, shipped_year_kg: 0, permit_kg: 0, quota_pct: null, quota_level: 'ok' };
const dtTH = s => { const d = new Date(s.slice(0, 10) + 'T00:00:00'); return d.toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: '2-digit' }) + ' ' + s.slice(11, 16) + ' น.'; };
const dateTH = s => new Date(s.slice(0, 10) + 'T00:00:00').toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: 'numeric' });
const nowParts = () => { const d = new Date(); const p = n => String(n).padStart(2, '0'); return { date: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`, time: `${p(d.getHours())}:${p(d.getMinutes())}` }; };
const receivers = () => [...new Set(S.types.map(t => t.receiver))];
const shortRecv = r => String(r || '').replace(/^(บจก\.|บมจ\.|บริษัท)\s*/, '').replace('บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม', 'บริหารและพัฒนาฯ (GENCO)');
const typeLabel = t => t ? `${t.code} · ${t.name}` : '—';
const lvName = { ok: 'ปกติ', warn: 'ใกล้เต็มโควตา', over: 'เกินโควตา' };
// Weight text: up to 2 decimals, > 0 (also checked by the server).
const parseKg = v => { const s = String(v ?? '').replace(/,/g, '').trim(); return /^\d+(\.\d{1,2})?$/.test(s) ? Number(s) : NaN; };

/* ---------- container illustrations (inline SVG) ---------- */
function containerSVG(k, mono = false) {
  const W = mono ? '#fff' : null;
  if (k === 'drum200') {
    const body = W || '#0D9488', top = mono ? 'rgba(255,255,255,.75)' : '#5EEAD4', band = mono ? 'rgba(4,120,87,.55)' : '#115E59';
    return `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M15 12v40c0 3 7.6 5.5 17 5.5S49 55 49 52V12" fill="${body}"/>
      <path d="M15 25c0 3 7.6 5.5 17 5.5S49 28 49 25M15 40c0 3 7.6 5.5 17 5.5S49 43 49 40" stroke="${band}" stroke-width="2.4" fill="none"/>
      <ellipse cx="32" cy="12" rx="17" ry="5.5" fill="${top}"/><ellipse cx="38" cy="11.5" rx="3" ry="1.2" fill="${band}"/>
      <rect x="20" y="16" width="3.5" height="34" rx="1.7" fill="#fff" opacity=".28"/></svg>`;
  }
  if (k === 'ibc') {
    const cage = mono ? '#fff' : '#64748B', tank = mono ? 'rgba(255,255,255,.22)' : '#F1F5F9', liq = mono ? 'rgba(255,255,255,.7)' : '#F59E0B', base = mono ? 'rgba(255,255,255,.85)' : '#A16207';
    return `<svg viewBox="0 0 64 64" aria-hidden="true"><rect x="11" y="12" width="42" height="38" rx="4" fill="${tank}"/>
      <rect x="13" y="29" width="38" height="19" rx="3" fill="${liq}"/>
      <path d="M11 12h42v38H11zM25 12v38M39 12v38M11 25h42M11 38h42" stroke="${cage}" stroke-width="2" fill="none"/>
      <rect x="27" y="7" width="10" height="5" rx="1.5" fill="${cage}"/><rect x="8" y="50" width="48" height="5" rx="1.5" fill="${base}"/>
      <rect x="10" y="55" width="6" height="4" fill="${base}"/><rect x="29" y="55" width="6" height="4" fill="${base}"/><rect x="48" y="55" width="6" height="4" fill="${base}"/></svg>`;
  }
  const bag = mono ? 'rgba(255,255,255,.9)' : '#ECFDF5', edge = mono ? '#fff' : '#059669', loop = mono ? '#FDE68A' : '#F59E0B';
  return `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M17 17c0-7 6-9 6-9M47 17c0-7-6-9-6-9" stroke="${loop}" stroke-width="3.5" fill="none" stroke-linecap="round"/>
    <path d="M14 17h36l4 36c.3 3-1.5 5-4 5H14c-2.5 0-4.3-2-4-5z" fill="${bag}" stroke="${edge}" stroke-width="2.2"/>
    <path d="M14 17l-1 9M50 17l1 9" stroke="${edge}" stroke-width="1.6"/><path d="M21 30h22M20 38h24M19 46h26" stroke="${edge}" stroke-width="1.4" stroke-dasharray="3 3" opacity=".7"/>
    <circle cx="32" cy="38" r="6" fill="none" stroke="${loop}" stroke-width="2"/><path d="M29.5 38l2 2 3.5-4" stroke="${loop}" stroke-width="2" fill="none"/></svg>`;
}
const ecoLogo = (cls = '') => `<div class="eco-logo ${cls}" aria-hidden="true"><span class="ring"></span><span class="core"><i class="fa-solid fa-recycle"></i></span></div>`;
const loader = (msg = 'กำลังโหลดข้อมูล...') => `<div class="loader">${ecoLogo()}<small>${msg}</small></div>`;

/* ---------- UI effects ---------- */
document.addEventListener('pointerdown', e => {
  const b = e.target.closest('.btn, .tpick, .quick button, .seg button');
  if (!b) return;
  const r = b.getBoundingClientRect(), d = Math.max(r.width, r.height);
  const sp = document.createElement('span');
  sp.className = 'ripple';
  Object.assign(sp.style, { width: d + 'px', height: d + 'px', left: e.clientX - r.left - d / 2 + 'px', top: e.clientY - r.top - d / 2 + 'px' });
  if (getComputedStyle(b).position === 'static') b.style.position = 'relative';
  b.style.overflow = 'hidden';
  b.appendChild(sp);
  setTimeout(() => sp.remove(), 700);
});
function toast(msg, type = 'ok') {
  const m = { ok: ['fa-circle-check', '#10B981'], err: ['fa-circle-xmark', '#EF4444'], info: ['fa-recycle', '#0D9488'] }[type];
  const t = document.createElement('div');
  t.className = 'toast';
  t.style.setProperty('--c', m[1]);
  t.innerHTML = `<i class="fa-solid ${m[0]}"></i><div>${esc(msg)}</div>`;
  $('#toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 400); }, type === 'err' ? 5500 : 3500);
}
const reduceMotion = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
function confetti() {
  if (reduceMotion()) return;
  const cv = $('#confetti'), cx = cv.getContext('2d');
  cv.width = innerWidth; cv.height = innerHeight;
  const cols = ['#10B981', '#14B8A6', '#F59E0B', '#34D399', '#047857', '#FCD34D', '#6EE7B7'];
  const ps = Array.from({ length: 150 }, () => ({
    x: innerWidth / 2 + (Math.random() - .5) * 200, y: innerHeight / 2.2,
    vx: (Math.random() - .5) * 18, vy: -Math.random() * 18 - 6, s: Math.random() * 8 + 5,
    r: Math.random() * 6, vr: (Math.random() - .5) * .4, c: cols[Math.random() * cols.length | 0], leaf: Math.random() < .35,
  }));
  let f = 0;
  const run = S.confettiRun = (S.confettiRun || 0) + 1;
  setTimeout(() => { if (S.confettiRun === run) { S.confettiRun++; cx.clearRect(0, 0, cv.width, cv.height); } }, 2600);
  (function loop() {
    if (S.confettiRun !== run) return;
    cx.clearRect(0, 0, cv.width, cv.height);
    ps.forEach(p => {
      p.vy += .45; p.vx *= .99; p.x += p.vx; p.y += p.vy; p.r += p.vr;
      cx.save(); cx.translate(p.x, p.y); cx.rotate(p.r); cx.fillStyle = p.c; cx.globalAlpha = Math.max(0, 1 - f / 140);
      if (p.leaf) { cx.beginPath(); cx.ellipse(0, 0, p.s / 1.4, p.s / 3, 0, 0, Math.PI * 2); cx.fill(); } else cx.fillRect(-p.s / 2, -p.s / 4, p.s, p.s / 2);
      cx.restore();
    });
    if (++f < 140) requestAnimationFrame(loop); else cx.clearRect(0, 0, cv.width, cv.height);
  })();
}
function modal(html, wide = false) {
  $('#modalCard').className = 'modal-card' + (wide ? ' wide' : '');
  $('#modalCard').innerHTML = html;
  $('#modal').hidden = false;
  return $('#modalCard');
}
function closeModal() {
  if ($('#modal').hidden) return;
  $('#modal').hidden = true;
  const fn = S.onModalClose; S.onModalClose = null;
  if (fn) fn();
}
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal' || e.target.closest('[data-close]')) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
function countUp(root = view) {
  $$('[data-count]', root).forEach(el => {
    const to = +el.dataset.count, dec = +(el.dataset.dec || 0), t0 = performance.now(), dur = reduceMotion() ? 1 : 1200;
    const step = t => {
      const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 4);
      el.textContent = fmt(p < 1 ? to * e : to, dec);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}
function killCharts() { S.charts.forEach(c => c.destroy()); S.charts = []; }
document.addEventListener('animationend', e => {
  if (e.animationName === 'rise' && e.target.parentElement?.classList.contains('stagger')) {
    e.target.style.animation = 'none'; e.target.style.opacity = '1';
  }
});

/* clock */
setInterval(tickClock, 1000); tickClock();
function tickClock() {
  const d = new Date();
  $('#clock').textContent = d.toLocaleTimeString('th-TH', { hour12: false });
  $('#clockDate').textContent = d.toLocaleDateString('th-TH', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
}
$('#menuBtn').onclick = () => $('#sidebar').classList.toggle('open');
$('#navOverlay').onclick = () => $('#sidebar').classList.remove('open');

/* ================= ผู้ใช้งาน (ชื่อมุมขวาบน) ================= */
function paintUser() {
  $('#userName').textContent = S.user || 'ระบุชื่อผู้ใช้';
  $('#userAv').innerHTML = S.user ? esc(S.user.trim().charAt(0)) : '<i class="fa-solid fa-user"></i>';
  $('#userChip').classList.toggle('need', !S.user);
}
// required = opened because a save needs a name: keep the form, resolve when closed.
function askUserName(required = false) {
  return new Promise(resolve => {
    const m = modal(`<h3><i class="fa-solid fa-user-pen" style="color:var(--brand)"></i> ผู้ใช้งาน</h3>
      <p class="hint">${required ? '<b style="color:var(--bad)">กรุณาระบุชื่อผู้ใช้งานก่อนบันทึกรายการ</b> · ' : ''}ชื่อนี้จะถูกบันทึกเป็น “ผู้บันทึก” ในประวัติและ Log ทุกรายการ</p>
      <div class="field mt"><label>ชื่อ - นามสกุล</label><input id="uName" value="${esc(S.user)}" placeholder="เช่น สมชาย ใจดี" maxlength="60"></div>
      <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="uSave"><i class="fa-solid fa-check"></i>บันทึก</button></div>`);
    const inp = $('#uName', m); inp.focus(); inp.select();
    let done = false;
    const finish = () => { if (!done) { done = true; resolve(S.user); } };
    const save = () => {
      const name = inp.value.trim().slice(0, 60);
      if (required && !name) { inp.focus(); return; }
      S.user = name; store('set', 'wasteUser', S.user); paintUser(); closeModal(); finish();
      toast(S.user ? `สวัสดีคุณ${S.user}` : 'ล้างชื่อผู้ใช้แล้ว', 'info');
      if (!required) route();
    };
    $('#uSave', m).onclick = save;
    inp.onkeydown = e => { if (e.key === 'Enter') save(); };
    S.onModalClose = finish;
  });
}
$('#userChip').onclick = () => askUserName(false);
paintUser();
$('#treatLock').onclick = async () => { if (S.tt) lockTreat(); else if (await unlockTreat()) route(); };
paintTreat();

/* ================= ROUTER ================= */
const monthFirst = () => nowParts().date.slice(0, 8) + '01';
const histDefault = () => '&' + new URLSearchParams({ from: monthFirst(), to: nowParts().date, type: '', kind: '', receiver: '', q: '' });
function pageReads(p) {
  return ({ dashboard: [['dashboard', '']], history: [['history', histDefault()]] })[p] || [];
}
const currentPage = () => { const page = (location.hash.slice(1) || 'dashboard').split('?')[0]; return PAGES[page] ? page : 'dashboard'; };
const routeReads = (p = currentPage()) => [['bootstrap', ''], ...pageReads(p)];
const PAGE_FN = {
  dashboard: ([d]) => dashboard(d), history: ([h]) => history(h), balance, logs, settings,
  store: () => txForm('store'), ship: () => txForm('ship'), treat: () => txForm('treat'),
};
['input', 'change'].forEach(ev => view.addEventListener(ev, () => { S.dirty = true; }, true));
view.addEventListener('pointerdown', e => { if (e.target.closest('button, input, select, textarea, label, .ccard, .tpick')) S.dirty = true; }, true);
const sigOf = res => JSON.stringify(res.map((d, i) => (i ? d : { ...d, now: null })));
let routeSeq = 0;
const PW_PAGES = ['logs', 'settings'];
const NAME_PAGES = ['store', 'ship', 'treat'];
const INPUT_PAGES = ['store', 'ship', 'treat', 'history', 'settings'];
async function route() {
  const p = currentPage();
  if (p === 'treat' && !S.tt) {
    if (!(await unlockTreat())) { if (currentPage() === p) location.hash = 'dashboard'; return; }
  }
  if (PW_PAGES.includes(p)) {
    try { await askPassword(); }
    catch (e) { toast(e.message, 'err'); if (currentPage() === p) location.hash = 'dashboard'; return; }
  }
  if (NAME_PAGES.includes(p) && !S.user) {
    await askUserName(true);
    if (!S.user) { toast('กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนทำรายการ', 'err'); if (currentPage() === p) location.hash = 'dashboard'; return; }
  }
  if (currentPage() !== p) return;
  $$('.nav a').forEach(a => a.classList.toggle('active', a.dataset.page === p));
  $('#pageTitle').textContent = PAGES[p][0];
  $('#pageSub').textContent = PAGES[p][1];
  document.title = PAGES[p][0] + ' · Balance Waste';
  $('#sidebar').classList.remove('open');
  const seq = ++routeSeq;
  const reads = routeReads(p);
  const hits = reads.map(rcGet);
  const cached = hits.every(Boolean) ? hits.map(h => h.d) : null;
  const paint = async res => {
    killCharts(); view.onclick = null; S.onFresh = null; S.dirty = false;
    applyBoot(res[0]);
    updateBadge();
    await PAGE_FN[p](res.slice(1));
    countUp();
  };
  killCharts(); view.onclick = null; S.onFresh = null;
  view.innerHTML = loader();
  view.style.animation = 'none'; view.offsetHeight; view.style.animation = '';
  let shown = null;
  if (cached) {
    try { await paint(cached); shown = sigOf(cached); } catch { view.innerHTML = loader(); }
    if (shown && hits.every(h => Date.now() - h.t < RC_FRESH_MS)) return;
  }
  try {
    const res = await readMany(reads);
    if (seq !== routeSeq) return;
    res.forEach((d, i) => rcPut(reads[i], d));
    if (shown === sigOf(res)) { applyBoot(res[0]); return; }
    if (shown !== null && ((S.dirty && INPUT_PAGES.includes(p)) || p === 'logs')) { applyBoot(res[0]); updateBadge(); S.onFresh?.(); return; }
    await paint(res);
  } catch (e) {
    if (seq !== routeSeq) return;
    if (shown !== null) { toast(e.message, 'err'); return; }
    view.innerHTML = `<div class="card empty"><i class="fa-solid fa-plug-circle-xmark"></i>${esc(e.message)}</div>`;
  }
}
addEventListener('hashchange', () => route());
function alertCount() {
  const days = S.limits.storageAlertDays || 90;
  let n = 0, bad = false;
  S.types.forEach(t => {
    const b = balOf(t.id);
    if (b.quota_level === 'over') { n++; bad = true; } else if (b.quota_level === 'warn') n++;
    if (b.age_days != null && b.age_days > days) n++;
  });
  return { n, bad };
}
function updateBadge() {
  const { n, bad } = alertCount();
  const el = $('#alertBadge');
  el.textContent = n ? fmt(n) : '';
  el.classList.toggle('bad', bad);
}

/* ================= DASHBOARD ================= */
function needleGauge(pct) {
  const cx = 150, cy = 150, R = 118, MAX = 120;
  const ang = v => (-90 + Math.min(MAX, Math.max(0, v)) / MAX * 180) * Math.PI / 180;
  const pt = (v, r) => [cx + r * Math.sin(ang(v)), cy - r * Math.cos(ang(v))];
  const arc = (a, b, col) => { const [x1, y1] = pt(a, R), [x2, y2] = pt(b, R); return `<path d="M${x1.toFixed(1)} ${y1.toFixed(1)} A${R} ${R} 0 0 1 ${x2.toFixed(1)} ${y2.toFixed(1)}" stroke="${col}" stroke-width="20" fill="none" stroke-linecap="butt"/>`; };
  let ticks = '';
  for (let v = 0; v <= MAX; v += 10) {
    const [x1, y1] = pt(v, R - 16), [x2, y2] = pt(v, R - (v % 50 ? 22 : 28));
    ticks += `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="#94A3B8" stroke-width="${v % 50 ? 1.5 : 2.5}"/>`;
    if (!(v % 20)) { const [tx, ty] = pt(v, R - 40); ticks += `<text x="${tx.toFixed(1)}" y="${(ty + 4).toFixed(1)}" text-anchor="middle" font-size="11" fill="#64748B" font-family="Kanit">${v}</text>`; }
  }
  const deg = -90 + Math.min(MAX, Math.max(0, pct || 0)) / MAX * 180;
  return `<svg class="needle-gauge" viewBox="0 0 300 172" role="img" aria-label="ใช้โควตา ${fmt(pct, 1)}%">
    <defs><linearGradient id="ng" x1="0" x2="1"><stop offset="0" stop-color="#34D399"/><stop offset="1" stop-color="#059669"/></linearGradient></defs>
    ${arc(0, 80, 'url(#ng)')}${arc(80, 100, '#F59E0B')}${arc(100, 120, '#EF4444')}${ticks}
    <g class="tick-spin"><circle cx="150" cy="150" r="26" fill="none" stroke="#A7F3D0" stroke-width="2" stroke-dasharray="4 6"/></g>
    <g class="needle" data-deg="${deg.toFixed(1)}"><path d="M146 150 L150 44 L154 150 Z" fill="#0F2A23"/><circle cx="150" cy="150" r="5" fill="#F59E0B"/></g>
    <circle cx="150" cy="150" r="12" fill="#0F2A23"/><circle cx="150" cy="150" r="5" fill="#FCD34D"/></svg>`;
}
function ringCard(t) {
  const b = balOf(t.id), p = b.quota_pct == null ? 0 : b.quota_pct;
  return `<div class="ring-card ${b.quota_level}" style="--c:${esc(t.color)}" data-p="${Math.min(100, p)}" data-go="balance" title="${esc(typeLabel(t))} · ${esc(t.receiver)}">
    <div class="qring"><span class="spinner-arc"></span><div class="in"><i class="fa-solid ${esc(t.icon)}"></i><b>${fmt(p, 1)}%</b></div></div>
    <div class="rc-name">${esc(t.name)}</div><div class="rc-sub">${esc(t.code)} · ${fmtT(b.shipped_year_kg)} / ${fmtT(b.permit_kg)} ตัน</div></div>`;
}
function tlItem(m) {
  const t = typeById(m.type_id) || {}, k = KIND[m.kind] || KIND.store;
  return `<div class="tl-item"><div class="ti" style="--c:${k.c}"><i class="fa-solid ${k.icon}"></i></div>
    <div class="tx"><b>${esc(t.name || '—')}</b><small>${k.short} · ${esc(m.doc_no)} · ${dtTH(m.moved_at)} · ${esc(m.actor || '')}</small></div>
    <div class="kg ${k.cls}">${k.sign}${fmtKg(m.weight_kg)} กก.</div></div>`;
}
function alertItem(a) {
  const t = typeById(a.type_id) || {};
  if (a.kind === 'quota') {
    return `<div class="alert ${a.level === 'bad' ? 'bad' : ''}"><div class="ai"><i class="fa-solid ${a.level === 'bad' ? 'fa-circle-exclamation' : 'fa-triangle-exclamation'}"></i></div>
      <p><b>${esc(t.code)} ${esc(t.name)}</b> ${a.level === 'bad' ? 'ส่งกำจัดเกินปริมาณที่ได้รับอนุญาต' : 'ใช้โควตาใกล้เต็ม'} — ${fmt(a.pct, 1)}%<br>
      <small>${fmtKg(a.used_kg)} / ${fmtKg(a.permit_kg)} กก. · ${esc(shortRecv(t.receiver))}</small></p></div>`;
  }
  return `<div class="alert"><div class="ai"><i class="fa-solid fa-hourglass-half"></i></div>
    <p><b>${esc(t.code)} ${esc(t.name)}</b> จัดเก็บในโรงงานนาน ${fmt(a.days)} วัน<br><small>ล็อตเก่าสุดตั้งแต่ ${dateTH(a.since)} · คงเหลือ ${fmtKg(a.balance_kg)} กก.</small></p></div>`;
}
async function dashboard(d) {
  const k = d.kpi;
  const h = new Date().getHours();
  const greet = h < 12 ? ['สวัสดีตอนเช้า', 'fa-sun'] : h < 17 ? ['สวัสดีตอนบ่าย', 'fa-cloud-sun'] : ['สวัสดีตอนเย็น', 'fa-moon'];
  const withBal = S.types.filter(t => balOf(t.id).balance_kg > 0).length;
  const lv = k.quota_pct == null ? 'ok' : k.quota_pct > 100 ? 'over' : k.quota_pct >= (S.limits.quotaWarnPct || 80) ? 'warn' : 'ok';
  const kpis = [
    ['คงเหลือในโรงงานรวม', k.balance_kg, 'กก.', 2, 'fa-warehouse', '#059669', `≈ ${fmtT(k.balance_kg)} ตัน · ${withBal} ประเภท`],
    ['จัดเก็บเดือนนี้', k.month_in_kg, 'กก.', 2, 'fa-box-archive', '#14B8A6', `ทั้งปี ${fmtT(k.year_in_kg)} ตัน`],
    ['ส่งกำจัดปีนี้', k.year_out_kg, 'กก.', 2, 'fa-truck-arrow-right', '#F97316', `เดือนนี้ ${fmtKg(k.month_out_kg)} กก.`],
    ['% โควตาที่ใช้ (ปีนี้)', k.quota_pct || 0, '%', 1, 'fa-chart-pie', lv === 'over' ? '#EF4444' : lv === 'warn' ? '#F59E0B' : '#6366F1', `${fmtT(k.year_out_kg)} / ${fmtT(k.permit_kg)} ตัน`],
  ];
  const withPermit = S.types.filter(t => +t.active === 1 || balOf(t.id).balance_kg > 0);
  const cc = d.containers || {};
  const cTotal = CONT_KEYS.reduce((a, c) => a + (cc[c] || 0), 0);
  view.innerHTML = `
  <div class="reset-topbar"><a class="btn btn-primary btn-sm" href="#store"><i class="fa-solid fa-box-archive"></i>บันทึกจัดเก็บ</a><a class="btn btn-ship btn-sm" href="#ship"><i class="fa-solid fa-truck-arrow-right"></i>ส่งกำจัด</a>${treatOn() ? '<a class="btn btn-treat btn-sm" href="#treat"><i class="fa-solid fa-flask-vial"></i>บำบัดเอง</a>' : ''}<button class="btn btn-bad btn-sm" id="resetBtn">🗑 รีเซ็ตข้อมูล / Reset data</button></div>
  <div class="hero">
    <div class="hero-orb" aria-hidden="true"><span class="o1"></span><span class="o2"></span><span class="o3"><i class="fa-solid fa-recycle"></i></span>
      <span class="sat"><b><i class="fa-solid fa-leaf"></i></b></span><span class="sat s2"><b><i class="fa-solid fa-droplet"></i></b></span><span class="sat s3"><b><i class="fa-solid fa-seedling"></i></b></span></div>
    <div class="greet"><i class="fa-solid ${greet[1]}"></i>${greet[0]}${S.user ? ' คุณ' + esc(S.user) : ''} · ${new Date().toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</div>
    <h2>Balance <em>Waste</em> · ระบบบริหารจัดการของเสีย</h2>
    <p>${esc(S.company.name || '')}</p>
    <div class="chips"><span><i class="fa-solid fa-industry"></i>ทะเบียนโรงงาน ${esc(S.company.factoryReg || '')}</span><span><i class="fa-solid fa-file-shield"></i>${esc(S.company.permitNo || '')} · ปี ${esc(S.company.permitYear || '')}</span><span><i class="fa-solid fa-scale-balanced"></i>อนุญาตรวม ${fmtT(k.permit_kg)} ตัน</span></div>
  </div>

  <div class="grid g-kpi mt stagger">
    ${kpis.map(([l, n, u, dec, ic, c, tr]) => `
      <div class="card kpi" style="--c:${c}">
        <div class="top"><div class="orbit" style="--c:${c}"><span class="dot"></span><span class="dot d2"></span><span class="dot d3"></span><div class="ic"><i class="fa-solid ${ic}"></i></div></div><div class="kpi-label">${l}</div></div>
        <div class="kpi-num"><span data-count="${n}" data-dec="${dec}">0</span><small>${u}</small></div><div class="trend">${tr}</div>
      </div>`).join('')}
  </div>
  ${k.treated_kg !== undefined ? treatSummary(k) : ''}

  <div class="grid g-2 mt">
    <div class="card"><h3><i class="fa-solid fa-chart-column"></i>ปริมาณจัดเก็บ / ส่งกำจัด 12 เดือน<span class="sub">กิโลกรัม</span></h3><div class="chart-box"><canvas id="cMonths"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-gauge-high"></i>โควตารวมปีนี้<span class="sub">เทียบปริมาณที่ได้รับอนุญาต</span></h3>
      <div class="gauge-wrap">${needleGauge(k.quota_pct || 0)}<div class="gauge-val"><span data-count="${k.quota_pct || 0}" data-dec="1">0</span><small>%</small></div>
      <div class="gauge-sub">ส่งกำจัดแล้ว ${fmtKg(k.year_out_kg)} กก. จาก ${fmtKg(k.permit_kg)} กก.<br><span class="tag lv-${lv}">${lvName[lv]}</span></div></div></div>
  </div>

  <div class="sec-h"><i class="fa-solid fa-circle-notch fa-spin" style="color:var(--brand);--fa-animation-duration:3s"></i>การใช้โควตาแยกประเภท <small>ส่งกำจัดปีนี้ / ปริมาณที่ได้รับอนุญาต</small></div>
  <div class="card"><div class="rings stagger">${withPermit.map(ringCard).join('') || '<div class="empty">ยังไม่มีประเภทของเสีย</div>'}</div></div>

  <div class="grid g-2e mt">
    <div class="card"><h3><i class="fa-solid fa-handshake"></i>โควตาตามผู้รับดำเนินการ<span class="sub">ปีนี้</span></h3>
      <div class="recv">${(d.receivers || []).map(r => `<div class="recv-row"><div class="rh"><b>${esc(shortRecv(r.receiver))}</b><em style="color:${r.quota_level === 'over' ? 'var(--bad)' : r.quota_level === 'warn' ? 'var(--amber2)' : 'var(--brand2)'}">${fmt(r.quota_pct, 1)}%</em><span>${fmtT(r.shipped_year_kg)} / ${fmtT(r.permit_kg)} ตัน · ${r.types} ประเภท</span></div>
        <div class="qbar ${r.quota_level}"><span style="width:${Math.min(100, r.quota_pct || 0)}%"></span></div></div>`).join('')}</div></div>
    <div class="card"><h3><i class="fa-solid fa-cubes-stacked"></i>ภาชนะในพื้นที่จัดเก็บ<span class="sub">${fmt(cTotal)} ภาชนะ</span></h3>
      ${cTotal ? '<div class="chart-box sm"><canvas id="cCont"></canvas></div>' : `<div class="empty">${containerSVG('drum200').replace('<svg', '<svg style="width:70px;height:70px"')}<br>ยังไม่มีภาชนะในพื้นที่จัดเก็บ</div>`}</div>
  </div>

  <div class="grid g-2e mt">
    <div class="card"><h3><i class="fa-solid fa-bell"></i>แจ้งเตือน<span class="sub">โควตา ≥ ${S.limits.quotaWarnPct || 80}% · จัดเก็บเกิน ${S.limits.storageAlertDays || 90} วัน</span></h3>
      <div class="alerts">${(d.alerts || []).length ? d.alerts.map(alertItem).join('') : '<div class="all-good"><i class="fa-solid fa-recycle"></i><div><b>ทุกอย่างปกติ</b><br><small>ไม่มีประเภทใดใกล้เต็มโควตาหรือจัดเก็บนานเกินกำหนด</small></div></div>'}</div></div>
    <div class="card"><h3><i class="fa-solid fa-clock-rotate-left"></i>รายการล่าสุด<span class="sub"><a href="#history" style="color:var(--brand2)">ดูทั้งหมด →</a></span></h3>
      <div class="tl">${(d.recent || []).length ? d.recent.map(tlItem).join('') : '<div class="empty"><i class="fa-solid fa-inbox"></i>ยังไม่มีรายการ — เริ่มที่ “บันทึกจัดเก็บ” หรือบันทึก “ยอดยกมา”</div>'}</div></div>
  </div>`;
  view.onclick = e => { const g = e.target.closest('[data-go]'); if (g) location.hash = g.dataset.go; };
  $('#resetBtn').onclick = resetData;
  // animate rings + needle after the first paint
  requestAnimationFrame(() => requestAnimationFrame(() => {
    $$('.ring-card', view).forEach(el => el.style.setProperty('--p', el.dataset.p));
    const n = $('.needle-gauge .needle', view);
    if (n) n.style.transform = `rotate(${n.dataset.deg}deg)`;
  }));
  drawDashCharts(d);
}
/* ของเสียทั้งหมด = บำบัดเอง + ส่งกำจัด + คงเหลือ (สะสม) — แสดงเฉพาะตอนปลดล็อก */
function treatSummary(k) {
  const tot = k.total_in_kg || 0, pc = v => tot > 0 ? Math.max(0, v / tot * 100) : 0;
  const parts = [
    ['บำบัดเอง', k.treated_kg, 'fa-flask-vial', 'var(--treat)', 'treat', `ปีนี้ ${fmtT(k.year_treat_kg)} ตัน · เดือนนี้ ${fmtKg(k.month_treat_kg)} กก.`],
    ['ส่งกำจัด (ตัดโควตา)', k.shipped_kg, 'fa-truck-arrow-right', 'var(--out)', 'ship', `ปีนี้ ${fmtT(k.year_out_kg)} / ${fmtT(k.permit_kg)} ตัน · ${fmt(k.quota_pct || 0, 1)}%`],
    ['คงเหลือในโรงงาน', k.balance_kg, 'fa-warehouse', 'var(--brand)', 'bal', 'รอส่งกำจัด / บำบัด'],
  ];
  return `<div class="card mt tsum">
    <h3><i class="fa-solid fa-flask-vial brew"></i>ของเสียทั้งหมดในระบบ (สะสม)<span class="sub"><i class="fa-solid fa-lock-open"></i> ข้อมูลบำบัดเอง · แท็บนี้เท่านั้น</span></h3>
    <div class="tsum-row">
      <div class="tsum-total"><div class="treat-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-recycle"></i></span></div>
        <div><small>ของเสียทั้งหมด (จัดเก็บ + ยอดยกมา)</small><b><span data-count="${tot}" data-dec="2">0</span> <em>กก.</em></b><small>≈ ${fmtT(tot)} ตัน</small></div></div>
      <div class="tsum-parts">${parts.map(([l, v, ic, c, cls, sub]) => `<div class="tpart ${cls}" style="--c:${c}">
        <div class="tp-h"><span class="tp-ic"><i class="fa-solid ${ic}"></i></span><span class="tp-l">${l}</span><i class="pc">${fmt(pc(v), 1)}%</i></div>
        <b><span data-count="${v}" data-dec="2">0</span> <em>กก.</em></b><small class="tp-s" title="${esc(sub)}">${sub}</small></div>`).join('')}</div>
    </div>
    <div class="tbar" role="img" aria-label="สัดส่วน บำบัดเอง ${fmt(pc(k.treated_kg), 1)}% ส่งกำจัด ${fmt(pc(k.shipped_kg), 1)}% คงเหลือ ${fmt(pc(k.balance_kg), 1)}%">
      <span class="treat" style="width:${pc(k.treated_kg)}%"></span><span class="ship" style="width:${pc(k.shipped_kg)}%"></span><span class="bal" style="width:${pc(k.balance_kg)}%"></span></div>
    <p class="hint" style="margin-top:8px"><i class="fa-solid fa-circle-info"></i> โควตาตามหนังสืออนุญาต สก.2 นับเฉพาะ “ส่งกำจัด” — การบำบัดเองภายในโรงงานไม่ตัดโควตา</p>
  </div>`;
}
function drawDashCharts(d) {
  if (!window.Chart) return;
  Chart.defaults.font.family = 'Prompt, Kanit, sans-serif';
  Chart.defaults.color = '#5B6F68';
  const lbl = d.months.map(m => new Date(m.ym + '-01T00:00:00').toLocaleDateString('th-TH', { month: 'short', year: '2-digit' }));
  const cm = $('#cMonths');
  if (cm) {
    const g = cm.getContext('2d').createLinearGradient(0, 0, 0, 280);
    g.addColorStop(0, '#34D399'); g.addColorStop(1, '#047857');
    const g2 = cm.getContext('2d').createLinearGradient(0, 0, 0, 280);
    g2.addColorStop(0, '#FDBA74'); g2.addColorStop(1, '#EA580C');
    const g3 = c => { const g = c.getContext('2d').createLinearGradient(0, 0, 0, 280); g.addColorStop(0, '#C4B5FD'); g.addColorStop(1, '#6D28D9'); return g; };
    S.charts.push(new Chart(cm, {
      type: 'bar',
      data: { labels: lbl, datasets: [
        { label: 'จัดเก็บ (กก.)', data: d.months.map(m => m.in_kg), backgroundColor: g, borderRadius: 8, maxBarThickness: 26 },
        { label: 'ส่งกำจัด (กก.)', data: d.months.map(m => m.out_kg), backgroundColor: g2, borderRadius: 8, maxBarThickness: 26 },
        ...(d.months[0] && d.months[0].treat_kg !== undefined ? [{ label: 'บำบัดเอง (กก.)', data: d.months.map(m => m.treat_kg), backgroundColor: g3(cm), borderRadius: 8, maxBarThickness: 26 }] : []),
      ] },
      options: { responsive: true, maintainAspectRatio: false, animation: { duration: reduceMotion() ? 0 : 1200, easing: 'easeOutQuart' },
        plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, pointStyle: 'circle' } }, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmtKg(c.parsed.y)}` } } },
        scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: 'rgba(6,78,59,.06)' }, ticks: { callback: v => fmt(v) } } } },
    }));
  }
  const cc = $('#cCont');
  if (cc) {
    S.charts.push(new Chart(cc, {
      type: 'doughnut',
      data: { labels: CONT_KEYS.map(k => CONT[k].name), datasets: [{ data: CONT_KEYS.map(k => d.containers[k] || 0), backgroundColor: ['#0D9488', '#F59E0B', '#10B981'], borderWidth: 4, borderColor: '#fff', hoverOffset: 12 }] },
      options: { responsive: true, maintainAspectRatio: false, cutout: '64%', animation: { animateRotate: !reduceMotion(), duration: reduceMotion() ? 0 : 1600 },
        plugins: { legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'circle', padding: 14 } }, tooltip: { callbacks: { label: c => `${c.label}: ${fmt(c.parsed)} ภาชนะ` } } } },
    }));
  }
}

/* ================= STORE / SHIP FORM ================= */
const meLabel = () => S.user || 'ไม่ระบุ (ตั้งชื่อที่มุมขวาบน)';
function typePicker(sel, mode, q) {
  const qq = q.trim().toLowerCase();
  const list = S.types.filter(t => (mode === 'ship' ? true : +t.active === 1))
    .filter(t => !qq || `${t.code} ${t.name} ${t.receiver} ${t.mgmt_code}`.toLowerCase().includes(qq));
  if (!list.length) return '<div class="empty"><i class="fa-solid fa-magnifying-glass"></i>ไม่พบประเภทของเสีย</div>';
  return receivers().map(r => {
    const ts = list.filter(t => t.receiver === r);
    if (!ts.length) return '';
    return `<div class="rgroup"><h4><i class="fa-solid fa-truck-field"></i>${esc(r)} <small>· ${ts.length} ประเภท</small></h4><div class="tpicks">${ts.map(t => {
      const b = balOf(t.id), off = mode === 'ship' && !(b.balance_kg > 0);
      return `<button type="button" class="tpick ${+t.id === +sel ? 'sel' : ''}" data-type="${t.id}" style="--c:${esc(t.color)}" ${off ? 'disabled' : ''}>
        <span class="ti"><i class="fa-solid ${esc(t.icon)}"></i></span>
        <span class="tt"><b>${esc(t.name)}</b><small><span class="code">${esc(t.code)}</span> · จัดการ ${esc(t.mgmt_code)}</small>
        <span class="bal">${off ? 'ไม่มีคงเหลือ' : 'คงเหลือ ' + fmtKg(b.balance_kg) + ' กก.'}</span></span></button>`;
    }).join('')}</div></div>`;
  }).join('');
}
const TREAT_METHODS = ['ระบบบำบัดน้ำเสียภายในโรงงาน', 'เตาเผาภายในโรงงาน', 'ระบบปรับเสถียร / ทำให้เป็นของแข็ง', 'ระบบคัดแยก / รีไซเคิลภายในโรงงาน'];
function txForm(page) {
  const n = nowParts();
  const ship = page === 'ship', treat = page === 'treat', out = ship || treat; // out: takes stock out of the plant
  const F = { kind: ship ? 'ship' : treat ? 'treat' : 'store', type_id: null, weight: '', container: '', containers: 1, date: n.date, time: n.time, location: '', manifest_no: '', vehicle: '', method: treat ? TREAT_METHODS[0] : '', note: '', q: '' };
  const hashType = +new URLSearchParams(location.hash.split('?')[1] || '').get('type');
  if (hashType && typeById(hashType)) F.type_id = hashType;
  view.innerHTML = `<div class="wizard ${ship ? 'ship-mode' : treat ? 'treat-mode' : ''}">
    <div class="steps">
      ${treat ? `<div class="treat-note"><div class="treat-ic sm" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-flask-vial"></i></span></div>
        <div><b>บำบัดเองภายในโรงงาน</b><small>ตัดยอดคงเหลือในโรงงานเหมือนส่งกำจัด (ไม่เกินยอดคงเหลือ / จำนวนภาชนะ) แต่ <u>ไม่นับ</u> โควตาตามหนังสืออนุญาต สก.2 · แสดงเฉพาะโหมดปลดล็อก</small></div></div>` : ''}
      ${out ? '' : `<div class="seg" id="modeSeg"><button type="button" data-mode="store" class="on"><i class="fa-solid fa-box-archive"></i>บันทึกจัดเก็บ</button><button type="button" data-mode="opening"><i class="fa-solid fa-flag-checkered"></i>ยอดยกมา (Opening balance)</button></div>
        <p class="hint" id="modeHint" style="margin-top:-8px"></p>`}
      <div class="card"><div class="step-h"><span class="n">1</span><h3>เลือกประเภทของเสีย</h3><span class="sub">${S.types.length} ประเภท · แยกตามผู้รับดำเนินการ</span></div>
        <div class="picker-top"><div class="search"><i class="fa-solid fa-magnifying-glass"></i><input id="tQ" placeholder="ค้นหา รหัส / ชื่อ / ผู้รับ เช่น 150102 หรือ หลอดไฟ"></div></div>
        <div id="tPick"></div></div>
      <div class="card"><div class="step-h"><span class="n">2</span><h3>น้ำหนัก &amp; ภาชนะบรรจุ</h3></div>
        <div id="liveBal"></div>
        <div class="field"><label>น้ำหนัก (กิโลกรัม) · ทศนิยมไม่เกิน 2 ตำแหน่ง</label>
          <div class="weight-box" id="wBox"><input id="fW" inputmode="decimal" autocomplete="off" placeholder="0.00"><span>กก.</span></div>
          <div class="weight-help" id="wHelp">&nbsp;</div>
          <div class="quick" id="wQuick">${[5, 10, 50, 100, 500].map(v => `<button type="button" data-add="${v}">+${v}</button>`).join('')}<button type="button" data-add="0">ล้าง</button>${out ? '<button type="button" data-all="1">ทั้งหมดที่คงเหลือ</button>' : ''}</div></div>
        <div class="field mt"><label>ภาชนะบรรจุ</label><div class="ccards" id="cCards"></div></div>
        <div class="field mt"><label>จำนวนภาชนะ</label><div class="stepper"><button type="button" data-step="-1" aria-label="ลด"><i class="fa-solid fa-minus"></i></button><input id="fN" inputmode="numeric" value="1"><button type="button" data-step="1" aria-label="เพิ่ม"><i class="fa-solid fa-plus"></i></button></div><div class="weight-help" id="nHelp"></div></div>
      </div>
      <div class="card"><div class="step-h"><span class="n">3</span><h3>รายละเอียด</h3></div>
        <div class="fields">
          <div class="field"><label>วันที่</label><input type="date" id="fDate" value="${F.date}" max="${n.date}"></div>
          <div class="field"><label>เวลา</label><input type="time" id="fTime" value="${F.time}"></div>
          ${ship ? `<div class="field"><label>เลขที่ใบกำกับการขนส่ง (Manifest) — ไม่บังคับ</label><input id="fMan" maxlength="60" placeholder="เช่น MN-2569-000123"></div>
            <div class="field"><label>ทะเบียนรถ — ไม่บังคับ</label><input id="fVeh" maxlength="40" placeholder="เช่น 70-1234 ชลบุรี"></div>`
            : treat ? `<div class="field full"><label>วิธีบำบัด / ระบบ — ไม่บังคับ</label><input id="fMeth" maxlength="120" list="methList" value="${esc(F.method)}" placeholder="เช่น ระบบบำบัดน้ำเสียภายในโรงงาน"><datalist id="methList">${TREAT_METHODS.map(m => `<option value="${esc(m)}">`).join('')}</datalist></div>`
            : `<div class="field full"><label>สถานที่จัดเก็บ — ไม่บังคับ</label><input id="fLoc" maxlength="100" placeholder="เช่น โรงเก็บของเสียอันตราย A"></div>`}
          <div class="field full"><label>หมายเหตุ</label><textarea id="fNote" rows="2" maxlength="255"></textarea></div>
        </div></div>
    </div>
    <aside class="card summary" id="sumBox"></aside>
  </div>`;

  const cont = () => {
    const b = F.type_id ? balOf(F.type_id) : null;
    $('#cCards').innerHTML = CONT_KEYS.map(k => {
      const av = b ? b.containers[k] || 0 : 0;
      return `<div class="ccard ${F.container === k ? 'sel' : ''}" data-c="${k}" tabindex="0" role="radio" aria-checked="${F.container === k}" aria-label="${CONT[k].name}">
        <div class="face front">${containerSVG(k)}<b>${CONT[k].name}</b><small>${CONT[k].sub}</small>${out || (b && av) ? `<span class="avail">ในคลัง ${fmt(av)}</span>` : ''}</div>
        <div class="face back"><span class="chk"><i class="fa-solid fa-check"></i></span>${containerSVG(k, true)}<b>${CONT[k].name}</b>${out || (b && av) ? `<span class="avail">ในคลัง ${fmt(av)}</span>` : '<small>เลือกแล้ว</small>'}</div></div>`;
    }).join('');
  };
  const pick = () => { $('#tPick').innerHTML = typePicker(F.type_id, out ? 'ship' : 'store', F.q); };
  const errs = () => {
    const out = {};
    const w = parseKg(F.weight);
    if (!F.type_id) out.type = 'กรุณาเลือกประเภทของเสีย';
    if (F.weight === '') out.w = 'กรุณากรอกน้ำหนัก';
    else if (isNaN(w)) out.w = 'น้ำหนักต้องเป็นตัวเลข ทศนิยมไม่เกิน 2 ตำแหน่ง';
    else if (!(w > 0)) out.w = 'น้ำหนักต้องมากกว่า 0';
    if (!F.container) out.c = 'กรุณาเลือกภาชนะ';
    if (!/^\d+$/.test(String(F.containers)) || +F.containers < 1) out.n = 'จำนวนภาชนะต้องเป็นจำนวนเต็มตั้งแต่ 1';
    if (out && F.type_id) {
      const b = balOf(F.type_id);
      if (!out.w && w > b.balance_kg + 1e-9) out.w = `เกินยอดคงเหลือ (${fmtKg(b.balance_kg)} กก.)`;
      if (F.container && !out.n && +F.containers > (b.containers[F.container] || 0)) out.n = `${CONT[F.container].name}ในคลังมี ${fmt(b.containers[F.container] || 0)} ภาชนะ`;
    }
    return out;
  };
  const live = () => {
    const t = typeById(F.type_id), box = $('#liveBal');
    if (!t) { box.innerHTML = ''; return; }
    const b = balOf(t.id), w = parseKg(F.weight) || 0, after = out ? b.balance_kg - w : b.balance_kg + w;
    box.innerHTML = `<div class="live-bal ${out && after < -1e-9 ? 'low' : ''}"><span class="lb-ic"><i class="fa-solid fa-scale-balanced"></i></span>
      <div><small>คงเหลือในโรงงานตอนนี้ · ${esc(t.name)}</small><b>${fmtKg(b.balance_kg)}</b> กก. <small>หลังบันทึก ≈ ${fmtKg(after)} กก. · ${CONT_KEYS.map(k => `${CONT[k].name} ${fmt(b.containers[k] || 0)}`).join(' · ')}</small></div></div>`;
  };
  const summary = () => {
    const t = typeById(F.type_id), e = errs(), w = parseKg(F.weight), k = KIND[F.kind];
    const ok = !Object.keys(e).length;
    $('#sumBox').innerHTML = `${treat ? '<div class="treat-ic sum-ic" aria-hidden="true"><span class="tr"></span><span class="tc"><i class="fa-solid fa-flask-vial"></i></span></div>' : ecoLogo('sum-ic')}
      <h3 style="justify-content:center"><i class="fa-solid ${k.icon}" style="color:${k.c}"></i>${k.name}</h3>
      <div class="sum-total"><b>${isNaN(w) ? '0' : fmtKg(w)} <small style="font-size:16px">กก.</small></b><small>${isNaN(w) ? '' : '≈ ' + fmtT(w) + ' ตัน'}</small></div>
      <div class="sum-list">
        <div><span>ของเสีย</span><b>${t ? esc(t.name) : '—'}</b></div>
        <div><span>รหัสของเสีย</span><b>${t ? esc(t.code) + ' · จัดการ ' + esc(t.mgmt_code) : '—'}</b></div>
        ${treat ? `<div><span>วิธีบำบัด / ระบบ</span><b>${esc(F.method || '—')}</b></div>
        <div><span>โควตา สก.2</span><b style="color:var(--treat)">ไม่นับ (บำบัดภายในโรงงาน)</b></div>`
          : `<div><span>ผู้รับดำเนินการ</span><b>${t ? esc(t.receiver) : '—'}</b></div>`}
        <div><span>ภาชนะ</span><b>${F.container ? CONT[F.container].name + ' × ' + fmt(F.containers) : '—'}</b></div>
        <div><span>วันที่</span><b>${F.date ? dateTH(F.date) + ' ' + esc(F.time) + ' น.' : '—'}</b></div>
        <div><span>ผู้บันทึก</span><b>${esc(meLabel())}</b></div>
      </div>
      <button class="btn btn-lg ${ship ? 'btn-ship' : treat ? 'btn-treat' : 'btn-primary'}" id="fGo" ${ok ? '' : 'disabled'}><i class="fa-solid ${ship ? 'fa-truck-arrow-right' : treat ? 'fa-flask-vial' : 'fa-floppy-disk'}"></i>${ship ? 'ยืนยันส่งกำจัด' : treat ? 'ยืนยันบำบัดเอง' : F.kind === 'opening' ? 'บันทึกยอดยกมา' : 'บันทึกจัดเก็บ'}</button>
      <p class="hint" style="margin-top:10px;text-align:center">${ok ? '<i class="fa-solid fa-circle-check" style="color:var(--ok)"></i> พร้อมบันทึก' : esc(Object.values(e)[0])}</p>`;
    $('#fGo').onclick = submit;
    const we = e.w && F.weight !== '';
    $('#wBox').classList.toggle('bad', !!we);
    $('#wHelp').textContent = we ? e.w : (out && t ? `${treat ? 'บำบัด' : 'ส่ง'}ได้ไม่เกิน ${fmtKg(balOf(t.id).balance_kg)} กก.` : '\u00a0');
    $('#wHelp').classList.toggle('bad', !!we);
    $('#nHelp').textContent = e.n || '';
    $('#nHelp').className = 'weight-help' + (e.n ? ' bad' : '');
  };
  const refresh = () => { live(); summary(); };
  const setMode = m => {
    F.kind = m;
    $$('#modeSeg button').forEach(b => b.classList.toggle('on', b.dataset.mode === m));
    $('#modeHint').innerHTML = m === 'opening' ? '<i class="fa-solid fa-circle-info"></i> ยอดยกมา: ของเสียที่จัดเก็บอยู่ในโรงงานก่อนเริ่มใช้ระบบ — บันทึกตามที่ตรวจนับจริง (นับรวมในยอดคงเหลือ ไม่นับเป็นการจัดเก็บประจำเดือน)' : '<i class="fa-solid fa-circle-info"></i> บันทึกของเสียที่นำเข้าพื้นที่จัดเก็บของโรงงาน';
    refresh();
  };
  async function submit() {
    const go = $('#fGo');
    if (!go || go.disabled) return;
    const e = errs();
    if (Object.keys(e).length) { toast(Object.values(e)[0], 'err'); return; }
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังบันทึก...';
    const body = { type_id: F.type_id, weight_kg: String(parseKg(F.weight)), container: F.container, containers: String(F.containers), date: F.date, time: F.time, note: F.note };
    if (ship) Object.assign(body, { manifest_no: F.manifest_no, vehicle: F.vehicle });
    else if (treat) body.method = F.method;
    else body.location = F.location;
    try {
      const r = await writeApi(F.kind, body);
      const t = typeById(F.type_id);
      confetti();
      const m = modal(`<div class="success-ic ${ship ? 'ship' : treat ? 'treat' : ''}"><span class="sr"></span><span class="sc"><i class="fa-solid ${ship ? 'fa-truck-arrow-right' : treat ? 'fa-arrows-spin' : 'fa-check'}"></i></span></div>
        <h3 style="justify-content:center">${KIND[F.kind].name}สำเร็จ</h3>
        <div class="doc">${esc(r.doc_no)}</div>
        <div class="receipt"><div class="sum-list">
          <div><span>ของเสีย</span><b>${esc(typeLabel(t))}</b></div>
          <div><span>น้ำหนัก</span><b>${fmtKg(body.weight_kg)} กก.</b></div>
          <div><span>ภาชนะ</span><b>${CONT[F.container].name} × ${fmt(F.containers)}</b></div>
          ${treat && F.method ? `<div><span>วิธีบำบัด / ระบบ</span><b>${esc(F.method)}</b></div>` : ''}
          <div><span>คงเหลือในโรงงาน</span><b>${fmtKg(r.balance_kg)} กก.</b></div>
          <div><span>ผู้บันทึก</span><b>${esc(S.user)}</b></div></div></div>
        <div class="modal-acts"><a class="btn btn-ghost" href="#balance" data-close><i class="fa-solid fa-scale-balanced"></i>ดูคงเหลือ</a><button class="btn btn-primary" data-close><i class="fa-solid fa-plus"></i>ทำรายการต่อ</button></div>`);
      S.onModalClose = () => route();
      m.querySelector('[data-close].btn-primary')?.focus();
    } catch (err) {
      toast(err.message, 'err');
      if (err.code === 'TREAT_LOCKED') { location.hash = 'dashboard'; return; }
      summary();
    }
  }

  // events
  view.onclick = e => {
    const tp = e.target.closest('[data-type]');
    if (tp && !tp.disabled) { F.type_id = +tp.dataset.type; $$('.tpick', view).forEach(b => b.classList.toggle('sel', b === tp)); cont(); refresh(); return; }
    const cc = e.target.closest('[data-c]');
    if (cc) { F.container = cc.dataset.c; $$('.ccard', view).forEach(c => { const on = c === cc; c.classList.toggle('sel', on); c.setAttribute('aria-checked', on); }); refresh(); return; }
    const md = e.target.closest('[data-mode]');
    if (md) { setMode(md.dataset.mode); return; }
    const ad = e.target.closest('[data-add]');
    if (ad) {
      const v = +ad.dataset.add, cur = parseKg(F.weight) || 0;
      F.weight = v ? String(Math.round((cur + v) * 100) / 100) : '';
      $('#fW').value = F.weight; refresh(); return;
    }
    if (e.target.closest('[data-all]') && F.type_id) { F.weight = String(balOf(F.type_id).balance_kg); $('#fW').value = F.weight; refresh(); return; }
    const st = e.target.closest('[data-step]');
    if (st) { F.containers = Math.max(1, (parseInt(F.containers, 10) || 0) + +st.dataset.step); $('#fN').value = F.containers; refresh(); }
  };
  view.onkeydown = e => { const cc = e.target.closest('.ccard'); if (cc && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); cc.click(); } };
  $('#tQ').oninput = e => { F.q = e.target.value; pick(); };
  $('#fW').oninput = e => { F.weight = e.target.value.trim(); refresh(); };
  $('#fN').oninput = e => { F.containers = e.target.value.trim(); refresh(); };
  $('#fDate').onchange = e => { F.date = e.target.value; summary(); };
  $('#fTime').onchange = e => { F.time = e.target.value; summary(); };
  $('#fNote').oninput = e => { F.note = e.target.value; };
  if (ship) { $('#fMan').oninput = e => { F.manifest_no = e.target.value; }; $('#fVeh').oninput = e => { F.vehicle = e.target.value; }; }
  else if (treat) $('#fMeth').oninput = e => { F.method = e.target.value; summary(); };
  else $('#fLoc').oninput = e => { F.location = e.target.value; };
  // fresh data arrived while the user is typing: update balances only
  S.onFresh = () => { pick(); cont(); refresh(); };
  pick(); cont();
  if (out) refresh(); else setMode('store');
}

/* ================= BALANCE ================= */
function balance() {
  const days = S.limits.storageAlertDays || 90;
  // per-type treatment totals: only in the unlocked view (the server sends them only then)
  const showT = treatOn() && S.types.some(t => balOf(t.id).treated_kg !== undefined);
  let filter = 'all';
  const recv = receivers().map(r => {
    const ts = S.types.filter(t => t.receiver === r);
    const sum = k => ts.reduce((a, t) => a + Math.round((balOf(t.id)[k] || 0) * 100), 0) / 100;
    const permit = sum('permit_kg'), used = sum('shipped_year_kg');
    const pct = permit > 0 ? Math.round(used / permit * 1000) / 10 : null;
    return { r, ts, permit, used, bal: sum('balance_kg'), pct, lv: pct == null ? 'ok' : pct > 100 ? 'over' : pct >= (S.limits.quotaWarnPct || 80) ? 'warn' : 'ok' };
  });
  const totalBal = S.types.reduce((a, t) => a + Math.round(balOf(t.id).balance_kg * 100), 0) / 100;
  const cardOf = t => {
    const b = balOf(t.id), pct = b.quota_pct;
    return `<div class="bcard" style="--c:${esc(t.color)}">
      <div class="bh"><span class="bi"><i class="fa-solid ${esc(t.icon)}"></i></span><div class="bt"><b>${esc(t.name)} ${+t.active === 1 ? '' : '<span class="tag off">ปิดใช้งาน</span>'}</b><small>${esc(t.code)} · จัดการ ${esc(t.mgmt_code)} · ${esc(shortRecv(t.receiver))}</small></div></div>
      <div class="big"><b data-count="${b.balance_kg}" data-dec="2">0</b><span>กก. คงเหลือ</span></div>
      <div class="flow"><span><i class="fa-solid fa-arrow-down" style="color:var(--in)"></i> รับเข้า <b>${fmtKg(b.in_kg)}</b></span><span><i class="fa-solid fa-arrow-up" style="color:var(--out)"></i> ส่งกำจัด <b>${fmtKg(b.shipped_kg)}</b></span>${showT && b.treated_kg > 0 ? `<span class="ft"><i class="fa-solid fa-flask-vial"></i> บำบัดเอง <b>${fmtKg(b.treated_kg)}</b></span>` : ''}${b.opening_kg ? `<span>(ยกมา ${fmtKg(b.opening_kg)})</span>` : ''}</div>
      <div class="cchips">${CONT_KEYS.map(k => `<span class="cchip ${b.containers[k] ? '' : 'zero'}">${containerSVG(k)}${CONT[k].name} <b>${fmt(b.containers[k])}</b></span>`).join('')}</div>
      <div class="qline"><span>ส่งกำจัดปีนี้ <b>${fmtKg(b.shipped_year_kg)}</b> / ${fmtKg(b.permit_kg)} กก.</span><span class="tag lv-${b.quota_level}">${pct == null ? '—' : fmt(pct, 1) + '%'}</span></div>
      <div class="qbar ${b.quota_level}"><span style="width:${Math.min(100, pct || 0)}%"></span></div>
      ${b.age_days != null ? `<span class="agepill ${b.age_days > days ? 'old' : ''}"><i class="fa-solid fa-hourglass-half"></i>ล็อตเก่าสุด ${fmt(b.age_days)} วัน (ตั้งแต่ ${dateTH(b.oldest_at)})</span>` : ''}
      <div class="rec-acts" style="margin-top:10px;justify-content:flex-end"><a class="btn btn-sm btn-ghost" href="#store?type=${t.id}"><i class="fa-solid fa-box-archive"></i>จัดเก็บ</a>${b.balance_kg > 0 ? `<a class="btn btn-sm btn-ship" href="#ship?type=${t.id}"><i class="fa-solid fa-truck-arrow-right"></i>ส่งกำจัด</a>` : ''}${showT && b.balance_kg > 0 ? `<a class="btn btn-sm btn-treat" href="#treat?type=${t.id}"><i class="fa-solid fa-flask-vial"></i>บำบัด</a>` : ''}</div>
    </div>`;
  };
  const show = () => {
    const list = S.types.filter(t => {
      const b = balOf(t.id);
      if (filter === 'has') return b.balance_kg > 0;
      if (filter === 'alert') return b.quota_level !== 'ok' || (b.age_days != null && b.age_days > days);
      return +t.active === 1 || b.balance_kg > 0 || b.in_kg > 0;
    });
    $('#balGrid').innerHTML = list.length ? list.map(cardOf).join('') : '<div class="empty"><i class="fa-solid fa-leaf"></i>ไม่มีรายการตามตัวกรอง</div>';
    countUp($('#balGrid'));
  };
  view.innerHTML = `
    <div class="recv-cards stagger">${recv.map(x => `<div class="rcard"><h4><i class="fa-solid fa-truck-field"></i>${esc(x.r)}</h4>
      <div style="display:flex;align-items:baseline;gap:8px"><span class="pct" style="color:${x.lv === 'over' ? 'var(--bad)' : x.lv === 'warn' ? 'var(--amber2)' : 'var(--brand2)'}">${x.pct == null ? '—' : fmt(x.pct, 1) + '%'}</span><span class="hint">ใช้โควตาปีนี้</span></div>
      <div class="qbar ${x.lv}"><span style="width:${Math.min(100, x.pct || 0)}%"></span></div>
      <div class="nums"><span>ส่งแล้ว ${fmtT(x.used)} / ${fmtT(x.permit)} ตัน</span><span>คงเหลือในโรงงาน ${fmtKg(x.bal)} กก.</span></div></div>`).join('')}</div>
    <div class="card mt">
      <h3><i class="fa-solid fa-scale-balanced"></i>คงเหลือในโรงงานรวม <span style="font-family:Kanit;color:var(--brand2)">${fmtKg(totalBal)} กก.</span><span class="sub">≈ ${fmtT(totalBal)} ตัน</span></h3>
      <div class="seg" id="balSeg"><button data-f="all" class="on">ทั้งหมด</button><button data-f="has">มีคงเหลือ</button><button data-f="alert">ต้องติดตาม (${alertCount().n})</button></div>
    </div>
    <div class="bal-grid mt" id="balGrid"></div>
    <div class="card mt"><h3><i class="fa-solid fa-table"></i>ตารางสรุปตามหนังสืออนุญาต<span class="sub">ปี ${esc(S.company.permitYear || '')} · หน่วย กก.</span></h3>
      <div class="tbl-wrap"><table><thead><tr><th>#</th><th>รหัส</th><th>ของเสีย</th><th>จัดการ</th><th>ผู้รับดำเนินการ</th><th class="num">อนุญาต (ตัน)</th><th class="num">รับเข้าสะสม</th><th class="num">ส่งกำจัดสะสม</th>${showT ? '<th class="num th-treat">บำบัดเองสะสม</th>' : ''}<th class="num">คงเหลือ</th><th class="num">ส่งปีนี้</th><th class="num">% โควตา</th></tr></thead><tbody>
      ${S.types.map((t, i) => { const b = balOf(t.id); return `<tr><td>${i + 1}</td><td><span class="tcode" style="--c:${esc(t.color)}"><i class="fa-solid ${esc(t.icon)}"></i>${esc(t.code)}</span></td><td>${esc(t.name)}</td><td>${esc(t.mgmt_code)}</td><td>${esc(shortRecv(t.receiver))}</td>
        <td class="num">${fmt(t.permit_tons, 3)}</td><td class="num">${fmtKg(b.in_kg)}</td><td class="num">${fmtKg(b.shipped_kg)}</td>${showT ? `<td class="num" style="color:var(--treat)">${fmtKg(b.treated_kg)}</td>` : ''}<td class="num"><b>${fmtKg(b.balance_kg)}</b></td><td class="num">${fmtKg(b.shipped_year_kg)}</td><td class="num"><span class="tag lv-${b.quota_level}">${b.quota_pct == null ? '—' : fmt(b.quota_pct, 1) + '%'}</span></td></tr>`; }).join('')}
      </tbody></table></div></div>`;
  view.onclick = e => { const f = e.target.closest('[data-f]'); if (f) { filter = f.dataset.f; $$('#balSeg button').forEach(b => b.classList.toggle('on', b === f)); show(); } };
  show();
}

/* ================= HISTORY ================= */
const typeOptions = (sel, filter = () => true) => receivers().map(r => {
  const ts = S.types.filter(t => t.receiver === r && (filter(t) || +t.id === +sel));
  return ts.length ? `<optgroup label="${esc(r)}">${ts.map(t => `<option value="${t.id}" ${+t.id === +sel ? 'selected' : ''}>${esc(t.code)} · ${esc(t.name)}</option>`).join('')}</optgroup>` : '';
}).join('');
function history(initial) {
  const n = nowParts();
  view.innerHTML = `<div class="card">
    <div class="filters">
      <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="hFrom" value="${monthFirst()}"></div>
      <div class="field"><label>ถึงวันที่</label><input type="date" id="hTo" value="${n.date}"></div>
      <div class="field"><label>รายการ</label><select id="hKind"><option value="">ทั้งหมด</option>${kinds().map(([k, a]) => `<option value="${k}">${a.name}</option>`).join('')}</select></div>
      <div class="field"><label>ประเภทของเสีย</label><select id="hType"><option value="">ทั้งหมด</option>${typeOptions('')}</select></div>
      <div class="field"><label>ผู้รับดำเนินการ</label><select id="hRecv"><option value="">ทั้งหมด</option>${receivers().map(r => `<option>${esc(r)}</option>`).join('')}</select></div>
      <div class="field"><label>ค้นหา</label><input id="hQ" placeholder="เลขที่ / ผู้บันทึก / ใบกำกับ / หมายเหตุ"></div>
      <button class="btn btn-primary" id="hGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
      <button class="btn btn-amber" id="hCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
    </div>
    <div id="hRes">${loader()}</div></div>`;
  const params = () => '&' + new URLSearchParams({ from: $('#hFrom').value, to: $('#hTo').value, type: $('#hType').value, kind: $('#hKind').value, receiver: $('#hRecv').value, q: $('#hQ').value });
  let items = [], loadSeq = 0;
  const acts = m => `<div class="rec-acts"><button class="btn btn-sm btn-ghost rec-btn" data-medit="${m.id}" title="แก้ไขรายการ">✏️ แก้ไข</button><button class="btn btn-sm btn-ghost rec-btn" data-mdel="${m.id}" title="ลบรายการ">🗑 ลบ</button></div>`;
  const show = () => {
    const tot = k => items.filter(i => i.kind === k).reduce((a, i) => a + Math.round(i.weight_kg * 100), 0) / 100;
    $('#hRes').innerHTML = items.length ? `
      <div class="chips" style="margin-bottom:12px">${kinds().map(([k, a]) => `<span class="tag k-${k}"><i class="fa-solid ${a.icon}"></i>${a.short} ${fmtKg(tot(k))} กก.</span>`).join('')}<span class="tag off">${fmt(items.length)} รายการ</span></div>
      <div class="tbl-wrap tbl-desktop" style="max-height:64vh"><table><thead><tr><th>เลขที่เอกสาร</th><th>วันที่ / เวลา</th><th>รายการ</th><th>ของเสีย</th><th>ผู้รับดำเนินการ</th><th class="num">น้ำหนัก (กก.)</th><th>ภาชนะ</th><th>ที่เก็บ / ใบกำกับ</th><th>หมายเหตุ</th><th>ผู้บันทึก</th><th>จัดการ</th></tr></thead><tbody>
      ${items.map(m => { const k = KIND[m.kind] || KIND.store; return `<tr><td><b>${esc(m.doc_no)}</b></td><td style="white-space:nowrap">${dtTH(m.moved_at)}</td>
        <td><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.short}</span></td>
        <td><span class="tcode" style="--c:${esc(m.color)}"><i class="fa-solid ${esc(m.icon)}"></i><span><b>${esc(m.code)}</b> ${esc(m.type_name)}</span></span></td>
        <td style="font-size:12.5px">${m.kind === 'treat' ? '<span style="color:var(--treat)">บำบัดภายในโรงงาน</span>' : esc(shortRecv(m.receiver))}</td>
        <td class="num"><b class="kg ${k.cls}">${k.sign}${fmtKg(m.weight_kg)}</b></td>
        <td style="white-space:nowrap">${CONT[m.container]?.name || esc(m.container)} × ${fmt(m.containers)}</td>
        <td style="font-size:12.5px">${esc([m.location, m.manifest_no && 'ใบกำกับ ' + m.manifest_no, m.vehicle && 'รถ ' + m.vehicle, m.method && 'วิธี ' + m.method].filter(Boolean).join(' · ') || '—')}</td>
        <td style="max-width:220px;overflow:hidden;text-overflow:ellipsis">${esc(m.note || '')}</td>
        <td class="by-cell">${m.actor ? `<i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ${esc(m.actor)}` : '—'}</td><td>${acts(m)}</td></tr>`; }).join('')}
      </tbody></table></div>
      <div class="hcards">${items.map(m => { const k = KIND[m.kind] || KIND.store; return `<div class="hcard"><div class="r1"><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.short}</span><b>${esc(m.doc_no)}</b><span class="kg ${k.cls}">${k.sign}${fmtKg(m.weight_kg)} กก.</span></div>
        <p><span class="tcode" style="--c:${esc(m.color)}"><i class="fa-solid ${esc(m.icon)}"></i>${esc(m.code)}</span> ${esc(m.type_name)}</p>
        <small>${dtTH(m.moved_at)} · ${CONT[m.container]?.name || ''} × ${fmt(m.containers)} · ${esc(m.actor || '')}${m.manifest_no ? ' · ใบกำกับ ' + esc(m.manifest_no) : ''}${m.method ? ' · ' + esc(m.method) : ''}${m.note ? ' · ' + esc(m.note) : ''}</small>${acts(m)}</div>`; }).join('')}</div>`
      : '<div class="empty"><i class="fa-solid fa-magnifying-glass"></i>ไม่พบรายการตามเงื่อนไข</div>';
  };
  const load = async () => {
    const seq = ++loadSeq, box = $('#hRes');
    box.innerHTML = loader('กำลังค้นหา...');
    try {
      const r = await api('history', null, params());
      if (seq !== loadSeq || !box.isConnected) return;
      items = r.items; show();
    } catch (e) {
      if (seq !== loadSeq || !box.isConnected) return;
      box.innerHTML = `<div class="empty"><i class="fa-solid fa-plug-circle-xmark"></i>${esc(e.message)}</div>`;
    }
  };
  view.onclick = e => {
    const ed = e.target.closest('[data-medit]'), dl = e.target.closest('[data-mdel]');
    const m = ed || dl ? items.find(i => +i.id === +(ed || dl).dataset[ed ? 'medit' : 'mdel']) : null;
    if (!m) return;
    if (ed) editMovement(m); else deleteMovement(m);
  };
  $('#hGo').onclick = load;
  $('#hQ').onkeydown = e => { if (e.key === 'Enter') load(); };
  $('#hCsv').onclick = () => downloadCsv('export', params());
  if (initial) { items = initial.items; show(); } else load();
}

/* ================= EDIT / DELETE (รหัสผ่านผู้ดูแล) =================
   เซิร์ฟเวอร์ตรวจย้อนทุกรายการตามลำดับเวลา — ถ้ายอดใดจะติดลบจะปฏิเสธและแจ้งเหตุผล */
function recordForm(html, onSave) {
  const m = modal(html + `<div id="eErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px;white-space:pre-line"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="eSave"><i class="fa-solid fa-floppy-disk"></i>บันทึก</button></div>`, true);
  const go = $('#eSave', m), err = $('#eErr', m);
  go.onclick = async () => {
    if (go.disabled) return;
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังบันทึก...'; err.textContent = '';
    try {
      const r = await onSave(m);
      closeModal();
      toast(r.message, r.changed === false ? 'info' : 'ok');
      if (r.changed !== false) route();
    } catch (e) {
      err.textContent = e.message; toast(e.message, 'err');
      go.disabled = false; go.innerHTML = '<i class="fa-solid fa-floppy-disk"></i>บันทึก';
    }
  };
  return m;
}
async function editMovement(m) {
  try { await askPassword(); } catch (e) { toast(e.message, 'err'); return; }
  const k = KIND[m.kind] || KIND.store, ship = m.kind === 'ship', treat = m.kind === 'treat';
  recordForm(`<h3><i class="fa-solid fa-pen-to-square" style="color:var(--brand)"></i> แก้ไขรายการ ${esc(m.doc_no)}</h3>
    <p><span class="tag k-${m.kind}"><i class="fa-solid ${k.icon}"></i>${k.name}</span></p>
    <p class="hint" style="margin-top:6px"><i class="fa-solid fa-circle-info"></i> เปลี่ยนชนิดรายการไม่ได้ — หากต้องการให้ลบแล้วบันทึกใหม่ · ระบบตรวจยอดคงเหลือย้อนหลังทุกรายการ ไม่ให้ติดลบ</p>
    <p class="hint"><i class="fa-solid fa-user-check" style="color:var(--ok)"></i> บันทึกโดย: <b>${esc(m.actor || 'ไม่ระบุ')}</b> · การแก้ไขนี้จะบันทึกใน Log ในนาม <b>${esc(meLabel())}</b></p>
    <div class="fields mt">
      <div class="field full"><label>ประเภทของเสีย</label><select id="eType">${typeOptions(m.type_id, t => ship || treat || +t.active === 1)}</select></div>
      <div class="field"><label>น้ำหนัก (กก.)</label><input id="eW" inputmode="decimal" value="${+m.weight_kg}"></div>
      <div class="field"><label>ภาชนะ</label><select id="eC">${CONT_KEYS.map(c => `<option value="${c}" ${c === m.container ? 'selected' : ''}>${CONT[c].name}</option>`).join('')}</select></div>
      <div class="field"><label>จำนวนภาชนะ</label><input id="eN" inputmode="numeric" value="${+m.containers}"></div>
      <div class="field"><label>วันที่</label><input type="date" id="eDate" value="${m.moved_at.slice(0, 10)}"></div>
      <div class="field"><label>เวลา</label><input type="time" id="eTime" value="${m.moved_at.slice(11, 16)}"></div>
      ${ship ? `<div class="field"><label>เลขที่ใบกำกับการขนส่ง</label><input id="eMan" maxlength="60" value="${esc(m.manifest_no)}"></div><div class="field full"><label>ทะเบียนรถ</label><input id="eVeh" maxlength="40" value="${esc(m.vehicle)}"></div>`
        : treat ? `<div class="field full"><label>วิธีบำบัด / ระบบ</label><input id="eMeth" maxlength="120" list="methList" value="${esc(m.method)}"><datalist id="methList">${TREAT_METHODS.map(x => `<option value="${esc(x)}">`).join('')}</datalist></div>`
        : `<div class="field full"><label>สถานที่จัดเก็บ</label><input id="eLoc" maxlength="100" value="${esc(m.location)}"></div>`}
      <div class="field full"><label>หมายเหตุ</label><input id="eNote" maxlength="255" value="${esc(m.note)}"></div>
    </div>`, f => {
    const body = { id: m.id, doc_no: m.doc_no, type_id: $('#eType', f).value, weight_kg: $('#eW', f).value, container: $('#eC', f).value, containers: $('#eN', f).value, date: $('#eDate', f).value, time: $('#eTime', f).value, note: $('#eNote', f).value };
    if (ship) Object.assign(body, { manifest_no: $('#eMan', f).value, vehicle: $('#eVeh', f).value });
    else if (treat) body.method = $('#eMeth', f).value;
    else body.location = $('#eLoc', f).value;
    return recordApi('movement_update', body);
  });
}
async function deleteMovement(m) {
  const k = KIND[m.kind]?.name || m.kind;
  const msg = [`ลบรายการ ${m.doc_no}?`, '', `${k} ${m.code} ${m.type_name} ${fmtKg(m.weight_kg)} กก. · ${CONT[m.container]?.name || ''} × ${m.containers} · ${dtTH(m.moved_at)}`, '',
    'ยอดคงเหลือจะคำนวณใหม่ (ระบบไม่ยอมให้ยอดย้อนหลังติดลบ) · การลบย้อนกลับไม่ได้ ต้องการลบหรือไม่?'];
  try {
    await askPassword();
    if (!confirm(msg.join('\n'))) return;
    const r = await recordApi('movement_delete', { id: m.id, doc_no: m.doc_no });
    toast(r.message, 'info');
    route();
  } catch (e) { toast(e.message, 'err'); }
}

/* ================= LOGS ================= */
const LOGCAT = { waste: ['ของเสีย', 'fa-recycle', '#059669'], edit: ['แก้ไข / ลบ', 'fa-pen-to-square', '#8B5CF6'], setting: ['ตั้งค่า', 'fa-gear', '#0F2A23'], warn: ['ถูกปฏิเสธ', 'fa-ban', '#F59E0B'] };
const LOGACT = { store: 'fa-box-archive', ship: 'fa-truck-arrow-right', opening: 'fa-flag-checkered', treat: 'fa-flask-vial', treat_unlock: 'fa-lock-open', movement_update: 'fa-pen-to-square', movement_delete: 'fa-trash-can', type_save: 'fa-gear', reset_data: 'fa-broom' };
async function logs() {
  const n = nowParts();
  let cat = '', items = null, loadedWith = '';
  view.innerHTML = `
    <div class="log-kpi" id="logKpi"></div>
    <div class="card">
      <div class="filters">
        <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="lFrom" value="${monthFirst()}"></div>
        <div class="field"><label>ถึงวันที่</label><input type="date" id="lTo" value="${n.date}"></div>
        <div class="field" style="flex:2;min-width:200px"><label>ค้นหา</label><input id="lQ" placeholder="เลขที่เอกสาร / ผู้ทำรายการ / รายละเอียด"></div>
        <button class="btn btn-primary" id="lGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
        <button class="btn btn-amber" id="lCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
      </div>
      <div id="lRes">${loader()}</div>
    </div>`;
  const params = (withCat = true) => '&' + new URLSearchParams({ from: $('#lFrom').value, to: $('#lTo').value, q: $('#lQ').value, cat: withCat ? cat : '' });
  const show = () => {
    const cnt = k => items.filter(i => i.category === k).length;
    $('#logKpi').innerHTML = [['', 'ทั้งหมด', 'fa-list', '#0F2A23', items.length], ...Object.entries(LOGCAT).map(([k, [nm, ic, c]]) => [k, nm, ic, c, cnt(k)])]
      .map(([k, nm, ic, c, v]) => `<div class="lk ${k === cat ? 'on' : ''}" data-cat="${k}" style="--c:${c}"><b>${fmt(v)}</b><span><i class="fa-solid ${ic}"></i>${nm}</span></div>`).join('');
    const list = cat ? items.filter(i => i.category === cat) : items;
    if (!list.length) { $('#lRes').innerHTML = '<div class="empty"><i class="fa-solid fa-clipboard"></i>ไม่พบบันทึกประวัติตามเงื่อนไข</div>'; return; }
    const groups = {};
    list.forEach(i => (groups[i.created_at.slice(0, 10)] ||= []).push(i));
    let k = 0;
    $('#lRes').innerHTML = Object.entries(groups).map(([day, arr]) => `
      <div class="logday"><h4><i class="fa-regular fa-calendar"></i>${new Date(day + 'T00:00:00').toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })} · ${arr.length} รายการ</h4>
        <div class="logline">${arr.map(i => {
          const cc = LOGCAT[i.category] || ['', 'fa-circle'];
          return `<div class="logitem cat-${esc(i.category)}" style="animation-delay:${Math.min(k++, 20) * .03}s">
            <div class="ai"><i class="fa-solid ${LOGACT[i.action] || cc[1]}"></i></div>
            <div class="tx"><p>${esc(i.detail)}</p>
              <div class="mt2"><span><i class="fa-solid fa-tag"></i> ${cc[0]}</span>${i.ref ? `<span><i class="fa-solid fa-hashtag"></i> ${esc(i.ref)}</span>` : ''}
              <span><i class="fa-solid fa-user-check"></i> ผู้ทำรายการ / By: <b>${i.actor ? esc(i.actor) : 'ไม่ระบุ'}</b></span></div></div>
            <div class="tm">${i.created_at.slice(11, 16)} น.</div></div>`;
        }).join('')}</div></div>`).join('');
  };
  const load = async () => {
    $('#lRes').innerHTML = loader();
    try {
      const q = params(false);
      ({ items } = await logsApi('logs', q));
      loadedWith = q;
      show();
    } catch (e) {
      $('#lRes').innerHTML = `<div class="empty"><i class="fa-solid fa-lock"></i>${esc(e.message)}</div>`;
    }
  };
  view.onclick = e => { const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; if (items && params(false) === loadedWith) show(); else load(); } };
  $('#lGo').onclick = load;
  $('#lQ').onkeydown = e => { if (e.key === 'Enter') load(); };
  $('#lCsv').onclick = () => downloadCsv('logs_export', params());
  load();
}

/* ================= SETTINGS (waste types) ================= */
const ICONS = ['fa-recycle', 'fa-box-open', 'fa-bottle-water', 'fa-gears', 'fa-spray-can', 'fa-fill-drip', 'fa-biohazard', 'fa-lightbulb', 'fa-microchip', 'fa-jug-detergent', 'fa-note-sticky', 'fa-water', 'fa-faucet-drip', 'fa-oil-can', 'fa-flask', 'fa-battery-half', 'fa-dumpster', 'fa-trash-can', 'fa-leaf', 'fa-skull-crossbones', 'fa-fire', 'fa-droplet', 'fa-box'];
const COLORS = ['#059669', '#14B8A6', '#0EA5E9', '#6366F1', '#A855F7', '#EC4899', '#EF4444', '#F97316', '#F59E0B', '#EAB308', '#B7791F', '#8B5E3C', '#64748B', '#0F2A23'];
function settings() {
  view.innerHTML = `<div class="reset-topbar"><button class="btn btn-primary" id="tAdd"><i class="fa-solid fa-plus"></i>เพิ่มประเภทของเสีย</button><button class="btn btn-bad" id="resetBtn">🗑 รีเซ็ตข้อมูล / Reset data</button></div>
    <div class="card"><h3><i class="fa-solid fa-list-check"></i>ประเภทของเสียตามหนังสืออนุญาต<span class="sub">${S.types.length} รายการ · อนุญาตรวม ${fmt(S.types.filter(t => +t.active === 1).reduce((a, t) => a + Number(t.permit_tons), 0), 3)} ตัน</span></h3>
      <div class="tlist stagger">${S.types.map(t => `<div class="titem ${+t.active === 1 ? '' : 'off'}" style="--c:${esc(t.color)}"><span class="ti"><i class="fa-solid ${esc(t.icon)}"></i></span>
        <div class="tx"><b>${esc(t.code)} · ${esc(t.name)} ${+t.active === 1 ? '' : '<span class="tag off">ปิดใช้งาน</span>'}</b><small>อนุญาต ${fmt(t.permit_tons, 3)} ตัน · รหัสจัดการ ${esc(t.mgmt_code)} · ${esc(t.receiver)}${t.receiver_reg ? ' (ทะเบียน ' + esc(t.receiver_reg) + ')' : ''}</small></div>
        <button class="btn btn-sm btn-ghost" data-tedit="${t.id}"><i class="fa-solid fa-pen"></i>แก้ไข</button></div>`).join('')}</div></div>`;
  view.onclick = e => { const b = e.target.closest('[data-tedit]'); if (b) typeForm(typeById(b.dataset.tedit)); };
  $('#tAdd').onclick = () => typeForm(null);
  $('#resetBtn').onclick = resetData;
}
function typeForm(t) {
  let ic = t?.icon || ICONS[0], col = t?.color || COLORS[0];
  const recvList = receivers();
  const m = recordForm(`<h3><i class="fa-solid ${t ? 'fa-pen-to-square' : 'fa-plus'}" style="color:var(--brand)"></i> ${t ? 'แก้ไขประเภทของเสีย' : 'เพิ่มประเภทของเสีย'}</h3>
    <div class="fields mt">
      <div class="field"><label>รหัสของเสีย (6 หลัก)</label><input id="tCode" inputmode="numeric" maxlength="6" value="${esc(t?.code || '')}"></div>
      <div class="field"><label>รหัสวิธีการจัดการ (3 หลัก)</label><input id="tMg" inputmode="numeric" maxlength="3" value="${esc(t?.mgmt_code || '')}"></div>
      <div class="field full"><label>ชื่อของเสีย</label><input id="tName" maxlength="120" value="${esc(t?.name || '')}"></div>
      <div class="field"><label>ปริมาณที่ได้รับอนุญาต (ตัน/ปี)</label><input id="tTons" inputmode="decimal" value="${t ? +t.permit_tons : ''}"></div>
      <div class="field"><label>เลขทะเบียนผู้รับ</label><input id="tReg" inputmode="numeric" maxlength="20" value="${esc(t?.receiver_reg || '')}"></div>
      <div class="field full"><label>ผู้รับดำเนินการ</label><input id="tRecv" list="recvList" maxlength="120" value="${esc(t?.receiver || '')}"><datalist id="recvList">${recvList.map(r => `<option value="${esc(r)}">`).join('')}</datalist></div>
      <div class="field full"><label>ไอคอน</label><div class="icon-pick">${ICONS.map(i => `<button type="button" data-ic="${i}" class="${i === ic ? 'sel' : ''}"><i class="fa-solid ${i}"></i></button>`).join('')}</div></div>
      <div class="field full"><label>สี</label><div class="color-pick">${COLORS.map(c => `<button type="button" data-col="${c}" style="--c:${c}" class="${c.toUpperCase() === col.toUpperCase() ? 'sel' : ''}"></button>`).join('')}</div></div>
      <div class="field full"><label class="switch"><input type="checkbox" id="tAct" ${!t || +t.active === 1 ? 'checked' : ''}> ใช้งาน (แสดงในหน้าบันทึกจัดเก็บ)</label></div>
    </div>`, f => recordApi('type_save', { id: t?.id || '', code: $('#tCode', f).value, mgmt_code: $('#tMg', f).value, name: $('#tName', f).value, permit_tons: $('#tTons', f).value, receiver_reg: $('#tReg', f).value, receiver: $('#tRecv', f).value, icon: ic, color: col, active: $('#tAct', f).checked ? 1 : 0 }));
  m.addEventListener('click', e => {
    const i = e.target.closest('[data-ic]'), c = e.target.closest('[data-col]');
    if (i) { ic = i.dataset.ic; $$('[data-ic]', m).forEach(x => x.classList.toggle('sel', x === i)); }
    if (c) { col = c.dataset.col; $$('[data-col]', m).forEach(x => x.classList.toggle('sel', x === c)); }
  });
}

/* ================= RESET DATA ================= */
const RESET_CONFIRM = [
  'ยืนยันการรีเซ็ตข้อมูล?', '',
  'ระบบจะลบข้อมูลต่อไปนี้ทั้งหมด:',
  '• รายการจัดเก็บ / ส่งกำจัด / ยอดยกมา (movements) — ยอดคงเหลือทุกประเภทจะกลับเป็น 0',
  '• บันทึกประวัติ (audit_logs)', '',
  'ข้อมูลที่ยังเก็บไว้: ประเภทของเสีย (waste_types)', '',
  'การลบนี้ย้อนกลับไม่ได้ ต้องการดำเนินการต่อหรือไม่?',
].join('\n');
function resetData() {
  const m = modal(`<h3><i class="fa-solid fa-trash-can" style="color:var(--bad)"></i> รีเซ็ตข้อมูล / Reset data</h3>
    <p class="hint">ใส่รหัสผ่านผู้ดูแล (ตั้งโดยผู้ดูแลระบบใน Script Properties)</p>
    <div class="field mt"><label>รหัสผ่านผู้ดูแล</label><input type="password" id="rPw" autocomplete="off"></div>
    <div id="rErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-bad" id="rGo"><i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล</button></div>`);
  const inp = $('#rPw', m), go = $('#rGo', m), err = $('#rErr', m);
  const idle = () => { go.disabled = false; go.innerHTML = '<i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล'; };
  const submit = async () => {
    if (go.disabled) return;
    let pw = inp.value;
    if (!pw) { err.textContent = 'กรุณาใส่รหัสผ่าน'; inp.focus(); return; }
    if (!confirm(RESET_CONFIRM)) { inp.focus(); return; }
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-recycle fa-spin"></i> กำลังรีเซ็ต...'; err.textContent = '';
    try {
      const reads = routeReads();
      const r = await gasPost({ action: 'reset_data', resetPassword: pw, ...readsField(reads) });
      afterWrite(r, reads);
      pw = ''; inp.value = '';
      closeModal();
      const c = r.removed || {};
      toast(`รีเซ็ตข้อมูลแล้ว — ลบรายการ ${fmt(c.movements)} · Log ${fmt(c.audit_logs)} แถว`);
      route();
    } catch (e) {
      pw = ''; inp.value = '';
      err.textContent = e.message; idle(); inp.focus();
    }
  };
  go.onclick = submit;
  inp.onkeydown = e => { if (e.key === 'Enter') submit(); };
  inp.focus();
}

route();

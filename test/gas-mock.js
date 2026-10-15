/*
 * In-memory mocks of the Apps Script services used by apps-script/Code.gs (Balance Waste),
 * plus an end-to-end test of every API action through doGet/doPost.
 *
 *   node test/gas-mock.js
 */
"use strict";
const vm = require("vm");
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const crypto = require("crypto");
const { Buffer } = require("buffer"); // explicit: test/dev-server.js evaluates this file in a bare vm context

const ROOT = path.join(__dirname, "..");

/* ===================== mocks ===================== */
function createGas() {
  const state = { props: {}, cache: {}, cachePuts: [], cachePutAlls: [], cacheGetAlls: 0, flushes: 0, cacheDown: false, sheetReads: {}, rowReads: 0, events: [], sleeps: 0, locks: 0, unlocks: 0, lockFail: false, spreadsheets: {}, seq: 0, formatCalls: 0, created: 0, fetches: [], fetchStatus: 202, fetchError: null, logs: [] };

  class Range {
    constructor(sheet, row, col, nr, nc) {
      if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error("Range out of bounds");
      if (row + nr - 1 > sheet.maxRows) throw new Error("Range exceeds sheet rows: " + (row + nr - 1) + " > " + sheet.maxRows);
      if (col + nc - 1 > sheet.maxCols) throw new Error("Range exceeds sheet columns: " + (col + nc - 1) + " > " + sheet.maxCols);
      Object.assign(this, { sheet, row, col, nr, nc });
    }
    getValues() {
      if (this.nr === 1 && this.row > 1) state.rowReads++;
      const out = [];
      for (let r = 0; r < this.nr; r++) {
        const src = this.sheet.data[this.row - 1 + r] || [];
        const line = [];
        for (let c = 0; c < this.nc; c++) {
          const v = src[this.col - 1 + c];
          line.push(v === undefined || v === null ? "" : v);
        }
        out.push(line);
      }
      return out;
    }
    setValues(values) {
      state.events.push("write:" + this.sheet.name);
      assert.strictEqual(values.length, this.nr, "setValues row count");
      values.forEach((line, r) => {
        assert.strictEqual(line.length, this.nc, "setValues column count");
        const idx = this.row - 1 + r;
        while (this.sheet.data.length <= idx) this.sheet.data.push([]);
        line.forEach((v, c) => {
          if (typeof v === "string" && v.charAt(0) === "'") v = v.slice(1); // Sheets consumes the apostrophe prefix (stored as text)
          else if (typeof v === "string" && v.charAt(0) === "=") throw new Error("Formula written to sheet: " + v);
          this.sheet.data[idx][this.col - 1 + c] = v;
        });
      });
      return this;
    }
    setNumberFormats(f) {
      assert.strictEqual(f.length, this.nr, "setNumberFormats rows");
      f.forEach(line => assert.strictEqual(line.length, this.nc, "setNumberFormats cols"));
      state.formatCalls++;
      return this;
    }
    setFontWeight() { return this; }
    clearContent() {
      state.clearCalls = (state.clearCalls || 0) + 1;
      for (let r = 0; r < this.nr; r++) {
        const line = this.sheet.data[this.row - 1 + r];
        if (line) for (let c = 0; c < this.nc; c++) if (this.col - 1 + c < line.length) line[this.col - 1 + c] = "";
      }
      return this;
    }
  }

  class Sheet {
    constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; this.maxCols = 26; }
    getName() { return this.name; }
    setName(n) { this.name = n; return this; }
    getLastRow() {
      for (let r = this.data.length - 1; r >= 0; r--) {
        if ((this.data[r] || []).some(v => v !== "" && v !== undefined && v !== null)) return r + 1;
      }
      return 0;
    }
    getLastColumn() { return this.data.reduce((m, row) => Math.max(m, row.length), 0); }
    getMaxRows() { return this.maxRows; }
    insertRowsAfter(after, n) { this.maxRows += n; return this; }
    getMaxColumns() { return this.maxCols; }
    insertColumnsAfter(after, n) { this.maxCols += n; state.colInserts = (state.colInserts || 0) + 1; return this; }
    getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
    getDataRange() {
      state.sheetReads[this.name] = (state.sheetReads[this.name] || 0) + 1;
      return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
    }
    setFrozenRows(n) { this.frozen = n; return this; }
    getFrozenRows() { return this.frozen || 0; }
    deleteRows(pos, n) {
      if (pos < 1 || n < 1 || pos + n - 1 > this.maxRows) throw new Error("deleteRows out of bounds");
      if (this.maxRows - n <= (this.frozen || 0)) throw new Error("Sorry, it is not possible to delete all non-frozen rows.");
      state.deleteCalls = (state.deleteCalls || []).concat([[this.name, pos, n]]);
      this.data.splice(pos - 1, n);
      this.maxRows -= n;
      return this;
    }
    deleteRow(pos) { return this.deleteRows(pos, 1); }
  }

  class Spreadsheet {
    constructor(name) { this.id = "ss" + (++state.seq); this.name = name; this.sheets = [new Sheet("Sheet1")]; this.tz = null; }
    getId() { return this.id; }
    getUrl() { return "https://docs.google.com/spreadsheets/d/" + this.id + "/edit"; }
    getSheets() { return this.sheets.slice(); }
    getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
    insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
    setSpreadsheetTimeZone(tz) { this.tz = tz; }
  }

  const gas = {
    SpreadsheetApp: {
      create(name) { const ss = new Spreadsheet(name); state.spreadsheets[ss.id] = ss; state.created++; return ss; },
      openById(id) { const ss = state.spreadsheets[id]; if (!ss) throw new Error("Spreadsheet not found: " + id); state.opens = (state.opens || 0) + 1; return ss; },
      flush() { state.flushes++; state.events.push("flush"); }
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty: k => (k in state.props ? state.props[k] : null),
          setProperty: (k, v) => { state.props[k] = String(v); },
          setProperties: (o) => { Object.assign(state.props, o); },
          deleteProperty: k => { delete state.props[k]; }
        };
      }
    },
    CacheService: {
      getScriptCache() {
        return {
          get: k => { if (state.cacheDown) throw new Error("Cache down"); return k in state.cache ? state.cache[k] : null; },
          put: (k, v, ttl) => { state.cache[k] = String(v); state.cachePuts.push(ttl); },
          remove: k => { delete state.cache[k]; },
          // Batch calls (read cache / data versions); tracked apart from put().
          getAll: keys => { if (state.cacheDown) throw new Error("Cache down"); state.cacheGetAlls++; const o = {}; keys.forEach(k => { if (k in state.cache) o[k] = state.cache[k]; }); return o; },
          putAll: (map, ttl) => {
            if (state.cacheDown) throw new Error("Cache down");
            Object.keys(map).forEach(k => {
              if (/^WASTE_V_/.test(k)) state.events.push("version:" + k.slice(8));
              const v = String(map[k]);
              assert.ok(k.length <= 250, "cache key too long: " + k);
              assert.ok(Buffer.byteLength(v, "utf8") <= 100 * 1024, "cache value over 100 KB: " + k);
              state.cache[k] = v;
            });
            state.cachePutAlls.push(ttl);
          },
          removeAll: keys => { keys.forEach(k => { delete state.cache[k]; }); }
        };
      }
    },
    LockService: {
      getScriptLock() {
        return {
          waitLock() { if (state.lockFail) throw new Error("Lock timeout"); state.locks++; },
          tryLock() { state.locks++; return true; },
          releaseLock() { state.unlocks++; }
        };
      }
    },
    Utilities: {
      sleep(ms) { state.sleeps += ms; },
      DigestAlgorithm: { SHA_256: "SHA_256" },
      Charset: { UTF_8: "UTF_8" },
      getUuid() { return crypto.randomUUID(); },
      // Like Apps Script: string (charset) or byte array in, signed byte array out.
      computeDigest(alg, value) {
        assert.strictEqual(alg, "SHA_256");
        const buf = typeof value === "string" ? Buffer.from(value, "utf8") : Buffer.from(value.map(b => b & 255));
        return Array.from(crypto.createHash("sha256").update(buf).digest(), b => (b > 127 ? b - 256 : b));
      }
    },
    ContentService: {
      MimeType: { JSON: "application/json", JAVASCRIPT: "application/javascript", CSV: "text/csv", TEXT: "text/plain" },
      createTextOutput(text) {
        return { text, mime: null, setMimeType(m) { this.mime = m; return this; }, getContent() { return this.text; }, getMimeType() { return this.mime; } };
      }
    },
    Logger: { log() {} },
    // Everything Code.gs logs is kept in state.logs (checked for leaked secrets).
    console: {
      log(...a) { state.logs.push(["log", a.join(" ")]); },
      warn(...a) { state.logs.push(["warn", a.join(" ")]); },
      error(...a) { state.logs.push(["error", a.join(" ")]); console.error(...a); }
    },
    // Teams webhook: never touches the network. Records each call (and whether the
    // script lock was held at that moment); answers state.fetchStatus or throws
    // state.fetchError.
    UrlFetchApp: {
      fetch(url, opts) {
        state.fetches.push({ url, opts, lockHeld: state.locks !== state.unlocks });
        if (state.fetchError) throw new Error(state.fetchError);
        const code = state.fetchStatus;
        return { getResponseCode: () => code, getContentText: () => "" };
      }
    },
    // Project triggers (keep-warm). state.triggers: [{handler, everyMinutes, id}]
    ScriptApp: {
      getProjectTriggers() {
        return (state.triggers || []).map(t => ({ getHandlerFunction: () => t.handler, getUniqueId: () => t.id, _t: t }));
      },
      deleteTrigger(tr) { state.triggers = (state.triggers || []).filter(t => t !== tr._t); },
      newTrigger(handler) {
        const t = { handler, id: "trg" + (++state.seq) };
        const b = {
          timeBased() { t.timeBased = true; return b; },
          everyMinutes(n) { assert.ok([1, 5, 10, 15, 30].includes(n), "everyMinutes(" + n + ")"); t.everyMinutes = n; return b; },
          create() { assert.ok(t.timeBased, "time-based"); (state.triggers = state.triggers || []).push(t); return { getUniqueId: () => t.id }; }
        };
        return b;
      }
    }
  };
  return { gas, state };
}

function loadCode(gas) {
  const ctx = vm.createContext(Object.assign({}, gas));
  const src = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  vm.runInContext(src, ctx, { filename: "Code.gs" });
  return ctx;
}


/* ===================== test runner ===================== */
let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ok   " + name); }
  catch (e) { failed++; console.log("  FAIL " + name + "\n       " + String(e && e.stack || e).split("\n").slice(0, 4).join("\n       ")); }
}

const { gas, state } = createGas();
const ctx = loadCode(gas);
const RPW = "test-Reset-9"; // WASTE_RESET_PASSWORD of the in-memory mock (tests only)
const ACTOR = "Tester";
const LOG_ACTIONS = ["logs", "logs_export"];
const get = (action, params = {}) => {
  const out = ctx.doGet({ parameter: Object.assign({ action }, params) });
  assert.strictEqual(out.getMimeType(), "application/json");
  return JSON.parse(out.getContent());
};
const postRaw = (payload, who = ACTOR) =>
  JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(Object.assign(who ? { actor: who } : {}, payload)) } }).getContent());
const PW_ACTIONS = ["movement_update", "movement_delete", "type_save", "reset_data", "logs", "logs_export"];
const post = (action, body = {}, who = ACTOR) => postRaw(Object.assign(PW_ACTIONS.includes(action) ? { resetPassword: RPW } : {}, body, { action }), who);
const okData = res => { assert.strictEqual(res.ok, true, "expected ok, got: " + JSON.stringify(res)); return res.data; };
const err = (res, re) => {
  assert.strictEqual(res.ok, false, "expected error, got: " + JSON.stringify(res).slice(0, 300));
  if (re) assert.ok(re.test(res.error), "error '" + res.error + "' !~ " + re);
  return res;
};
const clearFails = () => Object.keys(state.cache).forEach(k => { if (/FAILURES/.test(k)) delete state.cache[k]; });
const p2 = n => String(n).padStart(2, "0");
const bkkNow = () => new Date(Date.now() + 7 * 3600 * 1000);
const fmtD = d => `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`;
const TODAY = fmtD(bkkNow());
const YMD = TODAY.slice(2).replace(/-/g, "");
// a moment `back` days ago at hh:mm (Asia/Bangkok)
const at = (back, hh = 9, mm = 0) => ({ date: fmtD(new Date(bkkNow().getTime() - back * 86400000)), time: p2(hh) + ":" + p2(mm) });
const sheet = n => state.spreadsheets[state.props.WASTE_SPREADSHEET_ID].getSheetByName(n);
const dataRows = n => sheet(n).getLastRow() - 1;
const boot = () => okData(get("bootstrap"));
const bal = id => boot().balance[id];
const logs = (p = {}) => okData(post("logs", p)).items;
const lastLog = () => logs()[0];
const mv = (kind, type_id, weight_kg, extra = {}) => post(kind, Object.assign({ type_id, weight_kg, container: "drum200", containers: 1 }, at(5), extra));

console.log("Balance Waste GAS backend tests (today " + TODAY + " Asia/Bangkok)");

/* ---------- setup ---------- */
test("API before setup: reads and writes report the setupSystem error", () => {
  err(get("bootstrap"), /setupSystem/);
  err(mv("store", 1, 10), /setupSystem/);
});

test("setupSystem creates spreadsheet, 3 sheets, headers and seeds the 13 permitted waste types", () => {
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, true);
  assert.strictEqual(r.seededTypes, 13);
  assert.strictEqual(r.resetPasswordConfigured, false);
  const ss = state.spreadsheets[state.props.WASTE_SPREADSHEET_ID];
  assert.strictEqual(ss.tz, "Asia/Bangkok");
  assert.strictEqual(ss.name, "Balance Waste Database");
  assert.deepStrictEqual(ss.getSheets().map(s => s.getName()), ["waste_types", "movements", "audit_logs"]);
  assert.deepStrictEqual(sheet("waste_types").data[0], ["id", "code", "name", "permit_tons", "mgmt_code", "receiver_reg", "receiver", "icon", "color", "active", "sort", "updated_at", "updated_by"]);
  assert.deepStrictEqual(sheet("movements").data[0], ["id", "doc_no", "kind", "type_id", "weight_kg", "container", "containers", "location", "manifest_no", "vehicle", "note", "moved_at", "created_at", "actor", "updated_at", "updated_by", "method"]);
  assert.strictEqual(dataRows("waste_types"), 13);
  assert.strictEqual(dataRows("movements"), 0);
  state.props.WASTE_RESET_PASSWORD = RPW;
});

test("seeded types match the สก.2 2569 permit table (codes keep leading zeros, total 194 t)", () => {
  const b = boot();
  const t = b.types;
  assert.strictEqual(t.length, 13);
  assert.deepStrictEqual(t.map(x => x.code), ["150101", "150102", "170405", "150111", "080409", "150202", "160215", "160213", "150102", "191201", "190813", "160215", "161001"]);
  assert.deepStrictEqual(t.map(x => x.permit_tons), [20, 30, 10, 2, 10, 30, 0.5, 1, 20, 20, 10, 0.5, 40]);
  assert.deepStrictEqual(t.map(x => x.mgmt_code), ["011", "011", "011", "073", "042", "042", "073", "073", "046", "046", "073", "073", "065"]);
  assert.strictEqual(t[4].code, "080409");
  assert.strictEqual(Math.round(t.reduce((a, x) => a + x.permit_tons, 0) * 1000) / 1000, 194);
  const byRecv = {};
  t.forEach(x => { byRecv[x.receiver] = (byRecv[x.receiver] || 0) + x.permit_tons; });
  assert.deepStrictEqual(byRecv, { "บจก. มั่นคงยั่งยืน รีไซเคิล": 60, "บมจ. เบตเตอร์ เวิลด์ กรีน": 83.5, "บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม": 50.5 });
  assert.strictEqual(t[10].receiver_reg, "10700001425472");
  assert.strictEqual(t[12].receiver_reg, "10103300225380");
  assert.ok(t.every(x => x.active === 1 && /^fa-/.test(x.icon) && /^#[0-9A-F]{6}$/.test(x.color)));
  // same code twice = two separate permit lines (ids)
  assert.notStrictEqual(t[6].id, t[11].id);
  assert.strictEqual(t[6].code, t[11].code);
  assert.notStrictEqual(t[6].receiver, t[11].receiver);
  assert.deepStrictEqual(b.containers, { drum200: "ถัง 200 ลิตร", ibc: "IBC", bigbag: "Bigbag" });
  assert.strictEqual(b.company.factoryReg, "82260007625693");
  assert.strictEqual(b.company.permitNo, "สก.2 เลขที่ 2569-8409");
  // no balances invented: every type starts at 0
  Object.values(b.balance).forEach(x => { assert.strictEqual(x.balance_kg, 0); assert.strictEqual(x.age_days, null); });
});

test("setupSystem is idempotent (no second spreadsheet, no duplicate seed)", () => {
  const id = state.props.WASTE_SPREADSHEET_ID;
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(r.seededTypes, 0);
  assert.strictEqual(state.props.WASTE_SPREADSHEET_ID, id);
  assert.strictEqual(state.created, 1);
  assert.strictEqual(dataRows("waste_types"), 13);
  assert.strictEqual(r.resetPasswordConfigured, true);
});

/* ---------- transport ---------- */
test("GET: service info, unknown action, writes and logs refused over GET", () => {
  assert.strictEqual(okData(get("")).service, "Balance Waste API");
  err(get("nope"), /Unknown action/);
  for (const a of ["store", "ship", "opening", "treat", "treat_unlock", "treat_lock", "movement_update", "movement_delete", "type_save", "reset_data", "logs", "logs_export", "verifyResetPassword"]) err(get(a), /POST/);
  const bad = JSON.parse(ctx.doPost({ postData: { contents: "{oops" } }).getContent());
  err(bad, /JSON/);
  err(postRaw({ action: "nope" }), /Unknown action/);
});

test("actor name required for every write (empty / whitespace / object rejected, nothing written)", () => {
  for (const who of [null, "", "   ", { x: 1 }]) {
    for (const a of ["store", "ship", "opening", "treat", "movement_update", "movement_delete", "type_save", "reset_data"]) {
      const r = postRaw(Object.assign({ action: a, type_id: 1, weight_kg: 5, container: "drum200", containers: 1, resetPassword: RPW }, who && typeof who === "object" ? { actor: who } : {}), typeof who === "string" ? who : null);
      err(r, /ระบุชื่อผู้ใช้งาน/);
      assert.strictEqual(r.code, "NAME_REQUIRED");
    }
  }
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("audit_logs"), 0);
});

/* ---------- store validations ---------- */
test("store: field validation (type, weight, decimals, container, count, date, text length)", () => {
  err(mv("store", 0, 10), /เลือกประเภทของเสีย/);
  err(mv("store", 99, 10), /เลือกประเภทของเสีย/);
  err(mv("store", 1, 0), /มากกว่า 0/);
  err(mv("store", 1, ""), /กรอกน้ำหนัก/);
  err(mv("store", 1, null), /กรอกน้ำหนัก/);
  err(mv("store", 1, -5), /ตัวเลข/);
  err(mv("store", 1, "abc"), /ตัวเลข/);
  err(mv("store", 1, "1e3"), /ตัวเลข/);
  err(mv("store", 1, "10.123"), /ทศนิยมได้ไม่เกิน 2/);
  err(mv("store", 1, 10.005), /ทศนิยมได้ไม่เกิน 2/);
  err(mv("store", 1, [5]), /กรอกน้ำหนัก/);
  err(mv("store", 1, "2000000"), /ไม่เกิน/);
  err(mv("store", 1, 10, { container: "box" }), /ภาชนะ/);
  err(mv("store", 1, 10, { container: "" }), /ภาชนะ/);
  err(mv("store", 1, 10, { containers: 0 }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { containers: "1.5" }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { containers: -2 }), /จำนวนเต็มตั้งแต่ 1/);
  err(mv("store", 1, 10, { date: "2026-02-30" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, { date: "15/10/2026" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, { time: "25:00" }), /วันที่\/เวลาไม่ถูกต้อง/);
  err(mv("store", 1, 10, at(-2)), /ไม่เกินเวลาปัจจุบัน/);
  err(mv("store", 1, 10, { note: "x".repeat(256) }), /หมายเหตุยาวเกิน 255/);
  err(mv("store", 1, 10, { location: "x".repeat(101) }), /สถานที่จัดเก็บยาวเกิน 100/);
  err(mv("store", 1, 10, { type_id: { a: 1 } }), /เลือกประเภท/);
  assert.strictEqual(dataRows("movements"), 0);
  // rejected writes are audited as warnings
  const w = logs({ cat: "warn" });
  assert.ok(w.length >= 10 && w.every(l => /^ปฏิเสธ/.test(l.detail) && l.actor === ACTOR));
});

test("store: saved with doc number, actor, 2-decimal weight; comma input accepted; audit row", () => {
  const r = okData(mv("store", 1, "1,234.5", { container: "bigbag", containers: 3, location: "โรงเก็บ A", note: "ทดสอบ" }));
  assert.strictEqual(r.doc_no, "ST-" + YMD + "-0001");
  assert.strictEqual(r.kind, "store");
  assert.strictEqual(r.balance_kg, 1234.5);
  assert.ok(/บันทึกจัดเก็บ/.test(r.message));
  const row = sheet("movements").data[1];
  assert.strictEqual(row[1], r.doc_no);
  assert.strictEqual(row[2], "store");
  assert.strictEqual(row[4], 1234.5);
  assert.strictEqual(row[5], "bigbag");
  assert.strictEqual(row[6], 3);
  assert.strictEqual(row[7], "โรงเก็บ A");
  assert.strictEqual(row[13], ACTOR);
  const l = lastLog();
  assert.strictEqual(l.category, "waste");
  assert.strictEqual(l.ref, r.doc_no);
  assert.ok(l.detail.includes("1,234.5 กก.") && l.detail.includes("Bigbag 3 ภาชนะ"));
  const r2 = okData(mv("store", 1, 0.1, { container: "bigbag", containers: 1 }));
  assert.strictEqual(r2.doc_no, "ST-" + YMD + "-0002");
});

test("balance math is exact to 0.01 kg (no float drift) and counts containers per type", () => {
  // type 2: 0.1 + 0.2 ... x10 = 3.00 exactly
  for (let i = 0; i < 10; i++) okData(mv("store", 2, i % 2 ? "0.2" : "0.1", { container: "ibc", containers: 1 }));
  const b = bal(2);
  assert.strictEqual(b.balance_kg, 1.5);
  assert.strictEqual(b.stored_kg, 1.5);
  assert.deepStrictEqual(b.containers, { drum200: 0, ibc: 10, bigbag: 0 });
  okData(mv("ship", 2, "0.3", { container: "ibc", containers: 2, ...at(4) }));
  assert.strictEqual(bal(2).balance_kg, 1.2);
  assert.strictEqual(bal(2).containers.ibc, 8);
  assert.strictEqual(bal(1).balance_kg, 1234.6);
  assert.strictEqual(bal(1).containers.bigbag, 4);
});

/* ---------- ship ---------- */
test("ship: cannot exceed the current balance or containers; exact remaining amount allowed", () => {
  okData(mv("store", 3, 100, { container: "bigbag", containers: 2, ...at(10) }));
  const before = dataRows("movements");
  const r = err(mv("ship", 3, "100.01", { container: "bigbag", containers: 1, ...at(2) }), /ส่งกำจัดเกินยอดคงเหลือ/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  assert.ok(/100 กก\. แต่ต้องนำออก 100\.01 กก\./.test(r.error), r.error);
  err(mv("ship", 3, 50, { container: "bigbag", containers: 3, ...at(2) }), /Bigbag.*มี 2 ภาชนะ แต่ต้องนำออก 3/);
  err(mv("ship", 3, 50, { container: "drum200", containers: 1, ...at(2) }), /ถัง 200 ลิตร/);
  err(mv("ship", 4, 1, { ...at(2) }), /มี 0 กก\./);
  assert.strictEqual(dataRows("movements"), before);
  const s = okData(mv("ship", 3, 100, { container: "bigbag", containers: 2, manifest_no: "MN-1", vehicle: "70-1234", ...at(2) }));
  assert.strictEqual(s.doc_no, "SH-" + YMD + "-0002"); // SH-0001 used by the balance-math test
  assert.strictEqual(s.balance_kg, 0);
  const b = bal(3);
  assert.strictEqual(b.balance_kg, 0);
  assert.strictEqual(b.containers.bigbag, 0);
  assert.strictEqual(b.age_days, null);
  const row = sheet("movements").data[dataRows("movements")];
  assert.strictEqual(row[8], "MN-1");
  assert.strictEqual(row[9], "70-1234");
  assert.ok(/ผู้รับ: บจก\. มั่นคงยั่งยืน รีไซเคิล/.test(lastLog().detail));
  err(mv("ship", 3, "0.01", { ...at(1) }), /ส่งกำจัดเกินยอดคงเหลือ/);
});

test("ship: a back-dated shipment before the stock existed is refused (chronological replay)", () => {
  okData(mv("store", 5, 50, { ...at(3) }));
  err(mv("ship", 5, 10, { ...at(6) }), /ส่งกำจัดเกินยอดคงเหลือ/);
  okData(mv("ship", 5, 10, { ...at(2) }));
  assert.strictEqual(bal(5).balance_kg, 40);
});

test("ship of a store-only field is ignored; store ignores manifest/vehicle", () => {
  const r = okData(mv("store", 6, 20, { manifest_no: "X", vehicle: "Y" }));
  const row = sheet("movements").data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[8], "");
  assert.strictEqual(row[9], "");
});

/* ---------- opening ---------- */
test("opening balance (ยอดยกมา) adds to balance but not to 'stored this month'", () => {
  const dash0 = okData(get("dashboard")).kpi;
  const r = okData(mv("opening", 13, 54000, { container: "ibc", containers: 54, ...at(1) }));
  assert.strictEqual(r.doc_no, "OB-" + YMD + "-0001");
  const b = bal(13);
  assert.strictEqual(b.balance_kg, 54000);
  assert.strictEqual(b.opening_kg, 54000);
  assert.strictEqual(b.stored_kg, 0);
  assert.strictEqual(b.in_kg, 54000);
  const dash = okData(get("dashboard")).kpi;
  assert.strictEqual(dash.month_in_kg, dash0.month_in_kg);
  assert.strictEqual(Math.round((dash.balance_kg - dash0.balance_kg) * 100) / 100, 54000);
});

/* ---------- quota ---------- */
test("quota %: shipped this year vs permit (tons -> kg), warn at >= 80 %, over at > 100 %", () => {
  // type 7: หลอดไฟ 0.5 t = 500 kg
  okData(mv("store", 7, 700, { containers: 5, ...at(3) }));
  okData(mv("ship", 7, 399, { ...at(2) }));
  let b = bal(7);
  assert.strictEqual(b.permit_kg, 500);
  assert.strictEqual(b.quota_pct, 79.8);
  assert.strictEqual(b.quota_level, "ok");
  okData(mv("ship", 7, 1, { ...at(2) }));
  b = bal(7);
  assert.strictEqual(b.quota_pct, 80);
  assert.strictEqual(b.quota_level, "warn");
  okData(mv("ship", 7, "100.01", { ...at(1) }));
  b = bal(7);
  assert.strictEqual(b.shipped_year_kg, 500.01);
  assert.strictEqual(b.quota_level, "over");
  const d = okData(get("dashboard"));
  const a = d.alerts.find(x => x.type_id === 7);
  assert.ok(a && a.kind === "quota" && a.level === "bad");
  assert.strictEqual(d.alerts[0].level, "bad"); // over-quota alerts first
  const bwg = d.receivers.find(r => r.receiver === "บมจ. เบตเตอร์ เวิลด์ กรีน");
  assert.strictEqual(bwg.permit_kg, 83500);
  assert.strictEqual(bwg.types, 7);
  assert.strictEqual(bwg.shipped_year_kg, 500.01 + 10); // + 10 kg of type 5
  assert.strictEqual(d.kpi.permit_kg, 194000);
  assert.deepStrictEqual(d.receivers.map(r => r.permit_kg), [60000, 83500, 50500]);
});

test("dashboard: 12 months oldest first, KPI totals, containers in storage, recent list", () => {
  const d = okData(get("dashboard"));
  assert.strictEqual(d.months.length, 12);
  assert.strictEqual(d.months[11].ym, TODAY.slice(0, 7));
  assert.ok(d.months[0].ym < d.months[11].ym);
  const b = boot().balance;
  const sum = Object.values(b).reduce((a, x) => a + Math.round(x.balance_kg * 100), 0) / 100;
  assert.strictEqual(d.kpi.balance_kg, sum);
  const yearOut = Object.values(b).reduce((a, x) => a + Math.round(x.shipped_year_kg * 100), 0) / 100;
  assert.strictEqual(d.kpi.year_out_kg, yearOut);
  assert.strictEqual(d.kpi.quota_pct, Math.round(yearOut / 194000 * 1000) / 10);
  assert.strictEqual(d.containers.ibc, Object.values(b).reduce((a, x) => a + x.containers.ibc, 0));
  assert.strictEqual(d.recent.length, 10);
  assert.ok(d.recent[0].moved_at >= d.recent[9].moved_at);
  assert.ok(d.logsToday > 0);
});

test("storage age (FIFO): oldest lot still in storage, alert after 90 days", () => {
  okData(mv("store", 8, 10, { ...at(120, 0) }));
  okData(mv("store", 8, 10, { ...at(30, 0) }));
  let b = bal(8);
  assert.strictEqual(b.age_days, 120);
  okData(mv("ship", 8, 10, { ...at(20) })); // ships the oldest lot first
  b = bal(8);
  assert.strictEqual(b.age_days, 30);
  okData(mv("store", 10, 5, { ...at(100, 0) }));
  const d = okData(get("dashboard"));
  const a = d.alerts.find(x => x.type_id === 10 && x.kind === "age");
  assert.ok(a && a.days === 100 && a.balance_kg === 5);
});

/* ---------- history / CSV ---------- */
test("history: newest first with type info; filters kind / type / receiver / dates / search", () => {
  const all = okData(get("history")).items;
  assert.strictEqual(all.length, dataRows("movements"));
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].moved_at >= all[i].moved_at);
  assert.ok(all.every(m => m.code && m.type_name && m.receiver && m.icon && m.color));
  assert.ok(okData(get("history", { kind: "ship" })).items.every(m => m.kind === "ship"));
  assert.strictEqual(okData(get("history", { kind: "opening" })).items.length, 1);
  assert.ok(okData(get("history", { type: "7" })).items.every(m => m.type_id === 7));
  const g = okData(get("history", { receiver: "บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม" })).items;
  assert.ok(g.length && g.every(m => [11, 12, 13].includes(m.type_id)));
  assert.strictEqual(okData(get("history", { q: "MN-1" })).items.length, 1);
  assert.ok(okData(get("history", { q: "หลอดไฟ" })).items.length >= 4);
  const d = at(30).date;
  assert.ok(okData(get("history", { from: d, to: d })).items.every(m => m.moved_at.slice(0, 10) === d));
  assert.strictEqual(okData(get("history", { from: "2001-01-01", to: "2001-01-02" })).items.length, 0);
});

test("CSV export: Thai headers, codes keep leading zeros, formula-like text neutralised", () => {
  okData(mv("store", 5, 1, { note: "=HYPERLINK(\"http://x\")", location: "+cmd", ...at(1) })); // also: never written as a formula
  const r = okData(get("export", { type: "5" }));
  assert.ok(/^waste_history_\d{8}_\d{6}\.csv$/.test(r.filename));
  const lines = r.csv.trim().split("\n");
  assert.ok(lines[0].startsWith('เลขที่เอกสาร,วันที่,เวลา,รายการ,รหัสของเสีย,ชื่อของเสีย,รหัสการจัดการ,ผู้รับดำเนินการ,"น้ำหนัก (กก.)",ภาชนะ,จำนวนภาชนะ'), lines[0]);
  assert.ok(lines[0].endsWith("ผู้บันทึก"));
  assert.ok(r.csv.includes('"=""080409"""'), "code as Excel text");
  assert.ok(r.csv.includes('"=""042"""'));
  assert.ok(r.csv.includes("'=HYPERLINK") && r.csv.includes("'+cmd"));
  assert.ok(!/,=HYPERLINK/.test(r.csv));
  assert.ok(r.csv.includes(",1.00,"));
  assert.strictEqual(lines.length - 1, okData(get("history", { type: "5" })).items.length);
  const row = sheet("movements").data.find(x => String(x[10]).includes("HYPERLINK"));
  assert.strictEqual(row[10], "=HYPERLINK(\"http://x\")"); // stored as text via apostrophe
});

/* ---------- in-house treatment (บำบัดเองภายใน) + hidden mode ---------- */
const unlock = () => okData(postRaw({ action: "treat_unlock", resetPassword: RPW }, null)).token;
// a read sent as POST with a treatment token (the page does this while unlocked)
const tread = (tok, action, params = {}) => okData(postRaw(Object.assign({ action, treatToken: tok }, params), null));
let TT = "";
let TR1 = null;

test("treat: POST only; locked without a token; unlock needs the admin password (same lockout), audited", () => {
  const before = dataRows("movements");
  let r = err(mv("treat", 11, 5), /ล็อก/);
  assert.strictEqual(r.code, "TREAT_LOCKED");
  err(mv("treat", 11, 5, { treatToken: "x".repeat(64) }), /ล็อก/);
  err(mv("treat", 11, 5, { treatToken: "a".repeat(64) }), /ล็อก/); // well-formed but never issued
  clearFails();
  r = err(postRaw({ action: "treat_unlock", resetPassword: "bad" }, null), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.cache.WASTE_RESET_FAILURES, "1"); // same counter as every admin password check
  err(postRaw({ action: "treat_unlock" }, null), /รหัสไม่ถูกต้อง/);
  state.cache.WASTE_RESET_FAILURES = "10";
  err(postRaw({ action: "treat_unlock", resetPassword: RPW }, null), /หลายครั้งเกินไป/);
  clearFails();
  const puts = state.cachePuts.length;
  TT = unlock();
  assert.ok(/^[0-9a-f]{64}$/.test(TT));
  assert.ok(state.cachePuts.slice(puts).includes(21600), "token kept 6 h (CacheService maximum)");
  assert.ok(Number(state.cache["WASTE_TT_" + TT]) > Date.now());
  assert.strictEqual(dataRows("movements"), before);
  assert.strictEqual(lastLog().action, "treat_unlock");
  assert.ok(!JSON.stringify(sheet("audit_logs").data).includes(TT), "token never in the audit log");
});

test("treat: lowers the in-plant balance like a shipment but is never counted in the สก.2 quota", () => {
  okData(mv("store", 12, 300, { container: "drum200", containers: 3, ...at(6) }));
  okData(mv("ship", 12, 100, { container: "drum200", containers: 1, ...at(4) }));
  const d0 = tread(TT, "dashboard");
  const r = okData(mv("treat", 12, 150, { container: "drum200", containers: 1, method: "ระบบบำบัดน้ำเสียภายในโรงงาน", note: "ทดสอบ", treatToken: TT, ...at(3) }));
  TR1 = r;
  assert.strictEqual(r.doc_no, "TR-" + YMD + "-0001");
  assert.strictEqual(r.kind, "treat");
  assert.strictEqual(r.balance_kg, 50);
  assert.ok(/บำบัดเองภายใน/.test(r.message));
  const row = sheet("movements").data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[2], "treat");
  assert.strictEqual(row[16], "ระบบบำบัดน้ำเสียภายในโรงงาน"); // method: appended column
  assert.strictEqual(row[8], ""); // no manifest / vehicle for a treatment
  // hidden (anonymous) view: true balance, quota counts the shipment only, no treatment totals
  const b = bal(12);
  assert.strictEqual(b.balance_kg, 50);
  assert.strictEqual(b.containers.drum200, 1);
  assert.strictEqual(b.shipped_kg, 100);
  assert.strictEqual(b.shipped_year_kg, 100);
  assert.strictEqual(b.permit_kg, 500);
  assert.strictEqual(b.quota_pct, 20);
  assert.ok(!("treated_kg" in b) && !("treated_year_kg" in b));
  assert.strictEqual(boot().treat, false);
  // unlocked view: the treatment totals add up
  const bt = tread(TT, "bootstrap");
  assert.strictEqual(bt.treat, true);
  const u = bt.balance[12];
  assert.strictEqual(u.treated_kg, 150);
  assert.strictEqual(u.treated_year_kg, 150);
  assert.strictEqual(u.quota_pct, 20);
  assert.strictEqual(Math.round((u.in_kg - u.shipped_kg - u.treated_kg) * 100) / 100, u.balance_kg);
  const d = tread(TT, "dashboard");
  assert.strictEqual(d.kpi.year_out_kg, d0.kpi.year_out_kg, "quota use unchanged by a treatment");
  assert.strictEqual(d.kpi.quota_pct, d0.kpi.quota_pct);
  assert.deepStrictEqual(d.receivers.map(x => x.shipped_year_kg), d0.receivers.map(x => x.shipped_year_kg));
  assert.strictEqual(Math.round((d.kpi.treated_kg - d0.kpi.treated_kg) * 100) / 100, 150);
  const k = d.kpi;
  assert.strictEqual(Math.round(k.total_in_kg * 100), Math.round(k.shipped_kg * 100) + Math.round(k.treated_kg * 100) + Math.round(k.balance_kg * 100));
  assert.ok(k.year_treat_kg >= 150 && k.month_treat_kg >= 0);
  assert.ok(d.months.reduce((a, m) => a + m.treat_kg, 0) >= 150);
  assert.ok(!d.alerts.some(a => a.type_id === 12 && a.kind === "quota"));
  const l = logs()[0];
  assert.strictEqual(l.action, "treat");
  assert.strictEqual(l.ref, r.doc_no);
  assert.ok(l.detail.includes("บำบัดเองภายใน") && l.detail.includes("วิธีบำบัด: ระบบบำบัดน้ำเสียภายในโรงงาน") && l.actor === ACTOR);
});

test("treat: cannot exceed the balance or containers (chronological replay); hidden error never names the treatment row", () => {
  const before = dataRows("movements");
  let r = err(mv("treat", 12, "50.01", { treatToken: TT, ...at(2) }), /บำบัดเกินยอดคงเหลือ/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  err(mv("treat", 12, 10, { containers: 2, treatToken: TT, ...at(2) }), /ถัง 200 ลิตร.*มี 1 ภาชนะ แต่ต้องนำออก 2/);
  err(mv("treat", 12, 10, { container: "ibc", treatToken: TT, ...at(2) }), /IBC/);
  err(mv("treat", 4, 1, { treatToken: TT, ...at(2) }), /มี 0 กก\./);
  err(mv("treat", 12, 10, { treatToken: TT, ...at(7) }), /บำบัดเกินยอดคงเหลือ/); // before the stock existed
  err(mv("treat", 12, 10, { method: "x".repeat(121), treatToken: TT, ...at(2) }), /วิธีบำบัด \/ ระบบยาวเกิน 120/);
  // a back-dated shipment that would make the later treatment negative
  r = err(mv("ship", 12, 60, { ...at(5) }), /ติดลบในรายการภายหลัง/);
  assert.ok(!/TR-|บำบัด/.test(r.error), r.error);
  r = err(mv("ship", 12, 60, { treatToken: TT, ...at(5) }), /ติดลบที่รายการ TR-.*บำบัดเองภายใน/);
  assert.strictEqual(dataRows("movements"), before);
  okData(mv("treat", 12, "0.5", { treatToken: TT, ...at(2) }));
  assert.strictEqual(bal(12).balance_kg, 49.5);
});

test("hidden mode: history / CSV / dashboard / batch leave out treatment rows; token reads (POST) include them", () => {
  const all = okData(get("history")).items;
  assert.ok(all.length > 0 && all.every(m => m.kind !== "treat"));
  assert.strictEqual(all.length, dataRows("movements") - 2);
  assert.strictEqual(okData(get("history", { kind: "treat" })).items.length, 0);
  assert.strictEqual(okData(get("history", { q: "บำบัดน้ำเสีย" })).items.length, 0);
  assert.strictEqual(okData(get("history", { q: "TR-" })).items.length, 0);
  // a token in a GET query is ignored (tokens travel in POST bodies only)
  assert.ok(okData(get("history", { treatToken: TT })).items.every(m => m.kind !== "treat"));
  assert.strictEqual(okData(get("bootstrap", { treatToken: TT })).treat, false);
  const dh = okData(get("dashboard"));
  assert.ok(!/treat|TR-/.test(JSON.stringify(dh)), "no treatment data in the hidden dashboard");
  assert.ok(dh.recent.every(m => m.kind !== "treat"));
  assert.ok(!/treated|TR-/.test(JSON.stringify(boot())));
  // unlocked
  const th = tread(TT, "history", { kind: "treat" }).items;
  assert.strictEqual(th.length, 2);
  assert.strictEqual(th[1].method, "ระบบบำบัดน้ำเสียภายในโรงงาน");
  assert.strictEqual(tread(TT, "history").items.length, dataRows("movements"));
  assert.strictEqual(tread(TT, "history", { q: "บำบัดน้ำเสีย" }).items.length, 1);
  assert.ok(tread(TT, "dashboard").recent.some(m => m.kind === "treat"));
  const bt = tread(TT, "batch", { reads: [{ action: "bootstrap" }, { action: "history", kind: "treat" }] });
  assert.strictEqual(bt.results[0].treat, true);
  assert.strictEqual(bt.results[1].items.length, 2);
  const bh = okData(get("batch", { reads: JSON.stringify([{ action: "history", kind: "treat" }]) }));
  assert.strictEqual(bh.results[0].items.length, 0);
  // CSV
  const c0 = okData(get("export", { type: "12" })).csv;
  assert.ok(!/TR-|บำบัด|วิธีบำบัด/.test(c0), "hidden CSV has no treatment rows / column");
  const c1 = tread(TT, "export", { type: "12" }).csv;
  const lines = c1.trim().split("\n");
  assert.ok(lines[0].endsWith(",ผู้บันทึก,\"วิธีบำบัด / ระบบ\""), lines[0]);
  const tr = lines.find(x => x.startsWith(TR1.doc_no));
  assert.ok(tr && tr.includes(",บำบัดเองภายใน,") && tr.includes(",บำบัดเองภายในโรงงาน,") && tr.endsWith(",ระบบบำบัดน้ำเสียภายในโรงงาน"), tr);
  assert.strictEqual(lines.length - 1, tread(TT, "history", { type: "12" }).items.length);
  // a write carrying reads answers in the mode of its token
  const w0 = okData(mv("store", 12, 1, { ...at(1), reads: [{ action: "bootstrap" }] }));
  assert.strictEqual(w0.reads[0].treat, false);
  assert.ok(!("treated_kg" in w0.reads[0].balance[12]));
  const w1 = okData(mv("store", 12, 1, { ...at(1), treatToken: TT, reads: [{ action: "bootstrap" }, { action: "history", kind: "treat" }] }));
  assert.strictEqual(w1.reads[0].treat, true);
  assert.strictEqual(w1.reads[0].balance[12].treated_kg, 150.5);
  assert.strictEqual(w1.reads[1].items.length, 2);
  // dashboard "logs today" count leaves out treatment audit rows when hidden
  assert.ok(okData(get("dashboard")).logsToday < tread(TT, "dashboard").logsToday);
});

test("treat token: lock drops it; an expired / evicted token is refused for reads and writes", () => {
  const t2 = unlock();
  assert.strictEqual(tread(t2, "bootstrap").treat, true);
  assert.strictEqual(okData(postRaw({ action: "treat_lock", treatToken: t2 }, null)).locked, true);
  assert.ok(!("WASTE_TT_" + t2 in state.cache));
  assert.strictEqual(tread(t2, "bootstrap").treat, false);
  assert.strictEqual(tread(t2, "history", { kind: "treat" }).items.length, 0);
  err(mv("treat", 12, 1, { treatToken: t2, ...at(1) }), /ล็อก/);
  okData(postRaw({ action: "treat_lock", treatToken: "nope" }, null)); // harmless
  const t3 = unlock();
  state.cache["WASTE_TT_" + t3] = String(Date.now() - 1); // expired
  assert.strictEqual(tread(t3, "bootstrap").treat, false);
  err(mv("treat", 12, 1, { treatToken: t3, ...at(1) }), /ล็อก/);
  delete state.cache["WASTE_TT_" + t3]; // evicted
  assert.strictEqual(tread(t3, "dashboard").kpi.treated_kg, undefined);
  state.cacheDown = true; // cache outage: locked, reads still answer
  try { assert.strictEqual(tread(TT, "bootstrap").treat, false); } finally { state.cacheDown = false; }
  assert.strictEqual(tread(TT, "bootstrap").treat, true); // the first token still works
});

test("treat rows: edit / delete with the admin password like any row (replay checked, audited)", () => {
  const id = TR1.id;
  err(postRaw({ action: "movement_update", id, weight_kg: 100 }), /รหัสไม่ถูกต้อง/);
  clearFails();
  // stock of type 12: 300 + 1 + 1 in, 100 shipped, 150 + 0.5 treated -> 51.5 left
  err(post("movement_update", { id, doc_no: TR1.doc_no, weight_kg: "201.51" }), /แก้ไขไม่ได้/);
  err(post("movement_update", { id, kind: "ship" }), /เปลี่ยนชนิดรายการไม่ได้/);
  const ok = okData(post("movement_update", { id, doc_no: TR1.doc_no, weight_kg: 120, method: "เตาเผาภายในโรงงาน" }));
  assert.strictEqual(ok.changed, true);
  assert.strictEqual(bal(12).balance_kg, 81.5);
  assert.strictEqual(bal(12).quota_pct, 20);
  const row = sheet("movements").data.find(x => x[1] === TR1.doc_no);
  assert.strictEqual(row[4], 120);
  assert.strictEqual(row[16], "เตาเผาภายในโรงงาน");
  const l = lastLog();
  assert.strictEqual(l.action, "movement_update");
  assert.ok(l.detail.includes("ระบบบำบัดน้ำเสียภายในโรงงาน") && l.detail.includes("เตาเผาภายในโรงงาน"));
  // a store the treatment depends on cannot be deleted; the treatment itself can
  const b6 = bal(6).balance_kg;
  const st = okData(mv("store", 6, 40, { container: "ibc", containers: 1, ...at(9) }));
  const tr = okData(mv("treat", 6, 40, { container: "ibc", containers: 1, treatToken: TT, ...at(8) }));
  err(post("movement_delete", { id: st.id, doc_no: st.doc_no }), /ลบไม่ได้.*ติดลบในรายการภายหลัง/);
  okData(post("movement_delete", { id: tr.id, doc_no: tr.doc_no }));
  assert.strictEqual(lastLog().action, "movement_delete");
  okData(post("movement_delete", { id: st.id, doc_no: st.doc_no }));
  assert.strictEqual(bal(6).balance_kg, b6);
});

test("migration: an older movements sheet without the method column keeps working (column appended on first use)", () => {
  const g = createGas();
  const c2 = loadCode(g.gas);
  c2.setupSystem();
  g.state.props.WASTE_RESET_PASSWORD = RPW;
  const ss = g.state.spreadsheets[g.state.props.WASTE_SPREADSHEET_ID];
  const ms = ss.getSheetByName("movements");
  const p = body => JSON.parse(c2.doPost({ postData: { contents: JSON.stringify(Object.assign({ actor: ACTOR }, body)) } }).getContent());
  okData(p({ action: "store", type_id: 13, weight_kg: 500, container: "ibc", containers: 1, ...at(3) }));
  // simulate the sheet as created by the previous version: no "method" column
  ms.data.forEach(row => { row.length = 16; });
  assert.ok(!ms.data[0].includes("method"));
  c2.clearReadCache();
  const g0 = JSON.parse(c2.doGet({ parameter: { action: "bootstrap" } }).getContent());
  assert.strictEqual(okData(g0).balance[13].balance_kg, 500);
  assert.strictEqual(ms.data[0][16], "method", "column appended to the header");
  const tok = okData(p({ action: "treat_unlock", resetPassword: RPW })).token;
  const r = okData(p({ action: "treat", type_id: 13, weight_kg: 200, container: "ibc", containers: 1, method: "ระบบบำบัดน้ำเสียภายในโรงงาน", treatToken: tok, ...at(2) }));
  assert.strictEqual(r.balance_kg, 300);
  const row = ms.data.find(x => x[1] === r.doc_no);
  assert.strictEqual(row[16], "ระบบบำบัดน้ำเสียภายในโรงงาน");
  assert.strictEqual(ms.data[1][16] || "", ""); // the old row reads the new column as empty
});

/* ---------- password gating ---------- */
test("reset password: missing property, wrong password, logs need POST + password", () => {
  const saved = state.props.WASTE_RESET_PASSWORD;
  delete state.props.WASTE_RESET_PASSWORD;
  const r = err(post("movement_delete", { id: 1 }), /WASTE_RESET_PASSWORD/);
  assert.strictEqual(r.passwordError, true);
  state.props.WASTE_RESET_PASSWORD = saved;
  err(postRaw({ action: "movement_delete", id: 1 }), /รหัสไม่ถูกต้อง/);
  err(postRaw({ action: "type_save", resetPassword: "nope" }), /รหัสไม่ถูกต้อง/);
  err(postRaw({ action: "logs" }, null), /รหัสไม่ถูกต้อง/);
  err(postRaw({ action: "logs_export", resetPassword: "x" }, null), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(okData(postRaw({ action: "verifyResetPassword", resetPassword: RPW }, null)).valid, true);
  err(postRaw({ action: "verifyResetPassword", resetPassword: "bad" }, null), /รหัสไม่ถูกต้อง/);
  assert.ok(state.sleeps >= 5000);
  clearFails();
});

test("reset password lockout: 10 wrong attempts lock even the right password (15 min, cache counter)", () => {
  clearFails();
  for (let i = 0; i < 10; i++) err(postRaw({ action: "verifyResetPassword", resetPassword: "wrong" + i }, null), /รหัสไม่ถูกต้อง/);
  const r = err(postRaw({ action: "verifyResetPassword", resetPassword: RPW }, null), /หลายครั้งเกินไป/);
  assert.strictEqual(r.passwordError, true);
  err(post("movement_delete", { id: 1 }), /หลายครั้งเกินไป/);
  err(post("logs"), /หลายครั้งเกินไป/);
  assert.strictEqual(state.cache.WASTE_RESET_FAILURES, "10");
  assert.ok(state.cachePuts.includes(900));
  clearFails();
  assert.strictEqual(okData(postRaw({ action: "verifyResetPassword", resetPassword: RPW }, null)).valid, true);
});

test("password never echoed in responses, logs or audit rows; no secrets in the source", () => {
  const all = JSON.stringify(state.logs) + JSON.stringify(sheet("audit_logs").data);
  assert.ok(!all.includes(RPW));
  const src = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  assert.ok(!/RESET_PASSWORD\s*=\s*["'][^"']+["']\s*;?\s*$/m.test(src.replace(/RESET_PASSWORD_PROPERTY = "WASTE_RESET_PASSWORD"/, "")));
  assert.ok(!src.includes("dev-reset") && !src.includes(RPW));
  assert.ok(!/AKfycb|https:\/\/hooks\./.test(src));
});

/* ---------- edit / delete integrity ---------- */
test("movement_update: edits with replay check; lowering a store that a later ship needs is refused", () => {
  const st = okData(mv("store", 9, 100, { container: "bigbag", containers: 2, ...at(8) }));
  const sh = okData(mv("ship", 9, 80, { container: "bigbag", containers: 1, ...at(4) }));
  let r = err(post("movement_update", { id: st.id, doc_no: st.doc_no, weight_kg: 79.99 }), /แก้ไขไม่ได้.*ติดลบที่รายการ SH-/);
  assert.strictEqual(r.code, "NEGATIVE_BALANCE");
  err(post("movement_update", { id: st.id, doc_no: st.doc_no, containers: 0 }), /จำนวนเต็มตั้งแต่ 1/);
  err(post("movement_update", { id: sh.id, doc_no: sh.doc_no, containers: 3 }), /แก้ไขไม่ได้/);
  err(post("movement_update", { id: st.id, ...at(3) }), /แก้ไขไม่ได้/); // moving the store after the ship
  err(post("movement_update", { id: st.id, type_id: 10 }), /แก้ไขไม่ได้/); // the ship of type 9 loses its stock
  err(post("movement_update", { id: st.id, doc_no: "ST-000000-9999", weight_kg: 120 }), /ไม่ตรงกัน/);
  err(post("movement_update", { id: 99999, weight_kg: 1 }), /ไม่พบรายการ/);
  err(post("movement_update", { id: st.id, kind: "ship" }), /เปลี่ยนชนิดรายการไม่ได้/);
  assert.strictEqual(okData(post("movement_update", { id: st.id, doc_no: st.doc_no })).changed, false);
  const ok = okData(post("movement_update", { id: st.id, doc_no: st.doc_no, weight_kg: "120.5", note: "แก้แล้ว", location: "B" }));
  assert.strictEqual(ok.changed, true);
  assert.strictEqual(bal(9).balance_kg, 40.5);
  const row = sheet("movements").data.find(x => x[1] === st.doc_no);
  assert.strictEqual(row[4], 120.5);
  assert.strictEqual(row[10], "แก้แล้ว");
  assert.strictEqual(row[15], ACTOR);
  assert.ok(row[14]);
  const l = lastLog();
  assert.strictEqual(l.category, "edit");
  assert.ok(l.detail.includes("เดิม:") && l.detail.includes("100 กก.") && l.detail.includes("120.5 กก."));
  // a ship edited above the stock is refused, within it accepted
  err(post("movement_update", { id: sh.id, weight_kg: 120.51 }), /แก้ไขไม่ได้/);
  okData(post("movement_update", { id: sh.id, weight_kg: 120.5, manifest_no: "MN-77" }));
  assert.strictEqual(bal(9).balance_kg, 0);
});

test("movement_delete: a store that a ship depends on cannot be deleted; deleting the ship first works", () => {
  const st = okData(mv("store", 4, 3, { ...at(9) }));
  const sh = okData(mv("ship", 4, 3, { ...at(8) }));
  const before = dataRows("movements");
  err(post("movement_delete", { id: st.id, doc_no: st.doc_no }), /ลบไม่ได้.*ติดลบ/);
  assert.strictEqual(dataRows("movements"), before);
  const r = okData(post("movement_delete", { id: sh.id, doc_no: sh.doc_no }));
  assert.ok(/ลบรายการ/.test(r.message));
  assert.strictEqual(bal(4).balance_kg, 3);
  okData(post("movement_delete", { id: st.id, doc_no: st.doc_no }));
  assert.strictEqual(bal(4).balance_kg, 0);
  assert.strictEqual(dataRows("movements"), before - 2);
  const l = lastLog();
  assert.strictEqual(l.action, "movement_delete");
  assert.strictEqual(l.ref, st.doc_no);
  // a deleted doc number is never handed out again
  const again = okData(mv("ship", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.notStrictEqual(again.doc_no, sh.doc_no);
});

/* ---------- waste type settings ---------- */
test("type_save: validation, add, edit, duplicate check, deactivate (no new store, ship still allowed)", () => {
  err(post("type_save", { code: "12345", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r" }), /6 หลัก/);
  err(post("type_save", { code: "123456", name: "", permit_tons: 1, mgmt_code: "011", receiver: "r" }), /ชื่อของเสีย/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: "abc", mgmt_code: "011", receiver: "r" }), /ตัวเลข/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "11", receiver: "r" }), /3 หลัก/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "" }), /ผู้รับดำเนินการ/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", receiver_reg: "12a" }), /ทะเบียนผู้รับ/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", icon: "evil\"><script>" }), /ไอคอน/);
  err(post("type_save", { code: "123456", name: "x", permit_tons: 1, mgmt_code: "011", receiver: "r", color: "red" }), /สี/);
  err(post("type_save", { code: "150101", name: "กระดาษลูกฟูก กระดาษลัง", permit_tons: 5, mgmt_code: "011", receiver_reg: "10110020125675", receiver: "บจก. มั่นคงยั่งยืน รีไซเคิล" }), /อยู่แล้ว/);
  const a = okData(post("type_save", { code: "130205", name: "น้ำมันเครื่องใช้แล้ว", permit_tons: "1.25", mgmt_code: "061", receiver_reg: "123", receiver: "บจก. ทดสอบ", icon: "fa-oil-can", color: "#0ea5e9" }));
  assert.strictEqual(a.created, true);
  assert.strictEqual(a.id, 14);
  let t = boot().types.find(x => x.id === 14);
  assert.strictEqual(t.permit_tons, 1.25);
  assert.strictEqual(t.color, "#0EA5E9");
  assert.strictEqual(t.sort, 14);
  assert.strictEqual(t.updated_by, ACTOR);
  assert.strictEqual(bal(14).permit_kg, 1250);
  assert.strictEqual(lastLog().category, "setting");
  okData(mv("store", 14, 10, { ...at(2) }));
  const e = okData(post("type_save", { id: 14, code: "130205", name: "น้ำมันเครื่องใช้แล้ว", permit_tons: 2, mgmt_code: "061", receiver_reg: "123", receiver: "บจก. ทดสอบ", active: 0 }));
  assert.strictEqual(e.created, false);
  t = boot().types.find(x => x.id === 14);
  assert.strictEqual(t.active, 0);
  assert.strictEqual(t.permit_tons, 2);
  assert.strictEqual(t.icon, "fa-oil-can"); // kept when not sent
  err(mv("store", 14, 1), /ปิดใช้งาน/);
  okData(mv("ship", 14, 10, { ...at(1) }));
  assert.ok(/เดิม:.*1\.25 ตัน.*ใช้งาน → ใหม่:.*2 ตัน.*ปิดใช้งาน/.test(logs({ cat: "setting" })[0].detail));
  err(post("type_save", { id: 999, code: "130205", name: "x", permit_tons: 1, mgmt_code: "061", receiver: "r" }), /ไม่พบ/);
  assert.strictEqual(dataRows("waste_types"), 14);
});

/* ---------- read cache / versions ---------- */
test("read cache: repeated reads served from CacheService; every write bumps the data version", () => {
  boot();
  const n0 = state.sheetReads.movements || 0;
  boot(); boot(); okData(get("history"));
  assert.strictEqual(state.sheetReads.movements || 0, n0, "cached reads must not read the sheet");
  const v0 = state.cache.WASTE_V_movements;
  assert.ok(v0);
  const ev0 = state.events.length;
  okData(mv("store", 1, 2, { container: "bigbag", containers: 1, ...at(1) }));
  const v1 = state.cache.WASTE_V_movements;
  assert.notStrictEqual(v1, v0);
  const evs = state.events.slice(ev0);
  assert.ok(evs.indexOf("flush") < evs.lastIndexOf("version:movements"), "flush before the version bump");
  const b = bal(1);
  assert.strictEqual(b.balance_kg, Math.round((1234.6 + 2 - 1) * 100) / 100); // fresh data after the write
  // waste_types version unchanged by a movement write
  const vt = state.cache.WASTE_V_waste_types;
  okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(state.cache.WASTE_V_waste_types, vt);
  // a rejected write changes nothing but the audit_logs version
  const vm0 = state.cache.WASTE_V_movements;
  err(mv("ship", 4, 999));
  assert.strictEqual(state.cache.WASTE_V_movements, vm0);
});

test("read cache: clearReadCache() forces fresh reads (hand edits); cache outage still answers", () => {
  boot();
  const n0 = state.sheetReads.waste_types || 0;
  sheet("waste_types").data[1][2] = "กระดาษลัง (แก้ด้วยมือ)";
  assert.notStrictEqual(boot().types[0].name, "กระดาษลัง (แก้ด้วยมือ)");
  ctx.clearReadCache();
  assert.strictEqual(boot().types[0].name, "กระดาษลัง (แก้ด้วยมือ)");
  assert.ok((state.sheetReads.waste_types || 0) > n0);
  sheet("waste_types").data[1][2] = "กระดาษลูกฟูก กระดาษลัง";
  ctx.clearReadCache();
  state.cacheDown = true;
  try {
    assert.strictEqual(boot().types[0].name, "กระดาษลูกฟูก กระดาษลัง");
    okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  } finally { state.cacheDown = false; }
  ctx.clearReadCache();
});

test("cache values stay under 100 KB per key (chunked) with many rows", () => {
  for (let i = 0; i < 40; i++) okData(mv("store", 6, 1, { note: "ข้อความภาษาไทยยาว ๆ ".repeat(10), ...at(1) }));
  ctx.clearReadCache();
  okData(get("history"));
  okData(get("history")); // asserts inside the mock putAll: <= 100 KB, key <= 250 chars
});

test("batch reads and write responses carrying post-write reads", () => {
  const r = okData(get("batch", { reads: JSON.stringify([{ action: "bootstrap" }, { action: "dashboard" }, { action: "history", kind: "ship" }]) }));
  assert.strictEqual(r.results.length, 3);
  assert.ok(r.results[0].types && r.results[1].kpi && r.results[2].items.every(m => m.kind === "ship"));
  err(get("batch", { reads: JSON.stringify([{ action: "logs" }]) }), /Unknown action/);
  err(get("batch", { reads: "[]" }), /reads/);
  err(get("batch", { reads: JSON.stringify(new Array(7).fill("bootstrap")) }), /reads/);
  const w = okData(mv("store", 2, 5, { container: "ibc", containers: 1, ...at(1), reads: [{ action: "bootstrap" }] }));
  assert.strictEqual(w.reads.length, 1);
  assert.strictEqual(w.reads[0].balance[2].balance_kg, bal(2).balance_kg);
  err(mv("store", 2, 5, { reads: [{ action: "logs" }] }), /Unknown action/);
});

test("LockService: writes take and release the script lock; lock timeout answers politely", () => {
  assert.strictEqual(state.locks, state.unlocks);
  const n = state.locks;
  okData(mv("store", 1, 1, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(state.locks, n + 1);
  assert.strictEqual(state.unlocks, n + 1);
  state.lockFail = true;
  try { err(mv("store", 1, 1), /กำลังบันทึกรายการอื่น/); } finally { state.lockFail = false; }
  err(mv("ship", 4, 1000));
  assert.strictEqual(state.locks, state.unlocks, "released after a failed write");
});

test("keepWarm / installKeepWarmTrigger / removeKeepWarmTrigger", () => {
  const w = ctx.keepWarm();
  assert.strictEqual(w.warmed, true);
  const r = ctx.installKeepWarmTrigger();
  assert.strictEqual(r.installed, true);
  assert.strictEqual(r.everyMinutes, 10);
  ctx.installKeepWarmTrigger();
  assert.strictEqual(state.triggers.filter(t => t.handler === "keepWarm").length, 1);
  assert.ok(state.cachePutAlls.includes(900), "keepWarm stores with the longer TTL");
  assert.strictEqual(ctx.removeKeepWarmTrigger().removed, 1);
  assert.strictEqual(state.triggers.length, 0);
});

/* ---------- audit log ---------- */
test("audit log: POST with password, filters by category / search / date, CSV", () => {
  const all = logs();
  assert.ok(all.length > 20);
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].id > all[i].id);
  assert.ok(logs({ cat: "edit" }).every(l => l.category === "edit"));
  assert.ok(logs({ q: "MN-1" }).length >= 1);
  assert.strictEqual(logs({ from: "2001-01-01", to: "2001-01-01" }).length, 0);
  const c = okData(post("logs_export", { cat: "setting" }));
  assert.ok(/^waste_log_\d{8}_\d{6}\.csv$/.test(c.filename));
  assert.ok(c.csv.startsWith("ลำดับ,วันที่,เวลา,หมวด,เลขที่อ้างอิง,รายละเอียด,ผู้ทำรายการ\n"));
  assert.ok(c.csv.includes("ตั้งค่า"));
});

/* ---------- reset data ---------- */
test("reset_data: password + name required; wipes movements and logs, keeps waste types; numbers restart", () => {
  err(postRaw({ action: "reset_data", resetPassword: "bad" }), /รหัสไม่ถูกต้อง/);
  clearFails();
  assert.ok(dataRows("movements") > 0);
  const r = okData(post("reset_data", { reads: [{ action: "bootstrap" }] }));
  assert.ok(r.removed.movements > 0 && r.removed.audit_logs > 0);
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("waste_types"), 14);
  const l = logs();
  assert.strictEqual(l.length, 1);
  assert.strictEqual(l[0].action, "reset_data");
  assert.ok(Object.values(r.reads[0].balance).every(b => b.balance_kg === 0));
  const s = okData(mv("store", 1, 5, { container: "bigbag", containers: 1, ...at(1) }));
  assert.strictEqual(s.doc_no, "ST-" + YMD + "-0001");
});

test("sheet storage: numbers / codes / dates keep their formats (no Date conversion, no formulas)", () => {
  const types = sheet("waste_types").data;
  assert.strictEqual(types[5][1], "080409");
  assert.strictEqual(typeof types[7][3], "number"); // 0.5 t
  const mvRow = sheet("movements").data[1];
  assert.ok(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:00$/.test(mvRow[11]));
  assert.strictEqual(typeof mvRow[4], "number");
});

/* ---------- frontend ---------- */
const readDoc = f => fs.readFileSync(path.join(ROOT, f), "utf8");
test("docs JS files parse; config shape; GET with _=ts; POST text/plain; actor sent; UTF-8 BOM CSV", () => {
  for (const f of ["docs/config.js", "docs/assets/app.js", "test/dev-server.js"]) new Function(readDoc(f).replace(/^#!.*\n/, ""));
  const cfg = {};
  new Function("window", readDoc("docs/config.js"))(cfg);
  assert.deepStrictEqual(Object.keys(cfg.WASTE_CONFIG), ["apiUrl"]);
  assert.ok(cfg.WASTE_CONFIG.apiUrl === "" || /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(cfg.WASTE_CONFIG.apiUrl));
  const app = readDoc("docs/assets/app.js");
  assert.ok(app.includes("window.WASTE_CONFIG"));
  assert.ok(app.includes("u.searchParams.set('_', Date.now())"));
  assert.ok(app.includes("'Content-Type': 'text/plain;charset=UTF-8'"));
  assert.ok(/actor: S\.user/.test(app) && app.includes("store('get', 'wasteUser')"));
  assert.ok(app.includes("'\\uFEFF' + r.csv"));
  for (const a of ["store", "ship", "opening", "movement_update", "movement_delete", "type_save", "reset_data"]) assert.ok(new RegExp("NAME_REQUIRED = \\[[^\\]]*'" + a + "'").test(app), a);
  for (const a of ["'movement_update'", "'movement_delete'", "'type_save'", "'verifyResetPassword'", "'logs_export'", "'batch'", "action: 'reset_data', resetPassword"]) assert.ok(app.includes(a), a);
  // the admin password is kept in memory only
  assert.ok(!/(local|session)Storage\.setItem\([^)]*([Pp]ass|[Rr]eset)/.test(app));
  assert.ok(!/store\('set', *'[^']*([Pp]ass|[Rr]eset)/.test(app));
  assert.ok(!/RC_PREFIX[^\n]*logs/.test(app));
  // treatment view: token only in POST bodies / sessionStorage, never in a URL or the localStorage read cache
  assert.ok(app.includes("sessionStorage.setItem(TT_KEY") && !/localStorage\.setItem\(TT_KEY/.test(app));
  assert.ok(!/searchParams\.set\([^)]*treat/i.test(app) && app.includes("if (S.tt) body.treatToken = S.tt"));
  assert.ok(app.includes("if (S.tt) return null;") && app.includes("if (!d || S.tt) return;"));
  for (const a of ["'treat_unlock'", "'treat_lock'", "treat: () => txForm('treat')", "k !== 'treat' || treatOn()"]) assert.ok(app.includes(a), a);
  // client-side checks mirror the server (2 decimals, ship <= balance)
  assert.ok(app.includes("/^\\d+(\\.\\d{1,2})?$/"));
  assert.ok(app.includes("เกินยอดคงเหลือ"));
});

test("index.html / style.css: brand, rainbow credit (7 colours, 21 s), rotating icons, reduced motion, CDNs", () => {
  const html = readDoc("docs/index.html"), css = readDoc("docs/assets/style.css");
  assert.ok(html.includes("<title>Balance Waste"));
  assert.strictEqual((html.match(/class="programming-credit">Programming by Duongruthai_Milk</g) || []).length, 2); // sidebar + footer
  assert.ok(css.includes(".programming-credit{font-size:11px;font-weight:700;letter-spacing:.02em;margin-top:2px;animation:rainbowText 21s ease-in-out infinite}"));
  assert.ok(css.includes("@keyframes rainbowText{0%,100%{color:#ef4444}14.28%{color:#f97316}28.56%{color:#eab308}42.84%{color:#22c55e}57.12%{color:#06b6d4}71.4%{color:#3b82f6}85.68%{color:#8b5cf6}}"));
  assert.ok(/prefers-reduced-motion:reduce[\s\S]*\.programming-credit\{animation:none/.test(css));
  assert.ok(css.includes("conic-gradient") && css.includes("@keyframes spin{to{transform:rotate(360deg)}}"));
  assert.ok(css.includes(".ccard.sel{transform:rotateY(180deg)}") && css.includes(".orbit .dot"));
  assert.ok(css.includes(".needle-gauge .needle"));
  // background animations: transform / opacity only (no animated blur)
  const kf = name => (css.match(new RegExp("@keyframes " + name + "\\{[^@]*?\\}\\}")) || [""])[0];
  for (const k of ["float", "drift", "spin", "pulse"]) { assert.ok(kf(k), k); assert.ok(!/filter|blur|width|height|top|left/.test(kf(k)), k + " animates layout/blur"); }
  assert.ok(html.includes("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css"));
  assert.ok(html.includes("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"));
  assert.ok(html.includes("family=Kanit") && html.includes("Prompt"));
  assert.ok(html.includes('id="userChip"') && html.includes('data-page="logs"') && html.includes('data-page="settings"'));
  assert.ok(/data-page="treat"[^>]*hidden/.test(html) && html.includes('id="treatLock"'), "treatment menu hidden by default + lock control");
  assert.ok(css.includes("--treat:#7C3AED") && css.includes("@keyframes brew") && css.includes(".btn-treat"));
  assert.ok(html.includes("assets/app.js?v=3") && html.includes("assets/style.css?v=3"));
  assert.ok(html.includes("82260007625693") && html.includes("2569-8409"));
  assert.ok(!/<script[^>]+src="(?!https:\/\/cdnjs\.cloudflare\.com\/|config\.js|assets\/app\.js)/.test(html));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);

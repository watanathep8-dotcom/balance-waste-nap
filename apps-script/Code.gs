/**
 * Balance Waste — ระบบบริหารจัดการของเสีย (Waste Management)
 * บริษัท แน็ป นิวตริซายส์ (ประเทศไทย) จำกัด
 * Backend for Google Apps Script (data in Google Sheets).
 *
 * Setup (once): run setupSystem() from the Apps Script editor. It creates the
 * "Balance Waste Database" spreadsheet (one sheet per table, header row on
 * top), seeds the 13 permitted waste types (สก.2 ปี 2569) into "waste_types"
 * and stores the spreadsheet ID in Script Properties (WASTE_SPREADSHEET_ID).
 * Then deploy as a web app (Execute as: me, Who has access: Anyone).
 *
 * No login: reads and the normal writes (store, ship, opening) need no
 * password. The name typed in the top-right of the page is sent as "actor" and
 * recorded on every movement and audit_logs row ("ผู้บันทึก").
 *
 * Script property WASTE_RESET_PASSWORD protects, with one lockout counter
 * (10 wrong / 15 min): reset_data (wipes movements and audit_logs; waste
 * types are kept), editing / deleting movements (movement_update,
 * movement_delete), editing / adding waste types (type_save) and reading the
 * audit log (logs, logs_export). The password is never in the source.
 *
 * Ledger integrity: every write that changes movements (ship, edit, delete,
 * also store/opening) is checked by replaying ALL movements in time order
 * (validateLedger_): the weight of every waste type and the container count of
 * every (waste type, container) pair must never go below zero at any moment.
 * Nothing is written unless the replay passes.
 *
 * Optional: run installKeepWarmTrigger() once from the editor; a 10-minute
 * trigger then runs keepWarm() (read-only) to keep the read cache hot.
 *
 * Transport:
 *   GET  ?action=<read action>&...params -> {ok:true,data} | {ok:false,error}
 *   POST text/plain JSON {action, actor, ...fields}  (no CORS preflight)
 */

var SPREADSHEET_ID_PROPERTY = "WASTE_SPREADSHEET_ID";
var SPREADSHEET_NAME = "Balance Waste Database";

// Reset password (Script Property only) with its own lockout counter.
var RESET_ACTION = "reset_data";
var RESET_PASSWORD_PROPERTY = "WASTE_RESET_PASSWORD";
var RESET_FAIL_CACHE_KEY = "WASTE_RESET_FAILURES";
var PASSWORD_MAX_FAILURES = 10;
var PASSWORD_LOCK_SECONDS = 900;
var RESET_TABLES = ["movements", "audit_logs"]; // waste_types are kept
var RESET_PASSWORD_ACTIONS = [RESET_ACTION, "movement_update", "movement_delete", "type_save"];
// Reading the audit log also needs the reset password (POST only).
var LOG_ACTIONS = ["logs", "logs_export"];

// Asia/Bangkok has no daylight saving time: fixed +07:00.
var TZ_OFFSET_MS = 7 * 60 * 60 * 1000;

var COMPANY = {
  name: "บริษัท แน็ป นิวตริซายส์ (ประเทศไทย) จำกัด",
  factoryReg: "82260007625693",
  permitNo: "สก.2 เลขที่ 2569-8409",
  permitYear: 2569
};

var CONTAINERS = { drum200: "ถัง 200 ลิตร", ibc: "IBC", bigbag: "Bigbag" };
var CONTAINER_KEYS = ["drum200", "ibc", "bigbag"];
var KIND_NAME = { store: "บันทึกจัดเก็บ", ship: "ส่งกำจัด / นำออก", opening: "ยอดยกมา" };
var KINDS = ["store", "ship", "opening"];
var PREFIX = { store: "ST", ship: "SH", opening: "OB" };
var ACT_NAME = {
  store: "บันทึกจัดเก็บ", ship: "ส่งกำจัด", opening: "ยอดยกมา",
  movement_update: "แก้ไขรายการ", movement_delete: "ลบรายการ", type_save: "ตั้งค่าประเภทของเสีย"
};
var LOG_CAT_NAME = { waste: "ของเสีย", edit: "แก้ไข / ลบ", setting: "ตั้งค่า", warn: "ถูกปฏิเสธ" };

var QUOTA_WARN_PCT = 80;        // >= 80 % of the permitted quantity: warning
var STORAGE_ALERT_DAYS = 90;    // oldest stored lot (FIFO) older than this: alert
var MAX_WEIGHT_KG = 1000000;    // one movement
var MAX_PERMIT_TONS = 100000;
var FUTURE_SLACK_MS = 10 * 60 * 1000; // client clock may be a few minutes ahead

/* Column definitions. Types: int, num (decimal), str, dt ("YYYY-MM-DD HH:MM:SS").
 * A trailing "?" marks columns that are null when empty. New columns are only
 * ever APPENDED (older sheets get them added on first use, see table_). */
var SCHEMA = {
  waste_types: [
    ["id", "int"], ["code", "str"], ["name", "str"], ["permit_tons", "num"], ["mgmt_code", "str"],
    ["receiver_reg", "str"], ["receiver", "str"], ["icon", "str"], ["color", "str"], ["active", "int"],
    ["sort", "int"], ["updated_at", "dt?"], ["updated_by", "str"]
  ],
  movements: [
    ["id", "int"], ["doc_no", "str"], ["kind", "str"], ["type_id", "int"], ["weight_kg", "num"],
    ["container", "str"], ["containers", "int"], ["location", "str"], ["manifest_no", "str"], ["vehicle", "str"],
    ["note", "str"], ["moved_at", "dt"], ["created_at", "dt"], ["actor", "str"],
    ["updated_at", "dt?"], ["updated_by", "str"]
  ],
  audit_logs: [
    ["id", "int"], ["category", "str"], ["action", "str"], ["ref", "str"], ["detail", "str"],
    ["actor", "str"], ["created_at", "dt"]
  ]
};
var TABLE_ORDER = ["waste_types", "movements", "audit_logs"];
var AUTO_CREATE_TABLES = [];

// [code, name, permit tons, management code, receiver registration, receiver, icon, colour]
// From Waste_Summary_Report_NAP_Nutriscience_2569.xlsx (สก.2 เลขที่ 2569-8409, ปี 2569).
var RCV_MKY = "บจก. มั่นคงยั่งยืน รีไซเคิล";
var RCV_BWG = "บมจ. เบตเตอร์ เวิลด์ กรีน";
var RCV_GENCO = "บมจ. บริหารและพัฒนาเพื่อการอนุรักษ์สิ่งแวดล้อม";
var SEED_TYPES = [
  ["150101", "กระดาษลูกฟูก กระดาษลัง", 20, "011", "10110020125675", RCV_MKY, "fa-box-open", "#B7791F"],
  ["150102", "ถุงพลาสติก ถังพลาสติก", 30, "011", "10110020125675", RCV_MKY, "fa-bottle-water", "#0EA5E9"],
  ["170405", "เศษเหล็ก", 10, "011", "10110020125675", RCV_MKY, "fa-gears", "#64748B"],
  ["150111", "กระป๋องสเปรย์", 2, "073", "20190300225401", RCV_BWG, "fa-spray-can", "#EF4444"],
  ["080409", "กากกาว", 10, "042", "10190000825494", RCV_BWG, "fa-fill-drip", "#A855F7"],
  ["150202", "วัสดุปนเปื้อน", 30, "042", "10190000825494", RCV_BWG, "fa-biohazard", "#F97316"],
  ["160215", "หลอดไฟ", 0.5, "073", "20190300225401", RCV_BWG, "fa-lightbulb", "#EAB308"],
  ["160213", "อุปกรณ์อิเล็กทรอนิกส์", 1, "073", "20190300225401", RCV_BWG, "fa-microchip", "#6366F1"],
  ["150102", "บรรจุภัณฑ์พลาสติก", 20, "046", "10190000825494", RCV_BWG, "fa-jug-detergent", "#06B6D4"],
  ["191201", "เศษสติ๊กเกอร์", 20, "046", "10190000825494", RCV_BWG, "fa-note-sticky", "#EC4899"],
  ["190813", "ตะกอนจากระบบบำบัดน้ำเสีย", 10, "073", "10700001425472", RCV_GENCO, "fa-water", "#8B5E3C"],
  ["160215", "หลอดไฟ", 0.5, "073", "10700001425472", RCV_GENCO, "fa-lightbulb", "#F59E0B"],
  ["161001", "น้ำล้างเครื่องจักร (น้ำแรก)", 40, "065", "10103300225380", RCV_GENCO, "fa-faucet-drip", "#14B8A6"]
];

var READ_ACTIONS = ["bootstrap", "dashboard", "history", "export", "logs", "logs_export", "batch"];
// "batch": several reads in one request; writes accept the same optional
// "reads" list and return the post-write results as data.reads.
var BATCH_ACTIONS = ["bootstrap", "dashboard", "history"];
var BATCH_MAX = 6;

/* Read cache (CacheService), keyed by a per-sheet data version that every
 * write bumps right after its changes are flushed (see flush_). Writes always
 * read the sheets live under the script lock. Hand edits in the spreadsheet
 * show up after READ_CACHE_TTL seconds, or at once after clearReadCache(). */
var CACHED_TABLES = ["waste_types", "movements", "audit_logs"];
var VERSION_PREFIX = "WASTE_V_";
var VERSION_TTL = 21600;
var RC_PREFIX = "WASTE_RC_";
var READ_CACHE_TTL = 600;
var KEEP_WARM_TTL = 900;
var KEEP_WARM_MINUTES = 10;
var RC_CHUNK = 30000;
var RC_MAX_CHUNKS = 30;
var WRITE_ACTIONS = ["store", "ship", "opening"];
var POST_ONLY_ACTIONS = ["verifyResetPassword"].concat(WRITE_ACTIONS, RESET_PASSWORD_ACTIONS, LOG_ACTIONS);

var MAX_PERSON = 100;
var MAX_TEXT = 255;

/* ===================== web app entry points ===================== */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = String(params.action || "");
  try {
    if (!action) return jsonResponse_({ ok: true, data: { service: "Balance Waste API" } });
    if (POST_ONLY_ACTIONS.indexOf(action) !== -1) throw new Error("คำสั่งนี้ต้องส่งแบบ POST");
    if (READ_ACTIONS.indexOf(action) === -1) fail_("Unknown action");
    resetRequest_("");
    return jsonResponse_({ ok: true, data: runRead_(action, params) });
  } catch (error) {
    return errorResponse_(error);
  }
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (parseError) {
    return jsonResponse_({ ok: false, error: "ข้อมูลที่ส่งมาไม่ใช่ JSON ที่ถูกต้อง" });
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) payload = {};
  var action = String(payload.action || "");

  var reads = null;
  try {
    if (action === "verifyResetPassword") {
      return jsonResponse_({ ok: true, data: verifyResetPassword(payload.resetPassword) });
    }
    if (READ_ACTIONS.indexOf(action) !== -1) {
      if (LOG_ACTIONS.indexOf(action) !== -1) assertResetPassword_(payload.resetPassword);
      resetRequest_("");
      return jsonResponse_({ ok: true, data: runRead_(action, payload) });
    }
    if (WRITE_ACTIONS.indexOf(action) === -1 && RESET_PASSWORD_ACTIONS.indexOf(action) === -1) fail_("Unknown action");
    if (payload.reads != null && payload.reads !== "") reads = parseReads_(payload.reads);
    // Every change must say who made it (name set in the page header).
    if (!actorIn_(payload)) fail_("กรุณาระบุชื่อผู้ใช้งาน (มุมขวาบน) ก่อนบันทึก", "NAME_REQUIRED");
    // Checked before taking the lock so guessing (1 s delay) never blocks writers.
    if (RESET_PASSWORD_ACTIONS.indexOf(action) !== -1) assertResetPassword_(payload.resetPassword);
  } catch (error) {
    return errorResponse_(error);
  }
  return runWrite_(action, payload, reads);
}

function actorIn_(input) {
  var v = input.actor;
  return mbSubstr_(phpTrim_(v !== null && typeof v === "object" ? "" : str_(v)), 0, MAX_PERSON);
}

function runWrite_(action, input, reads) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockError) {
    return jsonResponse_({ ok: false, error: "ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่อีกครั้ง" });
  }
  var actor = actorIn_(input);
  var result;
  try {
    resetRequest_(actor);
    try {
      result = handleWrite_(action, input);
      flush_();
      if (reads) versions_();
    } catch (error) {
      // Nothing buffered by the failed action is written (acts as a rollback).
      var actName = ACT_NAME[action];
      if (actName && action !== RESET_ACTION) {
        try {
          resetRequest_(actor);
          audit_("warn", action, "ปฏิเสธ" + actName + ": " + errorMessage_(error), "");
          flush_();
        } catch (ignored) {}
      }
      return errorResponse_(error);
    }
  } finally {
    lock.releaseLock();
  }
  if (reads) result.reads = afterReads_(reads);
  return jsonResponse_({ ok: true, data: result });
}

/* ===================== reads: single, batch, after a write ===================== */

function runRead_(action, p) {
  REQ_.cacheReads = true;
  return handleRead_(action, p);
}

function parseReads_(raw) {
  var list = raw;
  if (typeof raw === "string") {
    try { list = JSON.parse(raw); } catch (e) { fail_("reads ไม่ถูกต้อง"); }
  }
  if (!Array.isArray(list) || !list.length || list.length > BATCH_MAX) fail_("reads ไม่ถูกต้อง");
  return list.map(function (r) {
    if (typeof r === "string") r = { action: r };
    if (!r || typeof r !== "object" || Array.isArray(r)) fail_("reads ไม่ถูกต้อง");
    if (BATCH_ACTIONS.indexOf(String(r.action || "")) === -1) fail_("Unknown action");
    return r;
  });
}

function actionBatch_(p) {
  var reads = parseReads_(p.reads);
  return { results: reads.map(function (r) { return handleRead_(r.action, r); }) };
}

function afterReads_(reads) {
  try {
    REQ_.cacheReads = true;
    return reads.map(function (r) { return handleRead_(r.action, r); });
  } catch (error) {
    return null;
  }
}

/* ===================== reset password ===================== */

function verifyResetPassword(password) {
  assertResetPassword_(password);
  return { valid: true };
}

function assertResetPassword_(password) {
  checkPassword_(password, RESET_PASSWORD_PROPERTY, RESET_FAIL_CACHE_KEY,
    "ยังไม่ได้ตั้งรหัสผ่านผู้ดูแลใน Script Properties (" + RESET_PASSWORD_PROPERTY +
    ") กรุณาให้ผู้ดูแลระบบตั้งค่าที่ Project Settings > Script properties ก่อน");
}

// Property must be set; >10 wrong attempts per 15 min locks the check; every
// wrong attempt costs 1 s.
function checkPassword_(password, property, failKey, missingMessage) {
  var expected = PropertiesService.getScriptProperties().getProperty(property) || "";
  if (!expected) throw passwordError_(missingMessage);
  var cache = CacheService.getScriptCache();
  var failures = Number(cache.get(failKey) || 0);
  if (failures >= PASSWORD_MAX_FAILURES) {
    throw passwordError_("ใส่รหัสผิดหลายครั้งเกินไป กรุณารอ 15 นาที / Too many wrong attempts. Try again in 15 minutes.");
  }
  if (String(password == null ? "" : password) !== expected) {
    cache.put(failKey, String(failures + 1), PASSWORD_LOCK_SECONDS);
    Utilities.sleep(1000);
    throw passwordError_("รหัสไม่ถูกต้อง / Incorrect password.");
  }
}

function passwordError_(message) {
  var error = new Error(message);
  error.passwordError = true;
  return error;
}

function cache_() {
  return CacheService.getScriptCache();
}

/* ===================== router ===================== */

function handleRead_(action, p) {
  switch (action) {
    case "bootstrap": return actionBootstrap_();
    case "dashboard": return actionDashboard_();
    case "history": return { items: queryHistory_(p) };
    case "export": return exportHistory_(queryHistory_(p));
    case "logs": return { items: queryLogs_(p) };
    case "logs_export": return exportLogs_(queryLogs_(p));
    case "batch": return actionBatch_(p);
  }
  fail_("Unknown action");
}

function handleWrite_(action, input) {
  switch (action) {
    case "store": return actionMove_("store", input);
    case "ship": return actionMove_("ship", input);
    case "opening": return actionMove_("opening", input);
    case RESET_ACTION: return actionResetData_();
    case "movement_update": return actionMovementUpdate_(input);
    case "movement_delete": return actionMovementDelete_(input);
    case "type_save": return actionTypeSave_(input);
  }
  fail_("Unknown action");
}

/* ===================== balance math =====================
 * Weights are summed in integer hundredths of a kg ("cents") so 0.1 + 0.2
 * never drifts. stored = store + opening; balance = stored - shipped.
 */

function cents_(kg) {
  return Math.round(Number(kg || 0) * 100);
}

function kgOf_(cents) {
  return Math.round(cents) / 100;
}

function sortedTypes_() {
  return sortBy_(table_("waste_types").rows.slice(), function (a, b) { return a.sort - b.sort || a.id - b.id; });
}

// Per waste type: totals, containers in storage, this year's shipments vs the
// permitted quantity and the date of the oldest lot still in storage (FIFO).
function balances_() {
  var now = nowParts_();
  var year = now.date.slice(0, 4);
  var out = {};
  var fifo = {};
  table_("waste_types").rows.forEach(function (t) {
    out[t.id] = {
      type_id: t.id, stored_c: 0, opening_c: 0, shipped_c: 0, year_c: 0,
      containers: { drum200: 0, ibc: 0, bigbag: 0 }, last_at: null
    };
    fifo[t.id] = [];
  });
  sortBy_(table_("movements").rows.slice(), byMovedAsc_).forEach(function (m) {
    var b = out[m.type_id];
    if (!b) return;
    var c = cents_(m.weight_kg);
    var sign = m.kind === "ship" ? -1 : 1;
    if (m.kind === "store") b.stored_c += c;
    else if (m.kind === "opening") b.opening_c += c;
    else if (m.kind === "ship") {
      b.shipped_c += c;
      if (m.moved_at.slice(0, 4) === year) b.year_c += c;
    }
    if (b.containers[m.container] !== undefined) b.containers[m.container] += sign * (m.containers || 0);
    b.last_at = m.moved_at;
    // FIFO lots: what is shipped leaves the oldest lot first.
    var q = fifo[m.type_id];
    if (sign > 0) q.push({ at: m.moved_at, c: c });
    else {
      var left = c;
      while (left > 0 && q.length) {
        var take = Math.min(left, q[0].c);
        q[0].c -= take; left -= take;
        if (q[0].c <= 0) q.shift();
      }
    }
  });
  var nowMs = parseDt_(now.datetime);
  table_("waste_types").rows.forEach(function (t) {
    var b = out[t.id];
    var permitC = Math.round(Number(t.permit_tons || 0) * 100000);
    var bal = b.stored_c + b.opening_c - b.shipped_c;
    var oldest = bal > 0 && fifo[t.id].length ? fifo[t.id][0].at : null;
    out[t.id] = {
      type_id: t.id,
      stored_kg: kgOf_(b.stored_c),
      opening_kg: kgOf_(b.opening_c),
      in_kg: kgOf_(b.stored_c + b.opening_c),
      shipped_kg: kgOf_(b.shipped_c),
      balance_kg: kgOf_(bal),
      containers: b.containers,
      shipped_year_kg: kgOf_(b.year_c),
      permit_kg: kgOf_(permitC),
      quota_pct: permitC > 0 ? Math.round(b.year_c / permitC * 1000) / 10 : null,
      quota_level: quotaLevel_(b.year_c, permitC),
      oldest_at: oldest,
      age_days: oldest ? Math.max(0, Math.floor((nowMs - parseDt_(oldest)) / 86400000)) : null,
      last_at: b.last_at
    };
  });
  return out;
}

function quotaLevel_(usedC, permitC) {
  if (!(permitC > 0)) return usedC > 0 ? "over" : "ok";
  if (usedC > permitC) return "over";
  if (usedC * 100 >= permitC * QUOTA_WARN_PCT) return "warn";
  return "ok";
}

/* Chronological replay of every movement (time order, then id): the weight of
 * each waste type and the container count of each (type, container) pair must
 * never go below zero. Returns null when fine, else the first problem. */
function replay_(moves) {
  var w = {};
  var n = {};
  var list = sortBy_(moves.slice(), byMovedAsc_);
  for (var i = 0; i < list.length; i++) {
    var m = list[i];
    var c = cents_(m.weight_kg);
    var k = m.type_id + "|" + m.container;
    var sign = m.kind === "ship" ? -1 : 1;
    var beforeW = w[m.type_id] || 0;
    var beforeN = n[k] || 0;
    w[m.type_id] = beforeW + sign * c;
    n[k] = beforeN + sign * (m.containers || 0);
    if (w[m.type_id] < 0) return { move: m, what: "weight", have: beforeW, need: c };
    if (n[k] < 0) return { move: m, what: "containers", have: beforeN, need: m.containers };
  }
  return null;
}

// verb: what is being tried (e.g. "ส่งกำจัดไม่ได้"); newMove: the row being
// added / changed (its problem gets a direct message).
function validateLedger_(moves, verb, newMove) {
  var p = replay_(moves);
  if (!p) return;
  var t = typeLabel_(p.move.type_id);
  var when = dtTh_(p.move.moved_at);
  var detail = p.what === "weight"
    ? "น้ำหนักคงเหลือของ " + t + " ณ " + when + " มี " + fmtKg_(kgOf_(p.have)) + " กก. แต่ต้องนำออก " + fmtKg_(kgOf_(p.need)) + " กก."
    : "จำนวน" + (CONTAINERS[p.move.container] || p.move.container) + "ของ " + t + " ณ " + when + " มี " + p.have +
      " ภาชนะ แต่ต้องนำออก " + p.need + " ภาชนะ";
  if (newMove && p.move === newMove) fail_(verb + " — " + detail, "NEGATIVE_BALANCE");
  fail_(verb + " — ยอดคงเหลือจะติดลบที่รายการ " + (p.move.doc_no || "") + " (" + KIND_NAME[p.move.kind] + "): " + detail, "NEGATIVE_BALANCE");
}

function typeLabel_(id) {
  var t = typeMap_()[id];
  return t ? t.code + " " + t.name : "#" + id;
}

/* ===================== read actions ===================== */

function actionBootstrap_() {
  var now = nowParts_();
  var core = cachedDerived_("boot_" + now.date, ["waste_types", "movements"], function () {
    return { types: sortedTypes_().map(plain_), balance: balances_() };
  });
  return {
    types: core.types,
    balance: core.balance,
    containers: CONTAINERS,
    company: COMPANY,
    limits: { quotaWarnPct: QUOTA_WARN_PCT, storageAlertDays: STORAGE_ALERT_DAYS },
    now: { date: now.date, time: now.time }
  };
}

function actionDashboard_() {
  var now = nowParts_();
  var month = now.date.slice(0, 7);
  var year = now.date.slice(0, 4);
  var core = cachedDerived_("dash_" + now.date, ["waste_types", "movements", "audit_logs"], function () {
    var types = sortedTypes_();
    var bal = balances_();
    var moves = table_("movements").rows;

    // last 12 months (oldest first): kg stored / shipped
    var months = [];
    var monthIdx = {};
    var y = Number(year), mo = Number(now.date.slice(5, 7));
    for (var i = 11; i >= 0; i--) {
      var yy = y, mm = mo - i;
      while (mm < 1) { mm += 12; yy -= 1; }
      var key = yy + "-" + pad2_(mm);
      monthIdx[key] = months.length;
      months.push({ ym: key, in_c: 0, out_c: 0 });
    }
    var monthIn = 0, monthOut = 0, yearIn = 0;
    moves.forEach(function (m) {
      var c = cents_(m.weight_kg);
      var ym = m.moved_at.slice(0, 7);
      if (monthIdx[ym] !== undefined) {
        if (m.kind === "store") months[monthIdx[ym]].in_c += c;
        if (m.kind === "ship") months[monthIdx[ym]].out_c += c;
      }
      if (m.kind === "store" && ym === month) monthIn += c;
      if (m.kind === "ship" && ym === month) monthOut += c;
      if (m.kind === "store" && ym.slice(0, 4) === year) yearIn += c;
    });

    var totalBal = 0, totalYear = 0, totalPermit = 0;
    var containers = { drum200: 0, ibc: 0, bigbag: 0 };
    var receivers = {};
    var recvOrder = [];
    var alerts = [];
    types.forEach(function (t) {
      var b = bal[t.id];
      totalBal += cents_(b.balance_kg);
      totalYear += cents_(b.shipped_year_kg);
      totalPermit += cents_(b.permit_kg);
      CONTAINER_KEYS.forEach(function (k) { containers[k] += Math.max(0, b.containers[k]); });
      if (!receivers[t.receiver]) {
        receivers[t.receiver] = { receiver: t.receiver, permit_c: 0, year_c: 0, balance_c: 0, types: 0 };
        recvOrder.push(t.receiver);
      }
      var r = receivers[t.receiver];
      r.permit_c += cents_(b.permit_kg); r.year_c += cents_(b.shipped_year_kg); r.balance_c += cents_(b.balance_kg); r.types += 1;
      if (b.quota_level !== "ok") {
        alerts.push({ level: b.quota_level === "over" ? "bad" : "warn", kind: "quota", type_id: t.id, pct: b.quota_pct,
          used_kg: b.shipped_year_kg, permit_kg: b.permit_kg });
      }
      if (b.age_days !== null && b.age_days > STORAGE_ALERT_DAYS) {
        alerts.push({ level: "warn", kind: "age", type_id: t.id, days: b.age_days, since: b.oldest_at, balance_kg: b.balance_kg });
      }
    });
    alerts.sort(function (a, b) { return (a.level === "bad" ? 0 : 1) - (b.level === "bad" ? 0 : 1); });

    var recent = sortBy_(moves.slice(), byMovedDesc_).slice(0, 10).map(plain_);
    return {
      kpi: {
        balance_kg: kgOf_(totalBal),
        month_in_kg: kgOf_(monthIn),
        month_out_kg: kgOf_(monthOut),
        year_in_kg: kgOf_(yearIn),
        year_out_kg: kgOf_(totalYear),
        permit_kg: kgOf_(totalPermit),
        quota_pct: totalPermit > 0 ? Math.round(totalYear / totalPermit * 1000) / 10 : null,
        moves: moves.length
      },
      months: months.map(function (m) { return { ym: m.ym, in_kg: kgOf_(m.in_c), out_kg: kgOf_(m.out_c) }; }),
      containers: containers,
      receivers: recvOrder.map(function (k) {
        var r = receivers[k];
        return {
          receiver: r.receiver, types: r.types, permit_kg: kgOf_(r.permit_c), shipped_year_kg: kgOf_(r.year_c),
          balance_kg: kgOf_(r.balance_c), quota_pct: r.permit_c > 0 ? Math.round(r.year_c / r.permit_c * 1000) / 10 : null,
          quota_level: quotaLevel_(r.year_c, r.permit_c)
        };
      }),
      alerts: alerts,
      recent: recent,
      logsToday: table_("audit_logs").rows.filter(function (l) { return l.created_at.slice(0, 10) === now.date; }).length
    };
  });
  return core;
}

function queryHistory_(p) {
  var typeMap = typeMap_();
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var type = phpEmpty_(p.type) ? null : phpInt_(p.type);
  var kind = phpEmpty_(p.kind) ? null : str_(p.kind);
  var receiver = phpEmpty_(p.receiver) ? null : str_(p.receiver);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  var rows = table_("movements").rows.filter(function (m) {
    var t = typeMap[m.type_id];
    if (!t) return false;
    if (from !== null && !(m.moved_at >= from)) return false;
    if (to !== null && !(m.moved_at <= to)) return false;
    if (type !== null && m.type_id !== type) return false;
    if (kind !== null && m.kind !== kind) return false;
    if (receiver !== null && t.receiver !== receiver) return false;
    if (q && !(q(m.doc_no) || q(m.note) || q(m.actor) || q(m.location) || q(m.manifest_no) || q(m.vehicle) ||
      q(t.code) || q(t.name))) return false;
    return true;
  });
  return sortBy_(rows, byMovedDesc_).slice(0, 2000).map(function (m) {
    var t = typeMap[m.type_id];
    var o = plain_(m);
    o.code = t.code; o.type_name = t.name; o.receiver = t.receiver; o.mgmt_code = t.mgmt_code;
    o.icon = t.icon; o.color = t.color;
    return o;
  });
}

function exportHistory_(items) {
  var lines = [csvLine_(["เลขที่เอกสาร", "วันที่", "เวลา", "รายการ", "รหัสของเสีย", "ชื่อของเสีย", "รหัสการจัดการ",
    "ผู้รับดำเนินการ", "น้ำหนัก (กก.)", "ภาชนะ", "จำนวนภาชนะ", "สถานที่จัดเก็บ", "เลขที่ใบกำกับการขนส่ง",
    "ทะเบียนรถ", "หมายเหตุ", "ผู้บันทึก"])];
  items.forEach(function (r) {
    // Codes keep their leading zeros in Excel (="080409"; validated digits only).
    lines.push(csvLine_([r.doc_no, r.moved_at.slice(0, 10), r.moved_at.slice(11, 16), KIND_NAME[r.kind] || r.kind,
      excelText_(r.code), csvText_(r.type_name), excelText_(r.mgmt_code), csvText_(r.receiver), Number(r.weight_kg).toFixed(2),
      CONTAINERS[r.container] || r.container, r.containers, csvText_(r.location), csvText_(r.manifest_no),
      csvText_(r.vehicle), csvText_(r.note), csvText_(r.actor)]));
  });
  return { filename: "waste_history_" + stamp_() + ".csv", csv: lines.join("") };
}

// User text in a CSV cell: never starts like a spreadsheet formula.
function csvText_(v) {
  var s = v == null ? "" : String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function excelText_(digits) {
  return /^\d+$/.test(String(digits)) ? '="' + digits + '"' : csvText_(digits);
}

function queryLogs_(p) {
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var cat = phpEmpty_(p.cat) ? null : str_(p.cat);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  return table_("audit_logs").rows.filter(function (l) {
    if (from !== null && !(l.created_at >= from)) return false;
    if (to !== null && !(l.created_at <= to)) return false;
    if (cat !== null && !eqCi_(l.category, cat)) return false;
    if (q && !(q(l.detail) || q(l.ref) || q(l.actor))) return false;
    return true;
  }).sort(function (a, b) { return b.id - a.id; }).slice(0, 3000).map(plain_);
}

function exportLogs_(items) {
  var lines = [csvLine_(["ลำดับ", "วันที่", "เวลา", "หมวด", "เลขที่อ้างอิง", "รายละเอียด", "ผู้ทำรายการ"])];
  items.forEach(function (r) {
    lines.push(csvLine_([r.id, r.created_at.slice(0, 10), r.created_at.slice(11, 19), LOG_CAT_NAME[r.category] || r.category,
      r.ref, csvText_(r.detail), csvText_(r.actor)]));
  });
  return { filename: "waste_log_" + stamp_() + ".csv", csv: lines.join("") };
}

/* ===================== write actions ===================== */

// store / ship / opening. All three take: type_id, weight_kg, container,
// containers, date, time, location, note; ship also manifest_no, vehicle.
function actionMove_(kind, input) {
  var f = moveFields_(kind, input, null);
  var t = getType_(f.type_id);
  if (kind !== "ship" && t.active !== 1) fail_("ประเภทของเสียนี้ถูกปิดใช้งานแล้ว");
  var doc = nextNo_(PREFIX[kind], "movements", "doc_no");
  var row = insert_("movements", {
    doc_no: doc, kind: kind, type_id: f.type_id, weight_kg: f.weight_kg, container: f.container,
    containers: f.containers, location: f.location, manifest_no: f.manifest_no, vehicle: f.vehicle, note: f.note,
    moved_at: f.moved_at, created_at: nowParts_().datetime, actor: REQ_.actor, updated_at: null, updated_by: ""
  });
  validateLedger_(table_("movements").rows, (kind === "ship" ? "ส่งกำจัดเกินยอดคงเหลือ" : "บันทึกไม่ได้"), row);
  audit_("waste", kind, describeMove_(row), doc);
  var b = balances_()[f.type_id];
  return {
    id: row.id, doc_no: doc, kind: kind, balance_kg: b.balance_kg,
    message: KIND_NAME[kind] + " " + t.name + " " + fmtKg_(f.weight_kg) + " กก. เรียบร้อย (" + doc + ")"
  };
}

// Validates the fields of a movement. base: the existing row (edit) or null.
function moveFields_(kind, input, base) {
  var pick = function (k) { return base && input[k] === undefined ? base[k] : input[k]; };
  var typeId = phpInt_(pick("type_id"));
  if (!typeMap_()[typeId]) fail_("กรุณาเลือกประเภทของเสีย");
  var weight = numIn_(pick("weight_kg"), "น้ำหนัก (กก.)", 2);
  if (!(weight > 0)) fail_("น้ำหนักต้องมากกว่า 0 กก.");
  if (weight > MAX_WEIGHT_KG) fail_("น้ำหนักต่อรายการต้องไม่เกิน " + fmtKg_(MAX_WEIGHT_KG) + " กก.");
  var container = phpTrim_(safeStr_(pick("container")));
  if (CONTAINER_KEYS.indexOf(container) === -1) fail_("กรุณาเลือกภาชนะบรรจุ (ถัง 200 ลิตร / IBC / Bigbag)");
  var cRaw = phpTrim_(safeStr_(pick("containers")));
  if (!/^\d{1,6}$/.test(cRaw) || Number(cRaw) < 1) fail_("จำนวนภาชนะต้องเป็นจำนวนเต็มตั้งแต่ 1 ขึ้นไป");
  var at;
  if (base && input.date === undefined && input.time === undefined) at = base.moved_at;
  else at = moment_({ date: input.date == null && base ? base.moved_at.slice(0, 10) : input.date,
    time: input.time == null && base ? base.moved_at.slice(11, 16) : input.time });
  if (parseDt_(at) > Date.now() + TZ_OFFSET_MS + FUTURE_SLACK_MS) fail_("วันที่/เวลาต้องไม่เกินเวลาปัจจุบัน");
  var text = function (k, max, label) {
    var v = pick(k);
    return textIn_({ v: v == null ? "" : v }, "v", max, label);
  };
  var ship = kind === "ship";
  return {
    type_id: typeId, weight_kg: weight, container: container, containers: Number(cRaw), moved_at: at,
    location: text("location", 100, "สถานที่จัดเก็บ"),
    manifest_no: ship ? text("manifest_no", 60, "เลขที่ใบกำกับการขนส่ง") : "",
    vehicle: ship ? text("vehicle", 40, "ทะเบียนรถ") : "",
    note: text("note", MAX_TEXT, "หมายเหตุ")
  };
}

// Decimal input: number or text like "1,234.50"; at most `decimals` places.
function numIn_(v, label, decimals) {
  if (v === null || v === undefined || v === "" || typeof v === "boolean" || typeof v === "object") fail_("กรุณากรอก" + label);
  var s = typeof v === "number" ? String(v) : phpTrim_(String(v)).replace(/,/g, "");
  if (typeof v === "number" && !isFinite(v)) fail_(label + " ไม่ถูกต้อง");
  if (!/^\d+(\.\d+)?$/.test(s) && !(typeof v === "number" && /^\d+(\.\d+)?(e-?\d+)?$/.test(s))) fail_(label + " ต้องเป็นตัวเลข");
  var n = Number(s);
  if (!isFinite(n)) fail_(label + " ไม่ถูกต้อง");
  if (Math.abs(Math.round(n * Math.pow(10, decimals)) - n * Math.pow(10, decimals)) > 1e-6) {
    fail_(label + " ใส่ทศนิยมได้ไม่เกิน " + decimals + " ตำแหน่ง");
  }
  return Math.round(n * Math.pow(10, decimals)) / Math.pow(10, decimals);
}

function describeMove_(m) {
  var t = typeMap_()[m.type_id] || { code: "?", name: "?", receiver: "" };
  return KIND_NAME[m.kind] + " " + t.code + " " + t.name + " น้ำหนัก " + fmtKg_(m.weight_kg) + " กก. · " +
    (CONTAINERS[m.container] || m.container) + " " + m.containers + " ภาชนะ · เวลา " + dtTh_(m.moved_at) +
    (m.kind === "ship" ? " · ผู้รับ: " + t.receiver : "") +
    (m.location ? " · ที่เก็บ: " + m.location : "") +
    (m.manifest_no ? " · ใบกำกับ: " + m.manifest_no : "") +
    (m.vehicle ? " · ทะเบียนรถ: " + m.vehicle : "") +
    (m.note ? " · " + m.note : "");
}

// Wipes all data rows of movements + audit_logs (waste types are kept).
function actionResetData_() {
  var removed = {};
  RESET_TABLES.forEach(function (name) {
    var sheet = db_().getSheetByName(name);
    if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
    var n = Math.max(0, sheet.getLastRow() - 1);
    if (n > 0) {
      var width = Math.max(sheet.getLastColumn(), SCHEMA[name].length);
      sheet.getRange(2, 1, n, width).clearContent();
    }
    removed[name] = n;
    delete REQ_.tables[name];
    delete REQ_.appends[name];
    REQ_.dirty[name] = true;
  });
  REQ_.updates = REQ_.updates.filter(function (u) { return RESET_TABLES.indexOf(u.name) === -1; });
  REQ_.deletes = REQ_.deletes.filter(function (d) { return RESET_TABLES.indexOf(d.name) === -1; });
  audit_("setting", "reset_data", "รีเซ็ตข้อมูล: ล้างชีต movements (" + removed.movements + " แถว), audit_logs (" +
    removed.audit_logs + " แถว) · คงไว้: waste_types");
  return {
    removed: removed,
    message: "รีเซ็ตข้อมูลแล้ว — ลบรายการเคลื่อนไหว " + removed.movements + " แถว, บันทึกประวัติ " + removed.audit_logs + " แถว"
  };
}

/* ---------- edit / delete (WASTE_RESET_PASSWORD) ---------- */

function findMovement_(input) {
  var id = phpInt_(input.id);
  var m = null;
  table_("movements").rows.forEach(function (r) { if (r.id === id) m = r; });
  if (!m) fail_("ไม่พบรายการ (อาจถูกลบไปแล้ว)");
  if (!phpEmpty_(input.doc_no) && str_(input.doc_no) !== m.doc_no) fail_("ข้อมูลรายการไม่ตรงกัน กรุณาโหลดหน้าใหม่");
  return m;
}

function actionMovementUpdate_(input) {
  var m = findMovement_(input);
  if (input.kind !== undefined && str_(input.kind) !== m.kind) fail_("เปลี่ยนชนิดรายการไม่ได้ — ให้ลบแล้วบันทึกใหม่");
  var f = moveFields_(m.kind, input, m);
  var fields = ["type_id", "weight_kg", "container", "containers", "moved_at", "location", "manifest_no", "vehicle", "note"];
  var changes = [];
  fields.forEach(function (k) { if (f[k] !== m[k]) changes.push(k); });
  if (!changes.length) return { changed: false, message: "ไม่มีการเปลี่ยนแปลง" };
  if (f.type_id !== m.type_id && m.kind !== "ship" && getType_(f.type_id).active !== 1) fail_("ประเภทของเสียนี้ถูกปิดใช้งานแล้ว");
  var copy = plain_(m);
  fields.forEach(function (k) { copy[k] = f[k]; });
  var moves = table_("movements").rows.map(function (r) { return r === m ? copy : r; });
  validateLedger_(moves, "แก้ไขไม่ได้", copy);
  var before = describeMove_(m);
  var set = { updated_at: nowParts_().datetime, updated_by: REQ_.actor };
  fields.forEach(function (k) { set[k] = f[k]; });
  update_("movements", m, set);
  audit_("edit", "movement_update", "แก้ไขรายการ " + m.doc_no + " — เดิม: " + before + " → ใหม่: " + describeMove_(m), m.doc_no);
  return { changed: true, id: m.id, doc_no: m.doc_no, message: "แก้ไขรายการ " + m.doc_no + " เรียบร้อย" };
}

function actionMovementDelete_(input) {
  var m = findMovement_(input);
  var moves = table_("movements").rows.filter(function (r) { return r !== m; });
  validateLedger_(moves, "ลบไม่ได้", null);
  var detail = "ลบรายการ " + m.doc_no + " — " + describeMove_(m);
  delete_("movements", m);
  audit_("edit", "movement_delete", detail, m.doc_no);
  return { id: m.id, doc_no: m.doc_no, message: "ลบรายการ " + m.doc_no + " แล้ว" };
}

/* ---------- waste types (WASTE_RESET_PASSWORD) ---------- */

function actionTypeSave_(input) {
  var id = phpInt_(input.id);
  var existing = null;
  if (id) {
    existing = typeMap_()[id];
    if (!existing) fail_("ไม่พบประเภทของเสียนี้");
  }
  var code = phpTrim_(safeStr_(input.code));
  if (!/^\d{6}$/.test(code)) fail_("รหัสของเสียต้องเป็นตัวเลข 6 หลัก");
  var name = textIn_(input, "name", 120, "ชื่อของเสีย");
  if (!name) fail_("กรุณากรอกชื่อของเสีย");
  var tons = numIn_(input.permit_tons, "ปริมาณที่ได้รับอนุญาต (ตัน)", 3);
  if (tons > MAX_PERMIT_TONS) fail_("ปริมาณที่ได้รับอนุญาตสูงเกินไป");
  var mgmt = phpTrim_(safeStr_(input.mgmt_code));
  if (!/^\d{3}$/.test(mgmt)) fail_("รหัสวิธีการจัดการต้องเป็นตัวเลข 3 หลัก");
  var reg = phpTrim_(safeStr_(input.receiver_reg));
  if (!/^\d{0,20}$/.test(reg)) fail_("เลขทะเบียนผู้รับต้องเป็นตัวเลข (ไม่เกิน 20 หลัก)");
  var receiver = textIn_(input, "receiver", 120, "ผู้รับดำเนินการ");
  if (!receiver) fail_("กรุณากรอกชื่อผู้รับดำเนินการ");
  var icon = phpTrim_(safeStr_(input.icon)) || (existing ? existing.icon : "fa-recycle");
  if (!/^fa-[a-z0-9-]{1,40}$/.test(icon)) fail_("ไอคอนไม่ถูกต้อง");
  var color = phpTrim_(safeStr_(input.color)) || (existing ? existing.color : "#10B981");
  if (!/^#[0-9a-fA-F]{6}$/.test(color)) fail_("สีไม่ถูกต้อง");
  var active = input.active === undefined ? (existing ? existing.active : 1) : (phpTruthy_(input.active) && str_(input.active) !== "false" ? 1 : 0);
  var dup = table_("waste_types").rows.some(function (t) {
    return t !== existing && t.code === code && t.mgmt_code === mgmt && t.receiver_reg === reg && t.name === name;
  });
  if (dup) fail_("มีประเภทของเสียนี้ (รหัส ชื่อ วิธีจัดการ และผู้รับเดียวกัน) อยู่แล้ว");
  var fields = {
    code: code, name: name, permit_tons: tons, mgmt_code: mgmt, receiver_reg: reg, receiver: receiver,
    icon: icon, color: color.toUpperCase(), active: active, updated_at: nowParts_().datetime, updated_by: REQ_.actor
  };
  var desc = code + " " + name + " · อนุญาต " + tons + " ตัน · รหัสจัดการ " + mgmt + " · " + receiver +
    (reg ? " (" + reg + ")" : "") + " · " + (active ? "ใช้งาน" : "ปิดใช้งาน");
  if (existing) {
    var before = existing.code + " " + existing.name + " · อนุญาต " + existing.permit_tons + " ตัน · รหัสจัดการ " +
      existing.mgmt_code + " · " + existing.receiver + (existing.receiver_reg ? " (" + existing.receiver_reg + ")" : "") +
      " · " + (existing.active ? "ใช้งาน" : "ปิดใช้งาน");
    update_("waste_types", existing, fields);
    audit_("setting", "type_save", "แก้ไขประเภทของเสีย #" + existing.id + " — เดิม: " + before + " → ใหม่: " + desc, "T" + existing.id);
    return { id: existing.id, created: false, message: "บันทึกประเภทของเสีย " + name + " แล้ว" };
  }
  var maxSort = 0;
  table_("waste_types").rows.forEach(function (t) { if (t.sort > maxSort) maxSort = t.sort; });
  fields.sort = maxSort + 1;
  var row = insert_("waste_types", fields);
  audit_("setting", "type_save", "เพิ่มประเภทของเสีย #" + row.id + ": " + desc, "T" + row.id);
  return { id: row.id, created: true, message: "เพิ่มประเภทของเสีย " + name + " แล้ว" };
}

/* ===================== business helpers ===================== */

function getType_(id) {
  var t = typeMap_()[id];
  if (!t) fail_("ไม่พบประเภทของเสีย");
  return t;
}

function typeMap_() {
  var map = {};
  table_("waste_types").rows.forEach(function (t) { map[t.id] = t; });
  return map;
}

// date/time from the client (default: now), "YYYY-MM-DD" + "HH:MM".
function moment_(input) {
  var now = nowParts_();
  var d = input.date == null || input.date === "" ? now.date : safeStr_(input.date);
  var t = input.time == null || input.time === "" ? now.time : safeStr_(input.time);
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})$/.exec(d + " " + t.slice(0, 5));
  if (!m) fail_("วันที่/เวลาไม่ถูกต้อง");
  var Y = Number(m[1]), M = Number(m[2]), D = Number(m[3]), h = Number(m[4]), mi = Number(m[5]);
  var dt = new Date(Date.UTC(Y, M - 1, D, h, mi, 0));
  if (Y < 2000 || dt.getUTCMonth() !== M - 1 || dt.getUTCDate() !== D || h > 23 || mi > 59) fail_("วันที่/เวลาไม่ถูกต้อง");
  return formatUtcParts_(dt).datetime.slice(0, 16) + ":00";
}

// Next running number for today's tag (audit refs scanned too, so a number
// freed by a delete is never handed out again).
function nextNo_(prefix, table, col) {
  var tag = prefix + "-" + nowParts_().ymd + "-";
  var max = 0;
  var scan = function (v) {
    var s = String(v || "").toUpperCase();
    if (s.indexOf(tag) !== 0) return;
    var n = parseInt(s.slice(tag.length), 10);
    if (n > max) max = n;
  };
  table_(table).rows.forEach(function (r) { scan(r[col]); });
  table_("audit_logs").rows.forEach(function (l) { scan(l.ref); });
  var next = max + 1;
  return tag + ("0000" + next).slice(-Math.max(4, String(next).length));
}

function audit_(category, action, detail, ref) {
  insert_("audit_logs", {
    category: category, action: action, ref: ref || "",
    detail: mbSubstr_(detail, 0, 500),
    actor: REQ_.actor || "",
    created_at: nowParts_().datetime
  });
}

function textIn_(input, key, max, label) {
  var value = phpTrim_(safeStr_(input[key]));
  if (mbLen_(value) > max) fail_(label + "ยาวเกิน " + max + " ตัวอักษร");
  return value;
}

function fmtKg_(n) {
  var s = (Math.round(Number(n) * 100) / 100).toFixed(2);
  var parts = s.split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  var dec = parts[1].replace(/0+$/, "");
  return dec ? parts[0] + "." + dec : parts[0];
}

function dtTh_(s) {
  return s.slice(8, 10) + "/" + s.slice(5, 7) + "/" + (Number(s.slice(0, 4)) + 543) + " " + s.slice(11, 16) + " น.";
}

/* ===================== sheet storage ===================== */

var REQ_ = null; // per-request state: spreadsheet, loaded tables, pending writes

// actor: the name typed in the page header ("" when none).
function resetRequest_(actor) {
  REQ_ = {
    ss: null, tables: {}, appends: {}, updates: [], deletes: [], actor: actor || "",
    dirty: {},          // sheets changed outside the buffers (reset_data) -> version bump on flush
    cacheReads: false,  // true only while answering read actions
    refreshCache: false, // keepWarm: read the sheets live and store them again (fresh TTL)
    vers: null          // data versions read for this request (see versions_)
  };
}

function db_() {
  if (!REQ_.ss) {
    var id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
    if (!id) throw new Error("ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ใน Apps Script ก่อน");
    REQ_.ss = SpreadsheetApp.openById(id);
  }
  return REQ_.ss;
}

// Reads the whole sheet once per request and converts rows to typed objects.
// Sheets in AUTO_CREATE_TABLES are created when missing, and columns added to
// SCHEMA after a sheet was created are appended to its header row (existing
// rows read them as empty), so an older live spreadsheet upgrades itself.
function table_(name) {
  if (REQ_.tables[name]) return REQ_.tables[name];
  var cacheable = REQ_.cacheReads && CACHED_TABLES.indexOf(name) !== -1;
  var ver = null;
  if (cacheable) {
    ver = versions_()[name] || null; // taken BEFORE the sheet is read
    var hit = ver && !REQ_.refreshCache ? cachedRows_(name, ver) : null;
    if (hit) return (REQ_.tables[name] = hit);
  }
  var sheet = db_().getSheetByName(name);
  if (!sheet && AUTO_CREATE_TABLES.indexOf(name) !== -1) sheet = prepareSheet_(db_(), name, null);
  if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
  var values = sheet.getDataRange().getValues();
  var header = (values[0] || []).map(function (h) { return String(h).trim(); });
  var cols = SCHEMA[name];
  var idx = cols.map(function (c) { return header.indexOf(c[0]); });
  if (idx.indexOf(-1) !== -1) {
    if (!header.some(function (h) { return h !== ""; })) {
      throw new Error('หัวตารางของชีต "' + name + '" ไม่ครบ กรุณารัน setupSystem()');
    }
    header = addMissingColumns_(sheet, name, header);
    idx = cols.map(function (c) { return header.indexOf(c[0]); });
  }
  var rows = [];
  var maxId = 0;
  for (var r = 1; r < values.length; r++) {
    var raw = values[r];
    var empty = true;
    for (var k = 0; k < raw.length; k++) if (raw[k] !== "" && raw[k] !== null) { empty = false; break; }
    if (empty) continue;
    var obj = {};
    cols.forEach(function (col, i) { obj[col[0]] = fromCell_(col[1], raw[idx[i]]); });
    Object.defineProperty(obj, "_row", { value: r + 1, writable: true, enumerable: false });
    if (obj.id > maxId) maxId = obj.id;
    rows.push(obj);
  }
  // values: the raw sheet rows, so flush_ can update a row without reading it again.
  REQ_.tables[name] = { name: name, sheet: sheet, header: header, idx: idx, rows: rows, maxId: maxId, values: values };
  if (ver) storeRows_(name, ver, rows);
  return REQ_.tables[name];
}

/* ---------- read cache helpers (see CACHED_TABLES) ---------- */

function newVersion_() {
  return Date.now().toString(36) + "." + Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

// The current data version of every cached sheet (one CacheService call per
// request). A missing version (evicted / first use) gets a new random one, so
// nothing stored under an older version can match it. Cache errors disable the
// read cache for the request ({}).
function versions_() {
  if (REQ_.vers) return REQ_.vers;
  var vers = {};
  try {
    var keys = CACHED_TABLES.map(function (n) { return VERSION_PREFIX + n; });
    var got = cache_().getAll(keys) || {};
    var missing = {};
    CACHED_TABLES.forEach(function (n) {
      var v = got[VERSION_PREFIX + n];
      if (!v) { v = newVersion_(); missing[VERSION_PREFIX + n] = v; }
      vers[n] = v;
    });
    if (Object.keys(missing).length) cache_().putAll(missing, VERSION_TTL);
  } catch (error) {
    vers = {};
  }
  REQ_.vers = vers;
  return vers;
}

// Gives every named (cached) sheet a new data version. Called after the
// changes are flushed to the spreadsheet. If the cache cannot be updated the
// versions are removed instead (readers then start new ones).
function bumpVersions_(names) {
  var map = {};
  names.forEach(function (n) { if (CACHED_TABLES.indexOf(n) !== -1) map[VERSION_PREFIX + n] = newVersion_(); });
  var keys = Object.keys(map);
  if (!keys.length) return;
  try {
    cache_().putAll(map, VERSION_TTL);
  } catch (error) {
    try { cache_().removeAll(keys); } catch (ignored) {}
  }
  if (REQ_ && REQ_.vers) REQ_.vers = null; // re-read on the next cached read of this request
}

// Stores text under key in chunks of RC_CHUNK characters: key holds "<n>|<chunk 0>",
// key_1..key_<n-1> the rest. Too big (> RC_MAX_CHUNKS) or cache errors: not stored.
function rcPut_(key, text) {
  var ttl = REQ_ && REQ_.refreshCache ? KEEP_WARM_TTL : READ_CACHE_TTL;
  var n = Math.max(1, Math.ceil(text.length / RC_CHUNK));
  if (n > RC_MAX_CHUNKS) return;
  var map = {};
  for (var i = 0; i < n; i++) {
    var part = text.substr(i * RC_CHUNK, RC_CHUNK);
    map[i ? key + "_" + i : key] = i ? part : n + "|" + part;
  }
  try { cache_().putAll(map, ttl); } catch (ignored) {}
}

function rcGet_(key) {
  try {
    var head = cache_().get(key);
    if (!head) return null;
    var bar = head.indexOf("|");
    var n = parseInt(head.slice(0, bar), 10);
    if (!(n >= 1)) return null;
    var parts = [head.slice(bar + 1)];
    if (n > 1) {
      var keys = [];
      for (var i = 1; i < n; i++) keys.push(key + "_" + i);
      var got = cache_().getAll(keys) || {};
      for (var j = 0; j < keys.length; j++) {
        if (typeof got[keys[j]] !== "string") return null;
        parts.push(got[keys[j]]);
      }
    }
    return parts.join("");
  } catch (error) {
    return null;
  }
}

// Rows are stored as arrays in SCHEMA order, normalised exactly like a sheet
// round trip (toCell_ -> fromCell_), so cached and live reads are identical.
function storeRows_(name, ver, rows) {
  var cols = SCHEMA[name];
  var data = rows.map(function (r) {
    return cols.map(function (c) { return fromCell_(c[1], toCell_(c[1], r[c[0]])); });
  });
  rcPut_(RC_PREFIX + name + "_" + ver, JSON.stringify(data));
}

function cachedRows_(name, ver) {
  var text = rcGet_(RC_PREFIX + name + "_" + ver);
  if (!text) return null;
  var data;
  try { data = JSON.parse(text); } catch (error) { return null; }
  var cols = SCHEMA[name];
  var maxId = 0;
  var rows = data.map(function (a) {
    var obj = {};
    cols.forEach(function (c, i) { obj[c[0]] = a[i]; });
    Object.defineProperty(obj, "_row", { value: 0, writable: true, enumerable: false });
    if (obj.id > maxId) maxId = obj.id;
    return obj;
  });
  // No sheet: tables from the cache are read-only (writes always read live).
  return { name: name, sheet: null, header: null, idx: null, rows: rows, maxId: maxId, values: null, cached: true };
}

// Result of fn() cached under the versions of the sheets it depends on.
function cachedDerived_(id, deps, fn) {
  if (!REQ_.cacheReads) return fn();
  var vers = versions_();
  var parts = [];
  for (var i = 0; i < deps.length; i++) {
    if (!vers[deps[i]]) return fn();
    parts.push(vers[deps[i]]);
  }
  var key = RC_PREFIX + id + "_" + parts.join("_");
  var text = REQ_.refreshCache ? null : rcGet_(key);
  if (text) {
    try { return JSON.parse(text); } catch (ignored) {}
  }
  var value = fn();
  rcPut_(key, JSON.stringify(value));
  return value;
}

/**
 * Run from the Apps Script editor after editing the spreadsheet by hand: gives
 * every cached sheet a new data version, so the next reads load the sheets.
 */
function clearReadCache() {
  bumpVersions_(CACHED_TABLES);
  return { cleared: CACHED_TABLES };
}

/* ---------- keep-warm (time-driven trigger) ----------
 * keepWarm() runs the reads the pages open with (bootstrap, dashboard)
 * the way a web request would, but reads every sheet live and stores
 * it again in the read cache under the CURRENT data versions with a TTL longer
 * than the trigger period, so the next web request is answered from
 * CacheService. It only reads: no sheet row is written, no version is bumped,
 * no lock is taken (a missing version is created exactly as by any read).
 * Hand edits in the spreadsheet are picked up at each run.
 * Run installKeepWarmTrigger() once from the Apps Script editor.
 */
function keepWarm() {
  try {
    resetRequest_("");
    REQ_.cacheReads = true;
    REQ_.refreshCache = true;
    handleRead_("bootstrap", {});
    handleRead_("dashboard", {});
    return { warmed: true };
  } catch (error) {
    console.log("keepWarm: " + errorMessage_(error));
    return { warmed: false, error: errorMessage_(error) };
  } finally {
    REQ_ = null;
  }
}

/** Run once from the editor: (re)creates the one keepWarm trigger (every 10 minutes). */
function installKeepWarmTrigger() {
  removeKeepWarmTrigger();
  ScriptApp.newTrigger("keepWarm").timeBased().everyMinutes(KEEP_WARM_MINUTES).create();
  var warm = keepWarm();
  var result = { installed: true, everyMinutes: KEEP_WARM_MINUTES, warmed: warm.warmed };
  console.log(JSON.stringify(result));
  return result;
}

/** Deletes every keepWarm trigger of this project (other triggers are left alone). */
function removeKeepWarmTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "keepWarm") { ScriptApp.deleteTrigger(t); removed++; }
  });
  return { removed: removed };
}

function insert_(name, obj) {
  var t = table_(name);
  var row = {};
  SCHEMA[name].forEach(function (col) { row[col[0]] = obj[col[0]] === undefined ? null : obj[col[0]]; });
  t.maxId += 1;
  row.id = t.maxId;
  Object.defineProperty(row, "_row", { value: 0, writable: true, enumerable: false });
  t.rows.push(row);
  (REQ_.appends[name] = REQ_.appends[name] || []).push(row);
  return row;
}

function update_(name, row, fields) {
  Object.keys(fields).forEach(function (k) { row[k] = fields[k]; });
  if (row._row > 0) {
    for (var i = 0; i < REQ_.updates.length; i++) if (REQ_.updates[i].row === row) return;
    REQ_.updates.push({ name: name, row: row });
  } // rows inserted in this request are written with their latest values on flush
}

// Removes a row from the in-memory table and buffers the sheet row deletion.
function delete_(name, row) {
  var t = table_(name);
  var i = t.rows.indexOf(row);
  if (i !== -1) t.rows.splice(i, 1);
  REQ_.updates = REQ_.updates.filter(function (u) { return u.row !== row; });
  if (row._row > 0) {
    REQ_.deletes.push({ name: name, row: row });
  } else if (REQ_.appends[name]) {
    REQ_.appends[name] = REQ_.appends[name].filter(function (r) { return r !== row; });
  }
}

// Writes buffered changes in a safe order: updates (row numbers are still the
// ones read at the start of the request), then deletions (highest row first,
// each run of consecutive rows in one deleteRows call), then appends (one
// setNumberFormats + one setValues per table). Tables that lost rows are
// dropped from the per-request cache because their row numbers have shifted.
// Afterwards every changed sheet gets a new data version (read cache), also when
// a write fails half-way.
function flush_() {
  var changed = Object.keys(REQ_.dirty);
  REQ_.updates.forEach(function (u) { changed.push(u.name); });
  REQ_.deletes.forEach(function (d) { changed.push(d.name); });
  Object.keys(REQ_.appends).forEach(function (n) { if (REQ_.appends[n].length) changed.push(n); });
  REQ_.dirty = {};
  try {
    flushBuffers_();
  } finally {
    if (changed.length) {
      // Make the writes visible to other executions before announcing them.
      try { SpreadsheetApp.flush(); } catch (ignored) {}
      bumpVersions_(changed.filter(function (n, i, a) { return a.indexOf(n) === i; }));
    }
  }
}

function flushBuffers_() {
  REQ_.updates.forEach(function (u) {
    var t = REQ_.tables[u.name];
    var range = t.sheet.getRange(u.row._row, 1, 1, t.header.length);
    // The row as read at the start of this request (under the lock), else read it now.
    var current = t.values && t.values[u.row._row - 1] ? t.values[u.row._row - 1] : range.getValues()[0];
    var next = rowValues_(u.name, t, u.row, t.header.length, current);
    range.setValues([next]);
    if (t.values && t.values[u.row._row - 1]) t.values[u.row._row - 1] = next;
  });
  REQ_.updates = [];

  var byTable = {};
  REQ_.deletes.forEach(function (d) { (byTable[d.name] = byTable[d.name] || []).push(d.row._row); });
  Object.keys(byTable).forEach(function (name) {
    var sheet = REQ_.tables[name].sheet;
    var rows = byTable[name]
      .filter(function (r, i, a) { return r > 1 && a.indexOf(r) === i; })
      .sort(function (a, b) { return b - a; });
    // Sheets refuses to delete every non-frozen row: keep one spare row.
    if (sheet.getMaxRows() - rows.length < 2) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    var i = 0;
    while (i < rows.length) {
      var j = i;
      while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
      sheet.deleteRows(rows[j], j - i + 1); // rows[j] = lowest row number of the run
      i = j + 1;
    }
  });
  REQ_.deletes = [];

  Object.keys(REQ_.appends).forEach(function (name) {
    var list = REQ_.appends[name];
    if (!list.length) return;
    var t = REQ_.tables[name];
    var width = t.header.length;
    var start = Math.max(t.sheet.getLastRow(), 1) + 1;
    var needed = start + list.length - 1;
    var maxRows = t.sheet.getMaxRows();
    if (needed > maxRows) t.sheet.insertRowsAfter(maxRows, needed - maxRows);
    var formats = [];
    var values = list.map(function (row) {
      formats.push(rowFormats_(name, t));
      row._row = start + formats.length - 1;
      return rowValues_(name, t, row, width);
    });
    var range = t.sheet.getRange(start, 1, list.length, width);
    range.setNumberFormats(formats);
    range.setValues(values);
  });
  REQ_.appends = {};
  Object.keys(byTable).forEach(function (name) { delete REQ_.tables[name]; });
}

// Appends the SCHEMA columns missing from the header row (written at once, not
// buffered; plain-text / number format for the whole new columns). Returns the
// new header.
function addMissingColumns_(sheet, name, header) {
  while (header.length && header[header.length - 1] === "") header.pop();
  var missing = SCHEMA[name].filter(function (c) { return header.indexOf(c[0]) === -1; });
  if (!missing.length) return header;
  var start = header.length + 1;
  var needCols = start + missing.length - 1;
  var maxCols = sheet.getMaxColumns();
  if (needCols > maxCols) sheet.insertColumnsAfter(maxCols, needCols - maxCols);
  sheet.getRange(1, start, 1, missing.length).setValues([missing.map(function (c) { return c[0]; })]);
  var maxRows = sheet.getMaxRows();
  var rowFormat = missing.map(function (c) { return cellFormat_(c[1]); });
  var formats = [];
  for (var r = 0; r < maxRows; r++) formats.push(rowFormat);
  sheet.getRange(1, start, maxRows, missing.length).setNumberFormats(formats);
  sheet.getRange(1, start, 1, missing.length).setFontWeight("bold");
  bumpVersions_([name]);
  return header.concat(missing.map(function (c) { return c[0]; }));
}

function rowValues_(name, t, row, width, base) {
  var out = base ? base.slice() : [];
  while (out.length < width) out.push("");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = toCell_(col[1], row[col[0]]); });
  return out;
}

function rowFormats_(name, t) {
  var out = [];
  for (var i = 0; i < t.header.length; i++) out.push("@");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = cellFormat_(col[1]); });
  return out;
}

// Sheet number format of a column type: whole numbers, decimals or plain text.
function cellFormat_(type) {
  if (type.indexOf("int") === 0) return "0";
  if (type.indexOf("num") === 0) return "0.00#";
  return "@";
}

function toCell_(type, v) {
  if (v === null || v === undefined) return "";
  if (type.indexOf("int") === 0 || type.indexOf("num") === 0) return Number(v);
  var s = String(v);
  // Never let user text become a formula (cells are also plain-text formatted).
  if (s.charAt(0) === "=") s = "'" + s;
  return s;
}

function fromCell_(type, v) {
  var nullable = type.charAt(type.length - 1) === "?";
  var base = nullable ? type.slice(0, -1) : type;
  if (v === "" || v === null || v === undefined) return nullable ? null : (base === "int" || base === "num" ? 0 : "");
  if (base === "num") {
    var x = Number(v);
    return isFinite(x) ? Math.round(x * 1000) / 1000 : 0;
  }
  if (base === "int") {
    var n = Number(v);
    return isFinite(n) ? Math.trunc(n) : 0;
  }
  if (Object.prototype.toString.call(v) === "[object Date]") {
    // A cell Sheets converted to a Date (spreadsheet time zone is Asia/Bangkok).
    return formatUtcParts_(new Date(v.getTime() + TZ_OFFSET_MS)).datetime;
  }
  var s = String(v);
  if (s.charAt(0) === "'" && s.charAt(1) === "=") s = s.slice(1);
  return s;
}

/* ===================== setup ===================== */

/**
 * One-time setup. Run from the Apps Script editor and approve permissions.
 * If WASTE_SPREADSHEET_ID is already set, the existing spreadsheet is reused
 * (missing sheets/headers are added; waste types are seeded only when the
 * waste_types sheet is empty) instead of creating another.
 */
function setupSystem() {
  var properties = PropertiesService.getScriptProperties();
  var id = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  var ss;
  var created = false;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (error) {
      throw new Error("เปิดสเปรดชีตเดิม (" + id + ") ไม่ได้: " + errorMessage_(error) +
        " — ถ้าต้องการสร้างใหม่ ให้ลบ Script Property " + SPREADSHEET_ID_PROPERTY + " แล้วรันอีกครั้ง");
    }
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    ss.setSpreadsheetTimeZone("Asia/Bangkok");
    created = true;
    properties.setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
  }

  var defaultSheets = created ? ss.getSheets() : [];
  TABLE_ORDER.forEach(function (name, i) {
    var sheet = ss.getSheetByName(name);
    if (!sheet && created && i === 0 && defaultSheets.length) {
      sheet = defaultSheets[0];
      sheet.setName(name);
    }
    prepareSheet_(ss, name, sheet);
  });

  resetRequest_("setupSystem");
  REQ_.ss = ss;
  TABLE_ORDER.forEach(function (name) { table_(name); REQ_.dirty[name] = true; });
  var seeded = 0;
  if (!table_("waste_types").rows.length) {
    var at = nowParts_().datetime;
    SEED_TYPES.forEach(function (t, i) {
      insert_("waste_types", {
        code: t[0], name: t[1], permit_tons: t[2], mgmt_code: t[3], receiver_reg: t[4], receiver: t[5],
        icon: t[6], color: t[7], active: 1, sort: i + 1, updated_at: at, updated_by: "setupSystem"
      });
      seeded++;
    });
  }
  flush_();
  REQ_ = null;

  var result = {
    created: created,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    seededTypes: seeded,
    resetPasswordConfigured: Boolean(properties.getProperty(RESET_PASSWORD_PROPERTY))
  };
  console.log(JSON.stringify(result));
  return result;
}

// Creates the sheet when missing, writes the header row on an empty sheet,
// bolds/freezes it and sets number / plain-text formats per column so Sheets
// never converts values (e.g. "080409" into 80409 or a date text into a Date).
function prepareSheet_(ss, name, sheet) {
  if (!sheet) sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var cols = SCHEMA[name];
  var headers = cols.map(function (c) { return c[0]; });
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  var maxRows = sheet.getMaxRows();
  var formats = [];
  var rowFormat = cols.map(function (c) { return cellFormat_(c[1]); });
  for (var r = 0; r < maxRows; r++) formats.push(rowFormat);
  sheet.getRange(1, 1, maxRows, headers.length).setNumberFormats(formats);
  return sheet;
}

/* ===================== small utilities ===================== */

function fail_(message, code) {
  var error = new Error(message);
  if (code) error.code = code;
  throw error;
}

function errorMessage_(error) {
  return (error && error.message) ? String(error.message) : String(error);
}

function errorResponse_(error) {
  var body = { ok: false, error: errorMessage_(error) };
  if (error && error.code) body.code = error.code;
  if (error && error.passwordError) body.passwordError = true;
  return jsonResponse_(body);
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function plain_(obj) {
  var out = {};
  Object.keys(obj).forEach(function (k) { out[k] = obj[k]; });
  return out;
}

function sortBy_(arr, cmp) { // stable sort
  return arr.map(function (v, i) { return [v, i]; })
    .sort(function (a, b) { return cmp(a[0], b[0]) || a[1] - b[1]; })
    .map(function (p) { return p[0]; });
}

function byMovedAsc_(a, b) {
  if (a.moved_at !== b.moved_at) return a.moved_at < b.moved_at ? -1 : 1;
  return a.id - b.id;
}

function byMovedDesc_(a, b) {
  if (a.moved_at !== b.moved_at) return a.moved_at < b.moved_at ? 1 : -1;
  return b.id - a.id;
}

function str_(v) {
  if (v === null || v === undefined || v === false) return "";
  if (v === true) return "1";
  return String(v);
}

// Like str_ but rejects arrays/objects (PHP trim() on an array throws).
function safeStr_(v) {
  if (v !== null && typeof v === "object") fail_("ข้อมูลที่ส่งมาไม่ถูกต้อง");
  return str_(v);
}

function phpTrim_(s) {
  return String(s).replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

function phpInt_(v) {
  if (v === null || v === undefined || v === false) return 0;
  if (v === true) return 1;
  if (typeof v === "number") return isFinite(v) ? Math.trunc(v) : 0;
  if (typeof v === "object") return Array.isArray(v) && v.length ? 1 : 0;
  var m = /^[ \t\n\r\v\f]*([+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?)/.exec(String(v));
  if (!m) return 0;
  var n = Number(m[1]);
  return isFinite(n) ? Math.trunc(n) : 0;
}

function phpTruthy_(v) {
  if (v === null || v === undefined || v === false || v === 0 || v === "" || v === "0") return false;
  if (Array.isArray(v)) return v.length > 0;
  return true;
}

function phpEmpty_(v) {
  return !phpTruthy_(v);
}

function mbLen_(s) {
  return Array.from(String(s)).length;
}

function mbSubstr_(s, start, len) {
  return Array.from(String(s)).slice(start, start + len).join("");
}

function eqCi_(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return false;
  return String(a).toLowerCase() === String(b).toLowerCase();
}

// MySQL "col LIKE '%q%'" (case-insensitive, % and _ wildcards, \ escape).
function likeMatcher_(q) {
  var re = "";
  var chars = Array.from(q);
  for (var i = 0; i < chars.length; i++) {
    var ch = chars[i];
    if (ch === "\\" && i + 1 < chars.length) { re += escapeRe_(chars[++i]); continue; }
    if (ch === "%") re += "[\\s\\S]*";
    else if (ch === "_") re += "[\\s\\S]";
    else re += escapeRe_(ch);
  }
  var rx = new RegExp(re, "i");
  return function (v) { return v !== null && v !== undefined && rx.test(String(v)); };
}

function escapeRe_(ch) {
  return ch.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
}

// PHP fputcsv() with default delimiter/enclosure/escape, "\n" line ending.
function csvLine_(fields) {
  return fields.map(function (f) {
    var s = f === null || f === undefined ? "" : (f === true ? "1" : f === false ? "" : String(f));
    if (!/[,"\\\n\r\t ]/.test(s)) return s;
    var out = '"';
    var escaped = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === "\\") escaped = true;
      else if (!escaped && ch === '"') out += '"';
      else escaped = false;
      out += ch;
    }
    return out + '"';
  }).join(",") + "\n";
}

function pad2_(n) {
  return (n < 10 ? "0" : "") + n;
}

function formatUtcParts_(d) {
  var y = String(d.getUTCFullYear());
  while (y.length < 4) y = "0" + y;
  var date = y + "-" + pad2_(d.getUTCMonth() + 1) + "-" + pad2_(d.getUTCDate());
  var time = pad2_(d.getUTCHours()) + ":" + pad2_(d.getUTCMinutes());
  return {
    date: date,
    time: time,
    datetime: date + " " + time + ":" + pad2_(d.getUTCSeconds()),
    ymd: y.slice(-2) + pad2_(d.getUTCMonth() + 1) + pad2_(d.getUTCDate())
  };
}

// Current wall-clock time in Asia/Bangkok.
function nowParts_() {
  return formatUtcParts_(new Date(Date.now() + TZ_OFFSET_MS));
}

function parseDt_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(s));
  if (!m) return NaN;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

function shiftDate_(date, days) {
  return formatUtcParts_(new Date(parseDt_(date) + days * 86400000)).date;
}

function shiftDateTime_(datetime, days) {
  return formatUtcParts_(new Date(parseDt_(datetime) + days * 86400000)).datetime;
}

function stamp_() {
  var p = nowParts_();
  return p.date.replace(/-/g, "") + "_" + p.datetime.slice(11).replace(/:/g, "");
}

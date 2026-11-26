/**
 * نظام الحضور بالـQR - سيرفر Google Apps Script (نسخة متعددة التدريبات)
 *
 * سيرفر واحد بيخدم أكتر من تدريب (كل تدريب = شيت منفصل).
 * بيتحط جوه "شيت التحكم" (أي شيت عندك) من Extensions > Apps Script.
 * شيت التحكم فيه تابات بيعملها السكريبت لوحده:
 *   Trainings : قايمة التدريبات (اسم + رابط الشيت + الأيام + الموقع)
 *   Devices   : ربط كل متدرب بموبايله
 *   Log       : سجل كل محاولة تسجيل
 */

// ====== إعدادات ======
var CONFIG = {
  HEADER_ROW: 1,
  PHONE_HEADERS: ['رقم التليفون', 'رقم الموبايل', 'رقم الهاتف', 'رقم المحمول', 'رقم الجوال', 'Phone', 'Mobile', 'Phone Number'],
  NAME_HEADERS: ['الاسم رباعي باللغة العربية', 'الاسم رباعي', 'الاسم بالكامل', 'الاسم', 'Name', 'Full Name'],
  PROTECTED_HEADERS: ['email', 'timestamp', 'الطابع الزمني', 'البريد الإلكتروني'],
  PRESENT_VALUE: 1,
  GRACE_SLOTS: 1,
  MIN_ROTATION: 15,
  MAX_ROTATION: 120,
  OUT_SUFFIX: ' - Out',
  SUMMARY_HEADER: 'نسبة الحضور',
  SUMMARY_THRESHOLD: 0.7,
  SESSION_MAX_AGE_MS: 18 * 3600 * 1000,
  TRAINEES_CACHE_SEC: 300,
  TRAININGS_CACHE_SEC: 120,
  DEFAULT_RADIUS: 150,
  TZ: 'Africa/Cairo'
};

var TR_HEADERS = ['id', 'name', 'sheet', 'tab', 'days', 'lat', 'lng', 'radius', 'geo', 'created'];
var LOG_HEADERS = ['time', 'session', 'day', 'phone', 'device', 'result', 'training', 'name', 'geo'];
var DEV_HEADERS = ['phone', 'device', 'firstSeen', 'training'];

// ====== نقطة الدخول ======
function now_() { return Date.now(); }

function doGet() {
  return json_({ ok: true, msg: 'attendance backend is running', version: 3 });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = route_(req);
  } catch (err) {
    if (err && err.appCode) {
      out = { ok: false, code: err.appCode };
      if (err.detail !== undefined) out.detail = err.detail;
    } else {
      console.error(err && err.stack ? err.stack : String(err));
      out = { ok: false, code: 'SERVER_ERROR' };
    }
  }
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail_(code, detail) {
  var e = new Error(code);
  e.appCode = code;
  e.detail = detail;
  throw e;
}

function route_(req) {
  switch (req.action) {
    case 'ping':            return { ok: true };
    case 'info':            return info_(req);
    case 'checkin':         return checkin_(req);
    case 'init':            return init_(req);
    case 'start':           return start_(req);
    case 'token':           return token_(req);
    case 'stop':            return stop_(req);
    case 'stats':           return stats_(req);
    case 'absent':          return absent_(req);
    case 'manual':          return manual_(req);
    case 'unbind':          return unbind_(req);
    case 'summary':         return summary_(req);
    case 'training_save':   return trainingSave_(req);
    case 'training_delete': return trainingDelete_(req);
    default:                return { ok: false, code: 'BAD_ACTION' };
  }
}

// ====== أدوات مساعدة ======
function props_() { return PropertiesService.getScriptProperties(); }

function secret_() {
  var p = props_();
  var s = p.getProperty('SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    p.setProperty('SECRET', s);
  }
  return s;
}

function hostAuth_(req) {
  var pw = props_().getProperty('HOST_PASSWORD');
  if (!pw) fail_('NO_PASSWORD');
  if (String(req.pw || '') !== pw) {
    Utilities.sleep(1000); // يبطّئ أي تخمين للباسورد
    fail_('BAD_PASSWORD');
  }
}

function withLock_(fn, ms) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(ms || 25000); } catch (e) { fail_('BUSY'); }
  try { return fn(); } finally { lock.releaseLock(); }
}

function putCache_(key, obj, sec) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(obj), sec); } catch (e) { /* كبير أو غير متاح: مش مشكلة */ }
}
function getCache_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}
function delCache_(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}

function extraSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  } else if (sh.getLastColumn() < headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function firstName_(v) { return String(v == null ? '' : v).trim().split(/\s+/)[0]; }

// رقم الموبايل: بيقبل 010.. أو 10.. أو +2010.. وأرقام عربي. بنقارن بآخر 10 أرقام.
function normPhone_(v) {
  var s = String(v == null ? '' : v).replace(/[٠-٩]/g, function (c) {
    return String(c.charCodeAt(0) - 0x0660);
  }).replace(/[۰-۹]/g, function (c) {
    return String(c.charCodeAt(0) - 0x06F0);
  });
  var d = s.replace(/\D/g, '');
  return d.length >= 10 ? d.slice(-10) : '';
}
function dispPhone_(raw) {
  var d = normPhone_(raw);
  if (!d) return String(raw == null ? '' : raw);
  return d.charAt(0) === '1' ? '0' + d : d;
}

function timeStr_() { return Utilities.formatDate(new Date(now_()), CONFIG.TZ, 'HH:mm'); }
function stampStr_() { return Utilities.formatDate(new Date(now_()), CONFIG.TZ, 'yyyy-MM-dd HH:mm'); }

// ====== التدريبات ======
function parseDays_(v) {
  var out = [];
  String(v || '').split(/[,\n،]+/).forEach(function (x) {
    x = x.trim();
    if (x && out.indexOf(x) < 0) out.push(x);
  });
  return out;
}

function parseSheetId_(v) {
  v = String(v || '').trim();
  if (!v) return '';
  var m = v.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(v)) return v;
  fail_('BAD_SHEET');
}

function numOrNull_(v) {
  if (v === '' || v == null) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

function rowToTraining_(r) {
  var geo = String(r[8] || 'warn');
  if (geo !== 'off' && geo !== 'warn' && geo !== 'block') geo = 'warn';
  var rad = numOrNull_(r[7]);
  return {
    id: String(r[0]), name: String(r[1]), sheet: String(r[2] || ''), tab: String(r[3] || ''),
    days: parseDays_(r[4]), lat: numOrNull_(r[5]), lng: numOrNull_(r[6]),
    radius: rad > 0 ? rad : CONFIG.DEFAULT_RADIUS, geo: geo
  };
}

function trainings_() {
  var c = getCache_('trs');
  if (c) return c;
  var rows = extraSheet_('Trainings', TR_HEADERS).getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() && String(rows[i][1]).trim()) list.push(rowToTraining_(rows[i]));
  }
  putCache_('trs', list, CONFIG.TRAININGS_CACHE_SEC);
  return list;
}

function getTraining_(id) {
  var list = trainings_();
  for (var i = 0; i < list.length; i++) if (list[i].id === String(id || '')) return list[i];
  return null;
}

function openSheet_(tr) {
  var ss = tr.sheet ? SpreadsheetApp.openById(tr.sheet) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = tr.tab ? ss.getSheetByName(tr.tab) : ss.getSheets()[0];
  if (!sh) fail_('NO_TAB');
  return sh;
}

function findCol_(h, cands) {
  var low = h.map(function (x) { return String(x).toLowerCase(); });
  for (var i = 0; i < cands.length; i++) {
    var j = low.indexOf(cands[i].toLowerCase());
    if (j >= 0) return j;
  }
  return -1;
}

// سياق التدريب: العناوين + خريطة (رقم -> صف) بكاش 5 دقايق. fresh=true يتجاهل الكاش.
function loadCtx_(tr, fresh) {
  var sh = openSheet_(tr);
  var ck = 'tr_' + tr.id;
  if (!fresh) {
    var c = getCache_(ck);
    if (c) { c.sh = sh; c.fresh = false; return c; }
  }
  var hr = CONFIG.HEADER_ROW;
  var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if (lastCol < 1) fail_('MISSING_HEADERS', []);
  var h = sh.getRange(hr, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  var pc = findCol_(h, CONFIG.PHONE_HEADERS), nc = findCol_(h, CONFIG.NAME_HEADERS);
  if (pc < 0 || nc < 0) fail_('MISSING_HEADERS', h);
  var m = {};
  if (lastRow > hr) {
    var n = lastRow - hr;
    var ph = sh.getRange(hr + 1, pc + 1, n, 1).getValues();
    var nm = sh.getRange(hr + 1, nc + 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var p = normPhone_(ph[i][0]);
      if (p && !m[p]) m[p] = [hr + 1 + i, String(nm[i][0]).trim()];
    }
  }
  var ctx = { h: h, pc: pc, nc: nc, m: m };
  putCache_(ck, ctx, CONFIG.TRAINEES_CACHE_SEC);
  ctx.sh = sh; ctx.fresh = true;
  return ctx;
}

function ensureCol_(tr, header) {
  var ctx = loadCtx_(tr, true);
  if (ctx.h.indexOf(header) < 0) {
    ctx.sh.getRange(CONFIG.HEADER_ROW, ctx.h.length + 1).setValue(header);
    delCache_('tr_' + tr.id);
    ctx = loadCtx_(tr, true);
  }
  return ctx;
}

function isProtected_(name) {
  var l = name.toLowerCase();
  if (CONFIG.PROTECTED_HEADERS.indexOf(l) >= 0) return true;
  if (name === CONFIG.SUMMARY_HEADER) return true;
  if (name.slice(-CONFIG.OUT_SUFFIX.length) === CONFIG.OUT_SUFFIX) return true;
  var all = CONFIG.PHONE_HEADERS.concat(CONFIG.NAME_HEADERS);
  for (var i = 0; i < all.length; i++) if (all[i].toLowerCase() === l) return true;
  return false;
}

function colName_(s) { return s.kind === 'out' ? s.day + CONFIG.OUT_SUFFIX : s.day; }

function pubTraining_(t) {
  return { id: t.id, name: t.name, sheet: t.sheet, tab: t.tab, days: t.days, lat: t.lat, lng: t.lng, radius: t.radius, geo: t.geo };
}

function trainingSave_(req) {
  hostAuth_(req);
  var name = String(req.name || '').trim();
  if (!name || name.length > 80) fail_('BAD_NAME');
  var geo = String(req.geo || 'warn');
  if (geo !== 'off' && geo !== 'warn' && geo !== 'block') geo = 'warn';
  var sheetId = parseSheetId_(req.sheet);
  var lat = numOrNull_(req.lat), lng = numOrNull_(req.lng);
  if ((lat !== null && Math.abs(lat) > 90) || (lng !== null && Math.abs(lng) > 180)) fail_('BAD_LOCATION');
  var rad = numOrNull_(req.radius);
  if (!(rad > 0)) rad = CONFIG.DEFAULT_RADIUS;
  var days = parseDays_(req.days);
  var tab = String(req.tab || '').trim();

  return withLock_(function () {
    delCache_('trs');
    var sh = extraSheet_('Trainings', TR_HEADERS);
    var rows = sh.getDataRange().getValues();
    var id = String(req.id || ''), rowIdx = -1;
    for (var i = 1; i < rows.length; i++) {
      if (id && String(rows[i][0]) === id) rowIdx = i;
      if (String(rows[i][1]).trim() === name && String(rows[i][0]) !== id) fail_('DUP_NAME');
    }
    if (id && rowIdx < 0) fail_('NO_TRAINING');
    if (!id) id = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toLowerCase();

    var tr = { id: id, name: name, sheet: sheetId, tab: tab, days: days, lat: lat, lng: lng, radius: rad, geo: geo };
    var check;
    try {
      check = loadCtx_(tr, true);
    } catch (err) {
      if (err && err.appCode) throw err;
      console.error(err && err.stack ? err.stack : String(err));
      fail_('CANT_OPEN');
    }
    var vals = [id, name, sheetId, tab, days.join(','), lat === null ? '' : lat, lng === null ? '' : lng, rad, geo,
      rowIdx >= 0 ? rows[rowIdx][9] : new Date(now_())];
    if (rowIdx >= 0) sh.getRange(rowIdx + 1, 1, 1, TR_HEADERS.length).setValues([vals]);
    else sh.appendRow(vals);
    delCache_('trs');
    return { ok: true, training: pubTraining_(tr), trainees: Object.keys(check.m).length };
  });
}

function trainingDelete_(req) {
  hostAuth_(req);
  var id = String(req.id || '');
  return withLock_(function () {
    var sh = extraSheet_('Trainings', TR_HEADERS);
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === id) sh.deleteRow(i + 1);
    }
    delCache_('trs'); delCache_('tr_' + id);
    var all = sessions_(), changed = false;
    Object.keys(all).forEach(function (k) { if (all[k].tid === id) { delete all[k]; changed = true; } });
    if (changed) saveSessions_(all);
    return { ok: true };
  });
}

// ====== الجلسات ======
function sessions_() {
  var raw = props_().getProperty('SESSIONS');
  try { return raw ? JSON.parse(raw) : {}; } catch (e) { return {}; }
}
function saveSessions_(o) { props_().setProperty('SESSIONS', JSON.stringify(o)); }

function session_(sid) {
  var s = sessions_()[String(sid || '')];
  if (!s) return null;
  if (now_() - s.startedAt > CONFIG.SESSION_MAX_AGE_MS) return null;
  return s;
}

function pubSession_(s) {
  var t = getTraining_(s.tid);
  return { sid: s.sid, tid: s.tid, training: t ? t.name : '', day: s.day, rot: s.rot, kind: s.kind, startedAt: s.startedAt };
}

// الكود: HMAC من (رقم الجلسة + اليوم + رقم الفترة الزمنية)
function tokenFor_(sid, day, slot) {
  var sig = Utilities.computeHmacSha256Signature(sid + '|' + day + '|' + slot, secret_());
  var A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', out = '';
  for (var i = 0; i < 10; i++) out += A.charAt((sig[i] & 255) % 32);
  return out;
}
function slotOf_(s) { return Math.floor(now_() / (s.rot * 1000)); }
function validToken_(s, t) {
  var slot = slotOf_(s);
  for (var k = 0; k <= CONFIG.GRACE_SLOTS; k++) {
    if (t === tokenFor_(s.sid, s.day, slot - k)) return true;
  }
  return false;
}

// ====== أوامر المدرب ======
function init_(req) {
  hostAuth_(req);
  var list = trainings_().map(pubTraining_);
  var ss = sessions_(), act = [];
  Object.keys(ss).forEach(function (k) {
    if (now_() - ss[k].startedAt <= CONFIG.SESSION_MAX_AGE_MS && getTraining_(ss[k].tid)) act.push(pubSession_(ss[k]));
  });
  return { ok: true, trainings: list, sessions: act };
}

function start_(req) {
  hostAuth_(req);
  var tr = getTraining_(req.tid);
  if (!tr) fail_('NO_TRAINING');
  var day = String(req.day || '').trim();
  var kind = req.kind === 'out' ? 'out' : 'in';
  var rot = parseInt(req.rot, 10);
  if (!(rot >= CONFIG.MIN_ROTATION && rot <= CONFIG.MAX_ROTATION)) fail_('BAD_ROTATION');
  if (!day || isProtected_(day)) fail_('BAD_DAY');

  var ctx = loadCtx_(tr, true);
  if (ctx.h.indexOf(day) < 0) {
    if (tr.days.indexOf(day) >= 0) ctx = ensureCol_(tr, day);
    else fail_('NO_DAY_COLUMN');
  }
  if (kind === 'out') ensureCol_(tr, day + CONFIG.OUT_SUFFIX);

  return withLock_(function () {
    var all = sessions_();
    Object.keys(all).forEach(function (k) {
      var o = all[k];
      if (now_() - o.startedAt > CONFIG.SESSION_MAX_AGE_MS || (o.tid === tr.id && o.day === day && o.kind === kind)) delete all[k];
    });
    var s = {
      sid: Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase(),
      tid: tr.id, day: day, kind: kind, rot: rot, startedAt: now_()
    };
    all[s.sid] = s;
    saveSessions_(all);
    return { ok: true, session: pubSession_(s) };
  }, 10000);
}

function token_(req) {
  hostAuth_(req);
  var s = session_(req.sid);
  if (!s) fail_('NO_SESSION');
  var slot = slotOf_(s);
  return { ok: true, token: tokenFor_(s.sid, s.day, slot), serverNow: now_(), validUntil: (slot + 1) * s.rot * 1000 };
}

function stop_(req) {
  hostAuth_(req);
  return withLock_(function () {
    var all = sessions_();
    delete all[String(req.sid || '')];
    saveSessions_(all);
    return { ok: true };
  }, 10000);
}

// عداد الحضور اللحظي وآخر الأسماء
function stats_(req) {
  hostAuth_(req);
  var s = session_(req.sid);
  if (!s) fail_('NO_SESSION');
  var tr = getTraining_(s.tid);
  if (!tr) fail_('NO_TRAINING');
  var counts = count_(tr, s.day, s.kind);
  return { ok: true, present: counts.present, total: counts.total, recent: getCache_('rc_' + s.sid) || [], serverNow: now_() };
}

function count_(tr, day, kind) {
  var ctx = loadCtx_(tr, false);
  var col = ctx.h.indexOf(kind === 'out' ? day + CONFIG.OUT_SUFFIX : day);
  var inCol = ctx.h.indexOf(day);
  if (col < 0 || inCol < 0) { ctx = loadCtx_(tr, true); col = ctx.h.indexOf(kind === 'out' ? day + CONFIG.OUT_SUFFIX : day); inCol = ctx.h.indexOf(day); }
  if (col < 0 || inCol < 0) fail_('NO_DAY_COLUMN');
  var hr = CONFIG.HEADER_ROW, lastRow = ctx.sh.getLastRow(), present = 0, total = 0;
  if (lastRow > hr) {
    var rows = ctx.sh.getRange(hr + 1, 1, lastRow - hr, ctx.h.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (!normPhone_(rows[i][ctx.pc])) continue;
      var hasIn = String(rows[i][inCol]).trim() !== '';
      var has = String(rows[i][col]).trim() !== '';
      if (kind === 'out') { if (hasIn) { total++; if (has) present++; } }
      else { total++; if (has) present++; }
    }
  }
  return { present: present, total: total };
}

// الغايبين (أو اللي ماسجلوش خروج)
function absent_(req) {
  hostAuth_(req);
  var tr = getTraining_(req.tid);
  if (!tr) fail_('NO_TRAINING');
  var day = String(req.day || '').trim(), kind = req.kind === 'out' ? 'out' : 'in';
  var ctx = loadCtx_(tr, true);
  var col = ctx.h.indexOf(kind === 'out' ? day + CONFIG.OUT_SUFFIX : day), inCol = ctx.h.indexOf(day);
  if (col < 0 || inCol < 0) fail_('NO_DAY_COLUMN');
  var hr = CONFIG.HEADER_ROW, lastRow = ctx.sh.getLastRow(), list = [], present = 0, total = 0;
  if (lastRow > hr) {
    var rows = ctx.sh.getRange(hr + 1, 1, lastRow - hr, ctx.h.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var p = normPhone_(rows[i][ctx.pc]);
      if (!p) continue;
      var hasIn = String(rows[i][inCol]).trim() !== '';
      var has = String(rows[i][col]).trim() !== '';
      if (kind === 'out' && !hasIn) continue;
      total++;
      if (has) present++;
      else list.push({ name: String(rows[i][ctx.nc]).trim(), phone: dispPhone_(rows[i][ctx.pc]), key: p });
    }
  }
  return { ok: true, list: list, present: present, total: total };
}

// تسجيل يدوي من المدرب (من غير كود ولا ربط موبايل)
function manual_(req) {
  hostAuth_(req);
  var tr = getTraining_(req.tid);
  if (!tr) fail_('NO_TRAINING');
  var day = String(req.day || '').trim(), kind = req.kind === 'out' ? 'out' : 'in';
  var phone = normPhone_(req.phone);
  if (!phone) fail_('BAD_PHONE');
  return withLock_(function () {
    var ctx = loadCtx_(tr, true);
    var col = ctx.h.indexOf(kind === 'out' ? day + CONFIG.OUT_SUFFIX : day), inCol = ctx.h.indexOf(day);
    if (col < 0 || inCol < 0) fail_('NO_DAY_COLUMN');
    var rec = ctx.m[phone];
    if (!rec) fail_('NOT_FOUND');
    var s = { sid: 'MANUAL', day: day, kind: kind };
    if (kind === 'out' && String(ctx.sh.getRange(rec[0], inCol + 1).getValue()).trim() === '') fail_('NOT_CHECKED_IN');
    var cell = ctx.sh.getRange(rec[0], col + 1);
    var already = String(cell.getValue()).trim() !== '';
    if (!already) writeMark_(cell, kind, 'يدوي ');
    log_(s, tr, phone, '-', already ? 'MANUAL_ALREADY' : 'MANUAL', '', firstName_(rec[1]));
    return { ok: true, already: already, name: firstName_(rec[1]) };
  });
}

function writeMark_(cell, kind, prefix) {
  if (kind === 'out') {
    cell.setNumberFormat('@');
    cell.setValue(timeStr_());
  } else {
    cell.setValue(CONFIG.PRESENT_VALUE);
    cell.setNote((prefix || 'QR ') + stampStr_());
  }
}

// فك ربط موبايل متدرب (لو غيّر موبايله أو الموبايل وقع)
function unbind_(req) {
  hostAuth_(req);
  var phone = normPhone_(req.phone);
  if (!phone) fail_('BAD_PHONE');
  return withLock_(function () {
    var sh = extraSheet_('Devices', DEV_HEADERS);
    var rows = sh.getDataRange().getValues(), removed = 0;
    for (var i = rows.length - 1; i >= 1; i--) {
      if (normPhone_(rows[i][0]) === phone) { sh.deleteRow(i + 1); removed++; }
    }
    delCache_('dev_maps');
    return { ok: true, removed: removed };
  });
}

// عمود "نسبة الحضور" + تلوين اللي تحت الحد
function summary_(req) {
  hostAuth_(req);
  var tr = getTraining_(req.tid);
  if (!tr) fail_('NO_TRAINING');
  return withLock_(function () {
    var ctx = ensureCol_(tr, CONFIG.SUMMARY_HEADER);
    var hr = CONFIG.HEADER_ROW, sh = ctx.sh;
    var dayCols = [];
    tr.days.forEach(function (d) { var i = ctx.h.indexOf(d); if (i >= 0) dayCols.push(i + 1); });
    if (!dayCols.length) fail_('NO_DAYS');
    var sc = ctx.h.indexOf(CONFIG.SUMMARY_HEADER) + 1;
    var lastRow = Math.max(sh.getLastRow(), hr + 1), n = lastRow - hr;
    var pL = colLetter_(ctx.pc + 1), den = [], formulas = [];
    dayCols.forEach(function (c) { var L = colLetter_(c); den.push('(COUNTA(' + L + '$' + (hr + 1) + ':' + L + ')>0)'); });
    for (var i = 0; i < n; i++) {
      var r = hr + 1 + i, num = [];
      dayCols.forEach(function (c) { num.push('(' + colLetter_(c) + r + '<>"")'); });
      formulas.push(['=IF($' + pL + r + '="","",IFERROR((' + num.join('+') + ')/MAX(1,' + den.join('+') + '),0))']);
    }
    var range = sh.getRange(hr + 1, sc, n, 1);
    range.setFormulas(formulas);
    range.setNumberFormat('0%');
    var rules = sh.getConditionalFormatRules().filter(function (r) {
      var rg = r.getRanges()[0];
      return !rg || rg.getColumn() !== sc;
    });
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(CONFIG.SUMMARY_THRESHOLD).setBackground('#f4c7c3').setRanges([range]).build());
    sh.setConditionalFormatRules(rules);
    return { ok: true, column: CONFIG.SUMMARY_HEADER, days: dayCols.length, threshold: CONFIG.SUMMARY_THRESHOLD };
  });
}

// ====== جانب المتدرب ======
function info_(req) {
  var s = session_(req.sid);
  if (!s) fail_('SESSION_CLOSED');
  var tr = getTraining_(s.tid);
  if (!tr) fail_('SESSION_CLOSED');
  return { ok: true, training: tr.name, day: s.day, kind: s.kind, geo: tr.geo, tokenOk: validToken_(s, String(req.t || '')) };
}

function checkin_(req) {
  var s = session_(req.sid);
  if (!s) fail_('SESSION_CLOSED');
  if (!validToken_(s, String(req.t || ''))) fail_('TOKEN_EXPIRED');
  var phone = normPhone_(req.phone);
  if (!phone) fail_('BAD_PHONE');
  var device = String(req.device || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(device)) fail_('BAD_DEVICE');
  var g = null, la = Number(req.lat), ln = Number(req.lng);
  if (req.lat != null && req.lng != null && isFinite(la) && isFinite(ln) && Math.abs(la) <= 90 && Math.abs(ln) <= 180) {
    var ac = Number(req.acc);
    g = { lat: la, lng: ln, acc: isFinite(ac) && ac >= 0 ? ac : 0 };
  }
  return withLock_(function () {
    var tr = getTraining_(s.tid);
    if (!tr) fail_('SESSION_CLOSED');
    return process_(s, tr, phone, device, g);
  });
}

function dist_(a, b, c, d) {
  var R = 6371000, rad = Math.PI / 180;
  var dLat = (c - a) * rad, dLng = (d - b) * rad;
  var x = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(a * rad) * Math.cos(c * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}

function devMaps_() {
  var c = getCache_('dev_maps');
  if (c) return c;
  var rows = extraSheet_('Devices', DEV_HEADERS).getDataRange().getValues();
  var m = { p: {}, d: {} };
  for (var i = 1; i < rows.length; i++) {
    var p = normPhone_(rows[i][0]), d = String(rows[i][1]);
    if (!p || !d) continue;
    m.p[p] = d; m.d[d] = p;
  }
  putCache_('dev_maps', m, 21600);
  return m;
}

function process_(s, tr, phone, device, g) {
  var kind = s.kind, col = colName_(s);
  var ctx = loadCtx_(tr, false);
  var ci = ctx.h.indexOf(col), ii = ctx.h.indexOf(s.day);
  if (!ctx.m[phone] || ci < 0 || ii < 0) {
    if (!ctx.fresh) { ctx = loadCtx_(tr, true); ci = ctx.h.indexOf(col); ii = ctx.h.indexOf(s.day); }
  }
  if (ci < 0 || ii < 0) fail_('NO_DAY_COLUMN');
  var rec = ctx.m[phone];
  if (!rec) return done_(s, tr, phone, device, 'NOT_FOUND', '', '', { ok: false, code: 'NOT_FOUND' });
  var name = firstName_(rec[1]);

  // ربط الموبايل (عام لكل التدريبات): كل رقم بموبايل واحد وكل موبايل بمتدرب واحد
  var maps = devMaps_();
  var bound = maps.p[phone], owner = maps.d[device];
  if (bound && bound !== device)
    return done_(s, tr, phone, device, 'DEVICE_MISMATCH', '', name, { ok: false, code: 'DEVICE_MISMATCH' });
  if (!bound && owner && owner !== phone)
    return done_(s, tr, phone, device, 'DEVICE_USED_FOR_OTHER', '', name, { ok: false, code: 'DEVICE_USED_FOR_OTHER' });

  // الموقع: warn بيسجل بس، block بيرفض
  var geo = 'OFF', far = false;
  if (tr.geo !== 'off' && tr.lat !== null && tr.lng !== null) {
    if (!g) {
      geo = 'NOGPS';
      if (tr.geo === 'block') return done_(s, tr, phone, device, 'NO_LOCATION', geo, name, { ok: false, code: 'NO_LOCATION' });
    } else {
      var d = dist_(tr.lat, tr.lng, g.lat, g.lng);
      far = d > tr.radius + Math.min(g.acc, 150);
      geo = (far ? 'FAR ' : 'NEAR ') + Math.round(d) + 'm';
      if (far && tr.geo === 'block') return done_(s, tr, phone, device, 'TOO_FAR', geo, name, { ok: false, code: 'TOO_FAR' });
    }
  }

  if (kind === 'out' && String(ctx.sh.getRange(rec[0], ii + 1).getValue()).trim() === '')
    return done_(s, tr, phone, device, 'NOT_CHECKED_IN', geo, name, { ok: false, code: 'NOT_CHECKED_IN' });

  if (!bound) {
    extraSheet_('Devices', DEV_HEADERS).appendRow(["'" + phone, device, new Date(now_()), tr.name]);
    maps.p[phone] = device; maps.d[device] = phone;
    putCache_('dev_maps', maps, 21600);
  }

  var cell = ctx.sh.getRange(rec[0], ci + 1);
  var already = String(cell.getValue()).trim() !== '';
  if (!already) writeMark_(cell, kind, 'QR ');
  if (!already) pushRecent_(s.sid, name, far || geo === 'NOGPS');
  return done_(s, tr, phone, device, already ? 'ALREADY' : 'OK', geo, name,
    { ok: true, already: already, name: name, day: s.day, kind: kind, training: tr.name });
}

function pushRecent_(sid, name, flag) {
  var list = getCache_('rc_' + sid) || [];
  list.unshift({ n: name, t: timeStr_(), w: flag ? 1 : 0 });
  putCache_('rc_' + sid, list.slice(0, 8), 21600);
}

function log_(s, tr, phone, device, result, geo, name) {
  var lg = extraSheet_('Log', LOG_HEADERS);
  lg.appendRow([new Date(now_()), s.sid, s.day + (s.kind === 'out' ? CONFIG.OUT_SUFFIX : ''), "'" + phone, device, result, tr.name, name || '', geo || '']);
}
function done_(s, tr, phone, device, result, geo, name, out) {
  log_(s, tr, phone, device, result, geo, name);
  return out;
}

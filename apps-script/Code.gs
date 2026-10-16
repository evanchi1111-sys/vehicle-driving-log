/**
 * 車輛借用系統 後端（Google Apps Script）
 *
 * 資料存在一份 Google 試算表，每個集合一個工作表（欄位：id、json、updatedAt），
 * 另外會自動產生「行車紀錄（檢視用）」工作表，方便直接閱讀、列印與備份。
 *
 * 使用前請先修改下面兩個密碼，部署步驟見 docs/Google設定步驟.md。
 */

// ===== 一定要修改這兩個密碼（至少 6 個字元，不要使用常見的數字）=====
var USER_PIN  = 'CHANGE-ME-USER';   // 登記密碼：給同仁，可登記行車紀錄
var ADMIN_PIN = 'CHANGE-ME-ADMIN';  // 管理者密碼：可修改、刪除、管理清單、備份還原
// ====================================================================

// write: 'admin' = 只有管理者可寫入；'create' = 一般同仁只能新增，不能改或刪既有資料
// read:  'admin' = 只有管理者可讀取
var COLLECTIONS = {
  vehicles:       {write: 'admin'},
  vehicleLogs:    {write: 'create'},
  settings:       {write: 'admin'},
  vehicleRepairs: {write: 'admin'},
  roster:         {write: 'admin'},
  places:         {write: 'admin'},
  vehicleDeleted: {write: 'admin', read: 'admin'}
};
var VIEW_SHEET = '行車紀錄（檢視用）';
var MAX_JSON = 40000;       // 一個儲存格上限約 50000 字元
var MAX_BATCH = 200;
var MAX_FAILS = 30;         // 10 分鐘內密碼錯誤累計達此數，暫時鎖定 10 分鐘（防止猜密碼）
var ID_RE = /^[A-Za-z0-9_\-.~:@+]{1,200}$/;

/** 第一次使用請在編輯器執行一次：建立試算表與工作表，並在「執行記錄」顯示試算表網址 */
function setup() {
  var s = ss_();
  Object.keys(COLLECTIONS).forEach(function (n) { sheet_(n); });
  rebuildView_();
  Logger.log('試算表已建立：' + s.getUrl());
}

/** 選用：在編輯器執行，檢查後端是否正常（會新增再刪除一筆測試資料） */
function selfTest() {
  var r1 = handle_({pin: ADMIN_PIN, op: 'ping'});
  var r2 = handle_({pin: ADMIN_PIN, op: 'set', collection: 'places', id: 'selftest', data: {rows: ['測試']}});
  var r3 = handle_({pin: ADMIN_PIN, op: 'list', force: true});
  var r4 = handle_({pin: ADMIN_PIN, op: 'delete', collection: 'places', id: 'selftest'});
  Logger.log(JSON.stringify({ping: r1.ok && r1.role, set: r2.ok, list: r3.ok && Object.keys(r3.data), delete: r4.ok}));
}

function doGet() {
  return ContentService.createTextOutput('vehicle log api ok');
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    out = handle_(req);
  } catch (err) {
    out = {ok: false, error: err && err.code ? err.code : 'bad_request', message: String(err && err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- 核心

function handle_(req) {
  var role = auth_(req.pin);
  var op = req.op;
  if (op === 'ping') return {ok: true, role: role, rev: rev_()};
  if (op === 'list') {
    var rev = rev_();
    if (req.rev != null && Number(req.rev) === rev && !req.force) return {ok: true, changed: false, role: role, rev: rev};
    var data = {};
    Object.keys(COLLECTIONS).forEach(function (n) {
      if (COLLECTIONS[n].read === 'admin' && role !== 'admin') return;
      data[n] = readAll_(n);
    });
    return {ok: true, changed: true, role: role, rev: rev, data: data};
  }
  if (op === 'set' || op === 'delete') return write_(role, [req]);
  if (op === 'batch') {
    if (role !== 'admin') throw fail_('forbidden');
    if (!Array.isArray(req.ops) || req.ops.length > MAX_BATCH) throw fail_('bad_request', '批次筆數超過上限');
    return write_(role, req.ops);
  }
  throw fail_('bad_request', '未知的操作');
}

function write_(role, ops) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var touchedLogs = false;
    ops.forEach(function (o) {
      var name = o.collection, cfg = COLLECTIONS[name];
      if (!cfg || !Object.prototype.hasOwnProperty.call(COLLECTIONS, name)) throw fail_('bad_request', '未知的集合');
      if (typeof o.id !== 'string' || !ID_RE.test(o.id)) throw fail_('bad_request', '編號格式不正確');
      var sh = sheet_(name), row = findRow_(sh, o.id);
      if (o.op === 'delete') {
        if (role !== 'admin') throw fail_('forbidden');
        if (row) sh.deleteRow(row);
      } else if (o.op === 'set') {
        if (!o.data || typeof o.data !== 'object' || Array.isArray(o.data)) throw fail_('bad_request', '資料格式不正確');
        var json = JSON.stringify(o.data);
        if (json.length > MAX_JSON) throw fail_('too_large');
        var canWrite = role === 'admin' || (cfg.write === 'create' && !row);
        if (!canWrite) throw fail_('forbidden');
        var now = new Date().toISOString();
        if (row) sh.getRange(row, 1, 1, 3).setValues([[o.id, json, now]]);
        else sh.appendRow([o.id, json, now]);
      } else {
        throw fail_('bad_request', '未知的操作');
      }
      if (name === 'vehicleLogs') touchedLogs = true;
    });
    if (touchedLogs) rebuildView_();
    var rev = bump_();
    return {ok: true, rev: rev};
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- 密碼與鎖定

function auth_(pin) {
  if (String(USER_PIN).indexOf('CHANGE-ME') === 0 || String(ADMIN_PIN).indexOf('CHANGE-ME') === 0) throw fail_('not_configured', '請先修改程式裡的兩個密碼');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw fail_('locked', '密碼錯誤次數過多，請 10 分鐘後再試');
  pin = String(pin == null ? '' : pin);
  if (pin === String(ADMIN_PIN)) return 'admin';
  if (pin === String(USER_PIN)) return 'user';
  cache.put('fails', String(fails + 1), 600);
  throw fail_('auth', '密碼不正確');
}

function fail_(code, message) {
  var e = new Error(message || code);
  e.code = code;
  return e;
}

// ---------------------------------------------------------------- 試算表

var SS_CACHE = null;   // 一次執行只開啟一次試算表，加快速度

function ss_() {
  if (SS_CACHE) return SS_CACHE;
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) {
    SS_CACHE = SpreadsheetApp.openById(id);
    return SS_CACHE;
  }
  var s = SpreadsheetApp.create('車輛借用系統資料');
  props.setProperty('SHEET_ID', s.getId());
  SS_CACHE = s;
  return s;
}

function sheet_(name) {
  var s = ss_();
  var sh = s.getSheetByName(name);
  if (!sh) {
    sh = s.insertSheet(name);
    sh.getRange(1, 1, 1, 3).setValues([['id', 'json', 'updatedAt']]);
    sh.getRange('A:B').setNumberFormat('@');
  }
  return sh;
}

function findRow_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return 0;
}

function readAll_(name) {
  var sh = sheet_(name), n = sh.getLastRow();
  if (n < 2) return [];
  var v = sh.getRange(2, 1, n - 1, 2).getValues(), out = [];
  for (var i = 0; i < v.length; i++) {
    try {
      var o = JSON.parse(v[i][1]);
      o.id = String(v[i][0]);
      out.push(o);
    } catch (e) { /* 略過損毀的列 */ }
  }
  return out;
}

function rev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('REV') || 0);
}

function bump_() {
  var props = PropertiesService.getScriptProperties();
  var r = Number(props.getProperty('REV') || 0) + 1;
  props.setProperty('REV', String(r));
  return r;
}

/** 重新產生方便閱讀的「行車紀錄（檢視用）」工作表（依日期與時間排序，不可手動修改） */
function rebuildView_() {
  var logs = readAll_('vehicleLogs').sort(function (a, b) {
    var x = (a.date || '') + ' ' + (a.outTime || ''), y = (b.date || '') + ' ' + (b.outTime || '');
    return x < y ? -1 : x > y ? 1 : 0;
  });
  var rows = [['車號', '日期', '單位', '駕駛人', '往返地點', '出發里程數', '返回里程數', '出發時間', '返回時間', '外觀', '車況', '說明', '補登原因', '修改原因']];
  logs.forEach(function (l) {
    rows.push([l.plate || '', l.date || '', l.unit || '', l.driver || '', l.route || '', l.outMileage == null ? '' : l.outMileage,
      l.retMileage == null ? '' : l.retMileage, l.outTime || '', l.retTime || '',
      l.look === 'bad' ? '✕' : '✓', l.cond === 'bad' ? '✕' : '✓', l.note || '', l.backfillReason || '', l.editedReason || '']);
  });
  var s = ss_();
  var sh = s.getSheetByName(VIEW_SHEET) || s.insertSheet(VIEW_SHEET);
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

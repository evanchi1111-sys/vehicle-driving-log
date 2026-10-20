/**
 * 車輛借用系統 後端（Google Apps Script）
 *
 * 資料存在一份 Google 試算表，每個集合一個工作表（欄位：id、json、updatedAt），
 * 另外會自動產生「行車紀錄（檢視用）」工作表，方便直接閱讀、列印與備份。
 *
 * 使用前請先修改下面兩個密碼，部署步驟見 docs/Google設定步驟.md。
 */

// ===== 一定要修改這兩個密碼（至少 6 個字元，不要使用常見的數字）=====
var USER_PIN  = 'CHANGE-ME-USER';   // 登記密碼：給同仁。想讓同仁免密碼登記，改成 ''（兩個單引號，中間不放字）；仍只能新增行車紀錄
var ADMIN_PIN = 'CHANGE-ME-ADMIN';  // 管理者密碼（一定要改，至少 6 個字元）：可修改、刪除、管理清單、備份還原
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
  if ((pin == null || String(pin) === '') && String(USER_PIN) === '') return 'user';  // 免密碼登記：同仁只有新增行車紀錄的權限
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw fail_('locked', '密碼錯誤次數過多，請 10 分鐘後再試');
  pin = String(pin == null ? '' : pin);
  if (pin === String(ADMIN_PIN)) return 'admin';
  if (pin === String(USER_PIN)) return 'user';
  if (pin === '') throw fail_('auth', '需要密碼');  // 沒帶密碼的探測不計入錯誤次數
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

// ---------------------------------------------------------------- 每週自動備份：寄送行車紀錄（CSV）與完整備份（JSON）

/**
 * 收件人：預設寄給執行這個腳本的帳號（也就是你自己）。
 * 行車紀錄含人員姓名與行程，建議維持寄給自己，不要改成別人的信箱。
 */
var BACKUP_TO = '';   // 留空 = 寄給腳本擁有者

/** 只需在編輯器執行一次：設定每週一早上 7～8 點自動寄出。重複執行不會建立重複的排程 */
function installWeeklyBackup() {
  removeWeeklyBackup();
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  Logger.log('已設定：每週一 07:00～08:00 自動寄送備份。');
}

/** 取消每週自動寄送 */
function removeWeeklyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklyBackup') ScriptApp.deleteTrigger(t);
  });
}

/** 由排程呼叫，也可以在編輯器手動執行來測試（會立刻寄出一封信） */
function weeklyBackup() {
  var to = BACKUP_TO || Session.getEffectiveUser().getEmail();
  try {
    // 直接從原始資料產生表格（不經過試算表儲存格），避免 Google 試算表把 "08:30"、"2026-10-19"
    // 自動轉成日期時間物件，造成 CSV 裡的時間格式錯誤
    var values = csvRows_();
    var count = Math.max(0, values.length - 1);   // 扣掉標題列
    var csv = '﻿' + values.map(function (row) {
      return row.map(csvCell_).join(',');
    }).join('\r\n');   // 開頭的 BOM 讓 Excel 正確顯示中文
    // 完整備份：格式與網頁「備份全部資料」相同，可用網頁的「從備份檔還原」匯入
    var c = {}, lines = [], total = 0;
    Object.keys(COLLECTIONS).forEach(function (n) {
      var rows = readAll_(n);
      c[n] = rows;
      total += rows.length;
      lines.push(n + '：' + rows.length + ' 筆');
    });
    var json = JSON.stringify({app: 'eip-booking', version: 1, at: new Date().toISOString(), collections: c}, null, 1);
    var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    MailApp.sendEmail({
      to: to,
      subject: '【車輛借用系統】每週備份 ' + day,
      body: '附件有兩個檔案（' + day + '）：\n' +
        '1. 行車紀錄_' + day + '.csv：行車紀錄 ' + count + ' 筆，可用 Excel 開啟閱讀、列印，不能用來還原。\n' +
        '2. 車輛借用系統備份_' + day + '.json：完整備份，共 ' + total + ' 筆資料，可還原。\n\n' +
        lines.join('\n') + '\n\n' +
        '還原方式：管理者登入網頁 → 備份與還原 → 從備份檔還原（選 JSON 檔）。\n' +
        '檔案含人員姓名與行程，請存放在有權限控管的位置，不要上傳到公開網站或 GitHub。',
      attachments: [
        Utilities.newBlob(csv, 'text/csv', '行車紀錄_' + day + '.csv'),
        Utilities.newBlob(json, 'application/json', '車輛借用系統備份_' + day + '.json')
      ]
    });
  } catch (err) {
    MailApp.sendEmail(to, '【車輛借用系統】每週備份失敗', '錯誤訊息：' + (err && err.message || err));
    throw err;
  }
}

/** 行車紀錄表格資料（欄位與「行車紀錄（檢視用）」相同，日期與時間保持文字） */
function csvRows_() {
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
  return rows;
}

/** 轉成 CSV 儲存格；文字若以 = + - @ 開頭，前面加單引號，避免被 Excel 當成公式執行 */
function csvCell_(v) {
  var s = String(v == null ? '' : v);
  if (typeof v === 'string' && /^[=+\-@]/.test(s)) s = "'" + s;
  return '"' + s.replace(/"/g, '""') + '"';
}

// ---------------------------------------------------------------- 每天自動抓取中油 95 無鉛汽油牌價（用來估算油費）

var CPC_XML_URL  = 'https://vipmbr.cpc.com.tw/cpcstn/listpricewebservice.asmx/getCPCMainProdListPrice_XML';
var CPC_JSON_URL = 'https://www.cpc.com.tw/GetOilPriceJson.aspx?type=TodayOilPriceString';

/** 只需在編輯器執行一次：每天早上 9～10 點自動檢查中油牌價（牌價沒變動就不會新增資料） */
function installFuelPriceTrigger() {
  removeFuelPriceTrigger();
  ScriptApp.newTrigger('dailyFuelPrice').timeBased().everyDays(1).atHour(9).create();
  Logger.log('已設定：每天 09:00～10:00 自動更新中油 95 無鉛汽油牌價。');
}

/** 取消每天自動更新油價 */
function removeFuelPriceTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyFuelPrice') ScriptApp.deleteTrigger(t);
  });
}

/** 由排程呼叫。失敗時寄信通知，同一種失敗 7 天內只通知一次 */
function dailyFuelPrice() {
  try {
    updateFuelPrice();
  } catch (err) {
    var props = PropertiesService.getScriptProperties();
    var lastAlert = Number(props.getProperty('FUEL_ALERT') || 0);
    if (Date.now() - lastAlert > 7 * 24 * 3600 * 1000) {
      props.setProperty('FUEL_ALERT', String(Date.now()));
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【車輛借用系統】中油油價自動更新失敗',
        '錯誤訊息：' + (err && err.message || err) + '\n\n油費估算會繼續使用最近一次成功抓到的價格。' +
        '你也可以到網頁「油價與油耗」手動新增單價。');
    }
    throw err;
  }
}

/** 抓取目前牌價並寫入。也可以在編輯器手動執行，立刻抓一次（第一次設定時請先執行這個） */
function updateFuelPrice() {
  var got = fetchCpc95_();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_('settings'), row = findRow_(sh, 'fuelAuto'), doc = {rows: []};
    if (row) {
      try { doc = JSON.parse(sh.getRange(row, 2).getValue()); } catch (e) { doc = {rows: []}; }
    }
    doc.rows = doc.rows || [];
    var ex = null;
    doc.rows.forEach(function (r) { if (r.date === got.date) ex = r; });
    if (ex && Number(ex.p95) === got.p95) {
      Logger.log('牌價沒有變動：' + got.date + ' 起 ' + got.p95 + ' 元／公升');
      return got;
    }
    if (ex) ex.p95 = got.p95; else doc.rows.push({date: got.date, p95: got.p95});
    doc.rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    doc.rows = doc.rows.slice(-400);
    doc.at = new Date().toISOString();
    var json = JSON.stringify(doc), now = doc.at;
    if (row) sh.getRange(row, 1, 1, 3).setValues([['fuelAuto', json, now]]);
    else sh.appendRow(['fuelAuto', json, now]);
    bump_();   // 讓已開啟的網頁在下次檢查時取得新價格
    Logger.log('已更新：' + got.date + ' 起 95 無鉛汽油 ' + got.p95 + ' 元／公升');
    return got;
  } finally {
    lock.releaseLock();
  }
}

/** 從中油取得 95 無鉛汽油牌價與生效日期，先試 XML，失敗再試 JSON */
function fetchCpc95_() {
  var errs = [];
  try {
    var xml = cpcGet_(CPC_XML_URL);
    var blocks = xml.split('<Table>');
    for (var i = 0; i < blocks.length; i++) {
      if (blocks[i].indexOf('<產品名稱>95無鉛汽油</產品名稱>') < 0) continue;
      var p = /<參考牌價_金額>\s*([\d.]+)\s*<\/參考牌價_金額>/.exec(blocks[i]);
      var d = /<牌價生效日期>\s*(\d{7})\s*<\/牌價生效日期>/.exec(blocks[i]);
      if (p && d) {
        return checkCpc_({date: (Number(d[1].slice(0, 3)) + 1911) + '-' + d[1].slice(3, 5) + '-' + d[1].slice(5, 7), p95: Number(p[1])});
      }
    }
    errs.push('XML 找不到 95 無鉛汽油');
  } catch (e1) {
    errs.push('XML：' + (e1 && e1.message || e1));
  }
  try {
    var j = JSON.parse(cpcGet_(CPC_JSON_URL));
    var m = /(\d+)月(\d+)日/.exec(j.PriceUpdate || '');
    if (m && j.sPrice2) {
      var now = new Date(), y = now.getFullYear();
      if (new Date(y, Number(m[1]) - 1, Number(m[2])).getTime() > now.getTime() + 7 * 24 * 3600 * 1000) y -= 1;   // 跨年（例如 1 月看到 12 月的日期）
      return checkCpc_({date: y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2), p95: Number(j.sPrice2)});
    }
    errs.push('JSON 格式不符');
  } catch (e2) {
    errs.push('JSON：' + (e2 && e2.message || e2));
  }
  throw new Error('無法取得中油油價（' + errs.join('；') + '）');
}

function cpcGet_(url) {
  var res = UrlFetchApp.fetch(url, {muteHttpExceptions: true, followRedirects: true, headers: {'User-Agent': 'Mozilla/5.0'}});
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  return res.getContentText('UTF-8');
}

/** 基本檢查，避免抓到錯誤的數字：單價要在 10～100 元之間 */
function checkCpc_(r) {
  if (!(r.p95 >= 10 && r.p95 <= 100) || !/^\d{4}-\d{2}-\d{2}$/.test(r.date)) throw new Error('抓到的資料不合理：' + JSON.stringify(r));
  return r;
}

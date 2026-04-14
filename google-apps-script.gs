/**
 * BI Dashboard — Google Apps Script Web App
 * 用途：提供 users / login_logs / action_logs 的雲端同步 API
 * 部署：發布為 Web App（執行身份：自己，訪問權限：任何人）
 *
 * 使用前：
 * 1. 建立 Google Sheet，包含 3 個 tab：users / login_logs / action_logs
 * 2. 將此腳本綁定到該 Sheet
 * 3. 修改下方 CONFIG 中的 API_KEY
 * 4. 發布為 Web App
 */

// ══════════════════════════════════════
// 設定
// ══════════════════════════════════════
var CONFIG = {
  API_KEY: 'bi-dashboard-sync-2026-secure-01',  // ← 請修改為你自己的密鑰
  SHEET_USERS: 'users',
  SHEET_LOGIN_LOGS: 'login_logs',
  SHEET_ACTION_LOGS: 'action_logs'
};

// ══════════════════════════════════════
// 進入點
// ══════════════════════════════════════
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // API Key 校驗
    if (body.apiKey !== CONFIG.API_KEY) {
      return respond({ ok: false, error: 'INVALID_API_KEY' });
    }

    var action = body.action;
    var data = body.data || {};

    if (action === 'getUsers')      return respond(handleGetUsers());
    if (action === 'upsertUser')    return respond(handleUpsertUser(data));
    if (action === 'addLoginLog')   return respond(handleAddLoginLog(data));
    if (action === 'addActionLog')  return respond(handleAddActionLog(data));

    return respond({ ok: false, error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doGet(e) {
  var key = e.parameter.apiKey || '';
  if (key !== CONFIG.API_KEY) {
    return respond({ ok: false, error: 'INVALID_API_KEY' });
  }

  var action = e.parameter.action || 'getUsers';

  if (action === 'getLocalSchedule') {
    return respond(handleGetLocalSchedule(e.parameter.sheet));
  }

  return respond(handleGetUsers());
}

function handleGetLocalSchedule(sheetName) {
  var name = sheetName || LocalParser.TARGET_SHEET;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    return { ok: false, error: 'SHEET_NOT_FOUND', message: '找不到 Sheet: ' + name };
  }
  var result = LocalParser.parse(sheet);
  if (result.error) {
    return { ok: false, error: result.error, message: result.message };
  }
  // ── 诊断：对比 day10 vs day11 ──
  var d10 = result.daily['10'];
  var d11 = result.daily['11'];
  result._debug_day10_morning = d10 && d10.morning ? d10.morning : 'MISSING';
  result._debug_day11_morning = d11 && d11.morning ? d11.morning : 'MISSING';
  result._debug_parseLog = result._parseLog || [];
  return { ok: true, data: result };
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════
// Users
// ══════════════════════════════════════
var USER_COLUMNS = [
  'id','username','name','role','team','shift',
  'is_root','is_active','is_deleted','deleted_at',
  'must_change_password','password_status',
  'tfa_enabled','tfa_status','screenshot_status',
  'first_login_at','last_login_at',
  'created_at','updated_at'
];

function handleGetUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return { ok: false, error: 'Sheet not found: ' + CONFIG.SHEET_USERS };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, users: [] };

  var headers = data[0];
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    // 排除軟刪除
    if (row.is_deleted === true || row.is_deleted === 'TRUE') continue;
    users.push(row);
  }

  return { ok: true, users: users };
}

function handleUpsertUser(data) {
  if (!data.id || !data.username) return { ok: false, error: 'Missing id or username' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_USERS);
  if (!sheet) return { ok: false, error: 'Sheet not found' };

  // 確保表頭存在
  ensureHeaders(sheet, USER_COLUMNS);

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('id');

  // 尋找現有行
  var existingRow = -1;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][idCol] === data.id) {
      existingRow = i + 1; // Sheet 行號從 1 開始
      break;
    }
  }

  // 組裝行資料
  var rowData = USER_COLUMNS.map(function(col) {
    return data[col] !== undefined ? data[col] : '';
  });

  if (existingRow > 0) {
    // 更新現有行
    sheet.getRange(existingRow, 1, 1, USER_COLUMNS.length).setValues([rowData]);
    return { ok: true, action: 'updated', id: data.id };
  } else {
    // 新增行
    sheet.appendRow(rowData);
    return { ok: true, action: 'created', id: data.id };
  }
}

// ══════════════════════════════════════
// Login Logs
// ══════════════════════════════════════
var LOGIN_LOG_COLUMNS = ['id','username','time','result','reason','ip','user_agent'];

function handleAddLoginLog(data) {
  if (!data.id) return { ok: false, error: 'Missing id' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOGIN_LOGS);
  if (!sheet) return { ok: false, error: 'Sheet not found' };

  ensureHeaders(sheet, LOGIN_LOG_COLUMNS);

  var rowData = LOGIN_LOG_COLUMNS.map(function(col) {
    return data[col] !== undefined ? data[col] : '';
  });

  sheet.appendRow(rowData);
  return { ok: true, action: 'logged', id: data.id };
}

// ══════════════════════════════════════
// Action Logs
// ══════════════════════════════════════
var ACTION_LOG_COLUMNS = ['id','operator','operator_role','action','target','detail','time','result'];

function handleAddActionLog(data) {
  if (!data.id) return { ok: false, error: 'Missing id' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ACTION_LOGS);
  if (!sheet) return { ok: false, error: 'Sheet not found' };

  ensureHeaders(sheet, ACTION_LOG_COLUMNS);

  var rowData = ACTION_LOG_COLUMNS.map(function(col) {
    return data[col] !== undefined ? data[col] : '';
  });

  sheet.appendRow(rowData);
  return { ok: true, action: 'logged', id: data.id };
}

// ══════════════════════════════════════
// 工具函數
// ══════════════════════════════════════
function ensureHeaders(sheet, columns) {
  var firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
  var hasHeaders = false;
  for (var i = 0; i < firstRow.length; i++) {
    if (firstRow[i] !== '') { hasHeaders = true; break; }
  }
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  }
}

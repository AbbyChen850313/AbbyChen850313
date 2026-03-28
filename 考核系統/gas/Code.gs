// ============================================================
// Code.gs — 主路由 & LIFF Web App 入口
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1VKHfnnrv-xfdqj-36I6grY8K-YcuCd8WMIcNAvRA_eg',
  HR_SPREADSHEET_ID: '1hOBSm5BnCjsrp2rX51EN5kYVtEgLZ8FVIMF90_5BMqA',
  LINE_BOT_TOKEN: 'vC9j7A61kp6mlsd450SyLzHMmFB4fzF0piR/5skfHn4dDjGRSZU39pA72441l2gYKx6WSpFt+K63v87uF+KiKuPOe3yvqDeG4b5SQRAsJLm2nbauVyFwtb7b7azpw2Sdpd0xtxcEyFN3/OFrpiU0dAdB04t89/1O/w1cDnyilFU=',
  LIFF_ID: '2009611318-5UphK9JK',
};

// ============================================================
// Sheets 存取 Helper（所有 .gs 檔案共用）
// ============================================================

/** 取得後台 Spreadsheet */
function _ss() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/** 取得指定工作表（找不到回傳 null） */
function _sheet(name) {
  return _ss().getSheetByName(name);
}

/**
 * 取得工作表所有資料列（找不到或空表回傳空陣列）
 * 使用此 helper 可避免 null.getDataRange() 錯誤
 */
function _sheetRows(name) {
  const s = _sheet(name);
  return s ? s.getDataRange().getValues() : [];
}

// ============================================================
// 身份驗證 Helper
// ============================================================

/**
 * 驗證主管身份
 * @returns managerInfo 物件，或 { error: '身份驗證失敗' }
 */
function _verifyManager(lineUid) {
  const info = getManagerInfo(lineUid);
  return info || { error: '身份驗證失敗' };
}

/**
 * 驗證 HR 身份
 * @returns managerInfo 物件，或 { error: '無權限' }
 */
function _verifyHR(lineUid) {
  const info = getManagerInfo(lineUid);
  return (info && info.isHR) ? info : { error: '無權限' };
}

// ============================================================
// Web App 路由
// ============================================================

/**
 * GitHub Pages 前端透過 fetch() 呼叫 GAS API
 * 接收 POST body: { action: 'apiFnName', args: [...] }
 */
function doPost(e) {
  try {
    const { action, args } = JSON.parse(e.postData.contents);
    const API = {
      apiCheckBinding,
      apiBindByIdentity,
      apiGetScoreStatus,
      apiGetMyScores,
      apiGetEmployeesForManager,
      apiSaveDraft,
      apiSubmitScore,
      apiSyncEmployees,
      apiGetSettings,
      apiGetAllStatus,
      apiUpdateSettings,
      apiTriggerReminders,
      apiExportExcel,
    };
    if (!API[action]) return _jsonOut({ error: `Unknown action: ${action}` });
    const result = API[action](...(args || []));
    return _jsonOut(result);
  } catch (err) {
    return _jsonOut({ error: err.message });
  }
}

function _jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  const lineUid = e.parameter.uid || '';

  let template;
  switch (page) {
    case 'score':
      template = HtmlService.createTemplateFromFile('score');
      template.employeeId = e.parameter.eid || '';
      template.lineUid = lineUid;
      template.liffId = CONFIG.LIFF_ID;
      break;
    case 'admin':
      template = HtmlService.createTemplateFromFile('admin');
      template.lineUid = lineUid;
      template.liffId = CONFIG.LIFF_ID;
      break;
    case 'bind':
      template = HtmlService.createTemplateFromFile('bind');
      template.lineUid = '';
      template.liffId = CONFIG.LIFF_ID;
      break;
    default:
      template = HtmlService.createTemplateFromFile('dashboard');
      template.lineUid = lineUid;
      template.liffId = CONFIG.LIFF_ID;
  }

  return template.evaluate()
    .setTitle('考核系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** 計算當前季度（民國年，如 115Q1） */
function getCurrentQuarter() {
  const now = new Date();
  const rocYear = now.getFullYear() - 1911;
  const q = Math.ceil((now.getMonth() + 1) / 3);
  return `${rocYear}Q${q}`;
}

// ============================================================
// API 路由（前端透過 google.script.run 呼叫）
// 新增 API 請遵循以下模式：
//   1. 用 _verifyManager / _verifyHR 做身份驗證
//   2. 若驗證失敗直接 return error 物件
//   3. 業務邏輯委託給對應的 .gs 模組
// ============================================================

// --- Auth ---
// apiBindByIdentity, apiCheckBinding 定義於 Auth.gs

// --- Employees ---
function apiGetEmployeesForManager(lineUid) {
  const info = _verifyManager(lineUid);
  if (info.error) return info;
  return getEmployeesForManager(info);
}

function apiSyncEmployees(lineUid) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;
  return syncEmployees();
}

function apiGetSettings() {
  return getSettings();
}

// --- Scoring ---
function apiSaveDraft(data) {
  return saveDraft(data);
}

function apiSubmitScore(data) {
  return submitScore(data);
}

function apiGetMyScores(lineUid, quarter) {
  const info = getManagerInfo(lineUid);
  if (!info) return { error: '身份驗證失敗' };
  if (!info.isHR && info.responsibilities.length === 0) return { error: '無權限' };
  return getMyScores(lineUid, quarter || getCurrentQuarter());
}

function apiGetScoreStatus(lineUid) {
  const info = _verifyManager(lineUid);
  if (info.error) return info;
  if (info.isHR) return { isHR: true };
  const status = getScoreStatus(info, getCurrentQuarter());
  status.managerName = info.managerName;
  return status;
}

// --- Admin (HR only) ---
function apiGetAllStatus(lineUid) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;
  return getAllManagerStatus(getCurrentQuarter());
}

function apiUpdateSettings(lineUid, newSettings) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;
  return updateSettings(newSettings);
}

function apiTriggerReminders(lineUid) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;
  return sendReminderToAll(getCurrentQuarter());
}

function apiExportExcel(lineUid, quarter) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;
  return exportScores(quarter || getCurrentQuarter());
}

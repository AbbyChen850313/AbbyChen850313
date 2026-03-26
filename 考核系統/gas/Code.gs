// ============================================================
// Code.gs — 主路由 & LIFF Web App 入口
// ============================================================

// ★ 設定區：部署後填入以下值
const CONFIG = {
  SPREADSHEET_ID: '',          // 考核系統後台 Google Sheet ID
  HR_SPREADSHEET_ID: '',       // HR 員工基本資料 Sheet ID
  LINE_BOT_TOKEN: '',          // LINE Messaging API Channel Access Token
  LIFF_ID: '',                 // LINE LIFF App ID
};

/**
 * Web App 入口：根據 page 參數回傳對應 HTML
 */
function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  const lineUid = e.parameter.uid || '';

  let template;
  switch (page) {
    case 'score':
      template = HtmlService.createTemplateFromFile('score');
      template.employeeId = e.parameter.eid || '';
      template.lineUid = lineUid;
      break;
    case 'admin':
      template = HtmlService.createTemplateFromFile('admin');
      template.lineUid = lineUid;
      break;
    case 'bind':
      template = HtmlService.createTemplateFromFile('bind');
      break;
    default:
      template = HtmlService.createTemplateFromFile('dashboard');
      template.lineUid = lineUid;
  }

  return template.evaluate()
    .setTitle('考核系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 讓 HTML 模板引入共用 CSS/JS
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 取得設定（供前端 JS 呼叫）
 */
function getConfig() {
  return {
    liffId: CONFIG.LIFF_ID,
    currentQuarter: getCurrentQuarter(),
  };
}

/**
 * 計算當前季度字串，如 "115Q1"（民國年+Q1~Q4）
 */
function getCurrentQuarter() {
  const settings = getSettings();
  const now = new Date();
  const rocYear = now.getFullYear() - 1911;
  const month = now.getMonth() + 1;
  const q = Math.ceil(month / 3);
  return `${rocYear}Q${q}`;
}

/**
 * API 路由：前端呼叫 google.script.run 時使用的公開函式群
 * 以下為前端可呼叫的函式（GAS 只允許頂層函式）
 */

// --- Auth ---
function apiBindAccount(lineUid, displayName) {
  return bindLineAccount(lineUid, displayName);
}
function apiGetManagerInfo(lineUid) {
  return getManagerInfo(lineUid);
}

// --- Employees ---
function apiGetEmployeesForManager(lineUid) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo) return { error: '身份驗證失敗' };
  return getEmployeesForManager(managerInfo);
}
function apiSyncEmployees() {
  return syncEmployees();
}

// --- Scoring ---
function apiSaveDraft(data) {
  return saveDraft(data);
}
function apiSubmitScore(data) {
  return submitScore(data);
}
function apiGetMyScores(lineUid, quarter) {
  return getMyScores(lineUid, quarter || getCurrentQuarter());
}
function apiGetScoreStatus(lineUid) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo) return { error: '身份驗證失敗' };
  return getScoreStatus(managerInfo, getCurrentQuarter());
}

// --- Admin ---
function apiGetAllStatus(lineUid) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo || !managerInfo.isHR) return { error: '無權限' };
  return getAllManagerStatus(getCurrentQuarter());
}
function apiUpdateSettings(lineUid, newSettings) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo || !managerInfo.isHR) return { error: '無權限' };
  return updateSettings(newSettings);
}
function apiTriggerReminders(lineUid) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo || !managerInfo.isHR) return { error: '無權限' };
  return sendReminderToAll(getCurrentQuarter());
}
function apiExportExcel(lineUid, quarter) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo || !managerInfo.isHR) return { error: '無權限' };
  return exportScores(quarter || getCurrentQuarter());
}

// ============================================================
// Config.gs — 系統設定讀寫
// ============================================================

/**
 * 取得系統設定
 * @returns {Object} settings
 */
function getSettings() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('系統設定');
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key) settings[key] = value;
  }
  return settings;
}

/**
 * 更新系統設定（HR 專用）
 * @param {Object} newSettings - key-value 對應設定名稱與值
 */
function updateSettings(newSettings) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('系統設定');
  const data = sheet.getDataRange().getValues();

  for (const [key, value] of Object.entries(newSettings)) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, value]);
    }
  }
  return { success: true };
}

/**
 * 檢查目前是否在評分期間內
 * @returns {boolean}
 */
function isInScoringPeriod() {
  const settings = getSettings();
  const now = new Date();
  const start = new Date(settings['評分開始日']);
  const end = new Date(settings['評分截止日']);
  return now >= start && now <= end;
}

/**
 * 計算距截止日剩餘天數
 * @returns {number}
 */
function getDaysUntilDeadline() {
  const settings = getSettings();
  const now = new Date();
  const end = new Date(settings['評分截止日']);
  const diff = end - now;
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

/**
 * 取得評分期間描述，如 "115/1~3月"
 * @returns {string}
 */
function getScoringPeriodLabel() {
  const settings = getSettings();
  return settings['評分期間描述'] || '';
}

/**
 * 初始化系統設定工作表（首次使用時呼叫）
 */
function initSettingsSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('系統設定');
  if (!sheet) sheet = ss.insertSheet('系統設定');

  const defaults = [
    ['設定名稱', '設定值', '說明'],
    ['評分期間描述', '115/1~3月', '顯示在介面上的期間文字'],
    ['評分開始日', '2026/04/01', '評分開放日期'],
    ['評分截止日', '2026/04/10', '評分截止日期'],
    ['通知時間點1', '2026/04/05', '第一次提醒日期'],
    ['通知時間點2', '2026/04/09', '第二次提醒日期'],
    ['試用期天數', '90', '未滿幾天算試用期'],
    ['最低評分天數', '3', '到職滿幾天才納入評分'],
    ['當前季度', '115Q1', '當前評分季度'],
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, defaults.length, 3).setValues(defaults);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
}

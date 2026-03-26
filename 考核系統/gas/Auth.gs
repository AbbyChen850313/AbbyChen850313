// ============================================================
// Auth.gs — LINE Login 帳號綁定與身份驗證
// ============================================================

/**
 * 綁定 LINE 帳號（主管第一次登入時呼叫）
 * @param {string} lineUid - LINE User ID
 * @param {string} displayName - LINE 顯示名稱
 * @returns {Object} 綁定結果與主管資訊
 */
function bindLineAccount(lineUid, displayName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('LINE帳號');
  const data = sheet.getDataRange().getValues();

  // 檢查是否已綁定
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === lineUid) {
      return {
        success: true,
        alreadyBound: true,
        managerName: data[i][0],
        lineUid: lineUid,
      };
    }
  }

  // 新增綁定記錄
  const now = new Date();
  sheet.appendRow([displayName, lineUid, displayName, now, '待確認']);

  return {
    success: true,
    alreadyBound: false,
    managerName: displayName,
    lineUid: lineUid,
    message: '帳號已記錄，請等待 HR 確認身份後即可使用。',
  };
}

/**
 * 取得主管資訊
 * @param {string} lineUid
 * @returns {Object|null} 主管資訊，或 null 若未綁定/未授權
 */
function getManagerInfo(lineUid) {
  if (!lineUid) return null;

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const accountSheet = ss.getSheetByName('LINE帳號');
  const accountData = accountSheet.getDataRange().getValues();

  let managerName = null;
  let status = null;
  for (let i = 1; i < accountData.length; i++) {
    if (accountData[i][1] === lineUid) {
      managerName = accountData[i][0];
      status = accountData[i][4]; // 狀態欄：已授權/待確認/停用
      break;
    }
  }

  if (!managerName || status !== '已授權') return null;

  // 查詢主管權重表，找出該主管負責評哪些科別
  const weightSheet = ss.getSheetByName('主管權重');
  const weightData = weightSheet.getDataRange().getValues();

  const responsibilities = [];
  for (let i = 1; i < weightData.length; i++) {
    // 欄位: 被評科別 | 主管名稱 | 主管LINE_UID | 權重
    if (weightData[i][2] === lineUid || weightData[i][1] === managerName) {
      responsibilities.push({
        dept: weightData[i][0],   // 被評科別
        weight: weightData[i][3], // 權重
      });
    }
  }

  // 檢查是否為 HR（在 LINE帳號 表中角色欄為 HR）
  const isHR = accountData.some(row => row[1] === lineUid && row[5] === 'HR');

  return {
    lineUid,
    managerName,
    responsibilities,
    isHR,
  };
}

/**
 * 初始化 LINE帳號 工作表
 */
function initAccountSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('LINE帳號');
  if (!sheet) sheet = ss.insertSheet('LINE帳號');

  const headers = [['主管姓名', 'LINE_UID', 'LINE顯示名稱', '綁定時間', '狀態', '角色']];
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 6).setValues(headers);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
}

/**
 * HR 授權帳號（在 Apps Script 後台手動執行，或透過 Admin 頁面）
 * @param {string} lineUid
 * @param {string} managerName - HR 確認的正確姓名
 * @param {string} role - '主管' 或 'HR'
 */
function authorizeAccount(lineUid, managerName, role) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('LINE帳號');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === lineUid) {
      sheet.getRange(i + 1, 1).setValue(managerName); // 更正姓名
      sheet.getRange(i + 1, 5).setValue('已授權');
      sheet.getRange(i + 1, 6).setValue(role || '主管');
      return { success: true };
    }
  }
  return { error: '找不到該 LINE UID' };
}

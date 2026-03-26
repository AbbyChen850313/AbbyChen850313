// ============================================================
// Auth.gs — LINE Login 帳號綁定與身份驗證
// ============================================================

/**
 * 取得主管資訊
 */
function getManagerInfo(lineUid) {
  if (!lineUid) return null;

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const accountSheet = ss.getSheetByName('LINE帳號');
  const accountData = accountSheet.getDataRange().getValues();

  let managerName = null;
  let status = null;
  let role = null;
  for (let i = 1; i < accountData.length; i++) {
    if (accountData[i][1] === lineUid) {
      managerName = accountData[i][0];
      status = accountData[i][4];
      role = accountData[i][5];
      break;
    }
  }

  if (!managerName || status !== '已授權') return null;

  const weightSheet = ss.getSheetByName('主管權重');
  const weightData = weightSheet.getDataRange().getValues();

  const responsibilities = [];
  for (let i = 1; i < weightData.length; i++) {
    if (weightData[i][2] === lineUid || weightData[i][1] === managerName) {
      responsibilities.push({
        dept: weightData[i][0],
        weight: weightData[i][3],
      });
    }
  }

  const isHR = role === 'HR';

  return { lineUid, managerName, responsibilities, isHR };
}

// ============================================================
// 入會碼系統
// ============================================================

/**
 * HR 生成入會碼（一次產生 N 組）
 * @param {string} lineUid - HR 的 UID（驗證身份）
 * @param {number} count - 要產生幾組，預設 1
 */
function apiGenerateCodes(lineUid, count) {
  const managerInfo = getManagerInfo(lineUid);
  if (!managerInfo || !managerInfo.isHR) return { error: '無權限' };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('入會碼');

  const codes = [];
  for (let i = 0; i < (count || 1); i++) {
    const code = _generateCode();
    sheet.appendRow([code, '未使用', new Date(), '', '']);
    codes.push(code);
  }
  return { success: true, codes };
}

/**
 * 主管用入會碼完成帳號綁定（自助式，立即授權）
 * @param {string} lineUid
 * @param {string} displayName - LINE 顯示名稱
 * @param {string} name - 主管填寫的姓名（需與主管權重表一致）
 * @param {string} role - 'HR' 或 '主管'
 * @param {string} code - 入會碼
 */
function apiRegisterWithCode(lineUid, displayName, name, role, code) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1. 驗證入會碼
  const codeSheet = ss.getSheetByName('入會碼');
  const codeData = codeSheet.getDataRange().getValues();
  let codeRow = -1;
  for (let i = 1; i < codeData.length; i++) {
    if (String(codeData[i][0]).trim().toUpperCase() === code.trim().toUpperCase()
        && codeData[i][1] === '未使用') {
      codeRow = i + 1;
      break;
    }
  }
  if (codeRow < 0) return { error: '入會碼無效或已使用' };

  // 2. 檢查 LINE帳號 是否已有此 UID
  const accountSheet = ss.getSheetByName('LINE帳號');
  const accountData = accountSheet.getDataRange().getValues();
  for (let i = 1; i < accountData.length; i++) {
    if (accountData[i][1] === lineUid) {
      // 已存在：更新資料
      accountSheet.getRange(i + 1, 1).setValue(name);
      accountSheet.getRange(i + 1, 5).setValue('已授權');
      accountSheet.getRange(i + 1, 6).setValue(role);
      _markCodeUsed(codeSheet, codeRow, name);
      _updateWeightUid(ss, name, lineUid);
      return { success: true, message: '帳號已更新並授權' };
    }
  }

  // 3. 新增帳號（直接設為已授權）
  accountSheet.appendRow([name, lineUid, displayName, new Date(), '已授權', role]);

  // 4. 標記入會碼已使用
  _markCodeUsed(codeSheet, codeRow, name);

  // 5. 自動更新主管權重表的 LINE_UID
  _updateWeightUid(ss, name, lineUid);

  return { success: true, message: '帳號綁定成功！' };
}

/** 標記入會碼為已使用 */
function _markCodeUsed(sheet, row, usedBy) {
  sheet.getRange(row, 2).setValue('已使用');
  sheet.getRange(row, 4).setValue(new Date());
  sheet.getRange(row, 5).setValue(usedBy);
}

/** 自動填入主管權重表的 LINE_UID */
function _updateWeightUid(ss, name, lineUid) {
  const weightSheet = ss.getSheetByName('主管權重');
  const weightData = weightSheet.getDataRange().getValues();
  for (let i = 1; i < weightData.length; i++) {
    if (weightData[i][1] === name) {
      weightSheet.getRange(i + 1, 3).setValue(lineUid);
    }
  }
}

/** 產生 6 位英數入會碼 */
function _generateCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) {
    code += chars[Math.floor(Math.random() * chars.length)];
  }
  return code;
}

/**
 * 取得可用的主管姓名清單（從主管權重表抓）
 */
function apiGetManagerNames() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const weightSheet = ss.getSheetByName('主管權重');
  const weightData = weightSheet.getDataRange().getValues();
  const names = new Set();
  for (let i = 1; i < weightData.length; i++) {
    if (weightData[i][1]) names.add(weightData[i][1]);
  }
  return [...names].sort();
}

/**
 * 初始化 LINE帳號 與 入會碼 工作表
 */
function initAccountSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let accountSheet = ss.getSheetByName('LINE帳號');
  if (!accountSheet) accountSheet = ss.insertSheet('LINE帳號');
  accountSheet.clearContents();
  accountSheet.getRange(1, 1, 1, 6).setValues([['主管姓名', 'LINE_UID', 'LINE顯示名稱', '綁定時間', '狀態', '角色']]);
  accountSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');

  let codeSheet = ss.getSheetByName('入會碼');
  if (!codeSheet) codeSheet = ss.insertSheet('入會碼');
  codeSheet.clearContents();
  codeSheet.getRange(1, 1, 1, 5).setValues([['入會碼', '狀態', '建立時間', '使用時間', '使用者']]);
  codeSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
}

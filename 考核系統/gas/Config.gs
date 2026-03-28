// ============================================================
// Config.gs — 系統設定讀寫
// ============================================================

/**
 * 取得所有系統設定（key-value 物件）
 * 若工作表不存在回傳空物件，不會 crash。
 * 若「當前季度」或「評分期間描述」未設定，自動從當前時間推算填入。
 */
function getSettings() {
  const rows = _sheetRows('系統設定');
  const settings = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
  }

  // 自動推算：當前季度（若 Sheet 未填則用當下時間計算）
  if (!settings['當前季度']) {
    settings['當前季度'] = getCurrentQuarter();
  }
  // 自動推算：評分期間描述（若 Sheet 未填則由季度推算）
  if (!settings['評分期間描述']) {
    settings['評分期間描述'] = _quarterToDescription(settings['當前季度']);
  }

  return settings;
}

/**
 * 更新系統設定（HR 專用）
 * 若 key 已存在則更新，否則新增一列
 * @param {Object} newSettings - { 設定名稱: 設定值 }
 */
function updateSettings(newSettings) {
  const sheet = _sheet('系統設定');
  const rows = sheet.getDataRange().getValues();

  for (const [key, value] of Object.entries(newSettings)) {
    const existingRowIndex = rows.findIndex((r, i) => i > 0 && r[0] === key);
    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex + 1, 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  }
  return { success: true };
}

/** 檢查目前是否在評分期間內（未設定日期時預設開放） */
function isInScoringPeriod() {
  const { 評分開始日: start, 評分截止日: end } = getSettings();
  if (!start || !end) return true;
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) return true;
  const now = new Date();
  return now >= startDate && now <= endDate;
}

/** 計算距截止日剩餘天數 */
function getDaysUntilDeadline() {
  const { 評分截止日: end } = getSettings();
  return Math.ceil((new Date(end) - new Date()) / (1000 * 60 * 60 * 24));
}

/**
 * 將季度代碼轉為人類可讀的期間描述
 * @param {string} quarter - 如 "115Q1"
 * @returns {string} 如 "115/1~3月"
 */
function _quarterToDescription(quarter) {
  if (!quarter || quarter.length < 5) return quarter || '';
  const rocYear = quarter.substring(0, 3);
  const q = parseInt(quarter.charAt(4));
  const monthRanges = { 1: '1~3月', 2: '4~6月', 3: '7~9月', 4: '10~12月' };
  return `${rocYear}/${monthRanges[q] || ''}`;
}

/**
 * 取得目前作用中的環境設定（單一入口，所有程式碼從這裡取 token/liffId）
 * 切換環境只需改 系統設定 工作表的「使用測試Channel」即可，不需動程式碼
 * @returns {{ isTest: boolean, botToken: string, liffId: string, label: string }}
 */
function getActiveEnv() {
  const settings = getSettings();
  const isTest = settings['使用測試Channel'] === true || settings['使用測試Channel'] === 'true';
  return {
    isTest,
    botToken: isTest ? CONFIG.LINE_BOT_TOKEN_TEST : CONFIG.LINE_BOT_TOKEN,
    liffId:   isTest ? CONFIG.LIFF_ID_TEST        : CONFIG.LIFF_ID,
    label:    isTest ? '測試Channel' : '正式Channel',
  };
}

/** 啟用測試 Channel（執行一次即可） */
function enableTestChannel() {
  updateSettings({ '使用測試Channel': 'true' });
  Logger.log('✅ 已設定使用測試Channel = true');
}

/**
 * 初始化系統設定工作表（首次使用時呼叫）
 * 含所有預設值，HR 可在工作表手動調整
 */
function initSettingsSheet() {
  const ss = _ss();
  let sheet = ss.getSheetByName('系統設定');
  if (!sheet) sheet = ss.insertSheet('系統設定');

  const now = new Date();
  const quarter = getCurrentQuarter();

  const defaults = [
    ['設定名稱', '設定值', '說明'],
    ['當前季度', quarter, '當前評分季度（留空則自動依當下時間推算）'],
    ['評分期間描述', _quarterToDescription(quarter), '顯示在介面上的期間文字（留空則自動依季度推算）'],
    ['評分開始日', '', '評分開放日期（YYYY/MM/DD）'],
    ['評分截止日', '', '評分截止日期（YYYY/MM/DD）'],
    ['通知時間點1', '', '第一次提醒日期（YYYY/MM/DD）'],
    ['通知時間點2', '', '第二次提醒日期（YYYY/MM/DD）'],
    ['試用期天數', '90', '未滿幾天算試用期（黃底顯示）'],
    ['最低評分天數', '3', '到職滿幾天才納入評分'],
    ['RichMenu_A', '', '公開選單 richMenuId（setupRichMenus() 後自動填入）'],
    ['RichMenu_B', '', '同仁選單 richMenuId（setupRichMenus() 後自動填入）'],
    ['RichMenu_C1', '', '主管選單第一頁 richMenuId（setupRichMenus() 後自動填入）'],
    ['RichMenu_C2', '', '主管選單第二頁 richMenuId（setupRichMenus() 後自動填入）'],
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, defaults.length, 3).setValues(defaults);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 350);
}

/**
 * 建立三張說明文件工作表，讓後續維護者不需要額外說明即可理解系統
 * - 權限設定：各角色能操作的功能
 * - 系統說明：各工作表的欄位結構
 * - 操作手冊：HR 每季操作步驟
 */
function initDocumentationSheets() {
  _initPermissionSheet();
  _initSystemDocSheet();
  _initManualSheet();
  Logger.log('說明文件工作表建立完成');
}

/** 建立「權限設定」工作表 */
function _initPermissionSheet() {
  const ss = _ss();
  let sheet = ss.getSheetByName('權限設定');
  if (!sheet) sheet = ss.insertSheet('權限設定');
  sheet.clearContents();

  const data = [
    ['功能', '一般同仁', '主管（經理/廠長/協理/董事長）', 'HR', '系統管理員', '說明'],
    ['查看自己填寫的評分記錄', '❌', '✅（自己填的那份）', '✅', '✅', '主管查看自己對員工的評分草稿與送出記錄'],
    ['對員工進行評分', '❌', '✅（負責科別）', '❌', '❌', '依主管權重表決定負責科別'],
    ['儲存評分草稿', '❌', '✅', '❌', '❌', '截止前可反覆修改'],
    ['查看所有主管評分進度', '❌', '❌', '✅', '✅', '管理後台 admin 頁'],
    ['手動發送提醒通知', '❌', '❌', '✅', '✅', '對未完成評分的主管推播'],
    ['修改系統設定', '❌', '❌', '✅', '✅', '評分期間、截止日、通知日期等'],
    ['同步員工名單', '❌', '❌', '✅', '✅', '從 HR Sheet 讀取最新員工資料'],
    ['匯出考核結果', '❌', '❌', '✅', '✅', '產生 Google Sheet 格式的結果表'],
    ['查看/管理 LINE 帳號綁定', '❌', '❌', '✅', '✅', '可在 LINE帳號 工作表勾選 I欄刪除帳號'],
    ['重置他人帳號綁定', '❌', '❌', '✅', '✅', 'apiResetAccount'],
    ['切換測試/正式環境', '❌', '❌', '❌', '✅', 'LINE Bot 傳「啟用測試」/「啟用正式」'],
    ['建立 Rich Menu', '❌', '❌', '❌', '✅', 'LINE Bot 傳「建立選單」'],
    ['', '', '', '', '', ''],
    ['角色判定說明', '', '', '', '', ''],
    ['角色依 HR Sheet「(人工打)總表」O欄（職稱類別）決定', '', '', '', '', ''],
    ['董事長、經理、廠長、協理 → 主管（Rich Menu C）', '', '', '', '', ''],
    ['HR → HR 角色（Rich Menu C，進入後自動轉到管理後台）', '', '', '', '', ''],
    ['其他 → 一般同仁（Rich Menu B）', '', '', '', '', ''],
    ['未綁定 / 外部人員 → 公開選單（Rich Menu A）', '', '', '', '', ''],
    ['系統管理員 → 僅可手動在 LINE帳號 H欄設定', '', '', '', '', ''],
  ];

  sheet.getRange(1, 1, data.length, 6).setValues(data);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.getRange(15, 1, 1, 6).setFontWeight('bold').setBackground('#e8f4f8');
  sheet.autoResizeColumns(1, 6);
}

/** 建立「系統說明」工作表 */
function _initSystemDocSheet() {
  const ss = _ss();
  let sheet = ss.getSheetByName('系統說明');
  if (!sheet) sheet = ss.insertSheet('系統說明');
  sheet.clearContents();

  const data = [
    ['工作表', '欄位', '說明'],
    ['LINE帳號', 'A 姓名', '員工姓名（從 HR Sheet 取得）'],
    ['LINE帳號', 'B LINE_UID', 'LINE 使用者唯一識別碼（系統自動取得）'],
    ['LINE帳號', 'C LINE顯示名稱', '使用者在 LINE 的暱稱'],
    ['LINE帳號', 'D 綁定時間', '完成綁定的日期時間'],
    ['LINE帳號', 'E 狀態', '已授權 = 可使用系統'],
    ['LINE帳號', 'F 職稱', '從 HR Sheet M欄（職稱）取得'],
    ['LINE帳號', 'G 電話', '使用者綁定時填寫的手機號碼（必填）'],
    ['LINE帳號', 'H 角色', '系統管理員／HR／主管／同仁（可手動修改）'],
    ['LINE帳號', 'I 清除帳號', '勾選後執行「clearCheckedAccounts()」即可刪除'],
    ['', '', ''],
    ['主管權重', 'A 被評科別', '接受考核的科別（如：品管科、財務科）'],
    ['主管權重', 'B 職稱', '負責評分的主管職稱（用職稱而非姓名，人員異動時不需修改）'],
    ['主管權重', 'C 姓名', '目前擔任該職位者的姓名（方便人工核對，系統自動填入）'],
    ['主管權重', 'D LINE_UID', '主管綁定後系統自動填入（請勿手動修改）'],
    ['主管權重', 'E 權重', '評分佔比，同一科別的所有主管權重加總須 = 1.0'],
    ['', '', ''],
    ['員工資料', 'A 員工編號', 'HR 告知員工用於 LINE 綁定身分核對'],
    ['員工資料', 'B 姓名', ''],
    ['員工資料', 'C 部門', ''],
    ['員工資料', 'D 科別', ''],
    ['員工資料', 'E 到職日', ''],
    ['員工資料', 'F 離職日', '空白 = 在職中'],
    ['', '', ''],
    ['評分記錄', 'A 季度', '如 115Q1'],
    ['評分記錄', 'B 評分主管（姓名）', ''],
    ['評分記錄', 'C 被評人員', ''],
    ['評分記錄', 'D 被評科別', ''],
    ['評分記錄', 'E 主管權重', ''],
    ['評分記錄', 'F~K 六項評分', '職能專業度、工作效率、成本意識、部門合作、責任感、主動積極'],
    ['評分記錄', 'L 原始平均分', '六項平均'],
    ['評分記錄', 'M 特殊加減分', '主管手動調整分數'],
    ['評分記錄', 'N 調整後分數', 'L + M'],
    ['評分記錄', 'O 加權分數', 'N × E（權重）'],
    ['評分記錄', 'P 備註', ''],
    ['評分記錄', 'Q 狀態', '草稿 / 已送出'],
    ['評分記錄', 'R 最後更新', '每次存草稿或送出都更新'],
    ['', '', ''],
    ['系統設定', '當前季度', '留空則自動依當下時間推算（如 115Q1 = 民國115年第一季）'],
    ['系統設定', '評分期間描述', '留空則自動由季度推算（如 115Q1 → 115/1~3月）'],
    ['系統設定', 'RichMenu_A/B/C1/C2', '執行 setupRichMenus() 後自動填入，請勿手動修改'],
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  // 各工作表段落標題加底色
  [2, 11, 17, 23, 36].forEach(r => {
    if (data[r - 1] && data[r - 1][0]) {
      sheet.getRange(r, 1, 1, 3).setBackground('#e8f4f8').setFontWeight('bold');
    }
  });
  sheet.autoResizeColumns(1, 3);
}

/** 建立「操作手冊」工作表 */
function _initManualSheet() {
  const ss = _ss();
  let sheet = ss.getSheetByName('操作手冊');
  if (!sheet) sheet = ss.insertSheet('操作手冊');
  sheet.clearContents();

  const data = [
    ['每季 HR 操作步驟', '', ''],
    ['步驟', '操作', '說明'],
    ['1', '更新員工資料', '在 admin 後台「系統設定」頁按「從 HR Sheet 同步員工名單」'],
    ['2', '設定評分期間', '在「系統設定」填入評分開始日、截止日、兩個通知時間點'],
    ['3', '確認主管權重表', '開啟「主管權重」工作表，確認各科別的主管職稱與權重正確'],
    ['4', '通知主管綁定帳號', '主管打開 LINE Bot，輸入姓名 + 員工編號完成帳號綁定'],
    ['5', '評分期間開始', '主管打開考核系統 LIFF 進行評分'],
    ['6', '發送提醒通知', '在 admin 後台「評分進度」頁按「發送提醒通知」'],
    ['7', '評分截止後匯出', '在 admin 後台「匯出」頁選擇季度後按「匯出」'],
    ['', '', ''],
    ['首次部署步驟', '', ''],
    ['步驟', '操作', '說明'],
    ['D1', '建立所有工作表', '在 GAS 執行「initAllSheets()」'],
    ['D2', '建立說明文件', '在 GAS 執行「initDocumentationSheets()」'],
    ['D3', '設定 LINE Rich Menu', '在 RichMenu.gs 填入圖片 Drive ID 和按鈕 URL，執行「setupRichMenus()」'],
    ['D4', '設定定時提醒觸發器', '在 GAS 執行「setupTriggers()」（只需執行一次）'],
    ['', '', ''],
    ['常用 GAS 函式', '', ''],
    ['函式', '說明', ''],
    ['initAllSheets()', '一鍵建立所有工作表（首次部署）', ''],
    ['initDocumentationSheets()', '建立/更新三張說明文件工作表', ''],
    ['syncEmployees()', '從 HR Sheet 同步員工名單到「員工資料」', ''],
    ['setupRichMenus()', '建立 LINE Rich Menu（換圖時重新執行）', ''],
    ['setupTriggers()', '設定每日提醒排程觸發器', ''],
    ['clearCheckedAccounts()', '刪除 LINE帳號 I欄打勾的帳號', ''],
    ['setupAccountCheckboxes()', '補齊 LINE帳號 G欄的勾選框（修復用）', ''],
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  sheet.getRange(2, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.getRange(11, 1, 1, 3).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  sheet.getRange(12, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.getRange(18, 1, 1, 3).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  sheet.getRange(19, 1, 1, 3).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, 3);
}

// ============================================================
// Config.gs - Configuration read/write via Google Sheets
// ============================================================

var CONFIG_SHEET_NAME = '系統設定';
var MAIN_SPREADSHEET_ID_KEY = 'MAIN_SS_ID'; // stored in Script Properties

/**
 * Get the main spreadsheet. First run: creates it and saves ID.
 */
function getMainSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty(MAIN_SPREADSHEET_ID_KEY);

  if (!ssId) {
    // First run: create the spreadsheet
    var ss = SpreadsheetApp.create('員工考核系統資料庫');
    ssId = ss.getId();
    props.setProperty(MAIN_SPREADSHEET_ID_KEY, ssId);
    initializeSpreadsheet(ss);
    return ss;
  }
  return SpreadsheetApp.openById(ssId);
}

/**
 * Initialize spreadsheet with required sheets on first run.
 */
function initializeSpreadsheet(ss) {
  // Rename default Sheet1 to 系統設定
  var sheet1 = ss.getSheets()[0];
  sheet1.setName(CONFIG_SHEET_NAME);

  // Create other sheets
  var sheetsNeeded = ['考核分數', '員工名單快取', '考核期間'];
  sheetsNeeded.forEach(function(name) {
    ss.insertSheet(name);
  });

  // Write default config
  var defaultConfig = getDefaultConfig();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  configSheet.getRange('A1').setValue('CONFIG_JSON');
  configSheet.getRange('B1').setValue(JSON.stringify(defaultConfig));

  // Set up 考核分數 headers
  var scoresSheet = ss.getSheetByName('考核分數');
  scoresSheet.getRange('A1:I1').setValues([[
    '考核期間ID', '被評員工ID', '被評員工姓名', '被評部門',
    '評分主管角色', '評分主管Email', '評分項目', '分數', '提交時間'
  ]]);

  // Set up 員工名單快取 headers
  var empSheet = ss.getSheetByName('員工名單快取');
  empSheet.getRange('A1:H1').setValues([[
    '員工ID', '姓名', '部門', '到職日', '離職日', '在職狀態', '年資', '更新時間'
  ]]);

  Logger.log('Spreadsheet initialized: ' + ss.getUrl());
}

/**
 * Get default configuration.
 */
function getDefaultConfig() {
  return {
    adminEmails: [],
    employeeSheetId: '1hOBSm5BnCjsrp2rX51EN5kYVtEgLZ8FVIMF90_5BMqA',
    supervisors: [
      { role: '營運部協理',      email: '' },
      { role: '儲運科經理',      email: '' },
      { role: '生產科廠長',      email: '' },
      { role: '廠務部協理',      email: '' },
      { role: '永續發展科經理',  email: '' },
      { role: '業務人員A',       email: '' },
      { role: '業務人員B',       email: '' },
      { role: '業務人員C',       email: '' }
    ],
    deptReviewConfig: [
      {
        dept: '品管科',
        reviewers: [
          { role: '營運部協理',   weight: 70 },
          { role: '儲運科經理',   weight: 15 },
          { role: '生產科廠長',   weight: 15 }
        ]
      },
      {
        dept: '業務科',
        reviewers: [
          { role: '營運部協理',   weight: 70 },
          { role: '廠務部協理',   weight: 15 },
          { role: '永續發展科經理', weight: 15 }
        ]
      },
      {
        dept: '儲運科',
        reviewers: [
          { role: '儲運科經理',   weight: 70 },
          { role: '廠務部協理',   weight: 15 },
          { role: '營運部協理',   weight: 15 }
        ]
      },
      {
        dept: '儲運科經理',
        reviewers: [
          { role: '廠務部協理',   weight: 70 },
          { role: '營運部協理',   weight: 15 },
          { role: '永續發展科經理', weight: 15 }
        ]
      },
      {
        dept: '生產科',
        reviewers: [
          { role: '生產科廠長',   weight: 70 },
          { role: '廠務部協理',   weight: 15 },
          { role: '營運部協理',   weight: 15 }
        ]
      },
      {
        dept: '生產科廠長',
        reviewers: [
          { role: '廠務部協理',   weight: 70 },
          { role: '營運部協理',   weight: 15 },
          { role: '永續發展科經理', weight: 15 }
        ]
      },
      {
        dept: '財務科',
        reviewers: [
          { role: '永續發展科經理', weight: 70 },
          { role: '業務人員A',    weight: 10 },
          { role: '業務人員B',    weight: 10 },
          { role: '業務人員C',    weight: 10 }
        ]
      },
      {
        dept: '永續發展科',
        reviewers: [
          { role: '永續發展科經理', weight: 70 },
          { role: '營運部協理',   weight: 30 }
        ]
      }
    ],
    scoringCategories: [
      {
        dept: 'ALL',
        items: [
          { name: '工作品質與績效', weight: 30, description: '工作成果品質、達成目標程度' },
          { name: '工作態度與責任心', weight: 25, description: '積極度、責任感、出勤狀況' },
          { name: '溝通協調能力', weight: 20, description: '與同仁、上司溝通協作' },
          { name: '學習與成長', weight: 15, description: '接受新知、自我提升' },
          { name: '遵守規章制度', weight: 10, description: '遵守公司規定、保密義務' }
        ]
      }
    ],
    reviewPeriods: [
      {
        id: '115Q1',
        periodName: '115年第一季 (1-3月)',
        startDate: '2026-01-01',
        endDate: '2026-03-31',
        deadlineDate: '2026-04-10',
        notificationDates: ['2026/04/01', '2026/04/10'],
        notificationTime: '09:00',
        isOpen: true
      }
    ],
    notificationSettings: {
      senderName: '人資部門',
      reminderSubject: '【提醒】員工考核評分即將截止',
      reminderTemplate: '親愛的 {supervisorRole}，\n\n本季員工考核評分截止日為 {deadline}，您尚有 {pendingCount} 位同仁未完成評分。\n\n請盡速至考核系統完成評分：{url}\n\n人資部門 敬啟'
    }
  };
}

/**
 * Read config from spreadsheet.
 */
function getConfig() {
  try {
    var ss = getMainSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!sheet) {
      initializeSpreadsheet(ss);
      sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    }
    var jsonStr = sheet.getRange('B1').getValue();
    if (!jsonStr) return getDefaultConfig();
    return JSON.parse(jsonStr);
  } catch (e) {
    Logger.log('Config read error: ' + e);
    return getDefaultConfig();
  }
}

/**
 * Save config to spreadsheet.
 */
function saveConfig(configJson) {
  var ss = getMainSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var jsonStr = typeof configJson === 'string' ? configJson : JSON.stringify(configJson);
  sheet.getRange('A1').setValue('CONFIG_JSON');
  sheet.getRange('B1').setValue(jsonStr);

  // Re-install triggers when config is saved
  try { installTriggers(); } catch(e) {}
}

/**
 * Get scoring items for a specific department.
 * Falls back to ALL dept config if dept-specific not found.
 */
function getScoringItemsByDept(config, dept) {
  var categories = config.scoringCategories || [];
  // Try dept-specific first
  for (var i = 0; i < categories.length; i++) {
    if (categories[i].dept === dept) return categories[i].items;
  }
  // Fall back to ALL
  for (var j = 0; j < categories.length; j++) {
    if (categories[j].dept === 'ALL') return categories[j].items;
  }
  return [];
}

/**
 * Return the spreadsheet URL for admin display.
 */
function getSpreadsheetUrl() {
  requireAdmin();
  return getMainSpreadsheet().getUrl();
}

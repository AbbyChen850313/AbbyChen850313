// ============================================================
// Code.gs — 主路由 & LIFF Web App 入口
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1VKHfnnrv-xfdqj-36I6grY8K-YcuCd8WMIcNAvRA_eg',
  HR_SPREADSHEET_ID: '1hOBSm5BnCjsrp2rX51EN5kYVtEgLZ8FVIMF90_5BMqA',
  // 正式 Messaging Channel（2009611318）的 Bot Token
  LINE_BOT_TOKEN: 'vC9j7A61kp6mlsd450SyLzHMmFB4fzF0piR/5skfHn4dDjGRSZU39pA72441l2gYKx6WSpFt+K63v87uF+KiKuPOe3yvqDeG4b5SQRAsJLm2nbauVyFwtb7b7azpw2Sdpd0xtxcEyFN3/OFrpiU0dAdB04t89/1O/w1cDnyilFU=',
  // 測試 Messaging Channel（2008337190）的 Bot Token
  LINE_BOT_TOKEN_TEST: '3nqiobdCVPhomyttwLtvGdaW37UE/hUXI9jICkGWJv5Vo2EMbzAGR61pVu5nj9/O2yjRVzC8+1amRpgPxtv431/mYTzZh20qQ/Z4M1nKSekcp1GNgPanrKgmq+ocQT6DTi9E9wot4P13uFr1R4bESQdB04t89/1O/w1cDnyilFU=',
  LIFF_ID:      '2009611318-5UphK9JK',  // 正式
  LIFF_ID_TEST: '2009619528-aJO34c6u',  // 測試
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

// ============================================================
// 系統日誌
// ============================================================

/**
 * 寫入一筆日誌到「系統日誌」工作表
 * @param {'INFO'|'WARN'|'ERROR'} level
 * @param {string} fn  - 函式名稱（方便定位）
 * @param {string} msg - 簡短說明
 * @param {*} [detail] - 附加資料（物件/字串都行）
 */
function _log(level, fn, msg, detail) {
  try {
    const ss = _ss();
    let sheet = ss.getSheetByName('系統日誌');
    if (!sheet) {
      sheet = ss.insertSheet('系統日誌');
      sheet.getRange(1, 1, 1, 5).setValues([['時間', '等級', '函式', '說明', '詳細資料']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(5, 400);
    }
    const detailStr = detail !== undefined
      ? (typeof detail === 'object' ? JSON.stringify(detail) : String(detail))
      : '';
    sheet.appendRow([new Date(), level, fn, msg, detailStr]);

    // 只保留最近 500 筆，避免表格過大
    const lastRow = sheet.getLastRow();
    if (lastRow > 501) {
      sheet.deleteRows(2, lastRow - 501);
    }
  } catch (_) {
    // log 本身不能崩潰
    Logger.log(`[_log failed] ${level} ${fn} ${msg}`);
  }
}

/**
 * GitHub Pages 前端透過 fetch() 呼叫 GAS API
 * 接收 POST body: { action: 'apiFnName', args: [...] }
 */
function doPost(e) {
  let action = '(unknown)';
  try {
    const body = JSON.parse(e.postData.contents);

    // LINE Webhook 事件（Bot 收到訊息）
    if (body.events) {
      _handleLineWebhook(body.events);
      return _jsonOut({ ok: true });
    }

    action = body.action;
    const args = body.args;
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
      apiGetAllAccounts,
      apiResetAccount,
      apiGetLogs,
      apiGetManagerDashboard,
    };
    if (!API[action]) {
      _log('WARN', 'doPost', `未知 action: ${action}`);
      return _jsonOut({ error: `Unknown action: ${action}` });
    }
    const result = API[action](...(args || []));
    // 業務邏輯回傳 error 時也記錄
    if (result && result.error) {
      _log('WARN', action, result.error, { args: _sanitizeArgs(action, args) });
    }
    return _jsonOut(result);
  } catch (err) {
    _log('ERROR', action, err.message, { stack: err.stack });
    return _jsonOut({ error: err.message });
  }
}

/** 遮蔽 args 裡的 lineUid（避免日誌洩漏身份） */
function _sanitizeArgs(action, args) {
  if (!args) return [];
  // 第一個參數通常是 lineUid，只保留最後 4 碼
  return args.map((a, i) => {
    if (i === 0 && typeof a === 'string' && a.length > 8) {
      return '…' + a.slice(-4);
    }
    return a;
  });
}

function _jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // LIFF bind page 的 REST API（GitHub Pages 呼叫用）
  if (e.parameter.action) {
    return _handleLiffBindAction(e.parameter);
  }

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

/**
 * LIFF 綁定頁面（GitHub Pages 靜態頁）專用 GET API
 * 支援 action: checkBinding / bindByIdentity / unbindSelf
 */
function _handleLiffBindAction(params) {
  try {
    const action = params.action;
    const uid    = params.uid || '';
    let result;

    if (action === 'checkBinding') {
      result = apiCheckBinding(uid);

    } else if (action === 'bindByIdentity') {
      result = apiBindByIdentity(
        uid,
        params.displayName || '',
        params.name        || '',
        params.eid         || '',
        params.phone       || ''
      );

    } else if (action === 'unbindSelf') {
      // 自行解除綁定（測試/重綁用）
      if (!uid) { result = { error: 'missing uid' }; }
      else {
        const accountSheet = _sheet('LINE帳號');
        if (accountSheet) {
          const data = accountSheet.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) {
            if (data[i][1] === uid) accountSheet.deleteRow(i + 1);
          }
        }
        const richMenuA = getSettings()['RichMenu_A'];
        if (richMenuA) _linkRichMenuToUser(uid, richMenuA);
        result = { success: true };
      }

    } else {
      result = { error: 'unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

/** HR 以指定主管 UID 查看其儀表板（員工列表＋評分狀態） */
function apiGetManagerDashboard(hrLineUid, targetManagerUid) {
  const info = _verifyHR(hrLineUid);
  if (info.error) return info;
  const managerInfo = getManagerInfo(targetManagerUid);
  if (!managerInfo) return { error: '查無此主管帳號' };
  const status = getScoreStatus(managerInfo, getCurrentQuarter());
  const scores = getMyScores(targetManagerUid, getCurrentQuarter());
  status.managerName = managerInfo.managerName;
  status.employees = status.employees.map(emp => ({
    ...emp,
    scoreStatus: (scores[emp.name] && scores[emp.name].status) || emp.scoreStatus,
  }));
  return status;
}

/** 取得最近 100 筆日誌（HR 專用） */
function apiGetLogs(lineUid) {
  const info = _verifyHR(lineUid);
  if (info.error) return info;

  const sheet = _sheet('系統日誌');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  // 回傳最新 100 筆（倒序，最新在前）
  return data.slice(1).reverse().slice(0, 100).map(r => ({
    time:   r[0] ? new Date(r[0]).toLocaleString('zh-TW') : '',
    level:  r[1] || '',
    fn:     r[2] || '',
    msg:    r[3] || '',
    detail: r[4] || '',
  }));
}

// ============================================================
// LINE Webhook 處理
// ============================================================

function _handleLineWebhook(events) {
  events.forEach(event => {
    if (event.type !== 'message' || event.message.type !== 'text') return;
    const text = event.message.text.trim();
    const replyToken = event.replyToken;
    const uid = event.source.userId;
    const settings = getSettings();

    if (text === 'ping') {
      const isTest = settings['使用測試Channel'] === true || settings['使用測試Channel'] === 'true';
      const reply = [
        '🤖 系統回應 OK',
        `環境：${isTest ? '✅ 測試Channel' : '⚠️ 正式Channel'}`,
        `季度：${settings['當前季度'] || '未設定'}`,
        `評分期間：${settings['評分期間描述'] || '未設定'}`,
      ].join('\n');
      _lineReply(replyToken, reply);

    } else if (text === '設定' || text === '綁定設定') {
      _lineReply(replyToken, `請點以下連結進行帳號綁定：\nhttps://liff.line.me/${CONFIG.LIFF_ID_TEST}`);

    } else if (text === '主管') {
      const richMenuId = settings['RichMenu_C1'];
      if (richMenuId) { _linkRichMenuToUser(uid, richMenuId); _lineReply(replyToken, '已切換到主管選單 (C1)'); }
      else { _lineReply(replyToken, '尚未設定 Rich Menu，請先執行 setupRichMenus()'); }

    } else if (text === '同仁') {
      const richMenuId = settings['RichMenu_B'];
      if (richMenuId) { _linkRichMenuToUser(uid, richMenuId); _lineReply(replyToken, '已切換到同仁選單 (B)'); }
      else { _lineReply(replyToken, '尚未設定 Rich Menu，請先執行 setupRichMenus()'); }

    } else if (text === '重置') {
      const richMenuId = settings['RichMenu_A'];
      if (richMenuId) { _linkRichMenuToUser(uid, richMenuId); _lineReply(replyToken, '已重置為雜人選單 (A)'); }
      else { _lineReply(replyToken, '尚未設定 Rich Menu，請先執行 setupRichMenus()'); }

    } else if (text === '啟用測試') {
      updateSettings({ '使用測試Channel': 'true' });
      _lineReply(replyToken, '✅ 已設定使用測試Channel = true\n再傳 ping 確認');
    }
  });
}

function _lineReply(replyToken, text) {
  // Webhook 回覆固定用測試 channel token（webhook 本來就只接測試 channel 的訊息）
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${CONFIG.LINE_BOT_TOKEN_TEST}` },
    payload: JSON.stringify({
      replyToken,
      messages: [{ type: 'text', text }],
    }),
    muteHttpExceptions: true,
  });
}

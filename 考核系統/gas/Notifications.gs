// ============================================================
// Notifications.gs — LINE Bot 通知
// ============================================================

/**
 * 傳送 LINE Push Message 給單一使用者
 * @param {string} lineUid
 * @param {string} message
 */
function sendReminder(lineUid, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = JSON.stringify({
    to: lineUid,
    messages: [{
      type: 'text',
      text: message,
    }],
  });

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${CONFIG.LINE_BOT_TOKEN}`,
    },
    payload: payload,
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code !== 200) {
    Logger.log(`LINE Push 失敗 (${lineUid}): ${response.getContentText()}`);
    return false;
  }
  return true;
}

/**
 * 對所有尚未完成評分的主管發送提醒
 * @param {string} quarter
 */
function sendReminderToAll(quarter) {
  const settings = getSettings();
  const deadline = settings['評分截止日'];
  const period = settings['評分期間描述'];

  const allStatus = getAllManagerStatus(quarter);
  let sent = 0;

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const accountSheet = ss.getSheetByName('LINE帳號');
  const accountData = accountSheet.getDataRange().getValues();

  // 建立 名稱→UID 對照
  const uidMap = {};
  for (let i = 1; i < accountData.length; i++) {
    uidMap[accountData[i][0]] = accountData[i][1];
  }

  for (const status of allStatus) {
    if (status.pending <= 0) continue;

    const uid = uidMap[status.managerName];
    if (!uid) continue;

    const message =
      `📋 考核評分提醒\n\n` +
      `${period} 考核評分尚未完成\n` +
      `・已評分：${status.scored}人\n` +
      `・待評分：${status.pending}人\n` +
      `截止日：${deadline}\n\n` +
      `請盡快完成評分，謝謝！`;

    sendReminder(uid, message);
    sent++;
    Utilities.sleep(200); // 避免 LINE API 速率限制
  }

  return { success: true, notifiedCount: sent };
}

/**
 * 設定定時觸發器（在 GAS 部署後執行一次）
 */
function setupTriggers() {
  // 清除舊觸發器
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledReminder') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 每天早上9點執行排程檢查
  ScriptApp.newTrigger('scheduledReminder')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

/**
 * 定時排程：檢查是否為通知時間點，若是則發送
 */
function scheduledReminder() {
  const settings = getSettings();
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  const notify1 = settings['通知時間點1'] ?
    Utilities.formatDate(new Date(settings['通知時間點1']), 'Asia/Taipei', 'yyyy/MM/dd') : '';
  const notify2 = settings['通知時間點2'] ?
    Utilities.formatDate(new Date(settings['通知時間點2']), 'Asia/Taipei', 'yyyy/MM/dd') : '';

  if (today === notify1 || today === notify2) {
    const quarter = settings['當前季度'] || getCurrentQuarter();
    sendReminderToAll(quarter);
    Logger.log(`[${today}] 已發送提醒通知`);
  }
}

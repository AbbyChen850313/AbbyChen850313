// ============================================================
// RichMenu.gs — LINE Rich Menu 角色分流設定
// ============================================================
//
// 使用方式：
//   Step 1. 將 4 張圖片上傳到 Google Drive
//   Step 2. 把各圖的 Drive 檔案 ID 填入下方 DRIVE_FILE_IDS
//   Step 3. 確認下方 ACTION_URLS 的連結正確
//   Step 4. 在 GAS 編輯器執行 setupRichMenus()（只需執行一次）
//
// ============================================================

// ── 圖片來源（Google Drive 檔案 ID）────────────────────────
// 上傳圖片到 Drive 後，右鍵 → 取得連結 → 複製 ID 填入
const RICH_MENU_IMAGES = {
  A:   '1w5zK-TqxFUyN_UxGIRkVMN9MoBZJOpA4',  // 雜人/未綁定選單
  B:   '1yB-aWx8781knytIrFXHYsmU_jAJ862zq',  // 同仁選單
  C1:  '1prP0pLQeIbbpLlpz1m0izeu4wl_M5oxu',  // 主管第一頁（頂部300px Tab亮）
  C2:  '1BpXaXmOxNSj-5LJX0LUh7YtT5_48y_Mn',  // 主管第二頁（頂部300px Tab亮）
};

// ── 各按鈕連結（由 AI 團隊填入實際 URL）────────────────────
const ACTION_URLS = {
  官網:         'https://www.liangchun.com.tw/article.php?lang=tw&tb=5',
  綁定帳號:     `https://liff.line.me/${CONFIG.LIFF_ID}`,
  我要請款:     'https://TODO_請款URL',      // ← 請填入實際網址
  查詢請款:     'https://TODO_查詢請款URL',  // ← 請填入實際網址
  重要表單QA:   'https://TODO_表單URL',      // ← 請填入實際網址
  公司活動報名: 'https://TODO_活動URL',      // ← 請填入實際網址
  讚賞幣:       'https://TODO_讚賞幣URL',    // ← 請填入實際網址
  出勤:         'https://TODO_出勤URL',      // ← 請填入實際網址
  考核系統:     `https://liff.line.me/${CONFIG.LIFF_ID}`,
};

// ── Rich Menu Alias 名稱（Tab 切換用，不需修改）────────────
const ALIAS_MANAGER_P1 = 'alias-manager-p1';
const ALIAS_MANAGER_P2 = 'alias-manager-p2';

// ============================================================
// 一次性設定函式（執行一次即可）
// ============================================================

/**
 * 建立所有 Rich Menu 並設定預設值
 * 執行完成後，在 GAS 執行記錄查看各 richMenuId
 */
function setupRichMenus() {
  Logger.log('=== 開始建立 Rich Menu ===');

  // 1. 建立 4 個 Rich Menu，取得 ID
  const idA  = _createRichMenu(_buildMenuA());
  const idB  = _createRichMenu(_buildMenuB());
  const idC1 = _createRichMenu(_buildMenuC1());
  const idC2 = _createRichMenu(_buildMenuC2());

  Logger.log(`A  (雜人)     richMenuId: ${idA}`);
  Logger.log(`B  (同仁)     richMenuId: ${idB}`);
  Logger.log(`C1 (主管Tab1) richMenuId: ${idC1}`);
  Logger.log(`C2 (主管Tab2) richMenuId: ${idC2}`);

  // 2. 上傳圖片到各 Rich Menu
  _uploadRichMenuImage(idA,  RICH_MENU_IMAGES.A);
  _uploadRichMenuImage(idB,  RICH_MENU_IMAGES.B);
  _uploadRichMenuImage(idC1, RICH_MENU_IMAGES.C1);
  _uploadRichMenuImage(idC2, RICH_MENU_IMAGES.C2);
  Logger.log('圖片上傳完成');

  // 3. 建立 Alias（Tab 切換要用）
  _createOrUpdateAlias(ALIAS_MANAGER_P1, idC1);
  _createOrUpdateAlias(ALIAS_MANAGER_P2, idC2);
  Logger.log('Alias 建立完成');

  // 4. 設定 A 為全域預設（所有人預設看到雜人選單）
  _setDefaultRichMenu(idA);
  Logger.log('預設選單設定完成（A）');

  // 5. 將 richMenuId 存到系統設定，方便後續查詢
  updateSettings({
    'RichMenu_A':  idA,
    'RichMenu_B':  idB,
    'RichMenu_C1': idC1,
    'RichMenu_C2': idC2,
  });

  Logger.log('=== Rich Menu 設定完成 ===');
}

// ============================================================
// 綁定後切換 Rich Menu（Auth.gs 綁定成功後呼叫）
// ============================================================

/**
 * 依職稱類別為使用者切換 Rich Menu
 * @param {string} lineUid
 * @param {string} titleCategory - HR Sheet O欄的職稱類別值（如 '經理', '協理', '董事長', 'HR'）
 */
function switchRichMenuByRole(lineUid, titleCategory) {
  const settings = getSettings();
  // 主管（含 HR）→ C1（有考核系統 tab）；一般同仁 → B
  const isManagerOrHR = MANAGER_TITLE_CATEGORIES.includes(titleCategory) || titleCategory === 'HR';
  const richMenuId = isManagerOrHR ? settings['RichMenu_C1'] : settings['RichMenu_B'];

  if (!richMenuId) {
    Logger.log('switchRichMenuByRole: 找不到 RichMenu ID，請先執行 setupRichMenus()');
    return;
  }

  _linkRichMenuToUser(lineUid, richMenuId);
}

// ============================================================
// Rich Menu JSON 定義
// ============================================================

/** A — 雜人/未綁定（2格，全高，無 Tab） */
function _buildMenuA() {
  return {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: 'menu_a_public',
    chatBarText: '選單',
    areas: [
      _area(0, 0, 1250, 1686, { type: 'uri', uri: ACTION_URLS.官網 }),
      _area(1250, 0, 1250, 1686, { type: 'uri', uri: ACTION_URLS.綁定帳號 }),
    ],
  };
}

/** B — 一般同仁（6格，2列×3欄，無 Tab） */
function _buildMenuB() {
  return {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: 'menu_b_employee',
    chatBarText: '選單',
    areas: _sixCellAreas(0),
  };
}

/** C-1 — 主管第一頁（Tab bar + 6格，Tab1 選中） */
function _buildMenuC1() {
  return {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: 'menu_c1_manager_p1',
    chatBarText: '選單',
    areas: [
      // Tab bar：左側 Tab1（已選中，無動作）、右側 Tab2（切換到 C-2）
      _area(0,    0, 1250, 300, { type: 'postback', data: 'tab=1' }),
      _area(1250, 0, 1250, 300, { type: 'richmenuswitch', richMenuAliasId: ALIAS_MANAGER_P2, data: 'tab=2' }),
      // 六宮格內容（從 y=300 開始）
      ..._sixCellAreas(300),
    ],
  };
}

/** C-2 — 主管第二頁（Tab bar + 1格考核系統，Tab2 選中） */
function _buildMenuC2() {
  return {
    size: { width: 2500, height: 1686 },
    selected: true,
    name: 'menu_c2_manager_p2',
    chatBarText: '選單',
    areas: [
      // Tab bar：左側 Tab1（切換到 C-1）、右側 Tab2（已選中，無動作）
      _area(0,    0, 1250, 300, { type: 'richmenuswitch', richMenuAliasId: ALIAS_MANAGER_P1, data: 'tab=1' }),
      _area(1250, 0, 1250, 300, { type: 'postback', data: 'tab=2' }),
      // 整塊大按鈕：考核系統
      _area(0, 300, 2500, 1386, { type: 'uri', uri: ACTION_URLS.考核系統 }),
    ],
  };
}

// ============================================================
// Helper：Rich Menu 座標與六宮格
// ============================================================

/**
 * 建立 area 物件
 * @param {number} x
 * @param {number} y
 * @param {number} w - 寬度
 * @param {number} h - 高度
 * @param {Object} action - LINE action 物件
 */
function _area(x, y, w, h, action) {
  return {
    bounds: { x, y, width: w, height: h },
    action,
  };
}

/**
 * 建立標準六宮格（2列 × 3欄）的 areas 陣列
 * @param {number} startY - 內容區起始 Y 座標（有 Tab 時傳 300，無 Tab 時傳 0）
 */
function _sixCellAreas(startY) {
  const totalH = 1686 - startY;
  const rowH = Math.floor(totalH / 2);
  const colW = [833, 833, 834]; // 三欄寬（總和 2500）

  const actions = [
    { type: 'uri', uri: ACTION_URLS.我要請款 },
    { type: 'uri', uri: ACTION_URLS.查詢請款 },
    { type: 'uri', uri: ACTION_URLS.重要表單QA },
    { type: 'uri', uri: ACTION_URLS.公司活動報名 },
    { type: 'uri', uri: ACTION_URLS.讚賞幣 },
    { type: 'uri', uri: ACTION_URLS.出勤 },
  ];

  const areas = [];
  for (let row = 0; row < 2; row++) {
    let xOffset = 0;
    for (let col = 0; col < 3; col++) {
      areas.push(_area(
        xOffset,
        startY + row * rowH,
        colW[col],
        rowH,
        actions[row * 3 + col]
      ));
      xOffset += colW[col];
    }
  }
  return areas;
}

// ============================================================
// LINE API 呼叫
// ============================================================

/**
 * 根據呼叫方是正式還是測試環境，回傳對應的 Bot Token
 * 測試 LIFF（2009619528-aJO34c6u）→ 測試 channel token
 * 正式 LIFF（2009611318-5UphK9JK）→ 正式 channel token
 * 判斷依據：Line帳號 中是否有對應設定，預設用正式
 */
function _getBotToken() {
  try {
    const settings = getSettings();
    if (settings['使用測試Channel'] === true || settings['使用測試Channel'] === 'true') {
      return CONFIG.LINE_BOT_TOKEN_TEST;
    }
  } catch (_) {}
  return CONFIG.LINE_BOT_TOKEN;
}

function _lineApiPost(path, payload) {
  const response = UrlFetchApp.fetch(`https://api.line.me${path}`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${_getBotToken()}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code !== 200) throw new Error(`LINE API ${path} 失敗 (${code}): ${text}`);
  return JSON.parse(text);
}

/** 建立 Rich Menu，回傳 richMenuId */
function _createRichMenu(menuDef) {
  const result = _lineApiPost('/v2/bot/richmenu', menuDef);
  return result.richMenuId;
}

/** 上傳圖片到 Rich Menu（從 Google Drive 讀取） */
function _uploadRichMenuImage(richMenuId, driveFileId) {
  const file = DriveApp.getFileById(driveFileId);
  const blob = file.getBlob();
  const mimeType = blob.getContentType();

  const response = UrlFetchApp.fetch(
    `https://api-data.line.me/v2/bot/richmenu/${richMenuId}/content`,
    {
      method: 'post',
      contentType: mimeType,
      headers: { Authorization: `Bearer ${_getBotToken()}` },
      payload: blob.getBytes(),
      muteHttpExceptions: true,
    }
  );
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`圖片上傳失敗 (${richMenuId}): ${response.getContentText()}`);
  }
}

/** 建立 Rich Menu Alias（若已存在則先刪除再建立） */
function _createOrUpdateAlias(aliasId, richMenuId) {
  // 嘗試刪除舊的（若不存在會失敗，忽略即可）
  try {
    UrlFetchApp.fetch(`https://api.line.me/v2/bot/richmenu/alias/${aliasId}`, {
      method: 'delete',
      headers: { Authorization: `Bearer ${_getBotToken()}` },
      muteHttpExceptions: true,
    });
  } catch (e) { /* 忽略 */ }

  _lineApiPost('/v2/bot/richmenu/alias', {
    richMenuAliasId: aliasId,
    richMenuId: richMenuId,
  });
}

/** 設定全域預設 Rich Menu */
function _setDefaultRichMenu(richMenuId) {
  UrlFetchApp.fetch(
    `https://api.line.me/v2/bot/user/all/richmenu/${richMenuId}`,
    {
      method: 'post',
      headers: { Authorization: `Bearer ${_getBotToken()}` },
      muteHttpExceptions: true,
    }
  );
}

/** 將指定 Rich Menu 綁定給特定使用者 */
function _linkRichMenuToUser(lineUid, richMenuId) {
  UrlFetchApp.fetch(
    `https://api.line.me/v2/bot/user/${lineUid}/richmenu/${richMenuId}`,
    {
      method: 'post',
      headers: { Authorization: `Bearer ${_getBotToken()}` },
      muteHttpExceptions: true,
    }
  );
}

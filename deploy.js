// ============================================================
// deploy.js — 工作表初始化與版本遷移
// 全新部署：在 GAS 編輯器手動執行 deployAllSheets()
// 既有資料升級：執行 migrateV2()（冪等，可重複執行）
// ============================================================

function deployAllSheets() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);

  _deployDocSheet(ss);
  _deployClosureSheet(ss);
  _deployAuditSheet(ss);
  _deployTagSheet(ss);
  _deployDocTagSheet(ss);
  _deployGrantSheet(ss);
  _deployGroupGrantSheet(ss);
  _deployFileVersionSheet(ss);
  _ensureDocFilesFolder();

  SpreadsheetApp.flush();
  Logger.log('✅ 所有工作表初始化完成');
}

// ── V2 → V3 遷移：新增標籤主檔 / 文件標籤 / 使用者授權三張表 ──
// 冪等設計：表頭已存在就跳過，可重複執行。
function migrateV3() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  _deployTagSheet(ss);
  _deployDocTagSheet(ss);
  _deployGrantSheet(ss);
  SpreadsheetApp.flush();
  Logger.log('✅ migrateV3 完成（標籤主檔 / 文件標籤 / 使用者授權）');
}

// ── V1 → V2 遷移：補文件清單 J–M 欄與異動紀錄表 ──────────────
// 冪等設計：表頭已存在就跳過，不動既有資料列。
function migrateV2() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAMES.DOCS}，請先執行 deployAllSheets()`);

  const newHeaders = [
    { col: DOC_COL.OWNER_EMAIL + 1,  name: 'owner_email' },
    { col: DOC_COL.PUBLISHED_AT + 1, name: 'published_at' },
    { col: DOC_COL.NEXT_REVIEW + 1,  name: 'next_review_date' },
    { col: DOC_COL.REVIEW_CYCLE + 1, name: 'review_cycle_months' },
  ];

  newHeaders.forEach(h => {
    const cell = sheet.getRange(1, h.col);
    if (cell.getValue() === h.name) {
      Logger.log(`欄位已存在，跳過：${h.name}`);
      return;
    }
    cell.setValue(h.name)
        .setFontWeight('bold')
        .setBackground('#1a3a5c')
        .setFontColor('#ffffff');
    Logger.log(`✅ 補上欄位：${h.name}（第 ${h.col} 欄）`);
  });

  _applyDocSheetFormats(sheet);
  _deployAuditSheet(ss);

  SpreadsheetApp.flush();
  Logger.log('✅ migrateV2 完成');
}

// ── 建立「文件清單」工作表 ─────────────────────────────────────
function _deployDocSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.DOCS);
    Logger.log(`建立工作表：${SHEET_NAMES.DOCS}`);
  }

  const headers = [
    'doc_id', 'title', 'category', 'status',
    'owner', 'updated_at', 'version', 'google_drive_location',
    'owner_email', 'published_at', 'next_review_date', 'review_cycle_months',
    'file_id', 'pending_file_id', 'pending_version', 'pending_file_name'
  ];

  // 寫入 Header（第一列）
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a3a5c');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // 欄位格式：doc_id 強制文字避免被轉型
  sheet.getRange('A:A').setNumberFormat('@');

  // 欄寬設定
  const colWidths = [120, 240, 100, 80, 100, 120, 80, 280, 200, 110, 130, 90, 160, 160, 90, 200];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 資料驗證：status 下拉
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DOC_STATUS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(statusRule);

  // 資料驗證：category 下拉
  _applyCategoryDataValidation(sheet);

  // 條件格式：status 顏色標示
  const statusColors = {
    '草稿':  { bg: '#FEF3C7', font: '#92400E' },
    '審核中': { bg: '#DBEAFE', font: '#1E40AF' },
    '已發布': { bg: '#D1FAE5', font: '#065F46' },
    '已廢止': { bg: '#F3F4F6', font: '#6B7280' },
  };
  Object.entries(statusColors).forEach(([status, color]) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(color.bg)
      .setFontColor(color.font)
      .setRanges([sheet.getRange('D2:D1000')])
      .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  });

  _applyDocSheetFormats(sheet);

  Logger.log(`✅ ${SHEET_NAMES.DOCS} 初始化完成`);
}

// category (C) 下拉選單驗證（deployAllSheets 與 migrateDocCategoryDropdown 共用）
// 直接覆寫既有規則，天生冪等——每次都套用 DOC_CATEGORIES 目前的值，
// 不需要「已存在就跳過」的判斷。
function _applyCategoryDataValidation(sheet) {
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DOC_CATEGORIES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C2:C1000').setDataValidation(catRule);
}

// 新欄位的格式與驗證（deployAllSheets 與 migrate 共用）
function _applyDocSheetFormats(sheet) {
  // owner_email (I)、日期欄 (J, K) 強制文字，避免日期被序列化成 Date 物件
  sheet.getRange('I:I').setNumberFormat('@');
  sheet.getRange('J:K').setNumberFormat('@');
  // file_id / pending_* (M-P) 強制文字（V6）
  sheet.getRange('M:P').setNumberFormat('@');

  // 資料驗證：review_cycle_months (L) 下拉
  const cycleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(REVIEW_CYCLES.map(String), true)
    .setAllowInvalid(true) // 允許空值（未發布文件可不填）
    .build();
  sheet.getRange('L2:L1000').setDataValidation(cycleRule);
}

// ── V4 → V5 遷移：移除 owner_ID 欄位 ─────────────────────────
// 冪等設計：若第 6 欄 (原 F 欄) 表頭為 owner_ID，則將其刪除，後續欄位左移。
function migrateV5() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) {
    Logger.log(`找不到工作表：${SHEET_NAMES.DOCS}`);
    return;
  }

  const header = sheet.getRange(1, 6).getValue();
  if (String(header).toLowerCase().includes('owner_id')) {
    sheet.deleteColumn(6);
    Logger.log('✅ 已刪除 owner_ID 欄位 (原 F 欄)，後續欄位自動左移。');
  } else {
    Logger.log('✅ owner_ID 欄位已不存在，跳過刪除動作。');
  }

  // 重新套用最新版的格式與資料驗證
  _applyDocSheetFormats(sheet);

  SpreadsheetApp.flush();
  Logger.log('✅ migrateV5 完成（移除 owner_ID 欄位）');
}

// ── 建立「文件關聯」工作表 ─────────────────────────────────────
function _deployClosureSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CLOSURE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CLOSURE);
    Logger.log(`建立工作表：${SHEET_NAMES.CLOSURE}`);
  }

  const headers = ['doc_id', 'descendant_id', 'depth', 'relation_type', '說明'];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#2d5016');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // 強制文字格式
  sheet.getRange('A:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('0'); // depth 為整數

  // 欄寬
  const colWidths = [120, 120, 60, 120, 260];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 資料驗證：relation_type 下拉
  const relRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(RELATION_TYPES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(relRule);

  Logger.log(`✅ ${SHEET_NAMES.CLOSURE} 初始化完成`);
}

// ── 建立「異動紀錄」工作表（版本歷史與操作稽核合一）────────────
function _deployAuditSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.AUDIT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.AUDIT);
    Logger.log(`建立工作表：${SHEET_NAMES.AUDIT}`);
  } else if (sheet.getRange(1, 1).getValue() === '時間') {
    Logger.log(`${SHEET_NAMES.AUDIT} 已初始化，跳過`);
    return;
  }

  const headers = ['時間', '操作者', '動作', 'doc_id', '版本', '變更摘要'];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#5c1a1a');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // 時間與 doc_id 強制文字，避免 Date 序列化問題
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('@');

  const colWidths = [150, 200, 100, 120, 80, 360];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  Logger.log(`✅ ${SHEET_NAMES.AUDIT} 初始化完成`);
}

// ── V3 三張表共用建立器（冪等）─────────────────────────────
// 表頭第一格已等於預期值 → 視為已初始化，直接回傳（可重複執行）。
function _deployV3Sheet(ss, name, headers, textColLetters, bg) {
  let sheet = ss.getSheetByName(name);
  if (sheet && sheet.getRange(1, 1).getValue() === headers[0]) {
    Logger.log(`${name} 已初始化，跳過`);
    return sheet;
  }
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log(`建立工作表：${name}`);
  }
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground(bg);
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  // 指定欄位強制文字（tag_id / doc_id / email / parent_id）避免被轉型
  textColLetters.forEach(c => sheet.getRange(`${c}:${c}`).setNumberFormat('@'));
  return sheet;
}

// ── 建立「標籤主檔」工作表 ─────────────────────────────────────
function _deployTagSheet(ss) {
  const sheet = _deployV3Sheet(
    ss, SHEET_NAMES.TAGS,
    ['tag_id', 'name', 'parent_id', 'sort'],
    ['A', 'C'], '#4a2d5c');
  sheet.getRange('D:D').setNumberFormat('0'); // sort 為整數
  const colWidths = [120, 220, 120, 60];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  Logger.log(`✅ ${SHEET_NAMES.TAGS} 初始化完成`);
}

// ── 建立「文件標籤」工作表 ─────────────────────────────────────
function _deployDocTagSheet(ss) {
  const sheet = _deployV3Sheet(
    ss, SHEET_NAMES.DOC_TAGS,
    ['doc_id', 'tag_id'],
    ['A', 'B'], '#2d4a5c');
  const colWidths = [140, 140];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  Logger.log(`✅ ${SHEET_NAMES.DOC_TAGS} 初始化完成`);
}

// ── 建立「使用者授權」工作表 ───────────────────────────────────
function _deployGrantSheet(ss) {
  const sheet = _deployV3Sheet(
    ss, SHEET_NAMES.GRANTS,
    ['email', 'tag_id', 'permission'],
    ['A', 'B', 'C'], '#5c4a2d');
  const colWidths = [260, 140, 100];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  Logger.log(`✅ ${SHEET_NAMES.GRANTS} 初始化完成`);
}

// ── 建立「群組授權」工作表（V4；V5 加 permission 欄）─────────
function _deployGroupGrantSheet(ss) {
  const sheet = _deployV3Sheet(
    ss, SHEET_NAMES.GROUP_GRANTS,
    ['org_code', 'title', 'tag_id', 'permission'],
    ['A', 'B', 'C', 'D'], '#2d5c4a');
  const colWidths = [160, 160, 140, 100];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  Logger.log(`✅ ${SHEET_NAMES.GROUP_GRANTS} 初始化完成`);
}

// ── 建立「檔案版本」工作表（V6）───────────────────────────────
function _deployFileVersionSheet(ss) {
  const sheet = _deployV3Sheet(
    ss, SHEET_NAMES.FILE_VERSIONS,
    ['doc_id', 'version', 'file_id', 'file_name', 'uploaded_by', 'uploaded_at'],
    ['A', 'C'], '#5c2d2d');
  const colWidths = [120, 80, 220, 240, 200, 150];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  Logger.log(`✅ ${SHEET_NAMES.FILE_VERSIONS} 初始化完成`);
}

// 確保中央檔案資料夾存在；未設定 Script Property 時自動建立並寫回。
function _ensureDocFilesFolder() {
  const existing = _getProp(PROP_KEYS.DOC_FILES_FOLDER_ID);
  if (existing) {
    try {
      DriveApp.getFolderById(existing); // 驗證仍可存取
      return existing;
    } catch (e) {
      Logger.log(`⚠️ DOC_FILES_FOLDER_ID（${existing}）已無法存取，將建立新資料夾：${e}`);
    }
  }
  const folder = DriveApp.createFolder('文件管理系統檔案庫');
  PropertiesService.getScriptProperties().setProperty(PROP_KEYS.DOC_FILES_FOLDER_ID, folder.getId());
  Logger.log(`✅ 已建立中央檔案資料夾並寫入 Script Properties：${folder.getId()}`);
  return folder.getId();
}

// ── V3 → V4 遷移：新增群組授權表 ─────────────────────────────
// 冪等設計：表頭已存在就跳過，可重複執行。
function migrateV4() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  _deployGroupGrantSheet(ss);
  SpreadsheetApp.flush();
  Logger.log('✅ migrateV4 完成（群組授權）');
}

// ── V4 → V5 遷移：兩張授權表補 permission 欄 ─────────────────
// 冪等設計：表頭已存在就跳過。空白 permission 由讀取端視為 read，
// 先跑本函式或先部署程式碼皆安全，無需資料搬遷。
function migrateV6() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const targets = [
    { name: SHEET_NAMES.GRANTS,       col: GRANT_COL.PERMISSION + 1,      bg: '#5c4a2d' },
    { name: SHEET_NAMES.GROUP_GRANTS, col: GROUPGRANT_COL.PERMISSION + 1, bg: '#2d5c4a' },
  ];
  targets.forEach(t => {
    const sheet = ss.getSheetByName(t.name);
    if (!sheet) {
      Logger.log(`找不到工作表：${t.name}，請先執行 migrateV3()/migrateV4()`);
      return;
    }
    const cell = sheet.getRange(1, t.col);
    if (cell.getValue() === 'permission') {
      Logger.log(`欄位已存在，跳過：${t.name}.permission`);
      return;
    }
    cell.setValue('permission')
        .setFontWeight('bold')
        .setBackground(t.bg)
        .setFontColor('#ffffff');
    // permission 欄強制文字格式，與其他授權欄一致
    const letter = String.fromCharCode(64 + t.col); // C 或 D
    sheet.getRange(`${letter}:${letter}`).setNumberFormat('@');
    sheet.setColumnWidth(t.col, 100);
    Logger.log(`✅ 補上欄位：${t.name}.permission`);
  });
  SpreadsheetApp.flush();
  Logger.log('✅ migrateV6 完成（授權表 permission 欄）');
}

// ── V5 → V6 遷移：文件清單補 M–P 欄、建檔案版本表、確保中央資料夾 ──
// 冪等設計：表頭已存在就跳過，不動既有資料列；可重複執行。
function migrateV7() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAMES.DOCS}，請先執行 deployAllSheets()`);

  const newHeaders = [
    { col: DOC_COL.FILE_ID + 1,           name: 'file_id' },
    { col: DOC_COL.PENDING_FILE_ID + 1,   name: 'pending_file_id' },
    { col: DOC_COL.PENDING_VERSION + 1,   name: 'pending_version' },
    { col: DOC_COL.PENDING_FILE_NAME + 1, name: 'pending_file_name' },
  ];

  newHeaders.forEach(h => {
    const cell = sheet.getRange(1, h.col);
    if (cell.getValue() === h.name) {
      Logger.log(`欄位已存在，跳過：${h.name}`);
      return;
    }
    cell.setValue(h.name)
        .setFontWeight('bold')
        .setBackground('#1a3a5c')
        .setFontColor('#ffffff');
    Logger.log(`✅ 補上欄位：${h.name}（第 ${h.col} 欄）`);
  });

  _applyDocSheetFormats(sheet);
  _deployFileVersionSheet(ss);
  const folderId = _ensureDocFilesFolder();

  SpreadsheetApp.flush();
  Logger.log(`✅ migrateV7 完成（檔案版本管理；中央資料夾 ID：${folderId}）`);
}

// ── V6 選用工具：搬移既有手動貼的 Drive 連結進中央資料夾 ─────
// 管理員在 GAS 編輯器手動執行一次。掃描有 google_drive_location 但無
// file_id 的文件：複製檔案進中央資料夾（不動原檔），登記 file_id 與
// 檔案版本列。無法存取的連結（權限不足／格式無法辨識）記錄在 log，
// 不中斷其餘文件的處理。
function migrateV7ImportLegacyFiles() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAMES.DOCS}`);

  const folderId = _ensureDocFilesFolder();
  const rootFolder = DriveApp.getFolderById(folderId);
  const rows = sheet.getDataRange().getDisplayValues();
  const fileVerSheet = _getSheet(SHEET_NAMES.FILE_VERSIONS);

  const skipped = [];
  let migrated = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const docId = row[DOC_COL.DOC_ID];
    if (!docId) continue;
    const driveLoc = row[DOC_COL.GOOGLE_DRIVE_LOC];
    const existingFileId = row[DOC_COL.FILE_ID];
    if (!driveLoc || existingFileId) continue; // 無連結或已搬移過，跳過

    const m = String(driveLoc).match(/[-\w]{25,}/);
    if (!m) {
      skipped.push(`${docId}：無法從連結解析出檔案 ID（${driveLoc}）`);
      continue;
    }
    const sourceId = m[0];

    try {
      const sourceFile = DriveApp.getFileById(sourceId);
      let docFolder;
      const existing = rootFolder.getFoldersByName(docId);
      docFolder = existing.hasNext() ? existing.next() : rootFolder.createFolder(docId);
      const copy = sourceFile.makeCopy(sourceFile.getName(), docFolder);
      const version = row[DOC_COL.VERSION] || '0.1';

      sheet.getRange(i + 1, DOC_COL.FILE_ID + 1).setValue(copy.getId());
      fileVerSheet.appendRow([docId, version, copy.getId(), sourceFile.getName(), 'migrateV7ImportLegacyFiles', Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')]);
      migrated++;
      Logger.log(`✅ ${docId}：已搬移「${sourceFile.getName()}」`);
    } catch (e) {
      skipped.push(`${docId}：搬移失敗（${e.message || e}）`);
    }
  }

  SpreadsheetApp.flush();
  Logger.log(`✅ migrateV7ImportLegacyFiles 完成：搬移 ${migrated} 筆，跳過 ${skipped.length} 筆`);
  if (skipped.length) {
    Logger.log('未搬移清單（需人工處理）：\n' + skipped.join('\n'));
  }
}

// ── 選用工具：重新套用文件類別下拉選單 ──────────────────────
// C 欄的下拉選單驗證是建表當下寫死到儲存格的規則，不會因為改了
// env.js 的 DOC_CATEGORIES 常數就自動更新既有試算表。改動
// DOC_CATEGORIES 後，在既有已部署的試算表上執行本函式一次即可
// 讓下拉選單同步成目前的清單；天生冪等，可重複執行。
// 注意：僅更新下拉選單本身，不會改寫既有文件列上已經寫入的舊分類
// 字串——那些需要另外盤點、人工決定新分類的對應規則。
function migrateDocCategoryDropdown() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAMES.DOCS}，請先執行 deployAllSheets()`);

  _applyCategoryDataValidation(sheet);

  SpreadsheetApp.flush();
  Logger.log(`✅ migrateDocCategoryDropdown 完成，目前分類清單：${DOC_CATEGORIES.join('、')}`);
}

// ── 寫入測試資料（選用）──────────────────────────────────────
function seedSampleData() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const docSheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  const clsSheet = ss.getSheetByName(SHEET_NAMES.CLOSURE);

  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
  // 範例的下次審查日：發布日 + 12 個月
  const nextYear = Utilities.formatDate(
    new Date(new Date().setFullYear(new Date().getFullYear() + 1)),
    'Asia/Taipei', 'yyyy/MM/dd');

  const docs = [
    ['DOC-001', '資訊安全政策總綱', 'ISO管理系統', '已發布', '王小明', 'U001', today, '1.0', '', 'ming@example.com',  today, nextYear, 12],
    ['DOC-002', '存取控制程序書',   '標準作業流程(cSOP)', '已發布', '李大華', 'U002', today, '2.1', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-003', '事件回應程序書',   '標準作業流程(cSOP)', '審核中', '陳美玲', 'U003', today, '1.3', '', 'mei@example.com',   '', '', 12],
    ['DOC-004', '稽核作業程序書',   '管理辦法(C)', '草稿',   '王小明', 'U001', today, '0.2', '', 'ming@example.com',  '', '', 12],
    ['DOC-005', '帳號申請表單',     '紀錄', '已發布', '李大華', 'U002', today, '1.0', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-006', '權限異動紀錄表',   '紀錄', '已發布', '李大華', 'U002', today, '1.0', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-007', '帳號停用申請單',   '紀錄', '草稿',   '陳美玲', 'U003', today, '0.1', '', 'mei@example.com',   '', '', 12],
  ];

  docSheet.getRange(2, 1, docs.length, docs[0].length).setValues(docs);

  // Closure Table 資料（self + 關聯）
  const closures = [
    // self records
    ['DOC-001','DOC-001',0,'references','自身'],
    ['DOC-002','DOC-002',0,'references','自身'],
    ['DOC-003','DOC-003',0,'references','自身'],
    ['DOC-004','DOC-004',0,'references','自身'],
    ['DOC-005','DOC-005',0,'references','自身'],
    ['DOC-006','DOC-006',0,'references','自身'],
    ['DOC-007','DOC-007',0,'references','自身'],
    // DOC-001 → DOC-002, 003, 004（直接，depth=1）
    ['DOC-001','DOC-002',1,'references','政策總綱引用存取控制程序書'],
    ['DOC-001','DOC-003',1,'references','政策總綱引用事件回應程序書'],
    ['DOC-001','DOC-004',1,'references','政策總綱引用稽核作業程序書'],
    // DOC-002 → DOC-005, 006, 007（直接，depth=1）
    ['DOC-002','DOC-005',1,'references','存取控制引用帳號申請表單'],
    ['DOC-002','DOC-006',1,'references','存取控制引用權限異動紀錄表'],
    ['DOC-002','DOC-007',1,'references','存取控制引用帳號停用申請單'],
    // DOC-001 → DOC-005, 006, 007（間接，depth=2）
    ['DOC-001','DOC-005',2,'references','政策總綱間接關聯（透過存取控制）'],
    ['DOC-001','DOC-006',2,'references','政策總綱間接關聯（透過存取控制）'],
    ['DOC-001','DOC-007',2,'references','政策總綱間接關聯（透過存取控制）'],
  ];

  clsSheet.getRange(2, 1, closures.length, closures[0].length).setValues(closures);

  // ── V3 標籤範例（選用）：兩層標籤樹＋為部分文件貼標 ──────────
  const tagSheet = ss.getSheetByName(SHEET_NAMES.TAGS);
  const docTagSheet = ss.getSheetByName(SHEET_NAMES.DOC_TAGS);
  if (tagSheet && docTagSheet && tagSheet.getLastRow() < 2) {
    const tags = [
      ['TAG-001', '資訊安全',   '',        1],
      ['TAG-002', '政策層',     'TAG-001', 1],
      ['TAG-003', '程序層',     'TAG-001', 2],
      ['TAG-004', '表單',       '',        2],
    ];
    tagSheet.getRange(2, 1, tags.length, tags[0].length).setValues(tags);

    const docTags = [
      ['DOC-001', 'TAG-002'],
      ['DOC-002', 'TAG-003'],
      ['DOC-003', 'TAG-003'],
      ['DOC-004', 'TAG-003'],
      ['DOC-005', 'TAG-004'],
      ['DOC-006', 'TAG-004'],
      ['DOC-007', 'TAG-004'],
    ];
    docTagSheet.getRange(2, 1, docTags.length, docTags[0].length).setValues(docTags);
    Logger.log('✅ 標籤範例寫入完成');
  }

  SpreadsheetApp.flush();
  Logger.log('✅ 範例資料寫入完成');
}

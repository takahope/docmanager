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
    'owner_email', 'published_at', 'next_review_date', 'review_cycle_months'
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
  const colWidths = [120, 240, 100, 80, 100, 120, 80, 280, 200, 110, 130, 90];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 資料驗證：status 下拉
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DOC_STATUS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(statusRule);

  // 資料驗證：category 下拉
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DOC_CATEGORIES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('C2:C1000').setDataValidation(catRule);

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

// 新欄位的格式與驗證（deployAllSheets 與 migrate 共用）
function _applyDocSheetFormats(sheet) {
  // owner_email (I)、日期欄 (J, K) 強制文字，避免日期被序列化成 Date 物件
  sheet.getRange('I:I').setNumberFormat('@');
  sheet.getRange('J:K').setNumberFormat('@');

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
    ['DOC-001', '資訊安全政策總綱', 'ISMS', '已發布', '王小明', 'U001', today, '1.0', '', 'ming@example.com',  today, nextYear, 12],
    ['DOC-002', '存取控制程序書',   'ISMS', '已發布', '李大華', 'U002', today, '2.1', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-003', '事件回應程序書',   'ISMS', '審核中', '陳美玲', 'U003', today, '1.3', '', 'mei@example.com',   '', '', 12],
    ['DOC-004', '稽核作業程序書',   'ISMS', '草稿',   '王小明', 'U001', today, '0.2', '', 'ming@example.com',  '', '', 12],
    ['DOC-005', '帳號申請表單',     '表單', '已發布', '李大華', 'U002', today, '1.0', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-006', '權限異動紀錄表',   '表單', '已發布', '李大華', 'U002', today, '1.0', '', 'hua@example.com',   today, nextYear, 12],
    ['DOC-007', '帳號停用申請單',   '表單', '草稿',   '陳美玲', 'U003', today, '0.1', '', 'mei@example.com',   '', '', 12],
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

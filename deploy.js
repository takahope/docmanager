// ============================================================
// deploy.js — 工作表初始化
// 在 GAS 編輯器手動執行 deployAllSheets() 以建立工作表與 Header
// ============================================================

function deployAllSheets() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);

  _deployDocSheet(ss);
  _deployClosureSheet(ss);

  SpreadsheetApp.flush();
  Logger.log('✅ 所有工作表初始化完成');
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
    'owner', 'owner_ID', 'updated_at', 'version', 'google_drive_location'
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
  sheet.getRange('F:F').setNumberFormat('@'); // owner_ID 也強制文字

  // 欄寬設定
  const colWidths = [120, 240, 100, 80, 100, 120, 120, 80, 280];
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

  Logger.log(`✅ ${SHEET_NAMES.DOCS} 初始化完成`);
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

// ── 寫入測試資料（選用）──────────────────────────────────────
function seedSampleData() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const docSheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  const clsSheet = ss.getSheetByName(SHEET_NAMES.CLOSURE);

  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');

  const docs = [
    ['DOC-001', '資訊安全政策總綱', 'ISMS', '已發布', '王小明', 'U001', today, '1.0', ''],
    ['DOC-002', '存取控制程序書',   'ISMS', '已發布', '李大華', 'U002', today, '2.1', ''],
    ['DOC-003', '事件回應程序書',   'ISMS', '審核中', '陳美玲', 'U003', today, '1.3', ''],
    ['DOC-004', '稽核作業程序書',   'ISMS', '草稿',   '王小明', 'U001', today, '0.2', ''],
    ['DOC-005', '帳號申請表單',     '表單', '已發布', '李大華', 'U002', today, '1.0', ''],
    ['DOC-006', '權限異動紀錄表',   '表單', '已發布', '李大華', 'U002', today, '1.0', ''],
    ['DOC-007', '帳號停用申請單',   '表單', '草稿',   '陳美玲', 'U003', today, '0.1', ''],
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
  SpreadsheetApp.flush();
  Logger.log('✅ 範例資料寫入完成');
}

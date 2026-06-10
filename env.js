// ============================================================
// env.js — 環境設定、工作表名稱、欄位索引集中定義
// 所有讀寫試算表的程式碼必須透過此處的常數操作
// ============================================================

// ── 試算表 ID（部署後填入）──────────────────────────────────
const ENV = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(), // 當前試算表
};

// ── 工作表名稱 ────────────────────────────────────────────────
const SHEET_NAMES = {
  DOCS:     '文件清單',
  CLOSURE:  '文件關聯',
};

// ── 文件清單欄位索引（0-based，對應 getValues() 陣列）────────
const DOC_COL = {
  DOC_ID:              0,  // A: doc_id
  TITLE:               1,  // B: title
  CATEGORY:            2,  // C: category
  STATUS:              3,  // D: status
  OWNER:               4,  // E: owner
  OWNER_ID:            5,  // F: owner_ID
  UPDATED_AT:          6,  // G: updated_at
  VERSION:             7,  // H: version
  GOOGLE_DRIVE_LOC:    8,  // I: google_drive_location
};

// ── 文件關聯欄位索引（0-based）───────────────────────────────
const CLS_COL = {
  ANCESTOR_ID:    0,  // A: doc_id（祖先）
  DESCENDANT_ID:  1,  // B: descendant_id（子孫）
  DEPTH:          2,  // C: depth
  RELATION_TYPE:  3,  // D: relation_type
  DESCRIPTION:    4,  // E: 說明
};

// ── 文件狀態選項 ──────────────────────────────────────────────
const DOC_STATUS = ['草稿', '審核中', '已發布', '已廢止'];

// ── 關聯類型選項 ──────────────────────────────────────────────
const RELATION_TYPES = ['references', 'supersedes', 'derived_from', 'related'];

// ── 文件類別選項 ──────────────────────────────────────────────
const DOC_CATEGORIES = ['ISMS', 'PIMS', '表單', 'SOP', '政策', '指引'];

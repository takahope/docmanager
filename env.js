// ============================================================
// env.js — 環境設定、工作表名稱、欄位索引集中定義
// 所有讀寫試算表的程式碼必須透過此處的常數操作
// 機敏值（HR 試算表 ID、管理員名單）一律放 Script Properties，
// 本檔只放 key 名稱常數，不放實際值。
// ============================================================

// ── 試算表 ID ────────────────────────────────────────────────
const ENV = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(), // 當前試算表
};

// ── Script Properties key 名稱 ───────────────────────────────
// HR_SPREADSHEET_ID：HR 主資料試算表 ID（白名單、負責人名單來源）
// ADMIN_EMAILS：管理員信箱，逗號分隔
const PROP_KEYS = {
  HR_SPREADSHEET_ID: 'HR_SPREADSHEET_ID',
  ADMIN_EMAILS:      'ADMIN_EMAILS',
};

// ── HR 主表「人員主檔」結構（依 ECOSYSTEM.md 契約：A 信箱、B 姓名、C 狀態）──
const HR_SHEET_NAME = '人員主檔';
const HR_COL = {
  EMAIL:  0,  // A: 信箱（主鍵）
  NAME:   1,  // B: 姓名
  STATUS: 2,  // C: 狀態
};

// 白名單過濾採「排除法」：中文狀態值域多且精確比對不可靠，
// 只排除確定不該有存取權的狀態，其餘一律放行。
const EXCLUDED_HR_STATUS = ['離職'];

// ── 工作表名稱 ────────────────────────────────────────────────
const SHEET_NAMES = {
  DOCS:     '文件清單',
  CLOSURE:  '文件關聯',
  AUDIT:    '異動紀錄',
};

// ── 文件清單欄位索引（0-based，對應 getValues() 陣列）────────
const DOC_COL = {
  DOC_ID:              0,   // A: doc_id
  TITLE:               1,   // B: title
  CATEGORY:            2,   // C: category
  STATUS:              3,   // D: status
  OWNER:               4,   // E: owner
  OWNER_ID:            5,   // F: owner_ID
  UPDATED_AT:          6,   // G: updated_at
  VERSION:             7,   // H: version
  GOOGLE_DRIVE_LOC:    8,   // I: google_drive_location
  OWNER_EMAIL:         9,   // J: owner_email（權限比對用）
  PUBLISHED_AT:        10,  // K: published_at（轉「已發布」時自動填）
  NEXT_REVIEW:         11,  // L: next_review_date（發布日 + 審查週期）
  REVIEW_CYCLE:        12,  // M: review_cycle_months（審查週期，月）
};
const DOC_COL_COUNT = 13;   // 文件清單總欄數（appendRow / setValues 用）

// ── 異動紀錄欄位索引（0-based）───────────────────────────────
const AUDIT_COL = {
  TIMESTAMP: 0,  // A: 時間
  OPERATOR:  1,  // B: 操作者
  ACTION:    2,  // C: 動作
  DOC_ID:    3,  // D: doc_id
  VERSION:   4,  // E: 版本
  SUMMARY:   5,  // F: 變更摘要
};
const AUDIT_COL_COUNT = 6;

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

// ── 狀態流轉表（Map 查表取代 if-else 分支）──────────────────
// key = 目前狀態，value = 允許轉入的下一個狀態
// 「已廢止 → 草稿」僅限管理員（於 code.js 檢查角色）
const STATUS_TRANSITIONS = {
  '草稿':   ['審核中'],
  '審核中': ['已發布', '草稿'],
  '已發布': ['已廢止', '審核中'],
  '已廢止': ['草稿'],
};

// ── 審查週期選項（月）────────────────────────────────────────
const REVIEW_CYCLES = [6, 12, 24, 36];
const DEFAULT_REVIEW_CYCLE = 12;

// ── 關聯類型選項 ──────────────────────────────────────────────
const RELATION_TYPES = ['references', 'supersedes', 'derived_from', 'related'];

// ── 文件類別選項 ──────────────────────────────────────────────
const DOC_CATEGORIES = ['ISMS', 'PIMS', '表單', 'SOP', '政策', '指引'];

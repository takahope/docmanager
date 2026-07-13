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
  HR_SPREADSHEET_ID:   'HR_SPREADSHEET_ID',
  ADMIN_EMAILS:        'ADMIN_EMAILS',
  DOC_FILES_FOLDER_ID: 'DOC_FILES_FOLDER_ID', // V6：文件檔案中央資料夾 ID（migrateV7 可自動建立並寫回）
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

// ── HR 主表「組織架構樹」結構（依 docs/HR_sheet.md 契約）─────
// V4 群組授權用；動用前先跑 debugGetSystemData() 核對實表表頭。
const HR_ORG_SHEET_NAME = '組織架構樹';
const HR_ORG_COL = {
  TYPE:          0,  // A: 組織類型（ORG/TF/PARTNER/GOV）
  LEVEL:         1,  // B: 層級（1=EGC … 6=駐站）
  CODE:          2,  // C: 代碼（主鍵，如 DEPT-BIO）
  NAME:          3,  // D: 名稱
  ALIAS:         4,  // E: 別名
  PARENT_CODE:   5,  // F: 上級代碼（根節點空）
  MANAGER_EMAIL: 6,  // G: 管理人員信箱
  MANAGER_NAME:  7,  // H: 管理人員姓名
};

// ── HR 主表「人員職務配置」結構（依 docs/HR_sheet.md 契約）───
// 矩陣兼任＝一人多列；V4 群組授權的成員名冊唯一來源。
const HR_ASSIGN_SHEET_NAME = '人員職務配置';
const HR_ASSIGN_COL = {
  EMAIL:         0,  // A: 信箱（FK 人員主檔）
  NAME:          1,  // B: 姓名（冗餘欄）
  ORG_CODE:      2,  // C: 所屬組別代碼（FK 組織架構樹.代碼）
  ORG_NAME:      3,  // D: 所屬組別（冗餘欄）
  TITLE:         4,  // E: 職稱
  MANAGER_EMAIL: 5,  // F: 主管信箱
  MANAGER_NAME:  6,  // G: 直屬主管（冗餘欄）
};

// ── 工作表名稱 ────────────────────────────────────────────────
const SHEET_NAMES = {
  DOCS:     '文件清單',
  CLOSURE:  '文件關聯',
  AUDIT:    '異動紀錄',
  TAGS:     '標籤主檔',   // V3：標籤樹（adjacency list）
  DOC_TAGS: '文件標籤',   // V3：文件↔標籤多對多
  GRANTS:   '使用者授權', // V3：使用者↔標籤授權
  GROUP_GRANTS: '群組授權', // V4：群組（組織/職稱）↔ 標籤授權
  FILE_VERSIONS: '檔案版本', // V6：每個正式生效版本一列（doc_id, version, file_id, ...）
};

// ── 文件清單欄位索引（0-based，對應 getValues() 陣列）────────
const DOC_COL = {
  DOC_ID:              0,   // A: doc_id
  TITLE:               1,   // B: title
  CATEGORY:            2,   // C: category
  STATUS:              3,   // D: status
  OWNER:               4,   // E: owner
  UPDATED_AT:          5,   // F: updated_at
  VERSION:             6,   // G: version
  GOOGLE_DRIVE_LOC:    7,   // H: google_drive_location（V6：legacy 手動連結，僅顯示，不再接受輸入）
  OWNER_EMAIL:         8,   // I: owner_email（權限比對用）
  PUBLISHED_AT:        9,   // J: published_at（轉「已發布」時自動填）
  NEXT_REVIEW:         10,  // K: next_review_date（發布日 + 審查週期）
  REVIEW_CYCLE:        11,  // L: review_cycle_months（審查週期，月）
  FILE_ID:             12,  // M: file_id（V6：現行生效版的 Drive 檔案 ID）
  PENDING_FILE_ID:     13,  // N: pending_file_id（V6：待核版檔案 ID，核准發布時 promote）
  PENDING_VERSION:     14,  // O: pending_version（V6：待核版版號）
  PENDING_FILE_NAME:   15,  // P: pending_file_name（V6：待核版原始檔名）
};
const DOC_COL_COUNT = 16;   // 文件清單總欄數（appendRow / setValues 用）

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

// ── 標籤主檔欄位索引（0-based）（V3）────────────────────────
// adjacency list：parent_id 空＝根節點；標籤數量小，前端遞迴建樹。
const TAG_COL = {
  TAG_ID:    0,  // A: tag_id（TAG-001 格式，文字格式）
  NAME:      1,  // B: name（標籤名稱）
  PARENT_ID: 2,  // C: parent_id（父標籤 tag_id，空＝根）
  SORT:      3,  // D: sort（同層排序）
};
const TAG_COL_COUNT = 4;

// ── 文件標籤欄位索引（0-based）（V3）────────────────────────
// 多對多，一列一組 (doc_id, tag_id)。
const DOCTAG_COL = {
  DOC_ID: 0,  // A: doc_id
  TAG_ID: 1,  // B: tag_id
};
const DOCTAG_COL_COUNT = 2;

// ── 使用者授權欄位索引（0-based）（V3；V5 加 permission）────
// 純標籤授權，一列一組 (email, tag_id, permission)；email 小寫。
// permission：'edit' 或 'read'；空白／其他值一律視為 read（fail-closed）。
const GRANT_COL = {
  EMAIL:      0,  // A: email（小寫）
  TAG_ID:     1,  // B: tag_id
  PERMISSION: 2,  // C: permission（read/edit；空白＝read）
};
const GRANT_COL_COUNT = 3;

// ── 群組授權欄位索引（0-based）（V4；V5 加 permission 欄）────
// (org_code, title) 至少一欄非空；一列一組授權。
// org_code 精確比對職務配置（僅直屬成員，不含子單位——與標籤樹的
// 父含子繼承「相反」，這是設計定案，勿順手改成展開子樹）。
const GROUPGRANT_COL = {
  ORG_CODE:   0,  // A: org_code（HR 組織架構樹.代碼；可空）
  TITLE:      1,  // B: title（HR 人員職務配置.職稱；可空）
  TAG_ID:     2,  // C: tag_id（標籤主檔）
  PERMISSION: 3,  // D: permission（read/edit；空白＝read）（V5）
};
const GROUPGRANT_COL_COUNT = 4;

// ── 檔案版本欄位索引（0-based）（V6）─────────────────────────
// 每個「正式生效」版本一列；pending（待核）版不入表，核准發布時才 append。
const FILEVER_COL = {
  DOC_ID:      0,  // A: doc_id
  VERSION:     1,  // B: version（生效版號）
  FILE_ID:     2,  // C: file_id（Drive 檔案 ID）
  FILE_NAME:   3,  // D: file_name（原始檔名）
  UPLOADED_BY: 4,  // E: uploaded_by（上傳者信箱）
  UPLOADED_AT: 5,  // F: uploaded_at（核准發布時間）
};
const FILEVER_COL_COUNT = 6;

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

// ── 版號改版類型（V6，上傳檔案時二選一，白名單驗證）────────────
// 'start'（管理員自訂起始版號，僅文件首次上傳可用）刻意不列入此白名單——
// 它不經過 _bumpVersion，是 apiUploadDocFile 內獨立的分流與檢查，
// 讓 _bumpVersion 永遠只認 minor/major，避免日後被誤接進遞增邏輯。
const VERSION_BUMP_TYPES = ['minor', 'major']; // minor: X.Y→X.(Y+1)；major: X.Y→(X+1).0

// ── 文件類別選項 ──────────────────────────────────────────────
const DOC_CATEGORIES = ['ISO管理系統', '管理辦法(C)', '準則(cPOT)', '標準作業流程(cSOP)', '其他(cDOC)', '外來文件', '紀錄'];

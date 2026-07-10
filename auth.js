// ============================================================
// auth.js — 權限層：HR 白名單、角色判定、後端存取檢查
//
// 角色模型（由高至低）：
//   管理員（Script Properties ADMIN_EMAILS）：全部操作
//   文件負責人（owner_email === 登入者）：編輯自己負責的文件與其關聯
//   一般白名單使用者：唯讀＋建立新文件（自動成為負責人）
//   非白名單：拒絕進入
//
// 依 PLAYBOOK P6：HR 試算表 ID 與管理員名單放 Script Properties，
// 不寫進程式碼。key 名稱見 env.js 的 PROP_KEYS。
// ============================================================

const HR_CACHE_KEY = 'hr_people_v1';
const HR_CACHE_SECONDS = 600; // 10 分鐘：白名單檢查頻繁，避免每次跨表整表讀取

function _getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

// ── HR 人員名單（白名單來源＋負責人下拉選項）──────────────────
// 回傳 [{email, name}]，已排除 EXCLUDED_HR_STATUS 狀態者。
// 中文狀態值域多且精確比對不可靠 → 採排除法而非白名單法。
function _getHrPeople() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(HR_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const hrId = _getProp(PROP_KEYS.HR_SPREADSHEET_ID);
  if (!hrId) throw new Error('尚未設定 Script Properties：HR_SPREADSHEET_ID');

  const sheet = SpreadsheetApp.openById(hrId).getSheetByName(HR_SHEET_NAME);
  if (!sheet) throw new Error(`HR 主表找不到工作表：${HR_SHEET_NAME}`);

  const rows = sheet.getDataRange().getDisplayValues();
  const people = rows.slice(1)
    .filter(r => r[HR_COL.EMAIL])
    .filter(r => !EXCLUDED_HR_STATUS.includes(String(r[HR_COL.STATUS]).trim()))
    .map(r => ({
      email: String(r[HR_COL.EMAIL]).trim().toLowerCase(),
      name:  String(r[HR_COL.NAME]).trim(),
    }));

  cache.put(HR_CACHE_KEY, JSON.stringify(people), HR_CACHE_SECONDS);
  return people;
}

function _getAdminEmails() {
  return _getProp(PROP_KEYS.ADMIN_EMAILS)
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);
}

function _getCurrentEmail() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

// ── 使用者情境（doGet 與 apiGetInitData 共用）────────────────
function getUserContext() {
  const email = _getCurrentEmail();
  const admins = _getAdminEmails();
  const isAdmin = admins.includes(email);

  let isWhitelisted = isAdmin; // 管理員即使不在 HR 名單也放行
  let name = '';
  if (email) {
    const person = _getHrPeople().find(p => p.email === email);
    if (person) {
      isWhitelisted = true;
      name = person.name;
    }
  }

  return { email: email, name: name, isAdmin: isAdmin, isWhitelisted: isWhitelisted };
}

// ── 後端存取檢查（每個寫入型 API 進入時呼叫）─────────────────
// 檢查失敗直接 throw，由各 API 的 try/catch 轉成 {success:false}。

function _assertWhitelisted() {
  const ctx = getUserContext();
  if (!ctx.isWhitelisted) {
    throw new Error(`無存取權限：${ctx.email || '（無法取得登入信箱）'} 不在人員名單中`);
  }
  return ctx;
}

// 可編輯此文件？管理員一律可；文件負責人限自己負責的文件。
function _assertCanEditDoc(docId) {
  const ctx = _assertWhitelisted();
  if (ctx.isAdmin) return ctx;

  const doc = _readDocs().find(d => d.doc_id === docId);
  if (!doc) throw new Error('找不到文件：' + docId);
  if (String(doc.owner_email).toLowerCase() !== ctx.email) {
    throw new Error('權限不足：只有管理員或文件負責人可以編輯此文件');
  }
  return ctx;
}

function _assertAdmin() {
  const ctx = _assertWhitelisted();
  if (!ctx.isAdmin) throw new Error('權限不足：此操作僅限管理員');
  return ctx;
}

// ============================================================
// V3 標籤式可見性層
//   _getVisibleDocIds(ctx) 是「唯一的可見性事實來源」，
//   所有讀取 API 都必須透過它過濾，任何 API 不得繞過。
// ============================================================

// 讀取「標籤主檔」→ [{tag_id, name, parent_id, sort}]
function _readTags() {
  const sheet = _getSheet(SHEET_NAMES.TAGS);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1)
    .filter(r => r[TAG_COL.TAG_ID])
    .map(r => ({
      tag_id:    r[TAG_COL.TAG_ID],
      name:      r[TAG_COL.NAME],
      parent_id: r[TAG_COL.PARENT_ID] || '',
      sort:      parseInt(r[TAG_COL.SORT], 10) || 0,
    }));
}

// 讀取「文件標籤」→ [{doc_id, tag_id}]
function _readDocTags() {
  const sheet = _getSheet(SHEET_NAMES.DOC_TAGS);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1)
    .filter(r => r[DOCTAG_COL.DOC_ID] && r[DOCTAG_COL.TAG_ID])
    .map(r => ({
      doc_id: r[DOCTAG_COL.DOC_ID],
      tag_id: r[DOCTAG_COL.TAG_ID],
    }));
}

// 讀取「使用者授權」→ [{email, tag_id}]（email 一律小寫）
function _readGrants() {
  const sheet = _getSheet(SHEET_NAMES.GRANTS);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1)
    .filter(r => r[GRANT_COL.EMAIL] && r[GRANT_COL.TAG_ID])
    .map(r => ({
      email:  String(r[GRANT_COL.EMAIL]).trim().toLowerCase(),
      tag_id: r[GRANT_COL.TAG_ID],
    }));
}

// 把一組標籤展開為「含全部子孫」的 tag_id 集合（BFS on parent_id）。
// 授權父標籤即涵蓋整棵子樹，這裡即實現「父含子」的樹狀繼承。
function _expandTagWithDescendants(tagIds, allTags) {
  const childrenMap = new Map();
  allTags.forEach(t => {
    if (!childrenMap.has(t.parent_id)) childrenMap.set(t.parent_id, []);
    childrenMap.get(t.parent_id).push(t.tag_id);
  });
  const result = new Set();
  const queue = Array.from(tagIds || []);
  while (queue.length > 0) {
    const id = queue.shift();
    if (!id || result.has(id)) continue;
    result.add(id);
    (childrenMap.get(id) || []).forEach(c => {
      if (!result.has(c)) queue.push(c);
    });
  }
  return result;
}

// 【唯一可見性事實來源】回傳當前使用者可見的 doc_id 集合（Set）。
//   1. 管理員 → 全部文件
//   2. 自己是 owner_email 的文件 → 可見（不受標籤限制）
//   3. 文件的任一標籤 ∈ 使用者授權標籤的子孫展開集 → 可見
//   4. 無標籤文件 → 僅落入規則 1、2（deny by default）
function _getVisibleDocIds(ctx) {
  const docs = _readDocs();
  if (ctx && ctx.isAdmin) {
    return new Set(docs.map(d => d.doc_id));
  }

  const email = ctx ? String(ctx.email || '').toLowerCase() : '';

  // 使用者被授權的標籤 → 展開子孫
  const grantedTagIds = _readGrants()
    .filter(g => g.email === email)
    .map(g => g.tag_id);
  const grantedExpanded = _expandTagWithDescendants(grantedTagIds, _readTags());

  // doc_id → [tag_id...]
  const docTagMap = new Map();
  _readDocTags().forEach(dt => {
    if (!docTagMap.has(dt.doc_id)) docTagMap.set(dt.doc_id, []);
    docTagMap.get(dt.doc_id).push(dt.tag_id);
  });

  const visible = new Set();
  docs.forEach(d => {
    // 規則 2：文件負責人
    if (email && String(d.owner_email || '').toLowerCase() === email) {
      visible.add(d.doc_id);
      return;
    }
    // 規則 3：任一標籤落在授權子孫集
    const tags = docTagMap.get(d.doc_id) || [];
    if (tags.some(t => grantedExpanded.has(t))) {
      visible.add(d.doc_id);
    }
    // 規則 4：無標籤文件不加入（deny by default）
  });
  return visible;
}

// 單文件可見性檢查（apiGetDocHistory、祖先/子孫查詢入口用）。
// 不可見時 throw，且訊息不洩漏文件是否存在。
function _assertCanViewDoc(docId) {
  const ctx = _assertWhitelisted();
  const visible = _getVisibleDocIds(ctx);
  if (!visible.has(docId)) {
    throw new Error('無存取權限：找不到文件或您沒有檢視此文件的權限');
  }
  return ctx;
}

// ── 管理員工具 ────────────────────────────────────────────────

// 清除 HR 名單快取（HR 主表更新後可手動執行，或等 10 分鐘自動過期）
function clearHrCache() {
  CacheService.getScriptCache().remove(HR_CACHE_KEY);
  Logger.log('✅ HR 名單快取已清除');
}

// 新增 HR 試算表讀取 scope 後，在 GAS 編輯器手動執行一次以觸發授權
function authorizeOnce() {
  const email = Session.getActiveUser().getEmail();
  const hrId = _getProp(PROP_KEYS.HR_SPREADSHEET_ID);
  if (!hrId) {
    Logger.log('⚠️ 請先在 Script Properties 設定 HR_SPREADSHEET_ID 再執行');
    return;
  }
  const name = SpreadsheetApp.openById(hrId).getName();
  Logger.log(`✅ 授權完成。目前帳號：${email}，可讀取 HR 試算表：${name}`);
}

// 診斷函式：核對 HR 主表實際表頭與白名單筆數
// （依 ECOSYSTEM.md 要求：動跨專案欄位前先核對實表結構）
function debugGetSystemData() {
  const hrId = _getProp(PROP_KEYS.HR_SPREADSHEET_ID);
  Logger.log(`HR_SPREADSHEET_ID 已設定：${hrId ? '是' : '否'}`);
  Logger.log(`ADMIN_EMAILS：${_getAdminEmails().join(', ') || '（未設定）'}`);

  if (!hrId) return;

  const sheet = SpreadsheetApp.openById(hrId).getSheetByName(HR_SHEET_NAME);
  if (!sheet) {
    Logger.log(`⚠️ HR 主表找不到工作表：${HR_SHEET_NAME}`);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, 5).getDisplayValues()[0];
  Logger.log(`人員主檔表頭（前 5 欄）：${JSON.stringify(headers)}`);

  clearHrCache();
  const people = _getHrPeople();
  Logger.log(`白名單筆數（排除 ${EXCLUDED_HR_STATUS.join('/')}）：${people.length}`);
  Logger.log(`目前登入者情境：${JSON.stringify(getUserContext())}`);
}

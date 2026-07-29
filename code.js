// ============================================================
// code.js — 核心業務邏輯
// 文件 CRUD + Closure Table 關聯維護 + 前端 API
// 權限檢查見 auth.js、異動紀錄見 audit.js
// ============================================================

// ── Web App 進入點（白名單閘門）──────────────────────────────
function doGet() {
  let ctx;
  try {
    ctx = getUserContext();
  } catch (e) {
    // HR_SPREADSHEET_ID 未設定等組態錯誤 → 顯示設定指引而非白屏
    return _renderMessagePage('SPREADSHEET_ID尚未完成設定', String(e.message || e) +
      '<br><br>請管理員於「專案設定 → 指令碼資訊」設定 Script Properties。');
  }

  if (!ctx.isWhitelisted) {
    return _renderMessagePage('無存取權限',
      `您的登入信箱 <b>${ctx.email || '（無法取得）'}</b> 不在人員名單中。<br>` +
      '如需使用本功能，請聯絡管理員。');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('文件管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _renderMessagePage(title, bodyHtml) {
  const html = `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">
    <style>body{font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif;
    display:flex;align-items:center;justify-content:center;min-height:90vh;background:#f5f6f8;color:#1f2937}
    .box{background:#fff;border:1px solid #dde2e8;border-radius:8px;padding:32px 40px;max-width:480px;text-align:center}
    h1{font-size:18px;margin-bottom:12px}p{font-size:14px;line-height:1.7;color:#6b7280}</style></head>
    <body><div class="box"><h1>${title}</h1><p>${bodyHtml}</p></div></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('文件管理');
}

// ── 工具函式 ──────────────────────────────────────────────────

function _getSheet(name) {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`找不到工作表：${name}`);
  return sheet;
}

// 讀取文件清單（過濾空列，回傳物件陣列）
function _readDocs() {
  const sheet = _getSheet(SHEET_NAMES.DOCS);
  const rows = sheet.getDataRange().getDisplayValues(); // 字串格式，避免轉型
  return rows.slice(1)
    .filter(r => r[DOC_COL.DOC_ID])
    .map(r => ({
      doc_id:             r[DOC_COL.DOC_ID],
      title:              r[DOC_COL.TITLE],
      category:           r[DOC_COL.CATEGORY],
      status:             r[DOC_COL.STATUS],
      owner:              r[DOC_COL.OWNER],
      updated_at:         r[DOC_COL.UPDATED_AT],
      version:            r[DOC_COL.VERSION],
      drive_loc:          r[DOC_COL.GOOGLE_DRIVE_LOC],
      owner_email:        r[DOC_COL.OWNER_EMAIL] || '',
      published_at:       r[DOC_COL.PUBLISHED_AT] || '',
      next_review:        r[DOC_COL.NEXT_REVIEW] || '',
      review_cycle:       r[DOC_COL.REVIEW_CYCLE] || '',
      file_id:            r[DOC_COL.FILE_ID] || '',
      pending_file_id:    r[DOC_COL.PENDING_FILE_ID] || '',
      pending_version:    r[DOC_COL.PENDING_VERSION] || '',
      pending_file_name:  r[DOC_COL.PENDING_FILE_NAME] || '',
      security_level:     r[DOC_COL.SECURITY_LEVEL] || '一般',
    }));
}

// 讀取閉包表（過濾空列）
function _readClosure() {
  const sheet = _getSheet(SHEET_NAMES.CLOSURE);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1)
    .filter(r => r[CLS_COL.ANCESTOR_ID])
    .map(r => ({
      ancestor_id:   r[CLS_COL.ANCESTOR_ID],
      descendant_id: r[CLS_COL.DESCENDANT_ID],
      depth:         parseInt(r[CLS_COL.DEPTH], 10),
      relation_type: r[CLS_COL.RELATION_TYPE],
      description:   r[CLS_COL.DESCRIPTION],
    }));
}

function _now() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
}

// 今天 + N 個月，回傳 yyyy/MM/dd 字串（下次審查日計算用）
function _addMonthsFromToday(months) {
  const d = new Date();
  d.setMonth(d.getMonth() + months);
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd');
}

// 把文件物件轉成完整的一列（apiCreateDoc / apiUpdateDoc 共用）
function _docToRow(doc) {
  const row = new Array(DOC_COL_COUNT).fill('');
  row[DOC_COL.DOC_ID]           = doc.doc_id;
  row[DOC_COL.TITLE]            = doc.title || '';
  row[DOC_COL.CATEGORY]         = doc.category || '';
  row[DOC_COL.STATUS]           = doc.status || '草稿';
  row[DOC_COL.OWNER]            = doc.owner || '';
  row[DOC_COL.UPDATED_AT]       = _now();
  row[DOC_COL.VERSION]          = doc.version || '0.1';
  row[DOC_COL.GOOGLE_DRIVE_LOC] = doc.drive_loc || '';
  row[DOC_COL.OWNER_EMAIL]      = doc.owner_email || '';
  row[DOC_COL.PUBLISHED_AT]     = doc.published_at || '';
  row[DOC_COL.NEXT_REVIEW]      = doc.next_review || '';
  row[DOC_COL.REVIEW_CYCLE]     = doc.review_cycle || '';
  row[DOC_COL.FILE_ID]          = doc.file_id || '';
  row[DOC_COL.PENDING_FILE_ID]  = doc.pending_file_id || '';
  row[DOC_COL.PENDING_VERSION]  = doc.pending_version || '';
  row[DOC_COL.PENDING_FILE_NAME] = doc.pending_file_name || '';
  row[DOC_COL.SECURITY_LEVEL]   = doc.security_level || '一般';
  return row;
}

// 讀取「檔案版本」（過濾空列，新→舊排序由呼叫端決定）
function _readFileVersions() {
  const sheet = _getSheet(SHEET_NAMES.FILE_VERSIONS);
  const rows = sheet.getDataRange().getDisplayValues();
  return rows.slice(1)
    .filter(r => r[FILEVER_COL.DOC_ID])
    .map(r => ({
      doc_id:      r[FILEVER_COL.DOC_ID],
      version:     r[FILEVER_COL.VERSION],
      file_id:     r[FILEVER_COL.FILE_ID],
      file_name:   r[FILEVER_COL.FILE_NAME],
      uploaded_by: r[FILEVER_COL.UPLOADED_BY],
      uploaded_at: r[FILEVER_COL.UPLOADED_AT],
    }));
}

// 去除路徑分隔字元，避免檔名污染 Drive 命名或造成混淆
function _sanitizeFileName(name) {
  return String(name || '').replace(/[\/\\]/g, '_').trim().slice(0, 200) || 'untitled';
}

// 版號遞增：X.Y 格式解析；minor → X.(Y+1)，major → (X+1).0。
// 無法解析（自由文字舊資料）視為 0.0 起算。
function _bumpVersion(current, bumpType) {
  if (!VERSION_BUMP_TYPES.includes(bumpType)) {
    throw new Error(`不支援的改版類型：${bumpType}（允許：${VERSION_BUMP_TYPES.join('、')}）`);
  }
  const m = String(current || '').match(/^(\d+)\.(\d+)$/);
  const major = m ? parseInt(m[1], 10) : 0;
  const minor = m ? parseInt(m[2], 10) : 0;
  return bumpType === 'major' ? `${major + 1}.0` : `${major}.${minor + 1}`;
}

// 取得（或建立）中央資料夾下該文件的子資料夾
function _getOrCreateDocFolder(docId) {
  const rootId = _getProp(PROP_KEYS.DOC_FILES_FOLDER_ID);
  if (!rootId) throw new Error('尚未設定文件檔案庫，請管理員執行 migrateV7()（或 deployAllSheets()）');
  const root = DriveApp.getFolderById(rootId);
  const existing = root.getFoldersByName(docId);
  return existing.hasNext() ? existing.next() : root.createFolder(docId);
}

// ============================================================
// 前端 API：文件 CRUD
// ============================================================

// 取得所有文件（含選項清單與使用者情境，供前端初始化一次取完）
// docs 僅回傳可見文件（透過 _getVisibleDocIds 過濾，防資訊洩漏）。
function apiGetInitData() {
  // 白名單閘門：與其他讀取 API 一致。避免已離職（不在 HR 名單）但仍掛在
  // owner_email 的使用者，繞過 doGet 頁面直接呼叫本 API 取回文件與標籤全樹。
  const ctx = _assertWhitelisted();
  const visible = _getVisibleDocIds(ctx);
  const docs = _readDocs().filter(d => visible.has(d.doc_id));

  // 標籤全樹（資料夾樹需要名稱；僅 id/name/parent/sort，無機密）
  const tags = _readTags().map(t => ({
    tag_id: t.tag_id, name: t.name, parent_id: t.parent_id, sort: t.sort,
  }));
  // 文件標籤：僅回傳可見文件的貼標，避免經由標籤推知不可見文件存在
  const docTags = _readDocTags().filter(dt => visible.has(dt.doc_id));
  // 當前使用者的有效授權標籤（個人∪群組，read 含 edit；前端資料夾樹用）
  const grantedTagIds = Array.from(_getEffectiveGrantedTagIds(ctx).read);
  // 可編輯文件集（V5）：前端據此顯示編輯入口；每個寫入 API 仍逐一重新斷言（IDOR 防護）
  const editable = _getEditableDocIds(ctx);

  return {
    docs: docs,
    statuses: DOC_STATUS,
    categories: DOC_CATEGORIES,
    securityLevels: SECURITY_LEVELS,
    relationTypes: RELATION_TYPES,
    reviewCycles: REVIEW_CYCLES,
    statusTransitions: STATUS_TRANSITIONS,
    user: Object.assign({}, ctx, { grantedTagIds: grantedTagIds }),
    // 負責人下拉選項：白名單即可看到（僅姓名與信箱，無其他個資）
    hrPeople: ctx.isWhitelisted ? _getHrPeople() : [],
    tags: tags,
    docTags: docTags,
    editableDocIds: docs.filter(d => editable.has(d.doc_id)).map(d => d.doc_id),
  };
}

// 新增文件（自動產生 doc_id + 寫入 self closure 記錄）
// 白名單使用者皆可建立；未指定負責人信箱時預設為建立者本人。
function apiCreateDoc(doc) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ctx = _assertWhitelisted();
    const sheet = _getSheet(SHEET_NAMES.DOCS);

    // 產生新 doc_id：掃描現有最大序號
    const existing = _readDocs().map(d => d.doc_id);
    let maxNum = 0;
    existing.forEach(id => {
      const m = id.match(/^DOC-(\d+)$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
    const newId = 'DOC-' + String(maxNum + 1).padStart(3, '0');

    const newDoc = Object.assign({}, doc, {
      doc_id: newId,
      status: doc.status || '草稿',
      owner_email: doc.owner_email || ctx.email,
      review_cycle: doc.review_cycle || DEFAULT_REVIEW_CYCLE,
      published_at: '',
      next_review: '',
    });
    sheet.appendRow(_docToRow(newDoc));

    // Closure Table：寫入自身記錄（depth=0）
    const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
    clsSheet.appendRow([newId, newId, 0, 'references', '自身']);

    // 可選：建立時同時貼標（僅接受既有標籤）
    const wantTags = (doc.tagIds || []).map(String).filter(Boolean);
    if (wantTags.length > 0) {
      const validTagIds = new Set(_readTags().map(t => t.tag_id));
      const rows = wantTags.filter(t => validTagIds.has(t)).map(t => [newId, t]);
      if (rows.length > 0) {
        const dtSheet = _getSheet(SHEET_NAMES.DOC_TAGS);
        dtSheet.getRange(dtSheet.getLastRow() + 1, 1, rows.length, DOCTAG_COL_COUNT)
          .setValues(rows);
      }
    }

    _logAudit('建立', newId, newDoc.version || '0.1', `建立文件「${newDoc.title || ''}」`);
    SpreadsheetApp.flush();
    return { success: true, doc_id: newId };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 上傳新版檔案（V6）：寫入 pending 欄並將文件轉入「審核中」，
// 待管理員核准發布時（apiUpdateDoc promote 邏輯）才正式生效。
// bumpType：'minor' 或 'major'，決定 pending_version 怎麼算。
function apiUploadDocFile(docId, fileName, base64, mimeType, bumpType, customVersion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ctx = _assertCanEditDoc(docId);
    const oldDoc = _readDocs().find(d => d.doc_id === docId);
    if (!oldDoc) return { success: false, error: '找不到文件：' + docId };
    if (oldDoc.status === '已廢止') {
      return { success: false, error: '已廢止文件不可上傳新版檔案' };
    }
    if (!base64) return { success: false, error: '未提供檔案內容' };

    const safeName = _sanitizeFileName(fileName);
    let pendingVersion;
    if (bumpType === 'start') {
      // 管理員自訂起始版號：僅文件首次上傳（尚無現行版或待核版）可用，
      // 不經過 _bumpVersion，直接採信通過格式驗證的字串。
      if (!ctx.isAdmin) {
        return { success: false, error: '權限不足：僅管理員可設定自訂起始版號' };
      }
      if (oldDoc.file_id || oldDoc.pending_file_id) {
        return { success: false, error: '此文件已有檔案，僅能使用小改版／大改版' };
      }
      const v = String(customVersion || '').trim();
      if (!/^\d+\.\d+$/.test(v)) {
        return { success: false, error: '起始版號格式錯誤，需為「數字.數字」，例如 3.2' };
      }
      pendingVersion = v;
    } else {
      pendingVersion = _bumpVersion(oldDoc.version, bumpType);
    }

    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream',
      `${docId}_v${pendingVersion}_${safeName}`);
    const folder = _getOrCreateDocFolder(docId);
    const file = folder.createFile(blob);

    // 若已有待核檔案（重傳），舊的丟垃圾桶避免資料夾堆積
    if (oldDoc.pending_file_id) {
      try { DriveApp.getFileById(oldDoc.pending_file_id).setTrashed(true); } catch (e) { /* 已不存在則略過 */ }
    }

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
    if (idx < 1) return { success: false, error: '找不到文件：' + docId };

    sheet.getRange(idx + 1, DOC_COL.PENDING_FILE_ID + 1, 1, 3)
      .setValues([[file.getId(), pendingVersion, safeName]]);

    let newStatus = oldDoc.status;
    if (oldDoc.status !== '審核中') {
      const allowed = STATUS_TRANSITIONS[oldDoc.status] || [];
      if (allowed.includes('審核中')) {
        sheet.getRange(idx + 1, DOC_COL.STATUS + 1).setValue('審核中');
        newStatus = '審核中';
      }
    }

    const bumpLabel = bumpType === 'major' ? '大改版' : bumpType === 'start' ? '自訂起始版號' : '小改版';
    _logAudit('上傳檔案', docId, pendingVersion,
      `上傳待核檔案「${safeName}」（${bumpLabel}，待核准後生效為 v${pendingVersion}）`);

    SpreadsheetApp.flush();
    return { success: true, pending_version: pendingVersion, pending_file_name: safeName, status: newStatus };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 管理員批次匯入直接生效之單檔處理 API
// 不經過 _bumpVersion 與待核欄位，直接寫入 file_id / version / published_at，並將狀態設定為「已發布」
function apiBatchImportDirectFile(docId, fileName, base64, mimeType, targetVersion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ctx = _assertAdmin();
    const oldDoc = _readDocs().find(d => d.doc_id === docId);
    if (!oldDoc) return { success: false, error: '找不到文件：' + docId };
    if (oldDoc.status === '已廢止') {
      return { success: false, error: '已廢止文件不可批次匯入檔案' };
    }
    if (!base64) return { success: false, error: '未提供檔案內容' };

    const safeName = _sanitizeFileName(fileName);
    const ver = String(targetVersion || oldDoc.version || '1.0').trim();

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
    if (idx < 1) return { success: false, error: '找不到試算表列：' + docId };

    if (oldDoc.pending_file_id) {
      try { DriveApp.getFileById(oldDoc.pending_file_id).setTrashed(true); } catch (e) { /* 忽略已不存在的情況 */ }
    }

    // 解碼 Base64 並存入 Google Drive
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream',
      `${docId}_v${ver}_${safeName}`);
    const folder = _getOrCreateDocFolder(docId);
    const file = folder.createFile(blob);

    // 準備新版資料，以 oldDoc 為基礎覆寫
    const newStatus = '已發布';
    const cycle = parseInt(oldDoc.review_cycle, 10) || DEFAULT_REVIEW_CYCLE;
    const pubDate = (oldDoc.status !== '已發布' || !oldDoc.published_at) ? _now() : oldDoc.published_at;
    const nextRev = (oldDoc.status !== '已發布' || !oldDoc.next_review) ? _addMonthsFromToday(cycle) : oldDoc.next_review;

    const merged = Object.assign({}, oldDoc, {
      status: newStatus,
      version: ver,
      published_at: pubDate,
      next_review: nextRev,
      file_id: file.getId(),
      pending_file_id: '',
      pending_version: '',
      pending_file_name: '',
    });

    sheet.getRange(idx + 1, 1, 1, DOC_COL_COUNT).setValues([_docToRow(merged)]);

    // 登記至「檔案版本」表
    const fileVerSheet = _getSheet(SHEET_NAMES.FILE_VERSIONS);
    fileVerSheet.appendRow([docId, ver, file.getId(), safeName, _getCurrentEmail(), _nowWithTime()]);

    // 記錄異動審計
    _logAudit('批次匯入', docId, ver, `管理員批次初始化/直接合位正式檔案「${safeName}」（生效為 v${ver}）`);

    SpreadsheetApp.flush();
    return { success: true, docId: docId, version: ver, status: newStatus, file_id: file.getId() };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 更新文件（依 doc_id 定位列，整列寫回）
// 狀態變更走 STATUS_TRANSITIONS 查表驗證；轉「已發布」自動填發布日與下次審查日。
function apiUpdateDoc(doc) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ctx = _assertCanEditDoc(doc.doc_id);

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === doc.doc_id);
    if (idx < 1) return { success: false, error: '找不到文件：' + doc.doc_id };

    const oldDoc = _readDocs().find(d => d.doc_id === doc.doc_id);

    // V5：更換負責人屬權限管理行為，僅管理員或現任負責人可為之
    // （edit 授權者送回的 owner_email 未變動時不觸發；含清空也算變動，fail-closed）
    const newOwnerEmail = String(doc.owner_email || '').trim().toLowerCase();
    const oldOwnerEmail = String(oldDoc.owner_email || '').trim().toLowerCase();
    if (newOwnerEmail !== oldOwnerEmail) {
      _assertOwnerOrAdmin(doc.doc_id);
    }

    const oldStatus = oldDoc.status;
    const newStatus = doc.status || oldStatus;

    // 狀態流轉驗證（Map 查表；同狀態不需檢查）
    if (newStatus !== oldStatus) {
      const allowed = STATUS_TRANSITIONS[oldStatus] || [];
      if (!allowed.includes(newStatus)) {
        return { success: false, error: `狀態不可由「${oldStatus}」轉為「${newStatus}」（允許：${allowed.join('、') || '無'}）` };
      }
      // 廢止文件復活僅限管理員
      if (oldStatus === '已廢止' && !ctx.isAdmin) {
        return { success: false, error: '「已廢止」文件僅限管理員恢復為草稿' };
      }
    }

    // 發布日／下次審查日：預設沿用舊值；轉「已發布」時重算
    // V6：version / file_id / pending_* / drive_loc 一律沿用 oldDoc，前端送來的值一律忽略——
    // 唯一改動 version/file_id/pending_* 的路徑是下方「核准發布時 promote」與 apiUploadDocFile；
    // drive_loc（V6 起表單已移除該欄位，前端不再送出）純粹沿用舊值，避免編輯時被清空。
    const merged = Object.assign({}, doc, {
      published_at: oldDoc.published_at,
      next_review:  oldDoc.next_review,
      status: newStatus,
      version: oldDoc.version,
      drive_loc: oldDoc.drive_loc,
      file_id: oldDoc.file_id,
      pending_file_id: oldDoc.pending_file_id,
      pending_version: oldDoc.pending_version,
      pending_file_name: oldDoc.pending_file_name,
    });
    if (newStatus === '已發布' && oldStatus !== '已發布') {
      const cycle = parseInt(doc.review_cycle, 10) || DEFAULT_REVIEW_CYCLE;
      merged.published_at = _now();
      merged.next_review  = _addMonthsFromToday(cycle);
    }

    // V6：核准發布時，若有待核檔案，promote 為正式版並登記檔案版本表
    let promotedVersion = '';
    if (newStatus === '已發布' && oldStatus !== '已發布' && oldDoc.pending_file_id) {
      promotedVersion = oldDoc.pending_version;
      merged.version = oldDoc.pending_version;
      merged.file_id = oldDoc.pending_file_id;
      merged.pending_file_id = '';
      merged.pending_version = '';
      merged.pending_file_name = '';
    }

    sheet.getRange(idx + 1, 1, 1, DOC_COL_COUNT).setValues([_docToRow(merged)]);

    if (promotedVersion) {
      const fileVerSheet = _getSheet(SHEET_NAMES.FILE_VERSIONS);
      fileVerSheet.appendRow([doc.doc_id, promotedVersion, merged.file_id, oldDoc.pending_file_name, _getCurrentEmail(), _nowWithTime()]);
    }

    const action = (newStatus !== oldStatus) ? '狀態變更' : '更新';
    const summary = (newStatus !== oldStatus)
      ? `${oldStatus} → ${newStatus}` +
        (merged.published_at !== oldDoc.published_at ? `（發布日 ${merged.published_at}，下次審查 ${merged.next_review}）` : '') +
        (promotedVersion ? `｜檔案 v${promotedVersion} 生效` : '')
      : (_diffSummary(oldDoc, merged) || '（無欄位變更）');
    _logAudit(action, doc.doc_id, merged.version, summary);

    SpreadsheetApp.flush();
    return { success: true, published_at: merged.published_at, next_review: merged.next_review, version: merged.version };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 批次核准文件（審核中 → 已發布）：管理員專用。逐筆獨立判斷可否核准，
// 成功的照常生效（含 promote 待核檔案為正式版），失敗的記原因、不影響其他筆。
// 與 apiUpdateDoc 核准發布分支邏輯重複但刻意不抽共用——apiUpdateDoc 承擔任意
// 欄位更新與多種狀態轉移已經夠複雜，這裡只做單一轉移，各自獨立更清楚。
function apiBatchApproveDocuments(docIds) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();
    const ids = Array.from(new Set((docIds || []).map(String).filter(Boolean)));
    if (ids.length === 0) return { success: true, approved: [], failed: [] };

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const docs = _readDocs();
    const fileVerSheet = _getSheet(SHEET_NAMES.FILE_VERSIONS);

    const approved = [];
    const failed = [];

    ids.forEach(docId => {
      const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
      const oldDoc = docs.find(d => d.doc_id === docId);
      if (idx < 1 || !oldDoc) {
        failed.push({ doc_id: docId, error: '找不到文件' });
        return;
      }
      if (oldDoc.status !== '審核中') {
        failed.push({ doc_id: docId, error: `狀態為「${oldDoc.status}」，非審核中，無法核准` });
        return;
      }

      const cycle = parseInt(oldDoc.review_cycle, 10) || DEFAULT_REVIEW_CYCLE;
      const merged = Object.assign({}, oldDoc, {
        status: '已發布',
        published_at: _now(),
        next_review: _addMonthsFromToday(cycle),
      });

      let promotedVersion = '';
      if (oldDoc.pending_file_id) {
        promotedVersion = oldDoc.pending_version;
        merged.version = oldDoc.pending_version;
        merged.file_id = oldDoc.pending_file_id;
        merged.pending_file_id = '';
        merged.pending_version = '';
        merged.pending_file_name = '';
      }

      sheet.getRange(idx + 1, 1, 1, DOC_COL_COUNT).setValues([_docToRow(merged)]);

      if (promotedVersion) {
        fileVerSheet.appendRow([docId, promotedVersion, merged.file_id, oldDoc.pending_file_name, _getCurrentEmail(), _nowWithTime()]);
      }

      const summary = `審核中 → 已發布（發布日 ${merged.published_at}，下次審查 ${merged.next_review}）` +
        (promotedVersion ? `｜檔案 v${promotedVersion} 生效` : '') + '（批次核准）';
      _logAudit('狀態變更', docId, merged.version, summary);

      approved.push(docId);
    });

    SpreadsheetApp.flush();
    return { success: true, approved: approved, failed: failed };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 代理下載文件檔案（V6）：檔案不設任何 Drive 共用，一律經此 API
// 依現有可見性／可編輯性權限回傳內容，避免繞過標籤權限直接用連結存取。
// fileId 必須屬於該文件（現行 file_id、pending_file_id，或其歷史版本），
// 否則視為越權存取拒絕（IDOR 防護）；pending 檔額外要求編輯權（審核者為 admin，天然通過）。
function apiDownloadDocFile(docId, fileId) {
  const ctx = _assertCanViewDoc(docId);
  const doc = _readDocs().find(d => d.doc_id === docId);
  if (!doc) throw new Error('找不到文件：' + docId);

  const historyIds = _readFileVersions().filter(v => v.doc_id === docId).map(v => v.file_id);
  const validIds = new Set([doc.file_id, doc.pending_file_id, ...historyIds].filter(Boolean));
  if (!validIds.has(fileId)) {
    throw new Error('無效的檔案：此檔案不屬於指定文件');
  }
  if (fileId === doc.pending_file_id) {
    _assertCanEditDoc(docId); // 待核檔案僅編輯權者（含審核管理員）可下載
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  return {
    success: true,
    fileName: file.getName(),
    mimeType: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes()),
  };
}

// 取得文件的檔案版本歷史（V6，新→舊排序）
function apiGetDocFileVersions(docId) {
  _assertCanViewDoc(docId);
  return _readFileVersions()
    .filter(v => v.doc_id === docId)
    .reverse();
}

// 刪除文件（同步刪除閉包表中所有相關記錄）——不可逆，僅限管理員
function apiDeleteDoc(docId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();

    // 1. 刪除文件清單中的列
    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
    if (idx < 1) return { success: false, error: '找不到文件：' + docId };
    const title = rows[idx][DOC_COL.TITLE];
    sheet.deleteRow(idx + 1);

    // 2. 刪除閉包表中所有 ancestor 或 descendant 為此文件的列（由下往上刪）
    const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
    const clsRows = clsSheet.getDataRange().getDisplayValues();
    for (let i = clsRows.length - 1; i >= 1; i--) {
      if (clsRows[i][CLS_COL.ANCESTOR_ID] === docId ||
          clsRows[i][CLS_COL.DESCENDANT_ID] === docId) {
        clsSheet.deleteRow(i + 1);
      }
    }

    // 3. 級聯清除「文件標籤」中此文件的貼標列（由下往上刪）
    const dtSheet = _getSheet(SHEET_NAMES.DOC_TAGS);
    const dtRows = dtSheet.getDataRange().getDisplayValues();
    for (let i = dtRows.length - 1; i >= 1; i--) {
      if (dtRows[i][DOCTAG_COL.DOC_ID] === docId) {
        dtSheet.deleteRow(i + 1);
      }
    }

    // 4. 級聯清除「檔案版本」中此文件的紀錄列（由下往上刪）——避免 doc_id
    // 序號重用時，新文件繼承到舊文件不相干的版本歷史（V6 起才有此表，
    // 尚未執行 migrateV7()/deployAllSheets() 的舊試算表可能還沒有，安全跳過）。
    const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
    const fvSheet = ss.getSheetByName(SHEET_NAMES.FILE_VERSIONS);
    if (fvSheet) {
      const fvRows = fvSheet.getDataRange().getDisplayValues();
      for (let i = fvRows.length - 1; i >= 1; i--) {
        if (fvRows[i][FILEVER_COL.DOC_ID] === docId) {
          fvSheet.deleteRow(i + 1);
        }
      }
    }

    _logAudit('刪除', docId, '', `刪除文件「${title}」（含所有關聯與標籤記錄）`);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 前端 API：Closure Table 關聯操作
// ============================================================

// 查詢某文件的所有關聯（向下：我引用了誰）
// 入口先驗證可見性，結果再過濾掉不可見的子孫（完全隱藏）。
function apiGetDescendants(docId, maxDepth) {
  const ctx = _assertCanViewDoc(docId);
  const visible = _getVisibleDocIds(ctx);
  const closure = _readClosure();
  const docsMap = new Map(_readDocs().map(d => [d.doc_id, d]));
  const limit = maxDepth || 99;

  return closure
    .filter(c => c.ancestor_id === docId && c.depth > 0 && c.depth <= limit &&
                 visible.has(c.descendant_id))
    .map(c => {
      const doc = docsMap.get(c.descendant_id);
      return {
        doc_id: c.descendant_id,
        depth: c.depth,
        relation_type: c.relation_type,
        description: c.description,
        title:  doc ? doc.title : '（文件不存在）',
        status: doc ? doc.status : '',
        owner:  doc ? doc.owner : '',
      };
    })
    .sort((a, b) => a.depth - b.depth || a.doc_id.localeCompare(b.doc_id));
}

// 反查某文件被誰關聯（向上：誰引用了我）
// 入口先驗證可見性，結果再過濾掉不可見的祖先（完全隱藏）。
function apiGetAncestors(docId) {
  const ctx = _assertCanViewDoc(docId);
  const visible = _getVisibleDocIds(ctx);
  const closure = _readClosure();
  const docsMap = new Map(_readDocs().map(d => [d.doc_id, d]));

  return closure
    .filter(c => c.descendant_id === docId && c.depth > 0 &&
                 visible.has(c.ancestor_id))
    .map(c => {
      const doc = docsMap.get(c.ancestor_id);
      return {
        doc_id: c.ancestor_id,
        depth: c.depth,
        relation_type: c.relation_type,
        title:  doc ? doc.title : '（文件不存在）',
        status: doc ? doc.status : '',
      };
    })
    .sort((a, b) => a.depth - b.depth);
}

// 新增關聯（核心：自動維護閉包表的間接路徑）
// 建立 A → B 時：
//   1. 檢查循環（B 的子孫不可包含 A）
//   2. 寫入 A → B (depth=1)
//   3. 對 A 的每個祖先 X (depth=dX)：寫入 X → B (depth=dX+1)
//   4. 對 B 的每個子孫 Y (depth=dY)：寫入 A → Y (depth=dY+1)
//   5. 交叉：X → Y (depth = dX + 1 + dY)
// 另：relation_type = supersedes 且目標為「已發布」→ 自動將目標標為「已廢止」
function apiAddRelation(ancestorId, descendantId, relationType, description) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertCanEditDoc(ancestorId);

    if (ancestorId === descendantId) {
      return { success: false, error: '不可關聯自身' };
    }

    // 目標文件也必須可見，否則會形成「存在性 oracle」：回傳值可反推
    // 不可見文件是否存在／是否已發布。
    // V5 收緊：supersedes 會自動廢止對方已發布文件，屬對後代端的實質寫入，
    // 需具備後代端的「編輯權」（edit 蘊含可見，檢查已涵蓋可見性）。
    if (relationType === 'supersedes') {
      _assertCanEditDoc(descendantId);
    } else {
      _assertCanViewDoc(descendantId);
    }

    const closure = _readClosure();

    // 重複檢查
    const exists = closure.some(c =>
      c.ancestor_id === ancestorId && c.descendant_id === descendantId && c.depth === 1);
    if (exists) return { success: false, error: '此關聯已存在' };

    // 循環檢查：若 descendant 的子孫鏈中已包含 ancestor，會形成循環
    const wouldCycle = closure.some(c =>
      c.ancestor_id === descendantId && c.descendant_id === ancestorId && c.depth >= 0);
    if (wouldCycle) {
      return { success: false, error: `無法建立關聯：${descendantId} 已（直接或間接）關聯到 ${ancestorId}，會形成循環` };
    }

    // 取得 A 的所有祖先（含 A 自身 depth=0 視為起點）
    const ancestorsOfA = closure.filter(c =>
      c.descendant_id === ancestorId && c.depth >= 0);
    // 取得 B 的所有子孫（含 B 自身 depth=0）
    const descendantsOfB = closure.filter(c =>
      c.ancestor_id === descendantId && c.depth >= 0);

    // 笛卡兒積：每個 (X → A) × (B → Y) 產生 X → Y, depth = dX + 1 + dY
    const newRows = [];
    const desc = description || '';
    ancestorsOfA.forEach(xa => {
      descendantsOfB.forEach(by => {
        const newDepth = xa.depth + 1 + by.depth;
        // 跳過已存在的相同路徑
        const dup = closure.some(c =>
          c.ancestor_id === xa.ancestor_id &&
          c.descendant_id === by.descendant_id &&
          c.depth === newDepth);
        if (!dup) {
          const isDirect = (newDepth === 1);
          newRows.push([
            xa.ancestor_id,
            by.descendant_id,
            newDepth,
            relationType || 'references',
            isDirect ? desc : `間接關聯（透過 ${ancestorId}→${descendantId}）`,
          ]);
        }
      });
    });

    if (newRows.length > 0) {
      const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
      clsSheet.getRange(clsSheet.getLastRow() + 1, 1, newRows.length, 5)
        .setValues(newRows);
    }

    _logAudit('關聯新增', ancestorId, '',
      `${ancestorId} → ${descendantId}（${relationType || 'references'}）${desc ? '：' + desc : ''}`);

    // supersedes 自動廢止：新版文件取代已發布的舊版 → 舊版標「已廢止」
    let deprecated = null;
    if (relationType === 'supersedes') {
      const target = _readDocs().find(d => d.doc_id === descendantId);
      if (target && target.status === '已發布') {
        const sheet = _getSheet(SHEET_NAMES.DOCS);
        const rows = sheet.getDataRange().getDisplayValues();
        const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === descendantId);
        if (idx >= 1) {
          sheet.getRange(idx + 1, DOC_COL.STATUS + 1).setValue('已廢止');
          sheet.getRange(idx + 1, DOC_COL.UPDATED_AT + 1).setValue(_now());
          _logAudit('狀態變更', descendantId, target.version,
            `已發布 → 已廢止（被 ${ancestorId} supersedes 自動廢止）`);
          deprecated = descendantId;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, added: newRows.length, deprecated: deprecated };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 刪除直接關聯（同步清除因此關聯產生的間接路徑）
// 刪除 A → B (depth=1) 時，需重建受影響的路徑：
// 簡化策略：刪除所有「經過 A→B」的路徑後重算
function apiRemoveRelation(ancestorId, descendantId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertCanEditDoc(ancestorId);

    const closure = _readClosure();

    // 確認直接關聯存在
    const direct = closure.find(c =>
      c.ancestor_id === ancestorId && c.descendant_id === descendantId && c.depth === 1);
    if (!direct) return { success: false, error: '找不到此直接關聯' };

    // 找出可能受影響的路徑：
    // X → Y 其中 X ∈ {A 的祖先 ∪ A}，Y ∈ {B 的子孫 ∪ B}
    const ancestorsOfA = new Set(
      closure.filter(c => c.descendant_id === ancestorId && c.depth >= 0)
        .map(c => c.ancestor_id));
    const descendantsOfB = new Set(
      closure.filter(c => c.ancestor_id === descendantId && c.depth >= 0)
        .map(c => c.descendant_id));

    // 重建策略：刪除所有疑似受影響的列，再從剩餘 depth=1 邊重算閉包
    const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
    const allRows = clsSheet.getDataRange().getDisplayValues();

    // 由下往上刪除受影響的路徑（depth > 0 且落在影響範圍）
    for (let i = allRows.length - 1; i >= 1; i--) {
      const r = allRows[i];
      const anc = r[CLS_COL.ANCESTOR_ID];
      const des = r[CLS_COL.DESCENDANT_ID];
      const dep = parseInt(r[CLS_COL.DEPTH], 10);
      if (dep > 0 && ancestorsOfA.has(anc) && descendantsOfB.has(des)) {
        clsSheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();

    // 從剩餘的 depth=1 邊重算受影響範圍的閉包
    _rebuildClosurePaths(Array.from(ancestorsOfA), Array.from(descendantsOfB));

    _logAudit('關聯移除', ancestorId, '', `移除 ${ancestorId} → ${descendantId}（間接路徑已重算）`);

    return { success: true };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 從 depth=1 邊重算指定範圍的間接路徑（BFS）
function _rebuildClosurePaths(affectedAncestors, affectedDescendants) {
  const closure = _readClosure();

  // 建立鄰接表（只用 depth=1 直接邊）
  const adj = new Map();
  closure.filter(c => c.depth === 1).forEach(c => {
    if (!adj.has(c.ancestor_id)) adj.set(c.ancestor_id, []);
    adj.get(c.ancestor_id).push(c.descendant_id);
  });

  // 已存在的路徑集合（避免重複寫入）
  const existing = new Set(closure.map(c =>
    `${c.ancestor_id}|${c.descendant_id}|${c.depth}`));

  const newRows = [];

  // 對每個受影響的 ancestor 做 BFS，找出深度 >= 2 的可達節點
  affectedAncestors.forEach(start => {
    const queue = [[start, 0]];
    const visited = new Map([[start, 0]]);

    while (queue.length > 0) {
      const [node, depth] = queue.shift();
      const neighbors = adj.get(node) || [];
      neighbors.forEach(next => {
        const nextDepth = depth + 1;
        if (!visited.has(next) || visited.get(next) > nextDepth) {
          visited.set(next, nextDepth);
          queue.push([next, nextDepth]);

          // 只補回 depth >= 2 的間接路徑（depth=1 直接邊本來就在）
          if (nextDepth >= 2 && affectedDescendants.includes(next)) {
            const key = `${start}|${next}|${nextDepth}`;
            if (!existing.has(key)) {
              existing.add(key);
              newRows.push([start, next, nextDepth, 'references', '間接關聯（重算）']);
            }
          }
        }
      });
    }
  });

  if (newRows.length > 0) {
    const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
    clsSheet.getRange(clsSheet.getLastRow() + 1, 1, newRows.length, 5)
      .setValues(newRows);
    SpreadsheetApp.flush();
  }
}

// 取得整體關聯圖資料（供前端畫關聯樹/網路圖）
// nodes 與 edges 都先過可見集：不可見文件與其連線完全不出現。
function apiGetGraphData() {
  const ctx = _assertWhitelisted();
  const visible = _getVisibleDocIds(ctx);
  const docs = _readDocs().filter(d => visible.has(d.doc_id));
  const closure = _readClosure();

  return {
    nodes: docs.map(d => ({
      id: d.doc_id, title: d.title, status: d.status, category: d.category,
    })),
    // 只回傳 depth=1 的直接邊，且兩端皆可見
    edges: closure
      .filter(c => c.depth === 1 &&
                   visible.has(c.ancestor_id) && visible.has(c.descendant_id))
      .map(c => ({
        from: c.ancestor_id, to: c.descendant_id,
        type: c.relation_type, description: c.description,
      })),
  };
}

// ============================================================
// 前端 API：標籤貼標 / 標籤樹管理 / 使用者授權（V3）
// 全部走 LockService＋_assert*＋_logAudit＋SpreadsheetApp.flush()
// ============================================================

// 整組覆寫某文件的標籤（僅管理員或該文件負責人）。
// V5：edit 授權者不可貼標——改標籤＝改可見範圍，屬權限管理行為。
function apiSetDocTags(docId, tagIds) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertOwnerOrAdmin(docId);

    // 僅接受既有標籤
    const validTagIds = new Set(_readTags().map(t => t.tag_id));
    const wanted = Array.from(new Set((tagIds || []).map(String).filter(Boolean)))
      .filter(t => validTagIds.has(t));

    const sheet = _getSheet(SHEET_NAMES.DOC_TAGS);
    const rows = sheet.getDataRange().getDisplayValues();
    // 刪除此文件既有貼標列（由下往上）
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][DOCTAG_COL.DOC_ID] === docId) sheet.deleteRow(i + 1);
    }
    // 寫入新貼標
    if (wanted.length > 0) {
      const newRows = wanted.map(t => [docId, t]);
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, DOCTAG_COL_COUNT)
        .setValues(newRows);
    }

    _logAudit('貼標', docId, '', `設定標籤：${wanted.join(', ') || '（清空）'}`);
    SpreadsheetApp.flush();
    return { success: true, tagIds: wanted };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 新增標籤（僅管理員）。parentId 空＝根節點。
function apiCreateTag(name, parentId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();
    const tagName = String(name || '').trim();
    if (!tagName) return { success: false, error: '標籤名稱不可為空' };

    const tags = _readTags();
    const parent = parentId ? String(parentId) : '';
    if (parent && !tags.some(t => t.tag_id === parent)) {
      return { success: false, error: '找不到父標籤：' + parent };
    }

    // 產生新 tag_id
    let maxNum = 0;
    tags.forEach(t => {
      const m = t.tag_id.match(/^TAG-(\d+)$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
    const newId = 'TAG-' + String(maxNum + 1).padStart(3, '0');

    // sort：同層最大 + 1
    const maxSort = tags.filter(t => t.parent_id === parent)
      .reduce((m, t) => Math.max(m, t.sort), 0);

    const sheet = _getSheet(SHEET_NAMES.TAGS);
    sheet.appendRow([newId, tagName, parent, maxSort + 1]);

    _logAudit('標籤管理', '', '',
      `新增標籤「${tagName}」(${newId})${parent ? ' 於 ' + parent : '（根節點）'}`);
    SpreadsheetApp.flush();
    return { success: true, tag_id: newId };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 標籤改名（僅管理員）。
function apiRenameTag(tagId, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();
    const tagName = String(name || '').trim();
    if (!tagName) return { success: false, error: '標籤名稱不可為空' };

    const sheet = _getSheet(SHEET_NAMES.TAGS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[TAG_COL.TAG_ID] === tagId);
    if (idx < 1) return { success: false, error: '找不到標籤：' + tagId };

    const oldName = rows[idx][TAG_COL.NAME];
    sheet.getRange(idx + 1, TAG_COL.NAME + 1).setValue(tagName);

    _logAudit('標籤管理', '', '', `標籤改名：${oldName} → ${tagName}(${tagId})`);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 搬移標籤到新的父節點（僅管理員）；拒絕搬到自己的子孫底下（防環）。
function apiMoveTag(tagId, newParentId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();

    const tags = _readTags();
    if (!tags.some(t => t.tag_id === tagId)) {
      return { success: false, error: '找不到標籤：' + tagId };
    }
    const parent = newParentId ? String(newParentId) : '';
    if (parent) {
      if (parent === tagId) {
        return { success: false, error: '不可將標籤搬到自己底下' };
      }
      if (!tags.some(t => t.tag_id === parent)) {
        return { success: false, error: '找不到父標籤：' + parent };
      }
      // 新父不可為自己的子孫（含自身）→ 會形成環
      const subtree = _expandTagWithDescendants([tagId], tags);
      if (subtree.has(parent)) {
        return { success: false, error: '不可搬移到自己的子孫底下（會形成環）' };
      }
    }

    const sheet = _getSheet(SHEET_NAMES.TAGS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[TAG_COL.TAG_ID] === tagId);
    if (idx < 1) return { success: false, error: '找不到標籤：' + tagId };
    sheet.getRange(idx + 1, TAG_COL.PARENT_ID + 1).setValue(parent);

    _logAudit('標籤管理', '', '', `搬移標籤 ${tagId} → 父 ${parent || '（根節點）'}`);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 刪除標籤（僅管理員）：連同整棵子樹一併刪除，
// 並級聯清除「文件標籤」與「使用者授權」中對這些標籤的引用。
function apiDeleteTag(tagId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();

    const tags = _readTags();
    if (!tags.some(t => t.tag_id === tagId)) {
      return { success: false, error: '找不到標籤：' + tagId };
    }
    // 待刪集合 = 此標籤 + 全部子孫（避免留下孤兒節點）
    const toDelete = _expandTagWithDescendants([tagId], tags);

    // 1. 刪除標籤主檔中的列
    const tagSheet = _getSheet(SHEET_NAMES.TAGS);
    const tagRows = tagSheet.getDataRange().getDisplayValues();
    for (let i = tagRows.length - 1; i >= 1; i--) {
      if (toDelete.has(tagRows[i][TAG_COL.TAG_ID])) tagSheet.deleteRow(i + 1);
    }

    // 2. 級聯清除文件標籤
    const dtSheet = _getSheet(SHEET_NAMES.DOC_TAGS);
    const dtRows = dtSheet.getDataRange().getDisplayValues();
    for (let i = dtRows.length - 1; i >= 1; i--) {
      if (toDelete.has(dtRows[i][DOCTAG_COL.TAG_ID])) dtSheet.deleteRow(i + 1);
    }

    // 3. 級聯清除使用者授權
    const gSheet = _getSheet(SHEET_NAMES.GRANTS);
    const gRows = gSheet.getDataRange().getDisplayValues();
    for (let i = gRows.length - 1; i >= 1; i--) {
      if (toDelete.has(gRows[i][GRANT_COL.TAG_ID])) gSheet.deleteRow(i + 1);
    }

    _logAudit('標籤管理', '', '',
      `刪除標籤 ${tagId} 及其子樹（共 ${toDelete.size} 個），並清除相關貼標與授權`);
    SpreadsheetApp.flush();
    return { success: true, deleted: Array.from(toDelete) };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 授權項參數正規化（V5）：接受 [{tagId, permission}]（字串項視為 read 舊格式），
// 同 tagId 去重取最高權限，僅保留既有標籤。
function _normGrantItems(grants) {
  const validTagIds = new Set(_readTags().map(t => t.tag_id));
  const byTag = {};
  (grants || []).forEach(g => {
    const isObj = !!g && typeof g === 'object';
    const tagId = String(isObj ? (g.tagId || '') : (g || '')).trim();
    if (!tagId || !validTagIds.has(tagId)) return;
    const perm = _normPermission(isObj ? g.permission : '');
    if (perm === 'edit' || !byTag[tagId]) byTag[tagId] = perm;
  });
  return Object.keys(byTag).map(t => ({ tagId: t, permission: byTag[t] }));
}

// 授權項的稽核摘要標示（edit 才標註，read 為預設不贅述）
function _grantAuditLabel(items) {
  return items.map(i => i.permission === 'edit' ? `${i.tagId}（編輯）` : i.tagId).join(', ');
}

// 整組覆寫某使用者的標籤授權（僅管理員）。V5：grants=[{tagId, permission}]。
function apiSetUserGrants(email, grants) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();

    const target = String(email || '').trim().toLowerCase();
    if (!target) return { success: false, error: '未指定使用者信箱' };

    const items = _normGrantItems(grants);

    const sheet = _getSheet(SHEET_NAMES.GRANTS);
    const rows = sheet.getDataRange().getDisplayValues();
    // 刪除此使用者既有授權列（由下往上）
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][GRANT_COL.EMAIL]).trim().toLowerCase() === target) {
        sheet.deleteRow(i + 1);
      }
    }
    // 寫入新授權
    if (items.length > 0) {
      const newRows = items.map(it => [target, it.tagId, it.permission]);
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, GRANT_COL_COUNT)
        .setValues(newRows);
    }

    _logAudit('授權', '', '',
      `設定 ${target} 授權標籤：${_grantAuditLabel(items) || '（清空）'}`);
    SpreadsheetApp.flush();
    return { success: true, email: target, grants: items };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 取得所有使用者授權（僅管理員；供權限管理分頁）。
// V5 回傳 [{email, grants:[{tagId, permission}]}]。讀取型：失敗直接 throw。
function apiGetAllGrants() {
  _assertAdmin();
  const map = {};
  _readGrants().forEach(g => {
    if (!map[g.email]) map[g.email] = [];
    map[g.email].push({ tagId: g.tag_id, permission: g.permission });
  });
  return Object.keys(map).map(email => ({ email: email, grants: map[email] }));
}

// ============================================================
// 前端 API：群組授權（V4，僅管理員）
// 群組成員名冊由 HR「人員職務配置」權威維護，本系統只存
// (org_code, title) → tag_id 對映。org_code 不展開子樹（設計定案）。
// ============================================================

// 全部群組授權列（附組織名稱供顯示）。讀取型：失敗直接 throw。
function apiGetGroupGrants() {
  _assertAdmin();
  const orgNameByCode = {};
  _getHrOrgTree().forEach(o => { orgNameByCode[o.code] = o.name; });
  return _readGroupGrants().map(g => ({
    org_code: g.org_code,
    org_name: g.org_code ? (orgNameByCode[g.org_code] || '（查無組織）') : '',
    title:    g.title,
    tag_id:   g.tag_id,
    permission: g.permission,
  }));
}

// 整組覆寫 (orgCode, title) 組合的授權；grants=[] 即刪除該組合（V5 帶 permission）。
function apiSetGroupGrants(orgCode, title, grants) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    _assertAdmin();

    const org = String(orgCode || '').trim();
    const ttl = String(title || '').trim();
    if (!org && !ttl) return { success: false, error: '組織與職稱至少擇一' };

    const items = _normGrantItems(grants);

    const sheet = _getSheet(SHEET_NAMES.GROUP_GRANTS);
    const rows = sheet.getDataRange().getDisplayValues();
    // 刪除該 (org, title) 組合既有授權列（由下往上）
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][GROUPGRANT_COL.ORG_CODE]).trim() === org &&
          String(rows[i][GROUPGRANT_COL.TITLE]).trim() === ttl) {
        sheet.deleteRow(i + 1);
      }
    }
    if (items.length > 0) {
      const newRows = items.map(it => [org, ttl, it.tagId, it.permission]);
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, GROUPGRANT_COL_COUNT)
        .setValues(newRows);
    }

    const label = `${org || '（不限組織）'}／${ttl || '（不限職稱）'}`;
    _logAudit('授權', '', '',
      `設定群組授權 ${label} → ${_grantAuditLabel(items) || '（清空）'}`);
    SpreadsheetApp.flush();
    return { success: true, org_code: org, title: ttl, grants: items };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

// 組織節點與現存職稱清單（管理 UI 下拉；職稱只能選不能打字，防打錯字）。
function apiGetOrgOptions() {
  _assertAdmin();
  const orgs = _getHrOrgTree();
  const titles = Array.from(new Set(
    _getHrAssignments().map(a => a.title).filter(Boolean)
  )).sort();
  return { orgs: orgs, titles: titles };
}

// 有效權限預覽：某使用者最終取得的標籤、最高權限與各來源（僅管理員）。
// 在後端計算，避免把全公司職務配置送到前端。
function apiPreviewUserTags(email) {
  _assertAdmin();
  const target = String(email || '').trim().toLowerCase();
  if (!target) return [];

  const byTag = {};   // tag_id → { permission, sources: [] }
  const add = (tagId, permission, source) => {
    if (!byTag[tagId]) byTag[tagId] = { permission: 'read', sources: [] };
    if (permission === 'edit') byTag[tagId].permission = 'edit';
    const label = source + (permission === 'edit' ? '（編輯）' : '（讀取）');
    if (byTag[tagId].sources.indexOf(label) < 0) byTag[tagId].sources.push(label);
  };

  _readGrants().forEach(g => {
    if (g.email === target) add(g.tag_id, g.permission, '個人授權');
  });

  const hits = _groupGrantHits(target);
  if (hits.length > 0) {
    const orgNameByCode = {};
    _getHrOrgTree().forEach(o => { orgNameByCode[o.code] = o.name; });
    hits.forEach(gg => {
      const orgLabel = gg.org_code
        ? (orgNameByCode[gg.org_code] || gg.org_code) : '不限組織';
      add(gg.tag_id, gg.permission, `群組授權：${orgLabel}／${gg.title || '不限職稱'}`);
    });
  }

  return Object.keys(byTag).map(tagId => ({
    tag_id: tagId,
    permission: byTag[tagId].permission,
    sources: byTag[tagId].sources,
  }));
}

function apiGetDocxExportData(tagId) {
  const ctx = getUserContext();
  const visibleDocIds = _getVisibleDocIds(ctx);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get all documents
  const docSheet = ss.getSheetByName(ENV.SHEET_DOC);
  const docData = docSheet.getDataRange().getDisplayValues();
  const allDocs = {};
  for (let i = 1; i < docData.length; i++) {
    const row = docData[i];
    const docId = row[DOC_COL.DOC_ID];
    if (visibleDocIds.has(docId)) {
      allDocs[docId] = {
        doc_id: docId,
        title: row[DOC_COL.TITLE],
        category: row[DOC_COL.CATEGORY],
        security_level: row[DOC_COL.SECURITY_LEVEL] || '一般',
        version: row[DOC_COL.VERSION],
        published_at: row[DOC_COL.PUBLISHED_AT]
      };
    }
  }

  // 2. Get tags for documents to filter by tagId
  const docTagSheet = ss.getSheetByName(ENV.SHEET_DOCTAG);
  const docTagData = docTagSheet.getDataRange().getValues();
  const docsWithTag = new Set();
  for (let i = 1; i < docTagData.length; i++) {
    const dId = String(docTagData[i][DOCTAG_COL.DOC_ID]).trim();
    const tId = String(docTagData[i][DOCTAG_COL.TAG_ID]).trim();
    if (tId === tagId && allDocs[dId]) {
      docsWithTag.add(dId);
    }
  }

  // 3. Get closure relationships
  const clsSheet = ss.getSheetByName(ENV.SHEET_CLS);
  const clsData = clsSheet.getDataRange().getValues();
  
  const parentToChildren = {};
  const allChildren = new Set();
  
  for (let i = 1; i < clsData.length; i++) {
    const row = clsData[i];
    const ancestor = String(row[CLS_COL.ANCESTOR_ID]).trim();
    const descendant = String(row[CLS_COL.DESCENDANT_ID]).trim();
    const depth = parseInt(row[CLS_COL.DEPTH], 10);
    
    if (depth === 1 && docsWithTag.has(ancestor) && docsWithTag.has(descendant)) {
      if (!parentToChildren[ancestor]) {
        parentToChildren[ancestor] = [];
      }
      parentToChildren[ancestor].push(descendant);
      allChildren.add(descendant);
    }
  }

  // 4. Flatten relationships
  const flattenedData = [];
  
  // Sort docsWithTag for stable output if needed, but we'll just iterate
  // It's better to sort by docId so the table looks organized
  const sortedDocsWithTag = Array.from(docsWithTag).sort();
  
  for (const docId of sortedDocsWithTag) {
    // If it's a child to another doc in this tag, it shouldn't be treated as a parent
    if (allChildren.has(docId)) continue;
    
    const parentDoc = allDocs[docId];
    const childrenIds = parentToChildren[docId] || [];
    // Sort children for stable output
    childrenIds.sort();
    
    if (childrenIds.length === 0) {
      flattenedData.push({
        doc_id: parentDoc.doc_id,
        title: parentDoc.title,
        category: parentDoc.category,
        security_level: parentDoc.security_level,
        version: parentDoc.version,
        published_at: parentDoc.published_at,
        form_id: "",
        form_title: "",
        form_version: "",
        form_published_at: ""
      });
    } else {
      childrenIds.forEach((childId, index) => {
        const childDoc = allDocs[childId];
        if (index === 0) {
          flattenedData.push({
            doc_id: parentDoc.doc_id,
            title: parentDoc.title,
            category: parentDoc.category,
            security_level: parentDoc.security_level,
            version: parentDoc.version,
            published_at: parentDoc.published_at,
            form_id: childDoc.doc_id,
            form_title: childDoc.title,
            form_version: childDoc.version,
            form_published_at: childDoc.published_at
          });
        } else {
          flattenedData.push({
            doc_id: "",
            title: "",
            category: "",
            security_level: "",
            version: "",
            published_at: "",
            form_id: childDoc.doc_id,
            form_title: childDoc.title,
            form_version: childDoc.version,
            form_published_at: childDoc.published_at
          });
        }
      });
    }
  }

  // 5. Fetch template from Drive
  const templateId = _getProp(PROP_KEYS.DOCX_TEMPLATE_FILE_ID);
  if (!templateId) {
    throw new Error('系統尚未設定 Docx 範本檔案 (DOCX_TEMPLATE_FILE_ID)。');
  }
  
  let templateBase64 = "";
  try {
    const file = DriveApp.getFileById(templateId);
    templateBase64 = Utilities.base64Encode(file.getBlob().getBytes());
  } catch (e) {
    throw new Error('無法讀取 Docx 範本檔案，請檢查檔案 ID 或權限。');
  }

  const outputFolderId = _getProp(PROP_KEYS.DOCX_OUTPUT_FOLDER_ID);
  let recordNo = "";
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const year = Utilities.formatDate(now, tz, 'yyyy');
  const month = Utilities.formatDate(now, tz, 'MM');
  const day = Utilities.formatDate(now, tz, 'dd');
  const dateKey = Utilities.formatDate(now, tz, 'yyyyMMdd');

  if (outputFolderId) {
    const prefix = _getProp(PROP_KEYS.RECORD_NUMBER_PREFIX) || 'IS-R-032';
    recordNo = _createRecordNoFromFolder(outputFolderId, prefix, dateKey);
  }

  return {
    templateBase64: templateBase64,
    data: flattenedData,
    year: year,
    month: month,
    day: day,
    recordNo: recordNo
  };
}

function _escapeRegExp(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function _createRecordNoFromFolder(folderId, prefix, dateKey) {
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch(e) {
    throw new Error('找不到輸出資料夾 (DOCX_OUTPUT_FOLDER_ID) 或無權限。');
  }

  const escapedPrefix = _escapeRegExp(prefix);
  const pattern = new RegExp('^' + escapedPrefix + '-' + dateKey + '-(\\d+)');

  let maxSerial = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = String(file.getName() || '').trim();
    const match = name.match(pattern);
    if (!match) continue;

    const serial = parseInt(match[1], 10);
    if (!isNaN(serial) && serial > maxSerial) {
      maxSerial = serial;
    }
  }

  const nextSerial = maxSerial + 1;
  const serialText = nextSerial < 100 ? ('0' + nextSerial).slice(-2) : String(nextSerial);
  return prefix + '-' + dateKey + '-' + serialText;
}

function apiSaveDocxToDrive(base64Data, fileName) {
  const ctx = getUserContext();
  if (!ctx.isWhitelisted) {
    throw new Error("無存取權限");
  }

  const folderId = _getProp(PROP_KEYS.DOCX_OUTPUT_FOLDER_ID);
  if (!folderId) {
    throw new Error("系統尚未設定 DOCX_OUTPUT_FOLDER_ID");
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data), 
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
      fileName
    );
    const newFile = folder.createFile(blob);
    return {
      success: true,
      fileId: newFile.getId(),
      url: newFile.getUrl()
    };
  } catch (err) {
    throw new Error("存檔失敗：" + err.message);
  }
}

// 批次更新文件屬性 (機密等級與發行日期)
function apiBatchUpdateMetadata(updateList) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ctx = _assertAdmin(); // 僅限管理員
    if (!updateList || !updateList.length) return { success: true, updatedCount: 0 };

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    
    // 建立 doc_id 對應 row index 的 Map (跳過表頭)
    const docRowMap = new Map();
    for (let i = 1; i < rows.length; i++) {
      const dId = String(rows[i][DOC_COL.DOC_ID]).trim();
      if (dId) docRowMap.set(dId, i);
    }

    let updatedCount = 0;
    // 走訪並更新
    for (const item of updateList) {
      const rowIndex = docRowMap.get(item.doc_id);
      if (rowIndex !== undefined) {
        // 更新記憶體中的 rows 陣列
        if (item.security_level !== undefined) {
          // 允許清空或為合法值
          if (item.security_level === '' || SECURITY_LEVELS.includes(item.security_level)) {
            rows[rowIndex][DOC_COL.SECURITY_LEVEL] = item.security_level;
          }
        }
        if (item.published_at !== undefined) {
          rows[rowIndex][DOC_COL.PUBLISHED_AT] = item.published_at;
        }
        
        // 寫入異動紀錄
        _logAudit('批次更新屬性', item.doc_id, rows[rowIndex][DOC_COL.VERSION], `更新由 Excel 匯入`);
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      // 一次性寫回所有資料
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      SpreadsheetApp.flush();
    }

    return { success: true, updatedCount: updatedCount };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

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
      doc_id:       r[DOC_COL.DOC_ID],
      title:        r[DOC_COL.TITLE],
      category:     r[DOC_COL.CATEGORY],
      status:       r[DOC_COL.STATUS],
      owner:        r[DOC_COL.OWNER],
      owner_id:     r[DOC_COL.OWNER_ID],
      updated_at:   r[DOC_COL.UPDATED_AT],
      version:      r[DOC_COL.VERSION],
      drive_loc:    r[DOC_COL.GOOGLE_DRIVE_LOC],
      owner_email:  r[DOC_COL.OWNER_EMAIL] || '',
      published_at: r[DOC_COL.PUBLISHED_AT] || '',
      next_review:  r[DOC_COL.NEXT_REVIEW] || '',
      review_cycle: r[DOC_COL.REVIEW_CYCLE] || '',
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
  row[DOC_COL.OWNER_ID]         = doc.owner_id || '';
  row[DOC_COL.UPDATED_AT]       = _now();
  row[DOC_COL.VERSION]          = doc.version || '0.1';
  row[DOC_COL.GOOGLE_DRIVE_LOC] = doc.drive_loc || '';
  row[DOC_COL.OWNER_EMAIL]      = doc.owner_email || '';
  row[DOC_COL.PUBLISHED_AT]     = doc.published_at || '';
  row[DOC_COL.NEXT_REVIEW]      = doc.next_review || '';
  row[DOC_COL.REVIEW_CYCLE]     = doc.review_cycle || '';
  return row;
}

// ============================================================
// 前端 API：文件 CRUD
// ============================================================

// 取得所有文件（含選項清單與使用者情境，供前端初始化一次取完）
function apiGetInitData() {
  const ctx = getUserContext();
  return {
    docs: _readDocs(),
    statuses: DOC_STATUS,
    categories: DOC_CATEGORIES,
    relationTypes: RELATION_TYPES,
    reviewCycles: REVIEW_CYCLES,
    statusTransitions: STATUS_TRANSITIONS,
    user: ctx,
    // 負責人下拉選項：白名單即可看到（僅姓名與信箱，無其他個資）
    hrPeople: ctx.isWhitelisted ? _getHrPeople() : [],
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

    _logAudit('建立', newId, newDoc.version || '0.1', `建立文件「${newDoc.title || ''}」`);
    SpreadsheetApp.flush();
    return { success: true, doc_id: newId };
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
    const merged = Object.assign({}, doc, {
      published_at: oldDoc.published_at,
      next_review:  oldDoc.next_review,
      status: newStatus,
    });
    if (newStatus === '已發布' && oldStatus !== '已發布') {
      const cycle = parseInt(doc.review_cycle, 10) || DEFAULT_REVIEW_CYCLE;
      merged.published_at = _now();
      merged.next_review  = _addMonthsFromToday(cycle);
    }

    sheet.getRange(idx + 1, 1, 1, DOC_COL_COUNT).setValues([_docToRow(merged)]);

    const action = (newStatus !== oldStatus) ? '狀態變更' : '更新';
    const summary = (newStatus !== oldStatus)
      ? `${oldStatus} → ${newStatus}` + (merged.published_at !== oldDoc.published_at ? `（發布日 ${merged.published_at}，下次審查 ${merged.next_review}）` : '')
      : (_diffSummary(oldDoc, doc) || '（無欄位變更）');
    _logAudit(action, doc.doc_id, doc.version || oldDoc.version, summary);

    SpreadsheetApp.flush();
    return { success: true, published_at: merged.published_at, next_review: merged.next_review };
  } catch (e) {
    return { success: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
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

    _logAudit('刪除', docId, '', `刪除文件「${title}」（含所有關聯記錄）`);
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
function apiGetDescendants(docId, maxDepth) {
  const closure = _readClosure();
  const docsMap = new Map(_readDocs().map(d => [d.doc_id, d]));
  const limit = maxDepth || 99;

  return closure
    .filter(c => c.ancestor_id === docId && c.depth > 0 && c.depth <= limit)
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
function apiGetAncestors(docId) {
  const closure = _readClosure();
  const docsMap = new Map(_readDocs().map(d => [d.doc_id, d]));

  return closure
    .filter(c => c.descendant_id === docId && c.depth > 0)
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
function apiGetGraphData() {
  const docs = _readDocs();
  const closure = _readClosure();

  return {
    nodes: docs.map(d => ({
      id: d.doc_id, title: d.title, status: d.status, category: d.category,
    })),
    // 只回傳 depth=1 的直接邊（圖形繪製只需要直接邊）
    edges: closure
      .filter(c => c.depth === 1)
      .map(c => ({
        from: c.ancestor_id, to: c.descendant_id,
        type: c.relation_type, description: c.description,
      })),
  };
}

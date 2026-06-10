// ============================================================
// code.js — 核心業務邏輯
// 文件 CRUD + Closure Table 關聯維護 + 前端 API
// ============================================================

// ── Web App 進入點 ────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('文件管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
      doc_id:     r[DOC_COL.DOC_ID],
      title:      r[DOC_COL.TITLE],
      category:   r[DOC_COL.CATEGORY],
      status:     r[DOC_COL.STATUS],
      owner:      r[DOC_COL.OWNER],
      owner_id:   r[DOC_COL.OWNER_ID],
      updated_at: r[DOC_COL.UPDATED_AT],
      version:    r[DOC_COL.VERSION],
      drive_loc:  r[DOC_COL.GOOGLE_DRIVE_LOC],
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

// ============================================================
// 前端 API：文件 CRUD
// ============================================================

// 取得所有文件（含選項清單，供前端初始化一次取完）
function apiGetInitData() {
  return {
    docs: _readDocs(),
    statuses: DOC_STATUS,
    categories: DOC_CATEGORIES,
    relationTypes: RELATION_TYPES,
  };
}

// 新增文件（自動產生 doc_id + 寫入 self closure 記錄）
function apiCreateDoc(doc) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = _getSheet(SHEET_NAMES.DOCS);

    // 產生新 doc_id：掃描現有最大序號
    const existing = _readDocs().map(d => d.doc_id);
    let maxNum = 0;
    existing.forEach(id => {
      const m = id.match(/^DOC-(\d+)$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
    const newId = 'DOC-' + String(maxNum + 1).padStart(3, '0');

    sheet.appendRow([
      newId,
      doc.title || '',
      doc.category || '',
      doc.status || '草稿',
      doc.owner || '',
      doc.owner_id || '',
      _now(),
      doc.version || '0.1',
      doc.drive_loc || '',
    ]);

    // Closure Table：寫入自身記錄（depth=0）
    const clsSheet = _getSheet(SHEET_NAMES.CLOSURE);
    clsSheet.appendRow([newId, newId, 0, 'references', '自身']);

    SpreadsheetApp.flush();
    return { success: true, doc_id: newId };
  } catch (e) {
    return { success: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

// 更新文件（依 doc_id 定位列，更新欄位）
function apiUpdateDoc(doc) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === doc.doc_id);
    if (idx < 1) return { success: false, error: '找不到文件：' + doc.doc_id };

    const rowNum = idx + 1; // 1-based
    sheet.getRange(rowNum, DOC_COL.TITLE + 1).setValue(doc.title);
    sheet.getRange(rowNum, DOC_COL.CATEGORY + 1).setValue(doc.category);
    sheet.getRange(rowNum, DOC_COL.STATUS + 1).setValue(doc.status);
    sheet.getRange(rowNum, DOC_COL.OWNER + 1).setValue(doc.owner);
    sheet.getRange(rowNum, DOC_COL.OWNER_ID + 1).setValue(doc.owner_id);
    sheet.getRange(rowNum, DOC_COL.UPDATED_AT + 1).setValue(_now());
    sheet.getRange(rowNum, DOC_COL.VERSION + 1).setValue(doc.version);
    sheet.getRange(rowNum, DOC_COL.GOOGLE_DRIVE_LOC + 1).setValue(doc.drive_loc);

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e) };
  } finally {
    lock.releaseLock();
  }
}

// 刪除文件（同步刪除閉包表中所有相關記錄）
function apiDeleteDoc(docId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // 1. 刪除文件清單中的列
    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
    if (idx < 1) return { success: false, error: '找不到文件：' + docId };
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

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: String(e) };
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
function apiAddRelation(ancestorId, descendantId, relationType, description) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
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
      SpreadsheetApp.flush();
    }

    return { success: true, added: newRows.length };
  } catch (e) {
    return { success: false, error: String(e) };
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

    return { success: true };
  } catch (e) {
    return { success: false, error: String(e) };
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

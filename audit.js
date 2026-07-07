// ============================================================
// audit.js — 異動紀錄（版本歷史與操作稽核合一）
//
// 設計取捨：不另設「版本歷史」表——每筆寫入操作都記一列，
// 文件的版本歷史 = 以 doc_id 過濾此表，一張表服務兩個需求。
// ============================================================

function _nowWithTime() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
}

// 寫入一筆異動紀錄。由各 API 在持有 lock 的區段內呼叫，
// 記錄失敗不應讓主操作跟著失敗 → 內部吞錯誤只留 log。
function _logAudit(action, docId, version, summary) {
  try {
    const sheet = _getSheet(SHEET_NAMES.AUDIT);
    sheet.appendRow([
      _nowWithTime(),
      _getCurrentEmail(),
      action,
      docId || '',
      version || '',
      summary || '',
    ]);
  } catch (e) {
    Logger.log('⚠️ 異動紀錄寫入失敗（主操作不受影響）：' + e);
  }
}

// 產生更新操作的變更摘要：只列出真的有變的欄位
function _diffSummary(oldDoc, newDoc) {
  const FIELD_LABELS = {
    title: '標題', category: '類別', status: '狀態',
    owner: '負責人', owner_id: '負責人ID', owner_email: '負責人信箱',
    version: '版本', drive_loc: 'Drive位置', review_cycle: '審查週期',
  };
  const changes = [];
  Object.keys(FIELD_LABELS).forEach(key => {
    const oldVal = String(oldDoc[key] ?? '');
    const newVal = String(newDoc[key] ?? '');
    if (oldVal !== newVal) {
      changes.push(`${FIELD_LABELS[key]}：${oldVal || '（空）'} → ${newVal || '（空）'}`);
    }
  });
  return changes.join('；');
}

// ── 前端 API：查詢單一文件的異動歷史（新→舊排序）─────────────
function apiGetDocHistory(docId) {
  const sheet = _getSheet(SHEET_NAMES.AUDIT);
  const rows = sheet.getDataRange().getDisplayValues();

  return rows.slice(1)
    .filter(r => r[AUDIT_COL.DOC_ID] === docId)
    .map(r => ({
      timestamp: r[AUDIT_COL.TIMESTAMP],
      operator:  r[AUDIT_COL.OPERATOR],
      action:    r[AUDIT_COL.ACTION],
      version:   r[AUDIT_COL.VERSION],
      summary:   r[AUDIT_COL.SUMMARY],
    }))
    .reverse();
}

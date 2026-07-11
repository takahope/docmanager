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
    owner: '負責人', owner_email: '負責人信箱',
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

// 遮蔽異動摘要中「使用者無權檢視」的 doc_id。
// 關聯類動作會把對象 doc_id 寫進摘要（如「A → DOC-B」），若 DOC-B 不在
// 可見集，直接回傳會洩漏其存在，違反「不可見文件完全隱藏」。
function _redactInvisibleDocIds(summary, visibleSet) {
  return String(summary || '').replace(/DOC-\d+/g, function(id) {
    return visibleSet.has(id) ? id : '（無權限文件）';
  });
}

// ── 前端 API：查詢單一文件的異動歷史（新→舊排序）─────────────
// 入口先過 _assertCanViewDoc：不可見文件的歷史完全不回傳；
// 可見文件的摘要再經 _redactInvisibleDocIds 遮蔽其中不可見的關聯對象。
function apiGetDocHistory(docId) {
  const ctx = _assertCanViewDoc(docId);
  const visible = _getVisibleDocIds(ctx);
  const sheet = _getSheet(SHEET_NAMES.AUDIT);
  const rows = sheet.getDataRange().getDisplayValues();

  return rows.slice(1)
    .filter(r => r[AUDIT_COL.DOC_ID] === docId)
    .map(r => ({
      timestamp: r[AUDIT_COL.TIMESTAMP],
      operator:  r[AUDIT_COL.OPERATOR],
      action:    r[AUDIT_COL.ACTION],
      version:   r[AUDIT_COL.VERSION],
      summary:   _redactInvisibleDocIds(r[AUDIT_COL.SUMMARY], visible),
    }))
    .reverse();
}

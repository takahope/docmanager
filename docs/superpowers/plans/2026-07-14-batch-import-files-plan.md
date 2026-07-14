# 管理員批次匯入正式檔案 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 為 Google Apps Script (GAS) 文件管理系統新增管理員專用「批次匯入正式檔案」功能，透過前端智能解析檔名對應 `doc_id` 與提取 `version`，提供上傳前對位預覽表，並透過前端循序呼叫後端 API 把檔案寫入現行生效版（及「已發布」狀態），避開 GAS 單次 Payload 限制與 6 分鐘超時。

**Architecture:** 
- **後端 (`code.js`)**：新增 `apiBatchImportDirectFile(docId, fileName, base64, mimeType, targetVersion)` 函式，執行身分查驗、 Drive 檔案儲存、整列原子更新試算表 M/G/J/K 欄位（`file_id`/`version`/`published_at`/`next_review`）及清空待核欄位 N/O/P，同步寫入「檔案版本 (`FILE_VERSIONS`)」歷史紀錄。
- **前端 (`index.html`)**：在列表頁新增管理員按鈕與 Modal，包含檔案選取 input、對照預覽表 (`#batchImportPreviewTable`) 與循序上傳進度列，並以 `for...of` 迴圈搭配 `FileReader Base64` 循序上傳完成多檔匯入。

**Tech Stack:** Google Apps Script (GAS), Vanilla JavaScript, Google Sheets API, HTML5/CSS3 (Bootstrap/Custom Design System).

## Global Constraints

- **GAS 專案特性**：無本地建置與單元測試框架，語法檢查需以 `node --check <file>.js` 執行（HTML 內嵌 JS 亦以嚴格語法檢視與人工核對為主）。
- **欄位常數依賴**：嚴格使用 `env.js` 的 `DOC_COL` 常數 (`STATUS`: 3, `VERSION`: 6, `PUBLISHED_AT`: 9, `NEXT_REVIEW`: 10, `REVIEW_CYCLE`: 11, `FILE_ID`: 12, `PENDING_FILE_ID`: 13, `PENDING_VERSION`: 14, `PENDING_FILE_NAME`: 15)。
- **鎖定與安全性**：所有寫入 API 均需以 `_assertAdmin()` 保護並調用 `LockService.getScriptLock()` 與 `SpreadsheetApp.flush()`。

---

### Task 1: 後端 `apiBatchImportDirectFile` API 實作

**Files:**
- Modify: `code.js:330-340` (在 `apiUploadDocFile` 後或適當 API 區域新增 `apiBatchImportDirectFile` 函式)

**Interfaces:**
- Consumes: `_assertAdmin()`, `_readDocs()`, `_getOrCreateDocFolder()`, `_getSheet(SHEET_NAMES.DOCS)`, `_getSheet(SHEET_NAMES.FILE_VERSIONS)`, `_logAudit()`, `_now()`, `_addMonthsFromToday()`, `_docToRow()`, `_getCurrentEmail()`, `_nowWithTime()`
- Produces: `apiBatchImportDirectFile(docId, fileName, base64, mimeType, targetVersion)` 回傳 `{ success: boolean, docId?: string, version?: string, status?: string, file_id?: string, error?: string }`

- [ ] **Step 1: 在 `code.js` 中新增 `apiBatchImportDirectFile` 實作程式碼**

在 `code.js` 的 `apiUploadDocFile` 結束後（約第 334 行），插入以下完整實作：

```javascript
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

    // 解碼 Base64 並存入 Google Drive
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream',
      `${docId}_v${ver}_${safeName}`);
    const folder = _getOrCreateDocFolder(docId);
    const file = folder.createFile(blob);

    const sheet = _getSheet(SHEET_NAMES.DOCS);
    const rows = sheet.getDataRange().getDisplayValues();
    const idx = rows.findIndex(r => r[DOC_COL.DOC_ID] === docId);
    if (idx < 1) return { success: false, error: '找不到試算表列：' + docId };

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
```

- [ ] **Step 2: 執行 Node 語法驗證**

Run: `node --check code.js`
Expected: 不產生任何錯誤訊息（exit code 0）。

- [ ] **Step 3: Commit 後端 API**

```bash
git add code.js
git commit -m "feat: add apiBatchImportDirectFile for admin direct file batch import"
```

---

### Task 2: 前端管理員批次匯入按鈕與 Modal 彈窗介面

**Files:**
- Modify: `index.html` (在列表視圖操作區新增按鈕，並於 Modal 集中區新增 `batchImportModal`)

**Interfaces:**
- Consumes: `state.user.isAdmin`
- Produces: DOM `<button id="btnOpenBatchImport">` 與 Modal `<div id="batchImportModal">`

- [ ] **Step 1: 於 `index.html` 新增「批次匯入檔案」按鈕 HTML 或渲染邏輯**

在 `index.html` 的列表工具列區塊（約於 `#docSearchInput` 或「新增文件」按鈕旁，搜尋 `onclick="openCreateModal()"` 所在區塊），插入以 `isAdmin` 控制顯示的按鈕：

```html
          <button id="btnOpenBatchImport" class="btn btn-outline-primary me-2" style="display:none;" onclick="openBatchImportModal()">
            <i class="fas fa-file-import"></i> 批次匯入檔案
          </button>
```

同時在 `renderDocList()` 或 `updateAuthUI()` (處理權限畫面展示的 JS 函式中，找到控制 `btnOpenCreateModal` 或管理員按鈕的地方)，加入同步顯示控制：

```javascript
  const btnBatchImport = document.getElementById('btnOpenBatchImport');
  if (btnBatchImport) {
    btnBatchImport.style.display = (state.user && state.user.isAdmin) ? 'inline-block' : 'none';
  }
```

- [ ] **Step 2: 於 `index.html` 底部 Modal 集中區新增 `batchImportModal` 對話框結構**

在 `index.html` 底部其他 modal 旁（如 `uploadFileModal` 附近），新增以下 HTML 結構：

```html
<!-- 管理員批次匯入檔案 Modal -->
<div class="modal fade" id="batchImportModal" tabindex="-1" aria-hidden="true" data-bs-backdrop="static">
  <div class="modal-dialog modal-xl modal-dialog-centered modal-dialog-scrollable">
    <div class="modal-content">
      <div class="modal-header bg-primary text-white">
        <h5 class="modal-title"><i class="fas fa-file-import me-2"></i>管理員批次匯入正式檔案 (直接合位與發布)</h5>
        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" id="btnCloseBatchImportModalTop"></button>
      </div>
      <div class="modal-body">
        <div class="alert alert-info py-2 mb-3">
          <i class="fas fa-info-circle me-1"></i> 
          系統將根據選取之 PDF 檔案名稱自動提取 <code>doc_id</code>（例如 <code>TWHB-ISMS-002-001</code>）與版本號（例如 <code>V2.1</code>），配對成功的檔案將直接取代該文件之生效檔案並轉為「已發布」。
        </div>

        <!-- 檔案選擇器 -->
        <div class="mb-3">
          <label for="batchImportFileInput" class="form-label fw-bold">選擇待匯入多檔案 (可多選 10~50 個檔)：</label>
          <input class="form-control" type="file" id="batchImportFileInput" multiple accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx">
        </div>

        <!-- 對應預覽清單 -->
        <div class="table-responsive mb-3" style="max-height: 380px;">
          <table class="table table-sm table-bordered table-hover align-middle mb-0">
            <thead class="table-light sticky-top">
              <tr>
                <th style="width: 40px;" class="text-center"><input type="checkbox" id="chkSelectAllBatchItems" checked onclick="toggleSelectAllBatchItems(this)"></th>
                <th>原始檔案名稱</th>
                <th>大小</th>
                <th>配對對應文件 (ID & 標題)</th>
                <th style="width: 140px;">現狀態 ➔ 目標</th>
                <th style="width: 110px;">生效版號</th>
                <th style="width: 130px;" class="text-center">狀態 / 檢查</th>
              </tr>
            </thead>
            <tbody id="batchImportPreviewTbody">
              <tr>
                <td colspan="7" class="text-center text-muted py-4">請點擊上方按鈕選擇檔案以產生配對預覽...</td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- 上傳進度條區塊 (初始隱藏) -->
        <div id="batchImportProgressArea" style="display: none;" class="mt-3">
          <div class="d-flex justify-content-between align-items-center mb-1">
            <span id="batchImportProgressText" class="fw-bold text-primary">正在準備開始批次上傳...</span>
            <span id="batchImportProgressPercent" class="badge bg-primary">0%</span>
          </div>
          <div class="progress" style="height: 20px;">
            <div id="batchImportProgressBar" class="progress-bar progress-bar-striped progress-bar-animated bg-primary" role="progressbar" style="width: 0%;"></div>
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal" id="btnCloseBatchImportModalBottom">取消 / 關閉</button>
        <button type="button" class="btn btn-primary fw-bold" id="btnStartBatchImportExecution" disabled onclick="startBatchImportExecution()">
          <i class="fas fa-cloud-upload-alt me-1"></i> 確認並開始批次匯入
        </button>
      </div>
    </div>
  </div>
</div>
```

- [ ] **Step 3: 執行 Node 語法檢查及 UI 架構確認**

Run: `node --check index.html` (如果有報錯由於含有 HTML 標籤屬正常，可以用 grep 確認新增 DOM 結構無拼寫錯誤：`grep -n "batchImportModal" index.html`)
Expected: 成功找到 `#batchImportModal`。

- [ ] **Step 4: Commit 前端 UI 結構**

```bash
git add index.html
git commit -m "feat: add UI modal and buttons for batch importing files"
```

---

### Task 3: 前端智能解析檔名對位與預覽表生成邏輯

**Files:**
- Modify: `index.html` (在 JS `<script>` 區塊新增解析與渲染相關函式)

**Interfaces:**
- Consumes: `state.allDocs` (或目前 `state.docs` 中之可見列表)
- Produces: `window.batchImportItems = []`, `openBatchImportModal()`, `parseAndMatchBatchFiles(files)`, `renderBatchImportPreviewTable()`

- [ ] **Step 1: 新增全域暫存變數與開啟 Modal 函式**

在 `index.html` 的 JS 全域變數區域（或專屬管理員 JS 函式區），新增以下程式碼：

```javascript
// 儲存批次匯入配對佇列與狀態
let batchImportItems = [];

// 開啟批次匯入 Modal
function openBatchImportModal() {
  if (!state.user || !state.user.isAdmin) {
    alert('僅限管理員執行此操作');
    return;
  }
  batchImportItems = [];
  const fileInput = document.getElementById('batchImportFileInput');
  if (fileInput) fileInput.value = '';
  const progressArea = document.getElementById('batchImportProgressArea');
  if (progressArea) progressArea.style.display = 'none';
  const btnStart = document.getElementById('btnStartBatchImportExecution');
  if (btnStart) btnStart.disabled = true;

  renderBatchImportPreviewTable();

  // 綁定選取檔案監聽事件
  if (fileInput && !fileInput.dataset.bound) {
    fileInput.addEventListener('change', function(e) {
      parseAndMatchBatchFiles(e.target.files || []);
    });
    fileInput.dataset.bound = 'true';
  }

  const modalEl = document.getElementById('batchImportModal');
  const modal = new bootstrap.Modal(modalEl);
  modal.show();
}
```

- [ ] **Step 2: 新增智能解析檔名與生成預覽清單實作**

加入 `parseAndMatchBatchFiles(files)` 與對應之 `renderBatchImportPreviewTable()`：

```javascript
// 正則轉義工具
function _escapeRegExpBatch(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// 解析多檔案與配對 logic
function parseAndMatchBatchFiles(fileList) {
  batchImportItems = [];
  const docs = state.allDocs || state.docs || [];

  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    const name = file.name;
    let matchedDoc = null;

    // 依據 doc_id 進行匹配：尋找以 doc_id 開頭且後面跟隨 _ / - / 空格的對象
    for (let d = 0; d < docs.length; d++) {
      const doc = docs[d];
      if (!doc || !doc.doc_id) continue;
      const regex = new RegExp('^' + _escapeRegExpBatch(doc.doc_id) + '([_\\-\\s]|$)', 'i');
      if (regex.test(name)) {
        matchedDoc = doc;
        break;
      }
    }

    // 解析提取版本號 (例如 V2.1)
    let extractedVersion = '';
    const verMatch = name.match(/[_-\s][vV]([0-9]+(?:\.[0-9]+)*)(?:[_-\s][^.]+)?\.[^.]+$/i);
    if (verMatch && verMatch[1]) {
      extractedVersion = verMatch[1];
    } else if (matchedDoc) {
      extractedVersion = matchedDoc.version || '1.0';
    } else {
      extractedVersion = '1.0';
    }

    // 檢查檔案大小是否大於 20MB
    const isOverSize = file.size > 20 * 1024 * 1024;
    let status = 'matched';
    let statusText = '✔ 完全對應';
    let badgeClass = 'bg-success';

    if (isOverSize) {
      status = 'error_size';
      statusText = '✘ 超過 20MB';
      badgeClass = 'bg-danger';
    } else if (!matchedDoc) {
      status = 'unmatched';
      statusText = '✘ 找不到文件 ID';
      badgeClass = 'bg-danger';
    }

    batchImportItems.push({
      id: 'batch_item_' + i,
      file: file,
      matchedDoc: matchedDoc,
      targetVersion: extractedVersion,
      status: status,
      statusText: statusText,
      badgeClass: badgeClass,
      selected: (status === 'matched')
    });
  }

  renderBatchImportPreviewTable();
}

// 渲染預覽表格
function renderBatchImportPreviewTable() {
  const tbody = document.getElementById('batchImportPreviewTbody');
  const btnStart = document.getElementById('btnStartBatchImportExecution');
  if (!tbody) return;

  if (batchImportItems.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-4">請點擊上方按鈕選擇檔案以產生配對預覽...</td></tr>`;
    if (btnStart) btnStart.disabled = true;
    return;
  }

  let html = '';
  let selectableCount = 0;

  batchImportItems.forEach((item, idx) => {
    const fileSizeMB = (item.file.size / (1024 * 1024)).toFixed(2) + ' MB';
    const docInfo = item.matchedDoc
      ? `<strong>${escapeHtml(item.matchedDoc.doc_id)}</strong> <span class="text-muted">(${escapeHtml(item.matchedDoc.title || '無標題')})</span>`
      : `<span class="text-danger">未找到吻合的 doc_id</span>`;
    const statusFlow = item.matchedDoc
      ? `<span class="badge bg-secondary">${escapeHtml(item.matchedDoc.status)}</span> ➔ <span class="badge bg-success">已發布</span>`
      : `—`;

    const isSelectable = (item.status === 'matched');
    if (isSelectable && item.selected) selectableCount++;

    html += `
      <tr class="${item.selected && isSelectable ? 'table-active' : ''}" id="tr_${item.id}">
        <td class="text-center">
          <input type="checkbox" class="form-check-input batch-chk-item" data-idx="${idx}" ${item.selected ? 'checked' : ''} ${isSelectable ? '' : 'disabled'} onchange="onBatchItemCheckChanged(${idx}, this.checked)">
        </td>
        <td><span class="text-break" title="${escapeHtml(item.file.name)}">${escapeHtml(item.file.name)}</span></td>
        <td class="text-nowrap">${fileSizeMB}</td>
        <td>${docInfo}</td>
        <td class="text-nowrap">${statusFlow}</td>
        <td>
          <input type="text" class="form-control form-control-sm text-center fw-bold" value="${escapeHtml(item.targetVersion)}" ${isSelectable ? '' : 'disabled'} onchange="onBatchItemVersionChanged(${idx}, this.value)">
        </td>
        <td class="text-center">
          <span class="badge ${item.badgeClass} w-100 py-1" id="badge_${item.id}">${item.statusText}</span>
        </td>
      </tr>
    `;
  });

  tbody.innerHTML = html;
  if (btnStart) btnStart.disabled = (selectableCount === 0);
}

// 單列 Checkbox 變更
function onBatchItemCheckChanged(idx, checked) {
  if (batchImportItems[idx] && batchImportItems[idx].status === 'matched') {
    batchImportItems[idx].selected = checked;
  }
  const anySelected = batchImportItems.some(item => item.selected && item.status === 'matched');
  const btnStart = document.getElementById('btnStartBatchImportExecution');
  if (btnStart) btnStart.disabled = !anySelected;
}

// 全選 Checkbox 切換
function toggleSelectAllBatchItems(masterChk) {
  const checked = masterChk.checked;
  batchImportItems.forEach((item, idx) => {
    if (item.status === 'matched') item.selected = checked;
  });
  renderBatchImportPreviewTable();
}

// 版號文字修改
function onBatchItemVersionChanged(idx, newVer) {
  if (batchImportItems[idx]) {
    batchImportItems[idx].targetVersion = (newVer || '1.0').trim();
  }
}
```

- [ ] **Step 3: 驗證 JS 解析邏輯**

執行語法檢查確認沒有少引號或大括號。
Run: `grep -n "parseAndMatchBatchFiles" index.html`
Expected: 成功顯示定義與調用位置。

- [ ] **Step 4: Commit 智能解析與預覽表格 JS**

```bash
git add index.html
git commit -m "feat: implement smart file matching and preview table rendering for batch import"
```

---

### Task 4: 前端循序批次上傳執行控制器與非同步通訊

**Files:**
- Modify: `index.html` (新增 `startBatchImportExecution()` 與迴圈控制邏輯)

**Interfaces:**
- Consumes: `batchImportItems`, `google.script.run.withSuccessHandler().withFailureHandler().apiBatchImportDirectFile()`
- Produces: 循序非同步上傳控制器與進度通知

- [ ] **Step 1: 在 `index.html` 新增 `startBatchImportExecution` 及 Base64 轉換非同步工具**

於 `index.html` 的 JS 區域插入以下完整的非同步循序執行控制器：

```javascript
// 將 File 物件包裝為 Base64 Promise
function _readFileAsBase64Promise(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const dataUrl = e.target.result || '';
      const base64 = dataUrl.split(',')[1] || '';
      resolve(base64);
    };
    reader.onerror = function(err) {
      reject(err || new Error('檔案讀取失敗'));
    };
    reader.readAsDataURL(file);
  });
}

// 循序執行批次匯入上傳
async function startBatchImportExecution() {
  const selectedItems = batchImportItems.filter(item => item.selected && item.status === 'matched');
  if (selectedItems.length === 0) {
    alert('未選擇任何有效的待上傳對應檔案');
    return;
  }

  if (!confirm(`確定要批次匯入這 ${selectedItems.length} 份正式檔案嗎？\n匯入後對應文件將直接更新為新版號並轉為「已發布」狀態。`)) {
    return;
  }

  // UI 鎖定
  const btnStart = document.getElementById('btnStartBatchImportExecution');
  const btnCloseTop = document.getElementById('btnCloseBatchImportModalTop');
  const btnCloseBottom = document.getElementById('btnCloseBatchImportModalBottom');
  const fileInput = document.getElementById('batchImportFileInput');
  const progressArea = document.getElementById('batchImportProgressArea');
  const progressText = document.getElementById('batchImportProgressText');
  const progressBar = document.getElementById('batchImportProgressBar');
  const progressPercent = document.getElementById('batchImportProgressPercent');

  if (btnStart) btnStart.disabled = true;
  if (btnCloseTop) btnCloseTop.disabled = true;
  if (btnCloseBottom) btnCloseBottom.disabled = true;
  if (fileInput) fileInput.disabled = true;
  if (progressArea) progressArea.style.display = 'block';

  // 將表格 Checkbox 及版號輸入框禁用
  document.querySelectorAll('.batch-chk-item, #chkSelectAllBatchItems, .batch-ver-input').forEach(el => el.disabled = true);

  let successCount = 0;
  let failCount = 0;
  const total = selectedItems.length;

  for (let i = 0; i < total; i++) {
    const item = selectedItems[i];
    const currentNum = i + 1;
    const percent = Math.round((i / total) * 100);

    // 更新進度條 UI
    if (progressText) progressText.textContent = `⏳ 正在上傳第 ${currentNum} / ${total} 筆：${item.file.name}...`;
    if (progressBar) progressBar.style.width = percent + '%';
    if (progressPercent) progressPercent.textContent = percent + '%';

    // 更新該行表格狀態
    const badgeEl = document.getElementById('badge_' + item.id);
    if (badgeEl) {
      badgeEl.className = 'badge bg-warning text-dark w-100 py-1';
      badgeEl.textContent = '⏳ 上傳中...';
    }

    try {
      // 1. 讀取 Base64
      const base64 = await _readFileAsBase64Promise(item.file);
      const mimeType = item.file.type || 'application/octet-stream';

      // 2. 呼叫後端 API
      const res = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(res => resolve(res))
          .withFailureHandler(err => reject(err))
          .apiBatchImportDirectFile(item.matchedDoc.doc_id, item.file.name, base64, mimeType, item.targetVersion);
      });

      if (res && res.success) {
        successCount++;
        item.status = 'uploaded';
        if (badgeEl) {
          badgeEl.className = 'badge bg-success w-100 py-1';
          badgeEl.textContent = `✔ 生效 v${res.version}`;
        }
        // 更新本地 state 快取，使其在背景即時反映
        if (item.matchedDoc) {
          item.matchedDoc.status = '已發布';
          item.matchedDoc.version = res.version;
          item.matchedDoc.file_id = res.file_id;
          item.matchedDoc.pending_file_id = '';
          item.matchedDoc.pending_version = '';
          item.matchedDoc.pending_file_name = '';
        }
      } else {
        failCount++;
        item.status = 'failed';
        const errMsg = (res && res.error) ? res.error : '未知錯誤';
        if (badgeEl) {
          badgeEl.className = 'badge bg-danger w-100 py-1';
          badgeEl.textContent = `❌ 失敗: ${errMsg}`;
        }
      }
    } catch (err) {
      failCount++;
      item.status = 'failed';
      if (badgeEl) {
        badgeEl.className = 'badge bg-danger w-100 py-1';
        badgeEl.textContent = `❌ 失敗: ${err.message || '網路異常'}`;
      }
    }
  }

  // 全部處理完畢，更新進度條為 100%
  if (progressText) progressText.textContent = `🎉 批次處理結束！成功：${successCount} 筆，失敗：${failCount} 筆。`;
  if (progressBar) {
    progressBar.style.width = '100%';
    progressBar.classList.remove('progress-bar-animated');
    if (failCount > 0 && successCount === 0) progressBar.className = 'progress-bar bg-danger';
    else if (failCount > 0) progressBar.className = 'progress-bar bg-warning';
    else progressBar.className = 'progress-bar bg-success';
  }
  if (progressPercent) progressPercent.textContent = '100%';

  // 恢復關閉按鈕並提示
  if (btnCloseTop) btnCloseTop.disabled = false;
  if (btnCloseBottom) {
    btnCloseBottom.disabled = false;
    btnCloseBottom.className = 'btn btn-primary fw-bold';
    btnCloseBottom.textContent = '完成 / 關閉';
  }
  if (fileInput) fileInput.disabled = false;

  alert(`批次匯入處理完畢！\n成功：${successCount} 筆\n失敗：${failCount} 筆`);

  // 重新渲染背景列表畫面
  if (typeof renderDocList === 'function') renderDocList();
  if (typeof renderStats === 'function') renderStats();
}
```

- [ ] **Step 2: 驗證前端 JS 控制器與 UI 綁定**

Run: `grep -n "startBatchImportExecution" index.html`
Expected: 成功找到定義處及按鈕 `onclick="startBatchImportExecution()"` 綁定處。

- [ ] **Step 3: Commit 批次循序上傳控制器**

```bash
git add index.html
git commit -m "feat: implement sequential batch upload controller and event bindings"
```

---

## Plan Self-Review

1. **Spec coverage:** 
   - 批次直接生效正式版 (`apiBatchImportDirectFile`) 寫入試算表 M/G/J/K 及清空待核欄位 N/O/P ➔ Covered in Task 1.
   - 寫入 `FILE_VERSIONS` 工作表與 `_logAudit` ➔ Covered in Task 1.
   - 前端檔案選取、智能匹配 (`doc_id` + `version` 正則拆分) 與對位預覽表格渲染 ➔ Covered in Task 2 & Task 3.
   - 前端循序進度條與非同步單一 API 呼叫，避開 GAS 6 分鐘與 Payload 限制 ➔ Covered in Task 4.
2. **Placeholder scan:** 無 TBD/TODO 或抽象描述，每一 Task 均有完整可以直接運用的具體程式碼片段。
3. **Type consistency:** 欄位名稱 (`doc_id`, `version`, `file_id`, `status` 等) 與對應的 Google Sheets 欄位常數對照一致。

---

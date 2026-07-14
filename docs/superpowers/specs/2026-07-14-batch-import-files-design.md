# 管理員批次匯入正式檔案與自動合位系統設計規格 (V6 Extension)

- **日期**：2026-07-14
- **作者**：Antigravity & User
- **狀態**：Approved for Implementation

---

## 1. 專案背景與目的

Google Apps Script (GAS) 文件管理系統目前在 V6 架構下支援單一檔案的上傳與版本控管（前台選擇小改版／大改版／自訂起始版號，進入 `pending_*` 待核狀態，再經由管理員審核發布 promoter 為正式版）。

當管理員面對一批已在試算表「文件清單」中建立基本屬性（如 `doc_id`, `title`, `category` 等）但尚未掛載實體檔案的既有或歷史文件時，逐點擊進各單筆文件詳情頁進行上傳效率低下。

**本專案目的**在於建立一套**管理員專用的「批次匯入正式檔案與自動對位系統」**：
1. 透過前端多檔案選取與智能解析檔名（提取 `doc_id` 與 `version`），自動對應現有試算表中的文件。
2. 讓管理員在上傳前透過「對位預覽表格」檢核與修正辨識的版號。
3. 採用前端進度條與「循序呼叫後端上傳 API」機制，避開 GAS 單次 Payload 過大與 6 分鐘執行超時限制。
4. 匯入後檔案直接生效為現行正式版（寫入 `file_id` 與 `version`）、將狀態自動設定為「已發布」，並寫入「檔案版本 (`FILE_VERSIONS`)」與異動審計紀錄。

---

## 2. 系統架構與業務規則

### 2.1 角色與授權
- **僅限管理員 (`isAdmin`)** 能夠存取批次匯入 UI 按鈕並呼叫 `apiBatchImportDirectFile` 後端 API。
- 若一般使用者嘗試呼叫 API，後端將透過 `_assertAdmin()` 直接阻斷並拋出例外。

### 2.2 檔名解析規則與智能配對
常規檔名格式範例：`TWHB-ISMS-002-001_資訊安全風險評鑑管理程序書_V2.1.pdf`

前端解析演算法邏輯：
1. **文件配對 (`doc_id` 與 `title`)**：
   - 遍歷目前有權限載入的所有文件 `allDocs`。
   - 對每一個選取的檔案，驗證 `file.name` 是否以 `doc.doc_id` 開頭且其後緊接 `_`、`-` 或空白 (即 `new RegExp('^' + _escapeRegExp(doc.doc_id) + '[_\\-\\s]', 'i')`)。若精準吻合，即視為對應成功。
2. **版本號提取 (`version`)**：
   - 使用正則表達式自檔名尾部提取版本號：`/[_-\s][vV]([0-9]+(?:\.[0-9]+)*)(?:[_-\s][^.]+)?\.[^.]+$/i`。
   - 例如從 `_V2.1.pdf` 或 `_v2.1_最終版.pdf` 中成功提取出 `2.1`。
   - 若檔名不符合該格式或未提取成功，預設採用該文件目前試算表上的版號或 `1.0`，同時在預覽表格標示，供管理員直接在表格輸入框內微調。

### 2.3 狀態與版本衝突覆寫規則
當管理員勾選確認匯入配對成功的檔案後，後端執行以下覆寫與更新：
1. **檔案覆寫**：無論該文件原本是否已有 `file_id` 或 `pending_file_id`，一律將新上傳之檔案寫入為該文件的現行生效版 `file_id` (M欄) 與 `version` (G欄)。
2. **清空待核欄位**：將 `pending_file_id` (N欄)、`pending_version` (O欄)、`pending_file_name` (P欄) 設為空字串，清除所有待核殘留狀態。
3. **自動轉為已發布**：若文件原有 `status` (D欄) 不為「已發布」，自動將該欄位變更為「已發布」；同時自動設定發布日 `published_at` (J欄) 為當日，並根據 `review_cycle` (L欄) 重算下次審查日 `next_review` (K欄)。
4. **歷史與軌跡同步**：同步在「檔案版本 (`FILE_VERSIONS`)」工作表新增一筆正式版本紀錄，供歷史版本查閱或代理下載。

---

## 3. 後端詳細規格 (`code.js`)

### 3.1 新增 API 函式：`apiBatchImportDirectFile`
```javascript
/**
 * 管理員批次匯入直接生效之單檔處理 API
 * @param {string} docId - 對應的文件 ID (如 'DOC-001' 或 'TWHB-ISMS-002-001')
 * @param {string} fileName - 原始檔案名稱
 * @param {string} base64 - 檔案內容的 Base64 字串
 * @param {string} mimeType - 檔案 MIME 類型
 * @param {string} targetVersion - 目標生效版本號 (如 '2.1')
 * @returns {Object} { success: true/false, docId, version, status, file_id, error? }
 */
function apiBatchImportDirectFile(docId, fileName, base64, mimeType, targetVersion)
```

**內部執行流程**：
1. 取得 `LockService.getScriptLock()` 並等待最長 10 秒。
2. 呼叫 `_assertAdmin()` 檢驗權限。
3. 讀取並定位試算表「文件清單 (`DOCS`)」中符合 `doc_id === docId` 的資料列，若不存在或為「已廢止」狀態則回傳 `{ success: false, error: '...' }`。
4. 將 Base64 解碼並寫入 Google Drive 中該文件的專屬資料夾（呼叫 `_getOrCreateDocFolder(docId)`），儲存檔名格式為：`${docId}_v${targetVersion}_${safeName}`。
5. 更新「文件清單」試算表對應列欄位：
   - `status` (D欄) = `'已發布'`
   - `version` (G欄) = `targetVersion`
   - `published_at` (J欄) = 今日 (`_now()`)
   - `next_review` (K欄) = 今日 + `review_cycle` 個月 (`_addMonthsFromToday(cycle)`)
   - `file_id` (M欄) = `file.getId()`
   - `pending_file_id` (N欄) = `''`
   - `pending_version` (O欄) = `''`
   - `pending_file_name` (P欄) = `''`
6. 將本次生效檔案寫入「檔案版本 (`FILE_VERSIONS`)」工作表末端。
7. 呼叫 `_logAudit('批次匯入', docId, targetVersion, ...)` 記錄異動。
8. 執行 `SpreadsheetApp.flush()` 後釋放鎖定並回傳成功結果。

---

## 4. 前端詳細規格 (`index.html`)

### 4.1 按鈕與對話框佈局
- 在列表頁上方操作按鈕區塊，若 `state.user.isAdmin === true`，動態渲染「<i class="fas fa-file-import"></i> 批次匯入檔案」按鈕。
- 點擊按鈕開啟專屬 Modal `batchImportModal`：
  - **檔案選擇器**：`<input type="file" id="batchImportFileInput" multiple class="form-control">`
  - **預覽表格 (`#batchImportPreviewTable`)**：
    - 欄位包含：選取勾選 (`<input type="checkbox">`)、原檔名、匹配對應文件（顯示 `doc_id` + `title`）、原本狀態 ➔ `已發布`、目標版本號（文字框 `<input type="text" class="batch-ver-input">`）、驗證 Badge (`✔ 對應成功` / `✘ 找不到文件 ID`)。
  - **進度區塊 (`#batchImportProgressArea`)**：
    - 進度條 `<div class="progress">...</div>` 與訊息計數提示 `<div id="batchImportProgressText">已完成 0 / 10 筆...</div>`。
  - **控制按鈕**：`「確認並開始匯入」` 與 `「關閉」` 按鈕。

### 4.2 前端循序批次上傳控制器
當點擊「確認並開始匯入」：
1. 蒐集預覽表格中所有已被勾選且 `status === 'matched'` 的行資料清單。
2. 將按鈕設定為禁用 (Disabled)，並顯示進度條 (`progressArea.style.display = 'block'`)。
3. 採用 `for...of` 異步迴圈，逐一對該清單進行處理：
   - 使用 `FileReader` 將該 `File` 物件轉為 Base64 字串。
   - 透過 `google.script.run.withSuccessHandler(...).withFailureHandler(...).apiBatchImportDirectFile(docId, file.name, base64, file.type, targetVersion)` 進行單一上傳。
   - 在表格中把當前行標示為 `「⏳ 匯入中...」`，成功後改為 `「✔ 匯入完成 (已發布 v2.1)」`。
   - 若發生失敗，該行標示 `「❌ 失敗：原因」`，計數器計入失敗次數，並繼續執行下一筆（容錯不中斷）。
4. 處理完全部佇列後，彈出總結提示並在用戶點擊或關閉視窗後，呼叫 `loadInitialData()`（或重載頁面）以刷新文件清單狀態。

---

## 5. 錯誤處理與邊界案例

1. **網路連線或單檔上傳失敗**：
   - 因為前端是逐一發起獨立的 Request，其中單一個大檔案或網路逾時不會影響已成功的檔案，且失敗品會在預覽表清楚標示，方便管理員針對失敗檔案稍後重試。
2. **大於 20MB 的單一超大檔案**：
   - 沿用目前系統中單檔 20MB 的安全檢查，若選取的單個檔案大於 20MB，前端在產生預覽列時立刻標示警告並禁用該筆勾選，避免呼叫後端導致 Base64 記憶體溢出。
3. **沒有可匹配的 `doc_id`**：
   - 自動在預覽表標示紅色 `✘ 找不到對應文件`，勾選框設為不可點選，保護資料不遺失或錯置。

---

## 6. 驗收檢查清單 (Acceptance Criteria)

- [ ] 管理員身分可見並開啟「批次匯入檔案」對話框；一般使用者不可見。
- [ ] 一次選擇多個例如 `TWHB-ISMS-002-001_程序書_V2.1.pdf` 檔案時，前端預覽表自動準確命中 ID `TWHB-ISMS-002-001` 並自動填充版本框 `2.1`。
- [ ] 管理員可在開始上傳前於預覽表中修改特定檔案的「目標版本號」。
- [ ] 上傳執行時進度條正確移動，成功完成後，對應文件之 `status` 在試算表與前台皆顯示為「已發布」，`file_id` 與 `version` 皆更新為新版本，且待核欄位皆已清空。
- [ ] 下載及版本歷史功能皆可正常讀取本次批次匯入的檔案。

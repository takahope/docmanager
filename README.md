# 文件管理系統（Closure Table 架構）

Google Apps Script 全端文件管理系統，使用 Google Sheets 作為後端資料庫，採用 Closure Table 設計處理文件間的多層關聯。

## 檔案結構

| 檔案 | 職責 |
|---|---|
| `env.js` | 環境設定、工作表名稱、欄位索引集中定義 |
| `deploy.js` | 工作表初始化（`deployAllSheets()`）與範例資料（`seedSampleData()`） |
| `code.js` | 核心業務邏輯：文件 CRUD、Closure Table 維護、前端 API |
| `index.html` | 前端主頁面（CSS/JS 內嵌單檔） |

## 部署步驟

1. 建立新的 Google Spreadsheet
2. 開啟「擴充功能 → Apps Script」
3. 依序建立四個檔案，貼入對應內容：
   - `env.gs`（貼入 env.js 內容）
   - `deploy.gs`（貼入 deploy.js 內容）
   - `code.gs`（貼入 code.js 內容）
   - `index.html`（HTML 檔案，新增時選「HTML」類型）
4. 在編輯器中手動執行 `deployAllSheets()` → 授權 → 建立兩張工作表
5. （選用）執行 `seedSampleData()` 寫入範例資料
6. 「部署 → 新增部署作業 → 網頁應用程式」
   - 執行身分：我
   - 存取權限：依需求設定（機構內部建議「僅限機構內使用者」）
7. 開啟部署網址即可使用

## 資料表結構

### 工作表：文件清單
| 欄位 | 說明 |
|---|---|
| doc_id | 文件編號（DOC-001 格式，自動產生，文字格式） |
| title | 標題 |
| category | 類別（ISMS / PIMS / 表單 / SOP / 政策 / 指引） |
| status | 狀態（草稿 / 審核中 / 已發布 / 已廢止） |
| owner | 負責人 |
| owner_ID | 負責人 ID |
| updated_at | 更新日期（yyyy/MM/dd） |
| version | 版本 |
| google_drive_location | Drive 連結 |

### 工作表：文件關聯（Closure Table）
| 欄位 | 說明 |
|---|---|
| doc_id | 祖先文件（ancestor） |
| descendant_id | 子孫文件（descendant） |
| depth | 層距（0=自身、1=直接、2+=間接） |
| relation_type | 關聯類型（references / supersedes / derived_from / related） |
| 說明 | 備註 |

## Closure Table 維護邏輯

### 新增關聯 A → B（apiAddRelation）
1. 循環檢查：若 B 已（直接或間接）關聯到 A，拒絕建立
2. 笛卡兒積寫入：對每個 (X → A) × (B → Y) 組合，寫入 X → Y，depth = dX + 1 + dY
3. 這樣 A 的所有祖先都能直接查到 B 與 B 的子孫

### 刪除關聯 A → B（apiRemoveRelation）
1. 刪除影響範圍內所有 depth > 0 的路徑
2. 從剩餘 depth=1 直接邊以 BFS 重算間接路徑（處理多路徑情況）

### 刪除文件（apiDeleteDoc）
同步刪除閉包表中所有 ancestor 或 descendant 為該文件的列。

## 主要 API

| 函式 | 用途 |
|---|---|
| `apiGetInitData()` | 取得所有文件與選項清單 |
| `apiCreateDoc(doc)` | 新增文件（自動寫入 self closure） |
| `apiUpdateDoc(doc)` | 更新文件 |
| `apiDeleteDoc(docId)` | 刪除文件（連動清除關聯） |
| `apiGetDescendants(docId, maxDepth)` | 查此文件引用了誰（含間接） |
| `apiGetAncestors(docId)` | 反查誰引用了此文件 |
| `apiAddRelation(anc, desc, type, note)` | 建立關聯（自動維護閉包） |
| `apiRemoveRelation(anc, desc)` | 移除關聯（自動重算閉包） |
| `apiGetGraphData()` | 取得節點與直接邊（供關聯圖擴充使用） |

## 已套用的 GAS 最佳實踐

- 所有欄位索引集中在 `env.js`，禁止硬編碼 `row[8]`
- 寫入後強制 `SpreadsheetApp.flush()`
- doc_id 與 owner_ID 欄位強制文字格式（`@`），避免被轉數字
- 讀取使用 `getDisplayValues()` 避免日期序列化為 null
- 高並發寫入使用 `LockService` 防撞
- 前端所有使用者輸入經 `escapeHtml()` 防 XSS
- 搜尋使用 debounce 防止大量資料卡頓

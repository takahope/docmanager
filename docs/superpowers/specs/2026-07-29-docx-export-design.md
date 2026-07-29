# Docx Export Feature Design

## Overview
在「文件管理系統 (GAS)」中新增一個「匯出文件一覽表」的功能。使用者可以選擇特定的標籤 (Tag)，系統會根據該 Tag 中的文件及其關聯，依照「父文件 (主文件) - 子文件 (表單)」的階層結構，生成一份與原 docx 範本格式一致的 Word 檔案。

## Architecture
1. **Frontend**: 在 `index.html` 增加「匯出 Word」按鈕與 Modal，負責觸發匯出流程，並利用 `docxtemplater`, `pizzip`, `FileSaver.js` CDN 套件在客戶端生成與下載 `.docx` 檔案。
2. **Backend (GAS)**: 新增一支 API `apiGetDocxExportData`，負責過濾特定 Tag 的文件、計算父子關聯，並將 Google Drive 上的範本檔案轉為 Base64。
3. **Data Storage**: Word 範本檔案存放在 Google Drive，其 File ID 記錄於 `env.js` 的 `PROP_KEYS` (例如 `DOCX_TEMPLATE_FILE_ID`)，讓系統能動態抓取。

## Data Flow & Processing Logic
1. **取得文件**: 依據 `tagId`，過濾出該 Tag 中使用者有權限瀏覽的所有文件。
2. **父子關係判定**: 
   - 找出該 Tag 內的所有關聯 (基於 Closure Table, 僅限 depth = 1 的直接關聯)。
   - **子文件 (表單)**: 如果一份文件在此關聯中作為 descendant (子)，且 ancestor (父) 也在這群文件中，它就是某個主文件的附屬表單。
   - **主文件 (父)**: 未被群內其他文件當作子文件的，即為主文件。
3. **列平坦化 (Flattening for Table Rows)**:
   針對每個主文件：
   - 若有 N 個表單，產生 N 列資料。
     - 第 1 列：填入主文件所有資訊 (doc_id, title, category, version, published_at)，以及第 1 個表單的資訊 (form_id, form_title, form_version, form_published_at)。
     - 第 2~N 列：主文件資訊留白 (為保持視覺上如合併儲存格般清爽)，填入第 2~N 個表單的資訊。
   - 若無表單，產生 1 列資料，表單資訊留白。
4. **傳回資料**: 將此陣列與範本 Base64 傳回前端。

## Template Variables
Word 範本內的表格僅需設計一列，並包在迴圈 `{#docs}` 和 `{/docs}` 內。支援以下欄位：
- `{{doc_id}}`: 文件編號 (父)
- `{{title}}`: 文件名稱 (父)
- `{{category}}`: 機密等級 (父)
- `{{version}}`: 文件版本 (父)
- `{{published_at}}`: 發行日期 (父)
- `{{form_id}}`: 表單編號 (子)
- `{{form_title}}`: 表單名稱 (子)
- `{{form_version}}`: 表單版本 (子)
- `{{form_published_at}}`: 表單發行日期 (子)

## UI Changes
- **Tool Bar**: 在 `tabPage-list` 增加按鈕「⬇ 匯出文件一覽表(Word)」。
- **Modal**: 增加 `docxExportModal`，提供 Tag 選擇下拉選單與確認按鈕。
- **Loading**: 利用現有 `loading` 遮罩提示「正在產生文件，請稍候...」。

## Error Handling
- 若 GAS 取不到範本檔案 (File ID 錯誤或無權限)，拋出明確錯誤讓前端以 Toast 顯示。
- 若特定 Tag 內無任何文件，提示「該標籤下無任何文件」。

## Testing
- 測試一個沒有表單的主文件。
- 測試一個具有多個表單的主文件，確認第二列的主文件欄位為空白。
- 測試某文件同時身兼其他人的子文件時，是否只出現在表單欄位中。

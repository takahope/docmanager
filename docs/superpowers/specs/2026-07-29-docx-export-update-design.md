# Docx Export Feature Update Design

## Overview
本次更新基於現有的「文件一覽表匯出」功能，加入範本自動填寫日期與系統編號的規則（參考 `createdoc.md`）。
架構採用「混合模式 (Hybrid)」：由後端負責掃描 Google Drive 目錄計算流水號，並交由前端原本強大的 `docxtemplater` 引擎進行字串與表格的替換。最終產生的 Word 檔會先回傳至後端存檔，再提供本地下載。

## 1. Environment & Configuration (env.js)
需要在 `PROP_KEYS` 或常數設定中增加以下項目，以便動態調整且不寫死在程式碼中：
- `DOCX_OUTPUT_FOLDER_ID`: 產出的 Word 檔要存放的目標 Google Drive 資料夾 ID。
- `RECORD_NUMBER_PREFIX`: 流水號前綴字串（例如預設為 `IS-R-032`）。

## 2. Backend Logic (code.js)

### 2.1 產生流水號與獲取匯出資料
修改現有的 `apiGetDocxExportData`，除了過濾資料與轉 Base64 外，新增以下處理：
1. **取得現在時間**：根據時區取得目前的西元年 (4 碼)、月 (2 碼)、日 (2 碼)，以及用於正規表達式比對的日期鍵值 (yyyyMMdd)。
2. **掃描資料夾計算流水號**：
   - 到 `DOCX_OUTPUT_FOLDER_ID` 資料夾下，抓取所有檔名。
   - 使用正規表示式 `/^{PREFIX}-{yyyyMMdd}-(\d+)/` 找出當日最大流水號 `maxSerial`。
   - 產生下一個序號，若小於 100 則補零。
   - 組合出 `recordNo`（例如 `IS-R-032-20260729-01`）。
3. **回傳資料格式擴充**：
   ```json
   {
     "templateBase64": "...",
     "data": [...],
     "year": "2026",
     "month": "07",
     "day": "29",
     "recordNo": "IS-R-032-20260729-01"
   }
   ```

### 2.2 存檔 API
新增一支 `apiSaveDocxToDrive(base64Data, fileName)` 函數：
- 接收前端傳來的 Base64 字串，轉換為 `Blob`。
- 設定 MIME Type 為 `application/vnd.openxmlformats-officedocument.wordprocessingml.document`。
- 儲存至 `DOCX_OUTPUT_FOLDER_ID` 目錄中，並回傳存檔成功的檔案 ID 或網址。

## 3. Frontend Logic (index.html)

### 3.1 渲染資料擴充
在呼叫 `doc.render()` 時，加入新的全域變數提供給範本使用：
```javascript
doc.render({
  年: result.year,
  月: result.month,
  日: result.day,
  紀錄編號: result.recordNo,
  docs: result.data
});
```
*這對應了 `createdoc.md` 提到的 `{{年}}`, `{{月}}`, `{{日}}`, `{{紀錄編號}}` 標記。*

### 3.2 存檔與下載流程 (`generateDocx` 函數修改)
1. 產生 `.docx` 的二進位檔案 (Blob) 之後，不立刻呼叫 `saveAs`。
2. 透過 `FileReader` 將 Blob 轉為 Base64 字串。
3. 顯示 Loading UI：「正在存檔至 Google Drive...」。
4. 呼叫 `google.script.run.apiSaveDocxToDrive(base64, fileName)`。
5. **成功回呼 (Success Handler)**：收到成功回應後，再呼叫 `saveAs(blob, fileName)` 提供本地下載，並隱藏 Loading UI。
6. **失敗回呼 (Failure Handler)**：提示存檔失敗，且基於容錯考量，可以詢問或直接提供本地下載以免資料遺失。

## 4. Error Handling
- 如果 `DOCX_OUTPUT_FOLDER_ID` 未設定或無存取權限，流水號將無法計算。需在 `apiGetDocxExportData` 中攔截錯誤並回傳明確的失敗訊息。
- 若前端產生完檔案但 `apiSaveDocxToDrive` 上傳逾時，提示使用者存檔失敗。

## 5. Testing Criteria
- 測試匯出時，Word 範本的首頁（或頁首尾）是否能成功替換 `{{年}}`, `{{月}}`, `{{日}}`, `{{紀錄編號}}`。
- 檢查指定的 Google Drive 資料夾是否有出現剛生成的檔案。
- 測試連續匯出兩次，確認流水號是否有正確遞增（例如 `-01` 變為 `-02`）。

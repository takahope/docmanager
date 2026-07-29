# Docx Export Update Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 為既有之 Docx 匯出功能加上 Google Drive 自動存檔與系統單號產生邏輯。

**Architecture:** 混合架構。後端負責去指定資料夾掃描以計算當日單號並取得系統日期；前端以此資料渲染 `docxtemplater`，轉出 Base64 後呼叫後端 API 存進 Drive，成功後再從本地下載。

**Tech Stack:** Google Apps Script, JavaScript, docxtemplater, FileSaver.js

## Global Constraints

- 不可修改既有的權限檢查邏輯 (`_assert*`)
- 變數與函式命名應遵循專案原本習慣（如後端私有函式加上底線 `_` 開頭）

---

### Task 1: 設定檔擴充 (env.js)

**Files:**
- Modify: `env.js`

**Interfaces:**
- Produces: `PROP_KEYS.DOCX_OUTPUT_FOLDER_ID`, `PROP_KEYS.RECORD_NUMBER_PREFIX`

- [ ] **Step 1: 新增 PROP_KEYS 定義**

在 `env.js` 的 `PROP_KEYS` 區塊中加入 `DOCX_OUTPUT_FOLDER_ID` 與 `RECORD_NUMBER_PREFIX`。

```javascript
const PROP_KEYS = {
  HR_SPREADSHEET_ID:   'HR_SPREADSHEET_ID',
  ADMIN_EMAILS:        'ADMIN_EMAILS',
  DOC_FILES_FOLDER_ID: 'DOC_FILES_FOLDER_ID',
  DOCX_TEMPLATE_FILE_ID: 'DOCX_TEMPLATE_FILE_ID',
  DOCX_OUTPUT_FOLDER_ID: 'DOCX_OUTPUT_FOLDER_ID',
  RECORD_NUMBER_PREFIX: 'RECORD_NUMBER_PREFIX'
};
```

- [ ] **Step 2: Commit**

```bash
git add env.js
git commit -m "feat: add DOCX_OUTPUT_FOLDER_ID and RECORD_NUMBER_PREFIX to PROP_KEYS"
```

---

### Task 2: 後端邏輯擴充 (code.js)

**Files:**
- Modify: `code.js`

**Interfaces:**
- Produces: `apiSaveDocxToDrive(base64Data, fileName)`, expanded `apiGetDocxExportData` response

- [ ] **Step 1: 新增輔助函式 `_escapeRegExp` 與 `_createRecordNoFromFolder`**

在 `code.js` 底部加入以下兩個輔助函式：

```javascript
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
```

- [ ] **Step 2: 新增存檔 API `apiSaveDocxToDrive`**

在 `code.js` 底部加入：

```javascript
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
```

- [ ] **Step 3: 修改 `apiGetDocxExportData` 回傳資料**

在 `code.js` 的 `apiGetDocxExportData` 內，原本的 `return { templateBase64, data };` 之前，新增日期與編號的產生邏輯。注意需讀取 `PROP_KEYS` 的變數。

尋找 `apiGetDocxExportData` 內的：
```javascript
  return {
    templateBase64: templateBase64,
    data: flattenedData
  };
```
替換為：
```javascript
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
```

- [ ] **Step 4: Commit**

```bash
git add code.js
git commit -m "feat: add apiSaveDocxToDrive and generate recordNo in code.js"
```

---

### Task 3: 前端儲存與下載邏輯 (index.html)

**Files:**
- Modify: `index.html`

- [ ] **Step 1: 修改 `generateDocx` 加入動態變數與上傳邏輯**

在 `index.html` 中，將 `generateDocx` 的參數從 `(templateBase64, data, tagName)` 改為接收完整 `result`，並加入 Blob 轉 Base64 上傳 Google Drive 的流程。

找到原本的 `generateDocx` 整個函式替換成：

```javascript
function generateDocx(result, tagName) {
  try {
    const arrayBuffer = base64ToArrayBuffer(result.templateBase64);
    const zip = new PizZip(arrayBuffer);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    
    doc.render({
      年: result.year || '',
      月: result.month || '',
      日: result.day || '',
      紀錄編號: result.recordNo || '',
      docs: result.data
    });
    
    const blob = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    
    const recordPrefix = result.recordNo ? `${result.recordNo}_` : '';
    const dateStr = new Date().toISOString().split('T')[0];
    const fileName = `${tagName}_資訊安全暨個人資料管理文件一覽表_${recordPrefix}${dateStr}.docx`;

    // 準備上傳至雲端
    document.getElementById('loading').classList.remove('hide');
    document.getElementById('loading').innerText = '正在存檔至 Google Drive...';

    const reader = new FileReader();
    reader.onload = function() {
      // FileReader 讀出的資料為 data:MIME;base64,... 我們只需要 base64 部分
      const base64Data = reader.result.split(',')[1];
      
      google.script.run
        .withSuccessHandler(function(saveRes) {
          document.getElementById('loading').classList.add('hide');
          showToast('存檔成功，即將開始下載！');
          saveAs(blob, fileName);
        })
        .withFailureHandler(function(err) {
          document.getElementById('loading').classList.add('hide');
          console.error(err);
          // 容錯：若雲端存檔失敗，仍提供本地下載以免白費
          showToast('雲端存檔失敗：' + err.message + '。仍會為您下載檔案。');
          saveAs(blob, fileName);
        })
        .apiSaveDocxToDrive(base64Data, fileName);
    };
    reader.onerror = function() {
      document.getElementById('loading').classList.add('hide');
      showToast('讀取檔案失敗，無法存檔。');
    };
    
    // 將 blob 讀取為 DataURL 以便取得 Base64
    reader.readAsDataURL(blob);

  } catch (error) {
    console.error(error);
    document.getElementById('loading').classList.add('hide');
    showToast('文件產生失敗，請檢查範本格式。');
  }
}
```

- [ ] **Step 2: 修改 `execDocxExport` 中呼叫 `generateDocx` 的部分**

在 `execDocxExport` 內，原本的：
```javascript
      generateDocx(result.templateBase64, result.data, tagName);
```
改為：
```javascript
      generateDocx(result, tagName);
```

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: generate docx with date and recordNo, save to Drive before download"
```

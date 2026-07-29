# Native Docx Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Migrate document export from frontend `docxtemplater` to a native Google Apps Script backend using `DocumentApp`.

**Architecture:** The backend will gather closure data, construct a 2D array, duplicate a Google Doc template, inject the array, and return the new document URL. Frontend drops three large CDN dependencies and simply opens the returned URL in a new tab.

**Tech Stack:** Google Apps Script (`DocumentApp`, `DriveApp`), HTML/JS (vanilla frontend).

## Global Constraints

- Never hardcode column indices; use constants from `env.js`.
- Use `node --check <file>.js` for syntax validation (no test runner exists).
- Frontend inputs/updates must remain optimistic or use `google.script.run` with success/failure handlers.

---

### Task 1: Add Helper Functions to Backend

**Files:**
- Modify: `/Users/kih/Desktop/docmanager/code.js`

**Interfaces:**
- Produces: `escapeRegExp_`, `createRecordNoFromFolder_`, `replaceTemplateTokens_` (internal to `code.js`)

- [ ] **Step 1: Implement `escapeRegExp_` and `createRecordNoFromFolder_`**

Add at the bottom of `code.js` (before any module.exports if they existed, but this is GAS so just global scope):

```javascript
/**
 * 轉義正規表示式特殊字元
 */
function escapeRegExp_(value) {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * 根據資料夾現有檔案自動生成流水編號
 */
function createRecordNoFromFolder_(folder, prefix, dateKey) {
  const escapedPrefix = escapeRegExp_(prefix);
  const pattern = new RegExp(escapedPrefix + '-' + dateKey + '-(\\d+)');

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

- [ ] **Step 2: Implement `replaceTemplateTokens_`**

Add to `code.js`:

```javascript
/**
 * 替換 Google Doc 中所有區塊（正文/頁首/頁尾）的 {{佔位符}}
 */
function replaceTemplateTokens_(doc, tokenMap) {
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);
  const keys = Object.keys(tokenMap || {});

  sections.forEach(function(section) {
    keys.forEach(function(key) {
      const pattern = '\\{\\{\\s*' + escapeRegExp_(key) + '\\s*\\}\\}';
      section.replaceText(pattern, String(tokenMap[key]));
    });
  });
}
```

- [ ] **Step 3: Syntax check**

Run: `node --check /Users/kih/Desktop/docmanager/code.js`
Expected: Passes without syntax errors.

- [ ] **Step 4: Commit**

```bash
git add /Users/kih/Desktop/docmanager/code.js
git commit -m "feat(export): add helper functions for native Google Doc generation"
```

---

### Task 2: Implement Main API `apiExportNativeDocument`

**Files:**
- Modify: `/Users/kih/Desktop/docmanager/code.js`

**Interfaces:**
- Consumes: `_assertCanViewDoc`, `_getVisibleDocIds`, `apiGetDescendants`, the helpers from Task 1.
- Produces: `apiExportNativeDocument(tagId)` returning `{success: true, url: string, recordNo: string}`

- [ ] **Step 1: Replace old `apiGetDocxExportData` with new logic**

Locate `apiGetDocxExportData(tagId)` in `code.js` and completely replace it with `apiExportNativeDocument(tagId)`. Retain the initial data-fetching logic, but format into a 2D array and generate the doc:

```javascript
function apiExportNativeDocument(tagId) {
  const ctx = _assertWhitelisted();
  const visibleDocIds = _getVisibleDocIds(ctx);

  const docSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DOCS);
  const docData = docSheet.getDataRange().getDisplayValues();
  const docHeaders = docData[0];
  
  const closureSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CLOSURE);
  const closureData = closureSheet.getDataRange().getDisplayValues();
  const closureHeaders = closureData[0];
  
  const docTagsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DOC_TAGS);
  const docTagsData = docTagsSheet.getDataRange().getDisplayValues();
  
  let targetDocIds = new Set();
  for (let i = 1; i < docTagsData.length; i++) {
    if (docTagsData[i][1] === tagId && visibleDocIds.has(docTagsData[i][0])) {
      targetDocIds.add(docTagsData[i][0]);
    }
  }

  const allDocs = {};
  for (let i = 1; i < docData.length; i++) {
    const id = docData[i][DOC_COL.doc_id];
    if (visibleDocIds.has(id)) {
      allDocs[id] = {
        doc_id: id,
        title: docData[i][DOC_COL.title],
        category: docData[i][DOC_COL.category],
        security_level: docData[i][DOC_COL.security_level] || "",
        version: docData[i][DOC_COL.version],
        published_at: docData[i][DOC_COL.published_at]
      };
    }
  }

  const tableData = [
    ['文件編號', '文件名稱', '機密等級', '版本', '發行日期', '表單編號', '表單名稱', '表單版本', '表單發行日期']
  ];

  for (const parentId of targetDocIds) {
    if (!allDocs[parentId]) continue;
    const parentDoc = allDocs[parentId];
    let childrenIds = [];
    
    for (let i = 1; i < closureData.length; i++) {
      if (closureData[i][CLS_COL.doc_id] === parentId && 
          closureData[i][CLS_COL.depth] === "1" && 
          closureData[i][CLS_COL.relation_type] === 'related' &&
          visibleDocIds.has(closureData[i][CLS_COL.descendant_id])) {
        childrenIds.push(closureData[i][CLS_COL.descendant_id]);
      }
    }

    if (childrenIds.length === 0) {
      tableData.push([
        parentDoc.doc_id, parentDoc.title, parentDoc.security_level, parentDoc.version, parentDoc.published_at,
        "", "", "", ""
      ]);
    } else {
      childrenIds.forEach((childId, index) => {
        const childDoc = allDocs[childId];
        if (index === 0) {
          tableData.push([
            parentDoc.doc_id, parentDoc.title, parentDoc.security_level, parentDoc.version, parentDoc.published_at,
            childDoc.doc_id, childDoc.title, childDoc.version, childDoc.published_at
          ]);
        } else {
          tableData.push([
            "", "", "", "", "",
            childDoc.doc_id, childDoc.title, childDoc.version, childDoc.published_at
          ]);
        }
      });
    }
  }

  const templateId = _getProp(PROP_KEYS.DOCX_TEMPLATE_FILE_ID);
  if (!templateId) {
    throw new Error('系統尚未設定 Docx 範本檔案 (DOCX_TEMPLATE_FILE_ID)。');
  }

  const outputFolderId = _getProp(PROP_KEYS.DOCX_OUTPUT_FOLDER_ID);
  if (!outputFolderId) {
    throw new Error('系統尚未設定 Docx 輸出資料夾 (DOCX_OUTPUT_FOLDER_ID)。');
  }

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const templateFile = DriveApp.getFileById(templateId);

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const year = Utilities.formatDate(now, tz, 'yyyy');
  const month = Utilities.formatDate(now, tz, 'MM');
  const day = Utilities.formatDate(now, tz, 'dd');
  const dateKey = Utilities.formatDate(now, tz, 'yyyyMMdd');

  const prefix = _getProp(PROP_KEYS.RECORD_NUMBER_PREFIX) || 'IS-R-032';
  const recordNo = createRecordNoFromFolder_(outputFolder, prefix, dateKey);

  const newFileName = prefix + '_' + recordNo;
  const copiedFile = templateFile.makeCopy(newFileName, outputFolder);
  
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();

  replaceTemplateTokens_(doc, {
    '年': year,
    '月': month,
    '日': day,
    '紀錄編號': recordNo
  });

  if (tableData.length > 1) {
    const found = body.findText('\\{\\{\\s*表格\\s*\\}\\}');
    let table;
    if (found) {
      const textElement = found.getElement().asText();
      textElement.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
      let paragraph = textElement.getParent();
      while (paragraph && paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) {
        paragraph = paragraph.getParent();
      }
      const index = paragraph ? body.getChildIndex(paragraph) : body.getNumChildren();
      table = body.insertTable(index + 1, tableData);
      if (paragraph && !paragraph.asParagraph().getText().trim()) {
        body.removeChild(paragraph);
      }
    } else {
      table = body.appendTable(tableData);
    }

    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(6).setPaddingRight(6);
        if (r === 0) {
          cell.setBackgroundColor('#e5e7eb');
          cell.editAsText().setBold(true).setForegroundColor('#111827');
        }
      }
    }
  }

  doc.saveAndClose();

  return {
    success: true,
    url: doc.getUrl(),
    recordNo: recordNo
  };
}
```

- [ ] **Step 2: Syntax check**

Run: `node --check /Users/kih/Desktop/docmanager/code.js`
Expected: Passes without syntax errors.

- [ ] **Step 3: Commit**

```bash
git add /Users/kih/Desktop/docmanager/code.js
git commit -m "feat(export): replace legacy export with apiExportNativeDocument for DocumentApp"
```

---

### Task 3: Refactor Frontend

**Files:**
- Modify: `/Users/kih/Desktop/docmanager/index.html`

**Interfaces:**
- Consumes: `apiExportNativeDocument`

- [ ] **Step 1: Remove external libraries**

In `index.html`, find and delete the following lines from the `<head>`:
```html
<script src="https://unpkg.com/pizzip@3.1.7/dist/pizzip.js"></script>
<script src="https://unpkg.com/pizzip@3.1.7/dist/pizzip-utils.js"></script>
<script src="https://unpkg.com/docxtemplater@3.52.0/build/docxtemplater.js"></script>
<script src="https://unpkg.com/file-saver@2.0.5/dist/FileSaver.min.js"></script>
```

- [ ] **Step 2: Remove old generator logic**

In `index.html`, locate the block:
```javascript
// ============================================================
// Docx Export Logic
// ============================================================
function base64ToArrayBuffer(base64) {
```
And delete the functions `base64ToArrayBuffer` and `generateDocx(result)`.

- [ ] **Step 3: Update `exportDocx()`**

Replace the current `exportDocx(tagId, tagName)` function with:

```javascript
function exportDocx(tagId, tagName) {
  const btn = document.getElementById(`export-btn-${tagId}`);
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<i class="bi bi-hourglass-split"></i> 產生中...`;
  }
  
  google.script.run
    .withSuccessHandler(function(res) {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = `<i class="bi bi-download"></i> 匯出關聯清單`;
      }
      if (res && res.success) {
        showToast('匯出成功，即將開啟新分頁！單號：' + res.recordNo);
        window.open(res.url, '_blank');
      } else {
        showToast('匯出失敗：' + (res.message || '未知錯誤'), true);
      }
    })
    .withFailureHandler(function(err) {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = `<i class="bi bi-download"></i> 匯出關聯清單`;
      }
      showToast('後端發生錯誤：' + err.message, true);
    })
    .apiExportNativeDocument(tagId);
}
```

- [ ] **Step 4: Commit**

```bash
git add /Users/kih/Desktop/docmanager/index.html
git commit -m "refactor(export): remove docxtemplater and use native GAS export api"
```

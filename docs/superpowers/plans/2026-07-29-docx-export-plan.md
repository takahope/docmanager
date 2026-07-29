# Docx Export Feature Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement a Docx export feature that generates a structured document report based on a selected tag, flattening parent and child (form) documents into table rows.

**Architecture:** Frontend triggers export via a new modal. Backend (`code.js`) fetches documents for the given tag, maps their parent-child relationships, flattens them, and reads a Docx template from Google Drive. Frontend uses `docxtemplater` and `pizzip` to render the document.

**Tech Stack:** Google Apps Script (GAS), docxtemplater, pizzip, FileSaver.js.

## Global Constraints

- Never hardcode column indices — always use `DOC_COL` / `CLS_COL` / `TAG_COL` from `env.js`.
- Escape all user input in the frontend with `escapeHtml()`.
- Use `LockService` for any writes, although this feature is read-only.
- `_getVisibleDocIds` (auth.js) is the single visibility authority.
- Avoid external API calls on the backend, only use `DriveApp` to read the template file.

---

### Task 1: Add configuration for Docx Template ID

**Files:**
- Modify: `env.js`

**Interfaces:**
- Produces: `PROP_KEYS.DOCX_TEMPLATE_FILE_ID` to be used in backend logic.

- [ ] **Step 1: Add the property key to `env.js`**

Modify `env.js` to add `DOCX_TEMPLATE_FILE_ID` to `PROP_KEYS`. Find `const PROP_KEYS = {` and add it at the end:

```javascript
const PROP_KEYS = {
  HR_SPREADSHEET_ID:   'HR_SPREADSHEET_ID',
  ADMIN_EMAILS:        'ADMIN_EMAILS',
  DOC_FILES_FOLDER_ID: 'DOC_FILES_FOLDER_ID',
  DOCX_TEMPLATE_FILE_ID: 'DOCX_TEMPLATE_FILE_ID',
};
```

- [ ] **Step 2: Commit**

```bash
git add env.js
git commit -m "feat: add DOCX_TEMPLATE_FILE_ID to PROP_KEYS"
```

### Task 2: Implement Backend API for Docx Export

**Files:**
- Modify: `code.js`

**Interfaces:**
- Consumes: `_getVisibleDocIds` from `auth.js`, `DOC_COL` from `env.js`.
- Produces: `apiGetDocxExportData(tagId)` returning `{ templateBase64: string, data: Array<Object> }`.

- [ ] **Step 1: Write `apiGetDocxExportData` in `code.js`**

Add this function to the end of `code.js`:

```javascript
function apiGetDocxExportData(tagId) {
  const ctx = getUserContext();
  const visibleDocIds = _getVisibleDocIds(ctx);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get all documents
  const docSheet = ss.getSheetByName(ENV.SHEET_DOC);
  const docData = docSheet.getDataRange().getDisplayValues();
  const allDocs = {};
  for (let i = 1; i < docData.length; i++) {
    const row = docData[i];
    const docId = row[DOC_COL.DOC_ID];
    if (visibleDocIds.has(docId)) {
      allDocs[docId] = {
        doc_id: docId,
        title: row[DOC_COL.TITLE],
        category: row[DOC_COL.CATEGORY],
        version: row[DOC_COL.VERSION],
        published_at: row[DOC_COL.PUBLISHED_AT]
      };
    }
  }

  // 2. Get tags for documents to filter by tagId
  const docTagSheet = ss.getSheetByName(ENV.SHEET_DOCTAG);
  const docTagData = docTagSheet.getDataRange().getValues();
  const docsWithTag = new Set();
  for (let i = 1; i < docTagData.length; i++) {
    const dId = String(docTagData[i][DOCTAG_COL.DOC_ID]).trim();
    const tId = String(docTagData[i][DOCTAG_COL.TAG_ID]).trim();
    if (tId === tagId && allDocs[dId]) {
      docsWithTag.add(dId);
    }
  }

  // 3. Get closure relationships
  const clsSheet = ss.getSheetByName(ENV.SHEET_CLS);
  const clsData = clsSheet.getDataRange().getValues();
  
  const parentToChildren = {};
  const allChildren = new Set();
  
  for (let i = 1; i < clsData.length; i++) {
    const row = clsData[i];
    const ancestor = String(row[CLS_COL.ANCESTOR_ID]).trim();
    const descendant = String(row[CLS_COL.DESCENDANT_ID]).trim();
    const depth = parseInt(row[CLS_COL.DEPTH], 10);
    
    if (depth === 1 && docsWithTag.has(ancestor) && docsWithTag.has(descendant)) {
      if (!parentToChildren[ancestor]) {
        parentToChildren[ancestor] = [];
      }
      parentToChildren[ancestor].push(descendant);
      allChildren.add(descendant);
    }
  }

  // 4. Flatten relationships
  const flattenedData = [];
  
  for (const docId of docsWithTag) {
    // If it's a child to another doc in this tag, it shouldn't be treated as a parent
    if (allChildren.has(docId)) continue;
    
    const parentDoc = allDocs[docId];
    const childrenIds = parentToChildren[docId] || [];
    
    if (childrenIds.length === 0) {
      flattenedData.push({
        doc_id: parentDoc.doc_id,
        title: parentDoc.title,
        category: parentDoc.category,
        version: parentDoc.version,
        published_at: parentDoc.published_at,
        form_id: "",
        form_title: "",
        form_version: "",
        form_published_at: ""
      });
    } else {
      childrenIds.forEach((childId, index) => {
        const childDoc = allDocs[childId];
        if (index === 0) {
          flattenedData.push({
            doc_id: parentDoc.doc_id,
            title: parentDoc.title,
            category: parentDoc.category,
            version: parentDoc.version,
            published_at: parentDoc.published_at,
            form_id: childDoc.doc_id,
            form_title: childDoc.title,
            form_version: childDoc.version,
            form_published_at: childDoc.published_at
          });
        } else {
          flattenedData.push({
            doc_id: "",
            title: "",
            category: "",
            version: "",
            published_at: "",
            form_id: childDoc.doc_id,
            form_title: childDoc.title,
            form_version: childDoc.version,
            form_published_at: childDoc.published_at
          });
        }
      });
    }
  }

  // 5. Fetch template from Drive
  const templateId = _getProp(PROP_KEYS.DOCX_TEMPLATE_FILE_ID);
  if (!templateId) {
    throw new Error('系統尚未設定 Docx 範本檔案 (DOCX_TEMPLATE_FILE_ID)。');
  }
  
  let templateBase64 = "";
  try {
    const file = DriveApp.getFileById(templateId);
    templateBase64 = Utilities.base64Encode(file.getBlob().getBytes());
  } catch (e) {
    throw new Error('無法讀取 Docx 範本檔案，請檢查檔案 ID 或權限。');
  }

  return {
    templateBase64: templateBase64,
    data: flattenedData
  };
}
```

- [ ] **Step 2: Commit**

```bash
git add code.js
git commit -m "feat: implement apiGetDocxExportData in code.js"
```

### Task 3: Update Frontend `index.html` for Docx Export Modal and Libraries

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: Backend API `apiGetDocxExportData`.
- Produces: Modal UI, script tags for docxtemplater, and export functions.

- [ ] **Step 1: Add CDN Libraries in `index.html` head**
Add this right before `</head>`:

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.55.7/docxtemplater.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.1.7/pizzip.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pizzip/3.1.7/pizzip-utils.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
```

- [ ] **Step 2: Add Export Button and Modal UI**
Find `<button class="btn-export" onclick="exportCsv()">⬇ 匯出 CSV</button>` and add the Docx export button next to it:

```html
<button class="btn-export" onclick="openDocxExportModal()" style="margin-left: 8px;">⬇ 匯出文件一覽表(Word)</button>
```

Find the modals section at the bottom (e.g., above `<div id="toast"></div>`) and add:

```html
<!-- 匯出 Docx Modal -->
<div class="modal-backdrop" id="docxExportModal">
  <div class="modal">
    <div class="modal-header">
      <span>匯出文件一覽表(Word)</span>
      <button class="btn-x" onclick="tryCloseModal('docxExportModal')">✕</button>
    </div>
    <div class="modal-body">
      <div class="form-group">
        <label>選擇標籤 (Tag)</label>
        <select id="docxExportTagSelect"></select>
        <div class="field-hint">報表將匯出此標籤內的文件，並依照父子關係排列。</div>
      </div>
    </div>
    <div class="modal-footer">
      <button onclick="tryCloseModal('docxExportModal')">取消</button>
      <button class="btn-primary" onclick="execDocxExport()">確認匯出</button>
    </div>
  </div>
</div>
```

- [ ] **Step 3: Add Frontend Logic for Docx Export**
Inside `<script>`, near other modal logic (e.g. at the bottom of the script tag), add:

```javascript
// Base64 to ArrayBuffer utility
function base64ToArrayBuffer(base64) {
  var binary_string = window.atob(base64);
  var len = binary_string.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) {
    bytes[i] = binary_string.charCodeAt(i);
  }
  return bytes.buffer;
}

function openDocxExportModal() {
  const sel = document.getElementById('docxExportTagSelect');
  sel.innerHTML = '<option value="">— 請選擇 —</option>';
  
  // Populate from window.allTags (available globally if apiGetInitData ran)
  if (window.allTags) {
    window.allTags.forEach(t => {
      sel.innerHTML += `<option value="${t.tag_id}">${escapeHtml(t.name)}</option>`;
    });
  }
  
  document.getElementById('docxExportModal').classList.add('show');
}

function execDocxExport() {
  const tagId = document.getElementById('docxExportTagSelect').value;
  if (!tagId) {
    showToast('請選擇一個標籤');
    return;
  }
  
  const tagOption = document.getElementById('docxExportTagSelect').options[document.getElementById('docxExportTagSelect').selectedIndex];
  const tagName = tagOption.text;
  
  document.getElementById('loading').classList.remove('hide');
  document.getElementById('loading').innerText = '正在產生文件，請稍候...';
  
  google.script.run
    .withSuccessHandler(function(result) {
      document.getElementById('loading').classList.add('hide');
      document.getElementById('loading').innerText = '處理中…';
      tryCloseModal('docxExportModal');
      
      if (!result.data || result.data.length === 0) {
        showToast('該標籤下無任何文件');
        return;
      }
      
      generateDocx(result.templateBase64, result.data, tagName);
    })
    .withFailureHandler(function(err) {
      document.getElementById('loading').classList.add('hide');
      document.getElementById('loading').innerText = '處理中…';
      showToast('匯出失敗：' + err.message);
    })
    .apiGetDocxExportData(tagId);
}

function generateDocx(templateBase64, data, tagName) {
  try {
    const arrayBuffer = base64ToArrayBuffer(templateBase64);
    const zip = new PizZip(arrayBuffer);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    
    doc.render({
      docs: data
    });
    
    const out = doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    
    const dateStr = new Date().toISOString().split('T')[0];
    saveAs(out, `${tagName}_資訊安全暨個人資料管理文件一覽表_${dateStr}.docx`);
    showToast('匯出成功！');
  } catch (error) {
    console.error(error);
    showToast('文件產生失敗，請檢查範本格式。');
  }
}
```

- [ ] **Step 4: Ensure global variable `window.allTags` is accessible**
Find the `onInitDataReady(res)` function inside `index.html` and add `window.allTags = res.tags;` inside it:

```javascript
function onInitDataReady(res) {
  window.allTags = res.tags;
  // ... existing code ...
}
```

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat: add frontend docx export logic, UI and libraries"
```

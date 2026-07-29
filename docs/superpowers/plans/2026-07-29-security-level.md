# 機密等級 (Security Level) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 為系統新增「機密等級」欄位，並支援在匯出 Word 一覽表時替換對應的佔位符。

**Architecture:** 擴充 `env.js` 中的 `DOC_COL` 常數，於 `deploy.js` 撰寫 `migrateV8()` 以補齊現有試算表欄位。修改 `code.js` 處理新增欄位，並於 `index.html` 加入表單與顯示邏輯。

**Tech Stack:** Google Apps Script, Vanilla JS, HTML/CSS, Docxtemplater

## Global Constraints
- No local test commands except `node --check <file>.js` for syntax validation.
- All column indices must be referenced via `DOC_COL` constants.
- The single visibility authority `_getVisibleDocIds` and editability authority `_getEditableDocIds` must not be bypassed.

---

### Task 1: Update Data Model (env.js & deploy.js)

**Files:**
- Modify: `env.js`
- Modify: `deploy.js`

**Interfaces:**
- Produces: `DOC_COL.SECURITY_LEVEL`, `SECURITY_LEVELS` array, `migrateV8()` function.

- [ ] **Step 1: Add constants to env.js**

```javascript
// env.js 中修改 DOC_COL 及新增 SECURITY_LEVELS
// 在 DOC_COL 的結尾加入 SECURITY_LEVEL: 16
// 更新 DOC_COL_COUNT = 17;
// 新增 const SECURITY_LEVELS = ['一般', '限閱', '密', '機密'];
// 將 DOC_COL 的最後幾欄修改為：
  PENDING_FILE_NAME:   15,  // P: pending_file_name
  SECURITY_LEVEL:      16,  // Q: security_level
};
const DOC_COL_COUNT = 17;
const SECURITY_LEVELS = ['一般', '限閱', '密', '機密'];
```

- [ ] **Step 2: Update sheet deployment in deploy.js**

```javascript
// deploy.js 中修改 _deployDocSheet 函式內的 headers 與 _applyDocSheetFormats
// 在 _deployDocSheet 的 headers 陣列末端加入 'security_level'
// 修改 colWidths 陣列增加對應寬度
// 在 _applyDocSheetFormats 中加入 SECURITY_LEVEL 的資料驗證 (比照 CATEGORY)
```

- [ ] **Step 3: Implement migrateV8 in deploy.js**

```javascript
// deploy.js 中新增 migrateV8 函式，用於補齊 M-Q 欄以及預設值「一般」
function migrateV8() {
  const ss = SpreadsheetApp.openById(ENV.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DOCS);
  if (!sheet) throw new Error(`找不到工作表：${SHEET_NAMES.DOCS}`);

  const col = DOC_COL.SECURITY_LEVEL + 1;
  const cell = sheet.getRange(1, col);
  if (cell.getValue() !== 'security_level') {
    cell.setValue('security_level')
        .setFontWeight('bold')
        .setBackground('#1a3a5c')
        .setFontColor('#ffffff');
    sheet.setColumnWidth(col, 100);
    Logger.log(`✅ 補上欄位：security_level`);
  }

  // 套用驗證與格式
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SECURITY_LEVELS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, 999).setDataValidation(rule);

  // 補齊舊資料預設值為 '一般'
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues();
    let modified = false;
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) {
        values[i][0] = '一般';
        modified = true;
      }
    }
    if (modified) {
      range.setValues(values);
      Logger.log('✅ 舊資料已補上預設機密等級「一般」');
    }
  }
  SpreadsheetApp.flush();
}
```

- [ ] **Step 4: Verify syntax**

Run: `node --check env.js && node --check deploy.js`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add env.js deploy.js
git commit -m "feat: add security level column to database schema"
```

---

### Task 2: Update Backend API (code.js)

**Files:**
- Modify: `code.js`

**Interfaces:**
- Consumes: `DOC_COL.SECURITY_LEVEL`, `SECURITY_LEVELS` from Task 1.
- Produces: API responses with `security_level`.

- [ ] **Step 1: Update apiGetInitData**

```javascript
// 在 apiGetInitData 中新增回傳 securityLevels
// 將回傳的物件增加 `securityLevels: SECURITY_LEVELS`
```

- [ ] **Step 2: Update document creation/update logic**

```javascript
// 修改 apiCreateDoc 與 apiUpdateDoc，在寫入或更新 rows 時加入 doc.security_level
// 在 apiCreateDoc 的新 row 陣列末端確保填寫 doc.security_level || '一般'
// 在 apiUpdateDoc 的欄位更新邏輯，確保更新 row[DOC_COL.SECURITY_LEVEL]
```

- [ ] **Step 3: Verify syntax**

Run: `node --check code.js`
Expected: PASS

- [ ] **Step 4: Commit**

```bash
git add code.js
git commit -m "feat: handle security level in backend CRUD APIs"
```

---

### Task 3: Update Frontend (index.html)

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `securityLevels` from `apiGetInitData`, `security_level` in doc object.

- [ ] **Step 1: Add UI element to docModal**

```html
<!-- 在 index.html 中找 docModal，在類別旁新增機密等級下拉選單 -->
<div class="form-group"><label>機密等級</label><select id="f_security_level"></select></div>
```

- [ ] **Step 2: Populate dropdown and read/write values**

```javascript
// 在 init() 函式的 successHandler 中：
// res.securityLevels.forEach(s => addOption('f_security_level', s, s));

// 在 openDocModal() 中載入文件資料：
// el('f_security_level').value = doc ? doc.security_level : '一般';

// 在 saveDoc() 中取得資料：
// doc.security_level = el('f_security_level').value;
```

- [ ] **Step 3: Display in table and detail panel**

```html
<!-- 修改 docTable 表頭 -->
<th>機密等級</th>
```

```javascript
// 修改 renderDocTable()，在表格加入 doc.security_level 顯示
// 修改 renderDetailPanel()，在 detail-meta 區塊顯示機密等級
```

- [ ] **Step 4: Update Docx export payload**

```javascript
// 修改 execDocxExport()
// 在 docs.map() 迴圈內產生的 obj 物件中加入：
// security_level: doc.security_level || '一般',
```

- [ ] **Step 5: Verify syntax**

Run: (No easy way to test HTML JS syntax statically, manual review applies)

- [ ] **Step 6: Commit**

```bash
git add index.html
git commit -m "feat: add security level to UI and word export"
```

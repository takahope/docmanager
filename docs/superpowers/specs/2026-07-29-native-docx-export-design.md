# Native Google Apps Script Docx Export Design

## 1. Overview
This specification details the transition of the document export functionality from a frontend-heavy third-party implementation (`docxtemplater` + `pizzip`) to a native Google Apps Script `DocumentApp` backend implementation. This approach improves performance, reduces frontend payload, and simplifies the codebase.

## 2. Architecture & Approach
The system will adopt **Option A: Full Backend Pipeline**. 
When a user clicks "Export", the frontend calls a single API. The backend fetches all necessary data, prepares it, duplicates a Google Doc template, injects data into it, and returns the URL of the newly created Google Doc.

### 2.1 Backend (`code.js`)
*   **New Main Entry Point**: `apiExportNativeDocument(tagId)` will replace `apiGetDocxExportData`.
*   **Data Formatting**: The hierarchical relation data will be flattened and formatted into a 2D array representing a table. 
    *   **Table Headers**: `['文件編號', '文件名稱', '機密等級', '版本', '發行日期', '表單編號', '表單名稱', '表單版本', '表單發行日期']`
*   **Helper Functions**: 
    *   `escapeRegExp_(value)`: To safely escape regex tokens.
    *   `createRecordNoFromFolder_(folder, prefix, dateKey)`: To calculate the next sequential record number based on existing files in the output directory.
    *   `replaceTemplateTokens_(doc, tokenMap)`: To substitute text placeholders (`{{年}}`, `{{月}}`, `{{日}}`, `{{紀錄編號}}`) across the Doc body, header, and footer.
*   **Document Generation Logic**: 
    1.  Copy the template Google Doc to the output folder.
    2.  Open the newly copied document.
    3.  Replace text tokens.
    4.  Locate `{{表格}}` and insert the 2D data array as a table.
    5.  Format the table (padding, header background color, bold text).
    6.  Save and close the document.
    7.  Return `{ success: true, url: string, recordNo: string }`.

### 2.2 Frontend (`index.html`)
*   **Library Removal**: Delete the `<script>` tags for `pizzip`, `docxtemplater`, and `FileSaver.js`.
*   **Code Cleanup**: Remove `generateDocx()` and `base64ToArrayBuffer()` functions.
*   **API Invocation**: Update the export trigger to call `apiExportNativeDocument`.
*   **Success Handling**: Upon success, call `window.open(res.url, '_blank')` to open the generated Google Doc in a new browser tab.

### 2.3 Environment & Constants (`env.js`)
*   No changes are strictly required here. The existing `DOCX_TEMPLATE_FILE_ID` and `DOCX_OUTPUT_FOLDER_ID` properties will be reused, with the understanding that the template ID must now point to a native Google Doc file.

## 3. Data Flow
1.  User clicks the "Export" button for a specific Tag in the frontend.
2.  Frontend sets loading state and invokes `google.script.run.apiExportNativeDocument(tagId)`.
3.  Backend authenticates the user, retrieves visible document rows, filters by the selected `tagId`, and identifies ancestor/descendant relationships.
4.  Backend flattens the relationships into a 2D array.
5.  Backend calculates the `recordNo` and makes a copy of the template Google Doc.
6.  Backend replaces placeholder text and injects the 2D array into the document as a formatted table.
7.  Backend responds with the URL of the generated document.
8.  Frontend resets the loading state, shows a success toast, and opens the returned URL in a new tab.

## 4. Error Handling
*   **Invalid Template ID**: If `DriveApp.getFileById(templateId)` fails, a clear error message is thrown indicating that the template file could not be found or permissions are insufficient.
*   **Invalid Output Folder**: If `DriveApp.getFolderById(folderId)` fails, a clear error message is thrown.
*   **Missing Table Token**: If `{{表格}}` is not found in the template, the table is appended to the very end of the document.
*   Frontend `withFailureHandler` captures and displays these errors in a Toast notification to the user without leaving them in a hanging "Loading..." state.

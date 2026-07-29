# Document Export Grouping & Vertical Cell Merging Design

## Objective
Enhance the existing Google Doc export functionality (`apiExportNativeDocument`) to output a highly formatted, enterprise-grade table. The solution will handle dynamic category headers (already implemented via `DocumentApp` row copying) and perform vertical cell merging (`rowSpan`) for parent documents that span multiple child forms.

## Context & Current State
- The user has already implemented the `DocumentApp` phase: 
  - Data is flattened into a one-to-many relationship (`parentRows`).
  - Pre-defined groups (`TWHB-ISMSPIMS-`, `TWHB-ISMS-`, `TWHB-PISM-`) and dynamic groups (`category` fallback) are correctly routed.
  - Template rows with horizontal spans and background colors are copied perfectly.
  - Child form rows output empty strings `""` for the first 5 columns.
- **New Advantage**: The user has successfully enabled the **Google Docs API Advanced Service** (identifier: `Docs`). This eliminates the need for raw `UrlFetchApp` REST calls.

## Architecture & Data Flow

The solution employs a three-phase hybrid approach:

### Phase 1: Table Generation (DocumentApp)
*(Already implemented by User)*
- Clone template, copy pre-formatted rows, populate text, and strip unused placeholders.
- `doc.saveAndClose()` must be called to persist the document structure.

### Phase 2: Vertical Merging (Docs Advanced Service)
- **Read**: Use `Docs.Documents.get(documentId)` to retrieve the document's JSON structure.
- **Scan**: Iterate through the main table's rows.
  - Check the first cell (Column 0).
  - If the text is a valid Document ID (not empty, not a header), it marks `parentStartRow`.
  - Look ahead: Count subsequent rows where Column 0 is empty (just `\n`) to determine the `rowSpan`.
- **Merge**: If `rowSpan > 1`, generate 5 `MergeTableCellsRequest` payloads (for Column 0 through 4).
- **Update**: Execute `Docs.Documents.batchUpdate({ requests }, documentId)` to apply all merges simultaneously.

### Phase 3: Post-Merge Cleanup (DocumentApp)
- **Problem**: When Docs API merges cells, it concatenates their content. Merging 3 cells containing empty paragraphs results in a single cell with 3 blank paragraphs, pushing the text off-center.
- **Solution**: Re-open the document using `DocumentApp.openById(documentId)`.
- Loop through the table. For each cell in columns 0-4, remove any trailing empty paragraphs so the text aligns vertically perfectly.
- `doc.saveAndClose()`.

## Error Handling
- If `Docs.Documents.get()` fails or the table cannot be found, catch the exception and log the error, but do not delete the generated document (the user still gets a usable unmerged table).
- Robustly handle empty categories and missing fields to prevent undefined errors during text scanning.

## Testing Strategy
- Export a known dataset containing a parent with 0 forms, 1 form, and 3 forms.
- Verify that the 3-form parent visually merges the left 5 columns.
- Verify that no extra newlines exist in the merged cells.

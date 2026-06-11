# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Google Apps Script (GAS) full-stack document management system using Google Sheets as the backend database. It uses a **Closure Table** design to model multi-level relationships between documents (e.g., "references", "supersedes", "derived_from", "related").

## File Structure & Responsibilities

| File | Responsibility |
|---|---|
| `env.js` | Centralized config: spreadsheet ID, sheet names, column index constants, dropdown option lists |
| `deploy.js` | One-time setup: `deployAllSheets()` creates the two sheets with headers/formatting; `seedSampleData()` inserts sample rows |
| `code.js` | Core business logic: document CRUD, Closure Table maintenance, all `apiXxx` functions exposed to the frontend |
| `index.html` | Single-file frontend (HTML/CSS/JS inlined) |

## Deployment (no build/test tooling — this is a GAS project)

There is no local build, lint, or test command. Development happens by copying file contents into the Apps Script editor:

1. Create a new Google Spreadsheet → Extensions → Apps Script
2. Create four files in the editor and paste matching content:
   - `env.gs` ← `env.js`
   - `deploy.gs` ← `deploy.js`
   - `code.gs` ← `code.js`
   - `index.html` (create as type "HTML")
3. Run `deployAllSheets()` manually once → authorize → creates the two sheets
4. (Optional) Run `seedSampleData()` to populate sample rows
5. Deploy → New deployment → Web app (execute as: me; access per org policy)

## Architecture

### Data Model (Google Sheets)

**Sheet "文件清單" (Documents)** — columns defined in `DOC_COL` (env.js):
`doc_id` (DOC-001 format, text-forced), `title`, `category`, `status`, `owner`, `owner_ID` (text-forced), `updated_at`, `version`, `google_drive_location`.

Option lists also centralized in env.js: `DOC_STATUS`, `RELATION_TYPES`, `DOC_CATEGORIES`.

**Sheet "文件關聯" (Closure Table)** — columns defined in `CLS_COL` (env.js):
`doc_id` (ancestor), `descendant_id`, `depth` (0=self, 1=direct, 2+=indirect), `relation_type`, `說明` (note).

### Closure Table Maintenance Logic (code.js)

This is the core algorithmic complexity of the project — read these functions together when modifying relation logic:

- **Adding A → B (`apiAddRelation`)**: First checks for cycles (reject if B already (in)directly relates to A). Then writes the cartesian product of (X → A) × (B → Y) pairs as X → Y with `depth = dX + 1 + dY`, so every ancestor of A can directly query B and B's descendants.
- **Removing A → B (`apiRemoveRelation`)**: Deletes all affected `depth > 0` paths, then BFS-rebuilds indirect paths from the remaining direct (`depth=1`) edges via `_rebuildClosurePaths` (handles multi-path cases).
- **Deleting a document (`apiDeleteDoc`)**: Cascades — removes all closure rows where the doc is either ancestor or descendant.

### Frontend ↔ Backend Communication

`index.html` calls backend functions via `google.script.run.withSuccessHandler(...).withFailureHandler(...)`. Main exposed API surface (all in code.js):

- `apiGetInitData()` — all docs + dropdown option lists
- `apiCreateDoc(doc)` / `apiUpdateDoc(doc)` / `apiDeleteDoc(docId)`
- `apiGetDescendants(docId, maxDepth)` — what this doc references (incl. indirect)
- `apiGetAncestors(docId)` — what references this doc
- `apiAddRelation(ancestorId, descendantId, relationType, description)`
- `apiRemoveRelation(ancestorId, descendantId)`
- `apiGetGraphData()` — nodes + direct edges for relation graph visualization

## Established Best Practices (must follow)

- **Never hardcode column indices** (e.g. `row[8]`) — always use `DOC_COL` / `CLS_COL` from `env.js`.
- Call `SpreadsheetApp.flush()` after writes.
- `doc_id` and `owner_ID` columns must be forced to text format (`@`) to avoid numeric coercion.
- Use `getDisplayValues()` when reading dates to avoid serialization to `null`.
- Use `LockService` to prevent race conditions on concurrent writes.
- Escape all user input in the frontend with `escapeHtml()` to prevent XSS.
- Search/filter inputs use debounce to avoid UI jank with large datasets.

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Google Apps Script (GAS) full-stack document management system using Google Sheets as the backend database. It uses a **Closure Table** design to model multi-level relationships between documents (e.g., "references", "supersedes", "derived_from", "related").

V2 (2026-07) added: HR-whitelist access control, role-based editing, status-flow enforcement, publish/review-date tracking, an audit trail sheet, a dashboard, an SVG relation graph, and CSV export.

## File Structure & Responsibilities

| File | Responsibility |
|---|---|
| `env.js` | Centralized config: sheet names, column index constants (`DOC_COL`/`CLS_COL`/`AUDIT_COL`/`HR_COL`), `STATUS_TRANSITIONS` map, option lists, Script Properties key names (`PROP_KEYS`) |
| `deploy.js` | `deployAllSheets()` creates the three sheets; `migrateV2()` idempotently adds J–M columns + audit sheet to a V1 spreadsheet; `seedSampleData()` |
| `auth.js` | Permission layer: HR whitelist (CacheService 10 min), `getUserContext()`, `_assertWhitelisted/_assertCanEditDoc/_assertAdmin`, `authorizeOnce()`, `debugGetSystemData()`, `clearHrCache()` |
| `audit.js` | `_logAudit()` (called inside locks; swallows its own errors), `_diffSummary()`, `apiGetDocHistory(docId)` |
| `code.js` | `doGet` whitelist gate, document CRUD, Closure Table maintenance, all other `apiXxx` functions |
| `index.html` | Single-file frontend: 3 tabs (dashboard / list+detail / SVG graph), fake progress bar on boot, CSV export, modal dirty-guard |

## Deployment (no build/test tooling — this is a GAS project)

There is no local build or test command; syntax can be checked with `node --check <file>.js`. Development happens by copying file contents into the Apps Script editor:

1. Create a new Google Spreadsheet → Extensions → Apps Script
2. Create six files in the editor and paste matching content:
   - `env.gs` ← `env.js`, `deploy.gs` ← `deploy.js`, `auth.gs` ← `auth.js`, `audit.gs` ← `audit.js`, `code.gs` ← `code.js`
   - `index.html` (create as type "HTML")
3. Script Properties (Project Settings): set `HR_SPREADSHEET_ID` (HR master spreadsheet) and `ADMIN_EMAILS` (comma-separated) — **never hardcode these in env.js (PLAYBOOK P6)**
4. Run `authorizeOnce()` once in the editor → authorize scopes (Sheets read on HR file)
5. Run `deployAllSheets()` (new install) or `migrateV2()` (upgrading a V1 sheet — idempotent, safe to re-run)
6. (Optional) `seedSampleData()`; run `debugGetSystemData()` to verify HR headers + whitelist count
7. Deploy → New deployment → Web app (execute as: me; access: domain). `clasp push` only updates /dev; /exec needs a new deployment version.

## Architecture

### Data Model (Google Sheets)

**Sheet "文件清單" (Documents)** — columns defined in `DOC_COL` (env.js):
`doc_id` (DOC-001 format, text-forced), `title`, `category`, `status`, `owner`, `owner_ID` (text-forced), `updated_at`, `version`, `google_drive_location`, `owner_email` (permission match), `published_at`, `next_review_date`, `review_cycle_months`. Columns J–M are V2 additions; date columns are text-forced (`@`) to avoid Date serialization issues.

**Sheet "文件關聯" (Closure Table)** — columns defined in `CLS_COL` (env.js):
`doc_id` (ancestor), `descendant_id`, `depth` (0=self, 1=direct, 2+=indirect), `relation_type`, `說明` (note).

**Sheet "異動紀錄" (Audit trail)** — columns in `AUDIT_COL`: 時間, 操作者, 動作, doc_id, 版本, 變更摘要. Serves double duty as per-document version history (filter by doc_id).

### Permission model (auth.js)

- Whitelist source: HR master spreadsheet (Script Properties `HR_SPREADSHEET_ID`), sheet 人員主檔 (A email / B name / C status), filtered by **exclusion** (`EXCLUDED_HR_STATUS = ['離職']`) because Chinese status values are unreliable for exact matching. Cached 10 min.
- Roles: admin (`ADMIN_EMAILS`) > document owner (`owner_email` === login email, can edit own docs/relations) > whitelisted user (read-only + can create docs, becoming owner) > outsider (doGet renders denial page).
- Frontend `isAdmin` (real identity) vs `isAdminView` (view toggle) are separate; **every mutating API re-asserts permissions server-side** (IDOR protection). Delete is admin-only.

### Status flow (env.js `STATUS_TRANSITIONS` + code.js `apiUpdateDoc`)

草稿→審核中→已發布→已廢止 (with 審核中→草稿 rollback, 已發布→審核中 for revision, 已廢止→草稿 admin-only). Transition to 已發布 auto-fills `published_at` and computes `next_review_date` = today + `review_cycle_months`. Creating a `supersedes` relation onto a published doc auto-deprecates it (frontend confirms first; backend logs it).

### Closure Table Maintenance Logic (code.js)

This is the core algorithmic complexity of the project — read these functions together when modifying relation logic:

- **Adding A → B (`apiAddRelation`)**: First checks for cycles (reject if B already (in)directly relates to A). Then writes the cartesian product of (X → A) × (B → Y) pairs as X → Y with `depth = dX + 1 + dY`, so every ancestor of A can directly query B and B's descendants.
- **Removing A → B (`apiRemoveRelation`)**: Deletes all affected `depth > 0` paths, then BFS-rebuilds indirect paths from the remaining direct (`depth=1`) edges via `_rebuildClosurePaths` (handles multi-path cases).
- **Deleting a document (`apiDeleteDoc`)**: Cascades — removes all closure rows where the doc is either ancestor or descendant.

### Frontend ↔ Backend Communication

`index.html` calls backend functions via `google.script.run.withSuccessHandler(...).withFailureHandler(...)` (every call must have a failure handler). Exposed API surface:

- `apiGetInitData()` — docs + option lists + `statusTransitions` + `user` context + `hrPeople` (owner dropdown)
- `apiCreateDoc(doc)` / `apiUpdateDoc(doc)` / `apiDeleteDoc(docId)` (admin only)
- `apiGetDescendants(docId, maxDepth)` / `apiGetAncestors(docId)`
- `apiAddRelation(...)` (returns `deprecated` doc_id when supersedes auto-deprecates) / `apiRemoveRelation(...)`
- `apiGetGraphData()` — nodes + direct edges; rendered as a pure-SVG layered DAG (longest-path layering — safe because closure table guarantees acyclicity)
- `apiGetDocHistory(docId)` — audit rows for one document

Frontend conventions: dashboard stats are computed **client-side** from the already-loaded docs (no extra API); saves are optimistic (local state updated from inputs, no full reload); CSV export is built client-side with a `﻿` BOM.

## Established Best Practices (must follow)

- **Never hardcode column indices** (e.g. `row[8]`) — always use `DOC_COL` / `CLS_COL` / `AUDIT_COL` from `env.js`.
- Call `SpreadsheetApp.flush()` after writes.
- `doc_id`, `owner_ID`, and the date/email columns (J–L) must be forced to text format (`@`) to avoid coercion/serialization issues.
- Use `getDisplayValues()` when reading dates to avoid serialization to `null`.
- Use `LockService` to prevent race conditions on concurrent writes; `_logAudit` is called inside the lock and must never fail the main operation.
- Every mutating API starts with an `_assert*` permission check from auth.js — do not add write APIs without one.
- Status changes must go through `STATUS_TRANSITIONS` — never bypass with a raw `setValue` on the status column (except the documented supersedes auto-deprecation path).
- Escape all user input in the frontend with `escapeHtml()` to prevent XSS.
- Search/filter inputs use debounce **and** composition-event guards (注音 IME); re-rendering must not rebuild the input node.
- Modals never close on backdrop click; dirty forms confirm before closing.
- Secrets/IDs live in Script Properties (`PROP_KEYS`), never in committed code.

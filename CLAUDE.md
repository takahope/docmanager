# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A Google Apps Script (GAS) full-stack document management system using Google Sheets as the backend database. It uses a **Closure Table** design to model multi-level relationships between documents (e.g., "references", "supersedes", "derived_from", "related").

V2 (2026-07) added: HR-whitelist access control, role-based editing, status-flow enforcement, publish/review-date tracking, an audit trail sheet, a dashboard, an SVG relation graph, and CSV export.

V5 (2026-07) added: two-level permission model (read/edit) on tags and group grants; `_getEditableDocIds` authority; edit-gated write APIs; edit-only document detail editing and tag-management UI.

## File Structure & Responsibilities

| File | Responsibility |
|---|---|
| `env.js` | Centralized config: sheet names, column index constants (`DOC_COL`/`CLS_COL`/`AUDIT_COL`/`HR_COL`), `STATUS_TRANSITIONS` map, option lists, Script Properties key names (`PROP_KEYS`) |
| `deploy.js` | `deployAllSheets()` creates the three sheets; `migrateV2()` idempotently adds J–M columns + audit sheet to a V1 spreadsheet; `migrateV6()` adds read/edit permission columns to 使用者授權 and 群組授權 sheets (V4→V5 upgrade); `seedSampleData()` |
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
5. Run `deployAllSheets()` (new install) or `migrateV2()` (V1→V2) or `migrateV6()` (V4→V5, adding permission columns — idempotent, safe to re-run)
6. (Optional) `seedSampleData()`; run `debugGetSystemData()` to verify HR headers + whitelist count
7. Deploy → New deployment → Web app (execute as: me; access: domain). `clasp push` only updates /dev; /exec needs a new deployment version.

## Architecture

### Data Model (Google Sheets)

**Sheet "文件清單" (Documents)** — columns defined in `DOC_COL` (env.js):
`doc_id` (DOC-001 format, text-forced), `title`, `category`, `status`, `owner`, `owner_ID` (text-forced), `updated_at`, `version`, `google_drive_location`, `owner_email` (permission match), `published_at`, `next_review_date`, `review_cycle_months`. Columns J–M are V2 additions; date columns are text-forced (`@`) to avoid Date serialization issues.

**Sheet "文件關聯" (Closure Table)** — columns defined in `CLS_COL` (env.js):
`doc_id` (ancestor), `descendant_id`, `depth` (0=self, 1=direct, 2+=indirect), `relation_type`, `說明` (note).

**Sheet "異動紀錄" (Audit trail)** — columns in `AUDIT_COL`: 時間, 操作者, 動作, doc_id, 版本, 變更摘要. Serves double duty as per-document version history (filter by doc_id).

**Sheet "標籤主檔" (Tags)** — columns in `TAG_COL` (V3): `tag_id` (TAG-001 format, text-forced), `name`, `parent_id` (parent tag_id, empty = root; adjacency list — tag count is small, frontend builds the tree recursively, no closure table needed), `sort` (same-level ordering number). Tags double as **folders**.

**Sheet "文件標籤" (Doc-Tags)** — columns in `DOCTAG_COL` (V3): `doc_id` | `tag_id` (many-to-many, one pair per row). A document with multiple tags appears in multiple folders without copying.

**Sheet "使用者授權" (User Grants)** — columns in `GRANT_COL` (V3, V5 adds column 3): `email` (lowercased) | `tag_id` | `permission` (one grant per row). `permission` field: empty = `read` (default), `'edit'` = edit access (fail-closed — only these two values). Grants are **not** HR-cached — they take effect immediately.

**Sheet "群組授權" (Group Grants)** — columns in `GROUPGRANT_COL` (V4, V5 adds column 4): `org_code` | `title` | `tag_id` | `permission` (one grant per row; at least one of org_code/title non-empty). `permission` field: empty = `read` (default), `'edit'` = edit access (fail-closed). Group membership is resolved live from the HR master's 組織架構樹 + 人員職務配置 sheets (cached 10 min) — this system stores only the (group → tag) mapping. **org_code matches direct members only (no org-subtree expansion)** — deliberately opposite to the tag tree's parent-includes-child inheritance.

### Permission model (auth.js)

- Whitelist source: HR master spreadsheet (Script Properties `HR_SPREADSHEET_ID`), sheet 人員主檔 (A email / B name / C status), filtered by **exclusion** (`EXCLUDED_HR_STATUS = ['離職']`) because Chinese status values are unreliable for exact matching. Cached 10 min.
- Roles: admin (`ADMIN_EMAILS`) > document owner (`owner_email` === login email, can edit own docs/relations) > whitelisted user (read-only + can create docs, becoming owner) > outsider (doGet renders denial page).
- Frontend `isAdmin` (real identity) vs `isAdminView` (view toggle) are separate; **every mutating API re-asserts permissions server-side** (IDOR protection). Delete is admin-only.

### Tag-permission / visibility model (V3, auth.js)

- **純標籤授權 (tag-only authorization)**: documents carry tags; admins grant tags to users; a user sees documents bearing a granted tag. No per-document grants.
- **群組授權 (V4, group grants)**: `_getEffectiveGrantedTagIds(ctx)` = personal grants ∪ group grants (via `_groupGrantHits`: each HR assignment row × each group-grant row; org_code empty-or-exact-match AND title empty-or-exact-match). `_getVisibleDocIds` consumes this union; everything downstream (tag-subtree expansion, rules 1–4) is unchanged. Title matching is exact (trimmed); admin UI only offers existing titles from a dropdown.
- **父含子繼承 (parent-includes-child inheritance)**: granting a parent tag makes the entire subtree of documents visible. `_expandTagWithDescendants(tagIds, allTags)` BFS-expands a grant set over `parent_id`.
- **deny-by-default for untagged docs**: a document with no tags is visible only to admins and its owner. Before go-live, admins must tag existing documents or ordinary users see nothing.
- `_getVisibleDocIds(ctx)` returns a `Set` and is the **single visibility authority**: (1) admin → all; (2) own `owner_email` docs → visible; (3) any doc tag ∈ the user's expanded grant set → visible; (4) untagged → rules 1 & 2 only. `_assertCanViewDoc(docId)` guards single-doc entry points (history, ancestor/descendant queries).
- Tag-tree maintenance (create/rename/move/delete) is admin-only (`_assertAdmin`); tagging a document uses `_assertCanEditDoc` (admin + owner). Invisible documents are **fully hidden** — absent from the graph, ancestor/descendant queries, and detail pages.
- **編輯授權（V5, edit permission level）**：授權列帶 `permission`（read/edit，空白＝read，fail-closed）。`_getEffectiveGrantedTagIds(ctx)` 回傳 `{read, edit}` 兩個 Set（read ⊇ edit，edit 蘊含 read）。`_getEditableDocIds(ctx)` 是**唯一的可編輯性事實來源**（與 `_getVisibleDocIds` 同構：admin 全部／owner_email／文件標籤 ∈ edit 子孫展開集／無標籤僅前二者）。edit 授權者可改欄位、走狀態流轉、維護關聯；**貼標籤與更換負責人仍限 `_assertOwnerOrAdmin`（admin + owner）**——改標籤＝改可見範圍，不隨 edit 權下放。supersedes 關聯因會自動廢止後代端文件，需具備後代端編輯權。授權（含等級）仍僅管理員可設定。

### Status flow (env.js `STATUS_TRANSITIONS` + code.js `apiUpdateDoc`)

草稿→審核中→已發布→已廢止 (with 審核中→草稿 rollback, 已發布→審核中 for revision, 已廢止→草稿 admin-only). Transition to 已發布 auto-fills `published_at` and computes `next_review_date` = today + `review_cycle_months`. Creating a `supersedes` relation onto a published doc auto-deprecates it (frontend confirms first; backend logs it).

### Closure Table Maintenance Logic (code.js)

This is the core algorithmic complexity of the project — read these functions together when modifying relation logic:

- **Adding A → B (`apiAddRelation`)**: First checks for cycles (reject if B already (in)directly relates to A). Then writes the cartesian product of (X → A) × (B → Y) pairs as X → Y with `depth = dX + 1 + dY`, so every ancestor of A can directly query B and B's descendants.
- **Removing A → B (`apiRemoveRelation`)**: Deletes all affected `depth > 0` paths, then BFS-rebuilds indirect paths from the remaining direct (`depth=1`) edges via `_rebuildClosurePaths` (handles multi-path cases).
- **Deleting a document (`apiDeleteDoc`)**: Cascades — removes all closure rows where the doc is either ancestor or descendant.

### Frontend ↔ Backend Communication

`index.html` calls backend functions via `google.script.run.withSuccessHandler(...).withFailureHandler(...)` (every call must have a failure handler). Exposed API surface:

- `apiGetInitData()` — docs (filtered to the visible set) + option lists + `statusTransitions` + `user` context (with `grantedTagIds`: read-tier tag ids, unchanged shape) + `editableDocIds` (V5: doc ids the user may edit, filtered to the editable set) + `hrPeople` (owner dropdown) + `tags` (all tags — folder tree needs names) + `docTags` (visible docs only)
- `apiCreateDoc(doc)` (may carry `tagIds`) / `apiUpdateDoc(doc)` / `apiDeleteDoc(docId)` (admin only; cascades doc-tag rows)
- `apiGetDescendants(docId, maxDepth)` / `apiGetAncestors(docId)` — `_assertCanViewDoc` at entry, results re-filtered to the visible set
- `apiAddRelation(...)` (returns `deprecated` doc_id when supersedes auto-deprecates) / `apiRemoveRelation(...)`
- `apiGetGraphData()` — nodes + direct edges, both filtered to the visible set; rendered as a pure-SVG layered DAG (longest-path layering — safe because closure table guarantees acyclicity)
- `apiGetDocHistory(docId)` — audit rows for one document (`_assertCanViewDoc` at entry)
- `apiSetDocTags(docId, tagIds)` (`_assertOwnerOrAdmin`, overwrites the doc's tag rows; V5: edit grantees may not call this) — V3
- `apiCreateTag(name, parentId)` / `apiRenameTag(tagId, name)` / `apiMoveTag(tagId, newParentId)` (cycle-checked) / `apiDeleteTag(tagId)` (cascades doc-tag + grant rows) — admin only, V3
- `apiSetUserGrants(email, grants)` / `apiGetAllGrants()` — admin only, V3 (V5: grants=[{tagId, permission}], full-set overwrite, `[]` deletes)
- `apiGetGroupGrants()` / `apiSetGroupGrants(orgCode, title, grants)` (full-set overwrite per combo, `[]` deletes; V5: grants=[{tagId, permission}]) / `apiGetOrgOptions()` / `apiPreviewUserTags(email)` (V5: returns {tagId, permission} pairs) — admin only, V4

Frontend conventions: dashboard stats are computed **client-side** from the already-loaded docs (no extra API); saves are optimistic (local state updated from inputs, no full reload); CSV export is built client-side with a `﻿` BOM.

## Established Best Practices (must follow)

- **Never hardcode column indices** (e.g. `row[8]`) — always use `DOC_COL` / `CLS_COL` / `AUDIT_COL` / `TAG_COL` / `DOCTAG_COL` / `GRANT_COL` from `env.js`.
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
- `_getVisibleDocIds` (auth.js) is the **single visibility authority** — every read API that returns docs/nodes/history must filter through it (or `_assertCanViewDoc` at a single-doc entry point); no API may bypass it and return docs directly. Group-grant resolution goes through `_getEffectiveGrantedTagIds` — never read 群組授權 rows directly in an API.
- `_getEditableDocIds` (auth.js, V5) is the **single editability authority** — every write API that modifies a document must check authorization via `_assertCanEditDoc` or `_assertOwnerOrAdmin` or `_assertAdmin`, never bypassing these guards. `permission` values in grants must be normalized through `_normPermission` (whitelist: empty/`'read'`/`'edit'`, default empty) — never accept raw user input for permission fields.

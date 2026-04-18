# Cloudflare Migration

## Why move off Google Drive
* API calls are slow
* Users can mess with their own files
* Locating the sheetable root folder is by filename
* Sharing is messy and hacky, and requires API permissions that would require verification with google to be approved for public release.

## Stack
* **Cloudflare Pages** — hosts the SPA + Worker API at the same origin (no CORS)
* **Cloudflare Workers** (Pages Functions) — API layer, auth, presigned URLs for R2
* **Cloudflare D1** — SQLite database for structured table data
* **Cloudflare R2** — blob storage for images

## Key simplifications
* **No local-only mode.** All data lives in D1. Remove localStorage-based workbook snapshots entirely.
* **No internal DataModel / transaction list.** The browser grid is the source of truth. Each completed cell edit, row insert, or row delete fires an async `fetch()` to the Worker API which writes to D1 immediately. No dirty tracking, no batch auto-save, no generation counters.
* **Single D1 database for all users.** D1 doesn't support dynamic runtime bindings, so one DB per book isn't practical. Instead, all data lives in one D1 database with `book_id` columns. A single R2 bucket with `{book_id}/` key prefixes. Sharing = checking `book_members` table before every query.
* **No data migration.** Only test data exists currently — clean slate for the new backend.
* **Lightweight undo.** A small client-side stack of `{rowId, column, oldValue}`. Ctrl+Z pops and fires a reverse write. ~20 lines, no DataModel class.
* **Client-generated UUID row IDs.** Avoids the insert→edit race condition where the browser wouldn't know the server-assigned ID yet.
* **Client + server dual validation.** Browser validates optimistically (type checks, unique keys, references) using its in-memory row data. Server re-validates authoritatively. On rare server rejection, revert + toast.
* **Per-row write queue.** Cell edits to the same row are serialized client-side to prevent out-of-order writes. Different rows fire concurrently.

## Free tier budget (no paid plan needed for small usage)
| Service   | Free limit                                  |
| --------- | ------------------------------------------- |
| Workers   | 100K requests/day, 10ms CPU/invocation      |
| D1        | 5 GB storage, 5M reads/day, 100K writes/day |
| R2        | 10 GB, 1M writes/mo, 10M reads/mo           |
| Pages     | Unlimited static bandwidth, 500 builds/mo   |

---

## D1 database schema

A single D1 database stores everything: users, book registry, memberships, and all user table data. Physical user tables are named with auto-incremented IDs (`t_1`, `t_2`, ...) to avoid collisions. Human-readable names live only in `_tables` metadata.

```sql
------------------------------------------------------------
-- Account & book management
------------------------------------------------------------

-- Users (one row per authenticated person)
CREATE TABLE users (
  id         TEXT PRIMARY KEY,          -- opaque ID (e.g. Google sub)
  email      TEXT NOT NULL UNIQUE,
  name       TEXT,
  picture    TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Book registry
CREATE TABLE books (
  id         TEXT PRIMARY KEY,          -- UUID
  name       TEXT NOT NULL,
  owner_id   TEXT NOT NULL REFERENCES users(id),
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Book membership (who can access which book)
CREATE TABLE book_members (
  book_id  TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  user_id  TEXT NOT NULL REFERENCES users(id),
  role     TEXT NOT NULL CHECK (role IN ('owner', 'editor', 'viewer')),
  PRIMARY KEY (book_id, user_id)
);

------------------------------------------------------------
-- Table/column metadata (scoped to book)
------------------------------------------------------------

-- Auto-increment counter for physical table names (t_1, t_2, ...)
-- Each row registers a user-visible table.
CREATE TABLE _tables (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,  -- physical name = t_{id}
  book_id         TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,                       -- user-visible name
  display_order   INTEGER NOT NULL,
  unique_keys     TEXT NOT NULL DEFAULT '[]',          -- JSON array of column names
  default_sort    TEXT,                                 -- JSON array [{column, direction}]
  draft_position  TEXT NOT NULL DEFAULT 'bottom',
  UNIQUE (book_id, name)
);

-- Column definitions
CREATE TABLE _columns (
  table_id        INTEGER NOT NULL REFERENCES _tables(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  display_name    TEXT,
  type            TEXT NOT NULL,  -- text|integer|decimal|date|datetime|bool|reference|image
  display_order   INTEGER NOT NULL,
  ref_table       TEXT,           -- name (not id) of referenced table within same book
  ref_display     TEXT,           -- JSON array
  ref_search      TEXT,           -- JSON array
  PRIMARY KEY (table_id, name)
);

-- Chart sheets
CREATE TABLE _chart_sheets (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  book_id         TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  table_name      TEXT,
  mode            TEXT NOT NULL DEFAULT 'edit',
  charts          TEXT NOT NULL DEFAULT '[]',          -- JSON blob
  display_order   INTEGER NOT NULL,
  UNIQUE (book_id, name)
);

------------------------------------------------------------
-- User data tables (created dynamically by Worker)
------------------------------------------------------------
-- When _tables row id=42 is created with columns [Name, Email, Phone]:
--
--   CREATE TABLE t_42 (
--     _rowId TEXT PRIMARY KEY,          -- client-generated UUID
--     "Name"  TEXT NOT NULL DEFAULT '',
--     "Email" TEXT NOT NULL DEFAULT '',
--     "Phone" TEXT NOT NULL DEFAULT ''
--   );
--
-- All values stored as TEXT. _rowId is a UUID generated by the browser.
-- Renames: UPDATE _tables SET name = ?  (no DDL needed)
-- Deletes: DROP TABLE t_42; DELETE FROM _tables WHERE id = 42;
```

## Worker API routes

All routes are under `/api/` and served by Pages Functions. Auth via session cookie (Google OAuth login flow handled by Worker).

### Auth
```
GET  /api/auth/login          -> redirect to Google OAuth
GET  /api/auth/callback       -> exchange code, set session cookie, upsert user
POST /api/auth/logout         -> clear session
GET  /api/auth/me             -> return current user info
```

### Books
```
GET    /api/books              -> list books for current user (via book_members)
POST   /api/books              -> create book + owner membership row
PATCH  /api/books/:id          -> rename book
DELETE /api/books/:id          -> delete book (cascade: tables, rows, charts, members)
POST   /api/books/:id/members  -> add/update member (sharing)
DELETE /api/books/:id/members/:userId -> remove member
```

### Tables (resolve name → t_{id} via _tables)
```
GET    /api/books/:bookId/tables                          -> list tables + schemas
POST   /api/books/:bookId/tables                          -> create table (INSERT _tables + CREATE TABLE t_{id})
PATCH  /api/books/:bookId/tables/:name                    -> rename table (UPDATE _tables, no DDL)
DELETE /api/books/:bookId/tables/:name                     -> DROP TABLE t_{id} + DELETE _tables row
PUT    /api/books/:bookId/tables/:name/schema              -> update columns/keys/sort (ALTER TABLE for col changes)
```

### Rows (resolve to t_{id}, _rowId is client-generated UUID)
```
GET    /api/books/:bookId/tables/:name/rows                -> SELECT * FROM t_{id}
PUT    /api/books/:bookId/tables/:name/rows/:rowId         -> UPDATE t_{id} SET col=? WHERE _rowId=?
POST   /api/books/:bookId/tables/:name/rows                -> INSERT INTO t_{id}
DELETE /api/books/:bookId/tables/:name/rows/:rowId         -> DELETE FROM t_{id} WHERE _rowId=?
POST   /api/books/:bookId/tables/:name/rows/bulk           -> bulk edit/delete (for multi-select)
```

### Charts (per-book DB)
```
GET    /api/books/:bookId/charts                           -> list chart sheets
POST   /api/books/:bookId/charts                           -> create chart sheet
PATCH  /api/books/:bookId/charts/:name                     -> update chart config
DELETE /api/books/:bookId/charts/:name                     -> delete chart sheet
```

### Images (R2, single bucket with book_id/ key prefix)
```
POST   /api/books/:bookId/images/upload-url                -> return presigned PUT URL for {bookId}/{uuid}.ext
GET    /api/books/:bookId/images/:key                      -> proxy or redirect to R2
```

### Import
```
POST   /api/books/:bookId/import/csv                       -> parse CSV, create table + rows
POST   /api/books/:bookId/import/google-sheet              -> Worker fetches sheet (holds API key), returns parsed CSV
```

## Browser-side data flow

### Cell edit (fire-and-forget with per-row queue)
```
User edits cell in AG Grid
  -> valueSetter fires
  -> push {rowId, column, oldValue} onto undo stack
  -> validate client-side (type, unique key, reference)
  -> if invalid: show error, don't apply
  -> update local rowData state (optimistic)
  -> enqueue to per-row write queue
  -> fetch('PUT /api/.../rows/:rowId', { column, value })
  -> on server error: revert cell, pop undo entry, show toast
```

### Row insert (awaited for UUID confirmation)
```
User fills draft row
  -> generate UUID as _rowId client-side
  -> validate client-side
  -> add row to local rowData (optimistic, already has permanent ID)
  -> push onto undo stack as insert
  -> fetch('POST .../rows', { _rowId, ...row data })
  -> on server error: remove row, show toast
```

### Row delete (awaited for reference check)
```
User deletes row
  -> validate references client-side (any other rows pointing here?)
  -> if blocked: show error, don't delete
  -> await fetch('DELETE .../rows/:rowId')
  -> on success: remove from local rowData, push onto undo stack
  -> on server error: show toast (row stays in grid)
```

### Undo (Ctrl+Z)
```
Pop last entry from undo stack
  -> reverse the operation (restore old value / re-insert / re-delete)
  -> fire the corresponding async write to server
```

### Auth middleware (every /api/ route)
```
Extract session cookie → look up user
  -> For /api/books/:bookId/*:
     SELECT role FROM book_members WHERE book_id = ? AND user_id = ?
  -> No row = 403
  -> role = 'viewer' on write = 403
```

## What gets deleted from current codebase
* `drive.ts` — entirely
* `dataModel.ts` — entirely
* `config.ts` — entirely (config is D1 _tables/_columns)
* `FolderPicker.tsx` — entirely
* `useAppState.ts` — rewrite from ~1,050 lines to ~300 lines of thin fetch wrappers
* `types.ts` — remove `Transaction`, `AppState`, `DriveFolder`, `DriveFile`, dirty tracking types
* Local workbook localStorage logic — entirely
* `google-api.d.ts` — entirely (no more GAPI in browser)
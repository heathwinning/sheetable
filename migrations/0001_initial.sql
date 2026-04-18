------------------------------------------------------------
-- Account & book management
------------------------------------------------------------

CREATE TABLE users (
  id         TEXT PRIMARY KEY,
  email      TEXT NOT NULL UNIQUE,
  name       TEXT,
  -- picture column removed
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE books (
  id         TEXT PRIMARY KEY,
  name       TEXT NOT NULL,
  owner_id   TEXT NOT NULL REFERENCES users(id),
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE book_members (
  book_id  TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  user_id  TEXT NOT NULL REFERENCES users(id),
  role     TEXT NOT NULL CHECK (role IN ('owner', 'editor', 'viewer')),
  PRIMARY KEY (book_id, user_id)
);

------------------------------------------------------------
-- Table/column metadata
------------------------------------------------------------

CREATE TABLE _tables (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  book_id         TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  display_order   INTEGER NOT NULL,
  unique_keys     TEXT NOT NULL DEFAULT '[]',
  default_sort    TEXT,
  draft_position  TEXT NOT NULL DEFAULT 'bottom',
  UNIQUE (book_id, name)
);

CREATE TABLE _columns (
  table_id        INTEGER NOT NULL REFERENCES _tables(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  display_name    TEXT,
  type            TEXT NOT NULL,
  display_order   INTEGER NOT NULL,
  ref_table       TEXT,
  ref_display     TEXT,
  ref_search      TEXT,
  PRIMARY KEY (table_id, name)
);

CREATE TABLE _chart_sheets (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  book_id         TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  table_name      TEXT,
  mode            TEXT NOT NULL DEFAULT 'edit',
  charts          TEXT NOT NULL DEFAULT '[]',
  display_order   INTEGER NOT NULL,
  UNIQUE (book_id, name)
);

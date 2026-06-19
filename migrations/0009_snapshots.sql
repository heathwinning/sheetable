-- Snapshots table for in-app backup/restore
-- Stores full book snapshots as JSON blobs for point-in-time recovery
CREATE TABLE _snapshots (
  id          TEXT PRIMARY KEY,
  book_id     TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  label       TEXT,
  created_at  TEXT NOT NULL DEFAULT (datetime('now')),
  table_count INTEGER NOT NULL DEFAULT 0,
  row_count   INTEGER NOT NULL DEFAULT 0,
  view_count  INTEGER NOT NULL DEFAULT 0,
  chart_count INTEGER NOT NULL DEFAULT 0,
  data        TEXT NOT NULL
);
CREATE INDEX idx_snapshots_book ON _snapshots(book_id);

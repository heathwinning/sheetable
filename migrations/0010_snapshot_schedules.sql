-- Snapshot schedule configuration
CREATE TABLE _snapshot_schedules (
  book_id     TEXT PRIMARY KEY REFERENCES books(id) ON DELETE CASCADE,
  enabled     INTEGER NOT NULL DEFAULT 0,
  frequency   TEXT NOT NULL DEFAULT 'daily' CHECK (frequency IN ('daily', 'weekly', 'monthly')),
  next_run_at TEXT,
  updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE _report_sheets (
  id TEXT PRIMARY KEY,
  book_id TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name TEXT NOT NULL,
  definition TEXT NOT NULL,
  revision INTEGER NOT NULL DEFAULT 1,
  display_order INTEGER NOT NULL,
  UNIQUE (book_id, name)
);

CREATE INDEX _report_sheets_order ON _report_sheets (book_id, display_order);

CREATE TABLE _view_sheets (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  book_id       TEXT NOT NULL REFERENCES books(id) ON DELETE CASCADE,
  name          TEXT NOT NULL,
  table_name    TEXT NOT NULL,
  view_type     TEXT NOT NULL DEFAULT 'calendar',
  date_column   TEXT,
  display_order INTEGER NOT NULL,
  UNIQUE (book_id, name)
);

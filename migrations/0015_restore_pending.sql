-- Keep a single-table restore out of sheet listings until its rows are ready.
ALTER TABLE _tables ADD COLUMN restore_pending INTEGER NOT NULL DEFAULT 0;

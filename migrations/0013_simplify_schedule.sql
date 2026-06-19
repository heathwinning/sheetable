-- Replace frequency/retention columns with simpler interval_days + retention_days
ALTER TABLE _snapshot_schedules ADD COLUMN interval_days INTEGER NOT NULL DEFAULT 1;
ALTER TABLE _snapshot_schedules ADD COLUMN retention_days INTEGER NOT NULL DEFAULT 30;

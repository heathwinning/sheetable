-- Configurable retention for snapshot schedules
ALTER TABLE _snapshot_schedules ADD COLUMN daily_retention_days INTEGER NOT NULL DEFAULT 30;
ALTER TABLE _snapshot_schedules ADD COLUMN keep_monthly INTEGER NOT NULL DEFAULT 1;

-- Move snapshot data from D1 to R2, add retention type
-- WARNING: this drops the data column. Existing D1-based snapshots will be lost.
-- Apply after deploying new code that writes to R2.
ALTER TABLE _snapshots ADD COLUMN data_key TEXT;
ALTER TABLE _snapshots ADD COLUMN type TEXT NOT NULL DEFAULT 'daily' CHECK (type IN ('daily', 'monthly'));
-- Drop old JSON blob column (data now lives in R2)
ALTER TABLE _snapshots DROP COLUMN data;

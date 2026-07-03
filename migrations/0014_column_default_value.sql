-- Add default_value column to _columns for auto-populating new rows
ALTER TABLE _columns ADD COLUMN default_value TEXT;

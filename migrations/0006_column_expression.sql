-- Add expression and show_in_grid columns to _columns for calculated column type
ALTER TABLE _columns ADD COLUMN expression TEXT;
ALTER TABLE _columns ADD COLUMN show_in_grid INTEGER NOT NULL DEFAULT 0;

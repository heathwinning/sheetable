-- Add calculated_columns column to _tables for storing expr-eval virtual columns
ALTER TABLE _tables ADD COLUMN calculated_columns TEXT;

// Internal row ID column name — hidden from user, used for stable row identity
export const INTERNAL_ROW_ID = '_rowId';

// Column types supported by the system
export type ColumnType = 'text' | 'integer' | 'decimal' | 'date' | 'datetime' | 'bool' | 'reference' | 'image';

export interface ColumnDef {
  name: string;
  displayName?: string; // optional display name shown in column header
  type: ColumnType;
  // For reference columns
  refTable?: string;
  refDisplayColumns?: string[]; // columns shown in the cell for a referenced row
  refSearchColumns?: string[];  // columns shown/searched in the dropdown editor
}

export interface TableSchema {
  name: string;
  columns: ColumnDef[];
  // Backing CSV filename for this table (default: <table name>.csv)
  csvFileName?: string;
  // Column names that together form the unique key (can be one or more)
  uniqueKeys: string[];
  // Default sort when opening the table
  defaultSort?: { column: string; direction: 'asc' | 'desc' }[];
  // Where the new-row draft appears: 'top' or 'bottom' (default: 'bottom')
  draftRowPosition?: 'top' | 'bottom';
}

export type Row = Record<string, string>;

export interface TableData {
  schema: TableSchema;
  rows: Row[];
}

export interface Transaction {
  id: number;
  tableId: string;
  type: 'update' | 'insert' | 'delete';
  rowIndex?: number; // for update/delete
  rowId?: string; // stable row identifier (preferred for delete/undo)
  columnName?: string; // for update
  oldValue?: string;
  newValue?: string;
  row?: Row; // for insert
  timestamp: number;
}

export interface ValidationError {
  message: string;
  rowIndex: number;
  columnName?: string;
}

export interface AppState {
  tables: Record<string, TableData>;
  schemas: Record<string, TableSchema>;
  generation: Record<string, number>; // dirty tracking per table
  lastSavedGeneration: Record<string, number>;
}

// Google Drive types
export interface DriveFolder {
  id: string;
  name: string;
}

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
}

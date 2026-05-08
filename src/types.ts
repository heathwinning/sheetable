// Internal row ID column name — hidden from user, used for stable row identity
export const INTERNAL_ROW_ID = '_rowId';

// Column types supported by the system
export type ColumnType = 'text' | 'integer' | 'decimal' | 'date' | 'datetime' | 'bool' | 'reference' | 'image';

export interface ColumnDef {
  name: string;
  displayName?: string;
  type: ColumnType;
  width?: number;
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
}

export interface TableSchema {
  name: string;
  columns: ColumnDef[];
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: 'asc' | 'desc' }[];
  draftRowPosition?: 'top' | 'bottom';
}

export type Row = Record<string, string>;

export interface TableData {
  schema: TableSchema;
  rows: Row[];
}

export interface ValidationError {
  message: string;
  rowIndex: number;
  columnName?: string;
}

// Chart sheet — stores Graphic Walker chart configurations
export interface ChartSheet {
  name: string;
  tableName?: string;
  mode?: 'edit' | 'display';
  charts: unknown[];

// View sheet — a named tab that shows a table in a specific view type
export interface ViewSheet {
  name: string;
  tableName: string;
  viewType: 'grid' | 'calendar' | 'schedule';
  dateColumn?: string;
}
}

// User session from API
export interface SessionUser {
  id: string;
  email: string;
  name: string;
}

// Book info from API
export interface BookInfo {
  id: string;
  name: string;
  owner_id: string;
  role: string;
  created_at: string;
}

// Book member from API
export interface BookMember {
  userId: string;
  email: string;
  name: string;
  role: string;
}

// Pending invite from API
export interface BookInvite {
  email: string;
  role: string;
  createdAt: string;
}

// Undo entry
export interface UndoEntry {
  type: 'update' | 'insert' | 'delete';
  tableId: string;
  rowId: string;
  column?: string;
  oldValue?: string;
  newValue?: string;
  row?: Row;
}

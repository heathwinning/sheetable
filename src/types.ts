// Internal row ID column name — hidden from user, used for stable row identity
export const INTERNAL_ROW_ID = '_rowId';

// Column types supported by the system
export type ColumnType = 'text' | 'integer' | 'decimal' | 'date' | 'datetime' | 'bool' | 'reference' | 'image' | 'calculated' | 'list';

export type ListItemType = 'text' | 'integer' | 'decimal' | 'date' | 'datetime' | 'bool' | 'image' | 'reference';

export const typeOptions: { value: ColumnType; label: string }[] = [
  { value: 'text', label: 'Text' },
  { value: 'integer', label: 'Integer' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'date', label: 'Date' },
  { value: 'datetime', label: 'Datetime' },
  { value: 'bool', label: 'Boolean' },
  { value: 'reference', label: 'Reference' },
  { value: 'image', label: 'Image' },
  { value: 'calculated', label: 'Calculated' },
  { value: 'list', label: 'List…' },
];

export const listItemTypeOptions: { value: ListItemType; label: string }[] = [
  { value: 'text', label: 'Text' },
  { value: 'integer', label: 'Integer' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'date', label: 'Date' },
  { value: 'datetime', label: 'Datetime' },
  { value: 'bool', label: 'Boolean' },
  { value: 'image', label: 'Image' },
  { value: 'reference', label: 'Reference' },
];

export interface ColumnDef {
  name: string;
  displayName?: string;
  type: ColumnType;
  width?: number;
  /** When true, the column keeps a fixed width and truncates long content with ellipsis rather than auto-fitted. */
  truncate?: boolean;
  // Reference columns
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
  // Calculated columns (type === 'calculated')
  expression?: string;
  showInGrid?: boolean;
  // List columns (type === 'list')
  listOf?: ListItemType;
}

/** @deprecated Use columns with type === 'calculated' instead. Kept for backward-compat migration. */
export interface CalculatedColumn {
  name: string;
  expression: string;
  showInGrid?: boolean;
}

export interface TableSchema {
  name: string;
  columns: ColumnDef[];
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: 'asc' | 'desc' }[];
  draftRowPosition?: 'top' | 'bottom';
  /** @deprecated Migrated to columns with type === 'calculated' on load. */
  calculatedColumns?: CalculatedColumn[];
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

// Chart types
export type ChartType = 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'table';
export type DateFeature = 'year' | 'quarter' | 'yearmonth' | 'month' | 'monthnum' | 'week' | 'dayofweek' | 'day' | 'hour';
// Column expression encoding: plain column name, or "colname:datefeature" for date/datetime columns
// e.g. "sale_date:year", "created_at:month"
export type AggregateFunc = 'sum' | 'count' | 'avg' | 'min' | 'max' | 'none';
export type FilterOperator = 'eq' | 'neq' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains' | 'is_empty' | 'is_not_empty';

/** @deprecated – replaced by valueCalc / valueFormat string fields on ChartConfig */
export interface ColumnModifier {
  multiplier?: number;
  divisor?: number;
  thousands?: boolean;
  decimals?: number;
  prefix?: string;
  suffix?: string;
}

export interface ChartConfig {
  id: string;
  title: string;
  type: ChartType;
  table: string;
  xColumn: string;
  yColumn: string;
  groupBy?: string;
  aggregate: AggregateFunc;
  xLabel?: string;
  yLabel?: string;
  tableRows?: string[];
  tableColumns?: string[];
  tableSort?: { key: string; dir: 'asc' | 'desc' };
  /** @deprecated use tableRowDimSort for per-dimension control */
  rowOrder?: 'natural' | 'label-asc' | 'label-desc' | 'value-asc' | 'value-desc';
  /** @deprecated use tableColDimSort for per-dimension control */
  colOrder?: 'natural' | 'label-asc' | 'label-desc' | 'value-asc' | 'value-desc';
  /** Per-dimension sort for table row dimensions. Each entry is 'asc', 'desc', or 'none'. */
  tableRowDimSort?: ('asc' | 'desc' | 'none')[];
  /** Per-dimension sort for table column dimensions. Each entry is 'asc', 'desc', or 'none'. */
  tableColDimSort?: ('asc' | 'desc' | 'none')[];
  /** Handlebars template for display. Variables: `value` (the aggregated number), `date`, row fields.
   *  Supports {{dateFormat date 'MMM D, YYYY'}}. e.g. "{{value}} kg" */
  valueFormat?: string;
  /** @deprecated use valueCalc / valueFormat */
  yModifier?: ColumnModifier;
  stacked?: boolean; // for bar charts only
  /** @deprecated use filters[] */
  filterColumn?: string;
  /** @deprecated use filters[] */
  filterOperator?: FilterOperator;
  /** @deprecated use filters[] */
  filterValue?: string;
  filters?: { column: string }[];
}

export interface ChartLayoutItem {
  i: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

// Chart sheet — a named dashboard of Recharts plots on a react-grid-layout canvas
export interface ChartSheet {
  name: string;
  charts: ChartConfig[];
  layout: ChartLayoutItem[];
}

// View sheet — a named tab that shows a table in a specific view type
export interface ViewSheet {
  name: string;
  tableName: string;
  viewType: 'grid' | 'calendar' | 'schedule';
  dateColumn?: string;
  hideSourceTableTab?: boolean;
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
  sheet_order?: string;
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

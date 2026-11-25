import type { ReportDefinition, ReportSummary } from './reportTypes';

export interface BackupColumn {
  name: string;
  displayName?: string;
  type: string;
  width?: number;
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
  expression?: string;
  showInGrid?: boolean;
  listOf?: string;
  defaultValue?: string;
}

interface ExportTableSchema {
  columns: BackupColumn[];
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: string }[];
  draftRowPosition?: string;
}

export interface BackupRow {
  _rowId: string;
  _resolved?: Record<string, string>;
  [column: string]: string | Record<string, string> | undefined;
}

export interface BackupTable {
  name: string;
  schema: ExportTableSchema | null; // set after resolveAllReferences
  rows: BackupRow[];
}

interface ExportView {
  name: string;
  tableName: string;
  viewType: string;
  dateColumn?: string;
}

interface ExportChart {
  name: string;
  tableName?: string;
  mode?: string;
  charts?: unknown;
}

export interface BackupPayload {
  version: number;
  exportedAt?: string;
  bookName: string;
  sheetOrder?: string;
  tables: BackupTable[];
  views?: ExportView[];
  charts?: ExportChart[];
  reports?: (ReportSummary & { definition: ReportDefinition })[];
}

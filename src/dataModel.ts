import type { TableData, TableSchema, Row, Transaction, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';

export class DataModel {
  private tables: Map<string, TableData> = new Map();
  private generation: Map<string, number> = new Map();
  private lastSavedGeneration: Map<string, number> = new Map();
  private transactionId = 0;
  private rowIdCounters: Map<string, number> = new Map();

  private nextRowId(tableId: string): string {
    const counter = (this.rowIdCounters.get(tableId) ?? 0) + 1;
    this.rowIdCounters.set(tableId, counter);
    return String(counter);
  }

  getTable(tableId: string): TableData | undefined {
    return this.tables.get(tableId);
  }

  getAllTables(): Map<string, TableData> {
    return this.tables;
  }

  getTableIds(): string[] {
    return Array.from(this.tables.keys());
  }

  isDirty(tableId: string): boolean {
    const gen = this.generation.get(tableId) ?? 0;
    const saved = this.lastSavedGeneration.get(tableId) ?? 0;
    return gen !== saved;
  }

  markSaved(tableId: string): void {
    const gen = this.generation.get(tableId) ?? 0;
    this.lastSavedGeneration.set(tableId, gen);
  }

  private bumpGeneration(tableId: string): void {
    const gen = this.generation.get(tableId) ?? 0;
    this.generation.set(tableId, gen + 1);
  }

  createTable(schema: TableSchema, rows: Row[] = []): void {
    // Preserve existing _rowId values when loading from storage, and assign
    // only for missing/duplicate IDs to keep references stable across reloads.
    let maxCounter = 0;

    for (const row of rows) {
      const existingId = row[INTERNAL_ROW_ID]?.trim();
      if (!existingId) continue;
      const n = Number(existingId);
      if (Number.isInteger(n) && n > maxCounter) {
        maxCounter = n;
      }
    }

    this.rowIdCounters.set(schema.name, maxCounter);
    const seen = new Set<string>();
    for (const row of rows) {
      const existingId = row[INTERNAL_ROW_ID]?.trim();
      if (existingId && !seen.has(existingId)) {
        row[INTERNAL_ROW_ID] = existingId;
        seen.add(existingId);
      } else {
        const nextId = this.nextRowId(schema.name);
        row[INTERNAL_ROW_ID] = nextId;
        seen.add(nextId);
      }
    }

    this.tables.set(schema.name, { schema, rows });
    this.generation.set(schema.name, 1);
    this.lastSavedGeneration.set(schema.name, 0);
  }

  deleteTable(tableId: string): void {
    this.tables.delete(tableId);
    this.generation.delete(tableId);
    this.lastSavedGeneration.delete(tableId);
    this.rowIdCounters.delete(tableId);
  }

  renameTable(oldName: string, newName: string): void {
    const table = this.tables.get(oldName);
    if (!table) return;
    table.schema.name = newName;
    this.tables.delete(oldName);
    this.tables.set(newName, table);
    // Move generation tracking
    const gen = this.generation.get(oldName) ?? 0;
    this.generation.delete(oldName);
    this.generation.set(newName, gen + 1);
    const savedGen = this.lastSavedGeneration.get(oldName) ?? 0;
    this.lastSavedGeneration.delete(oldName);
    this.lastSavedGeneration.set(newName, savedGen);
    const counter = this.rowIdCounters.get(oldName) ?? 0;
    this.rowIdCounters.delete(oldName);
    this.rowIdCounters.set(newName, counter);
    // Update references in other tables
    for (const [, otherTable] of this.tables) {
      for (const col of otherTable.schema.columns) {
        if (col.refTable === oldName) {
          col.refTable = newName;
        }
      }
    }
  }

  updateSchema(tableId: string, newSchema: TableSchema): void {
    const table = this.tables.get(tableId);
    if (!table) return;
    const oldColumns = new Set(table.schema.columns.map(c => c.name));
    const newColumns = new Set(newSchema.columns.map(c => c.name));
    // Add empty values for new columns in existing rows
    for (const row of table.rows) {
      for (const col of newSchema.columns) {
        if (!oldColumns.has(col.name) && !(col.name in row)) {
          row[col.name] = '';
        }
      }
      // Remove data for deleted columns
      for (const oldCol of oldColumns) {
        if (!newColumns.has(oldCol)) {
          delete row[oldCol];
        }
      }
    }
    table.schema = newSchema;
    this.bumpGeneration(tableId);
  }

  // Apply a transaction and return validation errors (empty = success)
  applyTransaction(tx: Transaction): ValidationError[] {
    const table = this.tables.get(tx.tableId);
    if (!table) {
      return [{ message: `Table "${tx.tableId}" not found`, rowIndex: -1 }];
    }

    switch (tx.type) {
      case 'update':
        return this.applyUpdate(table, tx);
      case 'insert':
        return this.applyInsert(table, tx);
      case 'delete':
        return this.applyDelete(table, tx);
      default:
        return [{ message: `Unknown transaction type`, rowIndex: -1 }];
    }
  }

  private applyUpdate(table: TableData, tx: Transaction): ValidationError[] {
    if (tx.rowIndex === undefined || tx.columnName === undefined || tx.newValue === undefined) {
      return [{ message: 'Invalid update transaction', rowIndex: -1 }];
    }

    const col = table.schema.columns.find(c => c.name === tx.columnName);
    if (!col) {
      return [{ message: `Column "${tx.columnName}" not found`, rowIndex: tx.rowIndex }];
    }

    // Validate type constraint
    const typeErrors = this.validateType(col.type, tx.newValue, tx.rowIndex, tx.columnName);
    if (typeErrors.length > 0) return typeErrors;

    // Validate unique key constraint (if this column is part of the unique key)
    if ((table.schema.uniqueKeys ?? []).includes(tx.columnName)) {
      // Build the composite key for the row after applying the change
      const keyValues: Record<string, string> = {};
      for (const keyCol of (table.schema.uniqueKeys ?? [])) {
        keyValues[keyCol] = keyCol === tx.columnName ? tx.newValue : table.rows[tx.rowIndex][keyCol];
      }
      const errors = this.validateUniqueKey(table, keyValues, tx.rowIndex);
      if (errors.length > 0) return errors;
    }

    // Validate reference constraint (stored value is the _rowId of the referenced row)
    if (col.type === 'reference' && col.refTable && tx.newValue !== '') {
      const errors = this.validateReference(col.refTable, tx.newValue, tx.rowIndex);
      if (errors.length > 0) return errors;
    }

    table.rows[tx.rowIndex][tx.columnName] = tx.newValue;
    this.bumpGeneration(tx.tableId);
    return [];
  }

  private applyInsert(table: TableData, tx: Transaction): ValidationError[] {
    if (!tx.row) {
      return [{ message: 'Invalid insert transaction', rowIndex: -1 }];
    }

    const newRowIndex = table.rows.length;

    // Validate type constraints on the new row
    for (const col of table.schema.columns) {
      const value = tx.row[col.name] ?? '';
      const typeErrors = this.validateType(col.type, value, newRowIndex, col.name);
      if (typeErrors.length > 0) return typeErrors;

      if (col.type === 'reference' && col.refTable && value !== '') {
        const errors = this.validateReference(col.refTable, value, newRowIndex);
        if (errors.length > 0) return errors;
      }
    }

    // Validate unique key constraint
    if ((table.schema.uniqueKeys ?? []).length > 0) {
      const keyValues: Record<string, string> = {};
      for (const keyCol of (table.schema.uniqueKeys ?? [])) {
        keyValues[keyCol] = tx.row[keyCol] ?? '';
      }
      const errors = this.validateUniqueKey(table, keyValues, -1);
      if (errors.length > 0) return errors;
    }

    // Fill in missing columns with empty strings, assign _rowId
    const completeRow: Row = {};
    completeRow[INTERNAL_ROW_ID] = this.nextRowId(tx.tableId);
    for (const col of table.schema.columns) {
      completeRow[col.name] = tx.row[col.name] ?? '';
    }

    table.rows.push(completeRow);
    this.bumpGeneration(tx.tableId);
    return [];
  }

  private applyDelete(table: TableData, tx: Transaction): ValidationError[] {
    if (tx.rowIndex === undefined || tx.rowIndex < 0 || tx.rowIndex >= table.rows.length) {
      return [{ message: 'Invalid delete transaction', rowIndex: tx.rowIndex ?? -1 }];
    }

    // Check if any reference columns in other tables point to this row's _rowId
    const deletedRowId = table.rows[tx.rowIndex][INTERNAL_ROW_ID];
    for (const [, otherTable] of this.tables) {
      for (const col of otherTable.schema.columns) {
        if (col.type === 'reference' && col.refTable === table.schema.name) {
          for (let i = 0; i < otherTable.rows.length; i++) {
            if (otherTable.rows[i][col.name] === deletedRowId) {
              return [{
                message: `Cannot delete this row because it is referenced by the "${otherTable.schema.name}" table`,
                rowIndex: tx.rowIndex
              }];
            }
          }
        }
      }
    }

    table.rows.splice(tx.rowIndex, 1);
    this.bumpGeneration(tx.tableId);
    return [];
  }

  private validateUniqueKey(table: TableData, keyValues: Record<string, string>, excludeRow: number): ValidationError[] {
    const keyColumns = Object.keys(keyValues);

    // All key columns must be non-empty
    for (const colName of keyColumns) {
      if (keyValues[colName] === '') {
        return [{ message: `Key column "${colName}" cannot be empty`, rowIndex: excludeRow, columnName: colName }];
      }
    }

    // Check for duplicate composite key
    for (let i = 0; i < table.rows.length; i++) {
      if (i === excludeRow) continue;
      const matches = keyColumns.every(col => table.rows[i][col] === keyValues[col]);
      if (matches) {
        const keyStr = keyColumns.map(c => `${c}="${keyValues[c]}"`).join(', ');
        return [{
          message: keyColumns.length === 1
            ? `Duplicate key "${keyValues[keyColumns[0]]}" in column "${keyColumns[0]}"`
            : `Duplicate composite key: ${keyStr}`,
          rowIndex: excludeRow,
          columnName: keyColumns[0]
        }];
      }
    }
    return [];
  }

  private validateType(type: string, value: string, rowIndex: number, columnName: string): ValidationError[] {
    if (value === '') return []; // empty is always ok

    switch (type) {
      case 'integer':
        if (!/^-?\d+$/.test(value)) {
          return [{ message: `"${value}" is not a valid integer`, rowIndex, columnName }];
        }
        break;
      case 'decimal':
        if (isNaN(Number(value)) || value.trim() === '') {
          return [{ message: `"${value}" is not a valid decimal`, rowIndex, columnName }];
        }
        break;
      case 'date':
        if (!/^\d{4}-\d{2}-\d{2}$/.test(value) || isNaN(Date.parse(value))) {
          return [{ message: `"${value}" is not a valid date (YYYY-MM-DD)`, rowIndex, columnName }];
        }
        break;
      case 'datetime':
        if (isNaN(Date.parse(value))) {
          return [{ message: `"${value}" is not a valid datetime`, rowIndex, columnName }];
        }
        break;
      case 'bool':
        if (!['true', 'false', '1', '0', 'yes', 'no'].includes(value.toLowerCase())) {
          return [{ message: `"${value}" is not a valid boolean (true/false)`, rowIndex, columnName }];
        }
        break;
    }
    return [];
  }

  private validateReference(refTable: string, rowId: string, rowIndex: number): ValidationError[] {
    const target = this.tables.get(refTable);
    if (!target) {
      return [{ message: `Referenced table "${refTable}" not found`, rowIndex }];
    }
    const exists = target.rows.some(r => r[INTERNAL_ROW_ID] === rowId);
    if (!exists) {
      return [{ message: `Referenced row not found in "${refTable}"`, rowIndex }];
    }
    return [];
  }

  // Get all rows from a referenced table (for dropdown options)
  getReferenceRows(refTable: string): Row[] {
    const target = this.tables.get(refTable);
    if (!target) return [];
    return target.rows;
  }

  // Look up a single referenced row by _rowId
  getReferencedRow(refTable: string, rowId: string): Row | undefined {
    const target = this.tables.get(refTable);
    if (!target) return undefined;
    return target.rows.find(r => r[INTERNAL_ROW_ID] === rowId);
  }

  /**
   * Resolve a dot-notation column path through reference chains.
   * E.g. "city.name" on table "customers" means: look up the "city" column (a reference),
   * follow it to the referenced row, and return the "name" column value.
   * Supports arbitrary depth (e.g. "city.country.name").
   * Returns '' if any link in the chain is missing.
   */
  resolveColumnPath(tableName: string, row: Row, path: string): string {
    const parts = path.split('.');
    const tableData = this.tables.get(tableName);
    if (!tableData) return '';

    const colName = parts[0];
    const value = row[colName] ?? '';

    if (parts.length === 1) {
      // Single column — but check if this column is a reference and auto-resolve
      const col = tableData.schema.columns.find(c => c.name === colName);
      if (col?.type === 'reference' && col.refTable && value) {
        const refRow = this.getReferencedRow(col.refTable, value);
        if (!refRow) return '';
        // Use the reference's own display columns to resolve
        const displayCols = col.refDisplayColumns ?? [];
        if (displayCols.length > 0) {
          return displayCols.map(dc => this.resolveColumnPath(col.refTable!, refRow, dc)).filter(Boolean).join(' · ');
        }
        return value;
      }
      return value;
    }

    // Multi-part path: first part must be a reference column
    const col = tableData.schema.columns.find(c => c.name === colName);
    if (!col || col.type !== 'reference' || !col.refTable || !value) return '';

    const refRow = this.getReferencedRow(col.refTable, value);
    if (!refRow) return '';

    // Recurse with remaining path parts
    return this.resolveColumnPath(col.refTable, refRow, parts.slice(1).join('.'));
  }

  /**
   * Get available column paths for a table, expanding reference columns one level deep.
   * Returns paths like ["name", "city.name", "city.population"] for use in ref config.
   */
  getColumnPaths(tableName: string): { path: string; label: string }[] {
    const tableData = this.tables.get(tableName);
    if (!tableData) return [];

    const result: { path: string; label: string }[] = [];
    for (const col of tableData.schema.columns) {
      if (col.type === 'reference' && col.refTable) {
        const refTableData = this.tables.get(col.refTable);
        if (refTableData) {
          for (const refCol of refTableData.schema.columns) {
            result.push({
              path: `${col.name}.${refCol.name}`,
              label: `${col.name} → ${refCol.name}`,
            });
          }
        }
      } else {
        result.push({ path: col.name, label: col.name });
      }
    }
    return result;
  }

  // @deprecated — use getReferenceRows instead
  getReferenceValues(refTable: string): string[] {
    const target = this.tables.get(refTable);
    if (!target) return [];
    return target.rows.map(r => r[INTERNAL_ROW_ID]).filter(v => v !== '');
  }

  nextTransactionId(): number {
    return ++this.transactionId;
  }
}

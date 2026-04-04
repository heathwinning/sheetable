import type { TableSchema, Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { normalizeTemporalString } from './dateFormat';

// Parse CSV text into rows
export function parseCSV(text: string): string[][] {
  const rows: string[][] = [];
  let current = '';
  let inQuotes = false;
  let row: string[] = [];

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        current += '"';
        i++; // skip next quote
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        row.push(current);
        current = '';
      } else if (ch === '\n') {
        row.push(current);
        current = '';
        if (row.length > 0 && !(row.length === 1 && row[0] === '')) {
          rows.push(row);
        }
        row = [];
      } else if (ch === '\r') {
        // skip carriage return
      } else {
        current += ch;
      }
    }
  }

  // Last field/row
  row.push(current);
  if (row.length > 0 && !(row.length === 1 && row[0] === '')) {
    rows.push(row);
  }

  return rows;
}

// Convert parsed CSV to Row objects given a schema
export function csvToRows(csvText: string, schema: TableSchema): Row[] {
  const parsed = parseCSV(csvText);
  if (parsed.length === 0) return [];

  const headers = parsed[0];
  const columnTypeByName = new Map(schema.columns.map(col => [col.name, col.type]));
  const rows: Row[] = [];

  for (let i = 1; i < parsed.length; i++) {
    const row: Row = {};
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j];
      const raw = parsed[i][j] ?? '';
      const colType = columnTypeByName.get(header);
      row[header] = colType ? normalizeTemporalString(raw, colType) : raw;
    }
    // Fill in any schema columns not in CSV
    for (const col of schema.columns) {
      if (!(col.name in row)) {
        row[col.name] = '';
      }
    }
    // Ensure the internal row ID key exists for downstream logic.
    if (!(INTERNAL_ROW_ID in row)) {
      row[INTERNAL_ROW_ID] = '';
    }
    rows.push(row);
  }

  return rows;
}

// Serialize rows to CSV string
export function rowsToCSV(schema: TableSchema, rows: Row[]): string {
  // Persist hidden internal IDs so references remain stable across reloads.
  const headers = [INTERNAL_ROW_ID, ...schema.columns.map(c => c.name)];
  const columnTypeByName = new Map(schema.columns.map(col => [col.name, col.type]));
  const lines: string[] = [];

  lines.push(headers.map(escapeCSVField).join(','));

  for (const row of rows) {
    const fields = headers.map(h => {
      const raw = row[h] ?? '';
      const colType = columnTypeByName.get(h);
      const normalized = colType ? normalizeTemporalString(raw, colType) : raw;
      return escapeCSVField(normalized);
    });
    lines.push(fields.join(','));
  }

  return lines.join('\n') + '\n';
}

function escapeCSVField(value: string): string {
  if (value.includes(',') || value.includes('"') || value.includes('\n')) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}

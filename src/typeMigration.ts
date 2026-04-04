import type { ColumnType, Row } from './types';
import { DATE_FORMATS } from './dateFormatsList';

export interface ConversionResult {
  value: string;
  error?: string;
}

export interface MigrationPreview {
  columnName: string;
  fromType: ColumnType;
  toType: ColumnType;
  totalRows: number;
  /** Number of non-empty values */
  nonEmptyCount: number;
  /** Successfully convertible values */
  successCount: number;
  /** Values that will be blanked out */
  errorCount: number;
  /** Sample rows: [original, converted, error?] */
  samples: { original: string; converted: string; error?: string }[];
}

export interface PendingMigration {
  columnName: string;
  fromType: ColumnType;
  toType: ColumnType;
}

/**
 * Convert a single value from one type to another.
 * Empty strings always convert to empty strings.
 */
export function convertValue(value: string, fromType: ColumnType, toType: ColumnType, dateFormat?: string): ConversionResult {
  if (value === '') return { value: '' };
  if (fromType === toType) return { value };

  // Any type → text: keep as-is
  if (toType === 'text') return { value };

  // → integer
  if (toType === 'integer') {
    if (fromType === 'decimal') {
      const n = parseFloat(value);
      if (isNaN(n)) return { value: '', error: `"${value}" is not a number` };
      return { value: String(Math.round(n)) };
    }
    if (fromType === 'bool') {
      const lower = value.toLowerCase();
      if (['true', '1', 'yes'].includes(lower)) return { value: '1' };
      if (['false', '0', 'no'].includes(lower)) return { value: '0' };
      return { value: '', error: `"${value}" is not a valid boolean` };
    }
    // text or other → integer
    const trimmed = value.trim();
    // Handle decimal strings by rounding
    const n = Number(trimmed);
    if (trimmed === '' || isNaN(n)) return { value: '', error: `"${value}" is not a number` };
    return { value: String(Math.round(n)) };
  }

  // → decimal
  if (toType === 'decimal') {
    if (fromType === 'integer') return { value }; // integers are valid decimals
    if (fromType === 'bool') {
      const lower = value.toLowerCase();
      if (['true', '1', 'yes'].includes(lower)) return { value: '1' };
      if (['false', '0', 'no'].includes(lower)) return { value: '0' };
      return { value: '', error: `"${value}" is not a valid boolean` };
    }
    const n = Number(value.trim());
    if (isNaN(n) || value.trim() === '') return { value: '', error: `"${value}" is not a number` };
    return { value: String(n) };
  }

  // → date
  if (toType === 'date') {
    if (fromType === 'datetime') {
      // Strip time: take YYYY-MM-DD portion
      const match = value.match(/^\d{4}-\d{2}-\d{2}/);
      if (match) return { value: match[1] };
      const d = new Date(value);
      if (!isNaN(d.getTime())) {
        return { value: d.toISOString().slice(0, 10) };
      }
      return { value: '', error: `"${value}" is not a valid date` };
    }
    // text or other → date
    if (dateFormat) {
      const fmt = DATE_FORMATS.find(f => f.value === dateFormat);
      if (fmt) {
        const parsed = fmt.parse(value);
        if (parsed) return { value: parsed };
      }
      return { value: '', error: `"${value}" does not match selected format` };
    }
    return { value: '', error: `"${value}" cannot be parsed as a date` };
  }

  // → datetime
  if (toType === 'datetime') {
    if (fromType === 'date') {
      if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
        return { value: value + 'T00:00' };
      }
    }
    if (dateFormat) {
      const fmt = DATE_FORMATS.find(f => f.value === dateFormat);
      if (fmt) {
        const parsed = fmt.parse(value);
        if (parsed) return { value: parsed + 'T00:00' };
      }
      return { value: '', error: `"${value}" does not match selected format` };
    }
    return { value: '', error: `"${value}" cannot be parsed as a datetime` };
  }

  // → bool
  if (toType === 'bool') {
    const lower = value.toLowerCase().trim();
    if (['true', '1', 'yes'].includes(lower)) return { value: 'true' };
    if (['false', '0', 'no'].includes(lower)) return { value: 'false' };
    if (fromType === 'integer' || fromType === 'decimal') {
      const n = Number(value);
      if (!isNaN(n)) return { value: n !== 0 ? 'true' : 'false' };
    }
    return { value: '', error: `"${value}" cannot be converted to boolean` };
  }

  // → reference: can't auto-convert
  if (toType === 'reference') {
    return { value: '', error: 'Cannot auto-convert to reference' };
  }

  // → image: can't auto-convert
  if (toType === 'image') {
    if (fromType === 'text') {
      // If it looks like a URL, keep it
      if (/^https?:\/\//i.test(value.trim())) return { value };
    }
    return { value: '', error: 'Cannot auto-convert to image' };
  }

  // Fallback: keep value
  return { value };
}

/**
 * Generate a preview of what a type migration would do to table data.
 */
export function previewMigration(
  rows: Row[],
  columnName: string,
  fromType: ColumnType,
  toType: ColumnType,
  dateFormat?: string,
): MigrationPreview {
  let nonEmptyCount = 0;
  let successCount = 0;
  let errorCount = 0;
  const samples: MigrationPreview['samples'] = [];

  for (const row of rows) {
    const original = row[columnName] ?? '';
    if (original === '') continue;
    nonEmptyCount++;

    const result = convertValue(original, fromType, toType, dateFormat);
    if (result.error) {
      errorCount++;
    } else {
      successCount++;
    }

    // Collect up to 10 samples, prioritizing errors then successes
    if (samples.length < 10) {
      samples.push({
        original,
        converted: result.value,
        error: result.error,
      });
    }
  }

  // Sort samples: errors first, then successes
  samples.sort((a, b) => {
    if (a.error && !b.error) return -1;
    if (!a.error && b.error) return 1;
    return 0;
  });

  return {
    columnName,
    fromType,
    toType,
    totalRows: rows.length,
    nonEmptyCount,
    successCount,
    errorCount,
    samples,
  };
}

/**
 * Apply a type migration to rows in-place.
 * Returns the number of values that were blanked due to conversion errors.
 */
export function applyMigration(
  rows: Row[],
  columnName: string,
  fromType: ColumnType,
  toType: ColumnType,
  dateFormat?: string,
): number {
  let errorCount = 0;
  for (const row of rows) {
    const original = row[columnName] ?? '';
    if (original === '') continue;
    const result = convertValue(original, fromType, toType, dateFormat);
    row[columnName] = result.value;
    if (result.error) errorCount++;
  }
  return errorCount;
}

export interface ExtractPreview {
  sourceColumns: string[];
  uniqueTuples: string[][];
  totalRows: number;
  nonEmptyCount: number;
}

/**
 * Preview extracting unique tuples from one or more columns into a new reference table.
 */
export function previewExtract(
  rows: Row[],
  columnNames: string[],
): ExtractPreview {
  const seen = new Set<string>();
  const tuples: string[][] = [];
  let nonEmptyCount = 0;
  for (const row of rows) {
    const values = columnNames.map(cn => row[cn] ?? '');
    if (values.every(v => v === '')) continue;
    nonEmptyCount++;
    const key = JSON.stringify(values);
    if (!seen.has(key)) {
      seen.add(key);
      tuples.push(values);
    }
  }
  tuples.sort((a, b) => {
    for (let i = 0; i < a.length; i++) {
      const cmp = a[i].localeCompare(b[i]);
      if (cmp !== 0) return cmp;
    }
    return 0;
  });
  return {
    sourceColumns: columnNames,
    uniqueTuples: tuples,
    totalRows: rows.length,
    nonEmptyCount,
  };
}

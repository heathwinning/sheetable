import React, { useState, useMemo, useCallback } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import Select from 'react-select';
import type { UseAppStateReturn } from './useAppState';
import type { ColumnDef, ColumnType, Row, TableSchema } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseCSV } from './csv';
import * as drive from './drive';
import { useAlert } from './DialogProvider';

interface ImportPageProps {
  state: UseAppStateReturn;
}

// Date format patterns for parsing
const DATE_FORMATS: { value: string; label: string; parse: (s: string) => string | null }[] = [
  {
    value: 'yyyy-mm-dd',
    label: 'YYYY-MM-DD (2026-04-03)',
    parse: (s) => {
      const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'dd/mm/yyyy',
    label: 'DD/MM/YYYY (03/04/2026)',
    parse: (s) => {
      const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'mm/dd/yyyy',
    label: 'MM/DD/YYYY (04/03/2026)',
    parse: (s) => {
      const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      return m ? `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'dd-mm-yyyy',
    label: 'DD-MM-YYYY (03-04-2026)',
    parse: (s) => {
      const m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
      return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'dd.mm.yyyy',
    label: 'DD.MM.YYYY (03.04.2026)',
    parse: (s) => {
      const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
      return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'yyyy/mm/dd',
    label: 'YYYY/MM/DD (2026/04/03)',
    parse: (s) => {
      const m = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
      return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'd-mmm-yyyy',
    label: 'D-Mon-YYYY (3-Apr-2026)',
    parse: (s) => {
      const months: Record<string, string> = { jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12' };
      const m = s.match(/^(\d{1,2})[- ](\w{3})[- ](\d{4})$/);
      if (!m) return null;
      const mon = months[m[2].toLowerCase()];
      return mon ? `${m[3]}-${mon}-${m[1].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'mmm-d-yyyy',
    label: 'Mon D, YYYY (Apr 3, 2026)',
    parse: (s) => {
      const months: Record<string, string> = { jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12' };
      const m = s.match(/^(\w{3})[- ](\d{1,2}),?\s*(\d{4})$/);
      if (!m) return null;
      const mon = months[m[1].toLowerCase()];
      return mon ? `${m[3]}-${mon}-${m[2].padStart(2, '0')}` : null;
    },
  },
  {
    value: 'auto',
    label: 'Auto-detect',
    parse: (s) => {
      // Try ISO first
      const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (iso) return `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`;
      // Try Date constructor
      const d = new Date(s);
      if (!isNaN(d.getTime())) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
      }
      return null;
    },
  },
];

interface ColumnMapping {
  sourceColumn: string;
  targetColumn: string;
  enabled: boolean;
  dateFormat?: string;
  // For reference columns: source column → ref table column pairs for matching
  refSourceMappings?: { sourceColumn: string; refColumn: string }[];
}

const selectStyles = {
  control: (base: Record<string, unknown>, s: { isFocused: boolean }) => ({
    ...base,
    background: 'var(--bg)',
    borderColor: s.isFocused ? 'var(--ref-color)' : 'var(--border)',
    boxShadow: s.isFocused ? '0 0 0 2px rgba(37, 99, 235, 0.2)' : 'none',
    borderRadius: 4,
    fontSize: 13,
    minHeight: 30,
    '&:hover': { borderColor: 'var(--ref-color)' },
  }),
  menu: (base: Record<string, unknown>) => ({
    ...base,
    background: '#fff',
    border: '1px solid var(--border)',
    borderRadius: 4,
    boxShadow: '0 4px 16px rgba(0,0,0,0.10)',
    zIndex: 10,
  }),
  option: (base: Record<string, unknown>, s: { isSelected: boolean; isFocused: boolean }) => ({
    ...base,
    background: s.isSelected ? 'var(--primary)' : s.isFocused ? 'var(--cell-selected)' : 'transparent',
    color: s.isSelected ? '#fff' : 'var(--text)',
    fontSize: 13,
    padding: '4px 10px',
    cursor: 'pointer',
  }),
  singleValue: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  input: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  placeholder: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text-muted)' }),
  indicatorSeparator: () => ({ display: 'none' }),
};

// Column type options for new table mode
const typeOptions: { value: ColumnType; label: string }[] = [
  { value: 'text', label: 'Text' },
  { value: 'integer', label: 'Integer' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'date', label: 'Date' },
  { value: 'datetime', label: 'Datetime' },
  { value: 'bool', label: 'Boolean' },
  { value: 'reference', label: 'Reference' },
];

// New table column definition (extends ColumnMapping)
interface NewTableColumn {
  sourceColumn: string;
  enabled: boolean;
  name: string;
  type: ColumnType;
  dateFormat?: string;
  refTable?: string;
  refSourceMappings?: { sourceColumn: string; refColumn: string }[];
}

function guessColumnType(values: string[]): ColumnType {
  const samples = values.filter(v => v.trim()).slice(0, 20);
  if (samples.length === 0) return 'text';
  if (samples.every(v => /^-?\d+$/.test(v.trim()))) return 'integer';
  if (samples.every(v => /^-?\d+\.?\d*$/.test(v.trim()))) return 'decimal';
  if (samples.every(v => /^\d{4}-\d{1,2}-\d{1,2}$/.test(v.trim()))) return 'date';
  if (samples.every(v => /^(true|false)$/i.test(v.trim()))) return 'bool';
  return 'text';
}

export const ImportPage: React.FC<ImportPageProps> = ({ state }) => {
  const { tableId } = useParams<{ tableId: string }>();
  const navigate = useNavigate();
  const showAlert = useAlert();
  const isNewTable = !tableId;

  // Source data
  const [sourceHeaders, setSourceHeaders] = useState<string[]>([]);
  const [sourceRows, setSourceRows] = useState<string[][]>([]);
  const [sourceFileName, setSourceFileName] = useState<string>('');

  // Column mappings (existing table mode)
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);

  // New table config
  const [newTableName, setNewTableName] = useState('');
  const [newTableColumns, setNewTableColumns] = useState<NewTableColumn[]>([]);

  // Import state
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState<{ imported: number; errors: string[] } | null>(null);
  const [showSheetDialog, setShowSheetDialog] = useState(false);
  const [sheetUrl, setSheetUrl] = useState('');
  const [sheetLoading, setSheetLoading] = useState(false);
  const [sheetTabs, setSheetTabs] = useState<drive.SheetTab[]>([]);
  const [selectedSheetGid, setSelectedSheetGid] = useState<number | undefined>(undefined);
  const [spreadsheetId, setSpreadsheetId] = useState<string | null>(null);
  const [createdTableId, setCreatedTableId] = useState<string | null>(null);

  const schema = tableId ? state.getSchema(tableId) : undefined;

  // Build target column options
  const targetColumnOptions = useMemo(() => {
    if (!schema) return [];
    return [
      { value: '', label: '(Skip)' },
      ...schema.columns.map(c => ({ value: c.name, label: c.displayName || c.name })),
    ];
  }, [schema]);

  // Source column options for reference matching pairs
  const sourceColumnOptions = useMemo(() =>
    sourceHeaders.map(h => ({ value: h, label: h })),
    [sourceHeaders]
  );

  // Parse uploaded CSV
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setSourceFileName(file.name);
    const reader = new FileReader();
    reader.onload = () => {
      const text = reader.result as string;
      loadCSVData(text);
    };
    reader.readAsText(file);
  };

  // Import from Google Sheet
  const handleGoogleSheetImport = useCallback(async () => {
    if (!state.folderId) {
      showAlert('Connect to Google Drive first');
      return;
    }
    setShowSheetDialog(true);
    setSheetUrl('');
    setSheetTabs([]);
    setSpreadsheetId(null);
  }, [state.folderId, showAlert]);

  const handleSheetUrlSubmit = async () => {
    if (!sheetUrl.trim()) return;
    const match = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) {
      showAlert('Invalid Google Sheets URL');
      return;
    }
    setSheetLoading(true);
    try {
      const ssId = match[1];
      // Try to fetch sheet tabs; fall back to direct import if Sheets API unavailable
      let tabs: drive.SheetTab[] = [];
      try {
        tabs = await drive.getSpreadsheetSheets(ssId);
      } catch (err) {
        // Sheets API may not be enabled — fall back to importing first sheet
        console.warn('Could not fetch sheet tabs (enable Google Sheets API in Cloud Console for sheet selection):', err);
        tabs = [];
      }
      setSpreadsheetId(ssId);

      if (tabs.length > 1) {
        // Multiple sheets — let user pick
        setSheetTabs(tabs);
        setSelectedSheetGid(tabs[0]?.sheetId);
        setSheetLoading(false);
      } else {
        // Single sheet or couldn't fetch tabs — import directly
        const gid = tabs.length === 1 ? tabs[0].sheetId : undefined;
        const sheetName = tabs.length === 1 ? tabs[0].title : undefined;
        const csvText = await drive.exportSheetAsCSV(ssId, gid);
        setSourceFileName(sheetName ? `Google Sheet — ${sheetName}` : 'Google Sheet');
        setShowSheetDialog(false);
        setSheetTabs([]);
        loadCSVData(csvText);
      }
    } catch (err) {
      console.error('Failed to import from Google Sheets:', err);
      showAlert('Failed to import from Google Sheets. Make sure the sheet is shared and you have access.');
      setSheetLoading(false);
    }
  };

  const handleSheetTabImport = async () => {
    if (!spreadsheetId || selectedSheetGid === undefined) return;
    setSheetLoading(true);
    try {
      const tab = sheetTabs.find(t => t.sheetId === selectedSheetGid);
      const csvText = await drive.exportSheetAsCSV(spreadsheetId, selectedSheetGid);
      setSourceFileName(`Google Sheet — ${tab?.title ?? 'Sheet'}`);
      setShowSheetDialog(false);
      setSheetTabs([]);
      loadCSVData(csvText);
    } catch (err) {
      console.error('Failed to import sheet tab:', err);
      showAlert('Failed to import the selected sheet.');
    } finally {
      setSheetLoading(false);
    }
  };

  const loadCSVData = (text: string) => {
    const parsed = parseCSV(text);
    if (parsed.length < 2) {
      showAlert('CSV file is empty or has no data rows');
      return;
    }
    const headers = parsed[0];
    const rows = parsed.slice(1);
    setSourceHeaders(headers);
    setSourceRows(rows);
    setImportResult(null);

    if (isNewTable) {
      // New table mode: auto-generate column definitions with guessed types
      const cols: NewTableColumn[] = headers.map((header, idx) => {
        const colValues = rows.map(r => r[idx] ?? '');
        return {
          sourceColumn: header,
          enabled: true,
          name: header,
          type: guessColumnType(colValues),
          dateFormat: guessColumnType(colValues) === 'date' ? 'auto' : undefined,
        };
      });
      setNewTableColumns(cols);
    } else if (schema) {
      // Existing table mode: auto-map source columns to target columns by name match
      const autoMappings: ColumnMapping[] = headers.map(header => {
        const matchedCol = schema.columns.find(
          c => c.name.toLowerCase() === header.toLowerCase() ||
               (c.displayName && c.displayName.toLowerCase() === header.toLowerCase())
        );
        const mapping: ColumnMapping = {
          sourceColumn: header,
          targetColumn: matchedCol?.name ?? '',
          enabled: !!matchedCol,
        };
        // Auto-set date format for date columns
        if (matchedCol && (matchedCol.type === 'date' || matchedCol.type === 'datetime')) {
          mapping.dateFormat = 'auto';
        }
        // Auto-set ref match columns for reference columns
        if (matchedCol && matchedCol.type === 'reference' && matchedCol.refTable) {
          const refSchema = state.getSchema(matchedCol.refTable);
          if (refSchema) {
            const matchCols = matchedCol.refDisplayColumns?.length
              ? matchedCol.refDisplayColumns
              : matchedCol.refSearchColumns?.length
                ? matchedCol.refSearchColumns
                : [refSchema.columns[0]?.name].filter(Boolean);
            mapping.refSourceMappings = matchCols.length > 0
              ? [{ sourceColumn: header, refColumn: matchCols[0] }]
              : [];
          }
        }
        return mapping;
      });
      setMappings(autoMappings);
    }
  };

  const updateMapping = (index: number, updates: Partial<ColumnMapping>) => {
    setMappings(prev => prev.map((m, i) => {
      if (i !== index) return m;
      const updated = { ...m, ...updates };
      // When target column changes, auto-configure
      if (updates.targetColumn !== undefined && schema) {
        const col = schema.columns.find(c => c.name === updates.targetColumn);
        if (col && (col.type === 'date' || col.type === 'datetime')) {
          updated.dateFormat = updated.dateFormat || 'auto';
        } else {
          delete updated.dateFormat;
        }
        if (col && col.type === 'reference' && col.refTable) {
          const refSchema = state.getSchema(col.refTable);
          if (refSchema && !updated.refSourceMappings?.length) {
            const matchCols = col.refDisplayColumns?.length
              ? col.refDisplayColumns
              : col.refSearchColumns?.length
                ? col.refSearchColumns
                : [refSchema.columns[0]?.name].filter(Boolean);
            updated.refSourceMappings = matchCols.length > 0
              ? [{ sourceColumn: updated.sourceColumn, refColumn: matchCols[0] }]
              : [];
          }
        } else {
          delete updated.refSourceMappings;
        }
      }
      updated.enabled = !!updated.targetColumn;
      return updated;
    }));
  };

  const updateNewTableColumn = (index: number, updates: Partial<NewTableColumn>) => {
    setNewTableColumns(prev => prev.map((c, i) => {
      if (i !== index) return c;
      const updated = { ...c, ...updates };
      if (updates.type !== undefined) {
        if (updates.type === 'date' || updates.type === 'datetime') {
          updated.dateFormat = updated.dateFormat || 'auto';
        } else {
          delete updated.dateFormat;
        }
        if (updates.type === 'reference') {
          // default to first other table
          if (!updated.refTable) {
            const otherTables = state.tableIds.filter(t => t !== tableId);
            updated.refTable = otherTables[0] ?? '';
          }
          // Initialize refSourceMappings when switching to reference type
          if (updated.refTable && !updated.refSourceMappings?.length) {
            const refSchema = state.getSchema(updated.refTable);
            const firstRefCol = refSchema?.columns[0]?.name;
            updated.refSourceMappings = firstRefCol
              ? [{ sourceColumn: updated.sourceColumn, refColumn: firstRefCol }]
              : [];
          }
        } else {
          delete updated.refTable;
          delete updated.refSourceMappings;
        }
      }
      // Reset refSourceMappings when ref table changes
      if (updates.refTable !== undefined && updated.type === 'reference') {
        const refSchema = state.getSchema(updates.refTable);
        const firstRefCol = refSchema?.columns[0]?.name;
        updated.refSourceMappings = firstRefCol
          ? [{ sourceColumn: updated.sourceColumn, refColumn: firstRefCol }]
          : [];
      }
      return updated;
    }));
  };

  // Resolve a reference to a _rowId by matching multiple source→ref column pairs
  const resolveReference = useCallback(
    (refTableId: string, pairs: { sourceValue: string; refColumn: string }[]): string | null => {
      const refTable = state.model.getTable(refTableId);
      if (!refTable) return null;
      if (pairs.length === 0) return null;

      // All source values empty → empty reference
      if (pairs.every(p => !p.sourceValue)) return '';

      // Helper: get display value for a column, following references if the column is itself a reference
      const getDisplayValue = (colName: string, rawValue: string): string => {
        if (!rawValue) return '';
        const colDef = refTable.schema.columns.find(c => c.name === colName);
        if (colDef?.type === 'reference' && colDef.refTable) {
          const nestedTable = state.model.getTable(colDef.refTable);
          if (nestedTable) {
            const nestedRow = nestedTable.rows.find(r => r[INTERNAL_ROW_ID] === rawValue);
            if (nestedRow) {
              const displayCols = colDef.refDisplayColumns ?? [nestedTable.schema.columns[0]?.name].filter(Boolean);
              return displayCols.map(c => nestedRow[c] ?? '').filter(Boolean).join(' · ');
            }
          }
          return '';
        }
        return rawValue;
      };

      // Find a ref row where all pairs match simultaneously
      for (const row of refTable.rows) {
        const allMatch = pairs.every(p => {
          const displayValue = getDisplayValue(p.refColumn, row[p.refColumn] ?? '');
          return displayValue.toLowerCase() === p.sourceValue.toLowerCase();
        });
        if (allMatch) {
          return row[INTERNAL_ROW_ID];
        }
      }
      return null;
    },
    [state.model],
  );

  // Parse a date value using selected format
  const parseDate = useCallback(
    (value: string, formatId: string): string | null => {
      const fmt = DATE_FORMATS.find(f => f.value === formatId);
      if (!fmt) return value;
      return fmt.parse(value.trim());
    },
    [],
  );

  // Execute import
  const handleImport = useCallback(async () => {
    if (isNewTable) {
      // New table mode: create table, then insert rows
      if (!newTableName.trim()) {
        showAlert('Please enter a table name');
        return;
      }
      const enabledCols = newTableColumns.filter(c => c.enabled);
      if (enabledCols.length === 0) {
        showAlert('Enable at least one column');
        return;
      }
      // Check for duplicate column names
      const names = enabledCols.map(c => c.name.trim().toLowerCase());
      if (names.some((n, i) => names.indexOf(n) !== i)) {
        showAlert('Column names must be unique');
        return;
      }
      // Check for empty column names
      if (enabledCols.some(c => !c.name.trim())) {
        showAlert('All column names must be non-empty');
        return;
      }
      // Check reference columns have refTable set
      if (enabledCols.some(c => c.type === 'reference' && !c.refTable)) {
        showAlert('All reference columns must have a target table selected');
        return;
      }

      setImporting(true);
      setImportResult(null);

      // Build schema
      const columns: ColumnDef[] = enabledCols.map(c => {
        const col: ColumnDef = { name: c.name.trim(), type: c.type };
        if (c.type === 'reference' && c.refTable) {
          col.refTable = c.refTable;
          const refSchema = state.getSchema(c.refTable);
          if (refSchema) {
            col.refDisplayColumns = [refSchema.columns[0]?.name].filter(Boolean);
            col.refSearchColumns = [refSchema.columns[0]?.name].filter(Boolean);
          }
        }
        return col;
      });
      const newSchema: TableSchema = {
        name: newTableName.trim(),
        columns,
        uniqueKeys: [],
      };

      state.createTable(newSchema);
      const newId = newSchema.name;
      setCreatedTableId(newId);

      const errors: string[] = [];
      let imported = 0;

      for (let i = 0; i < sourceRows.length; i++) {
        const sourceRow = sourceRows[i];
        const newRow: Row = {};

        for (const col of columns) {
          newRow[col.name] = '';
        }

        let skipRow = false;
        for (const colDef of enabledCols) {
          const schemaCol = columns.find(c => c.name === colDef.name.trim());
          if (!schemaCol) continue;

          // Reference resolution (multi-column matching)
          if (schemaCol.type === 'reference' && schemaCol.refTable) {
            const srcMappings = colDef.refSourceMappings ?? [];
            const pairs = srcMappings.map(sm => ({
              sourceValue: sourceRow[sourceHeaders.indexOf(sm.sourceColumn)] ?? '',
              refColumn: sm.refColumn,
            }));
            if (pairs.length > 0 && pairs.some(p => p.sourceValue)) {
              const resolved = resolveReference(schemaCol.refTable, pairs);
              if (resolved === null) {
                const desc = pairs.map(p => `${p.refColumn}="${p.sourceValue}"`).join(', ');
                errors.push(`Row ${i + 1}: Could not find reference (${desc}) for column "${schemaCol.name}"`);
                skipRow = true;
                break;
              }
              newRow[colDef.name.trim()] = resolved;
            }
            continue;
          }

          const sourceIdx = sourceHeaders.indexOf(colDef.sourceColumn);
          if (sourceIdx === -1) continue;
          let value = sourceRow[sourceIdx] ?? '';

          // Strip thousands separators from numbers
          if ((schemaCol.type === 'integer' || schemaCol.type === 'decimal') && value) {
            value = value.replace(/,/g, '');
          }

          // Date parsing
          if ((schemaCol.type === 'date' || schemaCol.type === 'datetime') && colDef.dateFormat && value) {
            const parsed = parseDate(value, colDef.dateFormat);
            if (parsed === null) {
              errors.push(`Row ${i + 1}: Could not parse date "${value}" for column "${schemaCol.name}"`);
              skipRow = true;
              break;
            }
            value = parsed;
          }

          newRow[colDef.name.trim()] = value;
        }

        if (skipRow) continue;

        const insertErrors = state.insertRow(newId, newRow);
        if (insertErrors.length > 0) {
          errors.push(`Row ${i + 1}: ${insertErrors[0].message}`);
        } else {
          imported++;
        }
      }

      setImportResult({ imported, errors });
      setImporting(false);
      if (imported > 0) {
        navigate(`/table/${encodeURIComponent(newId)}`);
      }
      return;
    }

    // Existing table mode
    if (!tableId || !schema) return;
    setImporting(true);
    setImportResult(null);

    const enabledMappings = mappings.filter(m => m.enabled && m.targetColumn);
    const errors: string[] = [];
    let imported = 0;

    for (let i = 0; i < sourceRows.length; i++) {
      const sourceRow = sourceRows[i];
      const newRow: Row = {};

      // Fill all columns with empty
      for (const col of schema.columns) {
        newRow[col.name] = '';
      }

      let skipRow = false;
      for (const mapping of enabledMappings) {
        const col = schema.columns.find(c => c.name === mapping.targetColumn);
        if (!col) continue;

        // Reference resolution (multi-column matching)
        if (col.type === 'reference' && col.refTable) {
          const srcMappings = mapping.refSourceMappings ?? [];
          const pairs = srcMappings.map(sm => ({
            sourceValue: sourceRow[sourceHeaders.indexOf(sm.sourceColumn)] ?? '',
            refColumn: sm.refColumn,
          }));
          if (pairs.length > 0 && pairs.some(p => p.sourceValue)) {
            const resolved = resolveReference(col.refTable, pairs);
            if (resolved === null) {
              const desc = pairs.map(p => `${p.refColumn}="${p.sourceValue}"`).join(', ');
              errors.push(`Row ${i + 1}: Could not find reference (${desc}) for column "${col.name}"`);
              skipRow = true;
              break;
            }
            newRow[mapping.targetColumn] = resolved;
          }
          continue;
        }

        const sourceIdx = sourceHeaders.indexOf(mapping.sourceColumn);
        if (sourceIdx === -1) continue;
        let value = sourceRow[sourceIdx] ?? '';

        // Strip thousands separators from numbers
        if ((col.type === 'integer' || col.type === 'decimal') && value) {
          value = value.replace(/,/g, '');
        }

        // Date parsing
        if ((col.type === 'date' || col.type === 'datetime') && mapping.dateFormat && value) {
          const parsed = parseDate(value, mapping.dateFormat);
          if (parsed === null) {
            errors.push(`Row ${i + 1}: Could not parse date "${value}" for column "${col.name}"`);
            skipRow = true;
            break;
          }
          value = parsed;
        }

        newRow[mapping.targetColumn] = value;
      }

      if (skipRow) continue;

      const insertErrors = state.insertRow(tableId, newRow);
      if (insertErrors.length > 0) {
        errors.push(`Row ${i + 1}: ${insertErrors[0].message}`);
      } else {
        imported++;
      }
    }

    setImportResult({ imported, errors });
    setImporting(false);
    if (imported > 0) {
      navigate(`/table/${encodeURIComponent(tableId)}`);
    }
  }, [tableId, schema, mappings, sourceRows, sourceHeaders, state, parseDate, resolveReference, isNewTable, newTableName, newTableColumns, showAlert, navigate]);

  // For existing table mode, require valid table  
  if (!isNewTable && !schema) {
    return (
      <div className="import-page">
        <div className="import-card">
          <h2>Table not found</h2>
          <button className="btn-secondary" onClick={() => navigate('/')}>Back</button>
        </div>
      </div>
    );
  }

  const getTargetCol = (targetName: string) => schema?.columns.find(c => c.name === targetName);

  // Table options for reference column target
  const tableOptions = state.tableIds.map(tid => {
    const s = state.getSchema(tid);
    return { value: tid, label: s?.name ?? tid };
  });

  // Determine if import button should be enabled
  const canImport = isNewTable
    ? newTableColumns.some(c => c.enabled) && !!newTableName.trim()
    : mappings.some(m => m.enabled && m.targetColumn);

  // Determine which columns to show in preview
  const previewColumns = isNewTable
    ? newTableColumns.filter(c => c.enabled)
    : mappings.filter(m => m.enabled && m.targetColumn);

  const goToTableId = isNewTable ? createdTableId : tableId;

  return (
    <div className="import-page">
      <div className="import-card">
        <div className="import-header">
          <h2>{isNewTable ? 'Import into New Table' : `Import into ${schema!.name}`}</h2>
          <button className="btn-secondary" onClick={() => navigate(tableId ? `/table/${encodeURIComponent(tableId)}` : '/')}>
            Cancel
          </button>
        </div>

        {/* Table name (new table mode only) */}
        {isNewTable && (
          <div className="import-section">
            <h3>Table Name</h3>
            <input
              className="import-table-name-input"
              type="text"
              placeholder="Enter table name..."
              value={newTableName}
              onChange={(e) => setNewTableName(e.target.value)}
            />
          </div>
        )}

        {/* Step 1: Source selection */}
        <div className="import-section">
          <h3>{isNewTable ? '1. Select Source' : '1. Select Source'}</h3>
          <div className="import-source-buttons">
            <label className="import-file-label">
              <input type="file" accept=".csv,.txt" onChange={handleFileUpload} hidden />
              <span className="btn-primary">Upload CSV</span>
            </label>
            {state.isSignedIn && (
              <button className="btn-secondary" onClick={handleGoogleSheetImport}>
                Import from Google Sheet
              </button>
            )}
          </div>
          {sourceFileName && (
            <div className="import-source-info">
              Source: <strong>{sourceFileName}</strong> — {sourceRows.length} rows, {sourceHeaders.length} columns
            </div>
          )}
        </div>

        {/* Google Sheet URL dialog */}
        {showSheetDialog && (
          <div className="app-dialog-overlay" onClick={() => { if (!sheetLoading) { setShowSheetDialog(false); setSheetTabs([]); } }}>
            <div className="app-dialog" onClick={(e) => e.stopPropagation()}>
              <h3 className="app-dialog-title">Import from Google Sheet</h3>
              {sheetTabs.length > 1 ? (
                <>
                  <p className="app-dialog-message">Select a sheet to import:</p>
                  <div className="import-sheet-tabs">
                    {sheetTabs.map(tab => (
                      <label key={tab.sheetId} className={`import-sheet-tab-option${selectedSheetGid === tab.sheetId ? ' selected' : ''}`}>
                        <input
                          type="radio"
                          name="sheetTab"
                          checked={selectedSheetGid === tab.sheetId}
                          onChange={() => setSelectedSheetGid(tab.sheetId)}
                        />
                        {tab.title}
                      </label>
                    ))}
                  </div>
                  <div className="app-dialog-actions">
                    <button
                      className="app-dialog-btn app-dialog-btn-secondary"
                      onClick={() => { setShowSheetDialog(false); setSheetTabs([]); }}
                      disabled={sheetLoading}
                    >
                      Cancel
                    </button>
                    <button
                      className="app-dialog-btn app-dialog-btn-primary"
                      onClick={handleSheetTabImport}
                      disabled={sheetLoading || selectedSheetGid === undefined}
                    >
                      {sheetLoading ? 'Loading...' : 'Import'}
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <p className="app-dialog-message">Paste the Google Sheets share URL below.</p>
                  <input
                    className="import-sheet-url-input"
                    type="url"
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                    value={sheetUrl}
                    onChange={(e) => setSheetUrl(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') handleSheetUrlSubmit(); }}
                    autoFocus
                    disabled={sheetLoading}
                  />
                  <div className="app-dialog-actions">
                    <button
                      className="app-dialog-btn app-dialog-btn-secondary"
                      onClick={() => setShowSheetDialog(false)}
                      disabled={sheetLoading}
                    >
                      Cancel
                    </button>
                    <button
                      className="app-dialog-btn app-dialog-btn-primary"
                      onClick={handleSheetUrlSubmit}
                      disabled={sheetLoading || !sheetUrl.trim()}
                    >
                      {sheetLoading ? 'Loading...' : 'Import'}
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* Step 2: Column configuration */}
        {sourceHeaders.length > 0 && (
          <div className="import-section">
            <h3>{isNewTable ? '2. Configure Columns' : '2. Map Columns'}</h3>
            <div className="import-mapping-table">
              {isNewTable ? (
                <>
                  <div className="import-mapping-header">
                    <span className="import-mapping-cell">Import</span>
                    <span className="import-mapping-cell">Source Column</span>
                    <span className="import-mapping-cell">→</span>
                    <span className="import-mapping-cell import-mapping-cell-wide">Column Name</span>
                    <span className="import-mapping-cell">Type</span>
                    <span className="import-mapping-cell import-mapping-cell-wide">Options</span>
                  </div>
                  {newTableColumns.map((col, i) => (
                    <div key={i} className="import-mapping-row">
                      <span className="import-mapping-cell">
                        <input
                          type="checkbox"
                          checked={col.enabled}
                          onChange={(e) => updateNewTableColumn(i, { enabled: e.target.checked })}
                        />
                      </span>
                      <span className="import-mapping-cell import-mapping-source">
                        {col.sourceColumn}
                        <span className="import-sample-value">
                          e.g. "{sourceRows[0]?.[sourceHeaders.indexOf(col.sourceColumn)] ?? ''}"
                        </span>
                      </span>
                      <span className="import-mapping-cell">→</span>
                      <span className="import-mapping-cell import-mapping-cell-wide">
                        <input
                          className="import-col-name-input"
                          type="text"
                          value={col.name}
                          onChange={(e) => updateNewTableColumn(i, { name: e.target.value })}
                        />
                      </span>
                      <span className="import-mapping-cell">
                        <Select
                          options={typeOptions}
                          value={typeOptions.find(o => o.value === col.type) ?? null}
                          onChange={(opt) => updateNewTableColumn(i, { type: opt?.value ?? 'text' })}
                          styles={selectStyles}
                          isClearable={false}
                          menuPortalTarget={document.body}
                        />
                      </span>
                      <span className="import-mapping-cell import-mapping-cell-wide">
                        {(col.type === 'date' || col.type === 'datetime') && (
                          <Select
                            options={DATE_FORMATS.map(f => ({ value: f.value, label: f.label }))}
                            value={DATE_FORMATS.map(f => ({ value: f.value, label: f.label })).find(
                              o => o.value === (col.dateFormat ?? 'auto')
                            )}
                            onChange={(opt) => updateNewTableColumn(i, { dateFormat: opt?.value ?? 'auto' })}
                            styles={selectStyles}
                            isClearable={false}
                            menuPortalTarget={document.body}
                          />
                        )}
                        {col.type === 'reference' && (
                          <div className="import-ref-config">
                            <span className="import-ref-label">Table:</span>
                            <Select
                              options={tableOptions}
                              value={tableOptions.find(o => o.value === col.refTable) ?? null}
                              onChange={(opt) => updateNewTableColumn(i, { refTable: opt?.value ?? '' })}
                              styles={selectStyles}
                              isClearable={false}
                              menuPortalTarget={document.body}
                            />
                            {col.refTable && (
                              <div className="import-ref-pairs">
                                <span className="import-ref-label">Match columns:</span>
                                {(col.refSourceMappings ?? []).map((pair, pi) => {
                                  const refSchema = state.getSchema(col.refTable!);
                                  const refColOptions = (refSchema?.columns ?? []).map(rc => ({
                                    value: rc.name,
                                    label: rc.displayName || rc.name,
                                  }));
                                  return (
                                    <div key={pi} className="import-ref-pair">
                                      <Select
                                        options={sourceColumnOptions}
                                        value={sourceColumnOptions.find(o => o.value === pair.sourceColumn) ?? null}
                                        onChange={(opt) => {
                                          const newMappings = [...(col.refSourceMappings ?? [])];
                                          newMappings[pi] = { ...newMappings[pi], sourceColumn: opt?.value ?? '' };
                                          updateNewTableColumn(i, { refSourceMappings: newMappings });
                                        }}
                                        styles={selectStyles}
                                        placeholder="Source column"
                                        menuPortalTarget={document.body}
                                      />
                                      <span className="import-ref-pair-arrow">→</span>
                                      <Select
                                        options={refColOptions}
                                        value={refColOptions.find(o => o.value === pair.refColumn) ?? null}
                                        onChange={(opt) => {
                                          const newMappings = [...(col.refSourceMappings ?? [])];
                                          newMappings[pi] = { ...newMappings[pi], refColumn: opt?.value ?? '' };
                                          updateNewTableColumn(i, { refSourceMappings: newMappings });
                                        }}
                                        styles={selectStyles}
                                        placeholder="Ref column"
                                        menuPortalTarget={document.body}
                                      />
                                      <button
                                        className="import-ref-pair-remove"
                                        onClick={() => {
                                          const newMappings = (col.refSourceMappings ?? []).filter((_, j) => j !== pi);
                                          updateNewTableColumn(i, { refSourceMappings: newMappings });
                                        }}
                                        title="Remove"
                                      >×</button>
                                    </div>
                                  );
                                })}
                                <button
                                  className="import-ref-pair-add"
                                  onClick={() => {
                                    const refSchema = state.getSchema(col.refTable!);
                                    const usedRefCols = (col.refSourceMappings ?? []).map(m => m.refColumn);
                                    const usedSourceCols = (col.refSourceMappings ?? []).map(m => m.sourceColumn);
                                    const nextRefCol = refSchema?.columns.find(rc => !usedRefCols.includes(rc.name))?.name ?? '';
                                    const nextSourceCol = sourceHeaders.find(h => !usedSourceCols.includes(h)) ?? '';
                                    updateNewTableColumn(i, {
                                      refSourceMappings: [
                                        ...(col.refSourceMappings ?? []),
                                        { sourceColumn: nextSourceCol, refColumn: nextRefCol },
                                      ],
                                    });
                                  }}
                                >+ Add column match</button>
                              </div>
                            )}
                          </div>
                        )}
                      </span>
                    </div>
                  ))}
                </>
              ) : (
                <>
                  <div className="import-mapping-header">
                    <span className="import-mapping-cell">Import</span>
                    <span className="import-mapping-cell">Source Column</span>
                    <span className="import-mapping-cell">→</span>
                    <span className="import-mapping-cell import-mapping-cell-wide">Target Column</span>
                    <span className="import-mapping-cell import-mapping-cell-wide">Options</span>
                  </div>
                  {mappings.map((mapping, i) => {
                    const targetCol = mapping.targetColumn ? getTargetCol(mapping.targetColumn) : null;
                    const isDate = targetCol && (targetCol.type === 'date' || targetCol.type === 'datetime');
                    const isRef = targetCol && targetCol.type === 'reference' && targetCol.refTable;

                    return (
                      <div key={i} className="import-mapping-row">
                        <span className="import-mapping-cell">
                          <input
                            type="checkbox"
                            checked={mapping.enabled}
                            onChange={(e) => updateMapping(i, { enabled: e.target.checked })}
                          />
                        </span>
                        <span className="import-mapping-cell import-mapping-source">
                          {mapping.sourceColumn}
                          <span className="import-sample-value">
                            e.g. "{sourceRows[0]?.[sourceHeaders.indexOf(mapping.sourceColumn)] ?? ''}"
                          </span>
                        </span>
                        <span className="import-mapping-cell">→</span>
                        <span className="import-mapping-cell import-mapping-cell-wide">
                          <Select
                            options={targetColumnOptions}
                            value={targetColumnOptions.find(o => o.value === mapping.targetColumn) ?? null}
                            onChange={(opt) => updateMapping(i, { targetColumn: opt?.value ?? '' })}
                            styles={selectStyles}
                            isClearable={false}
                            placeholder="Skip"
                            menuPortalTarget={document.body}
                          />
                        </span>
                        <span className="import-mapping-cell import-mapping-cell-wide">
                          {isDate && (
                            <Select
                              options={DATE_FORMATS.map(f => ({ value: f.value, label: f.label }))}
                              value={DATE_FORMATS.map(f => ({ value: f.value, label: f.label })).find(
                                o => o.value === (mapping.dateFormat ?? 'auto')
                              )}
                              onChange={(opt) => updateMapping(i, { dateFormat: opt?.value ?? 'auto' })}
                              styles={selectStyles}
                              isClearable={false}
                              menuPortalTarget={document.body}
                            />
                          )}
                          {isRef && targetCol.refTable && (
                            <div className="import-ref-config">
                              <div className="import-ref-pairs">
                                <span className="import-ref-label">Match columns:</span>
                                {(mapping.refSourceMappings ?? []).map((pair, pi) => {
                                  const refSchema = state.getSchema(targetCol.refTable!);
                                  const refColOptions = (refSchema?.columns ?? []).map(rc => ({
                                    value: rc.name,
                                    label: rc.displayName || rc.name,
                                  }));
                                  return (
                                    <div key={pi} className="import-ref-pair">
                                      <Select
                                        options={sourceColumnOptions}
                                        value={sourceColumnOptions.find(o => o.value === pair.sourceColumn) ?? null}
                                        onChange={(opt) => {
                                          const newMappings = [...(mapping.refSourceMappings ?? [])];
                                          newMappings[pi] = { ...newMappings[pi], sourceColumn: opt?.value ?? '' };
                                          updateMapping(i, { refSourceMappings: newMappings });
                                        }}
                                        styles={selectStyles}
                                        placeholder="Source column"
                                        menuPortalTarget={document.body}
                                      />
                                      <span className="import-ref-pair-arrow">→</span>
                                      <Select
                                        options={refColOptions}
                                        value={refColOptions.find(o => o.value === pair.refColumn) ?? null}
                                        onChange={(opt) => {
                                          const newMappings = [...(mapping.refSourceMappings ?? [])];
                                          newMappings[pi] = { ...newMappings[pi], refColumn: opt?.value ?? '' };
                                          updateMapping(i, { refSourceMappings: newMappings });
                                        }}
                                        styles={selectStyles}
                                        placeholder="Ref column"
                                        menuPortalTarget={document.body}
                                      />
                                      <button
                                        className="import-ref-pair-remove"
                                        onClick={() => {
                                          const newMappings = (mapping.refSourceMappings ?? []).filter((_, j) => j !== pi);
                                          updateMapping(i, { refSourceMappings: newMappings });
                                        }}
                                        title="Remove"
                                      >×</button>
                                    </div>
                                  );
                                })}
                                <button
                                  className="import-ref-pair-add"
                                  onClick={() => {
                                    const refSchema = state.getSchema(targetCol.refTable!);
                                    const usedRefCols = (mapping.refSourceMappings ?? []).map(m => m.refColumn);
                                    const usedSourceCols = (mapping.refSourceMappings ?? []).map(m => m.sourceColumn);
                                    const nextRefCol = refSchema?.columns.find(rc => !usedRefCols.includes(rc.name))?.name ?? '';
                                    const nextSourceCol = sourceHeaders.find(h => !usedSourceCols.includes(h)) ?? '';
                                    updateMapping(i, {
                                      refSourceMappings: [
                                        ...(mapping.refSourceMappings ?? []),
                                        { sourceColumn: nextSourceCol, refColumn: nextRefCol },
                                      ],
                                    });
                                  }}
                                >+ Add column match</button>
                              </div>
                            </div>
                          )}
                        </span>
                      </div>
                    );
                  })}
                </>
              )}
            </div>
          </div>
        )}

        {/* Step 3: Preview */}
        {sourceHeaders.length > 0 && previewColumns.length > 0 && (
          <div className="import-section">
            <h3>3. Preview</h3>
            <div className="import-preview">
              <table>
                <thead>
                  <tr>
                    {previewColumns.map((c, i) => (
                      <th key={i}>{'name' in c ? c.name : c.targetColumn}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sourceRows.slice(0, 5).map((row, ri) => (
                    <tr key={ri}>
                      {previewColumns.map((c, ci) => {
                        if (c.refSourceMappings && c.refSourceMappings.length > 0) {
                          const values = c.refSourceMappings
                            .map(m => row[sourceHeaders.indexOf(m.sourceColumn)] ?? '')
                            .filter(Boolean);
                          return <td key={ci}>{values.join(' · ')}</td>;
                        }
                        const sourceCol = 'sourceColumn' in c ? c.sourceColumn : '';
                        const sourceIdx = sourceHeaders.indexOf(sourceCol);
                        return <td key={ci}>{row[sourceIdx] ?? ''}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              {sourceRows.length > 5 && (
                <div className="import-preview-more">
                  ...and {sourceRows.length - 5} more rows
                </div>
              )}
            </div>
          </div>
        )}

        {/* Step 4: Import */}
        {sourceHeaders.length > 0 && (
          <div className="import-section">
            <button
              className="btn-primary import-btn"
              onClick={handleImport}
              disabled={importing || !canImport}
            >
              {importing ? 'Importing...' : `Import ${sourceRows.length} rows`}
            </button>
          </div>
        )}

        {/* Results */}
        {importResult && (
          <div className="import-section">
            <div className={`import-result ${importResult.errors.length > 0 ? 'import-result-warnings' : 'import-result-success'}`}>
              <strong>{importResult.imported} rows imported successfully.</strong>
              {importResult.errors.length > 0 && (
                <div className="import-errors">
                  <p>{importResult.errors.length} rows had errors:</p>
                  <ul>
                    {importResult.errors.slice(0, 20).map((err, i) => (
                      <li key={i}>{err}</li>
                    ))}
                    {importResult.errors.length > 20 && (
                      <li>...and {importResult.errors.length - 20} more</li>
                    )}
                  </ul>
                </div>
              )}
            </div>
            {goToTableId && (
              <button
                className="btn-primary"
                onClick={() => navigate(`/table/${encodeURIComponent(goToTableId)}`)}
              >
                Go to Table
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

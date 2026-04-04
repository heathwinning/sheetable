import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import Select from 'react-select';
import { DATE_FORMATS } from './dateFormatsList';
import type { ColumnDef, ColumnType, Row, TableSchema } from './types';
import { INTERNAL_ROW_ID } from './types';
import type { UseAppStateReturn } from './useAppState';
import { useDialog } from './DialogProvider';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, ValueSetterParams, CellClickedEvent } from 'ag-grid-community';
import type { CustomCellEditorProps } from 'ag-grid-react';
import { previewMigration, applyMigration, previewExtract } from './typeMigration';
import type { MigrationPreview } from './typeMigration';
import * as drive from './drive';

const typeOptions: { value: ColumnType; label: string }[] = [
  { value: 'text', label: 'Text' },
  { value: 'integer', label: 'Integer' },
  { value: 'decimal', label: 'Decimal' },
  { value: 'date', label: 'Date' },
  { value: 'datetime', label: 'Datetime' },
  { value: 'bool', label: 'Boolean' },
  { value: 'reference', label: 'Reference' },
  { value: 'image', label: 'Image' },
];

function TypeCellEditor({ value, onValueChange, stopEditing }: CustomCellEditorProps) {
  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    listRef.current?.focus();
  }, []);

  return (
    <div
      ref={listRef}
      tabIndex={0}
      className="ag-custom-component-popup"
      onBlur={() => stopEditing()}
      onKeyDown={(e) => {
        if (e.key === 'Escape') stopEditing();
      }}
      style={{
        background: '#fff',
        border: '1px solid var(--border)',
        borderRadius: 4,
        boxShadow: '0 4px 16px rgba(0,0,0,0.10)',
        overflow: 'auto',
        maxHeight: 240,
        minWidth: 120,
        outline: 'none',
      }}
    >
      {typeOptions.map(o => (
        <div
          key={o.value}
          onMouseDown={(e) => {
            e.preventDefault();
            onValueChange(o.value);
            stopEditing();
          }}
          style={{
            padding: '4px 10px',
            fontSize: 13,
            cursor: 'pointer',
            background: o.value === value ? 'var(--primary, #2563eb)' : 'transparent',
            color: o.value === value ? '#fff' : 'var(--text)',
          }}
          onMouseEnter={(e) => {
            if (o.value !== value) e.currentTarget.style.background = 'var(--cell-selected, #e0e7ff)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.background = o.value === value ? 'var(--primary, #2563eb)' : 'transparent';
          }}
        >
          {o.label}
        </div>
      ))}
    </div>
  );
}

const selectStyles = {
  control: (base: Record<string, unknown>, state: { isFocused: boolean }) => ({
    ...base,
    background: 'var(--bg)',
    borderColor: state.isFocused ? 'var(--ref-color)' : 'var(--border)',
    boxShadow: state.isFocused ? '0 0 0 2px rgba(37, 99, 235, 0.2)' : 'none',
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
  option: (base: Record<string, unknown>, state: { isSelected: boolean; isFocused: boolean }) => ({
    ...base,
    background: state.isSelected ? 'var(--primary)' : state.isFocused ? 'var(--cell-selected)' : 'transparent',
    color: state.isSelected ? '#fff' : 'var(--text)',
    fontSize: 13,
    padding: '4px 10px',
    cursor: 'pointer',
  }),
  singleValue: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  multiValue: (base: Record<string, unknown>) => ({ ...base, background: 'var(--cell-selected, #e0e7ff)' }),
  input: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  placeholder: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text-muted)' }),
  indicatorSeparator: () => ({ display: 'none' }),
};

interface EditTablePageProps {
  state: UseAppStateReturn;
}

export const EditTablePage: React.FC<EditTablePageProps> = ({ state }) => {
  const { tableId, bookId } = useParams<{ tableId: string; bookId?: string }>();
  const navigate = useNavigate();
  const { showDialog } = useDialog();
  const isCreate = !tableId;
  const schema = tableId ? state.getSchema(tableId) : undefined;
  const bookBase = bookId ? `/book/${encodeURIComponent(bookId)}` : '';
  const toBookPath = (suffix: string) => `${bookBase}${suffix}`;

  const [tableName, setTableName] = useState(schema?.name ?? '');
  const [csvFileName, setCsvFileName] = useState('');
  const [csvFileId, setCsvFileId] = useState(schema?.csvFileId ?? '');
  const [availableCsvFiles, setAvailableCsvFiles] = useState<Array<{ id: string; name: string }>>([]);
  const [pickingCsv, setPickingCsv] = useState(false);
  const [columns, setColumns] = useState<ColumnDef[]>(
    () => schema?.columns.map(c => ({ ...c })) ?? [
      { name: '', type: 'text' as ColumnType },
    ]
  );
  const [uniqueKeys, setUniqueKeys] = useState<string[]>(
    () => schema?.uniqueKeys ?? []
  );
  const [defaultSort, setDefaultSort] = useState<{ column: string; direction: 'asc' | 'desc' }[]>(
    () => schema?.defaultSort ?? []
  );
  const [draftRowPosition, setDraftRowPosition] = useState<'top' | 'bottom'>(
    () => schema?.draftRowPosition ?? 'bottom'
  );
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<{ kind: 'success' | 'info'; message: string } | null>(null);

  // Sync local state when schema loads asynchronously (e.g. Drive reload)
  const [schemaLoaded, setSchemaLoaded] = useState(!!schema);
  useEffect(() => {
    if (schema && !schemaLoaded) {
      setTableName(schema.name ?? '');
      setCsvFileId(schema.csvFileId ?? '');
      setCsvFileName(schema.csvFileName ?? '');
      setColumns(schema.columns.map(c => ({ ...c })));
      setUniqueKeys(schema.uniqueKeys ?? []);
      setDefaultSort(schema.defaultSort ?? []);
      setDraftRowPosition(schema.draftRowPosition ?? 'bottom');
      setSchemaLoaded(true);
    }
  }, [schema, schemaLoaded]);

  useEffect(() => {
    let cancelled = false;
    state.listWorkbookCsvFiles()
      .then(files => {
        if (cancelled) return;
        setAvailableCsvFiles(files);
        if (csvFileId) {
          const matched = files.find(f => f.id === csvFileId);
          if (matched) {
            setCsvFileName(matched.name);
          }
        }
      })
      .catch(() => {
        if (!cancelled) setAvailableCsvFiles([]);
      });
    return () => {
      cancelled = true;
    };
  }, [csvFileId, state.folderId, state.listWorkbookCsvFiles]);

  // Type migration preview (for existing tables only)
  const [migrationPreview, setMigrationPreview] = useState<{
    preview: MigrationPreview;
    colIndex: number;
    newType: ColumnType;
    refUpdates?: Partial<ColumnDef>;
    dateFormat?: string;
  } | null>(null);

  // Unified migrations/normalization dialog state
  const [migrationsDialogOpen, setMigrationsDialogOpen] = useState(false);
  const [migrationTargetColIdx, setMigrationTargetColIdx] = useState<number | null>(null);

  // Extract-to-table preview dialog open
  const [extractPreview, setExtractPreview] = useState(false);

  const otherTableIds = useMemo(
    () => state.tableIds.filter(id => id !== tableId),
    [state.tableIds, tableId]
  );

  // AG Grid theme matching SpreadsheetGrid
  const editGridTheme = useMemo(() => themeQuartz.withParams({
    backgroundColor: '#ffffff',
    foregroundColor: '#1e1e2e',
    headerBackgroundColor: '#f5f6f8',
    rowHoverColor: '#f0f4ff',
    selectedRowBackgroundColor: '#e0e7ff',
    borderColor: '#d4d4d8',
    cellHorizontalPaddingScale: 0.8,
    headerFontSize: 12,
    fontSize: 13,
    rowHeight: 32,
    headerHeight: 32,
    columnBorder: true,
  }), []);

  // Row data for AG Grid column editor: each row = one column definition
  // Add _idx field so we can map back to state array
  const columnRowData = useMemo(() =>
    columns.map((col, i) => ({
      _idx: i,
      name: col.name,
      displayName: col.displayName ?? '',
      type: col.type,
      isKey: uniqueKeys.includes(col.name.trim()),
      sortDir: (() => {
        const entry = defaultSort.find(s => s.column === col.name.trim());
        return entry ? entry.direction : '';
      })(),
      isRef: col.type === 'reference',
      refTable: col.refTable ?? '',
      refDisplayColumns: col.refDisplayColumns ?? [],
      refSearchColumns: col.refSearchColumns ?? [],
    })),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [columns, uniqueKeys, defaultSort]
  );

  const getRefTableColumns = useCallback((refTableId: string): string[] => {
    const refSchema = state.getSchema(refTableId);
    return refSchema?.columns.map(c => c.name) ?? [];
  }, [state]);

  const getRefTableRows = useCallback((refTableId: string): Row[] => {
    return state.getRows(refTableId);
  }, [state]);

  const resolveRefPathValue = useCallback((tableId: string, row: Row, path: string): string => {
    return state.model.resolveColumnPath(tableId, row, path);
  }, [state.model]);

  const getRefTableColumnPaths = useCallback((refTableId: string): { path: string; label: string }[] => {
    return state.model.getColumnPaths(refTableId);
  }, [state.model]);

  // Reference config dialog
  const openRefConfigDialog = useCallback(async (colIndex: number) => {
    const col = columns[colIndex];
    if (col.type !== 'reference') return;

    // We'll use a custom modal state for the ref config
    setRefDialogCol(colIndex);
  }, [columns]);

  const [refDialogCol, setRefDialogCol] = useState<number | null>(null);

  const columnGridDefs: ColDef[] = useMemo(() => [
    {
      headerName: 'Name',
      field: 'name',
      editable: true,
      flex: 2,
      minWidth: 100,
      valueSetter: (params: ValueSetterParams) => {
        const idx = params.data._idx;
        const oldName = columns[idx].name;
        const newName = params.newValue ?? '';
        updateColumn(idx, { name: newName });
        // Update unique keys if renamed
        if (oldName && uniqueKeys.includes(oldName)) {
          setUniqueKeys(prev => prev.map(k => k === oldName ? newName : k));
        }
        // Update default sort if renamed
        setDefaultSort(prev => prev.map(s => s.column === oldName ? { ...s, column: newName } : s));
        return true;
      },
    },
    {
      headerName: 'Display Name',
      field: 'displayName',
      editable: true,
      flex: 2,
      minWidth: 100,
      valueSetter: (params: ValueSetterParams) => {
        updateColumn(params.data._idx, { displayName: params.newValue || undefined });
        return true;
      },
    },
    {
      headerName: 'Type',
      field: 'type',
      editable: true,
      flex: 1,
      minWidth: 120,
      cellEditor: TypeCellEditor,
      cellEditorPopup: true,
      cellRenderer: (params: { value: string; data: { type: string; name: string; _idx: number } }) => {
        const label = typeOptions.find(o => o.value === params.value)?.label ?? params.value;
        const isRef = params.data.type === 'reference';
        return React.createElement('div', {
          style: { display: 'flex', alignItems: 'center', gap: 4, width: '100%' },
        },
          React.createElement('span', { style: { flex: 1 } }, label),
          isRef && React.createElement('span', {
            onMouseDown: (e: React.MouseEvent) => {
              e.stopPropagation();
              e.preventDefault();
              openRefConfigDialog(params.data._idx);
            },
            title: 'Configure reference',
            style: { cursor: 'pointer', color: 'var(--ref-color, #2563eb)', fontSize: 14, lineHeight: 1, padding: '0 2px' },
          }, '✎'),
        );
      },
      valueSetter: (params: ValueSetterParams) => {
        const newType = params.newValue as ColumnType;
        const idx = params.data._idx;
        const oldType = columns[idx].type;
        if (oldType === newType) return false;

        const refUpdates: Partial<ColumnDef> = {
          type: newType,
          refTable: newType === 'reference' ? (otherTableIds[0] ?? '') : undefined,
          refDisplayColumns: newType === 'reference' ? [] : undefined,
          refSearchColumns: newType === 'reference' ? [] : undefined,
        };

        // For existing tables with data, show migration preview
        if (!isCreate && tableId) {
          const rows = state.getRows(tableId);
          const colName = columns[idx].name.trim();
          if (rows.length > 0 && colName) {
            // Text -> reference needs user-configured matching table/column
            if (oldType === 'text' && newType === 'reference') {
              setMigrationTargetColIdx(idx);
              setMigrationsDialogOpen(true);
              return false;
            }
            const preview = previewMigration(rows, colName, oldType, newType);
            if (preview.nonEmptyCount > 0) {
              setMigrationPreview({ preview, colIndex: idx, newType, refUpdates });
              return false; // Don't apply yet — wait for dialog confirmation
            }
          }
        }

        // No data to migrate (new table or empty column) — apply immediately
        updateColumn(idx, refUpdates);
        if (newType === 'reference') {
          requestAnimationFrame(() => setRefDialogCol(idx));
        }
        return true;
      },
    },
    {
      headerName: 'Unique',
      field: 'isKey',
      width: 75,
      maxWidth: 75,
      cellRenderer: (params: { value: boolean }) => params.value ? '✓' : '',
      cellStyle: () => ({ textAlign: 'center', cursor: 'pointer' }),
      onCellClicked: (params: CellClickedEvent) => {
        const name = params.data.name?.trim();
        if (!name) return;
        setUniqueKeys(prev =>
          prev.includes(name) ? prev.filter(k => k !== name) : [...prev, name]
        );
      },
    },
    {
      headerName: 'Sort',
      field: 'sortDir',
      width: 80,
      maxWidth: 80,
      editable: true,
      cellEditor: 'agSelectCellEditor',
      cellEditorParams: {
        values: ['', 'asc', 'desc'],
      },
      valueFormatter: (params) => {
        if (params.value === 'asc') return '↑ Asc';
        if (params.value === 'desc') return '↓ Desc';
        return '';
      },
      valueSetter: (params: ValueSetterParams) => {
        const colName = params.data.name?.trim();
        if (!colName) return false;
        const newDir = params.newValue as string;
        if (newDir === 'asc' || newDir === 'desc') {
          setDefaultSort(prev => {
            const existing = prev.findIndex(s => s.column === colName);
            if (existing >= 0) {
              return prev.map((s, j) => j === existing ? { ...s, direction: newDir as 'asc' | 'desc' } : s);
            }
            return [...prev, { column: colName, direction: newDir as 'asc' | 'desc' }];
          });
        } else {
          setDefaultSort(prev => prev.filter(s => s.column !== colName));
        }
        return true;
      },
    },
    {
      headerName: '',
      width: 50,
      maxWidth: 50,
      editable: false,
      sortable: false,
      filter: false,
      cellRenderer: () => {
        return columns.length <= 1 ? '' : '×';
      },
      cellStyle: () => ({
        cursor: columns.length <= 1 ? 'default' : 'pointer',
        textAlign: 'center',
        color: '#888',
      }),
      onCellClicked: (params: CellClickedEvent) => {
        if (columns.length <= 1) return;
        removeColumn(params.data._idx);
      },
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
  ], [columns, uniqueKeys, defaultSort, otherTableIds, openRefConfigDialog, getRefTableColumns]);

  if (tableId && !schema) {
    return (
      <div className="edit-table-page">
        <div className="edit-table-card">
          <h2>Table not found</h2>
          <button className="btn-secondary" onClick={() => navigate(bookBase || '/')}>
            Back
          </button>
        </div>
      </div>
    );
  }

  const updateColumn = (index: number, updates: Partial<ColumnDef>) => {
    setColumns(prev => prev.map((c, i) => i === index ? { ...c, ...updates } : c));
  };

  // Build a TableSchema from current local state, optionally with column overrides
  const buildSchema = (columnOverrides?: ColumnDef[]): TableSchema => {
    const cols = columnOverrides ?? columns;
    const colNames = cols.map(c => c.name.trim());
    const trimmedTableName = tableName.trim();
    return {
      name: trimmedTableName,
      columns: cols.map(c => ({
        ...c,
        name: c.name.trim(),
        displayName: c.displayName?.trim() || undefined,
      })),
      csvFileId: csvFileId.trim() || undefined,
      uniqueKeys: uniqueKeys.filter(uk => colNames.includes(uk)),
      defaultSort: defaultSort.filter(s => colNames.includes(s.column)),
      draftRowPosition,
    };
  };

  // Immediately apply a type migration and save the schema
  const applyMigrationNow = (colIndex: number, newType: ColumnType, refUpdates: Partial<ColumnDef>, dateFormat?: string) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    const colName = columns[colIndex].name.trim();
    const oldType = columns[colIndex].type;
    // Migrate data in-place
    applyMigration(table.rows, colName, oldType, newType, dateFormat);
    // Update local column state
    const newColumns = columns.map((c, i) => i === colIndex ? { ...c, ...refUpdates } : c);
    setColumns(newColumns);
    // Save schema immediately
    state.updateSchema(tableId, buildSchema(newColumns));
  };

  // Migrate to reference by matching source->ref column pairs into a named result column.
  const applyTextToReferenceMigrationNow = (
    resultColName: string,
    refTableId: string,
    pairs: { sourceColumn: string; refColumn: string }[],
  ) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    const refTable = state.model.getTable(refTableId);
    if (!refTable) return;

    const colName = resultColName.trim();
    if (!colName) return;

    const sourceColumns = Array.from(new Set(pairs.map(p => p.sourceColumn))).filter(Boolean);

    let matched = 0;
    let unmatched = 0;
    for (const row of table.rows) {
      const sourceValues = pairs.map(p => String(row[p.sourceColumn] ?? '').trim());
      let refValue = '';

      if (!sourceValues.every(v => v === '')) {
        const refMatch = refTable.rows.find(refRow =>
          pairs.every((p, i) => {
            const left = sourceValues[i].toLowerCase();
            const right = String(state.model.resolveColumnPath(refTableId, refRow, p.refColumn) ?? '').trim().toLowerCase();
            return left === right;
          })
        );

        if (refMatch) {
          refValue = refMatch[INTERNAL_ROW_ID];
          matched++;
        } else {
          unmatched++;
        }
      }

      // Remove selected source columns, then keep only the resulting reference column.
      for (const sc of sourceColumns) delete row[sc];
      row[colName] = refValue;
    }

    const uniqueRefColumns = Array.from(new Set(pairs.map(p => p.refColumn))).filter(Boolean);

    const sourceSet = new Set(sourceColumns);
    const removedIndices = columns
      .map((c, i) => ({ name: c.name.trim(), i }))
      .filter(x => sourceSet.has(x.name))
      .map(x => x.i);
    const removedNames = new Set(removedIndices.map(i => columns[i].name.trim()));
    const insertAt = removedIndices.length > 0 ? Math.min(...removedIndices) : columns.length;

    const baseColumns = columns.filter((c) => {
      const n = c.name.trim();
      if (sourceSet.has(n)) return false;
      if (n === colName) return false;
      return true;
    });

    const newRefCol: ColumnDef = {
      name: colName,
      type: 'reference' as ColumnType,
      refTable: refTableId,
      refDisplayColumns: uniqueRefColumns,
      refSearchColumns: uniqueRefColumns,
    };
    const newColumns = [...baseColumns];
    newColumns.splice(Math.min(insertAt, newColumns.length), 0, newRefCol);

    setColumns(newColumns);
    setUniqueKeys(prev => prev.filter(k => !removedNames.has(k) && k !== colName));
    setDefaultSort(prev => prev.filter(s => !removedNames.has(s.column) && s.column !== colName));
    state.updateSchema(tableId, buildSchema(newColumns));
    setNotice({
      kind: unmatched > 0 ? 'info' : 'success',
      message:
      unmatched > 0
        ? `Matched ${matched} value${matched !== 1 ? 's' : ''}; ${unmatched} unmatched value${unmatched !== 1 ? 's were' : ' was'} cleared.`
        : `Matched ${matched} value${matched !== 1 ? 's' : ''}.`,
    });
  };

  const applyTrimNormalizationNow = (columnNames: string[]) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    let trimmed = 0;
    for (const row of table.rows) {
      for (const colName of columnNames) {
        const v = row[colName];
        if (typeof v === 'string') {
          const t = v.trim();
          if (t !== v) {
            row[colName] = t;
            trimmed++;
          }
        }
      }
    }
    setNotice({
      kind: trimmed > 0 ? 'success' : 'info',
      message: trimmed > 0
        ? `Trimmed whitespace in ${trimmed} cell${trimmed !== 1 ? 's' : ''}.`
        : 'No whitespace to trim.',
    });
  };

  // Extract unique tuples from selected columns into a new reference table
  const applyExtractNow = (selectedColIndices: number[], resultColumnNames: string[], newTableName: string, refColName: string) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    const sourceColNames = selectedColIndices.map(i => columns[i].name.trim());

    // Collect unique tuples
    const seen = new Set<string>();
    const tuples: string[][] = [];
    for (const row of table.rows) {
      const values = sourceColNames.map(cn => row[cn] ?? '');
      if (values.every(v => v === '')) continue;
      const key = JSON.stringify(values);
      if (!seen.has(key)) {
        seen.add(key);
        tuples.push(values);
      }
    }

    // Create the new table with the user-chosen column names
    const newSchema: TableSchema = {
      name: newTableName,
      columns: resultColumnNames.map(name => ({ name, type: 'text' as ColumnType })),
      uniqueKeys: resultColumnNames,
    };
    const newRows = tuples.map(tuple => {
      const row: Row = {};
      resultColumnNames.forEach((name, i) => { row[name] = tuple[i]; });
      return row;
    });
    state.createTable(newSchema, newRows);

    // Build a map from tuple → _rowId in the new table
    const newTable = state.model.getTable(newTableName);
    if (!newTable) return;
    const tupleToRowId = new Map<string, string>();
    for (const row of newTable.rows) {
      const values = resultColumnNames.map(cn => row[cn] ?? '');
      tupleToRowId.set(JSON.stringify(values), row[INTERNAL_ROW_ID]);
    }

    // Migrate source rows: remove extracted columns, then set the new reference value.
    // The order matters when refColName matches one of the extracted source column names.
    for (const row of table.rows) {
      const values = sourceColNames.map(cn => row[cn] ?? '');
      const key = JSON.stringify(values);
      const refValue = values.every(v => v === '') ? '' : (tupleToRowId.get(key) ?? '');
      for (const cn of sourceColNames) {
        delete row[cn];
      }
      row[refColName] = refValue;
    }

    // Update local column state: remove all selected columns, add new reference column
    const selectedSet = new Set(selectedColIndices);
    const removedNames = new Set(selectedColIndices.map(i => columns[i].name.trim()));
    const newRefCol: ColumnDef = {
      name: refColName,
      type: 'reference' as ColumnType,
      refTable: newTableName,
      refDisplayColumns: resultColumnNames,
      refSearchColumns: resultColumnNames,
    };
    // Insert reference column at the position of the first removed column
    const insertAt = Math.min(...selectedColIndices);
    const newColumns = columns.filter((_, i) => !selectedSet.has(i));
    newColumns.splice(insertAt, 0, newRefCol);

    setColumns(newColumns);

    // Clean up uniqueKeys and defaultSort for removed columns
    setUniqueKeys(prev => prev.filter(k => !removedNames.has(k)));
    setDefaultSort(prev => prev.filter(s => !removedNames.has(s.column)));

    // Save schema immediately
    state.updateSchema(tableId, buildSchema(newColumns));
  };

  const addColumn = () => {
    setColumns(prev => [
      ...prev,
      { name: '', type: 'text' as ColumnType },
    ]);
  };

  const removeColumn = (index: number) => {
    if (columns.length <= 1) return;
    const removedName = columns[index].name;
    setColumns(prev => prev.filter((_, i) => i !== index));
    setUniqueKeys(prev => prev.filter(k => k !== removedName));
  };

  const handleSave = async () => {
    setError(null);
    setNotice(null);

    // Validate table name
    const trimmedName = tableName.trim();
    if (!trimmedName) {
      setError('Table name is required');
      return;
    }

    // Validate columns
    for (const col of columns) {
      if (!col.name.trim()) {
        setError('All columns must have a name');
        return;
      }
      if (col.type === 'reference' && !col.refTable) {
        setError(`Reference column "${col.name}" must specify a table`);
        return;
      }
    }

    // Validate unique keys reference valid columns
    const colNames = columns.map(c => c.name.trim());
    for (const uk of uniqueKeys) {
      if (!colNames.includes(uk)) {
        setError(`Unique key column "${uk}" does not match any column`);
        return;
      }
    }

    // Check for duplicate column names
    const names = new Set<string>();
    for (const col of columns) {
      if (names.has(col.name.trim())) {
        setError(`Duplicate column name "${col.name.trim()}"`);
        return;
      }
      names.add(col.name.trim());
    }

    // Check table name uniqueness
    const nameConflict = isCreate
      ? state.tableIds.includes(trimmedName)
      : trimmedName !== tableId && state.tableIds.includes(trimmedName);
    if (nameConflict) {
      setError(`A table named "${trimmedName}" already exists`);
      return;
    }

    // Build schema
    const newSchema: TableSchema = {
      name: trimmedName,
      csvFileId: csvFileId.trim() || undefined,
      columns: columns.map(c => ({
        ...c,
        name: c.name.trim(),
        displayName: c.displayName?.trim() || undefined,
      })),
      uniqueKeys,
      defaultSort: defaultSort.filter(s => colNames.includes(s.column)),
      draftRowPosition,
    };

    if (isCreate) {
      state.createTable(newSchema);
      navigate(toBookPath(`/table/${encodeURIComponent(trimmedName)}`), { replace: true });
    } else {
      // Propagate simple index-based column renames before applying the schema.
      if (schema && tableId) {
        const oldColumns = schema.columns;
        const nextColumns = newSchema.columns;
        const max = Math.min(oldColumns.length, nextColumns.length);
        for (let i = 0; i < max; i++) {
          const oldName = oldColumns[i].name?.trim();
          const newName = nextColumns[i].name?.trim();
          if (!oldName || !newName || oldName === newName) continue;
          state.renameColumn(tableId, oldName, newName);
        }
      }

      // Apply schema update
      state.updateSchema(tableId!, newSchema);

      if (
        trimmedName !== tableId &&
        state.isSignedIn &&
        state.folderId &&
        !state.folderId.startsWith('local-') &&
        !!(schema?.csvFileId || csvFileId)
      ) {
        const suggestedFileName = `${trimmedName}.csv`;
        const currentFileName = csvFileName || '(selected CSV)';
        const shouldRenameCsv = await showDialog({
          title: 'Rename Drive CSV?',
          message: `Rename the bound Drive CSV from "${currentFileName}" to "${suggestedFileName}"?`,
          buttons: [
            { label: 'Keep Current CSV Name', value: 'keep', variant: 'secondary' },
            { label: 'Rename CSV', value: 'rename', variant: 'primary' },
          ],
        });

        if (shouldRenameCsv === 'rename') {
          try {
            await state.renameTableCsvFile(tableId!, suggestedFileName);
            setCsvFileName(suggestedFileName);
            setNotice({ kind: 'success', message: `Renamed CSV to ${suggestedFileName}.` });
          } catch (err) {
            setNotice({
              kind: 'info',
              message: err instanceof Error
                ? `Table renamed, but CSV rename failed: ${err.message}`
                : 'Table renamed, but CSV rename failed.',
            });
          }
        }
      }

      // Rename table if name changed
      if (trimmedName !== tableId) {
        state.renameTable(tableId!, trimmedName);
        navigate(toBookPath(`/table/${encodeURIComponent(trimmedName)}`), { replace: true });
      } else {
        navigate(toBookPath(`/table/${encodeURIComponent(tableId!)}`));
      }
    }
  };

  const handleDelete = async () => {
    if (!tableId) return;
    const buttons = [
      { label: 'Cancel', value: 'cancel', variant: 'secondary' as const },
      { label: 'Delete Table', value: 'delete', variant: 'danger' as const },
    ];
    if (state.folderId) {
      buttons.push({ label: 'Delete Table & CSV', value: 'delete-drive', variant: 'danger' as const });
    }
    const result = await showDialog({
      title: 'Delete Table',
      message: `Delete table "${tableId}"? This cannot be undone.`,
      buttons,
    });
    if (result === 'delete' || result === 'delete-drive') {
      state.deleteTable(tableId, result === 'delete-drive');
      navigate(bookBase || '/');
    }
  };

  const handlePickCsvFromDrive = async () => {
    if (!state.isSignedIn || !state.folderId || state.folderId.startsWith('local-')) {
      setError('Sign in and open a Drive-backed book to use Drive Picker.');
      return;
    }

    setError(null);
    setPickingCsv(true);
    try {
      const picked = await drive.pickCsvFileInFolder(state.folderId);
      if (!picked) return;
      if (!picked.name.toLowerCase().endsWith('.csv')) {
        setError('Please choose a .csv file.');
        return;
      }
      setCsvFileId(picked.id);
      setCsvFileName(picked.name);
      setAvailableCsvFiles(prev => prev.some(f => f.id === picked.id)
        ? prev
        : [...prev, { id: picked.id, name: picked.name }].sort((a, b) => a.name.localeCompare(b.name)));

      if (!isCreate && tableId) {
        await state.loadTableFromCsvFile(tableId, picked.id);
        const rowCount = state.getRows(tableId).length;
        setNotice({ kind: 'info', message: `Loaded ${rowCount} row${rowCount === 1 ? '' : 's'} from ${picked.name}.` });
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to open Drive Picker.');
    } finally {
      setPickingCsv(false);
    }
  };

  return (
    <div className="edit-table-page">
      <div className="edit-table-card">
        <div className="edit-table-header">
          <h2>{isCreate ? 'New Table' : 'Edit Table'}</h2>
          <button className="btn-secondary btn-sm" onClick={() => navigate(tableId ? toBookPath(`/table/${encodeURIComponent(tableId)}`) : (bookBase || '/'))}>
            {isCreate ? '← Back' : '← Back to Table'}
          </button>
        </div>

        {error && <div className="edit-table-error">{error}</div>}
        {notice && <div className={`edit-table-notice edit-table-notice-${notice.kind}`}>{notice.message}</div>}

        <div className="edit-table-field">
          <label>Table Name</label>
          <input
            type="text"
            value={tableName}
            onChange={e => setTableName(e.target.value)}
            className="edit-table-input"
          />
        </div>

        <div className="edit-table-field">
          <label>CSV File</label>
          <div style={{ display: 'flex', gap: 8 }}>
            <input
              type="text"
              value={csvFileName}
              className="edit-table-input"
              placeholder="Pick a CSV file"
              list="csv-file-options"
              style={{ flex: 1 }}
              readOnly
            />
            <button
              className="btn-secondary"
              onClick={() => { void handlePickCsvFromDrive(); }}
              disabled={pickingCsv || !state.isSignedIn || !state.folderId || state.folderId.startsWith('local-')}
              title="Pick a CSV file from the current book folder"
            >
              {pickingCsv ? 'Opening...' : 'Pick from Drive'}
            </button>
          </div>
          <datalist id="csv-file-options">
            {availableCsvFiles.map(file => (
              <option key={file.id || file.name} value={file.name} />
            ))}
          </datalist>
          {csvFileId && (
            <div className="book-settings-note" style={{ marginTop: 6 }}>
              File ID: {csvFileId}
            </div>
          )}
          {state.isSignedIn && state.folderId && !state.folderId.startsWith('local-') && (
            <div className="book-settings-note" style={{ marginTop: 6 }}>
              Picker opens directly in this book folder.
            </div>
          )}
        </div>

        {/* New Row Position */}
        <div className="edit-table-section">
          <label className="edit-table-label">New Row Position</label>
          <div style={{ maxWidth: 200 }}>
            <Select
              value={{ value: draftRowPosition, label: draftRowPosition === 'bottom' ? 'Bottom' : 'Top' }}
              onChange={opt => setDraftRowPosition(opt?.value ?? 'bottom')}
              options={[{ value: 'bottom', label: 'Bottom' }, { value: 'top', label: 'Top' }]}
              styles={selectStyles}
              isSearchable={false}
              menuPlacement="auto"
            />
          </div>
        </div>

        {/* AG Grid column editor */}
        <div className="edit-table-columns">
          <h3>Columns</h3>
          <div style={{ width: '100%' }}>
            <AgGridReact
              theme={editGridTheme}
              modules={[AllCommunityModule]}
              rowData={columnRowData}
              columnDefs={columnGridDefs}
              domLayout="autoHeight"
              singleClickEdit={true}
              stopEditingWhenCellsLoseFocus={true}
              getRowId={(params) => String(params.data._idx)}
              suppressRowDrag={true}
            />
          </div>
          <div style={{ display: 'flex', gap: 8, marginTop: 8 }}>
            <button className="btn-secondary btn-sm" onClick={addColumn}>
              + Add Column
            </button>
            {!isCreate && tableId && (
              <button
                className="btn-secondary btn-sm"
                onClick={() => {
                  setExtractPreview(true);
                }}
              >
                Extract to Reference Table
              </button>
            )}
            {!isCreate && tableId && (
              <button
                className="btn-secondary btn-sm"
                onClick={() => {
                  setMigrationsDialogOpen(true);
                }}
              >
                Migrations & Normalize
              </button>
            )}
          </div>
        </div>

        {/* Reference config dialog */}
        {refDialogCol !== null && columns[refDialogCol]?.type === 'reference' && (
          <RefConfigDialog
            col={columns[refDialogCol]}
            colIndex={refDialogCol}
            otherTableIds={otherTableIds}
            getRefTableColumns={getRefTableColumns}
            getRefTableColumnPaths={getRefTableColumnPaths}
            onUpdate={(idx, updates) => updateColumn(idx, updates)}
            onClose={() => setRefDialogCol(null)}
          />
        )}

        {/* Unified migrations & normalization dialog */}
        {migrationsDialogOpen && (
          <MigrationsToolsDialog
            columns={columns}
            rows={state.getRows(tableId!)}
            otherTableIds={otherTableIds}
            getRefTableColumns={getRefTableColumns}
            getRefTableRows={getRefTableRows}
            resolveRefPathValue={resolveRefPathValue}
            initialTargetColIdx={migrationTargetColIdx}
            initialResultColName={migrationTargetColIdx !== null ? (columns[migrationTargetColIdx]?.name.trim() ?? '') : ''}
            onRunNormalize={applyTrimNormalizationNow}
            onRunReference={applyTextToReferenceMigrationNow}
            onClose={() => {
              setMigrationsDialogOpen(false);
              setMigrationTargetColIdx(null);
            }}
          />
        )}

        {/* Migration preview dialog */}
        {migrationPreview && (
          <MigrationPreviewDialog
            preview={migrationPreview.preview}
            dateFormat={migrationPreview.dateFormat}
            onDateFormatChange={(fmt: string) => {
              setMigrationPreview(prev => {
                if (!prev) return prev;
                // Recompute preview with new dateFormat
                const table = tableId ? state.model.getTable(tableId) : undefined;
                const rows = table ? table.rows : [];
                const colName = columns[prev.colIndex].name.trim();
                const oldType = columns[prev.colIndex].type;
                const preview = previewMigration(rows, colName, oldType, prev.newType, fmt);
                return { ...prev, dateFormat: fmt, preview };
              });
            }}
            onConfirm={() => {
              const { colIndex, newType, refUpdates } = migrationPreview;
              const dateFormat = migrationPreview.dateFormat;
              applyMigrationNow(colIndex, newType, refUpdates!, dateFormat);
              setMigrationPreview(null);
              if (newType === 'reference') {
                requestAnimationFrame(() => setRefDialogCol(colIndex));
              }
            }}
            onCancel={() => setMigrationPreview(null)}
          />
        )}

        {/* Extract to reference table preview dialog */}
        {extractPreview && (
          <ExtractPreviewDialog
            columns={columns}
            rows={state.getRows(tableId!)}
            existingTables={state.tableIds}
            onConfirm={(selectedColIndices, resultColumnNames, newTableName, refColName) => {
              applyExtractNow(selectedColIndices, resultColumnNames, newTableName, refColName);
              setExtractPreview(false);
            }}
            onCancel={() => setExtractPreview(false)}
          />
        )}

        <div className="edit-table-actions">
          <button className="btn-secondary" onClick={() => navigate(tableId ? toBookPath(`/table/${encodeURIComponent(tableId)}`) : (bookBase || '/'))}>
            Cancel
          </button>
          {!isCreate && (
            <button className="btn-danger" onClick={handleDelete}>
              Delete Table
            </button>
          )}
          <div style={{ flex: 1 }} />
          <button className="btn-primary" onClick={() => { void handleSave(); }}>
            {isCreate ? 'Create Table' : 'Save Changes'}
          </button>
        </div>
      </div>
    </div>
  );
};

// Inline reference config dialog component
const RefConfigDialog: React.FC<{
  col: ColumnDef;
  colIndex: number;
  otherTableIds: string[];
  getRefTableColumns: (tableId: string) => string[];
  getRefTableColumnPaths: (tableId: string) => { path: string; label: string }[];
  onUpdate: (index: number, updates: Partial<ColumnDef>) => void;
  onClose: () => void;
}> = ({ col, colIndex, otherTableIds, getRefTableColumnPaths, onUpdate, onClose }) => {
  const columnPaths = col.refTable ? getRefTableColumnPaths(col.refTable) : [];

  // Build label lookup for currently selected paths
  const pathLabelMap = useMemo(() => {
    const map = new Map<string, string>();
    for (const cp of columnPaths) map.set(cp.path, cp.label);
    return map;
  }, [columnPaths]);

  return (
    <div className="app-dialog-overlay" onClick={onClose}>
      <div className="app-dialog" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 420 }}>
        <h3 className="app-dialog-title">Reference Config: {col.name || 'unnamed'}</h3>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12, padding: '8px 0' }}>
          <div>
            <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Referenced Table</label>
            <Select
              value={col.refTable ? { value: col.refTable, label: col.refTable } : null}
              onChange={opt => onUpdate(colIndex, {
                refTable: opt?.value ?? '',
                refDisplayColumns: [],
                refSearchColumns: [],
              })}
              options={otherTableIds.map(id => ({ value: id, label: id }))}
              placeholder="Select table..."
              styles={refDialogSelectStyles}
              isClearable
              menuPlacement="auto"
            />
          </div>
          {col.refTable && (
            <>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Display Columns</label>
                <Select
                  isMulti
                  value={(col.refDisplayColumns ?? []).map(cn => ({
                    value: cn,
                    label: pathLabelMap.get(cn) ?? cn,
                  }))}
                  onChange={opts => onUpdate(colIndex, { refDisplayColumns: opts.map(o => o.value) })}
                  options={columnPaths.map(cp => ({ value: cp.path, label: cp.label }))}
                  styles={refDialogSelectStyles}
                  placeholder="Select columns to display..."
                  menuPlacement="auto"
                />
              </div>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Search Columns</label>
                <Select
                  isMulti
                  value={(col.refSearchColumns ?? []).map(cn => ({
                    value: cn,
                    label: pathLabelMap.get(cn) ?? cn,
                  }))}
                  onChange={opts => onUpdate(colIndex, { refSearchColumns: opts.map(o => o.value) })}
                  options={columnPaths.map(cp => ({ value: cp.path, label: cp.label }))}
                  styles={refDialogSelectStyles}
                  placeholder="Select columns to search..."
                  menuPlacement="auto"
                />
              </div>
            </>
          )}
        </div>
        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-primary" onClick={onClose}>
            Done
          </button>
        </div>
      </div>
    </div>
  );
};

const refDialogSelectStyles = {
  control: (base: Record<string, unknown>, state: { isFocused: boolean }) => ({
    ...base,
    background: 'var(--bg)',
    borderColor: state.isFocused ? 'var(--ref-color)' : 'var(--border)',
    boxShadow: state.isFocused ? '0 0 0 2px rgba(37, 99, 235, 0.2)' : 'none',
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
  option: (base: Record<string, unknown>, state: { isSelected: boolean; isFocused: boolean }) => ({
    ...base,
    background: state.isSelected ? 'var(--primary)' : state.isFocused ? 'var(--cell-selected)' : 'transparent',
    color: state.isSelected ? '#fff' : 'var(--text)',
    fontSize: 13,
    padding: '4px 10px',
    cursor: 'pointer',
  }),
  singleValue: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  multiValue: (base: Record<string, unknown>) => ({ ...base, background: 'var(--cell-selected, #e0e7ff)' }),
  input: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  placeholder: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text-muted)' }),
  indicatorSeparator: () => ({ display: 'none' }),
};

const MigrationsToolsDialog: React.FC<{
  columns: ColumnDef[];
  rows: Row[];
  otherTableIds: string[];
  getRefTableColumns: (tableId: string) => string[];
  getRefTableRows: (tableId: string) => Row[];
  resolveRefPathValue: (tableId: string, row: Row, path: string) => string;
  initialTargetColIdx: number | null;
  initialResultColName: string;
  onRunNormalize: (columnNames: string[]) => void;
  onRunReference: (resultColName: string, refTableId: string, pairs: { sourceColumn: string; refColumn: string }[]) => void;
  onClose: () => void;
}> = ({
  columns,
  rows,
  otherTableIds,
  getRefTableColumns,
  getRefTableRows,
  resolveRefPathValue,
  initialTargetColIdx,
  initialResultColName,
  onRunNormalize,
  onRunReference,
  onClose,
}) => {
  const namedCols = useMemo(
    () => columns.map((c, i) => ({ idx: i, name: c.name.trim(), type: c.type })).filter(c => c.name),
    [columns]
  );

  const [normalizeCols, setNormalizeCols] = useState<string[]>([]);
  const [resultColName, setResultColName] = useState(initialResultColName);
  const [refTable, setRefTable] = useState(otherTableIds[0] ?? '');
  const [pairs, setPairs] = useState<{ sourceColumn: string; refColumn: string }[]>([]);

  const sourceColOptions = useMemo(
    () => namedCols.filter(c => c.type !== 'image').map(c => ({ value: c.name, label: c.name })),
    [namedCols]
  );

  const refColOptions = useMemo(
    () => (refTable ? getRefTableColumns(refTable) : []).map(c => ({ value: c, label: c })),
    [refTable, getRefTableColumns]
  );

  useEffect(() => {
    if (!refTable) {
      setPairs([]);
      return;
    }
    const firstSource = sourceColOptions[0]?.value ?? '';
    const firstRef = refColOptions[0]?.value ?? '';
    if (pairs.length === 0 && firstRef) {
      const initialSource = initialTargetColIdx !== null
        ? (columns[initialTargetColIdx]?.name.trim() || firstSource)
        : firstSource;
      if (initialSource) {
        setPairs([{ sourceColumn: initialSource, refColumn: firstRef }]);
      }
    }
  }, [refTable, sourceColOptions, refColOptions, pairs.length, initialTargetColIdx, columns]);

  const validPairs = pairs.filter(p => p.sourceColumn && p.refColumn);
  const canRunReference = !!resultColName.trim() && !!refTable && validPairs.length > 0;

  const referencePreview = useMemo(() => {
    if (!refTable || validPairs.length === 0) {
      return { matched: 0, unmatched: 0, empty: 0, total: rows.length };
    }

    const refRows = getRefTableRows(refTable);
    let matched = 0;
    let unmatched = 0;
    let empty = 0;

    for (const row of rows) {
      const sourceValues = validPairs.map(p => String(row[p.sourceColumn] ?? '').trim());
      if (sourceValues.every(v => v === '')) {
        empty++;
        continue;
      }

      const refMatch = refRows.find(refRow =>
        validPairs.every((p, i) => {
          const left = sourceValues[i].toLowerCase();
          const right = String(resolveRefPathValue(refTable, refRow, p.refColumn) ?? '').trim().toLowerCase();
          return left === right;
        })
      );

      if (refMatch) matched++;
      else unmatched++;
    }

    return { matched, unmatched, empty, total: rows.length };
  }, [refTable, validPairs, rows, getRefTableRows, resolveRefPathValue]);

  return (
    <div className="app-dialog-overlay" onClick={onClose}>
      <div className="app-dialog" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 640 }}>
        <h3 className="app-dialog-title">Migrations & Normalize</h3>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 16, padding: '8px 0' }}>
          <div style={{ border: '1px solid var(--border)', borderRadius: 8, padding: 12 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Normalize (Trim Whitespace)</div>
            <Select
              isMulti
              options={sourceColOptions}
              value={sourceColOptions.filter(o => normalizeCols.includes(o.value))}
              onChange={opts => setNormalizeCols(opts.map(o => o.value))}
              styles={refDialogSelectStyles}
              placeholder="Select columns to trim..."
              menuPlacement="auto"
            />
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 10 }}>
              <button
                className="app-dialog-btn app-dialog-btn-primary"
                disabled={normalizeCols.length === 0}
                onClick={() => {
                  onRunNormalize(normalizeCols);
                  onClose();
                }}
              >
                Run Normalize
              </button>
            </div>
          </div>

          <div style={{ border: '1px solid var(--border)', borderRadius: 8, padding: 12 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>Text to Reference Migration</div>
            <div style={{ display: 'grid', gap: 10 }}>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Resulting Reference Column Name</label>
                <input
                  className="edit-table-input"
                  value={resultColName}
                  onChange={e => setResultColName(e.target.value)}
                  placeholder="Enter resulting reference column name"
                />
              </div>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Referenced Table</label>
                <Select
                  options={otherTableIds.map(id => ({ value: id, label: id }))}
                  value={refTable ? { value: refTable, label: refTable } : null}
                  onChange={opt => {
                    setRefTable(opt?.value ?? '');
                    setPairs([]);
                  }}
                  styles={refDialogSelectStyles}
                  isClearable
                  menuPlacement="auto"
                />
              </div>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 6, display: 'block' }}>Match Pairs</label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                  {pairs.map((pair, i) => (
                    <div key={i} style={{ display: 'grid', gridTemplateColumns: '1fr auto 1fr auto', gap: 8, alignItems: 'center' }}>
                      <Select
                        options={sourceColOptions}
                        value={sourceColOptions.find(o => o.value === pair.sourceColumn) ?? null}
                        onChange={opt => setPairs(prev => prev.map((p, j) => j === i ? { ...p, sourceColumn: opt?.value ?? '' } : p))}
                        styles={refDialogSelectStyles}
                        placeholder="Source column"
                        menuPlacement="auto"
                      />
                      <span style={{ color: 'var(--text-muted)', fontSize: 12 }}>→</span>
                      <Select
                        options={refColOptions}
                        value={refColOptions.find(o => o.value === pair.refColumn) ?? null}
                        onChange={opt => setPairs(prev => prev.map((p, j) => j === i ? { ...p, refColumn: opt?.value ?? '' } : p))}
                        styles={refDialogSelectStyles}
                        placeholder="Ref column"
                        menuPlacement="auto"
                      />
                      <button
                        className="app-dialog-btn app-dialog-btn-secondary"
                        onClick={() => setPairs(prev => prev.filter((_, j) => j !== i))}
                        style={{ padding: '6px 10px' }}
                      >
                        Remove
                      </button>
                    </div>
                  ))}
                </div>
                <div style={{ marginTop: 8 }}>
                  <button
                    className="app-dialog-btn app-dialog-btn-secondary"
                    onClick={() => setPairs(prev => [
                      ...prev,
                      {
                        sourceColumn: sourceColOptions.find(o => !prev.some(p => p.sourceColumn === o.value))?.value ?? sourceColOptions[0]?.value ?? '',
                        refColumn: refColOptions.find(o => !prev.some(p => p.refColumn === o.value))?.value ?? refColOptions[0]?.value ?? '',
                      },
                    ])}
                  >
                    + Add Match Pair
                  </button>
                </div>
              </div>
              <div style={{ fontSize: 12, color: 'var(--text-muted)', background: 'var(--bg)', border: '1px solid var(--border)', borderRadius: 6, padding: '8px 10px' }}>
                Preview: {referencePreview.matched} matched, {referencePreview.unmatched} unmatched, {referencePreview.empty} empty of {referencePreview.total} rows.
              </div>
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 10 }}>
              <button
                className="app-dialog-btn app-dialog-btn-primary"
                disabled={!canRunReference}
                onClick={() => {
                  onRunReference(resultColName.trim(), refTable, validPairs);
                  onClose();
                }}
              >
                Run Reference Migration
              </button>
            </div>
          </div>
        </div>

        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-secondary" onClick={onClose}>
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

// Migration preview dialog component
const MigrationPreviewDialog: React.FC<{
  preview: MigrationPreview;
  dateFormat?: string;
  onDateFormatChange?: (fmt: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}> = ({ preview, dateFormat, onDateFormatChange, onConfirm, onCancel }) => {
  const fromLabel = typeOptions.find(o => o.value === preview.fromType)?.label ?? preview.fromType;
  const toLabel = typeOptions.find(o => o.value === preview.toType)?.label ?? preview.toType;

  const needsDateFormat = preview.fromType === 'text' && (preview.toType === 'date' || preview.toType === 'datetime');

  type OptionType = { value: string; label: string };

  const dateFormatOptions: OptionType[] = DATE_FORMATS.filter((f: any) => f.value !== 'auto').map((f: any) => ({ value: f.value, label: f.label }));
  const selectedDateFormat = dateFormatOptions.find((o: OptionType) => o.value === dateFormat) ?? null;

  return (
    <div className="app-dialog-overlay" onClick={onCancel}>
      <div className="app-dialog" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 560 }}>
        <h3 className="app-dialog-title">
          Convert "{preview.columnName}" from {fromLabel} → {toLabel}
        </h3>
        <div style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 12 }}>
          {preview.nonEmptyCount} value{preview.nonEmptyCount !== 1 ? 's' : ''} to convert
          {preview.errorCount > 0 && (
            <span style={{ color: 'var(--danger, #dc2626)', fontWeight: 500 }}>
              {' '}— {preview.errorCount} will be cleared (not convertible)
            </span>
          )}
          {preview.errorCount === 0 && preview.nonEmptyCount > 0 && (
            <span style={{ color: 'var(--success, #16a34a)', fontWeight: 500 }}>
              {' '}— all values convertible
            </span>
          )}
        </div>

        {/* Date format selection dropdown */}
        {needsDateFormat && (
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>
              Select date format for conversion
            </label>
            <Select<OptionType>
              options={dateFormatOptions}
              value={selectedDateFormat}
              onChange={opt => onDateFormatChange?.(opt?.value ?? '')}
              styles={selectStyles}
              placeholder="Choose date format..."
              menuPlacement="auto"
              isClearable={false}
            />
          </div>
        )}

        {preview.samples.length > 0 && (
          <div style={{ maxHeight: 240, overflow: 'auto', marginBottom: 12 }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
              <thead>
                <tr style={{ background: 'var(--surface-2, #f5f6f8)' }}>
                  <th style={{ padding: '4px 8px', textAlign: 'left', borderBottom: '1px solid var(--border)' }}>Current</th>
                  <th style={{ padding: '4px 8px', textAlign: 'left', borderBottom: '1px solid var(--border)' }}>Converted</th>
                  <th style={{ padding: '4px 8px', textAlign: 'left', borderBottom: '1px solid var(--border)' }}>Status</th>
                </tr>
              </thead>
              <tbody>
                {preview.samples.map((s, i) => (
                  <tr key={i} style={{ background: s.error ? 'rgba(220, 38, 38, 0.05)' : 'transparent' }}>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid var(--border)', fontFamily: 'monospace' }}>
                      {s.original}
                    </td>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid var(--border)', fontFamily: 'monospace' }}>
                      {s.error ? <span style={{ opacity: 0.4 }}>empty</span> : s.converted}
                    </td>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid var(--border)' }}>
                      {s.error ? (
                        <span style={{ color: 'var(--danger, #dc2626)', fontSize: 11 }}>{s.error}</span>
                      ) : (
                        <span style={{ color: 'var(--success, #16a34a)' }}>✓</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {preview.nonEmptyCount > 10 && (
          <div style={{ fontSize: 11, color: 'var(--text-muted)', marginBottom: 8 }}>
            Showing {preview.samples.length} of {preview.nonEmptyCount} values
          </div>
        )}

        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-secondary" onClick={onCancel}>
            Cancel
          </button>
          <button
            className="app-dialog-btn app-dialog-btn-primary"
            onClick={onConfirm}
            disabled={needsDateFormat && !dateFormat}
          >
            {preview.errorCount > 0 ? `Convert (${preview.errorCount} will be cleared)` : 'Convert'}
          </button>
        </div>
      </div>
    </div>
  );
};

// Extract to reference table preview dialog
const ExtractPreviewDialog: React.FC<{
  columns: ColumnDef[];
  rows: Row[];
  existingTables: string[];
  onConfirm: (selectedColIndices: number[], resultColumnNames: string[], newTableName: string, refColName: string) => void;
  onCancel: () => void;
}> = ({ columns, rows, existingTables, onConfirm, onCancel }) => {
  // Available columns for extraction: non-reference columns with names
  const availableCols = useMemo(() =>
    columns
      .map((c, i) => ({ col: c, idx: i }))
      .filter(({ col }) => col.type !== 'reference' && col.name.trim()),
    [columns]
  );

  const columnOptions = useMemo(() =>
    availableCols.map(({ col, idx }) => ({ value: idx, label: col.name })),
    [availableCols]
  );

  const [selectedIndices, setSelectedIndices] = useState<number[]>([]);
  const [newTableName, setNewTableName] = useState('');
  const [refColName, setRefColName] = useState('');
  const [resultNames, setResultNames] = useState<Record<number, string>>({});

  const sortedSelectedIndices = useMemo(() =>
    [...selectedIndices].sort((a, b) => a - b),
    [selectedIndices]
  );

  // Compute preview when selection changes
  const preview = useMemo(() => {
    const colNames = sortedSelectedIndices.map(i => columns[i].name.trim());
    if (colNames.length === 0) return null;
    return previewExtract(rows, colNames);
  }, [sortedSelectedIndices, columns, rows]);

  const getResultName = (idx: number) => resultNames[idx] ?? columns[idx].name.trim();

  const nameConflict = newTableName.trim() !== '' && existingTables.includes(newTableName.trim());
  const nameEmpty = !newTableName.trim();
  const refNameEmpty = !refColName.trim();
  const noColumns = selectedIndices.length === 0;

  const handleConfirm = () => {
    const resultColumnNamesList = sortedSelectedIndices.map(i => getResultName(i));
    onConfirm(sortedSelectedIndices, resultColumnNamesList, newTableName.trim(), refColName.trim());
  };

  return (
    <div className="app-dialog-overlay" onClick={onCancel}>
      <div className="app-dialog" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 560 }}>
        <h3 className="app-dialog-title">Extract to Reference Table</h3>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12, padding: '8px 0' }}>

          <div>
            <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Columns to Extract</label>
            <Select
              isMulti
              value={selectedIndices.map(idx => ({ value: idx, label: columns[idx].name }))}
              onChange={opts => {
                const indices = opts.map(o => o.value);
                setSelectedIndices(indices);
                // Default table name to first selected column
                if (indices.length > 0 && !newTableName.trim()) {
                  setNewTableName(columns[indices[0]].name.trim());
                }
                // Default ref column name to first selected column
                if (indices.length > 0 && !refColName.trim()) {
                  setRefColName(columns[indices[0]].name.trim());
                }
              }}
              options={columnOptions}
              styles={refDialogSelectStyles}
              placeholder="Select columns..."
              menuPlacement="auto"
            />
          </div>

          {selectedIndices.length > 0 && (
            <>
              {/* Rename columns in new table */}
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Column Names in New Table</label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {sortedSelectedIndices.map(idx => (
                    <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span style={{ fontSize: 12, color: 'var(--text-muted)', minWidth: 80 }}>{columns[idx].name} →</span>
                      <input
                        type="text"
                        value={getResultName(idx)}
                        onChange={e => setResultNames(prev => ({ ...prev, [idx]: e.target.value }))}
                        className="edit-table-input"
                        style={{ flex: 1, fontSize: 13 }}
                      />
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>New Table Name</label>
                <input
                  type="text"
                  value={newTableName}
                  onChange={e => setNewTableName(e.target.value)}
                  className="edit-table-input"
                  style={{ width: '100%' }}
                />
                {nameConflict && (
                  <div style={{ color: 'var(--danger, #dc2626)', fontSize: 11, marginTop: 4 }}>
                    A table named &quot;{newTableName}&quot; already exists
                  </div>
                )}
              </div>

              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Reference Column Name</label>
                <input
                  type="text"
                  value={refColName}
                  onChange={e => setRefColName(e.target.value)}
                  className="edit-table-input"
                  style={{ width: '100%' }}
                  placeholder="Name for the reference column in this table"
                />
                <div style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 4 }}>
                  The extracted columns will be replaced with this reference column.
                </div>
              </div>
            </>
          )}

          {/* Preview table */}
          {preview && preview.uniqueTuples.length > 0 && (
            <div>
              <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>
                {preview.uniqueTuples.length} unique {selectedIndices.length > 1 ? 'tuple' : 'value'}{preview.uniqueTuples.length !== 1 ? 's' : ''} from {preview.nonEmptyCount} non-empty row{preview.nonEmptyCount !== 1 ? 's' : ''}
              </label>
              <div style={{ maxHeight: 200, overflow: 'auto', border: '1px solid var(--border)', borderRadius: 4 }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: 'var(--surface-2, #f5f6f8)' }}>
                      {sortedSelectedIndices.map(idx => (
                        <th key={idx} style={{ padding: '4px 8px', textAlign: 'left', borderBottom: '1px solid var(--border)' }}>
                          {getResultName(idx)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.uniqueTuples.slice(0, 50).map((tuple, i) => (
                      <tr key={i}>
                        {tuple.map((v, j) => (
                          <td key={j} style={{ padding: '4px 8px', borderBottom: '1px solid var(--border)', fontFamily: 'monospace' }}>
                            {v || <span style={{ opacity: 0.3 }}>empty</span>}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {preview.uniqueTuples.length > 50 && (
                <div style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 4 }}>
                  Showing 50 of {preview.uniqueTuples.length} rows
                </div>
              )}
            </div>
          )}
        </div>

        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-secondary" onClick={onCancel}>
            Cancel
          </button>
          <button
            className="app-dialog-btn app-dialog-btn-primary"
            onClick={handleConfirm}
            disabled={nameConflict || nameEmpty || refNameEmpty || noColumns}
          >
            Extract
          </button>
        </div>
      </div>
    </div>
  );
};

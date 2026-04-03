import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import Select from 'react-select';
import type { ColumnDef, ColumnType, TableSchema } from './types';
import { INTERNAL_ROW_ID } from './types';
import type { UseAppStateReturn } from './useAppState';
import { useDialog } from './DialogProvider';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, ValueSetterParams, CellClickedEvent } from 'ag-grid-community';
import type { CustomCellEditorProps } from 'ag-grid-react';
import { previewMigration, applyMigration, previewExtract } from './typeMigration';
import type { MigrationPreview, ExtractPreview } from './typeMigration';

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
  input: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text)' }),
  placeholder: (base: Record<string, unknown>) => ({ ...base, color: 'var(--text-muted)' }),
  indicatorSeparator: () => ({ display: 'none' }),
};

interface EditTablePageProps {
  state: UseAppStateReturn;
}

export const EditTablePage: React.FC<EditTablePageProps> = ({ state }) => {
  const { tableId } = useParams<{ tableId: string }>();
  const navigate = useNavigate();
  const { showDialog } = useDialog();
  const isCreate = !tableId;
  const schema = tableId ? state.getSchema(tableId) : undefined;

  const [tableName, setTableName] = useState(schema?.name ?? '');
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

  // Type migration preview (for existing tables only)
  const [migrationPreview, setMigrationPreview] = useState<{
    preview: MigrationPreview;
    colIndex: number;
    newType: ColumnType;
    refUpdates?: Partial<ColumnDef>;
  } | null>(null);

  // Extract-to-table preview
  const [extractPreview, setExtractPreview] = useState<{
    preview: ExtractPreview;
    colIndex: number;
  } | null>(null);

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
        const canExtract = !isCreate && tableId && params.data.name?.trim() && !isRef;
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
          canExtract && React.createElement('span', {
            onMouseDown: (e: React.MouseEvent) => {
              e.stopPropagation();
              e.preventDefault();
              const idx = params.data._idx;
              const colName = columns[idx].name.trim();
              const rows = state.getRows(tableId!);
              const preview = previewExtract(rows, colName, colName);
              if (preview.uniqueValues.length === 0) {
                setError('No values to extract — column is empty');
                return;
              }
              setExtractPreview({ preview, colIndex: idx });
            },
            title: 'Extract to reference table',
            style: { cursor: 'pointer', color: 'var(--text-muted, #888)', fontSize: 14, lineHeight: 1, padding: '0 2px' },
          }, '⤴'),
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
  ], [columns, uniqueKeys, defaultSort, otherTableIds, openRefConfigDialog]);

  if (tableId && !schema) {
    return (
      <div className="edit-table-page">
        <div className="edit-table-card">
          <h2>Table not found</h2>
          <button className="btn-secondary" onClick={() => navigate('/')}>
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
    return {
      name: tableName.trim(),
      columns: cols.map(c => ({
        ...c,
        name: c.name.trim(),
        displayName: c.displayName?.trim() || undefined,
      })),
      uniqueKeys: uniqueKeys.filter(uk => colNames.includes(uk)),
      defaultSort: defaultSort.filter(s => colNames.includes(s.column)),
      draftRowPosition,
    };
  };

  // Immediately apply a type migration and save the schema
  const applyMigrationNow = (colIndex: number, newType: ColumnType, refUpdates: Partial<ColumnDef>) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    const colName = columns[colIndex].name.trim();
    const oldType = columns[colIndex].type;
    // Migrate data in-place
    applyMigration(table.rows, colName, oldType, newType);
    // Update local column state
    const newColumns = columns.map((c, i) => i === colIndex ? { ...c, ...refUpdates } : c);
    setColumns(newColumns);
    // Save schema immediately
    state.updateSchema(tableId, buildSchema(newColumns));
  };

  // Extract unique values from a column into a new reference table
  const applyExtractNow = (colIndex: number, newTableName: string) => {
    if (!tableId) return;
    const table = state.model.getTable(tableId);
    if (!table) return;
    const colName = columns[colIndex].name.trim();

    // Collect unique non-empty values
    const uniqueValues: string[] = [];
    const seen = new Set<string>();
    for (const row of table.rows) {
      const val = row[colName] ?? '';
      if (val !== '' && !seen.has(val)) {
        seen.add(val);
        uniqueValues.push(val);
      }
    }

    // Create the new table with a "value" column
    const newSchema: TableSchema = {
      name: newTableName,
      columns: [{ name: 'value', type: 'text' as ColumnType }],
      uniqueKeys: ['value'],
    };
    const newRows = uniqueValues.map(v => ({ value: v }));
    state.createTable(newSchema, newRows);

    // Build a map from value → _rowId in the new table
    const newTable = state.model.getTable(newTableName);
    if (!newTable) return;
    const valueToRowId = new Map<string, string>();
    for (const row of newTable.rows) {
      valueToRowId.set(row['value'], row[INTERNAL_ROW_ID]);
    }

    // Migrate source column: replace values with _rowId references
    for (const row of table.rows) {
      const val = row[colName] ?? '';
      if (val !== '') {
        row[colName] = valueToRowId.get(val) ?? '';
      }
    }

    // Update local column state to reference type
    const refUpdates: Partial<ColumnDef> = {
      type: 'reference' as ColumnType,
      refTable: newTableName,
      refDisplayColumns: ['value'],
      refSearchColumns: ['value'],
    };
    const newColumns = columns.map((c, i) => i === colIndex ? { ...c, ...refUpdates } : c);
    setColumns(newColumns);

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

  const handleSave = () => {
    setError(null);

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
      navigate(`/table/${encodeURIComponent(trimmedName)}`, { replace: true });
    } else {
      // Apply schema update
      state.updateSchema(tableId!, newSchema);

      // Rename table if name changed
      if (trimmedName !== tableId) {
        state.renameTable(tableId!, trimmedName);
        navigate(`/table/${encodeURIComponent(trimmedName)}`, { replace: true });
      } else {
        navigate(`/table/${encodeURIComponent(tableId!)}`);
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
      navigate('/');
    }
  };

  return (
    <div className="edit-table-page">
      <div className="edit-table-card">
        <div className="edit-table-header">
          <h2>{isCreate ? 'New Table' : 'Edit Table'}</h2>
          <button className="btn-secondary btn-sm" onClick={() => navigate(tableId ? `/table/${encodeURIComponent(tableId)}` : '/')}>
            {isCreate ? '← Back' : '← Back to Table'}
          </button>
        </div>

        {error && <div className="edit-table-error">{error}</div>}

        <div className="edit-table-field">
          <label>Table Name</label>
          <input
            type="text"
            value={tableName}
            onChange={e => setTableName(e.target.value)}
            className="edit-table-input"
          />
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
          <button className="btn-secondary btn-sm" onClick={addColumn} style={{ marginTop: 8 }}>
            + Add Column
          </button>
        </div>

        {/* Reference config dialog */}
        {refDialogCol !== null && columns[refDialogCol]?.type === 'reference' && (
          <RefConfigDialog
            col={columns[refDialogCol]}
            colIndex={refDialogCol}
            otherTableIds={otherTableIds}
            getRefTableColumns={getRefTableColumns}
            onUpdate={(idx, updates) => updateColumn(idx, updates)}
            onClose={() => setRefDialogCol(null)}
          />
        )}

        {/* Migration preview dialog */}
        {migrationPreview && (
          <MigrationPreviewDialog
            preview={migrationPreview.preview}
            onConfirm={() => {
              const { colIndex, newType, refUpdates } = migrationPreview;
              // Apply migration and save immediately
              applyMigrationNow(colIndex, newType, refUpdates!);
              setMigrationPreview(null);
              // Open ref dialog if switching to reference
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
            preview={extractPreview.preview}
            onConfirm={() => {
              const { colIndex, preview } = extractPreview;
              applyExtractNow(colIndex, preview.newTableName);
              setExtractPreview(null);
            }}
            onCancel={() => setExtractPreview(null)}
            existingTables={state.tableIds}
            onChangeTableName={(name) => setExtractPreview(prev => prev ? {
              ...prev,
              preview: { ...prev.preview, newTableName: name },
            } : null)}
          />
        )}

        <div className="edit-table-actions">
          <button className="btn-secondary" onClick={() => navigate(tableId ? `/table/${encodeURIComponent(tableId)}` : '/')}>
            Cancel
          </button>
          {!isCreate && (
            <button className="btn-danger" onClick={handleDelete}>
              Delete Table
            </button>
          )}
          <div style={{ flex: 1 }} />
          <button className="btn-primary" onClick={handleSave}>
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
  onUpdate: (index: number, updates: Partial<ColumnDef>) => void;
  onClose: () => void;
}> = ({ col, colIndex, otherTableIds, getRefTableColumns, onUpdate, onClose }) => {
  const refTableCols = col.refTable ? getRefTableColumns(col.refTable) : [];

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
                  value={(col.refDisplayColumns ?? []).map(cn => ({ value: cn, label: cn }))}
                  onChange={opts => onUpdate(colIndex, { refDisplayColumns: opts.map(o => o.value) })}
                  options={refTableCols.map(cn => ({ value: cn, label: cn }))}
                  styles={refDialogSelectStyles}
                  placeholder="Select columns to display..."
                  menuPlacement="auto"
                />
              </div>
              <div>
                <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>Search Columns</label>
                <Select
                  isMulti
                  value={(col.refSearchColumns ?? []).map(cn => ({ value: cn, label: cn }))}
                  onChange={opts => onUpdate(colIndex, { refSearchColumns: opts.map(o => o.value) })}
                  options={refTableCols.map(cn => ({ value: cn, label: cn }))}
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

// Migration preview dialog component
const MigrationPreviewDialog: React.FC<{
  preview: MigrationPreview;
  onConfirm: () => void;
  onCancel: () => void;
}> = ({ preview, onConfirm, onCancel }) => {
  const fromLabel = typeOptions.find(o => o.value === preview.fromType)?.label ?? preview.fromType;
  const toLabel = typeOptions.find(o => o.value === preview.toType)?.label ?? preview.toType;

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
          <button className="app-dialog-btn app-dialog-btn-primary" onClick={onConfirm}>
            {preview.errorCount > 0 ? `Convert (${preview.errorCount} will be cleared)` : 'Convert'}
          </button>
        </div>
      </div>
    </div>
  );
};

// Extract to reference table preview dialog
const ExtractPreviewDialog: React.FC<{
  preview: ExtractPreview;
  onConfirm: () => void;
  onCancel: () => void;
  existingTables: string[];
  onChangeTableName: (name: string) => void;
}> = ({ preview, onConfirm, onCancel, existingTables, onChangeTableName }) => {
  const nameConflict = existingTables.includes(preview.newTableName);
  const nameEmpty = !preview.newTableName.trim();

  return (
    <div className="app-dialog-overlay" onClick={onCancel}>
      <div className="app-dialog" onClick={(e) => e.stopPropagation()} style={{ maxWidth: 480 }}>
        <h3 className="app-dialog-title">
          Extract "{preview.columnName}" to Reference Table
        </h3>
        <div style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 12 }}>
          {preview.uniqueValues.length} unique value{preview.uniqueValues.length !== 1 ? 's' : ''} from {preview.nonEmptyCount} non-empty cell{preview.nonEmptyCount !== 1 ? 's' : ''}
        </div>

        <div style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 4, display: 'block' }}>New Table Name</label>
          <input
            type="text"
            value={preview.newTableName}
            onChange={e => onChangeTableName(e.target.value)}
            className="edit-table-input"
            style={{ width: '100%' }}
          />
          {nameConflict && (
            <div style={{ color: 'var(--danger, #dc2626)', fontSize: 11, marginTop: 4 }}>
              A table named "{preview.newTableName}" already exists
            </div>
          )}
        </div>

        <div style={{ maxHeight: 200, overflow: 'auto', marginBottom: 12 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
            <thead>
              <tr style={{ background: 'var(--surface-2, #f5f6f8)' }}>
                <th style={{ padding: '4px 8px', textAlign: 'left', borderBottom: '1px solid var(--border)' }}>Values to extract</th>
              </tr>
            </thead>
            <tbody>
              {preview.uniqueValues.map((v, i) => (
                <tr key={i}>
                  <td style={{ padding: '4px 8px', borderBottom: '1px solid var(--border)', fontFamily: 'monospace' }}>
                    {v}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={{ fontSize: 12, color: 'var(--text-muted)', marginBottom: 8 }}>
          This will create a new table "{preview.newTableName}" with a "value" column, and convert this column to a reference.
        </div>

        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-secondary" onClick={onCancel}>
            Cancel
          </button>
          <button
            className="app-dialog-btn app-dialog-btn-primary"
            onClick={onConfirm}
            disabled={nameConflict || nameEmpty}
          >
            Extract
          </button>
        </div>
      </div>
    </div>
  );
};

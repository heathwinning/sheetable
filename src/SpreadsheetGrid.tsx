import React, { useState, useCallback, useMemo, useRef } from 'react';
import type { TableSchema, Row, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { DataModel } from './dataModel';
import { log } from './DebugLogger';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, GetRowIdParams, ValueSetterParams, RowClassParams, SelectionChangedEvent, PostSortRowsParams, FilterChangedEvent, GetMainMenuItemsParams } from 'ag-grid-community';
import RefCellEditor from './RefCellEditor';
import DateCellEditor from './DateCellEditor';
import { ImageCellRenderer, useImageDialog } from './ImageCell';

const DRAFT_ROW_ID = '_draft';

interface SpreadsheetGridProps {
  schema: TableSchema;
  rows: Row[];
  model: DataModel;
  onEdit: (rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  onInsert: (row: Row) => ValidationError[];
  onDeleteRow: (rowIndex: number) => ValidationError[];
  revision: number;
  folderId: string | null;
}

const gridTheme = themeQuartz.withParams({
  backgroundColor: '#ffffff',
  foregroundColor: '#1e1e2e',
  headerBackgroundColor: '#f5f6f8',
  rowHoverColor: '#f0f4ff',
  selectedRowBackgroundColor: '#e0e7ff',
  borderColor: '#d4d4d8',
  cellHorizontalPaddingScale: 0.8,
  headerFontSize: 12,
  fontSize: 13,
  rowHeight: 28,
  headerHeight: 32,
  columnBorder: true,
});

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  schema,
  rows,
  model,
  onEdit,
  onInsert,
  onDeleteRow,
  revision,
  folderId,
}) => {
  const [error, setError] = useState<string | null>(null);
  const gridRef = useRef<AgGridReact>(null);
  const draftCounter = useRef(0);
  const { openDialog, dialogElement } = useImageDialog();
  const [filterActive, setFilterActive] = useState(false);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  const [bulkEditCol, setBulkEditCol] = useState('');
  const [bulkEditValue, setBulkEditValue] = useState('');

  const onSelectionChanged = useCallback((event: SelectionChangedEvent) => {
    const selected = event.api.getSelectedRows() as Row[];
    const ids = new Set(selected.map(r => r[INTERNAL_ROW_ID]).filter(id => id !== DRAFT_ROW_ID));
    setSelectedRowIds(ids);
  }, []);

  // Create a fresh draft row (local-only, not in DataModel)
  const makeDraftRow = useCallback((): Row => {
    draftCounter.current += 1;
    const row: Row = { [INTERNAL_ROW_ID]: DRAFT_ROW_ID };
    for (const col of schema.columns) {
      row[col.name] = '';
    }
    return row;
  }, [schema]);

  const draftPosition = schema.draftRowPosition ?? 'bottom';

  const rowData = useMemo(() => {
    const draft = filterActive ? null : makeDraftRow();
    if (!draft) return [...rows];
    return draftPosition === 'top' ? [draft, ...rows] : [...rows, draft];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, revision, makeDraftRow, draftPosition, filterActive]);

  // Keep draft row pinned at top or bottom regardless of sorting
  const postSortRows = useCallback((params: PostSortRowsParams) => {
    if (filterActive) return;
    const nodes = params.nodes;
    const draftIdx = nodes.findIndex(n => n.data?.[INTERNAL_ROW_ID] === DRAFT_ROW_ID);
    if (draftIdx < 0) return;
    const [draftNode] = nodes.splice(draftIdx, 1);
    if (draftPosition === 'top') {
      nodes.unshift(draftNode);
    } else {
      nodes.push(draftNode);
    }
  }, [draftPosition, filterActive]);

  const onFilterChanged = useCallback((event: FilterChangedEvent) => {
    const model = event.api.getFilterModel();
    setFilterActive(Object.keys(model).length > 0);
  }, []);

  // Map row _rowId to index in the real rows array (excluding draft)
  const rowIdToIndex = useMemo(() => {
    const map = new Map<string, number>();
    rows.forEach((row, i) => map.set(row[INTERNAL_ROW_ID], i));
    return map;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, revision]);

  const applyBulkEdit = useCallback(() => {
    if (!bulkEditCol || selectedRowIds.size === 0) return;
    let successCount = 0;
    const errors: string[] = [];
    for (const rowId of selectedRowIds) {
      const idx = rowIdToIndex.get(rowId);
      if (idx === undefined) continue;
      const errs = onEdit(idx, bulkEditCol, bulkEditValue);
      if (errs.length > 0) {
        errors.push(errs[0].message);
      } else {
        successCount++;
      }
    }
    if (errors.length > 0) {
      setError(`Bulk edit: ${successCount} updated, ${errors.length} failed — ${errors[0]}`);
    } else {
      setError(null);
    }
    setBulkEditValue('');
    gridRef.current?.api.deselectAll();
  }, [bulkEditCol, bulkEditValue, selectedRowIds, rowIdToIndex, onEdit]);

  const getRowId = useCallback((params: GetRowIdParams) => {
    return params.data[INTERNAL_ROW_ID] ?? 'fallback';
  }, []);

  // Build AG Grid column definitions with valueSetter for validation
  const columnDefs: ColDef[] = useMemo(() => {
    const cols: ColDef[] = schema.columns.map((col) => {
      const sortEntry = (schema.defaultSort ?? []).find(s => s.column === col.name);
      const sortIdx = (schema.defaultSort ?? []).findIndex(s => s.column === col.name);
      const def: ColDef = {
        field: col.name,
        headerName: col.displayName || col.name,
        editable: true,
        minWidth: 80,
        resizable: true,
        ...(sortEntry ? { sort: sortEntry.direction, sortIndex: sortIdx } : {}),
        valueSetter: (params: ValueSetterParams) => {
            const newValue = params.newValue ?? '';
            const oldValue = params.oldValue ?? '';
            log('valueSetter', col.name, 'old:', oldValue, 'new:', newValue);
            if (newValue === oldValue) return false;

            const rowId = params.data[INTERNAL_ROW_ID];

            // Draft row: insert it into the DataModel
            if (rowId === DRAFT_ROW_ID) {
              const newRow: Row = {};
              for (const c of schema.columns) {
                newRow[c.name] = c.name === col.name ? newValue : (params.data[c.name] ?? '');
              }
              // Auto-generate a placeholder key if the edited column isn't a key column
              for (const keyColName of (schema.uniqueKeys ?? [])) {
                if (!newRow[keyColName]) {
                  newRow[keyColName] = `new-${draftCounter.current}`;
                }
              }
              const errors = onInsert(newRow);
              if (errors.length > 0) {
                setError(errors[0].message);
                log('draft insert error:', errors[0].message);
                return false;
              }
              setError(null);
              log('draft promoted via', col.name, '->', newValue);
              return false; // Return false because the real data comes from the new DataModel row
            }

            // Normal row: apply edit via DataModel
            const idx = rowIdToIndex.get(rowId);
            if (idx === undefined) {
              log('valueSetter: row not found for', rowId);
              return false;
            }

            const errors = onEdit(idx, col.name, newValue);
            if (errors.length > 0) {
              setError(errors[0].message);
              log('valueSetter error:', errors[0].message);
              return false;
            }
            setError(null);
            log('valueSetter OK:', col.name, oldValue, '->', newValue);
            return true;
        },
      };

      // Reference columns: custom editor with search, display resolved values
      if (col.type === 'reference' && col.refTable) {
        const refTable = col.refTable;
        const displayCols = col.refDisplayColumns ?? [];
        const searchCols = col.refSearchColumns ?? [];

        const resolveRefDisplay = (rowId: string): string => {
          if (!rowId) return '';
          const refRow = model.getReferencedRow(refTable, rowId);
          if (!refRow) return `[missing: ${rowId}]`;
          const cols = displayCols.length > 0 ? displayCols : searchCols;
          if (cols.length === 0) return `Row ${rowId}`;
          return cols.map(c => model.resolveColumnPath(refTable, refRow, c)).filter(Boolean).join(' · ');
        };

        // Show display columns instead of raw _rowId, resolving nested references
        def.valueFormatter = (params) => resolveRefDisplay(params.value);

        // Filter on resolved display text, not raw _rowId
        def.filterValueGetter = (params) => resolveRefDisplay(params.data?.[col.name] ?? '');

        def.cellEditorSelector = () => {
          return {
            component: RefCellEditor,
            popup: true,
            popupPosition: 'under',
            params: {
              refRows: model.getReferenceRows(refTable),
              refTable,
              model,
              searchColumns: searchCols,
              displayColumns: displayCols,
            },
          };
        };
      }

      // Set filter type based on column type
      if (col.type === 'integer' || col.type === 'decimal') {
        def.filter = 'agNumberColumnFilter';
      } else if (col.type === 'date' || col.type === 'datetime') {
        def.filter = 'agDateColumnFilter';
      } else if (col.type !== 'image') {
        def.filter = 'agTextColumnFilter';
      }

      // Bool columns: use a select dropdown
      if (col.type === 'bool') {
        def.cellEditor = 'agSelectCellEditor';
        def.cellEditorParams = {
          values: ['', 'true', 'false'],
        };
      }

      // Date columns: react-datepicker popup
      if (col.type === 'date' || col.type === 'datetime') {
        def.cellEditor = DateCellEditor;
        def.cellEditorPopup = true;
        def.cellEditorPopupPosition = 'under';
      }

      // Image columns: thumbnail renderer, click to open dialog
      if (col.type === 'image') {
        def.cellRenderer = ImageCellRenderer;
        def.editable = false;
        def.onCellClicked = (params) => {
          if (!folderId) {
            setError('Connect to Google Drive to upload images');
            return;
          }
          const rowId = params.data[INTERNAL_ROW_ID];
          if (rowId === DRAFT_ROW_ID) return;

          const currentFileId = params.value || null;
          openDialog(currentFileId, folderId, schema.name, (newFileId) => {
            const idx = rowIdToIndex.get(rowId);
            if (idx !== undefined) {
              const errors = onEdit(idx, col.name, newFileId ?? '');
              if (errors.length > 0) {
                setError(errors[0].message);
              }
            }
          });
        };
      }

      return def;
    });

    // Delete button column
    cols.push({
      headerName: '',
      width: 50,
      maxWidth: 50,
      editable: false,
      sortable: false,
      filter: false,
      cellRenderer: (params: { data: Row }) => {
        if (params.data[INTERNAL_ROW_ID] === DRAFT_ROW_ID) return '';
        return '×';
      },
      cellStyle: (params): Record<string, string> => {
        if (params.data[INTERNAL_ROW_ID] === DRAFT_ROW_ID) return { cursor: 'default' };
        return { cursor: 'pointer', textAlign: 'center', color: '#888' };
      },
      onCellClicked: (params) => {
        if (params.data[INTERNAL_ROW_ID] === DRAFT_ROW_ID) return;
        const idx = rowIdToIndex.get(params.data[INTERNAL_ROW_ID]);
        if (idx !== undefined) {
          const errors = onDeleteRow(idx);
          if (errors.length > 0) {
            setError(errors[0].message);
          }
        }
      },
    });

    return cols;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [schema, model, revision, rowIdToIndex, onDeleteRow, onEdit, onInsert, folderId]);

  // Style draft row with dimmer text
  const getRowClass = useCallback((params: RowClassParams) => {
    if (params.data?.[INTERNAL_ROW_ID] === DRAFT_ROW_ID) {
      return 'draft-row';
    }
    return undefined;
  }, []);

  // Clear error after timeout
  React.useEffect(() => {
    if (error) {
      const timer = setTimeout(() => setError(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [error]);

  const autoSizeStrategy = useMemo(() => ({
    type: 'fitCellContents' as const,
  }), []);

  const isRowSelectable = useCallback((params: { data?: Row }) => {
    return params.data?.[INTERNAL_ROW_ID] !== DRAFT_ROW_ID;
  }, []);

  const rowSelectionConfig = useMemo(() => ({
    mode: 'multiRow' as const,
    checkboxes: true,
    headerCheckbox: true,
    selectAll: 'filtered' as const,
    isRowSelectable,
  }), [isRowSelectable]);

  // Editable column options for bulk edit dropdown
  const editableColumnOptions = useMemo(() =>
    schema.columns
      .filter(c => c.type !== 'image')
      .map(c => ({ value: c.name, label: c.displayName || c.name })),
    [schema],
  );

  // Normalize: trim whitespace for a column
  const trimColumn = useCallback((colName: string) => {
    let trimmed = 0;
    for (const row of rows) {
      const val = row[colName];
      if (typeof val === 'string' && val !== val.trim()) {
        const idx = rowIdToIndex.get(row[INTERNAL_ROW_ID]);
        if (idx !== undefined) {
          const errs = onEdit(idx, colName, val.trim());
          if (errs.length === 0) trimmed++;
        }
      }
    }
    if (trimmed > 0) {
      setError(`Trimmed whitespace in ${trimmed} cell${trimmed > 1 ? 's' : ''}`);
    } else {
      setError('No whitespace to trim');
    }
  }, [rows, rowIdToIndex, onEdit]);

  const getMainMenuItems = useCallback((params: GetMainMenuItemsParams) => {
    const defaults = params.defaultItems ?? [];
    const colId = params.column?.getColId();
    const col = schema.columns.find(c => c.name === colId);
    if (!col || col.type === 'image' || col.type === 'reference') return defaults;
    return [
      ...defaults,
      'separator' as const,
      {
        name: 'Trim whitespace',
        action: () => trimColumn(col.name),
      },
    ];
  }, [schema, trimColumn]);

  // For visible trim whitespace UI
  const [trimCol, setTrimCol] = useState('');
  const [showTrimDialog, setShowTrimDialog] = useState(false);
  const trimOptions = useMemo(
    () => schema.columns
      .filter(c => c.type !== 'image' && c.type !== 'reference')
      .map(c => ({ value: c.name, label: c.displayName || c.name })),
    [schema]
  );

  return (
    <div className="spreadsheet-container" style={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '8px 8px 0' }}>
        <button
          className="btn-secondary btn-sm"
          onClick={() => setShowTrimDialog(true)}
          disabled={trimOptions.length === 0}
        >
          Normalize Columns
        </button>
      </div>

      {error && (
        <div className="error-banner">
          {error}
        </div>
      )}

      {showTrimDialog && (
        <div className="app-dialog-overlay" onClick={() => setShowTrimDialog(false)}>
          <div className="app-dialog" onClick={(e) => e.stopPropagation()}>
            <h3 className="app-dialog-title">Trim Whitespace</h3>
            <p className="app-dialog-message">Select a column to normalize by trimming leading/trailing whitespace.</p>
            <select
              className="bulk-edit-select"
              value={trimCol}
              onChange={(e) => setTrimCol(e.target.value)}
              style={{ width: '100%', marginBottom: 12 }}
            >
              <option value="">Column...</option>
              {trimOptions.map(o => (
                <option key={o.value} value={o.value}>{o.label}</option>
              ))}
            </select>
            <div className="app-dialog-actions">
              <button
                className="app-dialog-btn app-dialog-btn-secondary"
                onClick={() => setShowTrimDialog(false)}
              >
                Cancel
              </button>
              <button
                className="app-dialog-btn app-dialog-btn-primary"
                disabled={!trimCol}
                onClick={() => {
                  trimColumn(trimCol);
                  setShowTrimDialog(false);
                }}
              >
                Trim Whitespace
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedRowIds.size > 0 && (
        <div className="bulk-edit-bar">
          <span className="bulk-edit-count">{selectedRowIds.size} row{selectedRowIds.size > 1 ? 's' : ''} selected</span>
          <select
            className="bulk-edit-select"
            value={bulkEditCol}
            onChange={(e) => setBulkEditCol(e.target.value)}
          >
            <option value="">Column...</option>
            {editableColumnOptions.map(o => (
              <option key={o.value} value={o.value}>{o.label}</option>
            ))}
          </select>
          <input
            className="bulk-edit-input"
            type="text"
            placeholder="New value"
            value={bulkEditValue}
            onChange={(e) => setBulkEditValue(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter') applyBulkEdit(); }}
            disabled={!bulkEditCol}
          />
          <button className="btn-primary btn-sm" onClick={applyBulkEdit} disabled={!bulkEditCol}>
            Apply
          </button>
          <button className="btn-secondary btn-sm" onClick={() => gridRef.current?.api.deselectAll()}>
            Clear
          </button>
        </div>
      )}
      {dialogElement}
      <div style={{ flex: 1, minHeight: 0 }}>
        <AgGridReact
          ref={gridRef}
          modules={[AllCommunityModule]}
          theme={gridTheme}
          rowData={rowData}
          columnDefs={columnDefs}
          getRowId={getRowId}
          getRowClass={getRowClass}
          autoSizeStrategy={autoSizeStrategy}
          singleClickEdit={true}
          stopEditingWhenCellsLoseFocus={true}
          enterNavigatesVertically={true}
          enterNavigatesVerticallyAfterEdit={true}
          suppressNoRowsOverlay={true}
          postSortRows={postSortRows}
          rowSelection={rowSelectionConfig}
          onSelectionChanged={onSelectionChanged}
          onFilterChanged={onFilterChanged}
          getMainMenuItems={getMainMenuItems}
        />
      </div>
    </div>
  );
};

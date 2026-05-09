import React, { useState, useCallback, useMemo, useRef } from 'react';
import type { TableSchema, Row, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { log } from './DebugLogger';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, GetRowIdParams, ValueSetterParams, RowClassParams, SelectionChangedEvent, PostSortRowsParams, FilterChangedEvent, ColumnMovedEvent, ColumnResizedEvent, FirstDataRenderedEvent } from 'ag-grid-community';
import RefCellEditor from './RefCellEditor';
import DateCellEditor from './DateCellEditor';
import { ImageCellRenderer, useImageDialog } from './ImageCell';
import { normalizeTemporalString, parseTemporalUnknown } from './dateFormat';

const DRAFT_ROW_ID = '_draft';

function toSortableDateEpoch(value: unknown): number {
  const parsed = parseTemporalUnknown(value);
  if (!parsed) return Number.POSITIVE_INFINITY;
  const ts = parsed.getTime();
  return Number.isNaN(ts) ? Number.POSITIVE_INFINITY : ts;
}

interface SpreadsheetGridProps {
  schema: TableSchema;
  rows: Row[];
  onEdit: (rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  onInsert: (row: Row) => ValidationError[];
  onDeleteRow: (rowIndex: number) => ValidationError[];
  onColumnOrderChange?: (orderedColumnNames: string[]) => void;
  onColumnWidthChange?: (columnWidths: Record<string, number>) => void;
  revision: number;
  bookId: string | null;
  readOnly?: boolean;
  // Reference helpers (from useAppState)
  getReferencedRow: (refTable: string, rowId: string) => Row | undefined;
  getReferenceRows: (refTable: string) => Row[];
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  resolveColumnPathLabel: (tableName: string, path: string) => string;
}

const gridTheme = themeQuartz.withParams({
  cellHorizontalPaddingScale: 0.5,
  headerFontSize: 12,
  fontSize: 13,
  rowHeight: 26,
  headerHeight: 28,
  columnBorder: true,
});

export const SpreadsheetGrid: React.FC<SpreadsheetGridProps> = ({
  schema,
  rows,
  onEdit,
  onInsert,
  onDeleteRow,
  onColumnOrderChange,
  onColumnWidthChange,
  revision,
  bookId,
  readOnly,
  getReferencedRow,
  getReferenceRows,
  resolveColumnPath,
  resolveColumnPathLabel,
}) => {
  const [error, setError] = useState<string | null>(null);
  const gridRef = useRef<AgGridReact>(null);
  const draftCounter = useRef(0);
  const { openDialog, dialogElement } = useImageDialog();
  const [filterActive, setFilterActive] = useState(false);
  const [displayedRowCount, setDisplayedRowCount] = useState<number | null>(null);
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());
  const [bulkEditCol, setBulkEditCol] = useState('');
  const [bulkEditValue, setBulkEditValue] = useState('');

  const onSelectionChanged = useCallback((event: SelectionChangedEvent) => {
    const selected = event.api.getSelectedRows() as Row[];
    const ids = new Set(selected.map(r => r[INTERNAL_ROW_ID]).filter(id => id !== DRAFT_ROW_ID));
    setSelectedRowIds(ids);
  }, []);

  // Create a fresh draft row (local-only, not yet persisted)
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
    if (readOnly) return [...rows];
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
    const isFiltered = Object.keys(model).length > 0;
    setFilterActive(isFiltered);
    setDisplayedRowCount(isFiltered ? event.api.getDisplayedRowCount() : null);
  }, []);

  const onFirstDataRendered = useCallback((event: FirstDataRenderedEvent) => {
    // If draft row is configured at bottom, open the table with viewport at the bottom.
    if (draftPosition !== 'bottom') return;
    const displayedCount = event.api.getDisplayedRowCount();
    if (displayedCount > 0) {
      event.api.ensureIndexVisible(displayedCount - 1, 'bottom');
    }
  }, [draftPosition]);

  const clearAllFilters = useCallback(() => {
    gridRef.current?.api.setFilterModel(null);
    setFilterActive(false);
  }, []);

  const onColumnMoved = useCallback((event: ColumnMovedEvent) => {
    if (!event.finished || !onColumnOrderChange) return;
    const schemaColNames = new Set(schema.columns.map(c => c.name));
    const ordered = (event.api.getAllGridColumns() ?? [])
      .map(col => col.getColId())
      .filter(id => schemaColNames.has(id));
    if (ordered.length === schema.columns.length) {
      onColumnOrderChange(ordered);
    }
  }, [onColumnOrderChange, schema.columns]);

  const onColumnResized = useCallback((event: ColumnResizedEvent) => {
    if (!event.finished || !onColumnWidthChange || event.source === 'api') return;
    const schemaColNames = new Set(schema.columns.map(c => c.name));
    const widths: Record<string, number> = {};
    for (const col of event.columns ?? []) {
      const id = col.getColId();
      if (schemaColNames.has(id)) {
        widths[id] = col.getActualWidth();
      }
    }
    if (Object.keys(widths).length > 0) {
      onColumnWidthChange(widths);
    }
  }, [onColumnWidthChange, schema.columns]);

  // Map row _rowId to index in the real rows array (excluding draft)
  const rowIdToIndex = useMemo(() => {
    const map = new Map<string, number>();
    rows.forEach((row, i) => map.set(row[INTERNAL_ROW_ID], i));
    return map;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows, revision]);

  const applyBulkEdit = useCallback(() => {
    if (!bulkEditCol || selectedRowIds.size === 0) return;
    const bulkColType = schema.columns.find(c => c.name === bulkEditCol)?.type;
    const normalizedBulkValue = bulkColType
      ? normalizeTemporalString(bulkEditValue, bulkColType)
      : bulkEditValue;
    let successCount = 0;
    const errors: string[] = [];
    for (const rowId of selectedRowIds) {
      const idx = rowIdToIndex.get(rowId);
      if (idx === undefined) continue;
      const errs = onEdit(idx, bulkEditCol, normalizedBulkValue);
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
  }, [bulkEditCol, bulkEditValue, schema.columns, selectedRowIds, rowIdToIndex, onEdit]);

  const deleteSelectedRows = useCallback(() => {
    if (selectedRowIds.size === 0) return;
    const indices = Array.from(selectedRowIds)
      .map(id => rowIdToIndex.get(id))
      .filter((idx): idx is number => idx !== undefined)
      .sort((a, b) => b - a);

    let deleted = 0;
    const errors: string[] = [];
    for (const idx of indices) {
      const errs = onDeleteRow(idx);
      if (errs.length > 0) {
        errors.push(errs[0].message);
      } else {
        deleted++;
      }
    }

    if (errors.length > 0) {
      setError(`Delete: ${deleted} removed, ${errors.length} failed — ${errors[0]}`);
    } else {
      setError(null);
    }
    gridRef.current?.api.deselectAll();
  }, [selectedRowIds, rowIdToIndex, onDeleteRow]);

  const getRowId = useCallback((params: GetRowIdParams) => {
    return params.data[INTERNAL_ROW_ID] ?? 'fallback';
  }, []);

  // Build AG Grid column definitions with valueSetter for validation
  const columnDefs: ColDef[] = useMemo(() => {
    const cols: ColDef[] = schema.columns.flatMap((col) => {
      const sortEntry = (schema.defaultSort ?? []).find(s => s.column === col.name);
      const sortIdx = (schema.defaultSort ?? []).findIndex(s => s.column === col.name);
      const def: ColDef = {
        field: col.name,
        headerName: col.displayName || col.name,
        editable: !readOnly,
        minWidth: 80,
        ...(col.width ? { width: col.width } : {}),
        resizable: true,
        ...(sortEntry ? { sort: sortEntry.direction, sortIndex: sortIdx } : {}),
        valueSetter: (params: ValueSetterParams) => {
          const rawValue = String(params.newValue ?? '');
          const newValue = normalizeTemporalString(rawValue, col.type);
            const oldValue = params.oldValue ?? '';
            log('valueSetter', col.name, 'old:', oldValue, 'new:', newValue);
            if (newValue === oldValue) return false;

            const rowId = params.data[INTERNAL_ROW_ID];

            // Draft row: insert it via API
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
              return false; // Return false because the real data comes from the newly inserted row
            }

            // Normal row: apply edit via API
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

        const setReferenceValue = (params: ValueSetterParams, newValueRaw: string) => {
          const newValue = newValueRaw ?? '';
          const oldValue = params.data?.[col.name] ?? '';
          if (newValue === oldValue) return false;

          const rowId = params.data?.[INTERNAL_ROW_ID];

          // Draft row: insert it via API
          if (rowId === DRAFT_ROW_ID) {
            const newRow: Row = {};
            for (const c of schema.columns) {
              if (c.name === col.name) {
                newRow[c.name] = newValue;
              } else {
                newRow[c.name] = params.data?.[c.name] ?? '';
              }
            }
            for (const keyColName of (schema.uniqueKeys ?? [])) {
              if (!newRow[keyColName]) {
                newRow[keyColName] = `new-${draftCounter.current}`;
              }
            }
            const errors = onInsert(newRow);
            if (errors.length > 0) {
              setError(errors[0].message);
              return false;
            }
            setError(null);
            return false;
          }

          const idx = rowIdToIndex.get(rowId);
          if (idx === undefined) return false;
          const errors = onEdit(idx, col.name, newValue);
          if (errors.length > 0) {
            setError(errors[0].message);
            return false;
          }
          setError(null);
          return true;
        };

        const resolveRefDisplay = (rowId: string): string => {
          if (!rowId) return '';
          const refRow = getReferencedRow(refTable, rowId);
          if (!refRow) return `[missing: ${rowId}]`;
          const cols = displayCols.length > 0 ? displayCols : searchCols;
          if (cols.length === 0) return `Row ${rowId}`;
          return cols.map(c => resolveColumnPath(refTable, refRow, c)).filter(Boolean).join(' · ');
        };

        // Show display columns instead of raw _rowId, resolving nested references
        def.valueFormatter = (params) => resolveRefDisplay(params.value);

        // Sort by resolved display text (what users see), not raw hidden _rowId.
        def.comparator = (a, b) => {
          const left = resolveRefDisplay(a ?? '').toLocaleLowerCase();
          const right = resolveRefDisplay(b ?? '').toLocaleLowerCase();
          return left.localeCompare(right);
        };

        // Filter on resolved display text, not raw _rowId
        def.filterValueGetter = (params) => resolveRefDisplay(params.data?.[col.name] ?? '');

        def.cellEditorSelector = () => {
          return {
            component: RefCellEditor,
            popup: true,
            popupPosition: 'under',
            params: {
              refRows: getReferenceRows(refTable),
              refTable,
              resolveColumnPath,
              searchColumns: searchCols,
              displayColumns: displayCols,
            },
          };
        };

        def.valueSetter = (params: ValueSetterParams) => setReferenceValue(params, params.newValue ?? '');

        const derivedDefs: ColDef[] = displayCols.map((displayPath) => ({
          colId: `${col.name}::${displayPath}`,
          headerName: `${col.displayName || col.name} → ${resolveColumnPathLabel(refTable, displayPath)}`,
          editable: true,
          minWidth: 120,
          resizable: true,
          filter: 'agTextColumnFilter',
          headerClass: 'reference-derived-header',
          cellClass: 'reference-derived-cell',
          valueGetter: (params) => {
            const rowId = params.data?.[col.name] ?? '';
            if (!rowId) return '';
            const refRow = getReferencedRow(refTable, rowId);
            if (!refRow) return '';
            return resolveColumnPath(refTable, refRow, displayPath);
          },
          comparator: (a, b) => String(a ?? '').toLowerCase().localeCompare(String(b ?? '').toLowerCase()),
          cellEditorSelector: def.cellEditorSelector,
          valueSetter: (params: ValueSetterParams) => setReferenceValue(params, params.newValue ?? ''),
        }));

        if (derivedDefs.length > 0) {
          const hiddenBackingDef: ColDef = {
            ...def,
            hide: true,
            lockVisible: true,
            suppressColumnsToolPanel: true,
          };
          return [hiddenBackingDef, ...derivedDefs];
        }
        return [def];
      }

      // Set filter type based on column type
      if (col.type === 'integer' || col.type === 'decimal') {
        def.filter = 'agNumberColumnFilter';
        def.comparator = (a, b) => {
          const na = a == null || a === '' ? -Infinity : Number(a);
          const nb = b == null || b === '' ? -Infinity : Number(b);
          return na - nb;
        };
      } else if (col.type === 'date' || col.type === 'datetime') {
        def.filter = 'agDateColumnFilter';
        def.comparator = (a, b) => toSortableDateEpoch(a) - toSortableDateEpoch(b);
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

      // Image columns: icon indicator in-cell, click to open upload/preview dialog
      if (col.type === 'image') {
        def.cellRenderer = ImageCellRenderer;
        def.editable = false;
        def.onCellClicked = (params) => {
          if (!bookId) {
            setError('Sign in and open a book to upload images');
            return;
          }
          const rowId = params.data[INTERNAL_ROW_ID];
          if (rowId === DRAFT_ROW_ID) return;

          const currentKey = params.value || null;
          openDialog(currentKey, bookId, schema.name, (newKey) => {
            const idx = rowIdToIndex.get(rowId);
            if (idx !== undefined) {
              const errors = onEdit(idx, col.name, newKey ?? '');
              if (errors.length > 0) {
                setError(errors[0].message);
              }
            }
          });
        };
      }

      return [def];
    });

    return cols;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [schema, revision, rowIdToIndex, onEdit, onInsert, bookId, getReferencedRow, getReferenceRows, resolveColumnPath, resolveColumnPathLabel]);

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

  const popupParent = useMemo(() => {
    return typeof document !== 'undefined' ? document.body : undefined;
  }, []);

  // Editable column options for bulk edit dropdown
  const editableColumnOptions = useMemo(() =>
    schema.columns
      .filter(c => c.type !== 'image')
      .map(c => ({ value: c.name, label: c.displayName || c.name })),
    [schema],
  );

  const ZOOM_KEY = 'sheetable-grid-zoom';
  const [zoom, setZoom] = useState(() => {
    const stored = localStorage.getItem(ZOOM_KEY);
    return stored ? Math.max(0.25, Math.min(2, Number(stored) || 1)) : 1;
  });
  const zoomRef = useRef(zoom);
  React.useEffect(() => { zoomRef.current = zoom; }, [zoom]);
  const changeZoom = useCallback((next: number) => {
    const clamped = Math.max(0.25, Math.min(2, Math.round(next * 100) / 100));
    setZoom(clamped);
    localStorage.setItem(ZOOM_KEY, String(clamped));
  }, []);

  // Pinch-to-zoom via Pointer Events
  const gridWrapperRef = useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    const el = gridWrapperRef.current;
    if (!el) return;
    const pointers = new Map<number, { x: number; y: number }>();
    let startDist = 0;
    let startZoom = 1;
    const dist = () => {
      const [a, b] = Array.from(pointers.values());
      return Math.hypot(a.x - b.x, a.y - b.y);
    };
    const onDown = (e: PointerEvent) => {
      pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pointers.size === 2) {
        startDist = dist();
        startZoom = zoomRef.current;
      }
    };
    const onMove = (e: PointerEvent) => {
      if (!pointers.has(e.pointerId)) return;
      pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
      if (pointers.size === 2 && startDist > 0) {
        changeZoom(startZoom * (dist() / startDist));
      }
    };
    const onUp = (e: PointerEvent) => {
      pointers.delete(e.pointerId);
      if (pointers.size < 2) startDist = 0;
    };
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      // deltaY is negative when scrolling up (zoom in), positive when down (zoom out)
      const delta = -e.deltaY * (e.deltaMode === 1 ? 0.05 : 0.001);
      changeZoom(zoomRef.current * (1 + delta));
    };
    el.addEventListener('pointerdown', onDown);
    el.addEventListener('pointermove', onMove);
    el.addEventListener('pointerup', onUp);
    el.addEventListener('pointercancel', onUp);
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => {
      el.removeEventListener('pointerdown', onDown);
      el.removeEventListener('pointermove', onMove);
      el.removeEventListener('pointerup', onUp);
      el.removeEventListener('pointercancel', onUp);
      el.removeEventListener('wheel', onWheel);
    };
  }, [changeZoom]);

  return (
    <div className="spreadsheet-container" style={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column' }}>
      {error && (
        <div className="error-banner">
          {error}
        </div>
      )}

      {filterActive && (
        <div className="bulk-edit-bar" style={{ marginBottom: 8 }}>
          <span className="bulk-edit-count">Filters are active</span>
          <button className="btn-secondary btn-sm" onClick={clearAllFilters}>
            Clear Filters
          </button>
        </div>
      )}

      {selectedRowIds.size > 0 && !readOnly && (
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
          <button className="btn-danger btn-sm" onClick={deleteSelectedRows}>
            Delete Selected
          </button>
          <button className="btn-secondary btn-sm" onClick={() => gridRef.current?.api.deselectAll()}>
            Clear
          </button>
        </div>
      )}
      {dialogElement}
      <div className="grid-status-bar">
        <span className="grid-row-count">
          {rows.length} row{rows.length !== 1 ? 's' : ''}{displayedRowCount !== null ? ` (${displayedRowCount} shown)` : ''}
        </span>
        <span className="grid-status-spacer" />
      </div>
      <div className="grid-wrapper" ref={gridWrapperRef} style={{ flex: 1, minHeight: 0, zoom, touchAction: 'pan-x pan-y' }}>
        <AgGridReact
          ref={gridRef}
          modules={[AllCommunityModule]}
          theme={gridTheme}
          popupParent={popupParent}
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
          onFirstDataRendered={onFirstDataRendered}
          onColumnMoved={onColumnMoved}
          onColumnResized={onColumnResized}
        />
      </div>
    </div>
  );
};

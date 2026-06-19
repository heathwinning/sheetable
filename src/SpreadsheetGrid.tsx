import React, { useState, useCallback, useMemo, useRef, useEffect } from 'react';
import type { TableSchema, Row, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { log } from './DebugLogger';
import { useConfirm } from './DialogProvider';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, themeQuartz } from 'ag-grid-community';
import type { ColDef, GetRowIdParams, ValueSetterParams, RowClassParams, SelectionChangedEvent, PostSortRowsParams, FilterChangedEvent, ColumnResizedEvent, FirstDataRenderedEvent, CellEditingStartedEvent, CellEditingStoppedEvent, DisplayedColumnsChangedEvent } from 'ag-grid-community';
import { SelectionSumBar } from './SelectionSumBar';
import type { SelectionStats } from './SelectionSumBar';
import type { CustomCellRendererProps } from 'ag-grid-react';
import RefCellEditor from './RefCellEditor';
import DateCellEditor from './DateCellEditor';
import ListTagsEditor from './ListTagsEditor';
import { ImageCellRenderer, useImageDialog } from './ImageCell';
import { normalizeTemporalString, parseTemporalUnknown, formatDateCanonical, formatDateTimeCanonical } from './dateFormat';
import { getCalc } from './chartFormat';
import { sharedDefaultColDef } from './gridDefaults';

const DRAFT_ROW_ID = '_draft';

function parseListItems(value: string): string[] {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed.map(String);
  } catch {
    // fall through
  }
  return value ? [value] : [];
}

const ListTagsRenderer: React.FC<CustomCellRendererProps> = ({ value }) => {
  const items = parseListItems(value ?? '');
  if (items.length === 0) return null;
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 3, alignItems: 'center', paddingTop: 2 }}>
      {items.map((item, i) => (
        <span
          key={i}
          style={{
            background: 'var(--ag-row-hover-color, #e8f0fe)',
            color: 'var(--ag-foreground-color, #333)',
            borderRadius: 3,
            padding: '0 6px',
            fontSize: 11,
            lineHeight: '17px',
            display: 'inline-block',
          }}
        >
          {item}
        </span>
      ))}
    </div>
  );
};

const OpenRecordButton: React.FC<CustomCellRendererProps & { onOpen: (row: Row) => void }> = ({ data, onOpen }) => {
  if (!data || data[INTERNAL_ROW_ID] === DRAFT_ROW_ID) return null;
  return (
    <button
      title="Open record"
      aria-label="Open record"
      onClick={(e) => { e.stopPropagation(); onOpen(data as Row); }}
      style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 16, padding: 0, lineHeight: 1, color: 'var(--color-text-muted,#888)', width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
    >
      ⤢
    </button>
  );
};

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
  onDeleteRows: (rowIndices: number[]) => ValidationError[];
  onColumnWidthChange?: (columnWidths: Record<string, number>) => void;
  /** Called when the user clicks the open-record button on a row. */
  onOpenRecord?: (row: Row) => void;
  revision: number;
  bookId: string | null;
  readOnly?: boolean;
  // Reference helpers (from useAppState)
  getReferencedRow: (refTable: string, rowId: string) => Row | undefined;
  getReferenceRows: (refTable: string) => Row[];
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  resolveColumnPathLabel: (tableName: string, path: string) => string;
  resolveColumnPathLeafLabel: (tableName: string, path: string) => string;
  onCreateReferenceRow?: (refTable: string, seedText: string) => Promise<string | null>;
  /** When true, the grid will not stop editing when cells lose focus (use while a create-reference modal is open) */
  keepEditorAlive?: boolean;
  /** Called whenever the cell selection changes with numeric stats, or null if nothing numeric is selected. */
  onCellSelectionStats?: (stats: SelectionStats | null) => void;
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
  onDeleteRows,
  onColumnWidthChange,
  onOpenRecord,
  revision,
  bookId,
  readOnly,
  getReferencedRow,
  getReferenceRows,
  resolveColumnPath,
  resolveColumnPathLabel,
  resolveColumnPathLeafLabel,
  onCreateReferenceRow,
  keepEditorAlive,
  onCellSelectionStats,
}) => {
  const [error, setError] = useState<string | null>(null);
  const gridRef = useRef<AgGridReact>(null);
  const gridWrapperRef = useRef<HTMLDivElement>(null);
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

  const [cellStats, setCellStats] = useState<SelectionStats | null>(null);

  // ── Select mode ────────────────────────────────────────────────────────────
  const [selectMode, setSelectMode] = useState(false);
  const selectModeRef = useRef(false);
  useEffect(() => { selectModeRef.current = selectMode; }, [selectMode]);

  // Current selection range stored as { minRow, maxRow, minColIdx, maxColIdx }
  const selectionRef = useRef<{ minRow: number; maxRow: number; minColIdx: number; maxColIdx: number } | null>(null);
  // Ordered list of column IDs (kept in sync with the grid)
  const colOrderRef = useRef<string[]>([]);

  const updateColOrder = useCallback(() => {
    const api = gridRef.current?.api;
    if (!api) return;
    const cols = api.getColumns();
    colOrderRef.current = cols ? cols.map(c => c.getColId()) : [];
  }, []);

  const clearCellSelection = useCallback(() => {
    selectionRef.current = null;
    gridRef.current?.api.refreshCells({ force: true });
    setCellStats(null);
    onCellSelectionStats?.(null);
  }, [onCellSelectionStats]);

  // ── Mobile keyboard avoidance helpers ────────────────────────────────────
  const editingCellPos = useRef<{ rowIndex: number; colId: string } | null>(null);
  const popupObserver = useRef<MutationObserver | null>(null);

  // Highlight / un-highlight the cell being edited (AG Grid doesn't highlight
  // cells that use popup editors).
  const setCellHighlight = useCallback((rowIndex: number | null, colId: string | null, on: boolean) => {
    if (rowIndex === null || !colId) return;
    const gridEl = gridWrapperRef.current;
    if (!gridEl) return;
    const cell = gridEl.querySelector(
      `.ag-row[row-index="${rowIndex}"] .ag-cell[col-id="${CSS.escape(colId)}"]`,
    ) as HTMLElement | null;
    if (cell) {
      cell.classList.toggle('ag-cell-popup-editing', on);
    }
  }, []);

  // Watch for popups and portal dropdowns to appear, then scroll them into view.
  const startPopupObserver = useCallback(() => {
    popupObserver.current?.disconnect();
    const obs = new MutationObserver(() => {
      const popup = document.querySelector('.ag-popup:not(.ag-hidden), .ref-editor-dropdown, .date-cell-popover');
      if (popup) {
        obs.disconnect();
        requestAnimationFrame(() => {
          const vv = window.visualViewport;
          const vh = vv ? vv.height : window.innerHeight;
          const bodyViewport = gridWrapperRef.current?.querySelector('.ag-body-viewport') as HTMLElement | null;
          if (!bodyViewport) return;
          const popupBottom = popup.getBoundingClientRect().bottom;
          if (popupBottom > vh - 8) {
            bodyViewport.scrollTop += popupBottom - vh + 24;
          }
        });
      }
    });
    obs.observe(document.body, { childList: true, subtree: true });
    popupObserver.current = obs;
  }, []);

  const scrollEditingCellIntoView = useCallback(() => {
    const gridEl = gridWrapperRef.current;
    if (!gridEl) return;
    const bodyViewport = gridEl.querySelector('.ag-body-viewport') as HTMLElement | null;
    if (!bodyViewport) return;
    const vv = window.visualViewport;
    const visibleHeight = vv ? vv.height : window.innerHeight;

    // Position the editing cell at ~30% from the top of the visible area.
    const targetTop = visibleHeight * 0.3;

    // ── Inline editors (text, number) ──────────────────────────────────────
    const inlineCell = gridEl.querySelector('.ag-cell-inline-editing') as HTMLElement | null;
    if (inlineCell) {
      const cellTop = inlineCell.getBoundingClientRect().top;
      if (cellTop > targetTop) {
        bodyViewport.scrollTop += cellTop - targetTop;
      }
      return;
    }

    // ── Popup editors (date, ref, list, bool) ──────────────────────────────
    const pos = editingCellPos.current;
    const api = gridRef.current?.api;
    if (!pos || !api) return;

    // Find the source cell. Visible rows are in the DOM; virtualised rows
    // need ensureIndexVisible to force-render them first.
    const cellSel = `.ag-row[row-index="${pos.rowIndex}"] .ag-cell[col-id="${CSS.escape(pos.colId)}"]`;
    const cell = gridEl.querySelector(cellSel) as HTMLElement | null;

    const scrollCellToTarget = (c: HTMLElement) => {
      const cellTop = c.getBoundingClientRect().top;
      if (cellTop > targetTop) {
        bodyViewport.scrollTop += cellTop - targetTop;
      }
    };

    if (cell) {
      scrollCellToTarget(cell);
    } else {
      api.ensureIndexVisible(pos.rowIndex, 'top');
      requestAnimationFrame(() => {
        const c = gridEl.querySelector(cellSel) as HTMLElement | null;
        if (c) scrollCellToTarget(c);
      });
    }

    // After the popup renders, check it isn't behind the keyboard.
    setTimeout(() => {
      const popup = document.querySelector('.ag-popup:not(.ag-hidden)') as HTMLElement | null;
      if (popup) {
        const popupBottom = popup.getBoundingClientRect().bottom;
        const vh = window.visualViewport?.height ?? window.innerHeight;
        if (popupBottom > vh - 8) {
          bodyViewport.scrollTop += popupBottom - vh + 24;
        }
      }
    }, 300);
  }, []);

  const onCellEditingStarted = useCallback((event: CellEditingStartedEvent) => {
    if (selectModeRef.current) {
      event.api.stopEditing(true);
      return;
    }
    const colId = event.column.getColId();
    const rowIndex = event.node.rowIndex;
    if (rowIndex !== null) {
      editingCellPos.current = { rowIndex, colId };
      // Highlight the cell (AG Grid only highlights inline editors by default).
      setCellHighlight(rowIndex, colId, true);
      // Watch for the popup to appear so we can scroll it above the keyboard.
      startPopupObserver();
    }
    // Delayed checks: mobile keyboards animate in over 200-600ms.
    // We retry at multiple intervals so the last check runs with the final
    // reduced viewport height.
    [200, 450, 750].forEach(delay => {
      setTimeout(() => scrollEditingCellIntoView(), delay);
    });
  }, [scrollEditingCellIntoView, setCellHighlight, startPopupObserver]);

  const onCellEditingStopped = useCallback((event: CellEditingStoppedEvent) => {
    const colId = event.column.getColId();
    const rowIndex = event.node.rowIndex;
    setCellHighlight(rowIndex ?? null, colId, false);
    popupObserver.current?.disconnect();
    popupObserver.current = null;
    editingCellPos.current = null;
  }, [setCellHighlight]);

  const onDisplayedColumnsChanged = useCallback((_event: DisplayedColumnsChangedEvent) => {
    updateColOrder();
  }, [updateColOrder]);

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

  const hasAutosized = useRef(false);
  const hasInitialDraftScroll = useRef(false);

  useEffect(() => {
    hasInitialDraftScroll.current = false;
  }, [schema.name, draftPosition]);

  const autosizeGridColumns = useCallback((api: FirstDataRenderedEvent['api']) => {
    // Autosize columns that lack an explicit user-set width and aren't truncate-mode.
    // Pass skipHeader=false so header text is also considered.
    const colIds: string[] = [];
    for (const col of schema.columns) {
      if (col.width || col.truncate || col.type === 'image') continue;
      if (col.type === 'reference' && col.refDisplayColumns && col.refDisplayColumns.length > 0) {
        for (const dp of col.refDisplayColumns) colIds.push(`${col.name}::${dp}`);
      } else if (col.type !== 'calculated') {
        colIds.push(col.name);
      } else if (col.showInGrid) {
        colIds.push(`__calc__${col.name}`);
      }
    }
    if (colIds.length > 0) api.autoSizeColumns(colIds, false);
  }, [schema.columns]);

  // Scroll to the last row accounting for the CSS padding-bottom on the viewport.
  const scrollToDraftBottom = useCallback((api: FirstDataRenderedEvent['api']) => {
    const displayedCount = api.getDisplayedRowCount();
    if (displayedCount === 0) return;
    api.ensureIndexVisible(displayedCount - 1, 'bottom');
    // AG Grid doesn't account for the CSS padding-bottom we add for the draft spacer.
    // Read it from the DOM and add it to the scroll position.
    requestAnimationFrame(() => {
      const viewport = gridWrapperRef.current?.querySelector<HTMLElement>('.ag-body-viewport');
      if (!viewport) return;
      const pad = parseFloat(getComputedStyle(viewport).paddingBottom) || 0;
      if (pad > 0) {
        viewport.scrollTop += pad;
      }
    });
  }, []);

  const onFirstDataRendered = useCallback((event: FirstDataRenderedEvent) => {
    // Scroll to bottom if draft is pinned there.
    if (draftPosition === 'bottom' && !readOnly && !filterActive && rows.length > 0) {
      scrollToDraftBottom(event.api);
      hasInitialDraftScroll.current = true;
    }
    autosizeGridColumns(event.api);
    hasAutosized.current = true;
    const cols = event.api.getColumns();
    colOrderRef.current = cols ? cols.map(c => c.getColId()) : [];
  }, [draftPosition, readOnly, filterActive, rows.length, autosizeGridColumns, scrollToDraftBottom]);

  // Handle async data: if rows arrive after onFirstDataRendered (e.g. network load),
  // run autosize once when the first non-empty batch appears.
  const onRowDataUpdated = useCallback((event: { api: FirstDataRenderedEvent['api'] }) => {
    if (!hasInitialDraftScroll.current && draftPosition === 'bottom' && !readOnly && !filterActive && rows.length > 0) {
      scrollToDraftBottom(event.api);
      hasInitialDraftScroll.current = true;
    }

    if (hasAutosized.current) return;
    if (event.api.getDisplayedRowCount() === 0) return;
    autosizeGridColumns(event.api);
    hasAutosized.current = true;
  }, [draftPosition, readOnly, filterActive, rows.length, autosizeGridColumns]);

  const clearAllFilters = useCallback(() => {
    gridRef.current?.api.setFilterModel(null);
    setFilterActive(false);
  }, []);

  const onColumnResized = useCallback((event: ColumnResizedEvent) => {
    // Only persist widths from explicit user drag-resizes, never from autosize or API calls.
    if (!event.finished || !onColumnWidthChange || event.source !== 'uiColumnResized') return;
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

  const confirm = useConfirm();

  const applyBulkEdit = useCallback(async () => {
    if (!bulkEditCol || selectedRowIds.size === 0) return;
    const isKeyCol = (schema.uniqueKeys ?? []).includes(bulkEditCol);
    if (isKeyCol) {
      const ok = await confirm(
        `"${bulkEditCol}" is part of the unique key for this table. Setting it to the same value on multiple rows will likely cause conflicts. Are you sure you want to bulk-edit this column?`,
        'Bulk-edit a key column?',
      );
      if (!ok) return;
    }
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
  }, [bulkEditCol, bulkEditValue, schema.columns, schema.uniqueKeys, selectedRowIds, rowIdToIndex, onEdit, confirm]);

  const deleteSelectedRows = useCallback(() => {
    if (selectedRowIds.size === 0) return;
    const indices = Array.from(selectedRowIds)
      .map(id => rowIdToIndex.get(id))
      .filter((idx): idx is number => idx !== undefined);

    const errors = onDeleteRows(indices);

    if (errors.length > 0) {
      const deleted = indices.length - errors.length;
      setError(`Delete: ${deleted} removed, ${errors.length} failed — ${errors[0].message}`);
    } else {
      setError(null);
    }
    gridRef.current?.api.deselectAll();
  }, [selectedRowIds, rowIdToIndex, onDeleteRows]);

  const getRowId = useCallback((params: GetRowIdParams) => {
    return params.data[INTERNAL_ROW_ID] ?? 'fallback';
  }, []);

  // Build AG Grid column definitions with valueSetter for validation
  const columnDefs: ColDef[] = useMemo(() => {
    const cols: ColDef[] = schema.columns.flatMap((col) => {
      if (col.type === 'calculated') return []; // handled separately below
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
            params: {
              refRows: getReferenceRows(refTable),
              refTable,
              resolveColumnPath,
              searchColumns: searchCols,
              displayColumns: displayCols,
              onCreateRecord: onCreateReferenceRow,
            },
          };
        };

        def.valueSetter = (params: ValueSetterParams) => setReferenceValue(params, params.newValue ?? '');

        const derivedDefs: ColDef[] = displayCols.map((displayPath) => {
          const leafResolvedLabel = resolveColumnPathLeafLabel(refTable, displayPath)
            || resolveColumnPathLabel(refTable, displayPath).split(' → ').pop()
            || displayPath;
          return ({
          colId: `${col.name}::${displayPath}`,
          headerName: leafResolvedLabel,
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
          });
        });

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
        def.valueFormatter = (params) => {
          const parsed = parseTemporalUnknown(params.value);
          if (!parsed) return String(params.value ?? '');
          return col.type === 'datetime' ? formatDateTimeCanonical(parsed) : formatDateCanonical(parsed);
        };
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

      // Date columns: inline editor with portal date picker
      if (col.type === 'date' || col.type === 'datetime') {
        def.cellEditor = DateCellEditor;
      }

      // Truncate mode: fixed display width, ellipsis for overflow (AG Grid default cell behaviour)
      if (col.truncate) {
        if (!col.width) def.width = 160;
        def.maxWidth = col.width ?? 160;
      }

      // List columns: chip/tag renderer + popup tag editor
      if (col.type === 'list') {
        def.cellRenderer = ListTagsRenderer;
        def.cellEditor = ListTagsEditor;
        def.cellEditorPopup = true;
        def.cellEditorPopupPosition = 'under';
        def.autoHeight = true;
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

    // Append read-only calculated columns (showInGrid === true)
    const dataColNames = schema.columns.filter(c => c.type !== 'calculated').map(c => c.name);
    for (const calc of schema.columns.filter(c => c.type === 'calculated' && c.showInGrid)) {
      if (!calc.expression) continue;
      const calcFn = getCalc(calc.expression);
      cols.push({
        field: `__calc__${calc.name}`,
        headerName: calc.displayName || calc.name,
        editable: false,
        suppressMovable: true,
        valueGetter: (params) => {
          if (!params.data || params.data[INTERNAL_ROW_ID] === DRAFT_ROW_ID) return '';
          const ctx: Record<string, number> = {};
          for (const n of dataColNames) ctx[n] = Number(params.data[n]) || 0;
          const result = calcFn(0, ctx);
          return Number.isFinite(result) ? result : '';
        },
        cellStyle: { color: 'var(--color-text-muted, #888)', fontStyle: 'italic' },
      });
    }

    // Open-record button column (leading, narrow)
    if (onOpenRecord) {
      cols.unshift({
        colId: '__open_record__',
        headerName: '',
        width: 32,
        maxWidth: 32,
        minWidth: 32,
        resizable: false,
        sortable: false,
        filter: false,
        editable: false,
        suppressMovable: true,
        cellRenderer: (params: CustomCellRendererProps) => (
          <OpenRecordButton {...params} onOpen={onOpenRecord} />
        ),
      });
    }

    return cols;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [schema, revision, rowIdToIndex, onEdit, onInsert, bookId, onOpenRecord, getReferencedRow, getReferenceRows, resolveColumnPath, resolveColumnPathLabel, onCreateReferenceRow]);

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

  const defaultColDef = useMemo(() => ({
    ...sharedDefaultColDef,
    suppressMovable: true,
    cellClassRules: {
      'cell-selected': (params: { node: { rowIndex: number | null }; column: { getColId: () => string } }) => {
        if (!selectModeRef.current || !selectionRef.current) return false;
        const rowIndex = params.node.rowIndex;
        if (rowIndex === null) return false;
        const sel = selectionRef.current;
        if (rowIndex < sel.minRow || rowIndex > sel.maxRow) return false;
        const colId = params.column.getColId();
        const colIdx = colOrderRef.current.indexOf(colId);
        return colIdx >= sel.minColIdx && colIdx <= sel.maxColIdx;
      },
    },
  }), []); // stable — reads from refs only

  // Editable column options for bulk edit dropdown
  const editableColumnOptions = useMemo(() =>
    schema.columns
      .filter(c => c.type !== 'image')
      .map(c => ({ value: c.name, label: c.displayName || c.name })),
    [schema],
  );

  const showBottomDraftSpacer = !readOnly && !filterActive && draftPosition === 'bottom';

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

  // ── Cell selection: tap-to-anchor, tap-to-extend model ───────────────────
  // First tap: anchor + single cell. Second tap outside selection: extend
  // rectangle from anchor to tapped cell. Tap inside selection: clear.
  useEffect(() => {
    const el = gridWrapperRef.current;
    if (!el) return;

    // Anchor persists across taps (local to this effect; cleared when selection clears)
    let anchorCell: { rowIndex: number; colId: string } | null = null;
    let downX = 0;
    let downY = 0;

    const getCellAt = (x: number, y: number): { rowIndex: number; colId: string } | null => {
      let node = document.elementFromPoint(x, y) as HTMLElement | null;
      while (node && !node.classList.contains('ag-cell')) node = node.parentElement;
      if (!node) return null;
      const colId = node.getAttribute('col-id');
      let row = node.parentElement;
      while (row && !row.classList.contains('ag-row')) row = row.parentElement;
      if (!row || !colId) return null;
      if (row.classList.contains('ag-row-pinned')) return null;
      const rowIndex = parseInt(row.getAttribute('row-index') ?? '', 10);
      if (isNaN(rowIndex) || rowIndex < 0) return null;
      return { rowIndex, colId };
    };

    const applyRange = (from: { rowIndex: number; colId: string }, to: { rowIndex: number; colId: string }) => {
      const api = gridRef.current?.api;
      if (!api) return;
      const colIds = colOrderRef.current;
      const fromColIdx = colIds.indexOf(from.colId);
      const toColIdx = colIds.indexOf(to.colId);
      if (fromColIdx < 0 || toColIdx < 0) return;
      selectionRef.current = {
        minRow: Math.min(from.rowIndex, to.rowIndex),
        maxRow: Math.max(from.rowIndex, to.rowIndex),
        minColIdx: Math.min(fromColIdx, toColIdx),
        maxColIdx: Math.max(fromColIdx, toColIdx),
      };
      api.refreshCells({ force: true });
      const sel = selectionRef.current;
      const nums: number[] = [];
      for (let r = sel.minRow; r <= sel.maxRow; r++) {
        const rowNode = api.getDisplayedRowAtIndex(r);
        if (!rowNode) continue;
        for (let ci = sel.minColIdx; ci <= sel.maxColIdx; ci++) {
          const cid = colIds[ci];
          if (!cid) continue;
          const raw = api.getCellValue({ rowNode, colKey: cid });
          if (raw === null || raw === undefined || raw === '') continue;
          const n = Number(raw);
          if (!isNaN(n)) nums.push(n);
        }
      }
      if (nums.length === 0) {
        setCellStats(null);
        onCellSelectionStats?.(null);
      } else {
        const sum = nums.reduce((a, b) => a + b, 0);
        setCellStats({ sum, count: nums.length, avg: sum / nums.length });
        onCellSelectionStats?.({ sum, count: nums.length, avg: sum / nums.length });
      }
    };

    const isCellInSelection = (cell: { rowIndex: number; colId: string }): boolean => {
      if (!selectionRef.current) return false;
      const sel = selectionRef.current;
      const colIdx = colOrderRef.current.indexOf(cell.colId);
      return (
        cell.rowIndex >= sel.minRow && cell.rowIndex <= sel.maxRow &&
        colIdx >= sel.minColIdx && colIdx <= sel.maxColIdx
      );
    };

    const onPointerDown = (e: PointerEvent) => {
      if (!selectModeRef.current || !e.isPrimary) return;
      downX = e.clientX;
      downY = e.clientY;
    };

    const onPointerUp = (e: PointerEvent) => {
      if (!selectModeRef.current || !e.isPrimary) return;
      // Only act on taps — ignore scrolls (movement > 8px)
      if (Math.hypot(e.clientX - downX, e.clientY - downY) > 8) return;

      const cell = getCellAt(e.clientX, e.clientY);
      if (!cell) {
        // Tapped outside grid — clear
        anchorCell = null;
        clearCellSelection();
        return;
      }

      if (!anchorCell) {
        // No anchor yet: set anchor and highlight single cell
        anchorCell = cell;
        applyRange(cell, cell);
      } else if (isCellInSelection(cell)) {
        // Tapped inside existing selection → clear
        anchorCell = null;
        clearCellSelection();
      } else {
        // Tapped outside selection → extend rectangle from anchor
        applyRange(anchorCell, cell);
      }
    };

    el.addEventListener('pointerdown', onPointerDown);
    el.addEventListener('pointerup', onPointerUp);
    return () => {
      el.removeEventListener('pointerdown', onPointerDown);
      el.removeEventListener('pointerup', onPointerUp);
    };
  }, [clearCellSelection, onCellSelectionStats]);

  // ── Mobile keyboard avoidance continued ──────────────────────────────────
  // Listen to visualViewport resize events (fires when mobile keyboard appears/disappears)
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    let prevHeight = vv.height;
    const KEYBOARD_THRESHOLD = 100;

    const onViewportResize = () => {
      const heightDiff = prevHeight - vv.height;
      prevHeight = vv.height;
      if (heightDiff > KEYBOARD_THRESHOLD) {
        // Keyboard just opened — scroll editor into view
        requestAnimationFrame(() => {
          scrollEditingCellIntoView();
        });
      }
    };

    vv.addEventListener('resize', onViewportResize);
    return () => vv.removeEventListener('resize', onViewportResize);
  }, [scrollEditingCellIntoView]);

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
        <button
          className={`btn-ghost btn-sm grid-sum-mode-btn${selectMode ? ' active' : ''}`}
          title={selectMode ? 'Clear sum selection' : 'Tap cells to sum'}
          onClick={() => {
            setSelectMode(m => {
              if (m) clearCellSelection();
              return !m;
            });
          }}
        >
          Σ {selectMode ? 'Summing' : 'Sum'}
        </button>
      </div>
      <div
        className={`grid-wrapper${showBottomDraftSpacer ? ' grid-wrapper--bottom-draft-spacer' : ''}`}
        ref={gridWrapperRef}
        style={{ flex: 1, minHeight: 0, zoom, touchAction: 'pan-x pan-y' }}
      >
        <AgGridReact
          ref={gridRef}
          modules={[AllCommunityModule]}
          theme={gridTheme}
          popupParent={popupParent}
          rowData={rowData}
          columnDefs={columnDefs}
          getRowId={getRowId}
          getRowClass={getRowClass}
          singleClickEdit={!selectMode}
          suppressCellFocus={selectMode}
          stopEditingWhenCellsLoseFocus={!keepEditorAlive}
          suppressScrollOnNewData={true}
          suppressScrollWhenPopupsAreOpen={true}
          enterNavigatesVertically={true}
          enterNavigatesVerticallyAfterEdit={true}
          suppressNoRowsOverlay={true}
          postSortRows={postSortRows}
          rowSelection={rowSelectionConfig}
          selectionColumnDef={{ width: 28, maxWidth: 28, minWidth: 28, pinned: false, suppressHeaderMenuButton: true }}
          onSelectionChanged={onSelectionChanged}
          onFilterChanged={onFilterChanged}
          onFirstDataRendered={onFirstDataRendered}
          onRowDataUpdated={onRowDataUpdated}
          defaultColDef={defaultColDef}
          onColumnResized={onColumnResized}
          onDisplayedColumnsChanged={onDisplayedColumnsChanged}
          onCellEditingStarted={onCellEditingStarted}
          onCellEditingStopped={onCellEditingStopped}
        />
      </div>
      <SelectionSumBar stats={cellStats} />
    </div>
  );
};

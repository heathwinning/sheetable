import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import type { TableSchema, Row, ValidationError, ChartSheet, ChartConfig, ChartLayoutItem, ViewSheet, UndoEntry, SessionUser, BookInfo } from './types';
import { INTERNAL_ROW_ID } from './types';
import * as api from './api';
import { log } from './DebugLogger';
import { applyChartValueFormat } from './chartFormat';

export interface UseAppStateReturn {
  // Auth
  user: SessionUser | null;
  isLoading: boolean;
  signIn: () => void;
  signOut: () => Promise<void>;

  // Books
  books: BookInfo[];
  activeBookId: string | null;
  activeBookName: string | null;
  activeBookRole: 'owner' | 'editor' | 'viewer' | null;
  switchBook: (bookId: string) => Promise<void>;
  refreshBooks: () => Promise<void>;
  createBook: (name: string) => Promise<string | null>;
  renameBook: (bookId: string, name: string) => Promise<void>;
  deleteBook: (bookId: string) => Promise<void>;

  // Tables
  tableIds: string[];
  activeTableId: string | null;
  setActiveTableId: (id: string | null) => void;
  getSchema: (tableId: string) => TableSchema | undefined;
  getRows: (tableId: string) => Row[];
  setRows: (tableId: string, rows: Row[]) => void;
  createTable: (schema: TableSchema, rows?: Row[]) => Promise<void>;
  deleteTable: (tableId: string) => Promise<void>;
  renameTable: (oldName: string, newName: string) => Promise<void>;
  renameColumn: (tableId: string, oldName: string, newName: string) => void;
  reorderTables: (fromIndex: number, toIndex: number) => void;
  reorderCharts: (fromIndex: number, toIndex: number) => void;
  reorderViews: (fromIndex: number, toIndex: number) => void;
  reorderTablesTo: (ids: string[]) => void;
  reorderChartsTo: (ids: string[]) => void;
  reorderViewsTo: (ids: string[]) => void;
  sortedSheets: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[];
  reorderAllSheetsTo: (items: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[]) => void;
  updateSchema: (tableId: string, schema: TableSchema) => Promise<void>;

  // Row operations
  applyEdit: (tableId: string, rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  insertRow: (tableId: string, row: Row) => ValidationError[];
  deleteRow: (tableId: string, rowIndex: number) => ValidationError[];
  deleteRows: (tableId: string, rowIndices: number[]) => ValidationError[];

  // Undo
  undo: () => ValidationError[];
  canUndo: boolean;

  // Charts
  chartSheetIds: string[];
  getChartSheet: (id: string) => ChartSheet | undefined;
  createChartSheet: (name: string) => Promise<void>;
  deleteChartSheet: (name: string) => Promise<void>;
  renameChartSheet: (oldName: string, newName: string) => Promise<void>;
  updateChartSheet: (name: string, charts: ChartConfig[], layout: ChartLayoutItem[]) => Promise<void>;

  // View sheets
  viewSheetIds: string[];
  getViewSheet: (id: string) => ViewSheet | undefined;
  createViewSheet: (name: string, tableName: string, viewType: ViewSheet['viewType'], dateColumn?: string) => Promise<void>;
  deleteViewSheet: (name: string) => Promise<void>;
  renameViewSheet: (oldName: string, newName: string) => Promise<void>;
  updateViewSheet: (name: string, updates: Partial<Pick<ViewSheet, 'name' | 'tableName' | 'viewType' | 'dateColumn' | 'hideSourceTableTab'>>) => Promise<void>;

  // Reference helpers (replacement for DataModel methods)
  getReferencedRow: (refTable: string, rowId: string) => Row | undefined;
  getReferenceRows: (refTable: string) => Row[];
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  resolveColumnPathLabel: (tableName: string, path: string) => string;
  getColumnPaths: (tableName: string) => { path: string; label: string }[];

  // Revision counter
  revision: number;
}

// ---- Validation helpers ----

function validateType(type: string, value: string, rowIndex: number, columnName: string): ValidationError[] {
  if (value === '') return [];
  switch (type) {
    case 'integer':
      if (!/^-?\d+$/.test(value))
        return [{ message: `"${value}" is not a valid integer`, rowIndex, columnName }];
      break;
    case 'decimal':
      if (isNaN(Number(value)) || value.trim() === '')
        return [{ message: `"${value}" is not a valid decimal`, rowIndex, columnName }];
      break;
    case 'date':
      if (!/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(value) || isNaN(Date.parse(value.replace(/\//g, '-'))))
        return [{ message: `"${value}" is not a valid date (YYYY/MM/DD)`, rowIndex, columnName }];
      break;
    case 'datetime':
      if (isNaN(Date.parse(value)))
        return [{ message: `"${value}" is not a valid datetime`, rowIndex, columnName }];
      break;
    case 'bool':
      if (!['true', 'false', '1', '0', 'yes', 'no'].includes(value.toLowerCase()))
        return [{ message: `"${value}" is not a valid boolean (true/false)`, rowIndex, columnName }];
      break;
    case 'list':
      try { JSON.parse(value); } catch {
        return [{ message: `"${value}" is not a valid list (must be JSON array)`, rowIndex, columnName }];
      }
      break;
  }
  return [];
}

function validateUniqueKey(rows: Row[], _schema: TableSchema, keyValues: Record<string, string>, excludeRow: number): ValidationError[] {
  const keyColumns = Object.keys(keyValues);
  for (const colName of keyColumns) {
    if (keyValues[colName] === '') {
      return [{ message: `Key column "${colName}" cannot be empty`, rowIndex: excludeRow, columnName: colName }];
    }
  }
  for (let i = 0; i < rows.length; i++) {
    if (i === excludeRow) continue;
    const matches = keyColumns.every(col => rows[i][col] === keyValues[col]);
    if (matches) {
      const keyStr = keyColumns.map(c => `${c}="${keyValues[c]}"`).join(', ');
      return [{
        message: keyColumns.length === 1
          ? `Duplicate key "${keyValues[keyColumns[0]]}" in column "${keyColumns[0]}"`
          : `Duplicate composite key: ${keyStr}`,
        rowIndex: excludeRow,
        columnName: keyColumns[0],
      }];
    }
  }
  return [];
}

// ---- Per-row write queue ----

const rowQueues = new Map<string, Promise<void>>();

function enqueueWrite(rowId: string, fn: () => Promise<void>): void {
  const prev = rowQueues.get(rowId) ?? Promise.resolve();
  const next = prev.then(fn, fn).then(() => {
    if (rowQueues.get(rowId) === next) rowQueues.delete(rowId);
  });
  rowQueues.set(rowId, next);
}

// ---- Main hook ----

export function useAppState(): UseAppStateReturn {
  const [user, setUser] = useState<SessionUser | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [books, setBooks] = useState<BookInfo[]>([]);
  const [activeBookId, setActiveBookId] = useState<string | null>(null);
  const [revision, setRevision] = useState(0);
  const [activeTableId, setActiveTableId] = useState<string | null>(null);
  const [tableOrder, setTableOrder] = useState<string[]>([]);
  const [chartSheetOrder, setChartSheetOrder] = useState<string[]>([]);
  const [viewSheetOrder, setViewSheetOrder] = useState<string[]>([]);

  // In-memory data store (replaces DataModel)
  const schemasRef = useRef<Map<string, TableSchema>>(new Map());
  const rowsRef = useRef<Map<string, Row[]>>(new Map());
  const chartSheetsRef = useRef<Map<string, ChartSheet>>(new Map());
    const viewSheetsRef = useRef<Map<string, ViewSheet>>(new Map());
  const undoStackRef = useRef<UndoEntry[]>([]);

  const bump = useCallback(() => setRevision(r => r + 1), []);

  // ---- Check auth on mount ----
  useEffect(() => {
    api.getMe()
      .then(u => { setUser(u); setIsLoading(false); })
      .catch(() => setIsLoading(false));
  }, []);

  // ---- Load books when user changes ----
  useEffect(() => {
    if (!user) { setBooks([]); setActiveBookId(null); return; }
    api.listBooks().then(b => {
      setBooks(b);
      if (b.length > 0 && !activeBookId) setActiveBookId(b[0].id);
    }).catch(console.error);
  }, [user]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---- Load tables + charts when active book changes ----
  useEffect(() => {
    if (!activeBookId) {
      schemasRef.current.clear();
      rowsRef.current.clear();
      chartSheetsRef.current.clear();
      viewSheetsRef.current.clear();
      setTableOrder([]);
      setChartSheetOrder([]);
      setViewSheetOrder([]);
      setActiveTableId(null);
      bump();
      return;
    }

    let cancelled = false;

    (async () => {
      try {
        const [schemas, charts, views] = await Promise.all([
          api.listTables(activeBookId),
          api.listCharts(activeBookId),
          api.listViews(activeBookId),
        ]);

        if (cancelled) return;

        schemasRef.current.clear();
        rowsRef.current.clear();

        const order: string[] = [];
        for (const schema of schemas) {
          schemasRef.current.set(schema.name, schema);
          order.push(schema.name);
        }
        setTableOrder(order);

        chartSheetsRef.current.clear();
        const chartOrder: string[] = [];
        for (const chart of charts) {
          const raw = (chart as { charts?: unknown }).charts;
          let parsedCharts: ChartConfig[] = [];
          let parsedLayout: ChartLayoutItem[] = [];
          if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
            const cfg = raw as { charts?: unknown; layout?: unknown };
            if (Array.isArray(cfg.charts)) parsedCharts = cfg.charts as ChartConfig[];
            if (Array.isArray(cfg.layout)) parsedLayout = cfg.layout as ChartLayoutItem[];
          }
          chartSheetsRef.current.set(chart.name, { name: chart.name, charts: parsedCharts, layout: parsedLayout });
          chartOrder.push(chart.name);
        }
        setChartSheetOrder(chartOrder);

        viewSheetsRef.current.clear();
        const viewOrder: string[] = [];
        for (const view of views) {
          viewSheetsRef.current.set(view.name, view);
          viewOrder.push(view.name);
        }
        setViewSheetOrder(viewOrder);

        // Load all row data in parallel
        const rowPromises = schemas.map(async (schema) => {
          const rows = await api.listRows(activeBookId, schema.name);
          if (!cancelled) rowsRef.current.set(schema.name, rows);
        });
        await Promise.all(rowPromises);

        if (!cancelled) {
          undoStackRef.current = [];
          if (order.length > 0) setActiveTableId(order[0]);
          bump();
        }
      } catch (err) {
        console.error('Failed to load book data:', err);
      }
    })();

    return () => { cancelled = true; };
  }, [activeBookId, bump]);

  const activeBookName = useMemo(() => {
    return books.find(b => b.id === activeBookId)?.name ?? null;
  }, [books, activeBookId]);

  const activeBookRole = useMemo(() => {
    return (books.find(b => b.id === activeBookId)?.role ?? null) as 'owner' | 'editor' | 'viewer' | null;
  }, [books, activeBookId]);

  // ---- Auth ----
  const signIn = useCallback(() => {
    // Save current hash route so we can redirect back after OAuth
    const hash = window.location.hash;
    if (hash && hash !== '#/' && hash !== '#') {
      localStorage.setItem('sheetable-post-login-redirect', hash);
    }
    window.location.href = api.loginUrl();
  }, []);

  const signOut = useCallback(async () => {
    await api.logout();
    setUser(null);
    setBooks([]);
    setActiveBookId(null);
  }, []);

  // ---- Books ----
  const refreshBooks = useCallback(async () => {
    const b = await api.listBooks();
    setBooks(b);
  }, []);

  const switchBook = useCallback(async (bookId: string) => {
    setActiveBookId(bookId);
  }, []);

  const createBook = useCallback(async (name: string): Promise<string | null> => {
    try {
      const book = await api.createBook(name);
      setBooks(prev => [...prev, { id: book.id, name: book.name, owner_id: user?.id ?? '', role: 'owner', created_at: new Date().toISOString() }]);
      setActiveBookId(book.id);
      return book.id;
    } catch (err) {
      console.error('Failed to create book:', err);
      return null;
    }
  }, [user]);

  const renameBook = useCallback(async (bookId: string, name: string) => {
    await api.renameBook(bookId, name);
    setBooks(prev => prev.map(b => b.id === bookId ? { ...b, name } : b));
  }, []);

  const doDeleteBook = useCallback(async (bookId: string) => {
    await api.deleteBook(bookId);
    setBooks(prev => {
      const next = prev.filter(b => b.id !== bookId);
      if (activeBookId === bookId) {
        setActiveBookId(next.length > 0 ? next[0].id : null);
      }
      return next;
    });
  }, [activeBookId]);

  // ---- Reference helpers (replaces DataModel) ----
  const getReferencedRow = useCallback((refTable: string, rowId: string): Row | undefined => {
    const rows = rowsRef.current.get(refTable);
    return rows?.find(r => r[INTERNAL_ROW_ID] === rowId);
  }, []);

  const getReferenceRows = useCallback((refTable: string): Row[] => {
    return rowsRef.current.get(refTable) ?? [];
  }, []);

  const resolveColumnPath = useCallback((tableName: string, row: Row, path: string): string => {
    const parts = path.split('.');
    const schema = schemasRef.current.get(tableName);
    if (!schema) return '';

    const colName = parts[0];
    const value = row[colName] ?? '';

    // Check if this is a calculated column (only at root level, non-dotted)
    if (parts.length === 1) {
      const calcCol = schema.columns.find(c => c.name === colName && c.type === 'calculated');
      if (calcCol?.expression) {
        // Build a context with all numeric column values from the row
        const ctx: Record<string, number> = {};
        for (const col of schema.columns) {
          const v = Number(row[col.name]);
          ctx[col.name] = isNaN(v) ? 0 : v;
        }
        return applyChartValueFormat(0, { valueCalc: calcCol.expression }, ctx);
      }
    }

    if (parts.length === 1) {
      const col = schema.columns.find(c => c.name === colName);
      if (col?.type === 'reference' && col.refTable && value) {
        const refRow = getReferencedRow(col.refTable, value);
        if (!refRow) return '';
        const displayCols = col.refDisplayColumns ?? [];
        if (displayCols.length > 0) {
          return displayCols.map(dc => resolveColumnPath(col.refTable!, refRow, dc)).filter(Boolean).join(' · ');
        }
        return value;
      }
      return value;
    }

    const col = schema.columns.find(c => c.name === colName);
    if (!col || col.type !== 'reference' || !col.refTable || !value) return '';

    const refRow = getReferencedRow(col.refTable, value);
    if (!refRow) return '';
    return resolveColumnPath(col.refTable, refRow, parts.slice(1).join('.'));
  }, [getReferencedRow]);

  const resolveColumnPathLabel = useCallback((tableName: string, path: string): string => {
    const parts = path.split('.').filter(Boolean);
    if (parts.length === 0) return '';

    const labels: string[] = [];
    let currentTable = tableName;

    for (const part of parts) {
      const schema = schemasRef.current.get(currentTable);
      if (!schema) { labels.push(part); break; }
      const col = schema.columns.find(c => c.name === part);
      if (!col) { labels.push(part); break; }
      labels.push(col.displayName || col.name);
      if (col.type !== 'reference' || !col.refTable) break;
      currentTable = col.refTable;
    }

    return labels.join(' → ');
  }, []);

  const getColumnPaths = useCallback((tableName: string): { path: string; label: string }[] => {
    const MAX_DEPTH = 3;
    const result: { path: string; label: string }[] = [];
    const seen = new Set<string>(); // prevent cycles

    const walk = (table: string, prefix: string, labelPrefix: string, depth: number) => {
      if (depth > MAX_DEPTH || seen.has(table)) return;
      seen.add(table);
      const schema = schemasRef.current.get(table);
      if (!schema) return;
      for (const col of schema.columns) {
        const path = prefix ? `${prefix}.${col.name}` : col.name;
        const colLabel = col.displayName || col.name;
        const label = labelPrefix ? `${labelPrefix} → ${colLabel}` : colLabel;
        if (col.type === 'reference' && col.refTable) {
          walk(col.refTable, path, label, depth + 1);
        } else {
          result.push({ path, label });
        }
      }
      seen.delete(table);
    };

    walk(tableName, '', '', 0);
    return result;
  }, []);

  // ---- Table CRUD ----
  const getSchema = useCallback((tableId: string) => schemasRef.current.get(tableId), []);
  const getRows = useCallback((tableId: string) => rowsRef.current.get(tableId) ?? [], []);

  const doCreateTable = useCallback(async (schema: TableSchema, rows?: Row[]) => {
    if (!activeBookId) return;
    await api.createTable(activeBookId, schema, rows);

    schemasRef.current.set(schema.name, schema);
    rowsRef.current.set(schema.name, rows ?? []);
    setTableOrder(prev => [...prev, schema.name]);
    setActiveTableId(schema.name);

    // Reload rows from server to get server-assigned IDs
    const serverRows = await api.listRows(activeBookId, schema.name);
    rowsRef.current.set(schema.name, serverRows);
    bump();
  }, [activeBookId, bump]);

  const doDeleteTable = useCallback(async (tableId: string) => {
    if (!activeBookId) return;
    await api.deleteTable(activeBookId, tableId);
    schemasRef.current.delete(tableId);
    rowsRef.current.delete(tableId);
    setTableOrder(prev => {
      const next = prev.filter(id => id !== tableId);
      if (activeTableId === tableId) {
        setActiveTableId(next.length > 0 ? next[0] : null);
      }
      return next;
    });
    bump();
  }, [activeBookId, activeTableId, bump]);

  const doRenameTable = useCallback(async (oldName: string, newName: string) => {
    if (!activeBookId) return;
    await api.renameTable(activeBookId, oldName, newName);

    const schema = schemasRef.current.get(oldName);
    if (schema) {
      schema.name = newName;
      schemasRef.current.delete(oldName);
      schemasRef.current.set(newName, schema);
    }
    const rows = rowsRef.current.get(oldName);
    if (rows) {
      rowsRef.current.delete(oldName);
      rowsRef.current.set(newName, rows);
    }
    // Update references in other tables
    for (const [, s] of schemasRef.current) {
      for (const col of s.columns) {
        if (col.refTable === oldName) col.refTable = newName;
      }
    }
    // Update view sheets
    for (const [, view] of viewSheetsRef.current) {
      if (view.tableName === oldName) {
        view.tableName = newName;
        api.updateView(activeBookId, view.name, { tableName: newName }).catch(console.error);
      }
    }
    // Update chart sheets
    for (const [, sheet] of chartSheetsRef.current) {
      let dirty = false;
      for (const chart of sheet.charts) {
        if (chart.table === oldName) { chart.table = newName; dirty = true; }
      }
      if (dirty) {
        api.updateChart(activeBookId, sheet.name, { charts: { charts: sheet.charts, layout: sheet.layout } }).catch(console.error);
      }
    }
    setTableOrder(prev => prev.map(id => id === oldName ? newName : id));
    if (activeTableId === oldName) setActiveTableId(newName);
    bump();
  }, [activeBookId, activeTableId, bump]);

  const doRenameColumn = useCallback((tableId: string, oldName: string, newName: string) => {
    if (!activeBookId || !oldName || !newName || oldName === newName) return;

    const schema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);
    if (!schema || !rows) return;

    // Rename in rows
    for (const row of rows) {
      if (oldName in row) {
        row[newName] = row[oldName] ?? '';
        delete row[oldName];
      }
    }

    // Rename in schema
    for (const col of schema.columns) {
      if (col.name === oldName) col.name = newName;
    }
    schema.uniqueKeys = schema.uniqueKeys.map(k => k === oldName ? newName : k);
    if (schema.defaultSort) {
      schema.defaultSort = schema.defaultSort.map(s => s.column === oldName ? { ...s, column: newName } : s);
    }

    // Update reference paths in other schemas
    for (const [, s] of schemasRef.current) {
      for (const col of s.columns) {
        if (col.type === 'reference' && col.refTable === tableId) {
          if (col.refDisplayColumns) {
            col.refDisplayColumns = col.refDisplayColumns.map(p => rewritePath(p, oldName, newName));
          }
          if (col.refSearchColumns) {
            col.refSearchColumns = col.refSearchColumns.map(p => rewritePath(p, oldName, newName));
          }
        }
      }
    }

    // Update view sheets
    for (const [, view] of viewSheetsRef.current) {
      if (view.tableName === tableId && view.dateColumn === oldName) {
        view.dateColumn = newName;
        api.updateView(activeBookId, view.name, { dateColumn: newName }).catch(console.error);
      }
    }
    // Update chart sheets
    for (const [, sheet] of chartSheetsRef.current) {
      let dirty = false;
      for (const chart of sheet.charts) {
        if (chart.table !== tableId) continue;
        const rw = (s: string) => rewriteColumnExpr(s, oldName, newName);
        if (chart.xColumn) { const r = rw(chart.xColumn); if (r !== chart.xColumn) { chart.xColumn = r; dirty = true; } }
        if (chart.yColumn) { const r = rw(chart.yColumn); if (r !== chart.yColumn) { chart.yColumn = r; dirty = true; } }
        if (chart.groupBy) { const r = rw(chart.groupBy); if (r !== chart.groupBy) { chart.groupBy = r; dirty = true; } }
        if (chart.filterColumn) { const r = rw(chart.filterColumn); if (r !== chart.filterColumn) { chart.filterColumn = r; dirty = true; } }
        if (chart.filters) {
          const nf = chart.filters.map(f => ({ ...f, column: rw(f.column) }));
          if (nf.some((f, i) => f.column !== chart.filters![i].column)) { chart.filters = nf; dirty = true; }
        }
        if (chart.tableRows) { const rr = chart.tableRows.map(rw); if (rr.some((v, i) => v !== chart.tableRows![i])) { chart.tableRows = rr; dirty = true; } }
        if (chart.tableColumns) { const rc = chart.tableColumns.map(rw); if (rc.some((v, i) => v !== chart.tableColumns![i])) { chart.tableColumns = rc; dirty = true; } }
        if (chart.tableSort) { const rk = rw(chart.tableSort.key); if (rk !== chart.tableSort.key) { chart.tableSort = { ...chart.tableSort, key: rk }; dirty = true; } }
      }
      if (dirty) {
        api.updateChart(activeBookId, sheet.name, { charts: { charts: sheet.charts, layout: sheet.layout } }).catch(console.error);
      }
    }
    // Note: no API call for schemas here — the caller (handleSave) follows with updateSchema which persists everything.
    bump();
  }, [activeBookId, bump]);

  const doReorderTablesTo = useCallback((ids: string[]) => {
    setTableOrder(ids);
    if (activeBookId) api.reorderSheets(activeBookId, ids).catch(console.error);
  }, [activeBookId]);

  const doReorderChartsTo = useCallback((ids: string[]) => {
    setChartSheetOrder(ids);
    if (activeBookId) api.reorderSheets(activeBookId, undefined, ids).catch(console.error);
  }, [activeBookId]);

  const doReorderViewsTo = useCallback((ids: string[]) => {
    setViewSheetOrder(ids);
    if (activeBookId) {
      Promise.all(ids.map((name, i) => api.updateView(activeBookId, name, { displayOrder: i }))).catch(console.error);
    }
  }, [activeBookId]);

  // ---- Global cross-type sheet ordering ----
  const sortedSheets = useMemo(() => {
    const book = books.find(b => b.id === activeBookId);
    const raw = book?.sheet_order;
    if (raw) {
      try {
        const parsed: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[] = JSON.parse(raw);
        const existing = new Set<string>();
        const result: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[] = [];
        for (const item of parsed) {
          const key = `${item.type}:${item.name}`;
          if (
            (item.type === 'table' && tableOrder.includes(item.name)) ||
            (item.type === 'chart' && chartSheetOrder.includes(item.name)) ||
            (item.type === 'view' && viewSheetOrder.includes(item.name))
          ) {
            result.push(item);
            existing.add(key);
          }
        }
        for (const name of tableOrder) {
          if (!existing.has(`table:${name}`)) result.push({ type: 'table', name });
        }
        for (const name of chartSheetOrder) {
          if (!existing.has(`chart:${name}`)) result.push({ type: 'chart', name });
        }
        for (const name of viewSheetOrder) {
          if (!existing.has(`view:${name}`)) result.push({ type: 'view', name });
        }
        return result;
      } catch { /* fall through */ }
    }
    return [
      ...tableOrder.map(name => ({ type: 'table' as const, name })),
      ...chartSheetOrder.map(name => ({ type: 'chart' as const, name })),
      ...viewSheetOrder.map(name => ({ type: 'view' as const, name })),
    ];
  }, [books, activeBookId, tableOrder, chartSheetOrder, viewSheetOrder]);

  const doReorderAllSheetsTo = useCallback((items: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[]) => {
    const newTableOrder = items.filter(i => i.type === 'table').map(i => i.name);
    const newChartOrder = items.filter(i => i.type === 'chart').map(i => i.name);
    const newViewOrder = items.filter(i => i.type === 'view').map(i => i.name);
    setTableOrder(newTableOrder);
    setChartSheetOrder(newChartOrder);
    setViewSheetOrder(newViewOrder);
    if (activeBookId) {
      setBooks(prev => prev.map(b => b.id === activeBookId ? { ...b, sheet_order: JSON.stringify(items) } : b));
      api.reorderAllSheets(activeBookId, items).catch(console.error);
    }
  }, [activeBookId]);

  const doReorderTables = useCallback((fromIndex: number, toIndex: number) => {
    setTableOrder(prev => {
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      if (activeBookId) {
        api.reorderSheets(activeBookId, next).catch(console.error);
      }
      return next;
    });
  }, [activeBookId]);

  const doReorderCharts = useCallback((fromIndex: number, toIndex: number) => {
    setChartSheetOrder(prev => {
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      if (activeBookId) {
        api.reorderSheets(activeBookId, undefined, next).catch(console.error);
      }
      return next;
    });
  }, [activeBookId]);

  const doReorderViews = useCallback((fromIndex: number, toIndex: number) => {
    setViewSheetOrder(prev => {
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      if (activeBookId) {
        Promise.all(next.map((name, i) => api.updateView(activeBookId, name, { displayOrder: i }))).catch(console.error);
      }
      return next;
    });
  }, [activeBookId]);

  const doUpdateSchema = useCallback(async (tableId: string, schema: TableSchema) => {
    if (!activeBookId) return;

    const oldSchema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);

    if (rows && oldSchema) {
      const oldCols = new Set(oldSchema.columns.map(c => c.name));
      const newCols = new Set(schema.columns.map(c => c.name));
      for (const row of rows) {
        for (const col of schema.columns) {
          if (!oldCols.has(col.name) && !(col.name in row)) row[col.name] = '';
        }
        for (const oldCol of oldCols) {
          if (!newCols.has(oldCol)) delete row[oldCol];
        }
      }
    }

    schemasRef.current.set(tableId, schema);
    await api.updateTableSchema(activeBookId, tableId, schema);
    bump();
  }, [activeBookId, bump]);

  // ---- Row operations with client-side validation + async server writes ----

  const applyEdit = useCallback((tableId: string, rowIndex: number, columnName: string, newValue: string): ValidationError[] => {
    const schema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);
    if (!schema || !rows || rowIndex < 0 || rowIndex >= rows.length) {
      return [{ message: 'Table or row not found', rowIndex }];
    }

    const col = schema.columns.find(c => c.name === columnName);
    if (!col) return [{ message: `Column "${columnName}" not found`, rowIndex }];

    // Type validation
    const typeErrors = validateType(col.type, newValue, rowIndex, columnName);
    if (typeErrors.length > 0) return typeErrors;

    // Unique key validation
    if (schema.uniqueKeys.includes(columnName)) {
      const keyValues: Record<string, string> = {};
      for (const keyCol of schema.uniqueKeys) {
        keyValues[keyCol] = keyCol === columnName ? newValue : rows[rowIndex][keyCol];
      }
      const errors = validateUniqueKey(rows, schema, keyValues, rowIndex);
      if (errors.length > 0) return errors;
    }

    // Reference validation
    if (col.type === 'reference' && col.refTable && newValue !== '') {
      const refRows = rowsRef.current.get(col.refTable);
      if (!refRows?.some(r => r[INTERNAL_ROW_ID] === newValue)) {
        return [{ message: `Referenced row not found in "${col.refTable}"`, rowIndex }];
      }
    }

    // Apply optimistic update
    const oldValue = rows[rowIndex][columnName] ?? '';
    const rowId = rows[rowIndex][INTERNAL_ROW_ID];
    rows[rowIndex][columnName] = newValue;

    // Push undo entry
    undoStackRef.current.push({ type: 'update', tableId, rowId, column: columnName, oldValue, newValue });
    log('applyEdit:', tableId, rowId, columnName, oldValue, '->', newValue);

    // Async server write (per-row queue)
    if (activeBookId) {
      const bookId = activeBookId;
      enqueueWrite(rowId, () =>
        api.updateRow(bookId, tableId, rowId, { [columnName]: newValue })
          .catch(err => {
            console.error('Server write failed, reverting:', err);
            const currentRows = rowsRef.current.get(tableId);
            const row = currentRows?.find(r => r[INTERNAL_ROW_ID] === rowId);
            if (row) {
              row[columnName] = oldValue;
              bump();
            }
          })
      );
    }

    bump();
    return [];
  }, [activeBookId, bump]);

  const doInsertRow = useCallback((tableId: string, row: Row): ValidationError[] => {
    const schema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);
    if (!schema || !rows) return [{ message: 'Table not found', rowIndex: -1 }];

    const newRowIndex = rows.length;

    // Validate types
    for (const col of schema.columns) {
      const value = row[col.name] ?? '';
      const typeErrors = validateType(col.type, value, newRowIndex, col.name);
      if (typeErrors.length > 0) return typeErrors;

      if (col.type === 'reference' && col.refTable && value !== '') {
        const refRows = rowsRef.current.get(col.refTable);
        if (!refRows?.some(r => r[INTERNAL_ROW_ID] === value)) {
          return [{ message: `Referenced row not found in "${col.refTable}"`, rowIndex: newRowIndex }];
        }
      }
    }

    // Validate unique key
    if (schema.uniqueKeys.length > 0) {
      const keyValues: Record<string, string> = {};
      for (const keyCol of schema.uniqueKeys) keyValues[keyCol] = row[keyCol] ?? '';
      const errors = validateUniqueKey(rows, schema, keyValues, -1);
      if (errors.length > 0) return errors;
    }

    // Sequential integer _rowId (max existing + 1)
    const completeRow: Row = {};
    const maxId = Math.max(0, ...rows.map(r => Number(r[INTERNAL_ROW_ID]) || 0));
    const rowId = String(maxId + 1);
    completeRow[INTERNAL_ROW_ID] = rowId;
    for (const col of schema.columns) {
      completeRow[col.name] = row[col.name] ?? '';
    }

    rows.push(completeRow);
    undoStackRef.current.push({ type: 'insert', tableId, rowId, row: { ...completeRow } });
    log('insertRow:', tableId, rowId);

    // Async server write
    if (activeBookId) {
      const bookId = activeBookId;
      api.insertRow(bookId, tableId, completeRow).catch(err => {
        console.error('Server insert failed, removing row:', err);
        const currentRows = rowsRef.current.get(tableId);
        if (currentRows) {
          const idx = currentRows.findIndex(r => r[INTERNAL_ROW_ID] === rowId);
          if (idx >= 0) currentRows.splice(idx, 1);
          bump();
        }
      });
    }

    bump();
    return [];
  }, [activeBookId, bump]);

  const doDeleteRow = useCallback((tableId: string, rowIndex: number): ValidationError[] => {
    const schema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);
    if (!schema || !rows || rowIndex < 0 || rowIndex >= rows.length) {
      return [{ message: 'Row not found', rowIndex }];
    }

    const rowId = rows[rowIndex][INTERNAL_ROW_ID];

    // Check references from other tables
    for (const [otherName, otherSchema] of schemasRef.current) {
      for (const col of otherSchema.columns) {
        if (col.type === 'reference' && col.refTable === tableId) {
          const otherRows = rowsRef.current.get(otherName);
          if (otherRows?.some(r => r[col.name] === rowId)) {
            return [{ message: `Cannot delete: row referenced by "${otherName}"`, rowIndex }];
          }
        }
      }
    }

    const deletedRow = { ...rows[rowIndex] };

    if (!activeBookId) {
      rows.splice(rowIndex, 1);
      undoStackRef.current.push({ type: 'delete', tableId, rowId, row: deletedRow });
      bump();
      return [];
    }

    // Awaited delete — perform server call, remove on success
    const bookId = activeBookId;
    api.deleteRow(bookId, tableId, rowId).then(() => {
      const currentRows = rowsRef.current.get(tableId);
      if (currentRows) {
        const idx = currentRows.findIndex(r => r[INTERNAL_ROW_ID] === rowId);
        if (idx >= 0) currentRows.splice(idx, 1);
      }
      undoStackRef.current.push({ type: 'delete', tableId, rowId, row: deletedRow });
      bump();
    }).catch(err => {
      console.error('Server delete failed:', err);
    });

    // Return empty to indicate no client error — the actual removal happens async
    return [];
  }, [activeBookId, bump]);

  const doDeleteRows = useCallback((tableId: string, rowIndices: number[]): ValidationError[] => {
    const schema = schemasRef.current.get(tableId);
    const rows = rowsRef.current.get(tableId);
    if (!schema || !rows) return [{ message: 'Table not found', rowIndex: 0 }];

    // Sort descending so indices stay valid during removal
    const sorted = [...rowIndices].sort((a, b) => b - a);

    const validRows: { idx: number; rowId: string; row: Row }[] = [];
    const errors: ValidationError[] = [];

    for (const idx of sorted) {
      if (idx < 0 || idx >= rows.length) {
        errors.push({ message: 'Row not found', rowIndex: idx });
        continue;
      }
      const rowId = rows[idx][INTERNAL_ROW_ID];

      // Check references from other tables
      let refError: string | null = null;
      for (const [otherName, otherSchema] of schemasRef.current) {
        for (const col of otherSchema.columns) {
          if (col.type === 'reference' && col.refTable === tableId) {
            const otherRows = rowsRef.current.get(otherName);
            if (otherRows?.some(r => r[col.name] === rowId)) {
              refError = `Cannot delete: row referenced by "${otherName}"`;
            }
          }
        }
      }
      if (refError) {
        errors.push({ message: refError, rowIndex: idx });
        continue;
      }

      validRows.push({ idx, rowId, row: { ...rows[idx] } });
    }

    if (validRows.length === 0) return errors;

    if (!activeBookId) {
      // Offline: remove all valid rows (already sorted desc)
      for (const { idx, rowId, row } of validRows) {
        rows.splice(idx, 1);
        undoStackRef.current.push({ type: 'delete', tableId, rowId, row });
      }
      bump();
      return errors;
    }

    // Online: single bulk request
    const bookId = activeBookId;
    const rowIds = validRows.map(v => v.rowId);
    api.deleteRows(bookId, tableId, rowIds).then(() => {
      const currentRows = rowsRef.current.get(tableId);
      if (currentRows) {
        const idSet = new Set(rowIds);
        for (let i = currentRows.length - 1; i >= 0; i--) {
          if (idSet.has(currentRows[i][INTERNAL_ROW_ID])) {
            const removed = currentRows.splice(i, 1)[0];
            undoStackRef.current.push({ type: 'delete', tableId, rowId: removed[INTERNAL_ROW_ID], row: removed });
          }
        }
      }
      bump();
    }).catch(err => {
      console.error('Server bulk delete failed:', err);
    });

    return errors;
  }, [activeBookId, bump]);
  const canUndo = undoStackRef.current.length > 0;

  const doUndo = useCallback((): ValidationError[] => {
    const entry = undoStackRef.current.pop();
    if (!entry) return [];

    const rows = rowsRef.current.get(entry.tableId);
    if (!rows) return [];

    if (entry.type === 'update' && entry.column && entry.oldValue !== undefined) {
      const row = rows.find(r => r[INTERNAL_ROW_ID] === entry.rowId);
      if (row) {
        row[entry.column] = entry.oldValue;
        if (activeBookId) {
          enqueueWrite(entry.rowId, () =>
            api.updateRow(activeBookId, entry.tableId, entry.rowId, { [entry.column!]: entry.oldValue! })
          );
        }
      }
    } else if (entry.type === 'insert') {
      const idx = rows.findIndex(r => r[INTERNAL_ROW_ID] === entry.rowId);
      if (idx >= 0) {
        rows.splice(idx, 1);
        if (activeBookId) {
          api.deleteRow(activeBookId, entry.tableId, entry.rowId).catch(console.error);
        }
      }
    } else if (entry.type === 'delete' && entry.row) {
      rows.push(entry.row);
      if (activeBookId) {
        api.insertRow(activeBookId, entry.tableId, entry.row).catch(console.error);
      }
    }

    bump();
    return [];
  }, [activeBookId, bump]);

  // ---- Chart sheets ----
  const getChartSheet = useCallback((id: string) => chartSheetsRef.current.get(id), []);

  const doCreateChartSheet = useCallback(async (name: string) => {
    if (!activeBookId) return;
    await api.createChart(activeBookId, name);
    chartSheetsRef.current.set(name, { name, charts: [], layout: [] });
    setChartSheetOrder(prev => [...prev, name]);
    bump();
  }, [activeBookId, bump]);

  const doDeleteChartSheet = useCallback(async (name: string) => {
    if (!activeBookId) return;
    await api.deleteChart(activeBookId, name);
    chartSheetsRef.current.delete(name);
    setChartSheetOrder(prev => prev.filter(id => id !== name));
    bump();
  }, [activeBookId, bump]);

  const doRenameChartSheet = useCallback(async (oldName: string, newName: string) => {
    if (!activeBookId) return;
    await api.updateChart(activeBookId, oldName, { name: newName });
    const chart = chartSheetsRef.current.get(oldName);
    if (chart) {
      chart.name = newName;
      chartSheetsRef.current.delete(oldName);
      chartSheetsRef.current.set(newName, chart);
    }
    setChartSheetOrder(prev => prev.map(id => id === oldName ? newName : id));
    bump();
  }, [activeBookId, bump]);

  const doUpdateChartSheet = useCallback(async (name: string, charts: ChartConfig[], layout: ChartLayoutItem[]) => {
    if (!activeBookId) return;
    const sheet = chartSheetsRef.current.get(name);
    if (sheet) { sheet.charts = charts; sheet.layout = layout; }
    await api.updateChart(activeBookId, name, { charts: { charts, layout } });
  }, [activeBookId]);

  // ---- View sheets ----
  const getViewSheet = useCallback((id: string) => viewSheetsRef.current.get(id), []);

  const doCreateViewSheet = useCallback(async (
    name: string,
    tableName: string,
    viewType: ViewSheet['viewType'],
    dateColumn?: string,
  ) => {
    if (!activeBookId) return;
    await api.createView(activeBookId, name, tableName, viewType, dateColumn);
    viewSheetsRef.current.set(name, { name, tableName, viewType, dateColumn });
    setViewSheetOrder(prev => [...prev, name]);
    bump();
  }, [activeBookId, bump]);

  const doDeleteViewSheet = useCallback(async (name: string) => {
    if (!activeBookId) return;
    await api.deleteView(activeBookId, name);
    viewSheetsRef.current.delete(name);
    setViewSheetOrder(prev => prev.filter(id => id !== name));
    bump();
  }, [activeBookId, bump]);

  const doRenameViewSheet = useCallback(async (oldName: string, newName: string) => {
    if (!activeBookId) return;
    await api.updateView(activeBookId, oldName, { name: newName });
    const view = viewSheetsRef.current.get(oldName);
    if (view) {
      view.name = newName;
      viewSheetsRef.current.delete(oldName);
      viewSheetsRef.current.set(newName, view);
    }
    setViewSheetOrder(prev => prev.map(id => id === oldName ? newName : id));
    bump();
  }, [activeBookId, bump]);

  const doUpdateViewSheet = useCallback(async (
    name: string,
    updates: Partial<Pick<ViewSheet, 'name' | 'tableName' | 'viewType' | 'dateColumn' | 'hideSourceTableTab'>>,
  ) => {
    if (!activeBookId) return;
    const view = viewSheetsRef.current.get(name);
    if (view) Object.assign(view, updates);
    await api.updateView(activeBookId, name, {
      ...(updates.name !== undefined ? { name: updates.name } : {}),
      ...(updates.tableName !== undefined ? { tableName: updates.tableName } : {}),
      ...(updates.viewType !== undefined ? { viewType: updates.viewType } : {}),
      ...('dateColumn' in updates ? { dateColumn: updates.dateColumn ?? null } : {}),
      ...('hideSourceTableTab' in updates ? { hideSourceTableTab: updates.hideSourceTableTab ?? false } : {}),
    });
    bump();
  }, [activeBookId, bump]);

  return {
    user,
    isLoading,
    signIn,
    signOut,

    books,
    activeBookId,
    activeBookName,
    activeBookRole,
    switchBook,
    refreshBooks,
    createBook,
    renameBook,
    deleteBook: doDeleteBook,

    tableIds: tableOrder,
    activeTableId,
    setActiveTableId,
    getSchema,
    getRows,
    setRows: (tableId: string, rows: Row[]) => { rowsRef.current.set(tableId, rows); bump(); },
    createTable: doCreateTable,
    deleteTable: doDeleteTable,
    renameTable: doRenameTable,
    renameColumn: doRenameColumn,
    reorderTables: doReorderTables,
    reorderCharts: doReorderCharts,
    reorderViews: doReorderViews,
    reorderTablesTo: doReorderTablesTo,
    reorderChartsTo: doReorderChartsTo,
    reorderViewsTo: doReorderViewsTo,
    sortedSheets,
    reorderAllSheetsTo: doReorderAllSheetsTo,
    updateSchema: doUpdateSchema,

    applyEdit,
    insertRow: doInsertRow,
    deleteRow: doDeleteRow,
    deleteRows: doDeleteRows,

    undo: doUndo,
    canUndo,

    chartSheetIds: chartSheetOrder,
    getChartSheet,
    createChartSheet: doCreateChartSheet,
    deleteChartSheet: doDeleteChartSheet,
    renameChartSheet: doRenameChartSheet,
    updateChartSheet: doUpdateChartSheet,

  viewSheetIds: viewSheetOrder,
  getViewSheet,
  createViewSheet: doCreateViewSheet,
  deleteViewSheet: doDeleteViewSheet,
  renameViewSheet: doRenameViewSheet,
  updateViewSheet: doUpdateViewSheet,

    getReferencedRow,
    getReferenceRows,
    resolveColumnPath,
    resolveColumnPathLabel,
    getColumnPaths,

    revision,
  };
}

function rewritePath(path: string, oldName: string, newName: string): string {
  return path.split('.').map(p => p === oldName ? newName : p).join('.');
}

// Rewrites a column expression that may include a date feature suffix, e.g. "col:year" or "ref.col:month"
function rewriteColumnExpr(expr: string, oldName: string, newName: string): string {
  const colonIdx = expr.indexOf(':');
  if (colonIdx === -1) return rewritePath(expr, oldName, newName);
  return rewritePath(expr.slice(0, colonIdx), oldName, newName) + expr.slice(colonIdx);
}

import type { ReportSummary, SheetOrderEntry } from '../shared/reportTypes';
import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import type { TableSchema, ColumnDef, Row, ValidationError, ChartSheet, ChartConfig, ChartLayoutItem, ViewSheet, UndoEntry, SessionUser, BookInfo } from './types';
import { INTERNAL_ROW_ID } from './types';
import * as api from './api';
import { log } from './DebugLogger';
import { applyChartValueFormat } from './chartFormat';

type RememberedAccount = Pick<SessionUser, 'id' | 'email'>;
const REMEMBERED_ACCOUNT_KEY = 'sheetable-remembered-account';

function readRememberedAccount(): RememberedAccount | null {
  try {
    const saved = localStorage.getItem(REMEMBERED_ACCOUNT_KEY);
    if (!saved) return null;
    const account = JSON.parse(saved) as Partial<RememberedAccount> | null;
    if (typeof account?.id === 'string' && account.id && typeof account.email === 'string' && account.email) {
      return { id: account.id, email: account.email };
    }
  } catch {
    // An unavailable or malformed browser preference should not block sign-in.
  }
  return null;
}

export interface UseAppStateReturn {
  // Auth
  user: SessionUser | null;
  isLoading: boolean;
  rememberedAccount: RememberedAccount | null;
  signIn: (loginHint?: string) => void;
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
  bookReady: boolean;
  bookLoadError: string | null;
  reloadBookData: () => void;
  getRequiredTables: (roots: string[]) => string[];
  hasRows: (tableId: string) => boolean;
  loadRows: (tables: string[]) => Promise<void>;
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
  sortedSheets: SheetOrderEntry[];
  reorderAllSheetsTo: (items: SheetOrderEntry[]) => void;
  updateSchema: (tableId: string, schema: TableSchema) => Promise<void>;

  // Row operations
  applyEdit: (tableId: string, rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  insertRow: (tableId: string, row: Row) => ValidationError[];
  deleteRow: (tableId: string, rowIndex: number) => Promise<ValidationError[]>;
  deleteRows: (tableId: string, rowIndices: number[]) => Promise<ValidationError[]>;

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

  reportSheets: ReportSummary[];
  updateReportTab: (bookId: string, report: ReportSummary) => void;
  removeReportTab: (bookId: string, id: string) => void;

  // Reference helpers (replacement for DataModel methods)
  getReferencedRow: (refTable: string, rowId: string) => Row | undefined;
  getReferenceRows: (refTable: string) => Row[];
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  resolveColumnPathLabel: (tableName: string, path: string) => string;
  resolveColumnPathLeafLabel: (tableName: string, path: string) => string;
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
      try {
        if (!Array.isArray(JSON.parse(value))) throw new Error('not an array');
      } catch {
        return [{ message: `"${value}" is not a valid list (must be JSON array)`, rowIndex, columnName }];
      }
      break;
  }
  return [];
}

function parseListIds(value: string): string[] {
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed.map(String) : [];
  } catch {
    return [];
  }
}

function columnReferencesRow(col: ColumnDef, cellValue: string | undefined, rowId: string): boolean {
  if (!cellValue) return false;
  if (col.type === 'reference') return cellValue === rowId;
  if (col.type === 'list' && col.listOf === 'reference') return parseListIds(cellValue).includes(rowId);
  return false;
}

function referencedTable(col: ColumnDef): string | undefined {
  if (col.type === 'reference' || (col.type === 'list' && col.listOf === 'reference')) return col.refTable;
  return undefined;
}

function missingReference(col: ColumnDef, value: string, refRows: Row[] | undefined): boolean {
  if (!value) return false;
  const ids = col.type === 'list' ? parseListIds(value) : [value];
  return ids.some(id => !refRows?.some(r => r[INTERNAL_ROW_ID] === id));
}

function validateUniqueKey(rows: Row[], _schema: TableSchema, keyValues: Record<string, string>, excludeRow: number): ValidationError[] {
  const keyColumns = Object.keys(keyValues);
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
const EMPTY_ROWS: Row[] = [];

interface PendingCell {
  confirmed: string;
  latest: symbol;
}

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
  const [rememberedAccount, setRememberedAccount] = useState(readRememberedAccount);
  const [isLoading, setIsLoading] = useState(true);
  const [books, setBooks] = useState<BookInfo[]>([]);
  const [activeBookId, setActiveBookId] = useState<string | null>(null);
  const [revision, setRevision] = useState(0);
  const [reloadCount, setReloadCount] = useState(0);
  const [loadedBook, setLoadedBook] = useState<{ id: string; reload: number } | null>(null);
  const [bookFailure, setBookFailure] = useState<{ id: string; reload: number; message: string } | null>(null);
  const bookLoadError = bookFailure?.id === activeBookId && bookFailure.reload === reloadCount ? bookFailure.message : null;
  const bookReady = activeBookId === null || (loadedBook?.id === activeBookId && loadedBook.reload === reloadCount);
  const reloadBookData = useCallback(() => {
    setReloadCount(count => count + 1);
  }, []);
  const [activeTableId, setActiveTableId] = useState<string | null>(null);
  const [tableOrder, setTableOrder] = useState<string[]>([]);
  const [chartSheetOrder, setChartSheetOrder] = useState<string[]>([]);
  const [viewSheetOrder, setViewSheetOrder] = useState<string[]>([]);
  const [reportsByBook, setReportsByBook] = useState<Record<string, ReportSummary[]>>({});
  const reportSheets = activeBookId ? reportsByBook[activeBookId] : undefined;

  const updateReportTab = useCallback((bookId: string, { id, name }: ReportSummary) => {
    setReportsByBook(current => {
      const reports = current[bookId] ?? [];
      return {
        ...current,
        [bookId]: reports.some(report => report.id === id)
          ? reports.map(report => report.id === id ? { id, name } : report)
          : [...reports, { id, name }],
      };
    });
  }, []);

  const removeReportTab = useCallback((bookId: string, id: string) => {
    setReportsByBook(current => ({
      ...current,
      [bookId]: (current[bookId] ?? []).filter(report => report.id !== id),
    }));
  }, []);

  // In-memory data store (replaces DataModel)
  const schemasRef = useRef<Map<string, TableSchema>>(new Map());
  const rowsRef = useRef<Map<string, Row[]>>(new Map());
  const chartSheetsRef = useRef<Map<string, ChartSheet>>(new Map());
  const viewSheetsRef = useRef<Map<string, ViewSheet>>(new Map());
  const undoStackRef = useRef<UndoEntry[]>([]);
  const rowIndexCacheRef = useRef<Map<string, { rows: Row[]; length: number; byId: Map<string, Row> }>>(new Map());
  const rowLoadsRef = useRef(new Map<string, Promise<void>>());
  const pendingCellsRef = useRef(new WeakMap<Row, Map<string, PendingCell>>());

  const bump = useCallback(() => setRevision(r => r + 1), []);

  // ---- Check auth on mount ----
  useEffect(() => {
    api.getMe()
      .then(u => {
        setUser(u);
        if (!u.id || !u.email) return;
        const account = { id: u.id, email: u.email };
        setRememberedAccount(account);
        try {
          localStorage.setItem(REMEMBERED_ACCOUNT_KEY, JSON.stringify(account));
        } catch {
          // The account can still be offered in this tab after sign-out.
        }
      })
      .catch(() => {})
      .finally(() => setIsLoading(false));
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
    schemasRef.current.clear();
    rowIndexCacheRef.current.clear();
    rowsRef.current = new Map();
    rowLoadsRef.current = new Map();
    chartSheetsRef.current.clear();
    viewSheetsRef.current.clear();
    undoStackRef.current = [];
    setLoadedBook(null);
    setTableOrder([]);
    setChartSheetOrder([]);
    setViewSheetOrder([]);
    setActiveTableId(null);
    bump();

    if (!activeBookId) return;
    let cancelled = false;

    (async () => {
      try {
        const [schemas, charts, views, reports] = await Promise.all([
          api.listTables(activeBookId),
          api.listCharts(activeBookId),
          api.listViews(activeBookId),
          api.listReports(activeBookId),
        ]);

        if (cancelled) return;

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
        setReportsByBook(current => ({ ...current, [activeBookId]: reports }));

        setLoadedBook({ id: activeBookId, reload: reloadCount });
        bump();
      } catch (err) {
        if (!cancelled) {
          setBookFailure({
            id: activeBookId,
            reload: reloadCount,
            message: err instanceof Error ? err.message : 'Failed to load workbook',
          });
        }
      }
    })();

    return () => {
      cancelled = true;
      rowsRef.current = new Map();
    };
  }, [activeBookId, reloadCount, bump]);

  const getRequiredTables = useCallback((roots: string[]): string[] => {
    const required = new Set<string>();
    const visit = (table: string) => {
      const schema = schemasRef.current.get(table);
      if (!schema || required.has(table)) return;
      required.add(table);
      for (const column of schema.columns) {
        const reference = referencedTable(column);
        if (reference) visit(reference);
      }
    };
    roots.forEach(visit);
    return [...required];
  }, []);

  const hasRows = useCallback((table: string) => rowsRef.current.has(table), []);

  const loadRows = useCallback(async (tables: string[]) => {
    if (!bookReady || !activeBookId) throw new Error('Workbook is still loading');
    const rowsByTable = rowsRef.current;
    const pending = rowLoadsRef.current;
    await Promise.all(tables.map(table => {
      if (rowsByTable.has(table) || !schemasRef.current.has(table)) return;
      const existing = pending.get(table);
      if (existing) return existing;
      const request = api.listRows(activeBookId, table).then(rows => {
        if (rowsRef.current === rowsByTable && !rowsByTable.has(table) && schemasRef.current.has(table)) {
          rowsByTable.set(table, rows);
          bump();
        }
      }).finally(() => pending.delete(table));
      pending.set(table, request);
      return request;
    }));
    if (rowsRef.current !== rowsByTable) throw new Error('Workbook changed while rows were loading');
  }, [activeBookId, bookReady, bump]);

  const activeBookName = useMemo(() => {
    return books.find(b => b.id === activeBookId)?.name ?? null;
  }, [books, activeBookId]);

  const activeBookRole = useMemo(() => {
    return (books.find(b => b.id === activeBookId)?.role ?? null) as 'owner' | 'editor' | 'viewer' | null;
  }, [books, activeBookId]);

  // ---- Auth ----
  const signIn = useCallback((loginHint?: string) => {
    // Save current hash route so we can redirect back after OAuth
    const hash = window.location.hash;
    if (hash && hash !== '#/' && hash !== '#') {
      localStorage.setItem('sheetable-post-login-redirect', hash);
    } else {
      localStorage.removeItem('sheetable-post-login-redirect');
    }
    window.location.href = api.loginUrl(loginHint);
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
    if (!rows) return undefined;
    let entry = rowIndexCacheRef.current.get(refTable);
    if (entry && entry.rows === rows && entry.length === rows.length) {
      const hit = entry.byId.get(rowId);
      if (hit && hit[INTERNAL_ROW_ID] === rowId) return hit;
      if (!hit) return rows.find(r => r[INTERNAL_ROW_ID] === rowId);
    }
    const byId = new Map<string, Row>();
    for (const row of rows) byId.set(row[INTERNAL_ROW_ID], row);
    entry = { rows, length: rows.length, byId };
    rowIndexCacheRef.current.set(refTable, entry);
    return byId.get(rowId);
  }, []);

  const getReferenceRows = useCallback((refTable: string): Row[] => {
    return rowsRef.current.get(refTable) ?? EMPTY_ROWS;
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

  // Resolve the effective leaf display label for a path, following terminal reference
  // columns into their configured display columns.
  const resolveColumnPathLeafLabel = useCallback((tableName: string, path: string): string => {
    const MAX_DEPTH = 8;

    const walk = (table: string, p: string, depth: number): string => {
      if (!p || depth > MAX_DEPTH) return '';
      const parts = p.split('.').filter(Boolean);
      if (parts.length === 0) return '';

      const schema = schemasRef.current.get(table);
      if (!schema) return parts[parts.length - 1];

      const [head, ...rest] = parts;
      const col = schema.columns.find(c => c.name === head);
      if (!col) return head;

      if (rest.length > 0) {
        if (col.type === 'reference' && col.refTable) {
          return walk(col.refTable, rest.join('.'), depth + 1);
        }
        return col.displayName || col.name;
      }

      if (col.type === 'reference' && col.refTable) {
        const displayCols = col.refDisplayColumns?.length
          ? col.refDisplayColumns
          : col.refSearchColumns?.length
            ? col.refSearchColumns
            : [];
        if (displayCols.length > 0) {
          const labels = displayCols.map(dp => walk(col.refTable!, dp, depth + 1)).filter(Boolean);
          if (labels.length > 0) return labels.join(' · ');
        }
      }

      return col.displayName || col.name;
    };

    return walk(tableName, path, 0);
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
  const getRows = useCallback((tableId: string) => rowsRef.current.get(tableId) ?? EMPTY_ROWS, []);

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
        if (referencedTable(col) === tableId) {
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
    if (!bookReady) return [];
    const reportOrder = reportSheets?.map(report => report.id) ?? [];
    const book = books.find(b => b.id === activeBookId);
    const raw = book?.sheet_order;
    if (raw) {
      try {
        const parsed: SheetOrderEntry[] = JSON.parse(raw);
        const existing = new Set<string>();
        const result: SheetOrderEntry[] = [];
        for (const item of parsed) {
          const key = `${item.type}:${item.name}`;
          if (
            (item.type === 'table' && tableOrder.includes(item.name)) ||
            (item.type === 'chart' && chartSheetOrder.includes(item.name)) ||
            (item.type === 'view' && viewSheetOrder.includes(item.name)) ||
            (item.type === 'report' && reportOrder.includes(item.name))
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
        for (const name of reportOrder) {
          if (!existing.has(`report:${name}`)) result.push({ type: 'report', name });
        }
        return result;
      } catch { /* fall through */ }
    }
    return [
      ...tableOrder.map(name => ({ type: 'table' as const, name })),
      ...chartSheetOrder.map(name => ({ type: 'chart' as const, name })),
      ...viewSheetOrder.map(name => ({ type: 'view' as const, name })),
      ...reportOrder.map(name => ({ type: 'report' as const, name })),
    ];
  }, [books, activeBookId, tableOrder, chartSheetOrder, viewSheetOrder, reportSheets, bookReady]);

  const doReorderAllSheetsTo = useCallback((items: SheetOrderEntry[]) => {
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

  const persistCellEdit = useCallback((
    table: string, row: Row, column: string, value: string, previous: string, entry?: UndoEntry,
  ) => {
    if (!activeBookId) return;
    const rowId = row[INTERNAL_ROW_ID];
    const queueKey = JSON.stringify([activeBookId, table, rowId]);
    const rowsByTable = rowsRef.current;
    const undoStack = undoStackRef.current;
    const cells = pendingCellsRef.current.get(row) ?? new Map<string, PendingCell>();
    pendingCellsRef.current.set(row, cells);
    const version = Symbol();
    const cell = cells.get(column) ?? { confirmed: previous, latest: version };
    cell.latest = version;
    cells.set(column, cell);

    enqueueWrite(queueKey, async () => {
      if (entry) entry.oldValue = cell.confirmed;
      try {
        await api.updateRow(activeBookId, table, rowId, { [column]: value });
        cell.confirmed = value;
      } catch (error) {
        console.error('Server write failed:', error);
        if (entry) {
          const index = undoStack.indexOf(entry);
          if (index !== -1) undoStack.splice(index, 1);
        }
      } finally {
        if (cell.latest === version) {
          cells.delete(column);
          row[column] = cell.confirmed;
        }
        if (rowsRef.current === rowsByTable) bump();
      }
    });
  }, [activeBookId, bump]);

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
    const editRefTable = referencedTable(col);
    if (editRefTable && newValue !== '' && missingReference(col, newValue, rowsRef.current.get(editRefTable))) {
      return [{ message: `Referenced row not found in "${editRefTable}"`, rowIndex }];
    }

    // Apply optimistic update
    const oldValue = rows[rowIndex][columnName] ?? '';
    const rowId = rows[rowIndex][INTERNAL_ROW_ID];
    rows[rowIndex][columnName] = newValue;

    const entry: UndoEntry = { type: 'update', tableId, rowId, column: columnName, oldValue, newValue };
    undoStackRef.current.push(entry);
    log('applyEdit:', tableId, rowId, columnName, oldValue, '->', newValue);
    persistCellEdit(tableId, rows[rowIndex], columnName, newValue, oldValue, entry);

    bump();
    return [];
  }, [bump, persistCellEdit]);

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

      const insertRefTable = referencedTable(col);
      if (insertRefTable && value !== '' && missingReference(col, value, rowsRef.current.get(insertRefTable))) {
        return [{ message: `Referenced row not found in "${insertRefTable}"`, rowIndex: newRowIndex }];
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
    const maxId = rows.reduce((max, row) => Math.max(max, Number(row[INTERNAL_ROW_ID]) || 0), 0);
    const rowId = String(maxId + 1);
    completeRow[INTERNAL_ROW_ID] = rowId;
    for (const col of schema.columns) {
      completeRow[col.name] = row[col.name] ?? '';
    }

    rows.push(completeRow);
    const entry: UndoEntry = { type: 'insert', tableId, rowId, row: { ...completeRow } };
    const undoStack = undoStackRef.current;
    undoStack.push(entry);
    log('insertRow:', tableId, rowId);

    // Async server write
    if (activeBookId) {
      const bookId = activeBookId;
      api.insertRow(bookId, tableId, completeRow).catch(err => {
        console.error('Server insert failed, removing row:', err);
        const currentRows = rowsRef.current.get(tableId);
        if (currentRows) {
          const idx = currentRows.indexOf(completeRow);
          if (idx >= 0) {
            currentRows.splice(idx, 1);
            const undoIndex = undoStack.indexOf(entry);
            if (undoIndex >= 0) undoStack.splice(undoIndex, 1);
            bump();
          }
        }
      });
    }

    bump();
    return [];
  }, [activeBookId, bump]);

  const doDeleteRows = useCallback(async (tableId: string, rowIndices: number[]): Promise<ValidationError[]> => {
    const rows = rowsRef.current.get(tableId);
    if (!schemasRef.current.has(tableId) || !rows) return [{ message: 'Table not found', rowIndex: 0 }];
    const rowsByTable = rowsRef.current;
    const selected = [...new Set(rowIndices)].map(index => ({ index, row: rows[index] }));
    const references = [...schemasRef.current].flatMap(([name, schema]) =>
      schema.columns.filter(col => referencedTable(col) === tableId).map(column => ({ name, column })),
    );
    try {
      if (activeBookId && references.length) await loadRows([...new Set(references.map(ref => ref.name))]);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to load reference rows';
      return selected.map(({ index }) => ({ message, rowIndex: index }));
    }

    const errors: ValidationError[] = [];
    const valid: { index: number; row: Row }[] = [];
    const presentRows = new Set(rowsByTable.get(tableId));
    for (const { index, row } of selected) {
      if (!row || !presentRows.has(row)) {
        errors.push({ message: 'Row not found', rowIndex: index });
        continue;
      }
      const rowId = row[INTERNAL_ROW_ID];
      const reference = references.find(({ name, column }) =>
        rowsByTable.get(name)?.some(other => columnReferencesRow(column, other[column.name], rowId)),
      );
      if (reference) {
        errors.push({ message: `Cannot delete: row referenced by "${reference.name}"`, rowIndex: index });
      } else {
        valid.push({ index, row });
      }
    }

    if (valid.length === 0) return errors;
    const rowIds = new Set(valid.map(({ row }) => row[INTERNAL_ROW_ID]));
    try {
      if (activeBookId) {
        if (rowIds.size === 1) await api.deleteRow(activeBookId, tableId, [...rowIds][0]);
        else await api.deleteRows(activeBookId, tableId, [...rowIds]);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Delete failed';
      return [...errors, ...valid.map(({ index }) => ({ message, rowIndex: index }))];
    }

    const currentRows = rowsByTable.get(tableId);
    if (rowsRef.current === rowsByTable && currentRows) {
      for (let i = currentRows.length - 1; i >= 0; i--) {
        const row = currentRows[i];
        if (rowIds.has(row[INTERNAL_ROW_ID])) {
          currentRows.splice(i, 1);
          undoStackRef.current.push({ type: 'delete', tableId, rowId: row[INTERNAL_ROW_ID], row: { ...row } });
        }
      }
      bump();
    }
    return errors;
  }, [activeBookId, bump, loadRows]);

  const doDeleteRow = useCallback((tableId: string, rowIndex: number) =>
    doDeleteRows(tableId, [rowIndex]), [doDeleteRows]);

  const canUndo = undoStackRef.current.length > 0;

  const doUndo = useCallback((): ValidationError[] => {
    const entry = undoStackRef.current.pop();
    if (!entry) return [];

    const rows = rowsRef.current.get(entry.tableId);
    if (!rows) return [];

    if (entry.type === 'update' && entry.column && entry.oldValue !== undefined) {
      const row = rows.find(r => r[INTERNAL_ROW_ID] === entry.rowId);
      if (row) {
        const previous = row[entry.column] ?? '';
        row[entry.column] = entry.oldValue;
        persistCellEdit(entry.tableId, row, entry.column, entry.oldValue, previous);
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
  }, [activeBookId, bump, persistCellEdit]);

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
    rememberedAccount,
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

    bookReady,
    bookLoadError,
    reloadBookData,
    getRequiredTables,
    hasRows,
    loadRows,
    tableIds: bookReady ? tableOrder : [],
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

    reportSheets: reportSheets ?? [],
    updateReportTab,
    removeReportTab,

    getReferencedRow,
    getReferenceRows,
    resolveColumnPath,
    resolveColumnPathLabel,
    resolveColumnPathLeafLabel,
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

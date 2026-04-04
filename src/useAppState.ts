import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { DataModel } from './dataModel';
import type { TableSchema, Row, Transaction, ValidationError, ChartSheet } from './types';
import { INTERNAL_ROW_ID } from './types';
import { csvToRows, rowsToCSV } from './csv';
import type { ProjectConfig } from './config';
import { serializeConfig, parseConfig, serializeBooksConfig, parseBooksConfig } from './config';
import * as drive from './drive';

export interface WorkbookInfo {
  id: string;
  name: string;
}

export interface WorkbookInvitePayload {
  workbookId: string;
  workbookName: string;
  emailAddress: string;
  role: 'reader' | 'writer';
}

interface LocalWorkbookSnapshot {
  name: string;
  tables: Array<{ schema: TableSchema; rows: Row[] }>;
  tableOrder: string[];
  activeTableId: string | null;
  chartSheets?: ChartSheet[];
}

interface LocalBooksStorage {
  activeWorkbookId: string;
  workbooks: Record<string, LocalWorkbookSnapshot>;
}

const LOCAL_DEFAULT_WORKBOOK: WorkbookInfo = {
  id: 'local-default',
  name: 'Untitled',
};

const LOCAL_BOOKS_STORAGE_KEY = 'sheetable_local_books_v1';
const DRIVE_BOOKS_CONFIG_FILENAME = 'sheetable.books.json';

function encodeInvitePayload(payload: WorkbookInvitePayload): string {
  const json = JSON.stringify(payload);
  const bytes = new TextEncoder().encode(json);
  let binary = '';
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function decodeInvitePayload(token: string): WorkbookInvitePayload {
  const normalized = token.replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
  const json = new TextDecoder().decode(bytes);
  const parsed = JSON.parse(json) as Partial<WorkbookInvitePayload>;

  if (!parsed.workbookId || !parsed.workbookName || !parsed.emailAddress || !parsed.role) {
    throw new Error('Invalid invite token.');
  }
  if (parsed.role !== 'reader' && parsed.role !== 'writer') {
    throw new Error('Invalid invite role.');
  }

  return {
    workbookId: parsed.workbookId,
    workbookName: parsed.workbookName,
    emailAddress: parsed.emailAddress,
    role: parsed.role,
  };
}

function loadLocalBooksStorage(): LocalBooksStorage | null {
  try {
    const raw = localStorage.getItem(LOCAL_BOOKS_STORAGE_KEY);
    if (!raw) return null;
    return JSON.parse(raw) as LocalBooksStorage;
  } catch {
    return null;
  }
}

function saveLocalBooksStorage(storage: LocalBooksStorage): void {
  localStorage.setItem(LOCAL_BOOKS_STORAGE_KEY, JSON.stringify(storage));
}

function toErrorMessage(err: unknown): string {
  if (err instanceof Error && err.message) return err.message;
  const maybe = err as { result?: { error?: { message?: string } }; body?: string; message?: string };
  return maybe?.result?.error?.message || maybe?.message || maybe?.body || 'Unknown error';
}

export interface UseAppStateReturn {
  // Data
  model: DataModel;
  tableIds: string[];
  activeTableId: string | null;
  setActiveTableId: (id: string | null) => void;

  // Table operations
  createTable: (schema: TableSchema, rows?: Row[]) => void;
  deleteTable: (tableId: string, alsoDeleteFromDrive?: boolean) => void;
  renameTable: (oldName: string, newName: string) => void;
  renameColumn: (tableId: string, oldName: string, newName: string) => void;
  reorderTables: (fromIndex: number, toIndex: number) => void;
  updateSchema: (tableId: string, schema: TableSchema) => void;
  getRows: (tableId: string) => Row[];
  getSchema: (tableId: string) => TableSchema | undefined;
  listWorkbookCsvFiles: () => Promise<Array<{ id: string; name: string }>>;
  loadTableFromCsvFile: (tableId: string, fileId: string) => Promise<void>;
  renameTableCsvFile: (tableId: string, nextFileName: string) => Promise<void>;

  // Transaction
  applyEdit: (tableId: string, rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  insertRow: (tableId: string, row: Row) => ValidationError[];
  deleteRow: (tableId: string, rowIndex: number) => ValidationError[];
  undo: () => ValidationError[];
  canUndo: boolean;

  // Dirty state
  isDirty: (tableId: string) => boolean;
  isAnyDirty: () => boolean;

  // Drive
  isSignedIn: boolean;
  folderId: string | null;
  folderName: string | null;
  workbooks: WorkbookInfo[];
  createWorkbook: (name: string) => Promise<string | null>;
  renameWorkbook: (workbookId: string, name: string) => Promise<void>;
  deleteWorkbook: (workbookId: string) => Promise<string | null>;
  switchWorkbook: (workbookId: string) => Promise<void>;
  shareWorkbook: (workbookId: string, emailAddress: string, role?: 'reader' | 'writer') => Promise<void>;
  createWorkbookInviteLink: (workbookId: string, emailAddress: string, role?: 'reader' | 'writer') => Promise<string>;
  acceptWorkbookInvite: (token: string) => Promise<WorkbookInfo>;
  signIn: () => void;
  signOut: () => void;
  isSaving: boolean;
  lastSaved: Date | null;
  initializeDrive: (clientId: string) => Promise<void>;
  driveReady: boolean;
  isConnecting: boolean;
  userInfo: drive.UserInfo | null;

  // Chart sheets
  chartSheetIds: string[];
  getChartSheet: (id: string) => ChartSheet | undefined;
  createChartSheet: (name: string) => void;
  deleteChartSheet: (name: string) => void;
  renameChartSheet: (oldName: string, newName: string) => void;
  updateChartSheet: (name: string, charts: unknown[]) => void;
  setChartSheetTable: (name: string, tableName: string) => void;
  setChartSheetMode: (name: string, mode: 'edit' | 'display') => void;

  // Revision counter to trigger re-renders
  revision: number;
}

export function useAppState(): UseAppStateReturn {
  const modelRef = useRef(new DataModel());
  const [revision, setRevision] = useState(0);
  const [activeTableId, setActiveTableId] = useState<string | null>(null);
  const [tableOrder, setTableOrder] = useState<string[]>([]);
  const [isSignedIn, setIsSignedIn] = useState(false);
  const [folderId, setFolderId] = useState<string | null>(LOCAL_DEFAULT_WORKBOOK.id);
  const [folderName, setFolderName] = useState<string | null>(LOCAL_DEFAULT_WORKBOOK.name);
  const [workbooks, setWorkbooks] = useState<WorkbookInfo[]>([LOCAL_DEFAULT_WORKBOOK]);
  const [driveReady, setDriveReady] = useState(false);
  const [lastSaved, setLastSaved] = useState<Date | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);
  const [userInfo, setUserInfo] = useState<drive.UserInfo | null>(null);

  // File ID tracking for Drive
  const fileIdsRef = useRef<Map<string, string>>(new Map()); // tableId -> driveFileId
  const configFileIdRef = useRef<string | null>(null);
  const booksConfigFileIdRef = useRef<string | null>(null);
  const configDirtyRef = useRef(false);
  const revisionRef = useRef(0);
  const undoStackRef = useRef<Transaction[]>([]);
  const chartSheetsRef = useRef<Map<string, ChartSheet>>(new Map());
  const [chartSheetOrder, setChartSheetOrder] = useState<string[]>([]);
  const rootFolderIdRef = useRef<string | null>(null);
  const localWorkbooksRef = useRef<Map<string, LocalWorkbookSnapshot>>(
    new Map([[LOCAL_DEFAULT_WORKBOOK.id, { name: LOCAL_DEFAULT_WORKBOOK.name, tables: [], tableOrder: [], activeTableId: null }]])
  );
  const localActiveWorkbookIdRef = useRef<string>(LOCAL_DEFAULT_WORKBOOK.id);
  const initializedLocalRef = useRef(false);

  const model = modelRef.current;

  const bump = useCallback(() => setRevision(r => {
    const next = r + 1;
    revisionRef.current = next;
    return next;
  }), []);

  const clearModel = useCallback(() => {
    const existing = [...model.getTableIds()];
    for (const id of existing) {
      model.deleteTable(id);
    }
    fileIdsRef.current.clear();
    configFileIdRef.current = null;
    configDirtyRef.current = false;
    undoStackRef.current = [];
    chartSheetsRef.current.clear();
    setChartSheetOrder([]);
    setTableOrder([]);
    setActiveTableId(null);
  }, [model]);

  const snapshotLocalWorkbook = useCallback((workbookId: string, workbookName: string) => {
    const tableIds = model.getTableIds();
    const tables = tableIds
      .map(tableId => model.getTable(tableId))
      .filter((t): t is NonNullable<typeof t> => !!t)
      .map(table => ({
        schema: JSON.parse(JSON.stringify(table.schema)) as TableSchema,
        rows: table.rows.map(row => ({ ...row })),
      }));

    localWorkbooksRef.current.set(workbookId, {
      name: workbookName,
      tables,
      tableOrder: [...tableOrder],
      activeTableId,
      chartSheets: [...chartSheetsRef.current.values()].map(cs => ({ ...cs, charts: [...cs.charts] })),
    });
  }, [activeTableId, model, tableOrder]);

  const loadLocalWorkbook = useCallback((workbookId: string, workbookName: string) => {
    clearModel();
    setFolderId(workbookId);
    setFolderName(workbookName);
    localActiveWorkbookIdRef.current = workbookId;

    const snapshot = localWorkbooksRef.current.get(workbookId);
    if (!snapshot) {
      localWorkbooksRef.current.set(workbookId, {
        name: workbookName,
        tables: [],
        tableOrder: [],
        activeTableId: null,
      });
      bump();
      return;
    }

    for (const table of snapshot.tables) {
      model.createTable(
        JSON.parse(JSON.stringify(table.schema)) as TableSchema,
        table.rows.map(r => ({ ...r })),
      );
      model.markSaved(table.schema.name);
    }

    setTableOrder([...snapshot.tableOrder]);
    setActiveTableId(snapshot.activeTableId);
    if (snapshot.chartSheets) {
      for (const cs of snapshot.chartSheets) {
        chartSheetsRef.current.set(cs.name, { ...cs, charts: [...cs.charts] });
      }
      setChartSheetOrder(snapshot.chartSheets.map(cs => cs.name));
    }
    bump();
  }, [bump, clearModel, model]);

  const persistLocalBooks = useCallback(() => {
    const obj: Record<string, LocalWorkbookSnapshot> = {};
    for (const [id, snap] of localWorkbooksRef.current.entries()) {
      obj[id] = snap;
    }
    saveLocalBooksStorage({
      activeWorkbookId: localActiveWorkbookIdRef.current,
      workbooks: obj,
    });
  }, []);

  const persistDriveBooksConfig = useCallback(async (books: WorkbookInfo[]) => {
    const rootId = rootFolderIdRef.current;
    if (!rootId) return;
    const payload = serializeBooksConfig({ books: books.map(b => ({ id: b.id, name: b.name })) });

    if (booksConfigFileIdRef.current) {
      await drive.updateFile(booksConfigFileIdRef.current, payload, 'application/json');
      return;
    }

    const rootFiles = await drive.listFilesInFolder(rootId);
    const existing = rootFiles.find(f => f.name === DRIVE_BOOKS_CONFIG_FILENAME);
    if (existing) {
      booksConfigFileIdRef.current = existing.id;
      await drive.updateFile(existing.id, payload, 'application/json');
      return;
    }

    booksConfigFileIdRef.current = await drive.createFile(DRIVE_BOOKS_CONFIG_FILENAME, payload, rootId, 'application/json');
  }, []);

  const loadWorkbook = useCallback(async (workbookId: string, workbookName: string) => {
    clearModel();
    setFolderId(workbookId);
    setFolderName(workbookName);

    const files = await drive.listFilesInFolder(workbookId);
    const configFile = files.find(f => f.name === 'sheetable.json');
    if (!configFile) {
      bump();
      return;
    }

    configFileIdRef.current = configFile.id;
    const configText = await drive.downloadFile(configFile.id);
    const config = parseConfig(configText);

    const loadedTables = await Promise.all(config.tables.map(async (rawSchema) => {
      const schema: TableSchema = { ...rawSchema };
      const csvFileId = schema.csvFileId?.trim();
      const legacyCsvFileName = schema.csvFileName?.trim();
      const fallbackCsvFileName = legacyCsvFileName || `${schema.name}.csv`;
      const csvFile = (csvFileId ? files.find(f => f.id === csvFileId) : undefined)
        ?? files.find(f => f.name === fallbackCsvFileName);
      let rows: Row[] = [];
      if (csvFile) {
        schema.csvFileId = csvFile.id;
        fileIdsRef.current.set(schema.name, csvFile.id);
        const csvText = await drive.downloadFile(csvFile.id);
        rows = csvToRows(csvText, schema);
      }

      return {
        schema,
        rows,
        needsConfigMigration: (!csvFileId && !!csvFile) || !!legacyCsvFileName,
      };
    }));

    for (const loadedTable of loadedTables) {
      model.createTable(loadedTable.schema, loadedTable.rows);
      model.markSaved(loadedTable.schema.name);
      if (loadedTable.needsConfigMigration) {
        configDirtyRef.current = true;
      }
    }

    setTableOrder(config.tables.map(t => t.name));

    // Load chart sheets
    if (config.chartSheets) {
      for (const cs of config.chartSheets) {
        chartSheetsRef.current.set(cs.name, { ...cs });
      }
      setChartSheetOrder(config.chartSheets.map(cs => cs.name));
    }

    if (config.tables.length > 0) {
      setActiveTableId(config.tables[0].name);
    }
    bump();
  }, [bump, clearModel, model]);

  const refreshWorkbooks = useCallback(async () => {
    const rootId = rootFolderIdRef.current;
    if (!rootId) return [] as WorkbookInfo[];

    const rootFiles = await drive.listFilesInFolder(rootId);
    const booksConfigFile = rootFiles.find(f => f.name === DRIVE_BOOKS_CONFIG_FILENAME);
    let next: WorkbookInfo[] = [];

    if (booksConfigFile) {
      booksConfigFileIdRef.current = booksConfigFile.id;
      try {
        const configText = await drive.downloadFile(booksConfigFile.id);
        const parsed = parseBooksConfig(configText);
        next = (parsed.books ?? [])
          .filter(b => !!b?.id && !!b?.name)
          .map(b => ({ id: b.id, name: b.name }));
      } catch (err) {
        console.warn('Invalid books config, rebuilding:', err);
      }
    }

    // One-time migration when config is missing/invalid: seed from existing subfolders.
    if (next.length === 0) {
      const folders = await drive.listSubfolders(rootId);
      if (folders.length > 0) {
        next = folders;
      } else {
        const defaultId = await drive.findOrCreateSubfolder('Untitled', rootId);
        next = [{ id: defaultId, name: 'Untitled' }];
      }
      await persistDriveBooksConfig(next);
    }

    setWorkbooks(next);
    return next;
  }, [persistDriveBooksConfig]);

  const createTable = useCallback((schema: TableSchema, rows?: Row[]) => {
    model.createTable(schema, rows);
    setTableOrder(prev => (prev.includes(schema.name) ? prev : [...prev, schema.name]));
    setActiveTableId(schema.name);
    configDirtyRef.current = true;
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const deleteTable = useCallback((tableId: string, alsoDeleteFromDrive?: boolean) => {
    if (alsoDeleteFromDrive) {
      const fileId = fileIdsRef.current.get(tableId);
      if (fileId) {
        drive.deleteFile(fileId).catch(err => console.error('Failed to delete from Drive:', err));
        fileIdsRef.current.delete(tableId);
      }
    }
    model.deleteTable(tableId);
    setTableOrder(prev => prev.filter(id => id !== tableId));
    configDirtyRef.current = true;
    setActiveTableId(prev => prev === tableId ? null : prev);
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const renameTable = useCallback((oldName: string, newName: string) => {
    model.renameTable(oldName, newName);
    setTableOrder(prev => prev.map(id => id === oldName ? newName : id));
    setActiveTableId(prev => prev === oldName ? newName : prev);
    configDirtyRef.current = true;
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const renameColumn = useCallback((tableId: string, oldName: string, newName: string) => {
    model.renameColumn(tableId, oldName, newName);
    configDirtyRef.current = true;
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const reorderTables = useCallback((fromIndex: number, toIndex: number) => {
    if (fromIndex === toIndex) return;
    setTableOrder(prev => {
      if (fromIndex < 0 || toIndex < 0 || fromIndex >= prev.length || toIndex >= prev.length) {
        return prev;
      }
      const next = [...prev];
      const [moved] = next.splice(fromIndex, 1);
      next.splice(toIndex, 0, moved);
      return next;
    });
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const updateSchema = useCallback((tableId: string, schema: TableSchema) => {
    model.updateSchema(tableId, schema);
    configDirtyRef.current = true;
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const getRows = useCallback((tableId: string): Row[] => {
    return [...(model.getTable(tableId)?.rows ?? [])];
  }, [model]);

  const getSchema = useCallback((tableId: string): TableSchema | undefined => {
    return model.getTable(tableId)?.schema;
  }, [model]);

  const listWorkbookCsvFiles = useCallback(async (): Promise<Array<{ id: string; name: string }>> => {
    if (!isSignedIn || !folderId) {
      const rows = model.getTableIds().map(id => {
        const schema = model.getTable(id)?.schema;
        return {
          id: schema?.csvFileId?.trim() || '',
          name: schema?.csvFileName?.trim() || `${id}.csv`,
        };
      });
      const deduped = new Map<string, { id: string; name: string }>();
      for (const row of rows) {
        if (!deduped.has(row.name)) deduped.set(row.name, row);
      }
      return Array.from(deduped.values()).sort((a, b) => a.name.localeCompare(b.name));
    }

    const files = await drive.listFilesInFolder(folderId);
    return files
      .filter(f => f.name.toLowerCase().endsWith('.csv'))
      .map(f => ({ id: f.id, name: f.name }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [folderId, isSignedIn, model]);

  const loadTableFromCsvFile = useCallback(async (tableId: string, fileId: string): Promise<void> => {
    const table = model.getTable(tableId);
    if (!table) {
      throw new Error(`Table "${tableId}" not found.`);
    }

    const csvText = await drive.downloadFile(fileId);
    const rows = csvToRows(csvText, table.schema);
    const nextSchema: TableSchema = {
      ...table.schema,
      csvFileId: fileId,
    };

    model.deleteTable(tableId);
    model.createTable(nextSchema, rows);
    model.markSaved(tableId);

    fileIdsRef.current.set(tableId, fileId);
    configDirtyRef.current = true;
    undoStackRef.current = [];
    bump();
  }, [bump, model]);

  const renameTableCsvFile = useCallback(async (tableId: string, nextFileName: string): Promise<void> => {
    const table = model.getTable(tableId);
    if (!table) {
      throw new Error(`Table "${tableId}" not found.`);
    }

    const fileId = table.schema.csvFileId?.trim() || fileIdsRef.current.get(tableId);
    if (!fileId) {
      throw new Error('This table is not bound to a Drive CSV file yet.');
    }

    const trimmed = nextFileName.trim();
    if (!trimmed) {
      throw new Error('CSV file name cannot be empty.');
    }
    if (!trimmed.toLowerCase().endsWith('.csv')) {
      throw new Error('CSV file name must end with .csv.');
    }

    await drive.renameFile(fileId, trimmed);
  }, [model]);

  // --- Chart sheet CRUD ---
  const chartSheetIds = useMemo(() => {
    void revision; // depend on revision for re-render
    return [...chartSheetOrder];
  }, [chartSheetOrder, revision]);

  const getChartSheet = useCallback((name: string): ChartSheet | undefined => {
    return chartSheetsRef.current.get(name);
  }, []);

  const createChartSheet = useCallback((name: string) => {
    if (chartSheetsRef.current.has(name)) return;
    const defaultTableName = activeTableId ?? model.getTableIds()[0] ?? undefined;
    const cs: ChartSheet = { name, tableName: defaultTableName, mode: 'edit', charts: [] };
    chartSheetsRef.current.set(name, cs);
    setChartSheetOrder(prev => [...prev, name]);
    configDirtyRef.current = true;
    bump();
  }, [activeTableId, bump, model]);

  const deleteChartSheet = useCallback((name: string) => {
    chartSheetsRef.current.delete(name);
    setChartSheetOrder(prev => prev.filter(n => n !== name));
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const renameChartSheet = useCallback((oldName: string, newName: string) => {
    const trimmed = newName.trim();
    if (!trimmed || oldName === trimmed) return;
    const cs = chartSheetsRef.current.get(oldName);
    if (!cs) return;
    chartSheetsRef.current.delete(oldName);
    cs.name = trimmed;
    chartSheetsRef.current.set(trimmed, cs);
    setChartSheetOrder(prev => prev.map(n => n === oldName ? trimmed : n));
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const updateChartSheet = useCallback((name: string, charts: unknown[]) => {
    const cs = chartSheetsRef.current.get(name);
    if (!cs) return;
    cs.charts = charts;
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const setChartSheetTable = useCallback((name: string, tableName: string) => {
    const cs = chartSheetsRef.current.get(name);
    if (!cs) return;
    if (cs.tableName === tableName) return;
    cs.tableName = tableName;
    // Reset chart specs when source table changes to avoid stale field bindings.
    cs.charts = [];
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const setChartSheetMode = useCallback((name: string, mode: 'edit' | 'display') => {
    const cs = chartSheetsRef.current.get(name);
    if (!cs) return;
    if (cs.mode === mode) return;
    cs.mode = mode;
    configDirtyRef.current = true;
    bump();
  }, [bump]);

  const applyEdit = useCallback((tableId: string, rowIndex: number, columnName: string, newValue: string): ValidationError[] => {
    const table = model.getTable(tableId);
    const oldValue = table?.rows[rowIndex]?.[columnName] ?? '';
    const tx: Transaction = {
      id: model.nextTransactionId(),
      tableId,
      type: 'update',
      rowIndex,
      columnName,
      oldValue,
      newValue,
      timestamp: Date.now(),
    };
    const errors = model.applyTransaction(tx);
    if (errors.length === 0) {
      undoStackRef.current.push({
        id: model.nextTransactionId(),
        tableId,
        type: 'update',
        rowIndex,
        columnName,
        newValue: oldValue,
        timestamp: Date.now(),
      });
      bump();
    }
    return errors;
  }, [model, bump]);

  const insertRow = useCallback((tableId: string, row: Row): ValidationError[] => {
    const tx: Transaction = {
      id: model.nextTransactionId(),
      tableId,
      type: 'insert',
      row,
      timestamp: Date.now(),
    };
    const errors = model.applyTransaction(tx);
    if (errors.length === 0) {
      const table = model.getTable(tableId);
      const inserted = table?.rows[table.rows.length - 1];
      const insertedRowId = inserted?.[INTERNAL_ROW_ID];
      if (insertedRowId) {
        undoStackRef.current.push({
          id: model.nextTransactionId(),
          tableId,
          type: 'delete',
          rowId: insertedRowId,
          timestamp: Date.now(),
        });
      }
      bump();
    }
    return errors;
  }, [model, bump]);

  const deleteRow = useCallback((tableId: string, rowIndex: number): ValidationError[] => {
    const table = model.getTable(tableId);
    const deletedRow = table?.rows[rowIndex] ? { ...table.rows[rowIndex] } : undefined;
    const tx: Transaction = {
      id: model.nextTransactionId(),
      tableId,
      type: 'delete',
      rowIndex,
      timestamp: Date.now(),
    };
    const errors = model.applyTransaction(tx);
    if (errors.length === 0) {
      if (deletedRow) {
        undoStackRef.current.push({
          id: model.nextTransactionId(),
          tableId,
          type: 'insert',
          row: deletedRow,
          timestamp: Date.now(),
        });
      }
      bump();
    }
    return errors;
  }, [model, bump]);

  const undo = useCallback((): ValidationError[] => {
    const tx = undoStackRef.current.pop();
    if (!tx) return [];
    const errors = model.applyTransaction(tx);
    if (errors.length > 0) {
      // Put it back if undo failed so user can retry after fixing dependencies.
      undoStackRef.current.push(tx);
      return errors;
    }
    bump();
    return [];
  }, [model, bump]);

  const isDirty = useCallback((tableId: string): boolean => {
    return model.isDirty(tableId);
  }, [model]);

  const isAnyDirty = useCallback((): boolean => {
    return model.getTableIds().some(id => model.isDirty(id));
  }, [model]);

  // Connect to root "sheetable" folder, then load a workbook subfolder
  const connectToFolder = useCallback(async () => {
    setIsConnecting(true);
    try {
      const rootFolder = await drive.findOrCreateFolder();
      rootFolderIdRef.current = rootFolder.id;

      const folders = await refreshWorkbooks();
      if (folders.length > 0) {
        await loadWorkbook(folders[0].id, folders[0].name);
      }
    } finally {
      setIsConnecting(false);
    }
  }, [loadWorkbook, refreshWorkbooks]);

  // Drive integration
  const initializeDrive = useCallback(async (clientId: string) => {
    drive.setClientId(clientId);
    await drive.initGapi();

    const onSignedIn = async (expiresIn?: number) => {
      setIsSignedIn(true);
      // Fetch and save user info
      const info = await drive.getUserInfo();
      setUserInfo(info);
      if (expiresIn) {
        drive.saveAuth(drive.getAccessToken()!, expiresIn, info);
      }
      // Auto-connect to sheetable folder
      try {
        await connectToFolder();
      } catch (err) {
        console.error('Failed to connect to sheetable folder:', err);
      }
    };

    await drive.initTokenClient(async (expiresIn: number) => {
      await onSignedIn(expiresIn);
    });

    // Try restoring a cached token
    const restored = drive.tryRestoreToken();
    if (restored) {
      setUserInfo(restored.userInfo);
      await onSignedIn();
    }

    setDriveReady(true);
  }, [connectToFolder]);

  const signIn = useCallback(() => {
    if (!isSignedIn && folderId && folderId.startsWith('local-')) {
      snapshotLocalWorkbook(folderId, folderName ?? 'Untitled');
    }
    drive.requestAccessToken();
  }, [folderId, folderName, isSignedIn, snapshotLocalWorkbook]);

  const signOutHandler = useCallback(() => {
    drive.signOut();
    setIsSignedIn(false);
    rootFolderIdRef.current = null;
    booksConfigFileIdRef.current = null;
    const localList = Array.from(localWorkbooksRef.current.entries()).map(([id, snap]) => ({ id, name: snap.name }));
    setWorkbooks(localList.length > 0 ? localList : [LOCAL_DEFAULT_WORKBOOK]);
    setUserInfo(null);
    setLastSaved(null);
    setIsSaving(false);

    const targetId = localActiveWorkbookIdRef.current;
    const target = localWorkbooksRef.current.get(targetId);
    loadLocalWorkbook(targetId, target?.name ?? 'Untitled');
  }, [loadLocalWorkbook]);

  const createWorkbook = useCallback(async (name: string): Promise<string | null> => {
    if (!isSignedIn) {
      const trimmed = name.trim();
      if (!trimmed) return null;
      const currentLocalId = localActiveWorkbookIdRef.current;
      const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? 'Untitled';
      snapshotLocalWorkbook(currentLocalId, currentLocalName);
      const workbookId = `local-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      localWorkbooksRef.current.set(workbookId, {
        name: trimmed,
        tables: [],
        tableOrder: [],
        activeTableId: null,
      });
      const next = Array.from(localWorkbooksRef.current.entries()).map(([id, snap]) => ({ id, name: snap.name }));
      setWorkbooks(next);
      loadLocalWorkbook(workbookId, trimmed);
      return workbookId;
    }

    const rootId = rootFolderIdRef.current;
    if (!rootId) return null;
    const workbookId = await drive.findOrCreateSubfolder(name, rootId);
    const existingById = workbooks.find(w => w.id === workbookId);
    const created = existingById ?? { id: workbookId, name: name.trim() };
    const next = existingById ? workbooks : [...workbooks, created];
    setWorkbooks(next);
    await persistDriveBooksConfig(next);
    await loadWorkbook(created.id, created.name);
    return created.id;
  }, [isSignedIn, loadLocalWorkbook, loadWorkbook, persistDriveBooksConfig, workbooks]);

  const renameWorkbook = useCallback(async (workbookId: string, name: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;

    if (!isSignedIn) {
      const snapshot = localWorkbooksRef.current.get(workbookId);
      if (!snapshot) return;
      localWorkbooksRef.current.set(workbookId, {
        ...snapshot,
        name: trimmed,
      });
      const next = Array.from(localWorkbooksRef.current.entries()).map(([id, snap]) => ({ id, name: snap.name }));
      setWorkbooks(next);
      if (folderId === workbookId) {
        setFolderName(trimmed);
      }
      bump();
      return;
    }

    await drive.renameFile(workbookId, trimmed);
    const next = workbooks.map(w => w.id === workbookId ? { ...w, name: trimmed } : w);
    setWorkbooks(next);
    await persistDriveBooksConfig(next);
    if (folderId === workbookId) {
      setFolderName(trimmed);
    }
  }, [bump, folderId, isSignedIn, persistDriveBooksConfig, workbooks]);

  const deleteWorkbook = useCallback(async (workbookId: string): Promise<string | null> => {
    if (!isSignedIn) {
      const currentLocalId = localActiveWorkbookIdRef.current;
      const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? folderName ?? 'Untitled';
      snapshotLocalWorkbook(currentLocalId, currentLocalName);

      localWorkbooksRef.current.delete(workbookId);

      if (localWorkbooksRef.current.size === 0) {
        localWorkbooksRef.current.set(LOCAL_DEFAULT_WORKBOOK.id, {
          name: LOCAL_DEFAULT_WORKBOOK.name,
          tables: [],
          tableOrder: [],
          activeTableId: null,
        });
      }

      const next = Array.from(localWorkbooksRef.current.entries()).map(([id, snap]) => ({ id, name: snap.name }));
      setWorkbooks(next);

      const nextId = workbookId === currentLocalId ? next[0].id : currentLocalId;
      const nextBook = localWorkbooksRef.current.get(nextId);
      loadLocalWorkbook(nextId, nextBook?.name ?? 'Untitled');
      persistLocalBooks();
      return nextBook?.name ?? 'Untitled';
    }

    await drive.deleteFile(workbookId);
    let next = workbooks.filter(w => w.id !== workbookId);
    if (next.length === 0) {
      const rootId = rootFolderIdRef.current;
      if (rootId) {
        const defaultId = await drive.findOrCreateSubfolder('Untitled', rootId);
        next = [{ id: defaultId, name: 'Untitled' }];
      }
    }

    setWorkbooks(next);
    await persistDriveBooksConfig(next);

    if (next.length === 0) return null;
    const target = next.find(w => w.id === folderId && w.id !== workbookId) ?? next[0];
    if (!target) return null;
    await loadWorkbook(target.id, target.name);
    return target.name;
  }, [folderId, folderName, isSignedIn, loadLocalWorkbook, loadWorkbook, persistDriveBooksConfig, persistLocalBooks, snapshotLocalWorkbook, workbooks]);

  const switchWorkbook = useCallback(async (workbookId: string) => {
    if (!isSignedIn) {
      const currentLocalId = localActiveWorkbookIdRef.current;
      const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? 'Untitled';
      snapshotLocalWorkbook(currentLocalId, currentLocalName);

      const target = workbooks.find(w => w.id === workbookId);
      if (!target) return;
      loadLocalWorkbook(target.id, target.name);
      return;
    }

    const target = workbooks.find(w => w.id === workbookId);
    if (!target) return;
    await loadWorkbook(target.id, target.name);
  }, [isSignedIn, loadLocalWorkbook, loadWorkbook, snapshotLocalWorkbook, workbooks]);

  const shareWorkbook = useCallback(async (workbookId: string, emailAddress: string, role: 'reader' | 'writer' = 'writer') => {
    if (!isSignedIn || !workbookId) return;
    await drive.shareFolderWithEmail(workbookId, emailAddress, role);
  }, [isSignedIn]);

  const createWorkbookInviteLink = useCallback(async (workbookId: string, emailAddress: string, role: 'reader' | 'writer' = 'writer') => {
    if (!isSignedIn || !workbookId) {
      throw new Error('Sign in is required to create a share invite.');
    }

    const trimmedEmail = emailAddress.trim().toLowerCase();
    if (!trimmedEmail) {
      throw new Error('Invite email is required.');
    }

    const workbook = workbooks.find(w => w.id === workbookId);
    if (!workbook) {
      throw new Error('Workbook not found.');
    }

    await shareWorkbook(workbookId, trimmedEmail, role);

    const token = encodeInvitePayload({
      workbookId,
      workbookName: workbook.name,
      emailAddress: trimmedEmail,
      role,
    });

    const url = new URL(window.location.href);
    url.hash = `/accept?invite=${encodeURIComponent(token)}`;
    return url.toString();
  }, [isSignedIn, shareWorkbook, workbooks]);

  const acceptWorkbookInvite = useCallback(async (token: string): Promise<WorkbookInfo> => {
    if (!isSignedIn) {
      throw new Error('Sign in is required to accept invites.');
    }

    const invite = decodeInvitePayload(token);
    let signedInEmail = userInfo?.email?.trim().toLowerCase();
    if (!signedInEmail) {
      const freshInfo = await drive.getUserInfo();
      if (freshInfo) {
        setUserInfo(freshInfo);
        signedInEmail = freshInfo.email?.trim().toLowerCase();
      }
    }
    if (!signedInEmail) {
      throw new Error('Unable to verify signed-in email. Please sign out and sign in again to grant email access.');
    }
    if (signedInEmail !== invite.emailAddress.trim().toLowerCase()) {
      throw new Error(`This invite is for ${invite.emailAddress}. Signed in as ${signedInEmail}.`);
    }

    let rootId = rootFolderIdRef.current;
    if (!rootId) {
      const rootFolder = await drive.findOrCreateFolder();
      rootId = rootFolder.id;
      rootFolderIdRef.current = rootId;
    }

    let workbookName = invite.workbookName?.trim() || 'Shared Book';
    try {
      const sharedMeta = await drive.getFileMeta(invite.workbookId);
      if (sharedMeta.name?.trim()) {
        workbookName = sharedMeta.name.trim();
      }
    } catch {
      // Some shared items may not expose metadata with current token scope.
      // Continue using invite payload name.
    }

    try {
      await drive.ensureShortcutInFolder(rootId, invite.workbookId, workbookName);
    } catch (err) {
      throw new Error(`Could not add shared book shortcut. ${toErrorMessage(err)}`);
    }

    const existing = workbooks.find(w => w.id === invite.workbookId);
    const nextWorkbook: WorkbookInfo = existing ?? { id: invite.workbookId, name: workbookName };
    const nextBooks = existing ? workbooks : [...workbooks, nextWorkbook];

    setWorkbooks(nextBooks);
    await persistDriveBooksConfig(nextBooks);
    try {
      await loadWorkbook(nextWorkbook.id, nextWorkbook.name);
    } catch (err) {
      throw new Error(`Shared book was added, but opening it failed. ${toErrorMessage(err)}`);
    }

    return nextWorkbook;
  }, [isSignedIn, loadWorkbook, persistDriveBooksConfig, userInfo?.email, workbooks]);



  const saveAll = useCallback(async () => {
    if (!folderId) return;
    setIsSaving(true);

    // Capture the revision at the start of save to detect concurrent changes
    const savedRevision = revisionRef.current;

    try {
      const existingIds = model.getTableIds();
      const orderedInConfig = tableOrder.filter(id => existingIds.includes(id));
      const missingInOrder = existingIds.filter(id => !orderedInConfig.includes(id));
      const tableIds = [...orderedInConfig, ...missingInOrder];
      const schemas: TableSchema[] = [];

      const filesInFolder = await drive.listFilesInFolder(folderId);
      const fileByName = new Map(filesInFolder.map(f => [f.name, f.id]));
      const fileById = new Set(filesInFolder.map(f => f.id));

      const persistedTables = await Promise.all(tableIds.map(async (tableId) => {
        const table = model.getTable(tableId);
        if (!table) return null;

        const csv = rowsToCSV(table.schema, table.rows);
        const csvFileName = table.schema.csvFileName?.trim() || `${tableId}.csv`;
        const schemaFileId = table.schema.csvFileId?.trim();
        const mappedFileId = fileIdsRef.current.get(tableId);
        const namedFileId = fileByName.get(csvFileName);
        const existingFileId = schemaFileId && fileById.has(schemaFileId)
          ? schemaFileId
          : (mappedFileId && fileById.has(mappedFileId) ? mappedFileId : namedFileId);

        let persistedFileId = existingFileId;
        if (existingFileId) {
          await drive.updateFile(existingFileId, csv);
        } else {
          persistedFileId = await drive.createFile(csvFileName, csv, folderId);
        }

        if (!persistedFileId) {
          throw new Error(`Failed to persist CSV file for table ${tableId}`);
        }

        fileIdsRef.current.set(tableId, persistedFileId);

        const { csvFileName: _legacyCsvFileName, ...restSchema } = table.schema;
        return {
          tableId,
          schema: {
            ...restSchema,
            csvFileId: persistedFileId,
          } as TableSchema,
        };
      }));

      for (const persistedTable of persistedTables) {
        if (!persistedTable) continue;
        schemas.push(persistedTable.schema);
        model.markSaved(persistedTable.tableId);
      }

      // Save config (including chart sheets)
      const chartSheets = chartSheetOrder
        .map(name => chartSheetsRef.current.get(name))
        .filter((cs): cs is ChartSheet => !!cs);
      const config: ProjectConfig = { tables: schemas, chartSheets: chartSheets.length > 0 ? chartSheets : undefined };
      const configText = serializeConfig(config);

      if (configFileIdRef.current) {
        await drive.updateFile(configFileIdRef.current, configText, 'application/json');
      } else {
        configFileIdRef.current = await drive.createFile('sheetable.json', configText, folderId, 'application/json');
      }

      setLastSaved(new Date());
      // Only clear configDirty if no changes happened during save
      if (revisionRef.current === savedRevision) {
        configDirtyRef.current = false;
      }
      bump();
    } finally {
      setIsSaving(false);
    }
  }, [folderId, model, bump, tableOrder]);

  const tableIds = useMemo(() => {
    const existingIds = model.getTableIds();
    const ordered = tableOrder.filter(id => existingIds.includes(id));
    const remainder = existingIds.filter(id => !ordered.includes(id));
    return [...ordered, ...remainder];
  }, [model, revision, tableOrder]);

  // Auto-save: debounce 2 seconds after any change
  useEffect(() => {
    if (!folderId || !isSignedIn) return;
    // Check if anything is actually dirty
    const dirty = model.getTableIds().some(id => model.isDirty(id)) || configDirtyRef.current;
    if (!dirty) return;

    const timer = setTimeout(() => {
      saveAll();
    }, 2000);
    return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revision, folderId, isSignedIn]);

  // Restore local books on first mount (signed-out mode default)
  useEffect(() => {
    if (initializedLocalRef.current) return;
    initializedLocalRef.current = true;

    const stored = loadLocalBooksStorage();
    if (!stored) return;

    const entries = Object.entries(stored.workbooks ?? {});
    if (entries.length === 0) return;

    localWorkbooksRef.current = new Map(entries);
    const list = entries.map(([id, snap]) => ({ id, name: snap.name }));
    setWorkbooks(list);

    const targetId = stored.activeWorkbookId && localWorkbooksRef.current.has(stored.activeWorkbookId)
      ? stored.activeWorkbookId
      : list[0].id;
    localActiveWorkbookIdRef.current = targetId;
    const target = localWorkbooksRef.current.get(targetId);
    loadLocalWorkbook(targetId, target?.name ?? 'Untitled');
  }, [loadLocalWorkbook]);

  // Persist local books whenever local state changes
  useEffect(() => {
    if (isSignedIn) return;
    const currentLocalId = localActiveWorkbookIdRef.current;
    const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? folderName ?? 'Untitled';
    snapshotLocalWorkbook(currentLocalId, currentLocalName);
    persistLocalBooks();
  }, [folderName, isSignedIn, persistLocalBooks, revision, snapshotLocalWorkbook]);

  // Warn before unload if there are unsaved changes
  useEffect(() => {
    const dirty = model.getTableIds().some(id => model.isDirty(id)) || configDirtyRef.current;
    if (!dirty) return;

    const handler = (e: BeforeUnloadEvent) => {
      e.preventDefault();
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [revision, model]);

  return {
    model,
    tableIds,
    activeTableId,
    setActiveTableId,
    createTable,
    deleteTable,
    renameTable,
    renameColumn,
    reorderTables,
    updateSchema,
    getRows,
    getSchema,
    listWorkbookCsvFiles,
    loadTableFromCsvFile,
    renameTableCsvFile,
    applyEdit,
    insertRow,
    deleteRow,
    undo,
    canUndo: undoStackRef.current.length > 0,
    isDirty,
    isAnyDirty,
    isSignedIn,
    folderId,
    folderName,
    workbooks,
    createWorkbook,
    renameWorkbook,
    deleteWorkbook,
    switchWorkbook,
    shareWorkbook,
    createWorkbookInviteLink,
    acceptWorkbookInvite,
    signIn,
    signOut: signOutHandler,
    isSaving,
    lastSaved,
    initializeDrive,
    driveReady,
    isConnecting,
    userInfo,
    chartSheetIds,
    getChartSheet,
    createChartSheet,
    deleteChartSheet,
    renameChartSheet,
    updateChartSheet,
    setChartSheetTable,
    setChartSheetMode,
    revision,
  };
}

import { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { DataModel } from './dataModel';
import type { TableSchema, Row, Transaction, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { csvToRows, rowsToCSV } from './csv';
import type { ProjectConfig } from './config';
import { serializeConfig, parseConfig } from './config';
import * as drive from './drive';

export interface WorkbookInfo {
  id: string;
  name: string;
}

interface LocalWorkbookSnapshot {
  name: string;
  tables: Array<{ schema: TableSchema; rows: Row[] }>;
  tableOrder: string[];
  activeTableId: string | null;
}

const LOCAL_DEFAULT_WORKBOOK: WorkbookInfo = {
  id: 'local-default',
  name: 'Personal',
};

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
  reorderTables: (fromIndex: number, toIndex: number) => void;
  updateSchema: (tableId: string, schema: TableSchema) => void;
  getRows: (tableId: string) => Row[];
  getSchema: (tableId: string) => TableSchema | undefined;

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
  switchWorkbook: (workbookId: string) => Promise<void>;
  shareWorkbook: (workbookId: string, emailAddress: string, role?: 'reader' | 'writer') => Promise<void>;
  signIn: () => void;
  signOut: () => void;
  isSaving: boolean;
  lastSaved: Date | null;
  initializeDrive: (clientId: string) => Promise<void>;
  driveReady: boolean;
  isConnecting: boolean;
  userInfo: drive.UserInfo | null;

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
  const configDirtyRef = useRef(false);
  const revisionRef = useRef(0);
  const undoStackRef = useRef<Transaction[]>([]);
  const rootFolderIdRef = useRef<string | null>(null);
  const localWorkbooksRef = useRef<Map<string, LocalWorkbookSnapshot>>(
    new Map([[LOCAL_DEFAULT_WORKBOOK.id, { name: LOCAL_DEFAULT_WORKBOOK.name, tables: [], tableOrder: [], activeTableId: null }]])
  );
  const localActiveWorkbookIdRef = useRef<string>(LOCAL_DEFAULT_WORKBOOK.id);

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
    bump();
  }, [bump, clearModel, model]);

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

    for (const schema of config.tables) {
      const csvFile = files.find(f => f.name === `${schema.name}.csv`);
      let rows: Row[] = [];
      if (csvFile) {
        fileIdsRef.current.set(schema.name, csvFile.id);
        const csvText = await drive.downloadFile(csvFile.id);
        rows = csvToRows(csvText, schema);
      }
      model.createTable(schema, rows);
      model.markSaved(schema.name);
    }

    setTableOrder(config.tables.map(t => t.name));
    if (config.tables.length > 0) {
      setActiveTableId(config.tables[0].name);
    }
    bump();
  }, [bump, clearModel, model]);

  const refreshWorkbooks = useCallback(async () => {
    if (!rootFolderIdRef.current) return [] as WorkbookInfo[];
    const folders = await drive.listSubfolders(rootFolderIdRef.current);
    let next = folders;
    if (next.length === 0) {
      const defaultId = await drive.findOrCreateSubfolder('Personal', rootFolderIdRef.current);
      next = [{ id: defaultId, name: 'Personal' }];
    }
    setWorkbooks(next);
    return next;
  }, []);

  const createTable = useCallback((schema: TableSchema, rows?: Row[]) => {
    model.createTable(schema, rows);
    setTableOrder(prev => (prev.includes(schema.name) ? prev : [...prev, schema.name]));
    setActiveTableId(schema.name);
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
    undoStackRef.current = [];
    bump();
  }, [model, bump]);

  const getRows = useCallback((tableId: string): Row[] => {
    return [...(model.getTable(tableId)?.rows ?? [])];
  }, [model]);

  const getSchema = useCallback((tableId: string): TableSchema | undefined => {
    return model.getTable(tableId)?.schema;
  }, [model]);

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
      snapshotLocalWorkbook(folderId, folderName ?? 'Personal');
    }
    drive.requestAccessToken();
  }, [folderId, folderName, isSignedIn, snapshotLocalWorkbook]);

  const signOutHandler = useCallback(() => {
    drive.signOut();
    setIsSignedIn(false);
    rootFolderIdRef.current = null;
    const localList = Array.from(localWorkbooksRef.current.entries()).map(([id, snap]) => ({ id, name: snap.name }));
    setWorkbooks(localList.length > 0 ? localList : [LOCAL_DEFAULT_WORKBOOK]);
    setUserInfo(null);
    setLastSaved(null);
    setIsSaving(false);

    const targetId = localActiveWorkbookIdRef.current;
    const target = localWorkbooksRef.current.get(targetId);
    loadLocalWorkbook(targetId, target?.name ?? 'Personal');
  }, [loadLocalWorkbook]);

  const createWorkbook = useCallback(async (name: string): Promise<string | null> => {
    if (!isSignedIn) {
      const trimmed = name.trim();
      if (!trimmed) return null;
      const currentLocalId = localActiveWorkbookIdRef.current;
      const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? 'Personal';
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
    const next = await refreshWorkbooks();
    const created = next.find(w => w.id === workbookId) ?? { id: workbookId, name };
    await loadWorkbook(created.id, created.name);
    return created.id;
  }, [isSignedIn, loadLocalWorkbook, loadWorkbook, refreshWorkbooks]);

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
    const next = await refreshWorkbooks();
    if (folderId === workbookId) {
      const active = next.find(w => w.id === workbookId);
      setFolderName(active?.name ?? trimmed);
    }
  }, [bump, folderId, isSignedIn, refreshWorkbooks]);

  const switchWorkbook = useCallback(async (workbookId: string) => {
    if (!isSignedIn) {
      const currentLocalId = localActiveWorkbookIdRef.current;
      const currentLocalName = localWorkbooksRef.current.get(currentLocalId)?.name ?? 'Personal';
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

      for (const tableId of tableIds) {
        const table = model.getTable(tableId);
        if (!table) continue;
        schemas.push(table.schema);

        const csv = rowsToCSV(table.schema, table.rows);
        const existingFileId = fileIdsRef.current.get(tableId);

        if (existingFileId) {
          await drive.updateFile(existingFileId, csv);
        } else {
          const newFileId = await drive.createFile(`${tableId}.csv`, csv, folderId);
          fileIdsRef.current.set(tableId, newFileId);
        }

        model.markSaved(tableId);
      }

      // Save config
      const config: ProjectConfig = { tables: schemas };
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
    reorderTables,
    updateSchema,
    getRows,
    getSchema,
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
    switchWorkbook,
    shareWorkbook,
    signIn,
    signOut: signOutHandler,
    isSaving,
    lastSaved,
    initializeDrive,
    driveReady,
    isConnecting,
    userInfo,
    revision,
  };
}

import { useState, useCallback, useRef, useEffect } from 'react';
import { DataModel } from './dataModel';
import type { TableSchema, Row, Transaction, ValidationError } from './types';
import { csvToRows, rowsToCSV } from './csv';
import type { ProjectConfig } from './config';
import { serializeConfig, parseConfig } from './config';
import * as drive from './drive';

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
  updateSchema: (tableId: string, schema: TableSchema) => void;
  getRows: (tableId: string) => Row[];
  getSchema: (tableId: string) => TableSchema | undefined;

  // Transaction
  applyEdit: (tableId: string, rowIndex: number, columnName: string, newValue: string) => ValidationError[];
  insertRow: (tableId: string, row: Row) => ValidationError[];
  deleteRow: (tableId: string, rowIndex: number) => ValidationError[];

  // Dirty state
  isDirty: (tableId: string) => boolean;
  isAnyDirty: () => boolean;

  // Drive
  isSignedIn: boolean;
  folderId: string | null;
  folderName: string | null;
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
  const [isSignedIn, setIsSignedIn] = useState(false);
  const [folderId, setFolderId] = useState<string | null>(null);
  const [folderName, setFolderName] = useState<string | null>(null);
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

  const bump = useCallback(() => setRevision(r => {
    const next = r + 1;
    revisionRef.current = next;
    return next;
  }), []);

  const model = modelRef.current;

  const createTable = useCallback((schema: TableSchema, rows?: Row[]) => {
    model.createTable(schema, rows);
    setActiveTableId(schema.name);
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
    configDirtyRef.current = true;
    setActiveTableId(prev => prev === tableId ? null : prev);
    bump();
  }, [model, bump]);

  const renameTable = useCallback((oldName: string, newName: string) => {
    model.renameTable(oldName, newName);
    setActiveTableId(prev => prev === oldName ? newName : prev);
    bump();
  }, [model, bump]);

  const updateSchema = useCallback((tableId: string, schema: TableSchema) => {
    model.updateSchema(tableId, schema);
    bump();
  }, [model, bump]);

  const getRows = useCallback((tableId: string): Row[] => {
    return [...(model.getTable(tableId)?.rows ?? [])];
  }, [model]);

  const getSchema = useCallback((tableId: string): TableSchema | undefined => {
    return model.getTable(tableId)?.schema;
  }, [model]);

  const applyEdit = useCallback((tableId: string, rowIndex: number, columnName: string, newValue: string): ValidationError[] => {
    const tx: Transaction = {
      id: model.nextTransactionId(),
      tableId,
      type: 'update',
      rowIndex,
      columnName,
      newValue,
      timestamp: Date.now(),
    };
    const errors = model.applyTransaction(tx);
    if (errors.length === 0) bump();
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
    if (errors.length === 0) bump();
    return errors;
  }, [model, bump]);

  const deleteRow = useCallback((tableId: string, rowIndex: number): ValidationError[] => {
    const tx: Transaction = {
      id: model.nextTransactionId(),
      tableId,
      type: 'delete',
      rowIndex,
      timestamp: Date.now(),
    };
    const errors = model.applyTransaction(tx);
    if (errors.length === 0) bump();
    return errors;
  }, [model, bump]);

  const isDirty = useCallback((tableId: string): boolean => {
    return model.isDirty(tableId);
  }, [model]);

  const isAnyDirty = useCallback((): boolean => {
    return model.getTableIds().some(id => model.isDirty(id));
  }, [model]);

  // Connect to the "sheetable" folder automatically
  const connectToFolder = useCallback(async () => {
    setIsConnecting(true);
    try {
      const folder = await drive.findOrCreateFolder();
      setFolderId(folder.id);
      setFolderName(folder.name);

      // Load existing data from folder
      const files = await drive.listFilesInFolder(folder.id);
      const configFile = files.find(f => f.name === 'sheetable.json');
      if (configFile) {
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
        bump();
        if (config.tables.length > 0) {
          setActiveTableId(config.tables[0].name);
        }
      }
    } finally {
      setIsConnecting(false);
    }
  }, [model, bump]);

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
    drive.requestAccessToken();
  }, []);

  const signOutHandler = useCallback(() => {
    drive.signOut();
    setIsSignedIn(false);
    setFolderId(null);
    setFolderName(null);
    setUserInfo(null);
  }, []);



  const saveAll = useCallback(async () => {
    if (!folderId) return;
    setIsSaving(true);

    // Capture the revision at the start of save to detect concurrent changes
    const savedRevision = revisionRef.current;

    try {
      const tableIds = model.getTableIds();
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
  }, [folderId, model, bump]);

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
    tableIds: model.getTableIds(),
    activeTableId,
    setActiveTableId,
    createTable,
    deleteTable,
    renameTable,
    updateSchema,
    getRows,
    getSchema,
    applyEdit,
    insertRow,
    deleteRow,
    isDirty,
    isAnyDirty,
    isSignedIn,
    folderId,
    folderName,
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

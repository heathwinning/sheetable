import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Routes, Route, useNavigate, useParams, Link, useLocation, Navigate, useSearchParams } from 'react-router-dom';
import { useAppState } from './useAppState';
import { SpreadsheetGrid } from './SpreadsheetGrid';
import { EditTablePage } from './EditTablePage';
import { ChartSheetPage } from './ChartSheetPage';
import { useAlert, usePromptInput, useConfirm } from './DialogProvider';
import { ImportPage } from './ImportPage';
import { rowsToCSV } from './csv';
import * as api from './api';
import type { UseAppStateReturn } from './useAppState';
import type { BookMember, BookInvite } from './types';
import CalendarView from './CalendarView';
import ScheduleView from './ScheduleView';

type ViewType = 'grid' | 'calendar' | 'schedule';
import './App.css';

const bookPrefix = (bookName?: string) => (bookName ? `/book/${encodeURIComponent(bookName)}` : '');
const withBook = (bookName: string | undefined, suffix: string) => `${bookPrefix(bookName)}${suffix}`;

// --- Add Sheet Menu (dropdown for "+ Spreadsheet" / "+ Chart") ---
const AddSheetMenu: React.FC<{ state: UseAppStateReturn; bookId?: string }> = ({ state, bookId }) => {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const btnRef = useRef<HTMLButtonElement>(null);
  const navigate = useNavigate();
  const promptInput = usePromptInput();
  // pos variable removed - setPos is called but pos is never read
  const [, setPos] = useState<{ top: number; left: number }>({ top: 0, left: 0 });

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node) &&
          btnRef.current && !btnRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const toggle = () => {
    if (!open && btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      setPos({ top: rect.bottom + 2, left: rect.left });
    }
    setOpen(o => !o);
  };

  const addChart = async () => {
    setOpen(false);
    const name = await promptInput('Enter a name for the chart sheet:', 'New Chart Sheet', 'Chart name');
    if (!name?.trim()) return;
    state.createChartSheet(name.trim());
    navigate(withBook(bookId, `/chart/${encodeURIComponent(name.trim())}`));
  };

  const addView = async (viewType: 'calendar' | 'schedule') => {
    setOpen(false);
    const tableIds = state.tableIds;
    if (tableIds.length === 0) return;
    const label = viewType === 'calendar' ? 'Calendar' : 'Schedule';
    const name = await promptInput(
      `Enter a name for the ${label} view:`,
      `${tableIds[0]} ${label}`,
      'View name',
    );
    if (!name?.trim()) return;
    const tableName =
      tableIds.find(id => state.getSchema(id)?.columns.some(c => c.type === 'date' || c.type === 'datetime'))
      ?? tableIds[0];
    await state.createViewSheet(name.trim(), tableName, viewType);
    navigate(withBook(bookId, `/view/${encodeURIComponent(name.trim())}`));
  };

  return (
    <>
      <button ref={btnRef} className="table-tab add-tab" onClick={toggle} title="Add sheet">
        +
      </button>
      {open && (
        <div className="header-action-dropdown-wrap" ref={menuRef} onClick={() => setOpen(false)}>
          <div className="header-action-dropdown">
            <Link className="header-action-dropdown-item" to={withBook(bookId, '/table/new')} onClick={() => setOpen(false)}>
              <span className="add-sheet-icon">📊</span> Spreadsheet
            </Link>
            <button className="header-action-dropdown-item" onClick={addChart}>
              <span className="add-sheet-icon">📈</span> Chart
            </button>
            {state.tableIds.length > 0 && (
              <>
                <button className="header-action-dropdown-item" onClick={() => { void addView('calendar'); }}>
                  📅 Calendar View
                </button>
                <button className="header-action-dropdown-item" onClick={() => { void addView('schedule'); }}>
                  🗓️ Schedule View
                </button>
              </>
            )}
          </div>
        </div>
      )}
    </>
  );
};

// --- Import Menu (dropdown combining import options) ---
const ImportMenu: React.FC<{ bookId?: string; tableId: string }> = ({ bookId, tableId }) => {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const btnRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node) &&
          btnRef.current && !btnRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  return (
    <div className="header-action-dropdown-wrap">
      <button ref={btnRef} className="header-action-btn" onClick={() => setOpen(o => !o)} title="Import data">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        <span className="header-action-label">Import</span>
      </button>
      {open && (
        <div className="header-action-dropdown" ref={menuRef} onClick={() => setOpen(false)}>
          <Link className="header-action-dropdown-item" to={withBook(bookId, `/table/${encodeURIComponent(tableId)}/import`)}>
            Import into "{tableId}"
          </Link>
          <Link className="header-action-dropdown-item" to={withBook(bookId, '/import')}>
            Import as new table
          </Link>
        </div>
      )}
    </div>
  );
};

// --- Table View Page ---
const TableViewPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { tableId, bookId } = useParams<{ tableId: string; bookId?: string }>();
  const showAlert = useAlert();
  const [searchParams, setSearchParams] = useSearchParams();
  const viewType = (searchParams.get('view') ?? 'grid') as ViewType;

  const viewStorageKey = `sheetable-view-${state.activeBookId}-${tableId}`;

  // On mount / table change: restore saved view if URL has no ?view param
  useEffect(() => {
    if (!searchParams.has('view')) {
      const saved = localStorage.getItem(viewStorageKey);
      if (saved && saved !== 'grid') {
        setSearchParams(prev => {
          const next = new URLSearchParams(prev);
          next.set('view', saved);
          return next;
        }, { replace: true });
      }
    }
  }, [tableId]); // eslint-disable-line react-hooks/exhaustive-deps

  // Persist view whenever it changes
  useEffect(() => {
    if (viewType === 'grid') {
      localStorage.removeItem(viewStorageKey);
    } else {
      localStorage.setItem(viewStorageKey, viewType);
    }
  }, [viewType, viewStorageKey]);

  // Sync URL param to active table
  useEffect(() => {
    if (tableId && state.tableIds.includes(tableId)) {
      state.setActiveTableId(tableId);
    }
  }, [tableId, state.tableIds]); // eslint-disable-line react-hooks/exhaustive-deps

  const activeSchema = tableId ? state.getSchema(tableId) : null;
  const activeRows = tableId ? state.getRows(tableId) : [];

  // Date column selection for calendar/schedule views
  const dateColumns = useMemo(
    () => activeSchema?.columns.filter(c => c.type === 'date' || c.type === 'datetime') ?? [],
    [activeSchema],
  );
  const storageKey = `sheetable-date-col-${state.activeBookId}-${tableId}`;
  const [dateColumnOverride, setDateColumnOverride] = useState<string | null>(
    () => localStorage.getItem(storageKey),
  );
  const dateColumn =
    (dateColumnOverride && dateColumns.some(c => c.name === dateColumnOverride))
      ? dateColumnOverride
      : (dateColumns[0]?.name ?? null);
  const handleDateColumnChange = (col: string) => {
    setDateColumnOverride(col);
    localStorage.setItem(storageKey, col);
  };

  const runUndo = () => {
    const errors = state.undo();
    if (errors.length > 0) {
      showAlert(errors[0].message, 'Undo Failed');
    }
  };

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const isUndo = (e.metaKey || e.ctrlKey) && !e.shiftKey && e.key.toLowerCase() === 'z';
      if (!isUndo) return;
      // Avoid hijacking undo while typing in editable fields.
      const target = e.target as HTMLElement | null;
      const tag = target?.tagName?.toLowerCase();
      if (tag === 'input' || tag === 'textarea' || target?.isContentEditable) return;
      e.preventDefault();
      runUndo();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [state]);

  // Decide effective view — fall back to grid if schema has no date columns
  const effectiveView: ViewType =
    (viewType === 'calendar' || viewType === 'schedule') && !dateColumn
      ? 'grid'
      : viewType;

  return (
    <div className="app-body">
      {/* Main content */}
      <div className="main-content">
        {tableId && activeSchema ? (
          effectiveView === 'calendar' && dateColumn ? (
            <CalendarView
              schema={activeSchema}
              rows={activeRows}
              dateColumn={dateColumn}
              onDateColumnChange={handleDateColumnChange}
            />
          ) : effectiveView === 'schedule' && dateColumn ? (
            <ScheduleView
              schema={activeSchema}
              rows={activeRows}
              dateColumn={dateColumn}
              onDateColumnChange={handleDateColumnChange}
            />
          ) : (
          <SpreadsheetGrid
            key={tableId}
            schema={activeSchema}
            rows={activeRows}
            readOnly={state.activeBookRole === 'viewer'}
            onEdit={(rowIndex, columnName, newValue) =>
              state.applyEdit(tableId, rowIndex, columnName, newValue)
            }
            onInsert={(row) => state.insertRow(tableId, row)}
            onDeleteRow={(rowIndex) => state.deleteRow(tableId, rowIndex)}
            onColumnOrderChange={(orderedColumnNames) => {
              const current = activeSchema.columns.map(c => c.name);
              if (orderedColumnNames.length !== current.length) return;
              if (orderedColumnNames.every((name, i) => name === current[i])) return;
              const byName = new Map(activeSchema.columns.map(c => [c.name, c]));
              const reordered = orderedColumnNames
                .map(name => byName.get(name))
                .filter((c): c is typeof activeSchema.columns[number] => !!c);
              if (reordered.length !== activeSchema.columns.length) return;
              state.updateSchema(tableId, {
                ...activeSchema,
                columns: reordered,
              });
            }}
            onColumnWidthChange={(widths) => {
              const updated = activeSchema.columns.map(c =>
                widths[c.name] !== undefined ? { ...c, width: Math.round(widths[c.name]) } : c
              );
              if (updated.every((c, i) => c.width === activeSchema.columns[i].width)) return;
              state.updateSchema(tableId, {
                ...activeSchema,
                columns: updated,
              });
            }}
            revision={state.revision}
            bookId={state.activeBookId ?? null}
            getReferencedRow={state.getReferencedRow}
            getReferenceRows={state.getReferenceRows}
            resolveColumnPath={state.resolveColumnPath}
            resolveColumnPathLabel={state.resolveColumnPathLabel}
          />
          )
        ) : (
          <div className="empty-state-main">
            <h2>No table selected</h2>
            <p>Create a new table to get started.</p>
            {state.isLoading ? (
              <button className="btn-primary" disabled style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
                Loading…
              </button>
            ) : (
              <Link className="btn-primary" to={withBook(bookId, '/table/new')}>
                Create Table
              </Link>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

// --- Home Page (redirect to first table or show empty state) ---
const HomePage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { bookId } = useParams<{ bookId?: string }>();
  const navigate = useNavigate();

  // Auto-redirect to first table or chart sheet
  useEffect(() => {
    if (state.tableIds.length > 0) {
      navigate(withBook(bookId, `/table/${encodeURIComponent(state.tableIds[0])}`), { replace: true });
    } else if (state.chartSheetIds.length > 0) {
      navigate(withBook(bookId, `/chart/${encodeURIComponent(state.chartSheetIds[0])}`), { replace: true });
    }
  }, [state.tableIds, state.chartSheetIds, navigate, bookId]);

  if (state.tableIds.length > 0 || state.chartSheetIds.length > 0) return null;

  return (
    <div className="app-body">
      <div className="main-content">
        <div className="empty-state-main">
          {bookId ? (
            <>
              <h2>No tables yet</h2>
              <p>Create a new table to get started.</p>
              {state.isLoading ? (
                <button className="btn-primary" disabled style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
                  Loading…
                </button>
              ) : (
                <Link className="btn-primary" to={withBook(bookId, '/table/new')}>
                  Create Table
                </Link>
              )}
              <Link className="btn-secondary" to={withBook(bookId, '/import')}>
                Import from CSV / Sheet
              </Link>
            </>
          ) : (
            <>
              <h2>Welcome to Sheetable</h2>
              {state.user ? (
                <>
                  <p>Create or open a book to get started.</p>
                  <Link className="btn-primary" to="/book/new/settings">
                    Create Book
                  </Link>
                </>
              ) : (
                <>
                  <p>Sign in to get started.</p>
                  <button className="btn-primary" onClick={state.signIn}>Sign in</button>
                </>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
};

// --- User Status Button ---
const UserButton: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);

  // Close on outside click
  useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  if (state.isLoading) {
    return (
      <span className="drive-btn drive-btn-loading">
        <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
        Loading…
      </span>
    );
  }

  if (!state.user) {
    return (
      <button onClick={state.signIn} className="drive-btn drive-btn-signin">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M15 3h4a2 2 0 012 2v14a2 2 0 01-2 2h-4M10 17l5-5-5-5M15 12H3" />
        </svg>
        Sign in
      </button>
    );
  }

  return (
    <div className="drive-status-wrapper" ref={menuRef}>
      <button
        className="drive-btn drive-btn-status"
        onClick={() => setOpen(o => !o)}
        title={state.user.name}
      >
        <svg className="drive-avatar-placeholder" viewBox="0 0 28 28" fill="none">
          <circle cx="14" cy="11" r="4" fill="white"/>
          <path d="M6 23c0-4.4 3.6-8 8-8s8 3.6 8 8" fill="white"/>
        </svg>
        <span className="drive-status-dot synced" />
      </button>
      {open && (
        <div className="drive-dropdown">
          <div className="drive-dropdown-user">
            <div className="drive-dropdown-name">{state.user.name}</div>
            {state.user.email && (
              <div className="drive-dropdown-email">{state.user.email}</div>
            )}
          </div>
          <div className="drive-dropdown-divider" />
          <button
            className="drive-dropdown-item"
            onClick={() => { void state.signOut(); setOpen(false); }}
          >
            Sign out
          </button>
        </div>
      )}
    </div>
  );
};

const BookSidebar: React.FC<{ state: UseAppStateReturn; onMinimize: () => void }> = ({ state, onMinimize }) => {
  const navigate = useNavigate();
  const location = useLocation();
  const routeBookMatch = location.pathname.match(/^\/book\/([^/]+)/);
  const routeBookName = routeBookMatch ? decodeURIComponent(routeBookMatch[1]) : state.activeBookName;

  const currentTail = (() => {
    const m = location.pathname.match(/^\/book\/[^/]+(\/.*)?$/);
    if (m) return m[1] ?? '';
    if (location.pathname === '/') return '';
    if (location.pathname.startsWith('/table/') || location.pathname === '/table/new' || location.pathname === '/import') {
      return location.pathname;
    }
    return '';
  })();

  const openBook = async (bookId: string) => {
    await state.switchBook(bookId);
    const book = state.books.find(b => b.id === bookId);
    const bookName = book?.name ?? state.activeBookName ?? '';
    if (!bookName) return;
    navigate(`/book/${encodeURIComponent(bookName)}${currentTail}` || `/book/${encodeURIComponent(bookName)}`);
  };

  return (
    <aside className="book-sidebar">
      <div className="book-sidebar-header">
        <button
          className="book-sidebar-toggle"
          onClick={onMinimize}
          title="Minimize books"
          aria-label="Minimize books"
        >
          ⟨
        </button>
        <span className="book-sidebar-title">Books</span>
      </div>
      <div className="book-sidebar-list">
        <button className="book-sidebar-nav-new" onClick={() => navigate('/book/new/settings')} disabled={state.isLoading}>
          + New Book
        </button>
        {state.books.map(book => {
          const isActive = book.name === routeBookName;
          return (
            <div key={book.id} className={`book-sidebar-item ${isActive ? 'active' : ''}`}>
              <button
                className="book-sidebar-link"
                onClick={() => { void openBook(book.id); }}
                title={book.name}
              >
                {book.name}
              </button>
              <button
                className="book-sidebar-edit"
                onClick={() => navigate(`/book/${encodeURIComponent(book.name)}/settings`)}
                title="Edit book"
              >
                Edit
              </button>
            </div>
          );
        })}
      </div>
    </aside>
  );
};

// --- View Sheet Page ---
const ViewSheetPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { viewId, bookId } = useParams<{ viewId: string; bookId?: string }>();
  const navigate = useNavigate();
  const confirm = useConfirm();
  const promptInput = usePromptInput();

  const viewSheet = viewId ? state.getViewSheet(viewId) : undefined;

  // Config editors
  const [editOpen, setEditOpen] = useState(false);
  const [editTable, setEditTable] = useState('');
  const [editViewType, setEditViewType] = useState<'grid' | 'calendar' | 'schedule'>('calendar');
  const [editDateCol, setEditDateCol] = useState('');

  useEffect(() => {
    if (!viewSheet) return;
    setEditTable(viewSheet.tableName);
    setEditViewType(viewSheet.viewType);
    setEditDateCol(viewSheet.dateColumn ?? '');
  }, [viewSheet?.name]); // eslint-disable-line react-hooks/exhaustive-deps

  const dateColumnsForTable = useMemo(
    () => state.getSchema(editTable)?.columns.filter(c => c.type === 'date' || c.type === 'datetime') ?? [],
    [editTable, state.revision], // eslint-disable-line react-hooks/exhaustive-deps
  );

  const saveConfig = async () => {
    if (!viewId || !viewSheet) return;
    await state.updateViewSheet(viewId, {
      tableName: editTable,
      viewType: editViewType,
      dateColumn: editDateCol || undefined,
    });
    setEditOpen(false);
  };

  const doRename = async () => {
    if (!viewId) return;
    const name = await promptInput('Rename view:', viewId, 'View name');
    if (!name?.trim() || name.trim() === viewId) return;
    await state.renameViewSheet(viewId, name.trim());
    navigate(withBook(bookId, `/view/${encodeURIComponent(name.trim())}`), { replace: true });
  };

  const doDelete = async () => {
    if (!viewId) return;
    const confirmed = await confirm(`Delete view "${viewId}"?`, 'Delete View');
    if (!confirmed) return;
    await state.deleteViewSheet(viewId);
    const first = state.tableIds[0];
    navigate(first ? withBook(bookId, `/table/${encodeURIComponent(first)}`) : withBook(bookId, ''), { replace: true });
  };

  if (!viewSheet) {
    return (
      <div className="app-body">
        <div className="main-content">
          <div className="empty-state-main">
            <h2>View not found</h2>
          </div>
        </div>
      </div>
    );
  }

  const schema = state.getSchema(viewSheet.tableName);
  const rows = state.getRows(viewSheet.tableName);
  const dateColumns = schema?.columns.filter(c => c.type === 'date' || c.type === 'datetime') ?? [];
  const dateColumn = (viewSheet.dateColumn && dateColumns.some(c => c.name === viewSheet.dateColumn))
    ? viewSheet.dateColumn
    : (dateColumns[0]?.name ?? null);

  const handleDateColumnChange = (col: string) => {
    void state.updateViewSheet(viewSheet.name, { dateColumn: col });
  };

  const canEdit = state.activeBookRole === 'owner' || state.activeBookRole === 'editor';

  return (
    <div className="app-body">
      {/* Config bar */}
      {canEdit && (
        <div className="view-sheet-bar">
          <span className="view-sheet-bar-label">
            Showing <strong>{viewSheet.tableName}</strong> as{' '}
            <strong>{viewSheet.viewType}</strong>
          </span>
          <button className="header-action-btn" onClick={() => setEditOpen(o => !o)}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
            Configure
          </button>
          <button className="header-action-btn" onClick={() => { void doRename(); }}>
            Rename
          </button>
          <button className="header-action-btn" style={{ color: 'var(--color-danger)' }} onClick={() => { void doDelete(); }}>
            Delete
          </button>
        </div>
      )}

      {/* Config panel */}
      {editOpen && (
        <div className="view-sheet-config">
          <label className="view-sheet-config-label">Table</label>
          <select className="calendar-col-select" value={editTable} onChange={e => { setEditTable(e.target.value); setEditDateCol(''); }}>
            {state.tableIds.map(id => <option key={id} value={id}>{id}</option>)}
          </select>
          <label className="view-sheet-config-label">View type</label>
          <select className="calendar-col-select" value={editViewType} onChange={e => setEditViewType(e.target.value as 'grid' | 'calendar' | 'schedule')}>
            <option value="grid">Grid</option>
            <option value="calendar">Calendar</option>
            <option value="schedule">Schedule</option>
          </select>
          {(editViewType === 'calendar' || editViewType === 'schedule') && dateColumnsForTable.length > 0 && (
            <>
              <label className="view-sheet-config-label">Date column</label>
              <select className="calendar-col-select" value={editDateCol || dateColumnsForTable[0]?.name} onChange={e => setEditDateCol(e.target.value)}>
                {dateColumnsForTable.map(c => <option key={c.name} value={c.name}>{c.displayName ?? c.name}</option>)}
              </select>
            </>
          )}
          <button className="btn-primary" style={{ marginLeft: 8 }} onClick={() => { void saveConfig(); }}>Save</button>
          <button className="btn-secondary" style={{ marginLeft: 4 }} onClick={() => setEditOpen(false)}>Cancel</button>
        </div>
      )}

      {/* View content */}
      <div className="main-content">
        {!schema ? (
          <div className="empty-state-main">
            <h2>Table &ldquo;{viewSheet.tableName}&rdquo; not found</h2>
            {canEdit && <p>Open Configure to pick a different table.</p>}
          </div>
        ) : viewSheet.viewType === 'calendar' && dateColumn ? (
          <CalendarView
            schema={schema}
            rows={rows}
            dateColumn={dateColumn}
            onDateColumnChange={handleDateColumnChange}
          />
        ) : viewSheet.viewType === 'schedule' && dateColumn ? (
          <ScheduleView
            schema={schema}
            rows={rows}
            dateColumn={dateColumn}
            onDateColumnChange={handleDateColumnChange}
          />
        ) : (
          <SpreadsheetGrid
            key={viewSheet.tableName}
            schema={schema}
            rows={rows}
            readOnly={true}
            onEdit={() => []}
            onInsert={() => []}
            onDeleteRow={() => []}
            revision={state.revision}
            bookId={state.activeBookId ?? null}
            getReferencedRow={state.getReferencedRow}
            getReferenceRows={state.getReferenceRows}
            resolveColumnPath={state.resolveColumnPath}
            resolveColumnPathLabel={state.resolveColumnPathLabel}
          />
        )}
      </div>
    </div>
  );
};

const BookSettingsPage: React.FC<{ state: UseAppStateReturn; createMode?: boolean }> = ({ state, createMode = false }) => {
  const { bookId } = useParams<{ bookId: string }>();
  const navigate = useNavigate();
  const confirm = useConfirm();
  const effectiveBookName = createMode ? '' : (bookId ?? state.activeBookName ?? '');

  // Track book by ID so renames don't lose it
  const [trackedBookId, setTrackedBookId] = useState<string | null>(() => {
    const book = state.books.find(b => b.name === effectiveBookName);
    return book?.id ?? null;
  });

  // Update tracked ID when navigating to a different book
  useEffect(() => {
    if (createMode) return;
    const book = state.books.find(b => b.name === effectiveBookName);
    if (book && book.id !== trackedBookId) {
      setTrackedBookId(book.id);
    }
  }, [createMode, effectiveBookName, state.books, trackedBookId]);

  const currentBook = trackedBookId
    ? state.books.find(b => b.id === trackedBookId)
    : state.books.find(b => b.name === effectiveBookName);

  const [name, setName] = useState(currentBook?.name ?? '');
  const [shareEmail, setShareEmail] = useState('');
  const [shareRole, setShareRole] = useState<'editor' | 'viewer'>('editor');
  const [members, setMembers] = useState<BookMember[]>([]);
  const [invites, setInvites] = useState<BookInvite[]>([]);
  const [status, setStatus] = useState<string | null>(null);

  useEffect(() => {
    if (createMode) {
      setName('');
      return;
    }
    if (!currentBook) return;
    setName(currentBook.name);
  }, [createMode, currentBook]);

  // Load members when viewing existing book
  useEffect(() => {
    if (createMode || !currentBook) return;
    api.listMembers(currentBook.id).then(data => {
      setMembers(data.members);
      setInvites(data.invites);
    }).catch(() => {});
  }, [createMode, currentBook]);

  useEffect(() => {
    if (createMode) return;
    if (!effectiveBookName || effectiveBookName === state.activeBookName) return;
    const target = state.books.find(w => w.name === effectiveBookName);
    if (!target) return;
    void state.switchBook(target.id);
  }, [createMode, effectiveBookName, state]);

  // Sync URL when book name changes (e.g. after rename)
  useEffect(() => {
    if (createMode || !currentBook) return;
    if (effectiveBookName !== currentBook.name) {
      navigate(`/book/${encodeURIComponent(currentBook.name)}/settings`, { replace: true });
    }
  }, [createMode, currentBook, effectiveBookName, navigate]);

  if (!createMode && !currentBook) {
    return (
      <div className="book-settings-page">
        <div className="book-settings-card">
          <h2>Book not found</h2>
          <button className="btn-secondary" onClick={() => navigate('/')}>Back</button>
        </div>
      </div>
    );
  }

  const saveName = async () => {
    const trimmed = name.trim();
    if (!trimmed) {
      setStatus('Book name cannot be empty.');
      return;
    }
    if (createMode) {
      const createdId = await state.createBook(trimmed);
      if (!createdId) {
        setStatus('Failed to create book.');
        return;
      }
      navigate(`/book/${encodeURIComponent(trimmed)}/settings`, { replace: true });
      return;
    }

    if (!currentBook) {
      setStatus('Book not found.');
      return;
    }

    await state.renameBook(currentBook.id, trimmed);
    setStatus('Book name updated.');
  };

  const doShare = async () => {
    if (createMode || !currentBook) {
      setStatus('Create the book first to enable sharing.');
      return;
    }
    if (!state.user) {
      setStatus('Sign in to share this book.');
      return;
    }
    if (!shareEmail.trim()) {
      setStatus('Enter an email address to share.');
      return;
    }
    try {
      const result = await api.addMember(currentBook.id, shareEmail.trim(), shareRole);
      if (result.invited) {
        const link = `${window.location.origin}/#/invite/${currentBook.id}`;
        void navigator.clipboard.writeText(link);
        setStatus(`Invited ${shareEmail.trim()} as ${shareRole}. Invite link copied to clipboard.`);
      } else {
        setStatus(`Added ${shareEmail.trim()} as ${shareRole}.`);
      }
      setShareEmail('');
      // Refresh member + invite list
      const updated = await api.listMembers(currentBook.id);
      setMembers(updated.members);
      setInvites(updated.invites);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to add member.';
      setStatus(message);
    }
  };

  const changeRole = async (email: string, newRole: string) => {
    if (!currentBook) return;
    try {
      await api.addMember(currentBook.id, email, newRole);
      const updated = await api.listMembers(currentBook.id);
      setMembers(updated.members);
      setInvites(updated.invites);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to change role.';
      setStatus(message);
    }
  };

  const removeMember = async (userId: string) => {
    if (!currentBook) return;
    try {
      await api.removeMember(currentBook.id, userId);
      setMembers(prev => prev.filter(m => m.userId !== userId));
      setStatus('Member removed.');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to remove member.';
      setStatus(message);
    }
  };

  const cancelInvite = async (email: string) => {
    if (!currentBook) return;
    try {
      await api.cancelInvite(currentBook.id, email);
      setInvites(prev => prev.filter(i => i.email !== email));
      setStatus('Invitation cancelled.');
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to cancel invitation.';
      setStatus(message);
    }
  };

  const copyInviteLink = (bookId: string) => {
    const link = `${window.location.origin}/#/invite/${bookId}`;
    void navigator.clipboard.writeText(link);
    setStatus('Invite link copied to clipboard.');
  };

  const doDelete = async () => {
    if (createMode || !currentBook) return;
    const confirmed = await confirm(`Delete book "${currentBook.name}"? This cannot be undone.`, 'Delete Book');
    if (!confirmed) return;
    await state.deleteBook(currentBook.id);
    navigate('/', { replace: true });
  };

  const isOwner = !createMode && currentBook?.role === 'owner';

  return (
    <div className="book-settings-page">
      <div className="book-settings-card">
        <div className="book-settings-header">
          <h2>{createMode ? 'New Book' : 'Book Settings'}</h2>
          <button className="btn-secondary btn-sm" onClick={() => navigate(createMode ? (state.activeBookName ? `/book/${encodeURIComponent(state.activeBookName)}` : '/') : (currentBook ? `/book/${encodeURIComponent(currentBook.name)}` : '/'))}>
            ← Back to Book
          </button>
        </div>

        {(createMode || isOwner) && (
        <div className="book-settings-section">
          <label className="book-settings-label">Book Name</label>
          <div className="book-settings-row">
            <input
              className="edit-table-input"
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="Book name"
            />
            <button className="btn-primary" onClick={() => { void saveName(); }}>
              {createMode ? 'Create Book' : 'Save Name'}
            </button>
          </div>
        </div>
        )}

        <div className="book-settings-section">
          <label className="book-settings-label">Members</label>
          <div className="members-list">
            {members.map(m => (
              <div key={m.userId} className="member-row">
                <span className="member-name">{m.name || m.email}</span>
                <span className="member-role">
                  {m.role === 'owner' ? (
                    <span className="text-text-muted">owner</span>
                  ) : isOwner ? (
                    <select
                      className="workbook-toolbar-select"
                      value={m.role}
                      onChange={(e) => { void changeRole(m.email, e.target.value); }}
                      style={{ fontSize: 12, padding: '2px 4px' }}
                    >
                      <option value="editor">Editor</option>
                      <option value="viewer">Viewer</option>
                    </select>
                  ) : (
                    <span className="text-text-muted">{m.role}</span>
                  )}
                </span>
                <span className="member-actions">
                  {m.role !== 'owner' && isOwner && (
                    <button className="btn-ghost btn-sm" onClick={() => { void removeMember(m.userId); }}>Remove</button>
                  )}
                </span>
              </div>
            ))}
            {isOwner && invites.map(inv => (
              <div key={inv.email} className="member-row">
                <span className="member-name">{inv.email} <span className="text-text-muted">(invited)</span></span>
                <span className="member-role">
                  <select
                    className="workbook-toolbar-select"
                    value={inv.role}
                    onChange={(e) => { void changeRole(inv.email, e.target.value); }}
                    style={{ fontSize: 12, padding: '2px 4px' }}
                  >
                    <option value="editor">Editor</option>
                    <option value="viewer">Viewer</option>
                  </select>
                </span>
                <span className="member-actions">
                  <button className="btn-ghost btn-sm" onClick={() => currentBook && copyInviteLink(currentBook.id)}>Copy Link</button>
                  <button className="btn-ghost btn-sm" onClick={() => { void cancelInvite(inv.email); }}>Cancel</button>
                </span>
              </div>
            ))}
            {(createMode || isOwner) && (
            <div className="member-row member-add-row">
              <input
                className="edit-table-input"
                type="email"
                value={shareEmail}
                onChange={(e) => setShareEmail(e.target.value)}
                placeholder="user@example.com"
                disabled={!state.user || createMode}
              />
              <span className="member-role">
                <select
                  className="workbook-toolbar-select"
                  value={shareRole}
                  onChange={(e) => setShareRole(e.target.value as 'editor' | 'viewer')}
                  disabled={!state.user || createMode}
                >
                  <option value="editor">Editor</option>
                  <option value="viewer">Viewer</option>
                </select>
              </span>
              <span className="member-actions">
                <button className="btn-secondary" onClick={() => { void doShare(); }} disabled={!state.user || createMode}>
                  Add Member
                </button>
              </span>
            </div>
            )}
          </div>
          {createMode && <div className="book-settings-note">Create this book first, then add members.</div>}
          {!createMode && !state.user && <div className="book-settings-note">Sign in to manage members.</div>}
        </div>

        {!createMode && currentBook && isOwner && (
          <div className="book-settings-section">
            <label className="book-settings-label">Danger Zone</label>
            <div className="book-settings-row">
              <button className="btn-danger" onClick={() => { void doDelete(); }}>
                Delete Book
              </button>
            </div>
          </div>
        )}

        {status && <div className="edit-table-notice edit-table-notice-info">{status}</div>}
      </div>
    </div>
  );
};

// --- Invite Accept Page ---
const InviteAcceptPage: React.FC<{ state: ReturnType<typeof useAppState> }> = ({ state }) => {
  const { bookId } = useParams<{ bookId: string }>();
  const navigate = useNavigate();
  const [info, setInfo] = useState<{ bookName: string; status: string; role?: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [accepting, setAccepting] = useState(false);

  useEffect(() => {
    if (!bookId) return;
    api.getInviteStatus(bookId).then(setInfo).catch(() => setError('Failed to load invitation.'));
  }, [bookId]);

  useEffect(() => {
    if (!info) return;
    if (info.status === 'already-member') {
      navigate(`/book/${encodeURIComponent(info.bookName)}`, { replace: true });
    }
  }, [info, navigate]);

  const doAccept = async () => {
    if (!bookId || !info) return;
    setAccepting(true);
    try {
      await api.acceptInvite(bookId);
      // Refresh books list then navigate
      await state.refreshBooks();
      navigate(`/book/${encodeURIComponent(info.bookName)}`, { replace: true });
    } catch {
      setError('Failed to accept invitation.');
      setAccepting(false);
    }
  };

  if (error) {
    return (
      <div className="book-settings-page">
        <div className="book-settings-card">
          <h2>Invitation</h2>
          <p>{error}</p>
          <button className="btn-secondary" onClick={() => navigate('/')}>Go Home</button>
        </div>
      </div>
    );
  }

  if (!info) {
    return (
      <div className="book-settings-page">
        <div className="book-settings-card">
          <h2>Loading invitation…</h2>
        </div>
      </div>
    );
  }

  return (
    <div className="book-settings-page">
      <div className="book-settings-card">
        <h2>You&apos;re invited to join &ldquo;{info.bookName}&rdquo;</h2>
        {info.status === 'sign-in-required' && (
          <>
            <p>Sign in with Google to accept this invitation.</p>
            <button className="btn-primary" onClick={state.signIn}>Sign in with Google</button>
          </>
        )}
        {info.status === 'pending' && (
          <>
            <p>You&apos;ve been invited as <strong>{info.role}</strong>.</p>
            <button className="btn-primary" onClick={() => { void doAccept(); }} disabled={accepting}>
              {accepting ? 'Joining…' : 'Accept Invitation'}
            </button>
          </>
        )}
        {info.status === 'no-invite' && (
          <>
            <p>No pending invitation found for {state.user?.email}. The book owner needs to invite your email address first.</p>
            <button className="btn-secondary" onClick={() => navigate('/')}>Go Home</button>
          </>
        )}
      </div>
    </div>
  );
};

// --- Main App Shell ---
const App: React.FC = () => {
  const state = useAppState();
  const [sidebarCollapsed, setSidebarCollapsed] = useState(true);
  const [searchParams, setSearchParams] = useSearchParams();
  const currentView = (searchParams.get('view') ?? 'grid') as ViewType;
  const [draggingTableId, setDraggingTableId] = useState<string | null>(null);
  const [draggingChartId, setDraggingChartId] = useState<string | null>(null);
  const location = useLocation();
  const navigate = useNavigate();

  // Sync AG Grid dark mode with system preference
  useEffect(() => {
    const mq = window.matchMedia('(prefers-color-scheme: dark)');
    const apply = () => {
      document.documentElement.setAttribute('data-ag-theme-mode', mq.matches ? 'dark' : 'light');
    };
    apply();
    mq.addEventListener('change', apply);
    return () => mq.removeEventListener('change', apply);
  }, []);

  // After login, redirect back to the page the user was on (e.g. invite page)
  useEffect(() => {
    if (!state.user) return;
    const redirect = localStorage.getItem('sheetable-post-login-redirect');
    if (redirect) {
      localStorage.removeItem('sheetable-post-login-redirect');
      const path = redirect.replace(/^#/, '');
      if (path && path !== '/' && path !== location.pathname) {
        navigate(path, { replace: true });
      }
    }
  }, [state.user]); // eslint-disable-line react-hooks/exhaustive-deps

  // Keyboard shortcut: Ctrl+S / Cmd+S to prevent default
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, []);

  // URL -> book state: switching /book/:bookId should switch active book.
  useEffect(() => {
    if (state.books.length === 0) return;
    const m = location.pathname.match(/^\/book\/([^/]+)/);
    if (!m) return;
    const routeBookName = decodeURIComponent(m[1]);
    if (routeBookName === state.activeBookName) return;
    const target = state.books.find(w => w.name === routeBookName);
    if (!target) return;
    void state.switchBook(target.id);
  }, [location.pathname, state]);

  // Redirect bare / to active book when one exists.
  useEffect(() => {
    if (!state.activeBookName) return;
    if (location.pathname === '/') {
      navigate(`/book/${encodeURIComponent(state.activeBookName)}`, { replace: true });
      return;
    }
    if (location.pathname.startsWith('/book/')) return;
    if (location.pathname === '/table/new' || location.pathname.startsWith('/table/') || location.pathname === '/import') {
      navigate(`/book/${encodeURIComponent(state.activeBookName)}${location.pathname}`, { replace: true });
    }
  }, [location.pathname, navigate, state.activeBookName]);

  // Derive current book route for header tabs
  const headerBookMatch = location.pathname.match(/^\/book\/([^/]+)/);
  const headerBookId = headerBookMatch ? decodeURIComponent(headerBookMatch[1]) : undefined;
  const isOnSheetRoute = /\/(table|chart|view)\//.test(location.pathname) || location.pathname.match(/^\/book\/[^/]+$/);

  // Derive active table ID for header actions
  const tableMatch = location.pathname.match(/\/table\/([^/]+)$/);
  const headerTableId = tableMatch ? decodeURIComponent(tableMatch[1]) : null;
  const isTableView = !!headerTableId && !location.pathname.includes('/edit') && !location.pathname.includes('/import');

  // Determine if the current table has date/datetime columns (for view switcher visibility)
  const headerTableHasDates = useMemo(() => {
    if (!headerTableId) return false;
    const schema = state.getSchema(headerTableId);
    return schema?.columns.some(c => c.type === 'date' || c.type === 'datetime') ?? false;
  }, [headerTableId, state.revision]); // eslint-disable-line react-hooks/exhaustive-deps

  const setView = (v: ViewType) => {
    setSearchParams(prev => {
      const next = new URLSearchParams(prev);
      if (v === 'grid') next.delete('view');
      else next.set('view', v);
      return next;
    }, { replace: true });
  };

  const handleTabDrop = (targetId: string) => {
    if (!draggingTableId || draggingTableId === targetId) return;
    const fromIndex = state.tableIds.indexOf(draggingTableId);
    const toIndex = state.tableIds.indexOf(targetId);
    if (fromIndex >= 0 && toIndex >= 0) {
      state.reorderTables(fromIndex, toIndex);
    }
    setDraggingTableId(null);
  };

  const handleChartTabDrop = (targetId: string) => {
    if (!draggingChartId || draggingChartId === targetId) return;
    const fromIndex = state.chartSheetIds.indexOf(draggingChartId);
    const toIndex = state.chartSheetIds.indexOf(targetId);
    if (fromIndex >= 0 && toIndex >= 0) {
      state.reorderCharts(fromIndex, toIndex);
    }
    setDraggingChartId(null);
  };

  const canEdit = state.activeBookRole === 'owner' || state.activeBookRole === 'editor';

  const showAlert = useAlert();
  const runUndo = () => {
    const errors = state.undo();
    if (errors.length > 0) {
      void showAlert(errors[0].message, 'Undo Failed');
    }
  };

  const exportCSV = () => {
    if (!headerTableId) return;
    const schema = state.getSchema(headerTableId);
    const rows = state.getRows(headerTableId);
    if (!schema) return;
    const csv = rowsToCSV(schema, rows);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${headerTableId}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="app">
      {/* Top bar */}
      <header className="app-header">
        <div className="header-left">
          <button
            className="book-sidebar-toggle"
            onClick={() => setSidebarCollapsed(c => !c)}
            title={sidebarCollapsed ? 'Show books' : 'Hide books'}
            aria-label={sidebarCollapsed ? 'Show books' : 'Hide books'}
          >
            ☰
          </button>
          <Link to="/" className="app-title-link">
            <h1 className="app-title">Sheetable</h1>
          </Link>
          {headerBookId && (
            <Link to={`/book/${encodeURIComponent(headerBookId)}/settings`} className="header-book-name" title={headerBookId}>
              {headerBookId}
            </Link>
          )}
          {state.activeBookRole && state.activeBookRole !== 'owner' && (
            <span className="header-role-badge">{state.activeBookRole}</span>
          )}
        </div>
        {isOnSheetRoute && (state.tableIds.length > 0 || state.chartSheetIds.length > 0 || state.viewSheetIds.length > 0) && (
          <div className="header-tabs">
            {state.tableIds.map(id => {
              const isActive = location.pathname.includes(`/table/${encodeURIComponent(id)}`);
              return (
                <Link
                  key={id}
                  className={`table-tab ${isActive ? 'active' : ''} ${id === draggingTableId ? 'dragging' : ''}`}
                  to={withBook(headerBookId, `/table/${encodeURIComponent(id)}`)}
                  draggable
                  onDragStart={(e) => {
                    setDraggingTableId(id);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', id);
                  }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    handleTabDrop(id);
                  }}
                  onDragEnd={() => setDraggingTableId(null)}
                >
                  {id}
                </Link>
              );
            })}
            {state.chartSheetIds.map(id => {
              const isActive = location.pathname.includes(`/chart/${encodeURIComponent(id)}`);
              return (
                <Link
                  key={`chart-${id}`}
                  className={`table-tab chart-tab ${isActive ? 'active' : ''} ${id === draggingChartId ? 'dragging' : ''}`}
                  to={withBook(headerBookId, `/chart/${encodeURIComponent(id)}`)}
                  draggable
                  onDragStart={(e) => {
                    setDraggingChartId(id);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', id);
                  }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    handleChartTabDrop(id);
                  }}
                  onDragEnd={() => setDraggingChartId(null)}
                >
                  📈 {id}
                </Link>
              );
            })}
            {state.viewSheetIds.map(id => {
              const isActive = location.pathname.includes(`/view/${encodeURIComponent(id)}`);
              const vs = state.getViewSheet(id);
              const icon = vs?.viewType === 'calendar' ? '📅' : vs?.viewType === 'schedule' ? '🗓️' : '📋';
              return (
                <Link
                  key={`view-${id}`}
                  className={`table-tab view-tab ${isActive ? 'active' : ''}`}
                  to={withBook(headerBookId, `/view/${encodeURIComponent(id)}`)}
                >
                  {icon} {id}
                </Link>
              );
            })}
            {state.isLoading ? (
              <span className="table-tab add-tab disabled" title="Loading…" style={{ opacity: 0.5, cursor: 'default' }}>
                <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
              </span>
            ) : canEdit ? (
              <AddSheetMenu state={state} bookId={headerBookId} />
            ) : null}
          </div>
        )}
        <div className="header-right">
          {isTableView && headerTableId && headerTableHasDates && (
            <div className="view-switcher">
              <button
                className={`view-switch-btn${currentView === 'grid' ? ' active' : ''}`}
                onClick={() => setView('grid')}
                title="Grid view"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
                  <rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>
                </svg>
              </button>
              <button
                className={`view-switch-btn${currentView === 'calendar' ? ' active' : ''}`}
                onClick={() => setView('calendar')}
                title="Calendar view"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/>
                  <line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
                </svg>
              </button>
              <button
                className={`view-switch-btn${currentView === 'schedule' ? ' active' : ''}`}
                onClick={() => setView('schedule')}
                title="Schedule view"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/>
                  <line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/>
                  <line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>
                </svg>
              </button>
            </div>
          )}
          {isTableView && headerTableId && (
            <div className="header-actions">
              {canEdit && (
                <>
                  <button
                    className="header-action-btn"
                    onClick={runUndo}
                    disabled={!state.canUndo}
                    title="Undo (Ctrl/Cmd+Z)"
                  >
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{display:'block',margin:'0 auto'}}>
                      <path d="M3 12h13a5 5 0 1 1 0 10h-1" />
                      <polyline points="8 17 3 12 8 7" />
                    </svg>
                  </button>
                  <ImportMenu bookId={headerBookId} tableId={headerTableId} />
                </>
              )}
              <button
                className="header-action-btn"
                onClick={exportCSV}
                title="Export as CSV"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                <span className="header-action-label">Export</span>
              </button>
              {canEdit && (
                <Link
                  className="header-action-btn"
                  to={withBook(headerBookId, `/table/${encodeURIComponent(headerTableId)}/edit`)}
                  title="Edit table schema"
                >
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                  <span className="header-action-label">Edit</span>
                </Link>
              )}
            </div>
          )}
          <UserButton state={state} />
        </div>
      </header>

      <div className="app-shell">
        {!sidebarCollapsed && (
          <>
            <div className="sidebar-overlay" onClick={() => setSidebarCollapsed(true)} />
            <BookSidebar state={state} onMinimize={() => setSidebarCollapsed(true)} />
          </>
        )}
        <div className="app-main">
          <Routes>
            <Route path="/" element={<HomePage state={state} />} />
            <Route path="/invite/:bookId" element={<InviteAcceptPage state={state} />} />
            <Route path="/book/new/settings" element={<BookSettingsPage state={state} createMode={true} />} />
            <Route path="/book/:bookId" element={<HomePage state={state} />} />
            <Route path="/book/:bookId/table/new" element={<EditTablePage state={state} />} />
            <Route path="/book/:bookId/table/:tableId" element={<TableViewPage state={state} />} />
            <Route path="/book/:bookId/chart/:chartId" element={<ChartSheetPage state={state} />} />
            <Route path="/book/:bookId/settings" element={<BookSettingsPage state={state} />} />
            <Route path="/book/:bookId/view/:viewId" element={<ViewSheetPage state={state} />} />
            <Route path="/book/:bookId/table/:tableId/edit" element={<EditTablePage state={state} />} />
            <Route path="/book/:bookId/table/:tableId/import" element={<ImportPage state={state} />} />
            <Route path="/book/:bookId/import" element={<ImportPage state={state} />} />
            <Route path="/table/new" element={<EditTablePage state={state} />} />
            <Route path="/table/:tableId" element={<TableViewPage state={state} />} />
            <Route path="/table/:tableId/edit" element={<EditTablePage state={state} />} />
            <Route path="/table/:tableId/import" element={<ImportPage state={state} />} />
            <Route path="/import" element={<ImportPage state={state} />} />
            <Route path="/chart/:chartId" element={<ChartSheetPage state={state} />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </div>
      </div>
    </div>
  );
};

export default App;

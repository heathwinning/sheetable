import React, { useState, useEffect, useRef } from 'react';
import { Routes, Route, useNavigate, useParams, Link, useLocation, Navigate } from 'react-router-dom';
import { useAppState } from './useAppState';
import { SpreadsheetGrid } from './SpreadsheetGrid';
import { EditTablePage } from './EditTablePage';
import { useAlert } from './DialogProvider';
import { ImportPage } from './ImportPage';
import type { UseAppStateReturn } from './useAppState';
import './App.css';

// Client ID should be configured per deployment
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || '';

const bookPrefix = (bookId?: string) => (bookId ? `/book/${encodeURIComponent(bookId)}` : '');
const withBook = (bookId: string | undefined, suffix: string) => `${bookPrefix(bookId)}${suffix}`;

// --- Table View Page ---
const TableViewPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { tableId, bookId } = useParams<{ tableId: string; bookId?: string }>();
  const [draggingTableId, setDraggingTableId] = useState<string | null>(null);
  const showAlert = useAlert();

  // Sync URL param to active table
  useEffect(() => {
    if (tableId && state.tableIds.includes(tableId)) {
      state.setActiveTableId(tableId);
    }
  }, [tableId, state.tableIds]); // eslint-disable-line react-hooks/exhaustive-deps

  const activeSchema = tableId ? state.getSchema(tableId) : null;
  const activeRows = tableId ? state.getRows(tableId) : [];

  const handleTabDrop = (targetId: string) => {
    if (!draggingTableId || draggingTableId === targetId) return;
    const fromIndex = state.tableIds.indexOf(draggingTableId);
    const toIndex = state.tableIds.indexOf(targetId);
    if (fromIndex >= 0 && toIndex >= 0) {
      state.reorderTables(fromIndex, toIndex);
    }
    setDraggingTableId(null);
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

  return (
    <div className="app-body">
      {/* Table tabs */}
      <div className="table-tabs-bar">
        <div className="table-tabs">
          {state.tableIds.map(id => (
            <Link
              key={id}
              className={`table-tab ${id === tableId ? 'active' : ''} ${id === draggingTableId ? 'dragging' : ''}`}
              to={withBook(bookId, `/table/${encodeURIComponent(id)}`)}
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
              {state.isDirty(id) && <span className="tab-dirty">●</span>}
            </Link>
          ))}
          {state.isConnecting ? (
            <span className="table-tab add-tab disabled" title="Loading tables from Drive…" style={{ opacity: 0.5, cursor: 'default', display: 'flex', alignItems: 'center', gap: 6 }}>
              <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
              Loading…
            </span>
          ) : (
            <Link
              className="table-tab add-tab"
              to={withBook(bookId, '/table/new')}
              title="Create new table"
            >
              +
            </Link>
          )}
        </div>
        {tableId && activeSchema && (
          <div className="table-tabs-actions">
            <button
              className="btn-secondary btn-sm"
              onClick={runUndo}
              disabled={!state.canUndo}
              title="Undo (Ctrl/Cmd+Z)"
            >
              Undo
            </button>
            <Link
              className="btn-secondary btn-sm"
              to={withBook(bookId, `/table/${encodeURIComponent(tableId)}/import`)}
            >
              Import
            </Link>
            <Link
              className="btn-secondary btn-sm"
              to={withBook(bookId, '/import')}
            >
              Import New
            </Link>
            <Link
              className="btn-secondary btn-sm table-tabs-edit"
              to={withBook(bookId, `/table/${encodeURIComponent(tableId)}/edit`)}
            >
              Edit Table
            </Link>
          </div>
        )}
      </div>

      {/* Main content */}
      <div className="main-content">
        {tableId && activeSchema ? (
          <SpreadsheetGrid
            key={tableId}
            schema={activeSchema}
            rows={activeRows}
            model={state.model}
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
            revision={state.revision}
            folderId={state.folderId}
          />
        ) : (
          <div className="empty-state-main">
            <h2>No table selected</h2>
            <p>Create a new table to get started, or connect to Google Drive to load existing data.</p>
            {state.isConnecting ? (
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

  // Auto-redirect to first table
  useEffect(() => {
    if (state.tableIds.length > 0) {
      navigate(withBook(bookId, `/table/${encodeURIComponent(state.tableIds[0])}`), { replace: true });
    }
  }, [state.tableIds, navigate, bookId]);

  if (state.tableIds.length > 0) return null;

  return (
    <div className="app-body">
      <div className="table-tabs">
        <Link
          className="table-tab add-tab"
          to={withBook(bookId, '/table/new')}
          title="Create new table"
        >
          +
        </Link>
      </div>
      <div className="main-content">
        <div className="empty-state-main">
          <h2>No tables yet</h2>
          <p>Create a new table to get started, or connect to Google Drive to load existing data.</p>
          {state.isConnecting ? (
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
        </div>
      </div>
    </div>
  );
};

// --- Drive Status Button ---
const DriveStatusButton: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
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

  // Not ready yet — show loading indicator
  if (!state.driveReady) {
    return (
      <span className="drive-btn drive-btn-loading">
        <span className="drive-status-dot connecting" style={{ position: 'static', border: 'none' }} />
        Connecting…
      </span>
    );
  }

  // Not signed in
  if (!state.isSignedIn) {
    return (
      <button onClick={state.signIn} className="drive-btn drive-btn-signin">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M15 3h4a2 2 0 012 2v14a2 2 0 01-2 2h-4M10 17l5-5-5-5M15 12H3" />
        </svg>
        Sign in
      </button>
    );
  }

  // Determine sync status
  let statusIcon: React.ReactNode;
  let statusText: string;
  if (state.isConnecting) {
    statusIcon = <span className="drive-status-dot connecting" />;
    statusText = 'Connecting…';
  } else if (state.isSaving) {
    statusIcon = <span className="drive-status-dot saving" />;
    statusText = 'Saving…';
  } else if (state.isAnyDirty()) {
    statusIcon = <span className="drive-status-dot dirty" />;
    statusText = 'Unsaved changes';
  } else if (state.lastSaved) {
    statusIcon = <span className="drive-status-dot synced" />;
    statusText = `Saved ${state.lastSaved.toLocaleTimeString()}`;
  } else {
    statusIcon = <span className="drive-status-dot synced" />;
    statusText = 'Connected';
  }

  return (
    <div className="drive-status-wrapper" ref={menuRef}>
      <button
        className="drive-btn drive-btn-status"
        onClick={() => setOpen(o => !o)}
        title={statusText}
      >
        {state.userInfo?.picture ? (
          <img
            src={state.userInfo.picture}
            alt=""
            className="drive-avatar"
            referrerPolicy="no-referrer"
          />
        ) : (
          <span className="drive-avatar-placeholder">
            {state.userInfo?.name?.[0]?.toUpperCase() ?? '?'}
          </span>
        )}
        {statusIcon}
      </button>
      {open && (
        <div className="drive-dropdown">
          {state.userInfo && (
            <div className="drive-dropdown-user">
              <div className="drive-dropdown-name">{state.userInfo.name}</div>
              {state.userInfo.email && (
                <div className="drive-dropdown-email">{state.userInfo.email}</div>
              )}
            </div>
          )}
          <div className="drive-dropdown-divider" />
          <div className="drive-dropdown-status">
            {statusIcon} {statusText}
          </div>
          <div className="drive-dropdown-divider" />
          <button
            className="drive-dropdown-item"
            onClick={() => { state.signOut(); setOpen(false); }}
          >
            Sign out
          </button>
        </div>
      )}
    </div>
  );
};

const BookSidebar: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const navigate = useNavigate();
  const location = useLocation();
  const routeBookMatch = location.pathname.match(/^\/book\/([^/]+)/);
  const routeBookId = routeBookMatch ? decodeURIComponent(routeBookMatch[1]) : state.folderId;

  const currentTail = (() => {
    const m = location.pathname.match(/^\/book\/[^/]+(\/.*)?$/);
    if (m) return m[1] ?? '';
    if (location.pathname === '/') return '';
    if (location.pathname.startsWith('/table/') || location.pathname === '/table/new' || location.pathname === '/import') {
      return location.pathname;
    }
    return '';
  })();

  const onCreate = async () => {
    const name = window.prompt('New workbook name:');
    if (!name || !name.trim()) return;
    const createdId = await state.createWorkbook(name.trim());
    if (createdId) {
      navigate(`/book/${encodeURIComponent(createdId)}${currentTail}` || `/book/${encodeURIComponent(createdId)}`);
    }
  };

  const openBook = async (bookId: string) => {
    await state.switchWorkbook(bookId);
    navigate(`/book/${encodeURIComponent(bookId)}${currentTail}` || `/book/${encodeURIComponent(bookId)}`);
  };

  return (
    <aside className="book-sidebar">
      <div className="book-sidebar-header">
        <span className="book-sidebar-title">Books</span>
        <button className="btn-secondary btn-sm" onClick={() => { void onCreate(); }} disabled={state.isConnecting}>
          New
        </button>
      </div>
      <div className="book-sidebar-list">
        {state.workbooks.map(book => {
          const isActive = book.id === routeBookId;
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
                onClick={() => navigate(`/book/${encodeURIComponent(book.id)}/settings`)}
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

const BookSettingsPage: React.FC<{ state: UseAppStateReturn }> = ({ state }) => {
  const { bookId } = useParams<{ bookId: string }>();
  const navigate = useNavigate();
  const effectiveBookId = bookId ?? state.folderId ?? '';
  const currentBook = state.workbooks.find(b => b.id === effectiveBookId);

  const [name, setName] = useState(currentBook?.name ?? '');
  const [shareEmail, setShareEmail] = useState('');
  const [shareRole, setShareRole] = useState<'reader' | 'writer'>('writer');
  const [status, setStatus] = useState<string | null>(null);

  useEffect(() => {
    if (!currentBook) return;
    setName(currentBook.name);
  }, [currentBook]);

  useEffect(() => {
    if (!effectiveBookId || effectiveBookId === state.folderId) return;
    if (!state.workbooks.some(w => w.id === effectiveBookId)) return;
    void state.switchWorkbook(effectiveBookId);
  }, [effectiveBookId, state]);

  if (!currentBook) {
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
    await state.renameWorkbook(currentBook.id, trimmed);
    setStatus('Book name updated.');
  };

  const doShare = async () => {
    if (!state.isSignedIn) {
      setStatus('Sign in to share this book.');
      return;
    }
    if (!shareEmail.trim()) {
      setStatus('Enter an email address to share.');
      return;
    }
    await state.shareWorkbook(currentBook.id, shareEmail.trim(), shareRole);
    setStatus(`Shared with ${shareEmail.trim()} as ${shareRole}.`);
    setShareEmail('');
  };

  return (
    <div className="book-settings-page">
      <div className="book-settings-card">
        <div className="book-settings-header">
          <h2>Book Settings</h2>
          <button className="btn-secondary btn-sm" onClick={() => navigate(`/book/${encodeURIComponent(currentBook.id)}`)}>
            Back to Book
          </button>
        </div>

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
              Save Name
            </button>
          </div>
        </div>

        <div className="book-settings-section">
          <label className="book-settings-label">Share Book</label>
          <div className="book-settings-row">
            <input
              className="edit-table-input"
              type="email"
              value={shareEmail}
              onChange={(e) => setShareEmail(e.target.value)}
              placeholder="user@example.com"
              disabled={!state.isSignedIn}
            />
            <select
              className="workbook-toolbar-select"
              value={shareRole}
              onChange={(e) => setShareRole(e.target.value as 'reader' | 'writer')}
              disabled={!state.isSignedIn}
            >
              <option value="writer">Editor</option>
              <option value="reader">Viewer</option>
            </select>
            <button className="btn-secondary" onClick={() => { void doShare(); }} disabled={!state.isSignedIn}>
              Share
            </button>
          </div>
          {!state.isSignedIn && <div className="book-settings-note">Sign in with Google Drive to enable sharing.</div>}
        </div>

        {status && <div className="edit-table-notice edit-table-notice-info">{status}</div>}
      </div>
    </div>
  );
};

// --- Main App Shell ---
const App: React.FC = () => {
  const state = useAppState();
  const showAlert = useAlert();
  const location = useLocation();
  const navigate = useNavigate();
  const [clientIdInput, setClientIdInput] = useState(GOOGLE_CLIENT_ID);
  const [setupDone, setSetupDone] = useState(false);

  // Initialize Drive API after setup
  useEffect(() => {
    if (GOOGLE_CLIENT_ID && !state.driveReady) {
      state.initializeDrive(GOOGLE_CLIENT_ID)
        .then(() => setSetupDone(true))
        .catch(err => {
          console.error('Failed to initialize Google Drive:', err);
          setSetupDone(true); // Still allow app usage without Drive
        });
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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

  const handleSetup = async () => {
    if (!clientIdInput.trim()) return;
    try {
      await state.initializeDrive(clientIdInput.trim());
      setSetupDone(true);
    } catch (err) {
      console.error('Failed to initialize Google Drive:', err);
      showAlert('Failed to connect to Google Drive. Check your Client ID and try again.', 'Connection Error');
    }
  };

  // URL -> workbook state: switching /book/:bookId should switch active workbook.
  useEffect(() => {
    if (state.workbooks.length === 0) return;
    const m = location.pathname.match(/^\/book\/([^/]+)/);
    if (!m) return;
    const routeBookId = decodeURIComponent(m[1]);
    if (routeBookId === state.folderId) return;
    if (!state.workbooks.some(w => w.id === routeBookId)) return;
    void state.switchWorkbook(routeBookId);
  }, [location.pathname, state]);

  // Preserve backward compatibility: old table/import routes get upgraded to /book/:bookId/... when possible.
  useEffect(() => {
    if (!state.folderId) return;
    if (location.pathname === '/') {
      navigate(`/book/${encodeURIComponent(state.folderId)}`, { replace: true });
      return;
    }
    if (location.pathname.startsWith('/book/')) return;
    if (location.pathname === '/table/new' || location.pathname.startsWith('/table/') || location.pathname === '/import') {
      navigate(`/book/${encodeURIComponent(state.folderId)}${location.pathname}`, { replace: true });
    }
  }, [location.pathname, navigate, state.folderId]);


  // Setup screen
  if (!setupDone && !GOOGLE_CLIENT_ID) {
    return (
      <div className="app">
        <div className="setup-screen">
          <h1>Sheetable</h1>
          <p>Spreadsheet-like editor with Google Drive persistence</p>
          <div className="setup-card">
            <h2>Setup Google Drive Access</h2>
            <p>
              Enter your Google OAuth Client ID to connect to Google Drive.
              You can create one in the{' '}
              <a href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noopener noreferrer">
                Google Cloud Console
              </a>.
            </p>
            <div className="setup-form">
              <input
                type="text"
                value={clientIdInput}
                onChange={e => setClientIdInput(e.target.value)}
                placeholder="your-client-id.apps.googleusercontent.com"
                className="client-id-input"
              />
              <button onClick={handleSetup} className="btn-primary" disabled={!clientIdInput.trim()}>
                Connect
              </button>
            </div>
            <p className="setup-note">
              Or set VITE_GOOGLE_CLIENT_ID in a .env file and restart.
            </p>
            <hr className="setup-divider" />
            <button
              className="btn-secondary"
              onClick={() => setSetupDone(true)}
            >
              Skip — Use without Google Drive
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="app">
      {/* Top bar */}
      <header className="app-header">
        <div className="header-left">
          <Link to="/" className="app-title-link">
            <h1 className="app-title">Sheetable</h1>
          </Link>
        </div>
        <div className="header-right">
          <DriveStatusButton state={state} />
        </div>
      </header>

      <div className="app-shell">
        <BookSidebar state={state} />
        <div className="app-main">
          <Routes>
            <Route path="/" element={<HomePage state={state} />} />
            <Route path="/book/:bookId" element={<HomePage state={state} />} />
            <Route path="/book/:bookId/table/new" element={<EditTablePage state={state} />} />
            <Route path="/book/:bookId/table/:tableId" element={<TableViewPage state={state} />} />
            <Route path="/book/:bookId/settings" element={<BookSettingsPage state={state} />} />
            <Route path="/book/:bookId/table/:tableId/edit" element={<EditTablePage state={state} />} />
            <Route path="/book/:bookId/table/:tableId/import" element={<ImportPage state={state} />} />
            <Route path="/book/:bookId/import" element={<ImportPage state={state} />} />
            <Route path="/table/new" element={<EditTablePage state={state} />} />
            <Route path="/table/:tableId" element={<TableViewPage state={state} />} />
            <Route path="/table/:tableId/edit" element={<EditTablePage state={state} />} />
            <Route path="/table/:tableId/import" element={<ImportPage state={state} />} />
            <Route path="/import" element={<ImportPage state={state} />} />
            <Route path="*" element={<Navigate to="/" replace />} />
          </Routes>
        </div>
      </div>
    </div>
  );
};

export default App;

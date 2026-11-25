import { useEffect, useRef, useState } from 'react';
import { Link, useNavigate, useParams, useSearchParams } from 'react-router-dom';
import type { ReportData, ReportSheet } from '../shared/reportTypes';
import { reportName } from '../shared/reportTypes';
import * as api from './api';
import { useConfirm, usePromptInput } from './DialogProvider';
import { ReportDocument } from './ReportDocument';
import { makeReportDefinition, reportTemplates } from './reportTemplates';
import type { ReportTemplate } from './reportTemplates';
import type { UseAppStateReturn } from './useAppState';

type Props = { state: UseAppStateReturn; bookId: string; bookPath: string };

function message(error: unknown, fallback: string) {
  return error instanceof Error ? error.message : fallback;
}

function CreateReport({ state, bookId, bookPath }: Props) {
  const navigate = useNavigate();
  const [name, setName] = useState('');
  const [template, setTemplate] = useState<ReportTemplate>('editorial');
  const [table, setTable] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  async function create() {
    const title = reportName(name);
    if (!title || busy) return;
    setBusy(true);
    setError('');
    try {
      const report = await api.createReport(bookId, title, makeReportDefinition(title, template, table));
      state.updateReportTab(bookId, report);
      navigate(`${bookPath}/report/${report.id}?edit=1`, { replace: true });
    } catch (cause) {
      setError(message(cause, 'Could not create the report.'));
      setBusy(false);
    }
  }

  return (
    <main className="report-create">
      <form onSubmit={event => { event.preventDefault(); void create(); }}>
        <h1>New report</h1>
        <label className="report-create-field">
          Report name
          <input required maxLength={100} value={name} onChange={event => setName(event.target.value)} autoFocus />
        </label>
        <fieldset>
          <legend>Template</legend>
          <div className="report-template-options">
            {reportTemplates.map(option => (
              <label key={option.id} className="report-template-option" data-template={option.id}>
                <input type="radio" name="template" value={option.id} checked={template === option.id}
                  onChange={() => setTemplate(option.id)} />
                <div className="report-template-thumbnail" aria-hidden="true"><span /><span /><span /></div>
                <strong>{option.name}</strong>
                <small>{option.description}</small>
              </label>
            ))}
          </div>
        </fieldset>
        {template !== 'blank' && state.tableIds.length > 0 && (
          <label className="report-create-field">
            Add a workbook table
            <select value={table} onChange={event => setTable(event.target.value)}>
              <option value="">No table</option>
              {state.tableIds.map(id => <option key={id} value={id}>{id}</option>)}
            </select>
          </label>
        )}
        {error && <p role="alert" className="report-error">{error}</p>}
        <div className="report-toolbar-actions">
          <button className="btn-primary" disabled={busy || !reportName(name)}>{busy ? 'Creating…' : 'Create report'}</button>
          <Link to={bookPath} className="btn-secondary">Cancel</Link>
        </div>
      </form>
    </main>
  );
}

function ReportWorkspace({ state, bookId, bookPath, initial }: Props & { initial: ReportSheet }) {
  const navigate = useNavigate();
  const [params, setParams] = useSearchParams();
  const prompt = usePromptInput();
  const confirm = useConfirm();
  const [saved, setSaved] = useState(initial);
  const [draft, setDraft] = useState(initial.definition.data);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');
  const previewRef = useRef<HTMLDivElement>(null);
  const pending = useRef(false);
  const canEdit = state.activeBookRole === 'owner' || state.activeBookRole === 'editor';
  const editing = canEdit && params.get('edit') === '1';
  const dirty = JSON.stringify(draft) !== JSON.stringify(saved.definition.data);

  useEffect(() => {
    if (!dirty) return;
    const beforeUnload = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', beforeUnload);
    return () => window.removeEventListener('beforeunload', beforeUnload);
  }, [dirty]);

  function setEditing(value: boolean) {
    setParams(value ? { edit: '1' } : {}, { replace: true });
  }

  async function mutate(action: () => Promise<void>) {
    if (pending.current) return;
    pending.current = true;
    setBusy(true);
    setError('');
    setStatus('');
    try {
      await action();
    } catch (cause) {
      setError(message(cause, 'Could not update the report.'));
    } finally {
      pending.current = false;
      setBusy(false);
    }
  }

  function save(data: ReportData, name = saved.name) {
    void mutate(async () => {
      const updated = await api.updateReport(bookId, { ...saved, name, definition: { version: 1, data } });
      setSaved(updated);
      state.updateReportTab(bookId, updated);
      setStatus('Report saved');
    });
  }

  async function rename() {
    const name = reportName(await prompt(`Current name: ${saved.name}`, 'Rename report', 'Report name'));
    if (name && name !== saved.name) save(saved.definition.data, name);
  }

  async function duplicate() {
    const name = reportName(await prompt('Give the copy a name.', 'Duplicate report', 'Report name'));
    if (!name) return;
    void mutate(async () => {
      const report = await api.createReport(bookId, name, { version: 1, data: draft });
      state.updateReportTab(bookId, report);
      navigate(`${bookPath}/report/${report.id}?edit=1`);
    });
  }

  async function remove() {
    if (!await confirm(`Delete “${saved.name}” from this workbook?`, 'Delete report')) return;
    void mutate(async () => {
      await api.deleteReport(bookId, saved.id);
      state.removeReportTab(bookId, saved.id);
      navigate(bookPath, { replace: true });
    });
  }

  async function print() {
    const page = previewRef.current;
    if (!page) return;
    setError('');
    if (page.querySelector('[data-report-data="error"]')) {
      setError('Retry loading the workbook data before printing.');
      return;
    }
    if (page.querySelector('[data-report-data="loading"]')) {
      setError('Wait for the workbook data to finish loading before printing.');
      return;
    }
    const images = [...page.querySelectorAll('img')];
    try {
      await Promise.all([document.fonts.ready, ...images.map(image => image.decode())]);
      if (previewRef.current === page) window.print();
    } catch {
      setError('A report image could not be loaded. Replace or remove it before printing.');
    }
  }

  return (
    <main className="report-workspace">
      <div className="report-toolbar">
        <div className="report-toolbar-title">
          <strong>{saved.name}</strong>
          {dirty && <span>Unsaved changes</span>}
          {status && !dirty && <span role="status">{status}</span>}
        </div>
        <div className="report-toolbar-actions">
          {editing ? (
            <>
              <button className="btn-secondary btn-sm" onClick={() => setEditing(false)}>Preview</button>
              <button className="btn-primary btn-sm" disabled={busy || !dirty} onClick={() => save(draft)}>Save report</button>
            </>
          ) : (
            <>
              {canEdit && <button className="btn-primary btn-sm" onClick={() => setEditing(true)}>Edit report</button>}
              {canEdit && dirty && <button className="btn-secondary btn-sm" disabled={busy}
                onClick={() => save(draft)}>Save report</button>}
              <button className="btn-secondary btn-sm" onClick={() => void print()}>Print / Save PDF</button>
              {canEdit && (
                <>
                  <button className="btn-secondary btn-sm" disabled={busy} onClick={() => void rename()}>Rename</button>
                  <button className="btn-secondary btn-sm" disabled={busy} onClick={() => void duplicate()}>Duplicate</button>
                  <button className="btn-secondary btn-sm" disabled={busy} onClick={() => void remove()}>Delete</button>
                </>
              )}
            </>
          )}
        </div>
      </div>
      {error && <p className="report-error" role="alert">{error}</p>}
      {busy && <div className="report-operation" role="status">Saving changes…</div>}
      {editing ? (
        <div className="report-editor">
          <ReportDocument state={state} bookId={bookId} data={draft} editing onChange={setDraft} />
        </div>
      ) : (
        <div className="report-preview" ref={previewRef}>
          <ReportDocument state={state} bookId={bookId} data={draft} />
        </div>
      )}
    </main>
  );
}

function LoadReport({ state, bookId, bookPath, id }: Props & { id: string }) {
  const [report, setReport] = useState<ReportSheet>();
  const [error, setError] = useState('');
  const [attempt, setAttempt] = useState(0);
  useEffect(() => {
    let current = true;
    api.getReport(bookId, id).then(result => {
      if (current) setReport(result);
    }).catch(cause => {
      if (current) setError(message(cause, 'Could not open the report.'));
    });
    return () => { current = false; };
  }, [bookId, id, attempt]);

  if (report) return <ReportWorkspace state={state} bookId={bookId} bookPath={bookPath} initial={report} />;
  if (error) {
    return (
      <div className="empty-state-main">
        <p role="alert">{error}</p>
        <button className="btn-secondary" onClick={() => { setError(''); setAttempt(attempt + 1); }}>Retry</button>
        <Link to={bookPath}>Back to workbook</Link>
      </div>
    );
  }
  return <div className="empty-state-main" role="status">Loading report…</div>;
}

export default function ReportSheetPage({ state }: { state: UseAppStateReturn }) {
  const { bookId: bookName, reportId } = useParams();
  const bookId = state.activeBookId;
  if (!bookId || !bookName || !reportId) return <div className="empty-state-main">Report not found</div>;
  const bookPath = `/book/${encodeURIComponent(bookName)}`;
  if (reportId === 'new') {
    return state.activeBookRole === 'owner' || state.activeBookRole === 'editor'
      ? <CreateReport state={state} bookId={bookId} bookPath={bookPath} />
      : <div className="empty-state-main">Only workbook editors can create a report.</div>;
  }
  return <LoadReport key={`${bookId}:${reportId}`} state={state} bookId={bookId} bookPath={bookPath} id={reportId} />;
}

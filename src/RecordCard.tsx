/**
 * RecordCard — generic modal for viewing and editing a single row.
 *
 * Reusable across any context that needs to open a row individually:
 * calendar date click, event click, future list-view row click, etc.
 */
import React, { useState, useCallback } from 'react';
import type { TableSchema, Row, ValidationError, ColumnDef } from './types';
import { INTERNAL_ROW_ID } from './types';
import * as api from './api';
import { ImageListEditor } from './ImageListEditor';
import RefCellEditor from './RefCellEditor';
import { RefListEditor, resolveRefLabel } from './RefListEditor';
import { parseListItems, serializeListItems, isValidListItem } from './listValue';

interface RecordCardProps {
  schema: TableSchema;
  /** Dialog heading */
  title: string;
  /** Pre-populated field values (for create: partial; for edit: full row) */
  initialValues: Row;
  /** When true, shows values without edit controls */
  readOnly?: boolean;
  /** Called with the final values. Return [] on success, or errors to display. */
  onSave?: (values: Row) => ValidationError[];
  /** Close / cancel */
  onClose: () => void;
  /** Provides options for reference-type columns */
  getReferenceRows: (refTable: string) => Row[];
  /** Resolves nested column paths used by refDisplayColumns in reference labels */
  resolveColumnPath?: (tableName: string, row: Row, path: string) => string;
  /** Opens a create-record flow for a referenced table and resolves with the new row id */
  onCreateReferenceRow?: (refTable: string, seedText: string) => Promise<string | null>;
  /** Needed to render image URLs */
  bookId?: string | null;
}

// ---- Reference display label ------------------------------------------------

function refDisplayLabel(refRow: Row, col: ColumnDef, refTable: string, resolveColumnPath?: (tableName: string, row: Row, path: string) => string): string {
  if (col.refDisplayColumns && col.refDisplayColumns.length > 0) {
    const parts = col.refDisplayColumns
      .map(c => {
        if (resolveColumnPath) return resolveColumnPath(refTable, refRow, c);
        return refRow[c] ?? '';
      })
      .filter(Boolean);
    if (parts.length) return parts.join(' ');
  }
  const firstKey = Object.keys(refRow).find(k => k !== INTERNAL_ROW_ID);
  return (firstKey ? refRow[firstKey] : undefined) ?? refRow[INTERNAL_ROW_ID] ?? '';
}

// ---- Image field -------------------------------------------------------------

const ImageField: React.FC<{
  value: string;
  bookId: string;
  onChange: (key: string) => void;
  readOnly: boolean;
}> = ({ value, bookId, onChange, readOnly }) => {
  const [uploading, setUploading] = useState(false);

  const handleUpload = useCallback(() => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.style.display = 'none';
    document.body.appendChild(input);
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      document.body.removeChild(input);
      if (!file) return;
      setUploading(true);
      try {
        const { key, uploadUrl } = await api.getUploadUrl(bookId, file.name);
        await api.uploadImage(uploadUrl, file);
        onChange(key);
      } catch (err) {
        console.error('Upload failed', err);
      } finally {
        setUploading(false);
      }
    });
    input.addEventListener('cancel', () => document.body.removeChild(input));
    input.click();
  }, [bookId, onChange]);

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
      {value ? (
        <img
          src={api.imageUrl(bookId, value)}
          alt=""
          style={{ height: 64, maxWidth: 160, objectFit: 'cover', borderRadius: 6, border: '1px solid var(--color-border)' }}
        />
      ) : (
        <div style={{ height: 64, width: 100, borderRadius: 6, border: '1px dashed var(--color-border)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, color: 'var(--color-text-muted)' }}>
          No image
        </div>
      )}
      {!readOnly && (
        <button onClick={handleUpload} disabled={uploading} className="app-dialog-btn app-dialog-btn-secondary btn-sm">
          {uploading ? 'Uploading…' : value ? 'Change' : 'Upload'}
        </button>
      )}
      {!readOnly && value && (
        <button onClick={() => onChange('')} className="app-dialog-btn app-dialog-btn-secondary btn-sm" style={{ color: 'var(--color-danger)' }}>
          Remove
        </button>
      )}
    </div>
  );
};

// ---- Plain list field (text / number / bool / date items) ---------------------

const ListField: React.FC<{ value: string; listOf: string; onChange: (value: string) => void; readOnly: boolean }> = ({ value, listOf, onChange, readOnly }) => {
  const items = parseListItems(value);
  const [draft, setDraft] = useState('');
  const [invalid, setInvalid] = useState(false);
  const inputType = listOf === 'integer' || listOf === 'decimal' ? 'number' : listOf === 'date' ? 'date' : listOf === 'datetime' ? 'datetime-local' : 'text';

  const add = () => {
    const item = draft.trim();
    if (!item) return;
    if (!isValidListItem(listOf, item)) { setInvalid(true); return; }
    onChange(serializeListItems([...items, item]));
    setDraft('');
    setInvalid(false);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
      {items.length > 0 ? (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
          {items.map((item, index) => (
            <span key={`${item}-${index}`} style={{ background: 'var(--ag-row-hover-color, #e8f0fe)', borderRadius: 3, padding: '1px 4px 1px 7px', fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 2, lineHeight: '18px' }}>
              {item}
              {!readOnly && (
                <button type="button" aria-label="Remove" onClick={() => onChange(serializeListItems(items.filter((_, i) => i !== index)))} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'inherit', padding: '0 1px', fontSize: 14, lineHeight: 1 }}>×</button>
              )}
            </span>
          ))}
        </div>
      ) : readOnly ? <span style={valueStyle}>—</span> : null}
      {!readOnly && (
        <div style={{ display: 'flex', gap: 6 }}>
          {listOf === 'bool' ? (
            <select className="app-dialog-input" style={{ ...fieldInputStyle, flex: 1 }} value={draft} onChange={e => { setDraft(e.target.value); setInvalid(false); }}>
              <option value="">Select…</option>
              <option value="true">Yes (true)</option>
              <option value="false">No (false)</option>
            </select>
          ) : (
            <input
              type={inputType}
              step={listOf === 'decimal' ? 'any' : undefined}
              value={draft}
              placeholder="Type and press Enter to add…"
              onChange={e => { setDraft(e.target.value); setInvalid(false); }}
              onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); add(); } }}
              className="app-dialog-input"
              style={{ ...fieldInputStyle, flex: 1, ...(invalid ? { borderColor: 'var(--color-danger)' } : {}) }}
            />
          )}
          <button type="button" className="app-dialog-btn app-dialog-btn-secondary btn-sm" onClick={add} disabled={!draft.trim()}>Add</button>
        </div>
      )}
    </div>
  );
};

// ---- Toggle for bool fields --------------------------------------------------

const BoolToggle: React.FC<{ value: boolean; onChange: (v: boolean) => void; readOnly: boolean }> = ({ value, onChange, readOnly }) => (
  <button
    type="button"
    disabled={readOnly}
    onClick={() => !readOnly && onChange(!value)}
    style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 8,
      background: 'none',
      border: 'none',
      cursor: readOnly ? 'default' : 'pointer',
      padding: 0,
    }}
  >
    <span style={{
      display: 'inline-block',
      width: 36,
      height: 20,
      borderRadius: 10,
      background: value ? 'var(--color-primary)' : 'var(--color-border)',
      position: 'relative',
      transition: 'background 0.15s',
    }}>
      <span style={{
        display: 'block',
        position: 'absolute',
        top: 2,
        left: value ? 18 : 2,
        width: 16,
        height: 16,
        borderRadius: '50%',
        background: '#fff',
        transition: 'left 0.15s',
        boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
      }} />
    </span>
    <span style={{ fontSize: 13, color: 'var(--color-text)' }}>{value ? 'Yes' : 'No'}</span>
  </button>
);

// ---- Main RecordCard --------------------------------------------------------

export const RecordCard: React.FC<RecordCardProps> = ({
  schema,
  title,
  initialValues,
  readOnly = false,
  onSave,
  onClose,
  getReferenceRows,
  resolveColumnPath,
  onCreateReferenceRow,
  bookId,
}) => {
  const [values, setValues] = useState<Row>({ ...initialValues });
  const [errors, setErrors] = useState<ValidationError[]>([]);

  const set = (name: string, value: string) =>
    setValues(prev => ({ ...prev, [name]: value }));

  const handleSave = () => {
    if (!onSave) return;
    const errs = onSave(values);
    if (errs.length > 0) {
      setErrors(errs);
    } else {
      setErrors([]);
    }
  };

  const editableColumns = schema.columns.filter(c => c.name !== INTERNAL_ROW_ID);

  return (
    <div
      className="app-dialog-overlay"
      onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        className="app-dialog record-card-dialog"
        style={{
          width: 'min(520px, 94vw)',
          maxHeight: '88vh',
          display: 'flex',
          flexDirection: 'column',
          color: 'var(--color-text)',
          overflow: 'hidden',
        }}
      >
        {/* Header */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '14px 20px',
          borderBottom: '1px solid var(--color-border)',
          flexShrink: 0,
        }}>
          <span style={{ fontWeight: 600, fontSize: 15 }}>{title}</span>
          <button
            onClick={onClose}
            className="app-dialog-close"
            aria-label="Close"
          >
            ×
          </button>
        </div>

        {/* Body */}
        <div style={{ padding: '16px 20px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: 14 }}>
          {editableColumns.map(col => {
            const label = col.displayName ?? col.name;
            const value = values[col.name] ?? '';

            return (
              <div key={col.name} style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                <label style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                  {label}
                </label>

                {col.type === 'bool' ? (
                  <BoolToggle
                    value={value === 'true' || value === '1'}
                    onChange={v => set(col.name, v ? 'true' : 'false')}
                    readOnly={readOnly}
                  />
                ) : col.type === 'image' ? (
                  bookId ? (
                    <ImageField
                      value={value}
                      bookId={bookId}
                      onChange={key => set(col.name, key)}
                      readOnly={readOnly}
                    />
                  ) : (
                    <span style={{ fontSize: 13, color: 'var(--color-text-muted)' }}>{value || '—'}</span>
                  )
                ) : col.type === 'list' && col.listOf === 'image' ? (
                  bookId ? (
                    <ImageListEditor value={value} bookId={bookId} onChange={v => set(col.name, v)} readOnly={readOnly} />
                  ) : (
                    <span style={{ fontSize: 13, color: 'var(--color-text-muted)' }}>{parseListItems(value).length ? `${parseListItems(value).length} image(s)` : '—'}</span>
                  )
                ) : col.type === 'list' && col.listOf === 'reference' && col.refTable ? (
                  readOnly ? (
                    <span style={valueStyle}>
                      {parseListItems(value).map(id => resolveRefLabel(id, getReferenceRows(col.refTable!), col.refTable!, resolveColumnPath ?? ((_t, row, path) => row[path] ?? ''), col.refDisplayColumns ?? [], col.refSearchColumns ?? [])).join(', ') || '—'}
                    </span>
                  ) : (
                    <RefListEditor
                      value={value}
                      onChange={v => set(col.name, v)}
                      refRows={getReferenceRows(col.refTable)}
                      refTable={col.refTable}
                      resolveColumnPath={resolveColumnPath ?? ((_t, row, path) => row[path] ?? '')}
                      searchColumns={col.refSearchColumns ?? []}
                      displayColumns={col.refDisplayColumns ?? []}
                      onCreateRecord={onCreateReferenceRow}
                      autoFocus={false}
                    />
                  )
                ) : col.type === 'list' ? (
                  <ListField
                    value={value}
                    listOf={col.listOf ?? 'text'}
                    onChange={v => set(col.name, v)}
                    readOnly={readOnly}
                  />
                ) : col.type === 'reference' && col.refTable ? (
                  readOnly ? (
                    <span style={valueStyle}>
                      {refDisplayLabel(
                        getReferenceRows(col.refTable).find(r => r[INTERNAL_ROW_ID] === value) ?? {},
                        col,
                        col.refTable,
                        resolveColumnPath,
                      ) || '—'}
                    </span>
                  ) : (
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <RefCellEditor
                        key={`${col.name}:${value}`}
                        value={value}
                        onValueChange={next => set(col.name, next ?? '')}
                        stopEditing={() => {}}
                        refRows={getReferenceRows(col.refTable)}
                        refTable={col.refTable}
                        resolveColumnPath={resolveColumnPath ?? ((_t, row, path) => row[path] ?? '')}
                        searchColumns={col.refSearchColumns ?? []}
                        displayColumns={col.refDisplayColumns ?? []}
                        onCreateRecord={onCreateReferenceRow}
                        autoFocus={false}
                        placeholder={value
                          ? refDisplayLabel(
                              getReferenceRows(col.refTable).find(refRow => refRow[INTERNAL_ROW_ID] === value) ?? {},
                              col,
                              col.refTable,
                              resolveColumnPath,
                            ) || 'Search…'
                          : 'None — search…'}
                        inputClassName="app-dialog-input"
                        inputStyle={{ marginBottom: 0, width: '100%', boxSizing: 'border-box' }}
                      />
                    </div>
                  )
                ) : col.type === 'date' ? (
                  readOnly ? (
                    <span style={valueStyle}>{value || '—'}</span>
                  ) : (
                    <input type="date" value={value} onChange={e => set(col.name, e.target.value)} className="app-dialog-input" style={fieldInputStyle} />
                  )
                ) : col.type === 'datetime' ? (
                  readOnly ? (
                    <span style={valueStyle}>{value ? new Date(value).toLocaleString() : '—'}</span>
                  ) : (
                    <input
                      type="datetime-local"
                      value={value ? value.replace('Z', '').slice(0, 16) : ''}
                      onChange={e => set(col.name, e.target.value ? new Date(e.target.value).toISOString() : '')}
                      className="app-dialog-input"
                      style={fieldInputStyle}
                    />
                  )
                ) : (
                  readOnly ? (
                    <span style={valueStyle}>{value || '—'}</span>
                  ) : (
                    <input
                      type={col.type === 'integer' || col.type === 'decimal' ? 'number' : 'text'}
                      value={value}
                      onChange={e => set(col.name, e.target.value)}
                      className="app-dialog-input"
                      style={fieldInputStyle}
                      step={col.type === 'decimal' ? 'any' : undefined}
                    />
                  )
                )}
              </div>
            );
          })}

          {errors.length > 0 && (
            <div style={{ padding: '8px 12px', background: 'var(--color-cell-error)', color: 'var(--color-danger)', borderRadius: 6, fontSize: 13 }}>
              {errors.map((e, i) => <div key={i}>{e.message}</div>)}
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          gap: 8,
          padding: '12px 20px',
          borderTop: '1px solid var(--color-border)',
          flexShrink: 0,
          background: 'var(--color-surface)',
        }}>
          <button onClick={onClose} className="app-dialog-btn app-dialog-btn-secondary">
            {readOnly ? 'Close' : 'Cancel'}
          </button>
          {!readOnly && onSave && (
            <button onClick={handleSave} className="app-dialog-btn app-dialog-btn-primary">
              Save
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

// ---- Shared styles ----------------------------------------------------------

const fieldInputStyle: React.CSSProperties = {
  marginBottom: 0,
};

const valueStyle: React.CSSProperties = {
  fontSize: 13,
  color: 'var(--color-text)',
  padding: '7px 0',
};


/**
 * RecordCard — generic modal for viewing and editing a single row.
 *
 * Reusable across any context that needs to open a row individually:
 * calendar date click, event click, future list-view row click, etc.
 */
import React, { useState, useCallback } from 'react';
import Select from 'react-select';
import { dialogSelectStyles } from './selectStyles';
import type { TableSchema, Row, ValidationError, ColumnDef } from './types';
import { INTERNAL_ROW_ID } from './types';
import * as api from './api';

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
  const [creatingReferenceCol, setCreatingReferenceCol] = useState<string | null>(null);
  const [refInputValues, setRefInputValues] = useState<Record<string, string>>({});

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
                      <Select
                        styles={dialogSelectStyles}
                        value={value
                          ? (() => {
                              const r = getReferenceRows(col.refTable).find(refRow => refRow[INTERNAL_ROW_ID] === value);
                              return r
                                ? { value: String(value), label: refDisplayLabel(r, col, col.refTable!, resolveColumnPath) }
                                : null;
                            })()
                          : null}
                        options={[
                          ...getReferenceRows(col.refTable).map(refRow => ({
                            value: String(refRow[INTERNAL_ROW_ID] ?? ''),
                            label: refDisplayLabel(refRow, col, col.refTable!, resolveColumnPath),
                          })),
                          ...(onCreateReferenceRow && (refInputValues[col.name] ?? '').trim()
                            ? [{ value: '__create__', label: `+ Add new record in ${col.refTable}` }]
                            : []),
                        ]}
                        onChange={opt => {
                          if (!opt || opt.value !== '__create__') {
                            set(col.name, opt?.value ?? '');
                            return;
                          }
                          if (!col.refTable || !onCreateReferenceRow) return;
                          const seed = (refInputValues[col.name] ?? '').trim();
                          setCreatingReferenceCol(col.name);
                          void (async () => {
                            try {
                              const createdRowId = await onCreateReferenceRow(col.refTable!, seed);
                              if (createdRowId) set(col.name, createdRowId);
                            } finally {
                              setCreatingReferenceCol(prev => prev === col.name ? null : prev);
                            }
                          })();
                        }}
                        onInputChange={(v, { action }) => {
                          if (action === 'input-change') setRefInputValues(prev => ({ ...prev, [col.name]: v }));
                        }}
                        inputValue={refInputValues[col.name] ?? undefined}
                        isClearable
                        isLoading={creatingReferenceCol === col.name}
                        placeholder="— none —"
                        menuPlacement="auto"
                        formatOptionLabel={(opt) =>
                          opt.value === '__create__'
                            ? <span style={{ color: 'var(--color-primary)', fontWeight: 600 }}>{opt.label}</span>
                            : <>{opt.label}</>
                        }
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


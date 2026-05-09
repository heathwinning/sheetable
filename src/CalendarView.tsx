import React, { useMemo, useState, useRef, useEffect } from 'react';
import type { TableSchema, Row, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseTemporalUnknown } from './dateFormat';
import { format } from 'date-fns/format';
import { CalendarScrollView, AgendaView } from './CalendarScrollView';
import { RecordCard } from './RecordCard';

const DEFAULT_RESOLVE = (_row: Row, _path: string) => '';

// Simplified event type (no BigCalendar dependency)
interface CalEvent {
  id: unknown;
  title: string;
  start: Date;
  allDay: boolean;
  resource: Row;
}

interface CalendarViewProps {
  schema: TableSchema;
  rows: Row[];
  dateColumn: string;
  onDateColumnChange: (col: string) => void;
  resolveColumnPath?: (row: Row, path: string) => string;
  onCreateRow?: (row: Row) => ValidationError[];
  onUpdateField?: (rowIndex: number, column: string, value: string) => ValidationError[];
  getReferenceRows?: (refTable: string) => Row[];
  readOnly?: boolean;
  bookId?: string | null;
  /** Stable key for persisting column config (e.g. bookId + tableId) */
  configKey?: string;
}

interface DialogState {
  open: boolean;
  title: string;
  initialValues: Row;
  rowIndex?: number; // defined when editing an existing row
}

export const CalendarView: React.FC<CalendarViewProps> = ({
  schema, rows, dateColumn, onDateColumnChange,
  resolveColumnPath = DEFAULT_RESOLVE,
  onCreateRow, onUpdateField, getReferenceRows,
  readOnly = false,
  bookId,
  configKey,
}) => {
  const [viewMode, setViewModeRaw] = useState<'calendar' | 'agenda'>(() => {
    if (!configKey) return 'calendar';
    const saved = localStorage.getItem(`sheetable-cal-mode-${configKey}`);
    return saved === 'agenda' ? 'agenda' : 'calendar';
  });
  const setViewMode = (mode: 'calendar' | 'agenda') => {
    setViewModeRaw(mode);
    if (configKey) localStorage.setItem(`sheetable-cal-mode-${configKey}`, mode);
  };
  const [dialog, setDialog] = useState<DialogState>({ open: false, title: '', initialValues: {} });
  const [colPickerOpen, setColPickerOpen] = useState(false);
  const colPickerRef = useRef<HTMLDivElement>(null);

  // Columns available for display on event chips (all non-date, non-internal)
  const displayableColumns = useMemo(
    () => schema.columns.filter(c => c.name !== INTERNAL_ROW_ID && c.type !== 'date' && c.type !== 'datetime' && c.type !== 'image'),
    [schema.columns],
  );

  // Persist chosen display columns per table
  const colStorageKey = configKey ? `sheetable-cal-cols-${configKey}` : null;
  const [displayColumnNames, setDisplayColumnNames] = useState<string[]>(() => {
    if (!colStorageKey) return [];
    try {
      const saved = localStorage.getItem(colStorageKey);
      if (saved) return JSON.parse(saved) as string[];
    } catch { /* ignore */ }
    return [];
  });

  // Default: first non-reference text column, else first displayable
  const effectiveDisplayColumns = useMemo(() => {
    const valid = displayColumnNames.filter(n => displayableColumns.some(c => c.name === n));
    if (valid.length > 0) return valid;
    const first = displayableColumns.find(c => c.type !== 'reference') ?? displayableColumns[0];
    return first ? [first.name] : [];
  }, [displayColumnNames, displayableColumns]);

  const toggleDisplayColumn = (name: string) => {
    setDisplayColumnNames(prev => {
      const next = prev.includes(name) ? prev.filter(n => n !== name) : [...prev, name];
      if (colStorageKey) localStorage.setItem(colStorageKey, JSON.stringify(next));
      return next;
    });
  };

  // Close col picker on outside click
  useEffect(() => {
    if (!colPickerOpen) return;
    const handler = (e: MouseEvent) => {
      if (colPickerRef.current && !colPickerRef.current.contains(e.target as Node)) {
        setColPickerOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [colPickerOpen]);

  const dateColumns = useMemo(
    () => schema.columns.filter(c => c.type === 'date' || c.type === 'datetime'),
    [schema.columns],
  );

  const availableYears = useMemo(() => {
    const years = new Set<number>();
    for (const row of rows) {
      const d = parseTemporalUnknown(row[dateColumn]);
      if (d) years.add(d.getFullYear());
    }
    const currentYear = new Date().getFullYear();
    years.add(currentYear);
    return [...years].sort((a, b) => a - b);
  }, [rows, dateColumn]);

  const yearStorageKey = configKey ? `sheetable-cal-year-${configKey}` : null;
  const [selectedYear, setSelectedYearRaw] = useState<number>(() => {
    const currentYear = new Date().getFullYear();
    if (!yearStorageKey) return currentYear;
    const saved = Number(localStorage.getItem(yearStorageKey));
    return Number.isFinite(saved) ? saved : currentYear;
  });
  const setSelectedYear = (year: number) => {
    setSelectedYearRaw(year);
    if (yearStorageKey) localStorage.setItem(yearStorageKey, String(year));
  };

  useEffect(() => {
    if (!availableYears.includes(selectedYear)) {
      setSelectedYear(availableYears[availableYears.length - 1]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [availableYears]);

  // Whether the active date column is date-only
  const isDateOnly = useMemo(
    () => schema.columns.find(c => c.name === dateColumn)?.type === 'date',
    [schema.columns, dateColumn],
  );

  const events = useMemo(() => rows.map(row => {
    const raw = row[dateColumn];
    const date = parseTemporalUnknown(raw);
    if (!date) return null;
    // Build title from configured display columns
    const parts = effectiveDisplayColumns.map(name => {
      const col = schema.columns.find(c => c.name === name);
      if (!col) return '';
      return col.type === 'reference'
        ? resolveColumnPath(row, name)
        : (row[name] ?? '');
    }).filter(Boolean);
    const title = parts.join(' · ') || format(date, 'd MMM yyyy');
    return {
      id: row[INTERNAL_ROW_ID] ?? Math.random(),
      title,
      start: date,
      end: date,
      allDay: isDateOnly,
      resource: row,
    };
  }).filter(Boolean) as CalEvent[], [rows, dateColumn, effectiveDisplayColumns, schema.columns, isDateOnly, resolveColumnPath]);

  const hasPicker = dateColumns.length > 1;

  // Build a blank row pre-filled with the given date
  const blankRowForDate = (date: Date): Row => {
    const isDateOnly = schema.columns.find(c => c.name === dateColumn)?.type === 'date';
    const dateValue = isDateOnly
      ? format(date, 'yyyy-MM-dd')
      : date.toISOString();
    const row: Row = {};
    for (const col of schema.columns) {
      if (col.name === INTERNAL_ROW_ID) continue;
      row[col.name] = col.name === dateColumn ? dateValue : '';
    }
    return row;
  };

  const handleSelectSlot = (date: Date) => {
    if (readOnly || !onCreateRow) return;
    setDialog({
      open: true,
      title: 'New Entry',
      initialValues: blankRowForDate(date),
    });
  };

  const handleSelectEvent = (ev: CalEvent) => {
    const row = ev.resource;
    if (!row) return;
    const rowIndex = rows.findIndex(r => r[INTERNAL_ROW_ID] === row[INTERNAL_ROW_ID]);
    if (rowIndex < 0) return;
    const initialValues: Row = {};
    for (const col of schema.columns) {
      if (col.name === INTERNAL_ROW_ID) continue;
      initialValues[col.name] = row[col.name] ?? '';
    }
    setDialog({ open: true, title: 'Edit Entry', initialValues, rowIndex });
  };

  const handleDialogSave = (values: Row): ValidationError[] => {
    if (dialog.rowIndex !== undefined && onUpdateField) {
      // Edit mode: apply each changed field
      let errs: ValidationError[] = [];
      for (const col of schema.columns) {
        if (col.name === INTERNAL_ROW_ID) continue;
        const newVal = values[col.name] ?? '';
        const oldVal = rows[dialog.rowIndex]?.[col.name] ?? '';
        if (newVal !== oldVal) {
          errs = onUpdateField(dialog.rowIndex, col.name, newVal);
          if (errs.length > 0) return errs;
        }
      }
      setDialog(d => ({ ...d, open: false }));
      return [];
    } else if (onCreateRow) {
      const errs = onCreateRow(values);
      if (errs.length === 0) setDialog(d => ({ ...d, open: false }));
      return errs;
    }
    return [];
  };

  const viewToggleStyle = (active: boolean): React.CSSProperties => ({
    padding: '4px 10px',
    fontSize: 13,
    fontWeight: active ? 600 : 500,
    background: 'transparent',
    color: active ? 'var(--color-primary)' : 'var(--color-text-muted)',
    border: '1px solid var(--color-border)',
    borderBottom: active ? '2px solid var(--color-primary)' : '1px solid var(--color-border)',
    borderRadius: 4,
    cursor: 'pointer',
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'var(--color-surface)' }}>
      {/* Toolbar row */}
      <div style={{ padding: '6px 12px', borderBottom: '1px solid var(--color-border)', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', background: 'var(--color-surface)' }}>
        <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
          <button style={viewToggleStyle(viewMode === 'calendar')} onClick={() => setViewMode('calendar')}>Calendar</button>
          <button style={viewToggleStyle(viewMode === 'agenda')} onClick={() => setViewMode('agenda')}>Agenda</button>
          <select
            value={selectedYear}
            onChange={(e) => setSelectedYear(Number(e.target.value))}
            style={{
              fontSize: 13,
              padding: '3px 8px',
              borderRadius: 4,
              border: '1px solid var(--color-border)',
              background: 'var(--color-surface)',
              color: 'var(--color-text)',
              cursor: 'pointer',
            }}
            title="Year"
          >
            {availableYears.map((year) => (
              <option key={year} value={year}>{year}</option>
            ))}
          </select>
          <button
            style={viewToggleStyle(false)}
            onClick={() => setSelectedYear(new Date().getFullYear())}
            title="Jump to current year"
          >
            This year
          </button>

          {/* Consolidated view/event configuration */}
          <div ref={colPickerRef} style={{ position: 'relative', marginLeft: 4 }}>
            <button
              onClick={() => setColPickerOpen(o => !o)}
              title="Configure view and event text"
              style={{
                ...viewToggleStyle(colPickerOpen),
                padding: '3px 10px',
              }}
            >Configure</button>
            {colPickerOpen && (
              <div style={{
                position: 'absolute',
                top: 'calc(100% + 4px)',
                right: 0,
                background: 'var(--color-surface)',
                border: '1px solid var(--color-border)',
                borderRadius: 8,
                boxShadow: '0 4px 16px rgba(0,0,0,0.15)',
                padding: '10px 14px',
                minWidth: 200,
                zIndex: 500,
              }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--color-text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 8 }}>
                  View
                </div>
                {hasPicker && (
                  <label style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12, fontSize: 13, color: 'var(--color-text)' }}>
                    <span>Date field</span>
                    <select
                      value={dateColumn}
                      onChange={e => onDateColumnChange(e.target.value)}
                      style={{ fontSize: 13, padding: '4px 8px', borderRadius: 6, border: '1px solid var(--color-border)', background: 'var(--color-surface)', color: 'var(--color-text)' }}
                    >
                      {dateColumns.map(c => (
                        <option key={c.name} value={c.name}>{c.displayName ?? c.name}</option>
                      ))}
                    </select>
                  </label>
                )}

                <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--color-text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 8 }}>
                  Event text
                </div>
                {displayableColumns.length === 0 ? (
                  <div style={{ fontSize: 13, color: 'var(--color-text-muted)' }}>No columns available</div>
                ) : displayableColumns.map(col => {
                  const checked = effectiveDisplayColumns.includes(col.name);
                  return (
                    <label key={col.name} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', cursor: 'pointer', fontSize: 13, color: 'var(--color-text)' }}>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleDisplayColumn(col.name)}
                        style={{ accentColor: 'var(--color-primary)', width: 15, height: 15 }}
                      />
                      {col.displayName ?? col.name}
                      <span style={{ fontSize: 11, color: 'var(--color-text-muted)', marginLeft: 'auto' }}>{col.type}</span>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>

      {viewMode === 'calendar' ? (
        <CalendarScrollView
          events={events}
          year={selectedYear}
          onSelectDate={!readOnly && onCreateRow ? handleSelectSlot : undefined}
          onSelectEvent={(scrollEv) => handleSelectEvent(scrollEv as CalEvent)}
        />
      ) : (
        <div style={{ flex: 1, minHeight: 0, overflowY: 'auto', padding: '12px 16px 24px', maxWidth: 720, margin: '0 auto', width: '100%', boxSizing: 'border-box', color: 'var(--color-text)' }}>
          <AgendaView
            events={events}
            year={selectedYear}
            todayRef={{ current: null }}
            onSelectEvent={(scrollEv) => handleSelectEvent(scrollEv as CalEvent)}
          />
        </div>
      )}
      {dialog.open && (
        <RecordCard
          schema={schema}
          title={dialog.title}
          initialValues={dialog.initialValues}
          onSave={handleDialogSave}
          onClose={() => setDialog(d => ({ ...d, open: false }))}
          getReferenceRows={getReferenceRows ?? (() => [])}
          bookId={bookId}
          readOnly={dialog.rowIndex !== undefined ? (readOnly || !onUpdateField) : (readOnly || !onCreateRow)}
        />
      )}
    </div>
  );
};

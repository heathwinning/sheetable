import React, { useMemo, useState } from 'react';
import { Calendar as BigCalendar, dateFnsLocalizer, type Event as RBCEvent, type View, type SlotInfo } from 'react-big-calendar';
import { format } from 'date-fns/format';
import { parse } from 'date-fns/parse';
import { startOfWeek } from 'date-fns/startOfWeek';
import { getDay } from 'date-fns/getDay';
import { enUS } from 'date-fns/locale/en-US';
import type { TableSchema, Row, ValidationError } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseTemporalUnknown } from './dateFormat';
import { CalendarScrollView, AgendaView } from './CalendarScrollView';
import { RecordCard } from './RecordCard';
import 'react-big-calendar/lib/css/react-big-calendar.css';

const DEFAULT_RESOLVE = (_row: Row, _path: string) => '';

const locales = { 'en-US': enUS };
const localizer = dateFnsLocalizer({
  format,
  parse,
  startOfWeek: () => startOfWeek(new Date(), { weekStartsOn: 0 }),
  getDay,
  locales,
});

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
}) => {
  const [currentDate, setCurrentDate] = useState(new Date());
  const [currentView, setCurrentView] = useState<View>('month');
  const [viewMode, setViewMode] = useState<'grid' | 'scroll' | 'agenda'>('grid');
  const [dialog, setDialog] = useState<DialogState>({ open: false, title: '', initialValues: {} });

  const dateColumns = useMemo(
    () => schema.columns.filter(c => c.type === 'date' || c.type === 'datetime'),
    [schema.columns],
  );
  // Prefer non-reference, non-image columns for label; fall back to any non-date column
  const labelColumn = useMemo(() => {
    const nonDate = schema.columns.filter(c => c.name !== INTERNAL_ROW_ID && c.type !== 'date' && c.type !== 'datetime' && c.type !== 'image');
    return nonDate.find(c => c.type !== 'reference') ?? nonDate[0] ?? null;
  }, [schema.columns]);

  // Whether the active date column is date-only (no time component)
  const isDateOnly = useMemo(
    () => schema.columns.find(c => c.name === dateColumn)?.type === 'date',
    [schema.columns, dateColumn],
  );

  // Date-only data: restrict to month + agenda (no hourly time grids)
  const availableViews: View[] = useMemo(
    () => isDateOnly ? ['month', 'agenda'] : ['month', 'week', 'day', 'agenda'],
    [isDateOnly],
  );

  // If the current view is not available (e.g. switched from datetime → date col), reset to month
  const safeView: View = availableViews.includes(currentView) ? currentView : 'month';

  const events = useMemo(() => rows.map(row => {
    const raw = row[dateColumn];
    const date = parseTemporalUnknown(raw);
    if (!date) return null;
    let title = '';
    if (labelColumn) {
      title = labelColumn.type === 'reference'
        ? resolveColumnPath(row, labelColumn.name)
        : (row[labelColumn.name] ?? '');
    }
    // Fall back to formatted date if label is empty
    if (!title) title = format(date, 'd MMM yyyy');
    return {
      id: row[INTERNAL_ROW_ID] ?? Math.random(),
      title,
      start: date,
      end: date,
      allDay: isDateOnly,
      resource: row,
    };
  }).filter(Boolean) as RBCEvent[], [rows, dateColumn, labelColumn, isDateOnly, resolveColumnPath]);

  const hasPicker = dateColumns.length > 1;

  // Build a blank row pre-filled with the given date for the date column
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

  const handleSelectSlot = (slotInfo: SlotInfo) => {
    if (readOnly || !onCreateRow) return;
    setDialog({
      open: true,
      title: 'New Entry',
      initialValues: blankRowForDate(slotInfo.start),
    });
  };

  const handleSelectEvent = (event: RBCEvent) => {
    const row = (event as RBCEvent & { resource: Row }).resource;
    if (!row) return;
    const rowIndex = rows.findIndex(r => r[INTERNAL_ROW_ID] === row[INTERNAL_ROW_ID]);
    if (rowIndex < 0) return;
    // Build initial values from the row (exclude _rowId)
    const initialValues: Row = {};
    for (const col of schema.columns) {
      if (col.name === INTERNAL_ROW_ID) continue;
      initialValues[col.name] = row[col.name] ?? '';
    }
    setDialog({
      open: true,
      title: 'Edit Entry',
      initialValues,
      rowIndex,
    });
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
    padding: '3px 10px',
    fontSize: 13,
    fontWeight: active ? 600 : 400,
    background: active ? 'var(--color-primary)' : 'transparent',
    color: active ? '#fff' : 'var(--color-text)',
    border: '1px solid',
    borderColor: active ? 'var(--color-primary)' : 'var(--color-border)',
    borderRadius: 4,
    cursor: 'pointer',
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'var(--color-surface)' }}>
      {/* Toolbar row */}
      <div style={{ padding: '6px 12px', borderBottom: '1px solid var(--color-border)', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', background: 'var(--color-surface)' }}>
        {hasPicker && (
          <>
            <label style={{ fontWeight: 500, fontSize: 13, color: 'var(--color-text)' }}>Date field:</label>
            <select
              value={dateColumn}
              onChange={e => onDateColumnChange(e.target.value)}
              style={{ fontSize: 13, padding: '2px 6px', borderRadius: 4, border: '1px solid var(--color-border)', marginRight: 8, background: 'var(--color-surface)', color: 'var(--color-text)' }}
            >
              {dateColumns.map(c => (
                <option key={c.name} value={c.name}>{c.displayName ?? c.name}</option>
              ))}
            </select>
          </>
        )}
        <div style={{ display: 'flex', gap: 4, marginLeft: hasPicker ? 0 : 'auto' }}>
          <button style={viewToggleStyle(viewMode === 'grid')} onClick={() => setViewMode('grid')}>Grid</button>
          <button style={viewToggleStyle(viewMode === 'scroll')} onClick={() => setViewMode('scroll')}>Scroll</button>
          <button style={viewToggleStyle(viewMode === 'agenda')} onClick={() => setViewMode('agenda')}>Agenda</button>
        </div>
      </div>

      {viewMode === 'scroll' ? (
        <CalendarScrollView
          events={events.map(e => ({ id: (e as RBCEvent & { id: unknown }).id, title: String(e!.title), start: e!.start as Date, allDay: !!e!.allDay, resource: (e as RBCEvent & { resource: Row }).resource }))}
          onSelectDate={!readOnly && onCreateRow ? (date) => handleSelectSlot({ start: date, end: date, slots: [date], action: 'click' }) : undefined}
          onSelectEvent={(scrollEv) => {
            const row = (scrollEv as { resource?: Row }).resource;
            if (!row) return;
            const fakeEvent = events.find(e => (e as RBCEvent & { id: unknown }).id === scrollEv.id);
            if (fakeEvent) handleSelectEvent(fakeEvent as RBCEvent);
          }}
        />
      ) : viewMode === 'grid' ? (
        <div style={{ flex: 1, minHeight: 0, padding: 8 }}>
          <BigCalendar
            localizer={localizer}
            events={events}
            startAccessor="start"
            endAccessor="end"
            style={{ height: '100%' }}
            date={currentDate}
            view={safeView}
            onNavigate={date => setCurrentDate(date)}
            onView={view => setCurrentView(view)}
            popup
            views={availableViews}
            selectable={!readOnly && !!onCreateRow}
            onSelectSlot={handleSelectSlot}
            onSelectEvent={handleSelectEvent}
          />
        </div>
      ) : (
        <div style={{ flex: 1, minHeight: 0, overflowY: 'auto', padding: '12px 16px 24px', maxWidth: 720, margin: '0 auto', width: '100%', boxSizing: 'border-box', color: 'var(--color-text)' }}>
          <AgendaView
            events={events.map(e => ({ id: (e as RBCEvent & { id: unknown }).id, title: String(e!.title), start: e!.start as Date, allDay: !!e!.allDay, resource: (e as RBCEvent & { resource: Row }).resource }))}
            todayRef={{ current: null }}
            onSelectEvent={(scrollEv) => {
              const row = (scrollEv as { resource?: Row }).resource;
              if (!row) return;
              const fakeEvent = events.find(ev => (ev as RBCEvent & { id: unknown }).id === scrollEv.id);
              if (fakeEvent) handleSelectEvent(fakeEvent as RBCEvent);
            }}
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
          readOnly={readOnly && dialog.rowIndex !== undefined}
        />
      )}
    </div>
  );
};

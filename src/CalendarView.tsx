import React, { useMemo, useState } from 'react';
import { Calendar as BigCalendar, dateFnsLocalizer, type Event as RBCEvent, type View } from 'react-big-calendar';
import { format } from 'date-fns/format';
import { parse } from 'date-fns/parse';
import { startOfWeek } from 'date-fns/startOfWeek';
import { getDay } from 'date-fns/getDay';
import { enUS } from 'date-fns/locale/en-US';
import type { TableSchema, Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseTemporalUnknown } from './dateFormat';
import { CalendarScrollView } from './CalendarScrollView';
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
}

export const CalendarView: React.FC<CalendarViewProps> = ({ schema, rows, dateColumn, onDateColumnChange, resolveColumnPath = DEFAULT_RESOLVE }) => {
  const [currentDate, setCurrentDate] = useState(new Date());
  const [currentView, setCurrentView] = useState<View>('month');
  const [scrollMode, setScrollMode] = useState(false);

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

  const viewToggleStyle = (active: boolean): React.CSSProperties => ({
    padding: '3px 10px',
    fontSize: 13,
    fontWeight: active ? 600 : 400,
    background: active ? '#4f46e5' : 'transparent',
    color: active ? '#fff' : '#374151',
    border: '1px solid',
    borderColor: active ? '#4f46e5' : '#d1d5db',
    borderRadius: 4,
    cursor: 'pointer',
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'white' }}>
      {/* Toolbar row */}
      <div style={{ padding: '6px 12px', borderBottom: '1px solid #e5e7eb', display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        {hasPicker && (
          <>
            <label style={{ fontWeight: 500, fontSize: 13 }}>Date field:</label>
            <select
              value={dateColumn}
              onChange={e => onDateColumnChange(e.target.value)}
              style={{ fontSize: 13, padding: '2px 6px', borderRadius: 4, border: '1px solid #d1d5db', marginRight: 8 }}
            >
              {dateColumns.map(c => (
                <option key={c.name} value={c.name}>{c.displayName ?? c.name}</option>
              ))}
            </select>
          </>
        )}
        <div style={{ display: 'flex', gap: 4, marginLeft: hasPicker ? 0 : 'auto' }}>
          <button style={viewToggleStyle(!scrollMode)} onClick={() => setScrollMode(false)}>
            Grid
          </button>
          <button style={viewToggleStyle(scrollMode)} onClick={() => setScrollMode(true)}>
            Scroll
          </button>
        </div>
      </div>

      {scrollMode ? (
        <CalendarScrollView events={events.map(e => ({ id: (e as RBCEvent & { id: unknown }).id, title: String(e!.title), start: e!.start as Date, allDay: !!e!.allDay }))} />
      ) : (
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
          />
        </div>
      )}
    </div>
  );
};

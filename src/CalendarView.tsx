import React, { useState, useMemo } from 'react';
import type { TableSchema, Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseTemporalUnknown } from './dateFormat';

const DAY_LABELS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

interface CalendarViewProps {
  schema: TableSchema;
  rows: Row[];
  dateColumn: string;
  onDateColumnChange: (col: string) => void;
}

function dateKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

const CalendarView: React.FC<CalendarViewProps> = ({
  schema,
  rows,
  dateColumn,
  onDateColumnChange,
}) => {
  const today = new Date();
  const [viewYear, setViewYear] = useState(today.getFullYear());
  const [viewMonth, setViewMonth] = useState(today.getMonth());

  const dateColumns = useMemo(
    () => schema.columns.filter(c => c.type === 'date' || c.type === 'datetime'),
    [schema.columns],
  );

  // First non-internal, non-date column used as the row label
  const labelColumn = useMemo(
    () => schema.columns.find(c => c.name !== INTERNAL_ROW_ID && c.type !== 'date' && c.type !== 'datetime'),
    [schema.columns],
  );

  const getRowLabel = (row: Row): string => {
    if (labelColumn) return row[labelColumn.name] || '(empty)';
    return row[dateColumn] || '';
  };

  // Build date → rows map
  const eventMap = useMemo(() => {
    const map = new Map<string, Row[]>();
    for (const row of rows) {
      const raw = row[dateColumn];
      if (!raw) continue;
      const d = parseTemporalUnknown(raw);
      if (!d) continue;
      const key = dateKey(d);
      const existing = map.get(key);
      if (existing) existing.push(row);
      else map.set(key, [row]);
    }
    return map;
  }, [rows, dateColumn]);

  const prevMonth = () => {
    if (viewMonth === 0) { setViewYear(y => y - 1); setViewMonth(11); }
    else setViewMonth(m => m - 1);
  };
  const nextMonth = () => {
    if (viewMonth === 11) { setViewYear(y => y + 1); setViewMonth(0); }
    else setViewMonth(m => m + 1);
  };

  const todayKey = dateKey(today);
  const firstDow = new Date(viewYear, viewMonth, 1).getDay();
  const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();

  // Build grid cells: leading blanks + day cells
  const cells = useMemo(() => {
    const result: Array<{ day: number; key: string } | null> = [];
    for (let i = 0; i < firstDow; i++) result.push(null);
    for (let d = 1; d <= daysInMonth; d++) {
      const key = `${viewYear}-${String(viewMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      result.push({ day: d, key });
    }
    return result;
  }, [viewYear, viewMonth, firstDow, daysInMonth]);

  return (
    <div className="calendar-view">
      {/* Navigation header */}
      <div className="calendar-nav">
        <button className="calendar-nav-btn" onClick={prevMonth} aria-label="Previous month">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="15 18 9 12 15 6" />
          </svg>
        </button>
        <h2 className="calendar-month-title">
          {MONTH_NAMES[viewMonth]} {viewYear}
        </h2>
        <button className="calendar-nav-btn" onClick={nextMonth} aria-label="Next month">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="9 18 15 12 9 6" />
          </svg>
        </button>
        {dateColumns.length > 1 && (
          <label className="calendar-col-picker">
            <span className="calendar-col-picker-label">Date field:</span>
            <select
              className="calendar-col-select"
              value={dateColumn}
              onChange={e => onDateColumnChange(e.target.value)}
            >
              {dateColumns.map(c => (
                <option key={c.name} value={c.name}>{c.displayName ?? c.name}</option>
              ))}
            </select>
          </label>
        )}
      </div>

      {/* Day-of-week headers */}
      <div className="calendar-grid">
        {DAY_LABELS.map(d => (
          <div key={d} className="calendar-dow-header">{d}</div>
        ))}

        {/* Day cells */}
        {cells.map((cell, i) => {
          if (!cell) {
            return <div key={`blank-${i}`} className="calendar-cell calendar-cell-blank" />;
          }
          const events = eventMap.get(cell.key) ?? [];
          const isToday = cell.key === todayKey;
          return (
            <div
              key={cell.key}
              className={`calendar-cell${isToday ? ' calendar-cell-today' : ''}`}
            >
              <span className={`calendar-day-num${isToday ? ' is-today' : ''}`}>{cell.day}</span>
              <div className="calendar-event-list">
                {events.slice(0, 3).map((row, ei) => (
                  <div
                    key={row[INTERNAL_ROW_ID] ?? ei}
                    className="calendar-event-chip"
                    title={getRowLabel(row)}
                  >
                    {getRowLabel(row)}
                  </div>
                ))}
                {events.length > 3 && (
                  <div className="calendar-event-overflow">+{events.length - 3} more</div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default CalendarView;

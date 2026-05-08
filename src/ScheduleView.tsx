import React, { useMemo, useRef, useCallback } from 'react';
import type { TableSchema, Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { parseTemporalUnknown } from './dateFormat';

const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_FULL = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const DAY_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

interface ScheduleViewProps {
  schema: TableSchema;
  rows: Row[];
  dateColumn: string;
  onDateColumnChange: (col: string) => void;
}

function monthKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function dateKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

const ScheduleView: React.FC<ScheduleViewProps> = ({
  schema,
  rows,
  dateColumn,
  onDateColumnChange,
}) => {
  const monthHeadingRefs = useRef<Map<string, HTMLElement>>(new Map());
  const scrollContainerRef = useRef<HTMLDivElement>(null);

  const dateColumns = useMemo(
    () => schema.columns.filter(c => c.type === 'date' || c.type === 'datetime'),
    [schema.columns],
  );

  // Columns to display on cards (everything except internal ID and the date column)
  const displayColumns = useMemo(
    () => schema.columns.filter(c => c.name !== INTERNAL_ROW_ID && c.name !== dateColumn),
    [schema.columns, dateColumn],
  );

  // Build sorted date groups
  const dateGroups = useMemo(() => {
    const withDates = rows
      .map(row => ({ row, date: parseTemporalUnknown(row[dateColumn]) }))
      .filter((x): x is { row: Row; date: Date } => !!x.date);

    withDates.sort((a, b) => a.date.getTime() - b.date.getTime());

    const groups = new Map<string, { date: Date; rows: Row[] }>();
    for (const { row, date } of withDates) {
      const key = dateKey(date);
      const g = groups.get(key);
      if (g) g.rows.push(row);
      else groups.set(key, { date, rows: [row] });
    }
    return Array.from(groups.values());
  }, [rows, dateColumn]);

  // Unique months that have data
  const availableMonths = useMemo(() => {
    const seen = new Set<string>();
    const result: { key: string; year: number; month: number }[] = [];
    for (const { date } of dateGroups) {
      const mk = monthKey(date);
      if (!seen.has(mk)) {
        seen.add(mk);
        result.push({ key: mk, year: date.getFullYear(), month: date.getMonth() });
      }
    }
    return result;
  }, [dateGroups]);

  const scrollToMonth = useCallback((mk: string) => {
    const el = monthHeadingRefs.current.get(mk);
    if (!el || !scrollContainerRef.current) return;
    // Offset for the sticky month bar height (~42px)
    const container = scrollContainerRef.current;
    const containerTop = container.getBoundingClientRect().top;
    const elTop = el.getBoundingClientRect().top;
    container.scrollTop += elTop - containerTop - 44;
  }, []);

  const setMonthRef = useCallback((mk: string, el: HTMLElement | null) => {
    if (el) monthHeadingRefs.current.set(mk, el);
    else monthHeadingRefs.current.delete(mk);
  }, []);

  // Track which months have had their heading rendered in this pass
  const seenMonths = new Set<string>();

  return (
    <div className="schedule-view">
      {/* Sticky month jump bar */}
      {availableMonths.length > 1 && (
        <div className="schedule-month-bar">
          {availableMonths.map(({ key: mk, year, month }) => (
            <button
              key={mk}
              className="schedule-month-btn"
              onClick={() => scrollToMonth(mk)}
            >
              {MONTH_SHORT[month]} {year}
            </button>
          ))}
        </div>
      )}

      {/* Scrollable content */}
      <div className="schedule-scroll" ref={scrollContainerRef}>
        {dateGroups.length === 0 ? (
          <div className="schedule-empty">
            <p>No dated items to display.</p>
            <p className="schedule-empty-hint">Add rows with a value in the "{dateColumn}" column to see them here.</p>
          </div>
        ) : (
          <>
            {dateColumns.length > 1 && (
              <label className="schedule-col-picker">
                <span className="schedule-col-picker-label">Date field:</span>
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
            {dateGroups.map(({ date, rows: groupRows }) => {
              const mk = monthKey(date);
              const dk = dateKey(date);
              const isNewMonth = !seenMonths.has(mk);
              if (isNewMonth) seenMonths.add(mk);

              return (
                <React.Fragment key={dk}>
                  {isNewMonth && (
                    <div
                      className="schedule-month-heading"
                      ref={el => setMonthRef(mk, el)}
                    >
                      {MONTH_FULL[date.getMonth()]} {date.getFullYear()}
                    </div>
                  )}
                  <div className="schedule-date-group">
                    <div className="schedule-date-header">
                      {DAY_FULL[date.getDay()]}, {MONTH_FULL[date.getMonth()]} {date.getDate()}
                    </div>
                    <div className="schedule-cards">
                      {groupRows.map((row, i) => (
                        <div key={row[INTERNAL_ROW_ID] ?? i} className="schedule-card">
                          {displayColumns.map(col => {
                            const val = row[col.name];
                            if (!val) return null;
                            return (
                              <div key={col.name} className="schedule-card-field">
                                <span className="schedule-card-label">{col.displayName ?? col.name}</span>
                                <span className="schedule-card-value">{val}</span>
                              </div>
                            );
                          })}
                          {displayColumns.every(col => !row[col.name]) && (
                            <div className="schedule-card-empty">(no details)</div>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                </React.Fragment>
              );
            })}
          </>
        )}
      </div>
    </div>
  );
};

export default ScheduleView;

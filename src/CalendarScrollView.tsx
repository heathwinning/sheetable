import React, { useEffect, useRef, useMemo, useState, useCallback } from 'react';
import { startOfMonth } from 'date-fns/startOfMonth';
import { endOfMonth } from 'date-fns/endOfMonth';
import { startOfWeek } from 'date-fns/startOfWeek';
import { endOfWeek } from 'date-fns/endOfWeek';
import { addDays } from 'date-fns/addDays';
import { addMonths } from 'date-fns/addMonths';
import { isSameMonth } from 'date-fns/isSameMonth';
import { isToday } from 'date-fns/isToday';
import { isBefore } from 'date-fns/isBefore';
import { format } from 'date-fns/format';
import { parseISO } from 'date-fns/parseISO';

export interface ScrollEvent {
  id: unknown;
  title: string;
  start: Date;
  allDay: boolean;
  resource?: unknown;
}

interface CalendarScrollViewProps {
  events: ScrollEvent[];
  monthsBefore?: number;
  monthsAfter?: number;
  onSelectDate?: (date: Date) => void;
  onSelectEvent?: (event: ScrollEvent) => void;
}

const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

function getMonthDays(monthStart: Date): Date[] {
  const gridStart = startOfWeek(monthStart, { weekStartsOn: 0 });
  const gridEnd = endOfWeek(endOfMonth(monthStart), { weekStartsOn: 0 });
  const days: Date[] = [];
  let d = gridStart;
  while (d <= gridEnd) { days.push(d); d = addDays(d, 1); }
  return days;
}

// ---- Month grid sub-component -----------------------------------------------

const MonthGrid: React.FC<{
  monthStart: Date;
  events: ScrollEvent[];
  todayRef: React.RefObject<HTMLDivElement | null>;
  onSelectDate?: (date: Date) => void;
  onSelectEvent?: (event: ScrollEvent) => void;
}> = ({ monthStart, events, todayRef, onSelectDate, onSelectEvent }) => {
  const days = useMemo(() => getMonthDays(monthStart), [monthStart]);

  const eventsByDay = useMemo(() => {
    const map = new Map<string, ScrollEvent[]>();
    for (const ev of events) {
      const key = format(ev.start, 'yyyy-MM-dd');
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(ev);
    }
    return map;
  }, [events]);

  return (
    <div style={{ marginBottom: 32 }}>
      {/* Day-of-week headers */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', marginBottom: 4 }}>
        {DAY_NAMES.map(d => (
          <div key={d} style={{ textAlign: 'center', fontSize: 11, fontWeight: 600, color: 'var(--color-text-muted)', textTransform: 'uppercase', paddingBottom: 4 }}>{d}</div>
        ))}
      </div>
      {/* Day cells */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '2px 0' }}>
        {days.map((day, _i) => {
          const inMonth = isSameMonth(day, monthStart);
          const today = isToday(day);
          const key = format(day, 'yyyy-MM-dd');
          const dayEvents = eventsByDay.get(key) ?? [];
          return (
            <div
              key={key}
              ref={today ? (todayRef as React.RefObject<HTMLDivElement>) : undefined}
              onClick={() => inMonth && onSelectDate?.(day)}
              style={{
                padding: '4px 2px',
                minHeight: 54,
                opacity: inMonth ? 1 : 0.25,
                borderTop: '1px solid var(--color-border)',
                cursor: inMonth && onSelectDate ? 'pointer' : 'default',
                borderRadius: 4,
                transition: 'background 0.1s',
              }}
              onMouseEnter={e => { if (inMonth && onSelectDate) (e.currentTarget as HTMLDivElement).style.background = 'var(--color-surface-2)'; }}
              onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = ''; }}
            >
              <div style={{ display: 'flex', justifyContent: 'center', marginBottom: 3 }}>
                <span style={{
                  display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                  width: 26, height: 26, borderRadius: '50%', fontSize: 13,
                  fontWeight: today ? 700 : 400,
                  color: today ? '#fff' : 'var(--color-text)',
                  background: today ? 'var(--color-primary)' : 'transparent',
                }}>
                  {format(day, 'd')}
                </span>
              </div>
              {dayEvents.slice(0, 3).map((ev, j) => (
                <div
                  key={j}
                  onClick={e => { e.stopPropagation(); onSelectEvent?.(ev); }}
                  style={{
                    background: 'var(--color-cell-selected)', color: 'var(--color-primary)',
                    fontSize: 10, fontWeight: 500, borderRadius: 3, padding: '1px 4px',
                    marginBottom: 2, overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis',
                    cursor: onSelectEvent ? 'pointer' : 'default',
                  }}
                >
                  {ev.title}
                </div>
              ))}
              {dayEvents.length > 3 && (
                <div style={{ fontSize: 10, color: 'var(--color-text-muted)', paddingLeft: 4 }}>+{dayEvents.length - 3} more</div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

// ---- Agenda sub-view --------------------------------------------------------

export const AgendaView: React.FC<{
  events: ScrollEvent[];
  todayRef: React.RefObject<HTMLDivElement | null>;
  onSelectEvent?: (event: ScrollEvent) => void;
}> = ({ events, todayRef, onSelectEvent }) => {
  const sorted = useMemo(() =>
    [...events].sort((a, b) => a.start.getTime() - b.start.getTime()),
    [events]
  );

  // Group by day
  const groups = useMemo(() => {
    const map = new Map<string, ScrollEvent[]>();
    for (const ev of sorted) {
      const key = format(ev.start, 'yyyy-MM-dd');
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(ev);
    }
    return [...map.entries()].map(([key, evs]) => ({ key, date: parseISO(key), evs }));
  }, [sorted]);

  const todayKey = format(new Date(), 'yyyy-MM-dd');

  if (groups.length === 0) {
    return <div style={{ padding: 32, textAlign: 'center', color: 'var(--color-text-muted)', fontSize: 14 }}>No events</div>;
  }

  return (
    <div>
      {groups.map(({ key, date, evs }) => {
        const today = key === todayKey;
        const past = isBefore(date, new Date()) && !today;
        return (
          <div key={key} ref={today ? (todayRef as React.RefObject<HTMLDivElement>) : undefined} style={{ marginBottom: 2 }}>
            <div style={{
              padding: '6px 0 4px',
              borderTop: '1px solid var(--color-border)',
              display: 'flex', alignItems: 'baseline', gap: 8,
              opacity: past ? 0.5 : 1,
            }}>
              <span style={{ fontWeight: 700, fontSize: 15, color: today ? 'var(--color-primary)' : 'var(--color-text)', minWidth: 28 }}>
                {format(date, 'd')}
              </span>
              <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-text-muted)', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                {format(date, 'EEE')}
              </span>
              <span style={{ fontSize: 12, color: 'var(--color-text-muted)' }}>
                {format(date, 'MMMM yyyy')}
              </span>
              {today && <span style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-primary)', background: 'var(--color-cell-selected)', borderRadius: 4, padding: '1px 6px' }}>Today</span>}
            </div>
            {evs.map((ev, j) => (
              <div
                key={j}
                onClick={() => onSelectEvent?.(ev)}
                style={{
                  display: 'flex', alignItems: 'center', gap: 10,
                  padding: '8px 10px', marginBottom: 3, borderRadius: 6,
                  background: 'var(--color-surface-2)',
                  cursor: onSelectEvent ? 'pointer' : 'default',
                  border: '1px solid var(--color-border)',
                  opacity: past ? 0.6 : 1,
                }}
                onMouseEnter={e => { if (onSelectEvent) (e.currentTarget as HTMLDivElement).style.background = 'var(--color-cell-selected)'; }}
                onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'var(--color-surface-2)'; }}
              >
                <div style={{ width: 4, height: 32, borderRadius: 2, background: 'var(--color-primary)', flexShrink: 0 }} />
                <span style={{ fontSize: 13, fontWeight: 500, color: 'var(--color-text)' }}>{ev.title}</span>
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
};

// ---- Main CalendarScrollView ------------------------------------------------

export const CalendarScrollView: React.FC<CalendarScrollViewProps> = ({
  events,
  monthsBefore = 36,
  monthsAfter = 60,
  onSelectDate,
  onSelectEvent,
}) => {
  const todayRef = useRef<HTMLDivElement | null>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const monthRefsMap = useRef<Map<string, HTMLDivElement>>(new Map());

  const thisMonth = useMemo(() => startOfMonth(new Date()), []);
  const [activeMonthKey, setActiveMonthKey] = useState(() => format(thisMonth, 'yyyy-MM'));

  const months = useMemo(() => {
    const result: Date[] = [];
    for (let i = -monthsBefore; i <= monthsAfter; i++) result.push(addMonths(thisMonth, i));
    return result;
  }, [thisMonth, monthsBefore, monthsAfter]);

  // IntersectionObserver — track which month is at the top of the viewport
  useEffect(() => {
    const container = scrollRef.current;
    if (!container) return;
    const observer = new IntersectionObserver(
      entries => {
        // Find topmost intersecting month
        const visible = entries
          .filter(e => e.isIntersecting)
          .sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
        if (visible.length > 0) {
          const key = (visible[0].target as HTMLDivElement).dataset.monthKey;
          if (key) {
            setActiveMonthKey(key);
          }
        }
      },
      { root: container, rootMargin: '-10% 0px -80% 0px', threshold: 0 },
    );
    monthRefsMap.current.forEach(el => observer.observe(el));
    return () => observer.disconnect();
  }, [months]);

  // Scroll today into view on mount
  useEffect(() => {
    if (todayRef.current && scrollRef.current) {
      const cellTop = todayRef.current.offsetTop;
      scrollRef.current.scrollTop = Math.max(0, cellTop - 160);
    }
  }, []);

const scrollToMonth = useCallback((key: string) => {
    const el = monthRefsMap.current.get(key);
    if (el && scrollRef.current) {
      scrollRef.current.scrollTop = Math.max(0, el.offsetTop - 80);
    }
  }, []);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, background: 'var(--color-surface)' }}>

      {/* ── Sticky nav ── */}
      <div style={{
        flexShrink: 0,
        borderBottom: '1px solid var(--color-border)',
        background: 'var(--color-surface)',
        padding: '6px 12px',
        display: 'flex',
        alignItems: 'center',
        gap: 8,
      }}>
        <select
          value={activeMonthKey.slice(0, 4)}
          onChange={e => {
            const year = e.target.value;
            const key = `${year}-01`;
            setActiveMonthKey(key);
            scrollToMonth(key);
          }}
          style={{
            fontSize: 13,
            padding: '3px 8px',
            borderRadius: 6,
            border: '1px solid var(--color-border)',
            background: 'var(--color-surface)',
            color: 'var(--color-text)',
            cursor: 'pointer',
          }}
        >
          {[...new Set(months.map(m => format(m, 'yyyy')))].map(y => (
            <option key={y} value={y}>{y}</option>
          ))}
        </select>
        <button
          onClick={() => {
            const key = format(thisMonth, 'yyyy-MM');
            setActiveMonthKey(key);
            scrollToMonth(key);
          }}
          style={{
            fontSize: 12,
            padding: '3px 10px',
            borderRadius: 6,
            border: '1px solid var(--color-border)',
            background: 'transparent',
            color: 'var(--color-text)',
            cursor: 'pointer',
          }}
        >Today</button>
      </div>

      {/* ── Scrollable content ── */}
      <div
        ref={scrollRef}
        style={{
          flex: 1,
          overflowY: 'auto',
          padding: '12px 16px 24px',
          maxWidth: 720,
          margin: '0 auto',
          width: '100%',
          boxSizing: 'border-box',
          color: 'var(--color-text)',
        }}
      >
        {months.map((m) => {
            const key = format(m, 'yyyy-MM');
            return (
              <div
                key={key}
                data-month-key={key}
                ref={el => {
                  if (el) monthRefsMap.current.set(key, el);
                  else monthRefsMap.current.delete(key);
                }}
              >
                {/* Month label */}
                <div style={{
                  fontSize: 15, fontWeight: 700, color: 'var(--color-text)',
                  margin: '8px 0 6px', paddingLeft: 2, letterSpacing: '-0.01em',
                }}>
                  {format(m, 'MMMM yyyy')}
                </div>
                <MonthGrid
                  monthStart={m}
                  events={events}
                  todayRef={todayRef}
                  onSelectDate={onSelectDate}
                  onSelectEvent={onSelectEvent}
                />
              </div>
            );
          })}
      </div>
    </div>
  );
};

// ---- Styles -----------------------------------------------------------------





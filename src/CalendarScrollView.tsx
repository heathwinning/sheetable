import React, { useEffect, useRef, useMemo, useState, useCallback } from 'react';
import { startOfMonth } from 'date-fns/startOfMonth';
import { endOfMonth } from 'date-fns/endOfMonth';
import { startOfWeek } from 'date-fns/startOfWeek';
import { endOfWeek } from 'date-fns/endOfWeek';
import { addDays } from 'date-fns/addDays';
import { addMonths } from 'date-fns/addMonths';
import { startOfYear } from 'date-fns/startOfYear';
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
  year: number;
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
    <div className="calendar-month-grid">
      {/* Day-of-week headers */}
      <div className="calendar-grid-headers">
        {DAY_NAMES.map(d => (
          <div key={d} className="calendar-grid-header">{d}</div>
        ))}
      </div>
      {/* Day cells */}
      <div className="calendar-grid-days">
        {days.map((day, _i) => {
          const inMonth = isSameMonth(day, monthStart);
          const today = isToday(day);
          const key = format(day, 'yyyy-MM-dd');
          const dayEvents = inMonth ? (eventsByDay.get(key) ?? []) : [];
          return (
            <div
              key={key}
              ref={today ? (todayRef as React.RefObject<HTMLDivElement>) : undefined}
              onClick={() => inMonth && onSelectDate?.(day)}
              className={`calendar-day-cell ${inMonth ? 'in-month' : 'not-in-month'}`}
            >
              {inMonth && (
                <>
                  <div className="calendar-day-number-container">
                    <span className={`calendar-day-number ${today ? 'today' : ''}`}>
                      {format(day, 'd')}
                    </span>
                  </div>
                  {dayEvents.slice(0, 3).map((ev, j) => (
                    <div
                      key={j}
                      onClick={e => { e.stopPropagation(); onSelectEvent?.(ev); }}
                      className={`calendar-event-item ${onSelectEvent ? '' : 'not-clickable'}`}
                    >
                      {ev.title}
                    </div>
                  ))}
                  {dayEvents.length > 3 && (
                    <div className="calendar-event-overflow">+{dayEvents.length - 3} more</div>
                  )}
                </>
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
  year: number;
  selectedDate?: Date | null;
  todayRef: React.RefObject<HTMLDivElement | null>;
  onSelectEvent?: (event: ScrollEvent) => void;
}> = ({ events, year, selectedDate = null, todayRef, onSelectEvent }) => {
  const selectedDateKey = selectedDate ? format(selectedDate, 'yyyy-MM-dd') : null;
  const sorted = useMemo(() =>
    [...events]
      .filter((ev) => ev.start.getFullYear() === year)
      .filter((ev) => !selectedDateKey || format(ev.start, 'yyyy-MM-dd') === selectedDateKey)
      .sort((a, b) => a.start.getTime() - b.start.getTime()),
    [events, year, selectedDateKey]
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
    return <div className="agenda-empty">No events</div>;
  }

  return (
    <div>
      {groups.map(({ key, date, evs }) => {
        const today = key === todayKey;
        const past = isBefore(date, new Date()) && !today;
        return (
          <div key={key} ref={today ? (todayRef as React.RefObject<HTMLDivElement>) : undefined} className={`agenda-day-group`}>
            <div className={`agenda-day-header ${past ? 'past' : ''}`}>
              <span className={`agenda-day-number ${today ? 'today' : ''}`}>
                {format(date, 'd')}
              </span>
              <span className="agenda-day-weekday">
                {format(date, 'EEE')}
              </span>
              <span className="agenda-day-date">
                {format(date, 'MMMM yyyy')}
              </span>
              {today && <span className="agenda-today-badge">Today</span>}
            </div>
            {evs.map((ev, j) => (
              <div
                key={j}
                onClick={() => onSelectEvent?.(ev)}
                className={`agenda-event-item ${past ? 'past' : ''} ${onSelectEvent ? '' : 'not-clickable'}`}
              >
                <div className="agenda-event-accent" />
                <span className="agenda-event-title">{ev.title}</span>
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
  year,
  onSelectDate,
  onSelectEvent,
}) => {
  const todayRef = useRef<HTMLDivElement | null>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);
  const monthRefsMap = useRef<Map<string, HTMLDivElement>>(new Map());

  const thisMonth = useMemo(() => startOfMonth(new Date()), []);
  const initialMonth = useMemo(
    () => (thisMonth.getFullYear() === year ? thisMonth : new Date(year, 0, 1)),
    [thisMonth, year],
  );
  const [_activeMonthKey, setActiveMonthKey] = useState(() => format(initialMonth, 'yyyy-MM'));
  void _activeMonthKey;

  const months = useMemo(() => {
    const result: Date[] = [];
    const yearStart = startOfYear(new Date(year, 0, 1));
    for (let i = 0; i < 12; i++) result.push(addMonths(yearStart, i));
    return result;
  }, [year]);

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

  // Scroll today/month into view on mount — position ~1/3 from top (≈ 3 weeks of context above)
  useEffect(() => {
    if (todayRef.current && scrollRef.current) {
      const cellTop = todayRef.current.offsetTop;
      const viewHeight = scrollRef.current.clientHeight;
      scrollRef.current.scrollTop = Math.max(0, cellTop - viewHeight / 3);
    }
  }, []);

  useEffect(() => {
    const next = format(initialMonth, 'yyyy-MM');
    setActiveMonthKey(next);
    const el = monthRefsMap.current.get(next);
    if (el && scrollRef.current) {
      scrollRef.current.scrollTop = Math.max(0, el.offsetTop - 12);
    }
  }, [initialMonth]);

const _scrollToMonth = useCallback((key: string) => {
    const el = monthRefsMap.current.get(key);
    if (el && scrollRef.current) {
      scrollRef.current.scrollTop = Math.max(0, el.offsetTop - 80);
    }
  }, []);
  void _scrollToMonth;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, background: 'var(--color-surface)' }}>
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
                <div className="calendar-month-label">
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





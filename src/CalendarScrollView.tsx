import React, { useEffect, useRef, useMemo } from 'react';
import { startOfMonth } from 'date-fns/startOfMonth';
import { endOfMonth } from 'date-fns/endOfMonth';
import { startOfWeek } from 'date-fns/startOfWeek';
import { endOfWeek } from 'date-fns/endOfWeek';
import { addDays } from 'date-fns/addDays';
import { addMonths } from 'date-fns/addMonths';
import { isSameMonth } from 'date-fns/isSameMonth';
import { isSameDay } from 'date-fns/isSameDay';
import { isToday } from 'date-fns/isToday';
import { format } from 'date-fns/format';

interface ScrollEvent {
  id: unknown;
  title: string;
  start: Date;
  allDay: boolean;
}

interface CalendarScrollViewProps {
  events: ScrollEvent[];
  // How many months before/after today to render
  monthsBefore?: number;
  monthsAfter?: number;
}

const DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

function getMonthDays(monthStart: Date): (Date | null)[] {
  // Grid from start-of-week of the first day through end-of-week of the last day
  const gridStart = startOfWeek(monthStart, { weekStartsOn: 0 });
  const gridEnd = endOfWeek(endOfMonth(monthStart), { weekStartsOn: 0 });
  const days: (Date | null)[] = [];
  let d = gridStart;
  while (d <= gridEnd) {
    days.push(d);
    d = addDays(d, 1);
  }
  return days;
}

const MonthGrid: React.FC<{
  monthStart: Date;
  events: ScrollEvent[];
  todayRef: React.RefObject<HTMLDivElement | null>;
}> = ({ monthStart, events, todayRef }) => {
  const days = useMemo(() => getMonthDays(monthStart), [monthStart]);

  // Map events to day keys
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
      <div style={{
        fontSize: 16,
        fontWeight: 700,
        color: '#111827',
        marginBottom: 8,
        paddingLeft: 2,
        letterSpacing: '-0.01em',
      }}>
        {format(monthStart, 'MMMM yyyy')}
      </div>
      {/* Day-of-week headers */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', marginBottom: 4 }}>
        {DAY_NAMES.map(d => (
          <div key={d} style={{
            textAlign: 'center',
            fontSize: 11,
            fontWeight: 600,
            color: '#9ca3af',
            textTransform: 'uppercase',
            paddingBottom: 4,
          }}>{d}</div>
        ))}
      </div>
      {/* Day cells */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', gap: '2px 0' }}>
        {days.map((day, i) => {
          if (!day) return <div key={i} />;
          const inMonth = isSameMonth(day, monthStart);
          const today = isToday(day);
          const key = format(day, 'yyyy-MM-dd');
          const dayEvents = eventsByDay.get(key) ?? [];
          return (
            <div
              key={key}
              ref={today ? (todayRef as React.RefObject<HTMLDivElement>) : undefined}
              style={{
                padding: '4px 2px',
                minHeight: 52,
                opacity: inMonth ? 1 : 0.3,
                borderTop: '1px solid #f3f4f6',
              }}
            >
              <div style={{
                display: 'flex',
                justifyContent: 'center',
                marginBottom: 3,
              }}>
                <span style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  width: 26,
                  height: 26,
                  borderRadius: '50%',
                  fontSize: 13,
                  fontWeight: today ? 700 : 400,
                  color: today ? '#fff' : '#374151',
                  background: today ? '#4f46e5' : 'transparent',
                }}>
                  {format(day, 'd')}
                </span>
              </div>
              {dayEvents.slice(0, 3).map((ev, j) => (
                <div key={j} style={{
                  background: '#e0e7ff',
                  color: '#3730a3',
                  fontSize: 10,
                  fontWeight: 500,
                  borderRadius: 3,
                  padding: '1px 4px',
                  marginBottom: 2,
                  overflow: 'hidden',
                  whiteSpace: 'nowrap',
                  textOverflow: 'ellipsis',
                }}>
                  {ev.title}
                </div>
              ))}
              {dayEvents.length > 3 && (
                <div style={{ fontSize: 10, color: '#6b7280', paddingLeft: 4 }}>
                  +{dayEvents.length - 3} more
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

export const CalendarScrollView: React.FC<CalendarScrollViewProps> = ({
  events,
  monthsBefore = 12,
  monthsAfter = 24,
}) => {
  const todayRef = useRef<HTMLDivElement | null>(null);
  const scrollRef = useRef<HTMLDivElement | null>(null);

  const today = useMemo(() => startOfMonth(new Date()), []);

  const months = useMemo(() => {
    const result: Date[] = [];
    for (let i = -monthsBefore; i <= monthsAfter; i++) {
      result.push(addMonths(today, i));
    }
    return result;
  }, [today, monthsBefore, monthsAfter]);

  // Scroll today into view on mount
  useEffect(() => {
    if (todayRef.current && scrollRef.current) {
      // Offset so the current month is near the top, not the exact cell
      const cellTop = todayRef.current.offsetTop;
      scrollRef.current.scrollTop = Math.max(0, cellTop - 120);
    }
  }, []);

  return (
    <div
      ref={scrollRef}
      style={{
        flex: 1,
        overflowY: 'auto',
        padding: '16px 20px',
        maxWidth: 700,
        margin: '0 auto',
        width: '100%',
        boxSizing: 'border-box',
      }}
    >
      {months.map((m, i) => (
        <MonthGrid
          key={i}
          monthStart={m}
          events={events}
          todayRef={todayRef}
        />
      ))}
    </div>
  );
};

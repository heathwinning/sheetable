import { useState, useRef, useEffect, useCallback } from 'react';
import { createPortal } from 'react-dom';
import type { CustomCellEditorProps } from 'ag-grid-react';
import DatePicker from 'react-datepicker';
import type { ReactDatePickerCustomHeaderProps } from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';
import { log } from './DebugLogger';
import { parseTemporalUnknown, formatDateCanonical } from './dateFormat';

function parseDate(value: string): Date | null {
  return parseTemporalUnknown(value);
}

function formatDate(date: Date | null): string {
  if (!date) return '';
  return formatDateCanonical(date);
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function range(from: number, to: number): number[] {
  return Array.from({ length: to - from + 1 }, (_, i) => from + i);
}

// ── Custom Popover Select ──────────────────────────────────────────────────

function PopoverSelect({
  value,
  options,
  onChange,
  className,
}: {
  value: string;
  options: { label: string; value: string }[];
  onChange: (val: string) => void;
  className?: string;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [open]);

  const activeLabel = options.find(o => o.value === value)?.label ?? value;

  return (
    <div ref={ref} style={{ position: 'relative', display: 'inline-flex' }}>
      <button
        type="button"
        className={`date-popover-trigger ${className ?? ''}`}
        onClick={() => setOpen(o => !o)}
      >
        {activeLabel}
      </button>
      {open && (
        <div className="date-popover-menu">
          {options.map(opt => (
            <button
              key={opt.value}
              type="button"
              className={`date-popover-option${opt.value === value ? ' date-popover-option--active' : ''}`}
              onClick={() => { onChange(opt.value); setOpen(false); }}
            >
              {opt.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Custom Header ──────────────────────────────────────────────────────────

function CustomHeader({
  date,
  changeYear,
  changeMonth,
  decreaseMonth,
  increaseMonth,
  prevMonthButtonDisabled,
  nextMonthButtonDisabled,
}: ReactDatePickerCustomHeaderProps) {
  const yearOptions = range(1900, 2100).map(y => ({ label: String(y), value: String(y) }));
  const monthOptions = MONTHS.map((m, i) => ({ label: m, value: String(i) }));

  return (
    <div className="date-custom-header">
      <button
        type="button"
        aria-label="Previous Month"
        className="date-custom-header__nav"
        disabled={prevMonthButtonDisabled}
        onClick={decreaseMonth}
      >
        <span className="date-custom-header__nav-icon date-custom-header__nav-icon--prev" />
      </button>

      <div className="date-custom-header__selects">
        <PopoverSelect
          value={String(date.getMonth())}
          options={monthOptions}
          onChange={v => changeMonth(Number(v))}
          className="date-popover-month"
        />
        <PopoverSelect
          value={String(date.getFullYear())}
          options={yearOptions}
          onChange={v => changeYear(Number(v))}
          className="date-popover-year"
        />
      </div>

      <button
        type="button"
        aria-label="Next Month"
        className="date-custom-header__nav"
        disabled={nextMonthButtonDisabled}
        onClick={increaseMonth}
      >
        <span className="date-custom-header__nav-icon date-custom-header__nav-icon--next" />
      </button>
    </div>
  );
}

// ── DateCellEditor (inline text + portal DatePicker) ──────────────────────

export default function DateCellEditor({ value, onValueChange, stopEditing }: CustomCellEditorProps) {
  const [selectedDate, setSelectedDate] = useState<Date | null>(parseDate(value ?? ''));
  const [text, setText] = useState(() => {
    const d = parseDate(value ?? '');
    return d ? formatDate(d) : (value ?? '');
  });
  const inputRef = useRef<HTMLInputElement>(null);
  const popoverRef = useRef<HTMLDivElement>(null);

  // Portal position (above / below the cell)
  const [popoverStyle, setPopoverStyle] = useState<React.CSSProperties>({});
  const [popoverAbove, setPopoverAbove] = useState(false);

  const updatePopoverPos = useCallback(() => {
    const el = inputRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const vv = window.visualViewport;
    const viewH = vv ? vv.height : window.innerHeight;
    // The inline date picker is ~300px tall
    const pickerH = 310;
    const belowRoom = viewH - rect.bottom;
    const showAbove = belowRoom < pickerH + 8;

    setPopoverAbove(showAbove);
    setPopoverStyle({
      position: 'fixed',
      left: Math.min(rect.left, window.innerWidth - 320),
      ...(showAbove
        ? { bottom: window.innerHeight - rect.top }
        : { top: rect.bottom }),
      zIndex: 9999,
    });
  }, []);

  useEffect(() => {
    inputRef.current?.focus();
    updatePopoverPos();
    const onScroll = () => updatePopoverPos();
    const vp = document.querySelector('.ag-body-viewport');
    vp?.addEventListener('scroll', onScroll);
    window.addEventListener('scroll', onScroll, true);
    window.addEventListener('resize', onScroll);
    return () => {
      vp?.removeEventListener('scroll', onScroll);
      window.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onScroll);
    };
  }, [updatePopoverPos]);

  const commitText = useCallback(() => {
    const d = parseDate(text);
    if (d) {
      const formatted = formatDate(d);
      log('DateCellEditor typed:', formatted);
      onValueChange(formatted);
      setSelectedDate(d);
      setText(formatted);
      // Don't stop editing — let the user keep typing or pick from calendar
    }
  }, [text, onValueChange]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitText();
      stopEditing();
    } else if (e.key === 'Escape') {
      stopEditing();
    }
  };

  const handlePick = (date: Date | null) => {
    setSelectedDate(date);
    const formatted = formatDate(date);
    log('DateCellEditor picked:', formatted);
    onValueChange(formatted);
    setText(formatted);
    setTimeout(() => stopEditing(), 0);
  };

  const popover = (
    <div
      ref={popoverRef}
      className={`date-cell-popover${popoverAbove ? ' date-cell-popover-above' : ''}`}
      style={popoverStyle}
      onMouseDown={(e) => { e.preventDefault(); }}
    >
      <DatePicker
        selected={selectedDate}
        onChange={handlePick}
        dateFormat="yyyy/MM/dd"
        inline
        renderCustomHeader={CustomHeader}
      />
    </div>
  );

  return (
    <div style={{ width: '100%', height: '100%', display: 'flex', alignItems: 'center' }}>
      <input
        ref={inputRef}
        type="text"
        value={text}
        onChange={e => { setText(e.target.value); setSelectedDate(parseDate(e.target.value)); }}
        onBlur={() => {
          commitText();
          // Delay stopEditing so clicks on the date picker can register first
          setTimeout(() => stopEditing(), 150);
        }}
        onKeyDown={handleKeyDown}
        placeholder="yyyy/MM/dd"
        style={{
          width: '100%',
          height: '100%',
          border: 'none',
          outline: 'none',
          background: 'transparent',
          fontSize: 'inherit',
          fontFamily: 'inherit',
          padding: '0 4px',
          boxSizing: 'border-box',
        }}
      />
      {createPortal(popover, document.body)}
    </div>
  );
}

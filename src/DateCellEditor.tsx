import { useState, useRef, useEffect } from 'react';
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

// ── DateCellEditor ─────────────────────────────────────────────────────────

export default function DateCellEditor({ value, onValueChange, stopEditing }: CustomCellEditorProps) {
  const [selectedDate, setSelectedDate] = useState<Date | null>(parseDate(value ?? ''));

  return (
    <div className="date-cell-editor ag-custom-component-popup">
      <DatePicker
        selected={selectedDate}
        onChange={(date: Date | null) => {
          setSelectedDate(date);
          const formatted = formatDate(date);
          log('DateCellEditor selected:', formatted);
          onValueChange(formatted);
          setTimeout(() => stopEditing(), 0);
        }}
        dateFormat="yyyy/MM/dd"
        inline
        renderCustomHeader={CustomHeader}
      />
    </div>
  );
}

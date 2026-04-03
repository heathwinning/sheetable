import { useState } from 'react';
import type { CustomCellEditorProps } from 'ag-grid-react';
import DatePicker from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';
import { log } from './DebugLogger';

function parseDate(value: string): Date | null {
  if (!value) return null;
  const d = new Date(value + 'T00:00:00');
  return isNaN(d.getTime()) ? null : d;
}

function formatDate(date: Date | null): string {
  if (!date) return '';
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

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
          stopEditing();
        }}
        dateFormat="yyyy-MM-dd"
        inline
        showMonthDropdown
        showYearDropdown
        dropdownMode="select"
      />
    </div>
  );
}

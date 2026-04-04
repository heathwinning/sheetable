import { useState } from 'react';
import type { CustomCellEditorProps } from 'ag-grid-react';
import DatePicker from 'react-datepicker';
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
        dateFormat="yyyy/MM/dd"
        inline
        showMonthDropdown
        showYearDropdown
        dropdownMode="select"
      />
    </div>
  );
}

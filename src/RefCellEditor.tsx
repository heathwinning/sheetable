import React, { useState, useRef, useEffect, useCallback } from 'react';
import type { Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { log } from './DebugLogger';

interface RefCellEditorProps {
  value: string;
  onValueChange: (value: string | null) => void;
  stopEditing: () => void;
  refRows: Row[];
  searchColumns: string[];
  displayColumns: string[];
}

export default function RefCellEditor(props: RefCellEditorProps) {
  const { refRows, searchColumns, displayColumns, onValueChange, stopEditing } = props;
  const [search, setSearch] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(-1);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const getDisplayText = useCallback((row: Row): string => {
    const cols = displayColumns.length > 0 ? displayColumns : searchColumns;
    return cols.map(c => row[c] ?? '').filter(Boolean).join(' · ');
  }, [displayColumns, searchColumns]);

  const getSearchText = useCallback((row: Row): string => {
    const cols = searchColumns.length > 0 ? searchColumns : displayColumns;
    return cols.map(c => (row[c] ?? '').toLowerCase()).join(' ');
  }, [searchColumns, displayColumns]);

  const filtered = search
    ? refRows.filter(row => getSearchText(row).includes(search.toLowerCase()))
    : refRows;

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  useEffect(() => {
    if (selectedIndex >= 0 && listRef.current) {
      const item = listRef.current.children[selectedIndex + 1] as HTMLElement;
      item?.scrollIntoView({ block: 'nearest' });
    }
  }, [selectedIndex]);

  const selectValue = useCallback((rowId: string) => {
    log('RefCellEditor selectValue:', rowId);
    onValueChange(rowId || null);
    stopEditing();
  }, [onValueChange, stopEditing]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    const totalItems = filtered.length + 1;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex(prev => Math.min(prev + 1, totalItems - 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex(prev => Math.max(prev - 1, -1));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (selectedIndex === 0) {
        selectValue('');
      } else if (selectedIndex > 0 && selectedIndex <= filtered.length) {
        selectValue(filtered[selectedIndex - 1][INTERNAL_ROW_ID]);
      }
    } else if (e.key === 'Escape') {
      stopEditing();
    }
  };

  return (
    <div className="ref-editor" onKeyDown={handleKeyDown} onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}>
      <input
        ref={inputRef}
        type="text"
        value={search}
        onChange={e => { setSearch(e.target.value); setSelectedIndex(-1); }}
        placeholder="Search…"
        className="ref-editor-search"
      />
      <div className="ref-editor-list" ref={listRef}>
        <div
          className={`ref-editor-option ref-editor-clear ${selectedIndex === 0 ? 'selected' : ''}`}
          onClick={() => selectValue('')}
        >
          <em>Clear</em>
        </div>
        {filtered.map((row, i) => (
          <div
            key={row[INTERNAL_ROW_ID]}
            className={`ref-editor-option ${selectedIndex === i + 1 ? 'selected' : ''} ${row[INTERNAL_ROW_ID] === props.value ? 'current' : ''}`}
            onClick={() => selectValue(row[INTERNAL_ROW_ID])}
          >
            {getDisplayText(row) || <span style={{ opacity: 0.5 }}>Row {row[INTERNAL_ROW_ID]}</span>}
          </div>
        ))}
        {filtered.length === 0 && (
          <div className="ref-editor-option" style={{ opacity: 0.5, cursor: 'default' }}>No matches</div>
        )}
      </div>
    </div>
  );
}

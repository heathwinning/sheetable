import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import type { Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import { log } from './DebugLogger';

interface RefCellEditorProps {
  value: string;
  onValueChange: (value: string | null) => void;
  stopEditing: () => void;
  refRows: Row[];
  refTable: string;
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  searchColumns: string[];
  displayColumns: string[];
  onCreateRecord?: (refTable: string, seedText: string) => Promise<string | null>;
}

export default function RefCellEditor(props: RefCellEditorProps) {
  const { refRows, refTable, resolveColumnPath, searchColumns, displayColumns, onValueChange, stopEditing, onCreateRecord } = props;
  const [search, setSearch] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(-1);
  const [isCreating, setIsCreating] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  // Pre-resolve display and search text for all rows (handles nested references)
  const resolvedRows = useMemo(() => {
    const dCols = displayColumns.length > 0 ? displayColumns : searchColumns;
    const sCols = searchColumns.length > 0 ? searchColumns : displayColumns;
    return refRows.map(row => ({
      row,
      displayText: dCols.map(c => resolveColumnPath(refTable, row, c)).filter(Boolean).join(' · '),
      searchText: sCols.map(c => resolveColumnPath(refTable, row, c).toLowerCase()).join(' '),
    }));
  }, [refRows, refTable, resolveColumnPath, displayColumns, searchColumns]);

  const filtered = useMemo(() =>
    search
      ? resolvedRows.filter(r => r.searchText.includes(search.toLowerCase()))
      : resolvedRows,
    [resolvedRows, search]
  );

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  useEffect(() => {
    if (selectedIndex >= 0 && listRef.current) {
      const item = listRef.current.querySelector(`[data-select-index="${selectedIndex}"]`) as HTMLElement | null;
      item?.scrollIntoView({ block: 'nearest' });
    }
  }, [selectedIndex]);

  const selectValue = useCallback((rowId: string) => {
    log('RefCellEditor selectValue:', rowId);
    onValueChange(rowId || null);
    setTimeout(() => stopEditing(), 0);
  }, [onValueChange, stopEditing]);

  const showCreateOption = !!onCreateRecord && search.trim().length > 0;
  const createOptionIndex = filtered.length + 1;

  const handleCreateRecord = useCallback(async () => {
    if (!onCreateRecord || isCreating) return;
    setIsCreating(true);
    try {
      const createdRowId = await onCreateRecord(refTable, search.trim());
      if (createdRowId) {
        selectValue(createdRowId);
      }
    } finally {
      setIsCreating(false);
    }
  }, [onCreateRecord, isCreating, refTable, search, selectValue]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    const totalItems = filtered.length + 1 + (showCreateOption ? 1 : 0);
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
        selectValue(filtered[selectedIndex - 1].row[INTERNAL_ROW_ID]);
      } else if (showCreateOption && selectedIndex === createOptionIndex) {
        void handleCreateRecord();
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
          data-select-index={0}
          onClick={() => selectValue('')}
        >
          <em>Clear</em>
        </div>
        {filtered.map((r, i) => (
          <div
            key={r.row[INTERNAL_ROW_ID]}
            className={`ref-editor-option ${selectedIndex === i + 1 ? 'selected' : ''} ${r.row[INTERNAL_ROW_ID] === props.value ? 'current' : ''}`}
            data-select-index={i + 1}
            onClick={() => selectValue(r.row[INTERNAL_ROW_ID])}
          >
            {r.displayText || <span style={{ opacity: 0.5 }}>Row {r.row[INTERNAL_ROW_ID]}</span>}
          </div>
        ))}
        {showCreateOption && (
          <div
            className={`ref-editor-option ref-editor-create ${selectedIndex === createOptionIndex ? 'selected' : ''}`}
            data-select-index={createOptionIndex}
            onClick={() => { void handleCreateRecord(); }}
          >
            {isCreating ? 'Creating…' : `+ Add new record in ${refTable}`}
          </div>
        )}
        {filtered.length === 0 && (
          <div className="ref-editor-option" style={{ opacity: 0.5, cursor: 'default' }}>No matches</div>
        )}
      </div>
    </div>
  );
}

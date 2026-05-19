import React, { useState, useRef, useEffect } from 'react';

interface ListTagsEditorProps {
  value: string;
  onValueChange: (value: string) => void;
  stopEditing: () => void;
}

function parseItems(value: string): string[] {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) return parsed.map(String);
  } catch {
    // fall through
  }
  return value ? [value] : [];
}

export default function ListTagsEditor({ value, onValueChange, stopEditing }: ListTagsEditorProps) {
  const [items, setItems] = useState<string[]>(() => parseItems(value));
  const [inputVal, setInputVal] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  // Sync items → parent value
  useEffect(() => {
    onValueChange(items.length > 0 ? JSON.stringify(items) : '');
  }, [items]); // eslint-disable-line react-hooks/exhaustive-deps

  const addItem = () => {
    const v = inputVal.trim();
    if (!v) return;
    setItems(prev => [...prev, v]);
    setInputVal('');
  };

  const removeItem = (i: number) => {
    setItems(prev => prev.filter((_, j) => j !== i));
    inputRef.current?.focus();
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      addItem();
    } else if (e.key === 'Escape') {
      stopEditing();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      addItem();
      stopEditing();
    } else if (e.key === 'Backspace' && !inputVal) {
      setItems(prev => prev.slice(0, -1));
    }
  };

  const handleBlur = (e: React.FocusEvent) => {
    if (containerRef.current?.contains(e.relatedTarget as Node)) return;
    addItem();
    stopEditing();
  };

  return (
    <div
      ref={containerRef}
      style={{
        background: 'var(--ag-background-color, #fff)',
        border: '2px solid var(--ag-input-focus-border-color, #2196f3)',
        borderRadius: 4,
        padding: '3px 6px',
        display: 'flex',
        flexWrap: 'wrap',
        gap: 4,
        alignItems: 'center',
        minWidth: 220,
        maxWidth: 400,
        boxSizing: 'border-box',
      }}
      onBlur={handleBlur}
    >
      {items.map((item, i) => (
        <span
          key={i}
          style={{
            background: 'var(--ag-row-hover-color, #e8f0fe)',
            color: 'var(--ag-foreground-color, #333)',
            borderRadius: 3,
            padding: '1px 4px 1px 7px',
            fontSize: 12,
            display: 'inline-flex',
            alignItems: 'center',
            gap: 2,
            lineHeight: '18px',
          }}
        >
          {item}
          <button
            tabIndex={-1}
            onMouseDown={(e) => { e.preventDefault(); removeItem(i); }}
            style={{
              background: 'none',
              border: 'none',
              cursor: 'pointer',
              color: 'var(--ag-secondary-foreground-color, #666)',
              padding: '0 1px',
              fontSize: 14,
              lineHeight: 1,
              display: 'flex',
              alignItems: 'center',
            }}
          >
            ×
          </button>
        </span>
      ))}
      <input
        ref={inputRef}
        value={inputVal}
        onChange={e => setInputVal(e.target.value)}
        onKeyDown={handleKeyDown}
        style={{
          border: 'none',
          outline: 'none',
          background: 'transparent',
          fontSize: 13,
          minWidth: 80,
          flex: 1,
          color: 'var(--ag-foreground-color, #333)',
          padding: '1px 0',
        }}
        placeholder={items.length === 0 ? 'type + Enter to add…' : 'add more…'}
      />
    </div>
  );
}

import React, { useCallback } from 'react';
import type { Row } from './types';
import { INTERNAL_ROW_ID } from './types';
import RefCellEditor from './RefCellEditor';
import { parseListItems, serializeListItems } from './listValue';

interface RefListEditorProps {
  value: string;
  onChange: (value: string) => void;
  refRows: Row[];
  refTable: string;
  resolveColumnPath: (tableName: string, row: Row, path: string) => string;
  searchColumns: string[];
  displayColumns: string[];
  onCreateRecord?: (refTable: string, seedText: string) => Promise<string | null>;
  stopEditing?: () => void;
  autoFocus?: boolean;
}

export function resolveRefLabel(
  rowId: string,
  refRows: Row[],
  refTable: string,
  resolveColumnPath: (tableName: string, row: Row, path: string) => string,
  displayColumns: string[],
  searchColumns: string[],
): string {
  const refRow = refRows.find(row => row[INTERNAL_ROW_ID] === rowId);
  if (!refRow) return `[missing: ${rowId}]`;
  const cols = displayColumns.length > 0 ? displayColumns : searchColumns;
  const firstKey = Object.keys(refRow).find(key => key !== INTERNAL_ROW_ID);
  if (cols.length === 0 && firstKey) return resolveColumnPath(refTable, refRow, firstKey) || `Row ${rowId}`;
  if (cols.length === 0) return `Row ${rowId}`;
  return cols.map(path => resolveColumnPath(refTable, refRow, path)).filter(Boolean).join(' · ') || `Row ${rowId}`;
}

export const RefListEditor: React.FC<RefListEditorProps> = ({
  value, onChange, refRows, refTable, resolveColumnPath, searchColumns, displayColumns, onCreateRecord, stopEditing, autoFocus,
}) => {
  const ids = parseListItems(value);
  const setIds = useCallback((next: string[]) => onChange(serializeListItems(next)), [onChange]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, minWidth: 240 }}>
      {ids.length > 0 && (
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
          {ids.map((id, index) => (
            <span
              key={`${id}-${index}`}
              style={{
                background: 'color-mix(in srgb, var(--color-primary, #4f6bed) 22%, transparent)', color: 'var(--color-text, #333)', border: '1px solid color-mix(in srgb, var(--color-primary, #4f6bed) 45%, transparent)', borderRadius: 3,
                padding: '1px 4px 1px 7px', fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 2, lineHeight: '18px',
              }}
            >
              {resolveRefLabel(id, refRows, refTable, resolveColumnPath, displayColumns, searchColumns)}
              <button
                type="button"
                aria-label="Remove"
                tabIndex={-1}
                onMouseDown={event => { event.preventDefault(); setIds(ids.filter((_, i) => i !== index)); }}
                style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'inherit', padding: '0 1px', fontSize: 14, lineHeight: 1 }}
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}
      <div style={{ height: 32 }}>
        <RefCellEditor
          value=""
          onValueChange={() => {}}
          stopEditing={stopEditing ?? (() => {})}
          refRows={refRows}
          refTable={refTable}
          resolveColumnPath={resolveColumnPath}
          searchColumns={searchColumns}
          displayColumns={displayColumns}
          onCreateRecord={onCreateRecord}
          autoFocus={autoFocus}
          multiple
          selectedIds={ids}
          onSelectedIdsChange={setIds}
          placeholder={ids.length > 0 ? 'Add more…' : 'Search to add…'}
          inputClassName="app-dialog-input"
          inputStyle={{ marginBottom: 0, width: '100%', height: '100%', boxSizing: 'border-box' }}
        />
      </div>
    </div>
  );
};

export function RefListCellEditor(props: Omit<RefListEditorProps, 'value' | 'onChange'> & {
  value: string;
  onValueChange: (value: string | null) => void;
}) {
  const { value, onValueChange, ...rest } = props;
  const [draft, setDraft] = React.useState(value ?? '');
  const handleChange = useCallback((next: string) => {
    setDraft(next);
    onValueChange(next);
  }, [onValueChange]);
  return (
    <div style={{
      background: 'var(--ag-background-color, #fff)', border: '2px solid var(--ag-input-focus-border-color, #2196f3)',
      borderRadius: 4, padding: 6, maxWidth: 420,
    }}>
      <RefListEditor {...rest} value={draft} onChange={handleChange} />
    </div>
  );
}

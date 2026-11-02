import { useCallback, useEffect, useId, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import Select from 'react-select';
import { dialogSelectStyles } from './selectStyles';
import { formatColumnExpression, getColumnTransformations, parseColumnExpression } from './columnExpressions';
import type { ColumnPath } from './columnExpressions';
import type { DateFeature } from './types';

interface ColumnOption { value: string; label: string; configurable: boolean }

function TransformationModal({ column, initialFeature, onApply, onClose }: {
  column: ColumnPath;
  initialFeature: DateFeature | null;
  onApply: (expression: string) => void;
  onClose: () => void;
}) {
  const options = getColumnTransformations(column.type);
  const [feature, setFeature] = useState(() => options.some(option => option.value === initialFeature) ? initialFeature! : '');
  const titleId = useId();
  const selectId = useId();
  const descriptionId = useId();
  const dialogRef = useRef<HTMLDivElement>(null);
  const selectRef = useRef<HTMLSelectElement>(null);
  const description = options.find(option => option.value === feature)?.description;

  useEffect(() => {
    const previousFocus = document.activeElement;
    selectRef.current?.focus();
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        onClose();
      }
      if (event.key === 'Tab') {
        const controls = dialogRef.current?.querySelectorAll<HTMLElement>('select, button');
        if (!controls?.length) return;
        const first = controls[0];
        const last = controls[controls.length - 1];
        if (event.shiftKey && document.activeElement === first) {
          event.preventDefault();
          last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault();
          first.focus();
        }
      }
    };
    document.addEventListener('keydown', handleKeyDown, true);
    return () => {
      document.removeEventListener('keydown', handleKeyDown, true);
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected) previousFocus.focus();
    };
  }, [onClose]);

  return createPortal(
    <div
      className="app-dialog-overlay"
      style={{ zIndex: 10000 }}
      onMouseDown={event => { event.stopPropagation(); if (event.target === event.currentTarget) onClose(); }}
      onClick={event => event.stopPropagation()}
    >
      <div ref={dialogRef} className="app-dialog" role="dialog" aria-modal="true" aria-labelledby={titleId} style={{ width: 'min(420px, 94vw)', maxHeight: '90dvh', overflowY: 'auto' }}>
        <h2 id={titleId} className="app-dialog-title">Configure column</h2>
        <p className="app-dialog-message">{column.label}</p>
        <div className="app-dialog-field">
          <label className="app-dialog-label" htmlFor={selectId}>Transformation</label>
          <select ref={selectRef} id={selectId} className="app-dialog-select" value={feature} onChange={event => setFeature(event.target.value)} aria-describedby={description ? descriptionId : undefined}>
            <option value="">Raw value (no transformation)</option>
            {options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}
          </select>
          {description && <p id={descriptionId} style={{ fontSize: 12, color: 'var(--color-text-muted)' }}>{description}</p>}
        </div>
        <div className="app-dialog-actions">
          <button type="button" className="app-dialog-btn app-dialog-btn-secondary" onClick={onClose}>Cancel</button>
          <button type="button" className="app-dialog-btn app-dialog-btn-primary" onClick={() => onApply(feature ? `${column.path}:${feature}` : column.path)}>Apply</button>
        </div>
      </div>
    </div>,
    document.body,
  );
}

export function ColumnExpressionSelect({ columns, value, onChange, label, placeholder = 'Select column...', allowTransformations = true }: {
  columns: ColumnPath[];
  value: string;
  onChange: (expression: string) => void;
  label: string;
  placeholder?: string;
  allowTransformations?: boolean;
}) {
  const [pending, setPending] = useState<{ column: ColumnPath; feature: DateFeature | null } | null>(null);
  const closeModal = useCallback(() => setPending(null), []);
  const { col, feature } = parseColumnExpression(value);
  const groups = new Map<string, { label: string; options: ColumnOption[] }>();
  for (const column of columns) {
    const groupName = column.path.includes('.') ? column.path.split('.')[0] : 'Columns';
    if (!groups.has(groupName)) groups.set(groupName, { label: groupName, options: [] });
    groups.get(groupName)!.options.push({
      value: column.path,
      label: column.label,
      configurable: allowTransformations && getColumnTransformations(column.type).length > 0,
    });
  }
  const options = [...groups.values()];
  const selected = options.flatMap(group => group.options).find(option => option.value === col);

  return (
    <div style={{ minWidth: 0 }}>
      <Select<ColumnOption>
        styles={dialogSelectStyles}
        aria-label={label}
        value={selected ?? null}
        options={options}
        formatOptionLabel={(option, context) => context.context === 'menu'
          ? <span>{option.label}{option.configurable && <span title="Opens a configuration dialog" style={{ color: 'var(--color-text-muted)', fontSize: 11 }}> · Configure…</span>}</span>
          : option.value === col ? formatColumnExpression(value, columns) : option.label}
        onChange={option => {
          const column = columns.find(candidate => candidate.path === option?.value);
          if (column && option?.configurable) {
            setPending({ column, feature: column.path === col ? feature : null });
          } else {
            onChange(option?.value ?? '');
          }
        }}
        placeholder={placeholder}
        isClearable
        hideSelectedOptions={false}
        menuPortalTarget={document.body}
        menuPlacement="auto"
      />
      {pending && (
        <TransformationModal
          column={pending.column}
          initialFeature={pending.feature}
          onApply={expression => { onChange(expression); setPending(null); }}
          onClose={closeModal}
        />
      )}
    </div>
  );
}

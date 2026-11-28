import { useId, useLayoutEffect, useRef, useState } from 'react';

export interface AgendaFilters {
  search: string;
  from: string;
  through: string;
}

export function AgendaFilterDialog({ filters, onApply, onClose }: {
  filters: AgendaFilters;
  onApply: (filters: AgendaFilters) => void;
  onClose: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const backdropPressed = useRef(false);
  const id = useId();
  const [draft, setDraft] = useState(filters);
  const invalidRange = Boolean(draft.from && draft.through && draft.from > draft.through);

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    dialog?.showModal();
    return () => dialog?.close();
  }, []);

  return (
    <dialog
      ref={dialogRef}
      className="app-dialog agenda-filter-dialog"
      aria-labelledby={`${id}-title`}
      onCancel={event => { event.preventDefault(); onClose(); }}
      onPointerDown={event => {
        const { left, right, top, bottom } = event.currentTarget.getBoundingClientRect();
        backdropPressed.current = event.target === event.currentTarget
          && (event.clientX < left || event.clientX >= right || event.clientY < top || event.clientY >= bottom);
      }}
      onClick={event => {
        if (event.target === event.currentTarget && backdropPressed.current) onClose();
        backdropPressed.current = false;
      }}
    >
      <h2 id={`${id}-title`} className="app-dialog-title">Filter agenda</h2>
      <form onSubmit={event => {
        event.preventDefault();
        if (!invalidRange) onApply({ ...draft, search: draft.search.trim() });
      }}>
        <label className="app-dialog-field">
          <span className="app-dialog-label">Search events</span>
          <input
            type="search"
            className="app-dialog-input"
            value={draft.search}
            onChange={event => setDraft({ ...draft, search: event.target.value })}
            onKeyDown={event => {
              if (event.key === 'Escape') {
                event.preventDefault();
                onClose();
              }
            }}
          />
        </label>
        <div className="agenda-filter-dates">
          <label className="app-dialog-field">
            <span className="app-dialog-label">From</span>
            <input
              type="date"
              className="app-dialog-input"
              max="9999-12-31"
              value={draft.from}
              onChange={event => setDraft({ ...draft, from: event.target.value })}
            />
          </label>
          <label className="app-dialog-field">
            <span className="app-dialog-label">Through</span>
            <input
              type="date"
              className="app-dialog-input"
              max="9999-12-31"
              value={draft.through}
              aria-invalid={invalidRange || undefined}
              aria-describedby={invalidRange ? `${id}-range-error` : undefined}
              onChange={event => setDraft({ ...draft, through: event.target.value })}
            />
          </label>
        </div>
        {invalidRange && (
          <p id={`${id}-range-error`} className="agenda-filter-error" role="alert">
            Through must be on or after From.
          </p>
        )}
        <div className="app-dialog-actions">
          <button type="button" className="app-dialog-btn app-dialog-btn-secondary" onClick={() => setDraft({ search: '', from: '', through: '' })}>
            Clear
          </button>
          <button type="button" className="app-dialog-btn app-dialog-btn-secondary" onClick={onClose}>Cancel</button>
          <button type="submit" className="app-dialog-btn app-dialog-btn-primary" disabled={invalidRange}>Apply</button>
        </div>
      </form>
    </dialog>
  );
}

/**
 * Shared react-select styles for all modal/dialog dropdowns.
 * Adapts to the app's CSS custom properties so it works in dark and light mode.
 */

type Base = Record<string, unknown>;

export const dialogSelectStyles = {
  control: (base: Base, state: { isFocused: boolean }) => ({
    ...base,
    background: 'var(--bg)',
    borderColor: state.isFocused ? 'var(--ref-color)' : 'var(--border)',
    boxShadow: state.isFocused ? '0 0 0 2px rgba(37, 99, 235, 0.2)' : 'none',
    borderRadius: 4,
    fontSize: 13,
    minHeight: 34,
    '&:hover': { borderColor: 'var(--ref-color)' },
  }),
  menu: (base: Base) => ({
    ...base,
    background: 'var(--color-surface)',
    border: '1px solid var(--border)',
    borderRadius: 4,
    boxShadow: '0 4px 16px rgba(0,0,0,0.10)',
    zIndex: 9999,
  }),
  menuPortal: (base: Base) => ({ ...base, zIndex: 9999 }),
  option: (base: Base, state: { isSelected: boolean; isFocused: boolean }) => ({
    ...base,
    background: state.isSelected
      ? 'var(--primary)'
      : state.isFocused
        ? 'var(--cell-selected)'
        : 'transparent',
    color: state.isSelected ? 'var(--color-surface)' : 'var(--text)',
    fontSize: 13,
    padding: '6px 10px',
    cursor: 'pointer',
  }),
  singleValue: (base: Base) => ({ ...base, color: 'var(--text)' }),
  multiValue: (base: Base) => ({ ...base, background: 'var(--cell-selected, #e0e7ff)' }),
  multiValueLabel: (base: Base) => ({ ...base, color: 'var(--text)' }),
  input: (base: Base) => ({ ...base, color: 'var(--text)' }),
  placeholder: (base: Base) => ({ ...base, color: 'var(--text-muted)' }),
  indicatorSeparator: () => ({ display: 'none' }),
};

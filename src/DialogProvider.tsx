import React, { createContext, useContext, useState, useCallback, useRef } from 'react';

interface DialogButton {
  label: string;
  value: string;
  variant?: 'primary' | 'danger' | 'secondary';
}

interface SelectOption {
  label: string;
  value: string;
}

interface DialogConfig {
  title: string;
  message: string;
  buttons: DialogButton[];
  inputPlaceholder?: string;
  selectLabel?: string;
  selectOptions?: SelectOption[];
  selectDefault?: string;
}

interface DialogContextType {
  showDialog: (config: DialogConfig) => Promise<string | null>;
}

const DialogContext = createContext<DialogContextType | null>(null);

export function useDialog() {
  const ctx = useContext(DialogContext);
  if (!ctx) throw new Error('useDialog must be used within DialogProvider');
  return ctx;
}

// Convenience helpers
export function useConfirm() {
  const { showDialog } = useDialog();
  return useCallback(
    (message: string, title = 'Confirm') =>
      showDialog({
        title,
        message,
        buttons: [
          { label: 'Cancel', value: 'cancel', variant: 'secondary' },
          { label: 'Confirm', value: 'confirm', variant: 'primary' },
        ],
      }).then((v) => v === 'confirm'),
    [showDialog],
  );
}

export function useAlert() {
  const { showDialog } = useDialog();
  return useCallback(
    (message: string, title = 'Notice') =>
      showDialog({
        title,
        message,
        buttons: [{ label: 'OK', value: 'ok', variant: 'primary' }],
      }).then(() => {}),
    [showDialog],
  );
}

export function usePromptInput() {
  const { showDialog } = useDialog();
  return useCallback(
    (message: string, title = 'Input', placeholder = '') =>
      showDialog({
        title,
        message,
        inputPlaceholder: placeholder,
        buttons: [
          { label: 'Cancel', value: 'cancel', variant: 'secondary' },
          { label: 'OK', value: 'ok', variant: 'primary' },
        ],
      }),
    [showDialog],
  );
}

export const DialogProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [dialog, setDialog] = useState<DialogConfig | null>(null);
  const [inputValue, setInputValue] = useState('');
  const [selectValue, setSelectValue] = useState('');
  const resolveRef = useRef<((value: string | null) => void) | null>(null);

  const showDialog = useCallback((config: DialogConfig): Promise<string | null> => {
    return new Promise((resolve) => {
      resolveRef.current = resolve;
      setInputValue('');
      setSelectValue(config.selectDefault ?? config.selectOptions?.[0]?.value ?? '');
      setDialog(config);
    });
  }, []);

  const handleClick = (value: string | null) => {
    if (value === 'cancel' || value === null) {
      setDialog(null);
      resolveRef.current?.(null);
      resolveRef.current = null;
      return;
    }
    // For 'ok', return a JSON payload when there are combined fields
    if (value === 'ok') {
      const hasInput = dialog?.inputPlaceholder !== undefined;
      const hasSelect = (dialog?.selectOptions?.length ?? 0) > 0;
      let result: string;
      if (hasInput && hasSelect) {
        result = JSON.stringify({ input: inputValue, select: selectValue });
      } else if (hasInput) {
        result = inputValue;
      } else if (hasSelect) {
        result = selectValue;
      } else {
        result = 'ok';
      }
      setDialog(null);
      resolveRef.current?.(result);
      resolveRef.current = null;
      return;
    }
    setDialog(null);
    resolveRef.current?.(value);
    resolveRef.current = null;
  };

  return (
    <DialogContext.Provider value={{ showDialog }}>
      {children}
      {dialog && (
        <div className="app-dialog-overlay" onClick={() => handleClick(null)}>
          <div className="app-dialog" onClick={(e) => e.stopPropagation()}>
            <h3 className="app-dialog-title">{dialog.title}</h3>
            <p className="app-dialog-message">{dialog.message}</p>
            {dialog.selectOptions && dialog.selectOptions.length > 0 && (
              <div className="app-dialog-field">
                {dialog.selectLabel && <label className="app-dialog-label">{dialog.selectLabel}</label>}
                <select
                  className="app-dialog-select"
                  value={selectValue}
                  onChange={(e) => setSelectValue(e.target.value)}
                >
                  {dialog.selectOptions.map((opt) => (
                    <option key={opt.value} value={opt.value}>{opt.label}</option>
                  ))}
                </select>
              </div>
            )}
            {dialog.inputPlaceholder !== undefined && (
              <input
                className="app-dialog-input"
                type="text"
                placeholder={dialog.inputPlaceholder}
                value={inputValue}
                onChange={(e) => setInputValue(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') handleClick('ok'); }}
                autoFocus
              />
            )}
            <div className="app-dialog-actions">
              {dialog.buttons.map((btn) => (
                <button
                  key={btn.value}
                  className={`app-dialog-btn app-dialog-btn-${btn.variant ?? 'secondary'}`}
                  onClick={() => handleClick(btn.value)}
                >
                  {btn.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </DialogContext.Provider>
  );
};

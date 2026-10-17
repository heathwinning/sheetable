import React, { useCallback, useRef, useState } from 'react';
import * as api from './api';
import { parseListItems, serializeListItems } from './listValue';

interface ImageListEditorProps {
  value: string;
  bookId: string;
  onChange: (value: string) => void;
  readOnly?: boolean;
}

export const ImageListEditor: React.FC<ImageListEditorProps> = ({ value, bookId, onChange, readOnly = false }) => {
  const keys = parseListItems(value);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragIndex, setDragIndex] = useState<number | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const keysRef = useRef(keys);
  keysRef.current = keys;

  const commit = (next: string[]) => onChange(serializeListItems(next));

  const handleFiles = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    const images = Array.from(files).filter(file => file.type.startsWith('image/'));
    if (images.length < files.length) setError('Non-image files were skipped');
    else setError(null);
    if (images.length === 0) return;
    setUploading(true);
    try {
      const results = await Promise.allSettled(images.map(async file => {
        const { key, uploadUrl } = await api.getUploadUrl(bookId, file.name);
        await api.uploadImage(uploadUrl, file);
        return key;
      }));
      const uploaded = results.flatMap(result => result.status === 'fulfilled' ? [result.value] : []);
      if (uploaded.length < images.length) setError(`${images.length - uploaded.length} upload(s) failed`);
      if (uploaded.length > 0) onChange(serializeListItems([...keysRef.current, ...uploaded]));
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }, [bookId, onChange]);

  const move = (from: number, to: number) => {
    if (to < 0 || to >= keys.length || from === to) return;
    const next = [...keys];
    const [item] = next.splice(from, 1);
    next.splice(to, 0, item);
    commit(next);
  };

  const iconButton: React.CSSProperties = {
    background: 'rgba(0,0,0,0.6)', color: '#fff', border: 'none', borderRadius: 3,
    width: 18, height: 18, fontSize: 11, lineHeight: '18px', padding: 0, cursor: 'pointer',
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
        {keys.length === 0 && (
          <div style={{ fontSize: 12, color: 'var(--color-text-muted)', padding: '6px 0' }}>No images</div>
        )}
        {keys.map((key, index) => (
          <div
            key={`${key}-${index}`}
            draggable={!readOnly}
            onDragStart={() => setDragIndex(index)}
            onDragOver={event => { if (dragIndex !== null) event.preventDefault(); }}
            onDrop={event => {
              event.preventDefault();
              if (dragIndex !== null) move(dragIndex, index);
              setDragIndex(null);
            }}
            onDragEnd={() => setDragIndex(null)}
            style={{
              position: 'relative', width: 84, height: 84, borderRadius: 6, overflow: 'hidden',
              border: '1px solid var(--color-border)', opacity: dragIndex === index ? 0.4 : 1,
              cursor: readOnly ? 'default' : 'grab',
            }}
          >
            <img src={api.imageUrl(bookId, key)} alt="" draggable={false} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            {!readOnly && (
              <>
                <button type="button" aria-label="Remove image" title="Remove" onClick={() => commit(keys.filter((_, i) => i !== index))} style={{ ...iconButton, position: 'absolute', top: 3, right: 3 }}>×</button>
                <div style={{ position: 'absolute', bottom: 3, left: 3, display: 'flex', gap: 3 }}>
                  <button type="button" aria-label="Move earlier" title="Move earlier" disabled={index === 0} onClick={() => move(index, index - 1)} style={{ ...iconButton, opacity: index === 0 ? 0.4 : 1 }}>‹</button>
                  <button type="button" aria-label="Move later" title="Move later" disabled={index === keys.length - 1} onClick={() => move(index, index + 1)} style={{ ...iconButton, opacity: index === keys.length - 1 ? 0.4 : 1 }}>›</button>
                </div>
              </>
            )}
          </div>
        ))}
      </div>
      {!readOnly && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <input ref={fileInputRef} type="file" accept="image/*" multiple style={{ display: 'none' }} onChange={event => { void handleFiles(event.target.files); }} />
          <button type="button" className="app-dialog-btn app-dialog-btn-secondary btn-sm" disabled={uploading} onClick={() => fileInputRef.current?.click()}>
            {uploading ? 'Uploading…' : keys.length > 0 ? 'Add images' : 'Upload images'}
          </button>
          {error && <span style={{ fontSize: 12, color: 'var(--color-danger)' }}>{error}</span>}
        </div>
      )}
    </div>
  );
};

interface ImageListDialogState {
  value: string;
  bookId: string;
  readOnly: boolean;
  onSave: (value: string) => void;
}

const ImageListDialogBody: React.FC<{ initial: ImageListDialogState; onClose: () => void }> = ({ initial, onClose }) => {
  const [draft, setDraft] = useState(initial.value);
  return (
    <div className="app-dialog-overlay" onClick={() => onClose()}>
      <div className="app-dialog" onClick={event => event.stopPropagation()} style={{ width: 'min(520px, 94vw)', color: 'var(--color-text)' }}>
        <h3 className="app-dialog-title">Images</h3>
        <ImageListEditor value={draft} bookId={initial.bookId} onChange={setDraft} readOnly={initial.readOnly} />
        <div className="app-dialog-actions">
          <button className="app-dialog-btn app-dialog-btn-secondary" onClick={() => onClose()}>{initial.readOnly ? 'Close' : 'Cancel'}</button>
          {!initial.readOnly && (
            <button className="app-dialog-btn app-dialog-btn-primary" onClick={() => { initial.onSave(draft); onClose(); }}>Save</button>
          )}
        </div>
      </div>
    </div>
  );
};

export function useImageListDialog() {
  const [state, setState] = useState<ImageListDialogState | null>(null);

  const openListDialog = useCallback(
    (value: string, bookId: string, readOnly: boolean, onSave: (value: string) => void) =>
      setState({ value, bookId, readOnly, onSave }),
    [],
  );

  const onClose = useCallback(() => setState(null), []);
  const listDialogElement = state ? <ImageListDialogBody initial={state} onClose={onClose} /> : null;
  return { openListDialog, listDialogElement };
}

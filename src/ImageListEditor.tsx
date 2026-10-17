import React, { useCallback, useRef, useState } from 'react';
import * as api from './api';
import { parseListItems, serializeListItems } from './listValue';
import { ImageViewer } from './ImageViewer';

interface ImageListEditorProps {
  value: string;
  bookId: string;
  onChange: (value: string) => void;
  readOnly?: boolean;
}

const iconButton: React.CSSProperties = {
  background: 'rgba(0,0,0,0.6)', color: '#fff', border: 'none', borderRadius: 3,
  width: 18, height: 18, fontSize: 11, lineHeight: '18px', padding: 0, cursor: 'pointer',
};

function moveItem<T>(items: T[], from: number, to: number): T[] {
  if (to < 0 || to >= items.length || from === to) return items;
  const next = [...items];
  const [item] = next.splice(from, 1);
  next.splice(to, 0, item);
  return next;
}

interface ImageOrderEditorProps {
  items: string[];
  bookId: string;
  onChange: (items: string[]) => void;
}

export const ImageOrderEditor: React.FC<ImageOrderEditorProps> = ({ items, bookId, onChange }) => {
  const [dragIndex, setDragIndex] = useState<number | null>(null);

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
      {items.map((key, index) => (
        <div
          key={`${key}-${index}`}
          draggable
          onDragStart={() => setDragIndex(index)}
          onDragOver={event => { if (dragIndex !== null) event.preventDefault(); }}
          onDrop={event => {
            event.preventDefault();
            if (dragIndex !== null) onChange(moveItem(items, dragIndex, index));
            setDragIndex(null);
          }}
          onDragEnd={() => setDragIndex(null)}
          style={{
            position: 'relative', width: 96, height: 96, borderRadius: 6, overflow: 'hidden',
            border: '1px solid var(--color-border)', opacity: dragIndex === index ? 0.4 : 1, cursor: 'grab',
          }}
        >
          <img src={api.imageUrl(bookId, key)} alt="" draggable={false} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
          <div style={{ position: 'absolute', bottom: 3, left: 3, right: 3, display: 'flex', justifyContent: 'space-between' }}>
            <button type="button" aria-label="Move earlier" title="Move earlier" disabled={index === 0} onClick={() => onChange(moveItem(items, index, index - 1))} style={{ ...iconButton, opacity: index === 0 ? 0.4 : 1 }}>‹</button>
            <span style={{ fontSize: 11, color: '#fff', background: 'rgba(0,0,0,0.6)', borderRadius: 3, padding: '0 4px', lineHeight: '18px' }}>{index + 1}</span>
            <button type="button" aria-label="Move later" title="Move later" disabled={index === items.length - 1} onClick={() => onChange(moveItem(items, index, index + 1))} style={{ ...iconButton, opacity: index === items.length - 1 ? 0.4 : 1 }}>›</button>
          </div>
        </div>
      ))}
    </div>
  );
};

export const ImageListEditor: React.FC<ImageListEditorProps> = ({ value, bookId, onChange, readOnly = false }) => {
  const keys = parseListItems(value);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [viewIndex, setViewIndex] = useState<number | null>(null);
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

  const iconLabel: React.CSSProperties = { ...iconButton, position: 'absolute', top: 3, right: 3 };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
        {keys.length === 0 && (
          <div style={{ fontSize: 12, color: 'var(--color-text-muted)', padding: '6px 0' }}>No images</div>
        )}
        {keys.map((key, index) => (
          <div
            key={`${key}-${index}`}
            style={{ position: 'relative', width: 84, height: 84, borderRadius: 6, overflow: 'hidden', border: '1px solid var(--color-border)' }}
          >
            <button
              type="button"
              aria-label={`View image ${index + 1}`}
              onClick={() => setViewIndex(index)}
              style={{ display: 'block', width: '100%', height: '100%', padding: 0, border: 'none', background: 'none', cursor: 'zoom-in' }}
            >
              <img src={api.imageUrl(bookId, key)} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
            </button>
            {!readOnly && (
              <button type="button" aria-label="Remove image" title="Remove" onClick={() => commit(keys.filter((_, i) => i !== index))} style={iconLabel}>×</button>
            )}
          </div>
        ))}
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        {!readOnly && (
          <>
            <input ref={fileInputRef} type="file" accept="image/*" multiple style={{ display: 'none' }} onChange={event => { void handleFiles(event.target.files); }} />
            <button type="button" className="app-dialog-btn app-dialog-btn-secondary btn-sm" disabled={uploading} onClick={() => fileInputRef.current?.click()}>
              {uploading ? 'Uploading…' : keys.length > 0 ? 'Add images' : 'Upload images'}
            </button>
          </>
        )}
        {error && <span style={{ fontSize: 12, color: 'var(--color-danger)' }}>{error}</span>}
      </div>
      {viewIndex !== null && keys.length > 0 && (
        <ImageViewer
          images={keys}
          index={Math.min(viewIndex, keys.length - 1)}
          bookId={bookId}
          onIndexChange={setViewIndex}
          onClose={() => setViewIndex(null)}
        />
      )}
    </div>
  );
};

export function useImageListDialog() {
  const [listState, setListState] = useState<{ value: string; bookId: string; readOnly: boolean; onSave: (value: string) => void } | null>(null);
  const [viewState, setViewState] = useState<{ images: string[]; index: number; bookId: string } | null>(null);

  const openListDialog = useCallback(
    (value: string, bookId: string, readOnly: boolean, onSave: (value: string) => void) =>
      setListState({ value, bookId, readOnly, onSave }),
    [],
  );
  const openImageViewer = useCallback((images: string[], index: number, bookId: string) => setViewState({ images, index, bookId }), []);

  const listDialogElement = listState ? (
    <ImageListDialog initial={listState} onClose={() => setListState(null)} />
  ) : null;
  const viewerElement = viewState ? (
    <ImageViewer
      images={viewState.images}
      index={viewState.index}
      bookId={viewState.bookId}
      onIndexChange={index => setViewState(prev => prev && ({ ...prev, index }))}
      onClose={() => setViewState(null)}
    />
  ) : null;
  return { openListDialog, openImageViewer, listDialogElement, viewerElement };
}

const ImageListDialog: React.FC<{
  initial: { value: string; bookId: string; readOnly: boolean; onSave: (value: string) => void };
  onClose: () => void;
}> = ({ initial, onClose }) => {
  const [draft, setDraft] = useState(initial.value);
  const [reordering, setReordering] = useState(false);
  const [order, setOrder] = useState<string[]>([]);

  return (
    <div className="app-dialog-overlay" onClick={onClose}>
      <div className="app-dialog" onClick={event => event.stopPropagation()} style={{ width: 'min(560px, 94vw)', color: 'var(--color-text)' }}>
        <h3 className="app-dialog-title">Images</h3>
        {reordering ? (
          <>
            <p className="app-dialog-message" style={{ fontSize: 12 }}>Drag images or use the arrows to change their order.</p>
            <ImageOrderEditor items={order} bookId={initial.bookId} onChange={setOrder} />
            <div className="app-dialog-actions">
              <button className="app-dialog-btn app-dialog-btn-secondary" onClick={() => setReordering(false)}>Back</button>
              <button className="app-dialog-btn app-dialog-btn-primary" onClick={() => { setDraft(serializeListItems(order)); setReordering(false); }}>Apply order</button>
            </div>
          </>
        ) : (
          <>
            <ImageListEditor value={draft} bookId={initial.bookId} onChange={setDraft} readOnly={initial.readOnly} />
            <div className="app-dialog-actions">
              {!initial.readOnly && parseListItems(draft).length > 1 && (
                <button className="app-dialog-btn app-dialog-btn-secondary" onClick={() => { setOrder(parseListItems(draft)); setReordering(true); }}>Reorder</button>
              )}
              <button className="app-dialog-btn app-dialog-btn-secondary" onClick={onClose}>{initial.readOnly ? 'Close' : 'Cancel'}</button>
              {!initial.readOnly && (
                <button className="app-dialog-btn app-dialog-btn-primary" onClick={() => { initial.onSave(draft); onClose(); }}>Save</button>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
};

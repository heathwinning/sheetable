import React, { useCallback, useEffect } from 'react';
import * as api from './api';

interface ImageViewerProps {
  images: string[];
  index: number;
  bookId: string;
  onIndexChange: (index: number) => void;
  onClose: () => void;
}

export const ImageViewer: React.FC<ImageViewerProps> = ({ images, index, bookId, onIndexChange, onClose }) => {
  const count = images.length;
  const prev = useCallback(() => onIndexChange((index + count - 1) % count), [index, count, onIndexChange]);
  const next = useCallback(() => onIndexChange((index + 1) % count), [index, count, onIndexChange]);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
      else if (event.key === 'ArrowLeft' && count > 1) prev();
      else if (event.key === 'ArrowRight' && count > 1) next();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [count, onClose, prev, next]);

  if (count === 0) return null;
  const navButton: React.CSSProperties = {
    background: 'rgba(0,0,0,0.5)', color: '#fff', border: 'none', borderRadius: 999,
    width: 40, height: 40, fontSize: 24, lineHeight: '40px', cursor: 'pointer', flexShrink: 0,
  };

  return (
    <div className="image-dialog-overlay" onClick={onClose}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, maxWidth: '96vw' }} onClick={event => event.stopPropagation()}>
        {count > 1 && <button type="button" aria-label="Previous image" onClick={prev} style={navButton}>‹</button>}
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
          <img
            src={api.imageUrl(bookId, images[index])}
            alt={count > 1 ? `Image ${index + 1} of ${count}` : ''}
            style={{ maxWidth: '86vw', maxHeight: '82vh', objectFit: 'contain', borderRadius: 6, background: 'var(--color-surface)' }}
          />
          {count > 1 && <span style={{ color: '#fff', fontSize: 12 }}>{index + 1} / {count}</span>}
        </div>
        {count > 1 && <button type="button" aria-label="Next image" onClick={next} style={navButton}>›</button>}
        <button type="button" aria-label="Close" onClick={onClose} style={{ ...navButton, position: 'absolute', top: 16, right: 16 }}>×</button>
      </div>
    </div>
  );
};

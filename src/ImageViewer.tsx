import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import * as api from './api';

interface ImageViewerProps {
  images: string[];
  index: number;
  bookId: string;
  onIndexChange: (index: number) => void;
  onClose: () => void;
}

const controlButton: React.CSSProperties = {
  position: 'absolute',
  background: 'rgba(0,0,0,0.5)',
  color: '#fff',
  border: 'none',
  borderRadius: 999,
  cursor: 'pointer',
  zIndex: 1,
};

export const ImageViewer: React.FC<ImageViewerProps> = ({ images, index, bookId, onIndexChange, onClose }) => {
  const trackRef = useRef<HTMLDivElement>(null);
  const initialIndex = useRef(index);
  const [current, setCurrent] = useState(index);
  const count = images.length;

  useLayoutEffect(() => {
    const track = trackRef.current;
    if (track) track.scrollLeft = initialIndex.current * track.clientWidth;
  }, []);

  const goTo = useCallback((next: number) => {
    const track = trackRef.current;
    if (!track) return;
    track.scrollTo({ left: next * track.clientWidth, behavior: 'smooth' });
  }, []);

  const handleScroll = () => {
    const track = trackRef.current;
    if (!track || track.clientWidth === 0) return;
    const next = Math.min(count - 1, Math.round(track.scrollLeft / track.clientWidth));
    if (next !== current) {
      setCurrent(next);
      onIndexChange(next);
    }
  };

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
      else if (event.key === 'ArrowLeft') goTo(Math.max(0, current - 1));
      else if (event.key === 'ArrowRight') goTo(Math.min(count - 1, current + 1));
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [current, count, goTo, onClose]);

  if (count === 0) return null;

  return (
    <div
      onClick={event => event.stopPropagation()}
      style={{ position: 'fixed', inset: 0, zIndex: 2100, background: 'rgba(0,0,0,0.92)', touchAction: 'pan-x' }}
    >
      <div
        ref={trackRef}
        onScroll={handleScroll}
        style={{
          display: 'flex',
          width: '100%',
          height: '100%',
          overflowX: 'auto',
          scrollSnapType: 'x mandatory',
          overscrollBehaviorX: 'contain',
          scrollbarWidth: 'none',
        }}
      >
        {images.map((key, slide) => (
          <div
            key={`${key}-${slide}`}
            onClick={event => { if (event.target === event.currentTarget) onClose(); }}
            style={{
              flex: '0 0 100%',
              height: '100%',
              scrollSnapAlign: 'center',
              scrollSnapStop: 'always',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              padding: 16,
              boxSizing: 'border-box',
            }}
          >
            <img
              src={api.imageUrl(bookId, key)}
              alt={`Image ${slide + 1} of ${count}`}
              draggable={false}
              style={{
                maxWidth: '100%',
                maxHeight: '100%',
                objectFit: 'contain',
                borderRadius: 6,
                userSelect: 'none',
                transform: slide === current ? 'scale(1)' : 'scale(0.92)',
                opacity: slide === current ? 1 : 0.5,
                transition: 'transform 0.25s ease, opacity 0.25s ease',
              }}
            />
          </div>
        ))}
      </div>
      {count > 1 && current > 0 && (
        <button
          type="button"
          aria-label="Previous image"
          onClick={() => goTo(current - 1)}
          style={{ ...controlButton, left: 12, top: '50%', transform: 'translateY(-50%)', width: 44, height: 44, fontSize: 26, lineHeight: '44px' }}
        >
          ‹
        </button>
      )}
      {count > 1 && current < count - 1 && (
        <button
          type="button"
          aria-label="Next image"
          onClick={() => goTo(current + 1)}
          style={{ ...controlButton, right: 12, top: '50%', transform: 'translateY(-50%)', width: 44, height: 44, fontSize: 26, lineHeight: '44px' }}
        >
          ›
        </button>
      )}
      <button
        type="button"
        aria-label="Close"
        onClick={onClose}
        style={{ ...controlButton, right: 12, top: 'max(12px, env(safe-area-inset-top))', width: 40, height: 40, fontSize: 22, lineHeight: '40px' }}
      >
        ×
      </button>
      {count > 1 && (
        <div style={{ position: 'absolute', left: 0, right: 0, bottom: 'max(16px, env(safe-area-inset-bottom))', textAlign: 'center', color: '#fff', fontSize: 13, pointerEvents: 'none' }}>
          {current + 1} / {count}
        </div>
      )}
    </div>
  );
};

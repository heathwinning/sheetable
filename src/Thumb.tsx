import React, { useCallback, useState } from 'react';
import * as api from './api';

const failedBackfills = new Set<string>();
const pending: Array<() => Promise<void>> = [];
let running = 0;
const MAX_CONCURRENT_BACKFILLS = 2;

function enqueue(task: () => Promise<void>) {
  pending.push(task);
  pump();
}

function pump() {
  while (running < MAX_CONCURRENT_BACKFILLS && pending.length) {
    const task = pending.shift()!;
    running++;
    task().finally(() => { running--; pump(); });
  }
}

async function backfillThumbnail(bookId: string, key: string): Promise<string | null> {
  const fullUrl = api.imageUrl(bookId, key);
  const res = await fetch(fullUrl);
  if (!res.ok) return null;
  const thumb = await api.createThumbnail(await res.blob());
  if (!thumb) return null;
  await api.uploadThumbnail(fullUrl, thumb);
  return URL.createObjectURL(thumb);
}

interface ThumbProps {
  bookId: string;
  imageKey: string;
  style?: React.CSSProperties;
  className?: string;
  title?: string;
  onClick?: React.MouseEventHandler<HTMLImageElement>;
  draggable?: boolean;
}

export const Thumb: React.FC<ThumbProps> = ({ bookId, imageKey, style, className, title, onClick, draggable }) => {
  const [src, setSrc] = useState(() => api.thumbnailUrl(bookId, imageKey));
  const [prevKey, setPrevKey] = useState(`${bookId}/${imageKey}`);
  if (prevKey !== `${bookId}/${imageKey}`) {
    setPrevKey(`${bookId}/${imageKey}`);
    setSrc(api.thumbnailUrl(bookId, imageKey));
  }

  const handleError = useCallback(() => {
    const id = `${bookId}/${imageKey}`;
    if (failedBackfills.has(id)) {
      setSrc(api.imageUrl(bookId, imageKey));
      return;
    }
    failedBackfills.add(id);
    enqueue(async () => {
      try {
        const url = await backfillThumbnail(bookId, imageKey);
        setSrc(url ?? api.imageUrl(bookId, imageKey));
      } catch {
        setSrc(api.imageUrl(bookId, imageKey));
      }
    });
  }, [bookId, imageKey]);

  return (
    <img
      src={src}
      alt=""
      loading="lazy"
      decoding="async"
      className={className}
      title={title}
      onClick={onClick}
      draggable={draggable}
      onError={handleError}
      style={{ background: 'var(--color-border)', ...style }}
    />
  );
};

import React, { useCallback, useEffect, useRef, useState } from 'react';
import * as api from './api';

const failedBackfills = new Set<string>();
interface BackfillJob { id: string; controller: AbortController; run: (signal: AbortSignal) => Promise<void>; }
const pending: BackfillJob[] = [];
let running = 0;
const MAX_CONCURRENT_BACKFILLS = 2;

function enqueue(job: BackfillJob) {
  pending.push(job);
  pump();
}

function cancel(job: BackfillJob) {
  job.controller.abort();
  const index = pending.indexOf(job);
  if (index >= 0) pending.splice(index, 1);
  failedBackfills.delete(job.id);
}

function pump() {
  while (running < MAX_CONCURRENT_BACKFILLS && pending.length) {
    const job = pending.shift()!;
    running++;
    job.run(job.controller.signal).finally(() => { running--; pump(); });
  }
}

async function backfillThumbnail(bookId: string, key: string, signal: AbortSignal): Promise<string | null> {
  const fullUrl = api.imageUrl(bookId, key);
  const res = await fetch(fullUrl, { signal });
  if (!res.ok) return null;
  const blob = await res.blob();
  if (signal.aborted) return null;
  const thumb = await api.createThumbnail(blob);
  if (!thumb || signal.aborted) return null;
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

  const jobRef = useRef<BackfillJob | null>(null);

  useEffect(() => () => {
    if (jobRef.current) cancel(jobRef.current);
    jobRef.current = null;
  }, []);

  const handleError = useCallback(() => {
    const id = `${bookId}/${imageKey}`;
    if (failedBackfills.has(id)) {
      setSrc(api.imageUrl(bookId, imageKey));
      return;
    }
    failedBackfills.add(id);
    const controller = new AbortController();
    const job: BackfillJob = {
      id,
      controller,
      run: async signal => {
        try {
          const url = await backfillThumbnail(bookId, imageKey, signal);
          if (!signal.aborted) setSrc(url ?? api.imageUrl(bookId, imageKey));
        } catch {
          if (!signal.aborted) setSrc(api.imageUrl(bookId, imageKey));
        }
      },
    };
    jobRef.current = job;
    enqueue(job);
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

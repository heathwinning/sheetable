import { useEffect, useState } from 'react';
import type { UseAppStateReturn } from './useAppState';

export function useSheetRows(state: UseAppStateReturn, roots: string[]) {
  const { activeBookId, bookReady, getRequiredTables, hasRows, loadRows } = state;
  const missing = bookReady ? getRequiredTables(roots).filter(table => !hasRows(table)) : [];
  const key = JSON.stringify(missing);
  const scope = JSON.stringify(roots);
  const [attempt, setAttempt] = useState(0);
  const [failure, setFailure] = useState<{ bookId: string; scope: string; attempt: number; message: string } | null>(null);

  useEffect(() => {
    if (!activeBookId || !bookReady || key === '[]') return;
    let cancelled = false;
    void loadRows(JSON.parse(key) as string[]).catch(error => {
      if (!cancelled) {
        setFailure({ bookId: activeBookId, scope, attempt, message: error instanceof Error ? error.message : 'Failed to load rows' });
      }
    });
    return () => { cancelled = true; };
  }, [activeBookId, attempt, bookReady, key, loadRows, scope]);

  return {
    loading: missing.length > 0,
    error: missing.length > 0 && failure?.bookId === activeBookId && failure.scope === scope && failure.attempt === attempt ? failure.message : null,
    retry: () => setAttempt(current => current + 1),
  };
}

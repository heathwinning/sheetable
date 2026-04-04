import type { ColumnType } from './types';

function makeDate(y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): Date | null {
  const dt = new Date(y, m - 1, d, hh, mm, ss);
  if (
    dt.getFullYear() !== y ||
    dt.getMonth() !== m - 1 ||
    dt.getDate() !== d ||
    dt.getHours() !== hh ||
    dt.getMinutes() !== mm ||
    dt.getSeconds() !== ss
  ) {
    return null;
  }
  return dt;
}

export function parseTemporalUnknown(raw: unknown): Date | null {
  if (raw instanceof Date) {
    return Number.isNaN(raw.getTime()) ? null : raw;
  }

  const text = String(raw ?? '').trim();
  if (!text) return null;

  // YYYY/MM/DD or YYYY-MM-DD with optional time.
  let m = text.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const hh = Number(m[4] ?? 0);
    const mm = Number(m[5] ?? 0);
    const ss = Number(m[6] ?? 0);
    return makeDate(y, mo, d, hh, mm, ss);
  }

  // DD/MM/YYYY with optional time (legacy compatibility).
  m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);
    const hh = Number(m[4] ?? 0);
    const mm = Number(m[5] ?? 0);
    const ss = Number(m[6] ?? 0);
    return makeDate(y, mo, d, hh, mm, ss);
  }

  return null;
}

export function formatDateCanonical(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

export function formatDateTimeCanonical(date: Date): string {
  const base = formatDateCanonical(date);
  const hh = String(date.getHours()).padStart(2, '0');
  const mm = String(date.getMinutes()).padStart(2, '0');
  const ss = String(date.getSeconds()).padStart(2, '0');
  return `${base} ${hh}:${mm}:${ss}`;
}

export function normalizeTemporalString(value: string, type: ColumnType): string {
  if (type !== 'date' && type !== 'datetime') return value;
  const parsed = parseTemporalUnknown(value);
  if (!parsed) return value;
  return type === 'datetime' ? formatDateTimeCanonical(parsed) : formatDateCanonical(parsed);
}

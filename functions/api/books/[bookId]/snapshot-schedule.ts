import type { Env, RequestData } from '../../../lib';
import { json, error, requireUser, requireEditor } from '../../../lib';

// GET /api/books/:bookId/snapshot-schedule
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);
  const bookId = context.params.bookId as string;

  const row = await context.env.DB.prepare(
    'SELECT enabled, interval_days, next_run_at, updated_at FROM _snapshot_schedules WHERE book_id = ?'
  ).bind(bookId).first<{
    enabled: number; interval_days: number;
    next_run_at: string | null; updated_at: string;
  }>();

  return json(row ? {
    enabled: row.enabled === 1,
    intervalDays: row.interval_days,
    nextRunAt: row.next_run_at,
    updatedAt: row.updated_at,
  } : {
    enabled: false,
    intervalDays: 1,
    nextRunAt: null,
    updatedAt: null,
  });
};

// PUT /api/books/:bookId/snapshot-schedule
export const onRequestPut: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);
  const bookId = context.params.bookId as string;
  const body = await context.request.json() as {
    enabled?: boolean;
    intervalDays?: number;
  };

  if (body.intervalDays !== undefined && (body.intervalDays < 1 || body.intervalDays > 365)) {
    return error('intervalDays must be between 1 and 365');
  }

  const db = context.env.DB;
  const now = new Date().toISOString();
  const interval = body.intervalDays ?? 1;

  let nextRunAt: string | null = null;
  if (body.enabled) {
    const next = new Date();
    next.setDate(next.getDate() + interval);
    nextRunAt = next.toISOString();
  }

  await db.prepare(
    `INSERT INTO _snapshot_schedules (book_id, enabled, interval_days, next_run_at, updated_at)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(book_id) DO UPDATE SET
       enabled = COALESCE(?, enabled),
       interval_days = COALESCE(?, interval_days),
       next_run_at = COALESCE(?, next_run_at),
       updated_at = ?`
  ).bind(
    bookId, body.enabled ? 1 : 0, interval,
    body.enabled ? nextRunAt : null, now,
    body.enabled !== undefined ? (body.enabled ? 1 : 0) : null,
    body.intervalDays ?? null,
    body.enabled ? nextRunAt : null, now,
  ).run();

  const row = await db.prepare(
    'SELECT enabled, interval_days, next_run_at, updated_at FROM _snapshot_schedules WHERE book_id = ?'
  ).bind(bookId).first<{
    enabled: number; interval_days: number;
    next_run_at: string | null; updated_at: string;
  }>();

  return json({
    enabled: row!.enabled === 1,
    intervalDays: row!.interval_days,
    nextRunAt: row!.next_run_at,
    updatedAt: row!.updated_at,
  });
};

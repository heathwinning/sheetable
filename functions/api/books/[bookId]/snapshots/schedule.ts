import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

interface ScheduleConfig {
  enabled: boolean;
  frequency: 'daily' | 'weekly' | 'monthly';
  nextRunAt: string | null;
}

async function getSchedule(db: D1Database, bookId: string): Promise<ScheduleConfig | null> {
  const row = await db.prepare(
    'SELECT enabled, frequency, next_run_at FROM _snapshot_schedules WHERE book_id = ?'
  ).bind(bookId).first<{ enabled: number; frequency: string; next_run_at: string | null }>();
  if (!row) return null;
  return {
    enabled: row.enabled === 1,
    frequency: row.frequency as 'daily' | 'weekly' | 'monthly',
    nextRunAt: row.next_run_at,
  };
}

// GET /api/books/:bookId/snapshots/schedule
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;
  const schedule = await getSchedule(context.env.DB, bookId);
  return json(schedule ?? { enabled: false, frequency: 'daily', nextRunAt: null });
};

// PUT /api/books/:bookId/snapshots/schedule
export const onRequestPut: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as { enabled?: boolean; frequency?: string };

  const validFrequencies = ['daily', 'weekly', 'monthly'];
  const frequency = body.frequency ?? 'daily';
  if (!validFrequencies.includes(frequency)) {
    return error(`Invalid frequency. Must be one of: ${validFrequencies.join(', ')}`);
  }

  const enabled = body.enabled !== false; // default to true when creating

  // Compute next run time
  const now = new Date();
  let nextRunAt: string;
  switch (frequency) {
    case 'daily':
      nextRunAt = new Date(now.getTime() + 24 * 60 * 60 * 1000).toISOString();
      break;
    case 'weekly':
      nextRunAt = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
      break;
    case 'monthly':
      nextRunAt = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString();
      break;
    default:
      nextRunAt = new Date(now.getTime() + 24 * 60 * 60 * 1000).toISOString();
  }

  await context.env.DB.prepare(
    `INSERT INTO _snapshot_schedules (book_id, enabled, frequency, next_run_at, updated_at)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(book_id) DO UPDATE SET
       enabled = excluded.enabled,
       frequency = excluded.frequency,
       next_run_at = excluded.next_run_at,
       updated_at = excluded.updated_at`
  ).bind(
    bookId,
    enabled ? 1 : 0,
    frequency,
    nextRunAt,
    new Date().toISOString(),
  ).run();

  return json({
    enabled,
    frequency,
    nextRunAt,
  });
};

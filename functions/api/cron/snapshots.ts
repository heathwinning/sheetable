import type { Env, RequestData } from '../../lib';
import { json, error } from '../../lib';
import { gatherExportPayload } from '../books/[bookId]/export';
import { createSnapshot } from '../books/[bookId]/snapshots/storage';

// GET /api/cron/snapshots — called by GitHub Actions scheduled workflow
// Requires Authorization: Bearer <CRON_SECRET> header
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  // Auth check — shared secret with GitHub Actions
  const auth = context.request.headers.get('Authorization') ?? '';
  const expected = `Bearer ${context.env.CRON_SECRET}`;
  if (!context.env.CRON_SECRET || auth !== expected) {
    return error('Unauthorized', 401);
  }
  const db = context.env.DB;
  const now = new Date().toISOString();

  // Find all enabled schedules whose next_run_at is in the past
  const { results: schedules } = await db.prepare(
    `SELECT book_id, interval_days FROM _snapshot_schedules
     WHERE enabled = 1 AND next_run_at <= ?`
  ).bind(now).all<{ book_id: string; interval_days: number }>();

  const results: { bookId: string; snapshotId: string; rowCount: number }[] = [];

  for (const sched of schedules) {
    try {
      const payload = await gatherExportPayload(context.env, sched.book_id);
      await createSnapshot(context.env, sched.book_id, payload, `Auto snapshot`);

      // Compute next run time
      const nextDate = new Date();
      nextDate.setDate(nextDate.getDate() + (sched.interval_days || 1));

      await db.prepare(
        'UPDATE _snapshot_schedules SET next_run_at = ?, updated_at = ? WHERE book_id = ?'
      ).bind(nextDate.toISOString(), now, sched.book_id).run();

      results.push({ bookId: sched.book_id, snapshotId: 'created', rowCount: 0 });
    } catch (err) {
      console.error(`Failed to snapshot book ${sched.book_id}:`, err);
    }
  }

  return json({ created: results.length, snapshots: results });
};

import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';
import { readSnapshotData, deleteSnapshotData } from './storage';

// ---- GET /api/books/:bookId/snapshots/:snapshotId ----

export const onRequestGet: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;

  const row = await context.env.DB.prepare(
    `SELECT id, book_id, label, created_at, table_count, row_count, view_count, chart_count, data_key
     FROM _snapshots WHERE id = ? AND book_id = ?`
  ).bind(snapshotId, bookId).first<{
    id: string; book_id: string; label: string | null;
    created_at: string; table_count: number; row_count: number;
    view_count: number; chart_count: number; data_key: string;
  }>();

  if (!row) return error('Snapshot not found', 404);

  const data = await readSnapshotData(context.env, row.data_key);
  if (!data) return error('Snapshot data not found in storage', 500);

  return json({
    id: row.id,
    bookId: row.book_id,
    label: row.label,
    createdAt: row.created_at,
    tableCount: row.table_count,
    rowCount: row.row_count,
    viewCount: row.view_count,
    chartCount: row.chart_count,
    data,
  });
};

// ---- DELETE /api/books/:bookId/snapshots/:snapshotId ----

export const onRequestDelete: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;

  // Get data_key before deleting
  const row = await context.env.DB.prepare(
    'SELECT data_key FROM _snapshots WHERE id = ? AND book_id = ?'
  ).bind(snapshotId, bookId).first<{ data_key: string }>();

  if (!row) return error('Snapshot not found', 404);

  // Delete from R2
  await deleteSnapshotData(context.env, row.data_key);

  // Delete from D1
  await context.env.DB.prepare(
    'DELETE FROM _snapshots WHERE id = ? AND book_id = ?'
  ).bind(snapshotId, bookId).run();

  return json({ ok: true });
};

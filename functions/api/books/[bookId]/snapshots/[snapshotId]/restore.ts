import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser } from '../../../../../lib';
import { readSnapshotData } from '../storage';
import { restoreBackupBook } from '../../../../../restoreBackup';

// ---- POST /api/books/:bookId/snapshots/:snapshotId/restore ----

export const onRequestPost: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  const user = requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;
  const db = context.env.DB;

  // Load the snapshot from R2
  const row = await db.prepare(
    'SELECT data_key FROM _snapshots WHERE id = ? AND book_id = ?'
  ).bind(snapshotId, bookId).first<{ data_key: string }>();

  if (!row) return error('Snapshot not found', 404);

  const payload = await readSnapshotData(context.env, row.data_key);
  if (!payload) return error('Snapshot data not found in storage', 500);

  const stats = await restoreBackupBook(db, user.id, payload, { bucket: context.env.BUCKET, sourceBookId: bookId });

  return json(stats, 201);
};

import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../lib';
import { removeRestoredTable, restoreBackupTable, validateBackupColumns } from '../../../../../restoreBackup';
import { findTable, replaceRestoredTable } from '../../../../../replaceRestoredTable';
import { readSnapshotData } from '../storage';

export const onRequestPost: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;
  const db = context.env.DB;
  const body = await context.request.json() as { tableName: string; replace?: boolean };
  const requestedName = body.tableName?.trim();
  if (!requestedName) return error('tableName is required');

  const row = await db.prepare(
    'SELECT data_key, created_at FROM _snapshots WHERE id = ? AND book_id = ?'
  ).bind(snapshotId, bookId).first<{ data_key: string; created_at: string }>();
  if (!row) return error('Snapshot not found', 404);

  const snapshot = await readSnapshotData(context.env, row.data_key);
  if (!snapshot) return error('Snapshot data not found in storage', 500);
  const source = snapshot.tables.find(table => table.name.toLowerCase() === requestedName.toLowerCase());
  if (!source) return error(`Table "${requestedName}" not found in snapshot`, 404);

  const validationError = validateBackupColumns(source.schema?.columns ?? []);
  if (validationError) return error(`${validationError} in snapshot table "${source.name}"`);

  const date = new Date(row.created_at).toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
  });
  const baseName = `${source.name} (restored from ${date})`;
  let name = baseName;
  for (let counter = 2; await db.prepare('SELECT id FROM _tables WHERE book_id = ? AND name = ?').bind(bookId, name).first(); counter++) {
    name = `${baseName} (${counter})`;
  }

  const original = body.replace ? await findTable(db, bookId, source.name) : null;
  const order = await db.prepare(
    'SELECT COALESCE(MAX(display_order), -1) + 1 AS next FROM _tables WHERE book_id = ?'
  ).bind(bookId).first<{ next: number }>();
  const displayOrder = order?.next ?? 0;
  const id = await restoreBackupTable(db, bookId, source, name, displayOrder, true);

  let replaced = false;
  try {
    if (original) {
      replaced = !!await replaceRestoredTable(db, bookId, original, { id, name, display_order: displayOrder });
    }
    if (!replaced) {
      await db.prepare('UPDATE _tables SET restore_pending = 0 WHERE book_id = ? AND id = ?')
        .bind(bookId, id).run();
    }
  } catch (cause) {
    await removeRestoredTable(db, id);
    throw cause;
  }

  if (original && !replaced) {
    return error(`Tables changed during replacement. The restored copy is available as "${name}".`, 409);
  }
  return json({ tableName: replaced ? source.name : name, rowCount: source.rows.length, replaced }, 201);
};

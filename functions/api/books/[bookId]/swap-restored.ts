import type { Env, RequestData } from '../../../lib';
import { json, error, requireUser, requireEditor } from '../../../lib';
import { findTable, replaceRestoredTable } from '../../../replaceRestoredTable';

export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const db = context.env.DB;
  const body = await context.request.json() as { restoredName: string };
  const restoredName = body.restoredName?.trim();
  if (!restoredName) return error('restoredName is required');

  const match = restoredName.match(/^(.+) \(restored from [^)]+\)(?: \(\d+\))?$/);
  if (!match) return error('Table name does not match restored pattern');
  const originalName = match[1];

  const [original, restored] = await Promise.all([
    findTable(db, bookId, originalName),
    findTable(db, bookId, restoredName),
  ]);
  if (!restored) return error('Restored table not found', 404);
  if (!original) return error(`Original table "${originalName}" not found`, 404);

  const oldName = await replaceRestoredTable(db, bookId, original, restored);
  if (!oldName) return error('Tables changed during replacement. Refresh and try again.', 409);
  return json({ originalName, newName: originalName, oldName });
};

import type { Env, RequestData } from '../../../lib';
import { json, error, requireUser, requireEditor } from '../../../lib';

// POST /api/books/:bookId/swap-restored
// Swaps a restored table ("X (restored from date)") with the original ("X").
// Updates all _columns.ref_table references to follow the swap.
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const db = context.env.DB;
  const body = await context.request.json() as { restoredName: string };
  if (!body.restoredName?.trim()) return error('restoredName is required');

  const restoredName = body.restoredName.trim();

  // Extract original name: "Orders (restored from Jun 19, 2026)" → "Orders"
  const match = restoredName.match(/^(.+?)\s*\(restored from .+\)$/);
  if (!match) return error('Table name does not match restored pattern', 400);
  const originalName = match[1].trim();

  const restored = await db.prepare(
    'SELECT id, display_order FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, restoredName).first<{ id: number; display_order: number }>();

  const original = await db.prepare(
    'SELECT id, display_order FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, originalName).first<{ id: number; display_order: number }>();

  if (!restored) return error('Restored table not found', 404);
  if (!original) return error(`Original table "${originalName}" not found`, 404);

  // Update all reference columns pointing to the original → point to the restored name
  await db.prepare(
    'UPDATE _columns SET ref_table = ? WHERE ref_table = ?'
  ).bind(restoredName, originalName).run();

  // Delete the original table (cascades to _columns and drops t_{id})
  await db.prepare('DELETE FROM _tables WHERE id = ?').bind(original.id).run();

  // Rename restored → original name, take over its display order
  await db.batch([
    db.prepare('UPDATE _tables SET display_order = ? WHERE id = ?').bind(original.display_order, restored.id),
    db.prepare('UPDATE _tables SET name = ? WHERE id = ?').bind(originalName, restored.id),
    // Update restored table's own ref_table columns to new name
    db.prepare('UPDATE _columns SET ref_table = ? WHERE ref_table = ?').bind(originalName, restoredName),
  ]);

  return json({ originalName, newName: originalName, deletedOriginal: original.id });
};

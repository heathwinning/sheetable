import type { Env, RequestData } from '../../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../../lib';

async function getTableId(db: D1Database, bookId: string, name: string): Promise<number | null> {
  const row = await db.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, name).first<{ id: number }>();
  return row?.id ?? null;
}

// PUT /api/books/:bookId/tables/:name/rows/:rowId → update a single cell or multiple columns
export const onRequestPut: PagesFunction<Env, 'bookId' | 'name' | 'rowId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);
  const rowId = context.params.rowId as string;

  const tableId = await getTableId(context.env.DB, bookId, name);
  if (tableId === null) return error('Table not found', 404);

  const body = await context.request.json() as Record<string, string>;

  // Get valid column names to prevent SQL injection
  const { results: cols } = await context.env.DB.prepare(
    'SELECT name FROM _columns WHERE table_id = ?'
  ).bind(tableId).all<{ name: string }>();
  const validCols = new Set(cols.map(c => c.name));

  const updates: string[] = [];
  const values: string[] = [];

  for (const [key, value] of Object.entries(body)) {
    if (validCols.has(key)) {
      updates.push(`"${key}" = ?`);
      values.push(value);
    }
  }

  if (updates.length === 0) return error('No valid columns to update');

  values.push(rowId);

  const result = await context.env.DB.prepare(
    `UPDATE t_${tableId} SET ${updates.join(', ')} WHERE _rowId = ?`
  ).bind(...values).run();

  if (result.meta.changes === 0) {
    return error('Row not found', 404);
  }

  return json({ ok: true });
};

// DELETE /api/books/:bookId/tables/:name/rows/:rowId → delete a row
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'name' | 'rowId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);
  const rowId = context.params.rowId as string;

  const tableId = await getTableId(context.env.DB, bookId, name);
  if (tableId === null) return error('Table not found', 404);

  // Check if any other table references this row
  const { results: refCols } = await context.env.DB.prepare(
    `SELECT c.table_id, c.name, t.name as tname
     FROM _columns c
     JOIN _tables t ON t.id = c.table_id
     WHERE c.ref_table = ? AND t.book_id = ?`
  ).bind(name, bookId).all<{ table_id: number; name: string; tname: string }>();

  for (const ref of refCols) {
    const refCheck = await context.env.DB.prepare(
      `SELECT 1 FROM t_${ref.table_id} WHERE "${ref.name}" = ? LIMIT 1`
    ).bind(rowId).first();
    if (refCheck) {
      return error(`Cannot delete: row is referenced by table "${ref.tname}" column "${ref.name}"`);
    }
  }

  const result = await context.env.DB.prepare(
    `DELETE FROM t_${tableId} WHERE _rowId = ?`
  ).bind(rowId).run();

  if (result.meta.changes === 0) {
    return error('Row not found', 404);
  }

  return json({ ok: true });
};

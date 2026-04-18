import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../lib';

// Helper to resolve table name → _tables.id
async function resolveTable(db: D1Database, bookId: string, name: string) {
  return db.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, name).first<{ id: number }>();
}

// PATCH /api/books/:bookId/tables/:name → rename table or reorder
export const onRequestPatch: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);
  const body = await context.request.json() as { name?: string; displayOrder?: number };

  const table = await resolveTable(context.env.DB, bookId, name);
  if (!table) return error('Table not found', 404);

  const stmts: D1PreparedStatement[] = [];

  if (body.name?.trim() && body.name.trim() !== name) {
    stmts.push(
      context.env.DB.prepare('UPDATE _tables SET name = ? WHERE id = ?')
        .bind(body.name.trim(), table.id)
    );
  }

  if (body.displayOrder !== undefined) {
    stmts.push(
      context.env.DB.prepare('UPDATE _tables SET display_order = ? WHERE id = ?')
        .bind(body.displayOrder, table.id)
    );
  }

  if (stmts.length > 0) {
    await context.env.DB.batch(stmts);
  }

  return json({ ok: true });
};

// DELETE /api/books/:bookId/tables/:name → drop table
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);

  const table = await resolveTable(context.env.DB, bookId, name);
  if (!table) return error('Table not found', 404);

  await context.env.DB.batch([
    context.env.DB.prepare(`DROP TABLE IF EXISTS t_${table.id}`),
    context.env.DB.prepare('DELETE FROM _columns WHERE table_id = ?').bind(table.id),
    context.env.DB.prepare('DELETE FROM _tables WHERE id = ?').bind(table.id),
  ]);

  return json({ ok: true });
};

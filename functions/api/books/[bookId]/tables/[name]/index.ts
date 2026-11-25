import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../lib';

// Helper to resolve table name → _tables.id
async function resolveTable(db: D1Database, bookId: string, name: string) {
  return db.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ? AND restore_pending = 0'
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

  const newName = body.name?.trim();
  if (newName && newName !== name) {
    stmts.push(
      context.env.DB.prepare('UPDATE _tables SET name = ? WHERE id = ?').bind(newName, table.id),
      context.env.DB.prepare(
        `UPDATE _report_sheets AS report
         SET definition = json_set(definition, '$.data.content', (
           SELECT json_group_array(json(CASE
             WHEN json_extract(block.value, '$.type') = 'Table' AND json_extract(block.value, '$.props.table') = ?
             THEN json_set(block.value, '$.props.table', ?)
             ELSE block.value END))
           FROM json_each(report.definition, '$.data.content') AS block
         )), revision = revision + 1
         WHERE report.book_id = ? AND EXISTS (
           SELECT 1 FROM json_each(report.definition, '$.data.content') AS block
           WHERE json_extract(block.value, '$.type') = 'Table' AND json_extract(block.value, '$.props.table') = ?
         )`
      ).bind(name, newName, bookId, name),
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

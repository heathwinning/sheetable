import type { Env, RequestData } from '../../../lib';
import { json, error, requireUser, requireOwner, requireEditor } from '../../../lib';

// PATCH /api/books/:bookId → rename book or reorder tables/charts
export const onRequestPatch: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as {
    name?: string;
    tableOrder?: string[];
    chartOrder?: string[];
    sheetOrder?: { type: 'table' | 'chart' | 'view'; name: string; hidden?: boolean }[];
  };

  // Renaming requires owner
  if (body.name !== undefined) {
    requireOwner(context.data);
    if (!body.name.trim()) return error('name is required');
    await context.env.DB.prepare(
      'UPDATE books SET name = ? WHERE id = ?'
    ).bind(body.name.trim(), bookId).run();
  }

  // Reordering tables requires editor
  if (body.tableOrder) {
    requireEditor(context.data);
    const stmts = body.tableOrder.map((name, i) =>
      context.env.DB.prepare(
        'UPDATE _tables SET display_order = ? WHERE book_id = ? AND name = ?'
      ).bind(i, bookId, name)
    );
    if (stmts.length > 0) await context.env.DB.batch(stmts);
  }

  // Reordering charts requires editor
  if (body.chartOrder) {
    requireEditor(context.data);
    const stmts = body.chartOrder.map((name, i) =>
      context.env.DB.prepare(
        'UPDATE _chart_sheets SET display_order = ? WHERE book_id = ? AND name = ?'
      ).bind(i, bookId, name)
    );
    if (stmts.length > 0) await context.env.DB.batch(stmts);
  }

  // Global cross-type sheet order
  if (body.sheetOrder) {
    requireEditor(context.data);
    const order = body.sheetOrder;
    const stmts: ReturnType<typeof context.env.DB.prepare>[] = [];
    // Save JSON to books record
    stmts.push(
      context.env.DB.prepare('UPDATE books SET sheet_order = ? WHERE id = ?')
        .bind(JSON.stringify(order), bookId)
    );
    // Also update per-type display_order with global position for consistent per-type sorting
    order.forEach((item, globalIdx) => {
      if (item.type === 'table') {
        stmts.push(context.env.DB.prepare(
          'UPDATE _tables SET display_order = ? WHERE book_id = ? AND name = ?'
        ).bind(globalIdx, bookId, item.name));
      } else if (item.type === 'chart') {
        stmts.push(context.env.DB.prepare(
          'UPDATE _chart_sheets SET display_order = ? WHERE book_id = ? AND name = ?'
        ).bind(globalIdx, bookId, item.name));
      } else if (item.type === 'view') {
        stmts.push(context.env.DB.prepare(
          'UPDATE _view_sheets SET display_order = ? WHERE book_id = ? AND name = ?'
        ).bind(globalIdx, bookId, item.name));
      }
    });
    if (stmts.length > 0) await context.env.DB.batch(stmts);
  }

  return json({ ok: true });
};

// DELETE /api/books/:bookId → delete book and all its data
export const onRequestDelete: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireOwner(context.data);

  const bookId = context.params.bookId as string;

  // Get all physical table IDs to drop
  const { results: tables } = await context.env.DB.prepare(
    'SELECT id FROM _tables WHERE book_id = ?'
  ).bind(bookId).all<{ id: number }>();

  const stmts: D1PreparedStatement[] = [];

  // Drop each physical table
  for (const t of tables) {
    stmts.push(context.env.DB.prepare(`DROP TABLE IF EXISTS t_${t.id}`));
  }

  // Delete metadata (cascades handle _columns, book_members, _chart_sheets)
  stmts.push(
    context.env.DB.prepare('DELETE FROM _tables WHERE book_id = ?').bind(bookId),
    context.env.DB.prepare('DELETE FROM _chart_sheets WHERE book_id = ?').bind(bookId),
    context.env.DB.prepare('DELETE FROM book_members WHERE book_id = ?').bind(bookId),
    context.env.DB.prepare('DELETE FROM books WHERE id = ?').bind(bookId),
  );

  await context.env.DB.batch(stmts);

  // Clean up R2 objects for this book (best effort)
  const listed = await context.env.BUCKET.list({ prefix: `${bookId}/` });
  for (const obj of listed.objects) {
    await context.env.BUCKET.delete(obj.key);
  }

  return json({ ok: true });
};

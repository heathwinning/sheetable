import type { Env, RequestData } from '../../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../../lib';

// Helper to get physical table ID
async function getTableId(db: D1Database, bookId: string, name: string): Promise<number | null> {
  const row = await db.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, name).first<{ id: number }>();
  return row?.id ?? null;
}

// GET /api/books/:bookId/tables/:name/rows → get all rows
export const onRequestGet: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);

  const tableId = await getTableId(context.env.DB, bookId, name);
  if (tableId === null) return error('Table not found', 404);

  const { results } = await context.env.DB.prepare(
    `SELECT * FROM t_${tableId}`
  ).all();

  // Stringify _rowId (INTEGER PRIMARY KEY returns as number from D1)
  return json(results.map(r => ({ ...r, _rowId: String(r._rowId) })));
};

// POST /api/books/:bookId/tables/:name/rows → insert a new row
export const onRequestPost: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);

  const tableId = await getTableId(context.env.DB, bookId, name);
  if (tableId === null) return error('Table not found', 404);

  const body = await context.request.json() as Record<string, string>;

  // Get column names from schema
  const { results: cols } = await context.env.DB.prepare(
    'SELECT name FROM _columns WHERE table_id = ? ORDER BY display_order'
  ).bind(tableId).all<{ name: string }>();

  const colNames = cols.map(c => c.name);

  if (body._rowId) {
    // Client-provided integer _rowId
    const allCols = ['_rowId', ...colNames];
    const placeholders = allCols.map(() => '?').join(', ');
    const values = allCols.map(c => body[c] ?? '');
    await context.env.DB.prepare(
      `INSERT INTO t_${tableId} (${allCols.map(c => `"${c}"`).join(', ')}) VALUES (${placeholders})`
    ).bind(...values).run();
    return json({ ok: true, rowId: Number(body._rowId) }, 201);
  } else {
    // Auto-generate _rowId
    const placeholders = colNames.map(() => '?').join(', ');
    const values = colNames.map(c => body[c] ?? '');
    const result = await context.env.DB.prepare(
      `INSERT INTO t_${tableId} (${colNames.map(c => `"${c}"`).join(', ')}) VALUES (${placeholders})`
    ).bind(...values).run();
    return json({ ok: true, rowId: result.meta.last_row_id }, 201);
  }
};

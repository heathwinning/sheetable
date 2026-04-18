import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// GET /api/books/:bookId/charts → list chart sheets
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  const bookId = context.params.bookId as string;

  const { results } = await context.env.DB.prepare(
    'SELECT id, name, table_name, mode, charts, display_order FROM _chart_sheets WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all();

  return json(results.map(r => ({
    name: r.name,
    tableName: r.table_name,
    mode: r.mode,
    charts: JSON.parse(r.charts as string),
    displayOrder: r.display_order,
  })));
};

// POST /api/books/:bookId/charts → create chart sheet
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as { name?: string; tableName?: string };

  if (!body.name?.trim()) return error('name is required');

  const maxOrder = await context.env.DB.prepare(
    'SELECT COALESCE(MAX(display_order), -1) as m FROM _chart_sheets WHERE book_id = ?'
  ).bind(bookId).first<{ m: number }>();

  await context.env.DB.prepare(
    `INSERT INTO _chart_sheets (book_id, name, table_name, display_order)
     VALUES (?, ?, ?, ?)`
  ).bind(bookId, body.name.trim(), body.tableName || null, (maxOrder?.m ?? -1) + 1).run();

  return json({ name: body.name.trim() }, 201);
};

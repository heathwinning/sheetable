import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// GET /api/books/:bookId/views → list view sheets
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  const bookId = context.params.bookId as string;

  const { results } = await context.env.DB.prepare(
    'SELECT name, table_name, view_type, date_column, display_order FROM _view_sheets WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all();

  return json(results.map(r => ({
    name: r.name,
    tableName: r.table_name,
    viewType: r.view_type,
    dateColumn: r.date_column ?? undefined,
    displayOrder: r.display_order,
  })));
};

// POST /api/books/:bookId/views → create view sheet
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as {
    name?: string;
    tableName?: string;
    viewType?: string;
    dateColumn?: string;
  };

  if (!body.name?.trim()) return error('name is required');
  if (!body.tableName?.trim()) return error('tableName is required');

  const viewType = body.viewType ?? 'calendar';
  if (!['grid', 'calendar', 'schedule'].includes(viewType)) return error('invalid viewType');

  const maxOrder = await context.env.DB.prepare(
    'SELECT COALESCE(MAX(display_order), -1) as m FROM _view_sheets WHERE book_id = ?'
  ).bind(bookId).first<{ m: number }>();

  await context.env.DB.prepare(
    `INSERT INTO _view_sheets (book_id, name, table_name, view_type, date_column, display_order)
     VALUES (?, ?, ?, ?, ?, ?)`
  ).bind(
    bookId,
    body.name.trim(),
    body.tableName.trim(),
    viewType,
    body.dateColumn ?? null,
    (maxOrder?.m ?? -1) + 1,
  ).run();

  return json({ name: body.name.trim() }, 201);
};

import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// PATCH /api/books/:bookId/views/:name → update view sheet
export const onRequestPatch: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const viewName = decodeURIComponent(context.params.name as string);

  const body = await context.request.json() as {
    name?: string;
    tableName?: string;
    viewType?: string;
    dateColumn?: string | null;
    displayOrder?: number;
  };

  const sets: string[] = [];
  const values: unknown[] = [];

  if (body.name !== undefined) { sets.push('name = ?'); values.push(body.name); }
  if (body.tableName !== undefined) { sets.push('table_name = ?'); values.push(body.tableName); }
  if (body.viewType !== undefined) {
    if (!['grid', 'calendar', 'schedule'].includes(body.viewType)) return error('invalid viewType');
    sets.push('view_type = ?'); values.push(body.viewType);
  }
  if ('dateColumn' in body) { sets.push('date_column = ?'); values.push(body.dateColumn ?? null); }
  if (body.displayOrder !== undefined) { sets.push('display_order = ?'); values.push(body.displayOrder); }

  if (sets.length === 0) return error('No fields to update');

  values.push(bookId, viewName);
  await context.env.DB.prepare(
    `UPDATE _view_sheets SET ${sets.join(', ')} WHERE book_id = ? AND name = ?`
  ).bind(...values).run();

  return json({ ok: true });
};

// DELETE /api/books/:bookId/views/:name → delete view sheet
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const viewName = decodeURIComponent(context.params.name as string);

  await context.env.DB.prepare(
    'DELETE FROM _view_sheets WHERE book_id = ? AND name = ?'
  ).bind(bookId, viewName).run();

  return json({ ok: true });
};

import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// PATCH /api/books/:bookId/charts/:name → update chart sheet
export const onRequestPatch: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const chartName = decodeURIComponent(context.params.name as string);

  const body = await context.request.json() as {
    name?: string;
    tableName?: string;
    mode?: string;
    charts?: unknown[];
    displayOrder?: number;
  };

  const sets: string[] = [];
  const values: unknown[] = [];

  if (body.name !== undefined) { sets.push('name = ?'); values.push(body.name); }
  if (body.tableName !== undefined) { sets.push('table_name = ?'); values.push(body.tableName); }
  if (body.mode !== undefined) { sets.push('mode = ?'); values.push(body.mode); }
  if (body.charts !== undefined) { sets.push('charts = ?'); values.push(JSON.stringify(body.charts)); }
  if (body.displayOrder !== undefined) { sets.push('display_order = ?'); values.push(body.displayOrder); }

  if (sets.length === 0) return error('No fields to update');

  values.push(bookId, chartName);
  await context.env.DB.prepare(
    `UPDATE _chart_sheets SET ${sets.join(', ')} WHERE book_id = ? AND name = ?`
  ).bind(...values).run();

  return json({ ok: true });
};

// DELETE /api/books/:bookId/charts/:name → delete chart sheet
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const chartName = decodeURIComponent(context.params.name as string);

  await context.env.DB.prepare(
    'DELETE FROM _chart_sheets WHERE book_id = ? AND name = ?'
  ).bind(bookId, chartName).run();

  return json({ ok: true });
};

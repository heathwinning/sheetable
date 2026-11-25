import type { Env, RequestData } from '../../../../lib';
import { error, json, requireEditor, requireUser } from '../../../../lib';
import { isReportDefinition, reportName } from '../../../../../shared/reportTypes';
import type { ReportSummary } from '../../../../../shared/reportTypes';

type ReportRequest = PagesFunction<Env, 'bookId', RequestData>;

export const onRequestGet: ReportRequest = async ({ env, params, data }) => {
  requireUser(data);
  const { results } = await env.DB.prepare(
    'SELECT id, name FROM _report_sheets WHERE book_id = ? ORDER BY display_order, id'
  ).bind(params.bookId).all<ReportSummary>();
  return json(results);
};

export const onRequestPost: ReportRequest = async ({ env, params, data, request }) => {
  requireEditor(data);
  const body = await request.json<{ name?: unknown; definition?: unknown }>().catch(() => null);
  const name = reportName(body?.name);
  if (!name || !isReportDefinition(body?.definition)) return error('Invalid report name or layout', 400);

  const id = crypto.randomUUID();
  const result = await env.DB.prepare(
    `INSERT INTO _report_sheets (id, book_id, name, definition, display_order)
     VALUES (?, ?, ?, ?, (SELECT COALESCE(MAX(display_order), -1) + 1 FROM _report_sheets WHERE book_id = ?))
     ON CONFLICT (book_id, name) DO NOTHING`
  ).bind(id, params.bookId, name, JSON.stringify(body.definition), params.bookId).run();
  if (!result.meta.changes) return error('A report with this name already exists', 409);
  return json({ id, name, definition: body.definition, revision: 1 }, 201);
};

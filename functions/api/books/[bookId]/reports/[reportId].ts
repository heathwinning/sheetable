import type { Env, RequestData } from '../../../../lib';
import { error, json, requireEditor, requireUser } from '../../../../lib';
import { isReportDefinition, reportName } from '../../../../../shared/reportTypes';

type ReportRequest = PagesFunction<Env, 'bookId' | 'reportId', RequestData>;
type ReportRow = { id: string; name: string; definition: string; revision: number };

export const onRequestGet: ReportRequest = async ({ env, params, data }) => {
  requireUser(data);
  const row = await env.DB.prepare(
    'SELECT id, name, definition, revision FROM _report_sheets WHERE book_id = ? AND id = ?'
  ).bind(params.bookId, params.reportId).first<ReportRow>();
  return row ? json({ ...row, definition: JSON.parse(row.definition) }) : error('Report not found', 404);
};

export const onRequestPatch: ReportRequest = async ({ env, params, data, request }) => {
  requireEditor(data);
  const body = await request.json<{ name?: unknown; definition?: unknown; revision?: number }>().catch(() => null);
  const name = reportName(body?.name);
  const revision = body?.revision;
  if (!name || !isReportDefinition(body?.definition) || typeof revision !== 'number' || !Number.isSafeInteger(revision) || revision < 1) {
    return error('Invalid report name or layout', 400);
  }

  const row = await env.DB.prepare(
    `UPDATE _report_sheets SET name = ?, definition = ?, revision = revision + 1
      WHERE book_id = ? AND id = ? AND revision = ?
      AND NOT EXISTS (SELECT 1 FROM _report_sheets WHERE book_id = ? AND name = ? AND id != ?)
      RETURNING revision`
  ).bind(name, JSON.stringify(body.definition), params.bookId, params.reportId, revision,
    params.bookId, name, params.reportId).first<{ revision: number }>();
  if (!row) {
    const current = await env.DB.prepare('SELECT revision FROM _report_sheets WHERE book_id = ? AND id = ?')
      .bind(params.bookId, params.reportId).first<{ revision: number }>();
    if (!current) return error('Report not found', 404);
    return error(current.revision === revision
      ? 'A report with this name already exists'
      : 'This report has changed elsewhere. Reopen it or save your edits as a copy.', 409);
  }
  return json({ id: params.reportId, name, definition: body.definition, revision: row.revision });
};

export const onRequestDelete: ReportRequest = async ({ env, params, data }) => {
  requireEditor(data);
  const result = await env.DB.prepare('DELETE FROM _report_sheets WHERE book_id = ? AND id = ?')
    .bind(params.bookId, params.reportId).run();
  return result.meta.changes ? json({ ok: true }) : error('Report not found', 404);
};

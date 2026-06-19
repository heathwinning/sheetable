import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';
import { gatherExportPayload } from '../export';
import { createSnapshot } from './storage';

// ---- Types ----

interface SnapshotMeta {
  id: string;
  bookId: string;
  label: string | null;
  createdAt: string;
  tableCount: number;
  rowCount: number;
  viewCount: number;
  chartCount: number;
}

// ---- GET /api/books/:bookId/snapshots ----

export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;

  const { results } = await context.env.DB.prepare(
    `SELECT id, book_id, label, created_at, table_count, row_count, view_count, chart_count
     FROM _snapshots WHERE book_id = ? ORDER BY created_at DESC`
  ).bind(bookId).all<{
    id: string; book_id: string; label: string | null;
    created_at: string; table_count: number; row_count: number;
    view_count: number; chart_count: number;
  }>();

  const snapshots: SnapshotMeta[] = results.map(r => ({
    id: r.id,
    bookId: r.book_id,
    label: r.label,
    createdAt: r.created_at,
    tableCount: r.table_count,
    rowCount: r.row_count,
    viewCount: r.view_count,
    chartCount: r.chart_count,
  }));

  return json(snapshots);
};

// ---- POST /api/books/:bookId/snapshots ----

export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json().catch(() => ({})) as { label?: string };
  const label = body.label?.trim() || null;

  const payload = await gatherExportPayload(context.env, bookId);
  const totalRows = payload.tables.reduce((sum, t) => sum + t.rows.length, 0);

  // Write to R2 + insert metadata + run retention cleanup
  const result = await createSnapshot(context.env, bookId, payload, label);

  return json({
    id: result.snapshotId,
    label,
    createdAt: new Date().toISOString(),
    tableCount: payload.tables.length,
    rowCount: totalRows,
    viewCount: payload.views.length,
    chartCount: payload.charts.length,
  }, 201);
};

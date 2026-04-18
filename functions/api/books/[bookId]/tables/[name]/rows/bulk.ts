import type { Env, RequestData } from '../../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../../lib';

interface BulkOp {
  type: 'insert' | 'update' | 'delete';
  rowId: string;
  data?: Record<string, string>; // for insert/update
}

async function getTableId(db: D1Database, bookId: string, name: string): Promise<number | null> {
  const row = await db.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, name).first<{ id: number }>();
  return row?.id ?? null;
}

// POST /api/books/:bookId/tables/:name/rows/bulk → bulk insert/update/delete
export const onRequestPost: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);

  const tableId = await getTableId(context.env.DB, bookId, name);
  if (tableId === null) return error('Table not found', 404);

  const body = await context.request.json() as { operations: BulkOp[] };
  if (!body.operations?.length) return error('operations array is required');

  // Get valid column names
  const { results: cols } = await context.env.DB.prepare(
    'SELECT name FROM _columns WHERE table_id = ? ORDER BY display_order'
  ).bind(tableId).all<{ name: string }>();
  const validCols = new Set(cols.map(c => c.name));
  const colNames = cols.map(c => c.name);

  const stmts: D1PreparedStatement[] = [];

  for (const op of body.operations) {
    if (!op.rowId) continue;

    if (op.type === 'delete') {
      stmts.push(
        context.env.DB.prepare(`DELETE FROM t_${tableId} WHERE _rowId = ?`).bind(op.rowId)
      );
    } else if (op.type === 'insert' && op.data) {
      const allCols = ['_rowId', ...colNames];
      const placeholders = allCols.map(() => '?').join(', ');
      const values = allCols.map(c => c === '_rowId' ? op.rowId : (op.data![c] ?? ''));
      stmts.push(
        context.env.DB.prepare(
          `INSERT INTO t_${tableId} (${allCols.map(c => `"${c}"`).join(', ')}) VALUES (${placeholders})`
        ).bind(...values)
      );
    } else if (op.type === 'update' && op.data) {
      const updates: string[] = [];
      const values: string[] = [];
      for (const [key, value] of Object.entries(op.data)) {
        if (validCols.has(key)) {
          updates.push(`"${key}" = ?`);
          values.push(value);
        }
      }
      if (updates.length > 0) {
        values.push(op.rowId);
        stmts.push(
          context.env.DB.prepare(
            `UPDATE t_${tableId} SET ${updates.join(', ')} WHERE _rowId = ?`
          ).bind(...values)
        );
      }
    }
  }

  if (stmts.length > 0) {
    // D1 batch limit: chunk into batches of 50 statements
    const BATCH_SIZE = 50;
    for (let i = 0; i < stmts.length; i += BATCH_SIZE) {
      await context.env.DB.batch(stmts.slice(i, i + BATCH_SIZE));
    }
  }

  return json({ ok: true, processed: stmts.length });
};

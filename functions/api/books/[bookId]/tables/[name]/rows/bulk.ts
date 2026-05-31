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

  // Get valid column names (exclude calculated columns — they have no backing SQL column)
  const { results: cols } = await context.env.DB.prepare(
    "SELECT name, type FROM _columns WHERE table_id = ? AND type != 'calculated' ORDER BY display_order"
  ).bind(tableId).all<{ name: string; type: string }>();
  const validCols = new Set(cols.map(c => c.name));
  const colNames = cols.map(c => c.name);

  const stmts: D1PreparedStatement[] = [];

  // Separate inserts for multi-row batching; handle updates/deletes individually
  const insertOps = body.operations.filter(op => op.type === 'insert' && op.rowId && op.data);
  const otherOps  = body.operations.filter(op => op.type !== 'insert' && op.rowId);

  for (const op of otherOps) {
    if (op.type === 'delete') {
      stmts.push(
        context.env.DB.prepare(`DELETE FROM t_${tableId} WHERE _rowId = ?`).bind(op.rowId)
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

  // Multi-row INSERTs: D1 allows max 100 bound params per statement.
  // Pack floor(100 / numCols) rows per INSERT to maximise throughput.
  if (insertOps.length > 0) {
    const allCols = ['_rowId', ...colNames];
    const rowsPerStmt = Math.max(1, Math.floor(100 / allCols.length));
    const colList = allCols.map(c => `"${c}"`).join(', ');
    const rowPlaceholder = `(${allCols.map(() => '?').join(', ')})`;
    for (let i = 0; i < insertOps.length; i += rowsPerStmt) {
      const chunk = insertOps.slice(i, i + rowsPerStmt);
      const values = chunk.flatMap(op => allCols.map(c => c === '_rowId' ? op.rowId : (op.data![c] ?? '')));
      stmts.push(
        context.env.DB.prepare(
          `INSERT INTO t_${tableId} (${colList}) VALUES ${chunk.map(() => rowPlaceholder).join(', ')}`
        ).bind(...values)
      );
    }
  }

  if (stmts.length > 0) {
    // D1 batch limit: 100 statements per batch
    const BATCH_SIZE = 100;
    for (let i = 0; i < stmts.length; i += BATCH_SIZE) {
      await context.env.DB.batch(stmts.slice(i, i + BATCH_SIZE));
    }
  }

  return json({ ok: true, processed: stmts.length });
};

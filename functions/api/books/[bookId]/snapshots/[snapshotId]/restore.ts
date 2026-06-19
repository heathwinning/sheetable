import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser } from '../../../../../lib';
import { readSnapshotData } from '../storage';

// ---- Types matching the snapshot/export format ----

interface ImportColumn {
  name: string;
  displayName?: string;
  type: string;
  width?: number;
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
  expression?: string;
  showInGrid?: boolean;
  listOf?: string;
}

interface ImportTableSchema {
  columns: ImportColumn[];
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: string }[];
  draftRowPosition?: string;
}

interface ImportRow {
  _rowId: string;
  _resolved?: Record<string, string>;
  [column: string]: string | Record<string, string> | undefined;
}

interface ImportTable {
  name: string;
  schema: ImportTableSchema | null;
  rows: ImportRow[];
}

interface ImportView {
  name: string;
  tableName: string;
  viewType: string;
  dateColumn?: string;
}

interface ImportChart {
  name: string;
  tableName?: string;
  mode?: string;
  charts?: unknown;
}

interface ImportPayload {
  version: number;
  exportedAt?: string;
  bookName: string;
  tables: ImportTable[];
  views?: ImportView[];
  charts?: ImportChart[];
}

const _VALID_TYPES = new Set(['text', 'integer', 'decimal', 'date', 'datetime', 'bool', 'reference', 'image', 'calculated', 'list']);

async function restoreFromPayload(
  db: D1Database,
  userId: string,
  body: ImportPayload,
): Promise<{ bookId: string; bookName: string; tableCount: number; rowCount: number; viewCount: number; chartCount: number }> {
  const bookId = crypto.randomUUID();

  // 1. Create the book
  await db.batch([
    db.prepare('INSERT INTO books (id, name, owner_id) VALUES (?, ?, ?)')
      .bind(bookId, body.bookName.trim(), userId),
    db.prepare('INSERT INTO book_members (book_id, user_id, role) VALUES (?, ?, ?)')
      .bind(bookId, userId, 'owner'),
  ]);

  // 2. Create tables and insert rows
  for (let ti = 0; ti < body.tables.length; ti++) {
    const table = body.tables[ti];
    const cols = table.schema?.columns ?? [];

    let columns: ImportColumn[];
    if (cols.length > 0) {
      columns = cols;
    } else {
      const firstRow = table.rows[0];
      if (!firstRow) {
        columns = [{ name: 'name', type: 'text' }];
      } else {
        columns = Object.keys(firstRow)
          .filter(k => k !== '_rowId' && k !== '_resolved')
          .map(k => ({ name: k, type: 'text' }));
      }
    }

    const insertResult = await db.prepare(
      `INSERT INTO _tables (book_id, name, display_order, unique_keys, default_sort, draft_position)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).bind(
      bookId,
      table.name.trim(),
      ti,
      JSON.stringify(table.schema?.uniqueKeys ?? []),
      table.schema?.defaultSort ? JSON.stringify(table.schema.defaultSort) : null,
      table.schema?.draftRowPosition ?? 'bottom',
    ).run();

    const tableId = insertResult.meta.last_row_id;

    const colStmts: D1PreparedStatement[] = columns.map((col, ci) =>
      db.prepare(
        `INSERT INTO _columns (table_id, name, display_name, type, display_order, width, ref_table, ref_display, ref_search, expression, show_in_grid, list_of)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        tableId,
        col.name.trim(),
        col.displayName || null,
        col.type,
        ci,
        col.width ?? null,
        col.refTable || null,
        col.refDisplayColumns ? JSON.stringify(col.refDisplayColumns) : null,
        col.refSearchColumns ? JSON.stringify(col.refSearchColumns) : null,
        col.expression || null,
        col.showInGrid ? 1 : 0,
        col.listOf || null,
      )
    );

    const physicalCols = columns.filter(c => c.type !== 'calculated');
    const colDefs = physicalCols.map(c => `"${c.name.trim()}" TEXT NOT NULL DEFAULT ''`).join(', ');
    const createSql = `CREATE TABLE t_${tableId} (_rowId TEXT PRIMARY KEY, ${colDefs})`;

    await db.batch([...colStmts, db.prepare(createSql)]);

    if (table.rows.length > 0) {
      const physicalColNames = physicalCols.map(c => c.name.trim());
      const allCols = ['_rowId', ...physicalColNames];
      const rowsPerStmt = Math.max(1, Math.floor(100 / allCols.length));
      const colList = allCols.map(c => `"${c}"`).join(', ');
      const rowPlaceholder = `(${allCols.map(() => '?').join(', ')})`;

      const cleanRows = table.rows.map(r => {
        const { _resolved, ...rest } = r as ImportRow;
        return rest as Record<string, string>;
      });

      const rowStmts: D1PreparedStatement[] = [];
      for (let i = 0; i < cleanRows.length; i += rowsPerStmt) {
        const chunk = cleanRows.slice(i, i + rowsPerStmt);
        const values = chunk.flatMap(row =>
          allCols.map(c => c === '_rowId' ? (row._rowId ?? '') : (row[c] ?? ''))
        );
        rowStmts.push(
          db.prepare(
            `INSERT INTO t_${tableId} (${colList}) VALUES ${chunk.map(() => rowPlaceholder).join(', ')}`
          ).bind(...values)
        );
      }

      const BATCH_SIZE = 100;
      for (let i = 0; i < rowStmts.length; i += BATCH_SIZE) {
        await db.batch(rowStmts.slice(i, i + BATCH_SIZE));
      }
    }
  }

  // 3. Create views
  const views = body.views ?? [];
  for (let vi = 0; vi < views.length; vi++) {
    const v = views[vi];
    await db.prepare(
      `INSERT INTO _view_sheets (book_id, name, table_name, view_type, date_column, display_order)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).bind(bookId, v.name, v.tableName, v.viewType, v.dateColumn || null, vi).run();
  }

  // 4. Create charts
  const charts = body.charts ?? [];
  for (let ci = 0; ci < charts.length; ci++) {
    const c = charts[ci];
    await db.prepare(
      `INSERT INTO _chart_sheets (book_id, name, table_name, mode, charts, display_order)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).bind(
      bookId,
      c.name,
      c.tableName || null,
      c.mode || null,
      c.charts ? JSON.stringify(c.charts) : null,
      ci,
    ).run();
  }

  return {
    bookId,
    bookName: body.bookName.trim(),
    tableCount: body.tables.length,
    rowCount: body.tables.reduce((sum, t) => sum + t.rows.length, 0),
    viewCount: views.length,
    chartCount: charts.length,
  };
}

// ---- POST /api/books/:bookId/snapshots/:snapshotId/restore ----

export const onRequestPost: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  const user = requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;
  const db = context.env.DB;

  // Load the snapshot from R2
  const row = await db.prepare(
    'SELECT data_key FROM _snapshots WHERE id = ? AND book_id = ?'
  ).bind(snapshotId, bookId).first<{ data_key: string }>();

  if (!row) return error('Snapshot not found', 404);

  const payload = await readSnapshotData(context.env, row.data_key);
  if (!payload) return error('Snapshot data not found in storage', 500);

  const stats = await restoreFromPayload(db, user.id, payload);

  return json(stats, 201);
};

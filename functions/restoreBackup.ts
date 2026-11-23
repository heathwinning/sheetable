import type { BackupColumn, BackupPayload, BackupTable } from '../shared/backupTypes';

const VALID_TYPES = new Set(['text', 'integer', 'decimal', 'date', 'datetime', 'bool', 'reference', 'image', 'calculated', 'list']);

export function validateBackupColumns(columns: BackupColumn[]): string | undefined {
  for (const col of columns) {
    const name = col.name?.trim();
    if (!name || name.length > 64 || name.includes('"') || name.toLowerCase() === '_rowid' || name === '_resolved') {
      return 'Invalid column name';
    }
    if (!VALID_TYPES.has(col.type)) return `Invalid column type "${col.type}"`;
  }
}

function quoteColumn(name: string): string {
  return `"${name}"`;
}

export async function removeRestoredTable(db: D1Database, tableId: number) {
  await db.batch([
    db.prepare(`DROP TABLE IF EXISTS t_${tableId}`),
    db.prepare('DELETE FROM _tables WHERE id = ?').bind(tableId),
  ]);
}

export async function restoreBackupTable(
  db: D1Database, bookId: string, table: BackupTable, name: string, order: number, pending = false,
): Promise<number> {
  const columns: BackupColumn[] = table.schema?.columns.length ? table.schema.columns : (
    table.rows[0]
      ? Object.keys(table.rows[0]).filter(key => key !== '_rowId' && key !== '_resolved').map(key => ({ name: key, type: 'text' }))
      : [{ name: 'name', type: 'text' }]
  );
  const validationError = validateBackupColumns(columns);
  if (validationError) throw new Error(`${validationError} in table "${name}"`);

  const { meta } = await db.prepare(
    `INSERT INTO _tables (book_id, name, display_order, unique_keys, default_sort, draft_position, restore_pending)
     VALUES (?, ?, ?, ?, ?, ?, ?)`
  ).bind(
    bookId, name, order, JSON.stringify(table.schema?.uniqueKeys ?? []),
    table.schema?.defaultSort ? JSON.stringify(table.schema.defaultSort) : null,
    table.schema?.draftRowPosition ?? 'bottom', pending ? 1 : 0,
  ).run();

  const tableId = meta.last_row_id;
  try {
    const columnStatements = columns.map((col, index) =>
      db.prepare(
        `INSERT INTO _columns (table_id, name, display_name, type, display_order, width, ref_table,
         ref_display, ref_search, expression, show_in_grid, list_of, default_value)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        tableId, col.name.trim(), col.displayName || null, col.type, index, col.width ?? null,
        col.refTable || null, col.refDisplayColumns ? JSON.stringify(col.refDisplayColumns) : null,
        col.refSearchColumns ? JSON.stringify(col.refSearchColumns) : null, col.expression || null,
        col.showInGrid ? 1 : 0, col.listOf || null, col.defaultValue ?? null,
      ),
    );
    const physicalColumns = columns.filter(col => col.type !== 'calculated').map(col => col.name.trim());
    const definitions = ['_rowId TEXT PRIMARY KEY', ...physicalColumns.map(col => `${quoteColumn(col)} TEXT NOT NULL DEFAULT ''`)];
    await db.batch([...columnStatements, db.prepare(`CREATE TABLE t_${tableId} (${definitions.join(', ')})`)]);

    const names = ['_rowId', ...physicalColumns];
    const columnList = names.map(quoteColumn).join(', ');
    const placeholders = `(${names.map(() => '?').join(', ')})`;
    const rowsPerStatement = Math.max(1, Math.floor(100 / names.length));
    const rowsPerBatch = rowsPerStatement * 100;
    for (let offset = 0; offset < table.rows.length; offset += rowsPerBatch) {
      const statements: D1PreparedStatement[] = [];
      const end = Math.min(offset + rowsPerBatch, table.rows.length);
      for (let i = offset; i < end; i += rowsPerStatement) {
        const rows = table.rows.slice(i, i + rowsPerStatement);
        const values = rows.flatMap(row => names.map(col => {
          const value = row[col];
          return typeof value === 'string' || typeof value === 'number' ? String(value) : '';
        }));
        statements.push(db.prepare(
          `INSERT INTO t_${tableId} (${columnList}) VALUES ${rows.map(() => placeholders).join(', ')}`
        ).bind(...values));
      }
      await db.batch(statements);
    }
  } catch (error) {
    await removeRestoredTable(db, tableId);
    throw error;
  }
  return tableId;
}

export async function restoreBackupBook(db: D1Database, userId: string, backup: BackupPayload) {
  const bookId = crypto.randomUUID();
  const bookName = backup.bookName.trim();
  await db.batch([
    db.prepare('INSERT INTO books (id, name, owner_id, sheet_order) VALUES (?, ?, ?, ?)')
      .bind(bookId, bookName, userId, backup.sheetOrder ?? ''),
    db.prepare('INSERT INTO book_members (book_id, user_id, role) VALUES (?, ?, ?)')
      .bind(bookId, userId, 'owner'),
  ]);

  const tableIds: number[] = [];
  const views = backup.views ?? [];
  const charts = backup.charts ?? [];
  try {
    for (const [index, table] of backup.tables.entries()) {
      tableIds.push(await restoreBackupTable(db, bookId, table, table.name.trim(), index));
    }
    for (const [index, view] of views.entries()) {
      await db.prepare(
        'INSERT INTO _view_sheets (book_id, name, table_name, view_type, date_column, display_order) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(bookId, view.name, view.tableName, view.viewType, view.dateColumn || null, index).run();
    }
    for (const [index, chart] of charts.entries()) {
      await db.prepare(
        'INSERT INTO _chart_sheets (book_id, name, table_name, mode, charts, display_order) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(
        bookId, chart.name, chart.tableName || null, chart.mode || 'edit',
        chart.charts ? JSON.stringify(chart.charts) : '[]', index,
      ).run();
    }
  } catch (error) {
    for (const tableId of tableIds) await removeRestoredTable(db, tableId);
    await db.prepare('DELETE FROM books WHERE id = ?').bind(bookId).run();
    throw error;
  }

  return {
    bookId, bookName, tableCount: backup.tables.length,
    rowCount: backup.tables.reduce((sum, table) => sum + table.rows.length, 0),
    viewCount: views.length, chartCount: charts.length,
  };
}

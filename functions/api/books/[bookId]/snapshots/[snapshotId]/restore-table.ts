import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../lib';
import { readSnapshotData } from '../storage';

const _VALID_TYPES = new Set(['text', 'integer', 'decimal', 'date', 'datetime', 'bool', 'reference', 'image', 'calculated', 'list']);

// ---- POST /api/books/:bookId/snapshots/:snapshotId/restore-table ----

export const onRequestPost: PagesFunction<Env, 'bookId' | 'snapshotId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const snapshotId = context.params.snapshotId as string;
  const db = context.env.DB;

  const body = await context.request.json() as { tableName: string; replace?: boolean };
  if (!body.tableName?.trim()) return error('tableName is required');
  const replaceExisting = body.replace === true;

  // Load the snapshot from R2
  const row = await db.prepare(
    `SELECT data_key, created_at FROM _snapshots WHERE id = ? AND book_id = ?`
  ).bind(snapshotId, bookId).first<{ data_key: string; created_at: string }>();

  if (!row) return error('Snapshot not found', 404);

  const snapshot = await readSnapshotData(context.env, row.data_key);
  if (!snapshot) return error('Snapshot data not found in storage', 500);

  // Find the requested table in the snapshot
  const sourceTable = snapshot.tables.find(
    t => t.name.toLowerCase() === body.tableName.trim().toLowerCase()
  );
  if (!sourceTable) return error(`Table "${body.tableName}" not found in snapshot`, 404);

  const cols = sourceTable.schema?.columns ?? [];

  // Generate a unique name: "Contacts (restored from Jun 19, 2026)"
  const dateStr = new Date(row.created_at).toLocaleDateString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
  });
  const baseName = `${sourceTable.name} (restored from ${dateStr})`;

  // Check for name conflicts and append a counter if needed
  let newName = baseName;
  let counter = 2;
  while (true) {
    const existing = await db.prepare(
      'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
    ).bind(bookId, newName).first();
    if (!existing) break;
    newName = `${baseName} (${counter})`;
    counter++;
  }

  // Get next display order
  const maxOrder = await db.prepare(
    'SELECT COALESCE(MAX(display_order), -1) as m FROM _tables WHERE book_id = ?'
  ).bind(bookId).first<{ m: number }>();

  // Validate column types
  for (const col of cols) {
    if (!col.name?.trim()) return error(`Invalid column in snapshot table "${sourceTable.name}"`);
    if (!_VALID_TYPES.has(col.type)) return error(`Invalid column type "${col.type}" in snapshot table "${sourceTable.name}"`);
  }

  // Insert _tables row
  const insertResult = await db.prepare(
    `INSERT INTO _tables (book_id, name, display_order, unique_keys, default_sort, draft_position)
     VALUES (?, ?, ?, ?, ?, ?)`
  ).bind(
    bookId,
    newName,
    (maxOrder?.m ?? -1) + 1,
    JSON.stringify(sourceTable.schema?.uniqueKeys ?? []),
    sourceTable.schema?.defaultSort ? JSON.stringify(sourceTable.schema.defaultSort) : null,
    sourceTable.schema?.draftRowPosition ?? 'bottom',
  ).run();

  const tableId = insertResult.meta.last_row_id;

  try {
    // Insert column definitions
    const colStmts: D1PreparedStatement[] = cols.map((col, ci) =>
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

    // Create the physical table
    const physicalCols = cols.filter(c => c.type !== 'calculated');
    const colDefs = physicalCols.map(c => `"${c.name.trim()}" TEXT NOT NULL DEFAULT ''`).join(', ');
    const createSql = `CREATE TABLE t_${tableId} (_rowId TEXT PRIMARY KEY, ${colDefs})`;

    await db.batch([...colStmts, db.prepare(createSql)]);

    // Insert rows
    if (sourceTable.rows.length > 0) {
      const physicalColNames = physicalCols.map(c => c.name.trim());
      const allCols = ['_rowId', ...physicalColNames];
      const rowsPerStmt = Math.max(1, Math.floor(100 / allCols.length));
      const colList = allCols.map(c => `"${c}"`).join(', ');
      const rowPlaceholder = `(${allCols.map(() => '?').join(', ')})`;

      const cleanRows = sourceTable.rows.map(r => {
        const { _resolved, ...rest } = r as ImportRow;
        return rest as Record<string, string>;
      });

      const rowStmts: D1PreparedStatement[] = [];
      for (let i = 0; i < cleanRows.length; i += rowsPerStmt) {
        const chunk = cleanRows.slice(i, i + rowsPerStmt);
        const values = chunk.flatMap(r =>
          allCols.map(c => c === '_rowId' ? (r._rowId ?? '') : (r[c] ?? ''))
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
  } catch (err) {
    // Clean up the _tables row on failure
    await db.prepare('DELETE FROM _tables WHERE id = ?').bind(tableId).run();
    throw err;
  }

  // If replace mode: swap the restored table in place of the original.
  // If a previously-restored table exists, use that; otherwise restore first.
  if (replaceExisting) {
    // Check if a previously-restored table already exists
    const existingRestored = await db.prepare(
      `SELECT id FROM _tables WHERE book_id = ? AND name = ?`
    ).bind(bookId, newName).first<{ id: number }>();

    let restoredId = tableId;
    let restoredName = newName;

    if (existingRestored) {
      // Already exists from a previous restore — use it, discard the one we just created
      restoredId = existingRestored.id;
      restoredName = newName;
      // Clean up the duplicate we just created
      if (tableId !== existingRestored.id) {
        await db.prepare('DELETE FROM _tables WHERE id = ?').bind(tableId).run();
      }
    }

    const originalTable = await db.prepare(
      'SELECT id, name, display_order FROM _tables WHERE book_id = ? AND name = ?'
    ).bind(bookId, sourceTable.name).first<{ id: number; name: string; display_order: number }>();

    if (originalTable) {
      const oldName = `${sourceTable.name} (old ${new Date().toISOString().split('T')[0]})`;

      // Update all reference columns in other tables pointing to the original
      await db.prepare(
        `UPDATE _columns SET ref_table = ? WHERE ref_table = ?`
      ).bind(restoredName, sourceTable.name).run();

      // Swap: original → old, restored → original
      await db.batch([
        db.prepare('UPDATE _tables SET name = ? WHERE id = ?').bind(oldName, originalTable.id),
        db.prepare('UPDATE _tables SET display_order = ? WHERE id = ?').bind(originalTable.display_order, restoredId),
        db.prepare('UPDATE _tables SET name = ? WHERE id = ?').bind(sourceTable.name, restoredId),
        db.prepare('UPDATE _columns SET ref_table = ? WHERE ref_table = ?').bind(sourceTable.name, restoredName),
      ]);
    }
  }

  return json({
    tableName: newName,
    rowCount: sourceTable.rows.length,
    replaced: replaceExisting,
  }, 201);
};
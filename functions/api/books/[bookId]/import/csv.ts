import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

interface CsvImportBody {
  tableName: string;
  csvText: string;
}

// POST /api/books/:bookId/import/csv → parse CSV and create table + rows
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as CsvImportBody;

  if (!body.tableName?.trim()) return error('tableName is required');
  if (!body.csvText?.trim()) return error('csvText is required');

  // Parse CSV
  const lines = parseCsv(body.csvText);
  if (lines.length < 1) return error('CSV must have at least a header row');

  const headers = lines[0];
  const dataRows = lines.slice(1);

  // Get next display order
  const maxOrder = await context.env.DB.prepare(
    'SELECT COALESCE(MAX(display_order), -1) as m FROM _tables WHERE book_id = ?'
  ).bind(bookId).first<{ m: number }>();

  // Create _tables row
  const insertResult = await context.env.DB.prepare(
    `INSERT INTO _tables (book_id, name, display_order, unique_keys, default_sort, draft_position)
     VALUES (?, ?, ?, '[]', NULL, 'bottom')`
  ).bind(bookId, body.tableName.trim(), (maxOrder?.m ?? -1) + 1).run();

  const tableId = insertResult.meta.last_row_id;

  // Create column definitions (all text type for CSV import)
  const colStmts: D1PreparedStatement[] = headers.map((h, i) =>
    context.env.DB.prepare(
      `INSERT INTO _columns (table_id, name, display_name, type, display_order, ref_table, ref_display, ref_search)
       VALUES (?, ?, NULL, 'text', ?, NULL, NULL, NULL)`
    ).bind(tableId, h, i)
  );

  // Create physical table
  const colDefs = headers.map(h => `"${h}" TEXT NOT NULL DEFAULT ''`).join(', ');
  const createSql = `CREATE TABLE t_${tableId} (_rowId INTEGER PRIMARY KEY, ${colDefs})`;

  await context.env.DB.batch([...colStmts, context.env.DB.prepare(createSql)]);

  // Multi-row INSERTs: D1 allows max 100 bound params per statement.
  // Pack floor(100 / numCols) rows per INSERT to maximise throughput.
  const rowsPerStmt = Math.max(1, Math.floor(100 / headers.length));
  const colList = headers.map(h => `"${h}"`).join(', ');
  const rowPlaceholder = `(${headers.map(() => '?').join(', ')})`;
  const insertStmts: D1PreparedStatement[] = [];
  for (let i = 0; i < dataRows.length; i += rowsPerStmt) {
    const chunk = dataRows.slice(i, i + rowsPerStmt);
    const values = chunk.flatMap(row => headers.map((_, ci) => row[ci] ?? ''));
    insertStmts.push(
      context.env.DB.prepare(
        `INSERT INTO t_${tableId} (${colList}) VALUES ${chunk.map(() => rowPlaceholder).join(', ')}`
      ).bind(...values)
    );
  }

  const BATCH_SIZE = 100;
  try {
    for (let i = 0; i < insertStmts.length; i += BATCH_SIZE) {
      await context.env.DB.batch(insertStmts.slice(i, i + BATCH_SIZE));
    }
  } catch {
    // Clean up the partially-created table and metadata so the import leaves no orphans
    try {
      await context.env.DB.batch([
        context.env.DB.prepare(`DROP TABLE IF EXISTS t_${tableId}`),
        context.env.DB.prepare('DELETE FROM _columns WHERE table_id = ?').bind(tableId),
        context.env.DB.prepare('DELETE FROM _tables WHERE id = ?').bind(tableId),
      ]);
    } catch {
      // best-effort cleanup
    }
    return error('Failed to import rows — no data was saved', 500);
  }

  return json({ name: body.tableName.trim(), rowCount: dataRows.length }, 201);
};

// Simple CSV parser (handles quoted fields)
function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let current: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
        i++;
      } else if (ch === ',') {
        current.push(field);
        field = '';
        i++;
      } else if (ch === '\r' || ch === '\n') {
        current.push(field);
        field = '';
        if (current.some(f => f !== '')) rows.push(current);
        current = [];
        if (ch === '\r' && i + 1 < text.length && text[i + 1] === '\n') i++;
        i++;
      } else {
        field += ch;
        i++;
      }
    }
  }

  // Last field/row
  current.push(field);
  if (current.some(f => f !== '')) rows.push(current);

  return rows;
}

import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

interface ColumnInput {
  name: string;
  displayName?: string;
  type: string;
  width?: number;
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
}

interface CreateTableBody {
  name?: string;
  columns?: ColumnInput[];
  uniqueKeys?: string[];
  defaultSort?: { column: string; direction: string }[];
  draftRowPosition?: string;
}

const VALID_TYPES = new Set(['text', 'integer', 'decimal', 'date', 'datetime', 'bool', 'reference', 'image']);

// GET /api/books/:bookId/tables → list all tables with columns
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  const bookId = context.params.bookId as string;

  const { results: tables } = await context.env.DB.prepare(
    'SELECT id, name, display_order, unique_keys, default_sort, draft_position FROM _tables WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all<{ id: number; name: string; display_order: number; unique_keys: string; default_sort: string | null; draft_position: string }>();

  const { results: allCols } = await context.env.DB.prepare(
    `SELECT c.table_id, c.name, c.display_name, c.type, c.display_order, c.width, c.ref_table, c.ref_display, c.ref_search
     FROM _columns c
     JOIN _tables t ON t.id = c.table_id
     WHERE t.book_id = ?
     ORDER BY c.display_order`
  ).bind(bookId).all();

  const colsByTable = new Map<number, typeof allCols>();
  for (const col of allCols) {
    const tid = col.table_id as number;
    if (!colsByTable.has(tid)) colsByTable.set(tid, []);
    colsByTable.get(tid)!.push(col);
  }

  const result = tables.map(t => ({
    name: t.name,
    displayOrder: t.display_order,
    uniqueKeys: JSON.parse(t.unique_keys) as string[],
    defaultSort: t.default_sort ? JSON.parse(t.default_sort) : undefined,
    draftRowPosition: t.draft_position,
    columns: (colsByTable.get(t.id) ?? []).map(c => ({
      name: c.name as string,
      displayName: c.display_name as string | undefined,
      type: c.type as string,
      width: (c.width as number | null) ?? undefined,
      refTable: c.ref_table as string | undefined,
      refDisplayColumns: c.ref_display ? JSON.parse(c.ref_display as string) : undefined,
      refSearchColumns: c.ref_search ? JSON.parse(c.ref_search as string) : undefined,
    })),
  }));

  return json(result);
};

// POST /api/books/:bookId/tables → create a new table
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as CreateTableBody;

  if (!body.name?.trim()) return error('name is required');
  if (!body.columns?.length) return error('at least one column is required');

  // Validate column types
  for (const col of body.columns) {
    if (!col.name?.trim()) return error('column name is required');
    if (!VALID_TYPES.has(col.type)) return error(`invalid column type: ${col.type}`);
  }

  // Get next display order
  const maxOrder = await context.env.DB.prepare(
    'SELECT COALESCE(MAX(display_order), -1) as m FROM _tables WHERE book_id = ?'
  ).bind(bookId).first<{ m: number }>();

  // Insert _tables row
  const insertResult = await context.env.DB.prepare(
    `INSERT INTO _tables (book_id, name, display_order, unique_keys, default_sort, draft_position)
     VALUES (?, ?, ?, ?, ?, ?)`
  ).bind(
    bookId,
    body.name.trim(),
    (maxOrder?.m ?? -1) + 1,
    JSON.stringify(body.uniqueKeys ?? []),
    body.defaultSort ? JSON.stringify(body.defaultSort) : null,
    body.draftRowPosition ?? 'bottom',
  ).run();

  const tableId = insertResult.meta.last_row_id;

  // Insert column definitions
  const colStmts: D1PreparedStatement[] = body.columns.map((col, i) =>
    context.env.DB.prepare(
      `INSERT INTO _columns (table_id, name, display_name, type, display_order, width, ref_table, ref_display, ref_search)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      tableId,
      col.name.trim(),
      col.displayName || null,
      col.type,
      i,
      col.width ?? null,
      col.refTable || null,
      col.refDisplayColumns ? JSON.stringify(col.refDisplayColumns) : null,
      col.refSearchColumns ? JSON.stringify(col.refSearchColumns) : null,
    )
  );

  // Create the physical table: t_{id} with _rowId + all columns as TEXT
  const colDefs = body.columns.map(c => `"${c.name.trim()}" TEXT NOT NULL DEFAULT ''`).join(', ');
  const createSql = `CREATE TABLE t_${tableId} (_rowId INTEGER PRIMARY KEY, ${colDefs})`;

  await context.env.DB.batch([
    ...colStmts,
    context.env.DB.prepare(createSql),
  ]);

  return json({ name: body.name.trim(), tableId }, 201);
};

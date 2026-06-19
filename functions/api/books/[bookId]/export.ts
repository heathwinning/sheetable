import type { Env, RequestData } from '../../../lib';
import { error, requireUser } from '../../../lib';

// ---- Types ----

interface ExportColumn {
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

interface ExportTableSchema {
  columns: ExportColumn[];
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: string }[];
  draftRowPosition?: string;
}

interface ExportRow {
  _rowId: string;
  [column: string]: string;
}

interface ResolvedRow extends ExportRow {
  _resolved: Record<string, string>;
}

interface ExportTable {
  name: string;
  schema: ExportTableSchema | null; // set after resolveAllReferences
  rows: ResolvedRow[];
}

interface ExportView {
  name: string;
  tableName: string;
  viewType: string;
  dateColumn?: string;
}

interface ExportChart {
  name: string;
  tableName?: string;
  mode?: string;
  charts?: unknown;
}

interface ExportPayload {
  version: 1;
  exportedAt: string;
  bookName: string;
  tables: ExportTable[];
  views: ExportView[];
  charts: ExportChart[];
}

// ---- Reference resolution ----

interface ColMeta {
  name: string;
  type: string;
  refTable: string | null;
  refDisplayColumns: string[];
}

/** Load all rows for a table keyed by _rowId. */
async function loadRows(
  db: D1Database,
  tableId: number,
  colNames: string[],
): Promise<Map<string, ExportRow>> {
  const { results } = await db.prepare(`SELECT * FROM t_${tableId}`).all<Record<string, unknown>>();
  const map = new Map<string, ExportRow>();
  for (const r of results) {
    const row: ExportRow = { _rowId: String(r._rowId) };
    for (const c of colNames) {
      row[c] = String(r[c] ?? '');
    }
    map.set(row._rowId, row);
  }
  return map;
}

/**
 * Resolve a single column value. If the column is a reference, follow the
 * display columns (which may themselves be references) to produce a display
 * string. Returns empty string if the referenced row can't be found.
 */
function resolveColumnValue(
  tableName: string,
  col: ColMeta,
  value: string,
  allRows: Map<string, Map<string, ExportRow>>,
  allCols: Map<string, ColMeta[]>,
  depth: number,
): string {
  if (depth > 8) return value; // safety limit
  if (!value) return '';

  if (col.type !== 'reference' || !col.refTable) return value;

  const refRows = allRows.get(col.refTable);
  if (!refRows) return value;

  const refRow = refRows.get(value);
  if (!refRow) return '';

  const displayCols = col.refDisplayColumns;
  if (!displayCols || displayCols.length === 0) return value;

  const refCols = allCols.get(col.refTable) ?? [];
  const refColMap = new Map(refCols.map(c => [c.name, c]));

  const parts: string[] = [];
  for (const dc of displayCols) {
    // dc may be a dotted path like "category.name"
    const dotIdx = dc.indexOf('.');
    if (dotIdx >= 0) {
      const head = dc.substring(0, dotIdx);
      const tail = dc.substring(dotIdx + 1);
      const headCol = refColMap.get(head);
      if (headCol && headCol.type === 'reference' && headCol.refTable) {
        const headVal = refRow[head] ?? '';
        const resolved = resolveColumnValue(
          col.refTable,
          headCol,
          headVal,
          allRows,
          allCols,
          depth + 1,
        );
        if (resolved) {
          // Now follow tail through the referenced row
          const headRefRows = allRows.get(headCol.refTable);
          const headRefRow = headRefRows?.get(headVal);
          if (headRefRow) {
            const tailResolved = resolveColumnValue(
              headCol.refTable,
              refColMap.get(tail) ?? { name: tail, type: 'text', refTable: null, refDisplayColumns: [] },
              headRefRow[tail] ?? '',
              allRows,
              allCols,
              depth + 1,
            );
            parts.push(tailResolved || headRefRow[tail] || '');
            continue;
          }
        }
      }
      parts.push(refRow[dc] ?? '');
    } else {
      const dcCol = refColMap.get(dc);
      if (dcCol && dcCol.type === 'reference' && dcCol.refTable) {
        const dcVal = refRow[dc] ?? '';
        const resolved = resolveColumnValue(
          col.refTable,
          dcCol,
          dcVal,
          allRows,
          allCols,
          depth + 1,
        );
        parts.push(resolved || dcVal);
      } else {
        parts.push(refRow[dc] ?? '');
      }
    }
  }

  return parts.filter(Boolean).join(' · ');
}

/** Resolve all reference columns in all rows across all tables. */
function resolveAllReferences(
  tables: { name: string; tableId: number; cols: ColMeta[]; rows: Map<string, ExportRow> }[],
): ExportTable[] {
  // Build lookup maps
  const allRows = new Map<string, Map<string, ExportRow>>();
  const allCols = new Map<string, ColMeta[]>();
  for (const t of tables) {
    allRows.set(t.name, t.rows);
    allCols.set(t.name, t.cols);
  }

  return tables.map(t => {
    const resolvedRows: ResolvedRow[] = [];
    for (const [_rowId, row] of t.rows) {
      const resolved: Record<string, string> = {};
      for (const col of t.cols) {
        if (col.type === 'reference' && col.refTable) {
          const rawVal = row[col.name] ?? '';
          resolved[col.name] = resolveColumnValue(
            t.name,
            col,
            rawVal,
            allRows,
            allCols,
            0,
          );
        }
      }
      resolvedRows.push({ ...row, _resolved: resolved });
    }
    // Sort by _rowId numerically for stable output
    resolvedRows.sort((a, b) => {
      const na = parseInt(a._rowId, 10);
      const nb = parseInt(b._rowId, 10);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a._rowId.localeCompare(b._rowId);
    });
    return { name: t.name, schema: null, rows: resolvedRows };
  });
}

// ---- Shared: gather all book data as an ExportPayload ----

export async function gatherExportPayload(env: Env, bookId: string): Promise<ExportPayload> {
  const db = env.DB;

  const book = await db.prepare('SELECT name FROM books WHERE id = ?')
    .bind(bookId).first<{ name: string }>();
  if (!book) throw new Error('Book not found');

  // Get all tables with their metadata
  const { results: tableRows } = await db.prepare(
    'SELECT id, name, unique_keys, default_sort, draft_position FROM _tables WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all<{
    id: number; name: string; unique_keys: string;
    default_sort: string | null; draft_position: string;
  }>();

  // Get all columns for all tables in this book
  const { results: allCols } = await db.prepare(
    `SELECT c.table_id, c.name, c.display_name, c.type, c.width, c.ref_table, c.ref_display, c.ref_search, c.expression, c.show_in_grid, c.list_of
     FROM _columns c JOIN _tables t ON t.id = c.table_id
     WHERE t.book_id = ? ORDER BY c.display_order`
  ).bind(bookId).all<{
    table_id: number; name: string; display_name: string | null;
    type: string; width: number | null; ref_table: string | null;
    ref_display: string | null; ref_search: string | null;
    expression: string | null; show_in_grid: number | null; list_of: string | null;
  }>();

  // Group columns by table
  const colsByTable = new Map<number, typeof allCols>();
  for (const col of allCols) {
    if (!colsByTable.has(col.table_id)) colsByTable.set(col.table_id, []);
    colsByTable.get(col.table_id)!.push(col);
  }

  // Build column metadata and load rows
  const tableData: { name: string; tableId: number; cols: ColMeta[]; rows: Map<string, ExportRow> }[] = [];
  for (const t of tableRows) {
    const rawCols = colsByTable.get(t.id) ?? [];
    const cols: ColMeta[] = rawCols.map(c => ({
      name: c.name,
      type: c.type,
      refTable: c.ref_table,
      refDisplayColumns: c.ref_display ? JSON.parse(c.ref_display) as string[] : [],
    }));
    const colNames = cols.map(c => c.name);
    const rows = await loadRows(db, t.id, colNames);
    tableData.push({ name: t.name, tableId: t.id, cols, rows });
  }

  // Resolve references (must happen after all rows are loaded)
  const tables = resolveAllReferences(tableData);

  // Augment schemas with full metadata from the DB
  for (const t of tables) {
    const td = tableData.find(td => td.name === t.name);
    if (!td) continue;
    const rawCols = colsByTable.get(td.tableId) ?? [];
    const dbTable = tableRows.find(tr => tr.id === td.tableId);
    t.schema = {
      columns: rawCols.map(c => ({
        name: c.name,
        displayName: c.display_name || undefined,
        type: c.type,
        width: c.width ?? undefined,
        refTable: c.ref_table || undefined,
        refDisplayColumns: c.ref_display ? JSON.parse(c.ref_display) as string[] : undefined,
        refSearchColumns: c.ref_search ? JSON.parse(c.ref_search) as string[] : undefined,
        expression: c.expression || undefined,
        showInGrid: c.show_in_grid ? true : undefined,
        listOf: c.list_of || undefined,
      })),
      uniqueKeys: dbTable ? JSON.parse(dbTable.unique_keys) as string[] : [],
      defaultSort: dbTable?.default_sort ? JSON.parse(dbTable.default_sort) : undefined,
      draftRowPosition: dbTable?.draft_position ?? 'bottom',
    };
  }

  // Get views
  const { results: views } = await db.prepare(
    'SELECT name, table_name, view_type, date_column FROM _view_sheets WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all<{ name: string; table_name: string; view_type: string; date_column: string | null }>();

  // Get charts
  const { results: charts } = await db.prepare(
    'SELECT name, table_name, mode, charts as chart_data FROM _chart_sheets WHERE book_id = ? ORDER BY display_order'
  ).bind(bookId).all<{ name: string; table_name: string | null; mode: string | null; chart_data: string | null }>();

  return {
    version: 1,
    exportedAt: new Date().toISOString(),
    bookName: book.name,
    tables,
    views: views.map(v => ({
      name: v.name,
      tableName: v.table_name,
      viewType: v.view_type,
      dateColumn: v.date_column || undefined,
    })),
    charts: charts.map(c => ({
      name: c.name,
      tableName: c.table_name || undefined,
      mode: c.mode || undefined,
      charts: c.chart_data ? JSON.parse(c.chart_data) : undefined,
    })),
  };
}

// ---- GET /api/books/:bookId/export ----

export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  if (!context.data.bookRole) return error('Forbidden', 403);

  const bookId = context.params.bookId as string;

  let payload: ExportPayload;
  try {
    payload = await gatherExportPayload(context.env, bookId);
  } catch (err) {
    if (err instanceof Error && err.message === 'Book not found') return error('Book not found', 404);
    throw err;
  }

  const filename = `${payload.bookName.replace(/[^a-zA-Z0-9_-]/g, '_')}-export-${new Date().toISOString().split('T')[0]}.json`;
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: {
      'Content-Type': 'application/json',
      'Content-Disposition': `attachment; filename="${filename}"`,
    },
  });
};

import type { Env, RequestData } from '../../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../../lib';

interface ColumnInput {
  name: string;
  displayName?: string;
  type: string;
  width?: number;
  refTable?: string;
  refDisplayColumns?: string[];
  refSearchColumns?: string[];
}

interface SchemaBody {
  columns?: ColumnInput[];
  uniqueKeys?: string[];
  defaultSort?: { column: string; direction: string }[];
  draftRowPosition?: string;
  calculatedColumns?: { name: string; displayName?: string; expression: string }[];
}

const VALID_TYPES = new Set(['text', 'integer', 'decimal', 'date', 'datetime', 'bool', 'reference', 'image']);

// PUT /api/books/:bookId/tables/:name/schema → update column definitions and table settings
export const onRequestPut: PagesFunction<Env, 'bookId' | 'name', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const name = decodeURIComponent(context.params.name as string);
  const body = await context.request.json() as SchemaBody;

  const table = await context.env.DB.prepare(
    'SELECT id FROM _tables WHERE book_id = ? AND name = ?'
  ).bind(bookId, name).first<{ id: number }>();

  if (!table) return error('Table not found', 404);

  const stmts: D1PreparedStatement[] = [];

  // Update table-level settings
  stmts.push(
    context.env.DB.prepare(
      'UPDATE _tables SET unique_keys = ?, default_sort = ?, draft_position = ?, calculated_columns = ? WHERE id = ?'
    ).bind(
      JSON.stringify(body.uniqueKeys ?? []),
      body.defaultSort ? JSON.stringify(body.defaultSort) : null,
      body.draftRowPosition ?? 'bottom',
      body.calculatedColumns ? JSON.stringify(body.calculatedColumns) : null,
      table.id,
    )
  );

  if (body.columns) {
    // Validate
    for (const col of body.columns) {
      if (!col.name?.trim()) return error('column name is required');
      if (!VALID_TYPES.has(col.type)) return error(`invalid column type: ${col.type}`);
    }

    // Get existing columns
    const { results: existing } = await context.env.DB.prepare(
      'SELECT name FROM _columns WHERE table_id = ?'
    ).bind(table.id).all<{ name: string }>();
    const existingNames = new Set(existing.map(c => c.name));
    const newNames = new Set(body.columns.map(c => c.name.trim()));

    // Delete removed columns
    stmts.push(
      context.env.DB.prepare('DELETE FROM _columns WHERE table_id = ?').bind(table.id)
    );

    // Insert all columns fresh (simpler than diffing)
    for (let i = 0; i < body.columns.length; i++) {
      const col = body.columns[i];
      stmts.push(
        context.env.DB.prepare(
          `INSERT INTO _columns (table_id, name, display_name, type, display_order, width, ref_table, ref_display, ref_search)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          table.id,
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
    }

    // ALTER physical table: add new columns, drop removed columns
    for (const col of body.columns) {
      if (!existingNames.has(col.name.trim())) {
        stmts.push(
          context.env.DB.prepare(`ALTER TABLE t_${table.id} ADD COLUMN "${col.name.trim()}" TEXT NOT NULL DEFAULT ''`)
        );
      }
    }

    for (const name of existingNames) {
      if (!newNames.has(name)) {
        stmts.push(
          context.env.DB.prepare(`ALTER TABLE t_${table.id} DROP COLUMN "${name}"`)
        );
      }
    }
  }

  await context.env.DB.batch(stmts);

  return json({ ok: true });
};

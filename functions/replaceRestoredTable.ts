interface TableLocation {
  id: number;
  name: string;
  display_order: number;
}

export function findTable(db: D1Database, bookId: string, name: string) {
  return db.prepare('SELECT id, name, display_order FROM _tables WHERE book_id = ? AND name = ? AND restore_pending = 0')
    .bind(bookId, name).first<TableLocation>();
}

export async function replaceRestoredTable(
  db: D1Database, bookId: string, original: TableLocation, restored: TableLocation,
): Promise<string | null> {
  const date = new Date().toISOString().slice(0, 10);
  const archiveName = `${original.name} (old ${date}) (${crypto.randomUUID().slice(0, 8)})`;
  const [archived, promoted] = await db.batch([
    db.prepare(
      `UPDATE _tables SET name = ?,
         display_order = (SELECT COALESCE(MAX(display_order), -1) + 1 FROM _tables WHERE book_id = ?)
       WHERE book_id = ? AND id = ? AND name = ?
       AND EXISTS (SELECT 1 FROM _tables WHERE book_id = ? AND id = ? AND name = ?)`
    ).bind(archiveName, bookId, bookId, original.id, original.name, bookId, restored.id, restored.name),
    db.prepare(
      `UPDATE _tables SET name = ?, display_order = ?, restore_pending = 0 WHERE book_id = ? AND id = ? AND name = ?
       AND EXISTS (SELECT 1 FROM _tables WHERE book_id = ? AND id = ? AND name = ?)`
    ).bind(original.name, original.display_order, bookId, restored.id, restored.name, bookId, original.id, archiveName),
    db.prepare(
      `UPDATE _columns SET ref_table = ? WHERE ref_table = ?
       AND table_id IN (SELECT id FROM _tables WHERE book_id = ?)
       AND EXISTS (SELECT 1 FROM _tables WHERE book_id = ? AND id = ? AND name = ?)
       AND EXISTS (SELECT 1 FROM _tables WHERE book_id = ? AND id = ? AND name = ?)`
    ).bind(original.name, restored.name, bookId, bookId, original.id, archiveName, bookId, restored.id, original.name),
  ]);
  return archived.meta.changes && promoted.meta.changes ? archiveName : null;
}

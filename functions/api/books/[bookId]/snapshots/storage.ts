import type { Env } from '../../../../lib';

const SNAPSHOT_R2_PREFIX = 'snapshots';

function snapshotKey(bookId: string, snapshotId: string): string {
  return `${SNAPSHOT_R2_PREFIX}/${bookId}/${snapshotId}.json`;
}

export interface ExportPayload {
  version: number;
  exportedAt: string;
  bookName: string;
  tables: unknown[];
  views: unknown[];
  charts: unknown[];
}

/** Write a snapshot to R2 and insert D1 metadata. Runs retention cleanup after. */
export async function createSnapshot(
  env: Env,
  bookId: string,
  payload: ExportPayload,
  label: string | null,
): Promise<{ snapshotId: string }> {
  const snapshotId = crypto.randomUUID();
  const now = new Date().toISOString();
  const nowDate = new Date();
  const totalRows = (payload.tables as { rows: unknown[] }[]).reduce((sum, t) => sum + t.rows.length, 0);

  // Classify: annual (Jan 1), monthly (1st of month), or daily
  const isFirstOfYear = nowDate.getMonth() === 0 && nowDate.getDate() === 1;
  const isFirstOfMonth = nowDate.getDate() === 1;
  const type = isFirstOfYear ? 'annual' : isFirstOfMonth ? 'monthly' : 'daily';

  // Write payload to R2 (gzip-compressed)
  const key = snapshotKey(bookId, snapshotId);
  const body = await gzip(JSON.stringify(payload));
  await env.BUCKET.put(key, body, {
    httpMetadata: { contentType: 'application/json', contentEncoding: 'gzip' },
  });

  // Insert metadata in D1
  await env.DB.prepare(
    `INSERT INTO _snapshots (id, book_id, label, created_at, table_count, row_count, view_count, chart_count, data_key, type)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).bind(
    snapshotId,
    bookId,
    label,
    now,
    payload.tables.length,
    totalRows,
    payload.views.length,
    payload.charts.length,
    key,
    type,
  ).run();

  // Run retention cleanup
  await cleanupSnapshots(env, bookId);

  return { snapshotId };
}

/** Read a snapshot payload from R2. Returns null if not found. */
export async function readSnapshotData(env: Env, snapshotKey: string): Promise<ExportPayload | null> {
  const obj = await env.BUCKET.get(snapshotKey);
  if (!obj) return null;
  const buf = await obj.arrayBuffer();
  const text = new TextDecoder().decode(await gunzip(buf));
  return JSON.parse(text) as ExportPayload;
}

async function gzip(data: string): Promise<Uint8Array> {
  const stream = new CompressionStream('gzip');
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();
  writer.write(new TextEncoder().encode(data));
  writer.close();

  const chunks: Uint8Array[] = [];
  let done = false;
  while (!done) {
    const result = await reader.read();
    if (result.value) chunks.push(result.value);
    done = result.done;
  }
  // Concatenate
  const totalLen = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(totalLen);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out;
}

async function gunzip(data: ArrayBuffer): Promise<ArrayBuffer> {
  const stream = new DecompressionStream('gzip');
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();
  writer.write(new Uint8Array(data));
  writer.close();

  const chunks: Uint8Array[] = [];
  let done = false;
  while (!done) {
    const result = await reader.read();
    if (result.value) chunks.push(result.value);
    done = result.done;
  }
  const totalLen = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(totalLen);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.length;
  }
  return out.buffer;
}

/** Delete snapshot data from R2. */
export async function deleteSnapshotData(env: Env, snapshotKey: string): Promise<void> {
  await env.BUCKET.delete(snapshotKey);
}

/**
 * Retention policy — count-based per tier:
 * - Daily: keep last 30
 * - Monthly: keep last 12
 * - Annual: unlimited
 */
export async function cleanupSnapshots(env: Env, bookId: string): Promise<void> {
  // Delete oldest daily snapshots beyond 30
  await purgeTier(env, bookId, 'daily', 30);
  // Delete oldest monthly snapshots beyond 12
  await purgeTier(env, bookId, 'monthly', 12);
}

async function _purgeOldTables(env: Env, bookId: string): Promise<void> {
  const db = env.DB;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);
  const cutoffStr = cutoff.toISOString().split('T')[0];

  const { results } = await db.prepare(
    `SELECT id, name FROM _tables
     WHERE book_id = ? AND name LIKE '% (old ____-__-__)'
       AND SUBSTR(name, -11, 10) < ?`
  ).bind(bookId, cutoffStr).all<{ id: number; name: string }>();

  for (const t of results) {
    await db.prepare('DELETE FROM _tables WHERE id = ?').bind(t.id).run();
  }
}

async function purgeTier(env: Env, bookId: string, type: string, keep: number): Promise<void> {
  const db = env.DB;

  // Get data_keys to delete from R2 before removing DB rows
  const { results } = await db.prepare(
    `SELECT data_key FROM _snapshots WHERE book_id = ? AND type = ? AND id IN (
       SELECT id FROM _snapshots WHERE book_id = ? AND type = ?
       ORDER BY created_at DESC LIMIT -1 OFFSET ?
     )`
  ).bind(bookId, type, bookId, type, keep).all<{ data_key: string }>();

  // Delete from R2
  for (const row of results) {
    await env.BUCKET.delete(row.data_key);
  }

  // Delete from D1
  await db.prepare(
    `DELETE FROM _snapshots WHERE book_id = ? AND type = ? AND id IN (
       SELECT id FROM _snapshots WHERE book_id = ? AND type = ?
       ORDER BY created_at DESC LIMIT -1 OFFSET ?
     )`
  ).bind(bookId, type, bookId, type, keep).run();
}

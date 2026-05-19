import type { SessionUser, BookInfo, BookMember, BookInvite, TableSchema, Row, ChartSheet, ViewSheet } from './types';

const BASE = '/api';

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...init?.headers,
    },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText })) as { error?: string };
    throw new Error(body.error ?? `HTTP ${res.status}`);
  }
  return res.json() as Promise<T>;
}

// ---- Auth ----

export function loginUrl(): string {
  return `${BASE}/auth/login`;
}

export async function getMe(): Promise<SessionUser> {
  return request('/auth/me');
}

export async function logout(): Promise<void> {
  await request('/auth/logout', { method: 'POST' });
}

// ---- Books ----

export async function listBooks(): Promise<BookInfo[]> {
  return request('/books');
}

export async function createBook(name: string): Promise<{ id: string; name: string }> {
  return request('/books', { method: 'POST', body: JSON.stringify({ name }) });
}

export async function renameBook(bookId: string, name: string): Promise<void> {
  await request(`/books/${bookId}`, { method: 'PATCH', body: JSON.stringify({ name }) });
}

export async function deleteBook(bookId: string): Promise<void> {
  await request(`/books/${bookId}`, { method: 'DELETE' });
}

export async function reorderSheets(bookId: string, tableOrder?: string[], chartOrder?: string[]): Promise<void> {
  await request(`/books/${bookId}`, {
    method: 'PATCH',
    body: JSON.stringify({ tableOrder, chartOrder }),
  });
}

// ---- Members ----

export async function listMembers(bookId: string): Promise<{ members: BookMember[]; invites: BookInvite[] }> {
  return request(`/books/${bookId}/members`);
}

export async function addMember(bookId: string, email: string, role: string): Promise<{ ok: boolean; invited?: boolean }> {
  return request(`/books/${bookId}/members`, { method: 'POST', body: JSON.stringify({ email, role }) });
}

export async function removeMember(bookId: string, userId: string): Promise<void> {
  await request(`/books/${bookId}/members/${userId}`, { method: 'DELETE' });
}

export async function cancelInvite(bookId: string, email: string): Promise<void> {
  await request(`/books/${bookId}/invites/${encodeURIComponent(email)}`, { method: 'DELETE' });
}

export async function getInviteStatus(bookId: string): Promise<{ bookName: string; status: string; role?: string }> {
  return request(`/books/${bookId}/invite`);
}

export async function acceptInvite(bookId: string): Promise<{ ok: boolean; role: string }> {
  return request(`/books/${bookId}/invite`, { method: 'POST' });
}

// ---- Tables ----

interface ApiTable {
  name: string;
  displayOrder: number;
  uniqueKeys: string[];
  defaultSort?: { column: string; direction: 'asc' | 'desc' }[];
  draftRowPosition: string;
  /** Legacy: migrated to columns with type='calculated' on load */
  calculatedColumns?: { name: string; expression: string; showInGrid?: boolean }[];
  columns: {
    name: string;
    displayName?: string;
    type: string;
    refTable?: string;
    refDisplayColumns?: string[];
    refSearchColumns?: string[];
    expression?: string;
    showInGrid?: boolean;
  }[];
}

export async function listTables(bookId: string): Promise<TableSchema[]> {
  const tables = await request<ApiTable[]>(`/books/${bookId}/tables`);
  return tables.map(t => {
    const baseCols = t.columns.map(c => ({
      name: c.name,
      displayName: c.displayName,
      type: c.type as TableSchema['columns'][0]['type'],
      refTable: c.refTable,
      refDisplayColumns: c.refDisplayColumns,
      refSearchColumns: c.refSearchColumns,
      expression: c.expression,
      showInGrid: c.showInGrid,
    }));
    // Migrate legacy calculatedColumns into the columns array
    const legacyCalcCols = (t.calculatedColumns ?? []).map(calc => ({
      name: calc.name,
      type: 'calculated' as const,
      expression: calc.expression,
      showInGrid: calc.showInGrid,
    }));
    const calcNames = new Set(legacyCalcCols.map(c => c.name));
    const columns = [
      ...baseCols.filter(c => !calcNames.has(c.name)),
      ...legacyCalcCols,
    ];
    return {
      name: t.name,
      columns,
      uniqueKeys: t.uniqueKeys,
      defaultSort: t.defaultSort,
      draftRowPosition: t.draftRowPosition as 'top' | 'bottom',
    };
  });
}

export async function createTable(bookId: string, schema: TableSchema, rows?: Row[]): Promise<void> {
  await request(`/books/${bookId}/tables`, {
    method: 'POST',
    body: JSON.stringify({
      name: schema.name,
      columns: schema.columns.map(c => ({
        name: c.name,
        displayName: c.displayName,
        type: c.type,
        width: c.width,
        refTable: c.refTable,
        refDisplayColumns: c.refDisplayColumns,
        refSearchColumns: c.refSearchColumns,
        expression: c.expression,
        showInGrid: c.showInGrid,
      })),
      uniqueKeys: schema.uniqueKeys,
      defaultSort: schema.defaultSort,
      draftRowPosition: schema.draftRowPosition,
    }),
  });

  // Insert rows in bulk if provided
  if (rows?.length) {
    const operations = rows.map((row, i) => ({
      type: 'insert' as const,
      rowId: row._rowId || String(i + 1),
      data: Object.fromEntries(
        Object.entries(row).filter(([k]) => k !== '_rowId').map(([k, v]) => [k, String(v ?? '')])
      ),
    }));
    await bulkRowOps(bookId, schema.name, operations);
  }
}

export async function renameTable(bookId: string, oldName: string, newName: string): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(oldName)}`, {
    method: 'PATCH',
    body: JSON.stringify({ name: newName }),
  });
}

export async function deleteTable(bookId: string, name: string): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(name)}`, { method: 'DELETE' });
}

export async function updateTableSchema(bookId: string, name: string, schema: TableSchema): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(name)}/schema`, {
    method: 'PUT',
    body: JSON.stringify({
      columns: schema.columns.map(c => ({
        name: c.name,
        displayName: c.displayName,
        type: c.type,
        width: c.width,
        refTable: c.refTable,
        refDisplayColumns: c.refDisplayColumns,
        refSearchColumns: c.refSearchColumns,
        expression: c.expression,
        showInGrid: c.showInGrid,
      })),
      uniqueKeys: schema.uniqueKeys,
      defaultSort: schema.defaultSort,
      draftRowPosition: schema.draftRowPosition,
      calculatedColumns: undefined,
    }),
  });
}

// ---- Rows ----

export async function listRows(bookId: string, tableName: string): Promise<Row[]> {
  return request(`/books/${bookId}/tables/${encodeURIComponent(tableName)}/rows`);
}

export async function insertRow(bookId: string, tableName: string, row: Row): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(tableName)}/rows`, {
    method: 'POST',
    body: JSON.stringify(row),
  });
}

export async function updateRow(bookId: string, tableName: string, rowId: string, data: Record<string, string>): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(tableName)}/rows/${encodeURIComponent(rowId)}`, {
    method: 'PUT',
    body: JSON.stringify(data),
  });
}

export async function deleteRow(bookId: string, tableName: string, rowId: string): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(tableName)}/rows/${encodeURIComponent(rowId)}`, {
    method: 'DELETE',
  });
}

export async function bulkRowOps(
  bookId: string,
  tableName: string,
  operations: Array<{ type: 'insert' | 'update' | 'delete'; rowId: string; data?: Record<string, string> }>,
): Promise<void> {
  await request(`/books/${bookId}/tables/${encodeURIComponent(tableName)}/rows/bulk`, {
    method: 'POST',
    body: JSON.stringify({ operations }),
  });
}

// ---- Charts ----

export async function listCharts(bookId: string): Promise<ChartSheet[]> {
  return request(`/books/${bookId}/charts`);
}

export async function createChart(bookId: string, name: string, tableName?: string): Promise<void> {
  await request(`/books/${bookId}/charts`, {
    method: 'POST',
    body: JSON.stringify({ name, tableName }),
  });
}

export async function updateChart(
  bookId: string,
  name: string,
  updates: Partial<{ name: string; tableName: string; mode: string; charts: unknown; displayOrder: number }>,
): Promise<void> {
  await request(`/books/${bookId}/charts/${encodeURIComponent(name)}`, {
    method: 'PATCH',
    body: JSON.stringify(updates),
  });
}

export async function deleteChart(bookId: string, name: string): Promise<void> {
  await request(`/books/${bookId}/charts/${encodeURIComponent(name)}`, { method: 'DELETE' });
}

// ---- View sheets ----

export async function listViews(bookId: string): Promise<ViewSheet[]> {
  return request(`/books/${bookId}/views`);
}

export async function createView(
  bookId: string,
  name: string,
  tableName: string,
  viewType: ViewSheet['viewType'],
  dateColumn?: string,
): Promise<void> {
  await request(`/books/${bookId}/views`, {
    method: 'POST',
    body: JSON.stringify({ name, tableName, viewType, dateColumn }),
  });
}

export async function updateView(
  bookId: string,
  name: string,
  updates: Partial<{ name: string; tableName: string; viewType: string; dateColumn: string | null; hideSourceTableTab: boolean; displayOrder: number }>,
): Promise<void> {
  await request(`/books/${bookId}/views/${encodeURIComponent(name)}`, {
    method: 'PATCH',
    body: JSON.stringify(updates),
  });
}

export async function deleteView(bookId: string, name: string): Promise<void> {
  await request(`/books/${bookId}/views/${encodeURIComponent(name)}`, { method: 'DELETE' });
}

// ---- Images ----

export async function getUploadUrl(bookId: string, filename: string): Promise<{ key: string; uploadUrl: string }> {
  return request(`/books/${bookId}/images/upload-url`, {
    method: 'POST',
    body: JSON.stringify({ filename }),
  });
}

export async function uploadImage(uploadUrl: string, file: File): Promise<void> {
  const res = await fetch(uploadUrl, {
    method: 'PUT',
    headers: { 'Content-Type': file.type },
    body: file,
  });
  if (!res.ok) throw new Error('Image upload failed');
}

export function imageUrl(bookId: string, key: string): string {
  return `${BASE}/books/${bookId}/images/${encodeURIComponent(key)}`;
}

// ---- Import ----

export async function importCsv(bookId: string, tableName: string, csvText: string): Promise<{ name: string; rowCount: number }> {
  return request(`/books/${bookId}/import/csv`, {
    method: 'POST',
    body: JSON.stringify({ tableName, csvText }),
  });
}

export async function fetchGoogleSheetList(bookId: string, spreadsheetId: string): Promise<{ title: string; sheets: { sheetId: number; title: string; index: number }[] }> {
  return request(`/books/${bookId}/import/google-sheet?spreadsheetId=${encodeURIComponent(spreadsheetId)}`);
}

export async function fetchGoogleSheet(bookId: string, spreadsheetId: string, gid?: number): Promise<{ csvText: string }> {
  return request(`/books/${bookId}/import/google-sheet`, {
    method: 'POST',
    body: JSON.stringify({ spreadsheetId, gid }),
  });
}

export async function fetchCsvFromUrl(bookId: string, url: string): Promise<{ csvText: string }> {
  return request(`/books/${bookId}/import/fetch-url`, {
    method: 'POST',
    body: JSON.stringify({ url }),
  });
}

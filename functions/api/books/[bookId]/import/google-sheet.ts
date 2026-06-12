import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// GET /api/books/:bookId/import/google-sheet?spreadsheetId=... → list sheets in a spreadsheet
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const url = new URL(context.request.url);
  const spreadsheetId = url.searchParams.get('spreadsheetId')?.trim();
  if (!spreadsheetId) return error('spreadsheetId is required');

  const apiKey = context.env.GOOGLE_API_KEY;
  if (!apiKey) return error('Google API key not configured', 500);

  const apiUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?fields=properties.title,sheets.properties&key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(apiUrl);
  if (!res.ok) {
    return error(`Failed to fetch spreadsheet metadata: ${res.status}`, 502);
  }

  const data = await res.json() as {
    properties?: { title?: string };
    sheets?: { properties?: { sheetId?: number; title?: string; index?: number } }[];
  };

  const sheets = (data.sheets ?? []).map((s) => ({
    sheetId: s.properties?.sheetId ?? 0,
    title: s.properties?.title ?? '',
    index: s.properties?.index ?? 0,
  }));

  return json({ title: data.properties?.title ?? '', sheets });
};

// POST /api/books/:bookId/import/google-sheet → fetch Google Sheet and return parsed data
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const body = await context.request.json() as { spreadsheetId?: string; gid?: number };

  if (!body.spreadsheetId?.trim()) return error('spreadsheetId is required');

  const gidParam = body.gid != null ? `&gid=${body.gid}` : '';
  const exportUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(body.spreadsheetId)}/export?format=csv${gidParam}`;

  const res = await fetch(exportUrl);
  if (!res.ok) {
    return error(`Failed to fetch Google Sheet: ${res.status}`, 502);
  }

  const csvText = await res.text();
  if (csvText.length > 10 * 1024 * 1024) {
    return error('Sheet is too large (max 10 MB)');
  }
  return json({ csvText });
};

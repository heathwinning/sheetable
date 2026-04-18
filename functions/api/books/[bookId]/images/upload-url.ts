import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// POST /api/books/:bookId/images/upload-url → return a presigned PUT URL for R2
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as { filename?: string; contentType?: string };

  if (!body.filename) return error('filename is required');

  const ext = body.filename.split('.').pop() ?? 'bin';
  const key = `${bookId}/${crypto.randomUUID()}.${ext}`;

  // We can't create presigned URLs with R2 binding directly.
  // Instead, we'll have the client PUT through our Worker as a proxy.
  // Return the key so the client can use PUT /api/books/:bookId/images/:key
  return json({ key, uploadUrl: `/api/books/${bookId}/images/${encodeURIComponent(key)}` });
};

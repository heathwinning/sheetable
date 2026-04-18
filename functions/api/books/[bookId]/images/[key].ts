import type { Env, RequestData } from '../../../../lib';
import { error, requireUser, requireEditor } from '../../../../lib';

// GET /api/books/:bookId/images/:key → serve image from R2
export const onRequestGet: PagesFunction<Env, 'bookId' | 'key', RequestData> = async (context) => {
  requireUser(context.data);

  const bookId = context.params.bookId as string;
  const key = `${bookId}/${decodeURIComponent(context.params.key as string)}`;

  const object = await context.env.BUCKET.get(key);
  if (!object) return error('Image not found', 404);

  const headers = new Headers();
  headers.set('Content-Type', object.httpMetadata?.contentType ?? 'application/octet-stream');
  headers.set('Cache-Control', 'public, max-age=31536000, immutable');

  return new Response(object.body, { headers });
};

// PUT /api/books/:bookId/images/:key → upload image to R2
export const onRequestPut: PagesFunction<Env, 'bookId' | 'key', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const bookId = context.params.bookId as string;
  const key = `${bookId}/${decodeURIComponent(context.params.key as string)}`;

  const contentType = context.request.headers.get('Content-Type') ?? 'application/octet-stream';

  await context.env.BUCKET.put(key, context.request.body, {
    httpMetadata: { contentType },
  });

  return new Response(JSON.stringify({ ok: true, key }), {
    status: 201,
    headers: { 'Content-Type': 'application/json' },
  });
};

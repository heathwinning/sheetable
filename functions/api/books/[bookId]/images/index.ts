import type { Env, RequestData } from '../../../../lib';
import { error, json, requireUser } from '../../../../lib';

export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async ({ env, params, data, request }) => {
  requireUser(data);
  const bookPath = `${params.bookId}/`;
  const cursor = new URL(request.url).searchParams.get('cursor') ?? undefined;
  const list = await env.BUCKET.list({ prefix: bookPath + bookPath, limit: 24, cursor }).catch(() => null);
  if (!list) return error('Unable to list workbook images', 400);
  return json({
    images: list.objects.map(object => object.key.slice(bookPath.length)),
    cursor: list.truncated ? list.cursor : undefined,
  });
};

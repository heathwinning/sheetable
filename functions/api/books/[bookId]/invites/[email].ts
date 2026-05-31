import type { Env, RequestData } from '../../../../lib';
import { json, requireUser, requireOwner } from '../../../../lib';

// DELETE /api/books/:bookId/invites/:email → cancel a pending invite
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'email', RequestData> = async (context) => {
  requireUser(context.data);
  requireOwner(context.data);

  const bookId = context.params.bookId as string;
  const email = decodeURIComponent(context.params.email as string);

  await context.env.DB.prepare(
    'DELETE FROM book_invites WHERE book_id = ? AND email = ?'
  ).bind(bookId, email).run();

  return json({ ok: true });
};

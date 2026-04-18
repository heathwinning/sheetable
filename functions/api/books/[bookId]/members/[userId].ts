import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireOwner } from '../../../../lib';

// DELETE /api/books/:bookId/members/:userId → remove member
export const onRequestDelete: PagesFunction<Env, 'bookId' | 'userId', RequestData> = async (context) => {
  requireUser(context.data);
  requireOwner(context.data);

  const bookId = context.params.bookId as string;
  const userId = context.params.userId as string;

  // Cannot remove the owner
  const book = await context.env.DB.prepare(
    'SELECT owner_id FROM books WHERE id = ?'
  ).bind(bookId).first<{ owner_id: string }>();

  if (book?.owner_id === userId) {
    return error('Cannot remove the book owner');
  }

  await context.env.DB.prepare(
    'DELETE FROM book_members WHERE book_id = ? AND user_id = ?'
  ).bind(bookId, userId).run();

  return json({ ok: true });
};

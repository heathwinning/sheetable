import type { Env, RequestData } from '../../../lib';
import { json, error, requireUser } from '../../../lib';

// GET /api/books/:bookId/invite → check invite status for current user
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  const bookId = context.params.bookId as string;

  const book = await context.env.DB.prepare(
    'SELECT name FROM books WHERE id = ?'
  ).bind(bookId).first<{ name: string }>();

  if (!book) return error('Book not found', 404);

  // If not signed in, just return book name
  if (!context.data.user) {
    return json({ bookName: book.name, status: 'sign-in-required' });
  }

  // Check if already a member
  const membership = await context.env.DB.prepare(
    'SELECT role FROM book_members WHERE book_id = ? AND user_id = ?'
  ).bind(bookId, context.data.user.id).first();

  if (membership) {
    return json({ bookName: book.name, status: 'already-member' });
  }

  // Check for pending invite
  const invite = await context.env.DB.prepare(
    'SELECT role FROM book_invites WHERE book_id = ? AND email = ?'
  ).bind(bookId, context.data.user.email).first<{ role: string }>();

  if (invite) {
    return json({ bookName: book.name, status: 'pending', role: invite.role });
  }

  return json({ bookName: book.name, status: 'no-invite' });
};

// POST /api/books/:bookId/invite/accept → accept a pending invite
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  const user = requireUser(context.data);
  const bookId = context.params.bookId as string;

  const invite = await context.env.DB.prepare(
    'SELECT role FROM book_invites WHERE book_id = ? AND email = ?'
  ).bind(bookId, user.email).first<{ role: string }>();

  if (!invite) return error('No pending invitation found', 404);

  // Add as member and delete invite in one batch
  await context.env.DB.batch([
    context.env.DB.prepare(
      `INSERT INTO book_members (book_id, user_id, role)
       VALUES (?, ?, ?)
       ON CONFLICT(book_id, user_id) DO UPDATE SET role = excluded.role`
    ).bind(bookId, user.id, invite.role),
    context.env.DB.prepare(
      'DELETE FROM book_invites WHERE book_id = ? AND email = ?'
    ).bind(bookId, user.email),
  ]);

  return json({ ok: true, role: invite.role });
};

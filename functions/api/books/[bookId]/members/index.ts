import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireOwner } from '../../../../lib';

// POST /api/books/:bookId/members → add or update a member
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireOwner(context.data);

  const bookId = context.params.bookId as string;
  const body = await context.request.json() as { email?: string; role?: string };

  if (!body.email?.trim()) return error('email is required');
  if (!body.role || !['editor', 'viewer'].includes(body.role)) {
    return error('role must be editor or viewer');
  }

  // Look up user by email
  const target = await context.env.DB.prepare(
    'SELECT id FROM users WHERE email = ?'
  ).bind(body.email.trim()).first<{ id: string }>();

  if (!target) {
    // User doesn't exist yet — create a pending invite
    await context.env.DB.prepare(
      `INSERT INTO book_invites (book_id, email, role, invited_by)
       VALUES (?, ?, ?, ?)
       ON CONFLICT(book_id, email) DO UPDATE SET role = excluded.role`
    ).bind(bookId, body.email.trim(), body.role, context.data.user!.id).run();

    return json({ ok: true, invited: true });
  }

  await context.env.DB.prepare(
    `INSERT INTO book_members (book_id, user_id, role)
     VALUES (?, ?, ?)
     ON CONFLICT(book_id, user_id) DO UPDATE SET role = excluded.role`
  ).bind(bookId, target.id, body.role).run();

  return json({ ok: true });
};

// GET /api/books/:bookId/members → list members + pending invites
export const onRequestGet: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  const bookId = context.params.bookId as string;

  const { results: members } = await context.env.DB.prepare(
    `SELECT u.id AS userId, u.email, u.name, bm.role
     FROM book_members bm
     JOIN users u ON u.id = bm.user_id
     WHERE bm.book_id = ?`
  ).bind(bookId).all();

  const { results: invites } = await context.env.DB.prepare(
    `SELECT email, role, created_at AS createdAt
     FROM book_invites
     WHERE book_id = ?`
  ).bind(bookId).all();

  return json({ members, invites });
};

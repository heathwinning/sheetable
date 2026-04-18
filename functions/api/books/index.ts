import type { Env, RequestData } from '../../lib';
import { json, error, requireUser } from '../../lib';

// GET /api/books → list books for current user
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  const user = requireUser(context.data);

  const { results } = await context.env.DB.prepare(
    `SELECT b.id, b.name, b.owner_id, bm.role, b.created_at
     FROM books b
     JOIN book_members bm ON bm.book_id = b.id
     WHERE bm.user_id = ?
     ORDER BY b.created_at DESC`
  ).bind(user.id).all();

  return json(results);
};

// POST /api/books → create a new book
export const onRequestPost: PagesFunction<Env, string, RequestData> = async (context) => {
  const user = requireUser(context.data);
  const body = await context.request.json() as { name?: string };

  if (!body.name?.trim()) return error('name is required');

  const id = crypto.randomUUID();

  await context.env.DB.batch([
    context.env.DB.prepare(
      'INSERT INTO books (id, name, owner_id) VALUES (?, ?, ?)'
    ).bind(id, body.name.trim(), user.id),
    context.env.DB.prepare(
      'INSERT INTO book_members (book_id, user_id, role) VALUES (?, ?, ?)'
    ).bind(id, user.id, 'owner'),
  ]);

  return json({ id, name: body.name.trim() }, 201);
};

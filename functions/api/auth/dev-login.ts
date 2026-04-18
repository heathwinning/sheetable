import type { Env, RequestData } from '../../lib';
import { signSession, setSessionCookie, error } from '../../lib';

// GET /api/auth/dev-login?key=<DEV_LOGIN_KEY> → create test user + sign in
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  const url = new URL(context.request.url);

  // Only allow on localhost; in non-local environments require DEV_LOGIN_KEY
  const isLocal = url.hostname === 'localhost' || url.hostname === '127.0.0.1';
  if (!isLocal) {
    const key = url.searchParams.get('key');
    if (!context.env.DEV_LOGIN_KEY || key !== context.env.DEV_LOGIN_KEY) {
      return error('Forbidden', 403);
    }
  }

  const testUser = {
    id: 'dev-test-user',
    email: 'test@sheetable.dev',
    name: 'Test User',
  };

  // Upsert test user
  await context.env.DB.prepare(
    `INSERT INTO users (id, email, name)
     VALUES (?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET email = excluded.email, name = excluded.name`
  ).bind(testUser.id, testUser.email, testUser.name).run();

  const token = await signSession(testUser, context.env.SESSION_SECRET);
  const response = Response.redirect(`${url.origin}/`, 302);
  return setSessionCookie(response, token);
};

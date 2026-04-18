import type { Env, RequestData } from '../lib';
import { verifySession, getSessionCookie, error } from '../lib';

// Auth middleware: attaches user to context.data for all /api/* routes.
// Public paths (login, callback) skip auth checks.
export const onRequest: PagesFunction<Env, string, RequestData> = async (context) => {
  const url = new URL(context.request.url);
  const path = url.pathname;

  // Public auth endpoints don't need a session
  if (path === '/api/auth/login' || path === '/api/auth/callback' || path === '/api/auth/dev-login') {
    return context.next();
  }

  // Invite check endpoint is semi-public (works with or without auth)
  const isInviteGet = context.request.method === 'GET' && /^\/api\/books\/[^/]+\/invite$/.test(path);

  const cookie = getSessionCookie(context.request);
  if (!cookie) {
    if (isInviteGet) return context.next();
    return error('Unauthorized', 401);
  }

  const user = await verifySession(cookie, context.env.SESSION_SECRET);
  if (!user) {
    if (isInviteGet) return context.next();
    return error('Unauthorized', 401);
  }

  context.data.user = user;

  // If this is a book-scoped route, check membership
  // Skip membership check for invite accept endpoint (user may not be a member yet)
  const isInviteAccept = /^\/api\/books\/[^/]+\/invite$/.test(path);
  const bookMatch = path.match(/^\/api\/books\/([^/]+)(?:\/|$)/);
  if (bookMatch && !isInviteAccept) {
    const bookId = decodeURIComponent(bookMatch[1]);
    const membership = await context.env.DB.prepare(
      'SELECT role FROM book_members WHERE book_id = ? AND user_id = ?'
    ).bind(bookId, user.id).first<{ role: string }>();

    if (!membership) {
      return error('Forbidden', 403);
    }
    context.data.bookRole = membership.role as 'owner' | 'editor' | 'viewer';
  }

  try {
    return await context.next();
  } catch (e) {
    if (e instanceof Response) return e;
    throw e;
  }
};

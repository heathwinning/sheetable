import type { Env, RequestData } from '../../lib';
import { signSession, setSessionCookie, error } from '../../lib';

const STATE_COOKIE = 'sheetable_oauth_state';

interface GoogleTokenResponse {
  access_token: string;
  id_token: string;
  token_type: string;
}

interface GoogleUserInfo {
  sub: string;
  email: string;
  name: string;
}

// GET /api/auth/callback → exchange Google auth code for session
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  const url = new URL(context.request.url);
  const code = url.searchParams.get('code');
  if (!code) return error('Missing code parameter');

  // Verify CSRF state
  const state = url.searchParams.get('state');
  const cookieHeader = context.request.headers.get('Cookie') ?? '';
  const stateMatch = cookieHeader.match(new RegExp(`(?:^|;\\s*)${STATE_COOKIE}=([^;]+)`));
  const savedState = stateMatch ? stateMatch[1] : null;
  if (!state || !savedState || state !== savedState) {
    return error('Invalid OAuth state', 403);
  }

  const redirectUri = `${url.origin}/api/auth/callback`;

  // Exchange code for tokens
  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: context.env.GOOGLE_CLIENT_ID,
      client_secret: context.env.GOOGLE_CLIENT_SECRET,
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
    }),
  });

  if (!tokenRes.ok) {
    return error('Failed to exchange auth code', 502);
  }

  const tokens = await tokenRes.json() as GoogleTokenResponse;

  // Get user info
  const userRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
    headers: { Authorization: `Bearer ${tokens.access_token}` },
  });

  if (!userRes.ok) {
    return error('Failed to fetch user info', 502);
  }

  const userInfo = await userRes.json() as GoogleUserInfo;

  // Upsert user in DB
  await context.env.DB.prepare(
    `INSERT INTO users (id, email, name)
     VALUES (?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET email = excluded.email, name = excluded.name`
  ).bind(userInfo.sub, userInfo.email, userInfo.name).run();

  // Redeem any pending book invitations for this email
  const { results: invites } = await context.env.DB.prepare(
    'SELECT book_id, role FROM book_invites WHERE email = ?'
  ).bind(userInfo.email).all<{ book_id: string; role: string }>();

  if (invites.length > 0) {
    const stmts = invites.map(inv =>
      context.env.DB.prepare(
        `INSERT INTO book_members (book_id, user_id, role)
         VALUES (?, ?, ?)
         ON CONFLICT(book_id, user_id) DO UPDATE SET role = excluded.role`
      ).bind(inv.book_id, userInfo.sub, inv.role)
    );
    stmts.push(
      context.env.DB.prepare(
        'DELETE FROM book_invites WHERE email = ?'
      ).bind(userInfo.email)
    );
    try {
      await context.env.DB.batch(stmts);
    } catch (e) {
      console.error('Failed to redeem invites for', userInfo.email, e);
      // Non-fatal: allow login to proceed
    }
  }

  // Sign session cookie
  const sessionUser = {
    id: userInfo.sub,
    email: userInfo.email,
    name: userInfo.name,
  };
  const token = await signSession(sessionUser, context.env.SESSION_SECRET);

  // Redirect to app root with session cookie set; clear OAuth state cookie
  const response = new Response(null, {
    status: 302,
    headers: { Location: `${url.origin}/` },
  });
  response.headers.append(
    'Set-Cookie',
    `${STATE_COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`,
  );
  return setSessionCookie(response, token);
};

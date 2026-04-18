import type { Env, RequestData } from '../../lib';

const STATE_COOKIE = 'sheetable_oauth_state';

// GET /api/auth/login → redirect to Google OAuth consent screen
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  const url = new URL(context.request.url);
  const redirectUri = `${url.origin}/api/auth/callback`;

  const state = crypto.randomUUID();

  const params = new URLSearchParams({
    client_id: context.env.GOOGLE_CLIENT_ID,
    redirect_uri: redirectUri,
    response_type: 'code',
    scope: 'openid email profile',
    access_type: 'online',
    prompt: 'select_account',
    state,
  });

  const response = new Response(null, {
    status: 302,
    headers: { Location: `https://accounts.google.com/o/oauth2/v2/auth?${params}` },
  });
  response.headers.append(
    'Set-Cookie',
    `${STATE_COOKIE}=${state}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=600`,
  );
  return response;
};

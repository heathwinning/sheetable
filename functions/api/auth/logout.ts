import type { Env, RequestData } from '../../lib';
import { clearSessionCookie } from '../../lib';

// POST /api/auth/logout → clear session cookie
export const onRequestPost: PagesFunction<Env, string, RequestData> = async () => {
  return clearSessionCookie();
};

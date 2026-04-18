import type { Env, RequestData } from '../../lib';
import { json, requireUser } from '../../lib';

// GET /api/auth/me → return current user info
export const onRequestGet: PagesFunction<Env, string, RequestData> = async (context) => {
  const user = requireUser(context.data);
  return json(user);
};

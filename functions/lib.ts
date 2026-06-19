export interface Env {
  DB: D1Database;
  BUCKET: R2Bucket;
  GOOGLE_CLIENT_ID: string;
  GOOGLE_CLIENT_SECRET: string;
  GOOGLE_API_KEY: string;
  SESSION_SECRET: string;
  DEV_LOGIN_KEY: string;
  CRON_SECRET?: string;
}

export interface SessionUser {
  id: string;
  email: string;
  name: string;
}

export type BookRole = 'owner' | 'editor' | 'viewer';

// Attached by auth middleware via context.data
export interface RequestData {
  user?: SessionUser;
  bookRole?: BookRole;
}

// Helper: JSON response
export function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

// Helper: error response
export function error(message: string, status = 400): Response {
  return json({ error: message }, status);
}

// ---- Session cookie (HMAC-signed JSON) ----

const COOKIE_NAME = 'sheetable_session';
const encoder = new TextEncoder();

async function getSigningKey(secret: string): Promise<CryptoKey> {
  return crypto.subtle.importKey(
    'raw',
    encoder.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
  false,
    ['sign', 'verify'],
  );
}

function toHex(buf: ArrayBuffer): string {
  return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, '0')).join('');
}

function fromHex(hex: string): Uint8Array {
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < hex.length; i += 2) {
    bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
  }
  return bytes;
}

export async function signSession(user: SessionUser, secret: string): Promise<string> {
  const payload = JSON.stringify(user);
  const key = await getSigningKey(secret);
  const sig = await crypto.subtle.sign('HMAC', key, encoder.encode(payload));
  return `${toHex(sig)}.${btoa(payload)}`;
}

export async function verifySession(cookie: string, secret: string): Promise<SessionUser | null> {
  const dot = cookie.indexOf('.');
  if (dot < 0) return null;
  const sigHex = cookie.substring(0, dot);
  const payloadB64 = cookie.substring(dot + 1);
  try {
    const payload = atob(payloadB64);
    const key = await getSigningKey(secret);
    const valid = await crypto.subtle.verify('HMAC', key, fromHex(sigHex), encoder.encode(payload));
    if (!valid) return null;
    return JSON.parse(payload) as SessionUser;
  } catch {
    return null;
  }
}

export function setSessionCookie(response: Response, token: string): Response {
  const headers = new Headers(response.headers);
  headers.append('Set-Cookie', `${COOKIE_NAME}=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=2592000`);
  return new Response(response.body, { status: response.status, headers });
}

export function clearSessionCookie(): Response {
  return new Response(JSON.stringify({ ok: true }), {
    status: 200,
    headers: {
      'Content-Type': 'application/json',
      'Set-Cookie': `${COOKIE_NAME}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`,
    },
  });
}

export function getSessionCookie(request: Request): string | null {
  const header = request.headers.get('Cookie') ?? '';
  const match = header.match(new RegExp(`(?:^|;\\s*)${COOKIE_NAME}=([^;]+)`));
  return match ? match[1] : null;
}

// Helper: require the user to be authenticated (from context.data)
export function requireUser(data: RequestData): SessionUser {
  if (!data.user) throw new Response(JSON.stringify({ error: 'Unauthorized' }), { status: 401, headers: { 'Content-Type': 'application/json' } });
  return data.user;
}

// Helper: require at least editor role on a book
export function requireEditor(data: RequestData): void {
  if (data.bookRole !== 'owner' && data.bookRole !== 'editor') {
    throw new Response(JSON.stringify({ error: 'Forbidden' }), { status: 403, headers: { 'Content-Type': 'application/json' } });
  }
}

// Helper: require owner role
export function requireOwner(data: RequestData): void {
  if (data.bookRole !== 'owner') {
    throw new Response(JSON.stringify({ error: 'Forbidden: owner only' }), { status: 403, headers: { 'Content-Type': 'application/json' } });
  }
}

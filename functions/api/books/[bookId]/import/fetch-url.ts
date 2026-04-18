import type { Env, RequestData } from '../../../../lib';
import { json, error, requireUser, requireEditor } from '../../../../lib';

// POST /api/books/:bookId/import/fetch-url → fetch CSV text from a remote URL
export const onRequestPost: PagesFunction<Env, 'bookId', RequestData> = async (context) => {
  requireUser(context.data);
  requireEditor(context.data);

  const body = await context.request.json() as { url: string };
  const url = body.url?.trim();
  if (!url) return error('url is required');

  // Validate URL format
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return error('Invalid URL');
  }

  // Only allow http/https
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    return error('Only http and https URLs are supported');
  }

  // Block private/internal IPs
  const hostname = parsed.hostname.toLowerCase();
  if (
    hostname === 'localhost' ||
    hostname === '127.0.0.1' ||
    hostname === '::1' ||
    hostname === '0.0.0.0' ||
    hostname.endsWith('.local') ||
    hostname.startsWith('10.') ||
    hostname.startsWith('192.168.') ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(hostname)
  ) {
    return error('Cannot fetch from private/internal addresses');
  }

  try {
    const resp = await fetch(url, {
      headers: { 'Accept': 'text/csv, text/plain, */*' },
      redirect: 'follow',
    });

    if (!resp.ok) {
      return error(`Failed to fetch URL: ${resp.status} ${resp.statusText}`);
    }

    // Limit response size (10 MB)
    const contentLength = resp.headers.get('content-length');
    if (contentLength && parseInt(contentLength, 10) > 10 * 1024 * 1024) {
      return error('File is too large (max 10 MB)');
    }

    const csvText = await resp.text();

    if (csvText.length > 10 * 1024 * 1024) {
      return error('File is too large (max 10 MB)');
    }

    return json({ csvText });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    return error(`Failed to fetch URL: ${message}`);
  }
};

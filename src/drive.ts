// Google Drive API integration
// Uses Google Identity Services (GIS) for OAuth2

const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/spreadsheets.readonly';
const DISCOVERY_DOC = 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest';
const SHEETS_DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4';

let tokenClient: google.accounts.oauth2.TokenClient | null = null;
let accessToken: string | null = null;

const STORAGE_KEY = 'sheetable_auth';

interface StoredAuth {
  token: string;
  expiresAt: number;
  userInfo: UserInfo | null;
}

export function saveAuth(token: string, expiresIn: number, userInfo: UserInfo | null): void {
  const data: StoredAuth = {
    token,
    expiresAt: Date.now() + expiresIn * 1000 - 60000, // 1 min buffer
    userInfo,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function loadAuth(): StoredAuth | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const data: StoredAuth = JSON.parse(raw);
    if (Date.now() >= data.expiresAt) {
      localStorage.removeItem(STORAGE_KEY);
      return null;
    }
    return data;
  } catch {
    localStorage.removeItem(STORAGE_KEY);
    return null;
  }
}

function clearAuth(): void {
  localStorage.removeItem(STORAGE_KEY);
}

// Try to restore a valid token from localStorage
export function tryRestoreToken(): { token: string; userInfo: UserInfo | null } | null {
  const stored = loadAuth();
  if (stored) {
    accessToken = stored.token;
    // Also set the token on the gapi client so gapi.client.drive calls work
    try {
      gapi.client.setToken({ access_token: stored.token });
    } catch { /* gapi not ready yet */ }
    return { token: stored.token, userInfo: stored.userInfo };
  }
  return null;
}

// These will be set from environment
let CLIENT_ID = '';

export function setClientId(id: string) {
  CLIENT_ID = id;
}

export function getClientId(): string {
  return CLIENT_ID;
}

export function isAuthenticated(): boolean {
  return accessToken !== null;
}

export function getAccessToken(): string | null {
  // Prefer the token from gapi client (may be refreshed)
  try {
    const gapiToken = gapi?.client?.getToken?.()?.access_token;
    if (gapiToken) {
      accessToken = gapiToken;
      return gapiToken;
    }
  } catch { /* gapi not loaded yet */ }
  return accessToken;
}

// Wait for a global to be defined (async script loading)
function waitForGlobal<T>(name: string, timeoutMs = 10000): Promise<T> {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const check = () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const val = (window as any)[name];
      if (val) {
        resolve(val);
      } else if (Date.now() - start > timeoutMs) {
        reject(new Error(`Timed out waiting for ${name} to load`));
      } else {
        setTimeout(check, 100);
      }
    };
    check();
  });
}

// Initialize GAPI client
export async function initGapi(): Promise<void> {
  await waitForGlobal('gapi');
  return new Promise((resolve, reject) => {
    gapi.load('client', async () => {
      try {
        await gapi.client.init({
          discoveryDocs: [DISCOVERY_DOC, SHEETS_DISCOVERY_DOC],
        });
        resolve();
      } catch (err) {
        reject(err);
      }
    });
  });
}

// Initialize GIS token client
export async function initTokenClient(onTokenReceived: (expiresIn: number) => void): Promise<void> {
  await waitForGlobal('google');
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: (response: google.accounts.oauth2.TokenResponse) => {
      if (response.error) {
        console.error('Token error:', response.error);
        return;
      }
      accessToken = response.access_token;
      const expiresIn = parseInt(String(response.expires_in), 10) || 3600;
      onTokenReceived(expiresIn);
    },
  });
}

// Request access token (triggers consent flow if needed)
export function requestAccessToken(): void {
  if (!tokenClient) {
    console.error('Token client not initialized');
    return;
  }
  tokenClient.requestAccessToken({ prompt: '' });
}

export function signOut(): void {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {
      accessToken = null;
    });
  }
  clearAuth();
}

export interface UserInfo {
  name: string;
  picture: string;
  email?: string;
}

export async function getUserInfo(): Promise<UserInfo | null> {
  const token = getAccessToken();
  if (!token) return null;
  try {
    const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return null;
    const data = await res.json();
    return { name: data.name, picture: data.picture, email: data.email };
  } catch {
    return null;
  }
}

// Find or create the "sheetable" folder in the user's Drive root
const FOLDER_NAME = 'sheetable';

export async function findOrCreateFolder(): Promise<{ id: string; name: string }> {
  // Search for existing folder by name
  const response = await gapi.client.drive.files.list({
    q: `name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
    pageSize: 1,
  });

  const existing = response.result.files?.[0];
  if (existing) {
    return { id: existing.id!, name: existing.name! };
  }

  // Create the folder
  const id = await createFolder(FOLDER_NAME);
  return { id, name: FOLDER_NAME };
}

// List files in a folder
export async function listFilesInFolder(folderId: string): Promise<Array<{ id: string; name: string; mimeType: string }>> {
  const response = await gapi.client.drive.files.list({
    q: `'${folderId}' in parents and trashed=false`,
    fields: 'files(id, name, mimeType)',
    pageSize: 100,
  });

  return (response.result.files ?? []) as Array<{ id: string; name: string; mimeType: string }>;
}

// Download file content
export async function downloadFile(fileId: string): Promise<string> {
  const response = await gapi.client.drive.files.get({
    fileId,
    alt: 'media',
  });
  return response.body;
}

// Create a new file
export async function createFile(name: string, content: string, folderId: string, mimeType = 'text/csv'): Promise<string> {
  const metadata = {
    name,
    mimeType,
    parents: [folderId],
  };

  const boundary = 'sheetable_boundary';
  const body =
    `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
    JSON.stringify(metadata) +
    `\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n` +
    content +
    `\r\n--${boundary}--`;

  const response = await gapi.client.request({
    path: '/upload/drive/v3/files',
    method: 'POST',
    params: { uploadType: 'multipart' },
    headers: {
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body,
  });

  return response.result.id as string;
}

// Update file content
export async function updateFile(fileId: string, content: string, mimeType = 'text/csv'): Promise<void> {
  await gapi.client.request({
    path: `/upload/drive/v3/files/${encodeURIComponent(fileId)}`,
    method: 'PATCH',
    params: { uploadType: 'media' },
    headers: {
      'Content-Type': mimeType,
    },
    body: content,
  });
}

// Create a folder
export async function createFolder(name: string, parentId?: string): Promise<string> {
  const metadata: Record<string, unknown> = {
    name,
    mimeType: 'application/vnd.google-apps.folder',
  };
  if (parentId) {
    metadata.parents = [parentId];
  }

  const response = await gapi.client.drive.files.create({
    resource: metadata,
    fields: 'id',
  });

  return response.result.id as string;
}

// Delete (trash) a file on Drive
export async function deleteFile(fileId: string): Promise<void> {
  await gapi.client.drive.files.update({
    fileId,
    resource: { trashed: true },
  });
}

// Rename a Drive file or folder
export async function renameFile(fileId: string, name: string): Promise<void> {
  await gapi.client.drive.files.update({
    fileId,
    resource: { name },
  });
}

// Find or create a subfolder inside a parent folder
export async function findOrCreateSubfolder(name: string, parentId: string): Promise<string> {
  const response = await gapi.client.drive.files.list({
    q: `name='${name}' and '${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id)',
    pageSize: 1,
  });

  const existing = response.result.files?.[0];
  if (existing) return existing.id!;

  return createFolder(name, parentId);
}

export interface DriveFolderEntry {
  id: string;
  name: string;
}

export interface DriveFileMeta {
  id: string;
  name: string;
  mimeType: string;
}

// List immediate subfolders in a folder
export async function listSubfolders(parentId: string): Promise<DriveFolderEntry[]> {
  const response = await gapi.client.drive.files.list({
    q: `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'files(id, name)',
    pageSize: 100,
    orderBy: 'name_natural',
  });
  return (response.result.files ?? [])
    .filter(f => f.id && f.name)
    .map(f => ({ id: f.id!, name: f.name! }));
}

// Share a Drive folder with a user by email
export async function shareFolderWithEmail(folderId: string, emailAddress: string, role: 'reader' | 'writer' = 'writer'): Promise<void> {
  await gapi.client.request({
    path: `/drive/v3/files/${folderId}/permissions`,
    method: 'POST',
    params: {
      sendNotificationEmail: 'true',
    },
    body: JSON.stringify({
      type: 'user',
      role,
      emailAddress,
    }),
  });
}

// Optional helper for receiver flows: create a shortcut to a folder/file in a target folder
export async function createShortcut(targetId: string, parentFolderId: string, name?: string): Promise<string> {
  const response = await gapi.client.drive.files.create({
    resource: {
      name,
      mimeType: 'application/vnd.google-apps.shortcut',
      parents: [parentFolderId],
      shortcutDetails: { targetId },
    },
    fields: 'id',
  });
  return response.result.id as string;
}

export async function getFileMeta(fileId: string): Promise<DriveFileMeta> {
  const response = await gapi.client.request({
    path: `/drive/v3/files/${encodeURIComponent(fileId)}`,
    method: 'GET',
    params: { fields: 'id,name,mimeType' },
  });

  return {
    id: response.result.id as string,
    name: response.result.name as string,
    mimeType: response.result.mimeType as string,
  };
}

export async function findShortcutInFolder(parentFolderId: string, targetId: string): Promise<string | null> {
  const response = await gapi.client.drive.files.list({
    q: `'${parentFolderId}' in parents and mimeType='application/vnd.google-apps.shortcut' and shortcutDetails.targetId='${targetId}' and trashed=false`,
    fields: 'files(id)',
    pageSize: 1,
  });

  const existing = response.result.files?.[0];
  return existing?.id ?? null;
}

export async function ensureShortcutInFolder(parentFolderId: string, targetId: string, name?: string): Promise<string> {
  const existing = await findShortcutInFolder(parentFolderId, targetId);
  if (existing) return existing;
  return createShortcut(targetId, parentFolderId, name);
}

// Upload a binary file (e.g. image) to Drive
export async function uploadBinaryFile(file: File, folderId: string): Promise<string> {
  const metadata = {
    name: file.name,
    mimeType: file.type,
    parents: [folderId],
  };

  // Google Drive multipart upload requires multipart/related, not multipart/form-data
  const boundary = 'sheetable_upload_boundary';
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelimiter = '\r\n--' + boundary + '--';

  // Read file as base64
  const base64Data = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(',')[1]); // strip data URL prefix
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

  const body =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: ' + file.type + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    base64Data +
    closeDelimiter;

  const token = getAccessToken();
  if (!token) throw new Error('Not authenticated');

  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'multipart/related; boundary=' + boundary,
    },
    body,
  });

  if (!res.ok) throw new Error(`Upload failed: ${res.status}`);
  const data = await res.json();
  return data.id as string;
}

// Get a thumbnail URL for a Drive file
export function getThumbnailUrl(fileId: string, size = 200): string {
  return `https://drive.google.com/thumbnail?id=${fileId}&sz=s${size}`;
}

// Export a Google Sheet as CSV
export interface SheetTab {
  sheetId: number;
  title: string;
}

export async function getSpreadsheetSheets(spreadsheetId: string): Promise<SheetTab[]> {
  const token = getAccessToken();
  if (!token) throw new Error('Not authenticated');
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?fields=${encodeURIComponent('sheets.properties(sheetId,title)')}`,
    { headers: { Authorization: `Bearer ${token}` } },
  );
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Failed to fetch spreadsheet: ${res.status} ${text}`);
  }
  const data = await res.json();
  return (data.sheets ?? []).map((s: { properties: { sheetId: number; title: string } }) => ({
    sheetId: s.properties.sheetId,
    title: s.properties.title,
  }));
}

export async function exportSheetAsCSV(spreadsheetId: string, gid?: number): Promise<string> {
  const token = getAccessToken();
  if (!token) throw new Error('Not authenticated');
  let url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=csv`;
  if (gid !== undefined) {
    url += `&gid=${gid}`;
  }
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Failed to export sheet: ${res.status}`);
  return res.text();
}

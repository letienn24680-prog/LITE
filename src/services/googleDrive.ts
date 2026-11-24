export const GOOGLE_TOKEN_STORAGE_KEY = 'lite_google_token';
export const GOOGLE_TOKEN_STORAGE_KEY_PREFIX = 'lite_gdrive_token_';
export const GOOGLE_TOKEN_CHANGED_EVENT = 'lite:gdrive-token-changed';
const GOOGLE_FOLDER_STORAGE_KEY_PREFIX = 'lite_gdrive_folder_';
const GOOGLE_FILEMAP_STORAGE_KEY_PREFIX = 'lite_gdrive_filemap_';
const LITE_WORKSPACE_FOLDER_NAME = 'LITE Workspace';

export type GoogleDriveTargetExt = 'docx' | 'xlsx' | 'pptx' | 'pdf';

export interface GoogleDriveFileItem {
  id: string;
  name: string;
  mimeType: string;
  modifiedTime?: string;
  size?: number;
  iconLink?: string;
  webViewLink?: string;
  isGoogleWorkspaceType: boolean;
  targetExtension: GoogleDriveTargetExt;
}

export class GoogleDriveAuthError extends Error {
  readonly statusCode?: number;

  constructor(message: string, statusCode?: number) {
    super(message);
    this.name = 'GoogleDriveAuthError';
    this.statusCode = statusCode;
  }
}

const GOOGLE_DRIVE_FILES_URL = 'https://www.googleapis.com/drive/v3/files';
const GOOGLE_DRIVE_UPLOAD_URL =
  'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink';

const SUPPORTED_DRIVE_MIME_TYPES = [
  'application/vnd.google-apps.document',
  'application/vnd.google-apps.spreadsheet',
  'application/vnd.google-apps.presentation',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'text/csv',
  'application/pdf',
] as const;

const WORKSPACE_EXPORT_MAP: Record<
  string,
  { targetExt: GoogleDriveTargetExt; exportMime: string }
> = {
  'application/vnd.google-apps.document': {
    targetExt: 'docx',
    exportMime:
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  },
  'application/vnd.google-apps.spreadsheet': {
    targetExt: 'xlsx',
    exportMime:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  },
  'application/vnd.google-apps.presentation': {
    targetExt: 'pptx',
    exportMime:
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  },
};

const BINARY_MIME_TARGET_MAP: Record<string, GoogleDriveTargetExt> = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
    'docx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation':
    'pptx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.ms-excel': 'xlsx',
  'text/csv': 'xlsx',
  'application/pdf': 'pdf',
};

const FALLBACK_MIME_BY_EXT: Record<GoogleDriveTargetExt, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
};

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

let activeGoogleDriveUserId: string | null = null;

function normalizeUserId(userId?: string | null): string | null {
  const trimmed = userId?.trim() ?? '';
  return trimmed.length > 0 ? trimmed : null;
}

function getScopedStorageKey(userId?: string | null): string {
  const effectiveUserId = normalizeUserId(userId) ?? activeGoogleDriveUserId;
  return effectiveUserId
    ? `${GOOGLE_TOKEN_STORAGE_KEY_PREFIX}${effectiveUserId}`
    : GOOGLE_TOKEN_STORAGE_KEY;
}

function removeOtherDriveTokenKeys(keepKey: string | null): void {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  const keysToRemove: string[] = [];
  for (let i = 0; i < window.localStorage.length; i += 1) {
    const key = window.localStorage.key(i);
    if (!key) continue;
    const isDriveKey =
      key === GOOGLE_TOKEN_STORAGE_KEY ||
      key.startsWith(GOOGLE_TOKEN_STORAGE_KEY_PREFIX);
    if (isDriveKey && key !== keepKey) {
      keysToRemove.push(key);
    }
  }
  for (const key of keysToRemove) {
    window.localStorage.removeItem(key);
  }
}

function emitGoogleTokenChanged(hasToken: boolean, userId: string | null): void {
  if (typeof window === 'undefined') {
    return;
  }
  try {
    window.dispatchEvent(
      new CustomEvent(GOOGLE_TOKEN_CHANGED_EVENT, {
        detail: { hasToken, userId },
      })
    );
  } catch {
    // Ignore CustomEvent dispatch errors
  }
}

export function setActiveGoogleDriveUser(userId: string | null): void {
  const nextUserId = normalizeUserId(userId);
  activeGoogleDriveUserId = nextUserId;
  try {
    const keepKey = nextUserId
      ? `${GOOGLE_TOKEN_STORAGE_KEY_PREFIX}${nextUserId}`
      : GOOGLE_TOKEN_STORAGE_KEY;
    removeOtherDriveTokenKeys(keepKey);
    emitGoogleTokenChanged(Boolean(getStoredGoogleDriveToken(nextUserId)), nextUserId);
  } catch {
    // Ignore storage restrictions
  }
}

export function getStoredGoogleDriveToken(userId?: string | null): string | null {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return null;
    }
    const key = getScopedStorageKey(userId);
    const raw = window.localStorage.getItem(key);
    const trimmed = raw?.trim() ?? '';
    return trimmed.length > 0 ? trimmed : null;
  } catch {
    return null;
  }
}

export function setStoredGoogleDriveToken(
  token: string | null,
  userId?: string | null
): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    const normalizedUser = normalizeUserId(userId);
    if (normalizedUser) {
      activeGoogleDriveUserId = normalizedUser;
    }
    const key = getScopedStorageKey(normalizedUser);
    removeOtherDriveTokenKeys(key);

    const trimmed = token?.trim() ?? '';
    if (trimmed.length > 0) {
      window.localStorage.setItem(key, trimmed);
      emitGoogleTokenChanged(true, normalizedUser ?? activeGoogleDriveUserId);
    } else {
      window.localStorage.removeItem(key);
      emitGoogleTokenChanged(false, normalizedUser ?? activeGoogleDriveUserId);
    }
  } catch {
    // Ignore storage quota or private mode restrictions
  }
}

export function clearStoredGoogleDriveToken(userId?: string | null): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    const normalizedUser = normalizeUserId(userId);
    if (normalizedUser) {
      window.localStorage.removeItem(
        `${GOOGLE_TOKEN_STORAGE_KEY_PREFIX}${normalizedUser}`
      );
      window.localStorage.removeItem(GOOGLE_TOKEN_STORAGE_KEY);
      if (activeGoogleDriveUserId === normalizedUser) {
        activeGoogleDriveUserId = null;
      }
      emitGoogleTokenChanged(false, normalizedUser);
      return;
    }
    activeGoogleDriveUserId = null;
    removeOtherDriveTokenKeys(null);
    emitGoogleTokenChanged(false, null);
  } catch {
    // Ignore storage restrictions
  }
}

/**
 * Clears Google Drive folderId and fileId map caches in localStorage for a logged-out user
 * (or all non-guest scopes) so no previous account metadata remains on the device.
 */
export function clearGoogleDriveUserCache(userId?: string | null): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    const normalizedUser = normalizeUserId(userId);
    if (normalizedUser) {
      window.localStorage.removeItem(`${GOOGLE_FOLDER_STORAGE_KEY_PREFIX}${normalizedUser}`);
      window.localStorage.removeItem(`${GOOGLE_FILEMAP_STORAGE_KEY_PREFIX}${normalizedUser}`);
    }

    const keysToRemove: string[] = [];
    for (let i = 0; i < window.localStorage.length; i += 1) {
      const key = window.localStorage.key(i);
      if (!key) continue;
      if (
        (key.startsWith(GOOGLE_FOLDER_STORAGE_KEY_PREFIX) &&
          key !== `${GOOGLE_FOLDER_STORAGE_KEY_PREFIX}guest`) ||
        (key.startsWith(GOOGLE_FILEMAP_STORAGE_KEY_PREFIX) &&
          key !== `${GOOGLE_FILEMAP_STORAGE_KEY_PREFIX}guest`)
      ) {
        keysToRemove.push(key);
      }
    }
    for (const key of keysToRemove) {
      window.localStorage.removeItem(key);
    }
  } catch {
    // Ignore storage restrictions
  }
}

function requireAccessToken(userId?: string | null): string {
  const token = getStoredGoogleDriveToken(userId);
  if (!token) {
    throw new GoogleDriveAuthError(
      'Chưa có quyền truy cập Google Drive. Vui lòng kết nối tài khoản Google.'
    );
  }
  return token;
}

function resolveEffectiveUserScope(userId?: string | null): string {
  return normalizeUserId(userId) ?? activeGoogleDriveUserId ?? 'guest';
}

function getFolderStorageKey(userId?: string | null): string {
  return `${GOOGLE_FOLDER_STORAGE_KEY_PREFIX}${resolveEffectiveUserScope(userId)}`;
}

function getFileMapStorageKey(userId?: string | null): string {
  return `${GOOGLE_FILEMAP_STORAGE_KEY_PREFIX}${resolveEffectiveUserScope(userId)}`;
}

function readDriveFileMap(userId?: string | null): Record<string, string> {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return {};
    }
    const raw = window.localStorage.getItem(getFileMapStorageKey(userId));
    if (!raw) {
      return {};
    }
    const parsed: unknown = JSON.parse(raw);
    if (!isPlainRecord(parsed)) {
      return {};
    }
    const clean: Record<string, string> = {};
    for (const [docId, driveId] of Object.entries(parsed)) {
      if (typeof driveId === 'string' && driveId.trim().length > 0) {
        clean[docId] = driveId.trim();
      }
    }
    return clean;
  } catch {
    return {};
  }
}

function writeDriveFileMap(
  map: Record<string, string>,
  userId?: string | null
): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    window.localStorage.setItem(
      getFileMapStorageKey(userId),
      JSON.stringify(map)
    );
  } catch {
    // Ignore storage quota or private browsing errors
  }
}

export function getMappedDriveFileId(
  docId: string,
  userId?: string | null
): string | null {
  const cleanDocId = docId.trim();
  if (!cleanDocId) {
    return null;
  }
  const map = readDriveFileMap(userId);
  return map[cleanDocId] ?? null;
}

export function setMappedDriveFileId(
  docId: string,
  driveFileId: string,
  userId?: string | null
): void {
  const cleanDocId = docId.trim();
  const cleanDriveId = driveFileId.trim();
  if (!cleanDocId || !cleanDriveId) {
    return;
  }
  const map = readDriveFileMap(userId);
  map[cleanDocId] = cleanDriveId;
  writeDriveFileMap(map, userId);
}

export function clearMappedDriveFileId(
  docId: string,
  userId?: string | null
): void {
  const cleanDocId = docId.trim();
  if (!cleanDocId) {
    return;
  }
  const map = readDriveFileMap(userId);
  if (cleanDocId in map) {
    delete map[cleanDocId];
    writeDriveFileMap(map, userId);
  }
}

function escapeDriveQueryValue(input: string): string {
  return input.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function inferTargetExtension(
  mimeType: string,
  fileName: string
): GoogleDriveTargetExt {
  const workspaceConfig = WORKSPACE_EXPORT_MAP[mimeType];
  if (workspaceConfig) {
    return workspaceConfig.targetExt;
  }
  const binaryMapped = BINARY_MIME_TARGET_MAP[mimeType];
  if (binaryMapped) {
    return binaryMapped;
  }
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.pptx')) {
    return 'pptx';
  }
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv')) {
    return 'xlsx';
  }
  if (lower.endsWith('.pdf')) {
    return 'pdf';
  }
  return 'docx';
}

function ensureSafeFileNameWithExt(file: GoogleDriveFileItem): string {
  const cleanBase = file.name.trim() || `drive-document-${file.id}`;
  const lower = cleanBase.toLowerCase();

  if (file.isGoogleWorkspaceType) {
    const ext = `.${file.targetExtension}`;
    return lower.endsWith(ext) ? cleanBase : `${cleanBase}${ext}`;
  }

  const validBinaryExts = ['.docx', '.pptx', '.xlsx', '.xls', '.csv', '.pdf'];
  if (validBinaryExts.some((ext) => lower.endsWith(ext))) {
    return cleanBase;
  }

  return `${cleanBase}.${file.targetExtension}`;
}

async function extractGoogleApiError(response: Response): Promise<string> {
  try {
    const payload: unknown = await response.json();
    if (isPlainRecord(payload) && isPlainRecord(payload.error)) {
      const msg = payload.error.message;
      if (typeof msg === 'string' && msg.trim().length > 0) {
        return msg.trim();
      }
    }
  } catch {
    // Fallback to HTTP status text
  }
  return `Lỗi Google Drive API (HTTP ${response.status})`;
}

export async function listGoogleDriveFiles(
  query?: string,
  options?: { folderOnly?: boolean }
): Promise<GoogleDriveFileItem[]> {
  const token = requireAccessToken();

  const mimeFilter = SUPPORTED_DRIVE_MIME_TYPES.map(
    (mime) => `mimeType = '${mime}'`
  ).join(' or ');
  let q = `trashed = false and (${mimeFilter})`;

  if (options?.folderOnly) {
    const folderId = await ensureLiteWorkspaceFolder();
    q += ` and '${escapeDriveQueryValue(folderId)}' in parents`;
  }

  const trimmedQuery = query?.trim() ?? '';
  if (trimmedQuery.length > 0) {
    const escaped = escapeDriveQueryValue(trimmedQuery);
    q += ` and name contains '${escaped}'`;
  }

  const params = new URLSearchParams({
    q,
    fields: 'files(id,name,mimeType,modifiedTime,size,iconLink,webViewLink)',
    orderBy: 'modifiedTime desc',
    pageSize: '40',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true',
  });

  const response = await fetch(`${GOOGLE_DRIVE_FILES_URL}?${params.toString()}`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (response.status === 401 || response.status === 403) {
    if (response.status === 401) {
      clearStoredGoogleDriveToken();
    }
    throw new GoogleDriveAuthError(
      'Phiên truy cập Google Drive đã hết hạn hoặc thiếu quyền đọc tệp. Vui lòng kết nối lại Google.',
      response.status
    );
  }

  if (!response.ok) {
    throw new Error(await extractGoogleApiError(response));
  }

  const data: unknown = await response.json();
  if (!isPlainRecord(data) || !Array.isArray(data.files)) {
    return [];
  }

  const items: GoogleDriveFileItem[] = [];
  for (const raw of data.files) {
    if (!isPlainRecord(raw)) {
      continue;
    }
    const id = typeof raw.id === 'string' ? raw.id : '';
    const name = typeof raw.name === 'string' ? raw.name : '';
    const mimeType = typeof raw.mimeType === 'string' ? raw.mimeType : '';
    if (!id || !name || !mimeType) {
      continue;
    }

    const isGoogleWorkspaceType = mimeType in WORKSPACE_EXPORT_MAP;
    const targetExtension = inferTargetExtension(mimeType, name);
    const parsedSize =
      typeof raw.size === 'string'
        ? Number.parseInt(raw.size, 10)
        : typeof raw.size === 'number'
          ? raw.size
          : undefined;

    items.push({
      id,
      name,
      mimeType,
      modifiedTime:
        typeof raw.modifiedTime === 'string' ? raw.modifiedTime : undefined,
      size:
        typeof parsedSize === 'number' && Number.isFinite(parsedSize)
          ? parsedSize
          : undefined,
      iconLink: typeof raw.iconLink === 'string' ? raw.iconLink : undefined,
      webViewLink:
        typeof raw.webViewLink === 'string' ? raw.webViewLink : undefined,
      isGoogleWorkspaceType,
      targetExtension,
    });
  }

  return items;
}

export async function downloadGoogleDriveFile(
  file: GoogleDriveFileItem
): Promise<File> {
  const token = requireAccessToken();

  const workspaceExport = WORKSPACE_EXPORT_MAP[file.mimeType];
  const targetMime = workspaceExport
    ? workspaceExport.exportMime
    : file.mimeType || FALLBACK_MIME_BY_EXT[file.targetExtension];

  const requestUrl = file.isGoogleWorkspaceType
    ? `${GOOGLE_DRIVE_FILES_URL}/${encodeURIComponent(file.id)}/export?mimeType=${encodeURIComponent(targetMime)}`
    : `${GOOGLE_DRIVE_FILES_URL}/${encodeURIComponent(file.id)}?alt=media`;

  const response = await fetch(requestUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (response.status === 401 || response.status === 403) {
    if (response.status === 401) {
      clearStoredGoogleDriveToken();
    }
    throw new GoogleDriveAuthError(
      'Không có quyền tải tệp hoặc phiên Google Drive đã hết hạn. Vui lòng kết nối lại Google.',
      response.status
    );
  }

  if (!response.ok) {
    throw new Error(await extractGoogleApiError(response));
  }

  const arrayBuffer = await response.arrayBuffer();
  const safeFileName = ensureSafeFileNameWithExt(file);
  return new File([arrayBuffer], safeFileName, { type: targetMime });
}

const inFlightFolderLookups = new Map<string, Promise<string>>();

function readCachedFolderId(userId?: string | null): string | null {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return null;
    }
    const raw = window.localStorage.getItem(getFolderStorageKey(userId));
    const trimmed = raw?.trim() ?? '';
    return trimmed.length > 0 ? trimmed : null;
  } catch {
    return null;
  }
}

function writeCachedFolderId(
  folderId: string,
  userId?: string | null
): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    window.localStorage.setItem(getFolderStorageKey(userId), folderId);
  } catch {
    // Ignore storage quota or private mode restrictions
  }
}

function clearCachedFolderId(userId?: string | null): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    window.localStorage.removeItem(getFolderStorageKey(userId));
  } catch {
    // Ignore storage restrictions
  }
}

function handleDriveAuthStatus(
  response: Response,
  authErrorMessage: string,
  userId?: string | null
): void {
  if (response.status === 401 || response.status === 403) {
    if (response.status === 401) {
      clearStoredGoogleDriveToken(userId);
    }
    throw new GoogleDriveAuthError(authErrorMessage, response.status);
  }
}

function buildDriveMultipartPayload(
  metadata: Record<string, unknown>,
  blob: Blob,
  safeMime: string
): { boundary: string; body: Blob } {
  const boundary = `lite_drive_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const metadataJson = JSON.stringify(metadata);

  const body = new Blob(
    [
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      metadataJson,
      `\r\n--${boundary}\r\nContent-Type: ${safeMime}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`,
    ],
    { type: `multipart/related; boundary=${boundary}` }
  );

  return { boundary, body };
}

async function parseUploadedDrivePayload(
  response: Response,
  fallbackName: string,
  userId?: string | null
): Promise<{ id: string; name: string; webViewLink?: string }> {
  handleDriveAuthStatus(
    response,
    'Phiên Google Drive đã hết hạn hoặc thiếu quyền lưu tệp lên Drive. Vui lòng kết nối lại Google.',
    userId
  );

  if (!response.ok) {
    throw new Error(await extractGoogleApiError(response));
  }

  const data: unknown = await response.json();
  if (!isPlainRecord(data) || typeof data.id !== 'string') {
    throw new Error('Phản hồi từ Google Drive không hợp lệ.');
  }

  return {
    id: data.id,
    name: typeof data.name === 'string' ? data.name : fallbackName,
    webViewLink:
      typeof data.webViewLink === 'string' ? data.webViewLink : undefined,
  };
}

export async function ensureLiteWorkspaceFolder(
  userId?: string | null,
  forceRefresh = false
): Promise<string> {
  const token = requireAccessToken(userId);
  const storageKey = getFolderStorageKey(userId);

  if (!forceRefresh) {
    const cached = readCachedFolderId(userId);
    if (cached) {
      return cached;
    }
  } else {
    clearCachedFolderId(userId);
  }

  const existingPromise = inFlightFolderLookups.get(storageKey);
  if (existingPromise) {
    return existingPromise;
  }

  const lookupPromise = (async () => {
    const query = `name = '${escapeDriveQueryValue(LITE_WORKSPACE_FOLDER_NAME)}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false and 'me' in owners`;
    const params = new URLSearchParams({
      q: query,
      fields: 'files(id,name)',
      orderBy: 'modifiedTime desc',
      pageSize: '5',
    });

    const listRes = await fetch(
      `${GOOGLE_DRIVE_FILES_URL}?${params.toString()}`,
      {
        method: 'GET',
        headers: {
          Authorization: `Bearer ${token}`,
        },
      }
    );

    handleDriveAuthStatus(
      listRes,
      'Phiên Google Drive đã hết hạn hoặc thiếu quyền truy cập thư mục. Vui lòng kết nối lại Google.',
      userId
    );

    if (!listRes.ok) {
      throw new Error(await extractGoogleApiError(listRes));
    }

    const listJson: unknown = await listRes.json();
    if (isPlainRecord(listJson) && Array.isArray(listJson.files)) {
      for (const entry of listJson.files) {
        if (
          isPlainRecord(entry) &&
          typeof entry.id === 'string' &&
          entry.id.trim().length > 0
        ) {
          const foundId = entry.id.trim();
          writeCachedFolderId(foundId, userId);
          return foundId;
        }
      }
    }

    const createRes = await fetch(`${GOOGLE_DRIVE_FILES_URL}?fields=id,name`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json; charset=UTF-8',
      },
      body: JSON.stringify({
        name: LITE_WORKSPACE_FOLDER_NAME,
        mimeType: 'application/vnd.google-apps.folder',
      }),
    });

    handleDriveAuthStatus(
      createRes,
      'Không có quyền tạo thư mục LITE Workspace trên Google Drive. Vui lòng kết nối lại Google.',
      userId
    );

    if (!createRes.ok) {
      throw new Error(await extractGoogleApiError(createRes));
    }

    const createJson: unknown = await createRes.json();
    if (!isPlainRecord(createJson) || typeof createJson.id !== 'string') {
      throw new Error('Không thể tạo thư mục LITE Workspace trên Google Drive.');
    }

    const createdId = createJson.id.trim();
    writeCachedFolderId(createdId, userId);
    return createdId;
  })();

  inFlightFolderLookups.set(storageKey, lookupPromise);
  try {
    return await lookupPromise;
  } finally {
    if (inFlightFolderLookups.get(storageKey) === lookupPromise) {
      inFlightFolderLookups.delete(storageKey);
    }
  }
}

export async function uploadFileToGoogleDrive(
  blob: Blob,
  fileName: string,
  mimeType: string
): Promise<{ id: string; name: string; webViewLink?: string }> {
  const token = requireAccessToken();
  const safeName = fileName.trim() || 'lite-document';
  const safeMime = mimeType.trim() || 'application/octet-stream';

  const { boundary, body } = buildDriveMultipartPayload(
    { name: safeName, mimeType: safeMime },
    blob,
    safeMime
  );

  const response = await fetch(GOOGLE_DRIVE_UPLOAD_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body,
  });

  return parseUploadedDrivePayload(response, safeName);
}

export async function upsertDocumentToLiteDrive(
  docId: string,
  blob: Blob,
  fileName: string,
  mimeType: string,
  userId?: string | null
): Promise<{
  id: string;
  name: string;
  webViewLink?: string;
  updatedExisting: boolean;
}> {
  const token = requireAccessToken(userId);
  const cleanDocId = docId.trim();
  const safeName = fileName.trim() || 'lite-document';
  const safeMime = mimeType.trim() || 'application/octet-stream';

  let folderId = await ensureLiteWorkspaceFolder(userId);
  const existingDriveId = cleanDocId
    ? getMappedDriveFileId(cleanDocId, userId)
    : null;

  if (existingDriveId) {
    const patchPayload = buildDriveMultipartPayload(
      { name: safeName, mimeType: safeMime },
      blob,
      safeMime
    );
    const patchResponse = await fetch(
      `https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(existingDriveId)}?uploadType=multipart&fields=id,name,webViewLink`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': `multipart/related; boundary=${patchPayload.boundary}`,
        },
        body: patchPayload.body,
      }
    );

    if (patchResponse.status !== 404) {
      const updated = await parseUploadedDrivePayload(
        patchResponse,
        safeName,
        userId
      );
      if (cleanDocId) {
        setMappedDriveFileId(cleanDocId, updated.id, userId);
      }
      return {
        ...updated,
        updatedExisting: true,
      };
    }

    if (cleanDocId) {
      clearMappedDriveFileId(cleanDocId, userId);
    }
  }

  const createPayload = buildDriveMultipartPayload(
    {
      name: safeName,
      mimeType: safeMime,
      parents: [folderId],
    },
    blob,
    safeMime
  );

  let createResponse = await fetch(GOOGLE_DRIVE_UPLOAD_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/related; boundary=${createPayload.boundary}`,
    },
    body: createPayload.body,
  });

  if (createResponse.status === 404) {
    folderId = await ensureLiteWorkspaceFolder(userId, true);
    const retryPayload = buildDriveMultipartPayload(
      {
        name: safeName,
        mimeType: safeMime,
        parents: [folderId],
      },
      blob,
      safeMime
    );
    createResponse = await fetch(GOOGLE_DRIVE_UPLOAD_URL, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${retryPayload.boundary}`,
      },
      body: retryPayload.body,
    });
  }

  const created = await parseUploadedDrivePayload(
    createResponse,
    safeName,
    userId
  );
  if (cleanDocId) {
    setMappedDriveFileId(cleanDocId, created.id, userId);
  }
  return {
    ...created,
    updatedExisting: false,
  };
}

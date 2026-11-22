export const GOOGLE_TOKEN_STORAGE_KEY = 'lite_google_token';

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

export function getStoredGoogleDriveToken(): string | null {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return null;
    }
    const raw = window.localStorage.getItem(GOOGLE_TOKEN_STORAGE_KEY);
    const trimmed = raw?.trim() ?? '';
    return trimmed.length > 0 ? trimmed : null;
  } catch {
    return null;
  }
}

export function setStoredGoogleDriveToken(token: string | null): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    const trimmed = token?.trim() ?? '';
    if (trimmed.length > 0) {
      window.localStorage.setItem(GOOGLE_TOKEN_STORAGE_KEY, trimmed);
    } else {
      window.localStorage.removeItem(GOOGLE_TOKEN_STORAGE_KEY);
    }
  } catch {
    // Ignore storage quota or private mode restrictions
  }
}

export function clearStoredGoogleDriveToken(): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    window.localStorage.removeItem(GOOGLE_TOKEN_STORAGE_KEY);
  } catch {
    // Ignore storage restrictions
  }
}

function requireAccessToken(): string {
  const token = getStoredGoogleDriveToken();
  if (!token) {
    throw new GoogleDriveAuthError(
      'Chưa có quyền truy cập Google Drive. Vui lòng kết nối tài khoản Google.'
    );
  }
  return token;
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
  query?: string
): Promise<GoogleDriveFileItem[]> {
  const token = requireAccessToken();

  const mimeFilter = SUPPORTED_DRIVE_MIME_TYPES.map(
    (mime) => `mimeType = '${mime}'`
  ).join(' or ');
  let q = `trashed = false and (${mimeFilter})`;

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

export async function uploadFileToGoogleDrive(
  blob: Blob,
  fileName: string,
  mimeType: string
): Promise<{ id: string; name: string; webViewLink?: string }> {
  const token = requireAccessToken();
  const safeName = fileName.trim() || 'lite-document';
  const safeMime = mimeType.trim() || 'application/octet-stream';

  const boundary = `lite_drive_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const metadataJson = JSON.stringify({
    name: safeName,
    mimeType: safeMime,
  });

  const multipartBody = new Blob(
    [
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      metadataJson,
      `\r\n--${boundary}\r\nContent-Type: ${safeMime}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`,
    ],
    { type: `multipart/related; boundary=${boundary}` }
  );

  const response = await fetch(GOOGLE_DRIVE_UPLOAD_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body: multipartBody,
  });

  if (response.status === 401 || response.status === 403) {
    if (response.status === 401) {
      clearStoredGoogleDriveToken();
    }
    throw new GoogleDriveAuthError(
      'Phiên Google Drive đã hết hạn hoặc thiếu quyền lưu tệp lên Drive. Vui lòng kết nối lại Google.',
      response.status
    );
  }

  if (!response.ok) {
    throw new Error(await extractGoogleApiError(response));
  }

  const data: unknown = await response.json();
  if (!isPlainRecord(data) || typeof data.id !== 'string') {
    throw new Error('Phản hồi từ Google Drive không hợp lệ.');
  }

  return {
    id: data.id,
    name: typeof data.name === 'string' ? data.name : safeName,
    webViewLink:
      typeof data.webViewLink === 'string' ? data.webViewLink : undefined,
  };
}

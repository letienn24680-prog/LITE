const R2_API_URL = (import.meta.env.VITE_R2_API_URL ?? '')
  .trim()
  .replace(/\/+$/, '');

const PLACEHOLDER_TOKENS = [
  'your-worker',
  'your-r2-api',
  'example.com',
  '<',
  '>',
];

function isValidR2ApiUrl(value: string): boolean {
  if (!/^https?:\/\/\S+$/i.test(value)) {
    return false;
  }
  const lower = value.toLowerCase();
  if (PLACEHOLDER_TOKENS.some((token) => lower.includes(token))) {
    return false;
  }
  try {
    const parsed = new URL(value);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:';
  } catch {
    return false;
  }
}

interface PresignUploadResponse {
  uploadUrl?: string;
  objectKey?: string;
  method?: string;
  error?: string;
  message?: string;
}

interface PresignDownloadResponse {
  downloadUrl?: string;
  viewUrl?: string;
  allowDownload?: boolean;
  error?: string;
  message?: string;
}

/** True when VITE_R2_API_URL is configured with a valid HTTP(S) endpoint. */
export const isR2ApiConfigured: boolean = isValidR2ApiUrl(R2_API_URL);

/** Returns the normalized Cloudflare Worker R2 API base URL (empty string if unconfigured). */
export function getR2ApiBaseUrl(): string {
  return R2_API_URL;
}

function extractApiErrorMessage(
  payload: unknown,
  fallbackMessage: string
): string {
  if (typeof payload === 'object' && payload !== null) {
    const record = payload as Record<string, unknown>;
    if (typeof record.error === 'string' && record.error.trim()) {
      return record.error.trim();
    }
    if (typeof record.message === 'string' && record.message.trim()) {
      return record.message.trim();
    }
  }
  return fallbackMessage;
}

function resolveAbsoluteUploadUrl(rawUrl: string): string {
  if (/^https?:\/\//i.test(rawUrl)) {
    return rawUrl;
  }
  const normalizedPath = rawUrl.startsWith('/') ? rawUrl : `/${rawUrl}`;
  return `${R2_API_URL}${normalizedPath}`;
}

function uploadViaXhr(params: {
  uploadUrl: string;
  method: string;
  file: File;
  contentType: string;
  accessToken: string;
  onProgress?: (percent: number) => void;
}): Promise<void> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open(params.method, params.uploadUrl, true);
    xhr.setRequestHeader('Content-Type', params.contentType);

    // Attach Bearer only when hitting the Worker proxy fallback directly;
    // S3 Presigned URLs reject unexpected Authorization headers.
    if (params.uploadUrl.startsWith(R2_API_URL)) {
      xhr.setRequestHeader('Authorization', `Bearer ${params.accessToken}`);
    }

    xhr.upload.onprogress = (event: ProgressEvent<EventTarget>) => {
      if (event.lengthComputable && event.total > 0) {
        const pct = Math.min(99, Math.round((event.loaded / event.total) * 100));
        params.onProgress?.(pct);
      }
    };

    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        params.onProgress?.(100);
        resolve();
        return;
      }
      reject(
        new Error(
          `Tải tệp lên Cloudflare R2 thất bại (HTTP ${xhr.status}). Vui lòng thử lại.`
        )
      );
    };

    xhr.onerror = () => {
      reject(
        new Error(
          'Không thể kết nối tới Cloudflare R2 để tải tệp lên. Vui lòng kiểm tra mạng hoặc cấu hình CORS của bucket.'
        )
      );
    };

    xhr.onabort = () => {
      reject(new Error('Quá trình tải tệp lên Cloudflare R2 đã bị hủy.'));
    };

    xhr.send(params.file);
  });
}

/**
 * Requests a short-lived R2 Presigned PUT URL from the Cloudflare Worker and uploads
 * the file directly from the browser while reporting upload percentage (0..100).
 */
export async function uploadFileViaR2PresignedUrl(params: {
  file: File;
  documentId: string;
  accessToken: string;
  onProgress?: (percent: number) => void;
}): Promise<{ objectKey: string }> {
  if (!isR2ApiConfigured) {
    throw new Error('Dịch vụ Cloudflare R2 chưa được cấu hình (VITE_R2_API_URL).');
  }
  if (!params.accessToken.trim()) {
    throw new Error('Cần đăng nhập tài khoản để tải tệp lên Cloudflare R2.');
  }

  const contentType = params.file.type || 'application/octet-stream';
  params.onProgress?.(0);

  const presignResponse = await fetch(`${R2_API_URL}/api/r2/presign-upload`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${params.accessToken}`,
    },
    body: JSON.stringify({
      documentId: params.documentId,
      fileName: params.file.name,
      contentType,
      fileSizeBytes: params.file.size,
    }),
  });

  const presignData = (await presignResponse
    .json()
    .catch(() => null)) as PresignUploadResponse | null;

  if (!presignResponse.ok || !presignData?.uploadUrl || !presignData.objectKey) {
    const detail = extractApiErrorMessage(
      presignData,
      `Không thể xin URL tải lên R2 (HTTP ${presignResponse.status}).`
    );
    throw new Error(detail);
  }

  const resolvedUploadUrl = resolveAbsoluteUploadUrl(presignData.uploadUrl);
  const uploadMethod = (presignData.method ?? 'PUT').toUpperCase();

  await uploadViaXhr({
    uploadUrl: resolvedUploadUrl,
    method: uploadMethod,
    file: params.file,
    contentType,
    accessToken: params.accessToken,
    onProgress: params.onProgress,
  });

  return {
    objectKey: presignData.objectKey,
  };
}

/**
 * Requests a short-lived R2 Presigned GET URL for an authenticated owner or a valid share token.
 */
export async function getR2PresignedDownloadUrl(params: {
  objectKey: string;
  accessToken?: string | null;
  shareToken?: string;
  passwordHash?: string;
}): Promise<{ downloadUrl: string; allowDownload: boolean }> {
  if (!isR2ApiConfigured) {
    throw new Error('Dịch vụ Cloudflare R2 chưa được cấu hình (VITE_R2_API_URL).');
  }
  if (!params.objectKey.trim()) {
    throw new Error('Khóa đối tượng Cloudflare R2 không hợp lệ.');
  }

  const query = new URLSearchParams();
  query.set('key', params.objectKey);
  if (params.shareToken) {
    query.set('shareToken', params.shareToken);
  }
  if (params.passwordHash) {
    query.set('passwordHash', params.passwordHash);
    query.set('password', params.passwordHash);
  }

  const headers: Record<string, string> = {};
  if (params.accessToken && params.accessToken.trim()) {
    headers.Authorization = `Bearer ${params.accessToken.trim()}`;
  }

  const response = await fetch(
    `${R2_API_URL}/api/r2/presign-download?${query.toString()}`,
    {
      method: 'GET',
      headers,
    }
  );

  const payload = (await response
    .json()
    .catch(() => null)) as PresignDownloadResponse | null;

  const resolvedUrl = payload?.downloadUrl ?? payload?.viewUrl;
  if (!response.ok || !resolvedUrl) {
    const detail = extractApiErrorMessage(
      payload,
      `Không thể lấy liên kết tải tệp từ Cloudflare R2 (HTTP ${response.status}).`
    );
    throw new Error(detail);
  }

  return {
    downloadUrl: resolveAbsoluteUploadUrl(resolvedUrl),
    allowDownload: payload?.allowDownload !== false,
  };
}

const MAX_UPLOAD_BYTES = 52_428_800; // 50 MB
const UPLOAD_PRESIGN_EXPIRES_SECONDS = 600;
const DOWNLOAD_PRESIGN_EXPIRES_SECONDS = 300;
const ALLOWED_EXTENSIONS = new Set([
  '.pptx',
  '.docx',
  '.xlsx',
  '.xls',
  '.csv',
  '.pdf',
]);
const VALID_FILE_TYPES = new Set(['pptx', 'docx', 'xlsx', 'pdf']);
const VALID_ACCESS_LEVELS = new Set(['private', 'link', 'password']);

const textEncoder = new TextEncoder();

export interface R2HttpMetadataLike {
  contentType?: string;
}

export interface R2PutOptionsLike {
  httpMetadata?: R2HttpMetadataLike;
}

export interface R2ObjectLike {
  key: string;
  size: number;
  etag: string;
  httpEtag?: string;
  httpMetadata?: R2HttpMetadataLike;
}

export interface R2ObjectBodyLike extends R2ObjectLike {
  body: ReadableStream<Uint8Array>;
}

export interface R2BucketBinding {
  get(key: string): Promise<R2ObjectBodyLike | null>;
  put(
    key: string,
    value: ReadableStream<Uint8Array> | ArrayBuffer | ArrayBufferView | string | null,
    options?: R2PutOptionsLike
  ): Promise<R2ObjectLike | null>;
  delete(keys: string | string[]): Promise<void>;
}

export interface WorkerExecutionContext {
  waitUntil(promise: Promise<unknown>): void;
  passThroughOnException(): void;
}

export interface Env {
  R2_BUCKET?: R2BucketBinding;
  ALLOWED_ORIGIN?: string;
  SUPABASE_URL?: string;
  SUPABASE_ANON_KEY?: string;
  R2_ACCOUNT_ID?: string;
  R2_BUCKET_NAME?: string;
  R2_ACCESS_KEY_ID?: string;
  R2_SECRET_ACCESS_KEY?: string;
}

interface AuthenticatedUser {
  id: string;
  email: string;
  accessToken: string;
}

interface SupabaseShareRpcResult {
  status: string;
  share_token?: string;
  document_id?: string;
  title?: string;
  file_type?: string;
  file_name?: string | null;
  file_size?: string | null;
  allow_download?: boolean;
  view_count?: number;
  r2_object_key?: string | null;
  storage_path?: string | null;
  snapshot_payload?: Record<string, unknown>;
}

interface DocumentShareRow {
  share_token: string;
  document_id: string;
  user_id?: string;
  title: string;
  file_type: string;
  file_name: string | null;
  file_size: string | null;
  r2_object_key: string | null;
  storage_path: string | null;
  access_level: string;
  password_hash?: string | null;
  allow_download: boolean;
  snapshot_payload?: Record<string, unknown>;
  view_count: number;
  revoked_at: string | null;
  updated_at: string;
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function getTrimmedEnv(value: string | undefined): string {
  return typeof value === 'string' ? value.trim() : '';
}

function isValidSupabaseConfig(env: Env): { url: string; anonKey: string } | null {
  const url = getTrimmedEnv(env.SUPABASE_URL).replace(/\/+$/, '');
  const anonKey = getTrimmedEnv(env.SUPABASE_ANON_KEY);
  if (
    !/^https?:\/\//i.test(url) ||
    url.includes('your-project-ref') ||
    !anonKey ||
    anonKey.includes('your-anon-key')
  ) {
    return null;
  }
  return { url, anonKey };
}

function getS3PresignConfig(env: Env): {
  accountId: string;
  bucketName: string;
  accessKeyId: string;
  secretAccessKey: string;
} | null {
  const accountId = getTrimmedEnv(env.R2_ACCOUNT_ID);
  const bucketName = getTrimmedEnv(env.R2_BUCKET_NAME) || 'lite-user-files';
  const accessKeyId = getTrimmedEnv(env.R2_ACCESS_KEY_ID);
  const secretAccessKey = getTrimmedEnv(env.R2_SECRET_ACCESS_KEY);

  if (!accountId || !bucketName || !accessKeyId || !secretAccessKey) {
    return null;
  }
  return { accountId, bucketName, accessKeyId, secretAccessKey };
}

function buildCorsHeaders(request: Request, env: Env): Record<string, string> {
  const rawAllowed = getTrimmedEnv(env.ALLOWED_ORIGIN) || '*';
  const requestOrigin = request.headers.get('Origin')?.trim() ?? '';

  let allowOriginHeader = '*';
  let includeVary = false;

  if (rawAllowed !== '*') {
    const allowedList = rawAllowed
      .split(',')
      .map((item) => item.trim().replace(/\/+$/, ''))
      .filter((item) => item.length > 0);

    const normalizedRequestOrigin = requestOrigin.replace(/\/+$/, '');
    if (allowedList.length === 0) {
      allowOriginHeader = '*';
    } else if (normalizedRequestOrigin && allowedList.includes(normalizedRequestOrigin)) {
      allowOriginHeader = requestOrigin;
      includeVary = true;
    } else {
      allowOriginHeader = allowedList[0];
      includeVary = true;
    }
  }

  const headers: Record<string, string> = {
    'Access-Control-Allow-Origin': allowOriginHeader,
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type, X-Requested-With, apikey',
    'Access-Control-Max-Age': '86400',
  };

  if (includeVary) {
    headers['Vary'] = 'Origin';
  }

  return headers;
}

function jsonResponse(
  request: Request,
  env: Env,
  status: number,
  payload: Record<string, unknown>
): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
      ...buildCorsHeaders(request, env),
    },
  });
}

function toHex(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 1) {
    out += bytes[i].toString(16).padStart(2, '0');
  }
  return out;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

async function sha256Hex(input: string | Uint8Array): Promise<string> {
  const data = typeof input === 'string' ? textEncoder.encode(input) : input;
  const digest = await crypto.subtle.digest('SHA-256', toArrayBuffer(data));
  return toHex(new Uint8Array(digest));
}

async function hmacSha256Bytes(key: Uint8Array, message: string): Promise<Uint8Array> {
  const cryptoKey = await crypto.subtle.importKey(
    'raw',
    toArrayBuffer(key),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const signature = await crypto.subtle.sign(
    'HMAC',
    cryptoKey,
    toArrayBuffer(textEncoder.encode(message))
  );
  return new Uint8Array(signature);
}

function constantTimeEquals(a: string, b: string): boolean {
  if (a.length !== b.length) {
    return false;
  }
  let diff = 0;
  for (let i = 0; i < a.length; i += 1) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function encodeRfc3986(value: string): string {
  return encodeURIComponent(value).replace(
    /[!'()*]/g,
    (ch) => `%${ch.charCodeAt(0).toString(16).toUpperCase()}`
  );
}

function formatAmzTimestamps(date: Date): { amzDate: string; dateStamp: string } {
  const iso = date.toISOString(); // YYYY-MM-DDTHH:mm:ss.sssZ
  const compact = iso.replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');
  return {
    amzDate: compact,
    dateStamp: compact.slice(0, 8),
  };
}

/**
 * Generates an AWS Signature V4 presigned URL for Cloudflare R2 S3 endpoints
 * using only Web Crypto (`crypto.subtle`).
 */
export async function createR2PresignedUrl(options: {
  method: 'GET' | 'PUT';
  accountId: string;
  bucketName: string;
  accessKeyId: string;
  secretAccessKey: string;
  objectKey: string;
  expiresInSeconds: number;
}): Promise<string> {
  const expires = Math.max(1, Math.min(604_800, Math.floor(options.expiresInSeconds)));
  const host = `${options.accountId}.r2.cloudflarestorage.com`;
  const { amzDate, dateStamp } = formatAmzTimestamps(new Date());
  const credentialScope = `${dateStamp}/auto/s3/aws4_request`;

  const keySegments = options.objectKey.split('/').filter((seg) => seg.length > 0);
  const canonicalUri = `/${encodeRfc3986(options.bucketName)}/${keySegments
    .map(encodeRfc3986)
    .join('/')}`;

  const queryPairs: [string, string][] = [
    ['X-Amz-Algorithm', 'AWS4-HMAC-SHA256'],
    ['X-Amz-Credential', `${options.accessKeyId}/${credentialScope}`],
    ['X-Amz-Date', amzDate],
    ['X-Amz-Expires', String(expires)],
    ['X-Amz-SignedHeaders', 'host'],
  ];

  const encodedPairs = queryPairs
    .map(([k, v]): [string, string] => [encodeRfc3986(k), encodeRfc3986(v)])
    .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0));

  const canonicalQueryString = encodedPairs.map(([k, v]) => `${k}=${v}`).join('&');
  const canonicalHeaders = `host:${host}\n`;
  const signedHeaders = 'host';
  const canonicalRequest = [
    options.method,
    canonicalUri,
    canonicalQueryString,
    canonicalHeaders,
    signedHeaders,
    'UNSIGNED-PAYLOAD',
  ].join('\n');

  const canonicalRequestHash = await sha256Hex(canonicalRequest);
  const stringToSign = [
    'AWS4-HMAC-SHA256',
    amzDate,
    credentialScope,
    canonicalRequestHash,
  ].join('\n');

  const kSecret = textEncoder.encode(`AWS4${options.secretAccessKey}`);
  const kDate = await hmacSha256Bytes(kSecret, dateStamp);
  const kRegion = await hmacSha256Bytes(kDate, 'auto');
  const kService = await hmacSha256Bytes(kRegion, 's3');
  const kSigning = await hmacSha256Bytes(kService, 'aws4_request');
  const signatureBytes = await hmacSha256Bytes(kSigning, stringToSign);
  const signature = toHex(signatureBytes);

  return `https://${host}${canonicalUri}?${canonicalQueryString}&X-Amz-Signature=${signature}`;
}

function getDirectProxySigningSecret(env: Env): string {
  return (
    getTrimmedEnv(env.R2_SECRET_ACCESS_KEY) ||
    getTrimmedEnv(env.SUPABASE_ANON_KEY) ||
    'lite-r2-local-binding-fallback-secret'
  );
}

async function signDirectProxyUrl(
  request: Request,
  env: Env,
  op: 'PUT' | 'GET',
  objectKey: string,
  expiresInSeconds: number
): Promise<string> {
  const reqUrl = new URL(request.url);
  const exp = Math.floor(Date.now() / 1000) + expiresInSeconds;
  const secret = getDirectProxySigningSecret(env);
  const payload = `${op}:${objectKey}:${exp}`;
  const sigBytes = await hmacSha256Bytes(textEncoder.encode(secret), payload);
  const sig = toHex(sigBytes);
  const encodedKey = objectKey.split('/').map(encodeURIComponent).join('/');
  return `${reqUrl.origin}/api/r2/direct/${encodedKey}?exp=${exp}&sig=${sig}`;
}

async function verifyDirectProxySignature(
  env: Env,
  op: 'PUT' | 'GET',
  objectKey: string,
  expParam: string | null,
  sigParam: string | null
): Promise<boolean> {
  if (!expParam || !sigParam || !/^\d+$/.test(expParam) || !/^[a-f0-9]{64}$/i.test(sigParam)) {
    return false;
  }
  const exp = Number.parseInt(expParam, 10);
  const nowSec = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(exp) || exp < nowSec) {
    return false;
  }
  const secret = getDirectProxySigningSecret(env);
  const expectedBytes = await hmacSha256Bytes(
    textEncoder.encode(secret),
    `${op}:${objectKey}:${exp}`
  );
  const expectedSig = toHex(expectedBytes);
  return constantTimeEquals(expectedSig, sigParam.toLowerCase());
}

function sanitizePathSegment(segment: string, fallback: string): string {
  const cleaned = segment
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, '_')
    .replace(/\.{2,}/g, '_')
    .slice(0, 120);
  return cleaned || fallback;
}

function sanitizeStorageFileName(originalName: string): {
  safeFileName: string;
  extension: string;
} {
  const trimmed = originalName.trim();
  const lastDotIdx = trimmed.lastIndexOf('.');
  const rawBase = lastDotIdx > 0 ? trimmed.slice(0, lastDotIdx) : trimmed;
  const rawExt = lastDotIdx > 0 ? trimmed.slice(lastDotIdx + 1) : '';

  const asciiBase = rawBase
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/-{2,}/g, '-')
    .replace(/^[-_]+|[-_]+$/g, '')
    .slice(0, 100);

  const safeBase = asciiBase || 'document';
  const safeExt = rawExt
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase()
    .slice(0, 12);

  const extension = safeExt ? `.${safeExt}` : '';
  const safeFileName = safeExt ? `${safeBase}.${safeExt}` : safeBase;
  return { safeFileName, extension };
}

function isSafeObjectKey(key: string): boolean {
  if (!key || key.length > 512 || key.startsWith('/') || key.endsWith('/')) {
    return false;
  }
  if (key.includes('\\') || key.includes('\0')) {
    return false;
  }
  const segments = key.split('/');
  return segments.every((seg) => seg.length > 0 && seg !== '.' && seg !== '..' && !seg.includes('..'));
}

function extractBearerToken(request: Request): string | null {
  const authHeader = request.headers.get('Authorization');
  if (!authHeader) {
    return null;
  }
  const match = authHeader.trim().match(/^Bearer\s+(.+)$/i);
  if (!match) {
    return null;
  }
  const token = match[1].trim();
  return token.length > 0 ? token : null;
}

/**
 * Verifies a Supabase user JWT via `${env.SUPABASE_URL}/auth/v1/user`.
 */
export async function authenticateSupabaseUser(
  request: Request,
  env: Env
): Promise<
  | { ok: true; user: AuthenticatedUser }
  | { ok: false; status: number; message: string }
> {
  const token = extractBearerToken(request);
  if (!token) {
    return {
      ok: false,
      status: 401,
      message: 'Thiếu mã xác thực người dùng (Bearer token).',
    };
  }

  const supabaseCfg = isValidSupabaseConfig(env);
  if (!supabaseCfg) {
    return {
      ok: false,
      status: 503,
      message: 'Worker chưa được cấu hình SUPABASE_URL và SUPABASE_ANON_KEY.',
    };
  }

  let authRes: Response;
  try {
    authRes = await fetch(`${supabaseCfg.url}/auth/v1/user`, {
      method: 'GET',
      headers: {
        apikey: supabaseCfg.anonKey,
        Authorization: `Bearer ${token}`,
      },
    });
  } catch {
    return {
      ok: false,
      status: 502,
      message: 'Không thể kết nối tới dịch vụ xác thực Supabase.',
    };
  }

  if (!authRes.ok) {
    return {
      ok: false,
      status: 401,
      message: 'Phiên đăng nhập không hợp lệ hoặc đã hết hạn.',
    };
  }

  let body: unknown;
  try {
    body = await authRes.json();
  } catch {
    return {
      ok: false,
      status: 401,
      message: 'Phản hồi xác thực Supabase không hợp lệ.',
    };
  }

  if (!isPlainRecord(body) || typeof body.id !== 'string' || body.id.trim().length === 0) {
    return {
      ok: false,
      status: 401,
      message: 'Không xác định được danh tính người dùng.',
    };
  }

  return {
    ok: true,
    user: {
      id: body.id.trim(),
      email: typeof body.email === 'string' ? body.email.trim() : '',
      accessToken: token,
    },
  };
}

async function readJsonBody(
  request: Request
): Promise<{ ok: true; data: Record<string, unknown> } | { ok: false; message: string }> {
  try {
    const parsed: unknown = await request.json();
    if (!isPlainRecord(parsed)) {
      return { ok: false, message: 'Dữ liệu JSON phải là một đối tượng hợp lệ.' };
    }
    return { ok: true, data: parsed };
  } catch {
    return { ok: false, message: 'Không thể đọc nội dung JSON gửi lên.' };
  }
}

async function callSupabaseShareRpc(
  env: Env,
  fnName: 'get_shared_document_meta' | 'unlock_shared_document',
  args: Record<string, unknown>
): Promise<
  | { ok: true; data: SupabaseShareRpcResult }
  | { ok: false; status: number; message: string }
> {
  const supabaseCfg = isValidSupabaseConfig(env);
  if (!supabaseCfg) {
    return {
      ok: false,
      status: 503,
      message: 'Worker chưa được cấu hình SUPABASE_URL và SUPABASE_ANON_KEY.',
    };
  }

  let res: Response;
  try {
    res = await fetch(`${supabaseCfg.url}/rest/v1/rpc/${fnName}`, {
      method: 'POST',
      headers: {
        apikey: supabaseCfg.anonKey,
        Authorization: `Bearer ${supabaseCfg.anonKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(args),
    });
  } catch {
    return {
      ok: false,
      status: 502,
      message: 'Không thể kết nối tới Supabase RPC.',
    };
  }

  if (!res.ok) {
    return {
      ok: false,
      status: res.status >= 500 ? 502 : res.status,
      message: `Lỗi truy vấn liên kết chia sẻ (${res.status}).`,
    };
  }

  let raw: unknown;
  try {
    raw = await res.json();
  } catch {
    return {
      ok: false,
      status: 502,
      message: 'Phản hồi từ Supabase RPC không hợp lệ.',
    };
  }

  if (!isPlainRecord(raw) || typeof raw.status !== 'string') {
    return {
      ok: false,
      status: 502,
      message: 'Định dạng dữ liệu chia sẻ từ Supabase không hợp lệ.',
    };
  }

  return {
    ok: true,
    data: raw as unknown as SupabaseShareRpcResult,
  };
}

async function generateDownloadAccessUrl(
  request: Request,
  env: Env,
  objectKey: string,
  expiresInSeconds = DOWNLOAD_PRESIGN_EXPIRES_SECONDS
): Promise<string | null> {
  const s3Cfg = getS3PresignConfig(env);
  if (s3Cfg) {
    return createR2PresignedUrl({
      method: 'GET',
      accountId: s3Cfg.accountId,
      bucketName: s3Cfg.bucketName,
      accessKeyId: s3Cfg.accessKeyId,
      secretAccessKey: s3Cfg.secretAccessKey,
      objectKey,
      expiresInSeconds,
    });
  }
  if (env.R2_BUCKET) {
    return signDirectProxyUrl(request, env, 'GET', objectKey, expiresInSeconds);
  }
  return null;
}

async function generateUploadAccessUrl(
  request: Request,
  env: Env,
  objectKey: string,
  expiresInSeconds = UPLOAD_PRESIGN_EXPIRES_SECONDS
): Promise<string | null> {
  const s3Cfg = getS3PresignConfig(env);
  if (s3Cfg) {
    return createR2PresignedUrl({
      method: 'PUT',
      accountId: s3Cfg.accountId,
      bucketName: s3Cfg.bucketName,
      accessKeyId: s3Cfg.accessKeyId,
      secretAccessKey: s3Cfg.secretAccessKey,
      objectKey,
      expiresInSeconds,
    });
  }
  if (env.R2_BUCKET) {
    return signDirectProxyUrl(request, env, 'PUT', objectKey, expiresInSeconds);
  }
  return null;
}

function generateShareToken(): string {
  const bytes = new Uint8Array(9);
  crypto.getRandomValues(bytes);
  const alphabet = 'abcdefghjkmnpqrstuvwxyz23456789';
  let suffix = '';
  for (let i = 0; i < bytes.length; i += 1) {
    suffix += alphabet[bytes[i] % alphabet.length];
  }
  return `lite-${suffix}`;
}

function normalizeShareRowResponse(row: DocumentShareRow): Record<string, unknown> {
  return {
    shareToken: row.share_token,
    documentId: row.document_id,
    title: row.title,
    fileType: row.file_type,
    fileName: row.file_name ?? undefined,
    fileSize: row.file_size ?? undefined,
    r2ObjectKey: row.r2_object_key ?? undefined,
    storagePath: row.storage_path ?? undefined,
    accessLevel: row.access_level,
    allowDownload: Boolean(row.allow_download),
    viewCount: typeof row.view_count === 'number' ? row.view_count : 0,
    revokedAt: row.revoked_at,
    updatedAt: row.updated_at,
  };
}

async function resolvePasswordHashInput(
  password: unknown,
  passwordHash: unknown
): Promise<string | null> {
  if (typeof passwordHash === 'string') {
    const trimmedHash = passwordHash.trim().toLowerCase();
    if (/^[a-f0-9]{64}$/.test(trimmedHash)) {
      return trimmedHash;
    }
  }
  if (typeof password === 'string' && password.length > 0) {
    return sha256Hex(password);
  }
  return null;
}

function handleHealth(request: Request, env: Env): Response {
  const hasS3Presign = getS3PresignConfig(env) !== null;
  const hasR2Binding = Boolean(env.R2_BUCKET);
  return jsonResponse(request, env, 200, {
    ok: true,
    service: 'lite-r2-api',
    hasS3Presign,
    hasR2Binding,
  });
}

async function handlePresignUpload(request: Request, env: Env): Promise<Response> {
  const authResult = await authenticateSupabaseUser(request, env);
  if (!authResult.ok) {
    return jsonResponse(request, env, authResult.status, { error: authResult.message });
  }

  const bodyResult = await readJsonBody(request);
  if (!bodyResult.ok) {
    return jsonResponse(request, env, 400, { error: bodyResult.message });
  }

  const { documentId, fileName, fileSizeBytes } = bodyResult.data;
  if (typeof documentId !== 'string' || documentId.trim().length === 0) {
    return jsonResponse(request, env, 400, { error: 'Thiếu documentId hợp lệ.' });
  }
  if (typeof fileName !== 'string' || fileName.trim().length === 0) {
    return jsonResponse(request, env, 400, { error: 'Thiếu fileName hợp lệ.' });
  }

  if (
    typeof fileSizeBytes !== 'number' ||
    !Number.isFinite(fileSizeBytes) ||
    fileSizeBytes < 0
  ) {
    return jsonResponse(request, env, 400, { error: 'Thiếu fileSizeBytes hợp lệ.' });
  }
  if (fileSizeBytes > MAX_UPLOAD_BYTES) {
    return jsonResponse(request, env, 413, {
      error: 'Dung lượng tệp vượt quá giới hạn tối đa 50 MB.',
    });
  }

  const { safeFileName, extension } = sanitizeStorageFileName(fileName);
  if (!ALLOWED_EXTENSIONS.has(extension)) {
    return jsonResponse(request, env, 400, {
      error:
        'Định dạng tệp không được hỗ trợ. Chỉ chấp nhận .pptx, .docx, .xlsx, .xls, .csv, .pdf.',
    });
  }

  const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
  const safeDocId = sanitizePathSegment(documentId, 'doc');
  const objectKey = `${safeUserId}/${safeDocId}/${safeFileName}`;

  const uploadUrl = await generateUploadAccessUrl(
    request,
    env,
    objectKey,
    UPLOAD_PRESIGN_EXPIRES_SECONDS
  );
  if (!uploadUrl) {
    return jsonResponse(request, env, 503, {
      error:
        'Kho lưu trữ Cloudflare R2 chưa được cấu hình (thiếu khóa S3 API hoặc R2_BUCKET binding).',
    });
  }

  return jsonResponse(request, env, 200, {
    uploadUrl,
    objectKey,
    method: 'PUT',
    expiresIn: UPLOAD_PRESIGN_EXPIRES_SECONDS,
  });
}

async function handlePresignDownload(request: Request, env: Env): Promise<Response> {
  const url = new URL(request.url);
  const key = (url.searchParams.get('key') ?? url.searchParams.get('objectKey') ?? '').trim();
  const shareToken = (url.searchParams.get('shareToken') ?? '').trim();
  const passwordParam = url.searchParams.get('password');
  const passwordHashParam = url.searchParams.get('passwordHash');

  if (!isSafeObjectKey(key)) {
    return jsonResponse(request, env, 400, { error: 'Khóa đối tượng R2 (key) không hợp lệ.' });
  }

  const bearerToken = extractBearerToken(request);
  if (bearerToken && !shareToken) {
    const authResult = await authenticateSupabaseUser(request, env);
    if (!authResult.ok) {
      return jsonResponse(request, env, authResult.status, { error: authResult.message });
    }
    const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
    if (!key.startsWith(`${authResult.user.id}/`) && !key.startsWith(`${safeUserId}/`)) {
      return jsonResponse(request, env, 403, {
        error: 'Bạn không có quyền truy cập tệp của tài khoản khác.',
      });
    }

    const downloadUrl = await generateDownloadAccessUrl(
      request,
      env,
      key,
      DOWNLOAD_PRESIGN_EXPIRES_SECONDS
    );
    if (!downloadUrl) {
      return jsonResponse(request, env, 503, {
        error: 'Kho lưu trữ Cloudflare R2 chưa được cấu hình.',
      });
    }

    return jsonResponse(request, env, 200, {
      downloadUrl,
      viewUrl: downloadUrl,
      allowDownload: true,
      downloadAllowed: true,
      expiresIn: DOWNLOAD_PRESIGN_EXPIRES_SECONDS,
    });
  }

  if (bearerToken && shareToken) {
    const authResult = await authenticateSupabaseUser(request, env);
    if (authResult.ok) {
      const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
      if (key.startsWith(`${authResult.user.id}/`) || key.startsWith(`${safeUserId}/`)) {
        const downloadUrl = await generateDownloadAccessUrl(
          request,
          env,
          key,
          DOWNLOAD_PRESIGN_EXPIRES_SECONDS
        );
        if (!downloadUrl) {
          return jsonResponse(request, env, 503, {
            error: 'Kho lưu trữ Cloudflare R2 chưa được cấu hình.',
          });
        }
        return jsonResponse(request, env, 200, {
          downloadUrl,
          viewUrl: downloadUrl,
          allowDownload: true,
          downloadAllowed: true,
          expiresIn: DOWNLOAD_PRESIGN_EXPIRES_SECONDS,
        });
      }
    }
  }

  if (!shareToken) {
    return jsonResponse(request, env, 401, {
      error: 'Cần xác thực người dùng hoặc shareToken hợp lệ để tải tệp.',
    });
  }

  const resolvedPasswordHash = await resolvePasswordHashInput(
    passwordParam,
    passwordHashParam
  );
  const rpcResult = resolvedPasswordHash
    ? await callSupabaseShareRpc(env, 'unlock_shared_document', {
        p_share_token: shareToken,
        p_password_hash: resolvedPasswordHash,
      })
    : await callSupabaseShareRpc(env, 'get_shared_document_meta', {
        p_share_token: shareToken,
      });

  if (!rpcResult.ok) {
    return jsonResponse(request, env, rpcResult.status, { error: rpcResult.message });
  }

  const meta = rpcResult.data;
  if (meta.status === 'not_found') {
    return jsonResponse(request, env, 404, {
      error: 'Liên kết chia sẻ không tồn tại.',
      status: 'not_found',
    });
  }
  if (meta.status === 'revoked' || meta.status === 'private') {
    return jsonResponse(request, env, 403, {
      error: 'Liên kết chia sẻ đã bị thu hồi hoặc đang ở chế độ Riêng tư.',
      status: meta.status,
    });
  }
  if (meta.status === 'password_required' || meta.status === 'password_invalid') {
    return jsonResponse(request, env, 401, {
      error:
        meta.status === 'password_invalid'
          ? 'Mật khẩu chia sẻ không chính xác.'
          : 'Liên kết chia sẻ yêu cầu mật khẩu.',
      status: meta.status,
    });
  }
  if (meta.status !== 'ok') {
    return jsonResponse(request, env, 403, {
      error: 'Không thể xác minh quyền truy cập liên kết chia sẻ.',
      status: meta.status,
    });
  }

  if (meta.r2_object_key !== key) {
    return jsonResponse(request, env, 403, {
      error: 'Khóa tệp không khớp với tài liệu được chia sẻ.',
    });
  }

  const allowDownload = Boolean(meta.allow_download);
  const isPdf = meta.file_type === 'pdf';
  if (!allowDownload && !isPdf) {
    return jsonResponse(request, env, 403, {
      error: 'Chủ sở hữu chỉ cho phép xem trực tuyến, không cho phép tải xuống tệp gốc.',
      allowDownload: false,
      downloadAllowed: false,
    });
  }

  const signedUrl = await generateDownloadAccessUrl(
    request,
    env,
    key,
    DOWNLOAD_PRESIGN_EXPIRES_SECONDS
  );
  if (!signedUrl) {
    return jsonResponse(request, env, 503, {
      error: 'Kho lưu trữ Cloudflare R2 chưa được cấu hình.',
    });
  }

  return jsonResponse(request, env, 200, {
    downloadUrl: signedUrl,
    viewUrl: signedUrl,
    allowDownload,
    downloadAllowed: allowDownload,
    expiresIn: DOWNLOAD_PRESIGN_EXPIRES_SECONDS,
  });
}

async function handleUpsertShare(request: Request, env: Env): Promise<Response> {
  const authResult = await authenticateSupabaseUser(request, env);
  if (!authResult.ok) {
    return jsonResponse(request, env, authResult.status, { error: authResult.message });
  }

  const supabaseCfg = isValidSupabaseConfig(env);
  if (!supabaseCfg) {
    return jsonResponse(request, env, 503, {
      error: 'Worker chưa được cấu hình Supabase.',
    });
  }

  const bodyResult = await readJsonBody(request);
  if (!bodyResult.ok) {
    return jsonResponse(request, env, 400, { error: bodyResult.message });
  }

  const body = bodyResult.data;
  const documentId = typeof body.documentId === 'string' ? body.documentId.trim() : '';
  const title = typeof body.title === 'string' && body.title.trim() ? body.title.trim() : 'Tài liệu LITE';
  const fileType = typeof body.fileType === 'string' ? body.fileType.trim().toLowerCase() : '';
  const accessLevel = typeof body.accessLevel === 'string' ? body.accessLevel.trim().toLowerCase() : '';
  const allowDownload = body.allowDownload === undefined ? true : Boolean(body.allowDownload);

  if (!documentId) {
    return jsonResponse(request, env, 400, { error: 'Thiếu documentId.' });
  }
  if (!VALID_FILE_TYPES.has(fileType)) {
    return jsonResponse(request, env, 400, { error: 'fileType không hợp lệ.' });
  }
  if (!VALID_ACCESS_LEVELS.has(accessLevel)) {
    return jsonResponse(request, env, 400, { error: 'accessLevel không hợp lệ.' });
  }

  const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
  const r2ObjectKey =
    typeof body.r2ObjectKey === 'string' && body.r2ObjectKey.trim().length > 0
      ? body.r2ObjectKey.trim()
      : null;
  if (r2ObjectKey) {
    if (
      !isSafeObjectKey(r2ObjectKey) ||
      (!r2ObjectKey.startsWith(`${authResult.user.id}/`) &&
        !r2ObjectKey.startsWith(`${safeUserId}/`))
    ) {
      return jsonResponse(request, env, 403, {
        error: 'r2ObjectKey không thuộc tài khoản hiện tại.',
      });
    }
  }

  const storagePath =
    typeof body.storagePath === 'string' && body.storagePath.trim().length > 0
      ? body.storagePath.trim()
      : null;
  if (storagePath) {
    if (
      !isSafeObjectKey(storagePath) ||
      (!storagePath.startsWith(`${authResult.user.id}/`) &&
        !storagePath.startsWith(`${safeUserId}/`))
    ) {
      return jsonResponse(request, env, 403, {
        error: 'storagePath không thuộc tài khoản hiện tại.',
      });
    }
  }

  const requestedToken =
    typeof body.shareToken === 'string' && body.shareToken.trim().length > 0
      ? body.shareToken.trim()
      : null;
  if (requestedToken && !/^[a-zA-Z0-9_-]{6,64}$/.test(requestedToken)) {
    return jsonResponse(request, env, 400, { error: 'Định dạng shareToken không hợp lệ.' });
  }

  // Look up existing share for (user_id, document_id) so we preserve shareToken & viewCount
  const existingUrl = `${supabaseCfg.url}/rest/v1/document_shares?user_id=eq.${encodeURIComponent(
    authResult.user.id
  )}&document_id=eq.${encodeURIComponent(documentId)}&select=share_token,password_hash,view_count&limit=1`;

  let existingShare: {
    share_token: string;
    password_hash: string | null;
    view_count: number;
  } | null = null;

  try {
    const existingRes = await fetch(existingUrl, {
      method: 'GET',
      headers: {
        apikey: supabaseCfg.anonKey,
        Authorization: `Bearer ${authResult.user.accessToken}`,
      },
    });
    if (existingRes.ok) {
      const rows: unknown = await existingRes.json();
      if (Array.isArray(rows) && rows.length > 0 && isPlainRecord(rows[0])) {
        const first = rows[0];
        if (typeof first.share_token === 'string') {
          existingShare = {
            share_token: first.share_token,
            password_hash:
              typeof first.password_hash === 'string' ? first.password_hash : null,
            view_count: typeof first.view_count === 'number' ? first.view_count : 0,
          };
        }
      }
    }
  } catch {
    return jsonResponse(request, env, 502, {
      error: 'Không thể kết nối tới Supabase để kiểm tra cấu hình chia sẻ.',
    });
  }

  let nextPasswordHash: string | null = null;
  if (accessLevel === 'password') {
    const providedHash = await resolvePasswordHashInput(body.password, body.passwordHash);
    if (providedHash) {
      nextPasswordHash = providedHash;
    } else if (existingShare?.password_hash) {
      nextPasswordHash = existingShare.password_hash;
    } else {
      return jsonResponse(request, env, 400, {
        error: 'Vui lòng cung cấp mật khẩu khi bật chế độ Khóa mật khẩu.',
      });
    }
  }

  const finalShareToken = requestedToken || existingShare?.share_token || generateShareToken();
  const snapshotPayload = isPlainRecord(body.snapshotPayload) ? body.snapshotPayload : {};
  const nowIso = new Date().toISOString();

  const rowPayload = {
    share_token: finalShareToken,
    document_id: documentId,
    user_id: authResult.user.id,
    title,
    file_type: fileType,
    file_name: typeof body.fileName === 'string' ? body.fileName : null,
    file_size: typeof body.fileSize === 'string' ? body.fileSize : null,
    r2_object_key: r2ObjectKey,
    storage_path: storagePath,
    access_level: accessLevel,
    password_hash: nextPasswordHash,
    allow_download: allowDownload,
    snapshot_payload: snapshotPayload,
    revoked_at: null,
    updated_at: nowIso,
  };

  let saveRes: Response;
  try {
    // If the document already has a share under a different token and caller requested a new token,
    // update by user_id + document_id via PATCH; otherwise upsert cleanly with on_conflict=share_token.
    if (existingShare && existingShare.share_token !== finalShareToken) {
      saveRes = await fetch(
        `${supabaseCfg.url}/rest/v1/document_shares?user_id=eq.${encodeURIComponent(
          authResult.user.id
        )}&document_id=eq.${encodeURIComponent(documentId)}`,
        {
          method: 'PATCH',
          headers: {
            apikey: supabaseCfg.anonKey,
            Authorization: `Bearer ${authResult.user.accessToken}`,
            'Content-Type': 'application/json',
            Prefer: 'return=representation',
          },
          body: JSON.stringify(rowPayload),
        }
      );
    } else {
      saveRes = await fetch(
        `${supabaseCfg.url}/rest/v1/document_shares?on_conflict=share_token`,
        {
          method: 'POST',
          headers: {
            apikey: supabaseCfg.anonKey,
            Authorization: `Bearer ${authResult.user.accessToken}`,
            'Content-Type': 'application/json',
            Prefer: 'resolution=merge-duplicates,return=representation',
          },
          body: JSON.stringify(rowPayload),
        }
      );
    }
  } catch {
    return jsonResponse(request, env, 502, {
      error: 'Không thể lưu cấu hình chia sẻ lên Supabase.',
    });
  }

  if (!saveRes.ok) {
    const errText = await saveRes.text().catch(() => '');
    const status = saveRes.status === 409 ? 409 : saveRes.status >= 500 ? 502 : saveRes.status;
    return jsonResponse(request, env, status, {
      error: errText || `Không thể lưu liên kết chia sẻ (${saveRes.status}).`,
    });
  }

  let savedRaw: unknown;
  try {
    savedRaw = await saveRes.json();
  } catch {
    return jsonResponse(request, env, 502, {
      error: 'Phản hồi lưu chia sẻ từ Supabase không hợp lệ.',
    });
  }

  const savedRow =
    Array.isArray(savedRaw) && savedRaw.length > 0 && isPlainRecord(savedRaw[0])
      ? (savedRaw[0] as unknown as DocumentShareRow)
      : (rowPayload as unknown as DocumentShareRow);

  return jsonResponse(request, env, 200, normalizeShareRowResponse(savedRow));
}

async function formatShareResolutionResponse(
  request: Request,
  env: Env,
  token: string,
  meta: SupabaseShareRpcResult
): Promise<Response> {
  if (meta.status === 'not_found') {
    return jsonResponse(request, env, 404, {
      status: 'not_found',
      shareToken: token,
      error: 'Liên kết chia sẻ không tồn tại hoặc đã bị xóa.',
    });
  }

  if (meta.status === 'revoked') {
    return jsonResponse(request, env, 403, {
      status: 'revoked',
      shareToken: token,
      error: 'Liên kết chia sẻ đã bị thu hồi bởi chủ sở hữu.',
    });
  }

  if (meta.status === 'private') {
    return jsonResponse(request, env, 403, {
      status: 'private',
      shareToken: token,
      error: 'Tài liệu đang ở chế độ Riêng tư.',
    });
  }

  if (meta.status === 'password_required') {
    return jsonResponse(request, env, 200, {
      status: 'password_required',
      shareToken: meta.share_token ?? token,
      title: meta.title ?? undefined,
      fileType: meta.file_type ?? undefined,
      fileName: meta.file_name ?? undefined,
      fileSize: meta.file_size ?? undefined,
      allowDownload:
        typeof meta.allow_download === 'boolean' ? meta.allow_download : undefined,
      viewCount: typeof meta.view_count === 'number' ? meta.view_count : 0,
    });
  }

  if (meta.status === 'password_invalid') {
    return jsonResponse(request, env, 401, {
      status: 'password_invalid',
      shareToken: meta.share_token ?? token,
      title: meta.title ?? undefined,
      fileType: meta.file_type ?? undefined,
      fileName: meta.file_name ?? undefined,
      fileSize: meta.file_size ?? undefined,
      allowDownload:
        typeof meta.allow_download === 'boolean' ? meta.allow_download : undefined,
      viewCount: typeof meta.view_count === 'number' ? meta.view_count : 0,
      error: 'Mật khẩu mở khóa không chính xác.',
    });
  }

  if (meta.status !== 'ok') {
    return jsonResponse(request, env, 403, {
      status: 'error',
      shareToken: token,
      error: 'Không thể truy cập liên kết chia sẻ.',
    });
  }

  const allowDownload = Boolean(meta.allow_download);
  const fileType = meta.file_type ?? 'docx';
  const r2ObjectKey =
    typeof meta.r2_object_key === 'string' && meta.r2_object_key.trim().length > 0
      ? meta.r2_object_key.trim()
      : undefined;

  let presignedFileUrl: string | undefined;
  if (r2ObjectKey && (allowDownload || fileType === 'pdf')) {
    const generated = await generateDownloadAccessUrl(
      request,
      env,
      r2ObjectKey,
      DOWNLOAD_PRESIGN_EXPIRES_SECONDS
    );
    if (generated) {
      presignedFileUrl = generated;
    }
  }

  return jsonResponse(request, env, 200, {
    ok: true,
    status: 'ready',
    shareToken: meta.share_token ?? token,
    documentId: meta.document_id ?? undefined,
    title: meta.title ?? 'Tài liệu chia sẻ',
    fileType,
    fileName: meta.file_name ?? undefined,
    fileSize: meta.file_size ?? undefined,
    allowDownload,
    viewCount: typeof meta.view_count === 'number' ? meta.view_count : 0,
    r2ObjectKey,
    storagePath: meta.storage_path ?? undefined,
    downloadUrl: presignedFileUrl,
    fileUrl: presignedFileUrl,
    viewUrl: presignedFileUrl,
    snapshotPayload: isPlainRecord(meta.snapshot_payload) ? meta.snapshot_payload : {},
  });
}

async function handleGetShare(
  request: Request,
  env: Env,
  token: string
): Promise<Response> {
  if (!/^[a-zA-Z0-9_-]{3,80}$/.test(token)) {
    return jsonResponse(request, env, 400, {
      status: 'error',
      error: 'Mã liên kết chia sẻ không hợp lệ.',
    });
  }

  const rpcResult = await callSupabaseShareRpc(env, 'get_shared_document_meta', {
    p_share_token: token,
  });
  if (!rpcResult.ok) {
    return jsonResponse(request, env, rpcResult.status, {
      status: 'error',
      error: rpcResult.message,
    });
  }

  return formatShareResolutionResponse(request, env, token, rpcResult.data);
}

async function handleUnlockShare(
  request: Request,
  env: Env,
  token: string
): Promise<Response> {
  if (!/^[a-zA-Z0-9_-]{3,80}$/.test(token)) {
    return jsonResponse(request, env, 400, {
      status: 'error',
      error: 'Mã liên kết chia sẻ không hợp lệ.',
    });
  }

  const bodyResult = await readJsonBody(request);
  if (!bodyResult.ok) {
    return jsonResponse(request, env, 400, {
      status: 'error',
      error: bodyResult.message,
    });
  }

  const passwordHash = await resolvePasswordHashInput(
    bodyResult.data.password,
    bodyResult.data.passwordHash
  );
  if (!passwordHash) {
    return jsonResponse(request, env, 400, {
      status: 'error',
      error: 'Vui lòng nhập mật khẩu.',
    });
  }

  const rpcResult = await callSupabaseShareRpc(env, 'unlock_shared_document', {
    p_share_token: token,
    p_password_hash: passwordHash,
  });
  if (!rpcResult.ok) {
    return jsonResponse(request, env, rpcResult.status, {
      status: 'error',
      error: rpcResult.message,
    });
  }

  return formatShareResolutionResponse(request, env, token, rpcResult.data);
}

async function handleDeleteShare(
  request: Request,
  env: Env,
  token: string
): Promise<Response> {
  if (!/^[a-zA-Z0-9_-]{3,80}$/.test(token)) {
    return jsonResponse(request, env, 400, { error: 'Mã liên kết chia sẻ không hợp lệ.' });
  }

  const authResult = await authenticateSupabaseUser(request, env);
  if (!authResult.ok) {
    return jsonResponse(request, env, authResult.status, { error: authResult.message });
  }

  const supabaseCfg = isValidSupabaseConfig(env);
  if (!supabaseCfg) {
    return jsonResponse(request, env, 503, {
      error: 'Worker chưa được cấu hình Supabase.',
    });
  }

  const nowIso = new Date().toISOString();
  let revokeRes: Response;
  try {
    revokeRes = await fetch(
      `${supabaseCfg.url}/rest/v1/document_shares?share_token=eq.${encodeURIComponent(
        token
      )}&user_id=eq.${encodeURIComponent(authResult.user.id)}`,
      {
        method: 'PATCH',
        headers: {
          apikey: supabaseCfg.anonKey,
          Authorization: `Bearer ${authResult.user.accessToken}`,
          'Content-Type': 'application/json',
          Prefer: 'return=representation',
        },
        body: JSON.stringify({
          revoked_at: nowIso,
          access_level: 'private',
          updated_at: nowIso,
        }),
      }
    );
  } catch {
    return jsonResponse(request, env, 502, {
      error: 'Không thể kết nối tới Supabase để thu hồi liên kết.',
    });
  }

  if (!revokeRes.ok) {
    return jsonResponse(
      request,
      env,
      revokeRes.status >= 500 ? 502 : revokeRes.status,
      { error: `Không thể thu hồi liên kết chia sẻ (${revokeRes.status}).` }
    );
  }

  let rows: unknown;
  try {
    rows = await revokeRes.json();
  } catch {
    rows = [];
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    return jsonResponse(request, env, 404, {
      error: 'Không tìm thấy liên kết chia sẻ thuộc tài khoản của bạn.',
    });
  }

  return jsonResponse(request, env, 200, {
    ok: true,
    revoked: true,
    shareToken: token,
  });
}

async function readStreamWithLimit(
  stream: ReadableStream<Uint8Array>,
  maxBytes: number
): Promise<Uint8Array | null> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        break;
      }
      if (!value) {
        continue;
      }
      total += value.byteLength;
      if (total > maxBytes) {
        await reader.cancel();
        return null;
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }

  const combined = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    combined.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return combined;
}

async function handleDirectR2Proxy(
  request: Request,
  env: Env,
  rawEncodedKey: string
): Promise<Response> {
  if (!env.R2_BUCKET) {
    return jsonResponse(request, env, 503, {
      error: 'R2_BUCKET binding chưa được khai báo trên Worker.',
    });
  }

  let objectKey: string;
  try {
    objectKey = rawEncodedKey
      .split('/')
      .map((seg) => decodeURIComponent(seg))
      .join('/');
  } catch {
    return jsonResponse(request, env, 400, { error: 'Khóa đối tượng không hợp lệ.' });
  }

  if (!isSafeObjectKey(objectKey)) {
    return jsonResponse(request, env, 400, { error: 'Khóa đối tượng không hợp lệ.' });
  }

  const url = new URL(request.url);
  const expParam = url.searchParams.get('exp');
  const sigParam = url.searchParams.get('sig');

  if (request.method === 'PUT') {
    let authorized = false;
    if (expParam && sigParam) {
      authorized = await verifyDirectProxySignature(env, 'PUT', objectKey, expParam, sigParam);
    } else {
      const authResult = await authenticateSupabaseUser(request, env);
      if (!authResult.ok) {
        return jsonResponse(request, env, authResult.status, { error: authResult.message });
      }
      const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
      authorized =
        objectKey.startsWith(`${authResult.user.id}/`) ||
        objectKey.startsWith(`${safeUserId}/`);
    }

    if (!authorized) {
      return jsonResponse(request, env, 403, {
        error: 'Không có quyền tải lên khóa R2 này.',
      });
    }

    const lastDot = objectKey.lastIndexOf('.');
    const ext = lastDot > 0 ? objectKey.slice(lastDot).toLowerCase() : '';
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      return jsonResponse(request, env, 400, {
        error: 'Phần mở rộng tệp không nằm trong danh sách cho phép.',
      });
    }

    const contentLengthHeader = request.headers.get('Content-Length');
    if (contentLengthHeader) {
      const declaredSize = Number.parseInt(contentLengthHeader, 10);
      if (!Number.isFinite(declaredSize) || declaredSize < 0) {
        return jsonResponse(request, env, 400, { error: 'Content-Length không hợp lệ.' });
      }
      if (declaredSize > MAX_UPLOAD_BYTES) {
        return jsonResponse(request, env, 413, {
          error: 'Dung lượng tệp vượt quá giới hạn tối đa 50 MB.',
        });
      }
    }

    if (!request.body) {
      return jsonResponse(request, env, 400, { error: 'Thiếu nội dung tệp tải lên.' });
    }

    const contentType =
      request.headers.get('Content-Type')?.trim() || 'application/octet-stream';

    if (contentLengthHeader) {
      await env.R2_BUCKET.put(objectKey, request.body, {
        httpMetadata: { contentType },
      });
    } else {
      const buffered = await readStreamWithLimit(request.body, MAX_UPLOAD_BYTES);
      if (!buffered) {
        return jsonResponse(request, env, 413, {
          error: 'Dung lượng tệp vượt quá giới hạn tối đa 50 MB.',
        });
      }
      await env.R2_BUCKET.put(objectKey, buffered, {
        httpMetadata: { contentType },
      });
    }

    return jsonResponse(request, env, 200, {
      ok: true,
      objectKey,
    });
  }

  if (request.method === 'GET') {
    let authorized = false;

    if (expParam && sigParam) {
      authorized = await verifyDirectProxySignature(env, 'GET', objectKey, expParam, sigParam);
    } else if (extractBearerToken(request)) {
      const authResult = await authenticateSupabaseUser(request, env);
      if (!authResult.ok) {
        return jsonResponse(request, env, authResult.status, { error: authResult.message });
      }
      const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
      authorized =
        objectKey.startsWith(`${authResult.user.id}/`) ||
        objectKey.startsWith(`${safeUserId}/`);
    } else {
      const shareToken = (url.searchParams.get('shareToken') ?? '').trim();
      if (shareToken) {
        const passwordHash = await resolvePasswordHashInput(
          url.searchParams.get('password'),
          url.searchParams.get('passwordHash')
        );
        const rpcResult = passwordHash
          ? await callSupabaseShareRpc(env, 'unlock_shared_document', {
              p_share_token: shareToken,
              p_password_hash: passwordHash,
            })
          : await callSupabaseShareRpc(env, 'get_shared_document_meta', {
              p_share_token: shareToken,
            });
        if (rpcResult.ok && rpcResult.data.status === 'ok' && rpcResult.data.r2_object_key === keyOrEmpty(objectKey)) {
          authorized = Boolean(rpcResult.data.allow_download) || rpcResult.data.file_type === 'pdf';
        }
      }
    }

    if (!authorized) {
      return jsonResponse(request, env, 403, {
        error: 'Liên kết tải trực tiếp không hợp lệ hoặc đã hết hạn.',
      });
    }

    const object = await env.R2_BUCKET.get(objectKey);
    if (!object) {
      return jsonResponse(request, env, 404, { error: 'Không tìm thấy tệp trong R2.' });
    }

    const headers = new Headers(buildCorsHeaders(request, env));
    headers.set(
      'Content-Type',
      object.httpMetadata?.contentType || 'application/octet-stream'
    );
    headers.set('Content-Length', String(object.size));
    headers.set('ETag', object.httpEtag || object.etag);
    headers.set('Cache-Control', 'private, max-age=300');

    return new Response(object.body, {
      status: 200,
      headers,
    });
  }

  if (request.method === 'DELETE') {
    const authResult = await authenticateSupabaseUser(request, env);
    if (!authResult.ok) {
      return jsonResponse(request, env, authResult.status, { error: authResult.message });
    }
    const safeUserId = sanitizePathSegment(authResult.user.id, 'user');
    if (
      !objectKey.startsWith(`${authResult.user.id}/`) &&
      !objectKey.startsWith(`${safeUserId}/`)
    ) {
      return jsonResponse(request, env, 403, {
        error: 'Bạn không có quyền xóa tệp của tài khoản khác.',
      });
    }

    await env.R2_BUCKET.delete(objectKey);
    return jsonResponse(request, env, 200, {
      ok: true,
      deleted: true,
      objectKey,
    });
  }

  return jsonResponse(request, env, 405, { error: 'Method not allowed' });
}

function keyOrEmpty(value: string): string {
  return value.trim();
}

export default {
  async fetch(
    request: Request,
    env: Env,
    _ctx: WorkerExecutionContext
  ): Promise<Response> {
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: buildCorsHeaders(request, env),
      });
    }

    try {
      const url = new URL(request.url);
      const pathname = url.pathname.replace(/\/+$/, '') || '/';

      if (pathname === '/api/health') {
        if (request.method !== 'GET') {
          return jsonResponse(request, env, 405, { error: 'Method not allowed' });
        }
        return handleHealth(request, env);
      }

      if (pathname === '/api/r2/presign-upload') {
        if (request.method !== 'POST') {
          return jsonResponse(request, env, 405, { error: 'Method not allowed' });
        }
        return await handlePresignUpload(request, env);
      }

      if (pathname === '/api/r2/presign-download') {
        if (request.method !== 'GET') {
          return jsonResponse(request, env, 405, { error: 'Method not allowed' });
        }
        return await handlePresignDownload(request, env);
      }

      if (pathname.startsWith('/api/r2/direct/')) {
        const rawKey = url.pathname.slice('/api/r2/direct/'.length);
        return await handleDirectR2Proxy(request, env, rawKey);
      }

      if (pathname === '/api/shares') {
        if (request.method !== 'POST') {
          return jsonResponse(request, env, 405, { error: 'Method not allowed' });
        }
        return await handleUpsertShare(request, env);
      }

      if (pathname.startsWith('/api/shares/')) {
        const rest = pathname.slice('/api/shares/'.length);
        if (rest.endsWith('/unlock')) {
          const tokenPart = rest.slice(0, -'/unlock'.length);
          if (!tokenPart || tokenPart.includes('/')) {
            return jsonResponse(request, env, 404, { error: 'Not found' });
          }
          if (request.method !== 'POST') {
            return jsonResponse(request, env, 405, { error: 'Method not allowed' });
          }
          return await handleUnlockShare(request, env, decodeURIComponent(tokenPart));
        }

        if (!rest.includes('/')) {
          const token = decodeURIComponent(rest);
          if (request.method === 'GET') {
            return await handleGetShare(request, env, token);
          }
          if (request.method === 'DELETE') {
            return await handleDeleteShare(request, env, token);
          }
          return jsonResponse(request, env, 405, { error: 'Method not allowed' });
        }
      }

      return jsonResponse(request, env, 404, { error: 'Endpoint không tồn tại.' });
    } catch {
      return jsonResponse(request, env, 500, {
        error: 'Đã xảy ra lỗi nội bộ trên Worker API.',
      });
    }
  },
};

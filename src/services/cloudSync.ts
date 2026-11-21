import { createThumbnailDataUri } from '../data/mockDocuments';
import { isSupabaseConfigured, supabase } from '../lib/supabase';
import type {
  CloudUserSummary,
  DocumentItem,
  DocumentType,
  OutlineSection,
  PdfPageItem,
  SlidePreviewItem,
  SpreadsheetSheet,
} from '../types/document';

const STORAGE_BUCKET = 'user-files';
const DOCUMENTS_TABLE = 'documents';

const DEFAULT_ACCENT_BY_TYPE: Record<DocumentType, string> = {
  pptx: '#4F46E5',
  docx: '#2563EB',
  xlsx: '#15803D',
  pdf: '#DB2777',
};

export interface CloudParsedContent {
  wordHtml?: string;
  slidesPreview: SlidePreviewItem[];
  sheets?: SpreadsheetSheet[];
  pdfPages?: PdfPageItem[];
  outline: OutlineSection[];
  wordParagraphs?: Record<string, string[]>;
}

interface CloudDocumentRow {
  id: string;
  user_id: string;
  title: string;
  file_type: string;
  file_name: string | null;
  file_size: string | null;
  file_size_bytes: number | null;
  storage_path: string | null;
  r2_object_key: string | null;
  folder: string | null;
  is_favorite: boolean;
  is_shared: boolean;
  page_count: number | null;
  content_preview: string | null;
  parsed_content: unknown;
  created_at: string | null;
  updated_at: string | null;
  last_accessed_at: string | null;
}

function getConfiguredSupabaseClient() {
  if (!isSupabaseConfigured || !supabase) {
    throw new Error(
      'Chế độ đám mây chưa được cấu hình. Vui lòng thiết lập VITE_SUPABASE_URL và VITE_SUPABASE_ANON_KEY.'
    );
  }
  return supabase;
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isValidDocumentType(value: unknown): value is DocumentType {
  return value === 'pptx' || value === 'docx' || value === 'xlsx' || value === 'pdf';
}

function formatByteSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return '0 B';
  }
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function parseFormattedFileSize(fileSize?: string): number | null {
  if (!fileSize) {
    return null;
  }
  const match = fileSize.trim().match(/^([\d.]+)\s*(B|KB|MB|GB)$/i);
  if (!match) {
    return null;
  }
  const amount = Number.parseFloat(match[1]);
  if (!Number.isFinite(amount) || amount < 0) {
    return null;
  }
  const unit = match[2].toUpperCase();
  if (unit === 'B') {
    return Math.round(amount);
  }
  if (unit === 'KB') {
    return Math.round(amount * 1024);
  }
  if (unit === 'MB') {
    return Math.round(amount * 1024 * 1024);
  }
  if (unit === 'GB') {
    return Math.round(amount * 1024 * 1024 * 1024);
  }
  return null;
}

function isValidIsoTimestamp(value?: string | null): value is string {
  if (!value) {
    return false;
  }
  const timestamp = Date.parse(value);
  return !Number.isNaN(timestamp);
}

function sanitizeSlidesForCloud(slides: SlidePreviewItem[]): SlidePreviewItem[] {
  return slides.map((slide) => {
    if (!slide.images || slide.images.length === 0) {
      return slide;
    }
    const persistableImages = slide.images.filter(
      (img) => typeof img.src === 'string' && !img.src.startsWith('blob:')
    );
    const nextSlide: SlidePreviewItem = { ...slide };
    if (persistableImages.length > 0) {
      nextSlide.images = persistableImages;
    } else {
      delete nextSlide.images;
    }
    return nextSlide;
  });
}

function normalizeWordParagraphs(
  value: unknown
): Record<string, string[]> | undefined {
  if (!isPlainRecord(value)) {
    return undefined;
  }
  const result: Record<string, string[]> = {};
  let hasEntry = false;
  for (const [key, entry] of Object.entries(value)) {
    if (Array.isArray(entry)) {
      result[key] = entry.filter((item): item is string => typeof item === 'string');
      hasEntry = true;
    }
  }
  return hasEntry ? result : undefined;
}

function normalizeParsedContent(raw: unknown): CloudParsedContent {
  if (!isPlainRecord(raw)) {
    return {
      slidesPreview: [],
      outline: [],
    };
  }

  const rawSlides = Array.isArray(raw.slidesPreview)
    ? (raw.slidesPreview as SlidePreviewItem[])
    : [];
  const rawOutline = Array.isArray(raw.outline)
    ? (raw.outline as OutlineSection[])
    : [];

  const normalized: CloudParsedContent = {
    slidesPreview: sanitizeSlidesForCloud(rawSlides),
    outline: rawOutline,
  };

  if (typeof raw.wordHtml === 'string') {
    normalized.wordHtml = raw.wordHtml;
  }
  if (Array.isArray(raw.sheets)) {
    normalized.sheets = raw.sheets as SpreadsheetSheet[];
  }
  if (Array.isArray(raw.pdfPages)) {
    normalized.pdfPages = raw.pdfPages as PdfPageItem[];
  }
  const normalizedParagraphs = normalizeWordParagraphs(raw.wordParagraphs);
  if (normalizedParagraphs) {
    normalized.wordParagraphs = normalizedParagraphs;
  }

  return normalized;
}

function sanitizeStorageFileName(originalName: string): string {
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

  return safeExt ? `${safeBase}.${safeExt}` : safeBase;
}

function sanitizePathSegment(segment: string, fallback: string): string {
  const cleaned = segment
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, '_')
    .replace(/\.{2,}/g, '_')
    .slice(0, 120);
  return cleaned || fallback;
}

function mapCloudRowToDocumentItem(row: CloudDocumentRow): DocumentItem {
  const docType: DocumentType = isValidDocumentType(row.file_type)
    ? row.file_type
    : 'docx';
  const parsed = normalizeParsedContent(row.parsed_content);
  const title = row.title?.trim() || 'Tài liệu chưa đặt tên';
  const accentFromSlide = parsed.slidesPreview[0]?.accentColor;
  const accentColor =
    docType === 'pptx' &&
    typeof accentFromSlide === 'string' &&
    /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(accentFromSlide.trim())
      ? accentFromSlide.trim()
      : DEFAULT_ACCENT_BY_TYPE[docType];

  const timestampForDisplay = row.updated_at ?? row.last_accessed_at ?? row.created_at;
  const computedFileSize =
    row.file_size ??
    (typeof row.file_size_bytes === 'number'
      ? formatByteSize(row.file_size_bytes)
      : undefined);

  const doc: DocumentItem = {
    id: row.id,
    title,
    type: docType,
    updatedAt: formatRelativeVietnameseTime(timestampForDisplay),
    isFavorite: Boolean(row.is_favorite),
    isShared: Boolean(row.is_shared),
    thumbnail: createThumbnailDataUri(docType, title, accentColor),
    slidesPreview: parsed.slidesPreview,
    outline: parsed.outline,
    contentPreview: row.content_preview ?? '',
    pageCount: typeof row.page_count === 'number' ? row.page_count : undefined,
    fileSize: computedFileSize,
    folder: row.folder ?? undefined,
    sheets: parsed.sheets,
    pdfPages: parsed.pdfPages,
    wordParagraphs: parsed.wordParagraphs,
    wordHtml: parsed.wordHtml,
    isUploaded: true,
    lastSavedAt: row.updated_at ?? undefined,
    lastAccessedAt: row.last_accessed_at ?? undefined,
    storagePath: row.storage_path ?? undefined,
    r2ObjectKey: row.r2_object_key ?? undefined,
    originalFileName: row.file_name ?? undefined,
    isSyncedToCloud: true,
    userId: row.user_id,
  };

  return doc;
}

/**
 * Translates common Supabase Auth error messages into clear Vietnamese feedback.
 */
export function mapAuthErrorToVietnamese(message: string): string {
  const normalized = message.trim().toLowerCase();
  if (!normalized) {
    return 'Đã xảy ra lỗi xác thực. Vui lòng thử lại.';
  }

  if (
    normalized.includes('invalid login credentials') ||
    normalized.includes('invalid_credentials') ||
    normalized.includes('invalid email or password')
  ) {
    return 'Email hoặc mật khẩu chưa chính xác. Vui lòng kiểm tra lại.';
  }

  if (normalized.includes('email not confirmed')) {
    return 'Email chưa được xác nhận. Vui lòng kiểm tra hộp thư (kể cả mục Spam) và bấm liên kết xác thực.';
  }

  if (
    normalized.includes('user already registered') ||
    normalized.includes('already been registered') ||
    normalized.includes('email already exists') ||
    normalized.includes('already exists')
  ) {
    return 'Email này đã được đăng ký. Vui lòng chuyển sang tab Đăng nhập.';
  }

  if (
    normalized.includes('password should be at least') ||
    normalized.includes('weak password') ||
    normalized.includes('at least 6 characters')
  ) {
    return 'Mật khẩu quá ngắn. Vui lòng nhập tối thiểu 6 ký tự.';
  }

  if (
    normalized.includes('unable to validate email address') ||
    normalized.includes('invalid email') ||
    normalized.includes('email address') && normalized.includes('invalid')
  ) {
    return 'Địa chỉ email không hợp lệ. Vui lòng kiểm tra lại định dạng.';
  }

  if (
    normalized.includes('rate limit') ||
    normalized.includes('too many requests') ||
    normalized.includes('only request this after') ||
    normalized.includes('over_email_send_rate_limit') ||
    normalized.includes('over_request_rate_limit')
  ) {
    return 'Bạn thao tác quá nhanh. Vui lòng đợi ít phút rồi thử lại.';
  }

  if (
    normalized.includes('failed to fetch') ||
    normalized.includes('networkerror') ||
    normalized.includes('network request failed') ||
    normalized.includes('load failed') ||
    normalized.includes('timeout') ||
    normalized.includes('network')
  ) {
    return 'Không thể kết nối tới máy chủ. Vui lòng kiểm tra kết nối mạng và thử lại.';
  }

  if (normalized.includes('signups not allowed')) {
    return 'Hệ thống hiện tạm khóa đăng ký tài khoản mới.';
  }

  if (
    normalized.includes('auth session missing') ||
    normalized.includes('refresh token') ||
    normalized.includes('jwt expired')
  ) {
    return 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.';
  }

  return `Không thể hoàn tất xác thực (${message.trim()}). Vui lòng thử lại.`;
}

/**
 * Builds a concise user summary from a Supabase Auth user object.
 */
export function buildCloudUserSummary(user: {
  id: string;
  email?: string | null;
  user_metadata?: Record<string, unknown>;
}): CloudUserSummary {
  const email = (user.email ?? '').trim();
  const metadata = user.user_metadata;

  let displayName = '';
  if (metadata) {
    const candidateKeys = ['full_name', 'name', 'display_name', 'username'];
    for (const key of candidateKeys) {
      const candidate = metadata[key];
      if (typeof candidate === 'string' && candidate.trim().length > 0) {
        displayName = candidate.trim();
        break;
      }
    }
  }

  if (!displayName) {
    const emailPrefix = email.includes('@') ? email.split('@')[0].trim() : email;
    displayName = emailPrefix || 'Người dùng LITE';
  }

  return {
    id: user.id,
    email,
    displayName,
  };
}

/**
 * Formats an ISO timestamp as relative Vietnamese time:
 * 'Vừa xong', 'X phút trước', 'Hôm nay, HH:MM', or 'dd/MM/yyyy HH:MM'.
 */
export function formatRelativeVietnameseTime(isoString?: string | null): string {
  if (!isoString) {
    return 'Chưa đồng bộ';
  }
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) {
    return 'Chưa đồng bộ';
  }

  const now = new Date();
  const diffMs = now.getTime() - date.getTime();

  if (diffMs < 60_000 && diffMs >= -60_000) {
    return 'Vừa xong';
  }

  if (diffMs >= 60_000 && diffMs < 3_600_000) {
    const minutes = Math.floor(diffMs / 60_000);
    return `${minutes} phút trước`;
  }

  const isSameDay =
    date.getFullYear() === now.getFullYear() &&
    date.getMonth() === now.getMonth() &&
    date.getDate() === now.getDate();

  if (isSameDay) {
    const timeFormatted = new Intl.DateTimeFormat('vi-VN', {
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    }).format(date);
    return `Hôm nay, ${timeFormatted}`;
  }

  const parts = new Intl.DateTimeFormat('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(date);

  const partMap = new Map<string, string>();
  for (const part of parts) {
    partMap.set(part.type, part.value);
  }

  const day = partMap.get('day') ?? '01';
  const month = partMap.get('month') ?? '01';
  const year = partMap.get('year') ?? String(date.getFullYear());
  const hour = partMap.get('hour') ?? '00';
  const minute = partMap.get('minute') ?? '00';

  return `${day}/${month}/${year} ${hour}:${minute}`;
}

/**
 * Fetches all cloud documents belonging to the user ordered by most recently accessed.
 */
export async function fetchUserDocuments(userId: string): Promise<DocumentItem[]> {
  const client = getConfiguredSupabaseClient();
  const { data, error } = await client
    .from(DOCUMENTS_TABLE)
    .select('*')
    .eq('user_id', userId)
    .order('last_accessed_at', { ascending: false });

  if (error) {
    throw new Error(
      `Không thể tải danh sách tài liệu từ đám mây: ${error.message}`
    );
  }

  const rows = Array.isArray(data) ? (data as CloudDocumentRow[]) : [];
  return rows.map(mapCloudRowToDocumentItem);
}

/**
 * Sanitizes file name into ASCII-safe Storage key and uploads file to 'user-files' bucket
 * at {user_id}/{document_id}/{file_name}.
 */
export async function uploadDocumentFileToStorage(
  userId: string,
  docId: string,
  file: File
): Promise<{ storagePath: string; fileName: string }> {
  const client = getConfiguredSupabaseClient();
  const safeUserId = sanitizePathSegment(userId, 'user');
  const safeDocId = sanitizePathSegment(docId, 'doc');
  const safeFileName = sanitizeStorageFileName(file.name);
  const storagePath = `${safeUserId}/${safeDocId}/${safeFileName}`;

  const { error } = await client.storage
    .from(STORAGE_BUCKET)
    .upload(storagePath, file, {
      upsert: true,
      contentType: file.type || undefined,
    });

  if (error) {
    throw new Error(
      `Không thể tải tệp "${file.name}" lên kho lưu trữ đám mây: ${error.message}`
    );
  }

  return {
    storagePath,
    fileName: safeFileName,
  };
}

/**
 * Downloads a file from Supabase Storage 'user-files' bucket and returns a browser Object URL + Blob.
 */
export async function getCloudFileObjectUrl(
  storagePath: string
): Promise<{ objectUrl: string; blob: Blob }> {
  if (!storagePath.trim()) {
    throw new Error('Đường dẫn tệp trên đám mây không hợp lệ.');
  }
  const client = getConfiguredSupabaseClient();
  const { data, error } = await client.storage
    .from(STORAGE_BUCKET)
    .download(storagePath);

  if (error || !data) {
    throw new Error(
      `Không thể tải tệp gốc từ đám mây${error?.message ? `: ${error.message}` : '.'}`
    );
  }

  const objectUrl = URL.createObjectURL(data);
  return {
    objectUrl,
    blob: data,
  };
}

/**
 * Downloads a cloud file from Supabase Storage and triggers browser file download.
 */
export async function downloadCloudFileToDevice(
  storagePath: string,
  downloadFileName: string
): Promise<void> {
  const { objectUrl } = await getCloudFileObjectUrl(storagePath);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = downloadFileName.trim() || 'tai-lieu-lite';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);

  try {
    anchor.click();
  } finally {
    document.body.removeChild(anchor);
    window.setTimeout(() => {
      URL.revokeObjectURL(objectUrl);
    }, 1000);
  }
}

/**
 * Uploads the source file if needed and upserts the document record into public.documents.
 */
export async function upsertCloudDocument(
  userId: string,
  doc: DocumentItem,
  options?: {
    file?: File | null;
    touchUpdatedAt?: boolean;
    forceUploadFile?: boolean;
  }
): Promise<DocumentItem> {
  const client = getConfiguredSupabaseClient();
  const hasExplicitFileOption =
    options !== undefined &&
    Object.prototype.hasOwnProperty.call(options, 'file');
  const sourceFile = hasExplicitFileOption
    ? (options.file ?? null)
    : (doc.sourceFile ?? null);
  let nextStoragePath = doc.storagePath;
  let nextOriginalFileName = doc.originalFileName ?? sourceFile?.name;

  if (
    sourceFile &&
    ((!nextStoragePath && !doc.r2ObjectKey) || options?.forceUploadFile)
  ) {
    const uploaded = await uploadDocumentFileToStorage(userId, doc.id, sourceFile);
    nextStoragePath = uploaded.storagePath;
    if (options?.forceUploadFile || !nextOriginalFileName) {
      nextOriginalFileName = sourceFile.name;
    }
  }

  const sanitizedSlides = sanitizeSlidesForCloud(doc.slidesPreview);
  const parsedContent: CloudParsedContent = {
    slidesPreview: sanitizedSlides,
    outline: doc.outline ?? [],
  };
  if (typeof doc.wordHtml === 'string') {
    parsedContent.wordHtml = doc.wordHtml;
  }
  if (doc.sheets) {
    parsedContent.sheets = doc.sheets;
  }
  if (doc.pdfPages) {
    parsedContent.pdfPages = doc.pdfPages;
  }
  if (doc.wordParagraphs) {
    parsedContent.wordParagraphs = doc.wordParagraphs;
  }

  const touchUpdatedAt = options?.touchUpdatedAt !== false;
  const nowIso = new Date().toISOString();
  const updatedAtIso =
    touchUpdatedAt || !isValidIsoTimestamp(doc.lastSavedAt)
      ? nowIso
      : doc.lastSavedAt;
  const lastAccessedIso =
    touchUpdatedAt || !isValidIsoTimestamp(doc.lastAccessedAt)
      ? nowIso
      : doc.lastAccessedAt;

  const computedFileSize =
    doc.fileSize ?? (sourceFile ? formatByteSize(sourceFile.size) : undefined);
  const fileSizeBytes = sourceFile
    ? sourceFile.size
    : parseFormattedFileSize(computedFileSize);

  const rowPayload = {
    id: doc.id,
    user_id: userId,
    title: doc.title.trim() || 'Tài liệu mới',
    file_type: doc.type,
    file_name: nextOriginalFileName ?? null,
    file_size: computedFileSize ?? null,
    file_size_bytes: fileSizeBytes,
    storage_path: nextStoragePath ?? null,
    r2_object_key: doc.r2ObjectKey ?? null,
    folder: doc.folder ?? null,
    is_favorite: Boolean(doc.isFavorite),
    is_shared: Boolean(doc.isShared),
    page_count: typeof doc.pageCount === 'number' ? doc.pageCount : null,
    content_preview: doc.contentPreview ?? '',
    parsed_content: parsedContent,
    updated_at: updatedAtIso,
    last_accessed_at: lastAccessedIso,
  };

  const { error } = await client
    .from(DOCUMENTS_TABLE)
    .upsert(rowPayload, { onConflict: 'user_id,id' });

  if (error) {
    throw new Error(
      `Không thể đồng bộ tài liệu "${doc.title}" lên đám mây: ${error.message}`
    );
  }

  return {
    ...doc,
    fileSize: computedFileSize,
    storagePath: nextStoragePath,
    originalFileName: nextOriginalFileName,
    lastSavedAt: updatedAtIso,
    lastAccessedAt: lastAccessedIso,
    updatedAt: touchUpdatedAt
      ? formatRelativeVietnameseTime(updatedAtIso)
      : doc.updatedAt,
    isSyncedToCloud: true,
    isUploaded: true,
    userId,
  };
}

/**
 * Updates last_accessed_at for a synced document when the user opens it.
 */
export async function touchDocumentLastAccessed(
  userId: string,
  docId: string
): Promise<void> {
  const client = getConfiguredSupabaseClient();
  const { error } = await client
    .from(DOCUMENTS_TABLE)
    .update({ last_accessed_at: new Date().toISOString() })
    .eq('user_id', userId)
    .eq('id', docId);

  if (error) {
    throw new Error(
      `Không thể cập nhật thời gian truy cập gần nhất: ${error.message}`
    );
  }
}

/**
 * Sequentially syncs local documents to the user's Supabase account.
 * Returns only documents that succeeded along with the count of failures.
 */
export async function syncLocalDocumentsToCloud(
  userId: string,
  docs: DocumentItem[]
): Promise<{ syncedDocs: DocumentItem[]; failedCount: number }> {
  const syncedDocs: DocumentItem[] = [];
  let failedCount = 0;

  for (const doc of docs) {
    try {
      const syncedDoc = await upsertCloudDocument(userId, doc, {
        file: doc.sourceFile ?? null,
        touchUpdatedAt: false,
      });
      syncedDocs.push(syncedDoc);
    } catch {
      failedCount += 1;
    }
  }

  return {
    syncedDocs,
    failedCount,
  };
}

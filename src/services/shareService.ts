import { createThumbnailDataUri } from '../data/mockDocuments';
import { isSupabaseConfigured, supabase } from '../lib/supabase';
import { getCloudFileObjectUrl } from './cloudSync';
import {
  getR2ApiBaseUrl,
  getR2PresignedDownloadUrl,
  isR2ApiConfigured,
} from './r2Storage';
import type {
  DocumentItem,
  DocumentType,
  OutlineSection,
  PdfPageItem,
  ShareAccessLevel,
  ShareConfigRecord,
  SharedDocumentResolution,
  SlidePreviewItem,
  SpreadsheetSheet,
} from '../types/document';

const LOCAL_SHARES_KEY = 'lite_local_shares_v1';
const SHARES_TABLE = 'document_shares';

const DEFAULT_ACCENT_BY_TYPE: Record<DocumentType, string> = {
  pptx: '#4F46E5',
  docx: '#2563EB',
  xlsx: '#15803D',
  pdf: '#DB2777',
};

interface LocalShareRecord extends ShareConfigRecord {
  title: string;
  fileType: DocumentType;
  fileName?: string;
  fileSize?: string;
  passwordHash?: string;
  snapshotPayload: Record<string, unknown>;
}

interface SupabaseShareRow {
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
  allow_download: boolean;
  view_count: number | null;
  revoked_at: string | null;
  updated_at: string | null;
}

interface ShareRpcPayload {
  status?: string;
  share_token?: string;
  shareToken?: string;
  document_id?: string;
  documentId?: string;
  title?: string;
  file_type?: string;
  fileType?: string;
  file_name?: string | null;
  fileName?: string | null;
  file_size?: string | null;
  fileSize?: string | null;
  allow_download?: boolean;
  allowDownload?: boolean;
  view_count?: number;
  viewCount?: number;
  r2_object_key?: string | null;
  r2ObjectKey?: string | null;
  storage_path?: string | null;
  storagePath?: string | null;
  downloadUrl?: string | null;
  fileUrl?: string | null;
  snapshot_payload?: unknown;
  snapshotPayload?: unknown;
  error?: string;
  message?: string;
}

const memoryShareFallback = new Map<string, LocalShareRecord>();

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isValidDocumentType(value: unknown): value is DocumentType {
  return value === 'pptx' || value === 'docx' || value === 'xlsx' || value === 'pdf';
}

function isValidAccessLevel(value: unknown): value is ShareAccessLevel {
  return value === 'private' || value === 'link' || value === 'password';
}

function sanitizeSlidesForShare(slides: SlidePreviewItem[]): SlidePreviewItem[] {
  return slides.map((slide) => {
    if (!slide.images || slide.images.length === 0) {
      return slide;
    }
    const cleanImages = slide.images.filter(
      (img) => typeof img.src === 'string' && !img.src.startsWith('blob:')
    );
    const nextSlide: SlidePreviewItem = { ...slide };
    if (cleanImages.length > 0) {
      nextSlide.images = cleanImages;
    } else {
      delete nextSlide.images;
    }
    return nextSlide;
  });
}

function generateShareToken(): string {
  const alphabet = 'abcdefghjkmnpqrstuvwxyz23456789';
  const length = 12;
  let suffix = '';
  if (typeof globalThis.crypto?.getRandomValues === 'function') {
    const bytes = new Uint8Array(length);
    globalThis.crypto.getRandomValues(bytes);
    for (let i = 0; i < length; i += 1) {
      suffix += alphabet[bytes[i] % alphabet.length];
    }
  } else {
    for (let i = 0; i < length; i += 1) {
      suffix += alphabet[Math.floor(Math.random() * alphabet.length)];
    }
  }
  return `lite-${suffix}`;
}

/** Pure TypeScript SHA-256 fallback for insecure HTTP LAN contexts where crypto.subtle is unavailable. */
function sha256HexFallback(input: string): string {
  const utf8 = new TextEncoder().encode(input);
  const K = new Uint32Array([
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1,
    0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
    0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786,
    0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
    0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
    0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b,
    0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a,
    0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
  ]);
  const H = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);

  const bitLen = utf8.length * 8;
  const paddedLen = Math.ceil((utf8.length + 9) / 64) * 64;
  const padded = new Uint8Array(paddedLen);
  padded.set(utf8);
  padded[utf8.length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(paddedLen - 8, Math.floor(bitLen / 0x100000000), false);
  view.setUint32(paddedLen - 4, bitLen >>> 0, false);

  const w = new Uint32Array(64);
  const rotr = (x: number, n: number) => (x >>> n) | (x << (32 - n));

  for (let offset = 0; offset < paddedLen; offset += 64) {
    for (let i = 0; i < 16; i += 1) {
      w[i] = view.getUint32(offset + i * 4, false);
    }
    for (let i = 16; i < 64; i += 1) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }

    let [a, b, c, d, e, f, g, h] = H;
    for (let i = 0; i < 64; i += 1) {
      const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + K[i] + w[i]) >>> 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;

      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    H[0] = (H[0] + a) >>> 0;
    H[1] = (H[1] + b) >>> 0;
    H[2] = (H[2] + c) >>> 0;
    H[3] = (H[3] + d) >>> 0;
    H[4] = (H[4] + e) >>> 0;
    H[5] = (H[5] + f) >>> 0;
    H[6] = (H[6] + g) >>> 0;
    H[7] = (H[7] + h) >>> 0;
  }

  return Array.from(H)
    .map((word) => word.toString(16).padStart(8, '0'))
    .join('');
}

/**
 * Hashes a share password using SHA-256 hex (Web Crypto with pure-TS fallback).
 */
export async function hashSharePassword(password: string): Promise<string> {
  if (typeof globalThis.crypto?.subtle?.digest === 'function') {
    const encoded = new TextEncoder().encode(password);
    const hashBuffer = await globalThis.crypto.subtle.digest('SHA-256', encoded);
    return Array.from(new Uint8Array(hashBuffer))
      .map((byte) => byte.toString(16).padStart(2, '0'))
      .join('');
  }
  return sha256HexFallback(password);
}

/**
 * Builds the public hash URL for sharing a document without requiring sign-in.
 */
export function buildPublicShareUrl(shareToken: string): string {
  const baseUrl = `${window.location.origin}${window.location.pathname}`;
  return `${baseUrl}#share/${encodeURIComponent(shareToken)}`;
}

/**
 * Builds a serializable snapshot of readable document content for public share links.
 */
export function buildShareSnapshotPayload(
  doc: DocumentItem
): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    id: doc.id,
    title: doc.title,
    fileType: doc.type,
    fileName: doc.originalFileName ?? null,
    fileSize: doc.fileSize ?? null,
    folder: doc.folder ?? null,
    pageCount: typeof doc.pageCount === 'number' ? doc.pageCount : null,
    contentPreview: doc.contentPreview ?? '',
    slidesPreview: sanitizeSlidesForShare(doc.slidesPreview ?? []),
    outline: doc.outline ?? [],
  };

  if (typeof doc.wordHtml === 'string') {
    payload.wordHtml = doc.wordHtml;
  }
  if (Array.isArray(doc.sheets)) {
    payload.sheets = doc.sheets;
  }
  if (Array.isArray(doc.pdfPages)) {
    payload.pdfPages = doc.pdfPages;
  }
  if (doc.wordParagraphs) {
    payload.wordParagraphs = doc.wordParagraphs;
  }

  return payload;
}

function readLocalShareMap(): Record<string, LocalShareRecord> {
  const merged: Record<string, LocalShareRecord> = {};
  memoryShareFallback.forEach((value, key) => {
    merged[key] = value;
  });

  try {
    const raw = window.localStorage.getItem(LOCAL_SHARES_KEY);
    if (!raw) {
      return merged;
    }
    const parsed: unknown = JSON.parse(raw);
    if (!isPlainRecord(parsed)) {
      return merged;
    }
    for (const [token, entry] of Object.entries(parsed)) {
      if (isPlainRecord(entry) && typeof entry.shareToken === 'string') {
        merged[token] = entry as unknown as LocalShareRecord;
      }
    }
  } catch {
    // Ignore storage read errors and use memory fallback
  }

  return merged;
}

function persistLocalShareRecord(record: LocalShareRecord): void {
  memoryShareFallback.set(record.shareToken, record);
  const current = readLocalShareMap();

  // Keep a single latest local share record per documentId to avoid stale tokens
  for (const [token, existing] of Object.entries(current)) {
    if (existing.documentId === record.documentId && token !== record.shareToken) {
      delete current[token];
      memoryShareFallback.delete(token);
    }
  }
  current[record.shareToken] = record;

  try {
    window.localStorage.setItem(LOCAL_SHARES_KEY, JSON.stringify(current));
  } catch {
    // Retry with minimal snapshot if localStorage quota is tight
    try {
      const compactRecord: LocalShareRecord = {
        ...record,
        snapshotPayload: {
          id: record.documentId,
          title: record.title,
          fileType: record.fileType,
          fileSize: record.fileSize ?? null,
          contentPreview:
            typeof record.snapshotPayload.contentPreview === 'string'
              ? record.snapshotPayload.contentPreview
              : '',
          outline: Array.isArray(record.snapshotPayload.outline)
            ? record.snapshotPayload.outline
            : [],
          slidesPreview: [],
        },
      };
      current[record.shareToken] = compactRecord;
      window.localStorage.setItem(LOCAL_SHARES_KEY, JSON.stringify(current));
    } catch {
      // Memory fallback already holds the full record for this session
    }
  }
}

function findLocalShareByDocumentId(documentId: string): LocalShareRecord | null {
  const all = readLocalShareMap();
  const matches = Object.values(all).filter(
    (item) => item.documentId === documentId
  );
  if (matches.length === 0) {
    return null;
  }
  matches.sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
  return matches[0];
}

function findLocalShareByToken(shareToken: string): LocalShareRecord | null {
  const all = readLocalShareMap();
  return all[shareToken] ?? null;
}

async function getCurrentAuthContext(): Promise<{
  userId: string;
  accessToken: string;
} | null> {
  if (!isSupabaseConfigured || !supabase) {
    return null;
  }
  try {
    const { data } = await supabase.auth.getSession();
    const session = data.session;
    if (!session?.user?.id || !session.access_token) {
      return null;
    }
    return {
      userId: session.user.id,
      accessToken: session.access_token,
    };
  } catch {
    return null;
  }
}

function mapSupabaseRowToShareConfig(row: SupabaseShareRow): ShareConfigRecord {
  return {
    shareToken: row.share_token,
    documentId: row.document_id,
    accessLevel: isValidAccessLevel(row.access_level)
      ? row.access_level
      : 'private',
    allowDownload: Boolean(row.allow_download),
    r2ObjectKey: row.r2_object_key ?? undefined,
    storagePath: row.storage_path ?? undefined,
    viewCount: typeof row.view_count === 'number' ? row.view_count : 0,
    revokedAt: row.revoked_at ?? null,
    updatedAt: row.updated_at ?? new Date().toISOString(),
  };
}

function buildDocumentFromSnapshot(params: {
  shareToken: string;
  documentId?: string;
  title?: string;
  fileType?: DocumentType;
  fileName?: string | null;
  fileSize?: string | null;
  r2ObjectKey?: string | null;
  storagePath?: string | null;
  allowDownload?: boolean;
  downloadUrl?: string;
  snapshotPayload: unknown;
}): DocumentItem {
  const snap = isPlainRecord(params.snapshotPayload)
    ? params.snapshotPayload
    : {};

  const resolvedType: DocumentType = isValidDocumentType(params.fileType)
    ? params.fileType
    : isValidDocumentType(snap.fileType)
      ? snap.fileType
      : 'docx';

  const resolvedTitle =
    (typeof params.title === 'string' && params.title.trim()) ||
    (typeof snap.title === 'string' && snap.title.trim()) ||
    'Tài liệu được chia sẻ';

  const slidesPreview = Array.isArray(snap.slidesPreview)
    ? sanitizeSlidesForShare(snap.slidesPreview as SlidePreviewItem[])
    : [];
  const outline = Array.isArray(snap.outline)
    ? (snap.outline as OutlineSection[])
    : [];

  const accentFromSlide = slidesPreview[0]?.accentColor;
  const accentColor =
    resolvedType === 'pptx' &&
    typeof accentFromSlide === 'string' &&
    /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(accentFromSlide.trim())
      ? accentFromSlide.trim()
      : DEFAULT_ACCENT_BY_TYPE[resolvedType];

  const doc: DocumentItem = {
    id:
      (typeof params.documentId === 'string' && params.documentId) ||
      (typeof snap.id === 'string' && snap.id) ||
      `shared-${params.shareToken}`,
    title: resolvedTitle,
    type: resolvedType,
    updatedAt: 'Liên kết chia sẻ',
    isFavorite: false,
    isShared: true,
    thumbnail: createThumbnailDataUri(resolvedType, resolvedTitle, accentColor),
    slidesPreview,
    outline,
    contentPreview:
      typeof snap.contentPreview === 'string' ? snap.contentPreview : '',
    pageCount: typeof snap.pageCount === 'number' ? snap.pageCount : undefined,
    fileSize:
      params.fileSize ??
      (typeof snap.fileSize === 'string' ? snap.fileSize : undefined),
    folder: typeof snap.folder === 'string' ? snap.folder : undefined,
    originalFileName:
      params.fileName ??
      (typeof snap.fileName === 'string' ? snap.fileName : undefined),
    r2ObjectKey: params.r2ObjectKey ?? undefined,
    storagePath: params.storagePath ?? undefined,
    shareToken: params.shareToken,
    shareAllowDownload: params.allowDownload,
    fileUrl: params.downloadUrl,
    objectUrl: params.downloadUrl,
    wordHtml: typeof snap.wordHtml === 'string' ? snap.wordHtml : undefined,
    sheets: Array.isArray(snap.sheets)
      ? (snap.sheets as SpreadsheetSheet[])
      : undefined,
    pdfPages: Array.isArray(snap.pdfPages)
      ? (snap.pdfPages as PdfPageItem[])
      : undefined,
    wordParagraphs: isPlainRecord(snap.wordParagraphs)
      ? (snap.wordParagraphs as Record<string, string[]>)
      : undefined,
  };

  return doc;
}

async function resolveDownloadUrlForShare(params: {
  r2ObjectKey?: string | null;
  storagePath?: string | null;
  shareToken: string;
  allowDownload: boolean;
  fileType?: DocumentType;
  passwordHash?: string;
  explicitDownloadUrl?: string | null;
}): Promise<string | undefined> {
  if (params.explicitDownloadUrl && params.explicitDownloadUrl.trim()) {
    const raw = params.explicitDownloadUrl.trim();
    if (/^https?:\/\//i.test(raw) || raw.startsWith('blob:')) {
      return raw;
    }
    if (isR2ApiConfigured) {
      const normalizedPath = raw.startsWith('/') ? raw : `/${raw}`;
      return `${getR2ApiBaseUrl()}${normalizedPath}`;
    }
  }

  if (!params.allowDownload && params.fileType !== 'pdf') {
    return undefined;
  }

  if (params.r2ObjectKey && isR2ApiConfigured) {
    try {
      const presigned = await getR2PresignedDownloadUrl({
        objectKey: params.r2ObjectKey,
        shareToken: params.shareToken,
        passwordHash: params.passwordHash,
      });
      return presigned.downloadUrl;
    } catch {
      // Fallback gracefully if R2 presign is unreachable
    }
  }

  if (params.storagePath && isSupabaseConfigured && supabase) {
    try {
      const authCtx = await getCurrentAuthContext();
      if (authCtx) {
        const { objectUrl } = await getCloudFileObjectUrl(params.storagePath);
        return objectUrl;
      }
    } catch {
      // Viewer may be anonymous and blocked by Storage RLS; snapshot still renders
    }
  }

  return undefined;
}

async function mapRpcPayloadToResolution(
  shareToken: string,
  payload: ShareRpcPayload,
  passwordHash?: string
): Promise<SharedDocumentResolution> {
  const token = payload.share_token ?? payload.shareToken ?? shareToken;
  const status = payload.status ?? 'error';
  const fileTypeRaw = payload.file_type ?? payload.fileType;
  const fileType = isValidDocumentType(fileTypeRaw) ? fileTypeRaw : undefined;
  const title = payload.title;
  const fileName = (payload.file_name ?? payload.fileName) ?? undefined;
  const fileSize = (payload.file_size ?? payload.fileSize) ?? undefined;
  const allowDownload =
    payload.allow_download ?? payload.allowDownload ?? true;
  const viewCount = payload.view_count ?? payload.viewCount ?? 0;
  const r2ObjectKey =
    (payload.r2_object_key ?? payload.r2ObjectKey) ?? undefined;
  const storagePath = (payload.storage_path ?? payload.storagePath) ?? undefined;

  if (status === 'not_found') {
    return { status: 'not_found', shareToken: token };
  }
  if (status === 'revoked') {
    return { status: 'revoked', shareToken: token, title, fileType };
  }
  if (status === 'private') {
    return { status: 'private', shareToken: token, title, fileType };
  }
  if (status === 'password_required') {
    return {
      status: 'password_required',
      shareToken: token,
      title,
      fileType,
      fileName,
      fileSize,
      allowDownload,
      viewCount,
    };
  }
  if (status === 'password_invalid') {
    return {
      status: 'password_required',
      shareToken: token,
      title,
      fileType,
      fileName,
      fileSize,
      allowDownload,
      viewCount,
      errorMessage: 'Mật khẩu chưa chính xác. Vui lòng thử lại.',
    };
  }

  if (status === 'ok' || status === 'ready') {
    const snapshot = payload.snapshot_payload ?? payload.snapshotPayload;
    const resolvedFileType =
      fileType ??
      (isPlainRecord(snapshot) && isValidDocumentType(snapshot.fileType)
        ? snapshot.fileType
        : undefined);

    const downloadUrl = await resolveDownloadUrlForShare({
      r2ObjectKey,
      storagePath,
      shareToken: token,
      allowDownload: Boolean(allowDownload),
      fileType: resolvedFileType,
      passwordHash,
      explicitDownloadUrl: payload.downloadUrl ?? payload.fileUrl,
    });

    const document = buildDocumentFromSnapshot({
      shareToken: token,
      documentId: payload.document_id ?? payload.documentId,
      title,
      fileType,
      fileName,
      fileSize,
      r2ObjectKey,
      storagePath,
      allowDownload,
      downloadUrl,
      snapshotPayload: snapshot,
    });

    return {
      status: 'ready',
      shareToken: token,
      title: document.title,
      fileType: document.type,
      fileName,
      fileSize,
      allowDownload,
      viewCount,
      r2ObjectKey,
      storagePath,
      downloadUrl,
      document,
    };
  }

  return {
    status: 'error',
    shareToken: token,
    errorMessage:
      payload.error ??
      payload.message ??
      'Không thể tải thông tin liên kết chia sẻ.',
  };
}

/**
 * Retrieves an existing share configuration for a document from Supabase, local storage,
 * or the document metadata fallback.
 */
export async function getExistingShareForDocument(
  documentId: string,
  fallbackDoc?: DocumentItem
): Promise<ShareConfigRecord | null> {
  const authCtx = await getCurrentAuthContext();
  if (authCtx && supabase) {
    try {
      const { data, error } = await supabase
        .from(SHARES_TABLE)
        .select(
          'share_token, document_id, user_id, title, file_type, file_name, file_size, r2_object_key, storage_path, access_level, allow_download, view_count, revoked_at, updated_at'
        )
        .eq('user_id', authCtx.userId)
        .eq('document_id', documentId)
        .maybeSingle();

      if (!error && data) {
        return mapSupabaseRowToShareConfig(data as SupabaseShareRow);
      }
    } catch {
      // Fall back to local store
    }
  }

  const localRecord = findLocalShareByDocumentId(documentId);
  if (localRecord) {
    return {
      shareToken: localRecord.shareToken,
      documentId: localRecord.documentId,
      accessLevel: localRecord.accessLevel,
      allowDownload: localRecord.allowDownload,
      r2ObjectKey: localRecord.r2ObjectKey,
      storagePath: localRecord.storagePath,
      viewCount: localRecord.viewCount,
      revokedAt: localRecord.revokedAt,
      updatedAt: localRecord.updatedAt,
    };
  }

  if (fallbackDoc?.shareToken) {
    return {
      shareToken: fallbackDoc.shareToken,
      documentId: fallbackDoc.id,
      accessLevel:
        fallbackDoc.shareAccess ?? (fallbackDoc.isShared ? 'link' : 'private'),
      allowDownload: fallbackDoc.shareAllowDownload ?? true,
      r2ObjectKey: fallbackDoc.r2ObjectKey,
      storagePath: fallbackDoc.storagePath,
      viewCount: 0,
      revokedAt: null,
      updatedAt: fallbackDoc.lastSavedAt ?? new Date().toISOString(),
    };
  }

  return null;
}

/**
 * Saves or updates share settings (private, public link, or password-protected link)
 * across R2 Worker API, Supabase document_shares, and local fallback store.
 */
export async function saveDocumentShareSettings(params: {
  document: DocumentItem;
  accessLevel: ShareAccessLevel;
  password?: string;
  allowDownload: boolean;
}): Promise<ShareConfigRecord> {
  const existingLocal = findLocalShareByDocumentId(params.document.id);
  const existingRemote = await getExistingShareForDocument(
    params.document.id,
    params.document
  );

  const shareToken =
    existingRemote?.shareToken ||
    existingLocal?.shareToken ||
    params.document.shareToken ||
    generateShareToken();

  const trimmedPassword = params.password?.trim() ?? '';
  let passwordHash: string | undefined;

  if (params.accessLevel === 'password') {
    if (trimmedPassword.length > 0) {
      passwordHash = await hashSharePassword(trimmedPassword);
    } else if (
      existingLocal?.accessLevel === 'password' &&
      existingLocal.passwordHash
    ) {
      passwordHash = existingLocal.passwordHash;
    } else if (existingRemote?.accessLevel !== 'password') {
      throw new Error('Vui lòng nhập mật khẩu bảo vệ cho liên kết chia sẻ.');
    }
  }

  const nowIso = new Date().toISOString();
  const snapshotPayload = buildShareSnapshotPayload(params.document);
  const nextRecord: LocalShareRecord = {
    shareToken,
    documentId: params.document.id,
    title: params.document.title.trim() || 'Tài liệu LITE',
    fileType: params.document.type,
    fileName: params.document.originalFileName,
    fileSize: params.document.fileSize,
    r2ObjectKey: params.document.r2ObjectKey,
    storagePath: params.document.storagePath,
    accessLevel: params.accessLevel,
    passwordHash,
    allowDownload: params.allowDownload,
    snapshotPayload,
    viewCount: existingRemote?.viewCount ?? existingLocal?.viewCount ?? 0,
    revokedAt: null,
    updatedAt: nowIso,
  };

  persistLocalShareRecord(nextRecord);

  const authCtx = await getCurrentAuthContext();
  if (authCtx && isR2ApiConfigured) {
    try {
      const response = await fetch(`${getR2ApiBaseUrl()}/api/shares`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${authCtx.accessToken}`,
        },
        body: JSON.stringify({
          shareToken,
          documentId: params.document.id,
          title: nextRecord.title,
          fileType: nextRecord.fileType,
          fileName: nextRecord.fileName ?? null,
          fileSize: nextRecord.fileSize ?? null,
          r2ObjectKey: nextRecord.r2ObjectKey ?? null,
          storagePath: nextRecord.storagePath ?? null,
          accessLevel: params.accessLevel,
          password: trimmedPassword || undefined,
          passwordHash,
          allowDownload: params.allowDownload,
          snapshotPayload,
        }),
      });

      if (response.ok) {
        const remoteData = (await response
          .json()
          .catch(() => null)) as Partial<ShareConfigRecord> | null;
        return {
          shareToken: remoteData?.shareToken ?? shareToken,
          documentId: params.document.id,
          accessLevel: remoteData?.accessLevel ?? params.accessLevel,
          allowDownload: remoteData?.allowDownload ?? params.allowDownload,
          r2ObjectKey: remoteData?.r2ObjectKey ?? params.document.r2ObjectKey,
          storagePath: remoteData?.storagePath ?? params.document.storagePath,
          viewCount: remoteData?.viewCount ?? nextRecord.viewCount,
          revokedAt: remoteData?.revokedAt ?? null,
          updatedAt: remoteData?.updatedAt ?? nowIso,
        };
      }
    } catch {
      // Fall through to direct Supabase upsert if Worker is unreachable
    }
  }

  if (authCtx && supabase) {
    const upsertRow: Record<string, unknown> = {
      share_token: shareToken,
      document_id: params.document.id,
      user_id: authCtx.userId,
      title: nextRecord.title,
      file_type: nextRecord.fileType,
      file_name: nextRecord.fileName ?? null,
      file_size: nextRecord.fileSize ?? null,
      r2_object_key: nextRecord.r2ObjectKey ?? null,
      storage_path: nextRecord.storagePath ?? null,
      access_level: params.accessLevel,
      allow_download: params.allowDownload,
      snapshot_payload: snapshotPayload,
      revoked_at: null,
      updated_at: nowIso,
    };

    if (params.accessLevel !== 'password') {
      upsertRow.password_hash = null;
    } else if (passwordHash) {
      upsertRow.password_hash = passwordHash;
    }

    const { error } = await supabase
      .from(SHARES_TABLE)
      .upsert(upsertRow, { onConflict: 'user_id,document_id' });

    if (error) {
      throw new Error(`Không thể lưu cài đặt chia sẻ: ${error.message}`);
    }
  }

  return {
    shareToken: nextRecord.shareToken,
    documentId: nextRecord.documentId,
    accessLevel: nextRecord.accessLevel,
    allowDownload: nextRecord.allowDownload,
    r2ObjectKey: nextRecord.r2ObjectKey,
    storagePath: nextRecord.storagePath,
    viewCount: nextRecord.viewCount,
    revokedAt: null,
    updatedAt: nowIso,
  };
}

/**
 * Revokes an existing share link immediately across R2 API, Supabase, and local store.
 */
export async function revokeDocumentShare(
  shareToken: string,
  documentId: string
): Promise<void> {
  const nowIso = new Date().toISOString();
  const localByToken = findLocalShareByToken(shareToken);
  const localByDoc = localByToken ?? findLocalShareByDocumentId(documentId);

  if (localByDoc) {
    persistLocalShareRecord({
      ...localByDoc,
      accessLevel: 'private',
      revokedAt: nowIso,
      updatedAt: nowIso,
    });
  }

  const authCtx = await getCurrentAuthContext();
  if (authCtx && isR2ApiConfigured) {
    try {
      const response = await fetch(
        `${getR2ApiBaseUrl()}/api/shares/${encodeURIComponent(shareToken)}`,
        {
          method: 'DELETE',
          headers: {
            Authorization: `Bearer ${authCtx.accessToken}`,
          },
        }
      );
      if (response.ok) {
        return;
      }
    } catch {
      // Fall through to direct Supabase update
    }
  }

  if (authCtx && supabase) {
    const { error } = await supabase
      .from(SHARES_TABLE)
      .update({
        revoked_at: nowIso,
        access_level: 'private',
        updated_at: nowIso,
      })
      .eq('user_id', authCtx.userId)
      .eq('document_id', documentId);

    if (error) {
      throw new Error(`Không thể thu hồi liên kết chia sẻ: ${error.message}`);
    }
  }
}

/**
 * Resolves a public share token from R2 API, Supabase RPC, or local preview store.
 */
export async function resolvePublicShareLink(
  shareToken: string
): Promise<SharedDocumentResolution> {
  const cleanToken = shareToken.trim();
  if (!cleanToken) {
    return { status: 'not_found', shareToken };
  }

  if (isR2ApiConfigured) {
    try {
      const response = await fetch(
        `${getR2ApiBaseUrl()}/api/shares/${encodeURIComponent(cleanToken)}`
      );
      const data = (await response.json().catch(() => null)) as ShareRpcPayload | null;

      if (data && data.status) {
        const mapped = await mapRpcPayloadToResolution(cleanToken, data);
        if (mapped.status !== 'not_found') {
          return mapped;
        }
      } else if (response.status === 403) {
        return { status: 'private', shareToken: cleanToken };
      }
    } catch {
      // Fall through to Supabase RPC / local store
    }
  }

  if (isSupabaseConfigured && supabase) {
    try {
      const { data, error } = await supabase.rpc('get_shared_document_meta', {
        p_share_token: cleanToken,
      });
      if (!error && isPlainRecord(data)) {
        const mapped = await mapRpcPayloadToResolution(
          cleanToken,
          data as ShareRpcPayload
        );
        if (mapped.status !== 'not_found') {
          return mapped;
        }
      }
    } catch {
      // Fall through to local store
    }
  }

  const localRecord = findLocalShareByToken(cleanToken);
  if (!localRecord) {
    return { status: 'not_found', shareToken: cleanToken };
  }

  if (localRecord.revokedAt) {
    return {
      status: 'revoked',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
    };
  }

  if (localRecord.accessLevel === 'private') {
    return {
      status: 'private',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
    };
  }

  if (localRecord.accessLevel === 'password') {
    return {
      status: 'password_required',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
      fileName: localRecord.fileName,
      fileSize: localRecord.fileSize,
      allowDownload: localRecord.allowDownload,
      viewCount: localRecord.viewCount,
    };
  }

  const nextViews = localRecord.viewCount + 1;
  persistLocalShareRecord({
    ...localRecord,
    viewCount: nextViews,
  });

  const downloadUrl = await resolveDownloadUrlForShare({
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    shareToken: cleanToken,
    allowDownload: localRecord.allowDownload,
    fileType: localRecord.fileType,
  });

  const document = buildDocumentFromSnapshot({
    shareToken: cleanToken,
    documentId: localRecord.documentId,
    title: localRecord.title,
    fileType: localRecord.fileType,
    fileName: localRecord.fileName,
    fileSize: localRecord.fileSize,
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    allowDownload: localRecord.allowDownload,
    downloadUrl,
    snapshotPayload: localRecord.snapshotPayload,
  });

  return {
    status: 'ready',
    shareToken: cleanToken,
    title: document.title,
    fileType: document.type,
    fileName: localRecord.fileName,
    fileSize: localRecord.fileSize,
    allowDownload: localRecord.allowDownload,
    viewCount: nextViews,
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    downloadUrl,
    document,
  };
}

/**
 * Unlocks a password-protected share link via R2 API, Supabase RPC, or local store.
 */
export async function unlockPasswordProtectedShare(
  shareToken: string,
  password: string
): Promise<SharedDocumentResolution> {
  const cleanToken = shareToken.trim();
  const cleanPassword = password.trim();
  if (!cleanToken) {
    return { status: 'not_found', shareToken };
  }
  if (!cleanPassword) {
    return {
      status: 'password_required',
      shareToken: cleanToken,
      errorMessage: 'Vui lòng nhập mật khẩu để mở tài liệu.',
    };
  }

  const passwordHash = await hashSharePassword(cleanPassword);

  if (isR2ApiConfigured) {
    try {
      const response = await fetch(
        `${getR2ApiBaseUrl()}/api/shares/${encodeURIComponent(cleanToken)}/unlock`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            password: cleanPassword,
            passwordHash,
          }),
        }
      );
      const data = (await response.json().catch(() => null)) as ShareRpcPayload | null;
      if (data && data.status) {
        const mapped = await mapRpcPayloadToResolution(
          cleanToken,
          data,
          passwordHash
        );
        if (mapped.status !== 'not_found') {
          return mapped;
        }
      } else if (response.status === 401) {
        return {
          status: 'password_required',
          shareToken: cleanToken,
          errorMessage: 'Mật khẩu chưa chính xác. Vui lòng thử lại.',
        };
      } else if (response.status === 403) {
        return {
          status: 'private',
          shareToken: cleanToken,
        };
      }
    } catch {
      // Fall through to Supabase RPC / local store
    }
  }

  if (isSupabaseConfigured && supabase) {
    try {
      const { data, error } = await supabase.rpc('unlock_shared_document', {
        p_share_token: cleanToken,
        p_password_hash: passwordHash,
      });
      if (!error && isPlainRecord(data)) {
        const mapped = await mapRpcPayloadToResolution(
          cleanToken,
          data as ShareRpcPayload,
          passwordHash
        );
        if (mapped.status !== 'not_found') {
          return mapped;
        }
      }
    } catch {
      // Fall through to local store
    }
  }

  const localRecord = findLocalShareByToken(cleanToken);
  if (!localRecord) {
    return { status: 'not_found', shareToken: cleanToken };
  }

  if (localRecord.revokedAt) {
    return {
      status: 'revoked',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
    };
  }

  if (localRecord.accessLevel === 'private') {
    return {
      status: 'private',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
    };
  }

  if (
    localRecord.accessLevel === 'password' &&
    (!localRecord.passwordHash || localRecord.passwordHash !== passwordHash)
  ) {
    return {
      status: 'password_required',
      shareToken: cleanToken,
      title: localRecord.title,
      fileType: localRecord.fileType,
      fileName: localRecord.fileName,
      fileSize: localRecord.fileSize,
      allowDownload: localRecord.allowDownload,
      viewCount: localRecord.viewCount,
      errorMessage: 'Mật khẩu chưa chính xác. Vui lòng thử lại.',
    };
  }

  const nextViews = localRecord.viewCount + 1;
  persistLocalShareRecord({
    ...localRecord,
    viewCount: nextViews,
  });

  const downloadUrl = await resolveDownloadUrlForShare({
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    shareToken: cleanToken,
    allowDownload: localRecord.allowDownload,
    fileType: localRecord.fileType,
    passwordHash,
  });

  const document = buildDocumentFromSnapshot({
    shareToken: cleanToken,
    documentId: localRecord.documentId,
    title: localRecord.title,
    fileType: localRecord.fileType,
    fileName: localRecord.fileName,
    fileSize: localRecord.fileSize,
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    allowDownload: localRecord.allowDownload,
    downloadUrl,
    snapshotPayload: localRecord.snapshotPayload,
  });

  return {
    status: 'ready',
    shareToken: cleanToken,
    title: document.title,
    fileType: document.type,
    fileName: localRecord.fileName,
    fileSize: localRecord.fileSize,
    allowDownload: localRecord.allowDownload,
    viewCount: nextViews,
    r2ObjectKey: localRecord.r2ObjectKey,
    storagePath: localRecord.storagePath,
    downloadUrl,
    document,
  };
}

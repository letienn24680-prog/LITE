import { createThumbnailDataUri } from '../data/mockDocuments';
import { isSupabaseConfigured, supabase } from '../lib/supabase';
import {
  formatRelativeVietnameseTime,
  getCloudFileObjectUrl,
} from './cloudSync';
import type {
  DocumentItem,
  DocumentType,
  OutlineSection,
  PdfPageItem,
  SlidePreviewItem,
  SpreadsheetSheet,
} from '../types/document';

export type CollabPermission = 'private' | 'view' | 'edit';

export interface CollaboratorPresence {
  sessionId: string;
  userId?: string;
  displayName: string;
  email?: string;
  color: string;
  cursorLabel?: string;
  cellCoord?: string;
  sheetId?: string;
  slideIndex?: number;
  isTyping?: boolean;
  updatedAt: number;
}

export type DocRealtimePatch = Record<string, any>;

export interface ShareLinkResult {
  shareToken: string;
  shareUrl: string;
  permission: 'view' | 'edit';
  ownerId: string;
  documentId: string;
}

export interface ResolvedCollabDocument {
  status: 'ready' | 'not_found' | 'revoked' | 'private' | 'error';
  shareToken: string;
  permission: 'view' | 'edit';
  ownerId: string;
  documentId: string;
  document?: DocumentItem;
  errorMessage?: string;
}

export interface CollabRoomIdentity {
  sessionId: string;
  userId?: string;
  displayName: string;
  email?: string;
  color?: string;
}

export interface JoinRoomCallbacks {
  onPresenceChange: (collaborators: CollaboratorPresence[]) => void;
  onRemotePatch: (patch: DocRealtimePatch) => void;
  onDatabaseDocumentUpdate?: (updatedDocPartial: Partial<DocumentItem>) => void;
  onStatusChange?: (
    status: 'connecting' | 'connected' | 'disconnected' | 'error'
  ) => void;
}

export interface RealtimeRoomHandle {
  sessionId: string;
  roomName: string;
  selfColor: string;
  updatePresence: (
    patch: Partial<Omit<CollaboratorPresence, 'sessionId' | 'color'>>
  ) => void;
  broadcastPatch: (
    patch: Omit<DocRealtimePatch, 'senderSessionId' | 'sentAt'>
  ) => void;
  schedulePersist: (docSnapshot: DocumentItem) => void;
  leave: () => void;
}

export const COLLAB_COLORS = [
  '#2563EB',
  '#16A34A',
  '#EA580C',
  '#9333EA',
  '#DB2777',
  '#0D9488',
] as const;

const SESSION_STORAGE_KEY = 'lite_collab_session_id_v1';
const LOCAL_COLLAB_SHARES_KEY = 'lite_collab_shares_v1';
const LOCAL_SAVED_DOCS_PREFIX = 'lite_saved_documents_v2_';
const SHARES_TABLE = 'document_shares';
const DOCUMENTS_TABLE = 'documents';
const COLLABORATORS_TABLE = 'document_collaborators';
const PERSIST_DEBOUNCE_MS = 700;
const SELF_PERSIST_ECHO_WINDOW_MS = 1200;

const DEFAULT_ACCENT_BY_TYPE: Record<DocumentType, string> = {
  pptx: '#4F46E5',
  docx: '#2563EB',
  xlsx: '#15803D',
  pdf: '#DB2777',
};

interface LocalCollabShareRecord {
  shareToken: string;
  ownerId: string;
  documentId: string;
  permission: 'view' | 'edit';
  updatedAt: string;
  documentSnapshot?: DocumentItem;
}

type LocalRoomEnvelope =
  | {
      kind: 'presence_update';
      presence: CollaboratorPresence;
    }
  | {
      kind: 'presence_leave';
      sessionId: string;
    }
  | {
      kind: 'presence_ping';
      fromSessionId: string;
    }
  | {
      kind: 'doc_patch';
      msgId: string;
      patch: DocRealtimePatch;
    };

const memoryCollabShareFallback = new Map<string, LocalCollabShareRecord>();
let cachedTabSessionId: string | null = null;

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isValidUuid(value: string | undefined | null): value is string {
  if (!value) {
    return false;
  }
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(
    value.trim()
  );
}

function isValidDocumentType(value: unknown): value is DocumentType {
  return value === 'pptx' || value === 'docx' || value === 'xlsx' || value === 'pdf';
}

function createSecureToken(length = 14): string {
  const alphabet = 'abcdefghjkmnpqrstuvwxyz23456789';
  let result = '';
  if (typeof globalThis.crypto?.getRandomValues === 'function') {
    const bytes = new Uint8Array(length);
    globalThis.crypto.getRandomValues(bytes);
    for (let i = 0; i < length; i += 1) {
      result += alphabet[bytes[i] % alphabet.length];
    }
  } else {
    for (let i = 0; i < length; i += 1) {
      result += alphabet[Math.floor(Math.random() * alphabet.length)];
    }
  }
  return `lite-${result}`;
}

export function getClientSessionId(): string {
  if (cachedTabSessionId) {
    return cachedTabSessionId;
  }

  try {
    const existing = window.sessionStorage.getItem(SESSION_STORAGE_KEY);
    if (existing && existing.trim().length > 0) {
      cachedTabSessionId = existing.trim();
      return cachedTabSessionId;
    }
  } catch {
    // Ignore sessionStorage restrictions
  }

  const generated =
    typeof globalThis.crypto?.randomUUID === 'function'
      ? `sess-${globalThis.crypto.randomUUID()}`
      : `sess-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;

  cachedTabSessionId = generated;
  try {
    window.sessionStorage.setItem(SESSION_STORAGE_KEY, generated);
  } catch {
    // Ignore sessionStorage write restrictions
  }
  return generated;
}

export function pickCollaboratorColor(seed: string): string {
  let hash = 2166136261;
  for (let i = 0; i < seed.length; i += 1) {
    hash ^= seed.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  const index = Math.abs(hash) % COLLAB_COLORS.length;
  return COLLAB_COLORS[index];
}

export const PUBLIC_WEB_ORIGIN =
  (import.meta.env.VITE_PUBLIC_APP_URL ?? '').trim().replace(/\/+$/, '') ||
  'https://lite-git-main-ai-check-thd.vercel.app';

export function resolvePublicWebOrigin(): string {
  if (typeof window === 'undefined') {
    return PUBLIC_WEB_ORIGIN;
  }
  const origin = (window.location.origin ?? '').trim();
  const protocol = (window.location.protocol ?? '').trim().toLowerCase();
  const lowerOrigin = origin.toLowerCase();

  if (
    !origin ||
    origin === 'null' ||
    lowerOrigin.includes('localhost') ||
    lowerOrigin.includes('127.0.0.1') ||
    protocol === 'capacitor:' ||
    protocol === 'file:' ||
    protocol === 'ionic:'
  ) {
    return PUBLIC_WEB_ORIGIN;
  }

  return origin.replace(/\/+$/, '');
}

export function buildCollabShareUrl(
  shareToken: string,
  mode?: 'view' | 'edit'
): string {
  const modeSuffix = mode === 'view' || mode === 'edit' ? `&mode=${mode}` : '';
  return `${resolvePublicWebOrigin()}/?share=${encodeURIComponent(shareToken)}${modeSuffix}`;
}

export function buildAppDeepLinkShareUrl(
  docIdOrToken: string,
  mode: 'view' | 'edit' = 'view'
): string {
  const safeMode = mode === 'edit' ? 'edit' : 'view';
  return `liteapp://open?share=${encodeURIComponent(docIdOrToken)}&mode=${safeMode}`;
}

function sanitizeSlidesForPersist(slides: SlidePreviewItem[]): SlidePreviewItem[] {
  return slides.map((slide) => {
    const hasImages = Boolean(slide.images && slide.images.length > 0);
    const hasBlobBg = Boolean(slide.backgroundImageSrc?.startsWith('blob:'));
    if (!hasImages && !hasBlobBg) {
      return slide;
    }
    const nextSlide: SlidePreviewItem = { ...slide };
    if (hasBlobBg) {
      delete nextSlide.backgroundImageSrc;
    }
    if (slide.images && slide.images.length > 0) {
      const cleanImages = slide.images.filter(
        (img) => typeof img.src === 'string' && !img.src.startsWith('blob:')
      );
      if (cleanImages.length > 0) {
        nextSlide.images = cleanImages;
      } else {
        delete nextSlide.images;
      }
    }
    return nextSlide;
  });
}

function sanitizeDocForLocalSnapshot(doc: DocumentItem): DocumentItem {
  const clean: DocumentItem = {
    ...doc,
    slidesPreview: sanitizeSlidesForPersist(doc.slidesPreview ?? []),
  };
  delete clean.sourceFile;
  delete clean.rawArrayBuffer;
  if (clean.objectUrl?.startsWith('blob:')) {
    delete clean.objectUrl;
  }
  if (clean.fileUrl?.startsWith('blob:')) {
    delete clean.fileUrl;
  }
  return clean;
}

function buildParsedContentPayload(doc: DocumentItem): Record<string, unknown> {
  const payload: Record<string, unknown> = {
    slidesPreview: sanitizeSlidesForPersist(doc.slidesPreview ?? []),
    outline: doc.outline ?? [],
  };
  if (typeof doc.wordHtml === 'string') {
    payload.wordHtml = doc.wordHtml;
  }
  if (
    typeof doc.slideAspectRatio === 'number' &&
    Number.isFinite(doc.slideAspectRatio) &&
    doc.slideAspectRatio > 0
  ) {
    payload.slideAspectRatio = doc.slideAspectRatio;
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

function buildCollabSnapshotPayload(doc?: DocumentItem): Record<string, unknown> {
  if (!doc) {
    return {};
  }
  return {
    id: doc.id,
    title: doc.title,
    fileType: doc.type,
    fileName: doc.originalFileName ?? null,
    fileSize: doc.fileSize ?? null,
    folder: doc.folder ?? null,
    pageCount: typeof doc.pageCount === 'number' ? doc.pageCount : null,
    contentPreview: doc.contentPreview ?? '',
    ...buildParsedContentPayload(doc),
  };
}

function readLocalCollabShareMap(): Record<string, LocalCollabShareRecord> {
  const merged: Record<string, LocalCollabShareRecord> = {};
  memoryCollabShareFallback.forEach((value, key) => {
    merged[key] = value;
  });

  try {
    const raw = window.localStorage.getItem(LOCAL_COLLAB_SHARES_KEY);
    if (!raw) {
      return merged;
    }
    const parsed: unknown = JSON.parse(raw);
    if (!isPlainRecord(parsed)) {
      return merged;
    }
    for (const [token, entry] of Object.entries(parsed)) {
      if (
        isPlainRecord(entry) &&
        typeof entry.shareToken === 'string' &&
        typeof entry.documentId === 'string'
      ) {
        merged[token] = {
          shareToken: entry.shareToken,
          ownerId: typeof entry.ownerId === 'string' ? entry.ownerId : 'local-owner',
          documentId: entry.documentId,
          permission: entry.permission === 'edit' ? 'edit' : 'view',
          updatedAt:
            typeof entry.updatedAt === 'string'
              ? entry.updatedAt
              : new Date().toISOString(),
          documentSnapshot: isPlainRecord(entry.documentSnapshot)
            ? (entry.documentSnapshot as unknown as DocumentItem)
            : undefined,
        };
      }
    }
  } catch {
    // Ignore localStorage read errors
  }

  return merged;
}

function writeLocalCollabShareMap(
  map: Record<string, LocalCollabShareRecord>
): void {
  memoryCollabShareFallback.clear();
  for (const [token, record] of Object.entries(map)) {
    memoryCollabShareFallback.set(token, record);
  }
  try {
    window.localStorage.setItem(LOCAL_COLLAB_SHARES_KEY, JSON.stringify(map));
  } catch {
    // Ignore storage quota errors
  }
}

function findLocalShareByDoc(
  ownerId: string,
  documentId: string
): LocalCollabShareRecord | null {
  const map = readLocalCollabShareMap();
  const entries = Object.values(map);
  return (
    entries.find(
      (entry) => entry.documentId === documentId && entry.ownerId === ownerId
    ) ??
    entries.find((entry) => entry.documentId === documentId) ??
    null
  );
}

function upsertLocalCollabShare(record: LocalCollabShareRecord): void {
  const map = readLocalCollabShareMap();
  for (const [token, existing] of Object.entries(map)) {
    if (
      existing.documentId === record.documentId &&
      existing.ownerId === record.ownerId &&
      token !== record.shareToken
    ) {
      delete map[token];
    }
  }
  map[record.shareToken] = record;
  writeLocalCollabShareMap(map);
}

function deleteLocalCollabShare(ownerId: string, documentId: string): void {
  const map = readLocalCollabShareMap();
  let changed = false;
  for (const [token, entry] of Object.entries(map)) {
    if (
      entry.documentId === documentId &&
      (entry.ownerId === ownerId || !ownerId || ownerId === 'local-owner')
    ) {
      delete map[token];
      changed = true;
    }
  }
  if (changed) {
    writeLocalCollabShareMap(map);
  }
}

function buildDocumentFromPayload(params: {
  documentId: string;
  ownerId: string;
  shareToken: string;
  permission: 'view' | 'edit';
  title?: string | null;
  fileType?: unknown;
  fileName?: string | null;
  fileSize?: string | null;
  folder?: string | null;
  pageCount?: number | null;
  contentPreview?: string | null;
  storagePath?: string | null;
  r2ObjectKey?: string | null;
  updatedAt?: string | null;
  parsedContent?: unknown;
  snapshotPayload?: unknown;
}): DocumentItem {
  const parsedObj = isPlainRecord(params.parsedContent)
    ? params.parsedContent
    : {};
  const snapshotObj = isPlainRecord(params.snapshotPayload)
    ? params.snapshotPayload
    : {};
  const source = Object.keys(parsedObj).length > 0 ? parsedObj : snapshotObj;

  const docType: DocumentType = isValidDocumentType(params.fileType)
    ? params.fileType
    : isValidDocumentType(snapshotObj.fileType)
      ? snapshotObj.fileType
      : 'docx';

  const title =
    (typeof params.title === 'string' && params.title.trim()) ||
    (typeof snapshotObj.title === 'string' && snapshotObj.title.trim()) ||
    'Tài liệu được chia sẻ';

  const slidesPreview = Array.isArray(source.slidesPreview)
    ? sanitizeSlidesForPersist(source.slidesPreview as SlidePreviewItem[])
    : [];
  const outline = Array.isArray(source.outline)
    ? (source.outline as OutlineSection[])
    : [];

  const accentFromSlide = slidesPreview[0]?.accentColor;
  const accentColor =
    docType === 'pptx' &&
    typeof accentFromSlide === 'string' &&
    /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(accentFromSlide.trim())
      ? accentFromSlide.trim()
      : DEFAULT_ACCENT_BY_TYPE[docType];

  const slideAspectRatio =
    typeof source.slideAspectRatio === 'number' &&
    Number.isFinite(source.slideAspectRatio) &&
    source.slideAspectRatio > 0
      ? source.slideAspectRatio
      : undefined;

  const doc: DocumentItem = {
    id: params.documentId,
    title,
    type: docType,
    updatedAt: params.updatedAt
      ? formatRelativeVietnameseTime(params.updatedAt)
      : 'Vừa mở',
    isFavorite: false,
    isShared: true,
    thumbnail: createThumbnailDataUri(docType, title, accentColor),
    slidesPreview,
    outline,
    contentPreview:
      (typeof params.contentPreview === 'string' && params.contentPreview) ||
      (typeof snapshotObj.contentPreview === 'string'
        ? snapshotObj.contentPreview
        : ''),
    pageCount:
      typeof params.pageCount === 'number'
        ? params.pageCount
        : typeof snapshotObj.pageCount === 'number'
          ? snapshotObj.pageCount
          : undefined,
    fileSize:
      params.fileSize ??
      (typeof snapshotObj.fileSize === 'string'
        ? snapshotObj.fileSize
        : undefined),
    folder:
      params.folder ??
      (typeof snapshotObj.folder === 'string' ? snapshotObj.folder : undefined),
    originalFileName:
      params.fileName ??
      (typeof snapshotObj.fileName === 'string'
        ? snapshotObj.fileName
        : undefined),
    storagePath: params.storagePath ?? undefined,
    r2ObjectKey: params.r2ObjectKey ?? undefined,
    shareToken: params.shareToken,
    shareAccess: 'link',
    collabPermission: params.permission,
    sharedOwnerId: params.ownerId,
    userId: params.ownerId,
    lastSavedAt: params.updatedAt ?? undefined,
    isUploaded: true,
  };

  if (slideAspectRatio !== undefined) {
    doc.slideAspectRatio = slideAspectRatio;
  }
  if (typeof source.wordHtml === 'string') {
    doc.wordHtml = source.wordHtml;
  }
  if (Array.isArray(source.sheets)) {
    doc.sheets = source.sheets as SpreadsheetSheet[];
  }
  if (Array.isArray(source.pdfPages)) {
    doc.pdfPages = source.pdfPages as PdfPageItem[];
  }
  if (isPlainRecord(source.wordParagraphs)) {
    const paragraphs: Record<string, string[]> = {};
    for (const [k, v] of Object.entries(source.wordParagraphs)) {
      if (Array.isArray(v)) {
        paragraphs[k] = v.filter((item): item is string => typeof item === 'string');
      }
    }
    doc.wordParagraphs = paragraphs;
  }

  return doc;
}

export async function getShareLinkInfo(
  ownerId: string,
  documentId: string
): Promise<{
  permission: CollabPermission;
  shareToken?: string;
  shareUrl?: string;
}> {
  if (isSupabaseConfigured && supabase && isValidUuid(ownerId)) {
    try {
      const { data, error } = await supabase
        .from(SHARES_TABLE)
        .select('share_token, access_level, permission, revoked_at')
        .eq('user_id', ownerId)
        .eq('document_id', documentId)
        .maybeSingle();

      if (!error && data) {
        const row = data as {
          share_token?: string;
          access_level?: string;
          permission?: string;
          revoked_at?: string | null;
        };
        if (
          row.share_token &&
          !row.revoked_at &&
          row.access_level !== 'private'
        ) {
          const perm: 'view' | 'edit' =
            row.permission === 'edit' ? 'edit' : 'view';
          return {
            permission: perm,
            shareToken: row.share_token,
            shareUrl: buildCollabShareUrl(row.share_token, perm),
          };
        }
        return { permission: 'private' };
      }
    } catch {
      // Fallback to local mirror below if Supabase query fails
    }
  }

  const local = findLocalShareByDoc(ownerId, documentId);
  if (local) {
    return {
      permission: local.permission,
      shareToken: local.shareToken,
      shareUrl: buildCollabShareUrl(local.shareToken, local.permission),
    };
  }

  return { permission: 'private' };
}

export async function createOrUpdateShareLink(
  ownerId: string,
  documentId: string,
  permission: 'view' | 'edit',
  docSnapshot?: DocumentItem
): Promise<ShareLinkResult> {
  const existingLocal = findLocalShareByDoc(ownerId, documentId);
  let shareToken = docSnapshot?.shareToken || existingLocal?.shareToken || '';
  const nowIso = new Date().toISOString();

  if (isSupabaseConfigured && supabase && isValidUuid(ownerId)) {
    try {
      const { data: existingRow } = await supabase
        .from(SHARES_TABLE)
        .select('share_token')
        .eq('user_id', ownerId)
        .eq('document_id', documentId)
        .maybeSingle();

      if (
        existingRow &&
        typeof (existingRow as { share_token?: unknown }).share_token ===
          'string'
      ) {
        shareToken = (existingRow as { share_token: string }).share_token;
      }

      if (!shareToken) {
        shareToken = createSecureToken(14);
      }

      const snapshotPayload = buildCollabSnapshotPayload(docSnapshot);
      const { error } = await supabase.from(SHARES_TABLE).upsert(
        {
          share_token: shareToken,
          document_id: documentId,
          user_id: ownerId,
          title: docSnapshot?.title || 'Tài liệu LITE',
          file_type: docSnapshot?.type || 'docx',
          file_name: docSnapshot?.originalFileName ?? null,
          file_size: docSnapshot?.fileSize ?? null,
          r2_object_key: docSnapshot?.r2ObjectKey ?? null,
          storage_path: docSnapshot?.storagePath ?? null,
          access_level: 'link',
          permission,
          allow_download: docSnapshot?.shareAllowDownload ?? true,
          snapshot_payload: snapshotPayload,
          revoked_at: null,
          updated_at: nowIso,
        },
        { onConflict: 'user_id,document_id' }
      );

      if (!error) {
        void supabase
          .from(DOCUMENTS_TABLE)
          .update({ is_shared: true })
          .eq('user_id', ownerId)
          .eq('id', documentId);
      }
    } catch {
      if (!shareToken) {
        shareToken = createSecureToken(14);
      }
    }
  }

  if (!shareToken) {
    shareToken = createSecureToken(14);
  }

  upsertLocalCollabShare({
    shareToken,
    ownerId,
    documentId,
    permission,
    updatedAt: nowIso,
    documentSnapshot: docSnapshot
      ? sanitizeDocForLocalSnapshot({
          ...docSnapshot,
          shareToken,
          isShared: true,
        })
      : undefined,
  });

  return {
    shareToken,
    shareUrl: buildCollabShareUrl(shareToken, permission),
    permission,
    ownerId,
    documentId,
  };
}

export async function revokeShareLink(
  ownerId: string,
  documentId: string
): Promise<void> {
  deleteLocalCollabShare(ownerId, documentId);

  if (isSupabaseConfigured && supabase && isValidUuid(ownerId)) {
    try {
      await supabase
        .from(SHARES_TABLE)
        .delete()
        .eq('user_id', ownerId)
        .eq('document_id', documentId);

      await supabase
        .from(DOCUMENTS_TABLE)
        .update({ is_shared: false })
        .eq('user_id', ownerId)
        .eq('id', documentId);
    } catch {
      // Ignore network/DB errors after local revocation
    }
  }
}

function readLocalSavedDocuments(currentUserId?: string): DocumentItem[] {
  try {
    const scope = currentUserId?.trim() || 'guest';
    const raw = window.localStorage.getItem(`${LOCAL_SAVED_DOCS_PREFIX}${scope}`);
    if (!raw) {
      return [];
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.filter(
      (entry): entry is DocumentItem =>
        isPlainRecord(entry) && typeof entry.id === 'string'
    );
  } catch {
    return [];
  }
}

export async function resolveSharedDocument(
  shareToken: string,
  currentUserId?: string,
  modeOrOptions?: 'view' | 'edit' | { mode?: 'view' | 'edit' }
): Promise<ResolvedCollabDocument> {
  const cleanToken = shareToken.trim();
  const requestedMode: 'view' | 'edit' | undefined =
    typeof modeOrOptions === 'string'
      ? modeOrOptions === 'edit' || modeOrOptions === 'view'
        ? modeOrOptions
        : undefined
      : modeOrOptions?.mode === 'edit' || modeOrOptions?.mode === 'view'
        ? modeOrOptions.mode
        : undefined;

  const applyRequestedMode = (serverPerm: 'view' | 'edit'): 'view' | 'edit' =>
    requestedMode === 'view' ? 'view' : serverPerm;

  if (!cleanToken) {
    return {
      status: 'not_found',
      shareToken: '',
      permission: 'view',
      ownerId: '',
      documentId: '',
      errorMessage: 'Liên kết chia sẻ không hợp lệ.',
    };
  }

  if (isSupabaseConfigured && supabase) {
    try {
      const resolveViaToken = async (tokenToResolve: string) =>
        await supabase!.rpc('resolve_collab_share', {
          p_share_token: tokenToResolve,
        });

      let activeToken = cleanToken;
      let { data, error } = await resolveViaToken(activeToken);

      if (
        !error &&
        isPlainRecord(data) &&
        data.status === 'not_found'
      ) {
        // Fallback: caller may have passed ?share=<docId> instead of <shareToken>
        const { data: byDocShare } = await supabase
          .from(SHARES_TABLE)
          .select('share_token')
          .eq('document_id', cleanToken)
          .is('revoked_at', null)
          .neq('access_level', 'private')
          .limit(1)
          .maybeSingle();

        if (
          byDocShare &&
          typeof (byDocShare as { share_token?: unknown }).share_token ===
            'string'
        ) {
          activeToken = (byDocShare as { share_token: string }).share_token;
          const retry = await resolveViaToken(activeToken);
          data = retry.data;
          error = retry.error;
        }
      }

      if (!error && isPlainRecord(data)) {
        const status = typeof data.status === 'string' ? data.status : 'error';
        const ownerId =
          typeof data.owner_id === 'string' ? data.owner_id : 'cloud-owner';
        const documentId =
          typeof data.document_id === 'string' ? data.document_id : '';
        const serverPermission: 'view' | 'edit' =
          data.permission === 'edit' ? 'edit' : 'view';
        const effectivePermission = applyRequestedMode(serverPermission);

        if (status === 'revoked') {
          return {
            status: 'revoked',
            shareToken: activeToken,
            permission: effectivePermission,
            ownerId,
            documentId,
            errorMessage: 'Liên kết chia sẻ này đã bị chủ sở hữu thu hồi.',
          };
        }

        if (status === 'private') {
          return {
            status: 'private',
            shareToken: activeToken,
            permission: effectivePermission,
            ownerId,
            documentId,
            errorMessage: 'Tài liệu hiện đang ở chế độ Hạn chế (Riêng tư).',
          };
        }

        if (status === 'ready' && documentId) {
          const doc = buildDocumentFromPayload({
            documentId,
            ownerId,
            shareToken: activeToken,
            permission: effectivePermission,
            title: typeof data.title === 'string' ? data.title : null,
            fileType: data.file_type,
            fileName:
              typeof data.file_name === 'string' ? data.file_name : null,
            fileSize:
              typeof data.file_size === 'string' ? data.file_size : null,
            folder: typeof data.folder === 'string' ? data.folder : null,
            pageCount:
              typeof data.page_count === 'number' ? data.page_count : null,
            contentPreview:
              typeof data.content_preview === 'string'
                ? data.content_preview
                : null,
            storagePath:
              typeof data.storage_path === 'string' ? data.storage_path : null,
            r2ObjectKey:
              typeof data.r2_object_key === 'string'
                ? data.r2_object_key
                : null,
            updatedAt:
              typeof data.updated_at === 'string' ? data.updated_at : null,
            parsedContent: data.parsed_content,
            snapshotPayload: data.snapshot_payload,
          });

          if (doc.type === 'pdf' && doc.storagePath && !doc.objectUrl) {
            try {
              const { objectUrl } = await getCloudFileObjectUrl(doc.storagePath);
              doc.objectUrl = objectUrl;
            } catch {
              // Guests without storage RLS fallback cleanly to snapshot / pages
            }
          }

          if (
            currentUserId &&
            isValidUuid(currentUserId) &&
            isValidUuid(ownerId) &&
            currentUserId !== ownerId
          ) {
            try {
              await supabase.from(COLLABORATORS_TABLE).upsert(
                {
                  owner_id: ownerId,
                  document_id: documentId,
                  collaborator_user_id: currentUserId,
                  share_token: activeToken,
                  permission: serverPermission,
                  last_opened_at: new Date().toISOString(),
                },
                { onConflict: 'collaborator_user_id,owner_id,document_id' }
              );
            } catch {
              // Ignore if table not created yet
            }
          }

          return {
            status: 'ready',
            shareToken: activeToken,
            permission: effectivePermission,
            ownerId,
            documentId,
            document: doc,
          };
        }

        if (status === 'not_found') {
          // Direct public.documents fallback when shared by docId
          const { data: sharedDocRow } = await supabase
            .from(DOCUMENTS_TABLE)
            .select('*')
            .eq('id', cleanToken)
            .eq('is_shared', true)
            .limit(1)
            .maybeSingle();

          if (sharedDocRow && isPlainRecord(sharedDocRow)) {
            const ownerIdFromRow =
              typeof sharedDocRow.user_id === 'string'
                ? sharedDocRow.user_id
                : 'cloud-owner';
            const fallbackPerm: 'view' | 'edit' =
              requestedMode === 'edit' ? 'edit' : 'view';
            const doc = buildDocumentFromPayload({
              documentId: cleanToken,
              ownerId: ownerIdFromRow,
              shareToken: cleanToken,
              permission: fallbackPerm,
              title:
                typeof sharedDocRow.title === 'string'
                  ? sharedDocRow.title
                  : null,
              fileType: sharedDocRow.file_type,
              fileName:
                typeof sharedDocRow.file_name === 'string'
                  ? sharedDocRow.file_name
                  : null,
              fileSize:
                typeof sharedDocRow.file_size === 'string'
                  ? sharedDocRow.file_size
                  : null,
              folder:
                typeof sharedDocRow.folder === 'string'
                  ? sharedDocRow.folder
                  : null,
              pageCount:
                typeof sharedDocRow.page_count === 'number'
                  ? sharedDocRow.page_count
                  : null,
              contentPreview:
                typeof sharedDocRow.content_preview === 'string'
                  ? sharedDocRow.content_preview
                  : null,
              storagePath:
                typeof sharedDocRow.storage_path === 'string'
                  ? sharedDocRow.storage_path
                  : null,
              r2ObjectKey:
                typeof sharedDocRow.r2_object_key === 'string'
                  ? sharedDocRow.r2_object_key
                  : null,
              updatedAt:
                typeof sharedDocRow.updated_at === 'string'
                  ? sharedDocRow.updated_at
                  : null,
              parsedContent: sharedDocRow.parsed_content,
            });
            return {
              status: 'ready',
              shareToken: cleanToken,
              permission: fallbackPerm,
              ownerId: ownerIdFromRow,
              documentId: cleanToken,
              document: doc,
            };
          }
        }
      } else if (error) {
        // Fallback to get_shared_document_meta RPC if resolve_collab_share hasn't been deployed
        const { data: legacyData, error: legacyErr } = await supabase.rpc(
          'get_shared_document_meta',
          { p_share_token: cleanToken }
        );
        if (!legacyErr && isPlainRecord(legacyData)) {
          const legacyStatus =
            typeof legacyData.status === 'string' ? legacyData.status : '';
          if (legacyStatus === 'revoked' || legacyStatus === 'private') {
            return {
              status: legacyStatus,
              shareToken: cleanToken,
              permission: 'view',
              ownerId: '',
              documentId:
                typeof legacyData.document_id === 'string'
                  ? legacyData.document_id
                  : '',
              errorMessage:
                legacyStatus === 'revoked'
                  ? 'Liên kết chia sẻ đã bị thu hồi.'
                  : 'Tài liệu đang ở chế độ riêng tư.',
            };
          }
          if (legacyStatus === 'ok') {
            const docId =
              typeof legacyData.document_id === 'string'
                ? legacyData.document_id
                : cleanToken;
            const effectivePerm = applyRequestedMode('view');
            const doc = buildDocumentFromPayload({
              documentId: docId,
              ownerId: 'cloud-owner',
              shareToken: cleanToken,
              permission: effectivePerm,
              title:
                typeof legacyData.title === 'string' ? legacyData.title : null,
              fileType: legacyData.file_type,
              fileName:
                typeof legacyData.file_name === 'string'
                  ? legacyData.file_name
                  : null,
              fileSize:
                typeof legacyData.file_size === 'string'
                  ? legacyData.file_size
                  : null,
              storagePath:
                typeof legacyData.storage_path === 'string'
                  ? legacyData.storage_path
                  : null,
              r2ObjectKey:
                typeof legacyData.r2_object_key === 'string'
                  ? legacyData.r2_object_key
                  : null,
              snapshotPayload: legacyData.snapshot_payload,
            });
            return {
              status: 'ready',
              shareToken: cleanToken,
              permission: effectivePerm,
              ownerId: 'cloud-owner',
              documentId: docId,
              document: doc,
            };
          }
        }
      }
    } catch {
      // Continue to local fallback check
    }
  }

  const localMap = readLocalCollabShareMap();
  const localRecord =
    localMap[cleanToken] ??
    Object.values(localMap).find((entry) => entry.documentId === cleanToken);

  if (localRecord?.documentSnapshot) {
    const effectivePermission = applyRequestedMode(localRecord.permission);
    const hydratedLocalDoc: DocumentItem = {
      ...localRecord.documentSnapshot,
      isShared: true,
      shareToken: localRecord.shareToken,
      collabPermission: effectivePermission,
      sharedOwnerId: localRecord.ownerId,
    };
    return {
      status: 'ready',
      shareToken: localRecord.shareToken,
      permission: effectivePermission,
      ownerId: localRecord.ownerId,
      documentId: localRecord.documentId,
      document: hydratedLocalDoc,
    };
  }

  // Fallback by documentId in local saved documents when shareAccess='link' or isShared=true
  const localSaved = readLocalSavedDocuments(currentUserId).find(
    (doc) =>
      doc.id === cleanToken && (doc.shareAccess === 'link' || Boolean(doc.isShared))
  );
  if (localSaved) {
    const fallbackPermission: 'view' | 'edit' =
      requestedMode === 'view'
        ? 'view'
        : requestedMode === 'edit'
          ? 'edit'
          : localSaved.collabPermission === 'edit'
            ? 'edit'
            : 'view';
    const ownerId =
      localSaved.sharedOwnerId || localSaved.userId || 'local-owner';
    const tokenForDoc = localSaved.shareToken || localSaved.id;
    return {
      status: 'ready',
      shareToken: tokenForDoc,
      permission: fallbackPermission,
      ownerId,
      documentId: localSaved.id,
      document: {
        ...localSaved,
        isShared: true,
        shareToken: tokenForDoc,
        collabPermission: fallbackPermission,
        sharedOwnerId: ownerId,
      },
    };
  }

  return {
    status: 'not_found',
    shareToken: cleanToken,
    permission: 'view',
    ownerId: '',
    documentId: '',
    errorMessage:
      'Không tìm thấy tài liệu hoặc liên kết chia sẻ đã hết hiệu lực.',
  };
}

export async function loadSharedWithMeDocuments(
  currentUserId: string
): Promise<DocumentItem[]> {
  if (!isSupabaseConfigured || !supabase || !isValidUuid(currentUserId)) {
    return [];
  }

  try {
    const { data, error } = await supabase.rpc('get_shared_with_me_documents');
    if (!error && Array.isArray(data)) {
      const results: DocumentItem[] = [];
      for (const entry of data) {
        if (!isPlainRecord(entry)) {
          continue;
        }
        const documentId =
          typeof entry.document_id === 'string' ? entry.document_id : '';
        const ownerId = typeof entry.owner_id === 'string' ? entry.owner_id : '';
        const shareToken =
          typeof entry.share_token === 'string' ? entry.share_token : '';
        if (!documentId || !ownerId || !shareToken) {
          continue;
        }
        const permission: 'view' | 'edit' =
          entry.permission === 'edit' ? 'edit' : 'view';

        results.push(
          buildDocumentFromPayload({
            documentId,
            ownerId,
            shareToken,
            permission,
            title: typeof entry.title === 'string' ? entry.title : null,
            fileType: entry.file_type,
            fileName:
              typeof entry.file_name === 'string' ? entry.file_name : null,
            fileSize:
              typeof entry.file_size === 'string' ? entry.file_size : null,
            folder: typeof entry.folder === 'string' ? entry.folder : null,
            pageCount:
              typeof entry.page_count === 'number' ? entry.page_count : null,
            contentPreview:
              typeof entry.content_preview === 'string'
                ? entry.content_preview
                : null,
            storagePath:
              typeof entry.storage_path === 'string' ? entry.storage_path : null,
            r2ObjectKey:
              typeof entry.r2_object_key === 'string'
                ? entry.r2_object_key
                : null,
            updatedAt:
              typeof entry.updated_at === 'string' ? entry.updated_at : null,
            parsedContent: entry.parsed_content,
            snapshotPayload: entry.snapshot_payload,
          })
        );
      }
      return results;
    }

    // Fallback direct table join when RPC isn't deployed
    const { data: collabRows, error: collabError } = await supabase
      .from(COLLABORATORS_TABLE)
      .select('owner_id, document_id, share_token, permission, last_opened_at')
      .eq('collaborator_user_id', currentUserId)
      .order('last_opened_at', { ascending: false });

    if (collabError || !Array.isArray(collabRows) || collabRows.length === 0) {
      return [];
    }

    const docs: DocumentItem[] = [];
    for (const rawRow of collabRows) {
      if (!isPlainRecord(rawRow)) {
        continue;
      }
      const token =
        typeof rawRow.share_token === 'string' ? rawRow.share_token : '';
      if (!token) {
        continue;
      }
      const resolved = await resolveSharedDocument(token, currentUserId);
      if (resolved.status === 'ready' && resolved.document) {
        docs.push(resolved.document);
      }
    }
    return docs;
  } catch {
    return [];
  }
}

function parseIncomingPatch(raw: unknown): DocRealtimePatch | null {
  if (!isPlainRecord(raw)) {
    return null;
  }
  const kind = raw.kind;
  const senderSessionId =
    typeof raw.senderSessionId === 'string' ? raw.senderSessionId : '';
  const sentAt =
    typeof raw.sentAt === 'number' && Number.isFinite(raw.sentAt)
      ? raw.sentAt
      : Date.now();

  if (!senderSessionId) {
    return null;
  }

  if (
    kind === 'excel_cell_patch' &&
    typeof raw.sheetId === 'string' &&
    typeof raw.rowIndex === 'number' &&
    typeof raw.colIndex === 'number' &&
    (typeof raw.value === 'string' || typeof raw.value === 'number')
  ) {
    return {
      kind,
      senderSessionId,
      sentAt,
      sheetId: raw.sheetId,
      rowIndex: raw.rowIndex,
      colIndex: raw.colIndex,
      value: raw.value,
      sheets: Array.isArray(raw.sheets)
        ? (raw.sheets as SpreadsheetSheet[])
        : undefined,
    };
  }

  if (kind === 'word_html_patch' && typeof raw.wordHtml === 'string') {
    return {
      kind,
      senderSessionId,
      sentAt,
      wordHtml: raw.wordHtml,
      contentPreview:
        typeof raw.contentPreview === 'string' ? raw.contentPreview : undefined,
    };
  }

  if (kind === 'slide_patch' && Array.isArray(raw.slidesPreview)) {
    return {
      kind,
      senderSessionId,
      sentAt,
      slidesPreview: raw.slidesPreview as SlidePreviewItem[],
      slideAspectRatio:
        typeof raw.slideAspectRatio === 'number'
          ? raw.slideAspectRatio
          : undefined,
    };
  }

  if (kind === 'title_patch' && typeof raw.title === 'string') {
    return {
      kind,
      senderSessionId,
      sentAt,
      title: raw.title,
    };
  }

  return null;
}

function parsePresenceRecord(raw: unknown): CollaboratorPresence | null {
  if (!isPlainRecord(raw)) {
    return null;
  }
  if (
    typeof raw.sessionId !== 'string' ||
    typeof raw.displayName !== 'string' ||
    typeof raw.color !== 'string'
  ) {
    return null;
  }
  return {
    sessionId: raw.sessionId,
    userId: typeof raw.userId === 'string' ? raw.userId : undefined,
    displayName: raw.displayName,
    email: typeof raw.email === 'string' ? raw.email : undefined,
    color: raw.color,
    cursorLabel:
      typeof raw.cursorLabel === 'string' ? raw.cursorLabel : undefined,
    cellCoord: typeof raw.cellCoord === 'string' ? raw.cellCoord : undefined,
    sheetId: typeof raw.sheetId === 'string' ? raw.sheetId : undefined,
    slideIndex:
      typeof raw.slideIndex === 'number' && Number.isFinite(raw.slideIndex)
        ? raw.slideIndex
        : undefined,
    isTyping: typeof raw.isTyping === 'boolean' ? raw.isTyping : undefined,
    updatedAt:
      typeof raw.updatedAt === 'number' && Number.isFinite(raw.updatedAt)
        ? raw.updatedAt
        : Date.now(),
  };
}

function openBroadcastChannelSafe(name: string): BroadcastChannel | null {
  try {
    if (typeof BroadcastChannel !== 'undefined') {
      return new BroadcastChannel(name);
    }
  } catch {
    // Ignore when BroadcastChannel is unsupported
  }
  return null;
}

export function joinDocumentRealtimeRoom(params: {
  ownerId: string;
  documentId: string;
  shareToken?: string;
  canEdit: boolean;
  identity: CollabRoomIdentity;
  callbacks: JoinRoomCallbacks;
}): RealtimeRoomHandle {
  const { ownerId, documentId, shareToken, canEdit, identity, callbacks } =
    params;
  const roomName = `lite-room:${ownerId}:${documentId}`;
  const selfColor =
    identity.color || pickCollaboratorColor(identity.sessionId);

  let destroyed = false;
  let supabaseSubscribed = false;
  let lastLocalPersistAt = 0;
  let persistTimerId: number | null = null;
  let pendingSnapshot: DocumentItem | null = null;

  const seenMsgIds = new Set<string>();
  const seenQueue: string[] = [];
  const localPeers = new Map<string, CollaboratorPresence>();
  const supabasePeers = new Map<string, CollaboratorPresence>();

  let currentPresence: CollaboratorPresence = {
    sessionId: identity.sessionId,
    userId: identity.userId,
    displayName: identity.displayName,
    email: identity.email,
    color: selfColor,
    cursorLabel: 'Đang xem tài liệu',
    isTyping: false,
    updatedAt: Date.now(),
  };

  const markAndCheckDuplicate = (msgId: string): boolean => {
    if (!msgId) {
      return false;
    }
    if (seenMsgIds.has(msgId)) {
      return true;
    }
    seenMsgIds.add(msgId);
    seenQueue.push(msgId);
    if (seenQueue.length > 240) {
      const oldest = seenQueue.shift();
      if (oldest) {
        seenMsgIds.delete(oldest);
      }
    }
    return false;
  };

  const emitMergedPresence = () => {
    if (destroyed) {
      return;
    }
    const merged = new Map<string, CollaboratorPresence>();
    const staleCutoff = Date.now() - 120_000;

    localPeers.forEach((peer, sid) => {
      if (peer.updatedAt >= staleCutoff) {
        merged.set(sid, peer);
      } else {
        localPeers.delete(sid);
      }
    });

    supabasePeers.forEach((peer, sid) => {
      const existing = merged.get(sid);
      if (!existing || peer.updatedAt >= existing.updatedAt) {
        merged.set(sid, peer);
      }
    });

    merged.set(identity.sessionId, currentPresence);
    callbacks.onPresenceChange(Array.from(merged.values()));
  };

  const bc = openBroadcastChannelSafe(roomName);
  const postBc = (envelope: LocalRoomEnvelope) => {
    if (!bc || destroyed) {
      return;
    }
    try {
      bc.postMessage(envelope);
    } catch {
      // Ignore closed BroadcastChannel errors
    }
  };

  if (bc) {
    bc.onmessage = (event: MessageEvent<unknown>) => {
      if (destroyed || !isPlainRecord(event.data)) {
        return;
      }
      const envelope = event.data as Partial<LocalRoomEnvelope>;
      if (envelope.kind === 'presence_update' && envelope.presence) {
        const parsed = parsePresenceRecord(envelope.presence);
        if (parsed && parsed.sessionId !== identity.sessionId) {
          localPeers.set(parsed.sessionId, parsed);
          emitMergedPresence();
        }
      } else if (
        envelope.kind === 'presence_leave' &&
        typeof envelope.sessionId === 'string'
      ) {
        if (localPeers.delete(envelope.sessionId)) {
          emitMergedPresence();
        }
      } else if (
        envelope.kind === 'presence_ping' &&
        envelope.fromSessionId !== identity.sessionId
      ) {
        postBc({ kind: 'presence_update', presence: currentPresence });
      } else if (envelope.kind === 'doc_patch') {
        const msgId =
          typeof envelope.msgId === 'string' ? envelope.msgId : '';
        if (markAndCheckDuplicate(msgId)) {
          return;
        }
        const parsedPatch = parseIncomingPatch(envelope.patch);
        if (
          parsedPatch &&
          parsedPatch.senderSessionId !== identity.sessionId
        ) {
          callbacks.onRemotePatch(parsedPatch);
        }
      }
    };

    postBc({ kind: 'presence_update', presence: currentPresence });
    postBc({ kind: 'presence_ping', fromSessionId: identity.sessionId });
  }

  const supabaseClient = isSupabaseConfigured ? supabase : null;
  const supabaseChannel = supabaseClient
    ? supabaseClient.channel(roomName, {
        config: {
          broadcast: { self: false },
          presence: { key: identity.sessionId },
        },
      })
    : null;

  if (supabaseChannel) {
    callbacks.onStatusChange?.('connecting');

    supabaseChannel
      .on('presence', { event: 'sync' }, () => {
        if (destroyed) {
          return;
        }
        supabasePeers.clear();
        const state = supabaseChannel.presenceState<Record<string, unknown>>();
        for (const entries of Object.values(state)) {
          for (const rawEntry of entries) {
            const parsed = parsePresenceRecord(rawEntry);
            if (parsed && parsed.sessionId !== identity.sessionId) {
              supabasePeers.set(parsed.sessionId, parsed);
            }
          }
        }
        emitMergedPresence();
      })
      .on<Record<string, unknown>>(
        'broadcast',
        { event: 'doc_patch' },
        ({ payload }) => {
          if (destroyed) {
            return;
          }
          const msgId =
            typeof payload._msgId === 'string' ? payload._msgId : '';
          if (markAndCheckDuplicate(msgId)) {
            return;
          }
          const parsedPatch = parseIncomingPatch(payload);
          if (
            parsedPatch &&
            parsedPatch.senderSessionId !== identity.sessionId
          ) {
            callbacks.onRemotePatch(parsedPatch);
          }
        }
      )
      .on<Record<string, unknown>>(
        'postgres_changes',
        {
          event: 'UPDATE',
          schema: 'public',
          table: DOCUMENTS_TABLE,
          filter: `id=eq.${documentId}`,
        },
        (change) => {
          if (destroyed || !callbacks.onDatabaseDocumentUpdate) {
            return;
          }
          if (
            Date.now() - lastLocalPersistAt <
            SELF_PERSIST_ECHO_WINDOW_MS
          ) {
            return;
          }
          const row = change.new;
          if (!isPlainRecord(row)) {
            return;
          }
          if (
            isValidUuid(ownerId) &&
            typeof row.user_id === 'string' &&
            row.user_id !== ownerId
          ) {
            return;
          }
          const parsed = isPlainRecord(row.parsed_content)
            ? row.parsed_content
            : {};
          const partial: Partial<DocumentItem> = {};
          if (typeof row.title === 'string' && row.title.trim()) {
            partial.title = row.title;
          }
          if (typeof row.content_preview === 'string') {
            partial.contentPreview = row.content_preview;
          }
          if (typeof parsed.wordHtml === 'string') {
            partial.wordHtml = parsed.wordHtml;
          }
          if (Array.isArray(parsed.sheets)) {
            partial.sheets = parsed.sheets as SpreadsheetSheet[];
          }
          if (Array.isArray(parsed.slidesPreview)) {
            partial.slidesPreview =
              parsed.slidesPreview as SlidePreviewItem[];
          }
          if (
            typeof parsed.slideAspectRatio === 'number' &&
            Number.isFinite(parsed.slideAspectRatio)
          ) {
            partial.slideAspectRatio = parsed.slideAspectRatio;
          }
          if (typeof row.updated_at === 'string') {
            partial.lastSavedAt = row.updated_at;
          }
          if (Object.keys(partial).length > 0) {
            callbacks.onDatabaseDocumentUpdate(partial);
          }
        }
      )
      .subscribe((status) => {
        if (destroyed) {
          return;
        }
        if (status === 'SUBSCRIBED') {
          supabaseSubscribed = true;
          callbacks.onStatusChange?.('connected');
          void supabaseChannel.track(currentPresence).catch(() => {});
        } else if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') {
          supabaseSubscribed = false;
          callbacks.onStatusChange?.(bc ? 'connected' : 'error');
        } else if (status === 'CLOSED') {
          supabaseSubscribed = false;
          callbacks.onStatusChange?.(bc ? 'connected' : 'disconnected');
        }
      });
  } else {
    callbacks.onStatusChange?.(bc ? 'connected' : 'disconnected');
  }

  emitMergedPresence();

  const persistSnapshotNow = async (snapshot: DocumentItem) => {
    if (!canEdit) {
      return;
    }
    const nowIso = new Date().toISOString();
    const parsedContent = buildParsedContentPayload(snapshot);

    // Mirror to local collab shares if present
    const localExisting = findLocalShareByDoc(ownerId, documentId);
    if (localExisting) {
      upsertLocalCollabShare({
        ...localExisting,
        updatedAt: nowIso,
        documentSnapshot: sanitizeDocForLocalSnapshot(snapshot),
      });
    }

    if (!isSupabaseConfigured || !supabase) {
      return;
    }

    lastLocalPersistAt = Date.now();

    try {
      const { data: sessionData } = await supabase.auth.getSession();
      const authUserId = sessionData.session?.user?.id;

      if (authUserId && isValidUuid(ownerId) && authUserId === ownerId) {
        await supabase
          .from(DOCUMENTS_TABLE)
          .update({
            title: snapshot.title,
            content_preview: snapshot.contentPreview ?? '',
            parsed_content: parsedContent,
            updated_at: nowIso,
          })
          .eq('user_id', ownerId)
          .eq('id', documentId);
        return;
      }

      const activeToken = shareToken || snapshot.shareToken;
      if (activeToken) {
        await supabase.rpc('update_shared_document_content', {
          p_share_token: activeToken,
          p_title: snapshot.title,
          p_content_preview: snapshot.contentPreview ?? '',
          p_parsed_content: parsedContent,
        });
      }
    } catch {
      // Ignore transient persist failures; next edit retries cleanly
    }
  };

  return {
    sessionId: identity.sessionId,
    roomName,
    selfColor,

    updatePresence(patch) {
      if (destroyed) {
        return;
      }
      currentPresence = {
        ...currentPresence,
        ...patch,
        sessionId: identity.sessionId,
        color: selfColor,
        updatedAt: Date.now(),
      };
      emitMergedPresence();
      postBc({ kind: 'presence_update', presence: currentPresence });
      if (supabaseChannel && supabaseSubscribed) {
        void supabaseChannel.track(currentPresence).catch(() => {});
      }
    },

    broadcastPatch(patch) {
      if (destroyed || !canEdit) {
        return;
      }
      const fullPatch: DocRealtimePatch = {
        ...patch,
        senderSessionId: identity.sessionId,
        sentAt: Date.now(),
      } as DocRealtimePatch;
      const msgId = `${identity.sessionId}-${fullPatch.sentAt}-${Math.random().toString(36).slice(2, 7)}`;

      postBc({ kind: 'doc_patch', msgId, patch: fullPatch });

      if (supabaseChannel && supabaseSubscribed) {
        void supabaseChannel
          .send({
            type: 'broadcast',
            event: 'doc_patch',
            payload: {
              ...fullPatch,
              _msgId: msgId,
            },
          })
          .catch(() => {});
      }
    },

    schedulePersist(docSnapshot: DocumentItem) {
      if (destroyed || !canEdit) {
        return;
      }
      pendingSnapshot = docSnapshot;
      if (persistTimerId !== null) {
        window.clearTimeout(persistTimerId);
      }
      persistTimerId = window.setTimeout(() => {
        persistTimerId = null;
        const snapshotToSave = pendingSnapshot;
        pendingSnapshot = null;
        if (snapshotToSave && !destroyed) {
          void persistSnapshotNow(snapshotToSave);
        }
      }, PERSIST_DEBOUNCE_MS);
    },

    leave() {
      if (destroyed) {
        return;
      }
      if (persistTimerId !== null) {
        window.clearTimeout(persistTimerId);
        persistTimerId = null;
        if (pendingSnapshot && canEdit) {
          void persistSnapshotNow(pendingSnapshot);
        }
        pendingSnapshot = null;
      }
      destroyed = true;

      if (bc) {
        try {
          bc.postMessage({
            kind: 'presence_leave',
            sessionId: identity.sessionId,
          } satisfies LocalRoomEnvelope);
          bc.close();
        } catch {
          // Ignore close errors
        }
      }

      if (supabaseChannel && supabaseClient) {
        void supabaseChannel.untrack().catch(() => {});
        void supabaseClient.removeChannel(supabaseChannel).catch(() => {});
      }
    },
  };
}

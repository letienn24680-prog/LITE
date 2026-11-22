import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { SidebarRail } from './components/layout/SidebarRail';
import { MobileTopHeader } from './components/layout/MobileTopHeader';
import { MobileBottomNav } from './components/layout/MobileBottomNav';
import { DocumentLibrary } from './components/workspace/DocumentLibrary';
import {
  LocationPanel,
  type LocationSourceKey,
} from './components/workspace/LocationPanel';
import { SlideWorkspace } from './components/workspaces/SlideWorkspace';
import { WordWorkspace } from './components/workspaces/WordWorkspace';
import { ExcelWorkspace } from './components/workspaces/ExcelWorkspace';
import { PdfWorkspace } from './components/workspaces/PdfWorkspace';
import { StageActionModal } from './components/modals/StageActionModal';
import { AuthModal } from './components/modals/AuthModal';
import { SettingsModal } from './components/modals/SettingsModal';
import { HelpModal } from './components/modals/HelpModal';
import { CloudActionModal } from './components/modals/CloudActionModal';
import { GoogleDriveBrowserModal } from './components/modals/GoogleDriveBrowserModal';
import { ShareCollabModal } from './components/modals/ShareCollabModal';
import { MobileRemoteView } from './components/remote/MobileRemoteView';
import { SharedDocumentView } from './components/share/SharedDocumentView';
import {
  createNewEditableDocument,
  LEGACY_MOCK_DOCUMENT_IDS,
  type CreateEditableDocumentParams,
} from './data/mockDocuments';
import { isSupabaseConfigured, supabase } from './lib/supabase';
import {
  buildCloudUserSummary,
  deleteCloudDocument,
  downloadCloudFileArrayBuffer,
  fetchUserDocuments,
  getCloudFileObjectUrl,
  syncLocalDocumentsToCloud,
  touchDocumentLastAccessed,
  upsertCloudDocument,
} from './services/cloudSync';
import {
  getR2PresignedDownloadUrl,
  isR2ApiConfigured,
  uploadFileViaR2PresignedUrl,
} from './services/r2Storage';
import { removeLocalSharesForDocumentIds } from './services/shareService';
import {
  getClientSessionId,
  joinDocumentRealtimeRoom,
  loadSharedWithMeDocuments,
  resolveSharedDocument,
  type CollabPermission,
  type CollaboratorPresence,
  type DocRealtimePatch,
  type RealtimeRoomHandle,
} from './services/realtimeCollab';
import {
  clearStoredGoogleDriveToken,
  getStoredGoogleDriveToken,
  setStoredGoogleDriveToken,
  type GoogleDriveFileItem,
} from './services/googleDrive';
import { parsePptxArrayBuffer, parseUploadedFile } from './utils/fileParser';
import type {
  CloudUserSummary,
  DocumentItem,
  FilterKey,
  ModalType,
  ShareAccessLevel,
  SidebarNavKey,
  SlidePreviewItem,
  SyncStatus,
  UserSettings,
  UtilityModalType,
  WordPaperTheme,
} from './types/document';

const STORAGE_KEY = 'lite_saved_documents_v1';
const SETTINGS_STORAGE_KEY = 'lite_user_settings_v1';

const LEGACY_MOCK_ID_SET = new Set<string>(LEGACY_MOCK_DOCUMENT_IDS);

const DEFAULT_USER_SETTINGS: UserSettings = {
  autoSyncOnEdit: true,
  defaultWordTheme: 'light',
};

function formatFileSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function safeDecodeSegment(raw: string): string {
  try {
    return decodeURIComponent(raw).trim();
  } catch {
    return raw.trim();
  }
}

function parseRemoteRoomFromHash(hash: string): string | null {
  const match = hash.match(/^#remote\/(.+)$/i);
  if (!match) {
    return null;
  }
  const decoded = safeDecodeSegment(match[1]);
  return decoded.length > 0 ? decoded : null;
}

function parseShareTokenFromHash(hash: string): string | null {
  const match = hash.match(/^#share\/(.+)$/i);
  if (!match) {
    return null;
  }
  const decoded = safeDecodeSegment(match[1]);
  return decoded.length > 0 ? decoded : null;
}

function parseDocIdFromHash(hash: string): string | null {
  const match = hash.match(/^#doc\/(.+)$/);
  if (!match) {
    return null;
  }
  const decoded = safeDecodeSegment(match[1]);
  return decoded.length > 0 ? decoded : null;
}

function parseCollabShareTokenFromSearch(search: string): string | null {
  const raw = new URLSearchParams(search).get('share');
  if (!raw) {
    return null;
  }
  const trimmed = raw.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function isValidWordTheme(value: unknown): value is WordPaperTheme {
  return value === 'light' || value === 'sepia' || value === 'dark';
}

function loadInitialUserSettings(): UserSettings {
  try {
    const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) {
      return DEFAULT_USER_SETTINGS;
    }
    const parsed: unknown = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') {
      return DEFAULT_USER_SETTINGS;
    }
    const candidate = parsed as Partial<UserSettings>;
    return {
      autoSyncOnEdit:
        typeof candidate.autoSyncOnEdit === 'boolean'
          ? candidate.autoSyncOnEdit
          : DEFAULT_USER_SETTINGS.autoSyncOnEdit,
      defaultWordTheme: isValidWordTheme(candidate.defaultWordTheme)
        ? candidate.defaultWordTheme
        : DEFAULT_USER_SETTINGS.defaultWordTheme,
    };
  } catch {
    return DEFAULT_USER_SETTINGS;
  }
}

function persistUserSettings(settings: UserSettings): void {
  try {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Ignore localStorage quota/privacy restrictions
  }
}

function isValidStoredDocument(value: unknown): value is DocumentItem {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const candidate = value as Partial<DocumentItem>;
  const validType =
    candidate.type === 'pptx' ||
    candidate.type === 'docx' ||
    candidate.type === 'xlsx' ||
    candidate.type === 'pdf';
  return (
    typeof candidate.id === 'string' &&
    candidate.id.length > 0 &&
    typeof candidate.title === 'string' &&
    validType &&
    Array.isArray(candidate.slidesPreview) &&
    Array.isArray(candidate.outline)
  );
}

function sanitizeDocumentForStorage(doc: DocumentItem): DocumentItem {
  const safeSlides = doc.slidesPreview.map((slide) => {
    const hasImages = Boolean(slide.images && slide.images.length > 0);
    const hasBlobBg = Boolean(slide.backgroundImageSrc?.startsWith('blob:'));
    if (!hasImages && !hasBlobBg) {
      return slide;
    }
    const nextSlide = { ...slide };
    if (nextSlide.images && nextSlide.images.length > 0) {
      const persistableImages = nextSlide.images.filter(
        (img) => !img.src.startsWith('blob:')
      );
      if (persistableImages.length > 0) {
        nextSlide.images = persistableImages;
      } else {
        delete nextSlide.images;
      }
    }
    if (nextSlide.backgroundImageSrc?.startsWith('blob:')) {
      delete nextSlide.backgroundImageSrc;
    }
    return nextSlide;
  });

  const nextDoc: DocumentItem = {
    ...doc,
    slidesPreview: safeSlides,
  };
  delete nextDoc.sourceFile;
  delete nextDoc.rawArrayBuffer;
  if (nextDoc.objectUrl?.startsWith('blob:')) {
    delete nextDoc.objectUrl;
  }
  if (nextDoc.fileUrl?.startsWith('blob:')) {
    delete nextDoc.fileUrl;
  }
  return nextDoc;
}

function persistDocumentsToStorage(docs: DocumentItem[]): {
  ok: boolean;
  quotaExceeded: boolean;
} {
  try {
    const payload = docs
      .filter((doc) => !LEGACY_MOCK_ID_SET.has(doc.id))
      .map(sanitizeDocumentForStorage);
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    return { ok: true, quotaExceeded: false };
  } catch (error) {
    const isQuota =
      error instanceof DOMException &&
      (error.name === 'QuotaExceededError' ||
        error.name === 'NS_ERROR_DOM_QUOTA_REACHED' ||
        error.code === 22);
    return { ok: false, quotaExceeded: isQuota };
  }
}

function loadInitialDocuments(): DocumentItem[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const seenIds = new Set<string>();
    const merged: DocumentItem[] = [];
    let filteredLegacyMock = false;

    for (const entry of parsed) {
      if (!isValidStoredDocument(entry)) {
        continue;
      }
      if (LEGACY_MOCK_ID_SET.has(entry.id)) {
        filteredLegacyMock = true;
        continue;
      }
      if (seenIds.has(entry.id)) {
        continue;
      }
      seenIds.add(entry.id);
      merged.push(sanitizeDocumentForStorage(entry));
    }

    if (filteredLegacyMock) {
      persistDocumentsToStorage(merged);
      removeLocalSharesForDocumentIds(LEGACY_MOCK_DOCUMENT_IDS);
    }

    return merged;
  } catch {
    return [];
  }
}

function preserveInMemoryRuntimeFields(
  incoming: DocumentItem,
  existing?: DocumentItem
): DocumentItem {
  if (!existing) {
    return incoming;
  }
  const rawArrayBuffer = existing.rawArrayBuffer ?? incoming.rawArrayBuffer;
  const slideAspectRatio =
    existing.slideAspectRatio ?? incoming.slideAspectRatio;
  return {
    ...incoming,
    ...(existing.sourceFile ? { sourceFile: existing.sourceFile } : {}),
    ...(rawArrayBuffer ? { rawArrayBuffer } : {}),
    ...(typeof slideAspectRatio === 'number' ? { slideAspectRatio } : {}),
    ...(existing.fileUrl && !incoming.fileUrl
      ? { fileUrl: existing.fileUrl }
      : {}),
    ...(existing.objectUrl && !incoming.objectUrl
      ? { objectUrl: existing.objectUrl }
      : {}),
    ...(existing.r2ObjectKey && !incoming.r2ObjectKey
      ? { r2ObjectKey: existing.r2ObjectKey }
      : {}),
  };
}

function needsCloudFileHydration(doc: DocumentItem): boolean {
  if (!doc.r2ObjectKey && !doc.storagePath) {
    return false;
  }
  if (doc.type === 'pdf') {
    return !doc.fileUrl && !doc.objectUrl;
  }
  if (doc.type === 'docx' || doc.type === 'pptx') {
    return !doc.rawArrayBuffer;
  }
  return false;
}

function mergeHydratedPptxSlides(
  currentSlides: SlidePreviewItem[],
  hydratedSlides: SlidePreviewItem[]
): SlidePreviewItem[] {
  if (currentSlides.length === 0) {
    return hydratedSlides;
  }
  return currentSlides.map((currentSlide, idx) => {
    const hydratedSlide = hydratedSlides[idx];
    if (!hydratedSlide) {
      return currentSlide;
    }

    const mergedBlocks =
      currentSlide.blocks && currentSlide.blocks.length > 0
        ? currentSlide.blocks.map((block, bIdx) => {
            const hydratedBlock = hydratedSlide.blocks?.[bIdx];
            if (!hydratedBlock) {
              return block;
            }
            return {
              ...hydratedBlock,
              ...block,
              xPct: block.xPct ?? hydratedBlock.xPct,
              yPct: block.yPct ?? hydratedBlock.yPct,
              wPct: block.wPct ?? hydratedBlock.wPct,
              hPct: block.hPct ?? hydratedBlock.hPct,
              fillColor: block.fillColor ?? hydratedBlock.fillColor,
              borderColor: block.borderColor ?? hydratedBlock.borderColor,
              fontFamily: block.fontFamily ?? hydratedBlock.fontFamily,
              valign: block.valign ?? hydratedBlock.valign,
              kind: block.kind ?? hydratedBlock.kind,
              tableRows: block.tableRows ?? hydratedBlock.tableRows,
            };
          })
        : hydratedSlide.blocks;

    const existingCustomImages = (currentSlide.images ?? []).filter(
      (img) => !img.src.startsWith('blob:')
    );
    const hydratedImages = hydratedSlide.images ?? [];
    const mergedImages =
      hydratedImages.length > 0
        ? [...hydratedImages, ...existingCustomImages]
        : currentSlide.images;

    return {
      ...hydratedSlide,
      ...currentSlide,
      title: currentSlide.title || hydratedSlide.title,
      subtitle: currentSlide.subtitle ?? hydratedSlide.subtitle,
      bullets:
        currentSlide.bullets.length > 0
          ? currentSlide.bullets
          : hydratedSlide.bullets,
      speakerNote: currentSlide.speakerNote ?? hydratedSlide.speakerNote,
      backgroundColor:
        currentSlide.backgroundColor ?? hydratedSlide.backgroundColor,
      backgroundGradient:
        currentSlide.backgroundGradient ?? hydratedSlide.backgroundGradient,
      backgroundImageSrc:
        currentSlide.backgroundImageSrc ?? hydratedSlide.backgroundImageSrc,
      slideWidthRatio:
        currentSlide.slideWidthRatio ?? hydratedSlide.slideWidthRatio,
      isMasterPositioned:
        currentSlide.isMasterPositioned ?? hydratedSlide.isMasterPositioned,
      blocks: mergedBlocks,
      images: mergedImages,
    };
  });
}

function isUnsyncedLocalCandidate(doc: DocumentItem): boolean {
  return (
    !LEGACY_MOCK_ID_SET.has(doc.id) &&
    !doc.isSyncedToCloud &&
    (Boolean(doc.isUploaded) || Boolean(doc.lastSavedAt))
  );
}

export function App() {
  const [documents, setDocuments] = useState<DocumentItem[]>(() =>
    loadInitialDocuments()
  );
  const [remoteRoomCode, setRemoteRoomCode] = useState<string | null>(() =>
    parseRemoteRoomFromHash(window.location.hash)
  );
  const [sharedToken, setSharedToken] = useState<string | null>(() => {
    if (parseRemoteRoomFromHash(window.location.hash)) {
      return null;
    }
    return parseShareTokenFromHash(window.location.hash);
  });
  const [collabShareQueryToken, setCollabShareQueryToken] = useState<
    string | null
  >(() => parseCollabShareTokenFromSearch(window.location.search));
  const [collabSharedDoc, setCollabSharedDoc] = useState<DocumentItem | null>(
    null
  );
  const [isResolvingCollabShare, setIsResolvingCollabShare] = useState<boolean>(
    () => Boolean(parseCollabShareTokenFromSearch(window.location.search))
  );
  const [collabShareError, setCollabShareError] = useState<string | null>(null);
  const [collabCollaborators, setCollabCollaborators] = useState<
    CollaboratorPresence[]
  >([]);
  const [latestRemotePatch, setLatestRemotePatch] =
    useState<DocRealtimePatch | null>(null);
  const [activeWorkspaceDocId, setActiveWorkspaceDocId] = useState<
    string | null
  >(() => {
    if (
      parseRemoteRoomFromHash(window.location.hash) ||
      parseShareTokenFromHash(window.location.hash)
    ) {
      return null;
    }
    const hashId = parseDocIdFromHash(window.location.hash);
    if (!hashId) {
      return null;
    }
    return documents.some((doc) => doc.id === hashId) ? hashId : null;
  });
  const [lastActiveDocId, setLastActiveDocId] = useState<string | null>(() => {
    const hashId = parseDocIdFromHash(window.location.hash);
    if (hashId && documents.some((doc) => doc.id === hashId)) {
      return hashId;
    }
    return documents[0]?.id ?? null;
  });
  const [activeNav, setActiveNav] = useState<SidebarNavKey>('home');
  const [activeLocation, setActiveLocation] = useState<LocationSourceKey>('all');
  const [isLocationPanelOpen, setIsLocationPanelOpen] = useState<boolean>(false);
  const [isDriveModalOpen, setIsDriveModalOpen] = useState<boolean>(false);
  const [hasGoogleToken, setHasGoogleToken] = useState<boolean>(() =>
    Boolean(getStoredGoogleDriveToken())
  );
  const [activeFilter, setActiveFilter] = useState<FilterKey>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeModal, setActiveModal] = useState<ModalType>(null);
  const [modalDocOverride, setModalDocOverride] =
    useState<DocumentItem | null>(null);
  const [uploadNotice, setUploadNotice] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [uploadProgress, setUploadProgress] = useState<number | null>(null);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);

  // Account, Settings & Cloud Sync States
  const [userSettings, setUserSettings] = useState<UserSettings>(() =>
    loadInitialUserSettings()
  );
  const [currentUser, setCurrentUser] = useState<CloudUserSummary | null>(null);
  const [syncStatus, setSyncStatus] = useState<SyncStatus>(() =>
    isSupabaseConfigured ? 'idle' : 'local'
  );
  const [lastSyncedAt, setLastSyncedAt] = useState<string | null>(null);
  const [activeUtilityModal, setActiveUtilityModal] =
    useState<UtilityModalType>(null);
  const [pendingMigrationDocs, setPendingMigrationDocs] = useState<
    DocumentItem[]
  >([]);
  const [isMigratingLocal, setIsMigratingLocal] = useState<boolean>(false);
  const [isLoggingOut, setIsLoggingOut] = useState<boolean>(false);
  const [hydratingPdfDocId, setHydratingPdfDocId] = useState<string | null>(
    null
  );

  const createdObjectUrlsRef = useRef<string[]>([]);
  const deviceFileInputRef = useRef<HTMLInputElement | null>(null);
  const activeNavRef = useRef<SidebarNavKey>(activeNav);
  activeNavRef.current = activeNav;
  const documentsRef = useRef<DocumentItem[]>(documents);
  documentsRef.current = documents;

  const currentUserRef = useRef<CloudUserSummary | null>(currentUser);
  currentUserRef.current = currentUser;

  const lastLoadedUserIdRef = useRef<string | null>(null);
  const inFlightSyncUserIdRef = useRef<string | null>(null);
  const hydratingPdfDocIdRef = useRef<string | null>(null);
  const autoHydrateAttemptedRef = useRef<Set<string>>(new Set());
  const saveQueueByDocRef = useRef<Map<string, Promise<void>>>(new Map());
  const saveSeqByDocRef = useRef<Map<string, number>>(new Map());
  const realtimeRoomRef = useRef<RealtimeRoomHandle | null>(null);
  const selfSessionId = useMemo(() => getClientSessionId(), []);
  const collabSharedDocRef = useRef<DocumentItem | null>(collabSharedDoc);
  collabSharedDocRef.current = collabSharedDoc;
  const activeCollabPermissionRef = useRef<'owner' | 'edit' | 'view'>('owner');

  useEffect(() => {
    return () => {
      createdObjectUrlsRef.current.forEach((url) => {
        URL.revokeObjectURL(url);
      });
      createdObjectUrlsRef.current = [];
    };
  }, []);

  // Priority Hash & Query Routing: ?share=<token> -> #remote/<roomCode> -> #share/<shareToken> -> #doc/<id>
  useEffect(() => {
    const syncFromLocation = () => {
      const nextQueryShare = parseCollabShareTokenFromSearch(
        window.location.search
      );
      setCollabShareQueryToken(nextQueryShare);
      if (!nextQueryShare) {
        setCollabSharedDoc(null);
        setCollabShareError(null);
        setIsResolvingCollabShare(false);
      }

      const currentHash = window.location.hash;
      const parsedRemote = parseRemoteRoomFromHash(currentHash);
      if (parsedRemote) {
        setRemoteRoomCode(parsedRemote);
        setSharedToken(null);
        setActiveWorkspaceDocId(null);
        setActiveModal(null);
        setModalDocOverride(null);
        setActiveUtilityModal(null);
        return;
      }

      const parsedShare = parseShareTokenFromHash(currentHash);
      if (parsedShare) {
        setRemoteRoomCode(null);
        setSharedToken(parsedShare);
        setActiveWorkspaceDocId(null);
        setActiveModal(null);
        setModalDocOverride(null);
        setActiveUtilityModal(null);
        return;
      }

      setRemoteRoomCode(null);
      setSharedToken(null);

      if (nextQueryShare) {
        return;
      }

      const hashId = parseDocIdFromHash(currentHash);
      if (hashId && documentsRef.current.some((doc) => doc.id === hashId)) {
        setActiveWorkspaceDocId(hashId);
        setLastActiveDocId(hashId);
        setActiveModal(null);
        setModalDocOverride(null);
      } else {
        setActiveWorkspaceDocId(null);
        setActiveModal(null);
        setModalDocOverride(null);
      }
    };

    window.addEventListener('hashchange', syncFromLocation);
    window.addEventListener('popstate', syncFromLocation);
    return () => {
      window.removeEventListener('hashchange', syncFromLocation);
      window.removeEventListener('popstate', syncFromLocation);
    };
  }, []);

  // Resolve ?share=<token> collaboration invite link (works for both guests and signed-in users)
  useEffect(() => {
    if (!collabShareQueryToken) {
      setIsResolvingCollabShare(false);
      setCollabShareError(null);
      return;
    }

    let cancelled = false;
    setIsResolvingCollabShare(true);
    setCollabShareError(null);

    void resolveSharedDocument(
      collabShareQueryToken,
      currentUserRef.current?.id
    )
      .then((resolved) => {
        if (cancelled) {
          return;
        }

        if (resolved.status === 'ready' && resolved.document) {
          const sharedDoc: DocumentItem = {
            ...resolved.document,
            collabPermission: resolved.permission,
            sharedOwnerId: resolved.ownerId,
            shareToken: resolved.shareToken,
            isShared: true,
          };

          setCollabSharedDoc(sharedDoc);
          setActiveWorkspaceDocId(sharedDoc.id);
          setLastActiveDocId(sharedDoc.id);
          setActiveModal(null);
          setModalDocOverride(null);

          if (currentUserRef.current) {
            setDocuments((prev) => {
              const existingIdx = prev.findIndex((d) => d.id === sharedDoc.id);
              const nextDocs =
                existingIdx >= 0
                  ? prev.map((d) =>
                      d.id === sharedDoc.id
                        ? { ...d, ...sharedDoc, isShared: true }
                        : d
                    )
                  : [sharedDoc, ...prev];
              persistDocumentsToStorage(nextDocs);
              return nextDocs;
            });
          }
        } else {
          setCollabSharedDoc(null);
          const errorByStatus: Record<string, string> = {
            revoked: 'Liên kết chia sẻ này đã bị chủ sở hữu thu hồi.',
            private: 'Tài liệu này đang ở chế độ Hạn chế (Riêng tư).',
            not_found: 'Liên kết chia sẻ không tồn tại hoặc đã hết hạn.',
            error:
              resolved.errorMessage ||
              'Không thể mở liên kết chia sẻ lúc này.',
          };
          setCollabShareError(
            errorByStatus[resolved.status] ??
              'Không thể mở liên kết chia sẻ lúc này.'
          );
        }
      })
      .catch((error: unknown) => {
        if (cancelled) {
          return;
        }
        setCollabSharedDoc(null);
        setCollabShareError(
          error instanceof Error
            ? error.message
            : 'Không thể mở liên kết chia sẻ lúc này.'
        );
      })
      .finally(() => {
        if (!cancelled) {
          setIsResolvingCollabShare(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [collabShareQueryToken]);

  const syncSessionForUser = useCallback(async (user: CloudUserSummary) => {
    if (
      lastLoadedUserIdRef.current === user.id ||
      inFlightSyncUserIdRef.current === user.id
    ) {
      return;
    }

    inFlightSyncUserIdRef.current = user.id;
    setSyncStatus('syncing');

    try {
      const [fetchedCloudDocs, sharedWithMeDocs] = await Promise.all([
        fetchUserDocuments(user.id),
        loadSharedWithMeDocuments(user.id).catch(() => [] as DocumentItem[]),
      ]);
      if (currentUserRef.current?.id !== user.id) {
        return;
      }

      // Silently purge any legacy mock records that were previously synced to Supabase
      const legacyCloudDocs = fetchedCloudDocs.filter((cloudDoc) =>
        LEGACY_MOCK_ID_SET.has(cloudDoc.id)
      );
      if (legacyCloudDocs.length > 0) {
        removeLocalSharesForDocumentIds(legacyCloudDocs.map((d) => d.id));
        legacyCloudDocs.forEach((legacyDoc) => {
          void deleteCloudDocument(user.id, legacyDoc).catch(() => undefined);
        });
      }

      const existingById = new Map<string, DocumentItem>(
        documentsRef.current.map((doc) => [doc.id, doc])
      );
      const cloudDocs = fetchedCloudDocs
        .filter((cloudDoc) => !LEGACY_MOCK_ID_SET.has(cloudDoc.id))
        .map((cloudDoc) =>
          preserveInMemoryRuntimeFields(
            cloudDoc,
            existingById.get(cloudDoc.id)
          )
        );
      const cloudIds = new Set<string>(cloudDocs.map((doc) => doc.id));

      const sharedDocs = sharedWithMeDocs
        .filter(
          (sharedDoc) =>
            !LEGACY_MOCK_ID_SET.has(sharedDoc.id) && !cloudIds.has(sharedDoc.id)
        )
        .map((sharedDoc) =>
          preserveInMemoryRuntimeFields(
            { ...sharedDoc, isShared: true },
            existingById.get(sharedDoc.id)
          )
        );
      sharedDocs.forEach((doc) => {
        cloudIds.add(doc.id);
      });

      const activeCollabDoc = collabSharedDocRef.current;
      if (
        activeCollabDoc &&
        !LEGACY_MOCK_ID_SET.has(activeCollabDoc.id) &&
        !cloudIds.has(activeCollabDoc.id)
      ) {
        sharedDocs.unshift({ ...activeCollabDoc, isShared: true });
        cloudIds.add(activeCollabDoc.id);
        if (activeCollabDoc.shareToken) {
          void resolveSharedDocument(
            activeCollabDoc.shareToken,
            user.id
          ).catch(() => undefined);
        }
      }

      // Keep unsynced local candidates (excluding legacy mocks) untouched in memory while asking user to migrate
      const localCandidates = documentsRef.current.filter(
        (doc) =>
          !LEGACY_MOCK_ID_SET.has(doc.id) &&
          isUnsyncedLocalCandidate(doc) &&
          !cloudIds.has(doc.id)
      );

      const mergedDocs = [...cloudDocs, ...sharedDocs, ...localCandidates];
      setDocuments(mergedDocs);
      persistDocumentsToStorage(mergedDocs);

      const currentHash = window.location.hash;
      const isSpecialRoute =
        Boolean(parseRemoteRoomFromHash(currentHash)) ||
        Boolean(parseShareTokenFromHash(currentHash)) ||
        Boolean(parseCollabShareTokenFromSearch(window.location.search));

      if (!isSpecialRoute) {
        const hashId = parseDocIdFromHash(currentHash);
        if (hashId && mergedDocs.some((doc) => doc.id === hashId)) {
          setActiveWorkspaceDocId(hashId);
          setLastActiveDocId(hashId);
        } else {
          setActiveWorkspaceDocId(null);
          setLastActiveDocId((prev) =>
            prev && mergedDocs.some((doc) => doc.id === prev)
              ? prev
              : (mergedDocs[0]?.id ?? null)
          );
        }
      }

      lastLoadedUserIdRef.current = user.id;
      setLastSyncedAt(new Date().toISOString());
      setSyncStatus('synced');

      if (localCandidates.length > 0 && !isSpecialRoute) {
        setPendingMigrationDocs(localCandidates);
        setActiveUtilityModal('migrateLocal');
      } else {
        setPendingMigrationDocs([]);
      }
    } catch (error) {
      setSyncStatus('error');
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể tải danh sách tài liệu từ tài khoản đám mây.';
      setUploadNotice(message);
    } finally {
      if (inFlightSyncUserIdRef.current === user.id) {
        inFlightSyncUserIdRef.current = null;
      }
    }
  }, []);

  // Initialize Supabase Auth session and subscribe to auth state transitions
  useEffect(() => {
    if (!isSupabaseConfigured || !supabase) {
      setSyncStatus('local');
      return;
    }

    let isMounted = true;

    const handleSessionUser = (
      sessionUser: {
        id: string;
        email?: string | null;
        user_metadata?: Record<string, unknown>;
      } | null
    ) => {
      if (!isMounted) {
        return;
      }

      if (sessionUser) {
        const summary = buildCloudUserSummary(sessionUser);
        currentUserRef.current = summary;
        setCurrentUser(summary);
        void syncSessionForUser(summary);
      } else {
        const hadUser =
          currentUserRef.current !== null ||
          lastLoadedUserIdRef.current !== null;
        currentUserRef.current = null;
        lastLoadedUserIdRef.current = null;
        inFlightSyncUserIdRef.current = null;
        setCurrentUser(null);
        setPendingMigrationDocs([]);
        setLastSyncedAt(null);
        setSyncStatus('idle');
        if (hadUser) {
          clearStoredGoogleDriveToken();
          setHasGoogleToken(false);
          autoHydrateAttemptedRef.current.clear();
          const localOnly = documentsRef.current.filter(
            (doc) =>
              !doc.isSyncedToCloud &&
              !doc.userId &&
              !LEGACY_MOCK_ID_SET.has(doc.id)
          );
          persistDocumentsToStorage(localOnly);
          const localDocs = loadInitialDocuments();
          setDocuments(localDocs);
          setActiveWorkspaceDocId(null);
          setLastActiveDocId(localDocs[0]?.id ?? null);
          setActiveModal(null);
          setModalDocOverride(null);
          setActiveUtilityModal(null);
          const isSpecialRoute =
            Boolean(parseRemoteRoomFromHash(window.location.hash)) ||
            Boolean(parseShareTokenFromHash(window.location.hash));
          if (window.location.hash && !isSpecialRoute) {
            window.history.pushState(
              null,
              '',
              `${window.location.pathname}${window.location.search}`
            );
          }
        }
      }
    };

    void supabase.auth
      .getSession()
      .then(({ data }) => {
        if (data.session?.provider_token) {
          setStoredGoogleDriveToken(data.session.provider_token);
          setHasGoogleToken(true);
        }
        handleSessionUser(data.session?.user ?? null);
      })
      .catch(() => {
        if (isMounted) {
          setSyncStatus('error');
        }
      });

    const { data: authSubscription } = supabase.auth.onAuthStateChange(
      (event, session) => {
        if (session?.provider_token) {
          setStoredGoogleDriveToken(session.provider_token);
          setHasGoogleToken(true);
        } else if (event === 'SIGNED_OUT' || !session) {
          clearStoredGoogleDriveToken();
          setHasGoogleToken(false);
        }
        handleSessionUser(session?.user ?? null);
      }
    );

    return () => {
      isMounted = false;
      authSubscription.subscription.unsubscribe();
    };
  }, [syncSessionForUser]);

  const hydrateCloudPdfDocument = useCallback(async (doc: DocumentItem) => {
    const hasR2Source = Boolean(doc.r2ObjectKey) && isR2ApiConfigured;
    const hasSupabaseSource = Boolean(doc.storagePath) && isSupabaseConfigured;
    if (!hasR2Source && !hasSupabaseSource) {
      return;
    }
    if (hydratingPdfDocIdRef.current === doc.id) {
      return;
    }

    hydratingPdfDocIdRef.current = doc.id;
    setHydratingPdfDocId(doc.id);

    try {
      if (doc.type === 'pdf') {
        let resolvedObjectUrl: string | null = null;

        if (doc.r2ObjectKey && isR2ApiConfigured) {
          try {
            const sessionRes = supabase
              ? await supabase.auth.getSession()
              : null;
            const accessToken = sessionRes?.data.session?.access_token ?? null;
            const { downloadUrl } = await getR2PresignedDownloadUrl({
              objectKey: doc.r2ObjectKey,
              accessToken,
            });
            const fileResponse = await fetch(downloadUrl);
            if (!fileResponse.ok) {
              throw new Error(
                `Không thể tải tệp PDF từ R2 (HTTP ${fileResponse.status}).`
              );
            }
            const pdfBlob = await fileResponse.blob();
            resolvedObjectUrl = URL.createObjectURL(pdfBlob);
          } catch (r2Error) {
            if (!doc.storagePath || !isSupabaseConfigured) {
              throw r2Error;
            }
          }
        }

        if (!resolvedObjectUrl && doc.storagePath && isSupabaseConfigured) {
          const { objectUrl } = await getCloudFileObjectUrl(doc.storagePath);
          resolvedObjectUrl = objectUrl;
        }

        if (!resolvedObjectUrl) {
          return;
        }

        createdObjectUrlsRef.current.push(resolvedObjectUrl);
        const finalObjectUrl = resolvedObjectUrl;

        setDocuments((prev) =>
          prev.map((item) =>
            item.id === doc.id
              ? {
                  ...item,
                  fileUrl: finalObjectUrl,
                  objectUrl: finalObjectUrl,
                }
              : item
          )
        );
        setCollabSharedDoc((prev) =>
          prev && prev.id === doc.id
            ? {
                ...prev,
                fileUrl: finalObjectUrl,
                objectUrl: finalObjectUrl,
              }
            : prev
        );
        return;
      }

      if (doc.type === 'docx' || doc.type === 'pptx') {
        let resolvedBuffer: ArrayBuffer | null = null;

        if (doc.r2ObjectKey && isR2ApiConfigured) {
          try {
            const sessionRes = supabase
              ? await supabase.auth.getSession()
              : null;
            const accessToken = sessionRes?.data.session?.access_token ?? null;
            const { downloadUrl } = await getR2PresignedDownloadUrl({
              objectKey: doc.r2ObjectKey,
              accessToken,
            });
            const fileResponse = await fetch(downloadUrl);
            if (!fileResponse.ok) {
              throw new Error(
                `Không thể tải tệp gốc từ R2 (HTTP ${fileResponse.status}).`
              );
            }
            resolvedBuffer = await fileResponse.arrayBuffer();
          } catch (r2Error) {
            if (!doc.storagePath || !isSupabaseConfigured) {
              throw r2Error;
            }
          }
        }

        if (!resolvedBuffer && doc.storagePath && isSupabaseConfigured) {
          resolvedBuffer = await downloadCloudFileArrayBuffer(doc.storagePath);
        }

        if (!resolvedBuffer) {
          return;
        }

        const buffer = resolvedBuffer;

        if (doc.type === 'docx') {
          setDocuments((prev) =>
            prev.map((item) =>
              item.id === doc.id
                ? {
                    ...item,
                    rawArrayBuffer: buffer,
                  }
                : item
            )
          );
          setCollabSharedDoc((prev) =>
            prev && prev.id === doc.id
              ? {
                  ...prev,
                  rawArrayBuffer: buffer,
                }
              : prev
          );
          return;
        }

        // doc.type === 'pptx'
        const parsedPptx = await parsePptxArrayBuffer(buffer.slice(0), {
          docIdPrefix: doc.id,
        });
        if (parsedPptx.createdObjectUrls.length > 0) {
          createdObjectUrlsRef.current.push(...parsedPptx.createdObjectUrls);
        }

        setDocuments((prev) =>
          prev.map((item) => {
            if (item.id !== doc.id) {
              return item;
            }
            return {
              ...item,
              rawArrayBuffer: buffer,
              slideAspectRatio: parsedPptx.slideAspectRatio,
              slidesPreview: mergeHydratedPptxSlides(
                item.slidesPreview,
                parsedPptx.slidesPreview
              ),
              outline:
                item.outline.length > 0 ? item.outline : parsedPptx.outline,
            };
          })
        );
        setCollabSharedDoc((prev) => {
          if (!prev || prev.id !== doc.id) {
            return prev;
          }
          return {
            ...prev,
            rawArrayBuffer: buffer,
            slideAspectRatio: parsedPptx.slideAspectRatio,
            slidesPreview: mergeHydratedPptxSlides(
              prev.slidesPreview,
              parsedPptx.slidesPreview
            ),
            outline:
              prev.outline.length > 0 ? prev.outline : parsedPptx.outline,
          };
        });
      }
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể tải tệp gốc từ đám mây.';
      setUploadNotice(message);
    } finally {
      if (hydratingPdfDocIdRef.current === doc.id) {
        hydratingPdfDocIdRef.current = null;
      }
      setHydratingPdfDocId((prev) => (prev === doc.id ? null : prev));
    }
  }, []);

  const handleUpdateSettings = useCallback((patch: Partial<UserSettings>) => {
    setUserSettings((prev) => {
      const next: UserSettings = { ...prev, ...patch };
      persistUserSettings(next);
      return next;
    });
  }, []);

  const handleSelectNav = useCallback((nav: SidebarNavKey) => {
    if (nav === 'folders') {
      setIsLocationPanelOpen((prev) =>
        activeNavRef.current === 'folders' ? !prev : true
      );
      setActiveNav('folders');
      return;
    }

    setIsLocationPanelOpen(false);
    setActiveNav(nav);
    if (nav === 'home') {
      setActiveFilter('all');
      setActiveLocation('all');
    } else if (nav === 'shared') {
      setActiveFilter('all');
      setActiveLocation('shared');
    } else {
      setActiveLocation((prevLoc) => (prevLoc === 'shared' ? 'all' : prevLoc));
      if (nav === 'slide') {
        setActiveFilter('pptx');
      } else if (nav === 'word') {
        setActiveFilter('docx');
      } else if (nav === 'excel') {
        setActiveFilter('xlsx');
      } else if (nav === 'pdf') {
        setActiveFilter('pdf');
      }
    }
  }, []);

  const handleFilterChange = useCallback((filter: FilterKey) => {
    setActiveFilter(filter);
    setActiveNav((prevNav) => {
      if (prevNav === 'folders' || prevNav === 'shared') {
        return prevNav;
      }
      if (filter === 'pptx') {
        return 'slide';
      }
      if (filter === 'docx') {
        return 'word';
      }
      if (filter === 'xlsx') {
        return 'excel';
      }
      if (filter === 'pdf') {
        return 'pdf';
      }
      return 'home';
    });
  }, []);

  const handleOpenDocument = useCallback(
    (doc: DocumentItem) => {
      setRemoteRoomCode(null);
      setSharedToken(null);
      setActiveWorkspaceDocId(doc.id);
      setLastActiveDocId(doc.id);
      setActiveModal(null);
      setModalDocOverride(null);

      const targetHash = `#doc/${encodeURIComponent(doc.id)}`;
      if (window.location.hash !== targetHash) {
        window.location.hash = targetHash;
      }

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured && doc.isSyncedToCloud) {
        const accessedIso = new Date().toISOString();
        void touchDocumentLastAccessed(activeUser.id, doc.id).catch(() => {
          // Non-blocking telemetry timestamp update
        });
        setDocuments((prev) =>
          prev.map((item) =>
            item.id === doc.id
              ? { ...item, lastAccessedAt: accessedIso }
              : item
          )
        );
      }

      if (needsCloudFileHydration(doc)) {
        autoHydrateAttemptedRef.current.add(doc.id);
        void hydrateCloudPdfDocument(doc);
      }
    },
    [hydrateCloudPdfDocument]
  );

  const handleBackHome = useCallback(() => {
    setRemoteRoomCode(null);
    setSharedToken(null);
    setCollabShareQueryToken(null);
    setCollabSharedDoc(null);
    setCollabShareError(null);
    setIsResolvingCollabShare(false);
    setActiveWorkspaceDocId(null);
    setActiveModal(null);
    setModalDocOverride(null);

    const url = new URL(window.location.href);
    const hasShareParam = url.searchParams.has('share');
    if (hasShareParam || window.location.hash) {
      url.searchParams.delete('share');
      window.history.pushState(null, '', `${url.pathname}${url.search}`);
    }
  }, []);

  const enqueueCloudUpsert = useCallback(
    (
      userId: string,
      docToSync: DocumentItem,
      options?: {
        file?: File | null;
        touchUpdatedAt?: boolean;
        forceUploadFile?: boolean;
      }
    ) => {
      const nextSeq = (saveSeqByDocRef.current.get(docToSync.id) ?? 0) + 1;
      saveSeqByDocRef.current.set(docToSync.id, nextSeq);

      const previousTask =
        saveQueueByDocRef.current.get(docToSync.id) ?? Promise.resolve();

      const nextTask = previousTask
        .catch(() => {
          // Continue queue even if an earlier write failed
        })
        .then(async () => {
          if (currentUserRef.current?.id !== userId) {
            return;
          }
          setSyncStatus('syncing');
          try {
            const synced = await upsertCloudDocument(
              userId,
              docToSync,
              options
            );
            if (currentUserRef.current?.id !== userId) {
              return;
            }
            const isLatestForDoc =
              saveSeqByDocRef.current.get(docToSync.id) === nextSeq;
            const nextDocs = documentsRef.current.map((item) => {
              if (item.id !== docToSync.id) {
                return item;
              }
              return {
                ...item,
                storagePath: synced.storagePath ?? item.storagePath,
                r2ObjectKey: synced.r2ObjectKey ?? item.r2ObjectKey,
                originalFileName:
                  synced.originalFileName ?? item.originalFileName,
                isSyncedToCloud: true,
                userId,
                ...(isLatestForDoc
                  ? {
                      lastSavedAt: synced.lastSavedAt ?? item.lastSavedAt,
                      lastAccessedAt:
                        synced.lastAccessedAt ?? item.lastAccessedAt,
                    }
                  : {}),
              };
            });
            setDocuments(nextDocs);
            persistDocumentsToStorage(nextDocs);
            setLastSyncedAt(new Date().toISOString());
            setSyncStatus('synced');
          } catch {
            if (currentUserRef.current?.id === userId) {
              setSyncStatus('error');
            }
          }
        });

      saveQueueByDocRef.current.set(docToSync.id, nextTask);
      return nextTask;
    },
    []
  );

  const handleToggleFavorite = useCallback(
    (docId: string) => {
      const existingDoc = documentsRef.current.find((doc) => doc.id === docId);
      if (!existingDoc) {
        return;
      }

      const toggledDoc: DocumentItem = {
        ...existingDoc,
        isFavorite: !existingDoc.isFavorite,
      };
      const nextDocs = documentsRef.current.map((doc) =>
        doc.id === docId ? toggledDoc : doc
      );
      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured && toggledDoc.isSyncedToCloud) {
        void enqueueCloudUpsert(activeUser.id, toggledDoc, {
          file: null,
          touchUpdatedAt: false,
        });
      }
    },
    [enqueueCloudUpsert]
  );

  const handleSaveDocument = useCallback(
    (updatedDoc: DocumentItem) => {
      if (activeCollabPermissionRef.current === 'view') {
        return;
      }

      const previousDoc =
        documentsRef.current.find((doc) => doc.id === updatedDoc.id) ??
        (collabSharedDocRef.current?.id === updatedDoc.id
          ? collabSharedDocRef.current
          : undefined);

      const savedDoc: DocumentItem = {
        ...updatedDoc,
        updatedAt: 'Vừa lưu',
        lastSavedAt: new Date().toISOString(),
      };

      if (previousDoc && previousDoc.title !== savedDoc.title) {
        realtimeRoomRef.current?.broadcastPatch({
          kind: 'title_patch',
          title: savedDoc.title,
        });
      }

      const existsInDocs = documentsRef.current.some(
        (doc) => doc.id === savedDoc.id
      );
      if (existsInDocs) {
        const nextDocs = documentsRef.current.map((doc) =>
          doc.id === savedDoc.id ? savedDoc : doc
        );
        setDocuments(nextDocs);

        const storageResult = persistDocumentsToStorage(nextDocs);
        if (storageResult.quotaExceeded) {
          setUploadNotice(
            `Đã lưu tạm "${savedDoc.title}" trong phiên hiện tại, nhưng bộ nhớ trình duyệt đã đầy nên không thể lưu vĩnh viễn toàn bộ dữ liệu lớn.`
          );
        }
      }

      if (collabSharedDocRef.current?.id === savedDoc.id) {
        setCollabSharedDoc(savedDoc);
      }

      realtimeRoomRef.current?.schedulePersist(savedDoc);

      const activeUser = currentUserRef.current;
      if (
        activeUser &&
        isSupabaseConfigured &&
        activeCollabPermissionRef.current === 'owner'
      ) {
        void enqueueCloudUpsert(activeUser.id, savedDoc, {
          file: savedDoc.sourceFile ?? null,
          touchUpdatedAt: true,
        });
      }
    },
    [enqueueCloudUpsert]
  );

  const handleShareStateChanged = useCallback(
    (info: {
      documentId: string;
      permission: CollabPermission;
      shareToken?: string;
      isShared: boolean;
    }) => {
      const shareAccess: ShareAccessLevel =
        info.permission === 'private' ? 'private' : 'link';

      const nextDocs = documentsRef.current.map((doc) =>
        doc.id === info.documentId
          ? {
              ...doc,
              isShared: info.isShared,
              shareToken: info.shareToken ?? doc.shareToken,
              shareAccess,
            }
          : doc
      );
      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      setCollabSharedDoc((prev) =>
        prev && prev.id === info.documentId
          ? {
              ...prev,
              isShared: info.isShared,
              shareToken: info.shareToken ?? prev.shareToken,
              shareAccess,
            }
          : prev
      );
    },
    []
  );

  const handleShareConfigUpdated = useCallback(
    (
      docId: string,
      config: {
        shareToken?: string;
        shareAccess: ShareAccessLevel;
        shareAllowDownload: boolean;
        isShared: boolean;
      }
    ) => {
      const existingDoc = documentsRef.current.find((doc) => doc.id === docId);
      if (!existingDoc) {
        return;
      }

      const updatedDoc: DocumentItem = {
        ...existingDoc,
        shareToken: config.shareToken ?? existingDoc.shareToken,
        shareAccess: config.shareAccess,
        shareAllowDownload: config.shareAllowDownload,
        isShared: config.isShared,
      };

      const nextDocs = documentsRef.current.map((doc) =>
        doc.id === docId ? updatedDoc : doc
      );

      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured) {
        void enqueueCloudUpsert(activeUser.id, updatedDoc, {
          file: null,
          touchUpdatedAt: false,
        });
      }
    },
    [enqueueCloudUpsert]
  );

  const handleReplacePdfFile = useCallback(
    (docId: string, file: File) => {
      const existingDoc = documentsRef.current.find((doc) => doc.id === docId);
      if (!existingDoc) {
        return;
      }

      const newUrl = URL.createObjectURL(file);
      createdObjectUrlsRef.current.push(newUrl);
      const fileSize = formatFileSize(file.size);

      const updatedPdfDoc: DocumentItem = {
        ...existingDoc,
        fileUrl: newUrl,
        objectUrl: newUrl,
        sourceFile: file,
        originalFileName: file.name,
        fileSize,
        updatedAt: 'Vừa nạp lại',
      };

      const nextDocs = documentsRef.current.map((doc) =>
        doc.id === docId ? updatedPdfDoc : doc
      );

      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured) {
        if (isR2ApiConfigured && supabase) {
          void (async () => {
            try {
              const sessionRes = await supabase.auth.getSession();
              const accessToken = sessionRes.data.session?.access_token ?? '';
              if (accessToken) {
                const uploaded = await uploadFileViaR2PresignedUrl({
                  file,
                  documentId: docId,
                  accessToken,
                });
                const docWithR2: DocumentItem = {
                  ...updatedPdfDoc,
                  r2ObjectKey: uploaded.objectKey,
                };
                void enqueueCloudUpsert(activeUser.id, docWithR2, {
                  file: null,
                  touchUpdatedAt: true,
                });
                return;
              }
            } catch {
              // Fall back to Supabase Storage below
            }
            void enqueueCloudUpsert(activeUser.id, updatedPdfDoc, {
              file,
              touchUpdatedAt: true,
              forceUploadFile: true,
            });
          })();
        } else {
          void enqueueCloudUpsert(activeUser.id, updatedPdfDoc, {
            file,
            touchUpdatedAt: true,
            forceUploadFile: true,
          });
        }
      }
    },
    [enqueueCloudUpsert]
  );

  const handleDeleteDocuments = useCallback(
    async (docsToDelete: DocumentItem[]) => {
      if (docsToDelete.length === 0) {
        return;
      }

      setIsDeleting(true);
      try {
        const deletedIds = docsToDelete.map((doc) => doc.id);
        const deletedIdSet = new Set<string>(deletedIds);

        // Revoke any in-memory blob URLs associated with the deleted documents
        const revokedUrls = new Set<string>();
        docsToDelete.forEach((doc) => {
          if (doc.fileUrl?.startsWith('blob:')) {
            revokedUrls.add(doc.fileUrl);
          }
          if (doc.objectUrl?.startsWith('blob:')) {
            revokedUrls.add(doc.objectUrl);
          }
          doc.slidesPreview.forEach((slide) => {
            if (slide.backgroundImageSrc?.startsWith('blob:')) {
              revokedUrls.add(slide.backgroundImageSrc);
            }
            slide.images?.forEach((img) => {
              if (img.src.startsWith('blob:')) {
                revokedUrls.add(img.src);
              }
            });
          });
        });
        revokedUrls.forEach((url) => {
          URL.revokeObjectURL(url);
        });
        if (revokedUrls.size > 0) {
          createdObjectUrlsRef.current = createdObjectUrlsRef.current.filter(
            (url) => !revokedUrls.has(url)
          );
        }

        deletedIds.forEach((id) => {
          saveQueueByDocRef.current.delete(id);
          saveSeqByDocRef.current.delete(id);
          autoHydrateAttemptedRef.current.delete(id);
        });

        const nextDocs = documentsRef.current.filter(
          (doc) => !deletedIdSet.has(doc.id)
        );
        setDocuments(nextDocs);
        persistDocumentsToStorage(nextDocs);
        setPendingMigrationDocs((prev) =>
          prev.filter((doc) => !deletedIdSet.has(doc.id))
        );
        removeLocalSharesForDocumentIds(deletedIds);

        setActiveWorkspaceDocId((prevId) => {
          if (prevId && deletedIdSet.has(prevId)) {
            if (parseDocIdFromHash(window.location.hash)) {
              window.history.pushState(
                null,
                '',
                `${window.location.pathname}${window.location.search}`
              );
            }
            return null;
          }
          return prevId;
        });

        setLastActiveDocId((prevId) =>
          prevId && !deletedIdSet.has(prevId)
            ? prevId
            : (nextDocs[0]?.id ?? null)
        );

        setModalDocOverride((prevOverride) => {
          if (prevOverride && deletedIdSet.has(prevOverride.id)) {
            setActiveModal(null);
            return null;
          }
          return prevOverride;
        });

        if (nextDocs.length === 0) {
          setActiveModal(null);
          setModalDocOverride(null);
        }

        const activeUser = currentUserRef.current;
        if (activeUser && isSupabaseConfigured) {
          let accessToken: string | null = null;
          if (isR2ApiConfigured && supabase) {
            try {
              const sessionRes = await supabase.auth.getSession();
              accessToken = sessionRes.data.session?.access_token ?? null;
            } catch {
              accessToken = null;
            }
          }

          setSyncStatus('syncing');
          const results = await Promise.allSettled(
            docsToDelete.map((doc) =>
              deleteCloudDocument(activeUser.id, doc, { accessToken })
            )
          );
          if (currentUserRef.current?.id === activeUser.id) {
            const hadCloudFailure = results.some(
              (res, idx) =>
                res.status === 'rejected' && docsToDelete[idx].isSyncedToCloud
            );
            setLastSyncedAt(new Date().toISOString());
            setSyncStatus(hadCloudFailure ? 'error' : 'synced');
          }
        }

        if (docsToDelete.length === 1) {
          setUploadNotice(
            `Đã xóa tài liệu "${docsToDelete[0].title}" khỏi thư viện.`
          );
        } else {
          setUploadNotice(
            `Đã xóa ${docsToDelete.length} tài liệu khỏi thư viện.`
          );
        }
      } finally {
        setIsDeleting(false);
      }
    },
    []
  );

  const handleCreateDocument = useCallback(
    (params: CreateEditableDocumentParams) => {
      const newDoc = createNewEditableDocument(params);
      const nextDocs = [newDoc, ...documentsRef.current];
      documentsRef.current = nextDocs;
      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      setRemoteRoomCode(null);
      setSharedToken(null);
      setLastActiveDocId(newDoc.id);
      setActiveWorkspaceDocId(newDoc.id);
      setActiveModal(null);
      setModalDocOverride(null);

      const targetHash = `#doc/${encodeURIComponent(newDoc.id)}`;
      if (window.location.hash !== targetHash) {
        window.location.hash = targetHash;
      }

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured) {
        void enqueueCloudUpsert(activeUser.id, newDoc, {
          file: null,
          touchUpdatedAt: true,
        });
      }
    },
    [enqueueCloudUpsert]
  );

  const handleUploadFiles = useCallback(async (files: FileList | File[]) => {
    const fileArray = Array.from(files);
    if (fileArray.length === 0) {
      return;
    }

    setIsUploading(true);
    setUploadProgress(null);
    const acceptedDocs: DocumentItem[] = [];
    const errorMessages: string[] = [];

    try {
      for (const file of fileArray) {
        const result = await parseUploadedFile(file);
        if (result.createdObjectUrls.length > 0) {
          createdObjectUrlsRef.current.push(...result.createdObjectUrls);
        }
        if (result.document) {
          acceptedDocs.push({
            ...result.document,
            originalFileName: result.document.originalFileName ?? file.name,
          });
        } else if (result.error) {
          errorMessages.push(result.error);
        }
      }

      let quotaNote = '';
      if (acceptedDocs.length > 0) {
        const primaryDoc = acceptedDocs[0];
        const nextDocs = [...acceptedDocs, ...documentsRef.current];
        setDocuments(nextDocs);
        setLastActiveDocId(primaryDoc.id);

        const storageResult = persistDocumentsToStorage(nextDocs);
        if (storageResult.quotaExceeded) {
          quotaNote =
            ' (Bộ nhớ trình duyệt đã đầy — tệp vẫn sẵn sàng dùng trong phiên hiện tại.)';
        }

        // Keep user on Home Hub and ensure the newly uploaded card is immediately visible
        setSearchQuery('');
        setActiveNav((prevNav) => {
          if (prevNav === 'folders') {
            return prevNav;
          }
          if (prevNav === 'slide' && primaryDoc.type === 'pptx') {
            return prevNav;
          }
          if (prevNav === 'word' && primaryDoc.type === 'docx') {
            return prevNav;
          }
          if (prevNav === 'excel' && primaryDoc.type === 'xlsx') {
            return prevNav;
          }
          if (prevNav === 'pdf' && primaryDoc.type === 'pdf') {
            return prevNav;
          }
          return 'home';
        });
        setActiveFilter((prevFilter) => {
          if (prevFilter === 'all' || prevFilter === primaryDoc.type) {
            return prevFilter;
          }
          return 'all';
        });
        setActiveLocation((prevLoc) =>
          prevLoc === 'shared' || prevLoc === 'cloud' ? 'all' : prevLoc
        );
      }

      // Cloud sync immediately after local parsing when signed in (R2 Presigned URL -> Supabase fallback)
      const activeUser = currentUserRef.current;
      let cloudSyncSummary = '';
      if (acceptedDocs.length > 0 && activeUser && isSupabaseConfigured) {
        setSyncStatus('syncing');
        const syncedById = new Map<string, DocumentItem>();
        let failedCloudCount = 0;

        let accessToken = '';
        if (isR2ApiConfigured && supabase) {
          const sessionRes = await supabase.auth.getSession();
          accessToken = sessionRes.data.session?.access_token ?? '';
        }

        for (let idx = 0; idx < acceptedDocs.length; idx += 1) {
          const parsedDoc = acceptedDocs[idx];
          let docForCloud: DocumentItem = parsedDoc;
          let uploadedToR2 = false;

          if (isR2ApiConfigured && accessToken && parsedDoc.sourceFile) {
            try {
              const r2Result = await uploadFileViaR2PresignedUrl({
                file: parsedDoc.sourceFile,
                documentId: parsedDoc.id,
                accessToken,
                onProgress: (percent) => {
                  const overall = Math.min(
                    100,
                    Math.round(
                      ((idx + percent / 100) / acceptedDocs.length) * 100
                    )
                  );
                  setUploadProgress(overall);
                },
              });
              docForCloud = {
                ...parsedDoc,
                r2ObjectKey: r2Result.objectKey,
              };
              uploadedToR2 = true;
            } catch {
              // Fall back seamlessly to Supabase user-files Storage
              setUploadProgress(null);
            }
          }

          try {
            const syncedDoc = await upsertCloudDocument(
              activeUser.id,
              docForCloud,
              {
                file: uploadedToR2 ? null : (parsedDoc.sourceFile ?? null),
                touchUpdatedAt: true,
              }
            );
            syncedById.set(parsedDoc.id, syncedDoc);
          } catch {
            failedCloudCount += 1;
          }
        }

        if (
          syncedById.size > 0 &&
          currentUserRef.current?.id === activeUser.id
        ) {
          const updatedDocs = documentsRef.current.map((doc) => {
            const synced = syncedById.get(doc.id);
            if (!synced) {
              return doc;
            }
            return {
              ...doc,
              storagePath: synced.storagePath ?? doc.storagePath,
              r2ObjectKey: synced.r2ObjectKey ?? doc.r2ObjectKey,
              originalFileName: synced.originalFileName ?? doc.originalFileName,
              lastSavedAt: synced.lastSavedAt ?? doc.lastSavedAt,
              lastAccessedAt: synced.lastAccessedAt ?? doc.lastAccessedAt,
              isSyncedToCloud: true,
              userId: activeUser.id,
            };
          });
          setDocuments(updatedDocs);
          persistDocumentsToStorage(updatedDocs);
        }

        if (currentUserRef.current?.id === activeUser.id) {
          if (failedCloudCount === 0) {
            setLastSyncedAt(new Date().toISOString());
            setSyncStatus('synced');
            cloudSyncSummary = ' Đã đồng bộ lên đám mây.';
          } else {
            setSyncStatus('error');
            cloudSyncSummary = ` (Chưa đồng bộ được ${failedCloudCount} tệp lên đám mây — vẫn lưu cục bộ.)`;
          }
        }
      }

      if (acceptedDocs.length > 0 && errorMessages.length === 0) {
        const notice =
          acceptedDocs.length === 1
            ? `Đã đọc tệp "${acceptedDocs[0].title}" (${acceptedDocs[0].type.toUpperCase()}). Bấm vào thẻ để mở và chỉnh sửa.${cloudSyncSummary}${quotaNote}`
            : `Đã đọc thành công ${acceptedDocs.length} tệp mới lên đầu thư viện.${cloudSyncSummary}${quotaNote}`;
        setUploadNotice(notice);
      } else if (acceptedDocs.length > 0 && errorMessages.length > 0) {
        setUploadNotice(
          `Đã đọc ${acceptedDocs.length} tệp hợp lệ.${cloudSyncSummary} Bỏ qua: ${errorMessages.join(' ')}${quotaNote}`
        );
      } else if (errorMessages.length > 0) {
        setUploadNotice(errorMessages.join(' '));
      }
    } finally {
      setIsUploading(false);
      setUploadProgress(null);
    }
  }, []);

  const handleConfirmMigrateLocal = useCallback(async () => {
    const activeUser = currentUserRef.current;
    if (
      !activeUser ||
      !isSupabaseConfigured ||
      pendingMigrationDocs.length === 0 ||
      isMigratingLocal
    ) {
      return;
    }

    setIsMigratingLocal(true);
    setSyncStatus('syncing');

    try {
      const validMigrationDocs = pendingMigrationDocs.filter(
        (doc) => !LEGACY_MOCK_ID_SET.has(doc.id)
      );
      const { syncedDocs, failedCount } = await syncLocalDocumentsToCloud(
        activeUser.id,
        validMigrationDocs
      );
      const syncedById = new Map<string, DocumentItem>(
        syncedDocs.map((doc) => [doc.id, doc])
      );

      const nextDocs = documentsRef.current
        .filter((doc) => !LEGACY_MOCK_ID_SET.has(doc.id))
        .map((doc) => {
          const synced = syncedById.get(doc.id);
          if (!synced) {
            return doc;
          }
          return preserveInMemoryRuntimeFields(synced, doc);
        });

      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);
      setPendingMigrationDocs([]);
      setActiveUtilityModal(null);
      setLastSyncedAt(new Date().toISOString());
      setSyncStatus(failedCount > 0 ? 'error' : 'synced');

      if (failedCount === 0) {
        setUploadNotice(
          `Đã đồng bộ thành công ${syncedDocs.length} tệp cục bộ lên tài khoản ${activeUser.email}.`
        );
      } else {
        setUploadNotice(
          `Đã đồng bộ ${syncedDocs.length} tệp lên đám mây. Có ${failedCount} tệp bị lỗi, bạn có thể thử đồng bộ thủ công trong Cài đặt.`
        );
      }
    } catch {
      setSyncStatus('error');
      setUploadNotice(
        'Không thể đồng bộ tệp cục bộ lên tài khoản lúc này. Vui lòng thử lại trong Cài đặt.'
      );
    } finally {
      setIsMigratingLocal(false);
    }
  }, [pendingMigrationDocs, isMigratingLocal]);

  const handleSkipMigrateLocal = useCallback(() => {
    const pendingIds = new Set<string>(pendingMigrationDocs.map((d) => d.id));
    const cloudDocs = documentsRef.current.filter(
      (doc) =>
        doc.isSyncedToCloud &&
        !pendingIds.has(doc.id) &&
        !LEGACY_MOCK_ID_SET.has(doc.id)
    );

    setDocuments(cloudDocs);
    persistDocumentsToStorage(cloudDocs);
    setPendingMigrationDocs([]);
    setActiveUtilityModal(null);
    setUploadNotice(
      'Đã giữ nguyên dữ liệu trên tài khoản đám mây (chưa tải tệp cục bộ lên).'
    );
  }, [pendingMigrationDocs]);

  const handleManualSync = useCallback(async () => {
    const activeUser = currentUserRef.current;
    if (!activeUser || !isSupabaseConfigured) {
      return;
    }

    setSyncStatus('syncing');
    try {
      const localCandidates = documentsRef.current.filter(
        isUnsyncedLocalCandidate
      );
      let failedCount = 0;
      if (localCandidates.length > 0) {
        const result = await syncLocalDocumentsToCloud(
          activeUser.id,
          localCandidates
        );
        failedCount = result.failedCount;
      }

      const [fetchedCloudDocs, sharedWithMeDocs] = await Promise.all([
        fetchUserDocuments(activeUser.id),
        loadSharedWithMeDocuments(activeUser.id).catch(
          () => [] as DocumentItem[]
        ),
      ]);
      if (currentUserRef.current?.id !== activeUser.id) {
        return;
      }

      const legacyCloudDocs = fetchedCloudDocs.filter((cloudDoc) =>
        LEGACY_MOCK_ID_SET.has(cloudDoc.id)
      );
      if (legacyCloudDocs.length > 0) {
        removeLocalSharesForDocumentIds(legacyCloudDocs.map((d) => d.id));
        legacyCloudDocs.forEach((legacyDoc) => {
          void deleteCloudDocument(activeUser.id, legacyDoc).catch(
            () => undefined
          );
        });
      }

      const existingById = new Map<string, DocumentItem>(
        documentsRef.current.map((doc) => [doc.id, doc])
      );
      const cloudDocs = fetchedCloudDocs
        .filter((cloudDoc) => !LEGACY_MOCK_ID_SET.has(cloudDoc.id))
        .map((cloudDoc) =>
          preserveInMemoryRuntimeFields(
            cloudDoc,
            existingById.get(cloudDoc.id)
          )
        );
      const cloudIds = new Set<string>(cloudDocs.map((doc) => doc.id));

      const sharedDocs = sharedWithMeDocs
        .filter(
          (sharedDoc) =>
            !LEGACY_MOCK_ID_SET.has(sharedDoc.id) && !cloudIds.has(sharedDoc.id)
        )
        .map((sharedDoc) =>
          preserveInMemoryRuntimeFields(
            { ...sharedDoc, isShared: true },
            existingById.get(sharedDoc.id)
          )
        );
      sharedDocs.forEach((doc) => {
        cloudIds.add(doc.id);
      });

      const remainingLocal =
        failedCount > 0
          ? localCandidates.filter((doc) => !cloudIds.has(doc.id))
          : [];

      const mergedDocs = [...cloudDocs, ...sharedDocs, ...remainingLocal];
      setDocuments(mergedDocs);
      persistDocumentsToStorage(mergedDocs);
      setPendingMigrationDocs([]);
      setLastSyncedAt(new Date().toISOString());
      setSyncStatus(failedCount > 0 ? 'error' : 'synced');

      if (failedCount === 0) {
        setUploadNotice(
          `Đã đồng bộ hoàn tất ${cloudDocs.length} tài liệu với tài khoản đám mây.`
        );
      } else {
        setUploadNotice(
          `Đã tải ${cloudDocs.length} tài liệu đám mây, còn ${failedCount} tệp cục bộ gặp lỗi khi tải lên.`
        );
      }
    } catch (error) {
      setSyncStatus('error');
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể đồng bộ thủ công lúc này.';
      setUploadNotice(message);
    }
  }, []);

  const handleConfirmLogout = useCallback(async () => {
    if (isLoggingOut) {
      return;
    }

    setIsLoggingOut(true);
    try {
      if (supabase) {
        await supabase.auth.signOut();
      }
    } catch {
      // Proceed to clear local session state even if network signOut fails
    } finally {
      clearStoredGoogleDriveToken();
      setHasGoogleToken(false);
      currentUserRef.current = null;
      lastLoadedUserIdRef.current = null;
      inFlightSyncUserIdRef.current = null;
      autoHydrateAttemptedRef.current.clear();

      // Strip cloud-owned documents from localStorage before returning to Guest Local Mode
      const pureLocalDocs = documentsRef.current.filter(
        (doc) =>
          !doc.isSyncedToCloud &&
          !doc.userId &&
          !LEGACY_MOCK_ID_SET.has(doc.id)
      );
      persistDocumentsToStorage(pureLocalDocs);

      const localDocs = loadInitialDocuments();
      setDocuments(localDocs);
      setCurrentUser(null);
      setPendingMigrationDocs([]);
      setLastSyncedAt(null);
      setSyncStatus(isSupabaseConfigured ? 'idle' : 'local');
      setActiveWorkspaceDocId(null);
      setLastActiveDocId(localDocs[0]?.id ?? null);
      setActiveModal(null);
      setModalDocOverride(null);
      setActiveUtilityModal(null);

      if (window.location.hash) {
        window.history.pushState(
          null,
          '',
          `${window.location.pathname}${window.location.search}`
        );
      }

      setUploadNotice(
        'Đã đăng xuất tài khoản. Ứng dụng đã chuyển về chế độ cục bộ (Local Mode).'
      );
      setIsLoggingOut(false);
    }
  }, [isLoggingOut]);

  const handleOpenModal = useCallback(
    (modal: Exclude<ModalType, null>, overrideDoc?: DocumentItem) => {
      setModalDocOverride(overrideDoc ?? null);
      setActiveModal(modal);
    },
    []
  );

  const handleShareDocument = useCallback(
    (doc: DocumentItem) => {
      handleOpenModal('share', doc);
    },
    [handleOpenModal]
  );

  const handleSwitchModal = useCallback((modal: Exclude<ModalType, null>) => {
    setActiveModal(modal);
  }, []);

  const handleCloseModal = useCallback(() => {
    setActiveModal(null);
    setModalDocOverride(null);
  }, []);

  const handleDismissNotice = useCallback(() => {
    setUploadNotice(null);
  }, []);

  const handleOpenSettings = useCallback(() => {
    setActiveUtilityModal('settings');
  }, []);

  const handleOpenHelp = useCallback(() => {
    setActiveUtilityModal('help');
  }, []);

  const handleOpenAuthModal = useCallback(() => {
    setActiveUtilityModal('auth');
  }, []);

  const handleRequestLogout = useCallback(() => {
    setActiveUtilityModal('logoutConfirm');
  }, []);

  const handleAuthAction = useCallback(() => {
    if (currentUserRef.current) {
      setActiveUtilityModal('logoutConfirm');
    } else {
      setActiveUtilityModal('auth');
    }
  }, []);

  const handleAuthSuccess = useCallback(() => {
    setActiveUtilityModal((prev) => (prev === 'auth' ? null : prev));
  }, []);

  const handleCloseUtilityModal = useCallback(() => {
    setActiveUtilityModal(null);
  }, []);

  const handleCloseCloudActionModal = useCallback(() => {
    if (isMigratingLocal || isLoggingOut) {
      return;
    }
    setActiveUtilityModal(null);
  }, [isMigratingLocal, isLoggingOut]);

  const handleMobileLogoClick = useCallback(() => {
    handleSelectNav('home');
  }, [handleSelectNav]);

  const handleConnectGoogleDrive = useCallback(async () => {
    if (!isSupabaseConfigured || !supabase) {
      setActiveUtilityModal('auth');
      return;
    }

    const oauthOptions = {
      redirectTo: window.location.origin,
      scopes:
        'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly',
      queryParams: {
        access_type: 'offline',
        prompt: 'consent',
      },
    };

    try {
      if (currentUserRef.current) {
        const { error: linkError } = await supabase.auth.linkIdentity({
          provider: 'google',
          options: oauthOptions,
        });
        if (!linkError) {
          return;
        }
      }

      const { error: signInError } = await supabase.auth.signInWithOAuth({
        provider: 'google',
        options: oauthOptions,
      });
      if (signInError) {
        setUploadNotice(
          `Không thể kết nối Google Drive: ${signInError.message}`
        );
      }
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể khởi tạo kết nối Google Drive.';
      setUploadNotice(message);
    }
  }, []);

  const handleOpenDriveFile = useCallback(
    async (file: File, driveItem: GoogleDriveFileItem) => {
      const result = await parseUploadedFile(file);
      if (result.createdObjectUrls.length > 0) {
        createdObjectUrlsRef.current.push(...result.createdObjectUrls);
      }

      if (!result.document || result.error) {
        setUploadNotice(
          result.error ??
            `Không thể đọc tệp "${driveItem.name}" từ Google Drive.`
        );
        setIsDriveModalOpen(false);
        return;
      }

      const parsedDoc: DocumentItem = {
        ...result.document,
        folder: 'Google Drive',
        originalFileName:
          result.document.originalFileName ?? driveItem.name ?? file.name,
      };

      const nextDocs = [parsedDoc, ...documentsRef.current];
      documentsRef.current = nextDocs;
      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured) {
        void enqueueCloudUpsert(activeUser.id, parsedDoc, {
          file: parsedDoc.sourceFile ?? null,
          touchUpdatedAt: true,
        });
      }

      setIsDriveModalOpen(false);
      handleOpenDocument(parsedDoc);
    },
    [enqueueCloudUpsert, handleOpenDocument]
  );

  const locationCounts = useMemo(
    () => ({
      recent: documents.filter((d) =>
        Boolean(d.lastAccessedAt || d.lastSavedAt || d.updatedAt)
      ).length,
      shared: documents.filter((d) => Boolean(d.isShared)).length,
      cloud: documents.filter((d) => Boolean(d.isSyncedToCloud)).length,
      device: documents.filter(
        (d) => !d.isSyncedToCloud || Boolean(d.isUploaded)
      ).length,
    }),
    [documents]
  );

  const filteredDocuments = useMemo(() => {
    const query = searchQuery.trim().toLowerCase();

    const matched = documents.filter((doc) => {
      if ((activeNav === 'shared' || activeLocation === 'shared') && !doc.isShared) {
        return false;
      }

      if (activeLocation === 'cloud' && !doc.isSyncedToCloud) {
        return false;
      }

      if (
        activeLocation === 'device' &&
        doc.isSyncedToCloud &&
        !doc.isUploaded
      ) {
        return false;
      }

      if (activeFilter === 'favorite' && !doc.isFavorite) {
        return false;
      }

      if (
        (activeFilter === 'pptx' ||
          activeFilter === 'docx' ||
          activeFilter === 'xlsx' ||
          activeFilter === 'pdf') &&
        doc.type !== activeFilter
      ) {
        return false;
      }

      if (query.length > 0) {
        const searchable = [
          doc.title,
          doc.type,
          doc.contentPreview,
          doc.folder ?? '',
          ...doc.outline.map((sec) => `${sec.romanIndex}. ${sec.title}`),
          ...(doc.sheets?.map((sheet) => sheet.name) ?? []),
        ]
          .join(' ')
          .toLowerCase();

        if (!searchable.includes(query)) {
          return false;
        }
      }

      return true;
    });

    if (activeLocation === 'recent') {
      return [...matched].sort((a, b) => {
        const timeA = Math.max(
          Date.parse(a.lastAccessedAt ?? '') || 0,
          Date.parse(a.lastSavedAt ?? '') || 0
        );
        const timeB = Math.max(
          Date.parse(b.lastAccessedAt ?? '') || 0,
          Date.parse(b.lastSavedAt ?? '') || 0
        );
        return timeB - timeA;
      });
    }

    if (activeNav === 'folders') {
      return [...matched].sort((a, b) =>
        (a.folder ?? 'Chưa phân loại').localeCompare(
          b.folder ?? 'Chưa phân loại',
          'vi'
        )
      );
    }

    return matched;
  }, [documents, activeNav, activeLocation, activeFilter, searchQuery]);

  const activeWorkspaceDoc = useMemo(() => {
    if (collabSharedDoc && collabSharedDoc.id === activeWorkspaceDocId) {
      return collabSharedDoc;
    }
    if (!activeWorkspaceDocId) {
      return null;
    }
    return documents.find((doc) => doc.id === activeWorkspaceDocId) ?? null;
  }, [collabSharedDoc, documents, activeWorkspaceDocId]);

  const activeCollabPermission: 'owner' | 'edit' | 'view' =
    activeWorkspaceDoc?.collabPermission ?? 'owner';
  activeCollabPermissionRef.current = activeCollabPermission;

  const activeWorkspaceDocIdForRoom = activeWorkspaceDoc?.id ?? null;
  const activeWorkspaceOwnerIdForRoom =
    activeWorkspaceDoc?.sharedOwnerId ||
    activeWorkspaceDoc?.userId ||
    currentUser?.id ||
    'local-owner';
  const activeWorkspaceShareTokenForRoom = activeWorkspaceDoc?.shareToken;

  // Join Supabase Realtime Collaboration Room (Presence + Broadcast + Postgres Changes)
  useEffect(() => {
    realtimeRoomRef.current?.leave();
    realtimeRoomRef.current = null;
    setCollabCollaborators([]);
    setLatestRemotePatch(null);

    if (!activeWorkspaceDocIdForRoom) {
      return;
    }

    const displayName =
      currentUser?.displayName ||
      (currentUser?.email ? currentUser.email.split('@')[0] : '') ||
      `Khách ${selfSessionId.slice(-4).toUpperCase()}`;

    const roomHandle = joinDocumentRealtimeRoom({
      ownerId: activeWorkspaceOwnerIdForRoom,
      documentId: activeWorkspaceDocIdForRoom,
      shareToken: activeWorkspaceShareTokenForRoom,
      canEdit: activeCollabPermission !== 'view',
      identity: {
        sessionId: selfSessionId,
        userId: currentUser?.id,
        displayName,
        email: currentUser?.email,
      },
      callbacks: {
        onPresenceChange: (list) => {
          setCollabCollaborators(list);
        },
        onRemotePatch: (patch) => {
          setLatestRemotePatch(patch);
          if (patch.kind === 'title_patch') {
            setDocuments((prev) =>
              prev.map((doc) =>
                doc.id === activeWorkspaceDocIdForRoom
                  ? { ...doc, title: patch.title }
                  : doc
              )
            );
            setCollabSharedDoc((prev) =>
              prev && prev.id === activeWorkspaceDocIdForRoom
                ? { ...prev, title: patch.title }
                : prev
            );
          }
        },
        onDatabaseDocumentUpdate: (updatedDocPartial) => {
          setDocuments((prev) =>
            prev.map((doc) =>
              doc.id === activeWorkspaceDocIdForRoom
                ? { ...doc, ...updatedDocPartial }
                : doc
            )
          );
          setCollabSharedDoc((prev) =>
            prev && prev.id === activeWorkspaceDocIdForRoom
              ? { ...prev, ...updatedDocPartial }
              : prev
          );
        },
      },
    });

    realtimeRoomRef.current = roomHandle;

    return () => {
      roomHandle.leave();
      if (realtimeRoomRef.current === roomHandle) {
        realtimeRoomRef.current = null;
      }
    };
  }, [
    activeWorkspaceDocIdForRoom,
    activeWorkspaceOwnerIdForRoom,
    activeWorkspaceShareTokenForRoom,
    activeCollabPermission,
    currentUser?.id,
    currentUser?.displayName,
    currentUser?.email,
    selfSessionId,
  ]);

  const handleCollabPresenceUpdate = useCallback(
    (patch: Partial<Omit<CollaboratorPresence, 'sessionId' | 'color'>>) => {
      realtimeRoomRef.current?.updatePresence(patch);
    },
    []
  );

  const handleCollabBroadcastPatch = useCallback(
    (patch: Omit<DocRealtimePatch, 'senderSessionId' | 'sentAt'>) => {
      if (activeCollabPermissionRef.current === 'view') {
        return;
      }
      realtimeRoomRef.current?.broadcastPatch(patch);
    },
    []
  );

  // Auto-hydrate cloud file (PDF / DOCX / PPTX) once if opened directly via URL hash after reload
  useEffect(() => {
    if (
      !activeWorkspaceDoc ||
      !needsCloudFileHydration(activeWorkspaceDoc) ||
      autoHydrateAttemptedRef.current.has(activeWorkspaceDoc.id)
    ) {
      return;
    }
    autoHydrateAttemptedRef.current.add(activeWorkspaceDoc.id);
    void hydrateCloudPdfDocument(activeWorkspaceDoc);
  }, [activeWorkspaceDoc, hydrateCloudPdfDocument]);

  const modalDocument = useMemo(() => {
    if (modalDocOverride) {
      return modalDocOverride;
    }
    if (activeWorkspaceDoc) {
      return activeWorkspaceDoc;
    }
    if (lastActiveDocId) {
      return (
        documents.find((doc) => doc.id === lastActiveDocId) ??
        documents[0] ??
        null
      );
    }
    return documents[0] ?? null;
  }, [modalDocOverride, activeWorkspaceDoc, documents, lastActiveDocId]);

  const cloudDocumentCount = useMemo(
    () => documents.filter((doc) => Boolean(doc.isSyncedToCloud)).length,
    [documents]
  );

  const isAnyModalOpen =
    activeModal !== null || activeUtilityModal !== null || isDriveModalOpen;

  const cloudActionModalMode: 'logoutConfirm' | 'migrateLocal' | null =
    activeUtilityModal === 'logoutConfirm' ||
    activeUtilityModal === 'migrateLocal'
      ? activeUtilityModal
      : null;

  // Priority Route 1: Mobile Realtime Remote Controller (#remote/<roomCode>)
  if (remoteRoomCode) {
    return (
      <MobileRemoteView
        roomCode={remoteRoomCode}
        onExitToHome={handleBackHome}
      />
    );
  }

  // Priority Route 2: Public / Password-Protected Share Viewer (#share/<token>)
  if (sharedToken) {
    return (
      <SharedDocumentView
        shareToken={sharedToken}
        onNavigateHome={handleBackHome}
      />
    );
  }

  // Priority Route 3: ?share=<token> Realtime Collaboration Link Loading / Error View
  if (collabShareQueryToken && !activeWorkspaceDoc) {
    if (isResolvingCollabShare) {
      return (
        <div className="min-h-[100dvh] w-full flex items-center justify-center bg-[#FAF8F5] px-4">
          <div className="w-full max-w-md rounded-2xl border border-[#EAE6DF] bg-white p-7 text-center shadow-sm">
            <div className="mx-auto mb-4 h-10 w-10 animate-spin rounded-full border-2 border-indigo-200 border-t-indigo-600" />
            <h2 className="text-base font-semibold text-slate-900">
              Đang mở tài liệu được chia sẻ...
            </h2>
            <p className="mt-1.5 text-sm text-slate-500">
              Đang đồng bộ nội dung và kết nối phòng cộng tác thời gian thực.
            </p>
          </div>
        </div>
      );
    }

    if (collabShareError) {
      return (
        <div className="min-h-[100dvh] w-full flex items-center justify-center bg-[#FAF8F5] px-4">
          <div className="w-full max-w-md rounded-2xl border border-[#EAE6DF] bg-white p-7 text-center shadow-sm">
            <div className="mx-auto mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-rose-50 text-rose-600 border border-rose-100 text-lg font-bold">
              !
            </div>
            <h2 className="text-base font-semibold text-slate-900">
              Không thể truy cập liên kết chia sẻ
            </h2>
            <p className="mt-2 text-sm text-slate-600">{collabShareError}</p>
            <button
              type="button"
              onClick={handleBackHome}
              className="mt-5 inline-flex items-center justify-center rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-indigo-700 transition-colors"
            >
              Quay về Trang chủ LITE
            </button>
          </div>
        </div>
      );
    }
  }

  return (
    <>
      {activeWorkspaceDoc ? (
        activeWorkspaceDoc.type === 'pptx' ? (
          <SlideWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            autoSaveEnabled={userSettings.autoSyncOnEdit}
            isHydratingCloudFile={hydratingPdfDocId === activeWorkspaceDoc.id}
            collaborators={collabCollaborators}
            selfSessionId={selfSessionId}
            collabPermission={activeCollabPermission}
            onPresenceUpdate={handleCollabPresenceUpdate}
            onBroadcastPatch={handleCollabBroadcastPatch}
            incomingRemotePatch={latestRemotePatch}
            onRetryCloudDownload={() => {
              void hydrateCloudPdfDocument(activeWorkspaceDoc);
            }}
            onBackHome={handleBackHome}
            onSaveDocument={handleSaveDocument}
            onOpenModal={handleOpenModal}
          />
        ) : activeWorkspaceDoc.type === 'docx' ? (
          <WordWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            defaultPaperTheme={userSettings.defaultWordTheme}
            autoSaveEnabled={userSettings.autoSyncOnEdit}
            isHydratingCloudFile={hydratingPdfDocId === activeWorkspaceDoc.id}
            collaborators={collabCollaborators}
            selfSessionId={selfSessionId}
            collabPermission={activeCollabPermission}
            onPresenceUpdate={handleCollabPresenceUpdate}
            onBroadcastPatch={handleCollabBroadcastPatch}
            incomingRemotePatch={latestRemotePatch}
            onRetryCloudDownload={() => {
              void hydrateCloudPdfDocument(activeWorkspaceDoc);
            }}
            onBackHome={handleBackHome}
            onSaveDocument={handleSaveDocument}
            onOpenModal={handleOpenModal}
          />
        ) : activeWorkspaceDoc.type === 'xlsx' ? (
          <ExcelWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            autoSaveEnabled={userSettings.autoSyncOnEdit}
            collaborators={collabCollaborators}
            selfSessionId={selfSessionId}
            collabPermission={activeCollabPermission}
            onPresenceUpdate={handleCollabPresenceUpdate}
            onBroadcastPatch={handleCollabBroadcastPatch}
            incomingRemotePatch={latestRemotePatch}
            onBackHome={handleBackHome}
            onSaveDocument={handleSaveDocument}
            onOpenModal={handleOpenModal}
          />
        ) : (
          <PdfWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            isHydratingCloudFile={hydratingPdfDocId === activeWorkspaceDoc.id}
            onRetryCloudDownload={() => {
              void hydrateCloudPdfDocument(activeWorkspaceDoc);
            }}
            onBackHome={handleBackHome}
            onOpenModal={handleOpenModal}
            onReplacePdfFile={handleReplacePdfFile}
          />
        )
      ) : (
        <div className="h-[100dvh] w-full overflow-hidden flex bg-[#FAF8F5] text-slate-900">
          <SidebarRail
            activeNav={activeNav}
            onSelectNav={handleSelectNav}
            currentUser={currentUser}
            syncStatus={syncStatus}
            onOpenSettings={handleOpenSettings}
            onOpenHelp={handleOpenHelp}
            onAuthAction={handleAuthAction}
          />

          <LocationPanel
            isOpen={isLocationPanelOpen}
            activeLocation={activeLocation}
            currentUser={currentUser}
            hasGoogleToken={hasGoogleToken}
            counts={locationCounts}
            onSelectLocation={(loc) => {
              setActiveLocation(loc);
            }}
            onPickDeviceFiles={() => {
              deviceFileInputRef.current?.click();
            }}
            onBrowseGoogleDrive={() => {
              setIsDriveModalOpen(true);
            }}
            onAddGoogleAccount={() => {
              void handleConnectGoogleDrive();
            }}
            onCloseMobile={() => {
              setIsLocationPanelOpen(false);
            }}
          />

          <div className="flex-1 min-w-0 h-full flex flex-col overflow-hidden">
            <MobileTopHeader
              currentUser={currentUser}
              syncStatus={syncStatus}
              onLogoClick={handleMobileLogoClick}
              onOpenHelp={handleOpenHelp}
              onOpenSettings={handleOpenSettings}
              onAuthAction={handleAuthAction}
            />

            <DocumentLibrary
              documents={filteredDocuments}
              activeDocId={lastActiveDocId}
              searchQuery={searchQuery}
              activeFilter={activeFilter}
              activeNav={activeNav}
              isUploading={isUploading}
              uploadProgress={uploadProgress}
              isDeleting={isDeleting}
              uploadNotice={uploadNotice}
              onSearchChange={setSearchQuery}
              onFilterChange={handleFilterChange}
              onOpenDocument={handleOpenDocument}
              onCreateDocument={handleCreateDocument}
              onShareDocument={handleShareDocument}
              onToggleFavorite={handleToggleFavorite}
              onUploadFiles={handleUploadFiles}
              onDeleteDocuments={handleDeleteDocuments}
              onDismissNotice={handleDismissNotice}
            />
          </div>

          <MobileBottomNav
            activeNav={activeNav}
            onSelectNav={handleSelectNav}
          />
        </div>
      )}

      <StageActionModal
        modal={activeModal === 'share' ? null : activeModal}
        document={modalDocument}
        onClose={handleCloseModal}
        onSwitchModal={handleSwitchModal}
        onShareConfigUpdated={handleShareConfigUpdated}
      />

      <ShareCollabModal
        isOpen={activeModal === 'share'}
        onClose={handleCloseModal}
        document={modalDocument}
        currentUser={currentUser}
        collaborators={collabCollaborators}
        selfSessionId={selfSessionId}
        onShareStateChanged={handleShareStateChanged}
      />

      <AuthModal
        isOpen={activeUtilityModal === 'auth'}
        onClose={handleAuthSuccess}
        onAuthSuccess={handleAuthSuccess}
      />

      <SettingsModal
        isOpen={activeUtilityModal === 'settings'}
        user={currentUser}
        settings={userSettings}
        syncStatus={syncStatus}
        cloudDocumentCount={cloudDocumentCount}
        lastSyncedAt={lastSyncedAt}
        isSupabaseConfigured={isSupabaseConfigured}
        onClose={handleCloseUtilityModal}
        onUpdateSettings={handleUpdateSettings}
        onManualSync={handleManualSync}
        onOpenAuth={handleOpenAuthModal}
        onRequestLogout={handleRequestLogout}
      />

      <HelpModal
        isOpen={activeUtilityModal === 'help'}
        onClose={handleCloseUtilityModal}
      />

      <CloudActionModal
        mode={cloudActionModalMode}
        userEmail={currentUser?.email}
        localCandidateCount={pendingMigrationDocs.length}
        isProcessing={isMigratingLocal || isLoggingOut}
        onClose={handleCloseCloudActionModal}
        onConfirmLogout={() => {
          void handleConfirmLogout();
        }}
        onConfirmMigrateLocal={() => {
          void handleConfirmMigrateLocal();
        }}
        onSkipMigrateLocal={handleSkipMigrateLocal}
      />

      <GoogleDriveBrowserModal
        isOpen={isDriveModalOpen}
        hasGoogleToken={hasGoogleToken}
        onClose={() => {
          setIsDriveModalOpen(false);
        }}
        onConnectGoogle={() => {
          void handleConnectGoogleDrive();
        }}
        onOpenDriveFile={handleOpenDriveFile}
      />

      <input
        ref={deviceFileInputRef}
        type="file"
        multiple
        accept=".pptx,.docx,.xlsx,.xls,.csv,.pdf"
        className="hidden"
        onChange={(event) => {
          const files = event.target.files;
          if (files && files.length > 0) {
            void handleUploadFiles(files);
          }
          event.target.value = '';
        }}
      />
    </>
  );
}

export default App;

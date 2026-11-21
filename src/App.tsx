import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { SidebarRail } from './components/layout/SidebarRail';
import { MobileTopHeader } from './components/layout/MobileTopHeader';
import { MobileBottomNav } from './components/layout/MobileBottomNav';
import { DocumentLibrary } from './components/workspace/DocumentLibrary';
import { SlideWorkspace } from './components/workspaces/SlideWorkspace';
import { WordWorkspace } from './components/workspaces/WordWorkspace';
import { ExcelWorkspace } from './components/workspaces/ExcelWorkspace';
import { PdfWorkspace } from './components/workspaces/PdfWorkspace';
import { StageActionModal } from './components/modals/StageActionModal';
import { AuthModal } from './components/modals/AuthModal';
import { SettingsModal } from './components/modals/SettingsModal';
import { HelpModal } from './components/modals/HelpModal';
import { CloudActionModal } from './components/modals/CloudActionModal';
import { MobileRemoteView } from './components/remote/MobileRemoteView';
import { SharedDocumentView } from './components/share/SharedDocumentView';
import { mockDocuments } from './data/mockDocuments';
import { isSupabaseConfigured, supabase } from './lib/supabase';
import {
  buildCloudUserSummary,
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
import { parseUploadedFile } from './utils/fileParser';
import type {
  CloudUserSummary,
  DocumentItem,
  FilterKey,
  ModalType,
  ShareAccessLevel,
  SidebarNavKey,
  SyncStatus,
  UserSettings,
  UtilityModalType,
  WordPaperTheme,
} from './types/document';

const STORAGE_KEY = 'lite_saved_documents_v1';
const SETTINGS_STORAGE_KEY = 'lite_user_settings_v1';

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
    if (!slide.images || slide.images.length === 0) {
      return slide;
    }
    const persistableImages = slide.images.filter(
      (img) => !img.src.startsWith('blob:')
    );
    const nextSlide = { ...slide };
    if (persistableImages.length > 0) {
      nextSlide.images = persistableImages;
    } else {
      delete nextSlide.images;
    }
    return nextSlide;
  });

  const nextDoc: DocumentItem = {
    ...doc,
    slidesPreview: safeSlides,
  };
  delete nextDoc.sourceFile;
  if (nextDoc.objectUrl?.startsWith('blob:')) {
    delete nextDoc.objectUrl;
  }
  if (nextDoc.fileUrl?.startsWith('blob:')) {
    delete nextDoc.fileUrl;
  }
  return nextDoc;
}

function loadInitialDocuments(): DocumentItem[] {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return mockDocuments;
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return mockDocuments;
    }

    const mockById = new Map<string, DocumentItem>(
      mockDocuments.map((doc) => [doc.id, doc])
    );
    const seenIds = new Set<string>();
    const merged: DocumentItem[] = [];

    for (const entry of parsed) {
      if (!isValidStoredDocument(entry) || seenIds.has(entry.id)) {
        continue;
      }
      seenIds.add(entry.id);
      const cleanEntry = sanitizeDocumentForStorage(entry);
      const mockMatch = mockById.get(cleanEntry.id);
      if (mockMatch) {
        merged.push({
          ...mockMatch,
          ...cleanEntry,
        });
      } else {
        merged.push(cleanEntry);
      }
    }

    for (const mockDoc of mockDocuments) {
      if (!seenIds.has(mockDoc.id)) {
        merged.push(mockDoc);
      }
    }

    return merged.length > 0 ? merged : mockDocuments;
  } catch {
    return mockDocuments;
  }
}

function persistDocumentsToStorage(docs: DocumentItem[]): {
  ok: boolean;
  quotaExceeded: boolean;
} {
  try {
    const payload = docs.map(sanitizeDocumentForStorage);
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

function preserveInMemoryRuntimeFields(
  incoming: DocumentItem,
  existing?: DocumentItem
): DocumentItem {
  if (!existing) {
    return incoming;
  }
  return {
    ...incoming,
    ...(existing.sourceFile ? { sourceFile: existing.sourceFile } : {}),
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

function isUnsyncedLocalCandidate(doc: DocumentItem): boolean {
  return (
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
  const [activeFilter, setActiveFilter] = useState<FilterKey>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeModal, setActiveModal] = useState<ModalType>(null);
  const [modalDocOverride, setModalDocOverride] =
    useState<DocumentItem | null>(null);
  const [uploadNotice, setUploadNotice] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [uploadProgress, setUploadProgress] = useState<number | null>(null);

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

  useEffect(() => {
    return () => {
      createdObjectUrlsRef.current.forEach((url) => {
        URL.revokeObjectURL(url);
      });
      createdObjectUrlsRef.current = [];
    };
  }, []);

  // Priority Hash Routing: #remote/<roomCode> -> #share/<shareToken> -> #doc/<id>
  useEffect(() => {
    const syncFromLocationHash = () => {
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

    window.addEventListener('hashchange', syncFromLocationHash);
    window.addEventListener('popstate', syncFromLocationHash);
    return () => {
      window.removeEventListener('hashchange', syncFromLocationHash);
      window.removeEventListener('popstate', syncFromLocationHash);
    };
  }, []);

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
      const fetchedCloudDocs = await fetchUserDocuments(user.id);
      if (currentUserRef.current?.id !== user.id) {
        return;
      }

      const existingById = new Map<string, DocumentItem>(
        documentsRef.current.map((doc) => [doc.id, doc])
      );
      const cloudDocs = fetchedCloudDocs.map((cloudDoc) =>
        preserveInMemoryRuntimeFields(cloudDoc, existingById.get(cloudDoc.id))
      );
      const cloudIds = new Set<string>(cloudDocs.map((doc) => doc.id));

      // Keep unsynced local candidates untouched in memory while asking user to migrate
      const localCandidates = documentsRef.current.filter(
        (doc) => isUnsyncedLocalCandidate(doc) && !cloudIds.has(doc.id)
      );
      const localCandidateIds = new Set<string>(
        localCandidates.map((doc) => doc.id)
      );

      const remainingMocks = mockDocuments.filter(
        (mockDoc) =>
          !cloudIds.has(mockDoc.id) && !localCandidateIds.has(mockDoc.id)
      );

      const mergedDocs = [...cloudDocs, ...localCandidates, ...remainingMocks];
      setDocuments(mergedDocs);
      persistDocumentsToStorage(mergedDocs);

      const currentHash = window.location.hash;
      const isSpecialRoute =
        Boolean(parseRemoteRoomFromHash(currentHash)) ||
        Boolean(parseShareTokenFromHash(currentHash));

      if (!isSpecialRoute) {
        const hashId = parseDocIdFromHash(currentHash);
        if (hashId && mergedDocs.some((doc) => doc.id === hashId)) {
          setActiveWorkspaceDocId(hashId);
          setLastActiveDocId(hashId);
        } else if (mergedDocs.length > 0) {
          setLastActiveDocId((prev) =>
            prev && mergedDocs.some((doc) => doc.id === prev)
              ? prev
              : mergedDocs[0].id
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
          autoHydrateAttemptedRef.current.clear();
          const localOnly = documentsRef.current.filter(
            (doc) => !doc.isSyncedToCloud && !doc.userId
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
        handleSessionUser(data.session?.user ?? null);
      })
      .catch(() => {
        if (isMounted) {
          setSyncStatus('error');
        }
      });

    const { data: authSubscription } = supabase.auth.onAuthStateChange(
      (_event, session) => {
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
      let resolvedObjectUrl: string | null = null;

      if (doc.r2ObjectKey && isR2ApiConfigured) {
        try {
          const sessionRes = supabase ? await supabase.auth.getSession() : null;
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
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'Không thể tải tệp PDF từ đám mây.';
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
    setActiveNav(nav);
    if (nav === 'home' || nav === 'shared') {
      setActiveFilter('all');
    } else if (nav === 'slide') {
      setActiveFilter('pptx');
    } else if (nav === 'word') {
      setActiveFilter('docx');
    } else if (nav === 'excel') {
      setActiveFilter('xlsx');
    } else if (nav === 'pdf') {
      setActiveFilter('pdf');
    }
    // For 'folders', keep activeFilter intact while showing folder grouping
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

      if (
        doc.type === 'pdf' &&
        (doc.r2ObjectKey || doc.storagePath) &&
        !doc.fileUrl &&
        !doc.objectUrl
      ) {
        autoHydrateAttemptedRef.current.add(doc.id);
        void hydrateCloudPdfDocument(doc);
      }
    },
    [hydrateCloudPdfDocument]
  );

  const handleBackHome = useCallback(() => {
    setRemoteRoomCode(null);
    setSharedToken(null);
    setActiveWorkspaceDocId(null);
    setActiveModal(null);
    setModalDocOverride(null);

    if (window.location.hash) {
      window.history.pushState(
        null,
        '',
        `${window.location.pathname}${window.location.search}`
      );
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
      let toggledDoc: DocumentItem | null = null;
      const nextDocs = documentsRef.current.map((doc) => {
        if (doc.id !== docId) {
          return doc;
        }
        toggledDoc = { ...doc, isFavorite: !doc.isFavorite };
        return toggledDoc;
      });
      setDocuments(nextDocs);
      persistDocumentsToStorage(nextDocs);

      const activeUser = currentUserRef.current;
      if (
        activeUser &&
        isSupabaseConfigured &&
        toggledDoc !== null &&
        (toggledDoc as DocumentItem).isSyncedToCloud
      ) {
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
      const savedDoc: DocumentItem = {
        ...updatedDoc,
        updatedAt: 'Vừa lưu',
        lastSavedAt: new Date().toISOString(),
      };

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

      const activeUser = currentUserRef.current;
      if (activeUser && isSupabaseConfigured) {
        void enqueueCloudUpsert(activeUser.id, savedDoc, {
          file: savedDoc.sourceFile ?? null,
          touchUpdatedAt: true,
        });
      }
    },
    [enqueueCloudUpsert]
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
      const { syncedDocs, failedCount } = await syncLocalDocumentsToCloud(
        activeUser.id,
        pendingMigrationDocs
      );
      const syncedById = new Map<string, DocumentItem>(
        syncedDocs.map((doc) => [doc.id, doc])
      );

      const nextDocs = documentsRef.current.map((doc) => {
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
      (doc) => doc.isSyncedToCloud && !pendingIds.has(doc.id)
    );
    const cloudIds = new Set<string>(cloudDocs.map((doc) => doc.id));
    const defaultMocks = mockDocuments.filter(
      (mockDoc) => !cloudIds.has(mockDoc.id)
    );
    const nextDocs = [...cloudDocs, ...defaultMocks];

    setDocuments(nextDocs);
    persistDocumentsToStorage(nextDocs);
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

      const fetchedCloudDocs = await fetchUserDocuments(activeUser.id);
      if (currentUserRef.current?.id !== activeUser.id) {
        return;
      }

      const existingById = new Map<string, DocumentItem>(
        documentsRef.current.map((doc) => [doc.id, doc])
      );
      const cloudDocs = fetchedCloudDocs.map((cloudDoc) =>
        preserveInMemoryRuntimeFields(cloudDoc, existingById.get(cloudDoc.id))
      );
      const cloudIds = new Set<string>(cloudDocs.map((doc) => doc.id));
      const remainingLocal =
        failedCount > 0
          ? localCandidates.filter((doc) => !cloudIds.has(doc.id))
          : [];
      const remainingLocalIds = new Set<string>(
        remainingLocal.map((doc) => doc.id)
      );
      const remainingMocks = mockDocuments.filter(
        (mockDoc) =>
          !cloudIds.has(mockDoc.id) && !remainingLocalIds.has(mockDoc.id)
      );

      const mergedDocs = [...cloudDocs, ...remainingLocal, ...remainingMocks];
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
      currentUserRef.current = null;
      lastLoadedUserIdRef.current = null;
      inFlightSyncUserIdRef.current = null;
      autoHydrateAttemptedRef.current.clear();

      // Strip cloud-owned documents from localStorage before returning to Guest Local Mode
      const pureLocalDocs = documentsRef.current.filter(
        (doc) => !doc.isSyncedToCloud && !doc.userId
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

  const filteredDocuments = useMemo(() => {
    const query = searchQuery.trim().toLowerCase();

    const matched = documents.filter((doc) => {
      if (activeNav === 'shared' && !doc.isShared) {
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

    if (activeNav === 'folders') {
      return [...matched].sort((a, b) =>
        (a.folder ?? 'Chưa phân loại').localeCompare(
          b.folder ?? 'Chưa phân loại',
          'vi'
        )
      );
    }

    return matched;
  }, [documents, activeNav, activeFilter, searchQuery]);

  const activeWorkspaceDoc = useMemo(() => {
    if (!activeWorkspaceDocId) {
      return null;
    }
    return documents.find((doc) => doc.id === activeWorkspaceDocId) ?? null;
  }, [documents, activeWorkspaceDocId]);

  // Auto-hydrate cloud PDF once if opened directly via URL hash after reload
  useEffect(() => {
    if (
      !activeWorkspaceDoc ||
      activeWorkspaceDoc.type !== 'pdf' ||
      (!activeWorkspaceDoc.r2ObjectKey && !activeWorkspaceDoc.storagePath) ||
      activeWorkspaceDoc.fileUrl ||
      activeWorkspaceDoc.objectUrl ||
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

  const isAnyModalOpen = activeModal !== null || activeUtilityModal !== null;

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

  return (
    <>
      {activeWorkspaceDoc ? (
        activeWorkspaceDoc.type === 'pptx' ? (
          <SlideWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            autoSaveEnabled={userSettings.autoSyncOnEdit}
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
            onBackHome={handleBackHome}
            onSaveDocument={handleSaveDocument}
            onOpenModal={handleOpenModal}
          />
        ) : activeWorkspaceDoc.type === 'xlsx' ? (
          <ExcelWorkspace
            document={activeWorkspaceDoc}
            isModalOpen={isAnyModalOpen}
            autoSaveEnabled={userSettings.autoSyncOnEdit}
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
              uploadNotice={uploadNotice}
              onSearchChange={setSearchQuery}
              onFilterChange={handleFilterChange}
              onOpenDocument={handleOpenDocument}
              onToggleFavorite={handleToggleFavorite}
              onUploadFiles={handleUploadFiles}
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
        modal={activeModal}
        document={modalDocument}
        onClose={handleCloseModal}
        onSwitchModal={handleSwitchModal}
        onShareConfigUpdated={handleShareConfigUpdated}
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
    </>
  );
}

export default App;

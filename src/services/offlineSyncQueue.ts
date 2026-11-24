import type { DocumentFolder, DocumentItem, SlidePreviewItem } from '../types/document';

const DB_NAME = 'lite_offline_sync_db_v1';
const DB_VERSION = 4;
const STORE_DOCUMENTS = 'documents';
const STORE_FOLDERS = 'folders';
const STORE_DRAFTS = 'drafts';
const STORE_PENDING = 'pendingUploads';
const STORE_VERSIONS = 'versions';
const STORE_ORIGINAL_BINARIES = 'originalBinaries';
const ALL_STORES = [
  STORE_DOCUMENTS,
  STORE_FOLDERS,
  STORE_DRAFTS,
  STORE_PENDING,
  STORE_VERSIONS,
  STORE_ORIGINAL_BINARIES,
] as const;
const MAX_VERSIONS_PER_DOC = 10;
export const AUTO_TRASH_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;

export type PendingUploadTarget = 'supabase' | 'drive' | 'both';
export type VersionSaveKind = 'manual' | 'auto';

export interface CachedFolderRecord {
  folderKey: string;
  userId: string;
  ownerScope: string;
  folderId: string;
  folder: DocumentFolder;
}

export interface CachedDocumentRecord {
  docKey: string;
  userId: string;
  ownerScope: string;
  docId: string;
  updatedAt: string;
  documentSnapshot: DocumentItem;
}

export interface LocalDraftRecord {
  docKey: string;
  userId: string;
  ownerScope?: string;
  docId: string;
  updatedAt: string;
  documentSnapshot: DocumentItem;
}

export interface PendingUploadTask {
  id: string;
  docKey: string;
  userId: string;
  ownerScope?: string;
  docId: string;
  target: PendingUploadTarget;
  documentSnapshot: DocumentItem;
  queuedAt: string;
  attemptCount: number;
  lastError?: string;
}

export interface DocumentVersionRecord {
  id: string;
  docKey: string;
  userId: string;
  ownerScope?: string;
  docId: string;
  title: string;
  savedAt: string;
  fileSizeBytes: number;
  savedByName: string;
  saveKind?: VersionSaveKind;
  documentSnapshot: DocumentItem;
}

export interface OriginalBinaryRecord {
  docKey: string;
  userId: string;
  ownerScope?: string;
  docId: string;
  fileName: string;
  mimeType: string;
  buffer: ArrayBuffer;
  savedAt: string;
}

// In-memory fallback when IndexedDB is unavailable (private mode / tests)
const memoryDocuments = new Map<string, CachedDocumentRecord>();
const memoryFolders = new Map<string, CachedFolderRecord>();
const memoryDrafts = new Map<string, LocalDraftRecord>();
const memoryPending = new Map<string, PendingUploadTask>();
const memoryVersions = new Map<string, DocumentVersionRecord>();
const memoryOriginalBinaries = new Map<string, OriginalBinaryRecord>();

let dbPromise: Promise<IDBDatabase | null> | null = null;

export function normalizeSyncUserId(userId?: string | null): string {
  const trimmed = userId?.trim();
  return trimmed && trimmed.length > 0 ? trimmed : 'guest';
}

export function makeUserDocKey(
  userId: string | null | undefined,
  docId: string
): string {
  return `${normalizeSyncUserId(userId)}:${docId}`;
}

function isBlobUrl(value?: string): boolean {
  return typeof value === 'string' && value.trim().startsWith('blob:');
}

function cloneArrayBufferSafe(buffer: ArrayBuffer | undefined): ArrayBuffer | undefined {
  if (!(buffer instanceof ArrayBuffer)) {
    return undefined;
  }
  try {
    if (buffer.byteLength === 0) {
      return undefined;
    }
    return buffer.slice(0);
  } catch {
    // Detached ArrayBuffer cannot be cloned
    return undefined;
  }
}

export function sanitizeDocumentForIndexedDb(doc: DocumentItem): DocumentItem {
  const sanitizedSlides: SlidePreviewItem[] = Array.isArray(doc.slidesPreview)
    ? doc.slidesPreview.map((slide) => {
        const cleanImages = slide.images
          ?.filter((img) => !isBlobUrl(img.src))
          .map((img) => ({ ...img }));
        const cleanBgImg = isBlobUrl(slide.backgroundImageSrc)
          ? undefined
          : slide.backgroundImageSrc;
        return {
          ...slide,
          blocks: slide.blocks?.map((block) => ({
            ...block,
            tableRows: block.tableRows?.map((row) => [...row]),
          })),
          images: cleanImages && cleanImages.length > 0 ? cleanImages : undefined,
          backgroundImageSrc: cleanBgImg,
        };
      })
    : [];

  const clonedBuffer = cloneArrayBufferSafe(doc.rawArrayBuffer);

  const sanitized: DocumentItem = {
    ...doc,
    sourceFile: null,
    objectUrl: isBlobUrl(doc.objectUrl) ? undefined : doc.objectUrl,
    fileUrl: isBlobUrl(doc.fileUrl) ? undefined : doc.fileUrl,
    slidesPreview: sanitizedSlides,
    outline: Array.isArray(doc.outline)
      ? doc.outline.map((item) => ({
          ...item,
          keyPoints: [...item.keyPoints],
        }))
      : [],
    sheets: doc.sheets?.map((sheet) => ({
      ...sheet,
      columns: [...sheet.columns],
      rows: sheet.rows.map((row) => [...row]),
      cellStyles: sheet.cellStyles ? { ...sheet.cellStyles } : undefined,
    })),
    pdfPages: doc.pdfPages?.map((page) => ({
      ...page,
      paragraphs: [...page.paragraphs],
      highlights: page.highlights ? [...page.highlights] : undefined,
    })),
    wordParagraphs: doc.wordParagraphs
      ? Object.fromEntries(
          Object.entries(doc.wordParagraphs).map(([k, v]) => [k, [...v]])
        )
      : undefined,
    rawArrayBuffer: clonedBuffer,
  };

  return sanitized;
}

function estimateSnapshotBytes(doc: DocumentItem): number {
  const rawBufferBytes =
    doc.rawArrayBuffer instanceof ArrayBuffer ? doc.rawArrayBuffer.byteLength : 0;
  try {
    const withoutBuffer: DocumentItem = {
      ...doc,
      rawArrayBuffer: undefined,
      sourceFile: null,
    };
    const json = JSON.stringify(withoutBuffer);
    const textBytes =
      typeof TextEncoder !== 'undefined'
        ? new TextEncoder().encode(json).byteLength
        : json.length * 2;
    return textBytes + rawBufferBytes;
  } catch {
    return rawBufferBytes;
  }
}

function getIndexedDbFactory(): IDBFactory | null {
  if (typeof window === 'undefined') {
    return null;
  }
  try {
    return window.indexedDB ?? null;
  } catch {
    return null;
  }
}

function openOfflineDb(): Promise<IDBDatabase | null> {
  if (dbPromise) {
    return dbPromise;
  }

  const idb = getIndexedDbFactory();
  if (!idb) {
    dbPromise = Promise.resolve(null);
    return dbPromise;
  }

  dbPromise = new Promise<IDBDatabase | null>((resolve) => {
    try {
      const request = idb.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = () => {
        const db = request.result;
        const tx = request.transaction;
        const ensureIndex = (
          storeName: string,
          indexName: string,
          keyPath: string
        ) => {
          const store = tx?.objectStore(storeName);
          if (store && !store.indexNames.contains(indexName)) {
            store.createIndex(indexName, keyPath, { unique: false });
          }
        };

        if (!db.objectStoreNames.contains(STORE_DOCUMENTS)) {
          const docsStore = db.createObjectStore(STORE_DOCUMENTS, {
            keyPath: 'docKey',
          });
          docsStore.createIndex('userId', 'userId', { unique: false });
          docsStore.createIndex('ownerScope', 'ownerScope', { unique: false });
        } else {
          ensureIndex(STORE_DOCUMENTS, 'userId', 'userId');
          ensureIndex(STORE_DOCUMENTS, 'ownerScope', 'ownerScope');
        }

        if (!db.objectStoreNames.contains(STORE_FOLDERS)) {
          const foldersStore = db.createObjectStore(STORE_FOLDERS, {
            keyPath: 'folderKey',
          });
          foldersStore.createIndex('userId', 'userId', { unique: false });
          foldersStore.createIndex('ownerScope', 'ownerScope', { unique: false });
        } else {
          ensureIndex(STORE_FOLDERS, 'userId', 'userId');
          ensureIndex(STORE_FOLDERS, 'ownerScope', 'ownerScope');
        }

        if (!db.objectStoreNames.contains(STORE_DRAFTS)) {
          const draftsStore = db.createObjectStore(STORE_DRAFTS, {
            keyPath: 'docKey',
          });
          draftsStore.createIndex('userId', 'userId', { unique: false });
          draftsStore.createIndex('ownerScope', 'ownerScope', { unique: false });
        } else {
          ensureIndex(STORE_DRAFTS, 'ownerScope', 'ownerScope');
        }

        if (!db.objectStoreNames.contains(STORE_PENDING)) {
          const pendingStore = db.createObjectStore(STORE_PENDING, {
            keyPath: 'id',
          });
          pendingStore.createIndex('userId', 'userId', { unique: false });
          pendingStore.createIndex('ownerScope', 'ownerScope', { unique: false });
          pendingStore.createIndex('docKey', 'docKey', { unique: false });
        } else {
          ensureIndex(STORE_PENDING, 'ownerScope', 'ownerScope');
        }

        if (!db.objectStoreNames.contains(STORE_VERSIONS)) {
          const versionsStore = db.createObjectStore(STORE_VERSIONS, {
            keyPath: 'id',
          });
          versionsStore.createIndex('docKey', 'docKey', { unique: false });
          versionsStore.createIndex('userId', 'userId', { unique: false });
          versionsStore.createIndex('ownerScope', 'ownerScope', { unique: false });
        } else {
          ensureIndex(STORE_VERSIONS, 'ownerScope', 'ownerScope');
        }

        if (!db.objectStoreNames.contains(STORE_ORIGINAL_BINARIES)) {
          const binaryStore = db.createObjectStore(STORE_ORIGINAL_BINARIES, {
            keyPath: 'docKey',
          });
          binaryStore.createIndex('userId', 'userId', { unique: false });
          binaryStore.createIndex('ownerScope', 'ownerScope', { unique: false });
          binaryStore.createIndex('docId', 'docId', { unique: false });
        } else {
          ensureIndex(STORE_ORIGINAL_BINARIES, 'ownerScope', 'ownerScope');
        }
      };

      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => {
          db.close();
          dbPromise = null;
        };
        resolve(db);
      };

      request.onerror = () => {
        resolve(null);
      };

      request.onblocked = () => {
        resolve(null);
      };
    } catch {
      resolve(null);
    }
  });

  return dbPromise;
}

function runIdbRequest<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'));
  });
}

function waitForTransaction(tx: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('IndexedDB transaction failed'));
    tx.onabort = () => reject(tx.error ?? new Error('IndexedDB transaction aborted'));
  });
}

function mergePendingTarget(
  existing: PendingUploadTarget,
  incoming: PendingUploadTarget
): PendingUploadTarget {
  if (existing === incoming) {
    return existing;
  }
  return 'both';
}

export async function saveLocalDraft(
  userId: string | null | undefined,
  doc: DocumentItem
): Promise<void> {
  const normalizedUser = normalizeSyncUserId(userId);
  const docKey = makeUserDocKey(normalizedUser, doc.id);
  const snapshot = {
    ...sanitizeDocumentForIndexedDb(doc),
    userId: doc.userId ?? normalizedUser,
    ownerScope: normalizedUser,
  };
  const record: LocalDraftRecord = {
    docKey,
    userId: normalizedUser,
    ownerScope: normalizedUser,
    docId: doc.id,
    updatedAt: new Date().toISOString(),
    documentSnapshot: snapshot,
  };

  memoryDrafts.set(docKey, record);

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const tx = db.transaction(STORE_DRAFTS, 'readwrite');
    tx.objectStore(STORE_DRAFTS).put(record);
    await waitForTransaction(tx);
  } catch {
    // Retry without rawArrayBuffer if structured clone fails for any edge reason
    try {
      const tx = db.transaction(STORE_DRAFTS, 'readwrite');
      tx.objectStore(STORE_DRAFTS).put({
        ...record,
        documentSnapshot: { ...snapshot, rawArrayBuffer: undefined },
      });
      await waitForTransaction(tx);
    } catch {
      // Keep memoryDrafts fallback
    }
  }
}

export async function loadLocalDraft(
  userId: string | null | undefined,
  docId: string
): Promise<DocumentItem | null> {
  const docKey = makeUserDocKey(userId, docId);
  const db = await openOfflineDb();

  if (db) {
    try {
      const tx = db.transaction(STORE_DRAFTS, 'readonly');
      const record = await runIdbRequest<LocalDraftRecord | undefined>(
        tx.objectStore(STORE_DRAFTS).get(docKey)
      );
      if (record?.documentSnapshot) {
        return record.documentSnapshot;
      }
    } catch {
      // Fallback to memory map
    }
  }

  return memoryDrafts.get(docKey)?.documentSnapshot ?? null;
}

export async function listLocalDrafts(
  userId?: string | null
): Promise<DocumentItem[]> {
  const normalizedUser = normalizeSyncUserId(userId);
  const db = await openOfflineDb();

  if (db) {
    try {
      const tx = db.transaction(STORE_DRAFTS, 'readonly');
      const store = tx.objectStore(STORE_DRAFTS);
      const index = store.index('userId');
      const records = await runIdbRequest<LocalDraftRecord[]>(
        index.getAll(normalizedUser)
      );
      return records
        .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt))
        .map((r) => r.documentSnapshot);
    } catch {
      // Fallback to memory
    }
  }

  return Array.from(memoryDrafts.values())
    .filter((r) => r.userId === normalizedUser)
    .sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt))
    .map((r) => r.documentSnapshot);
}

export async function deleteLocalDraft(
  userId: string | null | undefined,
  docId: string
): Promise<void> {
  const docKey = makeUserDocKey(userId, docId);
  memoryDrafts.delete(docKey);

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const tx = db.transaction(STORE_DRAFTS, 'readwrite');
    tx.objectStore(STORE_DRAFTS).delete(docKey);
    await waitForTransaction(tx);
  } catch {
    // Ignore cleanup errors
  }
}

export async function enqueuePendingUpload(
  userId: string | null | undefined,
  doc: DocumentItem,
  target: PendingUploadTarget,
  lastError?: string
): Promise<PendingUploadTask> {
  const normalizedUser = normalizeSyncUserId(userId);
  const docKey = makeUserDocKey(normalizedUser, doc.id);
  const snapshot = {
    ...sanitizeDocumentForIndexedDb(doc),
    userId: doc.userId ?? normalizedUser,
    ownerScope: normalizedUser,
  };
  const nowIso = new Date().toISOString();

  const existingMemory = Array.from(memoryPending.values()).find(
    (item) => item.docKey === docKey
  );

  let task: PendingUploadTask = {
    id: existingMemory?.id ?? `pending-${normalizedUser}-${doc.id}`,
    docKey,
    userId: normalizedUser,
    ownerScope: normalizedUser,
    docId: doc.id,
    target: existingMemory
      ? mergePendingTarget(existingMemory.target, target)
      : target,
    documentSnapshot: snapshot,
    queuedAt: nowIso,
    attemptCount: existingMemory?.attemptCount ?? 0,
    lastError,
  };

  memoryPending.set(task.id, task);

  const db = await openOfflineDb();
  if (!db) {
    return task;
  }

  try {
    const readTx = db.transaction(STORE_PENDING, 'readonly');
    const docKeyIndex = readTx.objectStore(STORE_PENDING).index('docKey');
    const existingList = await runIdbRequest<PendingUploadTask[]>(
      docKeyIndex.getAll(docKey)
    );
    const existing = existingList[0];

    if (existing) {
      task = {
        ...task,
        id: existing.id,
        target: mergePendingTarget(existing.target, target),
        attemptCount: existing.attemptCount,
      };
      memoryPending.set(task.id, task);
    }

    const writeTx = db.transaction(STORE_PENDING, 'readwrite');
    writeTx.objectStore(STORE_PENDING).put(task);
    await waitForTransaction(writeTx);
  } catch {
    try {
      const writeTx = db.transaction(STORE_PENDING, 'readwrite');
      writeTx.objectStore(STORE_PENDING).put({
        ...task,
        documentSnapshot: { ...snapshot, rawArrayBuffer: undefined },
      });
      await waitForTransaction(writeTx);
    } catch {
      // Keep memory fallback
    }
  }

  return task;
}

export async function getPendingUploads(
  userId?: string | null
): Promise<PendingUploadTask[]> {
  const filterByUser = userId !== undefined;
  const normalizedUser = filterByUser ? normalizeSyncUserId(userId) : null;

  const db = await openOfflineDb();
  if (db) {
    try {
      const tx = db.transaction(STORE_PENDING, 'readonly');
      const store = tx.objectStore(STORE_PENDING);
      const tasks =
        normalizedUser !== null
          ? await runIdbRequest<PendingUploadTask[]>(
              store.index('userId').getAll(normalizedUser)
            )
          : await runIdbRequest<PendingUploadTask[]>(store.getAll());

      return tasks.sort(
        (a, b) => Date.parse(a.queuedAt) - Date.parse(b.queuedAt)
      );
    } catch {
      // Fallback to memory
    }
  }

  return Array.from(memoryPending.values())
    .filter((item) =>
      normalizedUser !== null ? item.userId === normalizedUser : true
    )
    .sort((a, b) => Date.parse(a.queuedAt) - Date.parse(b.queuedAt));
}

export async function removePendingUpload(id: string): Promise<void> {
  memoryPending.delete(id);

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const tx = db.transaction(STORE_PENDING, 'readwrite');
    tx.objectStore(STORE_PENDING).delete(id);
    await waitForTransaction(tx);
  } catch {
    // Ignore errors
  }
}

export async function flushPendingUploads(
  userId: string | null | undefined,
  processor: (task: PendingUploadTask) => Promise<void>
): Promise<{ processed: number; failed: number; remaining: number }> {
  const tasks = await getPendingUploads(userId);
  let processed = 0;
  let failed = 0;

  for (const task of tasks) {
    try {
      await processor(task);
      await removePendingUpload(task.id);
      processed += 1;
    } catch (error) {
      failed += 1;
      const errMsg =
        error instanceof Error ? error.message : 'Đồng bộ ngoại tuyến thất bại.';
      const updatedTask: PendingUploadTask = {
        ...task,
        attemptCount: task.attemptCount + 1,
        lastError: errMsg,
      };
      memoryPending.set(updatedTask.id, updatedTask);

      const db = await openOfflineDb();
      if (db) {
        try {
          const tx = db.transaction(STORE_PENDING, 'readwrite');
          tx.objectStore(STORE_PENDING).put(updatedTask);
          await waitForTransaction(tx);
        } catch {
          // Ignore update error
        }
      }
    }
  }

  const remainingTasks = await getPendingUploads(userId);
  return {
    processed,
    failed,
    remaining: remainingTasks.length,
  };
}

function buildDocumentVersionFingerprint(doc: DocumentItem): string {
  try {
    return JSON.stringify({
      title: doc.title,
      wordHtml: doc.wordHtml ?? '',
      slidesPreview: doc.slidesPreview ?? [],
      slideAspectRatio: doc.slideAspectRatio ?? null,
      sheets: doc.sheets ?? null,
      pdfPages: doc.pdfPages ?? null,
      outline: doc.outline ?? [],
      wordComments: doc.wordComments ?? null,
      wordDrawings: doc.wordDrawings ?? null,
      wordProtected: Boolean(doc.wordProtected),
      slideProtected: Boolean(doc.slideProtected),
    });
  } catch {
    return `${doc.title}:${doc.lastSavedAt ?? doc.updatedAt}:${doc.fileSize ?? ''}`;
  }
}

export async function saveDocumentVersion(
  userId: string | null | undefined,
  doc: DocumentItem,
  savedByName?: string | null,
  options?: {
    saveKind?: VersionSaveKind;
    skipIfUnchanged?: boolean;
  }
): Promise<DocumentVersionRecord> {
  const normalizedUser = normalizeSyncUserId(userId);
  const docKey = makeUserDocKey(normalizedUser, doc.id);
  const saveKind: VersionSaveKind = options?.saveKind ?? 'manual';
  const snapshot = {
    ...sanitizeDocumentForIndexedDb(doc),
    userId: doc.userId ?? normalizedUser,
    ownerScope: normalizedUser,
  };

  if (options?.skipIfUnchanged) {
    const targetFingerprint = buildDocumentVersionFingerprint(snapshot);
    const dbExisting = await openOfflineDb();
    let latestVersion: DocumentVersionRecord | undefined;

    if (dbExisting) {
      try {
        const readTx = dbExisting.transaction(STORE_VERSIONS, 'readonly');
        const records = await runIdbRequest<DocumentVersionRecord[]>(
          readTx.objectStore(STORE_VERSIONS).index('docKey').getAll(docKey)
        );
        latestVersion = records.sort(
          (a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt)
        )[0];
      } catch {
        latestVersion = undefined;
      }
    }
    if (!latestVersion) {
      latestVersion = Array.from(memoryVersions.values())
        .filter((item) => item.docKey === docKey)
        .sort((a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt))[0];
    }
    if (
      latestVersion &&
      buildDocumentVersionFingerprint(latestVersion.documentSnapshot) ===
        targetFingerprint
    ) {
      return latestVersion;
    }
  }

  const savedAt = new Date().toISOString();
  const displaySaver =
    savedByName?.trim() ||
    (normalizedUser === 'guest' ? 'Khách (Thiết bị này)' : 'Bạn');

  const record: DocumentVersionRecord = {
    id: `ver-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    docKey,
    userId: normalizedUser,
    ownerScope: normalizedUser,
    docId: doc.id,
    title: doc.title,
    savedAt,
    fileSizeBytes: estimateSnapshotBytes(snapshot),
    savedByName: displaySaver,
    saveKind,
    documentSnapshot: snapshot,
  };

  memoryVersions.set(record.id, record);
  const memForDoc = Array.from(memoryVersions.values())
    .filter((item) => item.docKey === docKey)
    .sort((a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt));
  if (memForDoc.length > MAX_VERSIONS_PER_DOC) {
    for (const stale of memForDoc.slice(MAX_VERSIONS_PER_DOC)) {
      memoryVersions.delete(stale.id);
    }
  }

  const db = await openOfflineDb();
  if (!db) {
    return record;
  }

  try {
    const writeTx = db.transaction(STORE_VERSIONS, 'readwrite');
    const store = writeTx.objectStore(STORE_VERSIONS);
    try {
      store.put(record);
    } catch {
      store.put({
        ...record,
        documentSnapshot: { ...snapshot, rawArrayBuffer: undefined },
      });
    }
    await waitForTransaction(writeTx);

    const readTx = db.transaction(STORE_VERSIONS, 'readonly');
    const existing = await runIdbRequest<DocumentVersionRecord[]>(
      readTx.objectStore(STORE_VERSIONS).index('docKey').getAll(docKey)
    );
    const sorted = existing.sort(
      (a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt)
    );

    if (sorted.length > MAX_VERSIONS_PER_DOC) {
      const staleItems = sorted.slice(MAX_VERSIONS_PER_DOC);
      const pruneTx = db.transaction(STORE_VERSIONS, 'readwrite');
      const pruneStore = pruneTx.objectStore(STORE_VERSIONS);
      for (const stale of staleItems) {
        pruneStore.delete(stale.id);
      }
      await waitForTransaction(pruneTx);
    }
  } catch {
    // Keep memory fallback
  }

  return record;
}

export async function listDocumentVersions(
  userId: string | null | undefined,
  docId: string
): Promise<DocumentVersionRecord[]> {
  const docKey = makeUserDocKey(userId, docId);
  const db = await openOfflineDb();

  if (db) {
    try {
      const tx = db.transaction(STORE_VERSIONS, 'readonly');
      const records = await runIdbRequest<DocumentVersionRecord[]>(
        tx.objectStore(STORE_VERSIONS).index('docKey').getAll(docKey)
      );
      return records
        .sort((a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt))
        .slice(0, MAX_VERSIONS_PER_DOC);
    } catch {
      // Fallback to memory
    }
  }

  return Array.from(memoryVersions.values())
    .filter((item) => item.docKey === docKey)
    .sort((a, b) => Date.parse(b.savedAt) - Date.parse(a.savedAt))
    .slice(0, MAX_VERSIONS_PER_DOC);
}

export async function saveOriginalFileBinary(
  userId: string | null | undefined,
  docId: string,
  buffer: ArrayBuffer,
  fileName: string,
  mimeType: string
): Promise<void> {
  const safeBuffer = cloneArrayBufferSafe(buffer);
  if (!safeBuffer) {
    return;
  }

  const normalizedUser = normalizeSyncUserId(userId);
  const docKey = makeUserDocKey(normalizedUser, docId);
  const record: OriginalBinaryRecord = {
    docKey,
    userId: normalizedUser,
    ownerScope: normalizedUser,
    docId,
    fileName: fileName.trim() || `${docId}`,
    mimeType: mimeType.trim() || 'application/octet-stream',
    buffer: safeBuffer,
    savedAt: new Date().toISOString(),
  };

  memoryOriginalBinaries.set(docKey, record);

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const tx = db.transaction(STORE_ORIGINAL_BINARIES, 'readwrite');
    tx.objectStore(STORE_ORIGINAL_BINARIES).put(record);
    await waitForTransaction(tx);
  } catch {
    // Keep memory fallback
  }
}

export async function loadOriginalFileBinary(
  userId: string | null | undefined,
  docId: string
): Promise<OriginalBinaryRecord | null> {
  const docKey = makeUserDocKey(userId, docId);
  const db = await openOfflineDb();

  if (db) {
    try {
      const tx = db.transaction(STORE_ORIGINAL_BINARIES, 'readonly');
      const record = await runIdbRequest<OriginalBinaryRecord | undefined>(
        tx.objectStore(STORE_ORIGINAL_BINARIES).get(docKey)
      );
      if (record && record.buffer instanceof ArrayBuffer) {
        const cloned = cloneArrayBufferSafe(record.buffer);
        if (cloned) {
          return {
            ...record,
            buffer: cloned,
          };
        }
      }
    } catch {
      // Fallback to memory map
    }
  }

  const cached = memoryOriginalBinaries.get(docKey);
  if (!cached) {
    return null;
  }
  const cloned = cloneArrayBufferSafe(cached.buffer);
  if (!cloned) {
    return null;
  }
  return {
    ...cached,
    buffer: cloned,
  };
}

export async function deleteOriginalFileBinary(
  userId: string | null | undefined,
  docId: string
): Promise<void> {
  const docKey = makeUserDocKey(userId, docId);
  memoryOriginalBinaries.delete(docKey);

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const tx = db.transaction(STORE_ORIGINAL_BINARIES, 'readwrite');
    tx.objectStore(STORE_ORIGINAL_BINARIES).delete(docKey);
    await waitForTransaction(tx);
  } catch {
    // Ignore cleanup errors
  }
}

function isDocumentAllowedForOwnerScope(
  doc: DocumentItem,
  ownerScope: string
): boolean {
  const docScope = doc.ownerScope?.trim();
  const docUserId = doc.userId?.trim();
  if (ownerScope === 'guest') {
    if (doc.isSyncedToCloud || doc.sharedOwnerId) {
      return false;
    }
    if (docScope && docScope !== 'guest') {
      return false;
    }
    if (docUserId && docUserId !== 'guest') {
      return false;
    }
    return true;
  }
  if (docScope && docScope !== ownerScope) {
    return false;
  }
  if (docUserId && docUserId !== ownerScope && !doc.isShared) {
    return false;
  }
  return true;
}

/**
 * Replaces the IndexedDB cached documents list strictly within `ownerScope` (`user.id` or `'guest'`).
 */
export async function saveDocumentsForOwner(
  userId: string | null | undefined,
  docs: DocumentItem[]
): Promise<void> {
  const ownerScope = normalizeSyncUserId(userId);
  const records: CachedDocumentRecord[] = docs
    .filter((doc) => isDocumentAllowedForOwnerScope(doc, ownerScope))
    .map((doc) => {
      const snapshot: DocumentItem = {
        ...sanitizeDocumentForIndexedDb(doc),
        rawArrayBuffer: undefined,
        ownerScope,
        userId: doc.userId ?? ownerScope,
      };
      return {
        docKey: makeUserDocKey(ownerScope, doc.id),
        userId: ownerScope,
        ownerScope,
        docId: doc.id,
        updatedAt: snapshot.lastSavedAt ?? new Date().toISOString(),
        documentSnapshot: snapshot,
      };
    });

  for (const [key, entry] of memoryDocuments.entries()) {
    if ((entry.ownerScope || entry.userId) === ownerScope) {
      memoryDocuments.delete(key);
    }
  }
  for (const rec of records) {
    memoryDocuments.set(rec.docKey, rec);
  }

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const readTx = db.transaction(STORE_DOCUMENTS, 'readonly');
    const existing = await runIdbRequest<CachedDocumentRecord[]>(
      readTx.objectStore(STORE_DOCUMENTS).index('userId').getAll(ownerScope)
    );
    const keepKeys = new Set(records.map((r) => r.docKey));

    const writeTx = db.transaction(STORE_DOCUMENTS, 'readwrite');
    const store = writeTx.objectStore(STORE_DOCUMENTS);
    for (const stale of existing) {
      if (!keepKeys.has(stale.docKey)) {
        store.delete(stale.docKey);
      }
    }
    for (const rec of records) {
      store.put(rec);
    }
    await waitForTransaction(writeTx);
  } catch {
    // Keep memory fallback
  }
}

/**
 * Loads cached documents from IndexedDB strictly matching `userId === ownerScope`
 * (`user.id` when signed in, `'guest'` when signed out). Never leaks any other user's files.
 */
export async function loadDocuments(
  userId?: string | null
): Promise<DocumentItem[]> {
  const ownerScope = normalizeSyncUserId(userId);
  const db = await openOfflineDb();

  if (db) {
    try {
      const tx = db.transaction(STORE_DOCUMENTS, 'readonly');
      const store = tx.objectStore(STORE_DOCUMENTS);
      const records = await runIdbRequest<CachedDocumentRecord[]>(
        store.index('userId').getAll(ownerScope)
      );
      return records
        .filter(
          (rec) =>
            rec.userId === ownerScope &&
            (rec.ownerScope ?? rec.userId) === ownerScope &&
            isDocumentAllowedForOwnerScope(rec.documentSnapshot, ownerScope)
        )
        .map((rec) => ({
          ...rec.documentSnapshot,
          ownerScope,
          userId: rec.documentSnapshot.userId ?? ownerScope,
        }));
    } catch {
      // Fallback to memory
    }
  }

  return Array.from(memoryDocuments.values())
    .filter(
      (rec) =>
        rec.userId === ownerScope &&
        (rec.ownerScope ?? rec.userId) === ownerScope &&
        isDocumentAllowedForOwnerScope(rec.documentSnapshot, ownerScope)
    )
    .map((rec) => ({
      ...rec.documentSnapshot,
      ownerScope,
      userId: rec.documentSnapshot.userId ?? ownerScope,
    }));
}

export async function saveFoldersForOwner(
  userId: string | null | undefined,
  folders: DocumentFolder[]
): Promise<void> {
  const ownerScope = normalizeSyncUserId(userId);
  const records: CachedFolderRecord[] = folders.map((folder) => ({
    folderKey: makeUserDocKey(ownerScope, folder.id),
    userId: ownerScope,
    ownerScope,
    folderId: folder.id,
    folder: {
      ...folder,
      userId: folder.userId ?? ownerScope,
      ownerScope,
    },
  }));

  for (const [k, existing] of memoryFolders.entries()) {
    if (existing.ownerScope === ownerScope || existing.userId === ownerScope) {
      memoryFolders.delete(k);
    }
  }
  for (const rec of records) {
    memoryFolders.set(rec.folderKey, rec);
  }

  const db = await openOfflineDb();
  if (!db || !db.objectStoreNames.contains(STORE_FOLDERS)) {
    return;
  }

  try {
    const readTx = db.transaction(STORE_FOLDERS, 'readonly');
    const existing = await runIdbRequest<CachedFolderRecord[]>(
      readTx.objectStore(STORE_FOLDERS).index('userId').getAll(ownerScope)
    );
    const keepKeys = new Set(records.map((r) => r.folderKey));

    const writeTx = db.transaction(STORE_FOLDERS, 'readwrite');
    const store = writeTx.objectStore(STORE_FOLDERS);
    for (const stale of existing) {
      if (!keepKeys.has(stale.folderKey)) {
        store.delete(stale.folderKey);
      }
    }
    for (const rec of records) {
      store.put(rec);
    }
    await waitForTransaction(writeTx);
  } catch {
    // Keep memory fallback
  }
}

export async function loadFolders(
  userId?: string | null
): Promise<DocumentFolder[]> {
  const ownerScope = normalizeSyncUserId(userId);
  const db = await openOfflineDb();

  if (db && db.objectStoreNames.contains(STORE_FOLDERS)) {
    try {
      const tx = db.transaction(STORE_FOLDERS, 'readonly');
      const store = tx.objectStore(STORE_FOLDERS);
      const records = await runIdbRequest<CachedFolderRecord[]>(
        store.index('userId').getAll(ownerScope)
      );
      return records
        .filter(
          (rec) =>
            rec.userId === ownerScope &&
            (rec.ownerScope ?? rec.userId) === ownerScope
        )
        .map((rec) => ({
          ...rec.folder,
          ownerScope,
          userId: rec.folder.userId ?? ownerScope,
        }))
        .sort((a, b) => Date.parse(a.createdAt) - Date.parse(b.createdAt));
    } catch {
      // Fallback to memory
    }
  }

  return Array.from(memoryFolders.values())
    .filter(
      (rec) =>
        rec.userId === ownerScope &&
        (rec.ownerScope ?? rec.userId) === ownerScope
    )
    .map((rec) => ({
      ...rec.folder,
      ownerScope,
      userId: rec.folder.userId ?? ownerScope,
    }))
    .sort((a, b) => Date.parse(a.createdAt) - Date.parse(b.createdAt));
}

/**
 * Purges cached IndexedDB & in-memory records across all stores (`documents`, `drafts`,
 * `pendingUploads`, `versions`, `originalBinaries`) for the logged-out `userId` and any
 * non-`'guest'` owner scopes. Never deletes `'guest'` files.
 */
export async function clearUserSessionCaches(
  loggedOutUserId?: string | null
): Promise<void> {
  const targetUser = loggedOutUserId?.trim() || null;

  const shouldPurgeRecord = (rec: {
    userId?: string;
    ownerScope?: string;
    documentSnapshot?: DocumentItem;
  }): boolean => {
    const scope = (rec.ownerScope ?? rec.userId ?? '').trim();
    if (targetUser && (scope === targetUser || rec.userId === targetUser)) {
      return true;
    }
    if (scope && scope !== 'guest') {
      return true;
    }
    if (rec.documentSnapshot) {
      const snapUser = rec.documentSnapshot.userId?.trim();
      const snapScope = rec.documentSnapshot.ownerScope?.trim();
      if (
        (targetUser && (snapUser === targetUser || snapScope === targetUser)) ||
        (snapScope && snapScope !== 'guest') ||
        (snapUser && snapUser !== 'guest') ||
        Boolean(rec.documentSnapshot.isSyncedToCloud)
      ) {
        return true;
      }
    }
    return false;
  };

  for (const [k, rec] of memoryDocuments.entries()) {
    if (shouldPurgeRecord(rec)) memoryDocuments.delete(k);
  }
  for (const [k, rec] of memoryFolders.entries()) {
    if (shouldPurgeRecord(rec)) memoryFolders.delete(k);
  }
  for (const [k, rec] of memoryDrafts.entries()) {
    if (shouldPurgeRecord(rec)) memoryDrafts.delete(k);
  }
  for (const [k, rec] of memoryPending.entries()) {
    if (shouldPurgeRecord(rec)) memoryPending.delete(k);
  }
  for (const [k, rec] of memoryVersions.entries()) {
    if (shouldPurgeRecord(rec)) memoryVersions.delete(k);
  }
  for (const [k, rec] of memoryOriginalBinaries.entries()) {
    if (shouldPurgeRecord(rec)) memoryOriginalBinaries.delete(k);
  }

  const db = await openOfflineDb();
  if (!db) {
    return;
  }

  try {
    const existingStores = ALL_STORES.filter((name) =>
      db.objectStoreNames.contains(name)
    );
    if (existingStores.length === 0) {
      return;
    }
    const tx = db.transaction([...existingStores], 'readwrite');
    for (const storeName of existingStores) {
      const store = tx.objectStore(storeName);
      const cursorReq = store.openCursor();
      cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (!cursor) {
          return;
        }
        const value = cursor.value as {
          userId?: string;
          ownerScope?: string;
          documentSnapshot?: DocumentItem;
        };
        if (shouldPurgeRecord(value)) {
          cursor.delete();
        }
        cursor.continue();
      };
    }
    await waitForTransaction(tx);
  } catch {
    // Ignore IndexedDB cleanup errors
  }
}

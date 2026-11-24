import { useCallback, useEffect, useState } from 'react';
import {
  CheckCircle2,
  Clock,
  FileText,
  History,
  Loader2,
  RotateCcw,
  Save,
  User,
  X,
} from 'lucide-react';
import type { DocumentItem } from '../../types/document';
import {
  listDocumentVersions,
  saveDocumentVersion,
  type DocumentVersionRecord,
} from '../../services/offlineSyncQueue';

export interface VersionHistoryModalProps {
  isOpen: boolean;
  document: DocumentItem | null;
  userId?: string | null;
  savedByName?: string | null;
  onClose: () => void;
  onRestoreVersion: (
    restoredDoc: DocumentItem,
    version: DocumentVersionRecord
  ) => Promise<void> | void;
  onSaveCheckpoint?: (doc: DocumentItem) => Promise<void> | void;
}

const DATE_TIME_FORMATTER = new Intl.DateTimeFormat('vi-VN', {
  dateStyle: 'medium',
  timeStyle: 'short',
});

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

function formatVersionTimestamp(isoString: string): string {
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) {
    return 'Vừa lưu';
  }
  return DATE_TIME_FORMATTER.format(date);
}

export function VersionHistoryModal({
  isOpen,
  document,
  userId = null,
  savedByName = null,
  onClose,
  onRestoreVersion,
  onSaveCheckpoint,
}: VersionHistoryModalProps) {
  const [versions, setVersions] = useState<DocumentVersionRecord[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isSavingCheckpoint, setIsSavingCheckpoint] = useState(false);
  const [restoringId, setRestoringId] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const isBusy = isSavingCheckpoint || restoringId !== null;

  const refreshVersions = useCallback(async () => {
    if (!document) {
      setVersions([]);
      return;
    }
    setIsLoading(true);
    setErrorMessage(null);
    try {
      const list = await listDocumentVersions(userId, document.id);
      setVersions(list);
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Không thể tải danh sách lịch sử phiên bản.'
      );
    } finally {
      setIsLoading(false);
    }
  }, [document, userId]);

  useEffect(() => {
    if (!isOpen) {
      setStatusMessage(null);
      setErrorMessage(null);
      return;
    }
    void refreshVersions();
  }, [isOpen, refreshVersions]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    const handleKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape' && !isBusy) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isBusy, isOpen, onClose]);

  if (!isOpen || !document) {
    return null;
  }

  const handleCreateCheckpoint = async () => {
    if (isBusy) {
      return;
    }
    setIsSavingCheckpoint(true);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      if (onSaveCheckpoint) {
        await onSaveCheckpoint(document);
      } else {
        await saveDocumentVersion(userId, document, savedByName);
      }
      await refreshVersions();
      setStatusMessage('Đã lưu mốc phiên bản mới cho tài liệu này.');
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Không thể lưu mốc phiên bản hiện tại.'
      );
    } finally {
      setIsSavingCheckpoint(false);
    }
  };

  const handleRestore = async (version: DocumentVersionRecord) => {
    if (isBusy) {
      return;
    }
    setRestoringId(version.id);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      await onRestoreVersion(version.documentSnapshot, version);
      setStatusMessage(
        `Đã khôi phục về mốc ${formatVersionTimestamp(version.savedAt)}.`
      );
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Không thể khôi phục phiên bản đã chọn.'
      );
    } finally {
      setRestoringId(null);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="version-history-modal-title"
      onClick={() => {
        if (!isBusy) {
          onClose();
        }
      }}
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-6">
        <div
          onClick={(event) => event.stopPropagation()}
          className="flex max-h-full w-full max-w-2xl flex-col overflow-hidden rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] text-slate-900 shadow-2xl sm:rounded-3xl"
        >
          {/* Header */}
          <header className="flex shrink-0 items-center justify-between gap-3 border-b border-[#EAE5DC] bg-white px-5 py-4 sm:px-6">
            <div className="flex min-w-0 items-center gap-3">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-indigo-50 text-[#4F46E5]">
                <History className="h-5 w-5" aria-hidden="true" />
              </div>
              <div className="min-w-0">
                <h2
                  id="version-history-modal-title"
                  className="truncate text-base font-bold text-slate-900 sm:text-lg"
                >
                  Lịch sử phiên bản
                </h2>
                <p className="truncate text-xs text-slate-500">
                  {document.title} · Lưu tối đa 10 mốc gần nhất trên thiết bị
                </p>
              </div>
            </div>

            <div className="flex shrink-0 items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  void handleCreateCheckpoint();
                }}
                disabled={isBusy}
                className="inline-flex min-h-[38px] items-center gap-1.5 rounded-xl bg-[#4F46E5] px-3 py-1.5 text-xs font-semibold text-white shadow-sm shadow-indigo-600/20 transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-60"
              >
                {isSavingCheckpoint ? (
                  <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />
                ) : (
                  <Save className="h-3.5 w-3.5" aria-hidden="true" />
                )}
                <span>Lưu mốc ngay</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  if (!isBusy) {
                    onClose();
                  }
                }}
                disabled={isBusy}
                aria-label="Đóng lịch sử phiên bản"
                className="flex h-9 w-9 items-center justify-center rounded-xl text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-700 disabled:opacity-50"
              >
                <X className="h-4 w-4" aria-hidden="true" />
              </button>
            </div>
          </header>

          {/* Body */}
          <div className="min-h-0 flex-1 overflow-y-auto px-5 py-4 sm:px-6">
            {statusMessage && (
              <div className="mb-3 flex items-center gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3.5 py-2.5 text-xs font-medium text-emerald-800">
                <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-600" />
                <span>{statusMessage}</span>
              </div>
            )}

            {errorMessage && (
              <div className="mb-3 rounded-xl border border-rose-200 bg-rose-50 px-3.5 py-2.5 text-xs font-medium text-rose-700">
                {errorMessage}
              </div>
            )}

            {isLoading ? (
              <div className="flex flex-col items-center justify-center gap-2.5 py-14 text-slate-500">
                <Loader2 className="h-6 w-6 animate-spin text-[#4F46E5]" />
                <p className="text-xs font-medium">
                  Đang tải các mốc phiên bản...
                </p>
              </div>
            ) : versions.length === 0 ? (
              <div className="flex flex-col items-center justify-center rounded-2xl border border-dashed border-[#DDD6CA] bg-white/80 px-6 py-12 text-center">
                <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-indigo-50 text-[#4F46E5]">
                  <Clock className="h-6 w-6" aria-hidden="true" />
                </div>
                <p className="text-sm font-semibold text-slate-900">
                  Chưa có mốc phiên bản nào được lưu
                </p>
                <p className="mt-1 max-w-sm text-xs leading-relaxed text-slate-500">
                  Bấm <strong>Lưu mốc ngay</strong> hoặc nhấn{' '}
                  <strong>Lưu thay đổi</strong> trong trình chỉnh sửa để tạo điểm
                  khôi phục an toàn (tối đa 10 phiên bản gần nhất).
                </p>
              </div>
            ) : (
              <ol className="space-y-2.5">
                {versions.map((version, index) => {
                  const isRestoringThis = restoringId === version.id;
                  const isLatest = index === 0;
                  return (
                    <li
                      key={version.id}
                      className="flex flex-col justify-between gap-3 rounded-2xl border border-[#EAE5DC] bg-white p-3.5 shadow-[0_1px_2px_rgba(30,36,48,0.03)] sm:flex-row sm:items-center"
                    >
                      <div className="min-w-0 flex-1 space-y-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-900">
                            <Clock
                              className="h-3.5 w-3.5 text-[#4F46E5]"
                              aria-hidden="true"
                            />
                            {formatVersionTimestamp(version.savedAt)}
                          </span>
                          {version.saveKind === 'auto' ? (
                            <span className="inline-flex items-center rounded-full border border-amber-200 bg-amber-50 px-2 py-0.5 text-[10.5px] font-semibold text-amber-700">
                              Tự động lưu
                            </span>
                          ) : (
                            <span className="inline-flex items-center rounded-full border border-emerald-200 bg-emerald-50 px-2 py-0.5 text-[10.5px] font-semibold text-emerald-700">
                              Lưu thủ công
                            </span>
                          )}
                          {isLatest && (
                            <span className="inline-flex items-center rounded-full bg-indigo-50 px-2 py-0.5 text-[10.5px] font-semibold text-[#4F46E5]">
                              Mới nhất
                            </span>
                          )}
                        </div>

                        <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-slate-500">
                          <span className="inline-flex items-center gap-1 truncate max-w-[240px]">
                            <FileText className="h-3.5 w-3.5 shrink-0 text-slate-400" />
                            <span className="truncate font-medium text-slate-700">
                              {version.title}
                            </span>
                          </span>
                          <span className="inline-flex items-center gap-1">
                            <User className="h-3.5 w-3.5 text-slate-400" />
                            <span>{version.savedByName}</span>
                          </span>
                          <span className="tabular-nums text-slate-400">
                            {formatByteSize(version.fileSizeBytes)}
                          </span>
                        </div>
                      </div>

                      <button
                        type="button"
                        onClick={() => {
                          void handleRestore(version);
                        }}
                        disabled={isBusy}
                        className="inline-flex min-h-[38px] shrink-0 items-center justify-center gap-1.5 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5] px-3 py-1.5 text-xs font-semibold text-slate-800 transition-colors hover:border-[#4F46E5] hover:bg-indigo-50 hover:text-[#4F46E5] disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        {isRestoringThis ? (
                          <Loader2
                            className="h-3.5 w-3.5 animate-spin"
                            aria-hidden="true"
                          />
                        ) : (
                          <RotateCcw className="h-3.5 w-3.5" aria-hidden="true" />
                        )}
                        <span>
                          {isRestoringThis
                            ? 'Đang khôi phục...'
                            : 'Khôi phục phiên bản này'}
                        </span>
                      </button>
                    </li>
                  );
                })}
              </ol>
            )}
          </div>

          {/* Footer */}
          <footer className="flex shrink-0 items-center justify-between gap-3 border-t border-[#EAE5DC] bg-white px-5 py-3 text-xs text-slate-500 sm:px-6">
            <span>
              Hiển thị {versions.length}/10 mốc gần nhất cho tài khoản hiện tại
            </span>
            <button
              type="button"
              onClick={() => {
                if (!isBusy) {
                  onClose();
                }
              }}
              disabled={isBusy}
              className="inline-flex min-h-[36px] items-center justify-center rounded-xl border border-[#E2DDD5] bg-white px-3.5 py-1.5 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-50"
            >
              Đóng
            </button>
          </footer>
        </div>
      </div>
    </div>
  );
}

import { useCallback, useEffect, useState } from 'react';
import type { FormEvent } from 'react';
import {
  AlertCircle,
  BookOpen,
  FileSpreadsheet,
  FileText,
  Loader2,
  Presentation,
  RefreshCw,
  Search,
  ShieldAlert,
  X,
} from 'lucide-react';
import {
  downloadGoogleDriveFile,
  GoogleDriveAuthError,
  listGoogleDriveFiles,
  type GoogleDriveFileItem,
  type GoogleDriveTargetExt,
} from '../../services/googleDrive';
import { formatRelativeVietnameseTime } from '../../services/cloudSync';

export interface GoogleDriveBrowserModalProps {
  isOpen: boolean;
  hasGoogleToken: boolean;
  onClose: () => void;
  onConnectGoogle: () => void;
  onOpenDriveFile: (
    file: File,
    driveItem: GoogleDriveFileItem
  ) => Promise<void> | void;
}

const EXT_STYLE_MAP: Record<
  GoogleDriveTargetExt,
  { label: string; badgeClass: string; iconBg: string; icon: typeof FileText }
> = {
  docx: {
    label: 'DOCX',
    badgeClass: 'bg-blue-50 text-blue-700 border-blue-200/80',
    iconBg: 'bg-blue-50 text-blue-600 border-blue-100',
    icon: FileText,
  },
  xlsx: {
    label: 'XLSX',
    badgeClass: 'bg-green-100 text-green-800 border-green-300/80',
    iconBg: 'bg-emerald-50 text-emerald-600 border-emerald-100',
    icon: FileSpreadsheet,
  },
  pptx: {
    label: 'PPTX',
    badgeClass: 'bg-orange-50 text-orange-700 border-orange-200/80',
    iconBg: 'bg-orange-50 text-orange-600 border-orange-100',
    icon: Presentation,
  },
  pdf: {
    label: 'PDF',
    badgeClass: 'bg-pink-50 text-pink-700 border-pink-200/80',
    iconBg: 'bg-pink-50 text-pink-600 border-pink-100',
    icon: BookOpen,
  },
};

function formatDriveFileSize(bytes?: number, isWorkspace?: boolean): string {
  if (isWorkspace) {
    return 'Google Workspace';
  }
  if (typeof bytes !== 'number' || !Number.isFinite(bytes) || bytes <= 0) {
    return 'Không rõ dung lượng';
  }
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function getWorkspaceTypeLabel(mimeType: string): string | null {
  if (mimeType === 'application/vnd.google-apps.document') {
    return 'Google Docs';
  }
  if (mimeType === 'application/vnd.google-apps.spreadsheet') {
    return 'Google Sheets';
  }
  if (mimeType === 'application/vnd.google-apps.presentation') {
    return 'Google Slides';
  }
  return null;
}

function GoogleDriveTricolorIcon({ className }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 87.3 78"
      className={className}
      aria-hidden="true"
    >
      <path
        d="M6.6 66.85l3.85 6.65c.8 1.4 1.95 2.5 3.3 3.3l13.75-23.8H0c0 1.55.4 3.1 1.2 4.5l5.4 9.35z"
        fill="#0066DA"
      />
      <path
        d="M43.65 25L29.9 1.2c-1.35.8-2.5 1.9-3.3 3.3L1.2 48.5A9.06 9.06 0 0 0 0 53h27.5l16.15-28z"
        fill="#00AC47"
      />
      <path
        d="M73.55 76.8c1.35-.8 2.5-1.9 3.3-3.3l1.6-2.75 7.65-13.25c.8-1.4 1.2-2.95 1.2-4.5H59.8l5.85 11.5 7.9 12.3z"
        fill="#EA4335"
      />
      <path
        d="M43.65 25L57.4 1.2C56.05.4 54.5 0 52.9 0H34.4c-1.6 0-3.15.45-4.5 1.2L43.65 25z"
        fill="#00832D"
      />
      <path
        d="M59.8 53H27.5L13.75 76.8c1.35.8 2.9 1.2 4.5 1.2h50.8c1.6 0 3.15-.45 4.5-1.2L59.8 53z"
        fill="#2684FC"
      />
      <path
        d="M73.4 26.5l-12.7-22c-.8-1.4-1.95-2.5-3.3-3.3L43.65 25 59.8 53h27.45c0-1.55-.4-3.1-1.2-4.5L73.4 26.5z"
        fill="#FFBA00"
      />
    </svg>
  );
}

export function GoogleDriveBrowserModal({
  isOpen,
  hasGoogleToken,
  onClose,
  onConnectGoogle,
  onOpenDriveFile,
}: GoogleDriveBrowserModalProps) {
  const [searchInput, setSearchInput] = useState<string>('');
  const [appliedQuery, setAppliedQuery] = useState<string>('');
  const [files, setFiles] = useState<GoogleDriveFileItem[]>([]);
  const [isLoadingList, setIsLoadingList] = useState<boolean>(false);
  const [downloadingId, setDownloadingId] = useState<string | null>(null);
  const [authError, setAuthError] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const loadFiles = useCallback(async (queryText: string) => {
    setIsLoadingList(true);
    setErrorMessage(null);
    setAuthError(null);

    try {
      const results = await listGoogleDriveFiles(queryText);
      setFiles(results);
    } catch (error) {
      if (error instanceof GoogleDriveAuthError) {
        setAuthError(error.message);
        setFiles([]);
      } else {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : 'Không thể tải danh sách tệp từ Google Drive.'
        );
      }
    } finally {
      setIsLoadingList(false);
    }
  }, []);

  // Reset state whenever modal opens
  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setSearchInput('');
    setAppliedQuery('');
    setDownloadingId(null);
    setErrorMessage(null);
    setAuthError(null);

    if (hasGoogleToken) {
      void loadFiles('');
    } else {
      setFiles([]);
    }
  }, [isOpen, hasGoogleToken, loadFiles]);

  // Debounced search when user types
  useEffect(() => {
    if (!isOpen || !hasGoogleToken) {
      return;
    }
    const trimmed = searchInput.trim();
    if (trimmed === appliedQuery) {
      return;
    }
    const timer = window.setTimeout(() => {
      setAppliedQuery(trimmed);
      void loadFiles(trimmed);
    }, 360);

    return () => {
      window.clearTimeout(timer);
    };
  }, [isOpen, hasGoogleToken, searchInput, appliedQuery, loadFiles]);

  // Close on Escape when not downloading a file
  useEffect(() => {
    if (!isOpen) {
      return;
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && downloadingId === null) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, downloadingId, onClose]);

  if (!isOpen) {
    return null;
  }

  const needsReconnect = !hasGoogleToken || authError !== null;
  const isBusyDownloading = downloadingId !== null;

  const handleSearchSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!hasGoogleToken || isLoadingList) {
      return;
    }
    const trimmed = searchInput.trim();
    setAppliedQuery(trimmed);
    void loadFiles(trimmed);
  };

  const handleOpenItem = async (item: GoogleDriveFileItem) => {
    if (isBusyDownloading) {
      return;
    }
    setDownloadingId(item.id);
    setErrorMessage(null);

    try {
      const file = await downloadGoogleDriveFile(item);
      await Promise.resolve(onOpenDriveFile(file, item));
    } catch (error) {
      if (error instanceof GoogleDriveAuthError) {
        setAuthError(error.message);
      } else {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : `Không thể mở tệp "${item.name}" từ Google Drive.`
        );
      }
    } finally {
      setDownloadingId(null);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="google-drive-modal-title"
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget && !isBusyDownloading) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-2xl max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          {/* Header */}
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span className="h-9 w-9 shrink-0 rounded-xl bg-white border border-[#E6E0D6] flex items-center justify-center shadow-sm">
                <GoogleDriveTricolorIcon className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h2
                  id="google-drive-modal-title"
                  className="text-sm font-semibold text-slate-900 truncate"
                >
                  Duyệt Google Drive
                </h2>
                <p className="text-xs text-slate-500 truncate mt-0.5">
                  Mở trực tiếp Google Docs, Sheets, Slides, Office và PDF vào
                  LITE
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              disabled={isBusyDownloading}
              aria-label="Đóng trình duyệt Google Drive"
              className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Body */}
          {needsReconnect ? (
            <div className="flex-1 min-h-0 overflow-y-auto p-6 sm:p-8 flex flex-col items-center justify-center text-center space-y-4">
              <div className="h-14 w-14 rounded-2xl bg-[#FAF8F5] border border-[#E6E0D6] flex items-center justify-center shadow-sm">
                {authError ? (
                  <ShieldAlert className="h-7 w-7 text-amber-600" />
                ) : (
                  <GoogleDriveTricolorIcon className="h-8 w-8" />
                )}
              </div>

              <div className="max-w-md space-y-1.5">
                <h3 className="text-base font-bold text-slate-900">
                  {authError
                    ? 'Cần kết nối lại Google Drive'
                    : 'Kết nối tài khoản Google Drive'}
                </h3>
                <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                  {authError ??
                    'Liên kết Google để duyệt, mở trực tiếp Google Docs / Sheets / Slides / PDF từ Drive và lưu bản chỉnh sửa trở lại Google Drive.'}
                </p>
              </div>

              <button
                type="button"
                onClick={onConnectGoogle}
                className="min-h-[44px] px-5 rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold shadow-sm shadow-indigo-600/20 inline-flex items-center justify-center gap-2 transition-colors"
              >
                <GoogleDriveTricolorIcon className="h-4 w-4" />
                <span>Kết nối Google Drive ngay</span>
              </button>
            </div>
          ) : (
            <>
              {/* Search & Refresh Strip */}
              <div className="shrink-0 px-4 sm:px-6 py-3 bg-white border-b border-[#F0ECE4] flex items-center gap-2">
                <form
                  onSubmit={handleSearchSubmit}
                  className="flex-1 min-w-0 relative flex items-center"
                >
                  <Search
                    className="h-4 w-4 text-slate-400 absolute left-3 pointer-events-none"
                    aria-hidden="true"
                  />
                  <input
                    type="text"
                    value={searchInput}
                    disabled={isBusyDownloading}
                    onChange={(event) => setSearchInput(event.target.value)}
                    placeholder="Tìm tên tệp trên Google Drive..."
                    aria-label="Tìm kiếm tệp trên Google Drive"
                    className="w-full h-10 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/70 pl-9 pr-9 text-xs sm:text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none focus:ring-2 focus:ring-indigo-500/15"
                  />
                  {searchInput.length > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        setSearchInput('');
                        setAppliedQuery('');
                        void loadFiles('');
                      }}
                      aria-label="Xóa từ khóa tìm kiếm"
                      className="absolute right-2 h-7 w-7 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 flex items-center justify-center"
                    >
                      <X className="h-3.5 w-3.5" />
                    </button>
                  )}
                </form>

                <button
                  type="button"
                  onClick={() => void loadFiles(searchInput.trim())}
                  disabled={isLoadingList || isBusyDownloading}
                  title="Làm mới danh sách Google Drive"
                  aria-label="Làm mới danh sách Google Drive"
                  className="shrink-0 h-10 w-10 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5] hover:bg-[#EFECE6] text-slate-700 flex items-center justify-center transition-colors disabled:opacity-50"
                >
                  <RefreshCw
                    className={`h-4 w-4 ${isLoadingList ? 'animate-spin' : ''}`}
                  />
                </button>
              </div>

              {/* File List */}
              <div className="flex-1 min-h-[280px] overflow-y-auto p-4 sm:p-6 space-y-3">
                {errorMessage && (
                  <div
                    role="alert"
                    className="rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-2.5 flex items-start gap-2 text-xs text-rose-800"
                  >
                    <AlertCircle className="h-4 w-4 text-rose-600 shrink-0 mt-0.5" />
                    <span className="leading-relaxed">{errorMessage}</span>
                  </div>
                )}

                {isLoadingList ? (
                  <div className="py-12 flex flex-col items-center justify-center gap-2.5 text-slate-500">
                    <Loader2 className="h-6 w-6 animate-spin text-[#4F46E5]" />
                    <p className="text-xs font-medium">
                      Đang tải danh sách tệp từ Google Drive...
                    </p>
                  </div>
                ) : files.length === 0 ? (
                  <div className="py-12 text-center space-y-2">
                    <p className="text-sm font-semibold text-slate-800">
                      Không tìm thấy tài liệu phù hợp
                    </p>
                    <p className="text-xs text-slate-500 max-w-sm mx-auto leading-relaxed">
                      {appliedQuery
                        ? `Không có tệp nào khớp với "${appliedQuery}".`
                        : 'Chỉ hiển thị Google Docs, Sheets, Slides và các tệp .docx, .xlsx, .xls, .csv, .pptx, .pdf trên Drive.'}
                    </p>
                  </div>
                ) : (
                  <div className="space-y-2">
                    {files.map((item) => {
                      const extMeta = EXT_STYLE_MAP[item.targetExtension];
                      const ExtIcon = extMeta.icon;
                      const workspaceLabel = getWorkspaceTypeLabel(
                        item.mimeType
                      );
                      const isRowDownloading = downloadingId === item.id;

                      return (
                        <button
                          key={item.id}
                          type="button"
                          disabled={isBusyDownloading}
                          onClick={() => void handleOpenItem(item)}
                          className={`w-full text-left rounded-2xl border p-3 sm:p-3.5 flex items-center justify-between gap-3 transition-all ${
                            isRowDownloading
                              ? 'border-[#4F46E5] bg-indigo-50/60 ring-2 ring-indigo-500/15'
                              : 'border-[#EAE5DC] bg-[#FAF8F5]/70 hover:bg-white hover:border-[#D5CEC2] hover:shadow-sm'
                          } disabled:opacity-65 disabled:cursor-wait`}
                        >
                          <div className="flex items-center gap-3 min-w-0 flex-1">
                            <span
                              className={`h-10 w-10 rounded-xl border flex items-center justify-center shrink-0 ${extMeta.iconBg}`}
                            >
                              <ExtIcon className="h-5 w-5" />
                            </span>

                            <div className="min-w-0 flex-1">
                              <div className="flex items-center gap-2 min-w-0">
                                <span className="text-xs sm:text-sm font-semibold text-slate-900 truncate">
                                  {item.name}
                                </span>
                                <span
                                  className={`shrink-0 inline-flex items-center rounded border px-1.5 py-0 text-[10px] font-semibold ${extMeta.badgeClass}`}
                                >
                                  {extMeta.label}
                                </span>
                              </div>

                              <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[11px] text-slate-500 mt-0.5">
                                {workspaceLabel && (
                                  <>
                                    <span className="font-medium text-indigo-600">
                                      {workspaceLabel}
                                    </span>
                                    <span>•</span>
                                  </>
                                )}
                                <span>
                                  {item.modifiedTime
                                    ? formatRelativeVietnameseTime(
                                        item.modifiedTime
                                      )
                                    : 'Không rõ ngày sửa'}
                                </span>
                                <span>•</span>
                                <span>
                                  {formatDriveFileSize(
                                    item.size,
                                    item.isGoogleWorkspaceType
                                  )}
                                </span>
                              </div>
                            </div>
                          </div>

                          <div className="shrink-0">
                            {isRowDownloading ? (
                              <span className="inline-flex items-center gap-1.5 rounded-xl bg-[#4F46E5] text-white px-3 py-1.5 text-xs font-semibold">
                                <Loader2 className="h-3.5 w-3.5 animate-spin" />
                                <span>Đang mở...</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center rounded-xl border border-[#DDD6CA] bg-white px-3 py-1.5 text-xs font-semibold text-slate-700">
                                Mở tệp
                              </span>
                            )}
                          </div>
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

import { useEffect, useRef, useState } from 'react';
import type { ChangeEvent, DragEvent, KeyboardEvent } from 'react';
import {
  AlertCircle,
  ArrowUpRight,
  BookOpen,
  Check,
  CheckCircle2,
  Cloud,
  FileSearch,
  FileSpreadsheet,
  Folder,
  Loader2,
  Play,
  Plus,
  RotateCcw,
  Search,
  Share2,
  Star,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import type {
  DocumentItem,
  DocumentType,
  FilterKey,
  SidebarNavKey,
} from '../../types/document';

export interface DocumentLibraryProps {
  documents: DocumentItem[];
  activeDocId?: string | null;
  searchQuery: string;
  activeFilter: FilterKey;
  activeNav: SidebarNavKey;
  uploadNotice?: string | null;
  isUploading?: boolean;
  uploadProgress?: number | null;
  onSearchChange: (value: string) => void;
  onFilterChange: (filter: FilterKey) => void;
  onOpenDocument: (doc: DocumentItem) => void;
  onToggleFavorite: (docId: string) => void;
  onUploadFiles: (files: FileList | File[]) => void;
  onDismissNotice?: () => void;
  onDeleteDocuments?: (docs: DocumentItem[]) => Promise<void> | void;
  isDeleting?: boolean;
}

interface FilterOption {
  key: FilterKey;
  label: string;
}

const FILTER_OPTIONS: readonly FilterOption[] = [
  { key: 'all', label: 'Tất cả' },
  { key: 'pptx', label: 'Slide' },
  { key: 'docx', label: 'Word' },
  { key: 'xlsx', label: 'Excel' },
  { key: 'pdf', label: 'PDF' },
  { key: 'favorite', label: 'Yêu thích' },
];

const TYPE_BADGE_MAP: Record<DocumentType, { label: string; className: string }> = {
  pptx: {
    label: 'PPTX',
    className: 'bg-orange-50 text-orange-600 border-orange-200/80',
  },
  docx: {
    label: 'DOCX',
    className: 'bg-blue-50 text-blue-600 border-blue-200/80',
  },
  xlsx: {
    label: 'XLSX',
    className: 'bg-green-100 text-green-800 border-green-300/80',
  },
  pdf: {
    label: 'PDF',
    className: 'bg-pink-50 text-pink-600 border-pink-200/80',
  },
};

function isErrorNotice(message: string): boolean {
  const sanitized = message.replace(/"[^"]*"/g, '');
  return /không hỗ trợ|chưa hỗ trợ|không đúng|lỗi|không hợp lệ|thất bại|bỏ qua|chỉ hỗ trợ|không thể|bị hỏng|từ chối|đầy bộ nhớ/i.test(
    sanitized
  );
}

function getPageCountLabel(doc: DocumentItem): string {
  if (doc.type === 'pptx') {
    const count = doc.pageCount ?? doc.slidesPreview.length;
    return `${count} slide`;
  }
  if (doc.type === 'xlsx') {
    const count =
      doc.sheets && doc.sheets.length > 0 ? doc.sheets.length : (doc.pageCount ?? 1);
    return `${count} sheet`;
  }
  if (doc.type === 'pdf') {
    const count =
      doc.pageCount ??
      (doc.pdfPages && doc.pdfPages.length > 0 ? doc.pdfPages.length : doc.outline.length);
    return count > 0 ? `${count} trang` : 'PDF gốc';
  }
  const count = doc.pageCount ?? doc.outline.length;
  return count > 0 ? `${count} trang` : 'Văn bản';
}

function getPrimaryActionMeta(type: DocumentType): {
  label: string;
  buttonClassName: string;
} {
  if (type === 'pptx') {
    return {
      label: 'Trình chiếu',
      buttonClassName:
        'bg-indigo-600 shadow-indigo-600/15 hover:bg-indigo-700 focus-visible:ring-indigo-500/40',
    };
  }
  if (type === 'xlsx') {
    return {
      label: 'Mở bảng tính',
      buttonClassName:
        'bg-green-700 shadow-green-700/15 hover:bg-green-800 focus-visible:ring-green-600/40',
    };
  }
  return {
    label: 'Đọc',
    buttonClassName:
      'bg-emerald-600 shadow-emerald-600/15 hover:bg-emerald-700 focus-visible:ring-emerald-500/40',
  };
}

export function DocumentLibrary({
  documents,
  activeDocId = null,
  searchQuery,
  activeFilter,
  activeNav,
  uploadNotice,
  isUploading = false,
  uploadProgress = null,
  onSearchChange,
  onFilterChange,
  onOpenDocument,
  onToggleFavorite,
  onUploadFiles,
  onDismissNotice,
  onDeleteDocuments,
  isDeleting = false,
}: DocumentLibraryProps) {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const dragDepthRef = useRef(0);
  const [isDragging, setIsDragging] = useState(false);
  const [isSelectionMode, setIsSelectionMode] = useState(false);
  const [selectedDeleteIds, setSelectedDeleteIds] = useState<string[]>([]);
  const [confirmDeleteDocs, setConfirmDeleteDocs] = useState<DocumentItem[]>([]);

  useEffect(() => {
    if (documents.length === 0 && isSelectionMode) {
      setIsSelectionMode(false);
      setSelectedDeleteIds([]);
    }
  }, [documents.length, isSelectionMode]);

  useEffect(() => {
    if (confirmDeleteDocs.length === 0) return;
    const handleKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape' && !isDeleting) {
        setConfirmDeleteDocs([]);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [confirmDeleteDocs.length, isDeleting]);

  const hasActiveProgress =
    typeof uploadProgress === 'number' && Number.isFinite(uploadProgress);
  const isUploadBusy = isUploading || hasActiveProgress;
  const clampedUploadProgress = hasActiveProgress
    ? Math.max(0, Math.min(100, Math.round(uploadProgress)))
    : null;

  const openFilePicker = () => {
    if (isUploadBusy) return;
    fileInputRef.current?.click();
  };

  const handleFileInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files ? Array.from(event.target.files) : [];
    event.target.value = '';
    if (isUploadBusy || files.length === 0) {
      return;
    }
    onUploadFiles(files);
  };

  const handleDragEnter = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (isUploadBusy) return;
    dragDepthRef.current += 1;
    setIsDragging(true);
  };

  const handleDragOver = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (isUploadBusy) {
      event.dataTransfer.dropEffect = 'none';
      return;
    }
    event.dataTransfer.dropEffect = 'copy';
    if (!isDragging) {
      setIsDragging(true);
    }
  };

  const handleDragLeave = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    if (isUploadBusy) {
      dragDepthRef.current = 0;
      setIsDragging(false);
      return;
    }
    const nextTarget = event.relatedTarget as Node | null;
    if (!nextTarget || !event.currentTarget.contains(nextTarget)) {
      dragDepthRef.current = 0;
      setIsDragging(false);
      return;
    }
    dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
    if (dragDepthRef.current === 0) {
      setIsDragging(false);
    }
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.stopPropagation();
    dragDepthRef.current = 0;
    setIsDragging(false);
    if (isUploadBusy) return;

    const files = event.dataTransfer.files ? Array.from(event.dataTransfer.files) : [];
    if (files.length > 0) {
      onUploadFiles(files);
    }
  };

  const toggleSelectForDelete = (docId: string) => {
    setSelectedDeleteIds((prev) =>
      prev.includes(docId) ? prev.filter((id) => id !== docId) : [...prev, docId]
    );
  };

  const selectedIdSet = new Set(selectedDeleteIds);
  const selectedVisibleDocs = documents.filter((doc) => selectedIdSet.has(doc.id));
  const allVisibleSelected =
    documents.length > 0 && selectedVisibleDocs.length === documents.length;

  const handleToggleSelectAllVisible = () => {
    if (allVisibleSelected) {
      setSelectedDeleteIds([]);
    } else {
      setSelectedDeleteIds(documents.map((doc) => doc.id));
    }
  };

  const handleCancelSelectionMode = () => {
    setIsSelectionMode(false);
    setSelectedDeleteIds([]);
  };

  const handleConfirmDelete = async () => {
    if (confirmDeleteDocs.length === 0 || isDeleting) return;
    const targets = confirmDeleteDocs;
    try {
      await onDeleteDocuments?.(targets);
      const deletedIds = new Set(targets.map((item) => item.id));
      setSelectedDeleteIds((prev) => prev.filter((id) => !deletedIds.has(id)));
      setConfirmDeleteDocs([]);
      setIsSelectionMode(false);
    } catch {
      setConfirmDeleteDocs([]);
    }
  };

  const handleCardKeyDown = (event: KeyboardEvent<HTMLElement>, doc: DocumentItem) => {
    if (event.target !== event.currentTarget) return;
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      if (isSelectionMode) {
        toggleSelectForDelete(doc.id);
      } else {
        onOpenDocument(doc);
      }
    }
  };

  const handleResetFilters = () => {
    onSearchChange('');
    onFilterChange('all');
  };

  const hasActiveFilterOrSearch = searchQuery.trim().length > 0 || activeFilter !== 'all';

  const uniqueFolders = Array.from(
    new Set(documents.map((doc) => doc.folder).filter((folder): folder is string => Boolean(folder)))
  );

  const noticeIsError = uploadNotice ? isErrorNotice(uploadNotice) : false;

  return (
    <main className="h-full min-w-0 flex-1 overflow-y-auto bg-[#FAF8F5] px-4 py-4 pb-24 md:px-6 md:py-5 md:pb-6">
      <div className="mx-auto flex w-full max-w-[1600px] flex-col gap-4">
        {/* Hidden real file input */}
        <input
          ref={fileInputRef}
          type="file"
          accept=".pptx,.docx,.doc,.xlsx,.xls,.csv,.pdf"
          multiple
          disabled={isUploadBusy}
          onChange={handleFileInputChange}
          className="hidden"
          aria-hidden="true"
          tabIndex={-1}
        />

        {/* Top bar: Search + Upload CTA */}
        <header className="flex items-center gap-2.5">
          <div className="relative flex h-11 min-w-0 flex-1 items-center rounded-xl border border-[#E6E1D8] bg-white shadow-[0_1px_2px_rgba(30,36,48,0.03)] transition-all duration-150 focus-within:border-indigo-400 focus-within:ring-2 focus-within:ring-indigo-500/15">
            <Search
              className="pointer-events-none ml-3.5 h-4 w-4 shrink-0 text-slate-400"
              aria-hidden="true"
            />
            <input
              type="text"
              value={searchQuery}
              onChange={(event) => onSearchChange(event.target.value)}
              placeholder="Tìm kiếm tài liệu..."
              aria-label="Tìm kiếm tài liệu"
              className="h-full w-full min-w-0 bg-transparent pl-2.5 pr-9 text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none"
            />
            {searchQuery.length > 0 && (
              <button
                type="button"
                onClick={() => onSearchChange('')}
                aria-label="Xóa từ khóa tìm kiếm"
                className="absolute right-1.5 flex h-8 w-8 items-center justify-center rounded-lg text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-700"
              >
                <X className="h-3.5 w-3.5" />
              </button>
            )}
          </div>

          <button
            type="button"
            onClick={openFilePicker}
            disabled={isUploadBusy}
            aria-busy={isUploadBusy}
            className="inline-flex h-11 shrink-0 items-center gap-1.5 whitespace-nowrap rounded-xl bg-indigo-600 px-3.5 text-sm font-medium text-white shadow-sm shadow-indigo-600/20 transition-all duration-150 hover:bg-indigo-700 active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 disabled:cursor-not-allowed disabled:opacity-75 disabled:hover:bg-indigo-600 md:gap-2 md:px-4"
          >
            {isUploadBusy ? (
              <>
                <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                {clampedUploadProgress !== null ? (
                  <>
                    <span className="tabular-nums sm:hidden">{clampedUploadProgress}%</span>
                    <span className="hidden tabular-nums sm:inline">
                      Đang tải lên... {clampedUploadProgress}%
                    </span>
                  </>
                ) : (
                  <>
                    <span className="sm:hidden">Đang đọc...</span>
                    <span className="hidden sm:inline">Đang đọc nội dung tệp...</span>
                  </>
                )}
              </>
            ) : (
              <>
                <Plus className="h-4 w-4 md:hidden" aria-hidden="true" />
                <Upload className="hidden h-4 w-4 md:block" aria-hidden="true" />
                <span>Tải lên</span>
              </>
            )}
          </button>
        </header>

        {/* Quick upload dropzone (hidden on mobile; progress & notice always visible when active) */}
        <section
          aria-label="Tải lên nhanh"
          className={uploadNotice || clampedUploadProgress !== null ? undefined : 'hidden md:block'}
        >
          <div
            onDragEnter={handleDragEnter}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={openFilePicker}
            aria-busy={isUploadBusy}
            aria-disabled={isUploadBusy}
            className={`group hidden items-center justify-between gap-3 rounded-2xl border-2 border-dashed px-4 py-3 transition-all duration-150 md:flex ${
              isUploadBusy
                ? 'cursor-wait border-indigo-300 bg-indigo-50/50 opacity-85'
                : isDragging
                  ? 'cursor-pointer border-indigo-500 bg-indigo-50/80 ring-2 ring-indigo-500/20'
                  : 'cursor-pointer border-[#DDD6CA] bg-white/70 hover:border-indigo-400 hover:bg-white'
            }`}
          >
            <div
              className={`flex min-w-0 items-center gap-3 ${
                isDragging || isUploadBusy ? 'pointer-events-none' : ''
              }`}
            >
              <div
                className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-xl transition-colors duration-150 ${
                  isUploadBusy
                    ? 'bg-indigo-100 text-indigo-600'
                    : isDragging
                      ? 'bg-indigo-600 text-white'
                      : 'bg-[#F4F0E8] text-slate-500 group-hover:bg-indigo-50 group-hover:text-indigo-600'
                }`}
              >
                {isUploadBusy ? (
                  <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                ) : (
                  <Upload className="h-4 w-4" aria-hidden="true" />
                )}
              </div>

              <p className="truncate text-sm text-slate-600">
                <span>
                  {clampedUploadProgress !== null
                    ? `Đang tải tệp lên kho lưu trữ... ${clampedUploadProgress}%`
                    : isUploadBusy
                      ? 'Đang đọc nội dung tệp...'
                      : isDragging
                        ? 'Thả tệp tại đây để thêm vào không gian làm việc'
                        : 'Kéo và thả tệp vào đây hoặc '}
                </span>
                {!isDragging && !isUploadBusy && (
                  <button
                    type="button"
                    onClick={(event) => {
                      event.stopPropagation();
                      openFilePicker();
                    }}
                    className="font-semibold text-indigo-600 underline decoration-indigo-300 underline-offset-2 transition-colors hover:text-indigo-700 hover:decoration-indigo-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
                  >
                    Chọn tệp
                  </button>
                )}
              </p>
            </div>

            <span
              className={`hidden shrink-0 rounded-lg bg-[#F5F2EC] px-2.5 py-1 text-[11px] font-medium text-slate-500 sm:inline-flex ${
                isDragging || isUploadBusy ? 'pointer-events-none' : ''
              }`}
            >
              Hỗ trợ PPTX, DOCX, XLSX, XLS, CSV, PDF
            </span>
          </div>

          {clampedUploadProgress !== null && (
            <div
              role="progressbar"
              aria-valuenow={clampedUploadProgress}
              aria-valuemin={0}
              aria-valuemax={100}
              aria-label="Tiến trình tải tệp lên kho lưu trữ"
              className="rounded-xl border border-indigo-200/80 bg-white px-3.5 py-2.5 shadow-[0_1px_2px_rgba(79,70,229,0.06)] md:mt-2.5"
            >
              <div className="mb-1.5 flex items-center justify-between gap-2 text-xs font-medium text-indigo-700">
                <span className="truncate">Đang tải tệp lên kho lưu trữ...</span>
                <span className="shrink-0 tabular-nums font-semibold">
                  {clampedUploadProgress}%
                </span>
              </div>
              <div className="h-1.5 w-full overflow-hidden rounded-full bg-indigo-100">
                <div
                  className="h-full rounded-full bg-indigo-600 transition-[width] duration-200 ease-out"
                  style={{ width: `${clampedUploadProgress}%` }}
                />
              </div>
            </div>
          )}

          {uploadNotice && (
            <div
              role="status"
              className={`flex items-center justify-between gap-2 rounded-xl border px-3.5 py-2 text-xs font-medium transition-all ${
                clampedUploadProgress !== null ? 'mt-2.5' : 'md:mt-2.5'
              } ${
                noticeIsError
                  ? 'border-rose-200/90 bg-rose-50/90 text-rose-700'
                  : 'border-indigo-200/80 bg-indigo-50/80 text-indigo-700'
              }`}
            >
              <div className="flex min-w-0 items-center gap-2">
                {noticeIsError ? (
                  <AlertCircle className="h-4 w-4 shrink-0" aria-hidden="true" />
                ) : (
                  <CheckCircle2 className="h-4 w-4 shrink-0" aria-hidden="true" />
                )}
                <span className="min-w-0 leading-snug md:truncate">{uploadNotice}</span>
              </div>

              {onDismissNotice && (
                <button
                  type="button"
                  onClick={onDismissNotice}
                  aria-label="Đóng thông báo"
                  className="shrink-0 rounded-md p-1.5 opacity-70 transition-opacity hover:opacity-100"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
          )}
        </section>

        {/* Filter pills + subtle context line */}
        <section aria-label="Bộ lọc tài liệu" className="flex flex-col gap-2">
          <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
            <div className="-mx-4 flex items-center gap-2 overflow-x-auto whitespace-nowrap px-4 no-scrollbar md:mx-0 md:flex-wrap md:overflow-visible md:px-0">
              {FILTER_OPTIONS.map((option) => {
                const isActive = activeFilter === option.key;

                return (
                  <button
                    key={option.key}
                    type="button"
                    onClick={() => onFilterChange(option.key)}
                    aria-pressed={isActive}
                    className={`inline-flex min-h-[38px] shrink-0 items-center gap-1.5 rounded-full border px-3.5 py-1.5 text-xs font-medium transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/35 ${
                      isActive
                        ? 'border-[#1E2430] bg-[#1E2430] text-white shadow-sm'
                        : 'border-[#E6E1D8] bg-white text-slate-600 hover:border-[#CFC8BC] hover:bg-[#FDFCFA] hover:text-slate-900'
                    }`}
                  >
                    {option.key === 'favorite' && (
                      <Star
                        className={`h-3.5 w-3.5 ${
                          isActive ? 'fill-amber-400 text-amber-400' : 'text-amber-500'
                        }`}
                        aria-hidden="true"
                      />
                    )}
                    <span>{option.label}</span>
                  </button>
                );
              })}
            </div>

            <div className="flex flex-wrap items-center justify-between gap-2 md:justify-end">
              <span className="text-xs font-medium text-slate-400">
                {documents.length} tài liệu
              </span>

              {documents.length > 0 && (
                <div className="flex flex-wrap items-center gap-1.5">
                  {!isSelectionMode ? (
                    <button
                      type="button"
                      onClick={() => {
                        setIsSelectionMode(true);
                        setSelectedDeleteIds([]);
                      }}
                      disabled={isDeleting}
                      className="inline-flex min-h-[38px] items-center gap-1.5 rounded-full border border-[#E6E1D8] bg-white px-3 py-1.5 text-xs font-medium text-slate-600 transition-colors hover:border-rose-200 hover:bg-rose-50/70 hover:text-rose-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-rose-500/35 disabled:opacity-50"
                    >
                      <Trash2 className="h-3.5 w-3.5" aria-hidden="true" />
                      <span>Chọn xóa</span>
                    </button>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={handleToggleSelectAllVisible}
                        disabled={isDeleting}
                        className="inline-flex min-h-[38px] items-center gap-1.5 rounded-full border border-[#E6E1D8] bg-white px-3 py-1.5 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/35 disabled:opacity-50"
                      >
                        <span>{allVisibleSelected ? 'Bỏ chọn' : 'Chọn tất cả'}</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => setConfirmDeleteDocs(selectedVisibleDocs)}
                        disabled={selectedVisibleDocs.length === 0 || isDeleting}
                        className="inline-flex min-h-[38px] items-center gap-1.5 rounded-full bg-rose-600 px-3.5 py-1.5 text-xs font-medium text-white shadow-sm transition-colors hover:bg-rose-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-rose-500/40 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        <Trash2 className="h-3.5 w-3.5" aria-hidden="true" />
                        <span>Xóa đã chọn ({selectedVisibleDocs.length})</span>
                      </button>

                      <button
                        type="button"
                        onClick={handleCancelSelectionMode}
                        disabled={isDeleting}
                        className="inline-flex min-h-[38px] items-center rounded-full border border-[#E6E1D8] bg-white px-3 py-1.5 text-xs font-medium text-slate-600 transition-colors hover:bg-slate-100 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/35 disabled:opacity-50"
                      >
                        <span>Hủy</span>
                      </button>
                    </>
                  )}
                </div>
              )}
            </div>
          </div>

          {activeNav === 'folders' && (
            <div className="flex items-center gap-2 rounded-xl border border-[#EBE6DD] bg-white/75 px-3 py-1.5 text-xs text-slate-600">
              <Folder className="h-3.5 w-3.5 shrink-0 text-indigo-500" aria-hidden="true" />
              <span className="shrink-0 font-medium text-slate-700">Thư mục:</span>
              <div className="flex min-w-0 flex-wrap items-center gap-1.5">
                {uniqueFolders.length > 0 ? (
                  uniqueFolders.map((folderName) => (
                    <span
                      key={folderName}
                      className="rounded-md bg-[#F4F0E8] px-2 py-0.5 text-[11px] font-medium text-slate-600"
                    >
                      {folderName}
                    </span>
                  ))
                ) : (
                  <span className="text-slate-400">Chưa có thư mục phù hợp</span>
                )}
              </div>
            </div>
          )}

          {activeNav === 'shared' && (
            <div className="flex items-center gap-2 rounded-xl border border-[#EBE6DD] bg-white/75 px-3 py-1.5 text-xs text-slate-600">
              <Share2 className="h-3.5 w-3.5 shrink-0 text-indigo-500" aria-hidden="true" />
              <span className="shrink-0 font-medium text-slate-700">Tài liệu đang chia sẻ</span>
              <span className="hidden text-slate-400 sm:inline">•</span>
              <span className="hidden truncate text-slate-500 sm:inline">
                Hiển thị các tài liệu đã bật liên kết chia sẻ nhóm
              </span>
            </div>
          )}
        </section>

        {/* Document cards grid or clean empty state */}
        {documents.length === 0 ? (
          <section
            aria-label="Không có tài liệu"
            className="flex flex-col items-center justify-center rounded-2xl border border-[#EAE5DC] bg-white px-6 py-14 text-center shadow-[0_1px_2px_rgba(30,36,48,0.03)]"
          >
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[#F5F2EC] text-slate-400">
              <FileSearch className="h-5 w-5" aria-hidden="true" />
            </div>
            <h2 className="mt-4 text-[15px] font-semibold text-slate-900">
              {hasActiveFilterOrSearch
                ? 'Không tìm thấy tài liệu phù hợp'
                : 'Chưa có tài liệu nào trong thư viện'}
            </h2>
            <p className="mt-1.5 max-w-sm text-sm text-slate-500">
              {hasActiveFilterOrSearch
                ? 'Thử thay đổi từ khóa tìm kiếm hoặc đặt lại bộ lọc để xem toàn bộ tài liệu.'
                : 'Tải lên tệp PPTX, DOCX, XLSX, XLS, CSV hoặc PDF từ máy tính để bắt đầu làm việc ngay.'}
            </p>
            <div className="mt-5 flex flex-wrap items-center justify-center gap-2.5">
              {hasActiveFilterOrSearch && (
                <button
                  type="button"
                  onClick={handleResetFilters}
                  className="inline-flex min-h-[44px] items-center gap-1.5 rounded-xl border border-[#E2DDD5] bg-white px-4 py-2 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 md:min-h-[38px]"
                >
                  <RotateCcw className="h-3.5 w-3.5" aria-hidden="true" />
                  <span>Đặt lại bộ lọc</span>
                </button>
              )}
              <button
                type="button"
                onClick={openFilePicker}
                disabled={isUploadBusy}
                aria-busy={isUploadBusy}
                className="inline-flex min-h-[44px] items-center gap-1.5 rounded-xl bg-indigo-600 px-4 py-2 text-xs font-medium text-white shadow-sm shadow-indigo-600/20 transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-75 md:min-h-[38px]"
              >
                {isUploadBusy ? (
                  <>
                    <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />
                    <span className="tabular-nums">
                      {clampedUploadProgress !== null
                        ? `Đang tải lên... ${clampedUploadProgress}%`
                        : 'Đang đọc nội dung tệp...'}
                    </span>
                  </>
                ) : (
                  <>
                    <Upload className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Tải tệp lên</span>
                  </>
                )}
              </button>
            </div>
          </section>
        ) : (
          <section
            aria-label="Danh sách tài liệu"
            className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4"
          >
            {documents.map((doc) => {
              const isHighlighted = Boolean(activeDocId && doc.id === activeDocId);
              const isSelectedForDelete = isSelectionMode && selectedIdSet.has(doc.id);
              const badge = TYPE_BADGE_MAP[doc.type];
              const primaryAction = getPrimaryActionMeta(doc.type);

              return (
                <article
                  key={doc.id}
                  role="button"
                  tabIndex={0}
                  aria-label={
                    isSelectionMode
                      ? `${isSelectedForDelete ? 'Bỏ chọn' : 'Chọn'} xóa ${doc.title}`
                      : `Mở ${doc.title}`
                  }
                  onClick={() => {
                    if (isSelectionMode) {
                      toggleSelectForDelete(doc.id);
                    } else {
                      onOpenDocument(doc);
                    }
                  }}
                  onKeyDown={(event) => handleCardKeyDown(event, doc)}
                  className={`group flex cursor-pointer flex-col gap-3 rounded-2xl border bg-[#FFFFFF] p-3.5 transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 ${
                    isSelectedForDelete
                      ? 'border-rose-500 bg-rose-50/15 shadow-[0_6px_20px_rgba(225,29,72,0.08)] ring-2 ring-rose-500/20'
                      : isHighlighted
                        ? 'border-indigo-500 shadow-[0_6px_20px_rgba(79,70,229,0.09)] ring-2 ring-indigo-500/20'
                        : 'border-[#EAE5DC] shadow-[0_1px_2px_rgba(30,36,48,0.04)] hover:border-[#D9D2C5] hover:shadow-[0_8px_20px_rgba(30,36,48,0.06)]'
                  }`}
                >
                  {/* Card header: Title + Favorite star + Trash delete + Format badge */}
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex min-w-0 flex-1 items-center gap-2">
                      {isSelectionMode && (
                        <span
                          aria-hidden="true"
                          className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-md border transition-colors ${
                            isSelectedForDelete
                              ? 'border-rose-600 bg-rose-600 text-white'
                              : 'border-slate-300 bg-white text-transparent'
                          }`}
                        >
                          <Check className="h-3.5 w-3.5" strokeWidth={2.5} />
                        </span>
                      )}
                      <h3
                        title={doc.title}
                        className="min-w-0 flex-1 truncate text-[15px] font-semibold tracking-[-0.01em] text-slate-900"
                      >
                        {doc.title}
                      </h3>
                    </div>

                    <div className="flex shrink-0 items-center gap-1">
                      <button
                        type="button"
                        onClick={(event) => {
                          event.stopPropagation();
                          onToggleFavorite(doc.id);
                        }}
                        aria-label={
                          doc.isFavorite
                            ? `Bỏ yêu thích ${doc.title}`
                            : `Đánh dấu yêu thích ${doc.title}`
                        }
                        aria-pressed={doc.isFavorite}
                        className={`flex h-9 w-9 items-center justify-center rounded-lg transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40 md:h-7 md:w-7 ${
                          doc.isFavorite
                            ? 'text-amber-400 hover:bg-amber-50'
                            : 'text-slate-300 hover:bg-slate-100 hover:text-amber-400'
                        }`}
                      >
                        <Star
                          className={`h-4 w-4 ${doc.isFavorite ? 'fill-amber-400' : ''}`}
                        />
                      </button>

                      <button
                        type="button"
                        onClick={(event) => {
                          event.stopPropagation();
                          setConfirmDeleteDocs([doc]);
                        }}
                        disabled={isDeleting}
                        aria-label={`Xóa tài liệu ${doc.title}`}
                        title="Xóa tài liệu"
                        className="flex h-9 w-9 items-center justify-center rounded-lg text-slate-300 transition-colors duration-150 hover:bg-rose-50 hover:text-rose-600 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-rose-500/40 disabled:opacity-50 md:h-7 md:w-7"
                      >
                        <Trash2 className="h-4 w-4" aria-hidden="true" />
                      </button>

                      <span
                        className={`inline-flex items-center rounded-md border px-2 py-0.5 text-[11px] font-semibold tracking-wide ${badge.className}`}
                      >
                        {badge.label}
                      </span>
                    </div>
                  </div>

                  {/* Card middle: 16:10 thumbnail + subtle metadata */}
                  <div className="flex flex-col gap-2">
                    <div className="aspect-[16/10] w-full overflow-hidden rounded-xl border border-slate-100 bg-[#F8F6F2]">
                      <img
                        src={doc.thumbnail}
                        alt={doc.title}
                        className="h-full w-full object-cover transition-transform duration-300 group-hover:scale-[1.015]"
                      />
                    </div>

                    <div className="flex items-center justify-between gap-2 px-0.5 text-xs text-slate-500">
                      <div className="flex min-w-0 items-center gap-1.5 truncate">
                        {doc.isUploaded === true && (
                          <span className="inline-flex shrink-0 items-center rounded-md border border-indigo-200/80 bg-indigo-50/90 px-1.5 py-0.5 text-[10px] font-semibold leading-none text-indigo-700">
                            Tệp của bạn
                          </span>
                        )}
                        {(doc.isSyncedToCloud === true || Boolean(doc.r2ObjectKey)) && (
                          <span
                            title={
                              doc.r2ObjectKey
                                ? 'Đã đồng bộ tài khoản & Cloudflare R2'
                                : 'Đã đồng bộ lên tài khoản'
                            }
                            aria-label={
                              doc.r2ObjectKey
                                ? 'Đã đồng bộ tài khoản & Cloudflare R2'
                                : 'Đã đồng bộ lên tài khoản'
                            }
                            className="inline-flex shrink-0 items-center text-emerald-600"
                          >
                            <Cloud className="h-3.5 w-3.5" aria-hidden="true" />
                          </span>
                        )}
                        <span className="truncate">{doc.updatedAt}</span>
                        <span aria-hidden="true" className="text-slate-300">
                          •
                        </span>
                        <span className="shrink-0">{getPageCountLabel(doc)}</span>
                      </div>

                      {activeNav === 'folders' && doc.folder ? (
                        <span className="inline-flex shrink-0 items-center gap-1 rounded-md bg-[#F6F3EE] px-1.5 py-0.5 text-[11px] font-medium text-slate-600">
                          <Folder className="h-3 w-3 text-slate-400" aria-hidden="true" />
                          <span className="max-w-[100px] truncate">{doc.folder}</span>
                        </span>
                      ) : activeNav === 'shared' && doc.isShared ? (
                        <span className="inline-flex shrink-0 items-center gap-1 text-[11px] font-medium text-indigo-600">
                          <Share2 className="h-3 w-3" aria-hidden="true" />
                          <span>Đang chia sẻ</span>
                        </span>
                      ) : doc.fileSize ? (
                        <span className="shrink-0 text-[11px] text-slate-400">
                          {doc.fileSize}
                        </span>
                      ) : null}
                    </div>
                  </div>

                  {/* Card footer: Primary action + Secondary Open button */}
                  <div className="mt-auto flex items-center gap-2 pt-0.5">
                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        onOpenDocument(doc);
                      }}
                      className={`inline-flex min-h-[44px] flex-1 items-center justify-center gap-1.5 whitespace-nowrap rounded-xl px-3 py-2 text-xs font-medium text-white shadow-sm transition-all duration-150 active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 ${primaryAction.buttonClassName}`}
                    >
                      {doc.type === 'pptx' ? (
                        <Play className="h-3.5 w-3.5 shrink-0 fill-current" aria-hidden="true" />
                      ) : doc.type === 'xlsx' ? (
                        <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" aria-hidden="true" />
                      ) : (
                        <BookOpen className="h-3.5 w-3.5 shrink-0" aria-hidden="true" />
                      )}
                      <span>{primaryAction.label}</span>
                    </button>

                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        onOpenDocument(doc);
                      }}
                      className="inline-flex min-h-[44px] shrink-0 items-center justify-center gap-1 whitespace-nowrap rounded-xl border border-[#E2DDD5] bg-white px-3 py-2 text-xs font-medium text-slate-700 transition-colors duration-150 hover:bg-slate-50 hover:text-slate-900 active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/35"
                    >
                      <span>Mở</span>
                      <ArrowUpRight className="h-3.5 w-3.5 text-slate-500" aria-hidden="true" />
                    </button>
                  </div>
                </article>
              );
            })}
          </section>
        )}
      </div>

      {/* Delete confirmation modal */}
      {confirmDeleteDocs.length > 0 && (
        <div
          role="dialog"
          aria-modal="true"
          aria-labelledby="delete-confirm-title"
          onClick={() => {
            if (!isDeleting) {
              setConfirmDeleteDocs([]);
            }
          }}
          className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-[2px] pb-safe"
        >
          <div className="flex h-full w-full items-center justify-center p-4">
            <div
              onClick={(event) => event.stopPropagation()}
              className="w-full max-w-md rounded-2xl border border-[#EAE5DC] bg-white p-5 shadow-2xl"
            >
            <div className="flex items-start gap-3.5">
              <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-rose-50 text-rose-600">
                <Trash2 className="h-5 w-5" aria-hidden="true" />
              </div>

              <div className="min-w-0 flex-1">
                <h2
                  id="delete-confirm-title"
                  className="text-base font-semibold text-slate-900"
                >
                  {confirmDeleteDocs.length === 1
                    ? 'Xóa tài liệu khỏi thư viện?'
                    : `Xóa ${confirmDeleteDocs.length} tài liệu đã chọn?`}
                </h2>

                {confirmDeleteDocs.length === 1 ? (
                  <p className="mt-1.5 text-sm text-slate-600">
                    Bạn sắp xóa tệp{' '}
                    <span className="font-semibold text-slate-900">
                      "{confirmDeleteDocs[0].title}"
                    </span>{' '}
                    ({confirmDeleteDocs[0].type.toUpperCase()}).
                  </p>
                ) : (
                  <div className="mt-2 max-h-32 overflow-y-auto rounded-xl border border-[#EAE5DC] bg-[#FAF8F5] px-3 py-2 text-xs text-slate-700">
                    <ul className="space-y-1">
                      {confirmDeleteDocs.map((item) => (
                        <li key={item.id} className="flex items-center justify-between gap-2">
                          <span className="truncate font-medium">{item.title}</span>
                          <span className="shrink-0 text-[10px] font-semibold uppercase text-slate-400">
                            {item.type}
                          </span>
                        </li>
                      ))}
                    </ul>
                  </div>
                )}

                <p className="mt-2.5 text-xs leading-relaxed text-slate-500">
                  Tệp và liên kết chia sẻ liên quan sẽ bị xóa vĩnh viễn khỏi thư viện trên
                  thiết bị này cũng như tài khoản đám mây. Hành động này không thể hoàn tác.
                </p>
              </div>
            </div>

            <div className="mt-5 flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setConfirmDeleteDocs([])}
                disabled={isDeleting}
                className="inline-flex min-h-[44px] items-center justify-center rounded-xl border border-[#E2DDD5] bg-white px-4 py-2 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-50 md:min-h-[38px]"
              >
                Hủy
              </button>

              <button
                type="button"
                onClick={() => {
                  void handleConfirmDelete();
                }}
                disabled={isDeleting}
                className="inline-flex min-h-[44px] items-center justify-center gap-1.5 rounded-xl bg-rose-600 px-4 py-2 text-xs font-semibold text-white shadow-sm transition-colors hover:bg-rose-700 disabled:cursor-not-allowed disabled:opacity-70 md:min-h-[38px]"
              >
                {isDeleting ? (
                  <>
                    <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />
                    <span>Đang xóa...</span>
                  </>
                ) : (
                  <>
                    <Trash2 className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Xóa vĩnh viễn</span>
                  </>
                )}
              </button>
            </div>
          </div>
          </div>
        </div>
      )}
    </main>
  );
}

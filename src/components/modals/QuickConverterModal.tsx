import { useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent } from 'react';
import {
  ArrowDown,
  ArrowUp,
  CheckCircle2,
  Download,
  FileOutput,
  FilePlus2,
  Files,
  Loader2,
  Scissors,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import type { DocumentItem } from '../../types/document';
import { parseUploadedFile } from '../../utils/fileParser';
import {
  downloadDocumentAsPdf,
  inspectPdfPageCount,
  mergePdfFiles,
  parsePageRangeExpression,
  splitPdfFileByRanges,
} from '../../utils/quickConverter';
import { saveFileUniversal } from '../../utils/saveFileUniversal';

export type QuickConverterTab = 'export_pdf' | 'merge_pdf' | 'split_pdf';

export interface QuickConverterModalProps {
  isOpen: boolean;
  onClose: () => void;
  documents?: DocumentItem[];
  initialDocument?: DocumentItem | null;
  initialTab?: QuickConverterTab;
}

interface MergePdfEntry {
  id: string;
  file: File;
  pageCount: number | null;
}

async function documentToPdfFile(doc: DocumentItem): Promise<File | null> {
  if (doc.type !== 'pdf') {
    return null;
  }
  const safeName = doc.originalFileName?.toLowerCase().endsWith('.pdf')
    ? doc.originalFileName
    : `${doc.title || 'tai-lieu'}.pdf`;

  if (doc.sourceFile) {
    return doc.sourceFile;
  }
  if (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
    return new File([doc.rawArrayBuffer.slice(0)], safeName, {
      type: 'application/pdf',
    });
  }
  const url = doc.fileUrl ?? doc.objectUrl;
  if (url) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        return new File([buffer], safeName, { type: 'application/pdf' });
      }
    } catch {
      return null;
    }
  }
  return null;
}

function formatFileBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export function QuickConverterModal({
  isOpen,
  onClose,
  documents = [],
  initialDocument = null,
  initialTab,
}: QuickConverterModalProps) {
  const convertibleDocs = useMemo(
    () => documents.filter((d) => d.type !== 'pdf'),
    [documents]
  );

  const [activeTab, setActiveTab] = useState<QuickConverterTab>('export_pdf');
  const [selectedDocId, setSelectedDocId] = useState<string>('');
  const [localConvertDoc, setLocalConvertDoc] = useState<DocumentItem | null>(null);
  const [isExportingPdf, setIsExportingPdf] = useState(false);

  // Merge PDF state
  const [mergeEntries, setMergeEntries] = useState<MergePdfEntry[]>([]);
  const [mergeOutputTitle, setMergeOutputTitle] = useState('Tai-lieu-gop');
  const [isMerging, setIsMerging] = useState(false);

  // Split PDF state
  const [splitFile, setSplitFile] = useState<File | null>(null);
  const [splitTotalPages, setSplitTotalPages] = useState<number>(0);
  const [splitRangeInput, setSplitRangeInput] = useState<string>('1-2');
  const [isSplitting, setIsSplitting] = useState(false);

  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const localConvertInputRef = useRef<HTMLInputElement | null>(null);
  const mergeInputRef = useRef<HTMLInputElement | null>(null);
  const splitInputRef = useRef<HTMLInputElement | null>(null);

  const isBusy = isExportingPdf || isMerging || isSplitting;

  useEffect(() => {
    if (!isOpen) return;

    setStatusMessage(null);
    setErrorMessage(null);
    setLocalConvertDoc(null);

    const resolvedTab: QuickConverterTab =
      initialTab ??
      (initialDocument?.type === 'pdf' ? 'split_pdf' : 'export_pdf');
    setActiveTab(resolvedTab);

    if (initialDocument && initialDocument.type !== 'pdf') {
      setSelectedDocId(initialDocument.id);
    } else if (convertibleDocs.length > 0) {
      setSelectedDocId(convertibleDocs[0].id);
    } else {
      setSelectedDocId('');
    }

    if (initialDocument?.type === 'pdf') {
      void documentToPdfFile(initialDocument).then(async (pdfFile) => {
        if (!pdfFile) return;
        setSplitFile(pdfFile);
        try {
          const count = await inspectPdfPageCount(pdfFile);
          setSplitTotalPages(count);
          setSplitRangeInput(count > 1 ? `1-${Math.min(3, count)}` : '1');
        } catch {
          setSplitTotalPages(0);
        }
      });
    } else {
      setSplitFile(null);
      setSplitTotalPages(0);
    }
  }, [isOpen, initialDocument, initialTab, convertibleDocs]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !isBusy) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, isBusy, onClose]);

  if (!isOpen) {
    return null;
  }

  const targetExportDoc =
    localConvertDoc ??
    (initialDocument && initialDocument.id === selectedDocId
      ? initialDocument
      : convertibleDocs.find((d) => d.id === selectedDocId) ?? null);

  const handlePickLocalFileToConvert = async (
    event: ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;

    setErrorMessage(null);
    setStatusMessage(null);
    try {
      const parsed = await parseUploadedFile(file);
      if (parsed.error || !parsed.document) {
        setErrorMessage(parsed.error || 'Không thể đọc tệp vừa chọn.');
        return;
      }
      if (parsed.document.type === 'pdf') {
        setErrorMessage('Tệp này đã là PDF. Chuyển sang tab Gộp hoặc Tách PDF.');
        return;
      }
      setLocalConvertDoc(parsed.document);
      setStatusMessage(`Đã nạp tệp "${parsed.document.title}" sẵn sàng xuất PDF.`);
    } catch {
      setErrorMessage('Đã xảy ra lỗi khi đọc tệp từ thiết bị.');
    }
  };

  const handleRunExportPdf = async () => {
    if (!targetExportDoc || isBusy) return;
    setIsExportingPdf(true);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      const result = await downloadDocumentAsPdf(targetExportDoc);
      if (result.method === 'electron' && !result.savedPath) {
        setStatusMessage(result.message);
      } else {
        setStatusMessage(
          `${result.message} (${formatFileBytes(result.blob.size)})`
        );
      }
    } catch (err: unknown) {
      setErrorMessage(
        err instanceof Error
          ? err.message
          : 'Không thể xuất sang định dạng PDF. Vui lòng thử lại.'
      );
    } finally {
      setIsExportingPdf(false);
    }
  };

  const handleAddMergeFiles = async (event: ChangeEvent<HTMLInputElement>) => {
    const picked = event.target.files ? Array.from(event.target.files) : [];
    event.target.value = '';
    if (picked.length === 0) return;

    setErrorMessage(null);
    setStatusMessage(null);

    const added: MergePdfEntry[] = [];
    for (const file of picked) {
      let pageCount: number | null = null;
      try {
        pageCount = await inspectPdfPageCount(file);
      } catch {
        pageCount = null;
      }
      added.push({
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        file,
        pageCount,
      });
    }
    setMergeEntries((prev) => [...prev, ...added]);
  };

  const handleMoveMergeItem = (index: number, delta: -1 | 1) => {
    setMergeEntries((prev) => {
      const nextIdx = index + delta;
      if (nextIdx < 0 || nextIdx >= prev.length) return prev;
      const copy = [...prev];
      const [moved] = copy.splice(index, 1);
      copy.splice(nextIdx, 0, moved);
      return copy;
    });
  };

  const handleRemoveMergeItem = (id: string) => {
    setMergeEntries((prev) => prev.filter((item) => item.id !== id));
  };

  const handleRunMergePdf = async () => {
    if (mergeEntries.length < 2 || isBusy) return;
    setIsMerging(true);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      const result = await mergePdfFiles(
        mergeEntries.map((item) => item.file),
        mergeOutputTitle.trim() || 'Tai-lieu-gop'
      );
      const saveResult = await saveFileUniversal(result.blob, result.fileName, {
        mimeType: 'application/pdf',
        shareTitle: result.fileName,
      });
      if (saveResult.method === 'electron' && !saveResult.savedPath) {
        setStatusMessage(saveResult.message);
      } else {
        setStatusMessage(
          `Đã gộp thành công ${result.totalPages} trang. ${saveResult.message}`
        );
      }
    } catch (err: unknown) {
      setErrorMessage(
        err instanceof Error
          ? err.message
          : 'Không thể gộp các tệp PDF. Vui lòng kiểm tra lại tệp nguồn.'
      );
    } finally {
      setIsMerging(false);
    }
  };

  const handlePickSplitPdf = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;

    setErrorMessage(null);
    setStatusMessage(null);
    setSplitFile(file);
    try {
      const count = await inspectPdfPageCount(file);
      setSplitTotalPages(count);
      setSplitRangeInput(count > 1 ? `1-${Math.min(3, count)}` : '1');
    } catch {
      setSplitTotalPages(0);
      setErrorMessage('Không thể đọc số trang của tệp PDF này.');
    }
  };

  const splitPreviewValidation = (() => {
    if (!splitFile || splitTotalPages <= 0) return null;
    try {
      const zeroBased = parsePageRangeExpression(splitRangeInput, splitTotalPages);
      const displayPages = zeroBased.map((i) => i + 1);
      return {
        valid: true as const,
        pages: displayPages,
      };
    } catch (err: unknown) {
      return {
        valid: false as const,
        message: err instanceof Error ? err.message : 'Khoảng trang không hợp lệ.',
      };
    }
  })();

  const handleRunSplitPdf = async () => {
    if (!splitFile || isBusy) return;
    setIsSplitting(true);
    setErrorMessage(null);
    setStatusMessage(null);
    try {
      const result = await splitPdfFileByRanges(splitFile, splitRangeInput);
      const saveResult = await saveFileUniversal(result.blob, result.fileName, {
        mimeType: 'application/pdf',
        shareTitle: result.fileName,
      });
      if (saveResult.method === 'electron' && !saveResult.savedPath) {
        setStatusMessage(saveResult.message);
      } else {
        setStatusMessage(
          `Đã tách ${result.pageCount}/${result.totalSourcePages} trang. ${saveResult.message}`
        );
      }
    } catch (err: unknown) {
      setErrorMessage(
        err instanceof Error
          ? err.message
          : 'Không thể tách tệp PDF. Vui lòng kiểm tra khoảng trang.'
      );
    } finally {
      setIsSplitting(false);
    }
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="quick-converter-title"
      onClick={() => {
        if (!isBusy) onClose();
      }}
      className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-4">
        <div
          onClick={(event) => event.stopPropagation()}
          className="flex max-h-[90vh] w-full max-w-xl flex-col overflow-hidden rounded-2xl border border-[#EAE5DC] bg-white shadow-2xl"
        >
          {/* Modal Header */}
          <div className="flex shrink-0 items-center justify-between gap-3 border-b border-[#EFECE6] px-5 py-4">
            <div className="flex items-center gap-2.5 min-w-0">
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-indigo-50 text-indigo-600 border border-indigo-200/80">
                <FileOutput className="h-4 w-4" aria-hidden="true" />
              </span>
              <div className="min-w-0">
                <h2
                  id="quick-converter-title"
                  className="truncate text-base font-semibold text-slate-900"
                >
                  Chuyển đổi &amp; Xử lý PDF nhanh
                </h2>
                <p className="truncate text-xs text-slate-500">
                  Xuất Word/Excel/Slide sang PDF, Gộp hoặc Tách PDF ngay trên máy
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              disabled={isBusy}
              aria-label="Đóng hộp thoại chuyển đổi"
              className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-700 disabled:opacity-50"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* 3 Tabs */}
          <div
            role="tablist"
            aria-label="Công cụ chuyển đổi tệp"
            className="grid shrink-0 grid-cols-3 gap-1.5 border-b border-[#EFECE6] bg-[#FAF8F5] px-4 py-2.5"
          >
            {(
              [
                {
                  key: 'export_pdf',
                  label: 'Xuất sang PDF',
                  Icon: FileOutput,
                },
                {
                  key: 'merge_pdf',
                  label: 'Gộp PDF',
                  Icon: Files,
                },
                {
                  key: 'split_pdf',
                  label: 'Tách PDF',
                  Icon: Scissors,
                },
              ] as const
            ).map(({ key, label, Icon }) => {
              const active = activeTab === key;
              return (
                <button
                  key={key}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  onClick={() => {
                    setActiveTab(key);
                    setErrorMessage(null);
                    setStatusMessage(null);
                  }}
                  className={`inline-flex items-center justify-center gap-1.5 rounded-xl px-2.5 py-2 text-xs font-semibold transition-all ${
                    active
                      ? 'bg-white text-indigo-700 shadow-sm ring-1 ring-indigo-500/20'
                      : 'text-slate-600 hover:bg-white/60 hover:text-slate-900'
                  }`}
                >
                  <Icon className="h-3.5 w-3.5 shrink-0" aria-hidden="true" />
                  <span className="truncate">{label}</span>
                </button>
              );
            })}
          </div>

          {/* Scrollable Body */}
          <div className="flex-1 min-h-0 overflow-y-auto px-5 py-4 space-y-4">
            {activeTab === 'export_pdf' && (
              <div className="space-y-4">
                <div>
                  <label
                    htmlFor="converter-doc-select"
                    className="mb-1.5 block text-xs font-semibold text-slate-700"
                  >
                    Chọn tài liệu từ thư viện
                  </label>
                  <select
                    id="converter-doc-select"
                    value={localConvertDoc ? '' : selectedDocId}
                    onChange={(event) => {
                      setLocalConvertDoc(null);
                      setSelectedDocId(event.target.value);
                    }}
                    className="h-10 w-full rounded-xl border border-[#DDD6CA] bg-white px-3 text-sm text-slate-900 focus:border-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                  >
                    {convertibleDocs.length === 0 && !localConvertDoc && (
                      <option value="">Chưa có tệp Word/Excel/Slide nào trong thư viện</option>
                    )}
                    {convertibleDocs.map((doc) => (
                      <option key={doc.id} value={doc.id}>
                        [{doc.type.toUpperCase()}] {doc.title}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="flex items-center gap-3">
                  <div className="h-px flex-1 bg-slate-200" />
                  <span className="text-[11px] font-medium uppercase tracking-wider text-slate-400">
                    hoặc chọn từ thiết bị
                  </span>
                  <div className="h-px flex-1 bg-slate-200" />
                </div>

                <div>
                  <input
                    ref={localConvertInputRef}
                    type="file"
                    accept=".docx,.xlsx,.xls,.csv,.pptx"
                    onChange={(event) => void handlePickLocalFileToConvert(event)}
                    className="hidden"
                  />
                  <button
                    type="button"
                    onClick={() => localConvertInputRef.current?.click()}
                    className="flex w-full items-center justify-center gap-2 rounded-xl border border-dashed border-[#D6CFC3] bg-[#FAF8F5] px-4 py-3 text-xs font-semibold text-slate-700 transition-colors hover:border-indigo-400 hover:bg-indigo-50/40 hover:text-indigo-700"
                  >
                    <Upload className="h-4 w-4" aria-hidden="true" />
                    <span>
                      {localConvertDoc
                        ? `Đang chọn: ${localConvertDoc.title} (${localConvertDoc.type.toUpperCase()})`
                        : 'Chọn tệp .docx / .xlsx / .pptx từ máy để xuất PDF'}
                    </span>
                  </button>
                </div>

                {targetExportDoc && (
                  <div className="rounded-xl border border-[#EAE5DC] bg-[#FAF8F5] p-3.5 text-xs text-slate-700">
                    <div className="flex items-center justify-between gap-2">
                      <span className="font-semibold text-slate-900 truncate">
                        {targetExportDoc.title}
                      </span>
                      <span className="shrink-0 rounded-md bg-white px-2 py-0.5 text-[10px] font-bold uppercase text-indigo-700 border border-indigo-200/80">
                        {targetExportDoc.type} → PDF
                      </span>
                    </div>
                    <p className="mt-1 text-[11.5px] text-slate-500">
                      Hệ thống sẽ dựng bản in sắc nét chuẩn A4 / 16:9 hỗ trợ đầy đủ dấu tiếng Việt.
                    </p>
                  </div>
                )}

                <button
                  type="button"
                  disabled={!targetExportDoc || isBusy}
                  onClick={() => void handleRunExportPdf()}
                  className="inline-flex min-h-[42px] w-full items-center justify-center gap-2 rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-55"
                >
                  {isExportingPdf ? (
                    <>
                      <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                      <span>Đang kết xuất PDF...</span>
                    </>
                  ) : (
                    <>
                      <Download className="h-4 w-4" aria-hidden="true" />
                      <span>Xuất &amp; Tải xuống PDF</span>
                    </>
                  )}
                </button>
              </div>
            )}

            {activeTab === 'merge_pdf' && (
              <div className="space-y-4">
                <input
                  ref={mergeInputRef}
                  type="file"
                  accept=".pdf,application/pdf"
                  multiple
                  onChange={(event) => void handleAddMergeFiles(event)}
                  className="hidden"
                />

                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs font-semibold text-slate-700">
                    Danh sách tệp PDF cần gộp ({mergeEntries.length})
                  </span>
                  <button
                    type="button"
                    onClick={() => mergeInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 rounded-xl border border-indigo-200 bg-indigo-50 px-3 py-1.5 text-xs font-semibold text-indigo-700 transition-colors hover:bg-indigo-100"
                  >
                    <FilePlus2 className="h-3.5 w-3.5" aria-hidden="true" />
                    <span>Thêm tệp PDF</span>
                  </button>
                </div>

                {mergeEntries.length === 0 ? (
                  <button
                    type="button"
                    onClick={() => mergeInputRef.current?.click()}
                    className="flex w-full flex-col items-center justify-center gap-2 rounded-2xl border-2 border-dashed border-[#DDD6CA] bg-[#FAF8F5]/70 px-4 py-8 text-center transition-colors hover:border-indigo-400 hover:bg-white"
                  >
                    <Files className="h-6 w-6 text-slate-400" aria-hidden="true" />
                    <span className="text-xs font-semibold text-slate-700">
                      Chọn từ 2 tệp PDF trở lên để gộp thành 1 tệp duy nhất
                    </span>
                    <span className="text-[11px] text-slate-500">
                      Có thể sắp xếp thứ tự trên/dưới trước khi gộp
                    </span>
                  </button>
                ) : (
                  <div className="space-y-2 max-h-56 overflow-y-auto pr-1">
                    {mergeEntries.map((entry, idx) => (
                      <div
                        key={entry.id}
                        className="flex items-center justify-between gap-2 rounded-xl border border-[#EAE5DC] bg-[#FAF8F5] px-3 py-2"
                      >
                        <div className="min-w-0 flex-1">
                          <p className="truncate text-xs font-semibold text-slate-900">
                            {idx + 1}. {entry.file.name}
                          </p>
                          <p className="text-[11px] text-slate-500 tabular-nums">
                            {entry.pageCount !== null
                              ? `${entry.pageCount} trang`
                              : 'PDF'}{' '}
                            · {formatFileBytes(entry.file.size)}
                          </p>
                        </div>

                        <div className="flex shrink-0 items-center gap-1">
                          <button
                            type="button"
                            onClick={() => handleMoveMergeItem(idx, -1)}
                            disabled={idx === 0}
                            aria-label="Di chuyển lên"
                            className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-100 disabled:opacity-35"
                          >
                            <ArrowUp className="h-3.5 w-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleMoveMergeItem(idx, 1)}
                            disabled={idx === mergeEntries.length - 1}
                            aria-label="Di chuyển xuống"
                            className="flex h-7 w-7 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-100 disabled:opacity-35"
                          >
                            <ArrowDown className="h-3.5 w-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleRemoveMergeItem(entry.id)}
                            aria-label="Xóa khỏi danh sách gộp"
                            className="flex h-7 w-7 items-center justify-center rounded-lg text-slate-400 hover:bg-rose-50 hover:text-rose-600"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                <div>
                  <label
                    htmlFor="merge-output-name"
                    className="mb-1.5 block text-xs font-semibold text-slate-700"
                  >
                    Tên tệp sau khi gộp
                  </label>
                  <input
                    id="merge-output-name"
                    type="text"
                    value={mergeOutputTitle}
                    onChange={(event) => setMergeOutputTitle(event.target.value)}
                    placeholder="Tai-lieu-gop"
                    className="h-10 w-full rounded-xl border border-[#DDD6CA] bg-white px-3 text-sm text-slate-900 focus:border-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                  />
                </div>

                <button
                  type="button"
                  disabled={mergeEntries.length < 2 || isBusy}
                  onClick={() => void handleRunMergePdf()}
                  className="inline-flex min-h-[42px] w-full items-center justify-center gap-2 rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-55"
                >
                  {isMerging ? (
                    <>
                      <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                      <span>Đang gộp tệp PDF...</span>
                    </>
                  ) : (
                    <>
                      <Files className="h-4 w-4" aria-hidden="true" />
                      <span>Gộp &amp; Tải xuống PDF</span>
                    </>
                  )}
                </button>
              </div>
            )}

            {activeTab === 'split_pdf' && (
              <div className="space-y-4">
                <input
                  ref={splitInputRef}
                  type="file"
                  accept=".pdf,application/pdf"
                  onChange={(event) => void handlePickSplitPdf(event)}
                  className="hidden"
                />

                <button
                  type="button"
                  onClick={() => splitInputRef.current?.click()}
                  className="flex w-full items-center justify-between gap-3 rounded-xl border border-dashed border-[#D6CFC3] bg-[#FAF8F5] px-4 py-3 text-left transition-colors hover:border-indigo-400 hover:bg-indigo-50/40"
                >
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-xs font-semibold text-slate-900">
                      {splitFile ? splitFile.name : 'Chọn tệp PDF cần tách trang'}
                    </p>
                    <p className="mt-0.5 text-[11px] text-slate-500 tabular-nums">
                      {splitFile
                        ? `${splitTotalPages > 0 ? `${splitTotalPages} trang · ` : ''}${formatFileBytes(splitFile.size)}`
                        : 'Nhấn để chọn tệp .pdf từ thiết bị'}
                    </p>
                  </div>
                  <span className="shrink-0 rounded-lg bg-white px-2.5 py-1 text-xs font-semibold text-indigo-700 border border-indigo-200">
                    {splitFile ? 'Đổi tệp' : 'Chọn PDF'}
                  </span>
                </button>

                <div>
                  <label
                    htmlFor="split-range-input"
                    className="mb-1.5 block text-xs font-semibold text-slate-700"
                  >
                    Khoảng trang muốn tách (ví dụ: 1-3, 5)
                  </label>
                  <input
                    id="split-range-input"
                    type="text"
                    value={splitRangeInput}
                    onChange={(event) => setSplitRangeInput(event.target.value)}
                    placeholder="1-3, 5, 8-9"
                    className="h-10 w-full rounded-xl border border-[#DDD6CA] bg-white px-3 text-sm text-slate-900 focus:border-indigo-500 focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                  />

                  {splitTotalPages > 1 && (
                    <div className="mt-2 flex flex-wrap gap-1.5">
                      <button
                        type="button"
                        onClick={() => setSplitRangeInput('1')}
                        className="rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100"
                      >
                        Trang đầu (1)
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          setSplitRangeInput(
                            `1-${Math.max(1, Math.ceil(splitTotalPages / 2))}`
                          )
                        }
                        className="rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100"
                      >
                        Nửa đầu (1-{Math.ceil(splitTotalPages / 2)})
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          setSplitRangeInput(
                            `${Math.min(splitTotalPages, Math.ceil(splitTotalPages / 2) + 1)}-${splitTotalPages}`
                          )
                        }
                        className="rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-700 hover:bg-slate-100"
                      >
                        Nửa sau
                      </button>
                    </div>
                  )}

                  {splitPreviewValidation && (
                    <p
                      className={`mt-2 text-xs ${
                        splitPreviewValidation.valid
                          ? 'text-emerald-700'
                          : 'text-rose-600'
                      }`}
                    >
                      {splitPreviewValidation.valid
                        ? `Sẽ tách ${splitPreviewValidation.pages.length} trang: ${splitPreviewValidation.pages.slice(0, 14).join(', ')}${splitPreviewValidation.pages.length > 14 ? '...' : ''}`
                        : splitPreviewValidation.message}
                    </p>
                  )}
                </div>

                <button
                  type="button"
                  disabled={
                    !splitFile ||
                    (splitPreviewValidation !== null &&
                      !splitPreviewValidation.valid) ||
                    isBusy
                  }
                  onClick={() => void handleRunSplitPdf()}
                  className="inline-flex min-h-[42px] w-full items-center justify-center gap-2 rounded-xl bg-indigo-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-indigo-700 disabled:cursor-not-allowed disabled:opacity-55"
                >
                  {isSplitting ? (
                    <>
                      <Loader2 className="h-4 w-4 animate-spin" aria-hidden="true" />
                      <span>Đang tách tệp PDF...</span>
                    </>
                  ) : (
                    <>
                      <Scissors className="h-4 w-4" aria-hidden="true" />
                      <span>Tách &amp; Tải xuống PDF</span>
                    </>
                  )}
                </button>
              </div>
            )}

            {statusMessage && (
              <div
                role="status"
                className="flex items-start gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3.5 py-2.5 text-xs font-medium text-emerald-800"
              >
                <CheckCircle2 className="h-4 w-4 shrink-0 mt-0.5" aria-hidden="true" />
                <span className="break-words">{statusMessage}</span>
              </div>
            )}

            {errorMessage && (
              <div
                role="alert"
                className="rounded-xl border border-rose-200 bg-rose-50 px-3.5 py-2.5 text-xs font-medium text-rose-700"
              >
                {errorMessage}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

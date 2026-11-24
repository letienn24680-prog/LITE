import { useEffect, useMemo, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import {
  BookMarked,
  Check,
  CheckCircle2,
  Clock,
  Cloud,
  Copy,
  FileCheck2,
  FileText,
  Folder,
  HardDrive,
  Info,
  Lock,
  Quote,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  SpellCheck2,
  Unlock,
  Wand2,
  X,
} from 'lucide-react';
import type { DocumentItem, WordDrawingStroke } from '../../types/document';
import type { WordDocumentStats } from './WordInsertTools';

/* ============================================================================
 * Shared Helpers
 * ========================================================================== */

function escapeHtmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ============================================================================
 * 1. WordDrawingOverlay
 * ========================================================================== */

export interface WordDrawingOverlayProps {
  strokes: WordDrawingStroke[];
  activeTool: 'pen' | 'highlighter' | 'eraser' | null;
  strokeColor: string;
  strokeWidth: number;
  disabled?: boolean;
  onChangeStrokes: (next: WordDrawingStroke[]) => void;
}

function buildSvgPathFromNormalizedPoints(
  points: Array<{ x: number; y: number }>
): string {
  if (points.length === 0) {
    return '';
  }
  if (points.length === 1) {
    const p = points[0];
    return `M ${p.x.toFixed(1)} ${p.y.toFixed(1)} L ${(p.x + 0.6).toFixed(1)} ${(p.y + 0.6).toFixed(1)}`;
  }
  return points
    .map(
      (pt, idx) =>
        `${idx === 0 ? 'M' : 'L'} ${pt.x.toFixed(1)} ${pt.y.toFixed(1)}`
    )
    .join(' ');
}

export function WordDrawingOverlay({
  strokes,
  activeTool,
  strokeColor,
  strokeWidth,
  disabled = false,
  onChangeStrokes,
}: WordDrawingOverlayProps) {
  const svgRef = useRef<SVGSVGElement | null>(null);
  const [draftPoints, setDraftPoints] = useState<Array<{ x: number; y: number }>>(
    []
  );
  const isDrawingRef = useRef(false);
  const isErasingRef = useRef(false);

  const isDrawActive =
    !disabled && (activeTool === 'pen' || activeTool === 'highlighter');
  const isEraseActive = !disabled && activeTool === 'eraser';

  useEffect(() => {
    const handleWindowPointerUp = () => {
      isErasingRef.current = false;
    };
    window.addEventListener('pointerup', handleWindowPointerUp);
    window.addEventListener('pointercancel', handleWindowPointerUp);
    return () => {
      window.removeEventListener('pointerup', handleWindowPointerUp);
      window.removeEventListener('pointercancel', handleWindowPointerUp);
    };
  }, []);

  if (strokes.length === 0 && !isDrawActive) {
    return null;
  }

  const toNormalizedPoint = (
    event: ReactPointerEvent<SVGSVGElement>
  ): { x: number; y: number } => {
    const rect = svgRef.current?.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      return { x: 0, y: 0 };
    }
    const nx = Math.max(
      0,
      Math.min(1000, ((event.clientX - rect.left) / rect.width) * 1000)
    );
    const ny = Math.max(
      0,
      Math.min(1000, ((event.clientY - rect.top) / rect.height) * 1000)
    );
    return { x: nx, y: ny };
  };

  const handleSvgPointerDown = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!isDrawActive) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    isDrawingRef.current = true;
    setDraftPoints([toNormalizedPoint(event)]);
  };

  const handleSvgPointerMove = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!isDrawingRef.current || !isDrawActive) {
      return;
    }
    event.preventDefault();
    const nextPt = toNormalizedPoint(event);
    setDraftPoints((prev) => {
      const last = prev[prev.length - 1];
      if (last && Math.hypot(nextPt.x - last.x, nextPt.y - last.y) < 1.2) {
        return prev;
      }
      return [...prev, nextPt];
    });
  };

  const commitDraftStroke = () => {
    if (!isDrawingRef.current) {
      return;
    }
    isDrawingRef.current = false;
    if (
      isDrawActive &&
      (activeTool === 'pen' || activeTool === 'highlighter') &&
      draftPoints.length > 0
    ) {
      const normalized =
        draftPoints.length === 1
          ? [
              draftPoints[0],
              { x: draftPoints[0].x + 0.6, y: draftPoints[0].y + 0.6 },
            ]
          : draftPoints;

      const newStroke: WordDrawingStroke = {
        id: `wd-draw-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        tool: activeTool,
        color: strokeColor,
        width: strokeWidth,
        points: normalized,
      };
      onChangeStrokes([...strokes, newStroke]);
    }
    setDraftPoints([]);
  };

  const removeStrokeById = (strokeId: string) => {
    onChangeStrokes(strokes.filter((item) => item.id !== strokeId));
  };

  return (
    <svg
      ref={svgRef}
      viewBox="0 0 1000 1000"
      preserveAspectRatio="none"
      onPointerDown={handleSvgPointerDown}
      onPointerMove={handleSvgPointerMove}
      onPointerUp={commitDraftStroke}
      onPointerCancel={commitDraftStroke}
      aria-label="Lớp vẽ tự do trên văn bản"
      className={`absolute inset-0 w-full h-full z-20 ${
        isDrawActive
          ? 'pointer-events-auto touch-none cursor-crosshair'
          : 'pointer-events-none'
      }`}
    >
      {strokes.map((stroke) => {
        const isHighlighter = stroke.tool === 'highlighter';
        const computedWidth = isHighlighter
          ? Math.max(12, stroke.width * 2.4)
          : Math.max(1.5, stroke.width);

        return (
          <path
            key={stroke.id}
            d={buildSvgPathFromNormalizedPoints(stroke.points)}
            fill="none"
            stroke={stroke.color}
            strokeWidth={computedWidth}
            strokeOpacity={isHighlighter ? 0.35 : 0.92}
            strokeLinecap="round"
            strokeLinejoin="round"
            vectorEffect="non-scaling-stroke"
            onPointerDown={(event) => {
              if (!isEraseActive) {
                return;
              }
              event.stopPropagation();
              isErasingRef.current = true;
              removeStrokeById(stroke.id);
            }}
            onPointerEnter={(event) => {
              if (!isEraseActive) {
                return;
              }
              if (isErasingRef.current || event.buttons === 1) {
                removeStrokeById(stroke.id);
              }
            }}
            className={
              isEraseActive
                ? 'pointer-events-auto cursor-pointer hover:opacity-50'
                : 'pointer-events-none'
            }
          />
        );
      })}

      {draftPoints.length > 0 && isDrawActive && (
        <path
          d={buildSvgPathFromNormalizedPoints(draftPoints)}
          fill="none"
          stroke={strokeColor}
          strokeWidth={
            activeTool === 'highlighter'
              ? Math.max(12, strokeWidth * 2.4)
              : Math.max(1.5, strokeWidth)
          }
          strokeOpacity={activeTool === 'highlighter' ? 0.35 : 0.92}
          strokeLinecap="round"
          strokeLinejoin="round"
          vectorEffect="non-scaling-stroke"
          className="pointer-events-none"
        />
      )}
    </svg>
  );
}

/* ============================================================================
 * 2. WordFileInfoModal
 * ========================================================================== */

export interface WordFileInfoModalProps {
  isOpen: boolean;
  isDark: boolean;
  document: DocumentItem;
  stats: WordDocumentStats;
  isProtected: boolean;
  hasOriginalBinary: boolean;
  onClose: () => void;
}

export function WordFileInfoModal({
  isOpen,
  isDark,
  document,
  stats,
  isProtected,
  hasOriginalBinary,
  onClose,
}: WordFileInfoModalProps) {
  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const lastUpdatedLabel = document.lastSavedAt
    ? new Intl.DateTimeFormat('vi-VN', {
        dateStyle: 'medium',
        timeStyle: 'short',
      }).format(new Date(document.lastSavedAt))
    : document.updatedAt || 'Vừa cập nhật';

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="word-file-info-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-6">
        <div
          onClick={(event) => event.stopPropagation()}
          className={`w-full max-w-lg max-h-full flex flex-col overflow-hidden rounded-2xl border shadow-2xl ${
            isDark
              ? 'border-white/15 bg-[#1B2230] text-slate-100'
              : 'border-slate-200 bg-white text-slate-900'
          }`}
        >
          <header
            className={`flex shrink-0 items-center justify-between gap-3 border-b px-5 py-4 ${
              isDark
                ? 'border-white/10 bg-[#141923]'
                : 'border-slate-200 bg-[#F8FAFC]'
            }`}
          >
            <div className="flex items-center gap-2.5 min-w-0">
              <span
                className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-xl ${
                  isDark
                    ? 'bg-[#185ABD]/25 text-blue-300'
                    : 'bg-[#EEF4FF] text-[#185ABD]'
                }`}
              >
                <Info className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h2
                  id="word-file-info-title"
                  className="truncate text-base font-bold"
                >
                  Thông tin tài liệu Word
                </h2>
                <p
                  className={`truncate text-xs ${
                    isDark ? 'text-slate-400' : 'text-slate-500'
                  }`}
                >
                  Chi tiết tệp, trạng thái đồng bộ &amp; thống kê văn bản
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng thông tin tài liệu"
              className={`flex h-8 w-8 items-center justify-center rounded-xl transition-colors ${
                isDark
                  ? 'text-slate-400 hover:bg-white/10 hover:text-white'
                  : 'text-slate-400 hover:bg-slate-200/70 hover:text-slate-700'
              }`}
            >
              <X className="h-4 w-4" />
            </button>
          </header>

          <div className="min-h-0 flex-1 overflow-y-auto p-5 space-y-4">
            {/* Badges */}
            <div className="flex flex-wrap items-center gap-2">
              <span
                className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold border ${
                  hasOriginalBinary
                    ? isDark
                      ? 'bg-blue-500/20 border-blue-400/30 text-blue-200'
                      : 'bg-[#EEF4FF] border-[#185ABD]/25 text-[#185ABD]'
                    : isDark
                      ? 'bg-white/5 border-white/10 text-slate-300'
                      : 'bg-slate-100 border-slate-200 text-slate-700'
                }`}
              >
                <FileCheck2 className="h-3.5 w-3.5" />
                {hasOriginalBinary
                  ? 'Đã bảo toàn bản gốc .docx'
                  : 'Tạo mới trên LITE'}
              </span>

              <span
                className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold border ${
                  document.isSyncedToCloud
                    ? isDark
                      ? 'bg-emerald-500/20 border-emerald-400/30 text-emerald-200'
                      : 'bg-emerald-50 border-emerald-200 text-emerald-700'
                    : isDark
                      ? 'bg-white/5 border-white/10 text-slate-300'
                      : 'bg-slate-100 border-slate-200 text-slate-600'
                }`}
              >
                <Cloud className="h-3.5 w-3.5" />
                {document.isSyncedToCloud
                  ? 'Đã đồng bộ đám mây'
                  : 'Lưu trữ cục bộ'}
              </span>

              <span
                className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold border ${
                  isProtected
                    ? isDark
                      ? 'bg-amber-500/20 border-amber-400/30 text-amber-200'
                      : 'bg-amber-50 border-amber-200 text-amber-800'
                    : isDark
                      ? 'bg-white/5 border-white/10 text-slate-300'
                      : 'bg-slate-100 border-slate-200 text-slate-600'
                }`}
              >
                {isProtected ? (
                  <Lock className="h-3.5 w-3.5" />
                ) : (
                  <Unlock className="h-3.5 w-3.5" />
                )}
                {isProtected ? 'Đang khóa bảo vệ' : 'Đang mở chỉnh sửa'}
              </span>
            </div>

            {/* Metadata table */}
            <div
              className={`rounded-2xl border divide-y text-xs ${
                isDark
                  ? 'border-white/10 divide-white/10 bg-[#141923]/60'
                  : 'border-slate-200 divide-slate-100 bg-slate-50/70'
              }`}
            >
              <div className="flex items-center justify-between gap-3 px-4 py-2.5">
                <span className="inline-flex items-center gap-1.5 opacity-75">
                  <FileText className="h-3.5 w-3.5" /> Tên tài liệu
                </span>
                <span className="font-semibold text-right truncate max-w-[60%]">
                  {document.title}
                </span>
              </div>
              <div className="flex items-center justify-between gap-3 px-4 py-2.5">
                <span className="inline-flex items-center gap-1.5 opacity-75">
                  <HardDrive className="h-3.5 w-3.5" /> Tệp gốc &amp; Dung lượng
                </span>
                <span className="font-medium text-right truncate max-w-[60%]">
                  {document.originalFileName || `${document.title}.docx`} ·{' '}
                  {document.fileSize || '—'}
                </span>
              </div>
              <div className="flex items-center justify-between gap-3 px-4 py-2.5">
                <span className="inline-flex items-center gap-1.5 opacity-75">
                  <Folder className="h-3.5 w-3.5" /> Thư mục
                </span>
                <span className="font-medium text-right">
                  {document.folder || 'Tài liệu của tôi'}
                </span>
              </div>
              <div className="flex items-center justify-between gap-3 px-4 py-2.5">
                <span className="inline-flex items-center gap-1.5 opacity-75">
                  <Clock className="h-3.5 w-3.5" /> Cập nhật lần cuối
                </span>
                <span className="font-medium text-right">{lastUpdatedLabel}</span>
              </div>
            </div>

            {/* Stats grid */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
              {[
                { label: 'Số trang', value: stats.pages },
                { label: 'Số từ', value: stats.words },
                { label: 'Ký tự (có dấu cách)', value: stats.charsWithSpaces },
                { label: 'Đoạn văn', value: stats.paragraphs },
                { label: 'Số dòng ước tính', value: stats.lines },
                { label: 'Thời gian đọc', value: `~${stats.readingMinutes} phút` },
              ].map((item) => (
                <div
                  key={item.label}
                  className={`rounded-xl border p-3 ${
                    isDark
                      ? 'border-white/10 bg-white/[0.04]'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <p
                    className={`text-[11px] ${
                      isDark ? 'text-slate-400' : 'text-slate-500'
                    }`}
                  >
                    {item.label}
                  </p>
                  <p className="mt-1 text-base font-bold tabular-nums text-[#185ABD] dark:text-blue-300">
                    {typeof item.value === 'number'
                      ? item.value.toLocaleString('vi-VN')
                      : item.value}
                  </p>
                </div>
              ))}
            </div>
          </div>

          <footer
            className={`flex shrink-0 justify-end border-t px-5 py-3 ${
              isDark ? 'border-white/10 bg-[#141923]' : 'border-slate-200 bg-white'
            }`}
          >
            <button
              type="button"
              onClick={onClose}
              className="h-9 rounded-xl bg-[#185ABD] px-4 text-xs font-semibold text-white hover:bg-[#10489A]"
            >
              Đóng
            </button>
          </footer>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
 * 3. WordFootnoteModal
 * ========================================================================== */

export interface WordFootnoteModalProps {
  isOpen: boolean;
  isDark: boolean;
  nextIndex: number;
  onInsert: (footnoteText: string) => void;
  onClose: () => void;
}

export function WordFootnoteModal({
  isOpen,
  isDark,
  nextIndex,
  onInsert,
  onClose,
}: WordFootnoteModalProps) {
  const [text, setText] = useState('');

  useEffect(() => {
    if (isOpen) {
      setText('');
    }
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const handleConfirm = () => {
    const clean = text.trim();
    if (!clean) {
      return;
    }
    onInsert(clean);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="word-footnote-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-4">
        <div
          onClick={(event) => event.stopPropagation()}
          className={`w-full max-w-md rounded-2xl border p-5 shadow-2xl ${
            isDark
              ? 'border-white/15 bg-[#1B2230] text-slate-100'
              : 'border-slate-200 bg-white text-slate-900'
          }`}
        >
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2.5">
              <span
                className={`flex h-9 w-9 items-center justify-center rounded-xl font-bold text-xs ${
                  isDark
                    ? 'bg-[#185ABD]/25 text-blue-300'
                    : 'bg-[#EEF4FF] text-[#185ABD]'
                }`}
              >
                [{nextIndex}]
              </span>
              <div>
                <h2 id="word-footnote-title" className="text-sm font-bold">
                  Chèn chú thích cuối trang (Footnote)
                </h2>
                <p
                  className={`text-[11px] ${
                    isDark ? 'text-slate-400' : 'text-slate-500'
                  }`}
                >
                  Đặt số mũ [{nextIndex}] tại con trỏ và thêm mục giải thích ở cuối trang.
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng"
              className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 hover:text-slate-700"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          <div className="mt-4">
            <label
              htmlFor="word-footnote-input"
              className="mb-1.5 block text-xs font-semibold"
            >
              Nội dung chú thích [{nextIndex}]
            </label>
            <textarea
              id="word-footnote-input"
              autoFocus
              rows={3}
              value={text}
              onChange={(event) => setText(event.target.value)}
              placeholder="Ví dụ: Số liệu được tổng hợp từ báo cáo quý IV năm 2026..."
              className={`w-full rounded-xl border p-3 text-xs sm:text-sm focus:outline-none focus:ring-2 ${
                isDark
                  ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                  : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
              }`}
            />
          </div>

          <div className="mt-4 flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={onClose}
              className={`h-9 rounded-xl border px-3.5 text-xs font-semibold ${
                isDark
                  ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                  : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
              }`}
            >
              Hủy
            </button>
            <button
              type="button"
              disabled={!text.trim()}
              onClick={handleConfirm}
              className="h-9 rounded-xl bg-[#185ABD] px-4 text-xs font-semibold text-white hover:bg-[#10489A] disabled:opacity-50"
            >
              Chèn chú thích [{nextIndex}]
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
 * 4. WordCitationModal (APA / IEEE)
 * ========================================================================== */

export interface WordCitationModalProps {
  isOpen: boolean;
  isDark: boolean;
  nextBibNumber: number;
  onInsert: (payload: { inlineHtml: string; bibliographyHtml: string }) => void;
  onClose: () => void;
}

export function WordCitationModal({
  isOpen,
  isDark,
  nextBibNumber,
  onInsert,
  onClose,
}: WordCitationModalProps) {
  const [style, setStyle] = useState<'APA' | 'IEEE'>('APA');
  const [author, setAuthor] = useState('');
  const [year, setYear] = useState('2026');
  const [sourceTitle, setSourceTitle] = useState('');
  const [publisher, setPublisher] = useState('');

  useEffect(() => {
    if (isOpen) {
      setAuthor('');
      setYear('2026');
      setSourceTitle('');
      setPublisher('');
    }
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const cleanAuthor = author.trim() || 'Tác giả';
  const cleanYear = year.trim() || '2026';
  const cleanTitle = sourceTitle.trim() || 'Tên tài liệu tham khảo';
  const cleanPublisher = publisher.trim();

  const inlinePreview =
    style === 'APA' ? `(${cleanAuthor}, ${cleanYear})` : `[${nextBibNumber}]`;

  const bibPreview =
    style === 'APA'
      ? `${cleanAuthor} (${cleanYear}). ${cleanTitle}${cleanPublisher ? `. ${cleanPublisher}` : ''}.`
      : `[${nextBibNumber}] ${cleanAuthor}, "${cleanTitle}," ${cleanPublisher ? `${cleanPublisher}, ` : ''}${cleanYear}.`;

  const canSubmit = author.trim().length > 0 && sourceTitle.trim().length > 0;

  const handleInsert = () => {
    if (!canSubmit) {
      return;
    }
    const safeInline = escapeHtmlText(inlinePreview);
    const safeAuthor = escapeHtmlText(cleanAuthor);
    const safeYear = escapeHtmlText(cleanYear);
    const safeTitle = escapeHtmlText(cleanTitle);
    const safePublisher = escapeHtmlText(cleanPublisher);

    const inlineHtml = `<span data-lite-citation="true" style="color:#185ABD;font-weight:600;">${safeInline}</span>`;
    const bibliographyHtml =
      style === 'APA'
        ? `<p data-lite-bib-entry="true" style="margin:6px 0;padding-left:20px;text-indent:-20px;">${safeAuthor} (${safeYear}). <em>${safeTitle}</em>${safePublisher ? `. ${safePublisher}` : ''}.</p>`
        : `<p data-lite-bib-entry="true" style="margin:6px 0;">[${nextBibNumber}] ${safeAuthor}, &ldquo;${safeTitle},&rdquo; ${safePublisher ? `${safePublisher}, ` : ''}${safeYear}.</p>`;

    onInsert({ inlineHtml, bibliographyHtml });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="word-citation-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-6">
        <div
          onClick={(event) => event.stopPropagation()}
          className={`w-full max-w-lg max-h-full flex flex-col overflow-hidden rounded-2xl border shadow-2xl ${
            isDark
              ? 'border-white/15 bg-[#1B2230] text-slate-100'
              : 'border-slate-200 bg-white text-slate-900'
          }`}
        >
          <header
            className={`flex shrink-0 items-center justify-between gap-2 border-b px-5 py-3.5 ${
              isDark
                ? 'border-white/10 bg-[#141923]'
                : 'border-slate-200 bg-[#F8FAFC]'
            }`}
          >
            <div className="flex items-center gap-2.5">
              <span
                className={`flex h-9 w-9 items-center justify-center rounded-xl ${
                  isDark
                    ? 'bg-[#185ABD]/25 text-blue-300'
                    : 'bg-[#EEF4FF] text-[#185ABD]'
                }`}
              >
                <BookMarked className="h-4 w-4" />
              </span>
              <div>
                <h2 id="word-citation-title" className="text-sm font-bold">
                  Chèn Trích dẫn &amp; Tài liệu tham khảo
                </h2>
                <p
                  className={`text-[11px] ${
                    isDark ? 'text-slate-400' : 'text-slate-500'
                  }`}
                >
                  Hỗ trợ chuẩn học thuật APA và IEEE
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng"
              className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 hover:text-slate-700"
            >
              <X className="h-4 w-4" />
            </button>
          </header>

          <div className="min-h-0 flex-1 overflow-y-auto p-5 space-y-3.5">
            <div>
              <span className="mb-1.5 block text-xs font-semibold">
                Chuẩn trích dẫn
              </span>
              <div className="grid grid-cols-2 gap-2">
                {(['APA', 'IEEE'] as const).map((opt) => (
                  <button
                    key={opt}
                    type="button"
                    onClick={() => setStyle(opt)}
                    className={`rounded-xl border px-3 py-2 text-xs font-semibold transition-colors ${
                      style === opt
                        ? 'border-[#185ABD] bg-[#EEF4FF] text-[#185ABD]'
                        : isDark
                          ? 'border-white/10 bg-white/5 text-slate-300'
                          : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    {opt === 'APA' ? 'APA (Tác giả, Năm)' : `IEEE ([${nextBibNumber}])`}
                  </button>
                ))}
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
              <div className="sm:col-span-2">
                <label
                  htmlFor="cit-author"
                  className="mb-1 block text-xs font-semibold"
                >
                  Tác giả *
                </label>
                <input
                  id="cit-author"
                  autoFocus
                  type="text"
                  value={author}
                  onChange={(event) => setAuthor(event.target.value)}
                  placeholder="Nguyễn Văn A"
                  className={`h-9 w-full rounded-xl border px-3 text-xs focus:outline-none focus:ring-2 ${
                    isDark
                      ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                      : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
                  }`}
                />
              </div>
              <div>
                <label
                  htmlFor="cit-year"
                  className="mb-1 block text-xs font-semibold"
                >
                  Năm xuất bản
                </label>
                <input
                  id="cit-year"
                  type="text"
                  value={year}
                  onChange={(event) => setYear(event.target.value)}
                  placeholder="2026"
                  className={`h-9 w-full rounded-xl border px-3 text-xs focus:outline-none focus:ring-2 ${
                    isDark
                      ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                      : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
                  }`}
                />
              </div>
            </div>

            <div>
              <label
                htmlFor="cit-title"
                className="mb-1 block text-xs font-semibold"
              >
                Tên tài liệu / bài báo *
              </label>
              <input
                id="cit-title"
                type="text"
                value={sourceTitle}
                onChange={(event) => setSourceTitle(event.target.value)}
                placeholder="Nghiên cứu chuyển đổi số tài liệu doanh nghiệp"
                className={`h-9 w-full rounded-xl border px-3 text-xs focus:outline-none focus:ring-2 ${
                  isDark
                    ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                    : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
                }`}
              />
            </div>

            <div>
              <label
                htmlFor="cit-pub"
                className="mb-1 block text-xs font-semibold"
              >
                Nhà xuất bản / Tạp chí / URL
              </label>
              <input
                id="cit-pub"
                type="text"
                value={publisher}
                onChange={(event) => setPublisher(event.target.value)}
                placeholder="NXB Khoa học & Kỹ thuật hoặc https://..."
                className={`h-9 w-full rounded-xl border px-3 text-xs focus:outline-none focus:ring-2 ${
                  isDark
                    ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                    : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
                }`}
              />
            </div>

            <div
              className={`rounded-xl border p-3 space-y-2 text-xs ${
                isDark
                  ? 'border-white/10 bg-[#141923]/70'
                  : 'border-slate-200 bg-[#F8FAFC]'
              }`}
            >
              <div>
                <p className="text-[10.5px] font-semibold uppercase tracking-wider opacity-60">
                  Trích dẫn trong dòng
                </p>
                <p className="mt-0.5 font-semibold text-[#185ABD] dark:text-blue-300">
                  {inlinePreview}
                </p>
              </div>
              <div className="border-t border-slate-200/70 dark:border-white/10 pt-2">
                <p className="text-[10.5px] font-semibold uppercase tracking-wider opacity-60">
                  Mục Tài liệu tham khảo
                </p>
                <p className="mt-0.5 leading-relaxed">{bibPreview}</p>
              </div>
            </div>
          </div>

          <footer
            className={`flex shrink-0 items-center justify-end gap-2 border-t px-5 py-3 ${
              isDark ? 'border-white/10 bg-[#141923]' : 'border-slate-200 bg-white'
            }`}
          >
            <button
              type="button"
              onClick={onClose}
              className={`h-9 rounded-xl border px-3.5 text-xs font-semibold ${
                isDark
                  ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                  : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
              }`}
            >
              Hủy
            </button>
            <button
              type="button"
              disabled={!canSubmit}
              onClick={handleInsert}
              className="h-9 rounded-xl bg-[#185ABD] px-4 text-xs font-semibold text-white hover:bg-[#10489A] disabled:opacity-50"
            >
              Chèn trích dẫn
            </button>
          </footer>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
 * 5. WordAiToolsModal (100% client-side summary, Vietnamese grammar, expand)
 * ========================================================================== */

export interface WordAiToolsModalProps {
  isOpen: boolean;
  initialTab: 'summary' | 'grammar' | 'expand';
  isDark: boolean;
  documentTitle: string;
  plainText: string;
  selectedText: string;
  headings: string[];
  isReadOnly: boolean;
  onInsertHtmlAtTop: (html: string) => void;
  onInsertHtmlAtCursor: (html: string) => void;
  onApplyGrammarFixes: (fixes: Array<{ find: string; replace: string }>) => number;
  onClose: () => void;
}

interface GrammarIssue {
  id: string;
  label: string;
  find: string;
  replace: string;
  beforePreview: string;
  afterPreview: string;
}

type ExpandTone = 'formal' | 'academic' | 'report' | 'concise';
type ExpandKind = 'continue' | 'analyze' | 'conclude';

function buildSummarySentences(
  plainText: string,
  headings: string[]
): string[] {
  const cleaned = plainText.replace(/\s+/g, ' ').trim();
  if (!cleaned) {
    return [];
  }

  const rawSentences = cleaned
    .split(/(?<=[.!?])\s+|\n+/)
    .map((s) => s.trim())
    .filter((s) => s.length >= 24);

  if (rawSentences.length <= 5) {
    return rawSentences;
  }

  const headingKeywords = new Set(
    headings
      .join(' ')
      .toLowerCase()
      .split(/\s+/)
      .filter((w) => w.length >= 3)
  );

  const scored = rawSentences.map((sentence, idx) => {
    let score = 0;
    const words = sentence.toLowerCase().split(/\s+/);
    for (const w of words) {
      if (headingKeywords.has(w)) {
        score += 2;
      }
    }
    if (idx === 0 || idx === 1) score += 4;
    if (idx >= rawSentences.length - 2) score += 2.5;
    if (sentence.length >= 45 && sentence.length <= 180) score += 2;
    return { idx, sentence, score };
  });

  return scored
    .sort((a, b) => b.score - a.score)
    .slice(0, 5)
    .sort((a, b) => a.idx - b.idx)
    .map((item) => item.sentence);
}

function scanVietnameseGrammarIssues(plainText: string): GrammarIssue[] {
  const issues: GrammarIssue[] = [];
  const seenFind = new Set<string>();
  const text = plainText;

  const pushIssue = (
    label: string,
    find: string,
    replace: string,
    startIdx: number
  ) => {
    if (!find || find === replace || seenFind.has(`${find}->${replace}`)) {
      return;
    }
    seenFind.add(`${find}->${replace}`);
    const ctxStart = Math.max(0, startIdx - 16);
    const ctxEnd = Math.min(text.length, startIdx + find.length + 16);
    const prefix = ctxStart > 0 ? '…' : '';
    const suffix = ctxEnd < text.length ? '…' : '';
    const beforeSlice = text.slice(ctxStart, ctxEnd);
    const afterSlice =
      text.slice(ctxStart, startIdx) +
      replace +
      text.slice(startIdx + find.length, ctxEnd);

    issues.push({
      id: `gram-${issues.length + 1}`,
      label,
      find,
      replace,
      beforePreview: `${prefix}${beforeSlice}${suffix}`,
      afterPreview: `${prefix}${afterSlice}${suffix}`,
    });
  };

  // 1. Multiple spaces
  for (const match of text.matchAll(/\S {2,}\S/g)) {
    if (match.index !== undefined) {
      const m = match[0];
      pushIssue(
        'Khoảng trắng kép thừa',
        m,
        m.replace(/ {2,}/g, ' '),
        match.index
      );
    }
  }

  // 2. Space before punctuation
  for (const match of text.matchAll(/\S +([,.!?:;])/g)) {
    if (match.index !== undefined) {
      const m = match[0];
      pushIssue(
        'Khoảng trắng thừa trước dấu câu',
        m,
        m.replace(/ +([,.!?:;])/, '$1'),
        match.index
      );
    }
  }

  // 3. Missing space after punctuation (comma/semicolon/colon/question/exclamation, or period excluding URLs/decimals/TP.HCM)
  for (const match of text.matchAll(/([,!?;:])([A-Za-zÀ-ỹ])/g)) {
    if (match.index !== undefined) {
      const start = Math.max(0, match.index - 4);
      const end = Math.min(text.length, match.index + match[0].length + 4);
      const snippet = text.slice(start, end);
      if (/https?:|mailto:/i.test(snippet)) {
        continue;
      }
      pushIssue(
        'Thiếu khoảng trắng sau dấu câu',
        match[0],
        `${match[1]} ${match[2]}`,
        match.index
      );
    }
  }

  for (const match of text.matchAll(/([A-Za-zÀ-ỹ])\.([A-Za-zÀ-ỹ]{2,})/g)) {
    if (match.index !== undefined) {
      const start = Math.max(0, match.index - 8);
      const end = Math.min(text.length, match.index + match[0].length + 8);
      const snippet = text.slice(start, end);
      if (
        /(https?:\/\/|www\.|@|TP\.HCM|\.docx|\.xlsx|\.pptx|\.pdf|\.com|\.vn|\.org|\.net)/i.test(
          snippet
        )
      ) {
        continue;
      }
      pushIssue(
        'Thiếu khoảng trắng sau dấu chấm',
        match[0],
        `${match[1]}. ${match[2]}`,
        match.index
      );
    }
  }

  // 4. Lowercase after sentence end (. ! ?)
  for (const match of text.matchAll(/([.!?]\s+)([a-zàáảãạăắằẳẵặâấầẩẫậđèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵ])/g)) {
    if (match.index !== undefined) {
      pushIssue(
        'Chữ đầu câu chưa viết hoa',
        match[0],
        `${match[1]}${match[2].toUpperCase()}`,
        match.index
      );
    }
  }

  // 5. Repeated adjacent Vietnamese words
  for (const match of text.matchAll(
    /\b(và|của|các|những|trong|là|được|cho|với|tại|từ|đến|để|này|đó)\s+\1\b/gi
  )) {
    if (match.index !== undefined) {
      const firstWord = match[0].split(/\s+/)[0];
      pushIssue('Lặp từ liền kề', match[0], firstWord, match.index);
    }
  }

  // 6. Common abbreviations & Vietnamese proper nouns
  const dictRules: Array<{ regex: RegExp; replace: string; label: string }> = [
    { regex: /\bko\b/g, replace: 'không', label: 'Từ viết tắt "ko"' },
    { regex: /\bđc\b/g, replace: 'được', label: 'Từ viết tắt "đc"' },
    { regex: /\bvs\b/g, replace: 'với', label: 'Từ viết tắt "vs"' },
    {
      regex: /\bviệt nam\b/g,
      replace: 'Việt Nam',
      label: 'Tên riêng chưa viết hoa',
    },
    {
      regex: /\bhà nội\b/g,
      replace: 'Hà Nội',
      label: 'Tên riêng chưa viết hoa',
    },
    {
      regex: /\btphcm\b/gi,
      replace: 'TP.HCM',
      label: 'Chuẩn hóa viết tắt TP.HCM',
    },
  ];

  for (const rule of dictRules) {
    for (const match of text.matchAll(rule.regex)) {
      if (match.index !== undefined && match[0] !== rule.replace) {
        pushIssue(rule.label, match[0], rule.replace, match.index);
      }
    }
  }

  return issues.slice(0, 40);
}

function generateVietnameseExpansion(
  seedText: string,
  documentTitle: string,
  tone: ExpandTone,
  kind: ExpandKind,
  variantIdx: number
): string {
  const cleanSeed = seedText.trim().slice(0, 180);
  const topic = cleanSeed
    ? `"${cleanSeed.replace(/\s+/g, ' ').slice(0, 90)}${cleanSeed.length > 90 ? '…' : ''}"`
    : `"${documentTitle}"`;

  const toneLead: Record<ExpandTone, string> = {
    formal: 'Theo định hướng chuẩn hóa và nâng cao hiệu quả triển khai,',
    academic: 'Xét trên cơ sở phân tích hệ thống và đối chiếu thực tiễn,',
    report: 'Từ kết quả tổng hợp tiến độ và các số liệu ghi nhận thực tế,',
    concise: 'Trọng tâm tiếp theo tập trung vào ba điểm hành động rõ ràng:',
  };

  if (kind === 'conclude') {
    const options = [
      `${toneLead[tone]} toàn bộ nội dung liên quan đến ${topic} cần được rà soát đồng bộ trước khi ban hành chính thức. Các đầu mối phụ trách chủ động cập nhật tiến độ định kỳ nhằm đảm bảo mục tiêu đề ra đạt hiệu quả cao nhất.`,
      `Tổng kết lại, việc triển khai nhất quán các nội dung trong ${topic} sẽ giúp rút ngắn thời gian phối hợp, hạn chế sai sót phát sinh và tạo nền tảng vững chắc cho giai đoạn mở rộng tiếp theo.`,
    ];
    return options[variantIdx % options.length];
  }

  if (kind === 'analyze') {
    const options = [
      `${toneLead[tone]} khía cạnh ${topic} cho thấy mối liên hệ chặt chẽ giữa chất lượng dữ liệu đầu vào và tốc độ xử lý tổng thể. Khi chuẩn hóa quy trình rà soát theo từng bước cụ thể, các nhóm chuyên môn có thể kiểm soát rủi ro sớm và nâng cao độ tin cậy của kết quả cuối cùng.`,
      `Phân tích sâu hơn về ${topic} chỉ ra rằng việc phân tách trách nhiệm rõ ràng kết hợp cùng bộ chỉ số đo lường định lượng giúp phát hiện kịp thời các điểm nghẽn vận hành, từ đó tối ưu nguồn lực triển khai.`,
    ];
    return options[variantIdx % options.length];
  }

  const options = [
    `${toneLead[tone]} nội dung về ${topic} tiếp tục được cụ thể hóa bằng các hạng mục thực thi có mốc thời gian và người phụ trách rõ ràng. Điều này giúp đảm bảo tính liên thông giữa khâu lập kế hoạch, thực hiện và nghiệm thu kết quả.`,
    `Tiếp nối mạch trình bày của ${topic}, các bước triển khai kế tiếp ưu tiên tính đồng bộ trên nhiều thiết bị, giữ nguyên cấu trúc thông tin cốt lõi đồng thời tạo thuận lợi cho việc chia sẻ và rà soát chéo trong nhóm.`,
  ];
  return options[variantIdx % options.length];
}

export function WordAiToolsModal({
  isOpen,
  initialTab,
  isDark,
  documentTitle,
  plainText,
  selectedText,
  headings,
  isReadOnly,
  onInsertHtmlAtTop,
  onInsertHtmlAtCursor,
  onApplyGrammarFixes,
  onClose,
}: WordAiToolsModalProps) {
  const [activeTab, setActiveTab] = useState<'summary' | 'grammar' | 'expand'>(
    initialTab
  );
  const [copiedSummary, setCopiedSummary] = useState(false);
  const [copiedExpand, setCopiedExpand] = useState(false);
  const [appliedGrammarNotice, setAppliedGrammarNotice] = useState<string | null>(
    null
  );
  const [resolvedIssueIds, setResolvedIssueIds] = useState<Set<string>>(
    () => new Set()
  );

  const [expandTone, setExpandTone] = useState<ExpandTone>('formal');
  const [expandKind, setExpandKind] = useState<ExpandKind>('continue');
  const [expandVariantIdx, setExpandVariantIdx] = useState(0);
  const [expandDraft, setExpandDraft] = useState('');

  const seedForExpand = useMemo(() => {
    const sel = selectedText.trim();
    if (sel) {
      return sel;
    }
    const paragraphs = plainText
      .split(/\n+/)
      .map((p) => p.trim())
      .filter(Boolean);
    return paragraphs[paragraphs.length - 1] ?? '';
  }, [plainText, selectedText]);

  useEffect(() => {
    if (isOpen) {
      setActiveTab(initialTab);
      setCopiedSummary(false);
      setCopiedExpand(false);
      setAppliedGrammarNotice(null);
      setResolvedIssueIds(new Set());
      setExpandDraft(
        generateVietnameseExpansion(
          seedForExpand,
          documentTitle,
          expandTone,
          expandKind,
          0
        )
      );
      setExpandVariantIdx(0);
    }
  }, [documentTitle, initialTab, isOpen]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  const summarySentences = useMemo(
    () => buildSummarySentences(plainText, headings),
    [headings, plainText]
  );

  const wordCount = useMemo(() => {
    const t = plainText.trim();
    return t ? t.split(/\s+/).filter(Boolean).length : 0;
  }, [plainText]);

  const readingMinutes = Math.max(1, Math.ceil(wordCount / 200));

  const grammarIssues = useMemo(
    () => scanVietnameseGrammarIssues(plainText),
    [plainText]
  );

  const activeGrammarIssues = useMemo(
    () => grammarIssues.filter((issue) => !resolvedIssueIds.has(issue.id)),
    [grammarIssues, resolvedIssueIds]
  );

  if (!isOpen) {
    return null;
  }

  const regenerateExpand = (
    nextTone = expandTone,
    nextKind = expandKind,
    nextIdx = expandVariantIdx + 1
  ) => {
    setExpandTone(nextTone);
    setExpandKind(nextKind);
    setExpandVariantIdx(nextIdx);
    setExpandDraft(
      generateVietnameseExpansion(
        seedForExpand,
        documentTitle,
        nextTone,
        nextKind,
        nextIdx
      )
    );
  };

  const summaryTextToCopy = [
    `Tóm tắt: ${documentTitle}`,
    ...summarySentences.map((s) => `- ${s}`),
    headings.length > 0 ? `Các mục chính: ${headings.join(' | ')}` : '',
  ]
    .filter(Boolean)
    .join('\n');

  const handleCopySummary = async () => {
    try {
      await navigator.clipboard.writeText(summaryTextToCopy);
      setCopiedSummary(true);
      window.setTimeout(() => setCopiedSummary(false), 2000);
    } catch {
      // Ignore clipboard failure
    }
  };

  const handleInsertSummaryBox = () => {
    const itemsHtml = summarySentences
      .map((s) => `<li>${escapeHtmlText(s)}</li>`)
      .join('');
    const html = `<div data-lite-summary="true" style="margin:0 0 18px 0;padding:14px 16px;border-left:4px solid #185ABD;background:#EEF4FF;border-radius:8px;"><p style="margin:0 0 6px 0;font-weight:700;color:#185ABD;">Tóm tắt nội dung chính</p><ul style="margin:0;padding-left:20px;">${itemsHtml || `<li>${escapeHtmlText(documentTitle)}</li>`}</ul></div>`;
    onInsertHtmlAtTop(html);
    onClose();
  };

  const handleApplySingleFix = (issue: GrammarIssue) => {
    const count = onApplyGrammarFixes([
      { find: issue.find, replace: issue.replace },
    ]);
    setResolvedIssueIds((prev) => {
      const next = new Set(prev);
      next.add(issue.id);
      return next;
    });
    setAppliedGrammarNotice(
      count > 0
        ? `Đã sửa ${count} vị trí "${issue.find}" → "${issue.replace}".`
        : 'Đã cập nhật quy tắc sửa lỗi.'
    );
  };

  const handleApplyAllFixes = () => {
    if (activeGrammarIssues.length === 0) {
      return;
    }
    const fixes = activeGrammarIssues.map((item) => ({
      find: item.find,
      replace: item.replace,
    }));
    const count = onApplyGrammarFixes(fixes);
    setResolvedIssueIds((prev) => {
      const next = new Set(prev);
      activeGrammarIssues.forEach((item) => next.add(item.id));
      return next;
    });
    setAppliedGrammarNotice(
      `Đã tự động áp dụng ${count} sửa đổi chính tả & khoảng trắng.`
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="word-ai-modal-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-6">
        <div
          onClick={(event) => event.stopPropagation()}
          className={`flex max-h-full w-full max-w-3xl flex-col overflow-hidden rounded-2xl border shadow-2xl ${
            isDark
              ? 'border-white/15 bg-[#1B2230] text-slate-100'
              : 'border-slate-200 bg-white text-slate-900'
          }`}
        >
          {/* Header */}
          <header
            className={`flex shrink-0 items-center justify-between gap-3 border-b px-4 py-3 sm:px-5 ${
              isDark
                ? 'border-white/10 bg-[#141923]'
                : 'border-slate-200 bg-[#F8FAFC]'
            }`}
          >
            <div className="flex items-center gap-2.5 min-w-0">
              <span
                className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-xl ${
                  isDark
                    ? 'bg-[#185ABD]/25 text-blue-300'
                    : 'bg-[#EEF4FF] text-[#185ABD]'
                }`}
              >
                <Sparkles className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h2
                  id="word-ai-modal-title"
                  className="truncate text-sm sm:text-base font-bold"
                >
                  AI Tools — Trợ lý biên tập Word
                </h2>
                <p
                  className={`truncate text-xs ${
                    isDark ? 'text-slate-400' : 'text-slate-500'
                  }`}
                >
                  Xử lý tức thì 100% trên trình duyệt · Không gửi dữ liệu ra ngoài
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng AI Tools"
              className="flex h-8 w-8 items-center justify-center rounded-xl text-slate-400 hover:bg-slate-200/70 hover:text-slate-700"
            >
              <X className="h-4 w-4" />
            </button>
          </header>

          {/* Tabs */}
          <div
            className={`shrink-0 flex items-center gap-1.5 border-b px-4 pt-2 sm:px-5 overflow-x-auto no-scrollbar ${
              isDark
                ? 'border-white/10 bg-[#141923]'
                : 'border-slate-200 bg-white'
            }`}
          >
            {[
              { key: 'summary' as const, label: 'Tóm tắt văn bản', icon: Sparkles },
              {
                key: 'grammar' as const,
                label: `Sửa chính tả & ngữ pháp (${activeGrammarIssues.length})`,
                icon: SpellCheck2,
              },
              { key: 'expand' as const, label: 'Viết tiếp / Mở rộng', icon: Wand2 },
            ].map((tab) => {
              const Icon = tab.icon;
              const isActive = activeTab === tab.key;
              return (
                <button
                  key={tab.key}
                  type="button"
                  onClick={() => setActiveTab(tab.key)}
                  className={`relative inline-flex items-center gap-1.5 shrink-0 whitespace-nowrap px-3 py-2.5 text-xs font-semibold transition-colors ${
                    isActive
                      ? isDark
                        ? 'text-blue-300'
                        : 'text-[#185ABD]'
                      : isDark
                        ? 'text-slate-400 hover:text-slate-200'
                        : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  <Icon className="h-3.5 w-3.5" />
                  <span>{tab.label}</span>
                  {isActive && (
                    <span className="absolute inset-x-2 bottom-0 h-[2px] rounded-full bg-[#185ABD]" />
                  )}
                </button>
              );
            })}
          </div>

          {/* Body */}
          <div
            className={`min-h-0 flex-1 overflow-y-auto p-4 sm:p-5 ${
              isDark ? 'bg-[#12161F]' : 'bg-[#F8FAFC]'
            }`}
          >
            {activeTab === 'summary' && (
              <div className="space-y-4">
                <div className="flex flex-wrap items-center gap-2 text-xs">
                  <span
                    className={`rounded-lg border px-2.5 py-1 font-semibold tabular-nums ${
                      isDark
                        ? 'border-white/10 bg-white/5 text-slate-200'
                        : 'border-slate-200 bg-white text-slate-700'
                    }`}
                  >
                    {wordCount.toLocaleString('vi-VN')} từ
                  </span>
                  <span
                    className={`rounded-lg border px-2.5 py-1 font-semibold tabular-nums ${
                      isDark
                        ? 'border-white/10 bg-white/5 text-slate-200'
                        : 'border-slate-200 bg-white text-slate-700'
                    }`}
                  >
                    Thời gian đọc ~{readingMinutes} phút
                  </span>
                  <span
                    className={`rounded-lg border px-2.5 py-1 font-semibold tabular-nums ${
                      isDark
                        ? 'border-white/10 bg-white/5 text-slate-200'
                        : 'border-slate-200 bg-white text-slate-700'
                    }`}
                  >
                    {headings.length} tiêu đề mục
                  </span>
                </div>

                {headings.length > 0 && (
                  <div
                    className={`rounded-xl border p-3.5 ${
                      isDark
                        ? 'border-white/10 bg-[#1B2230]'
                        : 'border-slate-200 bg-white'
                    }`}
                  >
                    <p className="text-[11px] font-bold uppercase tracking-wider opacity-60 mb-2">
                      Cấu trúc mục lục phát hiện
                    </p>
                    <div className="flex flex-wrap gap-1.5">
                      {headings.map((h, i) => (
                        <span
                          key={`${h}-${i}`}
                          className={`rounded-lg px-2 py-1 text-xs font-medium ${
                            isDark
                              ? 'bg-[#185ABD]/25 text-blue-200'
                              : 'bg-[#EEF4FF] text-[#185ABD]'
                          }`}
                        >
                          {h}
                        </span>
                      ))}
                    </div>
                  </div>
                )}

                <div
                  className={`rounded-xl border p-4 space-y-2.5 ${
                    isDark
                      ? 'border-white/10 bg-[#1B2230]'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <p className="text-xs font-bold text-[#185ABD] dark:text-blue-300">
                    Các câu trọng tâm ({summarySentences.length})
                  </p>
                  {summarySentences.length === 0 ? (
                    <p className="text-xs opacity-70">
                      Văn bản còn ngắn. Hãy viết thêm nội dung để trích xuất tóm tắt tự động.
                    </p>
                  ) : (
                    <ul className="space-y-2 text-xs leading-relaxed list-disc pl-4">
                      {summarySentences.map((sentence, idx) => (
                        <li key={idx}>{sentence}</li>
                      ))}
                    </ul>
                  )}
                </div>

                <div className="flex flex-wrap items-center justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => void handleCopySummary()}
                    className={`inline-flex h-9 items-center gap-1.5 rounded-xl border px-3.5 text-xs font-semibold ${
                      isDark
                        ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                        : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                    }`}
                  >
                    {copiedSummary ? (
                      <Check className="h-3.5 w-3.5 text-emerald-500" />
                    ) : (
                      <Copy className="h-3.5 w-3.5" />
                    )}
                    <span>{copiedSummary ? 'Đã sao chép' : 'Sao chép tóm tắt'}</span>
                  </button>

                  <button
                    type="button"
                    disabled={isReadOnly || summarySentences.length === 0}
                    onClick={handleInsertSummaryBox}
                    className="inline-flex h-9 items-center gap-1.5 rounded-xl bg-[#185ABD] px-4 text-xs font-semibold text-white hover:bg-[#10489A] disabled:opacity-50"
                  >
                    <Quote className="h-3.5 w-3.5" />
                    <span>Chèn khung tóm tắt vào đầu văn bản</span>
                  </button>
                </div>
              </div>
            )}

            {activeTab === 'grammar' && (
              <div className="space-y-3.5">
                {appliedGrammarNotice && (
                  <div className="flex items-center gap-2 rounded-xl border border-emerald-200 bg-emerald-50 px-3.5 py-2.5 text-xs font-medium text-emerald-800">
                    <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-600" />
                    <span>{appliedGrammarNotice}</span>
                  </div>
                )}

                <div
                  className={`flex flex-wrap items-center justify-between gap-2 rounded-xl border p-3.5 ${
                    isDark
                      ? 'border-white/10 bg-[#1B2230]'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <p className="text-xs">
                    Phát hiện{' '}
                    <strong className="text-[#185ABD] dark:text-blue-300">
                      {activeGrammarIssues.length}
                    </strong>{' '}
                    lỗi khoảng trắng, dấu câu, lặp từ hoặc viết hoa tiếng Việt.
                  </p>

                  <button
                    type="button"
                    disabled={isReadOnly || activeGrammarIssues.length === 0}
                    onClick={handleApplyAllFixes}
                    className="inline-flex h-8 items-center gap-1.5 rounded-lg bg-[#185ABD] px-3 text-xs font-semibold text-white hover:bg-[#10489A] disabled:opacity-45"
                  >
                    <Check className="h-3.5 w-3.5" />
                    <span>Áp dụng tất cả</span>
                  </button>
                </div>

                {activeGrammarIssues.length === 0 ? (
                  <div
                    className={`rounded-xl border p-8 text-center text-xs ${
                      isDark
                        ? 'border-emerald-500/30 bg-emerald-500/10 text-emerald-200'
                        : 'border-emerald-200 bg-emerald-50 text-emerald-800'
                    }`}
                  >
                    Văn bản sạch lỗi khoảng trắng, dấu câu và chính tả phổ biến!
                  </div>
                ) : (
                  <ul className="space-y-2.5">
                    {activeGrammarIssues.map((issue) => (
                      <li
                        key={issue.id}
                        className={`rounded-xl border p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                          isDark
                            ? 'border-white/10 bg-[#1B2230]'
                            : 'border-slate-200 bg-white'
                        }`}
                      >
                        <div className="min-w-0 space-y-1">
                          <span className="inline-block rounded bg-[#EEF4FF] px-2 py-0.5 text-[10.5px] font-bold text-[#185ABD]">
                            {issue.label}
                          </span>
                          <div className="text-xs space-y-0.5">
                            <p className="text-rose-600 dark:text-rose-400 line-through break-words">
                              {issue.beforePreview}
                            </p>
                            <p className="text-emerald-700 dark:text-emerald-300 font-semibold break-words">
                              → {issue.afterPreview}
                            </p>
                          </div>
                        </div>

                        <button
                          type="button"
                          disabled={isReadOnly}
                          onClick={() => handleApplySingleFix(issue)}
                          className="shrink-0 h-8 rounded-lg border border-[#185ABD]/30 bg-[#EEF4FF] px-3 text-xs font-semibold text-[#185ABD] hover:bg-[#DFEAFE] disabled:opacity-45"
                        >
                          Áp dụng
                        </button>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}

            {activeTab === 'expand' && (
              <div className="space-y-4">
                <div
                  className={`rounded-xl border p-3.5 space-y-3 ${
                    isDark
                      ? 'border-white/10 bg-[#1B2230]'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <span className="mb-1.5 block text-[11px] font-semibold opacity-75">
                        Văn phong
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {(
                          [
                            ['formal', 'Trang trọng'],
                            ['academic', 'Học thuật'],
                            ['report', 'Báo cáo công việc'],
                            ['concise', 'Súc tích'],
                          ] as const
                        ).map(([key, label]) => (
                          <button
                            key={key}
                            type="button"
                            onClick={() =>
                              regenerateExpand(key, expandKind, expandVariantIdx)
                            }
                            className={`rounded-lg border px-2.5 py-1 text-[11px] font-semibold transition-colors ${
                              expandTone === key
                                ? 'border-[#185ABD] bg-[#EEF4FF] text-[#185ABD]'
                                : isDark
                                  ? 'border-white/10 bg-white/5 text-slate-300'
                                  : 'border-slate-200 bg-slate-50 text-slate-600'
                            }`}
                          >
                            {label}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div>
                      <span className="mb-1.5 block text-[11px] font-semibold opacity-75">
                        Kiểu mở rộng
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {(
                          [
                            ['continue', 'Viết tiếp đoạn văn'],
                            ['analyze', 'Mở rộng phân tích'],
                            ['conclude', 'Tạo đoạn kết luận'],
                          ] as const
                        ).map(([key, label]) => (
                          <button
                            key={key}
                            type="button"
                            onClick={() =>
                              regenerateExpand(expandTone, key, expandVariantIdx)
                            }
                            className={`rounded-lg border px-2.5 py-1 text-[11px] font-semibold transition-colors ${
                              expandKind === key
                                ? 'border-[#185ABD] bg-[#EEF4FF] text-[#185ABD]'
                                : isDark
                                  ? 'border-white/10 bg-white/5 text-slate-300'
                                  : 'border-slate-200 bg-slate-50 text-slate-600'
                            }`}
                          >
                            {label}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>

                <div
                  className={`rounded-xl border p-3.5 space-y-2 ${
                    isDark
                      ? 'border-white/10 bg-[#1B2230]'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold text-[#185ABD] dark:text-blue-300">
                      Đoạn văn gợi ý (có thể chỉnh sửa trước khi chèn)
                    </span>
                    <button
                      type="button"
                      onClick={() => regenerateExpand()}
                      className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-[11px] font-semibold text-[#185ABD] hover:bg-[#EEF4FF]"
                    >
                      <RefreshCw className="h-3 w-3" />
                      <span>Tạo gợi ý mới</span>
                    </button>
                  </div>

                  <textarea
                    rows={4}
                    value={expandDraft}
                    onChange={(event) => setExpandDraft(event.target.value)}
                    className={`w-full rounded-xl border p-3 text-xs sm:text-sm leading-relaxed focus:outline-none focus:ring-2 ${
                      isDark
                        ? 'border-white/15 bg-[#141923] text-white focus:border-blue-400 focus:ring-blue-400/20'
                        : 'border-slate-200 bg-slate-50 text-slate-900 focus:border-[#185ABD] focus:bg-white focus:ring-[#185ABD]/15'
                    }`}
                  />
                </div>

                <div className="flex flex-wrap items-center justify-end gap-2">
                  <button
                    type="button"
                    onClick={async () => {
                      try {
                        await navigator.clipboard.writeText(expandDraft);
                        setCopiedExpand(true);
                        window.setTimeout(() => setCopiedExpand(false), 2000);
                      } catch {
                        // Ignore clipboard failure
                      }
                    }}
                    className={`inline-flex h-9 items-center gap-1.5 rounded-xl border px-3.5 text-xs font-semibold ${
                      isDark
                        ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                        : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                    }`}
                  >
                    {copiedExpand ? (
                      <Check className="h-3.5 w-3.5 text-emerald-500" />
                    ) : (
                      <Copy className="h-3.5 w-3.5" />
                    )}
                    <span>{copiedExpand ? 'Đã sao chép' : 'Sao chép'}</span>
                  </button>

                  <button
                    type="button"
                    disabled={isReadOnly || !expandDraft.trim()}
                    onClick={() => {
                      onInsertHtmlAtCursor(`<p>${escapeHtmlText(expandDraft.trim())}</p>`);
                      onClose();
                    }}
                    className="inline-flex h-9 items-center gap-1.5 rounded-xl bg-[#185ABD] px-4 text-xs font-semibold text-white hover:bg-[#10489A] disabled:opacity-50"
                  >
                    <ShieldCheck className="h-3.5 w-3.5" />
                    <span>Chèn vào văn bản</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

import { useEffect, useMemo, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import {
  Check,
  Clock,
  Cloud,
  Copy,
  FileText,
  HardDrive,
  Image as ImageIcon,
  Layers,
  Lock,
  MessageSquare,
  Mic,
  Plus,
  Presentation,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  Trash2,
  Unlock,
  Wand2,
  X,
} from 'lucide-react';
import type {
  DocumentItem,
  SlideCommentItem,
  SlideDrawingStroke,
  SlidePreviewItem,
  SlideTextBlock,
} from '../../types/document';

// ============================================================================
// 1. SlideDrawingOverlay (Draw Tab: Pen, Highlighter, Eraser on 0..1000 SVG)
// ============================================================================

export interface SlideDrawingOverlayProps {
  strokes: SlideDrawingStroke[];
  activeTool: 'pen' | 'highlighter' | 'eraser' | null;
  strokeColor: string;
  strokeWidth: number;
  disabled?: boolean;
  onChangeStrokes: (next: SlideDrawingStroke[]) => void;
}

function buildSvgStrokePath(points: Array<{ x: number; y: number }>): string {
  if (points.length === 0) return '';
  if (points.length === 1) {
    const p = points[0];
    return `M ${p.x} ${p.y} L ${p.x + 0.2} ${p.y + 0.2}`;
  }
  const [first, ...rest] = points;
  const segments = rest.map((p) => `L ${p.x} ${p.y}`);
  return `M ${first.x} ${first.y} ${segments.join(' ')}`;
}

export function SlideDrawingOverlay({
  strokes,
  activeTool,
  strokeColor,
  strokeWidth,
  disabled = false,
  onChangeStrokes,
}: SlideDrawingOverlayProps) {
  const svgRef = useRef<SVGSVGElement | null>(null);
  const [draftStroke, setDraftStroke] = useState<SlideDrawingStroke | null>(null);
  const draftRef = useRef<SlideDrawingStroke | null>(null);
  const isErasingRef = useRef(false);

  const strokesRef = useRef(strokes);
  strokesRef.current = strokes;

  useEffect(() => {
    const handlePointerUp = () => {
      isErasingRef.current = false;
    };
    window.addEventListener('pointerup', handlePointerUp);
    window.addEventListener('pointercancel', handlePointerUp);
    return () => {
      window.removeEventListener('pointerup', handlePointerUp);
      window.removeEventListener('pointercancel', handlePointerUp);
    };
  }, []);

  // Cancel active draft if tool switches away or becomes disabled
  useEffect(() => {
    if (disabled || (activeTool !== 'pen' && activeTool !== 'highlighter')) {
      draftRef.current = null;
      setDraftStroke(null);
    }
  }, [activeTool, disabled]);

  const toNormalizedPoint = (clientX: number, clientY: number) => {
    const el = svgRef.current;
    if (!el) return { x: 0, y: 0 };
    const rect = el.getBoundingClientRect();
    const w = Math.max(1, rect.width);
    const h = Math.max(1, rect.height);
    const x = Math.max(0, Math.min(1000, ((clientX - rect.left) / w) * 1000));
    const y = Math.max(0, Math.min(1000, ((clientY - rect.top) / h) * 1000));
    return {
      x: Math.round(x * 10) / 10,
      y: Math.round(y * 10) / 10,
    };
  };

  const isDrawMode = !disabled && (activeTool === 'pen' || activeTool === 'highlighter');
  const isEraserMode = !disabled && activeTool === 'eraser';

  const handleSvgPointerDown = (e: ReactPointerEvent<SVGSVGElement>) => {
    if (!isDrawMode || e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch {
      // ignore capture issues in unsupported environments
    }
    const pt = toNormalizedPoint(e.clientX, e.clientY);
    const newStroke: SlideDrawingStroke = {
      id: `sdraw-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      tool: activeTool === 'highlighter' ? 'highlighter' : 'pen',
      color: strokeColor,
      width: strokeWidth,
      points: [pt],
    };
    draftRef.current = newStroke;
    setDraftStroke(newStroke);
  };

  const handleSvgPointerMove = (e: ReactPointerEvent<SVGSVGElement>) => {
    if (!isDrawMode || !draftRef.current) return;
    e.stopPropagation();
    const pt = toNormalizedPoint(e.clientX, e.clientY);
    const prevPts = draftRef.current.points;
    const last = prevPts[prevPts.length - 1];
    if (last && Math.hypot(pt.x - last.x, pt.y - last.y) < 1.2) return;
    const updated: SlideDrawingStroke = {
      ...draftRef.current,
      points: [...prevPts, pt],
    };
    draftRef.current = updated;
    setDraftStroke(updated);
  };

  const finishDraftStroke = (e: ReactPointerEvent<SVGSVGElement>) => {
    if (!draftRef.current) return;
    e.stopPropagation();
    try {
      if (e.currentTarget.hasPointerCapture(e.pointerId)) {
        e.currentTarget.releasePointerCapture(e.pointerId);
      }
    } catch {
      // ignore
    }
    const committed = draftRef.current;
    draftRef.current = null;
    setDraftStroke(null);
    onChangeStrokes([...strokesRef.current, committed]);
  };

  const eraseStrokeById = (id: string) => {
    if (!isEraserMode) return;
    const next = strokesRef.current.filter((s) => s.id !== id);
    if (next.length !== strokesRef.current.length) {
      onChangeStrokes(next);
    }
  };

  if (strokes.length === 0 && !isDrawMode && !draftStroke) {
    return null;
  }

  const allStrokes = draftStroke ? [...strokes, draftStroke] : strokes;

  return (
    <svg
      ref={svgRef}
      viewBox="0 0 1000 1000"
      preserveAspectRatio="none"
      onPointerDown={handleSvgPointerDown}
      onPointerMove={handleSvgPointerMove}
      onPointerUp={finishDraftStroke}
      onPointerCancel={finishDraftStroke}
      className={`absolute inset-0 w-full h-full z-30 select-none ${
        isDrawMode
          ? 'pointer-events-auto touch-none cursor-crosshair'
          : 'pointer-events-none'
      }`}
    >
      {allStrokes.map((stroke) => {
        const isHighlighter = stroke.tool === 'highlighter';
        return (
          <path
            key={stroke.id}
            d={buildSvgStrokePath(stroke.points)}
            fill="none"
            stroke={stroke.color}
            strokeWidth={isHighlighter ? Math.max(14, stroke.width * 2.5) : Math.max(1.5, stroke.width)}
            strokeOpacity={isHighlighter ? 0.36 : 0.95}
            strokeLinecap="round"
            strokeLinejoin="round"
            onPointerDown={
              isEraserMode
                ? (e) => {
                    e.stopPropagation();
                    e.preventDefault();
                    isErasingRef.current = true;
                    eraseStrokeById(stroke.id);
                  }
                : undefined
            }
            onPointerEnter={
              isEraserMode
                ? (e) => {
                    if (isErasingRef.current || e.buttons === 1) {
                      e.stopPropagation();
                      eraseStrokeById(stroke.id);
                    }
                  }
                : undefined
            }
            className={isEraserMode ? 'pointer-events-auto cursor-pointer hover:opacity-50 transition-opacity' : undefined}
          />
        );
      })}
    </svg>
  );
}

// ============================================================================
// 2. SlideCommentsPanel (Review/Collab Drawer)
// ============================================================================

export interface SlideCommentsPanelProps {
  slideIndex: number;
  slideTitle: string;
  comments: SlideCommentItem[];
  isReadOnly: boolean;
  onAddComment: (text: string) => void;
  onDeleteComment: (id: string) => void;
  onClose: () => void;
}

function formatCommentTime(iso: string): string {
  try {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return iso;
    return d.toLocaleString('vi-VN', {
      hour: '2-digit',
      minute: '2-digit',
      day: '2-digit',
      month: '2-digit',
    });
  } catch {
    return iso;
  }
}

export function SlideCommentsPanel({
  slideIndex,
  slideTitle,
  comments,
  isReadOnly,
  onAddComment,
  onDeleteComment,
  onClose,
}: SlideCommentsPanelProps) {
  const [draftText, setDraftText] = useState('');

  const handleSubmit = () => {
    const trimmed = draftText.trim();
    if (!trimmed || isReadOnly) return;
    onAddComment(trimmed);
    setDraftText('');
  };

  return (
    <aside className="w-[290px] max-w-[88vw] shrink-0 border-l border-slate-200 bg-white flex flex-col h-full shadow-sm z-20">
      {/* Header */}
      <div className="px-3.5 py-2.5 border-b border-slate-200 flex items-center justify-between gap-2 bg-[#FFF8F6]">
        <div className="min-w-0 flex items-center gap-2">
          <span className="w-7 h-7 rounded-lg bg-[#C43E1C]/10 text-[#C43E1C] flex items-center justify-center shrink-0">
            <MessageSquare className="w-4 h-4" />
          </span>
          <div className="min-w-0">
            <h3 className="text-xs font-bold text-slate-900 truncate">
              Bình luận Slide {slideIndex + 1}
            </h3>
            <p className="text-[11px] text-slate-500 truncate">
              {slideTitle || 'Trang chiếu chưa đặt tiêu đề'}
            </p>
          </div>
        </div>
        <button
          type="button"
          onClick={onClose}
          title="Đóng bảng bình luận"
          className="p-1 rounded-md text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* Add Comment Form */}
      <div className="p-3 border-b border-slate-100 bg-slate-50/70">
        {isReadOnly ? (
          <div className="flex items-center gap-2 px-2.5 py-2 rounded-lg bg-amber-50 border border-amber-200 text-[11px] text-amber-800">
            <Lock className="w-3.5 h-3.5 shrink-0" />
            <span>Đang ở chế độ chỉ đọc. Mở khóa để thêm bình luận.</span>
          </div>
        ) : (
          <div className="space-y-2">
            <textarea
              value={draftText}
              onChange={(e) => setDraftText(e.target.value)}
              onKeyDown={(e) => {
                if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                  e.preventDefault();
                  handleSubmit();
                }
              }}
              rows={3}
              placeholder={`Thêm nhận xét cho Slide ${slideIndex + 1}...`}
              className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-[#C43E1C] focus:ring-2 focus:ring-[#C43E1C]/15 resize-none"
            />
            <div className="flex items-center justify-between">
              <span className="text-[10px] text-slate-400">Ctrl + Enter để gửi</span>
              <button
                type="button"
                onClick={handleSubmit}
                disabled={!draftText.trim()}
                className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg bg-[#C43E1C] hover:bg-[#A63316] disabled:opacity-45 disabled:pointer-events-none text-white text-xs font-semibold transition-colors"
              >
                <Plus className="w-3.5 h-3.5" />
                Gửi bình luận
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Comment Thread */}
      <div className="flex-1 overflow-y-auto p-3 space-y-2.5">
        {comments.length === 0 ? (
          <div className="py-10 px-4 text-center">
            <div className="w-10 h-10 rounded-full bg-slate-100 text-slate-400 flex items-center justify-center mx-auto mb-2.5">
              <MessageSquare className="w-5 h-5" />
            </div>
            <p className="text-xs font-medium text-slate-600">Chưa có bình luận nào</p>
            <p className="text-[11px] text-slate-400 mt-1">
              Ghi chú góp ý nội dung, bố cục hoặc hiệu ứng cho trang chiếu này.
            </p>
          </div>
        ) : (
          comments.map((c) => (
            <div
              key={c.id}
              className="group rounded-xl border border-slate-200 bg-white p-2.5 shadow-sm hover:border-[#C43E1C]/35 transition-colors"
            >
              <div className="flex items-center justify-between gap-2 mb-1">
                <div className="flex items-center gap-1.5 min-w-0">
                  <span className="w-5 h-5 rounded-full bg-[#C43E1C]/15 text-[#C43E1C] text-[10px] font-bold flex items-center justify-center shrink-0">
                    {(c.author || 'B').charAt(0).toUpperCase()}
                  </span>
                  <span className="text-[11px] font-semibold text-slate-800 truncate">
                    {c.author || 'Người chỉnh sửa'}
                  </span>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <span className="text-[10px] text-slate-400">
                    {formatCommentTime(c.createdAt)}
                  </span>
                  {!isReadOnly && (
                    <button
                      type="button"
                      onClick={() => onDeleteComment(c.id)}
                      title="Xóa bình luận"
                      className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 opacity-80 group-hover:opacity-100 transition-colors"
                    >
                      <Trash2 className="w-3 h-3" />
                    </button>
                  )}
                </div>
              </div>
              <p className="text-xs text-slate-700 whitespace-pre-wrap break-words leading-relaxed pl-6">
                {c.text}
              </p>
            </div>
          ))
        )}
      </div>
    </aside>
  );
}

// ============================================================================
// 3. SlideFileInfoModal (File Tab -> Thông tin tệp)
// ============================================================================

export interface SlideFileInfoModalProps {
  isOpen: boolean;
  document: DocumentItem;
  slideCount: number;
  totalBlocks: number;
  totalImages: number;
  aspectRatioLabel: string;
  isProtected: boolean;
  hasOriginalBinary: boolean;
  onClose: () => void;
}

export function SlideFileInfoModal({
  isOpen,
  document,
  slideCount,
  totalBlocks,
  totalImages,
  aspectRatioLabel,
  isProtected,
  hasOriginalBinary,
  onClose,
}: SlideFileInfoModalProps) {
  if (!isOpen) return null;

  const syncLabel = document.isSyncedToCloud
    ? document.hasLocalEdits
      ? 'Đã đồng bộ Cloud (có thay đổi cục bộ chưa đẩy)'
      : 'Đã đồng bộ Cloud mới nhất'
    : 'Lưu trữ cục bộ trên thiết bị';

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-xs flex items-center justify-center p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-lg rounded-2xl border border-slate-200 bg-white shadow-2xl overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="px-5 py-4 border-b border-slate-200 bg-[#FFF8F6] flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-xl bg-[#C43E1C] text-white flex items-center justify-center shrink-0 shadow-sm">
              <Presentation className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h2 className="text-sm font-bold text-slate-900 truncate">
                Thông tin bài thuyết trình
              </h2>
              <p className="text-xs text-slate-500 truncate">{document.title}</p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="p-5 space-y-4">
          {/* Status Badges */}
          <div className="flex flex-wrap gap-2">
            <span
              className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-semibold border ${
                hasOriginalBinary
                  ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                  : 'bg-amber-50 text-amber-700 border-amber-200'
              }`}
            >
              <ShieldCheck className="w-3.5 h-3.5" />
              {hasOriginalBinary
                ? 'Bảo toàn dữ liệu gốc .pptx'
                : 'Bài thuyết trình dựng trong LITE'}
            </span>

            <span
              className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-semibold border ${
                isProtected
                  ? 'bg-rose-50 text-rose-700 border-rose-200'
                  : 'bg-slate-100 text-slate-700 border-slate-200'
              }`}
            >
              {isProtected ? <Lock className="w-3.5 h-3.5" /> : <Unlock className="w-3.5 h-3.5" />}
              {isProtected ? 'Đang khóa bảo vệ (Chỉ đọc)' : 'Cho phép chỉnh sửa'}
            </span>
          </div>

          {/* Metrics Grid */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
            <div className="rounded-xl border border-slate-200 bg-slate-50/80 p-3">
              <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                <Presentation className="w-3.5 h-3.5 text-[#C43E1C]" />
                <span>Số Slide</span>
              </div>
              <p className="text-lg font-bold text-slate-900 mt-0.5">{slideCount}</p>
            </div>
            <div className="rounded-xl border border-slate-200 bg-slate-50/80 p-3">
              <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                <Layers className="w-3.5 h-3.5 text-[#C43E1C]" />
                <span>Đối tượng</span>
              </div>
              <p className="text-lg font-bold text-slate-900 mt-0.5">{totalBlocks}</p>
            </div>
            <div className="rounded-xl border border-slate-200 bg-slate-50/80 p-3">
              <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                <ImageIcon className="w-3.5 h-3.5 text-[#C43E1C]" />
                <span>Hình ảnh</span>
              </div>
              <p className="text-lg font-bold text-slate-900 mt-0.5">{totalImages}</p>
            </div>
            <div className="rounded-xl border border-slate-200 bg-slate-50/80 p-3">
              <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                <FileText className="w-3.5 h-3.5 text-[#C43E1C]" />
                <span>Tỷ lệ khung</span>
              </div>
              <p className="text-lg font-bold text-slate-900 mt-0.5">{aspectRatioLabel}</p>
            </div>
          </div>

          {/* Detailed Metadata */}
          <div className="rounded-xl border border-slate-200 divide-y divide-slate-100 text-xs">
            <div className="px-3.5 py-2.5 flex items-center justify-between gap-4">
              <span className="text-slate-500 shrink-0">Tên tệp gốc</span>
              <span className="font-medium text-slate-800 truncate">
                {document.originalFileName || `${document.title}.pptx`}
              </span>
            </div>
            <div className="px-3.5 py-2.5 flex items-center justify-between gap-4">
              <span className="text-slate-500 shrink-0">Định dạng & Dung lượng</span>
              <span className="font-medium text-slate-800">
                PowerPoint (.pptx) {document.fileSize ? `· ${document.fileSize}` : ''}
              </span>
            </div>
            <div className="px-3.5 py-2.5 flex items-center justify-between gap-4">
              <span className="text-slate-500 shrink-0 flex items-center gap-1.5">
                {document.isSyncedToCloud ? (
                  <Cloud className="w-3.5 h-3.5 text-sky-600" />
                ) : (
                  <HardDrive className="w-3.5 h-3.5 text-slate-500" />
                )}
                Trạng thái đồng bộ
              </span>
              <span className="font-medium text-slate-800 text-right">{syncLabel}</span>
            </div>
            <div className="px-3.5 py-2.5 flex items-center justify-between gap-4">
              <span className="text-slate-500 shrink-0 flex items-center gap-1.5">
                <Clock className="w-3.5 h-3.5 text-slate-400" />
                Cập nhật gần nhất
              </span>
              <span className="font-medium text-slate-800">
                {document.lastSavedAt || document.updatedAt || 'Vừa xong'}
              </span>
            </div>
            <div className="px-3.5 py-2.5 flex items-center justify-between gap-4">
              <span className="text-slate-500 shrink-0">Thư mục lưu</span>
              <span className="font-medium text-slate-800">
                {document.folder || 'Chưa phân loại'}
              </span>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-5 py-3 border-t border-slate-100 bg-slate-50 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg bg-[#C43E1C] hover:bg-[#A63316] text-white text-xs font-semibold transition-colors"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 4. SlideAiToolsModal (AI Tools Tab — 100% Client-Side Vietnamese Engine)
// ============================================================================

export interface SlideAiToolsModalProps {
  isOpen: boolean;
  initialTab: 'generate' | 'summarize' | 'notes';
  documentTitle: string;
  slides: SlidePreviewItem[];
  activeSlideIndex: number;
  isReadOnly: boolean;
  onAppendSlides: (newSlides: SlidePreviewItem[]) => void;
  onUpdateSlideNote: (slideIndex: number, note: string) => void;
  onApplyAllEmptyNotes: (notesBySlideId: Record<string, string>) => void;
  onClose: () => void;
}

type AiDeckStyle = 'Chuyên nghiệp' | 'Báo cáo kinh doanh' | 'Học thuật' | 'Pitch Deck';

interface DeckPalette {
  accentColor: string;
  backgroundColor: string;
  backgroundGradient?: string;
  slideWidthRatio?: number;
  titleColor: string;
  bodyColor: string;
  isLight: boolean;
}

function isHexColorLight(hex?: string): boolean {
  if (!hex || !hex.startsWith('#')) return true;
  const clean = hex.slice(1);
  const full =
    clean.length === 3
      ? clean
          .split('')
          .map((c) => c + c)
          .join('')
      : clean.slice(0, 6);
  if (full.length !== 6) return true;
  const r = parseInt(full.slice(0, 2), 16);
  const g = parseInt(full.slice(2, 4), 16);
  const b = parseInt(full.slice(4, 6), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return true;
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance >= 0.55;
}

function inferDeckPalette(slides: SlidePreviewItem[], activeSlideIndex: number): DeckPalette {
  const ref = slides[activeSlideIndex] ?? slides[0];
  const accentColor = ref?.accentColor || '#C43E1C';
  const backgroundColor = ref?.backgroundColor || '#FFFFFF';
  const backgroundGradient = ref?.backgroundGradient;
  const slideWidthRatio = ref?.slideWidthRatio;
  const isLight = isHexColorLight(backgroundColor);

  return {
    accentColor,
    backgroundColor,
    backgroundGradient,
    slideWidthRatio,
    titleColor: isLight ? '#1E293B' : '#FFFFFF',
    bodyColor: isLight ? '#334155' : '#E2E8F0',
    isLight,
  };
}

function buildBlocksForGeneratedSlide(
  slideId: string,
  title: string,
  subtitle: string | undefined,
  bullets: string[],
  palette: DeckPalette,
  isHeroSlide: boolean
): SlideTextBlock[] {
  if (isHeroSlide) {
    return [
      {
        id: `${slideId}-accent-bar`,
        text: '',
        fontSize: 12,
        kind: 'shape',
        shapeKind: 'rect',
        fillColor: palette.accentColor,
        borderColor: palette.accentColor,
        borderWidthPx: 0,
        xPct: 35,
        yPct: 22,
        wPct: 30,
        hPct: 1.5,
      },
      {
        id: `${slideId}-title`,
        text: title,
        fontSize: 38,
        bold: true,
        color: palette.titleColor,
        align: 'center',
        valign: 'middle',
        role: 'title',
        xPct: 8,
        yPct: 27,
        wPct: 84,
        hPct: 22,
      },
      {
        id: `${slideId}-subtitle`,
        text: subtitle || bullets.join(' · '),
        fontSize: 18,
        color: palette.bodyColor,
        align: 'center',
        valign: 'top',
        role: 'subtitle',
        xPct: 14,
        yPct: 52,
        wPct: 72,
        hPct: 16,
      },
      ...(bullets.length > 0
        ? [
            {
              id: `${slideId}-points`,
              text: bullets.map((b) => `• ${b}`).join('\n'),
              fontSize: 15,
              color: palette.bodyColor,
              align: 'center' as const,
              valign: 'top' as const,
              role: 'body' as const,
              xPct: 16,
              yPct: 70,
              wPct: 68,
              hPct: 20,
            },
          ]
        : []),
    ];
  }

  return [
    {
      id: `${slideId}-accent-bar`,
      text: '',
      fontSize: 12,
      kind: 'shape',
      shapeKind: 'rect',
      fillColor: palette.accentColor,
      borderColor: palette.accentColor,
      borderWidthPx: 0,
      xPct: 6,
      yPct: 7,
      wPct: 1.2,
      hPct: 13,
    },
    {
      id: `${slideId}-title`,
      text: title,
      fontSize: 28,
      bold: true,
      color: palette.titleColor,
      align: 'left',
      valign: 'middle',
      role: 'title',
      xPct: 9,
      yPct: 6,
      wPct: 85,
      hPct: 12,
    },
    ...(subtitle
      ? [
          {
            id: `${slideId}-subtitle`,
            text: subtitle,
            fontSize: 15,
            italic: true,
            color: palette.accentColor,
            align: 'left' as const,
            valign: 'top' as const,
            role: 'subtitle' as const,
            xPct: 9,
            yPct: 18,
            wPct: 85,
            hPct: 7,
          },
        ]
      : []),
    {
      id: `${slideId}-body`,
      text: bullets.map((b) => `• ${b}`).join('\n'),
      fontSize: 18,
      color: palette.bodyColor,
      align: 'left',
      valign: 'top',
      role: 'body',
      lineHeight: 1.45,
      borderColor: `${palette.accentColor}33`,
      borderWidthPx: 1,
      borderRadius: '10px',
      fillColor: palette.isLight ? 'rgba(248, 250, 252, 0.85)' : 'rgba(255, 255, 255, 0.05)',
      paddingPct: { top: 3, right: 3, bottom: 3, left: 3 },
      xPct: 6,
      yPct: subtitle ? 28 : 23,
      wPct: 88,
      hPct: subtitle ? 60 : 65,
    },
  ];
}

function generateOutlineSlides(
  rawTopic: string,
  count: 3 | 5 | 7,
  style: AiDeckStyle,
  startIndex: number,
  palette: DeckPalette
): SlidePreviewItem[] {
  const topic = rawTopic.trim() || 'Chuyển đổi số và tối ưu quy trình';

  const blueprintByStyle: Record<
    AiDeckStyle,
    Array<{ title: string; subtitle: string; bullets: string[]; note: string }>
  > = {
    'Chuyên nghiệp': [
      {
        title: topic,
        subtitle: 'Tổng quan định hướng & Mục tiêu trọng tâm',
        bullets: [
          `Xác định phạm vi và giá trị cốt lõi của ${topic}`,
          'Đồng bộ mục tiêu giữa các phòng ban liên quan',
          'Thiết lập chuẩn mực thực thi và đo lường kết quả',
        ],
        note: `Mở đầu bài trình bày về chủ đề "${topic}", nêu rõ bối cảnh và lý do trọng tâm của cuộc họp hôm nay.`,
      },
      {
        title: 'Bối cảnh & Thách thức hiện tại',
        subtitle: 'Phân tích các điểm nghẽn cần ưu tiên giải quyết',
        bullets: [
          'Đánh giá hiện trạng quy trình, nguồn lực và công cụ đang sử dụng',
          'Nhận diện 3 rào cản lớn ảnh hưởng trực tiếp tới hiệu suất',
          'Cơ hội bứt phá khi chuẩn hóa dữ liệu và quy trình làm việc',
        ],
        note: 'Nhấn mạnh khoảng cách giữa hiện trạng và kỳ vọng để tạo sự đồng thuận về tính cấp thiết của giải pháp.',
      },
      {
        title: 'Giải pháp đề xuất & Kiến trúc triển khai',
        subtitle: 'Cách tiếp cận trọng tâm, tinh gọn và khả thi',
        bullets: [
          'Chuẩn hóa luồng công việc theo từng giai đoạn rõ ràng',
          'Ứng dụng tự động hóa cho các tác vụ lặp lại hàng ngày',
          'Thiết lập cơ chế phản hồi nhanh và kiểm soát chất lượng',
        ],
        note: 'Trình bày rõ cơ chế vận hành của giải pháp mới và lợi ích thực tế mà từng nhóm người dùng nhận được.',
      },
      {
        title: 'Phân tích tác động & Chỉ số KPI',
        subtitle: 'Thước đo hiệu quả định lượng và định tính',
        bullets: [
          'Rút ngắn 30–40% thời gian xử lý công việc trọng yếu',
          'Nâng cao độ chính xác báo cáo và tính minh bạch dữ liệu',
          'Theo dõi tiến độ theo tuần qua bảng chỉ số tập trung',
        ],
        note: 'Làm rõ các con số cam kết để ban điều hành dễ dàng theo dõi và nghiệm thu theo từng mốc.',
      },
      {
        title: 'Lộ trình triển khai & Phân công nguồn lực',
        subtitle: 'Kế hoạch hành động theo 3 giai đoạn',
        bullets: [
          'Giai đoạn 1 (Tuần 1–2): Thử nghiệm phạm vi hẹp và hiệu chỉnh',
          'Giai đoạn 2 (Tuần 3–6): Mở rộng toàn đội ngũ và đào tạo',
          'Giai đoạn 3 (Sau tuần 6): Tối ưu liên tục và nhân rộng',
        ],
        note: 'Xác nhận rõ người phụ trách từng mốc và các điểm kiểm tra (checkpoint) hàng tuần.',
      },
      {
        title: 'Quản trị rủi ro & Phương án dự phòng',
        subtitle: 'Chủ động kiểm soát biến số trong quá trình áp dụng',
        bullets: [
          'Rủi ro thay đổi thói quen: bố trí hướng dẫn 1-1 và tài liệu nhanh',
          'Rủi ro kỹ thuật: duy trì phương án vận hành song song an toàn',
          'Thiết lập kênh hỗ trợ trực tiếp trong 30 ngày đầu tiên',
        ],
        note: 'Trấn an các bên liên quan bằng kế hoạch giảm thiểu rủi ro cụ thể và thực tế.',
      },
      {
        title: 'Kết luận & Hành động tiếp theo',
        subtitle: 'Chốt quyết định và bước khởi động ngay tuần này',
        bullets: [
          `Phê duyệt phạm vi triển khai đợt đầu cho ${topic}`,
          'Thành lập nhóm nòng cốt và chốt lịch họp khởi động',
          'Hỏi đáp & tiếp thu ý kiến hoàn thiện từ các thành viên',
        ],
        note: 'Tóm tắt lại 1 thông điệp quan trọng nhất và kêu gọi quyết định phê duyệt ngay tại buổi trình bày.',
      },
    ],
    'Báo cáo kinh doanh': [
      {
        title: `Báo cáo: ${topic}`,
        subtitle: 'Tổng hợp kết quả kinh doanh & Định hướng kỳ tới',
        bullets: [
          'Điểm sáng tăng trưởng doanh thu và biên lợi nhuận',
          'Mức độ hoàn thành mục tiêu OKR/KPI trọng tâm',
          'Các quyết định điều chỉnh chiến lược trong quý tiếp theo',
        ],
        note: `Giới thiệu nhanh bức tranh tổng thể của báo cáo ${topic} trước khi đi vào từng chỉ số chi tiết.`,
      },
      {
        title: 'Tóm tắt điều hành (Executive Summary)',
        subtitle: 'Những con số nổi bật nhất cần lưu ý',
        bullets: [
          'Doanh thu đạt 112% kế hoạch đề ra nhờ tối ưu kênh chuyển đổi',
          'Chi phí thu hút khách hàng (CAC) giảm 14% so với kỳ trước',
          'Tỷ lệ giữ chân khách hàng trung thành duy trì trên 88%',
        ],
        note: 'Dành 60 giây nhấn mạnh các chỉ số vượt kỳ vọng và nguyên nhân cốt lõi tạo nên kết quả này.',
      },
      {
        title: 'Hiệu suất theo phân khúc & Kênh bán',
        subtitle: 'So sánh đóng góp giữa các mảng kinh doanh',
        bullets: [
          'Phân khúc doanh nghiệp (B2B) dẫn đầu với tốc độ tăng +28%',
          'Kênh đối tác chiến lược mở rộng thêm 15% thị phần khu vực',
          'Nhóm sản phẩm mới bắt đầu tạo dòng tiền dương ổn định',
        ],
        note: 'Chỉ rõ mảng nào đang là động lực chính và mảng nào cần tái phân bổ ngân sách.',
      },
      {
        title: 'Phân tích chi phí & Biên lợi nhuận',
        subtitle: 'Kiểm soát chi tiêu và tối ưu hiệu quả đầu tư',
        bullets: [
          'Biên lợi nhuận gộp cải thiện +3.5 điểm phần trăm',
          'Chi phí vận hành cố định được kiểm soát trong ngưỡng an toàn',
          'Tái đầu tư 20% ngân sách tiết kiệm vào tự động hóa chăm sóc khách hàng',
        ],
        note: 'Giải thích cân đối giữa tăng trưởng nhanh và kỷ luật tài chính.',
      },
      {
        title: 'Cơ hội thị trường & Động lực tăng trưởng',
        subtitle: 'Xu hướng nhu cầu và lợi thế cạnh tranh',
        bullets: [
          'Nhu cầu khách hàng doanh nghiệp vừa và nhỏ tăng mạnh',
          'Chuẩn hóa gói giải pháp trọn gói giúp rút ngắn chu kỳ chốt đơn',
          'Mở rộng liên minh tích hợp hệ sinh thái đối tác',
        ],
        note: 'Chia sẻ tín hiệu thị trường thực tế từ đội ngũ kinh doanh và phản hồi khách hàng.',
      },
      {
        title: 'Thách thức tồn đọng & Biện pháp khắc phục',
        subtitle: 'Những chỉ tiêu chưa đạt và phương án xử lý',
        bullets: [
          'Thời gian triển khai sau bán hàng còn chậm ở một số dự án lớn',
          'Bổ sung quy trình bàn giao chuẩn và nhân sự hỗ trợ kỹ thuật',
          'Theo dõi sát tỷ lệ chuyển đổi ở bước dùng thử sản phẩm',
        ],
        note: 'Thẳng thắn nhìn nhận điểm yếu và trình bày hành động khắc phục có thời hạn rõ ràng.',
      },
      {
        title: 'Kế hoạch hành động & Mục tiêu quý tới',
        subtitle: 'Cam kết con số và đầu việc ưu tiên',
        bullets: [
          'Mục tiêu tăng trưởng doanh thu +22% trong quý tiếp theo',
          'Ưu tiên 3 sáng kiến trọng điểm theo từng nhóm phụ trách',
          'Rà soát tiến độ 2 tuần/lần với ban điều hành',
        ],
        note: 'Kết thúc bằng bảng cam kết rõ ràng và mở phần trao đổi cùng ban lãnh đạo.',
      },
    ],
    'Học thuật': [
      {
        title: topic,
        subtitle: 'Nghiên cứu chuyên đề · Cơ sở lý thuyết & Thực nghiệm',
        bullets: [
          'Câu hỏi nghiên cứu và mục tiêu khoa học cốt lõi',
          'Phạm vi đối tượng khảo sát và phương pháp tiếp cận',
          'Đóng góp mới về mặt lý luận và ứng dụng thực tiễn',
        ],
        note: `Giới thiệu đề tài "${topic}", nêu rõ lý do chọn đề tài và tính cấp thiết trong lĩnh vực nghiên cứu.`,
      },
      {
        title: 'Đặt vấn đề & Tổng quan tài liệu',
        subtitle: 'Khoảng trống nghiên cứu từ các công trình trước',
        bullets: [
          'Tổng hợp các nghiên cứu tiêu biểu trong và ngoài nước gần đây',
          'Nhận diện giới hạn về dữ liệu, mô hình hoặc bối cảnh áp dụng',
          'Hình thành giả thuyết nghiên cứu (H1, H2, H3) làm nền tảng kiểm định',
        ],
        note: 'Chỉ rõ công trình này kế thừa điều gì và giải quyết khoảng trống nào mà tài liệu trước chưa chạm tới.',
      },
      {
        title: 'Phương pháp nghiên cứu & Dữ liệu',
        subtitle: 'Thiết kế thực nghiệm và công cụ phân tích',
        bullets: [
          'Kết hợp phương pháp định tính và định lượng trên tập mẫu chuẩn hóa',
          'Quy trình thu thập, làm sạch và kiểm định độ tin cậy dữ liệu',
          'Bộ tiêu chí đánh giá khách quan, đảm bảo khả năng tái lập',
        ],
        note: 'Trình bày minh bạch quy trình nghiên cứu để hội đồng đánh giá cao tính chặt chẽ khoa học.',
      },
      {
        title: 'Kết quả nghiên cứu & Phân tích số liệu',
        subtitle: 'Những phát hiện chính rút ra từ thực nghiệm',
        bullets: [
          'Kiểm chứng sự phù hợp của mô hình với mức ý nghĩa thống kê cao',
          'So sánh đối chứng giữa nhóm thử nghiệm và nhóm cơ sở',
          'Làm rõ các biến số có mức ảnh hưởng mạnh nhất tới kết quả',
        ],
        note: 'Đi thẳng vào phát hiện nổi bật nhất trước khi giải thích các bảng biểu chi tiết.',
      },
      {
        title: 'Thảo luận & Đối chiếu giả thuyết',
        subtitle: 'Ý nghĩa khoa học đằng sau các kết quả thu được',
        bullets: [
          'Giải thích cơ chế nguyên nhân – kết quả phù hợp với khung lý thuyết',
          'Phân tích các trường hợp ngoại lệ và điều kiện biên áp dụng',
          'Đối chiếu tương đồng và khác biệt với các nghiên cứu đi trước',
        ],
        note: 'Thể hiện tư duy phản biện khoa học và tính khách quan khi bàn luận giới hạn của nghiên cứu.',
      },
      {
        title: 'Hàm ý ứng dụng & Hạn chế đề tài',
        subtitle: 'Khuyến nghị thực tiễn và hướng mở rộng tương lai',
        bullets: [
          'Đề xuất giải pháp áp dụng thực tế cho đơn vị thụ hưởng',
          'Chỉ ra giới hạn về kích thước mẫu và khung thời gian khảo sát',
          'Định hướng nghiên cứu tiếp theo với dữ liệu đa chiều hơn',
        ],
        note: 'Nhấn mạnh giá trị chuyển giao thực tiễn và lộ trình nghiên cứu kế tiếp.',
      },
      {
        title: 'Kết luận & Tài liệu tham khảo',
        subtitle: 'Tổng kết đóng góp chính của chuyên đề',
        bullets: [
          `Khẳng định giá trị cốt lõi của nghiên cứu về ${topic}`,
          'Tóm lược 3 phát hiện then chốt trả lời trực tiếp câu hỏi nghiên cứu',
          'Trân trọng cảm ơn & sẵn sàng trao đổi phản biện',
        ],
        note: 'Kết thúc trang trọng, ngắn gọn và mời hội đồng/người nghe đặt câu hỏi thảo luận.',
      },
    ],
    'Pitch Deck': [
      {
        title: topic,
        subtitle: 'Giải pháp đột phá · Tầm nhìn & Cơ hội đầu tư',
        bullets: [
          'Sứ mệnh: Giải quyết bài toán lớn bằng trải nghiệm đơn giản gấp 10 lần',
          'Thời điểm vàng của thị trường và lợi thế công nghệ dẫn đầu',
          'Đội ngũ sáng lập giàu kinh nghiệm thực chiến',
        ],
        note: `Mở đầu bằng câu chuyện thật ấn tượng gắn với "${topic}" trong 30 giây đầu để thu hút nhà đầu tư.`,
      },
      {
        title: 'Vấn đề nhức nhối (The Problem)',
        subtitle: 'Nỗi đau lớn mà khách hàng đang phải trả giá mỗi ngày',
        bullets: [
          'Quy trình thủ công phân mảnh làm lãng phí hàng giờ làm việc mỗi tuần',
          'Chi phí sai sót cao và thiếu dữ liệu theo thời gian thực để ra quyết định',
          'Các công cụ cũ cồng kềnh, khó triển khai cho đội ngũ quy mô vừa và nhỏ',
        ],
        note: 'Hãy để nhà đầu tư cảm nhận rõ vấn đề này lớn, phổ biến và khách hàng sẵn sàng trả tiền để giải quyết.',
      },
      {
        title: 'Giải pháp & Sản phẩm (The Solution)',
        subtitle: 'Trải nghiệm mượt mà, triển khai trong vài phút',
        bullets: [
          `Nền tảng hợp nhất cho ${topic} trên một giao diện trực quan duy nhất`,
          'Tự động hóa quy trình lõi, giảm tới 70% thao tác thủ công',
          'Kiến trúc mở sẵn sàng tích hợp với hệ sinh thái công cụ sẵn có',
        ],
        note: 'Cho thấy sự khác biệt 10x so với cách làm cũ thay vì chỉ cải thiện nhỏ lẻ.',
      },
      {
        title: 'Quy mô thị trường (Market Size)',
        subtitle: 'TAM · SAM · SOM đầy tiềm năng tăng trưởng',
        bullets: [
          'TAM: Thị trường khu vực tăng trưởng kép 18%/năm',
          'SAM: Hàng trăm nghìn doanh nghiệp đang chuyển dịch lên nền tảng đám mây',
          'SOM: Mục tiêu chiếm lĩnh 5% phân khúc trọng tâm trong 24 tháng đầu',
        ],
        note: 'Chứng minh miếng bánh thị trường đủ lớn và chiến lược tiếp cận phân khúc đầu tiên thật sắc bén.',
      },
      {
        title: 'Mô hình kinh doanh & Độ kéo (Traction)',
        subtitle: 'Tăng trưởng doanh thu bền vững và chứng thực từ khách hàng',
        bullets: [
          'Doanh thu định kỳ (SaaS Subscription) với tỷ lệ gia hạn cao',
          'Chỉ số LTV/CAC đạt mức lành mạnh > 4.0x',
          'Tốc độ tăng trưởng người dùng hoạt động tăng đều 25% mỗi tháng',
        ],
        note: 'Trình bày bằng chứng thực tế rằng khách hàng yêu thích và sẵn sàng trả tiền định kỳ.',
      },
      {
        title: 'Lợi thế cạnh tranh (Unfair Advantage)',
        subtitle: 'Vì sao chúng tôi chiến thắng và khó bị sao chép',
        bullets: [
          'Dữ liệu và thuật toán tối ưu hóa chuyên biệt cho ngôn ngữ & quy trình bản địa',
          'Mạng lưới đối tác phân phối và cộng đồng người dùng tiên phong',
          'Tốc độ ra mắt tính năng vượt trội gấp 3 lần đối thủ truyền thống',
        ],
        note: 'Trả lời thẳng câu hỏi của nhà đầu tư: nếu đối thủ lớn nhảy vào, điều gì bảo vệ bạn?',
      },
      {
        title: 'Kế hoạch gọi vốn & Lộ trình 18 tháng',
        subtitle: 'Phân bổ vốn thông minh để bứt phá quy mô',
        bullets: [
          '45% phát triển sản phẩm lõi & nâng cấp trải nghiệm người dùng',
          '35% mở rộng tăng trưởng, bán hàng và tiếp cận thị trường',
          '20% vận hành, hạ tầng bảo mật và tuyển dụng nhân sự chủ chốt',
        ],
        note: 'Nêu rõ số vốn kêu gọi, các cột mốc đạt được sau vòng gọi vốn này và mời trao đổi chi tiết.',
      },
    ],
  };

  const pool = blueprintByStyle[style];
  let selectedTemplates: typeof pool;
  if (count === 3) {
    selectedTemplates = [pool[0], pool[2], pool[pool.length - 1]];
  } else if (count === 5) {
    selectedTemplates = [pool[0], pool[1], pool[2], pool[4], pool[pool.length - 1]];
  } else {
    selectedTemplates = pool;
  }

  const stamp = Date.now();
  return selectedTemplates.map((tpl, idx) => {
    const slideId = `ai-slide-${stamp}-${idx + 1}`;
    const isHero = idx === 0;
    return {
      id: slideId,
      index: startIndex + idx,
      title: tpl.title,
      subtitle: tpl.subtitle,
      bullets: [...tpl.bullets],
      accentColor: palette.accentColor,
      backgroundColor: palette.backgroundColor,
      backgroundGradient: palette.backgroundGradient,
      slideWidthRatio: palette.slideWidthRatio,
      speakerNote: tpl.note,
      isMasterPositioned: true,
      blocks: buildBlocksForGeneratedSlide(
        slideId,
        tpl.title,
        tpl.subtitle,
        tpl.bullets,
        palette,
        isHero
      ),
    };
  });
}

function collectSlidePlainLines(slide: SlidePreviewItem): string[] {
  const lines: string[] = [];
  if (slide.title?.trim()) lines.push(slide.title.trim());
  if (slide.subtitle?.trim()) lines.push(slide.subtitle.trim());
  if (Array.isArray(slide.bullets)) {
    for (const b of slide.bullets) {
      if (b?.trim()) lines.push(b.trim());
    }
  }
  if (Array.isArray(slide.blocks)) {
    for (const blk of slide.blocks) {
      if (blk.text?.trim()) {
        lines.push(blk.text.trim());
      } else if (blk.richParagraphs?.length) {
        const txt = blk.richParagraphs
          .map((p) => p.runs.map((r) => r.text).join(''))
          .join(' ')
          .trim();
        if (txt) lines.push(txt);
      } else if (blk.tableRows?.length) {
        const tbl = blk.tableRows
          .map((row) => row.join(' | '))
          .join(' ')
          .trim();
        if (tbl) lines.push(tbl);
      }
    }
  }
  if (slide.speakerNote?.trim()) {
    lines.push(slide.speakerNote.trim());
  }
  return Array.from(new Set(lines));
}

function buildVietnameseSpeakerNote(
  slide: SlidePreviewItem,
  slideIdx: number,
  totalSlides: number,
  nextSlide?: SlidePreviewItem
): string {
  const title = slide.title?.trim() || `Trang chiếu ${slideIdx + 1}`;
  const rawPoints: string[] = [];

  if (slide.subtitle?.trim()) {
    rawPoints.push(slide.subtitle.trim());
  }
  if (slide.bullets?.length) {
    for (const b of slide.bullets) {
      const clean = b.replace(/^[•\-*]\s*/, '').trim();
      if (clean) rawPoints.push(clean);
    }
  }
  if (rawPoints.length < 2 && slide.blocks?.length) {
    for (const blk of slide.blocks) {
      if (blk.role === 'title' || blk.kind === 'shape') continue;
      const lines = (blk.text || '')
        .split('\n')
        .map((l) => l.replace(/^[•\-*]\s*/, '').trim())
        .filter((l) => l.length > 0 && l !== title);
      for (const l of lines) {
        if (!rawPoints.includes(l)) rawPoints.push(l);
      }
    }
  }

  const keyPoints = rawPoints.slice(0, 4);
  const opening =
    slideIdx === 0
      ? `Xin chào anh/chị và các bạn. Hôm nay chúng ta cùng đi qua nội dung "${title}".`
      : slideIdx === totalSlides - 1
        ? `Đi tới phần tổng kết với chủ đề "${title}", đây là những điểm chốt quan trọng nhất.`
        : `Tiếp theo, chúng ta cùng xem xét mục "${title}" (Slide ${slideIdx + 1}/${totalSlides}).`;

  const body =
    keyPoints.length > 0
      ? keyPoints
          .map((pt, i) => `${i + 1}. Nhấn mạnh: ${pt}.`)
          .join('\n')
      : `Trình bày trực tiếp nội dung trực quan hiển thị trên màn hình và giải thích ví dụ minh họa.`;

  const transition = nextSlide
    ? `Câu chuyển: Sau khi nắm rõ phần "${title}", mời mọi người chuyển sang phần tiếp theo về "${nextSlide.title || `Slide ${slideIdx + 2}`}".`
    : `Kết thúc phần trình bày, cảm ơn mọi người đã lắng nghe và mở phiên trao đổi / hỏi đáp.`;

  const wordEstimate = opening.split(/\s+/).length + body.split(/\s+/).length + transition.split(/\s+/).length;
  const secEstimate = Math.max(25, Math.round((wordEstimate / 130) * 60));

  return [
    `[Mở dẫn] ${opening}`,
    '',
    '[Ý chính cần nhấn mạnh]',
    body,
    '',
    `[${transition}]`,
    `⏱ Thời lượng gợi ý: ~${secEstimate} giây`,
  ].join('\n');
}

export function SlideAiToolsModal({
  isOpen,
  initialTab,
  documentTitle,
  slides,
  activeSlideIndex,
  isReadOnly,
  onAppendSlides,
  onUpdateSlideNote,
  onApplyAllEmptyNotes,
  onClose,
}: SlideAiToolsModalProps) {
  const [activeTab, setActiveTab] = useState<'generate' | 'summarize' | 'notes'>(initialTab);
  const [topic, setTopic] = useState(documentTitle || 'Chiến lược phát triển sản phẩm 2026');
  const [slideCount, setSlideCount] = useState<3 | 5 | 7>(5);
  const [deckStyle, setDeckStyle] = useState<AiDeckStyle>('Chuyên nghiệp');
  const [generatedSlides, setGeneratedSlides] = useState<SlidePreviewItem[]>([]);
  const [copiedSummary, setCopiedSummary] = useState(false);
  const [draftNote, setDraftNote] = useState('');
  const [feedbackBanner, setFeedbackBanner] = useState<string | null>(null);

  const palette = useMemo(
    () => inferDeckPalette(slides, activeSlideIndex),
    [slides, activeSlideIndex]
  );

  const safeActiveIdx = Math.max(0, Math.min(slides.length - 1, activeSlideIndex));
  const currentSlide = slides[safeActiveIdx];

  // Sync initialTab and regenerate defaults when modal opens
  useEffect(() => {
    if (!isOpen) return;
    setActiveTab(initialTab);
    setFeedbackBanner(null);
    setCopiedSummary(false);

    const baseTopic = documentTitle?.trim() || 'Kế hoạch triển khai dự án';
    setTopic(baseTopic);
    setGeneratedSlides(
      generateOutlineSlides(baseTopic, slideCount, deckStyle, slides.length, palette)
    );

    if (currentSlide) {
      setDraftNote(
        buildVietnameseSpeakerNote(
          currentSlide,
          safeActiveIdx,
          slides.length,
          slides[safeActiveIdx + 1]
        )
      );
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, initialTab]);

  // Refresh speaker note suggestion when active slide changes while modal is open
  useEffect(() => {
    if (!isOpen || !currentSlide) return;
    setDraftNote(
      buildVietnameseSpeakerNote(
        currentSlide,
        safeActiveIdx,
        slides.length,
        slides[safeActiveIdx + 1]
      )
    );
  }, [isOpen, safeActiveIdx, slides.length, currentSlide?.id]);

  // Summary computation
  const summaryData = useMemo(() => {
    let totalWords = 0;
    const perSlide = slides.map((s, idx) => {
      const lines = collectSlidePlainLines(s);
      const words = lines.join(' ').trim().split(/\s+/).filter(Boolean).length;
      totalWords += words;
      const takeaway =
        s.bullets?.find((b) => b.trim()) ||
        s.subtitle?.trim() ||
        lines[1] ||
        'Trang chiếu nội dung trực quan';
      return {
        index: idx + 1,
        title: s.title?.trim() || `Slide ${idx + 1}`,
        takeaway: takeaway.replace(/^[•\-*]\s*/, ''),
        words,
      };
    });

    const estMinutes = Math.max(1, Math.round(Math.max(totalWords / 130, slides.length * 0.75)));
    const plainReport = [
      `TÓM TẮT BÀI THUYẾT TRÌNH: ${documentTitle || 'Bài thuyết trình'}`,
      `Tổng số slide: ${slides.length} · Tổng từ: ${totalWords} · Thời lượng ước tính: ~${estMinutes} phút`,
      '',
      ...perSlide.map((item) => `${item.index}. ${item.title}: ${item.takeaway}`),
    ].join('\n');

    return { totalWords, estMinutes, perSlide, plainReport };
  }, [slides, documentTitle]);

  const emptyNoteSlidesCount = useMemo(
    () => slides.filter((s) => !s.speakerNote || !s.speakerNote.trim()).length,
    [slides]
  );

  if (!isOpen) return null;

  const handleRegenerateOutline = () => {
    const next = generateOutlineSlides(topic, slideCount, deckStyle, slides.length, palette);
    setGeneratedSlides(next);
    setFeedbackBanner(null);
  };

  const handleUpdateGeneratedSlide = (
    idx: number,
    patch: { title?: string; subtitle?: string; bulletsText?: string; speakerNote?: string }
  ) => {
    setGeneratedSlides((prev) =>
      prev.map((item, i) => {
        if (i !== idx) return item;
        const nextTitle = patch.title !== undefined ? patch.title : item.title;
        const nextSubtitle = patch.subtitle !== undefined ? patch.subtitle : item.subtitle;
        const nextBullets =
          patch.bulletsText !== undefined
            ? patch.bulletsText
                .split('\n')
                .map((l) => l.replace(/^[•\-*]\s*/, '').trim())
                .filter(Boolean)
            : item.bullets;
        const nextNote = patch.speakerNote !== undefined ? patch.speakerNote : item.speakerNote;
        return {
          ...item,
          title: nextTitle,
          subtitle: nextSubtitle,
          bullets: nextBullets,
          speakerNote: nextNote,
          blocks: buildBlocksForGeneratedSlide(
            item.id,
            nextTitle,
            nextSubtitle,
            nextBullets,
            palette,
            i === 0
          ),
        };
      })
    );
  };

  const handleInsertGeneratedSlides = () => {
    if (isReadOnly || generatedSlides.length === 0) return;
    onAppendSlides(generatedSlides);
    setFeedbackBanner(`Đã chèn ${generatedSlides.length} slide mới vào cuối bài thuyết trình.`);
  };

  const handleCopySummary = async () => {
    try {
      await navigator.clipboard.writeText(summaryData.plainReport);
      setCopiedSummary(true);
      setTimeout(() => setCopiedSummary(false), 2000);
    } catch {
      setFeedbackBanner('Không thể truy cập clipboard trình duyệt.');
    }
  };

  const handleInsertSummarySlide = () => {
    if (isReadOnly) return;
    const stamp = Date.now();
    const slideId = `ai-summary-${stamp}`;
    const summaryBullets = summaryData.perSlide
      .slice(0, 6)
      .map((s) => `${s.title}: ${s.takeaway}`);
    if (summaryData.perSlide.length > 6) {
      summaryBullets.push(`...và ${summaryData.perSlide.length - 6} nội dung tiếp theo`);
    }
    const subtitle = `${slides.length} slide · ${summaryData.totalWords} từ · Thời lượng nói ~${summaryData.estMinutes} phút`;
    const newSlide: SlidePreviewItem = {
      id: slideId,
      index: slides.length,
      title: `Tổng kết: ${documentTitle || 'Bài thuyết trình'}`,
      subtitle,
      bullets: summaryBullets,
      accentColor: palette.accentColor,
      backgroundColor: palette.backgroundColor,
      backgroundGradient: palette.backgroundGradient,
      slideWidthRatio: palette.slideWidthRatio,
      speakerNote: summaryData.plainReport,
      isMasterPositioned: true,
      blocks: buildBlocksForGeneratedSlide(
        slideId,
        `Tổng kết: ${documentTitle || 'Bài thuyết trình'}`,
        subtitle,
        summaryBullets,
        palette,
        false
      ),
    };
    onAppendSlides([newSlide]);
    setFeedbackBanner('Đã chèn Slide Tóm tắt vào cuối bài thuyết trình.');
  };

  const handleApplyCurrentNote = () => {
    if (isReadOnly || !currentSlide) return;
    onUpdateSlideNote(safeActiveIdx, draftNote);
    setFeedbackBanner(`Đã lưu lời thoại diễn giả cho Slide ${safeActiveIdx + 1}.`);
  };

  const handleFillAllEmptyNotes = () => {
    if (isReadOnly) return;
    const notesBySlideId: Record<string, string> = {};
    slides.forEach((s, idx) => {
      if (!s.speakerNote || !s.speakerNote.trim()) {
        notesBySlideId[s.id] = buildVietnameseSpeakerNote(
          s,
          idx,
          slides.length,
          slides[idx + 1]
        );
      }
    });
    const count = Object.keys(notesBySlideId).length;
    if (count === 0) {
      setFeedbackBanner('Tất cả các slide đều đã có ghi chú diễn giả.');
      return;
    }
    onApplyAllEmptyNotes(notesBySlideId);
    setFeedbackBanner(`Đã tự động điền lời thoại cho ${count} slide đang trống.`);
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-3xl max-h-[88vh] rounded-2xl border border-slate-200 bg-white shadow-2xl flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header + Tab Switcher */}
        <div className="px-5 pt-4 pb-3 border-b border-slate-200 bg-[#FFF8F6] shrink-0">
          <div className="flex items-center justify-between gap-3 mb-3">
            <div className="flex items-center gap-2.5 min-w-0">
              <div className="w-9 h-9 rounded-xl bg-[#C43E1C] text-white flex items-center justify-center shrink-0 shadow-sm">
                <Sparkles className="w-4 h-4" />
              </div>
              <div className="min-w-0">
                <h2 className="text-sm font-bold text-slate-900">
                  Trợ lý AI PowerPoint (Client-Side)
                </h2>
                <p className="text-[11px] text-slate-500 truncate">
                  Tạo dàn ý slide, tóm tắt bài thuyết trình và soạn lời thoại diễn giả tức thì
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 transition-colors"
            >
              <X className="w-4 h-4" />
            </button>
          </div>

          <div className="flex flex-wrap gap-1.5">
            {(
              [
                { id: 'generate', label: 'Tạo dàn ý & Slide mới', Icon: Wand2 },
                { id: 'summarize', label: 'Tóm tắt toàn bài', Icon: FileText },
                { id: 'notes', label: 'Gợi ý lời thoại diễn giả', Icon: Mic },
              ] as const
            ).map(({ id, label, Icon }) => {
              const active = activeTab === id;
              return (
                <button
                  key={id}
                  type="button"
                  onClick={() => {
                    setActiveTab(id);
                    setFeedbackBanner(null);
                  }}
                  className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
                    active
                      ? 'bg-[#C43E1C] text-white shadow-xs'
                      : 'bg-white text-slate-600 border border-slate-200 hover:border-[#C43E1C]/40 hover:text-[#C43E1C]'
                  }`}
                >
                  <Icon className="w-3.5 h-3.5" />
                  {label}
                </button>
              );
            })}
          </div>
        </div>

        {/* Feedback / Read-only Banner */}
        {(isReadOnly || feedbackBanner) && (
          <div className="px-5 pt-3 shrink-0 space-y-2">
            {isReadOnly && (
              <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800">
                <Lock className="w-3.5 h-3.5 shrink-0" />
                <span>Bài thuyết trình đang khóa bảo vệ (Chỉ đọc). Bạn vẫn có thể xem tóm tắt nhưng không thể chèn/sửa slide.</span>
              </div>
            )}
            {feedbackBanner && (
              <div className="flex items-center justify-between gap-2 px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-800">
                <span className="flex items-center gap-1.5 font-medium">
                  <Check className="w-4 h-4 shrink-0" />
                  {feedbackBanner}
                </span>
                <button
                  type="button"
                  onClick={() => setFeedbackBanner(null)}
                  className="text-emerald-600 hover:text-emerald-900"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
            )}
          </div>
        )}

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* TAB 1: GENERATE SLIDES */}
          {activeTab === 'generate' && (
            <div className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-12 gap-3 p-3.5 rounded-xl border border-slate-200 bg-slate-50/70">
                <div className="md:col-span-6">
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Chủ đề bài thuyết trình
                  </label>
                  <input
                    type="text"
                    value={topic}
                    onChange={(e) => setTopic(e.target.value)}
                    placeholder="Ví dụ: Kế hoạch ra mắt sản phẩm Q3..."
                    className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-[#C43E1C] focus:ring-2 focus:ring-[#C43E1C]/15"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Số slide
                  </label>
                  <select
                    value={slideCount}
                    onChange={(e) => setSlideCount(Number(e.target.value) as 3 | 5 | 7)}
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-[#C43E1C]"
                  >
                    <option value={3}>3 slide</option>
                    <option value={5}>5 slide</option>
                    <option value={7}>7 slide</option>
                  </select>
                </div>
                <div className="md:col-span-4 flex items-end gap-2">
                  <div className="flex-1">
                    <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                      Phong cách
                    </label>
                    <select
                      value={deckStyle}
                      onChange={(e) => setDeckStyle(e.target.value as AiDeckStyle)}
                      className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-[#C43E1C]"
                    >
                      <option value="Chuyên nghiệp">Chuyên nghiệp</option>
                      <option value="Báo cáo kinh doanh">Báo cáo kinh doanh</option>
                      <option value="Học thuật">Học thuật</option>
                      <option value="Pitch Deck">Pitch Deck</option>
                    </select>
                  </div>
                  <button
                    type="button"
                    onClick={handleRegenerateOutline}
                    className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg border border-[#C43E1C] text-[#C43E1C] hover:bg-[#C43E1C]/5 text-xs font-semibold shrink-0 transition-colors"
                  >
                    <RefreshCw className="w-3.5 h-3.5" />
                    Tạo lại
                  </button>
                </div>
              </div>

              {/* Generated Slide Preview & Inline Editor */}
              <div className="space-y-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-700">
                    Xem trước & chỉnh sửa dàn ý ({generatedSlides.length} slide · màu nhấn{' '}
                    <span
                      className="inline-block w-2.5 h-2.5 rounded-full align-middle border border-slate-300"
                      style={{ backgroundColor: palette.accentColor }}
                    />
                    )
                  </span>
                </div>

                {generatedSlides.map((s, idx) => (
                  <div
                    key={s.id}
                    className="rounded-xl border border-slate-200 bg-white p-3 space-y-2 shadow-2xs"
                  >
                    <div className="flex items-center gap-2">
                      <span
                        className="px-2 py-0.5 rounded-md text-[10px] font-bold text-white shrink-0"
                        style={{ backgroundColor: palette.accentColor }}
                      >
                        Slide {idx + 1}
                      </span>
                      <input
                        type="text"
                        value={s.title}
                        onChange={(e) =>
                          handleUpdateGeneratedSlide(idx, { title: e.target.value })
                        }
                        className="flex-1 rounded-md border border-slate-200 px-2.5 py-1 text-xs font-semibold text-slate-900 focus:outline-none focus:border-[#C43E1C]"
                      />
                    </div>
                    <input
                      type="text"
                      value={s.subtitle || ''}
                      onChange={(e) =>
                        handleUpdateGeneratedSlide(idx, { subtitle: e.target.value })
                      }
                      placeholder="Phụ đề slide..."
                      className="w-full rounded-md border border-slate-200 px-2.5 py-1 text-[11px] text-slate-600 focus:outline-none focus:border-[#C43E1C]"
                    />
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      <div>
                        <label className="block text-[10px] font-medium text-slate-500 mb-0.5">
                          Các gạch đầu dòng (mỗi dòng 1 ý)
                        </label>
                        <textarea
                          rows={3}
                          value={s.bullets.join('\n')}
                          onChange={(e) =>
                            handleUpdateGeneratedSlide(idx, { bulletsText: e.target.value })
                          }
                          className="w-full rounded-md border border-slate-200 px-2.5 py-1.5 text-[11px] text-slate-700 focus:outline-none focus:border-[#C43E1C] resize-none"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-medium text-slate-500 mb-0.5">
                          Ghi chú diễn giả
                        </label>
                        <textarea
                          rows={3}
                          value={s.speakerNote || ''}
                          onChange={(e) =>
                            handleUpdateGeneratedSlide(idx, { speakerNote: e.target.value })
                          }
                          className="w-full rounded-md border border-slate-200 px-2.5 py-1.5 text-[11px] text-slate-600 focus:outline-none focus:border-[#C43E1C] resize-none"
                        />
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 2: SUMMARIZE DECK */}
          {activeTab === 'summarize' && (
            <div className="space-y-4">
              <div className="grid grid-cols-3 gap-3">
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
                  <p className="text-[11px] text-slate-500">Tổng số trang chiếu</p>
                  <p className="text-lg font-bold text-slate-900 mt-0.5">{slides.length} slide</p>
                </div>
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
                  <p className="text-[11px] text-slate-500">Tổng lượng nội dung</p>
                  <p className="text-lg font-bold text-slate-900 mt-0.5">
                    {summaryData.totalWords} từ
                  </p>
                </div>
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-3">
                  <p className="text-[11px] text-slate-500">Thời lượng thuyết trình</p>
                  <p className="text-lg font-bold text-[#C43E1C] mt-0.5">
                    ~{summaryData.estMinutes} phút
                  </p>
                </div>
              </div>

              <div className="rounded-xl border border-slate-200 divide-y divide-slate-100">
                {summaryData.perSlide.map((item) => (
                  <div key={item.index} className="p-3 flex items-start gap-3">
                    <span className="px-2 py-0.5 rounded-md bg-[#C43E1C]/10 text-[#C43E1C] text-[11px] font-bold shrink-0">
                      #{item.index}
                    </span>
                    <div className="min-w-0 flex-1">
                      <p className="text-xs font-bold text-slate-900">{item.title}</p>
                      <p className="text-xs text-slate-600 mt-0.5 leading-relaxed">
                        {item.takeaway}
                      </p>
                    </div>
                    <span className="text-[10px] text-slate-400 shrink-0">{item.words} từ</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 3: SPEAKER NOTES */}
          {activeTab === 'notes' && (
            <div className="space-y-4">
              {currentSlide ? (
                <>
                  <div className="flex items-center justify-between gap-3 p-3 rounded-xl border border-slate-200 bg-slate-50/80">
                    <div className="min-w-0">
                      <p className="text-xs font-bold text-slate-900 truncate">
                        Đang phân tích Slide {safeActiveIdx + 1}/{slides.length}:{' '}
                        {currentSlide.title || 'Chưa có tiêu đề'}
                      </p>
                      <p className="text-[11px] text-slate-500 mt-0.5">
                        Kịch bản gồm Mở dẫn, Ý chính cần nhấn mạnh và Câu chuyển slide kế tiếp.
                      </p>
                    </div>
                    <button
                      type="button"
                      onClick={() =>
                        setDraftNote(
                          buildVietnameseSpeakerNote(
                            currentSlide,
                            safeActiveIdx,
                            slides.length,
                            slides[safeActiveIdx + 1]
                          )
                        )
                      }
                      className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:border-[#C43E1C] text-xs font-medium text-slate-700 shrink-0 transition-colors"
                    >
                      <RefreshCw className="w-3.5 h-3.5" />
                      Soạn lại
                    </button>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      Lời thoại diễn giả đề xuất (có thể chỉnh sửa trước khi áp dụng)
                    </label>
                    <textarea
                      rows={9}
                      value={draftNote}
                      onChange={(e) => setDraftNote(e.target.value)}
                      className="w-full rounded-xl border border-slate-200 bg-white p-3 text-xs text-slate-800 leading-relaxed focus:outline-none focus:border-[#C43E1C] focus:ring-2 focus:ring-[#C43E1C]/15"
                    />
                  </div>
                </>
              ) : (
                <p className="text-xs text-slate-500">Chưa có slide nào để tạo ghi chú.</p>
              )}
            </div>
          )}
        </div>

        {/* Footer Actions */}
        <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-2 shrink-0">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-xs font-semibold text-slate-700 transition-colors"
          >
            Đóng
          </button>

          <div className="flex flex-wrap items-center gap-2">
            {activeTab === 'generate' && (
              <button
                type="button"
                onClick={handleInsertGeneratedSlides}
                disabled={isReadOnly || generatedSlides.length === 0}
                className="inline-flex items-center gap-1.5 px-4 py-1.5 rounded-lg bg-[#C43E1C] hover:bg-[#A63316] disabled:opacity-45 disabled:pointer-events-none text-white text-xs font-semibold transition-colors"
              >
                <Plus className="w-3.5 h-3.5" />
                Chèn vào bài thuyết trình
              </button>
            )}

            {activeTab === 'summarize' && (
              <>
                <button
                  type="button"
                  onClick={handleCopySummary}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:border-[#C43E1C] text-xs font-semibold text-slate-700 transition-colors"
                >
                  {copiedSummary ? (
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                  ) : (
                    <Copy className="w-3.5 h-3.5" />
                  )}
                  {copiedSummary ? 'Đã sao chép' : 'Sao chép tóm tắt'}
                </button>
                <button
                  type="button"
                  onClick={handleInsertSummarySlide}
                  disabled={isReadOnly}
                  className="inline-flex items-center gap-1.5 px-4 py-1.5 rounded-lg bg-[#C43E1C] hover:bg-[#A63316] disabled:opacity-45 disabled:pointer-events-none text-white text-xs font-semibold transition-colors"
                >
                  <Plus className="w-3.5 h-3.5" />
                  Chèn Slide Tóm tắt vào cuối bài
                </button>
              </>
            )}

            {activeTab === 'notes' && (
              <>
                <button
                  type="button"
                  onClick={handleFillAllEmptyNotes}
                  disabled={isReadOnly || emptyNoteSlidesCount === 0}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg border border-[#C43E1C]/40 bg-white hover:bg-[#C43E1C]/5 disabled:opacity-45 disabled:pointer-events-none text-xs font-semibold text-[#C43E1C] transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  Tự động điền ghi chú cho tất cả Slide đang trống ({emptyNoteSlidesCount})
                </button>
                <button
                  type="button"
                  onClick={handleApplyCurrentNote}
                  disabled={isReadOnly || !currentSlide}
                  className="inline-flex items-center gap-1.5 px-4 py-1.5 rounded-lg bg-[#C43E1C] hover:bg-[#A63316] disabled:opacity-45 disabled:pointer-events-none text-white text-xs font-semibold transition-colors"
                >
                  <Check className="w-3.5 h-3.5" />
                  Áp dụng cho Slide hiện tại
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

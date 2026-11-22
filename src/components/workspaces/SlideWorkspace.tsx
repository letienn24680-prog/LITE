import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type {
  ChangeEvent,
  CSSProperties,
  PointerEvent as ReactPointerEvent,
  MouseEvent as ReactMouseEvent,
} from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  CloudDownload,
  Columns2,
  Copy,
  Crosshair,
  Download,
  Image,
  ImagePlus,
  Italic,
  Layers,
  LayoutGrid,
  Loader2,
  Move,
  PaintBucket,
  Pause,
  Play,
  Plus,
  QrCode,
  RefreshCw,
  RotateCcw,
  Save,
  Sparkles,
  StickyNote,
  Trash2,
  Type,
  Underline,
  X,
} from 'lucide-react';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  SlideImageItem,
  SlidePreviewItem,
  SlideTextBlock,
} from '../../types/document';
import { exportSlidesToPptx } from '../../utils/fileExporter';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import {
  FloatingSelectionToolbar,
  type FloatingBlockTag,
} from './FloatingSelectionToolbar';
import { WorkspaceTopBar } from './WorkspaceTopBar';

export interface SlideWorkspaceProps {
  document: DocumentItem;
  onBackHome: () => void;
  onOpenModal: (
    modal: Exclude<ModalType, null>,
    overrideDoc?: DocumentItem
  ) => void;
  onSaveDocument: (updatedDoc: DocumentItem) => void;
  isModalOpen?: boolean;
  autoSaveEnabled?: boolean;
  isHydratingCloudFile?: boolean;
  onRetryCloudDownload?: () => void;
  collaborators?: CollaboratorPresence[];
  selfSessionId?: string;
  collabPermission?: 'owner' | 'edit' | 'view';
  onPresenceUpdate?: (
    patch: Partial<Omit<CollaboratorPresence, 'sessionId' | 'color'>>
  ) => void;
  onBroadcastPatch?: (
    patch: Omit<DocRealtimePatch, 'senderSessionId' | 'sentAt'>
  ) => void;
  incomingRemotePatch?: DocRealtimePatch | null;
}

interface LaserPosition {
  x: number;
  y: number;
  visible: boolean;
}

type SlideshowTransition = 'slide' | 'fade';

interface ActiveDragState {
  kind: 'block' | 'image';
  id: string;
  startClientX: number;
  startClientY: number;
  originXPct: number;
  originYPct: number;
  widthPct: number;
  heightPct: number;
  hasMoved: boolean;
}

const REFERENCE_SLIDE_WIDTH_PT = 960; // 13.333in * 72pt
const SLIDE_FONT_SIZES = [12, 14, 18, 24, 32, 40] as const;

const SLIDE_COLOR_SWATCHES = [
  { label: 'Trắng', value: '#FFFFFF' },
  { label: 'Bạc sáng', value: '#CBD5E1' },
  { label: 'Vàng nắng', value: '#FDE047' },
  { label: 'Xanh trời', value: '#38BDF8' },
  { label: 'Xanh ngọc', value: '#4ADE80' },
  { label: 'Hồng san hô', value: '#FB7185' },
  { label: 'Đen than', value: '#1E293B' },
] as const;

const SLIDE_BG_PRESETS = [
  { label: 'Đêm sâu', value: '#161D2E' },
  { label: 'Xanh hải quân', value: '#0F172A' },
  { label: 'Tím chàm', value: '#1E1B4B' },
  { label: 'Trắng tinh', value: '#FFFFFF' },
  { label: 'Giấy ấm', value: '#FAF8F5' },
] as const;

function formatTimer(totalSeconds: number): string {
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  return [hours, minutes, seconds]
    .map((unit) => String(unit).padStart(2, '0'))
    .join(':');
}

function isEditableElement(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  const tagName = target.tagName;
  return (
    tagName === 'INPUT' ||
    tagName === 'TEXTAREA' ||
    tagName === 'SELECT' ||
    target.isContentEditable
  );
}

function clampValue(val: number, min: number, max: number): number {
  if (val < min) {
    return min;
  }
  if (val > max) {
    return max;
  }
  return val;
}

function roundPct(val: number): number {
  return Math.round(val * 100) / 100;
}

function toSafeColorInputValue(
  color: string | undefined,
  fallback = '#FFFFFF'
): string {
  if (!color) {
    return fallback;
  }
  const trimmed = color.trim();
  if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) {
    return trimmed;
  }
  if (/^#[0-9a-fA-F]{3}$/.test(trimmed)) {
    const [, r, g, b] = trimmed;
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  return fallback;
}

function isLightHexBackground(hex: string | undefined): boolean {
  if (!hex) {
    return false;
  }
  const cleaned = hex.trim().replace('#', '');
  let r = 0;
  let g = 0;
  let b = 0;
  if (/^[0-9a-fA-F]{3}$/.test(cleaned)) {
    r = parseInt(cleaned[0] + cleaned[0], 16);
    g = parseInt(cleaned[1] + cleaned[1], 16);
    b = parseInt(cleaned[2] + cleaned[2], 16);
  } else if (/^[0-9a-fA-F]{6}$/.test(cleaned)) {
    r = parseInt(cleaned.slice(0, 2), 16);
    g = parseInt(cleaned.slice(2, 4), 16);
    b = parseInt(cleaned.slice(4, 6), 16);
  } else {
    return false;
  }
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.62;
}

function toRomanNumeral(num: number): string {
  const pairs: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.max(1, Math.floor(num));
  let result = '';
  for (const [value, symbol] of pairs) {
    while (remaining >= value) {
      result += symbol;
      remaining -= value;
    }
  }
  return result;
}

function isSlideMasterPositioned(slide: SlidePreviewItem): boolean {
  if (slide.isMasterPositioned) {
    return true;
  }
  const hasBlockCoords = (slide.blocks ?? []).some(
    (block) =>
      typeof block.xPct === 'number' && typeof block.yPct === 'number'
  );
  if (hasBlockCoords) {
    return true;
  }
  return (slide.images ?? []).some(
    (img) => typeof img.xPct === 'number' && typeof img.yPct === 'number'
  );
}

function buildSlideBackgroundStyle(slide: SlidePreviewItem): CSSProperties {
  const accent = slide.accentColor || '#4F46E5';
  if (slide.backgroundImageSrc) {
    return {
      backgroundColor: slide.backgroundColor ?? '#161D2E',
      backgroundImage: `url("${slide.backgroundImageSrc}")`,
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      backgroundRepeat: 'no-repeat',
    };
  }
  if (slide.backgroundGradient) {
    return {
      backgroundColor: slide.backgroundColor ?? '#161D2E',
      backgroundImage: slide.backgroundGradient,
    };
  }
  if (slide.backgroundColor) {
    return {
      backgroundColor: slide.backgroundColor,
    };
  }
  return {
    backgroundImage: `radial-gradient(circle at 86% 16%, ${accent}3A 0%, transparent 46%), radial-gradient(circle at 10% 90%, ${accent}1C 0%, transparent 40%), linear-gradient(145deg, #161D2E 0%, #1F293F 56%, #131926 100%)`,
  };
}

function getSafeAspectRatio(
  slide: SlidePreviewItem | undefined,
  document: DocumentItem
): number {
  const candidate = slide?.slideWidthRatio ?? document.slideAspectRatio;
  if (
    typeof candidate === 'number' &&
    Number.isFinite(candidate) &&
    candidate > 0.25 &&
    candidate < 4
  ) {
    return candidate;
  }
  return 16 / 9;
}

function normalizeSlideWithBlocks(
  slide: SlidePreviewItem,
  idx: number
): SlidePreviewItem {
  if (
    Array.isArray(slide.blocks) &&
    (slide.blocks.length > 0 || (slide.images && slide.images.length > 0))
  ) {
    return {
      ...slide,
      index: idx + 1,
      blocks: slide.blocks.map((block) => ({
        ...block,
        tableRows: block.tableRows
          ? block.tableRows.map((row) => [...row])
          : undefined,
      })),
      images: slide.images ? slide.images.map((img) => ({ ...img })) : undefined,
    };
  }

  const blocks: SlideTextBlock[] = [];
  blocks.push({
    id: `${slide.id}-blk-title`,
    text: slide.title || `Slide ${idx + 1}`,
    fontSize: 32,
    bold: true,
    italic: false,
    color: '#FFFFFF',
    align: 'left',
    role: 'title',
  });

  if (slide.subtitle && slide.subtitle.trim().length > 0) {
    blocks.push({
      id: `${slide.id}-blk-subtitle`,
      text: slide.subtitle,
      fontSize: 18,
      bold: false,
      italic: false,
      color: '#CBD5E1',
      align: 'left',
      role: 'subtitle',
    });
  }

  if (slide.bullets.length > 0) {
    blocks.push({
      id: `${slide.id}-blk-body`,
      text: slide.bullets.map((item) => `• ${item}`).join('\n'),
      fontSize: 18,
      bold: false,
      italic: false,
      color: '#E2E8F0',
      align: 'left',
      role: 'body',
    });
  }

  return {
    ...slide,
    index: idx + 1,
    blocks,
    images: slide.images ? slide.images.map((img) => ({ ...img })) : undefined,
  };
}

function syncSlideSummary(
  slide: SlidePreviewItem,
  idx: number
): SlidePreviewItem {
  const blocks = slide.blocks ?? [];
  const nonTableBlocks = blocks.filter((block) => block.kind !== 'table');
  const titleBlock =
    nonTableBlocks.find(
      (block) => block.role === 'title' && block.text.trim()
    ) ?? nonTableBlocks.find((block) => block.text.trim());
  const subtitleBlock = nonTableBlocks.find(
    (block) => block.role === 'subtitle' && block.text.trim()
  );
  const bodyBullets = nonTableBlocks
    .filter((block) => block !== titleBlock && block !== subtitleBlock)
    .flatMap((block) =>
      block.text
        .split('\n')
        .map((line) => line.replace(/^[•\-*]\s*/, '').trim())
        .filter(Boolean)
    );

  return {
    ...slide,
    index: idx + 1,
    title: titleBlock?.text.split('\n')[0]?.trim() || `Slide ${idx + 1}`,
    subtitle: subtitleBlock?.text.trim() || undefined,
    bullets: bodyBullets,
  };
}

function initializeDeckSlides(doc: DocumentItem): SlidePreviewItem[] {
  if (doc.slidesPreview.length > 0) {
    return doc.slidesPreview.map((slide, idx) =>
      syncSlideSummary(normalizeSlideWithBlocks(slide, idx), idx)
    );
  }

  const fallbackBullets =
    doc.outline.length > 0
      ? doc.outline.map((section) => `${section.romanIndex}. ${section.title}`)
      : [doc.contentPreview || 'Nhập nội dung trình chiếu của bạn tại đây.'];

  const fallbackSlide: SlidePreviewItem = {
    id: `${doc.id}-slide-1`,
    index: 1,
    title: doc.title,
    subtitle: doc.folder ?? 'Bản trình chiếu LITE',
    bullets: fallbackBullets,
    accentColor: '#4F46E5',
    speakerNote: doc.contentPreview,
  };

  return [syncSlideSummary(normalizeSlideWithBlocks(fallbackSlide, 0), 0)];
}

function readImageFileAsSrc(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    if (file.size <= 3 * 1024 * 1024) {
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === 'string' && reader.result) {
          resolve(reader.result);
        } else {
          reject(new Error('Không thể đọc dữ liệu ảnh.'));
        }
      };
      reader.onerror = () => {
        try {
          resolve(URL.createObjectURL(file));
        } catch {
          reject(new Error('Không thể nạp tệp ảnh.'));
        }
      };
      reader.readAsDataURL(file);
      return;
    }

    try {
      resolve(URL.createObjectURL(file));
    } catch (err) {
      reject(err instanceof Error ? err : new Error('Không thể nạp tệp ảnh.'));
    }
  });
}

interface AutoGrowTextareaProps {
  value: string;
  onChange: (value: string) => void;
  onFocus: () => void;
  disabled: boolean;
  placeholder: string;
  ariaLabel: string;
  scaledFontPx: number;
  bold: boolean;
  italic: boolean;
  underline?: boolean;
  color: string;
  align: 'left' | 'center' | 'right';
  fontFamily?: string;
  autoFocus?: boolean;
  maxHeightFull?: boolean;
}

function AutoGrowTextarea({
  value,
  onChange,
  onFocus,
  disabled,
  placeholder,
  ariaLabel,
  scaledFontPx,
  bold,
  italic,
  underline = false,
  color,
  align,
  fontFamily,
  autoFocus = false,
  maxHeightFull = false,
}: AutoGrowTextareaProps) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useLayoutEffect(() => {
    const el = textareaRef.current;
    if (!el) {
      return;
    }
    el.style.height = '0px';
    el.style.height = `${Math.max(el.scrollHeight, Math.ceil(scaledFontPx * 1.35))}px`;
  }, [value, scaledFontPx, bold, italic, underline, align, fontFamily]);

  useEffect(() => {
    const el = textareaRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const observer = new ResizeObserver(() => {
      el.style.height = '0px';
      el.style.height = `${Math.max(el.scrollHeight, Math.ceil(scaledFontPx * 1.35))}px`;
    });
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, [scaledFontPx]);

  return (
    <textarea
      ref={textareaRef}
      value={value}
      onChange={(event) => onChange(event.target.value)}
      onFocus={onFocus}
      onPointerDown={(event) => {
        if (!disabled) {
          event.stopPropagation();
        }
      }}
      disabled={disabled}
      autoFocus={autoFocus}
      rows={1}
      placeholder={placeholder}
      aria-label={ariaLabel}
      style={{
        fontSize: `${scaledFontPx.toFixed(2)}px`,
        color,
        textAlign: align,
        fontWeight: bold ? 700 : 400,
        fontStyle: italic ? 'italic' : 'normal',
        textDecoration: underline ? 'underline' : 'none',
        fontFamily: fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
        maxHeight: maxHeightFull ? '100%' : undefined,
      }}
      className={`w-full bg-transparent border-0 outline-none resize-none leading-[1.16] tracking-tight placeholder:opacity-45 disabled:pointer-events-none ${
        maxHeightFull ? 'overflow-y-auto no-scrollbar' : 'overflow-hidden'
      }`}
    />
  );
}

interface PositionedReadOnlySlideProps {
  doc: DocumentItem;
  slide: SlidePreviewItem;
  slideIndex: number;
  totalSlides: number;
  stageWidthPx: number;
  underlinedBlockIds?: Record<string, boolean>;
}

function ReadOnlySlideStage({
  doc,
  slide,
  slideIndex,
  totalSlides,
  stageWidthPx,
  underlinedBlockIds,
}: PositionedReadOnlySlideProps) {
  const fontScale =
    stageWidthPx > 0 ? stageWidthPx / REFERENCE_SLIDE_WIDTH_PT : 1;
  const isPositioned = isSlideMasterPositioned(slide);
  const hasLightBg = isLightHexBackground(slide.backgroundColor);
  const accent = slide.accentColor || '#4F46E5';
  const blocks = slide.blocks ?? [];
  const images = slide.images ?? [];
  const hasImages = images.length > 0;
  const hasMediaColumn = hasImages || Boolean(slide.metricHighlight);

  if (isPositioned) {
    return (
      <div
        className="relative w-full h-full overflow-hidden select-none"
        style={buildSlideBackgroundStyle(slide)}
      >
        {blocks.map((block, bIdx) => {
          const left = block.xPct ?? 6;
          const top = block.yPct ?? 10 + bIdx * 14;
          const width = block.wPct ?? 88;
          const height = block.hPct;
          const scaledFontPx = Math.max(
            4,
            (block.fontSize || (block.kind === 'table' ? 12 : 18)) * fontScale
          );
          const justifyClass =
            block.valign === 'middle'
              ? 'justify-center'
              : block.valign === 'bottom'
                ? 'justify-end'
                : 'justify-start';
          const textColor =
            block.color ?? (hasLightBg ? '#0F172A' : '#FFFFFF');

          if (
            block.kind === 'table' &&
            Array.isArray(block.tableRows) &&
            block.tableRows.length > 0
          ) {
            return (
              <div
                key={block.id}
                style={{
                  position: 'absolute',
                  left: `${left}%`,
                  top: `${top}%`,
                  width: `${width}%`,
                  height: height !== undefined ? `${height}%` : 'auto',
                  borderColor: block.borderColor,
                  fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                }}
                className={`overflow-hidden ${block.borderColor ? 'border' : ''}`}
              >
                <table className="w-full h-full border-collapse table-fixed">
                  <tbody>
                    {block.tableRows.map((row, rIdx) => (
                      <tr key={`${block.id}-ro-r-${rIdx}`}>
                        {row.map((cell, cIdx) => (
                          <td
                            key={`${block.id}-ro-c-${rIdx}-${cIdx}`}
                            className="border px-1.5 py-0.5 align-middle truncate"
                            style={{
                              borderColor:
                                block.borderColor ??
                                (hasLightBg
                                  ? 'rgba(15,23,42,0.22)'
                                  : 'rgba(255,255,255,0.22)'),
                              backgroundColor:
                                rIdx === 0
                                  ? block.fillColor ??
                                    (hasLightBg
                                      ? 'rgba(15,23,42,0.06)'
                                      : 'rgba(255,255,255,0.08)')
                                  : 'transparent',
                              fontSize: `${scaledFontPx.toFixed(2)}px`,
                              fontWeight:
                                rIdx === 0 || block.bold ? 700 : 400,
                              fontStyle: block.italic ? 'italic' : 'normal',
                              textDecoration: underlinedBlockIds?.[block.id]
                                ? 'underline'
                                : 'none',
                              color: textColor,
                              textAlign: block.align ?? 'left',
                              fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                            }}
                          >
                            {cell}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            );
          }

          return (
            <div
              key={block.id}
              style={{
                position: 'absolute',
                left: `${left}%`,
                top: `${top}%`,
                width: `${width}%`,
                height: height !== undefined ? `${height}%` : 'auto',
                backgroundColor: block.fillColor,
                borderColor: block.borderColor,
              }}
              className={`flex flex-col ${justifyClass} overflow-hidden px-1 py-0.5 ${
                block.borderColor ? 'border' : ''
              }`}
            >
              <p
                className="whitespace-pre-wrap break-words leading-[1.16] tracking-tight"
                style={{
                  fontSize: `${scaledFontPx.toFixed(2)}px`,
                  fontWeight: block.bold || block.role === 'title' ? 700 : 400,
                  fontStyle: block.italic ? 'italic' : 'normal',
                  textDecoration: underlinedBlockIds?.[block.id]
                    ? 'underline'
                    : 'none',
                  color: textColor,
                  textAlign: block.align ?? 'left',
                  fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                }}
              >
                {block.text}
              </p>
            </div>
          );
        })}

        {images.map((imageItem, iIdx) => (
          <div
            key={imageItem.id}
            style={{
              position: 'absolute',
              left: `${imageItem.xPct ?? 56}%`,
              top: `${imageItem.yPct ?? 20 + iIdx * 8}%`,
              width: `${imageItem.wPct ?? 36}%`,
              height: `${imageItem.hPct ?? 44}%`,
            }}
            className="overflow-hidden flex items-center justify-center"
          >
            <img
              src={imageItem.src}
              alt={imageItem.alt ?? `Hình ${iIdx + 1}`}
              draggable={false}
              className="w-full h-full object-contain"
            />
          </div>
        ))}
      </div>
    );
  }

  // Legacy Flow Slide Layout (with proportional fontScale)
  return (
    <div
      className="relative w-full h-full flex flex-col justify-between overflow-hidden select-none"
      style={{
        ...buildSlideBackgroundStyle(slide),
        padding: `${Math.max(8, Math.round(32 * fontScale))}px`,
      }}
    >
      <div className="shrink-0 flex items-center justify-between gap-2">
        <div className="inline-flex items-center gap-2 min-w-0">
          <span
            className="rounded-full shrink-0"
            style={{
              width: `${Math.max(12, Math.round(28 * fontScale))}px`,
              height: `${Math.max(3, Math.round(7 * fontScale))}px`,
              backgroundColor: accent,
            }}
          />
          <span
            className={`font-semibold uppercase tracking-widest truncate ${
              hasLightBg ? 'text-slate-600' : 'text-slate-300/90'
            }`}
            style={{ fontSize: `${Math.max(6, 11 * fontScale)}px` }}
          >
            {doc.folder ?? 'Bản trình chiếu'}
          </span>
        </div>
        <span
          className={`shrink-0 rounded-full border font-semibold tabular-nums ${
            hasLightBg
              ? 'bg-slate-900/10 border-slate-900/10 text-slate-800'
              : 'bg-black/45 border-white/10 text-white/95'
          }`}
          style={{
            fontSize: `${Math.max(6, 11 * fontScale)}px`,
            padding: `${Math.max(1, Math.round(3 * fontScale))}px ${Math.max(4, Math.round(10 * fontScale))}px`,
          }}
        >
          Slide {slideIndex + 1} / {totalSlides}
        </span>
      </div>

      <div
        className={`my-auto grid items-center min-h-0 overflow-hidden ${
          hasMediaColumn && blocks.length > 0 ? 'grid-cols-12' : 'grid-cols-1'
        }`}
        style={{ gap: `${Math.max(6, Math.round(20 * fontScale))}px` }}
      >
        {blocks.length > 0 && (
          <div
            className={hasMediaColumn ? 'col-span-8' : 'col-span-1'}
            style={{
              display: 'flex',
              flexDirection: 'column',
              gap: `${Math.max(4, Math.round(12 * fontScale))}px`,
            }}
          >
            {blocks.map((block) => (
              <p
                key={block.id}
                className="whitespace-pre-wrap break-words leading-[1.16]"
                style={{
                  fontSize: `${Math.max(5, (block.fontSize || 18) * fontScale)}px`,
                  fontWeight: block.bold || block.role === 'title' ? 700 : 400,
                  fontStyle: block.italic ? 'italic' : 'normal',
                  textDecoration: underlinedBlockIds?.[block.id]
                    ? 'underline'
                    : 'none',
                  color: block.color ?? (hasLightBg ? '#0F172A' : '#FFFFFF'),
                  textAlign: block.align ?? 'left',
                  fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                }}
              >
                {block.text}
              </p>
            ))}
          </div>
        )}

        {hasImages ? (
          <div
            className={`${
              blocks.length > 0 ? 'col-span-4' : 'col-span-1'
            } grid gap-2 ${images.length > 1 ? 'grid-cols-2' : 'grid-cols-1'}`}
          >
            {images.map((imageItem, idx) => (
              <div
                key={imageItem.id}
                className="rounded-xl overflow-hidden border border-white/15 bg-black/25 flex items-center justify-center p-1"
              >
                <img
                  src={imageItem.src}
                  alt={imageItem.alt ?? `Ảnh ${idx + 1}`}
                  className="w-full h-full max-h-[55vh] object-contain rounded-lg"
                />
              </div>
            ))}
          </div>
        ) : (
          slide.metricHighlight && (
            <div
              className="col-span-4 rounded-2xl border flex flex-col items-end justify-center text-right backdrop-blur-md"
              style={{
                backgroundColor: `${accent}24`,
                borderColor: `${accent}66`,
                padding: `${Math.max(6, Math.round(20 * fontScale))}px`,
              }}
            >
              <span
                className={`font-bold leading-none tracking-tight tabular-nums ${
                  hasLightBg ? 'text-slate-900' : 'text-white'
                }`}
                style={{ fontSize: `${Math.max(8, 30 * fontScale)}px` }}
              >
                {slide.metricHighlight.value}
              </span>
              <span
                className={`mt-1.5 leading-snug ${
                  hasLightBg ? 'text-slate-700' : 'text-slate-200/90'
                }`}
                style={{ fontSize: `${Math.max(5, 13 * fontScale)}px` }}
              >
                {slide.metricHighlight.label}
              </span>
            </div>
          )
        )}
      </div>

      <div className="shrink-0 flex items-center justify-between gap-2 pt-1">
        <div
          className={`flex-1 h-1 rounded-full overflow-hidden ${
            hasLightBg ? 'bg-slate-900/10' : 'bg-white/10'
          }`}
        >
          <div
            className="h-full rounded-full"
            style={{
              width: `${((slideIndex + 1) / totalSlides) * 100}%`,
              backgroundColor: accent,
            }}
          />
        </div>
        <span
          className={`truncate max-w-[45%] ${
            hasLightBg ? 'text-slate-500' : 'text-slate-400'
          }`}
          style={{ fontSize: `${Math.max(5, 11 * fontScale)}px` }}
        >
          {doc.title}
        </span>
      </div>
    </div>
  );
}

interface GridSlideThumbnailProps {
  doc: DocumentItem;
  slide: SlidePreviewItem;
  slideIndex: number;
  totalSlides: number;
  isActive: boolean;
  onSelect: () => void;
  underlinedBlockIds?: Record<string, boolean>;
}

function GridSlideThumbnail({
  doc,
  slide,
  slideIndex,
  totalSlides,
  isActive,
  onSelect,
  underlinedBlockIds,
}: GridSlideThumbnailProps) {
  const cardRef = useRef<HTMLDivElement | null>(null);
  const [widthPx, setWidthPx] = useState<number>(260);
  const aspectRatio = getSafeAspectRatio(slide, doc);

  useEffect(() => {
    const el = cardRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const update = () => {
      const w = el.clientWidth;
      if (w > 0) {
        setWidthPx(w);
      }
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, []);

  return (
    <button
      type="button"
      onClick={onSelect}
      className={`group text-left rounded-2xl p-2 border transition-all ${
        isActive
          ? 'bg-indigo-500/20 border-indigo-400 ring-2 ring-indigo-400/60'
          : 'bg-white/[0.04] hover:bg-white/[0.08] border-white/10 hover:border-white/25'
      }`}
    >
      <div
        ref={cardRef}
        style={{ aspectRatio: String(aspectRatio) }}
        className="w-full rounded-xl overflow-hidden border border-white/10 shadow-md pointer-events-none"
      >
        <ReadOnlySlideStage
          doc={doc}
          slide={slide}
          slideIndex={slideIndex}
          totalSlides={totalSlides}
          stageWidthPx={widthPx}
          underlinedBlockIds={underlinedBlockIds}
        />
      </div>
      <div className="mt-2 px-1 flex items-center justify-between gap-2">
        <span className="text-xs font-medium text-slate-200 truncate">
          {slide.title}
        </span>
        <span
          className={`shrink-0 px-1.5 py-0.5 rounded text-[10px] font-bold tabular-nums ${
            isActive
              ? 'bg-indigo-500 text-white'
              : 'bg-white/10 text-slate-300'
          }`}
        >
          {slideIndex + 1}
        </span>
      </div>
    </button>
  );
}

export function SlideWorkspace({
  document,
  onBackHome,
  onOpenModal,
  onSaveDocument,
  isModalOpen = false,
  autoSaveEnabled = false,
  isHydratingCloudFile = false,
  onRetryCloudDownload,
  collaborators: _collaborators = [],
  selfSessionId,
  collabPermission = 'owner',
  onPresenceUpdate,
  onBroadcastPatch,
  incomingRemotePatch = null,
}: SlideWorkspaceProps) {
  const isReadOnly = collabPermission === 'view';
  const [slides, setSlides] = useState<SlidePreviewItem[]>(() =>
    initializeDeckSlides(document)
  );
  const [activeSlideIdx, setActiveSlideIdx] = useState<number>(0);
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null);
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null);
  const [newlyAddedBlockId, setNewlyAddedBlockId] = useState<string | null>(
    null
  );
  const [isLaserActive, setIsLaserActive] = useState<boolean>(false);
  const [laserPos, setLaserPos] = useState<LaserPosition>({
    x: 50,
    y: 50,
    visible: false,
  });
  const [elapsedSeconds, setElapsedSeconds] = useState<number>(0);
  const [isTimerRunning, setIsTimerRunning] = useState<boolean>(false);
  const [isMobileDrawerOpen, setIsMobileDrawerOpen] = useState<boolean>(false);
  const [isMobileSheetOpen, setIsMobileSheetOpen] = useState<boolean>(false);
  const [underlinedBlockIds, setUnderlinedBlockIds] = useState<
    Record<string, boolean>
  >({});
  const [isNotesExpanded, setIsNotesExpanded] = useState<boolean>(true);
  const [isDirty, setIsDirty] = useState<boolean>(false);
  const [saveFeedback, setSaveFeedback] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportError, setExportError] = useState<string | null>(null);

  // ResizeObserver measured sizes
  const [editorViewportSize, setEditorViewportSize] = useState<{
    width: number;
    height: number;
  }>({ width: 960, height: 540 });
  const [slideshowViewportSize, setSlideshowViewportSize] = useState<{
    width: number;
    height: number;
  }>({ width: 1280, height: 720 });

  // Fullscreen Slideshow State
  const [isSlideshowActive, setIsSlideshowActive] = useState<boolean>(false);
  const [slideshowTransition, setSlideshowTransition] =
    useState<SlideshowTransition>('slide');
  const [slideshowDirection, setSlideshowDirection] = useState<1 | -1>(1);
  const [isSlideshowLaser, setIsSlideshowLaser] = useState<boolean>(false);
  const [slideshowLaserPos, setSlideshowLaserPos] = useState<LaserPosition>({
    x: 50,
    y: 50,
    visible: false,
  });
  const [isSlideshowGridOpen, setIsSlideshowGridOpen] =
    useState<boolean>(false);
  const [draggingId, setDraggingId] = useState<string | null>(null);

  const loadedDocIdRef = useRef<string>(document.id);
  const prevRawBufferRef = useRef<ArrayBuffer | undefined>(
    document.rawArrayBuffer
  );
  const wasHydratingRef = useRef<boolean>(isHydratingCloudFile);
  const isDirtyRef = useRef<boolean>(isDirty);
  isDirtyRef.current = isDirty;

  const uidRef = useRef<number>(1);
  const thumbnailRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const desktopThumbScrollRef = useRef<HTMLDivElement | null>(null);
  const saveTimerRef = useRef<number | null>(null);
  const exportErrorTimerRef = useRef<number | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const editorViewportRef = useRef<HTMLDivElement | null>(null);
  const editorCanvasRef = useRef<HTMLDivElement | null>(null);
  const slideshowViewportRef = useRef<HTMLDivElement | null>(null);
  const enteredFullscreenRef = useRef<boolean>(false);
  const justDraggedRef = useRef<boolean>(false);
  const activeDragCleanupRef = useRef<(() => void) | null>(null);
  const localMutationPendingRef = useRef<boolean>(false);
  const lastRemotePatchSentAtRef = useRef<number>(0);
  const lastBroadcastAtRef = useRef<number>(0);
  const broadcastThrottleTimerRef = useRef<number | null>(null);
  const pendingBroadcastSlidesRef = useRef<SlidePreviewItem[] | null>(null);
  const typingResetTimerRef = useRef<number | null>(null);

  const documentRef = useRef<DocumentItem>(document);
  documentRef.current = document;
  const onPresenceUpdateRef = useRef(onPresenceUpdate);
  onPresenceUpdateRef.current = onPresenceUpdate;
  const onBroadcastPatchRef = useRef(onBroadcastPatch);
  onBroadcastPatchRef.current = onBroadcastPatch;

  // Reset per-document state when opening another presentation, or sync hydrated cloud slides
  useEffect(() => {
    const isNewDoc = loadedDocIdRef.current !== document.id;
    const rawBufferArrived =
      Boolean(document.rawArrayBuffer) &&
      document.rawArrayBuffer !== prevRawBufferRef.current;
    const finishedHydrating =
      wasHydratingRef.current &&
      !isHydratingCloudFile &&
      document.slidesPreview.length > 0;

    prevRawBufferRef.current = document.rawArrayBuffer;
    wasHydratingRef.current = isHydratingCloudFile;

    if (isNewDoc) {
      loadedDocIdRef.current = document.id;
      setSlides(initializeDeckSlides(document));
      setActiveSlideIdx(0);
      setSelectedBlockId(null);
      setSelectedImageId(null);
      setNewlyAddedBlockId(null);
      setIsLaserActive(false);
      setLaserPos({ x: 50, y: 50, visible: false });
      setElapsedSeconds(0);
      setIsTimerRunning(false);
      setIsMobileDrawerOpen(false);
      setIsMobileSheetOpen(false);
      setUnderlinedBlockIds({});
      setIsDirty(false);
      setSaveFeedback(false);
      setExportError(null);
      setIsSlideshowActive(false);
      setIsSlideshowGridOpen(false);
      setIsSlideshowLaser(false);
      localMutationPendingRef.current = false;
      lastRemotePatchSentAtRef.current = 0;
      pendingBroadcastSlidesRef.current = null;
      if (broadcastThrottleTimerRef.current !== null) {
        window.clearTimeout(broadcastThrottleTimerRef.current);
        broadcastThrottleTimerRef.current = null;
      }
      if (typingResetTimerRef.current !== null) {
        window.clearTimeout(typingResetTimerRef.current);
        typingResetTimerRef.current = null;
      }
      return;
    }

    if (!isDirtyRef.current && (rawBufferArrived || finishedHydrating)) {
      localMutationPendingRef.current = false;
      setSlides(initializeDeckSlides(document));
      setSelectedBlockId(null);
      setSelectedImageId(null);
    }
  }, [document, isHydratingCloudFile]);

  // Clear selection/mobile sheet if switched to read-only permission
  useEffect(() => {
    if (isReadOnly) {
      setSelectedBlockId(null);
      setSelectedImageId(null);
      setIsMobileSheetOpen(false);
    }
  }, [isReadOnly]);

  // Real-time presentation stopwatch
  useEffect(() => {
    if (!isTimerRunning) {
      return;
    }
    const intervalId = window.setInterval(() => {
      setElapsedSeconds((prev) => prev + 1);
    }, 1000);
    return () => {
      window.clearInterval(intervalId);
    };
  }, [isTimerRunning]);

  // Clear laser pointer immediately when disabled
  useEffect(() => {
    if (!isLaserActive) {
      setLaserPos((prev) =>
        prev.visible ? { ...prev, visible: false } : prev
      );
    } else {
      setSelectedBlockId(null);
      setSelectedImageId(null);
    }
  }, [isLaserActive]);

  // Measure editor stage container with ResizeObserver
  useEffect(() => {
    const el = editorViewportRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateSize = () => {
      const w = el.clientWidth;
      const h = el.clientHeight;
      if (w > 0 && h > 0) {
        setEditorViewportSize({ width: w, height: h });
      }
    };
    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, []);

  // Measure slideshow viewport with ResizeObserver when slideshow is active
  useEffect(() => {
    if (!isSlideshowActive) {
      return;
    }
    const el = slideshowViewportRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateSize = () => {
      const rect = el.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        setSlideshowViewportSize({
          width: Math.max(200, Math.floor(rect.width - 16)),
          height: Math.max(150, Math.floor(rect.height - 16)),
        });
      }
    };
    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, [isSlideshowActive]);

  // Sync browser fullscreenchange -> slideshow state
  useEffect(() => {
    const handleFullscreenChange = () => {
      const hasFs = Boolean(window.document.fullscreenElement);
      if (hasFs) {
        enteredFullscreenRef.current = true;
      } else if (enteredFullscreenRef.current) {
        enteredFullscreenRef.current = false;
        setIsSlideshowGridOpen(false);
        setIsSlideshowLaser(false);
        setIsSlideshowActive(false);
      }
    };

    window.document.addEventListener(
      'fullscreenchange',
      handleFullscreenChange
    );
    return () => {
      window.document.removeEventListener(
        'fullscreenchange',
        handleFullscreenChange
      );
    };
  }, []);

  // Cleanup timers & active drag listeners on unmount
  useEffect(() => {
    return () => {
      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
      if (broadcastThrottleTimerRef.current !== null) {
        window.clearTimeout(broadcastThrottleTimerRef.current);
      }
      if (typingResetTimerRef.current !== null) {
        window.clearTimeout(typingResetTimerRef.current);
      }
      activeDragCleanupRef.current?.();
      if (
        enteredFullscreenRef.current &&
        window.document.fullscreenElement &&
        window.document.exitFullscreen
      ) {
        void window.document.exitFullscreen().catch(() => {});
      }
    };
  }, []);

  const safeSlideIdx =
    slides.length > 0 ? Math.min(activeSlideIdx, slides.length - 1) : 0;

  // Broadcast active slideIndex via Presence whenever current slide changes
  useEffect(() => {
    onPresenceUpdateRef.current?.({
      slideIndex: safeSlideIdx,
      cursorLabel: `Slide ${safeSlideIdx + 1}`,
    });
  }, [safeSlideIdx, document.id]);

  const notifyTypingActivity = useCallback(
    (activeNow = true) => {
      onPresenceUpdateRef.current?.({
        slideIndex: safeSlideIdx,
        cursorLabel: `Slide ${safeSlideIdx + 1}`,
        isTyping: activeNow,
      });
      if (typingResetTimerRef.current !== null) {
        window.clearTimeout(typingResetTimerRef.current);
        typingResetTimerRef.current = null;
      }
      if (activeNow) {
        typingResetTimerRef.current = window.setTimeout(() => {
          onPresenceUpdateRef.current?.({
            slideIndex: safeSlideIdx,
            cursorLabel: `Slide ${safeSlideIdx + 1}`,
            isTyping: false,
          });
          typingResetTimerRef.current = null;
        }, 1200);
      }
    },
    [safeSlideIdx]
  );

  // Throttled broadcast (~80ms) when local slide edits occur
  useEffect(() => {
    if (!localMutationPendingRef.current) {
      return;
    }
    localMutationPendingRef.current = false;
    if (isReadOnly || !onBroadcastPatchRef.current) {
      return;
    }

    const emitSlides = (targetSlides: SlidePreviewItem[]) => {
      lastBroadcastAtRef.current = Date.now();
      onBroadcastPatchRef.current?.({
        kind: 'slide_patch',
        slidesPreview: targetSlides.map((slide, idx) =>
          syncSlideSummary(slide, idx)
        ),
        slideAspectRatio:
          documentRef.current.slideAspectRatio ?? targetSlides[0]?.slideWidthRatio,
      });
    };

    const elapsed = Date.now() - lastBroadcastAtRef.current;
    if (elapsed >= 80 && broadcastThrottleTimerRef.current === null) {
      pendingBroadcastSlidesRef.current = null;
      emitSlides(slides);
      return;
    }

    pendingBroadcastSlidesRef.current = slides;
    if (broadcastThrottleTimerRef.current === null) {
      const waitMs = Math.max(16, 80 - elapsed);
      broadcastThrottleTimerRef.current = window.setTimeout(() => {
        broadcastThrottleTimerRef.current = null;
        const latest = pendingBroadcastSlidesRef.current;
        pendingBroadcastSlidesRef.current = null;
        if (latest) {
          emitSlides(latest);
        }
      }, waitMs);
    }
  }, [slides, isReadOnly]);

  // Apply incoming remote slide_patch without jumping viewer's active slide or echoing back
  useEffect(() => {
    if (!incomingRemotePatch || incomingRemotePatch.kind !== 'slide_patch') {
      return;
    }
    if (
      selfSessionId &&
      incomingRemotePatch.senderSessionId === selfSessionId
    ) {
      return;
    }
    if (incomingRemotePatch.sentAt <= lastRemotePatchSentAtRef.current) {
      return;
    }
    if (
      !Array.isArray(incomingRemotePatch.slidesPreview) ||
      incomingRemotePatch.slidesPreview.length === 0
    ) {
      return;
    }

    lastRemotePatchSentAtRef.current = incomingRemotePatch.sentAt;
    localMutationPendingRef.current = false;
    const normalized = incomingRemotePatch.slidesPreview.map((slide, idx) =>
      syncSlideSummary(normalizeSlideWithBlocks(slide, idx), idx)
    );
    setSlides(normalized);
    setActiveSlideIdx((prev) =>
      prev >= normalized.length ? Math.max(0, normalized.length - 1) : prev
    );
  }, [incomingRemotePatch, selfSessionId]);
  const currentSlide = slides[safeSlideIdx];
  const currentBlocks = currentSlide.blocks ?? [];
  const currentImages = currentSlide.images ?? [];
  const selectedBlock =
    currentBlocks.find((block) => block.id === selectedBlockId) ?? null;
  const selectedImage =
    currentImages.find((img) => img.id === selectedImageId) ?? null;
  const accentColor = currentSlide.accentColor || '#4F46E5';
  const hasLightSlideBg = isLightHexBackground(currentSlide.backgroundColor);
  const currentAspectRatio = getSafeAspectRatio(currentSlide, document);
  const isPositionedSlide = isSlideMasterPositioned(currentSlide);

  const editorFitWidthPx = Math.max(
    240,
    Math.min(
      editorViewportSize.width,
      editorViewportSize.height * currentAspectRatio,
      1140
    )
  );
  const editorFitHeightPx = editorFitWidthPx / currentAspectRatio;
  const editorFontScale = editorFitWidthPx / REFERENCE_SLIDE_WIDTH_PT;

  const slideshowFitWidthPx = Math.max(
    240,
    Math.min(
      slideshowViewportSize.width,
      slideshowViewportSize.height * currentAspectRatio
    )
  );
  const slideshowFitHeightPx = slideshowFitWidthPx / currentAspectRatio;

  const buildLiveDocument = useCallback((): DocumentItem => {
    const baseDoc = documentRef.current;
    const syncedSlides = slides.map((slide, idx) =>
      syncSlideSummary(slide, idx)
    );
    const derivedOutline: OutlineSection[] = syncedSlides.map((slide, idx) => ({
      id: `${baseDoc.id}-slide-sec-${idx + 1}`,
      romanIndex: toRomanNumeral(idx + 1),
      title: slide.title,
      summary: slide.subtitle ?? slide.bullets[0] ?? '',
      keyPoints: slide.bullets,
      pageNumber: idx + 1,
    }));

    return {
      ...baseDoc,
      slidesPreview: syncedSlides,
      outline: derivedOutline.length > 0 ? derivedOutline : baseDoc.outline,
      pageCount: syncedSlides.length,
      contentPreview:
        syncedSlides[0]?.bullets[0] ??
        syncedSlides[0]?.title ??
        baseDoc.contentPreview,
    };
  }, [slides]);

  const handleSave = useCallback(() => {
    if (isReadOnly) {
      return;
    }
    const updatedDoc = buildLiveDocument();
    onSaveDocument(updatedDoc);
    setIsDirty(false);
    setSaveFeedback(true);

    if (saveTimerRef.current !== null) {
      window.clearTimeout(saveTimerRef.current);
    }
    saveTimerRef.current = window.setTimeout(() => {
      setSaveFeedback(false);
      saveTimerRef.current = null;
    }, 2200);
  }, [buildLiveDocument, isReadOnly, onSaveDocument]);

  const handleRenameTitle = useCallback(
    (newTitle: string) => {
      if (isReadOnly) {
        return;
      }
      const liveDoc = buildLiveDocument();
      const renamedDoc: DocumentItem = {
        ...liveDoc,
        title: newTitle,
        thumbnail: createThumbnailDataUri(
          document.type,
          newTitle,
          accentColor
        ),
        originalFileName: `${newTitle}.${document.type}`,
        updatedAt: 'Vừa chỉnh sửa',
      };
      onSaveDocument(renamedDoc);
      setIsDirty(false);
      setSaveFeedback(true);

      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
      saveTimerRef.current = window.setTimeout(() => {
        setSaveFeedback(false);
        saveTimerRef.current = null;
      }, 2200);
    },
    [accentColor, buildLiveDocument, document.type, isReadOnly, onSaveDocument]
  );

  // Auto-save with 1800ms debounce when autoSaveEnabled is on
  useEffect(() => {
    if (isReadOnly || !autoSaveEnabled || !isDirty) {
      return;
    }
    const timerId = window.setTimeout(() => {
      handleSave();
    }, 1800);
    return () => {
      window.clearTimeout(timerId);
    };
  }, [autoSaveEnabled, isDirty, isReadOnly, slides, handleSave]);

  const startSlideshow = useCallback(() => {
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setIsLaserActive(false);
    setIsSlideshowLaser(false);
    setSlideshowLaserPos({ x: 50, y: 50, visible: false });
    setIsSlideshowGridOpen(false);
    setSlideshowDirection(1);
    setIsSlideshowActive(true);
    enteredFullscreenRef.current = false;

    try {
      const req = window.document.documentElement.requestFullscreen?.();
      if (req && typeof req.then === 'function') {
        void req
          .then(() => {
            enteredFullscreenRef.current = true;
          })
          .catch(() => {
            enteredFullscreenRef.current = false;
          });
      }
    } catch {
      enteredFullscreenRef.current = false;
    }
  }, []);

  const exitSlideshow = useCallback(() => {
    setIsSlideshowGridOpen(false);
    setIsSlideshowLaser(false);
    setIsSlideshowActive(false);
    enteredFullscreenRef.current = false;

    if (window.document.fullscreenElement && window.document.exitFullscreen) {
      try {
        void window.document.exitFullscreen().catch(() => {});
      } catch {
        // Ignore fullscreen exit errors
      }
    }
  }, []);

  // Keyboard navigation + Slideshow shortcuts + Ctrl/Cmd+S save
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (isModalOpen) {
        return;
      }

      if (isSlideshowActive) {
        if (event.key === 'Escape') {
          event.preventDefault();
          if (isSlideshowGridOpen) {
            setIsSlideshowGridOpen(false);
          } else {
            exitSlideshow();
          }
          return;
        }

        if (isSlideshowGridOpen) {
          return;
        }

        const target = event.target as HTMLElement | null;
        const isButtonFocused = target?.tagName === 'BUTTON';

        if (
          event.key === 'ArrowRight' ||
          event.key === 'PageDown' ||
          (event.key === ' ' && !isButtonFocused)
        ) {
          event.preventDefault();
          setSlideshowDirection(1);
          setActiveSlideIdx((prev) =>
            prev < slides.length - 1 ? prev + 1 : prev
          );
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault();
          setSlideshowDirection(-1);
          setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
        }
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        if (!isReadOnly) {
          handleSave();
        }
        return;
      }

      if (event.key === 'Escape') {
        if (isMobileSheetOpen) {
          setIsMobileSheetOpen(false);
          return;
        }
        if (isMobileDrawerOpen) {
          setIsMobileDrawerOpen(false);
          return;
        }
        if (isLaserActive) {
          setIsLaserActive(false);
          return;
        }
        if (selectedBlockId !== null || selectedImageId !== null) {
          setSelectedBlockId(null);
          setSelectedImageId(null);
        }
        return;
      }

      if (isEditableElement(event.target)) {
        return;
      }

      if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        event.preventDefault();
        setSelectedBlockId(null);
        setSelectedImageId(null);
        setActiveSlideIdx((prev) =>
          prev < slides.length - 1 ? prev + 1 : prev
        );
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setSelectedBlockId(null);
        setSelectedImageId(null);
        setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    slides.length,
    isMobileSheetOpen,
    isMobileDrawerOpen,
    isLaserActive,
    isModalOpen,
    isSlideshowActive,
    isSlideshowGridOpen,
    selectedBlockId,
    selectedImageId,
    handleSave,
    exitSlideshow,
  ]);

  // Keep active slide thumbnail smoothly scrolled into view
  useEffect(() => {
    const node = thumbnailRefs.current[currentSlide.id];
    node?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, [currentSlide.id]);

  const mutateSlideAt = useCallback(
    (targetIdx: number, updater: (slide: SlidePreviewItem) => SlidePreviewItem) => {
      if (isReadOnly) {
        return;
      }
      localMutationPendingRef.current = true;
      setSlides((prev) =>
        prev.map((slide, idx) =>
          idx === targetIdx ? syncSlideSummary(updater(slide), idx) : slide
        )
      );
      setIsDirty(true);
      setSaveFeedback(false);
      notifyTypingActivity(true);
    },
    [isReadOnly, notifyTypingActivity]
  );

  const handleAddSlide = () => {
    if (isReadOnly) {
      return;
    }
    uidRef.current += 1;
    const newSlideId = `${document.id}-custom-slide-${Date.now()}-${uidRef.current}`;
    const newTitleBlockId = `${newSlideId}-blk-title`;
    const inheritPositioned =
      isPositionedSlide || typeof document.slideAspectRatio === 'number';
    const inheritLight = isLightHexBackground(currentSlide.backgroundColor);

    const newSlide: SlidePreviewItem = syncSlideSummary(
      {
        id: newSlideId,
        index: safeSlideIdx + 2,
        title: 'Tiêu đề slide mới',
        bullets: ['Nhập nội dung chính tại đây...'],
        accentColor,
        speakerNote: '',
        ...(inheritPositioned
          ? {
              isMasterPositioned: true,
              slideWidthRatio:
                currentSlide.slideWidthRatio ?? document.slideAspectRatio,
              backgroundColor: currentSlide.backgroundColor ?? '#161D2E',
            }
          : {}),
        blocks: [
          {
            id: newTitleBlockId,
            text: 'Tiêu đề slide mới',
            fontSize: 32,
            bold: true,
            italic: false,
            color: inheritLight ? '#0F172A' : '#FFFFFF',
            align: 'left',
            role: 'title',
            ...(inheritPositioned
              ? { xPct: 6, yPct: 8, wPct: 88, hPct: 16, valign: 'middle' }
              : {}),
          },
          {
            id: `${newSlideId}-blk-body`,
            text: '• Nhập ý chính đầu tiên\n• Nhập ý chính tiếp theo',
            fontSize: 18,
            bold: false,
            italic: false,
            color: inheritLight ? '#334155' : '#E2E8F0',
            align: 'left',
            role: 'body',
            ...(inheritPositioned
              ? { xPct: 6, yPct: 28, wPct: 88, hPct: 58, valign: 'top' }
              : {}),
          },
        ],
      },
      safeSlideIdx + 1
    );

    setSlides((prev) => {
      const insertAt = Math.min(safeSlideIdx + 1, prev.length);
      const next = [...prev.slice(0, insertAt), newSlide, ...prev.slice(insertAt)];
      return next.map((item, idx) => ({ ...item, index: idx + 1 }));
    });
    setActiveSlideIdx(safeSlideIdx + 1);
    setSelectedBlockId(newTitleBlockId);
    setSelectedImageId(null);
    setNewlyAddedBlockId(newTitleBlockId);
    setIsLaserActive(false);
    setIsMobileDrawerOpen(false);
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleDuplicateSlide = () => {
    uidRef.current += 1;
    const cloneId = `${document.id}-dup-slide-${Date.now()}-${uidRef.current}`;
    const clonedSlide: SlidePreviewItem = {
      ...currentSlide,
      id: cloneId,
      blocks: currentBlocks.map((block, bIdx) => ({
        ...block,
        id: `${cloneId}-blk-${bIdx + 1}`,
        tableRows: block.tableRows
          ? block.tableRows.map((row) => [...row])
          : undefined,
      })),
      images: currentSlide.images
        ? currentSlide.images.map((img, iIdx) => ({
            ...img,
            id: `${cloneId}-img-${iIdx + 1}`,
          }))
        : undefined,
    };

    setSlides((prev) => {
      const insertAt = Math.min(safeSlideIdx + 1, prev.length);
      const next = [
        ...prev.slice(0, insertAt),
        clonedSlide,
        ...prev.slice(insertAt),
      ];
      return next.map((item, idx) => syncSlideSummary(item, idx));
    });
    setActiveSlideIdx(safeSlideIdx + 1);
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setIsMobileDrawerOpen(false);
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleDeleteSlide = () => {
    if (slides.length <= 1) {
      return;
    }
    setSlides((prev) => {
      const filtered = prev.filter((_, idx) => idx !== safeSlideIdx);
      return filtered.map((item, idx) => ({ ...item, index: idx + 1 }));
    });
    setActiveSlideIdx((prev) => Math.max(0, Math.min(prev, slides.length - 2)));
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleAddTextBlock = () => {
    uidRef.current += 1;
    const newBlockId = `${currentSlide.id}-blk-new-${Date.now()}-${uidRef.current}`;
    const defaultTextColor = hasLightSlideBg ? '#1E293B' : '#F1F5F9';
    const offset = (currentBlocks.length % 5) * 4;
    const newBlock: SlideTextBlock = {
      id: newBlockId,
      text: 'Khung nội dung mới',
      fontSize: 18,
      bold: false,
      italic: false,
      color: defaultTextColor,
      align: 'left',
      role: 'body',
      ...(isPositionedSlide
        ? {
            xPct: Math.min(54, 10 + offset),
            yPct: Math.min(72, 22 + offset),
            wPct: 44,
            hPct: 14,
            valign: 'top',
          }
        : {}),
    };

    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: [...(slide.blocks ?? []), newBlock],
    }));
    setIsLaserActive(false);
    setSelectedImageId(null);
    setSelectedBlockId(newBlockId);
    setNewlyAddedBlockId(newBlockId);
  };

  const handleTriggerAddImage = () => {
    setIsLaserActive(false);
    imageInputRef.current?.click();
  };

  const handleImageFileChange = async (
    event: ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) {
      return;
    }
    try {
      const src = await readImageFileAsSrc(file);
      uidRef.current += 1;
      const newImageId = `${currentSlide.id}-img-${Date.now()}-${uidRef.current}`;
      const offset = (currentImages.length % 4) * 5;
      const newImage: SlideImageItem = {
        id: newImageId,
        src,
        alt: file.name,
        ...(isPositionedSlide
          ? {
              xPct: Math.min(48, 24 + offset),
              yPct: Math.min(44, 20 + offset),
              wPct: 42,
              hPct: 42,
            }
          : {}),
      };
      mutateSlideAt(safeSlideIdx, (slide) => ({
        ...slide,
        images: [...(slide.images ?? []), newImage],
      }));
      setSelectedBlockId(null);
      setSelectedImageId(newImageId);
    } catch {
      setExportError('Không thể thêm tệp ảnh đã chọn.');
    }
  };

  const handleUpdateBlock = (
    blockId: string,
    patch: Partial<SlideTextBlock>
  ) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).map((block) =>
        block.id === blockId ? { ...block, ...patch } : block
      ),
    }));
  };

  const handleUpdateTableCell = (
    blockId: string,
    rowIdx: number,
    colIdx: number,
    nextVal: string
  ) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).map((block) => {
        if (block.id !== blockId || !block.tableRows) {
          return block;
        }
        const nextRows = block.tableRows.map((row, r) =>
          r === rowIdx ? row.map((cell, c) => (c === colIdx ? nextVal : cell)) : row
        );
        return {
          ...block,
          tableRows: nextRows,
          text: nextRows.map((row) => row.join(' | ')).join('\n'),
        };
      }),
    }));
  };

  const handleDeleteBlock = (blockId: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).filter((block) => block.id !== blockId),
    }));
    setUnderlinedBlockIds((prev) => {
      if (!prev[blockId]) {
        return prev;
      }
      const next = { ...prev };
      delete next[blockId];
      return next;
    });
    if (selectedBlockId === blockId) {
      setSelectedBlockId(null);
    }
  };

  const handleToggleBlockUnderline = (blockId: string) => {
    setUnderlinedBlockIds((prev) => ({
      ...prev,
      [blockId]: !prev[blockId],
    }));
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleApplyBlockPreset = (
    blockId: string,
    tag: FloatingBlockTag
  ) => {
    if (tag === 'h1') {
      handleUpdateBlock(blockId, { fontSize: 32, bold: true });
    } else if (tag === 'h2') {
      handleUpdateBlock(blockId, { fontSize: 24, bold: true });
    } else {
      handleUpdateBlock(blockId, { fontSize: 18, bold: false });
    }
  };

  const handleClearBlockFormat = (blockId: string) => {
    const defaultTextColor = hasLightSlideBg ? '#0F172A' : '#FFFFFF';
    handleUpdateBlock(blockId, {
      bold: false,
      italic: false,
      fillColor: undefined,
      color: defaultTextColor,
    });
    setUnderlinedBlockIds((prev) => {
      if (!prev[blockId]) {
        return prev;
      }
      const next = { ...prev };
      delete next[blockId];
      return next;
    });
  };

  const handlePreventMouseDown = (event: ReactMouseEvent<HTMLElement>) => {
    event.preventDefault();
  };

  const handleDeleteImage = (imageId: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      images: (slide.images ?? []).filter((img) => img.id !== imageId),
    }));
    if (selectedImageId === imageId) {
      setSelectedImageId(null);
    }
  };

  const handleChangeSlideBackground = (nextBgHex: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      backgroundColor: nextBgHex,
      backgroundGradient: undefined,
      backgroundImageSrc: undefined,
    }));
  };

  const handleSpeakerNoteChange = (value: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      speakerNote: value,
    }));
  };

  const handleExportPptx = async () => {
    if (isExporting) {
      return;
    }
    setIsExporting(true);
    setExportError(null);
    try {
      const syncedSlides = slides.map((slide, idx) =>
        syncSlideSummary(slide, idx)
      );
      const exportAspectRatio =
        document.slideAspectRatio ?? syncedSlides[0]?.slideWidthRatio;
      await exportSlidesToPptx(document.title, syncedSlides, exportAspectRatio);
    } catch {
      setExportError('Không thể xuất tệp .pptx. Vui lòng thử lại.');
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
      exportErrorTimerRef.current = window.setTimeout(() => {
        setExportError(null);
        exportErrorTimerRef.current = null;
      }, 3500);
    } finally {
      setIsExporting(false);
    }
  };

  const handlePrevSlide = () => {
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setNewlyAddedBlockId(null);
    setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
  };

  const handleNextSlide = () => {
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setNewlyAddedBlockId(null);
    setActiveSlideIdx((prev) => (prev < slides.length - 1 ? prev + 1 : prev));
  };

  const handleSelectSlide = (idx: number) => {
    setSelectedBlockId(null);
    setSelectedImageId(null);
    setNewlyAddedBlockId(null);
    setActiveSlideIdx(idx);
    setIsMobileDrawerOpen(false);
  };

  const handleToggleTimer = () => {
    setIsTimerRunning((prev) => !prev);
  };

  const handleResetTimer = () => {
    setIsTimerRunning(false);
    setElapsedSeconds(0);
  };

  const startElementDrag = (
    event: ReactPointerEvent<HTMLElement>,
    params: {
      kind: 'block' | 'image';
      id: string;
      xPct: number;
      yPct: number;
      wPct: number;
      hPct: number;
    }
  ) => {
    if (isLaserActive || isSlideshowActive || event.button !== 0) {
      return;
    }
    const canvasEl = editorCanvasRef.current;
    if (!canvasEl) {
      return;
    }
    event.stopPropagation();

    if (params.kind === 'block') {
      setSelectedBlockId(params.id);
      setSelectedImageId(null);
    } else {
      setSelectedImageId(params.id);
      setSelectedBlockId(null);
    }

    const dragState: ActiveDragState = {
      kind: params.kind,
      id: params.id,
      startClientX: event.clientX,
      startClientY: event.clientY,
      originXPct: params.xPct,
      originYPct: params.yPct,
      widthPct: params.wPct,
      heightPct: params.hPct,
      hasMoved: false,
    };

    const targetSlideIdx = safeSlideIdx;
    activeDragCleanupRef.current?.();

    const handleWindowPointerMove = (moveEvent: PointerEvent) => {
      const rect = editorCanvasRef.current?.getBoundingClientRect();
      if (!rect || rect.width <= 0 || rect.height <= 0) {
        return;
      }
      const dx = moveEvent.clientX - dragState.startClientX;
      const dy = moveEvent.clientY - dragState.startClientY;
      if (!dragState.hasMoved && Math.hypot(dx, dy) < 3) {
        return;
      }
      if (!dragState.hasMoved) {
        dragState.hasMoved = true;
        setDraggingId(dragState.id);
      }

      const maxX = Math.max(0, 100 - Math.min(dragState.widthPct, 95));
      const maxY = Math.max(0, 100 - Math.min(dragState.heightPct, 95));
      const nextX = roundPct(
        clampValue(
          dragState.originXPct + (dx / rect.width) * 100,
          0,
          maxX
        )
      );
      const nextY = roundPct(
        clampValue(
          dragState.originYPct + (dy / rect.height) * 100,
          0,
          maxY
        )
      );

      mutateSlideAt(targetSlideIdx, (slide) => {
        if (dragState.kind === 'block') {
          return {
            ...slide,
            blocks: (slide.blocks ?? []).map((block) =>
              block.id === dragState.id
                ? { ...block, xPct: nextX, yPct: nextY }
                : block
            ),
          };
        }
        return {
          ...slide,
          images: (slide.images ?? []).map((img) =>
            img.id === dragState.id ? { ...img, xPct: nextX, yPct: nextY } : img
          ),
        };
      });
    };

    const cleanupDrag = () => {
      window.removeEventListener('pointermove', handleWindowPointerMove);
      window.removeEventListener('pointerup', handleWindowPointerUp);
      window.removeEventListener('pointercancel', handleWindowPointerUp);
      activeDragCleanupRef.current = null;
      setDraggingId(null);
    };

    const handleWindowPointerUp = () => {
      if (dragState.hasMoved) {
        justDraggedRef.current = true;
        window.setTimeout(() => {
          justDraggedRef.current = false;
        }, 0);
      }
      cleanupDrag();
    };

    activeDragCleanupRef.current = cleanupDrag;
    window.addEventListener('pointermove', handleWindowPointerMove);
    window.addEventListener('pointerup', handleWindowPointerUp);
    window.addEventListener('pointercancel', handleWindowPointerUp);
  };

  const updateLaserFromPointer = (
    event: ReactPointerEvent<HTMLDivElement>
  ) => {
    if (!isLaserActive) {
      return;
    }
    const rect = event.currentTarget.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      return;
    }
    const x = clampValue(((event.clientX - rect.left) / rect.width) * 100, 0, 100);
    const y = clampValue(((event.clientY - rect.top) / rect.height) * 100, 0, 100);
    setLaserPos({ x, y, visible: true });
  };

  const handlePointerLeaveStage = () => {
    if (laserPos.visible) {
      setLaserPos((prev) => ({ ...prev, visible: false }));
    }
  };

  const updateSlideshowLaserFromPointer = (
    event: ReactPointerEvent<HTMLDivElement>
  ) => {
    if (!isSlideshowLaser) {
      return;
    }
    const rect = event.currentTarget.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      return;
    }
    const x = clampValue(((event.clientX - rect.left) / rect.width) * 100, 0, 100);
    const y = clampValue(((event.clientY - rect.top) / rect.height) * 100, 0, 100);
    setSlideshowLaserPos({ x, y, visible: true });
  };

  const renderSlideActionButtons = () => (
    <div className="shrink-0 p-2.5 border-b border-white/10 grid grid-cols-3 gap-1.5 bg-[#111620]">
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleAddSlide}
        title="Thêm slide mới"
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-indigo-600/90 hover:bg-indigo-500 text-[11px] font-semibold text-white transition-colors whitespace-nowrap"
      >
        <Plus className="h-3.5 w-3.5 shrink-0" />
        <span>Thêm</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDuplicateSlide}
        title="Nhân bản slide đang chọn"
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white/[0.07] hover:bg-white/[0.14] border border-white/10 text-[11px] font-medium text-slate-200 transition-colors whitespace-nowrap"
      >
        <Copy className="h-3 w-3 shrink-0" />
        <span>Nhân bản</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDeleteSlide}
        disabled={slides.length <= 1}
        title={
          slides.length <= 1
            ? 'Không thể xóa slide cuối cùng'
            : 'Xóa slide đang chọn'
        }
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white/[0.05] hover:bg-rose-500/20 border border-white/10 hover:border-rose-400/40 text-[11px] font-medium text-slate-300 hover:text-rose-200 transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed disabled:hover:bg-white/[0.05] disabled:hover:border-white/10 disabled:hover:text-slate-300"
      >
        <Trash2 className="h-3 w-3 shrink-0" />
        <span>Xóa</span>
      </button>
    </div>
  );

  const renderThumbnailList = (idPrefix: string) => (
    <div className="p-3 space-y-2.5">
      {slides.map((slide, idx) => {
        const isActive = idx === safeSlideIdx;
        const slideAccent = slide.accentColor || '#4F46E5';
        const isLightThumb = isLightHexBackground(slide.backgroundColor);
        const imageCount = slide.images?.length ?? 0;
        const thumbPositioned = isSlideMasterPositioned(slide);
        const thumbAspectRatio = getSafeAspectRatio(slide, document);

        return (
          <button
            key={`${idPrefix}-${slide.id}`}
            ref={
              idPrefix === 'desktop'
                ? (element) => {
                    thumbnailRefs.current[slide.id] = element;
                  }
                : undefined
            }
            type="button"
            onClick={() => handleSelectSlide(idx)}
            aria-current={isActive ? 'true' : undefined}
            aria-label={`Chọn Slide ${idx + 1}: ${slide.title}`}
            className={`group w-full text-left flex items-start gap-2 rounded-xl p-1.5 transition-all ${
              isActive
                ? 'bg-indigo-500/15 ring-2 ring-indigo-500'
                : 'hover:bg-white/[0.05] ring-1 ring-transparent'
            }`}
          >
            <span
              className={`w-5 shrink-0 pt-0.5 text-right text-[11px] font-semibold tabular-nums ${
                isActive
                  ? 'text-indigo-300'
                  : 'text-slate-500 group-hover:text-slate-300'
              }`}
            >
              {idx + 1}
            </span>

            <div className="flex-1 min-w-0">
              {thumbPositioned ? (
                <div
                  style={{ aspectRatio: String(thumbAspectRatio) }}
                  className={`w-full rounded-lg overflow-hidden border transition-colors pointer-events-none ${
                    isActive
                      ? 'border-indigo-400/80 shadow-md shadow-indigo-950/50'
                      : 'border-white/10 group-hover:border-white/25'
                  }`}
                >
                  <ReadOnlySlideStage
                    doc={document}
                    slide={slide}
                    slideIndex={idx}
                    totalSlides={slides.length}
                    stageWidthPx={164}
                    underlinedBlockIds={underlinedBlockIds}
                  />
                </div>
              ) : (
                <div
                  className={`aspect-video w-full rounded-lg overflow-hidden border p-2 flex flex-col justify-between transition-colors ${
                    isActive
                      ? 'border-indigo-400/80 shadow-md shadow-indigo-950/50'
                      : 'border-white/10 group-hover:border-white/25'
                  }`}
                  style={buildSlideBackgroundStyle(slide)}
                >
                  <div className="flex items-center justify-between gap-1">
                    <span
                      className="h-1 w-4 rounded-full"
                      style={{ backgroundColor: slideAccent }}
                    />
                    <div className="flex items-center gap-1">
                      {imageCount > 0 && (
                        <span
                          className={`inline-flex items-center gap-0.5 text-[8px] font-semibold px-1 py-0.5 rounded leading-none ${
                            isLightThumb
                              ? 'bg-slate-900/10 text-slate-700'
                              : 'bg-white/15 text-white/90'
                          }`}
                        >
                          <Image className="h-2.5 w-2.5" />
                          {imageCount}
                        </span>
                      )}
                      {slide.metricHighlight && (
                        <span className="text-[8px] font-bold text-white/90 px-1 py-0.5 rounded bg-white/10 tabular-nums leading-none">
                          {slide.metricHighlight.value}
                        </span>
                      )}
                    </div>
                  </div>

                  <div className="my-auto">
                    <p
                      className={`text-[9.5px] font-semibold leading-tight line-clamp-2 ${
                        isLightThumb ? 'text-slate-900' : 'text-white'
                      }`}
                    >
                      {slide.title}
                    </p>
                    <div className="mt-1.5 space-y-1">
                      <div
                        className={`h-1 w-4/5 rounded-full ${
                          isLightThumb ? 'bg-slate-400/40' : 'bg-white/20'
                        }`}
                      />
                      <div
                        className={`h-1 w-3/5 rounded-full ${
                          isLightThumb ? 'bg-slate-400/25' : 'bg-white/10'
                        }`}
                      />
                    </div>
                  </div>

                  <div className="h-0.5 w-full rounded-full bg-white/10 overflow-hidden">
                    <div
                      className="h-full rounded-full"
                      style={{
                        width: `${((idx + 1) / slides.length) * 100}%`,
                        backgroundColor: slideAccent,
                      }}
                    />
                  </div>
                </div>
              )}

              <p
                className={`mt-1.5 text-[11px] truncate ${
                  isActive ? 'text-white font-medium' : 'text-slate-400'
                }`}
              >
                {slide.title}
              </p>
            </div>
          </button>
        );
      })}
    </div>
  );

  const hasImages = currentImages.length > 0;
  const hasMediaColumn = hasImages || Boolean(currentSlide.metricHighlight);

  const slideshowAnimationStyle: CSSProperties =
    slideshowTransition === 'fade'
      ? { animation: 'liteSlideFadeIn 260ms cubic-bezier(0.22, 1, 0.36, 1) both' }
      : {
          animation: `${
            slideshowDirection > 0 ? 'liteSlideFromRight' : 'liteSlideFromLeft'
          } 280ms cubic-bezier(0.22, 1, 0.36, 1) both`,
        };

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#181D27] text-slate-100 overflow-hidden">
      {/* Hidden Image File Picker */}
      <input
        ref={imageInputRef}
        type="file"
        accept="image/*"
        onChange={(event) => void handleImageFileChange(event)}
        className="hidden"
        aria-label="Chọn ảnh để chèn vào slide"
      />

      {/* Shared Top Workspace Header */}
      <WorkspaceTopBar
        document={document}
        theme="dark"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument())}
        onRenameTitle={handleRenameTitle}
        leftExtra={
          <button
            type="button"
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở danh sách slide"
            title="Danh sách slide"
            className="md:hidden inline-flex items-center justify-center gap-1 min-h-[38px] px-2.5 rounded-xl bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 transition-colors"
          >
            <Layers className="h-4 w-4 text-indigo-300" />
            <span className="tabular-nums text-[11px]">
              {safeSlideIdx + 1}/{slides.length}
            </span>
          </button>
        }
        toolbar={
          <>
            {/* Presentation Stopwatch */}
            <div className="inline-flex items-center rounded-xl bg-[#1C2331] border border-white/10 p-0.5 shrink-0">
              <button
                type="button"
                onClick={handleToggleTimer}
                aria-label={
                  isTimerRunning
                    ? 'Tạm dừng đồng hồ trình chiếu'
                    : 'Bắt đầu bấm giờ trình chiếu'
                }
                title={
                  isTimerRunning
                    ? 'Bấm để tạm dừng đồng hồ'
                    : 'Bấm để chạy đồng hồ đếm giờ'
                }
                className="inline-flex items-center gap-1.5 min-h-[32px] px-2 rounded-[10px] hover:bg-white/[0.08] text-xs font-medium text-slate-100 transition-colors whitespace-nowrap"
              >
                <span
                  className={`h-2 w-2 rounded-full shrink-0 ${
                    isTimerRunning
                      ? 'bg-emerald-400 animate-pulse'
                      : elapsedSeconds > 0
                        ? 'bg-amber-400'
                        : 'bg-slate-500'
                  }`}
                />
                {isTimerRunning ? (
                  <Pause className="h-3.5 w-3.5 text-emerald-300 shrink-0" />
                ) : (
                  <Play className="h-3.5 w-3.5 text-slate-300 shrink-0" />
                )}
                <span className="font-mono text-[11.5px] tracking-tight tabular-nums">
                  {formatTimer(elapsedSeconds)}
                </span>
              </button>

              {elapsedSeconds > 0 && (
                <button
                  type="button"
                  onClick={handleResetTimer}
                  aria-label="Đặt lại đồng hồ về 00:00:00"
                  title="Đặt lại đồng hồ"
                  className="inline-flex items-center justify-center h-8 w-7 rounded-[10px] hover:bg-white/[0.1] text-slate-400 hover:text-white transition-colors"
                >
                  <RotateCcw className="h-3.5 w-3.5" />
                </button>
              )}
            </div>

            {/* Interactive Laser Pointer Toggle */}
            <button
              type="button"
              onClick={() => setIsLaserActive((prev) => !prev)}
              aria-pressed={isLaserActive}
              aria-label={
                isLaserActive ? 'Tắt bút Laser' : 'Bật bút Laser trên slide'
              }
              title={
                isLaserActive
                  ? 'Bút Laser đang bật (bấm để tắt)'
                  : 'Bật bút Laser chiếu điểm trên slide'
              }
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl text-xs font-medium border transition-all whitespace-nowrap ${
                isLaserActive
                  ? 'bg-rose-500/20 border-rose-400/60 text-rose-200 shadow-sm shadow-rose-950/50'
                  : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200 hover:text-white'
              }`}
            >
              <span className="relative flex items-center justify-center">
                <Crosshair className="h-3.5 w-3.5 shrink-0" />
                {isLaserActive && (
                  <span className="absolute -top-0.5 -right-0.5 h-2 w-2 rounded-full bg-rose-400 animate-ping" />
                )}
              </span>
              <span className="hidden xl:inline">Bút Laser</span>
            </button>

            {/* QR Remote Control */}
            <button
              type="button"
              onClick={() => onOpenModal('qr', buildLiveDocument())}
              aria-label="Mở điều khiển QR"
              title="Điều khiển trình chiếu qua mã QR"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 hover:text-white transition-colors whitespace-nowrap"
            >
              <QrCode className="h-3.5 w-3.5 text-indigo-300 shrink-0" />
              <span className="hidden 2xl:inline">QR Remote</span>
            </button>

            {/* Presenter View Modal Shortcut */}
            <button
              type="button"
              onClick={() => onOpenModal('present', buildLiveDocument())}
              aria-label="Mở Presenter View"
              title="Presenter View (Ghi chú diễn giả & Điều khiển từ xa)"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 hover:text-white transition-colors whitespace-nowrap"
            >
              <Columns2 className="h-3.5 w-3.5 text-indigo-300 shrink-0" />
              <span className="hidden 2xl:inline">Presenter View</span>
            </button>

            {/* Save Changes Button */}
            <button
              type="button"
              onClick={handleSave}
              aria-label="Lưu thay đổi bản trình chiếu"
              title="Lưu thay đổi (Ctrl+S)"
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
                saveFeedback
                  ? 'bg-emerald-500/20 border-emerald-400/50 text-emerald-200'
                  : isDirty
                    ? 'bg-amber-500/20 hover:bg-amber-500/30 border-amber-400/50 text-amber-100'
                    : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
              }`}
            >
              {saveFeedback ? (
                <>
                  <Check className="h-3.5 w-3.5 shrink-0" />
                  <span className="hidden md:inline">Đã lưu</span>
                </>
              ) : (
                <>
                  <span className="relative flex items-center justify-center">
                    <Save className="h-3.5 w-3.5 shrink-0" />
                    {isDirty && (
                      <span className="absolute -top-1 -right-1 h-2 w-2 rounded-full bg-amber-400" />
                    )}
                  </span>
                  <span className="hidden md:inline">Lưu thay đổi</span>
                </>
              )}
            </button>

            {/* Export Real .pptx Button */}
            <button
              type="button"
              onClick={() => void handleExportPptx()}
              disabled={isExporting}
              aria-label="Tải xuống tệp PowerPoint .pptx"
              title="Tải xuống .pptx"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl bg-orange-500/15 hover:bg-orange-500/25 border border-orange-400/35 text-xs font-semibold text-orange-200 transition-colors whitespace-nowrap disabled:opacity-60 disabled:cursor-wait"
            >
              {isExporting ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
              ) : (
                <Download className="h-3.5 w-3.5 shrink-0" />
              )}
              <span className="hidden lg:inline">
                {isExporting ? 'Đang xuất...' : 'Tải xuống .pptx'}
              </span>
            </button>

            {/* Primary Indigo CTA: Real Fullscreen Slideshow */}
            <button
              type="button"
              onClick={startSlideshow}
              aria-label="Trình chiếu toàn màn hình"
              title="Trình chiếu toàn màn hình (F5 / Toàn màn hình)"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white shadow-sm shadow-indigo-950/60 transition-all whitespace-nowrap"
            >
              <Play className="h-3.5 w-3.5 fill-white text-white shrink-0" />
              <span>Trình chiếu</span>
            </button>
          </>
        }
      />

      {/* Cloud Hydration Status Banner */}
      {isHydratingCloudFile && (
        <div className="shrink-0 px-4 py-2 bg-indigo-500/15 border-b border-indigo-400/30 text-xs text-indigo-200 flex items-center justify-center gap-2">
          <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0 text-indigo-300" />
          <span>Đang tải bản trình chiếu gốc từ đám mây...</span>
        </div>
      )}

      {/* Missing Cloud Buffer Retry Notice (only when no preview slides exist) */}
      {!isHydratingCloudFile &&
        document.slidesPreview.length === 0 &&
        Boolean(document.storagePath || document.r2ObjectKey) &&
        !document.rawArrayBuffer &&
        onRetryCloudDownload && (
          <div className="shrink-0 px-4 py-2 bg-amber-500/15 border-b border-amber-400/30 text-xs text-amber-200 flex items-center justify-between gap-2">
            <span className="inline-flex items-center gap-1.5">
              <CloudDownload className="h-3.5 w-3.5 shrink-0" />
              <span>Chưa tải được dữ liệu tệp trình chiếu gốc từ đám mây.</span>
            </span>
            <button
              type="button"
              onClick={onRetryCloudDownload}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-500/25 hover:bg-amber-500/40 text-xs font-semibold text-white transition-colors"
            >
              <RefreshCw className="h-3 w-3" />
              <span>Thử tải lại</span>
            </button>
          </div>
        )}

      {/* Export Error Notice */}
      {exportError && (
        <div className="shrink-0 px-4 py-2 bg-rose-500/15 border-b border-rose-400/30 text-xs text-rose-200 flex items-center justify-between gap-2">
          <span>{exportError}</span>
          <button
            type="button"
            onClick={() => setExportError(null)}
            aria-label="Đóng thông báo lỗi"
            className="p-1 rounded hover:bg-rose-500/20"
          >
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      )}

      {/* Main Split Workspace Body */}
      <div className="flex-1 min-h-0 flex relative overflow-hidden">
        {/* Mobile Slide-over Drawer Backdrop (< md) */}
        <div
          aria-hidden="true"
          onClick={() => setIsMobileDrawerOpen(false)}
          className={`md:hidden fixed inset-0 z-30 bg-slate-950/60 backdrop-blur-[2px] transition-opacity duration-200 ${
            isMobileDrawerOpen
              ? 'opacity-100 pointer-events-auto'
              : 'opacity-0 pointer-events-none'
          }`}
        />

        {/* Mobile Slide-over Drawer (< md, z-40) */}
        <aside
          aria-label="Danh sách slide di động"
          className={`md:hidden fixed inset-y-0 left-0 z-40 w-[268px] max-w-[84vw] bg-[#141923] border-r border-white/10 shadow-2xl flex flex-col transition-transform duration-200 ease-out ${
            isMobileDrawerOpen
              ? 'translate-x-0 pointer-events-auto'
              : '-translate-x-full pointer-events-none'
          }`}
        >
          <div className="h-14 shrink-0 px-4 border-b border-white/10 flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 min-w-0">
              <Layers className="h-4 w-4 text-indigo-400 shrink-0" />
              <span className="text-xs font-semibold text-white truncate">
                Danh sách slide
              </span>
              <span className="px-1.5 py-0.5 rounded-md bg-white/10 text-[10px] font-semibold text-slate-300 tabular-nums">
                {slides.length}
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsMobileDrawerOpen(false)}
              aria-label="Đóng danh sách slide"
              className="inline-flex items-center justify-center min-h-[38px] min-w-[38px] rounded-xl bg-white/[0.06] hover:bg-white/[0.14] text-slate-300 hover:text-white transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
          {renderSlideActionButtons()}
          <div className="flex-1 min-h-0 overflow-y-auto scrollbar-dark pb-safe">
            {renderThumbnailList('mobile')}
          </div>
        </aside>

        {/* Desktop Left Thumbnail Rail (md+, 220px) */}
        <aside
          aria-label="Danh sách slide thu nhỏ"
          className="hidden md:flex w-[220px] shrink-0 bg-[#141923] border-r border-white/10 flex-col overflow-hidden"
        >
          <div className="shrink-0 px-3.5 py-2.5 border-b border-white/10 flex items-center justify-between gap-2">
            <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
              Trang trình chiếu
            </span>
            <span className="px-1.5 py-0.5 rounded-md bg-white/[0.08] text-[10.5px] font-semibold text-slate-200 tabular-nums">
              {slides.length} slide
            </span>
          </div>
          {renderSlideActionButtons()}
          <div className="relative flex-1 min-h-0 flex overflow-hidden">
            <div
              ref={desktopThumbScrollRef}
              className="flex-1 min-h-0 overflow-y-auto scrollbar-none pr-4"
            >
              {renderThumbnailList('desktop')}
            </div>
            <CustomVerticalScrollbar
              scrollContainerRef={desktopThumbScrollRef}
              theme="dark"
              getDragLabel={({ percent }) => {
                const approxSlide = Math.max(
                  1,
                  Math.min(
                    slides.length,
                    Math.round(
                      (percent / 100) * Math.max(1, slides.length - 1)
                    ) + 1
                  )
                );
                return `Slide ${approxSlide}/${slides.length} · ${Math.round(percent)}%`;
              }}
            />
          </div>
        </aside>

        {/* Center Stage Editor + Bottom Speaker Notes */}
        <main className="flex-1 min-w-0 flex flex-col bg-[#181D27] overflow-hidden relative">
          {/* Desktop Slide Text Block, Image & Background Editing Toolbar (md+) */}
          <div className="hidden md:block shrink-0 bg-[#141923]/95 border-b border-white/10 px-3 sm:px-5 py-1.5">
            <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-h-[34px]">
              {/* + Thêm khung chữ */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleAddTextBlock}
                className="shrink-0 inline-flex items-center gap-1.5 min-h-[32px] px-2.5 rounded-lg bg-indigo-500/20 hover:bg-indigo-500/30 border border-indigo-400/35 text-xs font-semibold text-indigo-200 transition-colors whitespace-nowrap"
              >
                <Plus className="h-3.5 w-3.5 shrink-0" />
                <Type className="h-3.5 w-3.5 shrink-0" />
                <span>Thêm khung chữ</span>
              </button>

              {/* + Thêm ảnh */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleTriggerAddImage}
                className="shrink-0 inline-flex items-center gap-1.5 min-h-[32px] px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 transition-colors whitespace-nowrap"
              >
                <ImagePlus className="h-3.5 w-3.5 text-emerald-300 shrink-0" />
                <span>Thêm ảnh</span>
              </button>

              {/* Slide Background Swatches */}
              <div
                className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0"
                title="Đổi màu nền slide hiện tại"
              >
                <PaintBucket className="h-3 w-3 text-slate-400 mr-0.5 shrink-0" />
                {SLIDE_BG_PRESETS.map((bg) => {
                  const active =
                    !currentSlide.backgroundImageSrc &&
                    !currentSlide.backgroundGradient &&
                    (currentSlide.backgroundColor ?? '#161D2E').toLowerCase() ===
                      bg.value.toLowerCase();
                  return (
                    <button
                      key={bg.value}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleChangeSlideBackground(bg.value)}
                      title={`Nền ${bg.label}`}
                      aria-label={`Đổi nền slide thành ${bg.label}`}
                      className={`h-4 w-4 rounded border transition-transform ${
                        active
                          ? 'ring-2 ring-indigo-400 scale-110 border-white'
                          : 'border-white/25 hover:scale-110'
                      }`}
                      style={{ backgroundColor: bg.value }}
                    />
                  );
                })}
                <input
                  type="color"
                  value={toSafeColorInputValue(
                    currentSlide.backgroundColor,
                    '#161D2E'
                  )}
                  onChange={(event) =>
                    handleChangeSlideBackground(event.target.value)
                  }
                  aria-label="Chọn màu nền slide tùy chỉnh"
                  title="Màu nền slide tùy chỉnh"
                  className="h-4 w-5 ml-0.5 rounded cursor-pointer border-0 bg-transparent p-0"
                />
              </div>

              {selectedBlock ? (
                <>
                  <span
                    aria-hidden="true"
                    className="h-5 w-px mx-1 bg-white/10 shrink-0"
                  />

                  {/* H1 / H2 / P Preset Buttons */}
                  <div
                    role="group"
                    aria-label="Kiểu tiêu đề khung chữ"
                    className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-[#1C2331] border border-white/10 shrink-0"
                  >
                    {(
                      [
                        { tag: 'h1', label: 'H1', active: selectedBlock.fontSize >= 32 },
                        {
                          tag: 'h2',
                          label: 'H2',
                          active:
                            selectedBlock.fontSize >= 24 &&
                            selectedBlock.fontSize < 32,
                        },
                        { tag: 'p', label: 'P', active: selectedBlock.fontSize < 24 },
                      ] as const
                    ).map(({ tag, label, active }) => (
                      <button
                        key={tag}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() =>
                          handleApplyBlockPreset(selectedBlock.id, tag)
                        }
                        aria-pressed={active}
                        title={`Kiểu chữ ${label}`}
                        className={`min-h-[26px] px-2 rounded-md text-[11px] font-bold transition-colors ${
                          active
                            ? 'bg-indigo-600 text-white'
                            : 'text-slate-300 hover:bg-white/10 hover:text-white'
                        }`}
                      >
                        {label}
                      </button>
                    ))}
                  </div>

                  {/* Font Size Picker */}
                  <div
                    role="group"
                    aria-label="Cỡ chữ khung đang chọn"
                    className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-[#1C2331] border border-white/10 shrink-0"
                  >
                    {SLIDE_FONT_SIZES.map((size) => {
                      const isSizeActive = selectedBlock.fontSize === size;
                      return (
                        <button
                          key={size}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            handleUpdateBlock(selectedBlock.id, {
                              fontSize: size,
                            })
                          }
                          aria-pressed={isSizeActive}
                          className={`min-h-[26px] px-2 rounded-md text-[11px] font-semibold tabular-nums transition-colors ${
                            isSizeActive
                              ? 'bg-indigo-600 text-white'
                              : 'text-slate-300 hover:bg-white/10 hover:text-white'
                          }`}
                        >
                          {size}
                        </button>
                      );
                    })}
                  </div>

                  {/* Bold, Italic & Underline */}
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() =>
                      handleUpdateBlock(selectedBlock.id, {
                        bold: !selectedBlock.bold,
                      })
                    }
                    aria-pressed={Boolean(selectedBlock.bold)}
                    aria-label="In đậm khung chữ"
                    title="In đậm"
                    className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors ${
                      selectedBlock.bold
                        ? 'bg-indigo-500/30 border-indigo-400/50 text-white'
                        : 'bg-white/[0.04] hover:bg-white/[0.1] border-white/10 text-slate-300'
                    }`}
                  >
                    <Bold className="h-3.5 w-3.5" />
                  </button>

                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() =>
                      handleUpdateBlock(selectedBlock.id, {
                        italic: !selectedBlock.italic,
                      })
                    }
                    aria-pressed={Boolean(selectedBlock.italic)}
                    aria-label="In nghiêng khung chữ"
                    title="In nghiêng"
                    className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors ${
                      selectedBlock.italic
                        ? 'bg-indigo-500/30 border-indigo-400/50 text-white'
                        : 'bg-white/[0.04] hover:bg-white/[0.1] border-white/10 text-slate-300'
                    }`}
                  >
                    <Italic className="h-3.5 w-3.5" />
                  </button>

                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() =>
                      handleToggleBlockUnderline(selectedBlock.id)
                    }
                    aria-pressed={Boolean(underlinedBlockIds[selectedBlock.id])}
                    aria-label="Gạch chân khung chữ"
                    title="Gạch chân"
                    className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors ${
                      underlinedBlockIds[selectedBlock.id]
                        ? 'bg-indigo-500/30 border-indigo-400/50 text-white'
                        : 'bg-white/[0.04] hover:bg-white/[0.1] border-white/10 text-slate-300'
                    }`}
                  >
                    <Underline className="h-3.5 w-3.5" />
                  </button>

                  {/* Text Color Swatches + Custom Color Input */}
                  <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                    {SLIDE_COLOR_SWATCHES.map((swatch) => {
                      const isColorActive =
                        (selectedBlock.color ?? '#FFFFFF').toLowerCase() ===
                        swatch.value.toLowerCase();
                      return (
                        <button
                          key={swatch.value}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            handleUpdateBlock(selectedBlock.id, {
                              color: swatch.value,
                            })
                          }
                          title={`Màu ${swatch.label}`}
                          aria-label={`Đổi màu chữ thành ${swatch.label}`}
                          className={`h-5 w-5 rounded-full border transition-transform ${
                            isColorActive
                              ? 'ring-2 ring-indigo-400 scale-110 border-white'
                              : 'border-white/25 hover:scale-110'
                          }`}
                          style={{ backgroundColor: swatch.value }}
                        />
                      );
                    })}
                    <input
                      type="color"
                      value={toSafeColorInputValue(
                        selectedBlock.color,
                        hasLightSlideBg ? '#0F172A' : '#FFFFFF'
                      )}
                      onChange={(event) =>
                        handleUpdateBlock(selectedBlock.id, {
                          color: event.target.value,
                        })
                      }
                      aria-label="Chọn màu chữ tùy chỉnh"
                      title="Màu chữ tùy chỉnh"
                      className="h-5 w-6 ml-0.5 rounded cursor-pointer border-0 bg-transparent p-0"
                    />
                  </div>

                  {/* Block Fill Color Picker */}
                  <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                    <span className="text-[10px] text-slate-400 px-0.5">
                      Nền khối
                    </span>
                    <input
                      type="color"
                      value={toSafeColorInputValue(
                        selectedBlock.fillColor,
                        '#1E293B'
                      )}
                      onChange={(event) =>
                        handleUpdateBlock(selectedBlock.id, {
                          fillColor: event.target.value,
                        })
                      }
                      aria-label="Màu nền khung chữ"
                      title="Màu nền khung chữ"
                      className="h-5 w-5 rounded cursor-pointer border-0 bg-transparent p-0"
                    />
                    {selectedBlock.fillColor && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() =>
                          handleUpdateBlock(selectedBlock.id, {
                            fillColor: undefined,
                          })
                        }
                        title="Xóa màu nền khối"
                        className="text-[10px] text-slate-300 hover:text-white px-1 rounded bg-white/10"
                      >
                        Trong
                      </button>
                    )}
                  </div>

                  {/* Alignment: Left / Center / Right */}
                  <div className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                    {(
                      [
                        { value: 'left', label: 'Căn trái', Icon: AlignLeft },
                        {
                          value: 'center',
                          label: 'Căn giữa',
                          Icon: AlignCenter,
                        },
                        { value: 'right', label: 'Căn phải', Icon: AlignRight },
                      ] as const
                    ).map(({ value, label, Icon }) => {
                      const active = (selectedBlock.align ?? 'left') === value;
                      return (
                        <button
                          key={value}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            handleUpdateBlock(selectedBlock.id, {
                              align: value,
                            })
                          }
                          aria-pressed={active}
                          aria-label={label}
                          title={label}
                          className={`h-7 w-7 rounded-md flex items-center justify-center transition-colors ${
                            active
                              ? 'bg-indigo-600 text-white'
                              : 'text-slate-300 hover:bg-white/10'
                          }`}
                        >
                          <Icon className="h-3.5 w-3.5" />
                        </button>
                      );
                    })}
                  </div>

                  {/* Delete Selected Block */}
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => handleDeleteBlock(selectedBlock.id)}
                    aria-label="Xóa khung chữ đang chọn"
                    title="Xóa khung chữ"
                    className="shrink-0 inline-flex items-center gap-1 min-h-[32px] px-2.5 rounded-lg bg-rose-500/15 hover:bg-rose-500/25 border border-rose-400/35 text-xs font-medium text-rose-200 transition-colors whitespace-nowrap"
                  >
                    <Trash2 className="h-3.5 w-3.5 shrink-0" />
                    <span className="hidden lg:inline">Xóa khung</span>
                  </button>
                </>
              ) : selectedImage ? (
                <>
                  <span
                    aria-hidden="true"
                    className="h-5 w-px mx-1 bg-white/10 shrink-0"
                  />
                  <span className="text-[11.5px] text-indigo-200 whitespace-nowrap">
                    Đang chọn ảnh{isPositionedSlide ? ' · Kéo thả để di chuyển' : ''}
                  </span>
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => handleDeleteImage(selectedImage.id)}
                    aria-label="Xóa ảnh đang chọn"
                    className="shrink-0 inline-flex items-center gap-1 min-h-[32px] px-2.5 rounded-lg bg-rose-500/15 hover:bg-rose-500/25 border border-rose-400/35 text-xs font-medium text-rose-200 transition-colors whitespace-nowrap"
                  >
                    <Trash2 className="h-3.5 w-3.5 shrink-0" />
                    <span>Xóa ảnh</span>
                  </button>
                </>
              ) : (
                <p className="text-[11.5px] text-slate-400 truncate pl-1.5">
                  Bấm vào khung chữ/bảng để sửa nội dung, hoặc kéo thanh tay cầm
                  để di chuyển vị trí.
                </p>
              )}
            </div>
          </div>

          {/* Active Laser Mode Banner */}
          {isLaserActive && (
            <div className="shrink-0 px-3 sm:px-6 pt-2 flex justify-center">
              <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-rose-500/15 border border-rose-400/35 text-[11.5px] text-rose-200">
                <span className="h-2 w-2 rounded-full bg-rose-500 shadow-[0_0_8px_2px_rgba(244,63,94,0.8)] animate-pulse shrink-0" />
                <span className="truncate">
                  Bút Laser đang bật — Di chuyển hoặc chạm trên slide để chiếu
                  điểm
                </span>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setIsLaserActive(false)}
                  className="ml-1 px-2 py-0.5 rounded-full bg-rose-500/25 hover:bg-rose-500/40 text-[11px] font-semibold text-white transition-colors shrink-0"
                >
                  Tắt
                </button>
              </div>
            </div>
          )}

          {/* Main Editable Slide Stage Container (ResizeObserver proportional fit) */}
          <div className="flex-1 min-h-0 flex flex-col px-3 sm:px-6 lg:px-10 pt-3 sm:pt-4 pb-2 overflow-hidden">
            <div
              ref={editorViewportRef}
              className="flex-1 min-h-0 w-full flex items-center justify-center relative overflow-hidden"
            >
              <div
                ref={editorCanvasRef}
                onClick={() => {
                  if (justDraggedRef.current) {
                    return;
                  }
                  if (!isLaserActive) {
                    setSelectedBlockId(null);
                    setSelectedImageId(null);
                  }
                }}
                onPointerMove={updateLaserFromPointer}
                onPointerDown={updateLaserFromPointer}
                onPointerLeave={handlePointerLeaveStage}
                style={{
                  width: `${Math.round(editorFitWidthPx)}px`,
                  height: `${Math.round(editorFitHeightPx)}px`,
                  ...buildSlideBackgroundStyle(currentSlide),
                }}
                className={`relative rounded-2xl overflow-hidden border shadow-2xl shadow-black/55 transition-shadow duration-200 ${
                  isLaserActive
                    ? 'border-rose-400/55 ring-2 ring-rose-500/25 cursor-none touch-none select-none'
                    : 'border-white/15 ring-1 ring-white/[0.06]'
                }`}
              >
                {isPositionedSlide ? (
                  /* Master-Positioned OpenXML Coordinate Stage (xPct, yPct, wPct, hPct) */
                  <div className="absolute inset-0">
                    {currentBlocks.map((block, bIdx) => {
                      const isSelected =
                        !isLaserActive && selectedBlockId === block.id;
                      const isBeingDragged = draggingId === block.id;
                      const xPct = block.xPct ?? 6;
                      const yPct = block.yPct ?? 10 + bIdx * 14;
                      const wPct = block.wPct ?? 88;
                      const hPct = block.hPct ?? 16;
                      const scaledFontPx = Math.max(
                        8,
                        (block.fontSize ||
                          (block.kind === 'table' ? 12 : 18)) *
                          editorFontScale
                      );
                      const justifyClass =
                        block.valign === 'middle'
                          ? 'justify-center'
                          : block.valign === 'bottom'
                            ? 'justify-end'
                            : 'justify-start';
                      const textColor =
                        block.color ?? (hasLightSlideBg ? '#0F172A' : '#FFFFFF');

                      return (
                        <div
                          key={block.id}
                          onClick={(event) => {
                            if (isLaserActive) {
                              return;
                            }
                            event.stopPropagation();
                            setSelectedBlockId(block.id);
                            setSelectedImageId(null);
                          }}
                          onPointerDown={(event) => {
                            if (event.target === event.currentTarget) {
                              startElementDrag(event, {
                                kind: 'block',
                                id: block.id,
                                xPct,
                                yPct,
                                wPct,
                                hPct,
                              });
                            }
                          }}
                          style={{
                            position: 'absolute',
                            left: `${xPct}%`,
                            top: `${yPct}%`,
                            width: `${wPct}%`,
                            height:
                              block.hPct !== undefined ? `${hPct}%` : 'auto',
                            backgroundColor: block.fillColor,
                            borderColor: block.borderColor,
                            fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                          }}
                          className={`group flex flex-col ${justifyClass} rounded-lg px-1.5 py-1 transition-shadow ${
                            block.borderColor ? 'border' : ''
                          } ${
                            isSelected
                              ? 'ring-2 ring-indigo-500 z-10'
                              : isLaserActive
                                ? ''
                                : hasLightSlideBg
                                  ? 'hover:ring-1 hover:ring-slate-400/50'
                                  : 'hover:ring-1 hover:ring-white/30'
                          } ${isBeingDragged ? 'opacity-90 shadow-xl' : ''}`}
                        >
                          {/* Floating Drag Handle */}
                          {!isLaserActive && (
                            <button
                              type="button"
                              onPointerDown={(event) =>
                                startElementDrag(event, {
                                  kind: 'block',
                                  id: block.id,
                                  xPct,
                                  yPct,
                                  wPct,
                                  hPct,
                                })
                              }
                              aria-label="Kéo để di chuyển khung"
                              title="Kéo thả để di chuyển vị trí"
                              className={`absolute -top-3 left-1 z-20 inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-md bg-indigo-600 text-white text-[9px] font-semibold shadow-md cursor-grab active:cursor-grabbing transition-opacity ${
                                isSelected
                                  ? 'opacity-100'
                                  : 'opacity-0 group-hover:opacity-100'
                              }`}
                            >
                              <Move className="h-2.5 w-2.5" />
                              <span>Kéo</span>
                            </button>
                          )}

                          {block.kind === 'table' &&
                          Array.isArray(block.tableRows) &&
                          block.tableRows.length > 0 ? (
                            <div className="w-full h-full overflow-auto no-scrollbar">
                              <table className="w-full h-full border-collapse table-fixed">
                                <tbody>
                                  {block.tableRows.map((row, rIdx) => (
                                    <tr key={`${block.id}-r-${rIdx}`}>
                                      {row.map((cell, cIdx) => (
                                        <td
                                          key={`${block.id}-c-${rIdx}-${cIdx}`}
                                          className="border px-1.5 py-0.5 align-middle"
                                          style={{
                                            borderColor:
                                              block.borderColor ??
                                              (hasLightSlideBg
                                                ? 'rgba(15,23,42,0.22)'
                                                : 'rgba(255,255,255,0.22)'),
                                            backgroundColor:
                                              rIdx === 0
                                                ? block.fillColor ??
                                                  (hasLightSlideBg
                                                    ? 'rgba(15,23,42,0.06)'
                                                    : 'rgba(255,255,255,0.08)')
                                                : 'transparent',
                                          }}
                                        >
                                          <input
                                            type="text"
                                            value={cell}
                                            disabled={isLaserActive}
                                            onFocus={() => {
                                              if (!isLaserActive) {
                                                setSelectedBlockId(block.id);
                                                setSelectedImageId(null);
                                              }
                                            }}
                                            onPointerDown={(event) => {
                                              if (!isLaserActive) {
                                                event.stopPropagation();
                                              }
                                            }}
                                            onChange={(event) =>
                                              handleUpdateTableCell(
                                                block.id,
                                                rIdx,
                                                cIdx,
                                                event.target.value
                                              )
                                            }
                                            aria-label={`Ô bảng hàng ${rIdx + 1} cột ${cIdx + 1}`}
                                            style={{
                                              fontSize: `${scaledFontPx.toFixed(2)}px`,
                                              fontWeight:
                                                rIdx === 0 || block.bold
                                                  ? 700
                                                  : 400,
                                              fontStyle: block.italic
                                                ? 'italic'
                                                : 'normal',
                                              textDecoration:
                                                underlinedBlockIds[block.id]
                                                  ? 'underline'
                                                  : 'none',
                                              color: textColor,
                                              textAlign: block.align ?? 'left',
                                              fontFamily:
                                                block.fontFamily || undefined,
                                            }}
                                            className="w-full min-w-0 bg-transparent border-0 outline-none truncate"
                                          />
                                        </td>
                                      ))}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : (
                            <AutoGrowTextarea
                              value={block.text}
                              onChange={(nextText) =>
                                handleUpdateBlock(block.id, {
                                  text: nextText,
                                })
                              }
                              onFocus={() => {
                                if (!isLaserActive) {
                                  setSelectedBlockId(block.id);
                                  setSelectedImageId(null);
                                }
                              }}
                              disabled={isLaserActive}
                              autoFocus={newlyAddedBlockId === block.id}
                              placeholder={
                                block.role === 'title'
                                  ? 'Nhập tiêu đề slide...'
                                  : 'Nhập nội dung khung chữ...'
                              }
                              ariaLabel={`Khung chữ ${block.role ?? 'nội dung'} trên Slide ${safeSlideIdx + 1}`}
                              scaledFontPx={scaledFontPx}
                              bold={Boolean(block.bold)}
                              italic={Boolean(block.italic)}
                              underline={Boolean(underlinedBlockIds[block.id])}
                              color={textColor}
                              align={block.align ?? 'left'}
                              fontFamily={block.fontFamily}
                              maxHeightFull
                            />
                          )}
                        </div>
                      );
                    })}

                    {/* Positioned Draggable Images */}
                    {currentImages.map((imageItem, imgIdx) => {
                      const isSelected =
                        !isLaserActive && selectedImageId === imageItem.id;
                      const xPct = imageItem.xPct ?? 56;
                      const yPct = imageItem.yPct ?? 20 + imgIdx * 8;
                      const wPct = imageItem.wPct ?? 36;
                      const hPct = imageItem.hPct ?? 44;

                      return (
                        <div
                          key={imageItem.id}
                          onClick={(event) => {
                            if (isLaserActive) {
                              return;
                            }
                            event.stopPropagation();
                            setSelectedImageId(imageItem.id);
                            setSelectedBlockId(null);
                          }}
                          onPointerDown={(event) =>
                            startElementDrag(event, {
                              kind: 'image',
                              id: imageItem.id,
                              xPct,
                              yPct,
                              wPct,
                              hPct,
                            })
                          }
                          style={{
                            position: 'absolute',
                            left: `${xPct}%`,
                            top: `${yPct}%`,
                            width: `${wPct}%`,
                            height: `${hPct}%`,
                          }}
                          className={`group rounded-lg overflow-hidden flex items-center justify-center transition-shadow ${
                            isLaserActive
                              ? 'pointer-events-none'
                              : 'cursor-grab active:cursor-grabbing'
                          } ${
                            isSelected
                              ? 'ring-2 ring-indigo-500 z-10'
                              : 'hover:ring-1 hover:ring-indigo-400/60'
                          }`}
                        >
                          <img
                            src={imageItem.src}
                            alt={
                              imageItem.alt ??
                              `Hình ảnh ${imgIdx + 1} của Slide ${safeSlideIdx + 1}`
                            }
                            draggable={false}
                            className="w-full h-full object-contain pointer-events-none select-none"
                          />
                        </div>
                      );
                    })}

                    {currentBlocks.length === 0 && currentImages.length === 0 && (
                      <div className="w-full h-full flex flex-col items-center justify-center text-center opacity-75 p-6">
                        <Type className="h-7 w-7 mb-2 text-indigo-300" />
                        <p className="text-xs sm:text-sm font-medium">
                          Slide đang trống — bấm <strong>+ Thêm khung chữ</strong>{' '}
                          hoặc <strong>Thêm ảnh</strong> phía trên.
                        </p>
                      </div>
                    )}
                  </div>
                ) : (
                  /* Legacy Proportional Flow Canvas (for mock / fallback slides without coordinates) */
                  <div
                    className="absolute inset-0 flex flex-col justify-between"
                    style={{
                      padding: `${Math.max(12, Math.round(32 * editorFontScale))}px`,
                    }}
                  >
                    {/* Slide Top Header Strip */}
                    <div className="shrink-0 flex items-center justify-between gap-3 mb-2">
                      <div className="inline-flex items-center gap-2 min-w-0">
                        <span
                          className="h-1.5 sm:h-2 w-6 sm:w-8 rounded-full shrink-0"
                          style={{ backgroundColor: accentColor }}
                        />
                        <span
                          className={`text-[10px] sm:text-xs font-semibold uppercase tracking-widest truncate ${
                            hasLightSlideBg
                              ? 'text-slate-600'
                              : 'text-slate-300/90'
                          }`}
                        >
                          {document.folder ?? 'Bản trình chiếu'}
                        </span>
                      </div>

                      <span
                        className={`shrink-0 px-2.5 py-0.5 rounded-full border text-[10px] sm:text-xs font-semibold tabular-nums backdrop-blur-sm ${
                          hasLightSlideBg
                            ? 'bg-slate-900/10 border-slate-900/10 text-slate-800'
                            : 'bg-black/45 border-white/10 text-white/95'
                        }`}
                      >
                        Slide {safeSlideIdx + 1} / {slides.length}
                      </span>
                    </div>

                    {/* Main Editable Blocks + Images / Metric Area */}
                    <div
                      className={`flex-1 min-h-0 overflow-y-auto scrollbar-dark pr-1 grid gap-3 sm:gap-5 items-center ${
                        hasMediaColumn && currentBlocks.length > 0
                          ? 'grid-cols-12'
                          : 'grid-cols-1'
                      }`}
                    >
                      {currentBlocks.length > 0 && (
                        <div
                          className={`space-y-2 sm:space-y-3 ${
                            hasMediaColumn
                              ? 'col-span-7 sm:col-span-8'
                              : 'col-span-1'
                          }`}
                        >
                          {currentBlocks.map((block) => {
                            const isSelected =
                              !isLaserActive && selectedBlockId === block.id;
                            const scaledFontPx = Math.max(
                              10,
                              (block.fontSize || 18) * editorFontScale
                            );
                            return (
                              <div
                                key={block.id}
                                onClick={(event) => {
                                  if (isLaserActive) {
                                    return;
                                  }
                                  event.stopPropagation();
                                  setSelectedBlockId(block.id);
                                  setSelectedImageId(null);
                                }}
                                className={`rounded-xl px-2.5 py-1.5 -mx-2.5 transition-all ${
                                  isSelected
                                    ? hasLightSlideBg
                                      ? 'ring-2 ring-indigo-500 bg-indigo-500/[0.06]'
                                      : 'ring-2 ring-indigo-400/90 bg-white/[0.06]'
                                    : isLaserActive
                                      ? ''
                                      : hasLightSlideBg
                                        ? 'hover:ring-1 hover:ring-slate-400/40 hover:bg-black/[0.02] cursor-text'
                                        : 'hover:ring-1 hover:ring-white/25 hover:bg-white/[0.03] cursor-text'
                                }`}
                              >
                                <AutoGrowTextarea
                                  value={block.text}
                                  onChange={(nextText) =>
                                    handleUpdateBlock(block.id, {
                                      text: nextText,
                                    })
                                  }
                                  onFocus={() => {
                                    if (!isLaserActive) {
                                      setSelectedBlockId(block.id);
                                      setSelectedImageId(null);
                                    }
                                  }}
                                  disabled={isLaserActive}
                                  autoFocus={newlyAddedBlockId === block.id}
                                  placeholder={
                                    block.role === 'title'
                                      ? 'Nhập tiêu đề slide...'
                                      : 'Nhập nội dung khung chữ...'
                                  }
                                  ariaLabel={`Khung chữ ${block.role ?? 'nội dung'} trên Slide ${safeSlideIdx + 1}`}
                                  scaledFontPx={scaledFontPx}
                                  bold={Boolean(block.bold)}
                                  italic={Boolean(block.italic)}
                                  underline={Boolean(
                                    underlinedBlockIds[block.id]
                                  )}
                                  color={
                                    block.color ??
                                    (hasLightSlideBg ? '#0F172A' : '#FFFFFF')
                                  }
                                  align={block.align ?? 'left'}
                                  fontFamily={block.fontFamily}
                                />
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {hasImages && currentImages && (
                        <div
                          className={`${
                            currentBlocks.length > 0
                              ? 'col-span-5 sm:col-span-4'
                              : 'col-span-1 max-w-2xl mx-auto w-full'
                          } grid gap-2 ${
                            currentImages.length > 1
                              ? 'grid-cols-2'
                              : 'grid-cols-1'
                          }`}
                        >
                          {currentImages.map((imageItem, imgIdx) => {
                            const isImgSelected =
                              !isLaserActive &&
                              selectedImageId === imageItem.id;
                            return (
                              <figure
                                key={imageItem.id}
                                onClick={(event) => {
                                  if (isLaserActive) {
                                    return;
                                  }
                                  event.stopPropagation();
                                  setSelectedImageId(imageItem.id);
                                  setSelectedBlockId(null);
                                }}
                                className={`rounded-xl overflow-hidden border bg-black/25 flex items-center justify-center p-1.5 cursor-pointer transition-all ${
                                  isImgSelected
                                    ? 'border-indigo-400 ring-2 ring-indigo-500'
                                    : 'border-white/15 hover:border-white/35'
                                }`}
                              >
                                <img
                                  src={imageItem.src}
                                  alt={
                                    imageItem.alt ??
                                    `Hình ảnh ${imgIdx + 1} của Slide ${safeSlideIdx + 1}`
                                  }
                                  className="max-h-44 sm:max-h-56 w-full object-contain rounded-lg"
                                />
                              </figure>
                            );
                          })}
                        </div>
                      )}

                      {!hasImages && currentSlide.metricHighlight && (
                        <div
                          className="col-span-5 sm:col-span-4 rounded-xl sm:rounded-2xl p-2.5 sm:p-4 md:p-5 border backdrop-blur-md flex flex-col items-end justify-center text-right"
                          style={{
                            backgroundColor: `${accentColor}24`,
                            borderColor: `${accentColor}66`,
                          }}
                        >
                          <span className="text-base sm:text-2xl md:text-3xl font-bold text-white leading-none tracking-tight tabular-nums">
                            {currentSlide.metricHighlight.value}
                          </span>
                          <span className="mt-1 sm:mt-2 text-[9.5px] sm:text-xs text-slate-200/90 leading-snug line-clamp-3">
                            {currentSlide.metricHighlight.label}
                          </span>
                        </div>
                      )}

                      {currentBlocks.length === 0 && !hasImages && (
                        <div className="col-span-1 flex flex-col items-center justify-center text-center py-8 opacity-75">
                          <Type className="h-7 w-7 mb-2 text-indigo-300" />
                          <p className="text-xs sm:text-sm font-medium">
                            Slide đang trống — bấm{' '}
                            <strong>+ Thêm khung chữ</strong> phía trên để bắt
                            đầu soạn nội dung.
                          </p>
                        </div>
                      )}
                    </div>

                    {/* Slide Bottom Progress Strip */}
                    <div className="shrink-0 flex items-center justify-between gap-3 pt-2">
                      <div
                        className={`flex-1 h-1 rounded-full overflow-hidden ${
                          hasLightSlideBg ? 'bg-slate-900/10' : 'bg-white/10'
                        }`}
                      >
                        <div
                          className="h-full rounded-full transition-all duration-300"
                          style={{
                            width: `${((safeSlideIdx + 1) / slides.length) * 100}%`,
                            backgroundColor: accentColor,
                          }}
                        />
                      </div>
                      <span
                        className={`text-[10px] sm:text-xs truncate max-w-[45%] ${
                          hasLightSlideBg ? 'text-slate-500' : 'text-slate-400'
                        }`}
                      >
                        {document.title}
                      </span>
                    </div>
                  </div>
                )}

                {/* Live Glowing Red Laser Pointer Layer (Editor Stage) */}
                {isLaserActive && laserPos.visible && (
                  <div
                    aria-hidden="true"
                    className="pointer-events-none absolute z-20"
                    style={{
                      left: `${laserPos.x}%`,
                      top: `${laserPos.y}%`,
                    }}
                  >
                    <span className="absolute -translate-x-1/2 -translate-y-1/2 h-16 w-16 rounded-full bg-red-500/25 blur-md" />
                    <span className="absolute -translate-x-1/2 -translate-y-1/2 h-7 w-7 rounded-full border border-rose-400/75 bg-red-500/20" />
                    <span className="absolute -translate-x-1/2 -translate-y-1/2 h-3 w-3 rounded-full bg-red-500 ring-2 ring-white/95 shadow-[0_0_18px_5px_rgba(239,68,68,0.95)]" />
                  </div>
                )}

                {/* Left / Right Edge Slide Navigation Buttons */}
                {slides.length > 1 && (
                  <>
                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        handlePrevSlide();
                      }}
                      onPointerMove={(event) => event.stopPropagation()}
                      disabled={safeSlideIdx === 0}
                      aria-label="Slide trước"
                      className="cursor-pointer absolute left-2.5 sm:left-3.5 top-1/2 -translate-y-1/2 z-20 h-9 w-9 sm:h-10 sm:w-10 rounded-full bg-black/45 hover:bg-black/75 border border-white/15 text-white flex items-center justify-center backdrop-blur-sm transition disabled:opacity-25 disabled:cursor-not-allowed"
                    >
                      <ChevronLeft className="h-5 w-5" />
                    </button>
                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        handleNextSlide();
                      }}
                      onPointerMove={(event) => event.stopPropagation()}
                      disabled={safeSlideIdx >= slides.length - 1}
                      aria-label="Slide tiếp theo"
                      className="cursor-pointer absolute right-2.5 sm:right-3.5 top-1/2 -translate-y-1/2 z-20 h-9 w-9 sm:h-10 sm:w-10 rounded-full bg-black/45 hover:bg-black/75 border border-white/15 text-white flex items-center justify-center backdrop-blur-sm transition disabled:opacity-25 disabled:cursor-not-allowed"
                    >
                      <ChevronRight className="h-5 w-5" />
                    </button>
                  </>
                )}
              </div>
            </div>

            {/* Stage Footer Indicator & Quick Controls */}
            <div className="shrink-0 pt-2 flex items-center justify-between gap-2">
              <button
                type="button"
                onClick={() => setIsMobileDrawerOpen(true)}
                className="md:hidden inline-flex items-center gap-1.5 min-h-[34px] px-2.5 rounded-lg bg-[#141923] hover:bg-white/10 border border-white/10 text-[11.5px] font-medium text-slate-200 transition-colors"
              >
                <Layers className="h-3.5 w-3.5 text-indigo-400" />
                <span>Slide</span>
              </button>

              <p className="hidden md:block text-[11px] text-slate-400">
                Phím{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white/10 text-slate-200 font-mono text-[10px]">
                  ←
                </kbd>{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white/10 text-slate-200 font-mono text-[10px]">
                  →
                </kbd>{' '}
                chuyển slide ·{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white/10 text-slate-200 font-mono text-[10px]">
                  Ctrl+S
                </kbd>{' '}
                lưu nhanh
              </p>

              <div className="flex items-center gap-2.5 mx-auto md:mx-0">
                <button
                  type="button"
                  onClick={handlePrevSlide}
                  disabled={safeSlideIdx === 0}
                  aria-label="Chuyển slide trước"
                  className="h-7 w-7 rounded-lg bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 flex items-center justify-center text-slate-200 transition disabled:opacity-30 disabled:cursor-not-allowed"
                >
                  <ChevronLeft className="h-4 w-4" />
                </button>

                <span className="text-xs font-semibold text-slate-200 tabular-nums whitespace-nowrap">
                  Slide {safeSlideIdx + 1} / {slides.length}
                </span>

                <button
                  type="button"
                  onClick={handleNextSlide}
                  disabled={safeSlideIdx >= slides.length - 1}
                  aria-label="Chuyển slide kế tiếp"
                  className="h-7 w-7 rounded-lg bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 flex items-center justify-center text-slate-200 transition disabled:opacity-30 disabled:cursor-not-allowed"
                >
                  <ChevronRight className="h-4 w-4" />
                </button>
              </div>

              <div className="hidden sm:flex items-center gap-1.5 max-w-[160px] overflow-x-auto no-scrollbar">
                {slides.map((slideItem, idx) => (
                  <button
                    key={`${slideItem.id}-dot`}
                    type="button"
                    onClick={() => handleSelectSlide(idx)}
                    aria-label={`Đi tới Slide ${idx + 1}`}
                    className={`h-2 shrink-0 rounded-full transition-all ${
                      idx === safeSlideIdx
                        ? 'w-5 bg-indigo-400'
                        : 'w-2 bg-white/20 hover:bg-white/45'
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>

          {/* Bottom Speaker Notes Dock */}
          <section
            aria-label="Ghi chú diễn giả"
            className="shrink-0 border-t border-white/10 bg-[#141923] pb-14 md:pb-0"
          >
            <div className="px-3.5 sm:px-5 py-2 flex items-center justify-between gap-2">
              <div className="flex items-center gap-2 min-w-0">
                <StickyNote className="h-3.5 w-3.5 text-indigo-400 shrink-0" />
                <span className="text-xs font-semibold text-slate-200 whitespace-nowrap">
                  Ghi chú diễn giả
                </span>
                <span className="text-[11px] text-slate-500 tabular-nums whitespace-nowrap">
                  (Slide {safeSlideIdx + 1})
                </span>

                {!isNotesExpanded && currentSlide.speakerNote && (
                  <p className="hidden sm:block text-xs text-slate-400 truncate max-w-md ml-2">
                    — {currentSlide.speakerNote}
                  </p>
                )}
              </div>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => setIsNotesExpanded((prev) => !prev)}
                aria-expanded={isNotesExpanded}
                className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] text-[11px] font-medium text-slate-300 hover:text-white transition-colors"
              >
                <span>{isNotesExpanded ? 'Thu gọn' : 'Mở rộng'}</span>
                {isNotesExpanded ? (
                  <ChevronDown className="h-3.5 w-3.5" />
                ) : (
                  <ChevronUp className="h-3.5 w-3.5" />
                )}
              </button>
            </div>

            {isNotesExpanded && (
              <div className="px-3.5 sm:px-5 pb-3">
                <textarea
                  value={currentSlide.speakerNote ?? ''}
                  onChange={(event) =>
                    handleSpeakerNoteChange(event.target.value)
                  }
                  rows={2}
                  aria-label={`Nội dung ghi chú diễn giả cho Slide ${safeSlideIdx + 1}`}
                  placeholder="Nhập ghi chú diễn giả cho slide này (được lưu và xuất kèm tệp .pptx)..."
                  className="w-full h-[64px] sm:h-[74px] resize-none rounded-xl bg-[#181D27] border border-white/10 focus:border-indigo-400/70 focus:outline-none focus:ring-2 focus:ring-indigo-500/25 px-3.5 py-2 text-xs sm:text-[13px] text-slate-200 placeholder:text-slate-500 leading-relaxed transition-colors scrollbar-dark"
                />
              </div>
            )}
            <div className="pb-safe" />
          </section>

          {/* Floating Selection Formatting Toolbar (above selected text in slide frame) */}
          <FloatingSelectionToolbar
            containerRef={editorCanvasRef}
            disabled={isLaserActive || isSlideshowActive || isModalOpen}
            textColorSwatches={SLIDE_COLOR_SWATCHES}
            activeState={
              selectedBlock
                ? {
                    block:
                      selectedBlock.fontSize >= 32
                        ? 'h1'
                        : selectedBlock.fontSize >= 24
                          ? 'h2'
                          : 'p',
                    bold: Boolean(selectedBlock.bold),
                    italic: Boolean(selectedBlock.italic),
                    underline: Boolean(underlinedBlockIds[selectedBlock.id]),
                    textColor:
                      selectedBlock.color ??
                      (hasLightSlideBg ? '#0F172A' : '#FFFFFF'),
                    highlightColor: selectedBlock.fillColor ?? 'transparent',
                  }
                : undefined
            }
            onSelectBlock={(tag) => {
              if (selectedBlock) {
                handleApplyBlockPreset(selectedBlock.id, tag);
              }
            }}
            onToggleBold={() => {
              if (selectedBlock) {
                handleUpdateBlock(selectedBlock.id, {
                  bold: !selectedBlock.bold,
                });
              }
            }}
            onToggleItalic={() => {
              if (selectedBlock) {
                handleUpdateBlock(selectedBlock.id, {
                  italic: !selectedBlock.italic,
                });
              }
            }}
            onToggleUnderline={() => {
              if (selectedBlock) {
                handleToggleBlockUnderline(selectedBlock.id);
              }
            }}
            onSelectTextColor={(color) => {
              if (selectedBlock) {
                handleUpdateBlock(selectedBlock.id, { color });
              }
            }}
            onSelectHighlight={(color) => {
              if (selectedBlock) {
                handleUpdateBlock(selectedBlock.id, {
                  fillColor:
                    !color || color === 'transparent' ? undefined : color,
                });
              }
            }}
            onClearFormat={() => {
              if (selectedBlock) {
                handleClearBlockFormat(selectedBlock.id);
              }
            }}
          />

          {/* Mobile Collapsed 1-Row Bottom Bar (md:hidden) */}
          <div
            role="toolbar"
            aria-label="Thanh công cụ chỉnh sửa slide di động"
            className="md:hidden fixed bottom-0 inset-x-0 z-40 h-12 bg-[#141923] border-t border-white/15 flex items-center text-slate-100"
          >
            <div className="overflow-x-auto no-scrollbar flex items-center gap-1 px-2 flex-1 min-w-0">
              {/* + Khung chữ */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleAddTextBlock}
                aria-label="Thêm khung chữ"
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-indigo-600/30 border border-indigo-400/40 text-[11px] font-semibold text-indigo-100 whitespace-nowrap"
              >
                <Plus className="h-3.5 w-3.5" />
                <span>Khung chữ</span>
              </button>

              {/* + Ảnh */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleTriggerAddImage}
                aria-label="Thêm ảnh"
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white/[0.07] border border-white/10 text-[11px] font-medium text-slate-200 whitespace-nowrap"
              >
                <ImagePlus className="h-3.5 w-3.5 text-emerald-300" />
                <span>Ảnh</span>
              </button>

              <span className="h-5 w-px bg-white/15 shrink-0 mx-0.5" />

              {/* Bold */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                disabled={!selectedBlock}
                onClick={() => {
                  if (selectedBlock) {
                    handleUpdateBlock(selectedBlock.id, {
                      bold: !selectedBlock.bold,
                    });
                  }
                }}
                aria-pressed={Boolean(selectedBlock?.bold)}
                aria-label="In đậm"
                className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors disabled:opacity-35 ${
                  selectedBlock?.bold
                    ? 'bg-indigo-500/35 border-indigo-400/60 text-white'
                    : 'bg-white/[0.06] border-white/10 text-slate-200'
                }`}
              >
                <Bold className="h-3.5 w-3.5" />
              </button>

              {/* Italic */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                disabled={!selectedBlock}
                onClick={() => {
                  if (selectedBlock) {
                    handleUpdateBlock(selectedBlock.id, {
                      italic: !selectedBlock.italic,
                    });
                  }
                }}
                aria-pressed={Boolean(selectedBlock?.italic)}
                aria-label="In nghiêng"
                className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors disabled:opacity-35 ${
                  selectedBlock?.italic
                    ? 'bg-indigo-500/35 border-indigo-400/60 text-white'
                    : 'bg-white/[0.06] border-white/10 text-slate-200'
                }`}
              >
                <Italic className="h-3.5 w-3.5" />
              </button>

              {/* Underline */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                disabled={!selectedBlock}
                onClick={() => {
                  if (selectedBlock) {
                    handleToggleBlockUnderline(selectedBlock.id);
                  }
                }}
                aria-pressed={Boolean(
                  selectedBlock && underlinedBlockIds[selectedBlock.id]
                )}
                aria-label="Gạch chân"
                className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors disabled:opacity-35 ${
                  selectedBlock && underlinedBlockIds[selectedBlock.id]
                    ? 'bg-indigo-500/35 border-indigo-400/60 text-white'
                    : 'bg-white/[0.06] border-white/10 text-slate-200'
                }`}
              >
                <Underline className="h-3.5 w-3.5" />
              </button>

              {/* Màu chữ (cycles swatches or opens detail sheet) */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => {
                  if (!selectedBlock) {
                    setIsMobileSheetOpen(true);
                    return;
                  }
                  const currentHex = (
                    selectedBlock.color ??
                    (hasLightSlideBg ? '#1E293B' : '#FFFFFF')
                  ).toLowerCase();
                  const curIdx = SLIDE_COLOR_SWATCHES.findIndex(
                    (s) => s.value.toLowerCase() === currentHex
                  );
                  const nextSwatch =
                    SLIDE_COLOR_SWATCHES[
                      (curIdx + 1) % SLIDE_COLOR_SWATCHES.length
                    ];
                  handleUpdateBlock(selectedBlock.id, {
                    color: nextSwatch.value,
                  });
                }}
                aria-label="Đổi màu chữ"
                title="Chạm để đổi màu chữ nhanh"
                className="shrink-0 inline-flex flex-col items-center justify-center h-8 w-8 rounded-lg bg-white/[0.06] border border-white/10 text-slate-100"
              >
                <span className="text-[11px] font-bold leading-none">A</span>
                <span
                  className="mt-0.5 h-1 w-4 rounded-full"
                  style={{
                    backgroundColor:
                      selectedBlock?.color ??
                      (hasLightSlideBg ? '#0F172A' : '#FFFFFF'),
                  }}
                />
              </button>

              {/* Căn lề: Trái / Giữa / Phải */}
              {(
                [
                  { value: 'left', label: 'Căn trái', Icon: AlignLeft },
                  { value: 'center', label: 'Căn giữa', Icon: AlignCenter },
                  { value: 'right', label: 'Căn phải', Icon: AlignRight },
                ] as const
              ).map(({ value, label, Icon }) => {
                const active =
                  Boolean(selectedBlock) &&
                  (selectedBlock?.align ?? 'left') === value;
                return (
                  <button
                    key={value}
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    disabled={!selectedBlock}
                    onClick={() => {
                      if (selectedBlock) {
                        handleUpdateBlock(selectedBlock.id, { align: value });
                      }
                    }}
                    aria-pressed={active}
                    aria-label={label}
                    className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border transition-colors disabled:opacity-35 ${
                      active
                        ? 'bg-indigo-600 border-indigo-400 text-white'
                        : 'bg-white/[0.06] border-white/10 text-slate-200'
                    }`}
                  >
                    <Icon className="h-3.5 w-3.5" />
                  </button>
                );
              })}

              <span className="h-5 w-px bg-white/15 shrink-0 mx-0.5" />

              {/* + Slide */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleAddSlide}
                aria-label="Thêm slide mới"
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white/[0.07] border border-white/10 text-[11px] font-semibold text-slate-100 whitespace-nowrap"
              >
                <Plus className="h-3.5 w-3.5 text-indigo-300" />
                <span>Slide</span>
              </button>
            </div>

            {/* Right Fixed Expand Arrow ▲ */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => setIsMobileSheetOpen((prev) => !prev)}
              aria-expanded={isMobileSheetOpen}
              aria-label="Mở bảng công cụ chi tiết"
              title="Mở bảng công cụ chi tiết"
              className="shrink-0 h-full w-11 border-l border-white/15 flex items-center justify-center text-indigo-300 hover:bg-white/[0.08] transition-colors"
            >
              <ChevronUp className="h-5 w-5" />
            </button>
          </div>

          {/* Mobile Expanded Detail Bottom Sheet (md:hidden) */}
          {isMobileSheetOpen && (
            <div
              role="region"
              aria-label="Bảng công cụ chỉnh sửa slide chi tiết"
              className="md:hidden fixed bottom-0 inset-x-0 z-40 h-[44vh] bg-[#141923] border-t border-white/15 rounded-t-2xl shadow-2xl flex flex-col text-slate-100"
            >
              {/* Sheet Header */}
              <div className="shrink-0 px-3.5 py-2.5 border-b border-white/10 flex items-center justify-between gap-2">
                <div className="flex items-center gap-2 min-w-0">
                  <span className="h-2 w-2 rounded-full bg-indigo-400 shrink-0" />
                  <span className="text-xs font-semibold text-white truncate">
                    {selectedBlock
                      ? `Đang sửa khung chữ · Slide ${safeSlideIdx + 1}/${slides.length}`
                      : selectedImage
                        ? `Đang chọn ảnh · Slide ${safeSlideIdx + 1}/${slides.length}`
                        : `Công cụ Slide ${safeSlideIdx + 1} / ${slides.length}`}
                  </span>
                </div>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setIsMobileSheetOpen(false)}
                  aria-label="Thu gọn bảng công cụ"
                  title="Thu gọn bảng công cụ"
                  className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white/[0.08] hover:bg-white/[0.15] text-xs font-medium text-slate-200"
                >
                  <span>Thu gọn</span>
                  <ChevronDown className="h-4 w-4" />
                </button>
              </div>

              {/* Sheet Body */}
              <div className="flex-1 min-h-0 overflow-y-auto scrollbar-dark p-3.5 space-y-3.5">
                {/* Section 1: Định dạng khung chữ / ảnh */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                      Định dạng khung đang chọn
                    </span>
                    {selectedBlock && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleDeleteBlock(selectedBlock.id)}
                        className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-rose-500/20 border border-rose-400/40 text-[11px] font-semibold text-rose-200"
                      >
                        <Trash2 className="h-3 w-3" />
                        <span>Xóa khung</span>
                      </button>
                    )}
                    {selectedImage && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleDeleteImage(selectedImage.id)}
                        className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-rose-500/20 border border-rose-400/40 text-[11px] font-semibold text-rose-200"
                      >
                        <Trash2 className="h-3 w-3" />
                        <span>Xóa ảnh</span>
                      </button>
                    )}
                  </div>

                  {/* Row 1A: H1/H2/P + Font Sizes */}
                  <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
                    {(
                      [
                        {
                          tag: 'h1',
                          label: 'H1',
                          active: Boolean(
                            selectedBlock && selectedBlock.fontSize >= 32
                          ),
                        },
                        {
                          tag: 'h2',
                          label: 'H2',
                          active: Boolean(
                            selectedBlock &&
                              selectedBlock.fontSize >= 24 &&
                              selectedBlock.fontSize < 32
                          ),
                        },
                        {
                          tag: 'p',
                          label: 'P',
                          active: Boolean(
                            selectedBlock && selectedBlock.fontSize < 24
                          ),
                        },
                      ] as const
                    ).map(({ tag, label, active }) => (
                      <button
                        key={tag}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        disabled={!selectedBlock}
                        onClick={() => {
                          if (selectedBlock) {
                            handleApplyBlockPreset(selectedBlock.id, tag);
                          }
                        }}
                        className={`shrink-0 h-8 px-2.5 rounded-lg border text-xs font-bold disabled:opacity-35 ${
                          active
                            ? 'bg-indigo-600 border-indigo-400 text-white'
                            : 'bg-[#1C2331] border-white/10 text-slate-200'
                        }`}
                      >
                        {label}
                      </button>
                    ))}

                    <span className="h-5 w-px bg-white/15 shrink-0" />

                    {SLIDE_FONT_SIZES.map((size) => {
                      const active = selectedBlock?.fontSize === size;
                      return (
                        <button
                          key={size}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          disabled={!selectedBlock}
                          onClick={() => {
                            if (selectedBlock) {
                              handleUpdateBlock(selectedBlock.id, {
                                fontSize: size,
                              });
                            }
                          }}
                          className={`shrink-0 h-8 px-2.5 rounded-lg border text-xs font-semibold tabular-nums disabled:opacity-35 ${
                            active
                              ? 'bg-indigo-600 border-indigo-400 text-white'
                              : 'bg-[#1C2331] border-white/10 text-slate-200'
                          }`}
                        >
                          {size}
                        </button>
                      );
                    })}
                  </div>

                  {/* Row 1B: B / I / U + Alignment */}
                  <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      disabled={!selectedBlock}
                      onClick={() => {
                        if (selectedBlock) {
                          handleUpdateBlock(selectedBlock.id, {
                            bold: !selectedBlock.bold,
                          });
                        }
                      }}
                      className={`shrink-0 h-9 px-3 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-35 ${
                        selectedBlock?.bold
                          ? 'bg-indigo-600 border-indigo-400 text-white'
                          : 'bg-[#1C2331] border-white/10 text-slate-200'
                      }`}
                    >
                      <Bold className="h-3.5 w-3.5" />
                      <span>Đậm</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      disabled={!selectedBlock}
                      onClick={() => {
                        if (selectedBlock) {
                          handleUpdateBlock(selectedBlock.id, {
                            italic: !selectedBlock.italic,
                          });
                        }
                      }}
                      className={`shrink-0 h-9 px-3 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-35 ${
                        selectedBlock?.italic
                          ? 'bg-indigo-600 border-indigo-400 text-white'
                          : 'bg-[#1C2331] border-white/10 text-slate-200'
                      }`}
                    >
                      <Italic className="h-3.5 w-3.5" />
                      <span>Nghiêng</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      disabled={!selectedBlock}
                      onClick={() => {
                        if (selectedBlock) {
                          handleToggleBlockUnderline(selectedBlock.id);
                        }
                      }}
                      className={`shrink-0 h-9 px-3 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-35 ${
                        selectedBlock && underlinedBlockIds[selectedBlock.id]
                          ? 'bg-indigo-600 border-indigo-400 text-white'
                          : 'bg-[#1C2331] border-white/10 text-slate-200'
                      }`}
                    >
                      <Underline className="h-3.5 w-3.5" />
                      <span>Gạch chân</span>
                    </button>

                    {(
                      [
                        { value: 'left', label: 'Căn trái', Icon: AlignLeft },
                        { value: 'center', label: 'Căn giữa', Icon: AlignCenter },
                        { value: 'right', label: 'Căn phải', Icon: AlignRight },
                      ] as const
                    ).map(({ value, label, Icon }) => {
                      const active =
                        Boolean(selectedBlock) &&
                        (selectedBlock?.align ?? 'left') === value;
                      return (
                        <button
                          key={value}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          disabled={!selectedBlock}
                          onClick={() => {
                            if (selectedBlock) {
                              handleUpdateBlock(selectedBlock.id, {
                                align: value,
                              });
                            }
                          }}
                          aria-label={label}
                          className={`shrink-0 h-9 w-9 rounded-lg border flex items-center justify-center disabled:opacity-35 ${
                            active
                              ? 'bg-indigo-600 border-indigo-400 text-white'
                              : 'bg-[#1C2331] border-white/10 text-slate-200'
                          }`}
                        >
                          <Icon className="h-4 w-4" />
                        </button>
                      );
                    })}
                  </div>

                  {/* Row 1C: Text Color Swatches + Block Fill Color */}
                  <div className="flex items-center gap-2 overflow-x-auto no-scrollbar pt-0.5">
                    <span className="text-[11px] text-slate-400 shrink-0">
                      Màu chữ:
                    </span>
                    {SLIDE_COLOR_SWATCHES.map((swatch) => {
                      const active =
                        selectedBlock &&
                        (selectedBlock.color ?? '#FFFFFF').toLowerCase() ===
                          swatch.value.toLowerCase();
                      return (
                        <button
                          key={swatch.value}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          disabled={!selectedBlock}
                          onClick={() => {
                            if (selectedBlock) {
                              handleUpdateBlock(selectedBlock.id, {
                                color: swatch.value,
                              });
                            }
                          }}
                          aria-label={`Màu chữ ${swatch.label}`}
                          className={`shrink-0 h-6 w-6 rounded-full border disabled:opacity-35 ${
                            active
                              ? 'ring-2 ring-indigo-400 border-white scale-110'
                              : 'border-white/30'
                          }`}
                          style={{ backgroundColor: swatch.value }}
                        />
                      );
                    })}

                    <span className="h-5 w-px bg-white/15 shrink-0 mx-0.5" />

                    <span className="text-[11px] text-slate-400 shrink-0">
                      Nền khối:
                    </span>
                    <input
                      type="color"
                      disabled={!selectedBlock}
                      value={toSafeColorInputValue(
                        selectedBlock?.fillColor,
                        '#1E293B'
                      )}
                      onChange={(event) => {
                        if (selectedBlock) {
                          handleUpdateBlock(selectedBlock.id, {
                            fillColor: event.target.value,
                          });
                        }
                      }}
                      aria-label="Màu nền khối chữ"
                      className="h-6 w-7 rounded cursor-pointer border-0 bg-transparent p-0 disabled:opacity-35"
                    />
                    {selectedBlock?.fillColor && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() =>
                          handleUpdateBlock(selectedBlock.id, {
                            fillColor: undefined,
                          })
                        }
                        className="shrink-0 px-2 py-1 rounded bg-white/10 text-[11px] text-slate-200"
                      >
                        Trong suốt
                      </button>
                    )}
                  </div>
                </div>

                {/* Section 2: Chèn & Quản lý Slide */}
                <div className="space-y-1.5 pt-1 border-t border-white/10">
                  <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                    Chèn & Quản lý Slide
                  </span>
                  <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddTextBlock}
                      className="shrink-0 h-9 px-3 rounded-xl bg-indigo-600/25 border border-indigo-400/40 text-xs font-semibold text-indigo-100 inline-flex items-center gap-1.5"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      <Type className="h-3.5 w-3.5" />
                      <span>+ Khung chữ</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleTriggerAddImage}
                      className="shrink-0 h-9 px-3 rounded-xl bg-[#1C2331] border border-white/10 text-xs font-medium text-slate-200 inline-flex items-center gap-1.5"
                    >
                      <ImagePlus className="h-3.5 w-3.5 text-emerald-300" />
                      <span>+ Ảnh</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-[#1C2331] border border-white/10 text-xs font-medium text-slate-200 inline-flex items-center gap-1.5"
                    >
                      <Plus className="h-3.5 w-3.5 text-indigo-300" />
                      <span>+ Slide</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleDuplicateSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-[#1C2331] border border-white/10 text-xs font-medium text-slate-200 inline-flex items-center gap-1.5"
                    >
                      <Copy className="h-3.5 w-3.5" />
                      <span>Nhân bản</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      disabled={slides.length <= 1}
                      onClick={handleDeleteSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-rose-500/15 border border-rose-400/35 text-xs font-medium text-rose-200 inline-flex items-center gap-1.5 disabled:opacity-35"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      <span>Xóa Slide</span>
                    </button>
                  </div>
                </div>

                {/* Section 3: Màu nền Slide & Tiện ích */}
                <div className="space-y-1.5 pt-1 border-t border-white/10">
                  <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                    Màu nền Slide & Tiện ích
                  </span>
                  <div className="flex items-center gap-2 overflow-x-auto no-scrollbar">
                    {SLIDE_BG_PRESETS.map((bg) => {
                      const active =
                        !currentSlide.backgroundImageSrc &&
                        !currentSlide.backgroundGradient &&
                        (currentSlide.backgroundColor ?? '#161D2E').toLowerCase() ===
                          bg.value.toLowerCase();
                      return (
                        <button
                          key={bg.value}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleChangeSlideBackground(bg.value)}
                          className={`shrink-0 h-8 px-2.5 rounded-lg border text-[11px] font-medium inline-flex items-center gap-1.5 ${
                            active
                              ? 'border-indigo-400 bg-indigo-500/20 text-white'
                              : 'border-white/15 bg-[#1C2331] text-slate-200'
                          }`}
                        >
                          <span
                            className="h-3.5 w-3.5 rounded border border-white/30"
                            style={{ backgroundColor: bg.value }}
                          />
                          <span>{bg.label}</span>
                        </button>
                      );
                    })}

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsLaserActive((prev) => !prev)}
                      aria-pressed={isLaserActive}
                      className={`shrink-0 h-8 px-2.5 rounded-lg border text-[11px] font-medium inline-flex items-center gap-1.5 ${
                        isLaserActive
                          ? 'bg-rose-500/25 border-rose-400/60 text-rose-100'
                          : 'bg-[#1C2331] border-white/10 text-slate-200'
                      }`}
                    >
                      <Crosshair className="h-3.5 w-3.5" />
                      <span>Bút Laser</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsNotesExpanded((prev) => !prev)}
                      aria-pressed={isNotesExpanded}
                      className={`shrink-0 h-8 px-2.5 rounded-lg border text-[11px] font-medium inline-flex items-center gap-1.5 ${
                        isNotesExpanded
                          ? 'bg-indigo-500/25 border-indigo-400/60 text-indigo-100'
                          : 'bg-[#1C2331] border-white/10 text-slate-200'
                      }`}
                    >
                      <StickyNote className="h-3.5 w-3.5" />
                      <span>Ghi chú diễn giả</span>
                    </button>
                  </div>
                </div>
              </div>
              <div className="pb-safe" />
            </div>
          )}
        </main>
      </div>

      {/* Real Fullscreen Slideshow Mode (#000000) */}
      {isSlideshowActive && (
        <div
          ref={slideshowViewportRef}
          role="dialog"
          aria-modal="true"
          aria-label={`Trình chiếu toàn màn hình: ${document.title}`}
          className="fixed inset-0 z-50 bg-black flex items-center justify-center overflow-hidden select-none"
        >
          <style>{`
@keyframes liteSlideFromRight {
  from { opacity: 0.25; transform: translateX(48px); }
  to { opacity: 1; transform: translateX(0); }
}
@keyframes liteSlideFromLeft {
  from { opacity: 0.25; transform: translateX(-48px); }
  to { opacity: 1; transform: translateX(0); }
}
@keyframes liteSlideFadeIn {
  from { opacity: 0; }
  to { opacity: 1; }
}
`}</style>

          {/* Proportional Fullscreen Slide Stage */}
          <div
            key={`slideshow-${currentSlide.id}-${safeSlideIdx}`}
            onClick={() => {
              if (isSlideshowLaser || isSlideshowGridOpen) {
                return;
              }
              setSlideshowDirection(1);
              setActiveSlideIdx((prev) =>
                prev < slides.length - 1 ? prev + 1 : prev
              );
            }}
            onPointerMove={updateSlideshowLaserFromPointer}
            onPointerDown={updateSlideshowLaserFromPointer}
            onPointerLeave={() => {
              if (slideshowLaserPos.visible) {
                setSlideshowLaserPos((prev) => ({ ...prev, visible: false }));
              }
            }}
            style={{
              width: `${Math.round(slideshowFitWidthPx)}px`,
              height: `${Math.round(slideshowFitHeightPx)}px`,
              ...slideshowAnimationStyle,
            }}
            className={`relative overflow-hidden shadow-2xl ${
              isSlideshowLaser
                ? 'cursor-none touch-none'
                : 'cursor-pointer'
            }`}
          >
            <ReadOnlySlideStage
              doc={document}
              slide={currentSlide}
              slideIndex={safeSlideIdx}
              totalSlides={slides.length}
              stageWidthPx={slideshowFitWidthPx}
              underlinedBlockIds={underlinedBlockIds}
            />

            {/* Slideshow Red Laser Pointer */}
            {isSlideshowLaser && slideshowLaserPos.visible && (
              <div
                aria-hidden="true"
                className="pointer-events-none absolute z-30"
                style={{
                  left: `${slideshowLaserPos.x}%`,
                  top: `${slideshowLaserPos.y}%`,
                }}
              >
                <span className="absolute -translate-x-1/2 -translate-y-1/2 h-16 w-16 rounded-full bg-red-500/30 blur-md" />
                <span className="absolute -translate-x-1/2 -translate-y-1/2 h-7 w-7 rounded-full border border-rose-400/80 bg-red-500/20" />
                <span className="absolute -translate-x-1/2 -translate-y-1/2 h-3.5 w-3.5 rounded-full bg-red-500 ring-2 ring-white shadow-[0_0_20px_6px_rgba(239,68,68,0.95)]" />
              </div>
            )}
          </div>

          {/* Slide Grid Quick Jump Overlay */}
          {isSlideshowGridOpen && (
            <div
              onClick={(event) => {
                event.stopPropagation();
                if (event.target === event.currentTarget) {
                  setIsSlideshowGridOpen(false);
                }
              }}
              className="fixed inset-0 z-40 bg-black/85 backdrop-blur-md p-4 sm:p-8 flex flex-col overflow-hidden"
            >
              <div className="shrink-0 flex items-center justify-between gap-3 mb-4">
                <div className="flex items-center gap-2.5">
                  <LayoutGrid className="h-5 w-5 text-indigo-400" />
                  <h2 className="text-sm sm:text-base font-semibold text-white">
                    Lưới Slide ({slides.length} trang)
                  </h2>
                </div>
                <button
                  type="button"
                  onClick={() => setIsSlideshowGridOpen(false)}
                  aria-label="Đóng lưới slide"
                  className="h-9 px-3 rounded-xl bg-white/10 hover:bg-white/20 border border-white/15 text-xs font-medium text-white inline-flex items-center gap-1.5"
                >
                  <X className="h-4 w-4" />
                  <span>Đóng lưới</span>
                </button>
              </div>

              <div className="flex-1 min-h-0 overflow-y-auto scrollbar-dark grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3.5 pb-16">
                {slides.map((slideItem, idx) => (
                  <GridSlideThumbnail
                    key={`grid-${slideItem.id}`}
                    doc={document}
                    slide={slideItem}
                    slideIndex={idx}
                    totalSlides={slides.length}
                    isActive={idx === safeSlideIdx}
                    underlinedBlockIds={underlinedBlockIds}
                    onSelect={() => {
                      setSlideshowDirection(idx >= safeSlideIdx ? 1 : -1);
                      setActiveSlideIdx(idx);
                      setIsSlideshowGridOpen(false);
                    }}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Bottom Corner Translucent Slideshow Control HUD */}
          <div
            onClick={(event) => event.stopPropagation()}
            onPointerMove={(event) => event.stopPropagation()}
            className="fixed bottom-4 left-1/2 -translate-x-1/2 sm:left-auto sm:translate-x-0 sm:right-5 z-50 max-w-[96vw] overflow-x-auto no-scrollbar rounded-2xl bg-slate-900/75 hover:bg-slate-900/90 backdrop-blur-md border border-white/15 px-2.5 py-1.5 shadow-2xl flex items-center gap-1.5 text-xs text-slate-100 transition-colors pb-safe"
          >
            <button
              type="button"
              onClick={() => {
                setSlideshowDirection(-1);
                setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
              }}
              disabled={safeSlideIdx === 0}
              aria-label="Slide trước"
              className="h-8 px-2.5 rounded-xl bg-white/10 hover:bg-white/20 disabled:opacity-35 disabled:cursor-not-allowed inline-flex items-center gap-1 font-medium whitespace-nowrap"
            >
              <ChevronLeft className="h-4 w-4" />
              <span className="hidden sm:inline">Trước</span>
            </button>

            <span className="px-2 font-semibold tabular-nums text-slate-200 whitespace-nowrap">
              Slide {safeSlideIdx + 1} / {slides.length}
            </span>

            <button
              type="button"
              onClick={() => {
                setSlideshowDirection(1);
                setActiveSlideIdx((prev) =>
                  prev < slides.length - 1 ? prev + 1 : prev
                );
              }}
              disabled={safeSlideIdx >= slides.length - 1}
              aria-label="Slide tiếp"
              className="h-8 px-2.5 rounded-xl bg-white/10 hover:bg-white/20 disabled:opacity-35 disabled:cursor-not-allowed inline-flex items-center gap-1 font-medium whitespace-nowrap"
            >
              <span className="hidden sm:inline">Sau</span>
              <ChevronRight className="h-4 w-4" />
            </button>

            <span className="h-4 w-px bg-white/15 shrink-0" />

            {/* Transition Mode Toggle: Trượt / Mờ */}
            <button
              type="button"
              onClick={() =>
                setSlideshowTransition((prev) =>
                  prev === 'slide' ? 'fade' : 'slide'
                )
              }
              title="Đổi hiệu ứng chuyển slide (Trượt / Mờ)"
              className="h-8 px-2.5 rounded-xl bg-white/[0.07] hover:bg-white/15 border border-white/10 inline-flex items-center gap-1.5 font-medium whitespace-nowrap"
            >
              <Sparkles className="h-3.5 w-3.5 text-indigo-300" />
              <span>{slideshowTransition === 'slide' ? 'Trượt' : 'Mờ'}</span>
            </button>

            {/* Red Laser Toggle */}
            <button
              type="button"
              onClick={() => setIsSlideshowLaser((prev) => !prev)}
              aria-pressed={isSlideshowLaser}
              title="Bật / tắt bút Laser đỏ"
              className={`h-8 px-2.5 rounded-xl border inline-flex items-center gap-1.5 font-medium whitespace-nowrap transition-colors ${
                isSlideshowLaser
                  ? 'bg-rose-500/30 border-rose-400/60 text-rose-100'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <span
                className={`h-2 w-2 rounded-full ${
                  isSlideshowLaser ? 'bg-red-500 animate-ping' : 'bg-rose-400'
                }`}
              />
              <span>Laser</span>
            </button>

            {/* Slide Grid Toggle */}
            <button
              type="button"
              onClick={() => setIsSlideshowGridOpen((prev) => !prev)}
              aria-pressed={isSlideshowGridOpen}
              title="Mở lưới chọn nhanh slide"
              className={`h-8 px-2.5 rounded-xl border inline-flex items-center gap-1.5 font-medium whitespace-nowrap transition-colors ${
                isSlideshowGridOpen
                  ? 'bg-indigo-600 border-indigo-400 text-white'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <LayoutGrid className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">Lưới slide</span>
            </button>

            {/* Exit Slideshow */}
            <button
              type="button"
              onClick={exitSlideshow}
              aria-label="Thoát trình chiếu"
              className="h-8 px-2.5 rounded-xl bg-rose-500/20 hover:bg-rose-500/35 border border-rose-400/40 text-rose-100 font-semibold inline-flex items-center gap-1 whitespace-nowrap"
            >
              <X className="h-3.5 w-3.5" />
              <span>Thoát</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

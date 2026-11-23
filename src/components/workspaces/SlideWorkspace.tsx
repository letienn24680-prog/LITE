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
  BarChart3,
  Bold,
  BringToFront,
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
  Eraser,
  ExternalLink,
  Eye,
  EyeOff,
  FileDown,
  Group,
  Highlighter,
  Image,
  ImagePlus,
  Italic,
  Layers,
  LayoutGrid,
  LayoutTemplate,
  Link2,
  Loader2,
  MonitorPlay,
  Move,
  Music,
  PaintBucket,
  Palette,
  Pause,
  Pen,
  Pencil,
  Play,
  Plus,
  Presentation,
  QrCode,
  RefreshCw,
  RotateCcw,
  RotateCw,
  Save,
  SendToBack,
  Shapes,
  Sparkles,
  StickyNote,
  Table2,
  Trash2,
  Type,
  Underline,
  Ungroup,
  Video,
  Wand2,
  X,
} from 'lucide-react';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  SlideAnimationConfig,
  SlideChartData,
  SlideImageItem,
  SlidePreviewItem,
  SlideTextBlock,
  SlideTransitionConfig,
} from '../../types/document';
import {
  exportOriginalDocumentBinary,
  exportSlidesToPptx,
} from '../../utils/fileExporter';
import { useWorkspaceSync } from '../../contexts/WorkspaceSyncContext';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { useDocumentPinchZoom } from '../../hooks/useDocumentPinchZoom';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import {
  FloatingSelectionToolbar,
  type FloatingBlockTag,
} from './FloatingSelectionToolbar';
import { OfficeStatusBar } from './OfficeStatusBar';
import {
  PPT_COLOR_THEMES,
  PPT_SHAPE_OPTIONS,
  PPT_SLIDE_LAYOUTS,
  SlideChartInsertModal,
  SlideChartSvg,
  SlideLinkMediaModal,
  SlideShapeSvg,
  createBlocksForSlideLayout,
  type PptColorTheme,
  type PptShapeKind,
  type PptSlideLayoutId,
} from './SlideInsertChartShapes';
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
type SlideWorkspaceViewMode = 'native' | 'edit';
type PptRibbonTab = 'home' | 'insert' | 'design' | 'transitions' | 'slideshow';
type SlideshowDrawTool = 'none' | 'pen' | 'highlighter' | 'eraser';

interface SlideshowStroke {
  id: string;
  tool: 'pen' | 'highlighter';
  color: string;
  points: Array<{ x: number; y: number }>;
}

interface SmartGuidesState {
  xPct: number | null;
  yPct: number | null;
}

const PPT_RIBBON_TABS: ReadonlyArray<{ id: PptRibbonTab; label: string }> = [
  { id: 'home', label: 'Trang đầu' },
  { id: 'insert', label: 'Chèn' },
  { id: 'design', label: 'Thiết kế' },
  { id: 'transitions', label: 'Chuyển cảnh & Hoạt ảnh' },
  { id: 'slideshow', label: 'Trình chiếu' },
];

const PPT_TRANSITION_OPTIONS: ReadonlyArray<{
  type: SlideTransitionConfig['type'];
  label: string;
}> = [
  { type: 'none', label: 'Không' },
  { type: 'fade', label: 'Làm mờ' },
  { type: 'push', label: 'Đẩy sang' },
  { type: 'wipe', label: 'Quét ngang' },
  { type: 'split', label: 'Tách đôi' },
  { type: 'zoom', label: 'Phóng to' },
];

const PPT_ANIMATION_OPTIONS: ReadonlyArray<{
  type: SlideAnimationConfig['type'];
  label: string;
}> = [
  { type: 'none', label: 'Không' },
  { type: 'appear', label: 'Xuất hiện' },
  { type: 'fadeIn', label: 'Mờ dần vào' },
  { type: 'flyIn', label: 'Bay vào' },
  { type: 'zoomIn', label: 'Phóng vào' },
];

const PPT_GRADIENT_PRESETS: ReadonlyArray<{ label: string; value: string }> = [
  {
    label: 'Chàm đêm',
    value: 'linear-gradient(135deg, #0F172A 0%, #312E81 55%, #1E1B4B 100%)',
  },
  {
    label: 'Hoàng hôn cam',
    value: 'linear-gradient(135deg, #7C2D12 0%, #C43E1C 52%, #EA580C 100%)',
  },
  {
    label: 'Ngọc lục bảo',
    value: 'linear-gradient(135deg, #064E3B 0%, #047857 52%, #0F766E 100%)',
  },
  {
    label: 'Lụa sáng',
    value: 'linear-gradient(135deg, #FFFFFF 0%, #F1F5F9 55%, #E2E8F0 100%)',
  },
];

const PPT_FONT_FAMILIES = [
  { label: 'Inter', value: "'Inter', 'Segoe UI', system-ui, sans-serif" },
  { label: 'Segoe UI', value: "'Segoe UI', Roboto, sans-serif" },
  { label: 'Georgia', value: 'Georgia, serif' },
  { label: 'Times New Roman', value: "'Times New Roman', serif" },
  { label: 'Mono', value: "'JetBrains Mono', monospace" },
] as const;

function sanitizeExternalUrl(url: string | undefined): string | null {
  if (!url) {
    return null;
  }
  const trimmed = url.trim();
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return null;
}

function sortBlocksByZIndex(blocks: SlideTextBlock[]): SlideTextBlock[] {
  return [...blocks].sort((a, b) => (a.zIndex ?? 0) - (b.zIndex ?? 0));
}

function getSlideshowAnimStyle(
  anim: SlideAnimationConfig | undefined,
  revealedStep: number | undefined
): CSSProperties {
  if (
    typeof revealedStep !== 'number' ||
    !anim ||
    anim.type === 'none'
  ) {
    return {};
  }
  const order = Math.max(1, anim.order ?? 1);
  if (revealedStep < order) {
    return {
      opacity: 0,
      pointerEvents: 'none',
    };
  }
  if (anim.type === 'fadeIn') {
    return { animation: 'liteObjFadeIn 320ms ease-out both' };
  }
  if (anim.type === 'flyIn') {
    return { animation: 'liteObjFlyIn 360ms cubic-bezier(0.22, 1, 0.36, 1) both' };
  }
  if (anim.type === 'zoomIn') {
    return { animation: 'liteObjZoomIn 320ms cubic-bezier(0.22, 1, 0.36, 1) both' };
  }
  return {};
}

function renderCroppedSlideImage(
  imageItem: SlideImageItem,
  altText: string,
  extraClassName = ''
) {
  const crop = imageItem.cropRect;
  const hasCrop = Boolean(
    crop && (crop.l > 0 || crop.t > 0 || crop.r > 0 || crop.b > 0)
  );
  if (hasCrop && crop) {
    const visW = Math.max(5, 100 - crop.l - crop.r);
    const visH = Math.max(5, 100 - crop.t - crop.b);
    return (
      <img
        src={imageItem.src}
        alt={altText}
        draggable={false}
        style={{
          position: 'absolute',
          left: `${(-crop.l / visW) * 100}%`,
          top: `${(-crop.t / visH) * 100}%`,
          width: `${(100 / visW) * 100}%`,
          height: `${(100 / visH) * 100}%`,
          maxWidth: 'none',
          objectFit: 'fill',
        }}
        className={`pointer-events-none select-none ${extraClassName}`}
      />
    );
  }
  return (
    <img
      src={imageItem.src}
      alt={altText}
      draggable={false}
      className={`w-full h-full object-contain pointer-events-none select-none ${extraClassName}`}
    />
  );
}

function getBlockPaddingStyle(
  paddingPct: SlideTextBlock['paddingPct'],
  fontScale: number
): string | undefined {
  if (!paddingPct) {
    return undefined;
  }
  const t = Math.max(1, paddingPct.top * fontScale);
  const r = Math.max(2, paddingPct.right * fontScale);
  const b = Math.max(1, paddingPct.bottom * fontScale);
  const l = Math.max(2, paddingPct.left * fontScale);
  return `${t.toFixed(1)}px ${r.toFixed(1)}px ${b.toFixed(1)}px ${l.toFixed(1)}px`;
}

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


function RichTextFittedView({
  block,
  fontScale,
  textColor,
  underlined,
  onActivateEdit,
}: {
  block: SlideTextBlock;
  fontScale: number;
  textColor: string;
  underlined?: boolean;
  onActivateEdit?: () => void;
}) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [overflowScale, setOverflowScale] = useState<number>(1);
  const baseScalePct = (block.fontScalePct ?? 100) / 100;

  useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    setOverflowScale(1);
    const raf = window.requestAnimationFrame(() => {
      if (!containerRef.current) return;
      const node = containerRef.current;
      if (node.clientHeight > 8 && node.scrollHeight > node.clientHeight + 2) {
        const ratio = Math.max(0.68, Math.min(1, node.clientHeight / node.scrollHeight));
        setOverflowScale(ratio);
      }
    });
    return () => window.cancelAnimationFrame(raf);
  }, [block.text, block.richParagraphs, fontScale, block.wPct, block.hPct, block.fontSize]);

  const effectiveScale = fontScale * baseScalePct * overflowScale;
  const paragraphs = block.richParagraphs;

  if (!paragraphs || paragraphs.length === 0) {
    const scaledPx = Math.max(4, (block.fontSize || 18) * effectiveScale);
    return (
      <div
        ref={containerRef}
        onClick={(e) => {
          if (onActivateEdit) {
            e.stopPropagation();
            onActivateEdit();
          }
        }}
        className="w-full h-full overflow-hidden flex flex-col justify-center"
      >
        <p
          className="whitespace-pre-wrap break-words leading-[1.18] tracking-tight"
          style={{
            fontSize: `${scaledPx.toFixed(2)}px`,
            fontWeight: block.bold || block.role === "title" ? 700 : 400,
            fontStyle: block.italic ? "italic" : "normal",
            textDecoration: underlined ? "underline" : "none",
            color: textColor,
            textAlign: block.align ?? "left",
            fontFamily: block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
          }}
        >
          {block.text}
        </p>
      </div>
    );
  }

  return (
    <div
      ref={containerRef}
      onClick={(e) => {
        if (onActivateEdit) {
          e.stopPropagation();
          onActivateEdit();
        }
      }}
      className="w-full max-h-full overflow-hidden space-y-[0.22em]"
    >
      {paragraphs.map((para, pIdx) => {
        const firstRun = para.runs[0];
        const bulletPx = Math.max(4, ((firstRun?.fontSize ?? block.fontSize) || 16) * effectiveScale);
        return (
          <p
            key={`${block.id}-rp-${pIdx}`}
            className="whitespace-pre-wrap break-words leading-[1.18] tracking-tight"
            style={{
              textAlign: para.align ?? block.align ?? "left",
              paddingLeft: para.level ? `${para.level * 1.1}em` : undefined,
            }}
          >
            {para.bullet && (
              <span
                style={{
                  fontSize: `${bulletPx.toFixed(2)}px`,
                  color: firstRun?.color ?? textColor,
                  fontWeight: firstRun?.bold ? 700 : 400,
                  marginRight: "0.35em",
                }}
              >
                {para.bullet}
              </span>
            )}
            {para.runs.map((run, rIdx) => {
              const rPx = Math.max(4, ((run.fontSize ?? block.fontSize) || 16) * effectiveScale);
              return (
                <span
                  key={`${block.id}-rp-${pIdx}-r-${rIdx}`}
                  style={{
                    fontSize: `${rPx.toFixed(2)}px`,
                    fontWeight: (run.bold ?? block.bold ?? block.role === "title") ? 700 : 400,
                    fontStyle: (run.italic ?? block.italic) ? "italic" : "normal",
                    textDecoration: (run.underline || underlined) ? "underline" : "none",
                    color: run.color ?? textColor,
                    fontFamily: run.fontFamily || block.fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
                  }}
                >
                  {run.text}
                </span>
              );
            })}
          </p>
        );
      })}
    </div>
  );
}

type ResizeHandleDir = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

function PowerPointResizeHandles({
  onStartResize,
  onStartRotate,
}: {
  onStartResize: (e: ReactPointerEvent<HTMLButtonElement>, dir: ResizeHandleDir) => void;
  onStartRotate?: (e: ReactPointerEvent<HTMLButtonElement>) => void;
}) {
  const handles: Array<{ dir: ResizeHandleDir; cls: string; cursor: string }> = [
    { dir: "nw", cls: "-top-1.5 -left-1.5", cursor: "nwse-resize" },
    { dir: "n", cls: "-top-1.5 left-1/2 -translate-x-1/2", cursor: "ns-resize" },
    { dir: "ne", cls: "-top-1.5 -right-1.5", cursor: "nesw-resize" },
    { dir: "e", cls: "top-1/2 -translate-y-1/2 -right-1.5", cursor: "ew-resize" },
    { dir: "se", cls: "-bottom-1.5 -right-1.5", cursor: "nwse-resize" },
    { dir: "s", cls: "-bottom-1.5 left-1/2 -translate-x-1/2", cursor: "ns-resize" },
    { dir: "sw", cls: "-bottom-1.5 -left-1.5", cursor: "nesw-resize" },
    { dir: "w", cls: "top-1/2 -translate-y-1/2 -left-1.5", cursor: "ew-resize" },
  ];

  return (
    <>
      {onStartRotate && (
        <>
          <span
            aria-hidden="true"
            className="pointer-events-none absolute -top-5 left-1/2 -translate-x-1/2 w-px h-3.5 bg-[#C43E1C]"
          />
          <button
            type="button"
            onPointerDown={(e) => {
              e.stopPropagation();
              onStartRotate(e);
            }}
            aria-label="Xoay đối tượng"
            title="Kéo để xoay góc đối tượng"
            className="absolute -top-7 left-1/2 -translate-x-1/2 z-30 h-4 w-4 rounded-full bg-[#C43E1C] text-white border border-white shadow-md flex items-center justify-center cursor-grab active:cursor-grabbing hover:scale-110 transition-transform"
          >
            <RotateCw className="h-2.5 w-2.5" />
          </button>
        </>
      )}
      {handles.map((h) => (
        <button
          key={h.dir}
          type="button"
          onPointerDown={(e) => {
            e.stopPropagation();
            onStartResize(e, h.dir);
          }}
          aria-label={`Thay đổi kích thước (${h.dir})`}
          style={{ cursor: h.cursor }}
          className={`absolute ${h.cls} z-30 h-2.5 w-2.5 rounded-full bg-white border-2 border-[#C43E1C] shadow-sm hover:scale-125 transition-transform`}
        />
      ))}
    </>
  );
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
  revealedAnimStep?: number;
}

function ReadOnlySlideStage({
  doc,
  slide,
  slideIndex,
  totalSlides,
  stageWidthPx,
  underlinedBlockIds,
  revealedAnimStep,
}: PositionedReadOnlySlideProps) {
  const fontScale =
    stageWidthPx > 0 ? stageWidthPx / REFERENCE_SLIDE_WIDTH_PT : 1;
  const isPositioned = isSlideMasterPositioned(slide);
  const hasLightBg = isLightHexBackground(slide.backgroundColor);
  const accent = slide.accentColor || '#4F46E5';
  const blocks = sortBlocksByZIndex(slide.blocks ?? []);
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
            3.5,
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
          const blockPadding = getBlockPaddingStyle(block.paddingPct, fontScale);
          const animStyle = getSlideshowAnimStyle(block.animation, revealedAnimStep);
          const safeLink = sanitizeExternalUrl(block.linkUrl);

          if (block.kind === 'shape') {
            const shapeKind = block.shapeKind ?? 'rect';
            return (
              <div
                key={block.id}
                style={{
                  position: 'absolute',
                  left: `${left}%`,
                  top: `${top}%`,
                  width: `${width}%`,
                  height: height !== undefined ? `${height}%` : '18%',
                  zIndex: (block.zIndex ?? 0) + 1,
                  transform: block.rotationDeg
                    ? `rotate(${block.rotationDeg}deg)`
                    : undefined,
                  ...animStyle,
                }}
                className="relative flex items-center justify-center"
              >
                <SlideShapeSvg
                  shapeKind={shapeKind}
                  fillColor={block.fillColor}
                  borderColor={block.borderColor}
                  borderWidthPx={block.borderWidthPx}
                />
                {shapeKind !== 'line' && block.text && (
                  <span
                    className="relative z-10 px-2 text-center font-medium leading-tight pointer-events-none"
                    style={{
                      fontSize: `${scaledFontPx.toFixed(2)}px`,
                      color: textColor,
                      fontWeight: block.bold ? 700 : 500,
                    }}
                  >
                    {block.text}
                  </span>
                )}
              </div>
            );
          }

          if (block.kind === 'chart' && block.chartData) {
            return (
              <div
                key={block.id}
                style={{
                  position: 'absolute',
                  left: `${left}%`,
                  top: `${top}%`,
                  width: `${width}%`,
                  height: height !== undefined ? `${height}%` : '44%',
                  zIndex: (block.zIndex ?? 0) + 1,
                  transform: block.rotationDeg
                    ? `rotate(${block.rotationDeg}deg)`
                    : undefined,
                  ...animStyle,
                }}
              >
                <SlideChartSvg
                  chartData={block.chartData}
                  hasLightBg={hasLightBg}
                  fontScale={fontScale}
                />
              </div>
            );
          }

          if (block.kind === 'media' && block.mediaSrc) {
            return (
              <div
                key={block.id}
                onClick={(e) => e.stopPropagation()}
                onPointerDown={(e) => e.stopPropagation()}
                style={{
                  position: 'absolute',
                  left: `${left}%`,
                  top: `${top}%`,
                  width: `${width}%`,
                  height: height !== undefined ? `${height}%` : '34%',
                  zIndex: (block.zIndex ?? 0) + 1,
                  transform: block.rotationDeg
                    ? `rotate(${block.rotationDeg}deg)`
                    : undefined,
                  ...animStyle,
                }}
                className="rounded-lg overflow-hidden bg-slate-950/90 border border-white/20 flex flex-col justify-center p-1"
              >
                {block.mediaKind === 'audio' ? (
                  <div className="w-full px-1.5 flex flex-col gap-1">
                    {block.text && (
                      <span
                        className="truncate font-medium text-slate-200"
                        style={{ fontSize: `${Math.max(7, 11 * fontScale)}px` }}
                      >
                        {block.text}
                      </span>
                    )}
                    <audio controls src={block.mediaSrc} className="w-full h-7" />
                  </div>
                ) : (
                  <video
                    controls
                    playsInline
                    src={block.mediaSrc}
                    className="w-full h-full object-contain rounded"
                  />
                )}
              </div>
            );
          }

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
                  zIndex: (block.zIndex ?? 0) + 1,
                  borderColor: block.borderColor,
                  transform: block.rotationDeg
                    ? `rotate(${block.rotationDeg}deg)`
                    : undefined,
                  fontFamily:
                    block.fontFamily ||
                    "'Inter', 'Segoe UI', system-ui, sans-serif",
                  ...animStyle,
                }}
                className={`overflow-hidden ${block.borderColor ? 'border' : ''}`}
              >
                <table className="w-full h-full border-collapse table-fixed">
                  <tbody>
                    {block.tableRows.map((row, rIdx) => (
                      <tr key={`${block.id}-ro-r-${rIdx}`}>
                        {row.map((cell, cIdx) => {
                          const cellStyle = block.tableCellStyles?.[rIdx]?.[cIdx];
                          return (
                            <td
                              key={`${block.id}-ro-c-${rIdx}-${cIdx}`}
                              className="border px-1 py-0.5 align-middle truncate"
                              style={{
                                borderColor:
                                  block.borderColor ??
                                  (hasLightBg
                                    ? 'rgba(15,23,42,0.22)'
                                    : 'rgba(255,255,255,0.22)'),
                                backgroundColor:
                                  cellStyle?.fillColor ??
                                  (rIdx === 0
                                    ? block.fillColor ??
                                      (hasLightBg
                                        ? 'rgba(15,23,42,0.06)'
                                        : 'rgba(255,255,255,0.08)')
                                    : 'transparent'),
                                fontSize: `${scaledFontPx.toFixed(2)}px`,
                                fontWeight:
                                  (cellStyle?.bold ?? (rIdx === 0 || block.bold))
                                    ? 700
                                    : 400,
                                fontStyle: block.italic ? 'italic' : 'normal',
                                textDecoration: underlinedBlockIds?.[block.id]
                                  ? 'underline'
                                  : 'none',
                                color: cellStyle?.textColor ?? textColor,
                                textAlign:
                                  cellStyle?.align ?? block.align ?? 'left',
                                fontFamily:
                                  block.fontFamily ||
                                  "'Inter', 'Segoe UI', system-ui, sans-serif",
                              }}
                            >
                              {cell}
                            </td>
                          );
                        })}
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
              onClick={
                safeLink
                  ? (e) => {
                      e.stopPropagation();
                      window.open(safeLink, '_blank', 'noopener,noreferrer');
                    }
                  : undefined
              }
              style={{
                position: 'absolute',
                left: `${left}%`,
                top: `${top}%`,
                width: `${width}%`,
                height: height !== undefined ? `${height}%` : 'auto',
                zIndex: (block.zIndex ?? 0) + 1,
                backgroundColor: block.fillColor,
                backgroundImage: block.fillGradient,
                borderColor: block.borderColor,
                borderWidth: block.borderColor
                  ? `${block.borderWidthPx ?? 1}px`
                  : undefined,
                borderStyle: block.borderColor ? 'solid' : undefined,
                borderRadius: block.borderRadius,
                boxShadow: block.shadowCss,
                padding: blockPadding,
                transform: block.rotationDeg
                  ? `rotate(${block.rotationDeg}deg)`
                  : undefined,
                ...animStyle,
              }}
              className={`flex flex-col ${justifyClass} overflow-hidden ${
                blockPadding ? '' : 'px-1.5 py-0.5'
              } ${safeLink ? 'cursor-pointer hover:underline' : ''}`}
            >
              {block.richParagraphs && block.richParagraphs.length > 0 ? (
                <RichTextFittedView
                  block={block}
                  fontScale={fontScale}
                  textColor={safeLink ? '#38BDF8' : textColor}
                  underlined={Boolean(underlinedBlockIds?.[block.id] || safeLink)}
                />
              ) : (
                <p
                  className="whitespace-pre-wrap break-words leading-[1.16] tracking-tight"
                  style={{
                    fontSize: `${Math.max(
                      3.5,
                      scaledFontPx * ((block.fontScalePct ?? 100) / 100)
                    ).toFixed(2)}px`,
                    fontWeight:
                      block.bold || block.role === 'title' ? 700 : 400,
                    fontStyle: block.italic ? 'italic' : 'normal',
                    textDecoration:
                      underlinedBlockIds?.[block.id] || safeLink
                        ? 'underline'
                        : 'none',
                    color: safeLink ? '#38BDF8' : textColor,
                    textAlign: block.align ?? 'left',
                    fontFamily:
                      block.fontFamily ||
                      "'Inter', 'Segoe UI', system-ui, sans-serif",
                  }}
                >
                  {block.text}
                </p>
              )}
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
              borderRadius: imageItem.borderRadius,
              transform: imageItem.rotationDeg
                ? `rotate(${imageItem.rotationDeg}deg)`
                : undefined,
            }}
            className="relative overflow-hidden flex items-center justify-center"
          >
            {renderCroppedSlideImage(
              imageItem,
              imageItem.alt ?? `Hình ${iIdx + 1}`
            )}
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
  collaborators = [],
  selfSessionId,
  collabPermission = 'owner',
  onPresenceUpdate,
  onBroadcastPatch,
  incomingRemotePatch = null,
}: SlideWorkspaceProps) {
  const isReadOnly = collabPermission === 'view';
  const { onDownloadOriginalFile } = useWorkspaceSync();
  const [viewMode, setViewMode] = useState<SlideWorkspaceViewMode>(() =>
    collabPermission === 'view' ||
    Boolean(document.rawArrayBuffer || document.isUploaded)
      ? 'native'
      : 'edit'
  );
  const [hasLocalEdits, setHasLocalEdits] = useState<boolean>(() =>
    Boolean(document.hasLocalEdits)
  );
  const [slides, setSlides] = useState<SlidePreviewItem[]>(() =>
    initializeDeckSlides(document)
  );
  const [activeSlideIdx, setActiveSlideIdx] = useState<number>(0);
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null);
  const [editingBlockId, setEditingBlockId] = useState<string | null>(null);
  const [zoomPct, setZoomPct] = useState<number>(100);
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
  const [isDownloadingOriginal, setIsDownloadingOriginal] =
    useState<boolean>(false);
  const [downloadOriginalNotice, setDownloadOriginalNotice] = useState<
    string | null
  >(null);
  const [exportError, setExportError] = useState<string | null>(null);

  // ResizeObserver measured sizes
  const [editorViewportSize, setEditorViewportSize] = useState<{
    width: number;
    height: number;
  }>({ width: 960, height: 540 });
  const [measuredEditorCanvasWidth, setMeasuredEditorCanvasWidth] =
    useState<number>(0);
  const [slideshowViewportSize, setSlideshowViewportSize] = useState<{
    width: number;
    height: number;
  }>({ width: 1280, height: 720 });
  const [measuredSlideshowCanvasWidth, setMeasuredSlideshowCanvasWidth] =
    useState<number>(0);

  // Fullscreen Slideshow & PowerPoint Pro Ribbon State
  const [activeRibbonTab, setActiveRibbonTab] = useState<PptRibbonTab>('home');
  const [pptViewMode, setPptViewMode] = useState<'normal' | 'sorter'>('normal');
  const [selectedBlockIds, setSelectedBlockIds] = useState<string[]>([]);
  const [smartGuides, setSmartGuides] = useState<SmartGuidesState>({
    xPct: null,
    yPct: null,
  });
  const [thumbContextMenu, setThumbContextMenu] = useState<{
    x: number;
    y: number;
    slideIdx: number;
  } | null>(null);
  const [draggedSlideIdx, setDraggedSlideIdx] = useState<number | null>(null);
  const [dragOverSlideIdx, setDragOverSlideIdx] = useState<number | null>(null);
  const [notesHeightPx, setNotesHeightPx] = useState<number>(72);
  const [isLayoutPickerOpen, setIsLayoutPickerOpen] = useState<boolean>(false);
  const [isShapePickerOpen, setIsShapePickerOpen] = useState<boolean>(false);
  const [isChartModalOpen, setIsChartModalOpen] = useState<boolean>(false);
  const [linkMediaModalMode, setLinkMediaModalMode] = useState<
    'link' | 'video' | 'audio' | null
  >(null);

  const [isSlideshowActive, setIsSlideshowActive] = useState<boolean>(false);
  const [slideshowTransition] = useState<SlideshowTransition>('slide');
  const [slideshowDirection, setSlideshowDirection] = useState<1 | -1>(1);
  const [revealedAnimStep, setRevealedAnimStep] = useState<number>(0);
  const [isSlideshowLaser, setIsSlideshowLaser] = useState<boolean>(false);
  const [slideshowLaserPos, setSlideshowLaserPos] = useState<LaserPosition>({
    x: 50,
    y: 50,
    visible: false,
  });
  const [slideshowDrawTool, setSlideshowDrawTool] =
    useState<SlideshowDrawTool>('none');
  const [slideshowStrokesBySlide, setSlideshowStrokesBySlide] = useState<
    Record<string, SlideshowStroke[]>
  >({});
  const activeSlideshowStrokeIdRef = useRef<string | null>(null);
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
  const bgImageInputRef = useRef<HTMLInputElement | null>(null);
  const editorViewportRef = useRef<HTMLDivElement | null>(null);
  const editorCanvasRef = useRef<HTMLDivElement | null>(null);
  const slideshowViewportRef = useRef<HTMLDivElement | null>(null);
  const slideshowCanvasRef = useRef<HTMLDivElement | null>(null);
  const slideshowTouchStartRef = useRef<{ x: number; y: number } | null>(null);
  const justSwipedSlideshowRef = useRef<boolean>(false);
  const downloadNoticeTimerRef = useRef<number | null>(null);
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
      setViewMode(
        collabPermission === 'view' ||
          Boolean(document.rawArrayBuffer || document.isUploaded)
          ? 'native'
          : 'edit'
      );
      setHasLocalEdits(Boolean(document.hasLocalEdits));
      setActiveSlideIdx(0);
      setSelectedBlockId(null);
      setSelectedBlockIds([]);
      setEditingBlockId(null);
      setSelectedImageId(null);
      setNewlyAddedBlockId(null);
      setSmartGuides({ xPct: null, yPct: null });
      setThumbContextMenu(null);
      setPptViewMode('normal');
      setRevealedAnimStep(0);
      setSlideshowDrawTool('none');
      setSlideshowStrokesBySlide({});
      setDownloadOriginalNotice(null);
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

  // Clear selection/mobile sheet and force native view if switched to read-only permission
  useEffect(() => {
    if (isReadOnly) {
      setViewMode('native');
      setSelectedBlockId(null);
      setEditingBlockId(null);
      setSelectedImageId(null);
      setIsMobileSheetOpen(false);
    }
  }, [isReadOnly]);

  // Pinch / Ctrl+Wheel zoom strictly scoped to the slide viewport (25% - 200%)
  useDocumentPinchZoom(editorViewportRef, zoomPct, setZoomPct, {
    min: 25,
    max: 200,
    disabled:
      isSlideshowActive ||
      isModalOpen ||
      isChartModalOpen ||
      linkMediaModalMode !== null ||
      pptViewMode === 'sorter',
  });

  // Close thumbnail context menu on outside click / Escape
  useEffect(() => {
    if (!thumbContextMenu) {
      return;
    }
    const closeMenu = () => setThumbContextMenu(null);
    window.addEventListener('click', closeMenu);
    window.addEventListener('contextmenu', closeMenu);
    return () => {
      window.removeEventListener('click', closeMenu);
      window.removeEventListener('contextmenu', closeMenu);
    };
  }, [thumbContextMenu]);

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

  // Measure actual slide canvas width with ResizeObserver for exact fontScale = width / 960
  useEffect(() => {
    const el = editorCanvasRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateCanvasWidth = () => {
      const w = el.clientWidth;
      if (w > 0) {
        setMeasuredEditorCanvasWidth(w);
      }
    };
    updateCanvasWidth();
    const observer = new ResizeObserver(updateCanvasWidth);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, [zoomPct, activeSlideIdx, isHydratingCloudFile]);

  // Measure slideshow viewport and slideshow canvas with ResizeObserver when slideshow is active
  useEffect(() => {
    if (!isSlideshowActive) {
      setMeasuredSlideshowCanvasWidth(0);
      return;
    }
    const el = slideshowViewportRef.current;
    const canvasEl = slideshowCanvasRef.current;
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
      if (canvasEl && canvasEl.clientWidth > 0) {
        setMeasuredSlideshowCanvasWidth(canvasEl.clientWidth);
      }
    };
    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(el);
    if (canvasEl) {
      observer.observe(canvasEl);
    }
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
      if (downloadNoticeTimerRef.current !== null) {
        window.clearTimeout(downloadNoticeTimerRef.current);
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
  const zoomedWidthPx = editorFitWidthPx * (zoomPct / 100);
  const zoomedHeightPx = (editorFitWidthPx / currentAspectRatio) * (zoomPct / 100);
  const actualEditorCanvasWidth =
    measuredEditorCanvasWidth > 0 ? measuredEditorCanvasWidth : zoomedWidthPx;
  const editorFontScale = actualEditorCanvasWidth / REFERENCE_SLIDE_WIDTH_PT;

  const slideshowFitWidthPx = Math.max(
    240,
    Math.min(
      slideshowViewportSize.width,
      slideshowViewportSize.height * currentAspectRatio
    )
  );
  const slideshowFitHeightPx = slideshowFitWidthPx / currentAspectRatio;
  const actualSlideshowStageWidth =
    measuredSlideshowCanvasWidth > 0
      ? measuredSlideshowCanvasWidth
      : slideshowFitWidthPx;

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
      hasLocalEdits: hasLocalEdits || Boolean(baseDoc.hasLocalEdits),
      contentPreview:
        syncedSlides[0]?.bullets[0] ??
        syncedSlides[0]?.title ??
        baseDoc.contentPreview,
    };
  }, [hasLocalEdits, slides]);

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

  const getMaxAnimationStep = useCallback((slide: SlidePreviewItem | undefined): number => {
    if (!slide?.blocks) {
      return 0;
    }
    let maxOrder = 0;
    for (const blk of slide.blocks) {
      if (blk.animation && blk.animation.type !== 'none') {
        maxOrder = Math.max(maxOrder, blk.animation.order ?? 1);
      }
    }
    return maxOrder;
  }, []);

  const findNextVisibleSlideIdx = useCallback(
    (fromIdx: number, direction: 1 | -1): number => {
      let cursor = fromIdx + direction;
      while (cursor >= 0 && cursor < slides.length) {
        if (!slides[cursor]?.hidden) {
          return cursor;
        }
        cursor += direction;
      }
      return fromIdx;
    },
    [slides]
  );

  const startSlideshow = useCallback(
    (startFromIdx?: number) => {
      const requestedIdx =
        typeof startFromIdx === 'number'
          ? Math.max(0, Math.min(slides.length - 1, startFromIdx))
          : 0;
      let targetIdx = requestedIdx;
      if (typeof startFromIdx !== 'number' && slides[targetIdx]?.hidden) {
        const firstVisible = slides.findIndex((s) => !s.hidden);
        if (firstVisible >= 0) {
          targetIdx = firstVisible;
        }
      }
      setActiveSlideIdx(targetIdx);
      setSelectedBlockId(null);
      setSelectedBlockIds([]);
      setSelectedImageId(null);
      setIsLaserActive(false);
      setIsSlideshowLaser(false);
      setSlideshowDrawTool('none');
      setRevealedAnimStep(0);
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
    },
    [slides]
  );

  const exitSlideshow = useCallback(() => {
    setIsSlideshowGridOpen(false);
    setIsSlideshowLaser(false);
    setSlideshowDrawTool('none');
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

  const advanceSlideshowStep = useCallback(() => {
    const activeSlide = slides[safeSlideIdx];
    const maxStep = getMaxAnimationStep(activeSlide);
    if (revealedAnimStep < maxStep) {
      setRevealedAnimStep((prev) => prev + 1);
      return;
    }
    const nextIdx = findNextVisibleSlideIdx(safeSlideIdx, 1);
    if (nextIdx !== safeSlideIdx) {
      setSlideshowDirection(1);
      setRevealedAnimStep(0);
      setActiveSlideIdx(nextIdx);
    }
  }, [
    findNextVisibleSlideIdx,
    getMaxAnimationStep,
    revealedAnimStep,
    safeSlideIdx,
    slides,
  ]);

  const retreatSlideshowStep = useCallback(() => {
    if (revealedAnimStep > 0) {
      setRevealedAnimStep((prev) => Math.max(0, prev - 1));
      return;
    }
    const prevIdx = findNextVisibleSlideIdx(safeSlideIdx, -1);
    if (prevIdx !== safeSlideIdx) {
      setSlideshowDirection(-1);
      setRevealedAnimStep(0);
      setActiveSlideIdx(prevIdx);
    }
  }, [findNextVisibleSlideIdx, revealedAnimStep, safeSlideIdx]);

  // Keyboard navigation + Slideshow shortcuts + Ctrl/Cmd+S save
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (isModalOpen || isChartModalOpen || linkMediaModalMode !== null) {
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
          advanceSlideshowStep();
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
          event.preventDefault();
          retreatSlideshowStep();
        }
        return;
      }

      if (event.key === 'F5' && !isSlideshowActive) {
        event.preventDefault();
        startSlideshow(event.shiftKey ? safeSlideIdx : 0);
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
        if (thumbContextMenu) {
          setThumbContextMenu(null);
          return;
        }
        if (isLayoutPickerOpen || isShapePickerOpen) {
          setIsLayoutPickerOpen(false);
          setIsShapePickerOpen(false);
          return;
        }
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
        if (
          selectedBlockId !== null ||
          selectedBlockIds.length > 0 ||
          selectedImageId !== null
        ) {
          setSelectedBlockId(null);
          setSelectedBlockIds([]);
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
        setSelectedBlockIds([]);
        setSelectedImageId(null);
        setActiveSlideIdx((prev) =>
          prev < slides.length - 1 ? prev + 1 : prev
        );
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setSelectedBlockId(null);
        setSelectedBlockIds([]);
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
    safeSlideIdx,
    isMobileSheetOpen,
    isMobileDrawerOpen,
    isLaserActive,
    isModalOpen,
    isChartModalOpen,
    linkMediaModalMode,
    isLayoutPickerOpen,
    isShapePickerOpen,
    thumbContextMenu,
    isSlideshowActive,
    isSlideshowGridOpen,
    selectedBlockId,
    selectedBlockIds.length,
    selectedImageId,
    isReadOnly,
    startSlideshow,
    advanceSlideshowStep,
    retreatSlideshowStep,
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
      setHasLocalEdits(true);
      setSaveFeedback(false);
      notifyTypingActivity(true);
    },
    [isReadOnly, notifyTypingActivity]
  );

  const handleAddSlide = () => {
    if (isReadOnly) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
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
    if (isReadOnly) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
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
    if (isReadOnly || slides.length <= 1) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
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
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
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
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
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
    if (isReadOnly) {
      return;
    }
    setUnderlinedBlockIds((prev) => ({
      ...prev,
      [blockId]: !prev[blockId],
    }));
    setIsDirty(true);
    setHasLocalEdits(true);
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

  const handleDuplicateSlideAt = (targetIdx: number) => {
    if (isReadOnly) {
      return;
    }
    const targetSlide = slides[targetIdx];
    if (!targetSlide) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    uidRef.current += 1;
    const cloneId = `${document.id}-dup-slide-${Date.now()}-${uidRef.current}`;
    const clonedSlide: SlidePreviewItem = {
      ...targetSlide,
      id: cloneId,
      blocks: (targetSlide.blocks ?? []).map((block, bIdx) => ({
        ...block,
        id: `${cloneId}-blk-${bIdx + 1}`,
        tableRows: block.tableRows
          ? block.tableRows.map((row) => [...row])
          : undefined,
      })),
      images: targetSlide.images
        ? targetSlide.images.map((img, iIdx) => ({
            ...img,
            id: `${cloneId}-img-${iIdx + 1}`,
          }))
        : undefined,
    };

    setSlides((prev) => {
      const insertAt = Math.min(targetIdx + 1, prev.length);
      const next = [
        ...prev.slice(0, insertAt),
        clonedSlide,
        ...prev.slice(insertAt),
      ];
      return next.map((item, idx) => syncSlideSummary(item, idx));
    });
    setActiveSlideIdx(targetIdx + 1);
    setSelectedBlockId(null);
    setSelectedBlockIds([]);
    setSelectedImageId(null);
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleDeleteSlideAt = (targetIdx: number) => {
    if (isReadOnly || slides.length <= 1) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) => {
      const filtered = prev.filter((_, idx) => idx !== targetIdx);
      return filtered.map((item, idx) => ({ ...item, index: idx + 1 }));
    });
    setActiveSlideIdx((prev) => {
      if (prev > targetIdx) {
        return prev - 1;
      }
      return Math.max(0, Math.min(prev, slides.length - 2));
    });
    setSelectedBlockId(null);
    setSelectedBlockIds([]);
    setSelectedImageId(null);
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleToggleHideSlide = (targetIdx: number) => {
    if (isReadOnly) {
      return;
    }
    mutateSlideAt(targetIdx, (slide) => ({
      ...slide,
      hidden: !slide.hidden,
    }));
  };

  const handleReorderSlide = (fromIdx: number, toIdx: number) => {
    if (isReadOnly || fromIdx === toIdx) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    const activeId = currentSlide?.id;
    setSlides((prev) => {
      if (
        fromIdx < 0 ||
        toIdx < 0 ||
        fromIdx >= prev.length ||
        toIdx >= prev.length
      ) {
        return prev;
      }
      const next = [...prev];
      const [moved] = next.splice(fromIdx, 1);
      next.splice(toIdx, 0, moved);
      const reindexed = next.map((s, i) => ({ ...s, index: i + 1 }));
      if (activeId) {
        const nextActive = reindexed.findIndex((s) => s.id === activeId);
        if (nextActive >= 0) {
          setActiveSlideIdx(nextActive);
        }
      }
      return reindexed;
    });
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleApplySlideLayout = (layoutId: PptSlideLayoutId) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    setIsLayoutPickerOpen(false);
    const isLight = isLightHexBackground(currentSlide.backgroundColor);
    const newBlocks = createBlocksForSlideLayout(
      layoutId,
      `${currentSlide.id}-${Date.now()}`,
      {
        titleColor: isLight ? '#0F172A' : '#FFFFFF',
        bodyColor: isLight ? '#334155' : '#E2E8F0',
        accentColor,
      }
    );
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: newBlocks,
    }));
    setSelectedBlockId(newBlocks[0]?.id ?? null);
    setSelectedBlockIds(newBlocks[0] ? [newBlocks[0].id] : []);
    setSelectedImageId(null);
  };

  const handleInsertShape = (shapeKind: PptShapeKind) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    setIsShapePickerOpen(false);
    uidRef.current += 1;
    const shapeId = `${currentSlide.id}-shape-${Date.now()}-${uidRef.current}`;
    const offset = (currentBlocks.length % 5) * 4;
    const isLine = shapeKind === 'line';
    const newShapeBlock: SlideTextBlock = {
      id: shapeId,
      kind: 'shape',
      shapeKind,
      text: isLine ? '' : 'Nhấn đúp nhập chữ',
      fontSize: 16,
      bold: false,
      color: '#FFFFFF',
      align: 'center',
      fillColor: isLine ? undefined : '#C43E1C',
      borderColor: isLine ? '#F59E0B' : '#FFFFFF',
      borderWidthPx: isLine ? 4 : 2,
      xPct: Math.min(62, 24 + offset),
      yPct: Math.min(68, 28 + offset),
      wPct: isLine ? 36 : 24,
      hPct: isLine ? 6 : 22,
      zIndex: currentBlocks.length + 1,
    };
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: [...(slide.blocks ?? []), newShapeBlock],
    }));
    setSelectedImageId(null);
    setSelectedBlockId(shapeId);
    setSelectedBlockIds([shapeId]);
  };

  const handleInsertTable = (rows = 3, cols = 3) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    uidRef.current += 1;
    const tblId = `${currentSlide.id}-tbl-${Date.now()}-${uidRef.current}`;
    const tableRows: string[][] = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) =>
        r === 0 ? `Cột ${c + 1}` : `Dữ liệu ${r}.${c + 1}`
      )
    );
    const newTblBlock: SlideTextBlock = {
      id: tblId,
      kind: 'table',
      tableRows,
      text: tableRows.map((r) => r.join(' | ')).join('\n'),
      fontSize: 13,
      color: hasLightSlideBg ? '#0F172A' : '#FFFFFF',
      fillColor: hasLightSlideBg ? 'rgba(196,62,28,0.12)' : 'rgba(196,62,28,0.28)',
      borderColor: hasLightSlideBg ? 'rgba(15,23,42,0.25)' : 'rgba(255,255,255,0.25)',
      xPct: 12,
      yPct: 26,
      wPct: 76,
      hPct: 44,
      zIndex: currentBlocks.length + 1,
    };
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: [...(slide.blocks ?? []), newTblBlock],
    }));
    setSelectedImageId(null);
    setSelectedBlockId(tblId);
    setSelectedBlockIds([tblId]);
  };

  const handleInsertChart = (chartData: SlideChartData) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    setIsChartModalOpen(false);
    uidRef.current += 1;
    const chartId = `${currentSlide.id}-chart-${Date.now()}-${uidRef.current}`;
    const newChartBlock: SlideTextBlock = {
      id: chartId,
      kind: 'chart',
      chartData,
      text: chartData.title ?? 'Biểu đồ dữ liệu',
      fontSize: 14,
      xPct: 16,
      yPct: 22,
      wPct: 68,
      hPct: 54,
      zIndex: currentBlocks.length + 1,
    };
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: [...(slide.blocks ?? []), newChartBlock],
    }));
    setSelectedImageId(null);
    setSelectedBlockId(chartId);
    setSelectedBlockIds([chartId]);
  };

  const handleInsertLinkBlock = (label: string, url: string) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    setLinkMediaModalMode(null);
    uidRef.current += 1;
    const linkId = `${currentSlide.id}-link-${Date.now()}-${uidRef.current}`;
    const newLinkBlock: SlideTextBlock = {
      id: linkId,
      kind: 'text',
      text: label || url,
      linkUrl: url,
      fontSize: 16,
      bold: true,
      color: '#38BDF8',
      align: 'left',
      xPct: 12,
      yPct: 74,
      wPct: 48,
      hPct: 10,
      zIndex: currentBlocks.length + 1,
    };
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: [...(slide.blocks ?? []), newLinkBlock],
    }));
    setSelectedImageId(null);
    setSelectedBlockId(linkId);
    setSelectedBlockIds([linkId]);
  };

  const handleInsertMediaBlock = (
    mediaKind: 'video' | 'audio',
    src: string,
    label: string
  ) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    setLinkMediaModalMode(null);
    uidRef.current += 1;
    const mediaId = `${currentSlide.id}-media-${Date.now()}-${uidRef.current}`;
    const newMediaBlock: SlideTextBlock = {
      id: mediaId,
      kind: 'media',
      mediaKind,
      mediaSrc: src,
      text: label,
      fontSize: 13,
      xPct: 22,
      yPct: mediaKind === 'video' ? 24 : 40,
      wPct: 56,
      hPct: mediaKind === 'video' ? 46 : 16,
      zIndex: currentBlocks.length + 1,
    };
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: [...(slide.blocks ?? []), newMediaBlock],
    }));
    setSelectedImageId(null);
    setSelectedBlockId(mediaId);
    setSelectedBlockIds([mediaId]);
  };

  const handleApplyColorTheme = (theme: PptColorTheme) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      backgroundColor: theme.bgColor,
      backgroundGradient: theme.bgGradient,
      backgroundImageSrc: undefined,
      accentColor: theme.accentColor,
      blocks: (slide.blocks ?? []).map((blk) =>
        blk.kind === 'shape' || blk.kind === 'chart' || blk.kind === 'media'
          ? blk
          : {
              ...blk,
              color:
                blk.role === 'title' ? theme.titleColor : theme.bodyColor,
            }
      ),
    }));
  };

  const handleApplySlideGradient = (gradient: string) => {
    if (isReadOnly) {
      return;
    }
    setViewMode('edit');
    const isLightGrad = gradient.includes('#FFFFFF');
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      backgroundColor: isLightGrad ? '#FFFFFF' : '#161D2E',
      backgroundGradient: gradient,
      backgroundImageSrc: undefined,
    }));
  };

  const handleBackgroundImageFileChange = async (
    event: ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || isReadOnly) {
      return;
    }
    try {
      const src = await readImageFileAsSrc(file);
      mutateSlideAt(safeSlideIdx, (slide) => ({
        ...slide,
        backgroundImageSrc: src,
        backgroundGradient: undefined,
      }));
    } catch {
      setExportError('Không thể đặt ảnh nền cho slide.');
    }
  };

  const handleApplyBackgroundToAll = () => {
    if (isReadOnly) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((slide, idx) =>
        syncSlideSummary(
          {
            ...slide,
            backgroundColor: currentSlide.backgroundColor,
            backgroundGradient: currentSlide.backgroundGradient,
            backgroundImageSrc: currentSlide.backgroundImageSrc,
            accentColor: currentSlide.accentColor,
          },
          idx
        )
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleApplyTransitionToAll = () => {
    if (isReadOnly) {
      return;
    }
    const baseTransition = currentSlide.transition ?? {
      type: 'fade' as const,
      durationMs: 350,
    };
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((slide, idx) =>
        syncSlideSummary(
          {
            ...slide,
            transition: { ...baseTransition },
          },
          idx
        )
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleBringToFront = (blockId: string) => {
    if (isReadOnly) {
      return;
    }
    const maxZ = currentBlocks.reduce(
      (max, b) => Math.max(max, b.zIndex ?? 0),
      0
    );
    handleUpdateBlock(blockId, { zIndex: maxZ + 1 });
  };

  const handleSendToBack = (blockId: string) => {
    if (isReadOnly) {
      return;
    }
    const minZ = currentBlocks.reduce(
      (min, b) => Math.min(min, b.zIndex ?? 0),
      0
    );
    handleUpdateBlock(blockId, { zIndex: minZ - 1 });
  };

  const handleGroupSelected = () => {
    if (isReadOnly || selectedBlockIds.length < 2) {
      return;
    }
    const newGroupId = `grp-${Date.now()}`;
    const idSet = new Set(selectedBlockIds);
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).map((b) =>
        idSet.has(b.id) ? { ...b, groupId: newGroupId } : b
      ),
    }));
  };

  const handleUngroupSelected = () => {
    if (isReadOnly || !selectedBlock?.groupId) {
      return;
    }
    const targetGrp = selectedBlock.groupId;
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).map((b) =>
        b.groupId === targetGrp ? { ...b, groupId: undefined } : b
      ),
    }));
  };

  const handleSpeakerNoteChange = (value: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      speakerNote: value,
    }));
  };

  const handleDownloadOriginal = async () => {
    if (isDownloadingOriginal) {
      return;
    }
    setIsDownloadingOriginal(true);
    setExportError(null);
    setDownloadOriginalNotice(null);
    try {
      if (onDownloadOriginalFile) {
        await onDownloadOriginalFile();
        return;
      }

      const res = await exportOriginalDocumentBinary(
        buildLiveDocument(),
        document.rawArrayBuffer
      );
      setDownloadOriginalNotice(res.message);

      if (downloadNoticeTimerRef.current !== null) {
        window.clearTimeout(downloadNoticeTimerRef.current);
      }
      downloadNoticeTimerRef.current = window.setTimeout(() => {
        setDownloadOriginalNotice(null);
        downloadNoticeTimerRef.current = null;
      }, 4500);
    } catch {
      setExportError('Không thể tải xuống bản gốc. Vui lòng thử lại.');
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
      exportErrorTimerRef.current = window.setTimeout(() => {
        setExportError(null);
        exportErrorTimerRef.current = null;
      }, 3500);
    } finally {
      setIsDownloadingOriginal(false);
    }
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
    if (
      isReadOnly ||
      viewMode === 'native' ||
      isLaserActive ||
      isSlideshowActive ||
      event.button !== 0
    ) {
      return;
    }
    const canvasEl = editorCanvasRef.current;
    if (!canvasEl) {
      return;
    }
    event.stopPropagation();

    if (params.kind === 'block') {
      const targetBlock = currentBlocks.find((b) => b.id === params.id);
      if (event.shiftKey) {
        setSelectedBlockIds((prev) =>
          prev.includes(params.id) ? prev : [...prev, params.id]
        );
      } else if (targetBlock?.groupId) {
        const grpIds = currentBlocks
          .filter((b) => b.groupId === targetBlock.groupId)
          .map((b) => b.id);
        setSelectedBlockIds(grpIds);
      } else {
        setSelectedBlockIds([params.id]);
      }
      setSelectedBlockId(params.id);
      setSelectedImageId(null);
    } else {
      setSelectedImageId(params.id);
      setSelectedBlockId(null);
      setSelectedBlockIds([]);
    }

    const draggedBlock =
      params.kind === 'block'
        ? currentBlocks.find((b) => b.id === params.id)
        : undefined;
    const draggedGroupId = draggedBlock?.groupId;
    const originGroupPositions = new Map<string, { xPct: number; yPct: number }>();
    if (draggedGroupId) {
      for (const blk of currentBlocks) {
        if (blk.groupId === draggedGroupId) {
          originGroupPositions.set(blk.id, {
            xPct: blk.xPct ?? 6,
            yPct: blk.yPct ?? 10,
          });
        }
      }
    }

    const snapXTargets: number[] = [50];
    const snapYTargets: number[] = [50];
    for (const blk of currentBlocks) {
      if (blk.id === params.id || (draggedGroupId && blk.groupId === draggedGroupId)) {
        continue;
      }
      const bx = blk.xPct ?? 6;
      const by = blk.yPct ?? 10;
      const bw = blk.wPct ?? 40;
      const bh = blk.hPct ?? 16;
      snapXTargets.push(bx, bx + bw / 2, bx + bw);
      snapYTargets.push(by, by + bh / 2, by + bh);
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
      let nextX = clampValue(
        dragState.originXPct + (dx / rect.width) * 100,
        0,
        maxX
      );
      let nextY = clampValue(
        dragState.originYPct + (dy / rect.height) * 100,
        0,
        maxY
      );

      // Smart Guides Snap (center & object edges within 1.2%)
      let guideX: number | null = null;
      let guideY: number | null = null;
      const snapTol = 1.2;
      const xAnchors = [0, dragState.widthPct / 2, dragState.widthPct];
      for (const targetX of snapXTargets) {
        for (const offsetX of xAnchors) {
          if (Math.abs(nextX + offsetX - targetX) <= snapTol) {
            nextX = clampValue(targetX - offsetX, 0, maxX);
            guideX = roundPct(targetX);
            break;
          }
        }
        if (guideX !== null) {
          break;
        }
      }
      const yAnchors = [0, dragState.heightPct / 2, dragState.heightPct];
      for (const targetY of snapYTargets) {
        for (const offsetY of yAnchors) {
          if (Math.abs(nextY + offsetY - targetY) <= snapTol) {
            nextY = clampValue(targetY - offsetY, 0, maxY);
            guideY = roundPct(targetY);
            break;
          }
        }
        if (guideY !== null) {
          break;
        }
      }
      setSmartGuides({ xPct: guideX, yPct: guideY });

      const finalX = roundPct(nextX);
      const finalY = roundPct(nextY);
      const deltaX = finalX - dragState.originXPct;
      const deltaY = finalY - dragState.originYPct;

      mutateSlideAt(targetSlideIdx, (slide) => {
        if (dragState.kind === 'block') {
          return {
            ...slide,
            blocks: (slide.blocks ?? []).map((block) => {
              if (block.id === dragState.id) {
                return { ...block, xPct: finalX, yPct: finalY };
              }
              if (draggedGroupId && block.groupId === draggedGroupId) {
                const orig = originGroupPositions.get(block.id);
                if (orig) {
                  return {
                    ...block,
                    xPct: roundPct(clampValue(orig.xPct + deltaX, 0, 95)),
                    yPct: roundPct(clampValue(orig.yPct + deltaY, 0, 95)),
                  };
                }
              }
              return block;
            }),
          };
        }
        return {
          ...slide,
          images: (slide.images ?? []).map((img) =>
            img.id === dragState.id ? { ...img, xPct: finalX, yPct: finalY } : img
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
      setSmartGuides({ xPct: null, yPct: null });
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


  const startElementResize = (
    event: ReactPointerEvent<HTMLButtonElement>,
    dir: ResizeHandleDir,
    params: { kind: "block" | "image"; id: string; xPct: number; yPct: number; wPct: number; hPct: number }
  ) => {
    if (isReadOnly || isLaserActive || isSlideshowActive || event.button !== 0) return;
    event.stopPropagation();
    const startX = event.clientX;
    const startY = event.clientY;
    const targetSlideIdx = safeSlideIdx;

    const onMove = (moveEv: PointerEvent) => {
      const rect = editorCanvasRef.current?.getBoundingClientRect();
      if (!rect || rect.width <= 0 || rect.height <= 0) return;
      const dxPct = ((moveEv.clientX - startX) / rect.width) * 100;
      const dyPct = ((moveEv.clientY - startY) / rect.height) * 100;

      let { xPct, yPct, wPct, hPct } = params;
      if (dir.includes("e")) wPct = roundPct(clampValue(params.wPct + dxPct, 3, 100 - params.xPct));
      if (dir.includes("s")) hPct = roundPct(clampValue(params.hPct + dyPct, 3, 100 - params.yPct));
      if (dir.includes("w")) {
        const nextX = roundPct(clampValue(params.xPct + dxPct, 0, params.xPct + params.wPct - 3));
        wPct = roundPct(params.wPct + (params.xPct - nextX));
        xPct = nextX;
      }
      if (dir.includes("n")) {
        const nextY = roundPct(clampValue(params.yPct + dyPct, 0, params.yPct + params.hPct - 3));
        hPct = roundPct(params.hPct + (params.yPct - nextY));
        yPct = nextY;
      }

      mutateSlideAt(targetSlideIdx, (slide) => {
        if (params.kind === "block") {
          return {
            ...slide,
            blocks: (slide.blocks ?? []).map((b) => (b.id === params.id ? { ...b, xPct, yPct, wPct, hPct } : b)),
          };
        }
        return {
          ...slide,
          images: (slide.images ?? []).map((im) => (im.id === params.id ? { ...im, xPct, yPct, wPct, hPct } : im)),
        };
      });
    };

    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      justDraggedRef.current = true;
      window.setTimeout(() => { justDraggedRef.current = false; }, 0);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  };

  const startElementRotate = (
    event: ReactPointerEvent<HTMLButtonElement>,
    params: {
      kind: 'block' | 'image';
      id: string;
      xPct: number;
      yPct: number;
      wPct: number;
      hPct: number;
    }
  ) => {
    if (isReadOnly || isLaserActive || isSlideshowActive || event.button !== 0) {
      return;
    }
    const canvasRect = editorCanvasRef.current?.getBoundingClientRect();
    if (!canvasRect) {
      return;
    }
    event.stopPropagation();
    const centerX =
      canvasRect.left + ((params.xPct + params.wPct / 2) / 100) * canvasRect.width;
    const centerY =
      canvasRect.top + ((params.yPct + params.hPct / 2) / 100) * canvasRect.height;
    const targetSlideIdx = safeSlideIdx;

    const onMove = (moveEv: PointerEvent) => {
      const rad = Math.atan2(moveEv.clientY - centerY, moveEv.clientX - centerX);
      let deg = Math.round((rad * 180) / Math.PI + 90);
      if (moveEv.shiftKey) {
        deg = Math.round(deg / 15) * 15;
      }
      const normalized = ((deg % 360) + 360) % 360;
      const signedDeg = normalized > 180 ? normalized - 360 : normalized;
      mutateSlideAt(targetSlideIdx, (slide) => {
        if (params.kind === 'block') {
          return {
            ...slide,
            blocks: (slide.blocks ?? []).map((b) =>
              b.id === params.id ? { ...b, rotationDeg: signedDeg } : b
            ),
          };
        }
        return {
          ...slide,
          images: (slide.images ?? []).map((im) =>
            im.id === params.id ? { ...im, rotationDeg: signedDeg } : im
          ),
        };
      });
    };

    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      justDraggedRef.current = true;
      window.setTimeout(() => {
        justDraggedRef.current = false;
      }, 0);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  };

  const startNotesResize = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) {
      return;
    }
    event.preventDefault();
    const startY = event.clientY;
    const originH = notesHeightPx;
    const onMove = (moveEv: PointerEvent) => {
      const delta = startY - moveEv.clientY;
      setNotesHeightPx(clampValue(originH + delta, 44, 220));
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
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
        disabled={isReadOnly}
        title={isReadOnly ? 'Bạn chỉ có quyền xem' : 'Thêm slide mới'}
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-indigo-600/90 hover:bg-indigo-500 text-[11px] font-semibold text-white transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed"
      >
        <Plus className="h-3.5 w-3.5 shrink-0" />
        <span>Thêm</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDuplicateSlide}
        disabled={isReadOnly}
        title={isReadOnly ? 'Bạn chỉ có quyền xem' : 'Nhân bản slide đang chọn'}
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white/[0.07] hover:bg-white/[0.14] border border-white/10 text-[11px] font-medium text-slate-200 transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed"
      >
        <Copy className="h-3 w-3 shrink-0" />
        <span>Nhân bản</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDeleteSlide}
        disabled={isReadOnly || slides.length <= 1}
        title={
          isReadOnly
            ? 'Bạn chỉ có quyền xem'
            : slides.length <= 1
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
        const slideCollaborators = collaborators.filter(
          (c) => c.sessionId !== selfSessionId && c.slideIndex === idx
        );

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
            draggable={!isReadOnly}
            onDragStart={(e) => {
              if (isReadOnly) {
                return;
              }
              setDraggedSlideIdx(idx);
              e.dataTransfer.effectAllowed = 'move';
            }}
            onDragOver={(e) => {
              if (isReadOnly || draggedSlideIdx === null) {
                return;
              }
              e.preventDefault();
              if (dragOverSlideIdx !== idx) {
                setDragOverSlideIdx(idx);
              }
            }}
            onDrop={(e) => {
              e.preventDefault();
              if (!isReadOnly && draggedSlideIdx !== null && draggedSlideIdx !== idx) {
                handleReorderSlide(draggedSlideIdx, idx);
              }
              setDraggedSlideIdx(null);
              setDragOverSlideIdx(null);
            }}
            onDragEnd={() => {
              setDraggedSlideIdx(null);
              setDragOverSlideIdx(null);
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              setThumbContextMenu({
                x: Math.min(e.clientX, window.innerWidth - 210),
                y: Math.min(e.clientY, window.innerHeight - 200),
                slideIdx: idx,
              });
            }}
            onClick={() => handleSelectSlide(idx)}
            aria-current={isActive ? 'true' : undefined}
            aria-label={`Chọn Slide ${idx + 1}: ${slide.title}`}
            style={
              !isActive && slideCollaborators.length > 0
                ? { boxShadow: `0 0 0 1.5px ${slideCollaborators[0].color}99` }
                : undefined
            }
            className={`group w-full text-left flex items-start gap-2 rounded-xl p-1.5 transition-all relative ${
              isActive
                ? 'bg-[#C43E1C]/20 ring-2 ring-[#C43E1C]'
                : dragOverSlideIdx === idx
                  ? 'bg-white/10 ring-2 ring-amber-400/70'
                  : 'hover:bg-white/[0.05] ring-1 ring-transparent'
            } ${slide.hidden ? 'opacity-55' : ''} ${
              draggedSlideIdx === idx ? 'opacity-40 scale-[0.98]' : ''
            }`}
          >
            <div className="w-5 shrink-0 pt-0.5 flex flex-col items-end gap-0.5">
              <span
                className={`text-right text-[11px] font-semibold tabular-nums ${
                  isActive
                    ? 'text-orange-300'
                    : 'text-slate-500 group-hover:text-slate-300'
                }`}
              >
                {idx + 1}
              </span>
              {slide.hidden && (
                <EyeOff
                  className="h-3 w-3 text-slate-400"
                  aria-label="Slide đang bị ẩn"
                />
              )}
            </div>

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

              {slideCollaborators.length > 0 && (
                <div className="mt-1 flex flex-wrap items-center gap-1">
                  {slideCollaborators.slice(0, 3).map((collab) => {
                    const shortName =
                      collab.displayName.trim().split(/\s+/)[0] ||
                      collab.displayName;
                    return (
                      <span
                        key={collab.sessionId}
                        title={`${collab.displayName} đang xem Slide ${idx + 1}`}
                        style={{ borderColor: `${collab.color}88` }}
                        className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full bg-slate-900/90 border text-[9.5px] font-medium text-slate-100 max-w-[92px]"
                      >
                        <span
                          className="h-1.5 w-1.5 rounded-full shrink-0"
                          style={{ backgroundColor: collab.color }}
                        />
                        <span className="truncate">{shortName}</span>
                      </span>
                    );
                  })}
                  {slideCollaborators.length > 3 && (
                    <span className="px-1 py-0.5 rounded-full bg-white/10 text-[9px] font-semibold text-slate-300">
                      +{slideCollaborators.length - 3}
                    </span>
                  )}
                </div>
              )}
            </div>
          </button>
        );
      })}
    </div>
  );

  const hasImages = currentImages.length > 0;
  const hasMediaColumn = hasImages || Boolean(currentSlide.metricHighlight);

  const effectiveTransitionType: SlideTransitionConfig['type'] =
    currentSlide.transition?.type ??
    (slideshowTransition === 'fade' ? 'fade' : 'push');
  const transitionDuration = currentSlide.transition?.durationMs ?? 350;

  const slideshowAnimationStyle: CSSProperties = (() => {
    if (effectiveTransitionType === 'none') {
      return {};
    }
    if (effectiveTransitionType === 'fade') {
      return {
        animation: `liteSlideFadeIn ${transitionDuration}ms cubic-bezier(0.22, 1, 0.36, 1) both`,
      };
    }
    if (effectiveTransitionType === 'wipe') {
      return {
        animation: `liteSlideWipeIn ${transitionDuration}ms cubic-bezier(0.22, 1, 0.36, 1) both`,
      };
    }
    if (effectiveTransitionType === 'split') {
      return {
        animation: `liteSlideSplitIn ${transitionDuration}ms cubic-bezier(0.22, 1, 0.36, 1) both`,
      };
    }
    if (effectiveTransitionType === 'zoom') {
      return {
        animation: `liteSlideZoomIn ${transitionDuration}ms cubic-bezier(0.22, 1, 0.36, 1) both`,
      };
    }
    return {
      animation: `${
        slideshowDirection > 0 ? 'liteSlideFromRight' : 'liteSlideFromLeft'
      } ${transitionDuration}ms cubic-bezier(0.22, 1, 0.36, 1) both`,
    };
  })();

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
      <input
        ref={bgImageInputRef}
        type="file"
        accept="image/*"
        onChange={(event) => void handleBackgroundImageFileChange(event)}
        className="hidden"
        aria-label="Chọn ảnh nền cho slide"
      />

      {/* Shared Top Workspace Header */}
      <WorkspaceTopBar
        document={document}
        collaborators={collaborators}
        selfSessionId={selfSessionId}
        collabPermission={collabPermission}
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
            {/* Mode Switcher: Xem nguyên bản (Giống Microsoft Office) vs Chỉnh sửa nhanh */}
            <div
              role="group"
              aria-label="Chế độ xem và chỉnh sửa slide"
              className="inline-flex items-center rounded-xl bg-[#1C2331] border border-white/10 p-0.5 shrink-0"
            >
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => {
                  setSelectedBlockId(null);
                  setEditingBlockId(null);
                  setSelectedImageId(null);
                  setIsMobileSheetOpen(false);
                  setViewMode('native');
                }}
                aria-pressed={viewMode === 'native'}
                aria-label="Xem nguyên bản (Giống Microsoft Office)"
                title="Xem nguyên bản (Giống Microsoft Office)"
                className={`inline-flex items-center gap-1.5 min-h-[30px] px-2.5 rounded-[10px] text-xs font-semibold transition-colors whitespace-nowrap ${
                  viewMode === 'native'
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'text-slate-300 hover:bg-white/[0.08] hover:text-white'
                }`}
              >
                <Eye className="h-3.5 w-3.5 shrink-0" />
                <span>
                  Xem nguyên bản
                  <span className="hidden 2xl:inline">
                    {' '}
                    (Giống Microsoft Office)
                  </span>
                </span>
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                disabled={isReadOnly}
                onClick={() => {
                  if (!isReadOnly) {
                    setViewMode('edit');
                  }
                }}
                aria-pressed={viewMode === 'edit' && !isReadOnly}
                aria-label="Chỉnh sửa nhanh"
                title={
                  isReadOnly
                    ? 'Bạn chỉ có quyền xem tài liệu này'
                    : 'Chuyển sang chế độ Chỉnh sửa nhanh'
                }
                className={`inline-flex items-center gap-1.5 min-h-[30px] px-2.5 rounded-[10px] text-xs font-semibold transition-colors whitespace-nowrap disabled:opacity-40 disabled:cursor-not-allowed ${
                  viewMode === 'edit' && !isReadOnly
                    ? 'bg-indigo-600 text-white shadow-sm'
                    : 'text-slate-300 hover:bg-white/[0.08] hover:text-white'
                }`}
              >
                <Pencil className="h-3.5 w-3.5 shrink-0" />
                <span>Chỉnh sửa nhanh</span>
              </button>
            </div>

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
              disabled={isReadOnly}
              aria-label="Lưu thay đổi bản trình chiếu"
              title={
                isReadOnly ? 'Bạn chỉ có quyền xem' : 'Lưu thay đổi (Ctrl+S)'
              }
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap disabled:opacity-40 disabled:cursor-not-allowed ${
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
                    {isDirty && !isReadOnly && (
                      <span className="absolute -top-1 -right-1 h-2 w-2 rounded-full bg-amber-400" />
                    )}
                  </span>
                  <span className="hidden md:inline">Lưu thay đổi</span>
                </>
              )}
            </button>

            {/* Download Original File Button */}
            <button
              type="button"
              onClick={() => void handleDownloadOriginal()}
              disabled={isDownloadingOriginal}
              aria-label="Tải xuống bản gốc"
              title="Tải xuống tệp .pptx gốc ban đầu"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 rounded-xl bg-emerald-500/15 hover:bg-emerald-500/25 border border-emerald-400/35 text-xs font-semibold text-emerald-200 transition-colors whitespace-nowrap disabled:opacity-60 disabled:cursor-wait"
            >
              {isDownloadingOriginal ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
              ) : (
                <FileDown className="h-3.5 w-3.5 shrink-0" />
              )}
              <span className="hidden lg:inline">
                {isDownloadingOriginal ? 'Đang tải...' : 'Tải xuống bản gốc'}
              </span>
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

            {/* Primary PowerPoint CTA: Real Fullscreen Slideshow */}
            <button
              type="button"
              onClick={() => startSlideshow(0)}
              aria-label="Trình chiếu toàn màn hình"
              title="Trình chiếu toàn màn hình (F5)"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-[#C43E1C] hover:bg-[#a83316] text-xs font-semibold text-white shadow-sm shadow-black/40 transition-all whitespace-nowrap"
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

      {/* Download Original Success Notice */}
      {downloadOriginalNotice && (
        <div className="shrink-0 px-4 py-2 bg-emerald-500/15 border-b border-emerald-400/30 text-xs text-emerald-200 flex items-center justify-between gap-2">
          <span>{downloadOriginalNotice}</span>
          <button
            type="button"
            onClick={() => setDownloadOriginalNotice(null)}
            aria-label="Đóng thông báo tải xuống"
            className="p-1 rounded hover:bg-emerald-500/20"
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
          {/* Microsoft PowerPoint Ribbon (#C43E1C, horizontally scrollable on mobile) */}
          <div className="shrink-0 bg-[#141923] border-b border-white/10">
            {/* Ribbon Tab Bar */}
            <div className="bg-[#C43E1C] px-2 sm:px-4 flex items-center justify-between gap-2 overflow-x-auto no-scrollbar">
              <div
                role="tablist"
                aria-label="Thanh Ribbon PowerPoint"
                className="flex items-center gap-0.5 shrink-0"
              >
                {PPT_RIBBON_TABS.map((tab) => {
                  const isTabActive = activeRibbonTab === tab.id;
                  return (
                    <button
                      key={tab.id}
                      type="button"
                      role="tab"
                      aria-selected={isTabActive}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setActiveRibbonTab(tab.id);
                        setIsLayoutPickerOpen(false);
                        setIsShapePickerOpen(false);
                      }}
                      className={`px-2.5 sm:px-3.5 py-1.5 text-[11.5px] font-semibold whitespace-nowrap transition-colors rounded-t-lg mt-1 ${
                        isTabActive
                          ? 'bg-[#141923] text-white shadow-sm'
                          : 'text-white/90 hover:bg-white/15 hover:text-white'
                      }`}
                    >
                      {tab.label}
                    </button>
                  );
                })}
              </div>

              <div className="flex items-center gap-1.5 shrink-0 py-1">
                {!isReadOnly && selectedBlockIds.length >= 2 && (
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleGroupSelected}
                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-black/25 hover:bg-black/40 text-[11px] font-semibold text-white whitespace-nowrap"
                  >
                    <Group className="h-3 w-3" />
                    <span>Nhóm ({selectedBlockIds.length})</span>
                  </button>
                )}
                {!isReadOnly && selectedBlock?.groupId && (
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleUngroupSelected}
                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-black/25 hover:bg-black/40 text-[11px] font-semibold text-white whitespace-nowrap"
                  >
                    <Ungroup className="h-3 w-3" />
                    <span>Bỏ nhóm</span>
                  </button>
                )}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setPptViewMode((prev) =>
                      prev === 'sorter' ? 'normal' : 'sorter'
                    )
                  }
                  aria-pressed={pptViewMode === 'sorter'}
                  className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-semibold whitespace-nowrap transition-colors ${
                    pptViewMode === 'sorter'
                      ? 'bg-white text-[#C43E1C]'
                      : 'bg-black/20 hover:bg-black/35 text-white'
                  }`}
                >
                  <LayoutGrid className="h-3 w-3" />
                  <span className="hidden sm:inline">Sắp xếp slide</span>
                </button>
              </div>
            </div>

            {/* Active Ribbon Tab Command Strip */}
            <div className="px-2.5 sm:px-4 py-1.5 bg-[#141923]/95 relative">
              <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-h-[34px]">
                {activeRibbonTab === 'home' && (
                  <>
                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddSlide}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-[#C43E1C] hover:bg-[#a83316] text-xs font-semibold text-white whitespace-nowrap disabled:opacity-40"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      <span>Slide mới</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setIsLayoutPickerOpen((v) => !v);
                        setIsShapePickerOpen(false);
                      }}
                      aria-expanded={isLayoutPickerOpen}
                      className={`shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        isLayoutPickerOpen
                          ? 'bg-[#C43E1C]/25 border-[#C43E1C] text-orange-200'
                          : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                      }`}
                    >
                      <LayoutTemplate className="h-3.5 w-3.5 text-orange-300" />
                      <span>Bố cục</span>
                      <ChevronDown className="h-3 w-3 opacity-70" />
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddTextBlock}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Type className="h-3.5 w-3.5 text-indigo-300" />
                      <span>Khung chữ</span>
                    </button>

                    {selectedBlock && !isReadOnly ? (
                      <>
                        <span className="h-5 w-px mx-0.5 bg-white/10 shrink-0" />
                        <select
                          aria-label="Phông chữ"
                          value={
                            selectedBlock.fontFamily ?? PPT_FONT_FAMILIES[0].value
                          }
                          onChange={(e) =>
                            handleUpdateBlock(selectedBlock.id, {
                              fontFamily: e.target.value,
                            })
                          }
                          className="shrink-0 h-8 px-2 rounded-lg bg-[#1C2331] border border-white/15 text-xs text-slate-100 outline-none focus:border-[#C43E1C]"
                        >
                          {PPT_FONT_FAMILIES.map((f) => (
                            <option key={f.label} value={f.value}>
                              {f.label}
                            </option>
                          ))}
                        </select>

                        <div className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
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
                                className={`min-h-[26px] px-1.5 rounded-md text-[11px] font-semibold tabular-nums transition-colors ${
                                  isSizeActive
                                    ? 'bg-[#C43E1C] text-white'
                                    : 'text-slate-300 hover:bg-white/10 hover:text-white'
                                }`}
                              >
                                {size}
                              </button>
                            );
                          })}
                        </div>

                        <button
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            handleUpdateBlock(selectedBlock.id, {
                              bold: !selectedBlock.bold,
                            })
                          }
                          aria-pressed={Boolean(selectedBlock.bold)}
                          title="In đậm"
                          className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border ${
                            selectedBlock.bold
                              ? 'bg-[#C43E1C]/35 border-[#C43E1C] text-white'
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
                          title="In nghiêng"
                          className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border ${
                            selectedBlock.italic
                              ? 'bg-[#C43E1C]/35 border-[#C43E1C] text-white'
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
                          aria-pressed={Boolean(
                            underlinedBlockIds[selectedBlock.id]
                          )}
                          title="Gạch chân"
                          className={`shrink-0 inline-flex items-center justify-center h-8 w-8 rounded-lg border ${
                            underlinedBlockIds[selectedBlock.id]
                              ? 'bg-[#C43E1C]/35 border-[#C43E1C] text-white'
                              : 'bg-white/[0.04] hover:bg-white/[0.1] border-white/10 text-slate-300'
                          }`}
                        >
                          <Underline className="h-3.5 w-3.5" />
                        </button>

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
                                title={`Màu chữ ${swatch.label}`}
                                className={`h-4 w-4 rounded-full border ${
                                  isColorActive
                                    ? 'ring-2 ring-orange-400 scale-110 border-white'
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
                            aria-label="Màu chữ tùy chỉnh"
                            className="h-4 w-5 rounded cursor-pointer border-0 bg-transparent p-0"
                          />
                        </div>

                        <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                          <span className="text-[10px] text-slate-400">Nền</span>
                          <input
                            type="color"
                            value={toSafeColorInputValue(
                              selectedBlock.fillColor,
                              '#C43E1C'
                            )}
                            onChange={(event) =>
                              handleUpdateBlock(selectedBlock.id, {
                                fillColor: event.target.value,
                              })
                            }
                            aria-label="Màu nền khối"
                            className="h-4 w-5 rounded cursor-pointer border-0 bg-transparent p-0"
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
                              className="text-[10px] text-slate-300 hover:text-white px-1 rounded bg-white/10"
                            >
                              Xóa
                            </button>
                          )}
                        </div>

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
                          ).map(({ value, label, Icon }) => (
                            <button
                              key={value}
                              type="button"
                              onMouseDown={handlePreventMouseDown}
                              onClick={() =>
                                handleUpdateBlock(selectedBlock.id, {
                                  align: value,
                                })
                              }
                              title={label}
                              className={`h-7 w-7 rounded-md flex items-center justify-center ${
                                (selectedBlock.align ?? 'left') === value
                                  ? 'bg-[#C43E1C] text-white'
                                  : 'text-slate-300 hover:bg-white/10'
                              }`}
                            >
                              <Icon className="h-3.5 w-3.5" />
                            </button>
                          ))}
                        </div>

                        <div className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                          <button
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => handleBringToFront(selectedBlock.id)}
                            title="Đưa lên trên cùng (Bring to Front)"
                            className="h-7 px-2 rounded-md text-[11px] text-slate-200 hover:bg-white/10 inline-flex items-center gap-1"
                          >
                            <BringToFront className="h-3.5 w-3.5 text-orange-300" />
                            <span className="hidden lg:inline">Lên trên</span>
                          </button>
                          <button
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => handleSendToBack(selectedBlock.id)}
                            title="Đưa xuống dưới cùng (Send to Back)"
                            className="h-7 px-2 rounded-md text-[11px] text-slate-200 hover:bg-white/10 inline-flex items-center gap-1"
                          >
                            <SendToBack className="h-3.5 w-3.5 text-slate-300" />
                            <span className="hidden lg:inline">Xuống dưới</span>
                          </button>
                        </div>

                        <button
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleDeleteBlock(selectedBlock.id)}
                          title="Xóa đối tượng đang chọn"
                          className="shrink-0 inline-flex items-center gap-1 h-8 px-2 rounded-lg bg-rose-500/15 hover:bg-rose-500/25 border border-rose-400/35 text-xs font-medium text-rose-200"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </>
                    ) : selectedImage && !isReadOnly ? (
                      <>
                        <span className="h-5 w-px mx-0.5 bg-white/10 shrink-0" />
                        <span className="text-[11.5px] text-orange-200 whitespace-nowrap">
                          Đang chọn ảnh · Kéo chấm tròn trên đỉnh để xoay
                        </span>
                        <button
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleDeleteImage(selectedImage.id)}
                          className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-rose-500/15 hover:bg-rose-500/25 border border-rose-400/35 text-xs font-medium text-rose-200 whitespace-nowrap"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                          <span>Xóa ảnh</span>
                        </button>
                      </>
                    ) : (
                      <span className="text-[11.5px] text-slate-400 whitespace-nowrap pl-1">
                        {isReadOnly
                          ? 'Chế độ Chỉ xem · Chọn tab Trình chiếu để bắt đầu.'
                          : 'Bấm chọn khung/ảnh trên slide để định dạng, xoay, nhóm hoặc Shift+Click chọn nhiều.'}
                      </span>
                    )}
                  </>
                )}

                {activeRibbonTab === 'insert' && (
                  <>
                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddTextBlock}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-[#C43E1C]/20 hover:bg-[#C43E1C]/35 border border-[#C43E1C]/50 text-xs font-semibold text-orange-100 whitespace-nowrap disabled:opacity-40"
                    >
                      <Type className="h-3.5 w-3.5" />
                      <span>Hộp văn bản</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleTriggerAddImage}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <ImagePlus className="h-3.5 w-3.5 text-emerald-300" />
                      <span>Hình ảnh</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleInsertTable(3, 3)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Table2 className="h-3.5 w-3.5 text-sky-300" />
                      <span>Bảng (3×3)</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsChartModalOpen(true)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <BarChart3 className="h-3.5 w-3.5 text-amber-300" />
                      <span>Biểu đồ</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setIsShapePickerOpen((v) => !v);
                        setIsLayoutPickerOpen(false);
                      }}
                      aria-expanded={isShapePickerOpen}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        isShapePickerOpen
                          ? 'bg-[#C43E1C]/25 border-[#C43E1C] text-orange-200'
                          : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                      }`}
                    >
                      <Shapes className="h-3.5 w-3.5 text-orange-300" />
                      <span>Hình khối (11)</span>
                      <ChevronDown className="h-3 w-3 opacity-70" />
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('link')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Link2 className="h-3.5 w-3.5 text-cyan-300" />
                      <span>Liên kết</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('video')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Video className="h-3.5 w-3.5 text-rose-300" />
                      <span>Video</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('audio')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Music className="h-3.5 w-3.5 text-violet-300" />
                      <span>Âm thanh</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'design' && (
                  <>
                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Palette className="h-3.5 w-3.5 text-orange-300 mr-0.5 shrink-0" />
                      {PPT_COLOR_THEMES.map((theme) => {
                        const isActiveTheme =
                          (currentSlide.backgroundColor ?? '').toLowerCase() ===
                            theme.bgColor.toLowerCase() &&
                          (currentSlide.accentColor ?? '').toLowerCase() ===
                            theme.accentColor.toLowerCase();
                        return (
                          <button
                            key={theme.id}
                            type="button"
                            disabled={isReadOnly}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => handleApplyColorTheme(theme)}
                            title={theme.name}
                            className={`h-8 px-2 rounded-lg border inline-flex items-center gap-1 text-[11px] font-medium whitespace-nowrap disabled:opacity-40 ${
                              isActiveTheme
                                ? 'border-[#C43E1C] bg-[#C43E1C]/20 text-white'
                                : 'border-white/10 bg-[#1C2331] hover:border-white/30 text-slate-200'
                            }`}
                          >
                            <span className="flex items-center -space-x-1">
                              {theme.swatches.map((hex, i) => (
                                <span
                                  key={`${theme.id}-sw-${i}`}
                                  className="h-3 w-3 rounded-full border border-white/30"
                                  style={{ backgroundColor: hex }}
                                />
                              ))}
                            </span>
                            <span className="hidden lg:inline">{theme.name}</span>
                          </button>
                        );
                      })}
                    </div>

                    <span className="h-5 w-px mx-0.5 bg-white/10 shrink-0" />

                    <div className="inline-flex items-center gap-1 shrink-0">
                      {PPT_GRADIENT_PRESETS.map((g) => (
                        <button
                          key={g.label}
                          type="button"
                          disabled={isReadOnly}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleApplySlideGradient(g.value)}
                          title={`Nền chuyển sắc: ${g.label}`}
                          className="h-6 w-6 rounded border border-white/25 hover:scale-110 transition-transform disabled:opacity-40"
                          style={{ backgroundImage: g.value }}
                        />
                      ))}
                    </div>

                    <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                      <PaintBucket className="h-3 w-3 text-slate-400" />
                      {SLIDE_BG_PRESETS.map((bg) => (
                        <button
                          key={bg.value}
                          type="button"
                          disabled={isReadOnly}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleChangeSlideBackground(bg.value)}
                          title={`Màu đơn: ${bg.label}`}
                          className="h-4 w-4 rounded border border-white/25 hover:scale-110 transition-transform disabled:opacity-40"
                          style={{ backgroundColor: bg.value }}
                        />
                      ))}
                      <input
                        type="color"
                        disabled={isReadOnly}
                        value={toSafeColorInputValue(
                          currentSlide.backgroundColor,
                          '#161D2E'
                        )}
                        onChange={(event) =>
                          handleChangeSlideBackground(event.target.value)
                        }
                        aria-label="Màu nền slide tùy chỉnh"
                        className="h-4 w-5 rounded cursor-pointer border-0 bg-transparent p-0"
                      />
                    </div>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => bgImageInputRef.current?.click()}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap disabled:opacity-40"
                    >
                      <Image className="h-3.5 w-3.5 text-emerald-300" />
                      <span>Ảnh nền</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleApplyBackgroundToAll}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-[#C43E1C]/20 hover:bg-[#C43E1C]/35 border border-[#C43E1C]/50 text-xs font-semibold text-orange-100 whitespace-nowrap disabled:opacity-40"
                    >
                      <span>Áp dụng nền tất cả</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'transitions' && (
                  <>
                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Sparkles className="h-3.5 w-3.5 text-orange-300 mr-0.5 shrink-0" />
                      <span className="text-[11px] text-slate-400">Chuyển slide:</span>
                      {PPT_TRANSITION_OPTIONS.map((opt) => {
                        const active =
                          (currentSlide.transition?.type ?? 'fade') === opt.type;
                        return (
                          <button
                            key={opt.type}
                            type="button"
                            disabled={isReadOnly}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() =>
                              mutateSlideAt(safeSlideIdx, (s) => ({
                                ...s,
                                transition: {
                                  type: opt.type,
                                  durationMs: s.transition?.durationMs ?? 350,
                                },
                              }))
                            }
                            className={`h-7 px-2 rounded-md border text-[11px] font-medium whitespace-nowrap disabled:opacity-40 ${
                              active
                                ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                                : 'bg-[#1C2331] border-white/10 text-slate-200 hover:bg-white/10'
                            }`}
                          >
                            {opt.label}
                          </button>
                        );
                      })}
                    </div>

                    <select
                      aria-label="Thời lượng chuyển cảnh"
                      disabled={isReadOnly}
                      value={currentSlide.transition?.durationMs ?? 350}
                      onChange={(e) =>
                        mutateSlideAt(safeSlideIdx, (s) => ({
                          ...s,
                          transition: {
                            type: s.transition?.type ?? 'fade',
                            durationMs: Number(e.target.value),
                          },
                        }))
                      }
                      className="shrink-0 h-7 px-2 rounded-md bg-[#1C2331] border border-white/15 text-[11px] text-slate-100"
                    >
                      <option value={200}>0.2s</option>
                      <option value={350}>0.35s</option>
                      <option value={600}>0.6s</option>
                      <option value={1000}>1.0s</option>
                    </select>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleApplyTransitionToAll}
                      className="shrink-0 h-7 px-2.5 rounded-md bg-[#C43E1C]/20 hover:bg-[#C43E1C]/35 border border-[#C43E1C]/50 text-[11px] font-semibold text-orange-100 whitespace-nowrap disabled:opacity-40"
                    >
                      Áp dụng tất cả
                    </button>

                    <span className="h-5 w-px mx-0.5 bg-white/10 shrink-0" />

                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Wand2 className="h-3.5 w-3.5 text-amber-300 shrink-0" />
                      <span className="text-[11px] text-slate-400">
                        Hoạt ảnh khối:
                      </span>
                      {selectedBlock && !isReadOnly ? (
                        <>
                          {PPT_ANIMATION_OPTIONS.map((animOpt) => {
                            const active =
                              (selectedBlock.animation?.type ?? 'none') ===
                              animOpt.type;
                            return (
                              <button
                                key={animOpt.type}
                                type="button"
                                onMouseDown={handlePreventMouseDown}
                                onClick={() =>
                                  handleUpdateBlock(selectedBlock.id, {
                                    animation: {
                                      type: animOpt.type,
                                      order:
                                        animOpt.type === 'none'
                                          ? undefined
                                          : selectedBlock.animation?.order ?? 1,
                                    },
                                  })
                                }
                                className={`h-7 px-2 rounded-md border text-[11px] font-medium whitespace-nowrap ${
                                  active
                                    ? 'bg-indigo-600 border-indigo-400 text-white'
                                    : 'bg-[#1C2331] border-white/10 text-slate-200 hover:bg-white/10'
                                }`}
                              >
                                {animOpt.label}
                              </button>
                            );
                          })}
                          {selectedBlock.animation &&
                            selectedBlock.animation.type !== 'none' && (
                              <select
                                aria-label="Thứ tự click hiệu ứng"
                                value={selectedBlock.animation.order ?? 1}
                                onChange={(e) =>
                                  handleUpdateBlock(selectedBlock.id, {
                                    animation: {
                                      type:
                                        selectedBlock.animation?.type ?? 'fadeIn',
                                      order: Number(e.target.value),
                                    },
                                  })
                                }
                                className="h-7 px-1.5 rounded-md bg-[#1C2331] border border-white/15 text-[11px] text-slate-100"
                              >
                                {[1, 2, 3, 4, 5, 6].map((n) => (
                                  <option key={n} value={n}>
                                    Click #{n}
                                  </option>
                                ))}
                              </select>
                            )}
                        </>
                      ) : (
                        <span className="text-[11px] text-slate-500 whitespace-nowrap">
                          (Chọn 1 đối tượng trên slide)
                        </span>
                      )}
                    </div>
                  </>
                )}

                {activeRibbonTab === 'slideshow' && (
                  <>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => startSlideshow(0)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-3 rounded-lg bg-[#C43E1C] hover:bg-[#a83316] text-xs font-semibold text-white whitespace-nowrap"
                    >
                      <Play className="h-3.5 w-3.5 fill-white" />
                      <span>Từ đầu (F5)</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => startSlideshow(safeSlideIdx)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-3 rounded-lg bg-white/[0.08] hover:bg-white/[0.15] border border-white/15 text-xs font-semibold text-white whitespace-nowrap"
                    >
                      <MonitorPlay className="h-3.5 w-3.5 text-orange-300" />
                      <span>Từ trang hiện tại (Shift+F5)</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleToggleHideSlide(safeSlideIdx)}
                      aria-pressed={Boolean(currentSlide.hidden)}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        currentSlide.hidden
                          ? 'bg-amber-500/20 border-amber-400/50 text-amber-200'
                          : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                      }`}
                    >
                      <EyeOff className="h-3.5 w-3.5" />
                      <span>{currentSlide.hidden ? 'Bỏ ẩn slide' : 'Ẩn slide'}</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsLaserActive((p) => !p)}
                      aria-pressed={isLaserActive}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        isLaserActive
                          ? 'bg-rose-500/25 border-rose-400/60 text-rose-100'
                          : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                      }`}
                    >
                      <Crosshair className="h-3.5 w-3.5" />
                      <span>Bút Laser</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => onOpenModal('present', buildLiveDocument())}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap"
                    >
                      <Presentation className="h-3.5 w-3.5 text-indigo-300" />
                      <span>Presenter View</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => onOpenModal('qr', buildLiveDocument())}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap"
                    >
                      <QrCode className="h-3.5 w-3.5 text-cyan-300" />
                      <span>Điều khiển QR</span>
                    </button>
                  </>
                )}
              </div>

              {/* Popover: 6 Slide Layouts Picker */}
              {isLayoutPickerOpen && !isReadOnly && (
                <div className="mt-2 p-2.5 rounded-xl bg-[#1A202E] border border-white/15 shadow-2xl grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2 z-30 relative">
                  {PPT_SLIDE_LAYOUTS.map((layout) => (
                    <button
                      key={layout.id}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleApplySlideLayout(layout.id)}
                      className="text-left p-2 rounded-lg bg-white/[0.04] hover:bg-[#C43E1C]/20 border border-white/10 hover:border-[#C43E1C]/60 transition-colors"
                    >
                      <div className="text-[11px] font-semibold text-white">
                        {layout.label}
                      </div>
                      <div className="text-[10px] text-slate-400 mt-0.5 line-clamp-1">
                        {layout.desc}
                      </div>
                    </button>
                  ))}
                </div>
              )}

              {/* Popover: 11 SVG Shapes Picker */}
              {isShapePickerOpen && !isReadOnly && (
                <div className="mt-2 p-2.5 rounded-xl bg-[#1A202E] border border-white/15 shadow-2xl flex flex-wrap gap-1.5 z-30 relative">
                  {PPT_SHAPE_OPTIONS.map((shapeOpt) => {
                    const ShapeIcon = shapeOpt.Icon;
                    return (
                      <button
                        key={shapeOpt.kind}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleInsertShape(shapeOpt.kind)}
                        className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-white/[0.05] hover:bg-[#C43E1C]/25 border border-white/10 hover:border-[#C43E1C] text-[11px] font-medium text-slate-100 transition-colors"
                      >
                        <ShapeIcon className="h-3.5 w-3.5 text-orange-300" />
                        <span>{shapeOpt.label}</span>
                      </button>
                    );
                  })}
                </div>
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

          {/* Slide Sorter Grid OR Main Editable Slide Stage Container */}
          {pptViewMode === 'sorter' ? (
            <div className="flex-1 min-h-0 overflow-y-auto scrollbar-dark p-4 sm:p-6">
              <div className="flex items-center justify-between gap-3 mb-4">
                <div>
                  <h2 className="text-sm sm:text-base font-bold text-white flex items-center gap-2">
                    <LayoutGrid className="h-4 w-4 text-[#C43E1C]" />
                    <span>Chế độ sắp xếp trang chiếu (Slide Sorter)</span>
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Kéo thả để đổi thứ tự slide hoặc bấm đúp vào slide để chỉnh sửa chi tiết.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setPptViewMode('normal')}
                  className="px-3 py-1.5 rounded-xl bg-[#C43E1C] hover:bg-[#a83316] text-xs font-semibold text-white shrink-0"
                >
                  Về chế độ Bình thường
                </button>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                {slides.map((slideItem, idx) => {
                  const isActive = idx === safeSlideIdx;
                  return (
                    <div
                      key={`sorter-${slideItem.id}`}
                      draggable={!isReadOnly}
                      onDragStart={(e) => {
                        if (isReadOnly) {
                          return;
                        }
                        setDraggedSlideIdx(idx);
                        e.dataTransfer.effectAllowed = 'move';
                      }}
                      onDragOver={(e) => {
                        if (isReadOnly || draggedSlideIdx === null) {
                          return;
                        }
                        e.preventDefault();
                        if (dragOverSlideIdx !== idx) {
                          setDragOverSlideIdx(idx);
                        }
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        if (
                          !isReadOnly &&
                          draggedSlideIdx !== null &&
                          draggedSlideIdx !== idx
                        ) {
                          handleReorderSlide(draggedSlideIdx, idx);
                        }
                        setDraggedSlideIdx(null);
                        setDragOverSlideIdx(null);
                      }}
                      onDragEnd={() => {
                        setDraggedSlideIdx(null);
                        setDragOverSlideIdx(null);
                      }}
                      onClick={() => handleSelectSlide(idx)}
                      onDoubleClick={() => {
                        handleSelectSlide(idx);
                        setPptViewMode('normal');
                      }}
                      className={`group rounded-2xl p-2.5 border cursor-pointer transition-all flex flex-col ${
                        isActive
                          ? 'bg-[#C43E1C]/15 border-[#C43E1C] ring-2 ring-[#C43E1C]/60'
                          : dragOverSlideIdx === idx
                            ? 'bg-white/10 border-amber-400 ring-2 ring-amber-400/50'
                            : 'bg-[#141923] border-white/10 hover:border-white/25'
                      } ${slideItem.hidden ? 'opacity-55' : ''}`}
                    >
                      <div
                        style={{
                          aspectRatio: String(
                            getSafeAspectRatio(slideItem, document)
                          ),
                        }}
                        className="w-full rounded-xl overflow-hidden border border-white/10 pointer-events-none relative"
                      >
                        <ReadOnlySlideStage
                          doc={document}
                          slide={slideItem}
                          slideIndex={idx}
                          totalSlides={slides.length}
                          stageWidthPx={320}
                          underlinedBlockIds={underlinedBlockIds}
                        />
                        {slideItem.hidden && (
                          <span className="absolute top-2 right-2 px-1.5 py-0.5 rounded bg-black/75 text-[10px] font-semibold text-amber-200 inline-flex items-center gap-1">
                            <EyeOff className="h-3 w-3" />
                            <span>Đang ẩn</span>
                          </span>
                        )}
                      </div>

                      <div className="mt-2 flex items-center justify-between gap-2">
                        <div className="min-w-0 flex items-center gap-1.5">
                          <span
                            className={`px-1.5 py-0.5 rounded text-[10px] font-bold tabular-nums ${
                              isActive
                                ? 'bg-[#C43E1C] text-white'
                                : 'bg-white/10 text-slate-300'
                            }`}
                          >
                            {idx + 1}
                          </span>
                          <span className="text-xs font-medium text-slate-200 truncate">
                            {slideItem.title}
                          </span>
                        </div>

                        {!isReadOnly && (
                          <div
                            onClick={(e) => e.stopPropagation()}
                            className="flex items-center gap-1 shrink-0"
                          >
                            <button
                              type="button"
                              onClick={() => handleDuplicateSlideAt(idx)}
                              title="Nhân bản slide"
                              className="p-1 rounded hover:bg-white/10 text-slate-400 hover:text-white"
                            >
                              <Copy className="h-3.5 w-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleToggleHideSlide(idx)}
                              title={slideItem.hidden ? 'Hiện slide' : 'Ẩn slide'}
                              className="p-1 rounded hover:bg-white/10 text-slate-400 hover:text-amber-300"
                            >
                              {slideItem.hidden ? (
                                <Eye className="h-3.5 w-3.5 text-amber-300" />
                              ) : (
                                <EyeOff className="h-3.5 w-3.5" />
                              )}
                            </button>
                            <button
                              type="button"
                              disabled={slides.length <= 1}
                              onClick={() => handleDeleteSlideAt(idx)}
                              title="Xóa slide"
                              className="p-1 rounded hover:bg-rose-500/20 text-slate-400 hover:text-rose-300 disabled:opacity-30"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
          <div className="flex-1 min-h-0 flex flex-col px-3 sm:px-6 lg:px-10 pt-3 sm:pt-4 pb-2 overflow-hidden">
            <div
              ref={editorViewportRef}
              className="flex-1 min-h-0 w-full flex items-center justify-center relative overflow-auto"
            >
              <div
                ref={editorCanvasRef}
                onClick={() => {
                  if (justDraggedRef.current) {
                    return;
                  }
                  if (!isLaserActive) {
                    setSelectedBlockId(null);
                    setSelectedBlockIds([]);
                    setEditingBlockId(null);
                    setSelectedImageId(null);
                  }
                }}
                onPointerMove={updateLaserFromPointer}
                onPointerDown={updateLaserFromPointer}
                onPointerLeave={handlePointerLeaveStage}
                style={{
                  width: `${Math.round(zoomedWidthPx)}px`,
                  height: `${Math.round(zoomedHeightPx)}px`,
                  ...buildSlideBackgroundStyle(currentSlide),
                }}
                className={`relative rounded-2xl overflow-hidden border shadow-2xl shadow-black/55 transition-shadow duration-200 ${
                  isLaserActive
                    ? 'border-rose-400/55 ring-2 ring-rose-500/25 cursor-none touch-none select-none'
                    : 'border-white/15 ring-1 ring-white/[0.06]'
                }`}
              >
                {viewMode === 'native' || isReadOnly ? (
                  <ReadOnlySlideStage
                    doc={document}
                    slide={currentSlide}
                    slideIndex={safeSlideIdx}
                    totalSlides={slides.length}
                    stageWidthPx={actualEditorCanvasWidth}
                    underlinedBlockIds={underlinedBlockIds}
                  />
                ) : isPositionedSlide ? (
                  /* Master-Positioned OpenXML Coordinate Stage (xPct, yPct, wPct, hPct) */
                  <div className="absolute inset-0">
                    {sortBlocksByZIndex(currentBlocks).map((block, bIdx) => {
                      const isSelected =
                        !isLaserActive &&
                        (selectedBlockId === block.id ||
                          selectedBlockIds.includes(block.id));
                      const isBeingDragged = draggingId === block.id;
                      const xPct = block.xPct ?? 6;
                      const yPct = block.yPct ?? 10 + bIdx * 14;
                      const wPct = block.wPct ?? 88;
                      const hPct = block.hPct ?? 16;
                      const scaledFontPx = Math.max(
                        4,
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
                      const blockPadding = getBlockPaddingStyle(
                        block.paddingPct,
                        editorFontScale
                      );
                      const safeLink = sanitizeExternalUrl(block.linkUrl);
                      const isNonTextBlock =
                        block.kind === 'shape' ||
                        block.kind === 'chart' ||
                        block.kind === 'media';

                      return (
                        <div
                          key={block.id}
                          onClick={(event) => {
                            if (isLaserActive) {
                              return;
                            }
                            event.stopPropagation();
                            if (event.shiftKey) {
                              setSelectedBlockIds((prev) =>
                                prev.includes(block.id)
                                  ? prev
                                  : [...prev, block.id]
                              );
                            } else if (block.groupId) {
                              setSelectedBlockIds(
                                currentBlocks
                                  .filter((b) => b.groupId === block.groupId)
                                  .map((b) => b.id)
                              );
                            } else {
                              setSelectedBlockIds([block.id]);
                            }
                            setSelectedBlockId(block.id);
                            setSelectedImageId(null);
                          }}
                          onPointerDown={(event) => {
                            if (
                              event.target === event.currentTarget ||
                              block.kind === 'shape' ||
                              block.kind === 'chart'
                            ) {
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
                            zIndex: isSelected
                              ? (block.zIndex ?? 0) + 20
                              : (block.zIndex ?? 0) + 1,
                            backgroundColor:
                              block.kind === 'shape' ? undefined : block.fillColor,
                            backgroundImage:
                              block.kind === 'shape' ? undefined : block.fillGradient,
                            borderColor:
                              block.kind === 'shape' ? undefined : block.borderColor,
                            borderWidth:
                              block.kind !== 'shape' && block.borderColor
                                ? `${block.borderWidthPx ?? 1}px`
                                : undefined,
                            borderStyle:
                              block.kind !== 'shape' && block.borderColor
                                ? 'solid'
                                : undefined,
                            borderRadius: block.borderRadius ?? '6px',
                            boxShadow: block.shadowCss,
                            padding: isNonTextBlock ? undefined : blockPadding,
                            transform: block.rotationDeg
                              ? `rotate(${block.rotationDeg}deg)`
                              : undefined,
                            fontFamily:
                              block.fontFamily ||
                              "'Inter', 'Segoe UI', system-ui, sans-serif",
                          }}
                          className={`group flex flex-col ${justifyClass} ${
                            isNonTextBlock || blockPadding ? '' : 'px-1.5 py-1'
                          } transition-shadow ${
                            isSelected
                              ? 'ring-2 ring-[#C43E1C]'
                              : isLaserActive
                                ? ''
                                : hasLightSlideBg
                                  ? 'hover:ring-1 hover:ring-slate-400/50'
                                  : 'hover:ring-1 hover:ring-white/30'
                          } ${isBeingDragged ? 'opacity-90 shadow-xl' : ''}`}
                        >
                          {isSelected && !isReadOnly && (
                            <PowerPointResizeHandles
                              onStartResize={(ev, dir) =>
                                startElementResize(ev, dir, {
                                  kind: 'block',
                                  id: block.id,
                                  xPct,
                                  yPct,
                                  wPct,
                                  hPct,
                                })
                              }
                              onStartRotate={(ev) =>
                                startElementRotate(ev, {
                                  kind: 'block',
                                  id: block.id,
                                  xPct,
                                  yPct,
                                  wPct,
                                  hPct,
                                })
                              }
                            />
                          )}
                          {/* Floating Drag Handle + Group/Anim Indicator */}
                          {!isLaserActive && (
                            <div
                              className={`absolute -top-3 left-1 z-20 inline-flex items-center gap-1 transition-opacity ${
                                isSelected
                                  ? 'opacity-100'
                                  : 'opacity-0 group-hover:opacity-100'
                              }`}
                            >
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
                                className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-md bg-[#C43E1C] text-white text-[9px] font-semibold shadow-md cursor-grab active:cursor-grabbing"
                              >
                                <Move className="h-2.5 w-2.5" />
                                <span>Kéo</span>
                              </button>
                              {block.animation && block.animation.type !== 'none' && (
                                <span className="px-1 py-0.5 rounded bg-indigo-600 text-[8.5px] font-bold text-white shadow">
                                  #{block.animation.order ?? 1}
                                </span>
                              )}
                              {block.groupId && (
                                <span className="px-1 py-0.5 rounded bg-amber-500 text-[8.5px] font-bold text-slate-950 shadow">
                                  Nhóm
                                </span>
                              )}
                            </div>
                          )}

                          {block.kind === 'shape' ? (
                            <div
                              onDoubleClick={(e) => {
                                if (!isReadOnly && block.shapeKind !== 'line') {
                                  e.stopPropagation();
                                  setEditingBlockId(block.id);
                                }
                              }}
                              className="w-full h-full relative flex items-center justify-center"
                            >
                              <SlideShapeSvg
                                shapeKind={block.shapeKind ?? 'rect'}
                                fillColor={block.fillColor}
                                borderColor={block.borderColor}
                                borderWidthPx={block.borderWidthPx}
                              />
                              {block.shapeKind !== 'line' &&
                                editingBlockId !== block.id &&
                                block.text && (
                                  <span
                                    className="relative z-10 px-2 text-center leading-tight pointer-events-none"
                                    style={{
                                      fontSize: `${scaledFontPx.toFixed(2)}px`,
                                      color: textColor,
                                      fontWeight: block.bold ? 700 : 500,
                                    }}
                                  >
                                    {block.text}
                                  </span>
                                )}
                              {editingBlockId === block.id &&
                                block.shapeKind !== 'line' && (
                                  <input
                                    type="text"
                                    autoFocus
                                    value={block.text}
                                    onClick={(e) => e.stopPropagation()}
                                    onPointerDown={(e) => e.stopPropagation()}
                                    onChange={(e) =>
                                      handleUpdateBlock(block.id, {
                                        text: e.target.value,
                                      })
                                    }
                                    onBlur={() => setEditingBlockId(null)}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter' || e.key === 'Escape') {
                                        setEditingBlockId(null);
                                      }
                                    }}
                                    className="absolute inset-2 z-20 bg-black/65 text-white text-xs text-center rounded border border-white/40 outline-none px-1"
                                  />
                                )}
                            </div>
                          ) : block.kind === 'chart' && block.chartData ? (
                            <div className="w-full h-full">
                              <SlideChartSvg
                                chartData={block.chartData}
                                hasLightBg={hasLightSlideBg}
                                fontScale={editorFontScale}
                              />
                            </div>
                          ) : block.kind === 'media' && block.mediaSrc ? (
                            <div
                              onPointerDown={(e) => e.stopPropagation()}
                              className="w-full h-full rounded-lg overflow-hidden bg-slate-950/90 border border-white/20 flex flex-col justify-center p-1"
                            >
                              {block.mediaKind === 'audio' ? (
                                <div className="w-full px-1 flex flex-col gap-1">
                                  <span className="text-[10px] font-medium text-slate-200 truncate">
                                    {block.text || 'Tệp âm thanh'}
                                  </span>
                                  <audio
                                    controls
                                    src={block.mediaSrc}
                                    className="w-full h-7"
                                  />
                                </div>
                              ) : (
                                <video
                                  controls
                                  playsInline
                                  src={block.mediaSrc}
                                  className="w-full h-full object-contain rounded"
                                />
                              )}
                            </div>
                          ) : block.kind === 'table' &&
                            Array.isArray(block.tableRows) &&
                            block.tableRows.length > 0 ? (
                            <div className="w-full h-full overflow-auto no-scrollbar">
                              <table className="w-full h-full border-collapse table-fixed">
                                <tbody>
                                  {block.tableRows.map((row, rIdx) => (
                                    <tr key={`${block.id}-r-${rIdx}`}>
                                      {row.map((cell, cIdx) => {
                                        const cellStyle =
                                          block.tableCellStyles?.[rIdx]?.[cIdx];
                                        return (
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
                                              cellStyle?.fillColor ??
                                              (rIdx === 0
                                                ? block.fillColor ??
                                                  (hasLightSlideBg
                                                    ? 'rgba(15,23,42,0.06)'
                                                    : 'rgba(255,255,255,0.08)')
                                                : 'transparent'),
                                          }}
                                        >
                                          <input
                                            type="text"
                                            value={cell}
                                            disabled={isLaserActive || isReadOnly}
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
                                                (cellStyle?.bold ?? (rIdx === 0 || block.bold))
                                                  ? 700
                                                  : 400,
                                              fontStyle: block.italic
                                                ? 'italic'
                                                : 'normal',
                                              textDecoration:
                                                underlinedBlockIds[block.id]
                                                  ? 'underline'
                                                  : 'none',
                                              color: cellStyle?.textColor ?? textColor,
                                              textAlign: cellStyle?.align ?? block.align ?? 'left',
                                              fontFamily:
                                                block.fontFamily || undefined,
                                            }}
                                            className="w-full min-w-0 bg-transparent border-0 outline-none truncate"
                                          />
                                        </td>
                                        );
                                      })}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : block.text.trim().length === 0 && editingBlockId !== block.id && newlyAddedBlockId !== block.id ? (
                            <div
                              onDoubleClick={(e) => {
                                if (!isReadOnly && !isLaserActive) {
                                  e.stopPropagation();
                                  setSelectedBlockId(block.id);
                                  setEditingBlockId(block.id);
                                }
                              }}
                              className="w-full h-full"
                            />
                          ) : editingBlockId === block.id || newlyAddedBlockId === block.id || !block.richParagraphs || block.richParagraphs.length === 0 ? (
                            <AutoGrowTextarea
                              value={block.text}
                              onChange={(nextText) =>
                                handleUpdateBlock(block.id, {
                                  text: nextText,
                                  richParagraphs: undefined,
                                })
                              }
                              onFocus={() => {
                                if (!isLaserActive) {
                                  setSelectedBlockId(block.id);
                                  setEditingBlockId(block.id);
                                  setSelectedImageId(null);
                                }
                              }}
                              disabled={isLaserActive || isReadOnly}
                              autoFocus={editingBlockId === block.id || newlyAddedBlockId === block.id}
                              placeholder={
                                block.role === 'title'
                                  ? 'Nhập tiêu đề slide...'
                                  : 'Nhập nội dung khung chữ...'
                              }
                              ariaLabel={`Khung chữ ${block.role ?? 'nội dung'} trên Slide ${safeSlideIdx + 1}`}
                              scaledFontPx={scaledFontPx * ((block.fontScalePct ?? 100) / 100)}
                              bold={Boolean(block.bold)}
                              italic={Boolean(block.italic)}
                              underline={Boolean(underlinedBlockIds[block.id])}
                              color={textColor}
                              align={block.align ?? 'left'}
                              fontFamily={block.fontFamily}
                              maxHeightFull
                            />
                          ) : (
                            <RichTextFittedView
                              block={block}
                              fontScale={editorFontScale}
                              textColor={textColor}
                              underlined={Boolean(underlinedBlockIds[block.id])}
                              onActivateEdit={() => {
                                if (!isLaserActive) {
                                  setSelectedBlockId(block.id);
                                  setSelectedBlockIds([block.id]);
                                  setSelectedImageId(null);
                                  if (!isReadOnly) setEditingBlockId(block.id);
                                }
                              }}
                            />
                          )}
                          {safeLink && (
                            <a
                              href={safeLink}
                              target="_blank"
                              rel="noopener noreferrer"
                              onClick={(e) => e.stopPropagation()}
                              title={`Mở liên kết: ${safeLink}`}
                              className="mt-0.5 inline-flex items-center gap-1 text-[10px] text-sky-300 hover:underline self-start"
                            >
                              <ExternalLink className="h-2.5 w-2.5 shrink-0" />
                              <span className="truncate max-w-[160px]">{safeLink}</span>
                            </a>
                          )}
                        </div>
                      );
                    })}

                    {/* Smart Alignment Guides (while dragging elements) */}
                    {smartGuides.xPct !== null && (
                      <div
                        aria-hidden="true"
                        className="pointer-events-none absolute inset-y-0 z-40 border-l border-dashed border-cyan-400 shadow-[0_0_6px_rgba(34,211,238,0.85)]"
                        style={{ left: `${smartGuides.xPct}%` }}
                      />
                    )}
                    {smartGuides.yPct !== null && (
                      <div
                        aria-hidden="true"
                        className="pointer-events-none absolute inset-x-0 z-40 border-t border-dashed border-cyan-400 shadow-[0_0_6px_rgba(34,211,238,0.85)]"
                        style={{ top: `${smartGuides.yPct}%` }}
                      />
                    )}

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
                            setSelectedBlockIds([]);
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
                            borderRadius: imageItem.borderRadius ?? "6px",
                            transform: imageItem.rotationDeg ? `rotate(${imageItem.rotationDeg}deg)` : undefined,
                          }}
                          className={`group rounded-lg flex items-center justify-center transition-shadow ${
                            isLaserActive
                              ? 'pointer-events-none'
                              : 'cursor-grab active:cursor-grabbing'
                          } ${
                            isSelected
                              ? 'ring-2 ring-[#C43E1C] z-10'
                              : 'hover:ring-1 hover:ring-[#C43E1C]/60'
                          }`}
                        >
                          {isSelected && !isReadOnly && (
                            <PowerPointResizeHandles
                              onStartResize={(ev, dir) =>
                                startElementResize(ev, dir, { kind: "image", id: imageItem.id, xPct, yPct, wPct, hPct })
                              }
                              onStartRotate={(ev) =>
                                startElementRotate(ev, {
                                  kind: 'image',
                                  id: imageItem.id,
                                  xPct,
                                  yPct,
                                  wPct,
                                  hPct,
                                })
                              }
                            />
                          )}
                          <div
                            style={{ borderRadius: imageItem.borderRadius ?? '6px' }}
                            className="relative w-full h-full overflow-hidden flex items-center justify-center"
                          >
                            {renderCroppedSlideImage(
                              imageItem,
                              imageItem.alt ??
                                `Hình ảnh ${imgIdx + 1} của Slide ${safeSlideIdx + 1}`
                            )}
                          </div>
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
                              4,
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

              <div className="hidden xl:flex items-center gap-1.5 max-w-[140px] overflow-x-auto no-scrollbar">
                {slides.map((slideItem, idx) => (
                  <button
                    key={`${slideItem.id}-dot`}
                    type="button"
                    onClick={() => handleSelectSlide(idx)}
                    aria-label={`Đi tới Slide ${idx + 1}`}
                    className={`h-2 shrink-0 rounded-full transition-all ${
                      idx === safeSlideIdx
                        ? 'w-5 bg-[#C43E1C]'
                        : 'w-2 bg-white/20 hover:bg-white/45'
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>
          )}

          {/* Resizable Bottom Speaker Notes Dock */}
          <section
            aria-label="Ghi chú diễn giả"
            className="shrink-0 border-t border-white/10 bg-[#141923]"
          >
            {isNotesExpanded && (
              <div
                onPointerDown={startNotesResize}
                title="Kéo lên/xuống để thay đổi chiều cao ghi chú diễn giả"
                className="h-1.5 w-full cursor-row-resize hover:bg-[#C43E1C]/50 transition-colors flex items-center justify-center"
              >
                <span className="h-0.5 w-10 rounded-full bg-white/20" />
              </div>
            )}
            <div className="px-3.5 sm:px-5 py-1.5 flex items-center justify-between gap-2">
              <div className="flex items-center gap-2 min-w-0">
                <StickyNote className="h-3.5 w-3.5 text-[#C43E1C] shrink-0" />
                <span className="text-xs font-semibold text-slate-200 whitespace-nowrap">
                  Ghi chú diễn giả
                </span>
                <span className="text-[11px] text-slate-500 tabular-nums whitespace-nowrap">
                  (Slide {safeSlideIdx + 1})
                </span>
                {isReadOnly && (
                  <span className="text-[10.5px] text-amber-300/90 whitespace-nowrap">
                    · Chỉ xem
                  </span>
                )}

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
              <div className="px-3.5 sm:px-5 pb-2.5">
                <textarea
                  value={currentSlide.speakerNote ?? ''}
                  disabled={isReadOnly}
                  onChange={(event) =>
                    handleSpeakerNoteChange(event.target.value)
                  }
                  style={{ height: `${notesHeightPx}px` }}
                  aria-label={`Nội dung ghi chú diễn giả cho Slide ${safeSlideIdx + 1}`}
                  placeholder={
                    isReadOnly
                      ? 'Slide này ở chế độ Chỉ xem.'
                      : 'Nhập ghi chú diễn giả cho slide này (được lưu và xuất kèm tệp .pptx)...'
                  }
                  className="w-full resize-none rounded-xl bg-[#181D27] border border-white/10 focus:border-[#C43E1C]/70 focus:outline-none focus:ring-2 focus:ring-[#C43E1C]/25 px-3.5 py-2 text-xs sm:text-[13px] text-slate-200 placeholder:text-slate-500 leading-relaxed transition-colors scrollbar-dark disabled:opacity-70"
                />
              </div>
            )}
          </section>

          {/* Bottom PowerPoint Status Bar (25%-200% Zoom, View Switcher, Start Slideshow) */}
          <div className={!isReadOnly && viewMode === 'edit' ? 'pb-12 md:pb-0' : ''}>
            <OfficeStatusBar
              variant="powerpoint"
              theme="dark"
              zoomPercent={zoomPct}
              onZoomChange={setZoomPct}
              onFitToWidth={() => setZoomPct(100)}
              pageCurrent={safeSlideIdx + 1}
              pageTotal={slides.length}
              pptViewMode={pptViewMode}
              onChangePptViewMode={setPptViewMode}
              onStartSlideshow={() => startSlideshow(0)}
            />
          </div>

          {/* Floating Selection Formatting Toolbar (above selected text in slide frame) */}
          <FloatingSelectionToolbar
            containerRef={editorCanvasRef}
            disabled={
              isLaserActive ||
              isSlideshowActive ||
              isModalOpen ||
              isReadOnly ||
              viewMode === 'native'
            }
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

          {/* Mobile Collapsed 1-Row Bottom Bar (md:hidden, only in edit mode) */}
          {!isReadOnly && viewMode === 'edit' && (
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
          )}

          {/* Mobile Expanded Detail Bottom Sheet (md:hidden) */}
          {!isReadOnly && viewMode === 'edit' && isMobileSheetOpen && (
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

      {/* Thumbnail Right-Click Context Menu */}
      {thumbContextMenu && (
        <div
          role="menu"
          aria-label="Tùy chọn slide"
          onClick={(e) => e.stopPropagation()}
          style={{
            left: `${thumbContextMenu.x}px`,
            top: `${thumbContextMenu.y}px`,
          }}
          className="fixed z-50 w-48 rounded-xl bg-[#1C2331] border border-white/15 shadow-2xl py-1.5 text-xs text-slate-100"
        >
          <button
            type="button"
            disabled={isReadOnly}
            onClick={() => {
              setActiveSlideIdx(thumbContextMenu.slideIdx);
              handleAddSlide();
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#C43E1C]/25 flex items-center gap-2 disabled:opacity-40"
          >
            <Plus className="h-3.5 w-3.5 text-orange-300" />
            <span>Thêm slide mới</span>
          </button>
          <button
            type="button"
            disabled={isReadOnly}
            onClick={() => {
              handleDuplicateSlideAt(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#C43E1C]/25 flex items-center gap-2 disabled:opacity-40"
          >
            <Copy className="h-3.5 w-3.5 text-slate-300" />
            <span>Nhân bản slide</span>
          </button>
          <button
            type="button"
            disabled={isReadOnly}
            onClick={() => {
              handleToggleHideSlide(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#C43E1C]/25 flex items-center gap-2 disabled:opacity-40"
          >
            <EyeOff className="h-3.5 w-3.5 text-amber-300" />
            <span>
              {slides[thumbContextMenu.slideIdx]?.hidden
                ? 'Hiện slide khi chiếu'
                : 'Ẩn slide khi chiếu'}
            </span>
          </button>
          <div className="my-1 border-t border-white/10" />
          <button
            type="button"
            disabled={isReadOnly || slides.length <= 1}
            onClick={() => {
              handleDeleteSlideAt(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-rose-500/20 text-rose-200 flex items-center gap-2 disabled:opacity-35"
          >
            <Trash2 className="h-3.5 w-3.5" />
            <span>Xóa slide</span>
          </button>
        </div>
      )}

      {/* Insert Chart Modal */}
      <SlideChartInsertModal
        isOpen={isChartModalOpen}
        onClose={() => setIsChartModalOpen(false)}
        onInsertChart={handleInsertChart}
      />

      {/* Insert Hyperlink / Video / Audio Modal */}
      <SlideLinkMediaModal
        mode={linkMediaModalMode}
        onClose={() => setLinkMediaModalMode(null)}
        onInsertLink={handleInsertLinkBlock}
        onInsertMedia={handleInsertMediaBlock}
      />

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
@keyframes liteSlideWipeIn {
  from { clip-path: inset(0 100% 0 0); }
  to { clip-path: inset(0 0 0 0); }
}
@keyframes liteSlideSplitIn {
  from { clip-path: inset(0 50% 0 50%); opacity: 0.5; }
  to { clip-path: inset(0 0 0 0); opacity: 1; }
}
@keyframes liteSlideZoomIn {
  from { opacity: 0; transform: scale(0.84); }
  to { opacity: 1; transform: scale(1); }
}
@keyframes liteObjFadeIn {
  from { opacity: 0; }
  to { opacity: 1; }
}
@keyframes liteObjFlyIn {
  from { opacity: 0; transform: translateY(24px); }
  to { opacity: 1; transform: translateY(0); }
}
@keyframes liteObjZoomIn {
  from { opacity: 0; transform: scale(0.75); }
  to { opacity: 1; transform: scale(1); }
}
`}</style>

          {/* Proportional Fullscreen Slide Stage */}
          <div
            ref={slideshowCanvasRef}
            key={`slideshow-${currentSlide.id}-${safeSlideIdx}`}
            onClick={() => {
              if (justSwipedSlideshowRef.current) {
                justSwipedSlideshowRef.current = false;
                return;
              }
              if (
                isSlideshowLaser ||
                isSlideshowGridOpen ||
                slideshowDrawTool !== 'none'
              ) {
                return;
              }
              advanceSlideshowStep();
            }}
            onTouchStart={(event) => {
              if (
                isSlideshowLaser ||
                isSlideshowGridOpen ||
                slideshowDrawTool !== 'none' ||
                event.touches.length !== 1
              ) {
                slideshowTouchStartRef.current = null;
                return;
              }
              const touch = event.touches[0];
              slideshowTouchStartRef.current = {
                x: touch.clientX,
                y: touch.clientY,
              };
            }}
            onTouchEnd={(event) => {
              const start = slideshowTouchStartRef.current;
              slideshowTouchStartRef.current = null;
              if (
                !start ||
                isSlideshowLaser ||
                isSlideshowGridOpen ||
                slideshowDrawTool !== 'none' ||
                event.changedTouches.length === 0
              ) {
                return;
              }
              const touch = event.changedTouches[0];
              const dx = touch.clientX - start.x;
              const dy = touch.clientY - start.y;
              if (Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(dy) * 1.2) {
                justSwipedSlideshowRef.current = true;
                if (dx < 0) {
                  advanceSlideshowStep();
                } else {
                  retreatSlideshowStep();
                }
              }
            }}
            onPointerDown={(event) => {
              updateSlideshowLaserFromPointer(event);
              if (slideshowDrawTool === 'none' || isSlideshowGridOpen) {
                return;
              }
              event.stopPropagation();
              const rect = event.currentTarget.getBoundingClientRect();
              if (rect.width <= 0 || rect.height <= 0) {
                return;
              }
              const x = clampValue(((event.clientX - rect.left) / rect.width) * 100, 0, 100);
              const y = clampValue(((event.clientY - rect.top) / rect.height) * 100, 0, 100);
              if (slideshowDrawTool === 'eraser') {
                setSlideshowStrokesBySlide((prev) => {
                  const list = prev[currentSlide.id] ?? [];
                  const filtered = list.filter(
                    (st) => !st.points.some((pt) => Math.hypot(pt.x - x, pt.y - y) < 3.5)
                  );
                  return { ...prev, [currentSlide.id]: filtered };
                });
                return;
              }
              const strokeId = `ink-${Date.now()}-${Math.random()}`;
              activeSlideshowStrokeIdRef.current = strokeId;
              const newStroke: SlideshowStroke = {
                id: strokeId,
                tool: slideshowDrawTool,
                color: slideshowDrawTool === 'highlighter' ? '#FACC15' : '#EF4444',
                points: [{ x, y }],
              };
              setSlideshowStrokesBySlide((prev) => ({
                ...prev,
                [currentSlide.id]: [...(prev[currentSlide.id] ?? []), newStroke],
              }));
            }}
            onPointerMove={(event) => {
              updateSlideshowLaserFromPointer(event);
              if (slideshowDrawTool === 'none' || isSlideshowGridOpen) {
                return;
              }
              if (event.buttons === 0 && !activeSlideshowStrokeIdRef.current) {
                return;
              }
              const rect = event.currentTarget.getBoundingClientRect();
              if (rect.width <= 0 || rect.height <= 0) {
                return;
              }
              const x = clampValue(((event.clientX - rect.left) / rect.width) * 100, 0, 100);
              const y = clampValue(((event.clientY - rect.top) / rect.height) * 100, 0, 100);
              if (slideshowDrawTool === 'eraser') {
                if (event.buttons > 0) {
                  setSlideshowStrokesBySlide((prev) => {
                    const list = prev[currentSlide.id] ?? [];
                    const filtered = list.filter(
                      (st) => !st.points.some((pt) => Math.hypot(pt.x - x, pt.y - y) < 3.5)
                    );
                    return { ...prev, [currentSlide.id]: filtered };
                  });
                }
                return;
              }
              const activeId = activeSlideshowStrokeIdRef.current;
              if (!activeId) {
                return;
              }
              setSlideshowStrokesBySlide((prev) => ({
                ...prev,
                [currentSlide.id]: (prev[currentSlide.id] ?? []).map((st) =>
                  st.id === activeId
                    ? { ...st, points: [...st.points, { x, y }] }
                    : st
                ),
              }));
            }}
            onPointerUp={() => {
              activeSlideshowStrokeIdRef.current = null;
            }}
            onPointerCancel={() => {
              activeSlideshowStrokeIdRef.current = null;
            }}
            onPointerLeave={() => {
              activeSlideshowStrokeIdRef.current = null;
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
                : slideshowDrawTool !== 'none'
                  ? 'cursor-crosshair touch-none'
                  : 'cursor-pointer'
            }`}
          >
            <ReadOnlySlideStage
              doc={document}
              slide={currentSlide}
              slideIndex={safeSlideIdx}
              totalSlides={slides.length}
              stageWidthPx={actualSlideshowStageWidth}
              underlinedBlockIds={underlinedBlockIds}
              revealedAnimStep={revealedAnimStep}
            />

            {/* Slideshow Ink Annotations Layer (Pen & Highlighter) */}
            {(slideshowStrokesBySlide[currentSlide.id]?.length ?? 0) > 0 && (
              <svg
                viewBox="0 0 100 100"
                preserveAspectRatio="none"
                className="pointer-events-none absolute inset-0 w-full h-full z-20"
              >
                {(slideshowStrokesBySlide[currentSlide.id] ?? []).map((stroke) => (
                  <polyline
                    key={stroke.id}
                    fill="none"
                    stroke={stroke.color}
                    strokeWidth={stroke.tool === 'highlighter' ? 2.2 : 0.45}
                    strokeOpacity={stroke.tool === 'highlighter' ? 0.42 : 0.95}
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    points={stroke.points.map((p) => `${p.x},${p.y}`).join(' ')}
                  />
                ))}
              </svg>
            )}

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
                  <LayoutGrid className="h-5 w-5 text-[#C43E1C]" />
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
                      setRevealedAnimStep(0);
                      setActiveSlideIdx(idx);
                      setIsSlideshowGridOpen(false);
                    }}
                  />
                ))}
              </div>
            </div>
          )}

          {/* Bottom-Left Floating Translucent Slideshow Control Toolbar */}
          <div
            onClick={(event) => event.stopPropagation()}
            onPointerMove={(event) => event.stopPropagation()}
            className="fixed bottom-4 left-4 z-50 max-w-[94vw] overflow-x-auto no-scrollbar rounded-2xl bg-slate-900/75 hover:bg-slate-900/95 backdrop-blur-md border border-white/15 px-2.5 py-1.5 shadow-2xl flex items-center gap-1 text-xs text-slate-100 transition-colors"
          >
            <button
              type="button"
              onClick={retreatSlideshowStep}
              disabled={safeSlideIdx === 0 && revealedAnimStep === 0}
              aria-label="Slide trước"
              className="h-8 px-2 rounded-xl bg-white/10 hover:bg-white/20 disabled:opacity-35 disabled:cursor-not-allowed inline-flex items-center gap-1 font-medium whitespace-nowrap"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>

            <span className="px-1.5 font-semibold tabular-nums text-slate-200 whitespace-nowrap">
              {safeSlideIdx + 1} / {slides.length}
            </span>

            <button
              type="button"
              onClick={advanceSlideshowStep}
              disabled={
                safeSlideIdx >= slides.length - 1 &&
                revealedAnimStep >= getMaxAnimationStep(currentSlide)
              }
              aria-label="Slide tiếp"
              className="h-8 px-2 rounded-xl bg-white/10 hover:bg-white/20 disabled:opacity-35 disabled:cursor-not-allowed inline-flex items-center gap-1 font-medium whitespace-nowrap"
            >
              <ChevronRight className="h-4 w-4" />
            </button>

            <span className="h-4 w-px bg-white/15 shrink-0 mx-0.5" />

            {/* Red Laser Toggle */}
            <button
              type="button"
              onClick={() => {
                setIsSlideshowLaser((prev) => {
                  const next = !prev;
                  if (next) {
                    setSlideshowDrawTool('none');
                  }
                  return next;
                });
              }}
              aria-pressed={isSlideshowLaser}
              title="Bút Laser đỏ"
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
              <span className="hidden sm:inline">Laser</span>
            </button>

            {/* Pen Tool */}
            <button
              type="button"
              onClick={() => {
                setIsSlideshowLaser(false);
                setSlideshowDrawTool((prev) => (prev === 'pen' ? 'none' : 'pen'));
              }}
              aria-pressed={slideshowDrawTool === 'pen'}
              title="Bút vẽ (Đỏ)"
              className={`h-8 px-2.5 rounded-xl border inline-flex items-center gap-1 font-medium whitespace-nowrap transition-colors ${
                slideshowDrawTool === 'pen'
                  ? 'bg-[#C43E1C] border-orange-300 text-white'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <Pen className="h-3.5 w-3.5" />
              <span className="hidden md:inline">Bút</span>
            </button>

            {/* Highlighter Tool */}
            <button
              type="button"
              onClick={() => {
                setIsSlideshowLaser(false);
                setSlideshowDrawTool((prev) =>
                  prev === 'highlighter' ? 'none' : 'highlighter'
                );
              }}
              aria-pressed={slideshowDrawTool === 'highlighter'}
              title="Bút tô sáng (Highlighter)"
              className={`h-8 px-2.5 rounded-xl border inline-flex items-center gap-1 font-medium whitespace-nowrap transition-colors ${
                slideshowDrawTool === 'highlighter'
                  ? 'bg-amber-500/35 border-amber-300 text-amber-100'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <Highlighter className="h-3.5 w-3.5 text-amber-300" />
              <span className="hidden md:inline">Tô sáng</span>
            </button>

            {/* Eraser Tool */}
            <button
              type="button"
              onClick={() => {
                setIsSlideshowLaser(false);
                setSlideshowDrawTool((prev) =>
                  prev === 'eraser' ? 'none' : 'eraser'
                );
              }}
              aria-pressed={slideshowDrawTool === 'eraser'}
              title="Cục tẩy nét vẽ"
              className={`h-8 px-2 rounded-xl border inline-flex items-center gap-1 font-medium whitespace-nowrap transition-colors ${
                slideshowDrawTool === 'eraser'
                  ? 'bg-sky-500/30 border-sky-300 text-white'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <Eraser className="h-3.5 w-3.5" />
            </button>

            {(slideshowStrokesBySlide[currentSlide.id]?.length ?? 0) > 0 && (
              <button
                type="button"
                onClick={() =>
                  setSlideshowStrokesBySlide((prev) => ({
                    ...prev,
                    [currentSlide.id]: [],
                  }))
                }
                title="Xóa hết nét vẽ trên slide này"
                className="h-8 px-2 rounded-xl bg-white/[0.07] hover:bg-rose-500/25 border border-white/10 text-[11px] text-rose-200 whitespace-nowrap"
              >
                Xóa nét
              </button>
            )}

            {/* Slide Grid Toggle */}
            <button
              type="button"
              onClick={() => setIsSlideshowGridOpen((prev) => !prev)}
              aria-pressed={isSlideshowGridOpen}
              title="Nhảy nhanh tới slide (Grid)"
              className={`h-8 px-2.5 rounded-xl border inline-flex items-center gap-1.5 font-medium whitespace-nowrap transition-colors ${
                isSlideshowGridOpen
                  ? 'bg-[#C43E1C] border-orange-300 text-white'
                  : 'bg-white/[0.07] hover:bg-white/15 border-white/10 text-slate-200'
              }`}
            >
              <LayoutGrid className="h-3.5 w-3.5" />
              <span className="hidden sm:inline">Lưới</span>
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

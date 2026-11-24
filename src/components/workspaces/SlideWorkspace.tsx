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
  AlignJustify,
  AlignLeft,
  AlignRight,
  ArrowDownToLine,
  ArrowUpToLine,
  BarChart3,
  Bold,
  BringToFront,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Clipboard,
  CloudDownload,
  Columns,
  Copy,
  CopyPlus,
  Crosshair,
  Eraser,
  ExternalLink,
  Eye,
  EyeOff,
  FileDown,
  FileOutput,
  FileText,
  Group,
  Hash,
  Highlighter,
  History,
  Image,
  ImagePlus,
  IndentDecrease,
  IndentIncrease,
  Info,
  Italic,
  Layers,
  LayoutGrid,
  LayoutTemplate,
  Link2,
  List,
  ListOrdered,
  Loader2,
  Lock,
  MessageSquare,
  MonitorPlay,
  Move,
  Music,
  PaintBucket,
  Paintbrush,
  Palette,
  Pause,
  Pen,
  Play,
  Plus,
  Presentation,
  Printer,
  QrCode,
  RefreshCw,
  RotateCcw,
  RotateCw,
  Save,
  Scissors,
  SendToBack,
  Shapes,
  ShieldCheck,
  Sparkles,
  StickyNote,
  Strikethrough,
  Subscript,
  Superscript,
  Table2,
  Trash2,
  Type,
  Underline,
  Undo2,
  Ungroup,
  Unlock,
  Video,
  Wand2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  SlideAnimationConfig,
  SlideChartData,
  SlideCommentItem,
  SlideDrawingStroke,
  SlideImageItem,
  SlidePreviewItem,
  SlideTextBlock,
  SlideTransitionConfig,
} from '../../types/document';
import {
  exportOriginalDocumentBinary,
  exportSlidesToPptx,
} from '../../utils/fileExporter';
import { downloadDocumentAsPdf } from '../../utils/quickConverter';
import { useWorkspaceSync } from '../../contexts/WorkspaceSyncContext';
import { syncAndroidStatusBar } from '../../utils/androidSystemUi';
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
  PPT_QUICK_SHAPES_10,
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
import {
  SlideAiToolsModal,
  SlideCommentsPanel,
  SlideDrawingOverlay,
  SlideFileInfoModal,
} from './SlideOnlyOfficePanels';
import { WorkspaceTopBar } from './WorkspaceTopBar';
import { AIToolsDrawer } from '../ai/AIToolsDrawer';

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
type PptRibbonTab =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'design'
  | 'transitions'
  | 'animations'
  | 'collab'
  | 'protect'
  | 'view'
  | 'ai';
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
  { id: 'file', label: 'Tệp' },
  { id: 'home', label: 'Trang chủ' },
  { id: 'insert', label: 'Chèn' },
  { id: 'draw', label: 'Vẽ' },
  { id: 'design', label: 'Thiết kế' },
  { id: 'transitions', label: 'Chuyển cảnh' },
  { id: 'animations', label: 'Hoạt hình' },
  { id: 'collab', label: 'Cộng tác' },
  { id: 'protect', label: 'Bảo vệ' },
  { id: 'view', label: 'Xem' },
  { id: 'ai', label: 'AI Tools' },
];

const PPT_TRANSITION_OPTIONS: ReadonlyArray<{
  type: SlideTransitionConfig['type'];
  label: string;
}> = [
  { type: 'none', label: 'Không có' },
  { type: 'fade', label: 'Mờ dần' },
  { type: 'push', label: 'Đẩy sang' },
  { type: 'wipe', label: 'Gạt ngang' },
  { type: 'split', label: 'Tách đôi' },
  { type: 'zoom', label: 'Thu phóng' },
];

const PPT_ANIMATION_OPTIONS: ReadonlyArray<{
  type: SlideAnimationConfig['type'];
  label: string;
}> = [
  { type: 'none', label: 'Không có' },
  { type: 'appear', label: 'Xuất hiện' },
  { type: 'fadeIn', label: 'Mờ dần' },
  { type: 'flyIn', label: 'Bay vào' },
  { type: 'zoomIn', label: 'Phóng vào' },
  { type: 'pulse', label: 'Nhịp đập' },
];

const PPT_HIGHLIGHT_SWATCHES = [
  { label: 'Vàng chanh', value: '#FEF08A' },
  { label: 'Xanh lá sáng', value: '#BBF7D0' },
  { label: 'Xanh dương nhạt', value: '#BAE6FD' },
  { label: 'Hồng phấn', value: '#FBCFE8' },
  { label: 'Cam nhạt', value: '#FED7AA' },
] as const;

const PPT_DRAW_COLORS = [
  '#EF4444',
  '#C43E1C',
  '#F59E0B',
  '#10B981',
  '#2563EB',
  '#7C3AED',
  '#0F172A',
] as const;

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
  { label: 'Arial', value: 'Arial, Helvetica, sans-serif' },
  { label: 'Calibri', value: 'Calibri, Candara, Segoe UI, sans-serif' },
  { label: 'Georgia', value: 'Georgia, serif' },
  { label: 'Times New Roman', value: "'Times New Roman', serif" },
  { label: 'Mono', value: "'JetBrains Mono', monospace" },
] as const;

function buildTextDecorationCss(
  underline?: boolean,
  strikethrough?: boolean
): string {
  const parts: string[] = [];
  if (underline) parts.push('underline');
  if (strikethrough) parts.push('line-through');
  return parts.length > 0 ? parts.join(' ') : 'none';
}

function buildStaticSlideStrokePath(points: Array<{ x: number; y: number }>): string {
  if (points.length === 0) return '';
  if (points.length === 1) {
    const p = points[0];
    return `M ${p.x} ${p.y} L ${p.x + 0.2} ${p.y + 0.2}`;
  }
  const [first, ...rest] = points;
  return `M ${first.x} ${first.y} ${rest.map((p) => `L ${p.x} ${p.y}`).join(' ')}`;
}

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
  if (anim.type === 'pulse') {
    return { animation: 'liteObjPulse 560ms ease-in-out both' };
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
const SLIDE_FONT_SIZES = [12, 14, 18, 24, 32, 40, 48] as const;

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
  strikethrough?: boolean;
  highlightColor?: string;
  textScript?: 'normal' | 'super' | 'sub';
  lineHeight?: number;
  textColumns?: 1 | 2;
  color: string;
  align: 'left' | 'center' | 'right' | 'justify';
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
    const scriptScale =
      block.textScript === 'super' || block.textScript === 'sub' ? 0.72 : 1;
    const scaledPx = Math.max(4, (block.fontSize || 18) * effectiveScale * scriptScale);
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
          className="whitespace-pre-wrap break-words tracking-tight"
          style={{
            fontSize: `${scaledPx.toFixed(2)}px`,
            fontWeight: block.bold || block.role === "title" ? 700 : 400,
            fontStyle: block.italic ? "italic" : "normal",
            textDecoration: buildTextDecorationCss(
              Boolean(underlined || block.underline),
              block.strikethrough
            ),
            backgroundColor: block.highlightColor,
            lineHeight: block.lineHeight ?? 1.18,
            columnCount: block.textColumns === 2 ? 2 : undefined,
            columnGap: block.textColumns === 2 ? '1em' : undefined,
            transform:
              block.textScript === 'super'
                ? 'translateY(-0.22em)'
                : block.textScript === 'sub'
                  ? 'translateY(0.18em)'
                  : undefined,
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
      style={{
        columnCount: block.textColumns === 2 ? 2 : undefined,
        columnGap: block.textColumns === 2 ? '1em' : undefined,
      }}
      className="w-full max-h-full overflow-hidden space-y-[0.22em]"
    >
      {paragraphs.map((para, pIdx) => {
        const firstRun = para.runs[0];
        const bulletPx = Math.max(4, ((firstRun?.fontSize ?? block.fontSize) || 16) * effectiveScale);
        return (
          <p
            key={`${block.id}-rp-${pIdx}`}
            className="whitespace-pre-wrap break-words tracking-tight"
            style={{
              textAlign: para.align ?? block.align ?? "left",
              lineHeight: block.lineHeight ?? 1.18,
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
                    textDecoration: buildTextDecorationCss(
                      Boolean(run.underline || underlined || block.underline),
                      block.strikethrough
                    ),
                    backgroundColor: block.highlightColor,
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
  strikethrough = false,
  highlightColor,
  textScript = 'normal',
  lineHeight,
  textColumns = 1,
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
  }, [value, scaledFontPx, bold, italic, underline, strikethrough, lineHeight, align, fontFamily]);

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

  const effectiveScriptFontPx =
    textScript === 'super' || textScript === 'sub'
      ? Math.max(4, scaledFontPx * 0.72)
      : scaledFontPx;

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
        fontSize: `${effectiveScriptFontPx.toFixed(2)}px`,
        color,
        textAlign: align,
        fontWeight: bold ? 700 : 400,
        fontStyle: italic ? 'italic' : 'normal',
        textDecoration: buildTextDecorationCss(underline, strikethrough),
        backgroundColor:
          highlightColor && highlightColor !== 'transparent'
            ? highlightColor
            : undefined,
        lineHeight: lineHeight ?? 1.16,
        columnCount: textColumns === 2 ? 2 : undefined,
        columnGap: textColumns === 2 ? '1em' : undefined,
        transform:
          textScript === 'super'
            ? 'translateY(-0.22em)'
            : textScript === 'sub'
              ? 'translateY(0.18em)'
              : undefined,
        fontFamily: fontFamily || "'Inter', 'Segoe UI', system-ui, sans-serif",
        maxHeight: maxHeightFull ? '100%' : undefined,
      }}
      className={`w-full bg-transparent border-0 outline-none resize-none tracking-tight placeholder:opacity-45 disabled:pointer-events-none ${
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
  onActivateEditBlock?: (blockId?: string) => void;
}

function ReadOnlySlideStage({
  doc,
  slide,
  slideIndex,
  totalSlides,
  stageWidthPx,
  underlinedBlockIds,
  revealedAnimStep,
  onActivateEditBlock,
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
              onDoubleClick={
                onActivateEditBlock
                  ? (e) => {
                      e.stopPropagation();
                      onActivateEditBlock(block.id);
                    }
                  : undefined
              }
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
                  underlined={Boolean(
                    underlinedBlockIds?.[block.id] || block.underline || safeLink
                  )}
                />
              ) : (
                <p
                  className="whitespace-pre-wrap break-words tracking-tight"
                  style={{
                    fontSize: `${Math.max(
                      3.5,
                      scaledFontPx *
                        ((block.fontScalePct ?? 100) / 100) *
                        (block.textScript === 'super' || block.textScript === 'sub'
                          ? 0.72
                          : 1)
                    ).toFixed(2)}px`,
                    fontWeight:
                      block.bold || block.role === 'title' ? 700 : 400,
                    fontStyle: block.italic ? 'italic' : 'normal',
                    textDecoration: buildTextDecorationCss(
                      Boolean(
                        underlinedBlockIds?.[block.id] || block.underline || safeLink
                      ),
                      block.strikethrough
                    ),
                    backgroundColor: block.highlightColor,
                    lineHeight: block.lineHeight ?? 1.16,
                    columnCount: block.textColumns === 2 ? 2 : undefined,
                    columnGap: block.textColumns === 2 ? '1em' : undefined,
                    transform:
                      block.textScript === 'super'
                        ? 'translateY(-0.22em)'
                        : block.textScript === 'sub'
                          ? 'translateY(0.18em)'
                          : undefined,
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

        {/* Static Vector Drawings Layer (Pen / Highlighter strokes) */}
        {(slide.drawings?.length ?? 0) > 0 && (
          <svg
            viewBox="0 0 1000 1000"
            preserveAspectRatio="none"
            className="pointer-events-none absolute inset-0 w-full h-full z-20"
          >
            {(slide.drawings ?? []).map((stroke) => {
              const isHighlighter = stroke.tool === 'highlighter';
              return (
                <path
                  key={stroke.id}
                  d={buildStaticSlideStrokePath(stroke.points)}
                  fill="none"
                  stroke={stroke.color}
                  strokeWidth={
                    isHighlighter
                      ? Math.max(14, stroke.width * 2.5)
                      : Math.max(1.5, stroke.width)
                  }
                  strokeOpacity={isHighlighter ? 0.36 : 0.95}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              );
            })}
          </svg>
        )}

        {/* Slide Footer & Number Strip */}
        {(slide.footerText || slide.showSlideNumber) && (
          <div
            className={`pointer-events-none absolute inset-x-0 bottom-0 z-20 flex items-center justify-between gap-2 px-3 py-1 ${
              hasLightBg ? 'text-slate-600' : 'text-slate-300/90'
            }`}
            style={{ fontSize: `${Math.max(6, 10.5 * fontScale).toFixed(1)}px` }}
          >
            <span className="truncate">{slide.footerText ?? ''}</span>
            {slide.showSlideNumber && (
              <span className="shrink-0 font-semibold tabular-nums">
                {slideIndex + 1} / {totalSlides}
              </span>
            )}
          </div>
        )}
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
                className="whitespace-pre-wrap break-words"
                style={{
                  fontSize: `${Math.max(5, (block.fontSize || 18) * fontScale)}px`,
                  fontWeight: block.bold || block.role === 'title' ? 700 : 400,
                  fontStyle: block.italic ? 'italic' : 'normal',
                  textDecoration: buildTextDecorationCss(
                    Boolean(block.underline || underlinedBlockIds?.[block.id]),
                    Boolean(block.strikethrough)
                  ),
                  backgroundColor: block.highlightColor || 'transparent',
                  padding: block.highlightColor ? '1px 4px' : undefined,
                  borderRadius: block.highlightColor ? '3px' : undefined,
                  lineHeight: block.lineHeight ?? 1.16,
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
        {(slide.footerText || slide.showSlideNumber) ? (
          <div
            className={`flex items-center gap-2 truncate max-w-[60%] ${
              hasLightBg ? 'text-slate-500' : 'text-slate-400'
            }`}
            style={{ fontSize: `${Math.max(5, 10.5 * fontScale)}px` }}
          >
            {slide.footerText && <span className="truncate">{slide.footerText}</span>}
            {slide.showSlideNumber && (
              <span className="shrink-0 font-semibold tabular-nums">
                {slideIndex + 1} / {totalSlides}
              </span>
            )}
          </div>
        ) : (
          <span
            className={`truncate max-w-[45%] ${
              hasLightBg ? 'text-slate-500' : 'text-slate-400'
            }`}
            style={{ fontSize: `${Math.max(5, 11 * fontScale)}px` }}
          >
            {doc.title}
          </span>
        )}
      </div>

      {/* Static Ink Drawings Overlay (Legacy Flow Stage) */}
      {(slide.drawings?.length ?? 0) > 0 && (
        <svg
          viewBox="0 0 100 100"
          preserveAspectRatio="none"
          className="pointer-events-none absolute inset-0 z-20 w-full h-full"
        >
          {slide.drawings!.map((stroke) => {
            const d = buildStaticSlideStrokePath(stroke.points);
            if (!d) return null;
            const isHighlighter = stroke.tool === 'highlighter';
            return (
              <path
                key={stroke.id}
                d={d}
                fill="none"
                stroke={stroke.color}
                strokeWidth={
                  isHighlighter
                    ? Math.max(14, stroke.width * 2.5)
                    : Math.max(1.5, stroke.width)
                }
                strokeOpacity={isHighlighter ? 0.36 : 0.95}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            );
          })}
        </svg>
      )}
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
  const workspaceSync = useWorkspaceSync();
  const [isPresentationProtected, setIsPresentationProtected] =
    useState<boolean>(() => Boolean(document.slideProtected));
  const isEditingLocked = isReadOnly || isPresentationProtected;
  const hasOriginalBinary = Boolean(
    document.rawArrayBuffer ||
      document.isUploaded ||
      document.storagePath ||
      document.r2ObjectKey
  );
  const [viewMode, setViewMode] = useState<SlideWorkspaceViewMode>(() =>
    collabPermission === 'view' ? 'native' : 'edit'
  );
  const [hasLocalEdits, setHasLocalEdits] = useState<boolean>(() =>
    Boolean(document.hasLocalEdits)
  );
  const [slides, setSlides] = useState<SlidePreviewItem[]>(() =>
    initializeDeckSlides(document)
  );
  const [deckAspectRatio, setDeckAspectRatio] = useState<number | undefined>(
    () => document.slideAspectRatio
  );
  const [activeSlideIdx, setActiveSlideIdx] = useState<number>(0);
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null);
  const [editingBlockId, setEditingBlockId] = useState<string | null>(null);
  const [zoomPct, setZoomPct] = useState<number>(100);
  const setZoomPctClamped = useCallback((nextZoom: number) => {
    setZoomPct(Math.max(10, Math.min(300, Math.round(nextZoom))));
  }, []);
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
  const [isRibbonCollapsed, setIsRibbonCollapsed] = useState<boolean>(false);
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
  const [layoutPickerMode, setLayoutPickerMode] = useState<
    'add' | 'change' | null
  >(null);
  const [isShapePickerOpen, setIsShapePickerOpen] = useState<boolean>(false);
  const [isCaseMenuOpen, setIsCaseMenuOpen] = useState<boolean>(false);
  const [isHighlightMenuOpen, setIsHighlightMenuOpen] =
    useState<boolean>(false);
  const [clipboardItem, setClipboardItem] = useState<
    | { kind: 'block'; data: SlideTextBlock }
    | { kind: 'image'; data: SlideImageItem }
    | null
  >(null);
  const [copiedBlockStyle, setCopiedBlockStyle] =
    useState<Partial<SlideTextBlock> | null>(null);
  const [activeDrawTool, setActiveDrawTool] = useState<
    'pen' | 'highlighter' | 'eraser' | null
  >(null);
  const [drawColor, setDrawColor] = useState<string>('#EF4444');
  const [drawWidth, setDrawWidth] = useState<number>(4);
  const [previewTransitionKey, setPreviewTransitionKey] = useState<number>(0);
  const [previewAnimKey, setPreviewAnimKey] = useState<number>(0);
  const [isPreviewingAnims, setIsPreviewingAnims] = useState<boolean>(false);
  const [isCommentsOpen, setIsCommentsOpen] = useState<boolean>(false);
  const [isFileInfoOpen, setIsFileInfoOpen] = useState<boolean>(false);
  const [aiModalState, setAiModalState] = useState<{
    isOpen: boolean;
    initialTab: 'generate' | 'summarize' | 'notes';
  }>({ isOpen: false, initialTab: 'generate' });
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
      setDeckAspectRatio(document.slideAspectRatio);
      setIsPresentationProtected(Boolean(document.slideProtected));
      setViewMode(collabPermission === 'view' ? 'native' : 'edit');
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
      setLayoutPickerMode(null);
      setIsShapePickerOpen(false);
      setIsCaseMenuOpen(false);
      setIsHighlightMenuOpen(false);
      setCopiedBlockStyle(null);
      setActiveDrawTool(null);
      setPreviewTransitionKey(0);
      setIsPreviewingAnims(false);
      setIsCommentsOpen(false);
      setIsFileInfoOpen(false);
      setAiModalState({ isOpen: false, initialTab: 'generate' });
      setRevealedAnimStep(0);
      setSlideshowDrawTool('none');
      setSlideshowStrokesBySlide({});
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

  // Auto-reset active draw tool outside Draw tab or when locked/in Sorter/Slideshow/Native
  useEffect(() => {
    if (
      activeRibbonTab !== 'draw' ||
      isEditingLocked ||
      pptViewMode === 'sorter' ||
      isSlideshowActive ||
      viewMode === 'native'
    ) {
      setActiveDrawTool(null);
    }
  }, [activeRibbonTab, isEditingLocked, pptViewMode, isSlideshowActive, viewMode]);

  // Clear selection/mobile sheet and force native view if switched to read-only permission
  useEffect(() => {
    if (isReadOnly) {
      setViewMode('native');
      setSelectedBlockId(null);
      setEditingBlockId(null);
      setSelectedImageId(null);
      setIsMobileSheetOpen(false);
      setCopiedBlockStyle(null);
    }
  }, [isReadOnly]);

  // Pinch / Ctrl+Wheel zoom strictly scoped to the slide viewport (10% - 300%)
  useDocumentPinchZoom(editorViewportRef, zoomPct, setZoomPctClamped, {
    min: 10,
    max: 300,
    disabled:
      isSlideshowActive ||
      isModalOpen ||
      isChartModalOpen ||
      isFileInfoOpen ||
      aiModalState.isOpen ||
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
  const accentColor = currentSlide.accentColor || '#C43E1C';
  const hasLightSlideBg = isLightHexBackground(currentSlide.backgroundColor);
  const currentAspectRatio =
    deckAspectRatio ?? getSafeAspectRatio(currentSlide, document);
  const isPositionedSlide = isSlideMasterPositioned(currentSlide);

  const editorFitWidthPx = Math.max(
    160,
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

  const handleFitToWidth = useCallback(() => {
    const vp = editorViewportRef.current;
    if (!vp || editorFitWidthPx <= 0) {
      setZoomPctClamped(100);
      return;
    }
    const vw = Math.max(120, vp.clientWidth - 16);
    const vh = Math.max(100, vp.clientHeight - 16);
    const naturalH = editorFitWidthPx / currentAspectRatio;
    const fitScale = Math.min(vw / editorFitWidthPx, vh / naturalH);
    setZoomPctClamped(Math.floor(fitScale * 100));
  }, [currentAspectRatio, editorFitWidthPx, setZoomPctClamped]);

  const handleSwitchViewMode = useCallback(
    (nextMode: SlideWorkspaceViewMode) => {
      if (nextMode === 'native') {
        setSelectedBlockId(null);
        setSelectedBlockIds([]);
        setEditingBlockId(null);
        setSelectedImageId(null);
        setIsMobileSheetOpen(false);
        setViewMode('native');
        return;
      }
      if (!isEditingLocked) {
        setViewMode('edit');
      }
    },
    [isEditingLocked]
  );

  const handleUnlockDirectEdit = useCallback(
    (blockId?: string) => {
      if (isEditingLocked) {
        return;
      }
      setViewMode('edit');
      if (blockId) {
        setSelectedBlockId(blockId);
        setSelectedBlockIds([blockId]);
        setEditingBlockId(blockId);
        setSelectedImageId(null);
      }
    },
    [isEditingLocked]
  );

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
      slideAspectRatio: deckAspectRatio ?? baseDoc.slideAspectRatio,
      slideProtected: isPresentationProtected,
      hasLocalEdits: hasLocalEdits || Boolean(baseDoc.hasLocalEdits),
      contentPreview:
        syncedSlides[0]?.bullets[0] ??
        syncedSlides[0]?.title ??
        baseDoc.contentPreview,
    };
  }, [deckAspectRatio, hasLocalEdits, isPresentationProtected, slides]);

  const handleSave = useCallback(() => {
    if (isEditingLocked) {
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
  }, [buildLiveDocument, isEditingLocked, onSaveDocument]);

  const handleRenameTitle = useCallback(
    (newTitle: string) => {
      if (isEditingLocked) {
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
    [accentColor, buildLiveDocument, document.type, isEditingLocked, onSaveDocument]
  );

  // Auto-save with 1800ms debounce when autoSaveEnabled is on
  useEffect(() => {
    if (isEditingLocked || !autoSaveEnabled || !isDirty) {
      return;
    }
    const timerId = window.setTimeout(() => {
      handleSave();
    }, 1800);
    return () => {
      window.clearTimeout(timerId);
    };
  }, [autoSaveEnabled, isDirty, isEditingLocked, slides, handleSave]);

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
      if (
        isModalOpen ||
        isChartModalOpen ||
        isFileInfoOpen ||
        aiModalState.isOpen ||
        linkMediaModalMode !== null
      ) {
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
        if (!isEditingLocked) {
          handleSave();
        }
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        handlePrintDeck();
        return;
      }

      if (
        (event.ctrlKey || event.metaKey) &&
        !event.shiftKey &&
        event.key.toLowerCase() === 'z' &&
        !isEditableElement(event.target)
      ) {
        event.preventDefault();
        handleUndoSlideStroke();
        return;
      }

      if (event.key === 'Escape') {
        if (thumbContextMenu) {
          setThumbContextMenu(null);
          return;
        }
        if (copiedBlockStyle) {
          setCopiedBlockStyle(null);
          return;
        }
        if (activeDrawTool) {
          setActiveDrawTool(null);
          return;
        }
        if (
          layoutPickerMode !== null ||
          isShapePickerOpen ||
          isCaseMenuOpen ||
          isHighlightMenuOpen
        ) {
          setLayoutPickerMode(null);
          setIsShapePickerOpen(false);
          setIsCaseMenuOpen(false);
          setIsHighlightMenuOpen(false);
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

    const handleShortcutSlideshow = () => {
      if (!isSlideshowActive) {
        startSlideshow(0);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('lite:shortcut-slideshow', handleShortcutSlideshow);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('lite:shortcut-slideshow', handleShortcutSlideshow);
    };
  }, [
    slides.length,
    safeSlideIdx,
    isMobileSheetOpen,
    isMobileDrawerOpen,
    isLaserActive,
    isModalOpen,
    isChartModalOpen,
    isFileInfoOpen,
    aiModalState.isOpen,
    linkMediaModalMode,
    layoutPickerMode,
    isShapePickerOpen,
    isCaseMenuOpen,
    isHighlightMenuOpen,
    copiedBlockStyle,
    activeDrawTool,
    thumbContextMenu,
    isSlideshowActive,
    isSlideshowGridOpen,
    selectedBlockId,
    selectedBlockIds.length,
    selectedImageId,
    isEditingLocked,
    startSlideshow,
    advanceSlideshowStep,
    retreatSlideshowStep,
    handleSave,
    exitSlideshow,
  ]);

  useEffect(() => {
    if (isSlideshowActive) {
      syncAndroidStatusBar({ color: '#000000', isDarkBackground: true });
    } else {
      syncAndroidStatusBar({ color: '#FFFFFF', isDarkBackground: false });
    }
  }, [isSlideshowActive]);

  // Keep active slide thumbnail smoothly scrolled into view
  useEffect(() => {
    const node = thumbnailRefs.current[currentSlide.id];
    node?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, [currentSlide.id]);

  const mutateSlideAt = useCallback(
    (targetIdx: number, updater: (slide: SlidePreviewItem) => SlidePreviewItem) => {
      if (isEditingLocked) {
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
    [isEditingLocked, notifyTypingActivity]
  );

  const handleAddSlide = () => {
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked || slides.length <= 1) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
      return;
    }
    const targetBlock = currentBlocks.find((b) => b.id === blockId);
    const nextUnderline = !(
      Boolean(targetBlock?.underline) || Boolean(underlinedBlockIds[blockId])
    );
    handleUpdateBlock(blockId, { underline: nextUnderline });
    setUnderlinedBlockIds((prev) => ({
      ...prev,
      [blockId]: nextUnderline,
    }));
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
      underline: false,
      strikethrough: false,
      highlightColor: undefined,
      textScript: 'normal',
      lineHeight: undefined,
      textColumns: 1,
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
    if (isEditingLocked) {
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
    if (isEditingLocked || slides.length <= 1) {
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
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(targetIdx, (slide) => ({
      ...slide,
      hidden: !slide.hidden,
    }));
  };

  const handleReorderSlide = (fromIdx: number, toIdx: number) => {
    if (isEditingLocked || fromIdx === toIdx) {
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
    if (isEditingLocked) {
      return;
    }
    setViewMode('edit');
    const mode = layoutPickerMode ?? 'change';
    setLayoutPickerMode(null);
    const isLight = isLightHexBackground(currentSlide.backgroundColor);
    if (mode === 'add') {
      uidRef.current += 1;
      const newSlideId = `${document.id}-lay-slide-${Date.now()}-${uidRef.current}`;
      const layBlocks = createBlocksForSlideLayout(layoutId, newSlideId, {
        titleColor: isLight ? '#0F172A' : '#FFFFFF',
        bodyColor: isLight ? '#334155' : '#E2E8F0',
        accentColor,
      });
      const newSlide: SlidePreviewItem = syncSlideSummary(
        {
          id: newSlideId,
          index: safeSlideIdx + 2,
          title: 'Slide mới',
          bullets: [],
          accentColor,
          speakerNote: '',
          isMasterPositioned: true,
          slideWidthRatio: currentAspectRatio,
          backgroundColor: currentSlide.backgroundColor ?? '#161D2E',
          backgroundGradient: currentSlide.backgroundGradient,
          blocks: layBlocks,
        },
        safeSlideIdx + 1
      );
      localMutationPendingRef.current = true;
      setHasLocalEdits(true);
      setSlides((prev) => {
        const insertAt = Math.min(safeSlideIdx + 1, prev.length);
        const next = [...prev.slice(0, insertAt), newSlide, ...prev.slice(insertAt)];
        return next.map((item, idx) => ({ ...item, index: idx + 1 }));
      });
      setActiveSlideIdx(safeSlideIdx + 1);
      setSelectedBlockId(layBlocks[0]?.id ?? null);
      setSelectedBlockIds(layBlocks[0] ? [layBlocks[0].id] : []);
      setSelectedImageId(null);
      setIsDirty(true);
      setSaveFeedback(false);
      return;
    }

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

  const handleResetCurrentSlide = () => {
    if (isEditingLocked) {
      return;
    }
    const isLight = isLightHexBackground(currentSlide.backgroundColor);
    const existingTitle =
      currentBlocks.find((b) => b.role === 'title')?.text || currentSlide.title;
    const existingBody = currentBlocks
      .filter((b) => b.role !== 'title' && b.kind !== 'shape' && b.kind !== 'chart' && b.kind !== 'media')
      .map((b) => b.text)
      .filter(Boolean)
      .join('\n');
    const resetBlocks = createBlocksForSlideLayout(
      'title_content',
      `${currentSlide.id}-reset-${Date.now()}`,
      {
        titleColor: isLight ? '#0F172A' : '#FFFFFF',
        bodyColor: isLight ? '#334155' : '#E2E8F0',
        accentColor,
      }
    );
    if (resetBlocks[0] && existingTitle) {
      resetBlocks[0] = { ...resetBlocks[0], text: existingTitle };
    }
    if (resetBlocks[1] && existingBody) {
      resetBlocks[1] = { ...resetBlocks[1], text: existingBody };
    }
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      isMasterPositioned: true,
      blocks: resetBlocks,
    }));
    setSelectedBlockId(resetBlocks[0]?.id ?? null);
    setSelectedBlockIds(resetBlocks[0] ? [resetBlocks[0].id] : []);
  };

  const handleInsertShape = (shapeKind: PptShapeKind) => {
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
      return;
    }
    setViewMode('edit');
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((slide, idx) =>
        syncSlideSummary(
          {
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
          },
          idx
        )
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
    notifyTypingActivity(true);
  };

  const handleApplySlideGradient = (gradient: string) => {
    if (isEditingLocked) {
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
    if (!file || isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
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
    if (isEditingLocked) {
      return;
    }
    const maxZ = currentBlocks.reduce(
      (max, b) => Math.max(max, b.zIndex ?? 0),
      0
    );
    handleUpdateBlock(blockId, { zIndex: maxZ + 1 });
  };

  const handleSendToBack = (blockId: string) => {
    if (isEditingLocked) {
      return;
    }
    const minZ = currentBlocks.reduce(
      (min, b) => Math.min(min, b.zIndex ?? 0),
      0
    );
    handleUpdateBlock(blockId, { zIndex: minZ - 1 });
  };

  const handleGroupSelected = () => {
    if (isEditingLocked || selectedBlockIds.length < 2) {
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
    if (isEditingLocked || !selectedBlock?.groupId) {
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

  // --- ONLYOFFICE 11-Tab Pro Suite Handlers ---

  const handleCopyObject = () => {
    if (selectedBlock) {
      setClipboardItem({
        kind: 'block',
        data: {
          ...selectedBlock,
          tableRows: selectedBlock.tableRows
            ? selectedBlock.tableRows.map((r) => [...r])
            : undefined,
        },
      });
    } else if (selectedImage) {
      setClipboardItem({ kind: 'image', data: { ...selectedImage } });
    }
  };

  const handleCutObject = () => {
    if (isEditingLocked) {
      return;
    }
    handleCopyObject();
    if (selectedBlock) {
      handleDeleteBlock(selectedBlock.id);
    } else if (selectedImage) {
      handleDeleteImage(selectedImage.id);
    }
  };

  const handlePasteObject = () => {
    if (isEditingLocked || !clipboardItem) {
      return;
    }
    setViewMode('edit');
    uidRef.current += 1;
    if (clipboardItem.kind === 'block') {
      const pasteId = `${currentSlide.id}-paste-${Date.now()}-${uidRef.current}`;
      const pasted: SlideTextBlock = {
        ...clipboardItem.data,
        id: pasteId,
        xPct: Math.min(80, (clipboardItem.data.xPct ?? 12) + 3),
        yPct: Math.min(80, (clipboardItem.data.yPct ?? 18) + 3),
        zIndex: currentBlocks.length + 1,
      };
      mutateSlideAt(safeSlideIdx, (s) => ({
        ...s,
        isMasterPositioned: true,
        blocks: [...(s.blocks ?? []), pasted],
      }));
      setSelectedBlockId(pasteId);
      setSelectedBlockIds([pasteId]);
      setSelectedImageId(null);
    } else {
      const pasteImgId = `${currentSlide.id}-paste-img-${Date.now()}-${uidRef.current}`;
      const pastedImg: SlideImageItem = {
        ...clipboardItem.data,
        id: pasteImgId,
        xPct: Math.min(75, (clipboardItem.data.xPct ?? 24) + 3),
        yPct: Math.min(75, (clipboardItem.data.yPct ?? 20) + 3),
      };
      mutateSlideAt(safeSlideIdx, (s) => ({
        ...s,
        isMasterPositioned: true,
        images: [...(s.images ?? []), pastedImg],
      }));
      setSelectedImageId(pasteImgId);
      setSelectedBlockId(null);
    }
  };

  const handleToggleFormatPainter = () => {
    if (isEditingLocked) {
      return;
    }
    if (copiedBlockStyle) {
      setCopiedBlockStyle(null);
      return;
    }
    if (!selectedBlock) {
      return;
    }
    setCopiedBlockStyle({
      fontSize: selectedBlock.fontSize,
      fontFamily: selectedBlock.fontFamily,
      bold: selectedBlock.bold,
      italic: selectedBlock.italic,
      underline: Boolean(
        selectedBlock.underline || underlinedBlockIds[selectedBlock.id]
      ),
      strikethrough: selectedBlock.strikethrough,
      color: selectedBlock.color,
      highlightColor: selectedBlock.highlightColor,
      align: selectedBlock.align,
      valign: selectedBlock.valign,
      lineHeight: selectedBlock.lineHeight,
      textColumns: selectedBlock.textColumns,
      fillColor: selectedBlock.fillColor,
      borderColor: selectedBlock.borderColor,
    });
  };

  const applyCopiedStyleToBlock = (targetBlockId: string): boolean => {
    if (!copiedBlockStyle || isEditingLocked) {
      return false;
    }
    handleUpdateBlock(targetBlockId, copiedBlockStyle);
    if (typeof copiedBlockStyle.underline === 'boolean') {
      setUnderlinedBlockIds((prev) => ({
        ...prev,
        [targetBlockId]: Boolean(copiedBlockStyle.underline),
      }));
    }
    setCopiedBlockStyle(null);
    return true;
  };

  const handleTransformBlockCase = (
    mode: 'upper' | 'lower' | 'sentence' | 'title'
  ) => {
    if (isEditingLocked || !selectedBlock) {
      return;
    }
    setIsCaseMenuOpen(false);
    const src = selectedBlock.text;
    let nextText = src;
    if (mode === 'upper') {
      nextText = src.toLocaleUpperCase('vi-VN');
    } else if (mode === 'lower') {
      nextText = src.toLocaleLowerCase('vi-VN');
    } else if (mode === 'sentence') {
      const lower = src.toLocaleLowerCase('vi-VN');
      nextText = lower.replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (_, p1: string, ch: string) => `${p1}${ch.toLocaleUpperCase('vi-VN')}`);
    } else {
      nextText = src
        .toLocaleLowerCase('vi-VN')
        .replace(/(^|\s)(\p{L})/gu, (_, sp: string, ch: string) => `${sp}${ch.toLocaleUpperCase('vi-VN')}`);
    }
    handleUpdateBlock(selectedBlock.id, { text: nextText, richParagraphs: undefined });
  };

  const handleToggleBlockLinesList = (mode: 'bullet' | 'number') => {
    if (isEditingLocked || !selectedBlock) {
      return;
    }
    const lines = selectedBlock.text.split('\n');
    const allMatch =
      mode === 'bullet'
        ? lines.every((l) => !l.trim() || /^\s*[•\-*]\s+/.test(l))
        : lines.every((l) => !l.trim() || /^\s*\d+\.\s+/.test(l));
    const nextLines = lines.map((line, i) => {
      const cleaned = line.replace(/^(\s*)(?:[•\-*]|\d+\.)\s+/, '$1');
      if (!cleaned.trim() || allMatch) {
        return cleaned;
      }
      const indentMatch = cleaned.match(/^(\s*)(.*)$/);
      const indent = indentMatch?.[1] ?? '';
      const body = indentMatch?.[2] ?? cleaned;
      return mode === 'bullet'
        ? `${indent}• ${body}`
        : `${indent}${i + 1}. ${body}`;
    });
    handleUpdateBlock(selectedBlock.id, {
      text: nextLines.join('\n'),
      richParagraphs: undefined,
    });
  };

  const handleAdjustBlockIndent = (delta: 1 | -1) => {
    if (isEditingLocked || !selectedBlock) {
      return;
    }
    const lines = selectedBlock.text.split('\n');
    const nextLines = lines.map((line) =>
      delta > 0 ? `  ${line}` : line.replace(/^ {1,2}|\t/, '')
    );
    handleUpdateBlock(selectedBlock.id, {
      text: nextLines.join('\n'),
      richParagraphs: undefined,
    });
  };

  const handleToggleSlideNumber = () => {
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(safeSlideIdx, (s) => ({
      ...s,
      showSlideNumber: !s.showSlideNumber,
    }));
  };

  const handleUpdateSlideFooter = (footerText: string) => {
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(safeSlideIdx, (s) => ({
      ...s,
      footerText,
    }));
  };

  const handleApplyFooterToAll = () => {
    if (isEditingLocked) {
      return;
    }
    const footerText = currentSlide.footerText ?? '';
    const showSlideNumber = Boolean(currentSlide.showSlideNumber);
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((slide, idx) =>
        syncSlideSummary(
          {
            ...slide,
            footerText,
            showSlideNumber,
          },
          idx
        )
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
    notifyTypingActivity(true);
  };

  const handleChangeSlideDrawings = (nextStrokes: SlideDrawingStroke[]) => {
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(safeSlideIdx, (s) => ({
      ...s,
      drawings: nextStrokes,
    }));
  };

  const handleUndoSlideStroke = () => {
    if (isEditingLocked) {
      return;
    }
    const cur = currentSlide.drawings ?? [];
    if (cur.length === 0) {
      return;
    }
    handleChangeSlideDrawings(cur.slice(0, -1));
  };

  const handleClearSlideStrokes = () => {
    if (isEditingLocked || (currentSlide.drawings?.length ?? 0) === 0) {
      return;
    }
    handleChangeSlideDrawings([]);
  };

  const handleChangeSlideAspectRatio = (ratio: number) => {
    if (isEditingLocked) {
      return;
    }
    setDeckAspectRatio(ratio);
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((s, idx) =>
        syncSlideSummary({ ...s, slideWidthRatio: ratio }, idx)
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
    notifyTypingActivity(true);
  };

  const handlePreviewTransition = () => {
    setPreviewTransitionKey((k) => k + 1);
  };

  const handlePreviewAnimations = () => {
    setIsPreviewingAnims(false);
    window.requestAnimationFrame(() => {
      setPreviewAnimKey((k) => k + 1);
      setIsPreviewingAnims(true);
    });
  };

  const handleAddSlideComment = (text: string) => {
    if (isEditingLocked) {
      return;
    }
    const newComment: SlideCommentItem = {
      id: `cmt-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      text,
      author: 'Bạn',
      createdAt: new Date().toISOString(),
    };
    mutateSlideAt(safeSlideIdx, (s) => ({
      ...s,
      comments: [...(s.comments ?? []), newComment],
    }));
  };

  const handleDeleteSlideComment = (commentId: string) => {
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(safeSlideIdx, (s) => ({
      ...s,
      comments: (s.comments ?? []).filter((c) => c.id !== commentId),
    }));
  };

  const handleTogglePresentationProtection = () => {
    if (isReadOnly) {
      return;
    }
    const nextProtected = !isPresentationProtected;
    setIsPresentationProtected(nextProtected);
    if (nextProtected) {
      setSelectedBlockId(null);
      setSelectedBlockIds([]);
      setEditingBlockId(null);
      setSelectedImageId(null);
      setActiveDrawTool(null);
      setCopiedBlockStyle(null);
    }
    const live = buildLiveDocument();
    onSaveDocument({ ...live, slideProtected: nextProtected });
    setIsDirty(false);
    setSaveFeedback(true);
    if (saveTimerRef.current !== null) {
      window.clearTimeout(saveTimerRef.current);
    }
    saveTimerRef.current = window.setTimeout(() => {
      setSaveFeedback(false);
      saveTimerRef.current = null;
    }, 2200);
  };

  const handleSaveCopyPptx = async () => {
    if (isExporting) {
      return;
    }
    setIsExporting(true);
    setExportError(null);
    try {
      const syncedSlides = slides.map((s, idx) => syncSlideSummary(s, idx));
      await exportSlidesToPptx(
        `${document.title} (Bản sao)`,
        syncedSlides,
        currentAspectRatio
      );
    } catch {
      setExportError('Không thể lưu bản sao .pptx. Vui lòng thử lại.');
    } finally {
      setIsExporting(false);
    }
  };

  const handleDownloadOriginalFileAction = async () => {
    setExportError(null);
    try {
      if (workspaceSync.onDownloadOriginalFile) {
        await workspaceSync.onDownloadOriginalFile();
        return;
      }
      if (document.rawArrayBuffer && document.rawArrayBuffer.byteLength > 0) {
        await exportOriginalDocumentBinary(
          buildLiveDocument(),
          document.rawArrayBuffer
        );
        return;
      }
      await handleExportPptx();
    } catch {
      setExportError('Không thể tải bản gốc .pptx.');
    }
  };

  const handleExportPdfAction = async () => {
    setExportError(null);
    try {
      await downloadDocumentAsPdf(buildLiveDocument());
    } catch {
      setExportError('Không thể xuất bài thuyết trình sang PDF.');
    }
  };

  const handlePrintDeck = () => {
    window.print();
  };

  const handleAppendAiSlides = (newSlides: SlidePreviewItem[]) => {
    if (isEditingLocked || newSlides.length === 0) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    const startLength = slides.length;
    setSlides((prev) =>
      [...prev, ...newSlides].map((s, idx) => syncSlideSummary(s, idx))
    );
    setActiveSlideIdx(startLength);
    setIsDirty(true);
    setSaveFeedback(false);
    notifyTypingActivity(true);
  };

  const handleUpdateSlideNoteAt = (slideIndex: number, note: string) => {
    if (isEditingLocked) {
      return;
    }
    mutateSlideAt(slideIndex, (s) => ({ ...s, speakerNote: note }));
  };

  const handleApplyAllEmptyAiNotes = (notesBySlideId: Record<string, string>) => {
    if (isEditingLocked) {
      return;
    }
    localMutationPendingRef.current = true;
    setHasLocalEdits(true);
    setSlides((prev) =>
      prev.map((s, idx) => {
        const suggested = notesBySlideId[s.id];
        if (!suggested || (s.speakerNote ?? '').trim().length > 0) {
          return s;
        }
        return syncSlideSummary({ ...s, speakerNote: suggested }, idx);
      })
    );
    setIsDirty(true);
    setSaveFeedback(false);
    notifyTypingActivity(true);
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
        currentAspectRatio ??
        document.slideAspectRatio ??
        syncedSlides[0]?.slideWidthRatio;
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
      isEditingLocked ||
      activeDrawTool !== null ||
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
    if (isEditingLocked || activeDrawTool !== null || isLaserActive || isSlideshowActive || event.button !== 0) return;
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
    if (isEditingLocked || activeDrawTool !== null || isLaserActive || isSlideshowActive || event.button !== 0) {
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
    <div className="shrink-0 p-2.5 border-b border-slate-200/80 grid grid-cols-3 gap-1.5 bg-[#F8FAFC]">
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleAddSlide}
        disabled={isEditingLocked}
        title={
          isReadOnly
            ? 'Bạn chỉ có quyền xem'
            : isPresentationProtected
              ? 'Bài thuyết trình đang ở chế độ Chỉ đọc'
              : 'Thêm slide mới'
        }
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-[11px] font-semibold text-white shadow-xs transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed"
      >
        <Plus className="h-3.5 w-3.5 shrink-0" />
        <span>Thêm</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDuplicateSlide}
        disabled={isEditingLocked}
        title={
          isReadOnly
            ? 'Bạn chỉ có quyền xem'
            : isPresentationProtected
              ? 'Bài thuyết trình đang ở chế độ Chỉ đọc'
              : 'Nhân bản slide đang chọn'
        }
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white hover:bg-[#FFF4ED] border border-slate-200 hover:border-[#C43E1C]/40 text-[11px] font-medium text-slate-700 hover:text-[#C43E1C] transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed"
      >
        <Copy className="h-3 w-3 shrink-0" />
        <span>Nhân bản</span>
      </button>
      <button
        type="button"
        onMouseDown={handlePreventMouseDown}
        onClick={handleDeleteSlide}
        disabled={isEditingLocked || slides.length <= 1}
        title={
          isReadOnly
            ? 'Bạn chỉ có quyền xem'
            : isPresentationProtected
              ? 'Bài thuyết trình đang ở chế độ Chỉ đọc'
              : slides.length <= 1
                ? 'Không thể xóa slide cuối cùng'
                : 'Xóa slide đang chọn'
        }
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white hover:bg-rose-50 border border-slate-200 hover:border-rose-300 text-[11px] font-medium text-slate-700 hover:text-rose-600 transition-colors whitespace-nowrap disabled:opacity-35 disabled:cursor-not-allowed"
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
        const thumbAspectRatio =
          deckAspectRatio ?? getSafeAspectRatio(slide, document);
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
            draggable={!isEditingLocked}
            onDragStart={(e) => {
              if (isEditingLocked) {
                return;
              }
              setDraggedSlideIdx(idx);
              e.dataTransfer.effectAllowed = 'move';
            }}
            onDragOver={(e) => {
              if (isEditingLocked || draggedSlideIdx === null) {
                return;
              }
              e.preventDefault();
              if (dragOverSlideIdx !== idx) {
                setDragOverSlideIdx(idx);
              }
            }}
            onDrop={(e) => {
              e.preventDefault();
              if (!isEditingLocked && draggedSlideIdx !== null && draggedSlideIdx !== idx) {
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
                ? 'bg-[#FFF4ED]'
                : dragOverSlideIdx === idx
                  ? 'bg-amber-50 ring-2 ring-amber-400'
                  : 'hover:bg-slate-200/60'
            } ${slide.hidden ? 'opacity-55' : ''} ${
              draggedSlideIdx === idx ? 'opacity-40 scale-[0.98]' : ''
            }`}
          >
            <div className="w-5 shrink-0 pt-0.5 flex flex-col items-end gap-0.5">
              <span
                className={`text-right text-[11px] font-semibold tabular-nums ${
                  isActive
                    ? 'text-[#C43E1C] font-bold'
                    : 'text-slate-500 group-hover:text-slate-800'
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
                  className={`w-full rounded-lg bg-white shadow-xs overflow-hidden transition-all pointer-events-none ${
                    isActive
                      ? 'ring-2 ring-[#C43E1C] shadow-md'
                      : 'border border-slate-200/80 group-hover:border-slate-300'
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
                  className={`aspect-video w-full rounded-lg bg-white shadow-xs overflow-hidden p-2 flex flex-col justify-between transition-all ${
                    isActive
                      ? 'ring-2 ring-[#C43E1C] shadow-md'
                      : 'border border-slate-200/80 group-hover:border-slate-300'
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
                  isActive ? 'text-[#C43E1C] font-semibold' : 'text-slate-600'
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
    <div className="h-[100dvh] w-full flex flex-col bg-[#F3F4F6] text-slate-900 overflow-hidden">
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

      {/* Shared Top Workspace Header (Microsoft PowerPoint 365 Fluent Light) */}
      <WorkspaceTopBar
        document={document}
        collaborators={collaborators}
        selfSessionId={selfSessionId}
        collabPermission={collabPermission}
        theme="light"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument())}
        onRenameTitle={isEditingLocked ? undefined : handleRenameTitle}
        leftExtra={
          <button
            type="button"
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở danh sách slide"
            title="Danh sách slide"
            className="md:hidden inline-flex items-center justify-center gap-1 h-8 px-2 rounded-lg bg-slate-100 hover:bg-[#FFF4ED] border border-slate-200/80 text-xs font-medium text-slate-700 hover:text-[#C43E1C] transition-colors shrink-0"
          >
            <Layers className="h-3.5 w-3.5 text-[#C43E1C]" />
            <span className="tabular-nums text-[11px]">
              {safeSlideIdx + 1}/{slides.length}
            </span>
          </button>
        }
        toolbar={
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => startSlideshow(0)}
            title="Trình chiếu toàn màn hình (F5)"
            className="inline-flex items-center gap-1.5 h-8 px-2.5 sm:px-3 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs transition-colors whitespace-nowrap shrink-0"
          >
            <Play className="h-3.5 w-3.5 fill-white shrink-0" />
            <span className="hidden sm:inline">Trình chiếu</span>
            <span className="hidden lg:inline text-[10px] opacity-85 font-normal">(F5)</span>
          </button>
        }
        modeSwitch={{
          visible: hasOriginalBinary,
          mode: viewMode === 'native' ? 'original' : 'edit',
          onChange: (m) =>
            handleSwitchViewMode(m === 'original' ? 'native' : 'edit'),
          editDisabled: isEditingLocked,
          originalTitle: 'Xem nguyên bản (Giống Microsoft PowerPoint)',
          editTitle: 'Chỉnh sửa nhanh',
        }}
        onSave={isEditingLocked ? undefined : handleSave}
        isDirty={isDirty}
        isSaved={saveFeedback}
        onDownloadEdited={handleExportPptx}
        onExportPdf={handleExportPdfAction}
        extraMenuItems={[
          {
            id: 'ppt-slideshow',
            label: 'Trình chiếu (F5)',
            icon: <Play className="h-3.5 w-3.5 text-[#C43E1C]" />,
            onClick: () => startSlideshow(0),
          },
          {
            id: 'ppt-presenter',
            label: 'Presenter View',
            icon: <Presentation className="h-3.5 w-3.5 text-[#C43E1C]" />,
            onClick: () => onOpenModal('present', buildLiveDocument()),
          },
          {
            id: 'ppt-qr',
            label: 'Điều khiển QR',
            icon: <QrCode className="h-3.5 w-3.5 text-slate-600" />,
            onClick: () => onOpenModal('qr', buildLiveDocument()),
          },
          {
            id: 'ppt-comments',
            label: `Nhận xét slide (${currentSlide.comments?.length ?? 0})`,
            icon: <MessageSquare className="h-3.5 w-3.5 text-[#C43E1C]" />,
            onClick: () => setIsCommentsOpen((v) => !v),
          },
          {
            id: 'ppt-info',
            label: 'Thông tin tệp',
            icon: <Info className="h-3.5 w-3.5 text-slate-600" />,
            onClick: () => setIsFileInfoOpen(true),
          },
        ]}
      />

      {/* Cloud Hydration Status Banner */}
      {isHydratingCloudFile && (
        <div className="shrink-0 px-4 py-2 bg-[#FFF4ED] border-b border-[#C43E1C]/20 text-xs text-[#9E3014] flex items-center justify-center gap-2">
          <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0 text-[#C43E1C]" />
          <span>Đang tải bản trình chiếu gốc từ đám mây...</span>
        </div>
      )}

      {/* Missing Cloud Buffer Retry Notice (only when no preview slides exist) */}
      {!isHydratingCloudFile &&
        document.slidesPreview.length === 0 &&
        Boolean(document.storagePath || document.r2ObjectKey) &&
        !document.rawArrayBuffer &&
        onRetryCloudDownload && (
          <div className="shrink-0 px-4 py-2 bg-amber-50 border-b border-amber-200 text-xs text-amber-800 flex items-center justify-between gap-2">
            <span className="inline-flex items-center gap-1.5">
              <CloudDownload className="h-3.5 w-3.5 shrink-0" />
              <span>Chưa tải được dữ liệu tệp trình chiếu gốc từ đám mây.</span>
            </span>
            <button
              type="button"
              onClick={onRetryCloudDownload}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-600 hover:bg-amber-700 text-xs font-semibold text-white transition-colors"
            >
              <RefreshCw className="h-3 w-3" />
              <span>Thử tải lại</span>
            </button>
          </div>
        )}

      {/* Export Error Notice */}
      {exportError && (
        <div className="shrink-0 px-4 py-2 bg-rose-50 border-b border-rose-200 text-xs text-rose-700 flex items-center justify-between gap-2">
          <span>{exportError}</span>
          <button
            type="button"
            onClick={() => setExportError(null)}
            aria-label="Đóng thông báo lỗi"
            className="p-1 rounded hover:bg-rose-100"
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
          className={`md:hidden fixed inset-0 z-30 bg-slate-900/35 backdrop-blur-[2px] transition-opacity duration-200 ${
            isMobileDrawerOpen
              ? 'opacity-100 pointer-events-auto'
              : 'opacity-0 pointer-events-none'
          }`}
        />

        {/* Mobile Slide-over Drawer (< md, z-40) */}
        <aside
          aria-label="Danh sách slide di động"
          className={`md:hidden fixed inset-y-0 left-0 z-40 w-[268px] max-w-[84vw] bg-[#F8FAFC] border-r border-slate-200/80 shadow-2xl flex flex-col transition-transform duration-200 ease-out ${
            isMobileDrawerOpen
              ? 'translate-x-0 pointer-events-auto'
              : '-translate-x-full pointer-events-none'
          }`}
        >
          <div className="h-14 shrink-0 px-4 border-b border-slate-200/80 bg-white flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 min-w-0">
              <Layers className="h-4 w-4 text-[#C43E1C] shrink-0" />
              <span className="text-xs font-semibold text-slate-800 truncate">
                Danh sách slide
              </span>
              <span className="px-1.5 py-0.5 rounded-md bg-[#FFF4ED] text-[10px] font-semibold text-[#C43E1C] tabular-nums">
                {slides.length}
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsMobileDrawerOpen(false)}
              aria-label="Đóng danh sách slide"
              className="inline-flex items-center justify-center min-h-[38px] min-w-[38px] rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-900 transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
          {renderSlideActionButtons()}
          <div className="flex-1 min-h-0 overflow-y-auto scrollbar-none pb-safe">
            {renderThumbnailList('mobile')}
          </div>
        </aside>

        {/* Desktop Left Thumbnail Rail (md+, 224px, Fluent #F8FAFC) */}
        <aside
          aria-label="Danh sách slide thu nhỏ"
          className="hidden md:flex w-[224px] shrink-0 bg-[#F8FAFC] border-r border-slate-200/80 flex-col overflow-hidden"
        >
          <div className="shrink-0 px-3.5 py-2.5 border-b border-slate-200/80 flex items-center justify-between gap-2">
            <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
              Trang trình chiếu
            </span>
            <span className="px-1.5 py-0.5 rounded-md bg-[#FFF4ED] text-[10.5px] font-semibold text-[#C43E1C] tabular-nums">
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
              theme="light"
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

        {/* Center Stage Editor + Bottom Speaker Notes (Microsoft PowerPoint 365 Fluent #E5E7EB) */}
        <main className="flex-1 min-w-0 flex flex-col bg-[#E5E7EB] overflow-hidden relative">
          {/* Microsoft PowerPoint 365 Fluent Ribbon (#C43E1C, collapsible) */}
          <div className="shrink-0 bg-[#F8FAFC] border-b border-slate-200/80">
            {/* Fluent Ribbon Tab Bar (ONLYOFFICE 11 Tabs + Prominent Slideshow Play) */}
            <div className="px-2 sm:px-4 flex items-center justify-between gap-2 overflow-x-auto no-scrollbar border-b border-slate-200/70">
              <div
                role="tablist"
                aria-label="Thanh Ribbon PowerPoint ONLYOFFICE"
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
                        if (isRibbonCollapsed) {
                          setIsRibbonCollapsed(false);
                        }
                        setLayoutPickerMode(null);
                        setIsShapePickerOpen(false);
                        setIsCaseMenuOpen(false);
                        setIsHighlightMenuOpen(false);
                      }}
                      className={`px-2 sm:px-3 py-2 text-[11.5px] sm:text-[12px] whitespace-nowrap transition-colors border-b-2 ${
                        isTabActive
                          ? 'text-[#C43E1C] font-semibold border-[#C43E1C] bg-white/70'
                          : 'text-slate-600 font-medium border-transparent hover:text-slate-900 hover:bg-slate-200/50'
                      }`}
                    >
                      {tab.label}
                    </button>
                  );
                })}
              </div>

              <div className="flex items-center gap-1.5 shrink-0 py-1">
                {!isEditingLocked && selectedBlockIds.length >= 2 && (
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleGroupSelected}
                    className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-[#FFF4ED] hover:bg-[#C43E1C] text-[#C43E1C] hover:text-white border border-[#C43E1C]/30 text-[11px] font-semibold whitespace-nowrap transition-colors"
                  >
                    <Group className="h-3 w-3" />
                    <span>Nhóm ({selectedBlockIds.length})</span>
                  </button>
                )}
                {!isEditingLocked && selectedBlock?.groupId && (
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleUngroupSelected}
                    className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 text-[11px] font-semibold whitespace-nowrap transition-colors"
                  >
                    <Ungroup className="h-3 w-3" />
                    <span>Bỏ nhóm</span>
                  </button>
                )}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => startSlideshow(0)}
                  title="Trình chiếu từ đầu (F5)"
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-white text-[11px] font-semibold whitespace-nowrap shadow-xs transition-colors"
                >
                  <Play className="h-3 w-3 fill-white" />
                  <span>Trình chiếu</span>
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setPptViewMode((prev) =>
                      prev === 'sorter' ? 'normal' : 'sorter'
                    )
                  }
                  aria-pressed={pptViewMode === 'sorter'}
                  className={`inline-flex items-center gap-1 px-2 py-1 rounded-lg text-[11px] font-semibold whitespace-nowrap border transition-colors ${
                    pptViewMode === 'sorter'
                      ? 'bg-[#C43E1C] border-[#C43E1C] text-white shadow-xs'
                      : 'bg-white hover:bg-[#FFF4ED] border-slate-200 text-slate-700 hover:text-[#C43E1C]'
                  }`}
                >
                  <LayoutGrid className="h-3 w-3" />
                  <span className="hidden lg:inline">Sắp xếp</span>
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setIsRibbonCollapsed((prev) => !prev)}
                  aria-expanded={!isRibbonCollapsed}
                  aria-label={isRibbonCollapsed ? 'Mở rộng Ribbon' : 'Thu gọn Ribbon'}
                  title={isRibbonCollapsed ? 'Mở rộng thanh công cụ Ribbon' : 'Thu gọn thanh công cụ Ribbon'}
                  className="inline-flex items-center justify-center h-7 w-7 rounded-lg bg-white hover:bg-slate-100 border border-slate-200/80 text-slate-600 hover:text-[#C43E1C] transition-colors shrink-0"
                >
                  {isRibbonCollapsed ? (
                    <ChevronDown className="h-3.5 w-3.5" />
                  ) : (
                    <ChevronUp className="h-3.5 w-3.5" />
                  )}
                </button>
              </div>
            </div>

            {/* Active ONLYOFFICE 11-Tab Fluent Ribbon Strip */}
            {!isRibbonCollapsed && (
            <div className="rounded-xl bg-white shadow-[0_1px_3px_rgba(15,23,42,0.06)] border border-slate-200/80 mx-2 my-1.5 px-2.5 py-1.5 relative">
              <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-h-[38px]">
                {activeRibbonTab === 'file' && (
                  <>
                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleSave}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs whitespace-nowrap disabled:opacity-40"
                    >
                      <Save className="h-3.5 w-3.5" />
                      <span>Lưu</span>
                    </button>
                    <button
                      type="button"
                      disabled={isExporting}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => void handleSaveCopyPptx()}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-xs font-medium text-slate-700 hover:text-[#C43E1C] whitespace-nowrap disabled:opacity-40"
                    >
                      <CopyPlus className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Lưu bản sao</span>
                    </button>
                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => void handleDownloadOriginalFileAction()}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <FileDown className="h-3.5 w-3.5 text-sky-600" />
                      <span>Tải bản gốc (.pptx)</span>
                    </button>
                    <button
                      type="button"
                      disabled={isExporting}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => void handleExportPptx()}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <FileOutput className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Tải bản sửa (.pptx)</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => void handleExportPdfAction()}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <FileText className="h-3.5 w-3.5 text-rose-600" />
                      <span>Xuất PDF</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handlePrintDeck}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <Printer className="h-3.5 w-3.5 text-slate-600" />
                      <span>In</span>
                    </button>
                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsFileInfoOpen(true)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-xs font-medium text-slate-700 hover:text-[#C43E1C] whitespace-nowrap"
                    >
                      <Info className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Thông tin tệp</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'home' && (
                  <div className="flex items-stretch gap-2 shrink-0 py-0.5">
                    {/* Cluster 1: Clipboard & Format Painter */}
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        disabled={isEditingLocked || !clipboardItem}
                        onMouseDown={handlePreventMouseDown}
                        onClick={handlePasteObject}
                        title="Dán đối tượng đã sao chép"
                        className="shrink-0 inline-flex flex-col items-center justify-center h-12 px-2 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-[10px] font-medium text-slate-700 hover:text-[#C43E1C] disabled:opacity-40"
                      >
                        <Clipboard className="h-4 w-4 text-[#C43E1C]" />
                        <span>Dán</span>
                      </button>
                      <div className="grid grid-cols-1 gap-0.5">
                        <div className="flex items-center gap-0.5">
                          <button
                            type="button"
                            disabled={!selectedBlock && !selectedImage}
                            onMouseDown={handlePreventMouseDown}
                            onClick={handleCopyObject}
                            title="Sao chép đối tượng"
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <Copy className="h-3 w-3" />
                            <span className="hidden xl:inline">Chép</span>
                          </button>
                          <button
                            type="button"
                            disabled={isEditingLocked || (!selectedBlock && !selectedImage)}
                            onMouseDown={handlePreventMouseDown}
                            onClick={handleCutObject}
                            title="Cắt đối tượng"
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <Scissors className="h-3 w-3" />
                            <span className="hidden xl:inline">Cắt</span>
                          </button>
                        </div>
                        <button
                          type="button"
                          disabled={isEditingLocked || (!selectedBlock && !copiedBlockStyle)}
                          onMouseDown={handlePreventMouseDown}
                          onClick={handleToggleFormatPainter}
                          aria-pressed={Boolean(copiedBlockStyle)}
                          title="Chổi định dạng (Format Painter): bấm vào khối khác để chép kiểu chữ"
                          className={`h-6 px-1.5 rounded border text-[10.5px] inline-flex items-center gap-1 disabled:opacity-40 ${
                            copiedBlockStyle
                              ? 'bg-[#C43E1C] border-[#C43E1C] text-white font-semibold'
                              : 'bg-slate-50 hover:bg-[#FFF4ED] border-slate-200 text-slate-700'
                          }`}
                        >
                          <Paintbrush className="h-3 w-3" />
                          <span>Sao chép ĐD</span>
                        </button>
                      </div>
                    </div>

                    <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                    {/* Cluster 2: Slide Add / Layout / Reset / Duplicate / Delete */}
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        disabled={isEditingLocked}
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => {
                          setLayoutPickerMode((m) => (m === 'add' ? null : 'add'));
                          setIsShapePickerOpen(false);
                        }}
                        aria-expanded={layoutPickerMode === 'add'}
                        title="Thêm slide mới với bố cục tùy chọn"
                        className={`shrink-0 inline-flex flex-col items-center justify-center h-12 px-2.5 rounded-lg border text-[10px] font-semibold transition-colors disabled:opacity-40 ${
                          layoutPickerMode === 'add'
                            ? 'bg-[#9E3014] border-[#9E3014] text-white'
                            : 'bg-[#C43E1C] hover:bg-[#9E3014] border-[#C43E1C] text-white'
                        }`}
                      >
                        <Plus className="h-4 w-4" />
                        <span>Thêm slide</span>
                      </button>
                      <div className="grid grid-cols-1 gap-0.5">
                        <div className="flex items-center gap-0.5">
                          <button
                            type="button"
                            disabled={isEditingLocked}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => {
                              setLayoutPickerMode((m) => (m === 'change' ? null : 'change'));
                              setIsShapePickerOpen(false);
                            }}
                            aria-expanded={layoutPickerMode === 'change'}
                            title="Đổi bố cục slide hiện tại"
                            className={`h-6 px-2 rounded border text-[10.5px] font-medium inline-flex items-center gap-1 disabled:opacity-40 ${
                              layoutPickerMode === 'change'
                                ? 'bg-[#FFF4ED] border-[#C43E1C]/50 text-[#C43E1C]'
                                : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                            }`}
                          >
                            <LayoutTemplate className="h-3 w-3 text-[#C43E1C]" />
                            <span>Bố cục</span>
                          </button>
                          <button
                            type="button"
                            disabled={isEditingLocked}
                            onMouseDown={handlePreventMouseDown}
                            onClick={handleResetCurrentSlide}
                            title="Đặt lại bố cục chuẩn cho slide (giữ chữ)"
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <RotateCcw className="h-3 w-3" />
                            <span>Đặt lại</span>
                          </button>
                        </div>
                        <div className="flex items-center gap-0.5">
                          <button
                            type="button"
                            disabled={isEditingLocked}
                            onMouseDown={handlePreventMouseDown}
                            onClick={handleDuplicateSlide}
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <CopyPlus className="h-3 w-3" />
                            <span>Nhân bản</span>
                          </button>
                          <button
                            type="button"
                            disabled={isEditingLocked || slides.length <= 1}
                            onMouseDown={handlePreventMouseDown}
                            onClick={handleDeleteSlide}
                            className="h-6 px-1.5 rounded bg-rose-50 hover:bg-rose-100 border border-rose-200 text-[10.5px] text-rose-600 inline-flex items-center gap-1 disabled:opacity-35"
                          >
                            <Trash2 className="h-3 w-3" />
                            <span>Xóa</span>
                          </button>
                        </div>
                      </div>
                    </div>

                    <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                    {/* Cluster 3: Font (2 rows) */}
                    <div className="flex flex-col justify-between gap-1 shrink-0 relative">
                      <div className="flex items-center gap-1">
                        <select
                          aria-label="Phông chữ"
                          disabled={isEditingLocked || !selectedBlock}
                          value={
                            selectedBlock?.fontFamily ?? PPT_FONT_FAMILIES[0].value
                          }
                          onChange={(e) =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              fontFamily: e.target.value,
                            })
                          }
                          className="h-6 px-1.5 rounded bg-slate-50 border border-slate-200 text-[11px] text-slate-800 outline-none focus:border-[#C43E1C] disabled:opacity-40"
                        >
                          {PPT_FONT_FAMILIES.map((f) => (
                            <option key={f.label} value={f.value}>
                              {f.label}
                            </option>
                          ))}
                        </select>
                        <select
                          aria-label="Cỡ chữ"
                          disabled={isEditingLocked || !selectedBlock}
                          value={selectedBlock?.fontSize ?? 18}
                          onChange={(e) =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              fontSize: Number(e.target.value),
                            })
                          }
                          className="h-6 w-14 px-1 rounded bg-slate-50 border border-slate-200 text-[11px] tabular-nums text-slate-800 outline-none focus:border-[#C43E1C] disabled:opacity-40"
                        >
                          {[12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48].map((sz) => (
                            <option key={sz} value={sz}>
                              {sz}pt
                            </option>
                          ))}
                        </select>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              fontSize: Math.min(72, (selectedBlock.fontSize || 18) + 2),
                            })
                          }
                          title="Tăng cỡ chữ (+2pt)"
                          className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[11px] font-bold text-slate-700 disabled:opacity-40"
                        >
                          A↑
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              fontSize: Math.max(8, (selectedBlock.fontSize || 18) - 2),
                            })
                          }
                          title="Giảm cỡ chữ (-2pt)"
                          className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] font-bold text-slate-700 disabled:opacity-40"
                        >
                          A↓
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => {
                            setIsCaseMenuOpen((v) => !v);
                            setIsHighlightMenuOpen(false);
                          }}
                          title="Chuyển kiểu chữ HOA / thường"
                          className={`h-6 px-1.5 rounded border text-[11px] font-semibold disabled:opacity-40 ${
                            isCaseMenuOpen
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          Aa
                        </button>
                      </div>

                      <div className="flex items-center gap-0.5">
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              bold: !selectedBlock.bold,
                            })
                          }
                          aria-pressed={Boolean(selectedBlock?.bold)}
                          title="In đậm"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.bold
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Bold className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              italic: !selectedBlock.italic,
                            })
                          }
                          aria-pressed={Boolean(selectedBlock?.italic)}
                          title="In nghiêng"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.italic
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Italic className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock && handleToggleBlockUnderline(selectedBlock.id)
                          }
                          aria-pressed={Boolean(
                            selectedBlock &&
                              (selectedBlock.underline ||
                                underlinedBlockIds[selectedBlock.id])
                          )}
                          title="Gạch chân"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock &&
                            (selectedBlock.underline ||
                              underlinedBlockIds[selectedBlock.id])
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Underline className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              strikethrough: !selectedBlock.strikethrough,
                            })
                          }
                          aria-pressed={Boolean(selectedBlock?.strikethrough)}
                          title="Gạch ngang chữ"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.strikethrough
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Strikethrough className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              textScript:
                                selectedBlock.textScript === 'super'
                                  ? 'normal'
                                  : 'super',
                            })
                          }
                          aria-pressed={selectedBlock?.textScript === 'super'}
                          title="Chỉ số trên (x²)"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.textScript === 'super'
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Superscript className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              textScript:
                                selectedBlock.textScript === 'sub'
                                  ? 'normal'
                                  : 'sub',
                            })
                          }
                          aria-pressed={selectedBlock?.textScript === 'sub'}
                          title="Chỉ số dưới (H₂O)"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.textScript === 'sub'
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Subscript className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => {
                            setIsHighlightMenuOpen((v) => !v);
                            setIsCaseMenuOpen(false);
                          }}
                          title="Tô màu nổi bật chữ (Highlight)"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.highlightColor
                              ? 'bg-amber-100 border-amber-400 text-amber-800'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Highlighter className="h-3 w-3" />
                        </button>
                        <div className="inline-flex items-center gap-0.5 px-1 h-6 rounded bg-slate-100 border border-slate-200/80">
                          {SLIDE_COLOR_SWATCHES.slice(0, 4).map((swatch) => (
                            <button
                              key={swatch.value}
                              type="button"
                              disabled={isEditingLocked || !selectedBlock}
                              onMouseDown={handlePreventMouseDown}
                              onClick={() =>
                                selectedBlock &&
                                handleUpdateBlock(selectedBlock.id, {
                                  color: swatch.value,
                                })
                              }
                              title={`Màu chữ ${swatch.label}`}
                              className="h-3.5 w-3.5 rounded-full border border-slate-300 disabled:opacity-40"
                              style={{ backgroundColor: swatch.value }}
                            />
                          ))}
                          <input
                            type="color"
                            disabled={isEditingLocked || !selectedBlock}
                            value={toSafeColorInputValue(
                              selectedBlock?.color,
                              hasLightSlideBg ? '#0F172A' : '#FFFFFF'
                            )}
                            onChange={(e) =>
                              selectedBlock &&
                              handleUpdateBlock(selectedBlock.id, {
                                color: e.target.value,
                              })
                            }
                            aria-label="Màu chữ tùy chỉnh"
                            className="h-4 w-4 rounded cursor-pointer border-0 bg-transparent p-0 disabled:opacity-40"
                          />
                        </div>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock && handleClearBlockFormat(selectedBlock.id)
                          }
                          title="Xóa tất cả định dạng"
                          className="h-6 px-1.5 rounded bg-slate-50 hover:bg-rose-50 border border-slate-200 text-[10px] text-slate-600 hover:text-rose-600 disabled:opacity-40"
                        >
                          Xóa ĐD
                        </button>
                      </div>
                    </div>

                    <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                    {/* Cluster 4: Paragraph & Spacing (2 rows) */}
                    <div className="flex flex-col justify-between gap-1 shrink-0">
                      <div className="flex items-center gap-0.5">
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleToggleBlockLinesList('bullet')}
                          title="Danh sách dấu chấm (Bullets)"
                          className="h-6 w-6 rounded bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-slate-700 hover:text-[#C43E1C] flex items-center justify-center disabled:opacity-40"
                        >
                          <List className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleToggleBlockLinesList('number')}
                          title="Danh sách đánh số (1. 2. 3.)"
                          className="h-6 w-6 rounded bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-slate-700 hover:text-[#C43E1C] flex items-center justify-center disabled:opacity-40"
                        >
                          <ListOrdered className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleAdjustBlockIndent(-1)}
                          title="Giảm thụt lề"
                          className="h-6 w-6 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-700 flex items-center justify-center disabled:opacity-40"
                        >
                          <IndentDecrease className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleAdjustBlockIndent(1)}
                          title="Tăng thụt lề"
                          className="h-6 w-6 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-slate-700 flex items-center justify-center disabled:opacity-40"
                        >
                          <IndentIncrease className="h-3 w-3" />
                        </button>
                        <select
                          aria-label="Giãn dòng"
                          disabled={isEditingLocked || !selectedBlock}
                          value={selectedBlock?.lineHeight ?? 1.15}
                          onChange={(e) =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              lineHeight: Number(e.target.value),
                            })
                          }
                          title="Giãn dòng (Line Height)"
                          className="h-6 px-1 rounded bg-slate-50 border border-slate-200 text-[10.5px] text-slate-700 disabled:opacity-40"
                        >
                          <option value={1}>1.0</option>
                          <option value={1.15}>1.15</option>
                          <option value={1.5}>1.5</option>
                          <option value={2}>2.0</option>
                        </select>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, {
                              textColumns: selectedBlock.textColumns === 2 ? 1 : 2,
                            })
                          }
                          aria-pressed={selectedBlock?.textColumns === 2}
                          title="Chia 1 / 2 cột văn bản"
                          className={`h-6 px-1.5 rounded border text-[10.5px] inline-flex items-center gap-0.5 disabled:opacity-40 ${
                            selectedBlock?.textColumns === 2
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C] font-semibold'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Columns className="h-3 w-3" />
                          <span>{selectedBlock?.textColumns === 2 ? '2C' : '1C'}</span>
                        </button>
                      </div>

                      <div className="flex items-center gap-0.5">
                        {(
                          [
                            { value: 'left', label: 'Căn trái', Icon: AlignLeft },
                            { value: 'center', label: 'Căn giữa', Icon: AlignCenter },
                            { value: 'right', label: 'Căn phải', Icon: AlignRight },
                            { value: 'justify', label: 'Căn đều 2 bên', Icon: AlignJustify },
                          ] as const
                        ).map(({ value, label, Icon }) => (
                          <button
                            key={value}
                            type="button"
                            disabled={isEditingLocked || !selectedBlock}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() =>
                              selectedBlock &&
                              handleUpdateBlock(selectedBlock.id, { align: value })
                            }
                            title={label}
                            className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                              (selectedBlock?.align ?? 'left') === value
                                ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                                : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                            }`}
                          >
                            <Icon className="h-3 w-3" />
                          </button>
                        ))}
                        <span className="w-px h-4 bg-slate-200 mx-0.5" />
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, { valign: 'top' })
                          }
                          title="Căn lề trên"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            (selectedBlock?.valign ?? 'top') === 'top'
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <ArrowUpToLine className="h-3 w-3" />
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, { valign: 'middle' })
                          }
                          title="Căn giữa theo chiều dọc"
                          className={`h-6 px-1 rounded border text-[9.5px] font-semibold disabled:opacity-40 ${
                            selectedBlock?.valign === 'middle'
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          Giữa
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked || !selectedBlock}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() =>
                            selectedBlock &&
                            handleUpdateBlock(selectedBlock.id, { valign: 'bottom' })
                          }
                          title="Căn lề dưới"
                          className={`h-6 w-6 rounded border flex items-center justify-center disabled:opacity-40 ${
                            selectedBlock?.valign === 'bottom'
                              ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                              : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                          }`}
                        >
                          <ArrowDownToLine className="h-3 w-3" />
                        </button>
                      </div>
                    </div>

                    <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                    {/* Cluster 5: Quick Insert + 2x5 Quick Shapes + Arrange + Shape Fill/Outline */}
                    <div className="flex items-center gap-1.5 shrink-0">
                      <div className="grid grid-cols-1 gap-0.5">
                        <button
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={handleAddTextBlock}
                          className="h-6 px-2 rounded bg-[#FFF4ED] hover:bg-[#C43E1C] border border-[#C43E1C]/35 text-[10.5px] font-semibold text-[#C43E1C] hover:text-white inline-flex items-center gap-1 disabled:opacity-40"
                        >
                          <Type className="h-3 w-3" />
                          <span>Hộp chữ</span>
                        </button>
                        <button
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={handleTriggerAddImage}
                          className="h-6 px-2 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                        >
                          <ImagePlus className="h-3 w-3 text-emerald-600" />
                          <span>Chèn ảnh</span>
                        </button>
                      </div>

                      <div
                        className="grid grid-cols-5 gap-0.5 p-1 rounded-lg bg-slate-50 border border-slate-200/80"
                        aria-label="Chèn nhanh hình khối"
                      >
                        {PPT_QUICK_SHAPES_10.map((shapeOpt) => {
                          const ShapeIcon = shapeOpt.Icon;
                          return (
                            <button
                              key={shapeOpt.kind}
                              type="button"
                              disabled={isEditingLocked}
                              onMouseDown={handlePreventMouseDown}
                              onClick={() => handleInsertShape(shapeOpt.kind)}
                              title={`Chèn ${shapeOpt.label}`}
                              className="h-5 w-5 rounded hover:bg-[#FFF4ED] text-slate-700 hover:text-[#C43E1C] flex items-center justify-center disabled:opacity-40"
                            >
                              <ShapeIcon className="h-3.5 w-3.5" />
                            </button>
                          );
                        })}
                      </div>

                      <div className="grid grid-cols-1 gap-0.5">
                        <div className="flex items-center gap-0.5">
                          <button
                            type="button"
                            disabled={isEditingLocked || !selectedBlock}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => selectedBlock && handleBringToFront(selectedBlock.id)}
                            title="Đưa lên trên cùng"
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <BringToFront className="h-3 w-3 text-[#C43E1C]" />
                            <span className="hidden xl:inline">Lên trên</span>
                          </button>
                          <button
                            type="button"
                            disabled={isEditingLocked || !selectedBlock}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => selectedBlock && handleSendToBack(selectedBlock.id)}
                            title="Đưa xuống dưới cùng"
                            className="h-6 px-1.5 rounded bg-slate-50 hover:bg-slate-100 border border-slate-200 text-[10.5px] text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                          >
                            <SendToBack className="h-3 w-3 text-slate-500" />
                            <span className="hidden xl:inline">Xuống dưới</span>
                          </button>
                        </div>
                        <div className="flex items-center gap-1">
                          <label className="inline-flex items-center gap-0.5 h-6 px-1.5 rounded bg-slate-50 border border-slate-200 text-[10px] text-slate-600">
                            <span>Nền</span>
                            <input
                              type="color"
                              disabled={isEditingLocked || !selectedBlock}
                              value={toSafeColorInputValue(selectedBlock?.fillColor, '#C43E1C')}
                              onChange={(e) =>
                                selectedBlock &&
                                handleUpdateBlock(selectedBlock.id, { fillColor: e.target.value })
                              }
                              aria-label="Màu nền khối"
                              className="h-4 w-4 rounded cursor-pointer border-0 bg-transparent p-0 disabled:opacity-40"
                            />
                          </label>
                          <label className="inline-flex items-center gap-0.5 h-6 px-1.5 rounded bg-slate-50 border border-slate-200 text-[10px] text-slate-600">
                            <span>Viền</span>
                            <input
                              type="color"
                              disabled={isEditingLocked || !selectedBlock}
                              value={toSafeColorInputValue(selectedBlock?.borderColor, '#FFFFFF')}
                              onChange={(e) =>
                                selectedBlock &&
                                handleUpdateBlock(selectedBlock.id, {
                                  borderColor: e.target.value,
                                  borderWidthPx: selectedBlock.borderWidthPx ?? 2,
                                })
                              }
                              aria-label="Màu viền khối"
                              className="h-4 w-4 rounded cursor-pointer border-0 bg-transparent p-0 disabled:opacity-40"
                            />
                          </label>
                          {(selectedBlock || selectedImage) && (
                            <button
                              type="button"
                              disabled={isEditingLocked}
                              onMouseDown={handlePreventMouseDown}
                              onClick={() => {
                                if (selectedBlock) handleDeleteBlock(selectedBlock.id);
                                else if (selectedImage) handleDeleteImage(selectedImage.id);
                              }}
                              title="Xóa đối tượng đang chọn"
                              className="h-6 w-6 rounded bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-600 flex items-center justify-center disabled:opacity-40"
                            >
                              <Trash2 className="h-3 w-3" />
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {activeRibbonTab === 'insert' && (
                  <>
                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setLayoutPickerMode((m) => (m === 'add' ? null : 'add'));
                        setIsShapePickerOpen(false);
                      }}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs whitespace-nowrap disabled:opacity-40"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      <span>Slide mới</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleInsertTable(3, 3)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Table2 className="h-3.5 w-3.5 text-sky-600" />
                      <span>Bảng (3×3)</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleTriggerAddImage}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <ImagePlus className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Hình ảnh</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setIsShapePickerOpen((v) => !v);
                        setLayoutPickerMode(null);
                      }}
                      aria-expanded={isShapePickerOpen}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        isShapePickerOpen
                          ? 'bg-[#FFF4ED] border-[#C43E1C]/50 text-[#C43E1C] font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Shapes className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Hình khối ({PPT_SHAPE_OPTIONS.length})</span>
                      <ChevronDown className="h-3 w-3 opacity-70" />
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsChartModalOpen(true)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <BarChart3 className="h-3.5 w-3.5 text-amber-600" />
                      <span>Biểu đồ</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddTextBlock}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-[#FFF4ED] hover:bg-[#C43E1C] border border-[#C43E1C]/35 text-xs font-semibold text-[#C43E1C] hover:text-white transition-colors whitespace-nowrap disabled:opacity-40"
                    >
                      <Type className="h-3.5 w-3.5" />
                      <span>Hộp văn bản</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('link')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Link2 className="h-3.5 w-3.5 text-cyan-600" />
                      <span>Liên kết</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('video')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Video className="h-3.5 w-3.5 text-rose-600" />
                      <span>Video</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLinkMediaModalMode('audio')}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Music className="h-3.5 w-3.5 text-violet-600" />
                      <span>Âm thanh</span>
                    </button>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleToggleSlideNumber}
                      aria-pressed={Boolean(currentSlide.showSlideNumber)}
                      className={`shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        currentSlide.showSlideNumber
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C] font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Hash className="h-3.5 w-3.5" />
                      <span>Số slide</span>
                    </button>

                    <input
                      type="text"
                      disabled={isEditingLocked}
                      value={currentSlide.footerText ?? ''}
                      onChange={(e) => handleUpdateSlideFooter(e.target.value)}
                      placeholder="Nội dung chân trang (Footer)..."
                      aria-label="Nội dung chân trang slide"
                      className="shrink-0 h-8 w-44 px-2.5 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-800 outline-none focus:border-[#C43E1C] disabled:opacity-40"
                    />

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleApplyFooterToAll}
                      className="shrink-0 h-8 px-2.5 rounded-lg bg-[#FFF4ED] hover:bg-[#C43E1C] border border-[#C43E1C]/35 text-xs font-semibold text-[#C43E1C] hover:text-white transition-colors whitespace-nowrap disabled:opacity-40"
                    >
                      Áp dụng tất cả
                    </button>
                  </>
                )}

                {activeRibbonTab === 'draw' && (
                  <>
                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setViewMode('edit');
                        setActiveDrawTool((t) => (t === 'pen' ? null : 'pen'));
                      }}
                      aria-pressed={activeDrawTool === 'pen'}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        activeDrawTool === 'pen'
                          ? 'bg-[#C43E1C] border-[#C43E1C] text-white font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Pen className="h-3.5 w-3.5" />
                      <span>Bút vẽ</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setViewMode('edit');
                        setActiveDrawTool((t) => (t === 'highlighter' ? null : 'highlighter'));
                      }}
                      aria-pressed={activeDrawTool === 'highlighter'}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        activeDrawTool === 'highlighter'
                          ? 'bg-amber-500 border-amber-500 text-white font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Highlighter className="h-3.5 w-3.5 text-amber-600" />
                      <span>Tô sáng</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        setViewMode('edit');
                        setActiveDrawTool((t) => (t === 'eraser' ? null : 'eraser'));
                      }}
                      aria-pressed={activeDrawTool === 'eraser'}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap disabled:opacity-40 ${
                        activeDrawTool === 'eraser'
                          ? 'bg-sky-600 border-sky-600 text-white font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Eraser className="h-3.5 w-3.5" />
                      <span>Tẩy</span>
                    </button>

                    {activeDrawTool !== null && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => setActiveDrawTool(null)}
                        className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-800 hover:bg-slate-900 text-xs font-semibold text-white whitespace-nowrap"
                      >
                        <X className="h-3.5 w-3.5" />
                        <span>Thoát vẽ</span>
                      </button>
                    )}

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-slate-100 border border-slate-200/80 shrink-0">
                      {PPT_DRAW_COLORS.map((hex) => (
                        <button
                          key={hex}
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => setDrawColor(hex)}
                          aria-label={`Màu nét ${hex}`}
                          className={`h-4 w-4 rounded-full border disabled:opacity-40 ${
                            drawColor.toLowerCase() === hex.toLowerCase()
                              ? 'ring-2 ring-[#C43E1C] scale-110 border-white'
                              : 'border-slate-300'
                          }`}
                          style={{ backgroundColor: hex }}
                        />
                      ))}
                      <input
                        type="color"
                        disabled={isEditingLocked}
                        value={drawColor}
                        onChange={(e) => setDrawColor(e.target.value)}
                        aria-label="Màu bút tùy chỉnh"
                        className="h-4 w-5 rounded cursor-pointer border-0 bg-transparent p-0"
                      />
                    </div>

                    <div className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-slate-100 border border-slate-200/80 shrink-0">
                      {[2, 4, 6, 10].map((w) => (
                        <button
                          key={w}
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => setDrawWidth(w)}
                          className={`h-7 px-2 rounded-md text-[11px] font-semibold tabular-nums ${
                            drawWidth === w
                              ? 'bg-[#C43E1C] text-white'
                              : 'text-slate-700 hover:bg-white'
                          }`}
                        >
                          {w}px
                        </button>
                      ))}
                    </div>

                    <button
                      type="button"
                      disabled={isEditingLocked || (currentSlide.drawings?.length ?? 0) === 0}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleUndoSlideStroke}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Undo2 className="h-3.5 w-3.5" />
                      <span>Hoàn tác nét</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked || (currentSlide.drawings?.length ?? 0) === 0}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleClearSlideStrokes}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-rose-50 hover:bg-rose-100 border border-rose-200 text-xs font-medium text-rose-600 whitespace-nowrap disabled:opacity-40"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      <span>Xóa toàn bộ nét</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'design' && (
                  <>
                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Palette className="h-3.5 w-3.5 text-[#C43E1C] mr-0.5 shrink-0" />
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
                            disabled={isEditingLocked}
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => handleApplyColorTheme(theme)}
                            title={theme.name}
                            className={`h-8 px-2 rounded-lg border inline-flex items-center gap-1 text-[11px] font-medium whitespace-nowrap disabled:opacity-40 ${
                              isActiveTheme
                                ? 'border-[#C43E1C] bg-[#FFF4ED] text-[#C43E1C] font-semibold'
                                : 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700'
                            }`}
                          >
                            <span className="flex items-center -space-x-1">
                              {theme.swatches.map((hex, i) => (
                                <span
                                  key={`${theme.id}-sw-${i}`}
                                  className="h-3 w-3 rounded-full border border-slate-300"
                                  style={{ backgroundColor: hex }}
                                />
                              ))}
                            </span>
                            <span className="hidden lg:inline">{theme.name}</span>
                          </button>
                        );
                      })}
                    </div>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <div className="inline-flex items-center gap-1 shrink-0">
                      {PPT_GRADIENT_PRESETS.map((g) => (
                        <button
                          key={g.label}
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleApplySlideGradient(g.value)}
                          title={`Nền chuyển sắc: ${g.label}`}
                          className="h-6 w-6 rounded border border-slate-300 hover:scale-110 transition-transform disabled:opacity-40"
                          style={{ backgroundImage: g.value }}
                        />
                      ))}
                    </div>

                    <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-slate-100 border border-slate-200/80 shrink-0">
                      <PaintBucket className="h-3 w-3 text-slate-500" />
                      {SLIDE_BG_PRESETS.map((bg) => (
                        <button
                          key={bg.value}
                          type="button"
                          disabled={isEditingLocked}
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => handleChangeSlideBackground(bg.value)}
                          title={`Màu đơn: ${bg.label}`}
                          className="h-4 w-4 rounded border border-slate-300 hover:scale-110 transition-transform disabled:opacity-40"
                          style={{ backgroundColor: bg.value }}
                        />
                      ))}
                      <input
                        type="color"
                        disabled={isEditingLocked}
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
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => bgImageInputRef.current?.click()}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <Image className="h-3.5 w-3.5 text-emerald-600" />
                      <span>Ảnh nền</span>
                    </button>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleApplyBackgroundToAll}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-[#FFF4ED] hover:bg-[#C43E1C] border border-[#C43E1C]/35 text-xs font-semibold text-[#C43E1C] hover:text-white transition-colors whitespace-nowrap disabled:opacity-40"
                    >
                      <span>Áp dụng nền tất cả</span>
                    </button>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <div className="inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-slate-100 border border-slate-200/80 shrink-0">
                      <button
                        type="button"
                        disabled={isEditingLocked}
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleChangeSlideAspectRatio(16 / 9)}
                        className={`h-7 px-2 rounded-md text-[11px] font-semibold disabled:opacity-40 ${
                          Math.abs(currentAspectRatio - 4 / 3) >= 0.08
                            ? 'bg-[#C43E1C] text-white'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        16:9
                      </button>
                      <button
                        type="button"
                        disabled={isEditingLocked}
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleChangeSlideAspectRatio(4 / 3)}
                        className={`h-7 px-2 rounded-md text-[11px] font-semibold disabled:opacity-40 ${
                          Math.abs(currentAspectRatio - 4 / 3) < 0.08
                            ? 'bg-[#C43E1C] text-white'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        4:3
                      </button>
                    </div>
                  </>
                )}

                {activeRibbonTab === 'transitions' && (
                  <>
                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Sparkles className="h-3.5 w-3.5 text-[#C43E1C] mr-0.5 shrink-0" />
                      <span className="text-[11px] font-medium text-slate-500">Chuyển slide:</span>
                      {PPT_TRANSITION_OPTIONS.map((opt) => {
                        const active =
                          (currentSlide.transition?.type ?? 'fade') === opt.type;
                        return (
                          <button
                            key={opt.type}
                            type="button"
                            disabled={isEditingLocked}
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
                                : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100'
                            }`}
                          >
                            {opt.label}
                          </button>
                        );
                      })}
                    </div>

                    <select
                      aria-label="Thời lượng chuyển cảnh"
                      disabled={isEditingLocked}
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
                      className="shrink-0 h-7 px-2 rounded-md bg-slate-50 border border-slate-200 text-[11px] text-slate-800"
                    >
                      <option value={200}>0.2s</option>
                      <option value={350}>0.35s</option>
                      <option value={600}>0.6s</option>
                      <option value={1000}>1.0s</option>
                    </select>

                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleApplyTransitionToAll}
                      className="shrink-0 h-7 px-2.5 rounded-md bg-[#FFF4ED] hover:bg-[#C43E1C] border border-[#C43E1C]/35 text-[11px] font-semibold text-[#C43E1C] hover:text-white transition-colors whitespace-nowrap disabled:opacity-40"
                    >
                      Áp dụng tất cả
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handlePreviewTransition}
                      className="shrink-0 inline-flex items-center gap-1 h-7 px-2.5 rounded-md bg-slate-800 hover:bg-slate-900 text-[11px] font-semibold text-white whitespace-nowrap"
                    >
                      <Play className="h-3 w-3 fill-white" />
                      <span>Xem trước</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'animations' && (
                  <>
                    <div className="inline-flex items-center gap-1 shrink-0">
                      <Wand2 className="h-3.5 w-3.5 text-[#C43E1C] shrink-0" />
                      <span className="text-[11px] font-medium text-slate-500 whitespace-nowrap">
                        Hiệu ứng đối tượng:
                      </span>
                      {selectedBlock && !isEditingLocked ? (
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
                                    ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                                    : 'bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100'
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
                                className="h-7 px-1.5 rounded-md bg-slate-50 border border-slate-200 text-[11px] text-slate-800"
                              >
                                {[1, 2, 3, 4, 5, 6, 7, 8].map((n) => (
                                  <option key={n} value={n}>
                                    Click #{n}
                                  </option>
                                ))}
                              </select>
                            )}
                        </>
                      ) : (
                        <span className="text-[11px] text-slate-400 whitespace-nowrap">
                          (Chọn 1 khung/hình khối trên slide để gán hiệu ứng)
                        </span>
                      )}
                    </div>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handlePreviewAnimations}
                      className="shrink-0 inline-flex items-center gap-1 h-7 px-2.5 rounded-md bg-[#C43E1C] hover:bg-[#9E3014] text-[11px] font-semibold text-white whitespace-nowrap"
                    >
                      <Play className="h-3 w-3 fill-white" />
                      <span>Xem trước hiệu ứng</span>
                    </button>

                    {currentBlocks.some((b) => b.animation && b.animation.type !== 'none') && (
                      <div className="inline-flex items-center gap-1 shrink-0">
                        {currentBlocks
                          .filter((b) => b.animation && b.animation.type !== 'none')
                          .map((b) => (
                            <button
                              key={`anim-badge-${b.id}`}
                              type="button"
                              onMouseDown={handlePreventMouseDown}
                              onClick={() => {
                                setSelectedBlockId(b.id);
                                setSelectedBlockIds([b.id]);
                              }}
                              className="h-6 px-2 rounded bg-indigo-50 border border-indigo-200 text-[10.5px] font-semibold text-indigo-700 whitespace-nowrap"
                            >
                              #{b.animation?.order ?? 1} {b.kind === 'shape' ? 'Hình' : 'Chữ'}
                            </button>
                          ))}
                      </div>
                    )}
                  </>
                )}

                {activeRibbonTab === 'collab' && (
                  <>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => onOpenModal('share', buildLiveDocument())}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-3 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs whitespace-nowrap"
                    >
                      <ExternalLink className="h-3.5 w-3.5" />
                      <span>Chia sẻ tài liệu</span>
                    </button>
                    <button
                      type="button"
                      disabled={!workspaceSync.onOpenVersionHistory}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => workspaceSync.onOpenVersionHistory?.()}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap disabled:opacity-40"
                    >
                      <History className="h-3.5 w-3.5 text-sky-600" />
                      <span>Lịch sử phiên bản</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsCommentsOpen((v) => !v)}
                      aria-pressed={isCommentsOpen}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        isCommentsOpen
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C] font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <MessageSquare className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Nhận xét ({currentSlide.comments?.length ?? 0})</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'protect' && (
                  <>
                    <button
                      type="button"
                      disabled={isReadOnly}
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleTogglePresentationProtection}
                      aria-pressed={isPresentationProtected}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-3 rounded-lg border text-xs font-semibold whitespace-nowrap disabled:opacity-40 ${
                        isPresentationProtected
                          ? 'bg-amber-600 border-amber-600 text-white shadow-xs'
                          : 'bg-slate-50 hover:bg-[#FFF4ED] border-slate-200 text-slate-800 hover:text-[#C43E1C]'
                      }`}
                    >
                      {isPresentationProtected ? (
                        <Unlock className="h-3.5 w-3.5" />
                      ) : (
                        <Lock className="h-3.5 w-3.5 text-[#C43E1C]" />
                      )}
                      <span>
                        {isPresentationProtected
                          ? 'Mở khóa chỉnh sửa'
                          : 'Khóa chỉnh sửa (Chỉ đọc)'}
                      </span>
                    </button>
                    <span
                      className={`inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        isPresentationProtected
                          ? 'bg-amber-50 border-amber-200 text-amber-800'
                          : 'bg-emerald-50 border-emerald-200 text-emerald-700'
                      }`}
                    >
                      <ShieldCheck className="h-3.5 w-3.5" />
                      <span>
                        {isPresentationProtected
                          ? 'Đang khóa: ngăn sửa đổi & vẽ'
                          : 'Đang mở khóa chỉnh sửa'}
                      </span>
                    </span>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsFileInfoOpen(true)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <Info className="h-3.5 w-3.5 text-slate-600" />
                      <span>Thông tin bảo vệ</span>
                    </button>
                  </>
                )}

                {activeRibbonTab === 'view' && (
                  <>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setPptViewMode('normal')}
                      aria-pressed={pptViewMode === 'normal'}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        pptViewMode === 'normal'
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C] font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Presentation className="h-3.5 w-3.5" />
                      <span>Bình thường</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setPptViewMode('sorter')}
                      aria-pressed={pptViewMode === 'sorter'}
                      className={`shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        pptViewMode === 'sorter'
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C] font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <LayoutGrid className="h-3.5 w-3.5" />
                      <span>Sắp xếp slide</span>
                    </button>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => startSlideshow(0)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white whitespace-nowrap"
                    >
                      <Play className="h-3.5 w-3.5 fill-white" />
                      <span>Từ đầu (F5)</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => startSlideshow(safeSlideIdx)}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-xs font-medium text-slate-700 hover:text-[#C43E1C] whitespace-nowrap"
                    >
                      <MonitorPlay className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Trang hiện tại</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => onOpenModal('present', buildLiveDocument())}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <Presentation className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Presenter</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => onOpenModal('qr', buildLiveDocument())}
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <QrCode className="h-3.5 w-3.5 text-cyan-600" />
                      <span>QR</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsLaserActive((p) => !p)}
                      aria-pressed={isLaserActive}
                      className={`shrink-0 inline-flex items-center gap-1 h-8 px-2 rounded-lg border text-xs font-medium whitespace-nowrap ${
                        isLaserActive
                          ? 'bg-rose-50 border-rose-300 text-rose-700 font-semibold'
                          : 'bg-slate-50 hover:bg-slate-100 border-slate-200 text-slate-700'
                      }`}
                    >
                      <Crosshair className="h-3.5 w-3.5" />
                      <span>Laser</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setIsNotesExpanded((v) => !v)}
                      aria-pressed={isNotesExpanded}
                      className="shrink-0 inline-flex items-center gap-1 h-8 px-2 rounded-lg bg-slate-50 hover:bg-slate-100 border border-slate-200 text-xs font-medium text-slate-700 whitespace-nowrap"
                    >
                      <StickyNote className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Ghi chú</span>
                    </button>

                    <span className="h-5 w-px mx-0.5 bg-slate-200/80 shrink-0" />

                    <div className="shrink-0 inline-flex items-center gap-0.5 p-0.5 rounded-lg bg-slate-100 border border-slate-200/80">
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => setZoomPctClamped(zoomPct - 10)}
                        aria-label="Thu nhỏ"
                        className="h-7 w-7 rounded-md hover:bg-white flex items-center justify-center text-slate-700"
                      >
                        <ZoomOut className="h-3.5 w-3.5" />
                      </button>
                      <span className="px-1.5 text-[11px] font-semibold tabular-nums text-slate-800">
                        {zoomPct}%
                      </span>
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => setZoomPctClamped(zoomPct + 10)}
                        aria-label="Phóng to"
                        className="h-7 w-7 rounded-md hover:bg-white flex items-center justify-center text-slate-700"
                      >
                        <ZoomIn className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={handleFitToWidth}
                        className="h-7 px-2 rounded-md hover:bg-white text-[11px] font-medium text-slate-700"
                      >
                        Vừa khung
                      </button>
                    </div>

                    <div className="shrink-0 inline-flex items-center rounded-lg bg-slate-100 border border-slate-200/80 p-0.5">
                      <button
                        type="button"
                        onClick={handleToggleTimer}
                        className="inline-flex items-center gap-1 h-7 px-2 rounded-md hover:bg-white text-xs font-medium text-slate-800 whitespace-nowrap"
                      >
                        {isTimerRunning ? (
                          <Pause className="h-3 w-3 text-emerald-600 shrink-0" />
                        ) : (
                          <Play className="h-3 w-3 text-slate-600 shrink-0" />
                        )}
                        <span className="font-mono text-[11px] tabular-nums">
                          {formatTimer(elapsedSeconds)}
                        </span>
                      </button>
                      {elapsedSeconds > 0 && (
                        <button
                          type="button"
                          onClick={handleResetTimer}
                          aria-label="Đặt lại đồng hồ"
                          className="inline-flex items-center justify-center h-7 w-6 rounded-md hover:bg-white text-slate-500 hover:text-slate-900"
                        >
                          <RotateCcw className="h-3 w-3" />
                        </button>
                      )}
                    </div>
                  </>
                )}

                {activeRibbonTab === 'ai' && (
                  <>
                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        setAiModalState({ isOpen: true, initialTab: 'generate' })
                      }
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-3 rounded-lg bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs whitespace-nowrap disabled:opacity-40"
                    >
                      <Sparkles className="h-3.5 w-3.5" />
                      <span>Tạo slide bằng AI</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        setAiModalState({ isOpen: true, initialTab: 'summarize' })
                      }
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-xs font-medium text-slate-700 hover:text-[#C43E1C] whitespace-nowrap"
                    >
                      <Wand2 className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>Tóm tắt nội dung</span>
                    </button>
                    <button
                      type="button"
                      disabled={isEditingLocked}
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        setAiModalState({ isOpen: true, initialTab: 'notes' })
                      }
                      className="shrink-0 inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 text-xs font-medium text-slate-700 hover:text-[#C43E1C] whitespace-nowrap disabled:opacity-40"
                    >
                      <StickyNote className="h-3.5 w-3.5 text-amber-600" />
                      <span>Tạo Speaker Notes</span>
                    </button>
                  </>
                )}
              </div>

              {/* Popover: Case Transform (Aa) */}
              {isCaseMenuOpen && !isEditingLocked && selectedBlock && (
                <div className="mt-2 p-2 rounded-xl bg-white border border-slate-200 shadow-xl flex flex-wrap gap-1.5 z-30 relative">
                  {(
                    [
                      { id: 'sentence', label: 'Viết hoa đầu câu' },
                      { id: 'lower', label: 'chữ thường' },
                      { id: 'upper', label: 'CHỮ IN HOA' },
                      { id: 'title', label: 'Viết Hoa Từng Từ' },
                    ] as const
                  ).map((c) => (
                    <button
                      key={c.id}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleTransformBlockCase(c.id)}
                      className="px-2.5 py-1 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 hover:border-[#C43E1C]/50 text-xs font-medium text-slate-800 hover:text-[#C43E1C]"
                    >
                      {c.label}
                    </button>
                  ))}
                </div>
              )}

              {/* Popover: Text Highlight Swatches */}
              {isHighlightMenuOpen && !isEditingLocked && selectedBlock && (
                <div className="mt-2 p-2 rounded-xl bg-white border border-slate-200 shadow-xl flex flex-wrap items-center gap-1.5 z-30 relative">
                  <span className="text-[11px] font-medium text-slate-500 mr-1">
                    Màu tô sáng:
                  </span>
                  {PPT_HIGHLIGHT_SWATCHES.map((sw) => (
                    <button
                      key={sw.value}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => {
                        handleUpdateBlock(selectedBlock.id, {
                          highlightColor: sw.value,
                        });
                        setIsHighlightMenuOpen(false);
                      }}
                      title={sw.label}
                      className="h-6 px-2 rounded-md border border-slate-300 text-[11px] font-medium text-slate-800"
                      style={{ backgroundColor: sw.value }}
                    >
                      {sw.label}
                    </button>
                  ))}
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => {
                      handleUpdateBlock(selectedBlock.id, {
                        highlightColor: undefined,
                      });
                      setIsHighlightMenuOpen(false);
                    }}
                    className="h-6 px-2 rounded-md bg-slate-100 hover:bg-slate-200 text-[11px] text-slate-700"
                  >
                    Không màu
                  </button>
                </div>
              )}

              {/* Popover: 6 Slide Layouts Picker (Add / Change mode) */}
              {layoutPickerMode !== null && !isEditingLocked && (
                <div className="mt-2 p-2.5 rounded-xl bg-white border border-slate-200 shadow-xl z-30 relative">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-semibold text-slate-700">
                      {layoutPickerMode === 'add'
                        ? 'Chọn bố cục cho slide mới:'
                        : 'Đổi bố cục cho slide hiện tại:'}
                    </span>
                    <button
                      type="button"
                      onClick={() => setLayoutPickerMode(null)}
                      className="text-slate-400 hover:text-slate-700"
                    >
                      <X className="h-3.5 w-3.5" />
                    </button>
                  </div>
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2">
                    {PPT_SLIDE_LAYOUTS.map((layout) => (
                      <button
                        key={layout.id}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleApplySlideLayout(layout.id)}
                        className="text-left p-2 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 hover:border-[#C43E1C]/50 transition-colors"
                      >
                        <div className="text-[11px] font-semibold text-slate-900">
                          {layout.label}
                        </div>
                        <div className="text-[10px] text-slate-500 mt-0.5 line-clamp-1">
                          {layout.desc}
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* Popover: Full 13 SVG Shapes Picker */}
              {isShapePickerOpen && !isEditingLocked && (
                <div className="mt-2 p-2.5 rounded-xl bg-white border border-slate-200 shadow-xl flex flex-wrap gap-1.5 z-30 relative">
                  {PPT_SHAPE_OPTIONS.map((shapeOpt) => {
                    const ShapeIcon = shapeOpt.Icon;
                    return (
                      <button
                        key={shapeOpt.kind}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleInsertShape(shapeOpt.kind)}
                        className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50 hover:bg-[#FFF4ED] border border-slate-200 hover:border-[#C43E1C]/50 text-[11px] font-medium text-slate-800 hover:text-[#C43E1C] transition-colors"
                      >
                        <ShapeIcon className="h-3.5 w-3.5 text-[#C43E1C]" />
                        <span>{shapeOpt.label}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
            )}
          </div>

          {/* Active Laser Mode Banner */}
          {isLaserActive && (
            <div className="shrink-0 px-3 sm:px-6 pt-2 flex justify-center">
              <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-rose-50 border border-rose-200 text-[11.5px] text-rose-700 shadow-xs">
                <span className="h-2 w-2 rounded-full bg-rose-500 shadow-[0_0_8px_2px_rgba(244,63,94,0.65)] animate-pulse shrink-0" />
                <span className="truncate font-medium">
                  Bút Laser đang bật — Di chuyển hoặc chạm trên slide để chiếu điểm
                </span>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setIsLaserActive(false)}
                  className="ml-1 px-2 py-0.5 rounded-full bg-rose-600 hover:bg-rose-700 text-[11px] font-semibold text-white transition-colors shrink-0"
                >
                  Tắt
                </button>
              </div>
            </div>
          )}

          {/* Slide Sorter Grid OR Main Editable Slide Stage Container */}
          {pptViewMode === 'sorter' ? (
            <div className="flex-1 min-h-0 overflow-y-auto scrollbar-none bg-[#F3F4F6] p-4 sm:p-6">
              <div className="flex items-center justify-between gap-3 mb-4">
                <div>
                  <h2 className="text-sm sm:text-base font-bold text-slate-900 flex items-center gap-2">
                    <LayoutGrid className="h-4 w-4 text-[#C43E1C]" />
                    <span>Chế độ sắp xếp trang chiếu (Slide Sorter)</span>
                  </h2>
                  <p className="text-xs text-slate-500 mt-0.5">
                    Kéo thả để đổi thứ tự slide hoặc bấm đúp vào slide để chỉnh sửa chi tiết.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => setPptViewMode('normal')}
                  className="px-3 py-1.5 rounded-xl bg-[#C43E1C] hover:bg-[#9E3014] text-xs font-semibold text-white shadow-xs shrink-0"
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
                      draggable={!isEditingLocked}
                      onDragStart={(e) => {
                        if (isEditingLocked) {
                          return;
                        }
                        setDraggedSlideIdx(idx);
                        e.dataTransfer.effectAllowed = 'move';
                      }}
                      onDragOver={(e) => {
                        if (isEditingLocked || draggedSlideIdx === null) {
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
                          !isEditingLocked &&
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
                          ? 'bg-[#FFF4ED] border-[#C43E1C] ring-2 ring-[#C43E1C] shadow-md'
                          : dragOverSlideIdx === idx
                            ? 'bg-amber-50 border-amber-400 ring-2 ring-amber-400/60'
                            : 'bg-white border-slate-200/80 hover:border-slate-300 shadow-xs'
                      } ${slideItem.hidden ? 'opacity-60' : ''}`}
                    >
                      <div
                        style={{
                          aspectRatio: String(
                            getSafeAspectRatio(slideItem, document)
                          ),
                        }}
                        className="w-full rounded-xl overflow-hidden border border-slate-200/80 pointer-events-none relative"
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
                          <span className="absolute top-2 right-2 px-1.5 py-0.5 rounded bg-slate-900/80 text-[10px] font-semibold text-amber-200 inline-flex items-center gap-1">
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
                                : 'bg-slate-100 text-slate-600'
                            }`}
                          >
                            {idx + 1}
                          </span>
                          <span className="text-xs font-medium text-slate-800 truncate">
                            {slideItem.title}
                          </span>
                        </div>

                        {!isEditingLocked && (
                          <div
                            onClick={(e) => e.stopPropagation()}
                            className="flex items-center gap-1 shrink-0"
                          >
                            <button
                              type="button"
                              onClick={() => handleDuplicateSlideAt(idx)}
                              title="Nhân bản slide"
                              className="p-1 rounded hover:bg-slate-100 text-slate-500 hover:text-slate-900"
                            >
                              <Copy className="h-3.5 w-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleToggleHideSlide(idx)}
                              title={slideItem.hidden ? 'Hiện slide' : 'Ẩn slide'}
                              className="p-1 rounded hover:bg-slate-100 text-slate-500 hover:text-amber-600"
                            >
                              {slideItem.hidden ? (
                                <Eye className="h-3.5 w-3.5 text-amber-600" />
                              ) : (
                                <EyeOff className="h-3.5 w-3.5" />
                              )}
                            </button>
                            <button
                              type="button"
                              disabled={slides.length <= 1}
                              onClick={() => handleDeleteSlideAt(idx)}
                              title="Xóa slide"
                              className="p-1 rounded hover:bg-rose-50 text-slate-500 hover:text-rose-600 disabled:opacity-30"
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
            {isPresentationProtected && (
              <div className="shrink-0 mb-2 px-3 py-1.5 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-800 flex items-center justify-between gap-2 shadow-xs">
                <span className="inline-flex items-center gap-1.5 font-medium truncate">
                  <Lock className="h-3.5 w-3.5 text-amber-600 shrink-0" />
                  <span>Bài thuyết trình đang ở chế độ Chỉ đọc — mọi thao tác sửa nội dung và vẽ đều bị khóa.</span>
                </span>
                {!isReadOnly && (
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleTogglePresentationProtection}
                    className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-600 hover:bg-amber-700 text-[11px] font-semibold text-white transition-colors"
                  >
                    <Unlock className="h-3 w-3" />
                    <span>Mở khóa chỉnh sửa</span>
                  </button>
                )}
              </div>
            )}
            <div
              ref={editorViewportRef}
              className="flex-1 min-h-0 w-full flex relative overflow-auto p-2"
            >
              <div
                style={{
                  width: `${Math.round(zoomedWidthPx)}px`,
                  height: `${Math.round(zoomedHeightPx)}px`,
                }}
                className="m-auto shrink-0 relative"
              >
              <div
                ref={editorCanvasRef}
                key={`editor-stage-${currentSlide.id}-${previewTransitionKey}`}
                onClick={() => {
                  if (justDraggedRef.current || activeDrawTool !== null) {
                    return;
                  }
                  if (!isLaserActive) {
                    setSelectedBlockId(null);
                    setSelectedBlockIds([]);
                    setEditingBlockId(null);
                    setSelectedImageId(null);
                  }
                }}
                onDoubleClick={() => {
                  if (viewMode === 'native' && !isEditingLocked) {
                    handleUnlockDirectEdit();
                  }
                }}
                onPointerMove={updateLaserFromPointer}
                onPointerDown={updateLaserFromPointer}
                onPointerLeave={handlePointerLeaveStage}
                style={{
                  width: `${Math.round(zoomedWidthPx)}px`,
                  height: `${Math.round(zoomedHeightPx)}px`,
                  ...buildSlideBackgroundStyle(currentSlide),
                  ...(previewTransitionKey > 0 ? slideshowAnimationStyle : {}),
                }}
                className={`relative w-full h-full rounded-2xl overflow-hidden border shadow-[0_6px_28px_rgba(15,23,42,0.14)] transition-shadow duration-200 ${
                  isLaserActive
                    ? 'border-rose-500 ring-2 ring-rose-400/40 cursor-none touch-none select-none'
                    : copiedBlockStyle
                      ? 'border-[#C43E1C] ring-2 ring-[#C43E1C]/35 cursor-crosshair'
                      : 'border-slate-300/80 ring-1 ring-slate-900/5'
                }`}
              >
                {viewMode === 'native' || isEditingLocked ? (
                  <ReadOnlySlideStage
                    doc={document}
                    slide={currentSlide}
                    slideIndex={safeSlideIdx}
                    totalSlides={slides.length}
                    stageWidthPx={actualEditorCanvasWidth}
                    underlinedBlockIds={underlinedBlockIds}
                    onActivateEditBlock={
                      !isEditingLocked ? handleUnlockDirectEdit : undefined
                    }
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
                          key={`${block.id}-${previewAnimKey}`}
                          onClick={(event) => {
                            if (isLaserActive || activeDrawTool !== null) {
                              return;
                            }
                            event.stopPropagation();
                            if (copiedBlockStyle && !isEditingLocked) {
                              applyCopiedStyleToBlock(block.id);
                              return;
                            }
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
                            ...(isPreviewingAnims
                              ? getSlideshowAnimStyle(block.animation, 99)
                              : {}),
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
                          {isSelected && !isEditingLocked && (
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
                                if (!isEditingLocked && block.shapeKind !== 'line') {
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
                                            disabled={isLaserActive || isEditingLocked}
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
                                if (!isEditingLocked && !isLaserActive) {
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
                              disabled={isLaserActive || isEditingLocked}
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
                              underline={Boolean(
                                block.underline || underlinedBlockIds[block.id]
                              )}
                              strikethrough={Boolean(block.strikethrough)}
                              highlightColor={block.highlightColor}
                              textScript={block.textScript}
                              lineHeight={block.lineHeight}
                              textColumns={block.textColumns}
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
                              underlined={Boolean(
                                block.underline || underlinedBlockIds[block.id]
                              )}
                              onActivateEdit={() => {
                                if (!isLaserActive) {
                                  setSelectedBlockId(block.id);
                                  setSelectedBlockIds([block.id]);
                                  setSelectedImageId(null);
                                  if (!isEditingLocked) setEditingBlockId(block.id);
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
                          {isSelected && !isEditingLocked && (
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
                                  if (isLaserActive || activeDrawTool !== null) {
                                    return;
                                  }
                                  event.stopPropagation();
                                  if (copiedBlockStyle && !isEditingLocked) {
                                    applyCopiedStyleToBlock(block.id);
                                    return;
                                  }
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
                                  disabled={isLaserActive || isEditingLocked}
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
                                    block.underline || underlinedBlockIds[block.id]
                                  )}
                                  strikethrough={Boolean(block.strikethrough)}
                                  highlightColor={block.highlightColor}
                                  textScript={block.textScript}
                                  lineHeight={block.lineHeight}
                                  textColumns={block.textColumns}
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

                {/* Slide Drawing Ink Layer (rendered above stage in edit mode; ReadOnlySlideStage handles static rendering) */}
                {viewMode === 'edit' && !isEditingLocked && (
                  <SlideDrawingOverlay
                    strokes={currentSlide.drawings ?? []}
                    activeTool={activeDrawTool}
                    strokeColor={drawColor}
                    strokeWidth={drawWidth}
                    disabled={isEditingLocked || isLaserActive}
                    onChangeStrokes={handleChangeSlideDrawings}
                  />
                )}

                {/* Slide Footer & Number Strip (edit mode) */}
                {viewMode === 'edit' &&
                  !isEditingLocked &&
                  (currentSlide.footerText || currentSlide.showSlideNumber) && (
                    <div
                      className={`pointer-events-none absolute inset-x-0 bottom-0 z-20 flex items-center justify-between gap-2 px-3 py-1 ${
                        hasLightSlideBg ? 'text-slate-600' : 'text-slate-300/90'
                      }`}
                      style={{
                        fontSize: `${Math.max(6, 10.5 * editorFontScale).toFixed(1)}px`,
                      }}
                    >
                      <span className="truncate">{currentSlide.footerText ?? ''}</span>
                      {currentSlide.showSlideNumber && (
                        <span className="shrink-0 font-semibold tabular-nums">
                          {safeSlideIdx + 1} / {slides.length}
                        </span>
                      )}
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
            </div>

            {/* Stage Footer Indicator & Quick Controls */}
            <div className="shrink-0 pt-2 flex items-center justify-between gap-2">
              <button
                type="button"
                onClick={() => setIsMobileDrawerOpen(true)}
                className="md:hidden inline-flex items-center gap-1.5 min-h-[34px] px-2.5 rounded-lg bg-white hover:bg-[#FFF4ED] border border-slate-200 text-[11.5px] font-medium text-slate-700 hover:text-[#C43E1C] shadow-xs transition-colors"
              >
                <Layers className="h-3.5 w-3.5 text-[#C43E1C]" />
                <span>Slide</span>
              </button>

              <p className="hidden md:block text-[11px] text-slate-600">
                Phím{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-700 font-mono text-[10px]">
                  ←
                </kbd>{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-700 font-mono text-[10px]">
                  →
                </kbd>{' '}
                chuyển slide ·{' '}
                <kbd className="px-1.5 py-0.5 rounded bg-white border border-slate-200 text-slate-700 font-mono text-[10px]">
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
                  className="h-7 w-7 rounded-lg bg-white hover:bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-700 shadow-xs transition disabled:opacity-35 disabled:cursor-not-allowed"
                >
                  <ChevronLeft className="h-4 w-4" />
                </button>

                <span className="text-xs font-semibold text-slate-700 tabular-nums whitespace-nowrap">
                  Slide {safeSlideIdx + 1} / {slides.length}
                </span>

                <button
                  type="button"
                  onClick={handleNextSlide}
                  disabled={safeSlideIdx >= slides.length - 1}
                  aria-label="Chuyển slide kế tiếp"
                  className="h-7 w-7 rounded-lg bg-white hover:bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-700 shadow-xs transition disabled:opacity-35 disabled:cursor-not-allowed"
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
                        : 'w-2 bg-slate-400/50 hover:bg-slate-500'
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>
          )}

          {/* Resizable Bottom Speaker Notes Dock (Fluent Light) */}
          <section
            aria-label="Ghi chú diễn giả"
            className="shrink-0 border-t border-slate-200/80 bg-white"
          >
            {isNotesExpanded && (
              <div
                onPointerDown={startNotesResize}
                title="Kéo lên/xuống để thay đổi chiều cao ghi chú diễn giả"
                className="h-1.5 w-full cursor-row-resize hover:bg-[#C43E1C]/20 transition-colors flex items-center justify-center"
              >
                <span className="h-0.5 w-10 rounded-full bg-slate-300" />
              </div>
            )}
            <div className="px-3.5 sm:px-5 py-1.5 flex items-center justify-between gap-2">
              <div className="flex items-center gap-2 min-w-0">
                <StickyNote className="h-3.5 w-3.5 text-[#C43E1C] shrink-0" />
                <span className="text-xs font-semibold text-slate-800 whitespace-nowrap">
                  Ghi chú diễn giả
                </span>
                <span className="text-[11px] text-slate-500 tabular-nums whitespace-nowrap">
                  (Slide {safeSlideIdx + 1})
                </span>
                {isEditingLocked && (
                  <span className="text-[10.5px] text-amber-700 font-medium whitespace-nowrap">
                    · Chỉ xem
                  </span>
                )}

                {!isNotesExpanded && currentSlide.speakerNote && (
                  <p className="hidden sm:block text-xs text-slate-500 truncate max-w-md ml-2">
                    — {currentSlide.speakerNote}
                  </p>
                )}
              </div>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => setIsNotesExpanded((prev) => !prev)}
                aria-expanded={isNotesExpanded}
                className="shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-[#FFF4ED] text-[11px] font-medium text-slate-700 hover:text-[#C43E1C] transition-colors"
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
                  disabled={isEditingLocked}
                  onChange={(event) =>
                    handleSpeakerNoteChange(event.target.value)
                  }
                  style={{ height: `${notesHeightPx}px` }}
                  aria-label={`Nội dung ghi chú diễn giả cho Slide ${safeSlideIdx + 1}`}
                  placeholder={
                    isEditingLocked
                      ? 'Slide này ở chế độ Chỉ xem.'
                      : 'Nhập ghi chú diễn giả cho slide này (được lưu và xuất kèm tệp .pptx)...'
                  }
                  className="w-full resize-none rounded-xl bg-[#F8FAFC] border border-slate-200 focus:border-[#C43E1C] focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#C43E1C]/15 px-3.5 py-2 text-xs sm:text-[13px] text-slate-800 placeholder:text-slate-400 leading-relaxed transition-colors scrollbar-none disabled:opacity-70"
                />
              </div>
            )}
          </section>

          {/* Bottom PowerPoint Status Bar (10%-300% Zoom, Fit, View Switcher, Start Slideshow) */}
          <div className={!isEditingLocked && viewMode === 'edit' ? 'pb-12 md:pb-0' : ''}>
            <OfficeStatusBar
              variant="powerpoint"
              theme="light"
              zoomPercent={zoomPct}
              onZoomChange={setZoomPctClamped}
              onFitToWidth={handleFitToWidth}
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
              activeDrawTool !== null ||
              isSlideshowActive ||
              isModalOpen ||
              isEditingLocked ||
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
          {!isEditingLocked && viewMode === 'edit' && (
          <div
            role="toolbar"
            aria-label="Thanh công cụ chỉnh sửa slide di động"
            className="md:hidden fixed bottom-0 inset-x-0 z-40 h-12 bg-white border-t border-slate-200/80 shadow-[0_-2px_10px_rgba(15,23,42,0.06)] flex items-center text-slate-800"
          >
            <div className="overflow-x-auto no-scrollbar flex items-center gap-1 px-2 flex-1 min-w-0">
              {/* + Khung chữ */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleAddTextBlock}
                aria-label="Thêm khung chữ"
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-[#FFF4ED] border border-[#C43E1C]/35 text-[11px] font-semibold text-[#C43E1C] whitespace-nowrap"
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
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-50 border border-slate-200 text-[11px] font-medium text-slate-700 whitespace-nowrap"
              >
                <ImagePlus className="h-3.5 w-3.5 text-emerald-600" />
                <span>Ảnh</span>
              </button>

              <span className="h-5 w-px bg-slate-200 shrink-0 mx-0.5" />

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
                    ? 'bg-[#FFF4ED] border-[#C43E1C]/50 text-[#C43E1C]'
                    : 'bg-slate-50 border-slate-200 text-slate-700'
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
                    ? 'bg-[#FFF4ED] border-[#C43E1C]/50 text-[#C43E1C]'
                    : 'bg-slate-50 border-slate-200 text-slate-700'
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
                    ? 'bg-[#FFF4ED] border-[#C43E1C]/50 text-[#C43E1C]'
                    : 'bg-slate-50 border-slate-200 text-slate-700'
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
                className="shrink-0 inline-flex flex-col items-center justify-center h-8 w-8 rounded-lg bg-slate-50 border border-slate-200 text-slate-800"
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
                        ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                        : 'bg-slate-50 border-slate-200 text-slate-700'
                    }`}
                  >
                    <Icon className="h-3.5 w-3.5" />
                  </button>
                );
              })}

              <span className="h-5 w-px bg-slate-200 shrink-0 mx-0.5" />

              {/* + Slide */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleAddSlide}
                aria-label="Thêm slide mới"
                className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-50 border border-slate-200 text-[11px] font-semibold text-slate-800 whitespace-nowrap"
              >
                <Plus className="h-3.5 w-3.5 text-[#C43E1C]" />
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
              className="shrink-0 h-full w-11 border-l border-slate-200 flex items-center justify-center text-[#C43E1C] hover:bg-[#FFF4ED] transition-colors"
            >
              <ChevronUp className="h-5 w-5" />
            </button>
          </div>
          )}

          {/* Mobile Expanded Detail Bottom Sheet (md:hidden) */}
          {!isEditingLocked && viewMode === 'edit' && isMobileSheetOpen && (
            <div
              role="region"
              aria-label="Bảng công cụ chỉnh sửa slide chi tiết"
              className="md:hidden fixed bottom-0 inset-x-0 z-40 h-[44vh] bg-white border-t border-slate-200 rounded-t-2xl shadow-2xl flex flex-col text-slate-800"
            >
              {/* Sheet Header */}
              <div className="shrink-0 px-3.5 py-2.5 border-b border-slate-200 flex items-center justify-between gap-2 bg-[#F8FAFC] rounded-t-2xl">
                <div className="flex items-center gap-2 min-w-0">
                  <span className="h-2 w-2 rounded-full bg-[#C43E1C] shrink-0" />
                  <span className="text-xs font-semibold text-slate-900 truncate">
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
                  className="shrink-0 inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-medium text-slate-700"
                >
                  <span>Thu gọn</span>
                  <ChevronDown className="h-4 w-4" />
                </button>
              </div>

              {/* Sheet Body */}
              <div className="flex-1 min-h-0 overflow-y-auto scrollbar-none p-3.5 space-y-3.5">
                {/* Section 1: Định dạng khung chữ / ảnh */}
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                      Định dạng khung đang chọn
                    </span>
                    {selectedBlock && (
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleDeleteBlock(selectedBlock.id)}
                        className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-rose-50 border border-rose-200 text-[11px] font-semibold text-rose-600"
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
                        className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-rose-50 border border-rose-200 text-[11px] font-semibold text-rose-600"
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
                            ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                            : 'bg-slate-50 border-slate-200 text-slate-700'
                        }`}
                      >
                        {label}
                      </button>
                    ))}

                    <span className="h-5 w-px bg-slate-200 shrink-0" />

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
                              ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                              : 'bg-slate-50 border-slate-200 text-slate-700'
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
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
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
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
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
                          ? 'bg-[#FFF4ED] border-[#C43E1C] text-[#C43E1C]'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
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
                              ? 'bg-[#C43E1C] border-[#C43E1C] text-white'
                              : 'bg-slate-50 border-slate-200 text-slate-700'
                          }`}
                        >
                          <Icon className="h-4 w-4" />
                        </button>
                      );
                    })}
                  </div>

                  {/* Row 1C: Text Color Swatches + Block Fill Color */}
                  <div className="flex items-center gap-2 overflow-x-auto no-scrollbar pt-0.5">
                    <span className="text-[11px] text-slate-500 shrink-0">
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
                              ? 'ring-2 ring-[#C43E1C] border-white scale-110'
                              : 'border-slate-300'
                          }`}
                          style={{ backgroundColor: swatch.value }}
                        />
                      );
                    })}

                    <span className="h-5 w-px bg-slate-200 shrink-0 mx-0.5" />

                    <span className="text-[11px] text-slate-500 shrink-0">
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
                        className="shrink-0 px-2 py-1 rounded bg-slate-100 text-[11px] text-slate-700"
                      >
                        Trong suốt
                      </button>
                    )}
                  </div>
                </div>

                {/* Section 2: Chèn & Quản lý Slide */}
                <div className="space-y-1.5 pt-1 border-t border-slate-200">
                  <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                    Chèn & Quản lý Slide
                  </span>
                  <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddTextBlock}
                      className="shrink-0 h-9 px-3 rounded-xl bg-[#FFF4ED] border border-[#C43E1C]/35 text-xs font-semibold text-[#C43E1C] inline-flex items-center gap-1.5"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      <Type className="h-3.5 w-3.5" />
                      <span>+ Khung chữ</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleTriggerAddImage}
                      className="shrink-0 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-medium text-slate-700 inline-flex items-center gap-1.5"
                    >
                      <ImagePlus className="h-3.5 w-3.5 text-emerald-600" />
                      <span>+ Ảnh</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleAddSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-medium text-slate-700 inline-flex items-center gap-1.5"
                    >
                      <Plus className="h-3.5 w-3.5 text-[#C43E1C]" />
                      <span>+ Slide</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleDuplicateSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-medium text-slate-700 inline-flex items-center gap-1.5"
                    >
                      <Copy className="h-3.5 w-3.5" />
                      <span>Nhân bản</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      disabled={slides.length <= 1}
                      onClick={handleDeleteSlide}
                      className="shrink-0 h-9 px-3 rounded-xl bg-rose-50 border border-rose-200 text-xs font-medium text-rose-600 inline-flex items-center gap-1.5 disabled:opacity-35"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      <span>Xóa Slide</span>
                    </button>
                  </div>
                </div>

                {/* Section 3: Màu nền Slide & Tiện ích */}
                <div className="space-y-1.5 pt-1 border-t border-slate-200">
                  <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
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
                              ? 'border-[#C43E1C] bg-[#FFF4ED] text-[#C43E1C]'
                              : 'border-slate-200 bg-slate-50 text-slate-700'
                          }`}
                        >
                          <span
                            className="h-3.5 w-3.5 rounded border border-slate-300"
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
                          ? 'bg-rose-50 border-rose-300 text-rose-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
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
                          ? 'bg-[#FFF4ED] border-[#C43E1C]/40 text-[#C43E1C]'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
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
          className="fixed z-50 w-48 rounded-xl bg-white border border-slate-200 shadow-2xl py-1.5 text-xs text-slate-800"
        >
          <button
            type="button"
            disabled={isEditingLocked}
            onClick={() => {
              setActiveSlideIdx(thumbContextMenu.slideIdx);
              handleAddSlide();
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#FFF4ED] hover:text-[#C43E1C] flex items-center gap-2 disabled:opacity-40"
          >
            <Plus className="h-3.5 w-3.5 text-[#C43E1C]" />
            <span>Thêm slide mới</span>
          </button>
          <button
            type="button"
            disabled={isEditingLocked}
            onClick={() => {
              handleDuplicateSlideAt(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#FFF4ED] hover:text-[#C43E1C] flex items-center gap-2 disabled:opacity-40"
          >
            <Copy className="h-3.5 w-3.5 text-slate-500" />
            <span>Nhân bản slide</span>
          </button>
          <button
            type="button"
            disabled={isEditingLocked}
            onClick={() => {
              handleToggleHideSlide(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-[#FFF4ED] hover:text-[#C43E1C] flex items-center gap-2 disabled:opacity-40"
          >
            <EyeOff className="h-3.5 w-3.5 text-amber-600" />
            <span>
              {slides[thumbContextMenu.slideIdx]?.hidden
                ? 'Hiện slide khi chiếu'
                : 'Ẩn slide khi chiếu'}
            </span>
          </button>
          <div className="my-1 border-t border-slate-200" />
          <button
            type="button"
            disabled={isEditingLocked || slides.length <= 1}
            onClick={() => {
              handleDeleteSlideAt(thumbContextMenu.slideIdx);
              setThumbContextMenu(null);
            }}
            className="w-full px-3 py-1.5 text-left hover:bg-rose-50 text-rose-600 flex items-center gap-2 disabled:opacity-35"
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

      {/* Slide File Info Modal */}
      <SlideFileInfoModal
        isOpen={isFileInfoOpen}
        document={buildLiveDocument()}
        slideCount={slides.length}
        totalBlocks={slides.reduce((sum, s) => sum + (s.blocks?.length ?? 0), 0)}
        totalImages={slides.reduce((sum, s) => sum + (s.images?.length ?? 0), 0)}
        aspectRatioLabel={Math.abs(currentAspectRatio - 4 / 3) < 0.08 ? '4:3 (Chuẩn)' : '16:9 (Màn hình rộng)'}
        isProtected={isPresentationProtected}
        hasOriginalBinary={hasOriginalBinary}
        onClose={() => setIsFileInfoOpen(false)}
      />

      {/* PowerPoint AI Tools Drawer (Sinh dàn ý & tạo hàng loạt slide) */}
      <AIToolsDrawer
        isOpen={aiModalState.isOpen && aiModalState.initialTab === 'generate'}
        mode="pptx"
        initialAction="generate"
        isReadOnly={isEditingLocked}
        documentTitle={document.title}
        pptxSlideCount={slides.length}
        onPptxAppendSlides={(draftSlides) => {
          const refSlide = slides[safeSlideIdx] ?? slides[0];
          const accentColor = refSlide?.accentColor || '#C43E1C';
          const backgroundColor = refSlide?.backgroundColor || '#FFFFFF';
          const backgroundGradient = refSlide?.backgroundGradient;
          const slideWidthRatio = refSlide?.slideWidthRatio;
          const stamp = Date.now();

          const created: SlidePreviewItem[] = draftSlides.map((draft, idx) => {
            const slideId = `ai-slide-${stamp}-${idx + 1}`;
            const isHero = idx === 0;
            const blocks: SlideTextBlock[] = isHero
              ? [
                  {
                    id: `${slideId}-accent-bar`,
                    text: '',
                    fontSize: 12,
                    kind: 'shape',
                    shapeKind: 'rect',
                    fillColor: accentColor,
                    borderColor: accentColor,
                    borderWidthPx: 0,
                    xPct: 35,
                    yPct: 22,
                    wPct: 30,
                    hPct: 1.5,
                  },
                  {
                    id: `${slideId}-title`,
                    text: draft.title,
                    fontSize: 36,
                    bold: true,
                    color: '#1E293B',
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
                    text: draft.subtitle || draft.bullets.join(' · '),
                    fontSize: 18,
                    color: '#475569',
                    align: 'center',
                    valign: 'top',
                    role: 'subtitle',
                    xPct: 14,
                    yPct: 52,
                    wPct: 72,
                    hPct: 16,
                  },
                  {
                    id: `${slideId}-points`,
                    text: draft.bullets.map((b) => `• ${b}`).join('\n'),
                    fontSize: 15,
                    color: '#334155',
                    align: 'center',
                    valign: 'top',
                    role: 'body',
                    xPct: 14,
                    yPct: 70,
                    wPct: 72,
                    hPct: 22,
                  },
                ]
              : [
                  {
                    id: `${slideId}-accent-bar`,
                    text: '',
                    fontSize: 12,
                    kind: 'shape',
                    shapeKind: 'rect',
                    fillColor: accentColor,
                    borderColor: accentColor,
                    borderWidthPx: 0,
                    xPct: 6,
                    yPct: 8,
                    wPct: 1.2,
                    hPct: 12,
                  },
                  {
                    id: `${slideId}-title`,
                    text: draft.title,
                    fontSize: 28,
                    bold: true,
                    color: '#1E293B',
                    align: 'left',
                    valign: 'middle',
                    role: 'title',
                    xPct: 9,
                    yPct: 7,
                    wPct: 85,
                    hPct: 14,
                  },
                  ...(draft.subtitle
                    ? [
                        {
                          id: `${slideId}-subtitle`,
                          text: draft.subtitle,
                          fontSize: 15,
                          color: accentColor,
                          bold: true,
                          align: 'left' as const,
                          valign: 'top' as const,
                          role: 'subtitle' as const,
                          xPct: 9,
                          yPct: 22,
                          wPct: 85,
                          hPct: 8,
                        },
                      ]
                    : []),
                  {
                    id: `${slideId}-body`,
                    text: draft.bullets.map((b) => `• ${b}`).join('\n'),
                    fontSize: 18,
                    lineHeight: 1.55,
                    color: '#334155',
                    align: 'left',
                    valign: 'top',
                    role: 'body',
                    xPct: 9,
                    yPct: draft.subtitle ? 32 : 25,
                    wPct: 84,
                    hPct: 58,
                  },
                ];

            return {
              id: slideId,
              index: slides.length + idx + 1,
              title: draft.title,
              summary: draft.bullets.join(' · '),
              bullets: draft.bullets,
              speakerNote: draft.speakerNote,
              accentColor,
              backgroundColor,
              backgroundGradient,
              slideWidthRatio,
              blocks,
            };
          });
          handleAppendAiSlides(created);
        }}
        onPptxOpenOtherAiTab={(tab) =>
          setAiModalState({ isOpen: true, initialTab: tab })
        }
        onClose={() => setAiModalState((prev) => ({ ...prev, isOpen: false }))}
      />

      {/* Slide AI Tools Modal (Tóm tắt / Speaker Notes) */}
      <SlideAiToolsModal
        isOpen={aiModalState.isOpen && aiModalState.initialTab !== 'generate'}
        initialTab={aiModalState.initialTab}
        documentTitle={document.title}
        slides={slides}
        activeSlideIndex={safeSlideIdx}
        isReadOnly={isEditingLocked}
        onAppendSlides={handleAppendAiSlides}
        onUpdateSlideNote={handleUpdateSlideNoteAt}
        onApplyAllEmptyNotes={handleApplyAllEmptyAiNotes}
        onClose={() => setAiModalState((prev) => ({ ...prev, isOpen: false }))}
      />

      {/* Slide Comments Floating Side Drawer */}
      {isCommentsOpen && (
        <div className="fixed inset-y-14 right-3 z-40 w-[320px] max-w-[92vw] flex flex-col justify-center pointer-events-none">
          <div className="pointer-events-auto h-[520px] max-h-[82vh] rounded-2xl overflow-hidden shadow-2xl border border-slate-200 bg-white">
            <SlideCommentsPanel
              slideIndex={safeSlideIdx}
              slideTitle={currentSlide.title}
              comments={currentSlide.comments ?? []}
              isReadOnly={isEditingLocked}
              onAddComment={handleAddSlideComment}
              onDeleteComment={handleDeleteSlideComment}
              onClose={() => setIsCommentsOpen(false)}
            />
          </div>
        </div>
      )}

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
@keyframes liteObjPulse {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.06); }
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

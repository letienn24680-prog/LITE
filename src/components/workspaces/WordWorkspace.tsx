import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, MouseEvent as ReactMouseEvent } from 'react';
import { renderAsync } from 'docx-preview';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  BookOpen,
  Check,
  ChevronDown,
  ChevronRight,
  ChevronUp,
  CloudDownload,
  Download,
  FolderOpen,
  Highlighter,
  Italic,
  Lightbulb,
  List,
  ListOrdered,
  Loader2,
  Minus,
  Moon,
  Palette,
  PanelLeftClose,
  PanelLeftOpen,
  Quote,
  Redo2,
  RefreshCw,
  RemoveFormatting,
  Save,
  ScanText,
  Smartphone,
  Sparkles,
  SplitSquareVertical,
  Strikethrough,
  Sun,
  Table,
  Underline,
  Undo2,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  WordPaperTheme,
} from '../../types/document';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import {
  extractOutlineFromHtml,
  isLikelyStructuralOutlineTitle,
} from '../../utils/fileParser';
import { exportWordToDocx } from '../../utils/fileExporter';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { FloatingSelectionToolbar } from './FloatingSelectionToolbar';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { WorkspaceTopBar } from './WorkspaceTopBar';

export interface WordWorkspaceProps {
  document: DocumentItem;
  onBackHome: () => void;
  onOpenModal: (
    modal: Exclude<ModalType, null>,
    overrideDoc?: DocumentItem
  ) => void;
  onSaveDocument: (updatedDoc: DocumentItem) => void;
  isModalOpen?: boolean;
  defaultPaperTheme?: WordPaperTheme;
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

type WordPageLayout = 'print' | 'mobile';
const WORD_PAGE_LAYOUT_STORAGE_KEY = 'lite_word_page_layout';

function readStoredPageLayout(): WordPageLayout {
  try {
    const saved = window.localStorage.getItem(WORD_PAGE_LAYOUT_STORAGE_KEY);
    return saved === 'mobile' ? 'mobile' : 'print';
  } catch {
    return 'print';
  }
}

function saveStoredPageLayout(layout: WordPageLayout): void {
  try {
    window.localStorage.setItem(WORD_PAGE_LAYOUT_STORAGE_KEY, layout);
  } catch {
    // Ignore storage quota or private mode errors
  }
}

const FONT_SIZES = [12, 13, 14, 16, 18, 20, 24] as const;
type FontSizeOption = (typeof FONT_SIZES)[number];

const FONT_FAMILIES = ['Times New Roman', 'Arial'] as const;
type WordFontFamily = (typeof FONT_FAMILIES)[number];

const LINE_HEIGHTS = [1.15, 1.5, 2.0] as const;
type WordLineHeight = (typeof LINE_HEIGHTS)[number];

const ZOOM_PERCENTS = [85, 100, 115, 125] as const;
type WordZoomPercent = (typeof ZOOM_PERCENTS)[number];

type RibbonTab = 'home' | 'insert' | 'layout' | 'view';
type BlockTag = 'p' | 'h1' | 'h2' | 'h3';
type ActivePopover = 'block' | 'color' | 'highlight' | null;

const RIBBON_TABS: { key: RibbonTab; label: string }[] = [
  { key: 'home', label: 'Trang chủ' },
  { key: 'insert', label: 'Chèn' },
  { key: 'layout', label: 'Bố trí' },
  { key: 'view', label: 'Xem' },
];

const RIBBON_TAB_LABELS: Record<RibbonTab, string> = {
  home: 'Trang chủ',
  insert: 'Chèn',
  layout: 'Bố trí',
  view: 'Xem',
};

interface OutlineNavEntry {
  section: OutlineSection;
  level: 1 | 2 | 3;
}

interface FormatActiveState {
  block: BlockTag;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikeThrough: boolean;
  align: 'left' | 'center' | 'right' | 'justify';
  unorderedList: boolean;
  orderedList: boolean;
}

interface PaperThemeClasses {
  label: string;
  shortLabel: string;
  swatchClass: string;
  outerBg: string;
  ribbonBg: string;
  ribbonBorder: string;
  sidebarBg: string;
  sidebarBorder: string;
  paperBg: string;
  paperBorder: string;
  pageHexBg: string;
  pageHexText: string;
  textPrimary: string;
  textMuted: string;
  outlineActive: string;
  outlineIdle: string;
  calloutBg: string;
  calloutBorder: string;
  editorProse: string;
  scrollbarClass: string;
}

const PAPER_THEMES: Record<WordPaperTheme, PaperThemeClasses> = {
  light: {
    label: 'Trắng sáng',
    shortLabel: 'Sáng',
    swatchClass: 'bg-white border-slate-300',
    outerBg: 'bg-[#F3F1EC]',
    ribbonBg: 'bg-[#FFFFFF]',
    ribbonBorder: 'border-[#EAE6DF]',
    sidebarBg: 'bg-[#FAF9F6]',
    sidebarBorder: 'border-[#E6E2DA]',
    paperBg: 'bg-[#FFFFFF]',
    paperBorder: 'border-[#E7E3DC]',
    pageHexBg: '#FFFFFF',
    pageHexText: '#1E293B',
    textPrimary: 'text-[#1E293B]',
    textMuted: 'text-slate-500',
    outlineActive:
      'bg-indigo-50 text-indigo-800 font-semibold ring-1 ring-indigo-200/80',
    outlineIdle: 'text-slate-600 hover:bg-[#EFECE6] hover:text-slate-900',
    calloutBg: 'bg-[#F8F7F4]',
    calloutBorder: 'border-[#E7E3DC]',
    editorProse:
      '[&_h1]:text-slate-900 [&_h2]:text-slate-900 [&_h3]:text-slate-800 [&_h2]:border-[#ECE8E0] [&_blockquote]:bg-[#FAF8F5] [&_blockquote]:border-indigo-400 [&_blockquote]:text-slate-700 [&_th]:bg-[#F8F7F4] [&_th]:border-[#DDD7CE] [&_td]:border-[#E5E0D8] [&_a]:text-indigo-600',
    scrollbarClass: '',
  },
  sepia: {
    label: 'Giấy kem ấm',
    shortLabel: 'Kem ấm',
    swatchClass: 'bg-[#F7EFE5] border-[#CBB99F]',
    outerBg: 'bg-[#EFE6D8]',
    ribbonBg: 'bg-[#FAF4EA]',
    ribbonBorder: 'border-[#E4D8C5]',
    sidebarBg: 'bg-[#F4EBDD]',
    sidebarBorder: 'border-[#E2D4BF]',
    paperBg: 'bg-[#F7EFE5]',
    paperBorder: 'border-[#E5D7C2]',
    pageHexBg: '#F7EFE5',
    pageHexText: '#2C2620',
    textPrimary: 'text-[#2C2620]',
    textMuted: 'text-[#7A6B5B]',
    outlineActive:
      'bg-[#EAE0D0] text-[#221C16] font-semibold ring-1 ring-[#D7C4A8]',
    outlineIdle: 'text-[#5D5043] hover:bg-[#EBE1D2] hover:text-[#2C2620]',
    calloutBg: 'bg-[#EFE5D7]',
    calloutBorder: 'border-[#E2D4BF]',
    editorProse:
      '[&_h1]:text-[#221C16] [&_h2]:text-[#221C16] [&_h3]:text-[#3A3127] [&_h2]:border-[#E3D5BF] [&_blockquote]:bg-[#EFE5D7] [&_blockquote]:border-amber-600/60 [&_blockquote]:text-[#433A30] [&_th]:bg-[#EFE5D7] [&_th]:border-[#D9C9B1] [&_td]:border-[#DFD1BC] [&_a]:text-indigo-700',
    scrollbarClass: 'scrollbar-light',
  },
  dark: {
    label: 'Nền tối',
    shortLabel: 'Tối',
    swatchClass: 'bg-[#1E2430] border-slate-500',
    outerBg: 'bg-[#12161F]',
    ribbonBg: 'bg-[#171C27]',
    ribbonBorder: 'border-white/10',
    sidebarBg: 'bg-[#161B26]',
    sidebarBorder: 'border-white/10',
    paperBg: 'bg-[#1E2430]',
    paperBorder: 'border-white/10',
    pageHexBg: '#1E2430',
    pageHexText: '#E2E8F0',
    textPrimary: 'text-[#E2E8F0]',
    textMuted: 'text-slate-400',
    outlineActive:
      'bg-indigo-500/20 text-white font-semibold ring-1 ring-indigo-400/35',
    outlineIdle: 'text-slate-300 hover:bg-white/[0.06] hover:text-white',
    calloutBg: 'bg-[#171C27]',
    calloutBorder: 'border-white/10',
    editorProse:
      '[&_h1]:text-white [&_h2]:text-white [&_h3]:text-slate-100 [&_h2]:border-white/10 [&_blockquote]:bg-white/[0.04] [&_blockquote]:border-indigo-400/70 [&_blockquote]:text-slate-300 [&_th]:bg-white/[0.06] [&_th]:border-white/15 [&_td]:border-white/15 [&_a]:text-indigo-300',
    scrollbarClass: 'scrollbar-dark',
  },
};

const BLOCK_STYLE_OPTIONS: {
  tag: BlockTag;
  label: string;
  previewClass: string;
}[] = [
  { tag: 'p', label: 'Văn bản thường', previewClass: 'text-xs font-normal' },
  { tag: 'h1', label: 'Tiêu đề 1', previewClass: 'text-base font-bold' },
  { tag: 'h2', label: 'Tiêu đề 2', previewClass: 'text-sm font-bold' },
  { tag: 'h3', label: 'Tiêu đề 3', previewClass: 'text-xs font-semibold' },
];

const TEXT_COLOR_SWATCHES = [
  { label: 'Đen than', value: '#1E293B' },
  { label: 'Xám đậm', value: '#475569' },
  { label: 'Trắng sáng', value: '#F8FAFC' },
  { label: 'Indigo', value: '#4F46E5' },
  { label: 'Xanh dương', value: '#2563EB' },
  { label: 'Xanh lá', value: '#15803D' },
  { label: 'Cam đậm', value: '#EA580C' },
  { label: 'Đỏ nổi bật', value: '#DC2626' },
] as const;

const HIGHLIGHT_SWATCHES = [
  {
    label: 'Xóa nền',
    value: 'transparent',
    chipClass: 'bg-transparent border-dashed',
  },
  { label: 'Vàng', value: '#FEF08A', chipClass: 'bg-[#FEF08A]' },
  { label: 'Xanh lá', value: '#BBF7D0', chipClass: 'bg-[#BBF7D0]' },
  { label: 'Xanh dương', value: '#BFDBFE', chipClass: 'bg-[#BFDBFE]' },
  { label: 'Hồng', value: '#FBCFE8', chipClass: 'bg-[#FBCFE8]' },
] as const;

const INITIAL_FORMAT_STATE: FormatActiveState = {
  block: 'p',
  bold: false,
  italic: false,
  underline: false,
  strikeThrough: false,
  align: 'left',
  unorderedList: false,
  orderedList: false,
};

const DOCX_PREVIEW_SCOPED_CSS = `
.lite-word-host .lite-docx-wrapper {
  background: transparent !important;
  padding: 6px 0 28px 0 !important;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 24px;
}
.lite-word-host section.lite-docx {
  background: var(--lite-doc-page-bg, #ffffff) !important;
  color: var(--lite-doc-page-text, #1e293b);
  font-family: var(--lite-doc-font-family, 'Times New Roman', Times, serif) !important;
  line-height: var(--lite-doc-line-height, 1.5);
  box-shadow: 0 12px 32px -12px rgba(15, 23, 42, 0.18), 0 2px 8px -2px rgba(15, 23, 42, 0.08) !important;
  border: 1px solid rgba(148, 163, 184, 0.22);
  border-radius: 6px;
  margin: 0 auto !important;
  max-width: 100%;
  box-sizing: border-box;
}
.lite-word-host section.lite-docx p,
.lite-word-host section.lite-docx li,
.lite-word-host section.lite-docx td,
.lite-word-host section.lite-docx th {
  font-family: var(--lite-doc-font-family, 'Times New Roman', Times, serif);
  line-height: var(--lite-doc-line-height, 1.5);
}
.lite-word-host section.lite-docx table {
  max-width: 100%;
}
.lite-word-host section.lite-docx img {
  max-width: 100%;
  height: auto;
}
.lite-word-host [id^="heading-"] {
  scroll-margin-top: 24px;
}
@media (max-width: 840px) {
  .lite-word-host .lite-docx-wrapper {
    padding: 4px 0 20px 0 !important;
    gap: 16px;
  }
  .lite-word-host section.lite-docx {
    width: 100% !important;
    min-width: 0 !important;
    padding: 24px 16px !important;
    overflow-x: auto;
  }
}
`;

function getFontStack(family: WordFontFamily): string {
  return family === 'Arial'
    ? 'Arial, Helvetica, sans-serif'
    : "'Times New Roman', Times, serif";
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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

function stripInjectedStylesFromHtml(
  container: HTMLElement | null,
  fallbackHtml: string
): string {
  if (!container) {
    return fallbackHtml;
  }
  const clone = container.cloneNode(true) as HTMLElement;
  const styleNodes = clone.querySelectorAll('style');
  styleNodes.forEach((node) => node.remove());
  const cleaned = clone.innerHTML.trim();
  return cleaned.length > 0 ? cleaned : fallbackHtml;
}

function inferDomOutlineLevel(node: HTMLElement, text: string): 1 | 2 | 3 {
  const tag = node.tagName.toLowerCase();
  const cls = (
    typeof node.className === 'string' ? node.className : ''
  ).toLowerCase();

  if (tag === 'h1' || /(?:^|\s|_)heading[-_ ]?1\b/.test(cls)) {
    return 1;
  }
  if (tag === 'h2' || /(?:^|\s|_)heading[-_ ]?2\b/.test(cls)) {
    return 2;
  }
  if (
    tag === 'h3' ||
    tag === 'h4' ||
    /(?:^|\s|_)heading[-_ ]?[34]\b/.test(cls)
  ) {
    return 3;
  }

  if (/^\d+\.\d+\.\d+/.test(text)) {
    return 3;
  }
  if (/^\d+\.\d+\b/.test(text)) {
    return 2;
  }
  return 1;
}

function buildInitialWordHtml(doc: DocumentItem): string {
  if (doc.wordHtml && doc.wordHtml.trim().length > 0) {
    return extractOutlineFromHtml(doc.wordHtml).htmlWithIds;
  }

  const parts: string[] = [];
  parts.push(`<h1>${escapeHtml(doc.title)}</h1>`);

  if (doc.contentPreview && doc.contentPreview.trim().length > 0) {
    parts.push(
      `<blockquote>${escapeHtml(doc.contentPreview.trim())}</blockquote>`
    );
  }

  if (doc.outline.length > 0) {
    for (const section of doc.outline) {
      parts.push(
        `<h2 id="${escapeHtml(section.id)}">${escapeHtml(section.romanIndex)}. ${escapeHtml(section.title)}</h2>`
      );
      const customParagraphs = doc.wordParagraphs?.[section.id];
      const paragraphs =
        customParagraphs && customParagraphs.length > 0
          ? customParagraphs
          : section.summary
            ? [section.summary]
            : [];

      for (const paragraph of paragraphs) {
        parts.push(`<p>${escapeHtml(paragraph)}</p>`);
      }

      if (section.keyPoints.length > 0) {
        const itemsHtml = section.keyPoints
          .map((point) => `<li>${escapeHtml(point)}</li>`)
          .join('');
        parts.push(`<ul>${itemsHtml}</ul>`);
      }
    }
  } else if (!doc.contentPreview) {
    parts.push(
      '<p>Bắt đầu soạn thảo nội dung tài liệu của bạn tại đây...</p>'
    );
  }

  return extractOutlineFromHtml(parts.join('')).htmlWithIds;
}

function buildQuickTableHtml(rows: number, cols: number): string {
  const safeRows = Math.max(1, rows);
  const safeCols = Math.max(1, cols);
  const headCells = Array.from(
    { length: safeCols },
    (_, cIdx) => `<th>Cột ${cIdx + 1}</th>`
  ).join('');
  const bodyRows = Array.from({ length: Math.max(1, safeRows - 1) }, (_, rIdx) => {
    const cells = Array.from(
      { length: safeCols },
      (__, cIdx) => `<td>Nội dung ${rIdx + 1}.${cIdx + 1}</td>`
    ).join('');
    return `<tr>${cells}</tr>`;
  }).join('');

  return `<table><thead><tr>${headCells}</tr></thead><tbody>${bodyRows}</tbody></table><p><br></p>`;
}

export function WordWorkspace({
  document,
  onBackHome,
  onOpenModal,
  onSaveDocument,
  isModalOpen = false,
  defaultPaperTheme = 'light',
  autoSaveEnabled = false,
  isHydratingCloudFile = false,
  onRetryCloudDownload,
  collaborators,
  selfSessionId,
  collabPermission,
  onPresenceUpdate,
  onBroadcastPatch,
  incomingRemotePatch,
}: WordWorkspaceProps) {
  const isReadOnly = collabPermission === 'view';

  const [paperTheme, setPaperTheme] =
    useState<WordPaperTheme>(defaultPaperTheme);
  const [pageLayout, setPageLayout] = useState<WordPageLayout>(() =>
    readStoredPageLayout()
  );
  const [printZoom, setPrintZoom] = useState<number | null>(null);
  const [containerWidth, setContainerWidth] = useState<number>(0);
  const [contentHeight, setContentHeight] = useState<number>(0);
  const [fontFamily, setFontFamily] =
    useState<WordFontFamily>('Times New Roman');
  const [fontSizeIndex, setFontSizeIndex] = useState<number>(3); // 16px default in [12,13,14,16,18,20,24]
  const [lineHeight, setLineHeight] = useState<WordLineHeight>(1.5);
  const [zoomPercent, setZoomPercent] = useState<WordZoomPercent>(100);
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTab>('home');

  const [isDesktopSidebarOpen, setIsDesktopSidebarOpen] =
    useState<boolean>(true);
  const [isMobileDrawerOpen, setIsMobileDrawerOpen] = useState<boolean>(false);
  const [isMobileSheetExpanded, setIsMobileSheetExpanded] =
    useState<boolean>(false);
  const [isMobileTabMenuOpen, setIsMobileTabMenuOpen] =
    useState<boolean>(false);
  const [isMobileCommandSearchOpen, setIsMobileCommandSearchOpen] =
    useState<boolean>(false);
  const [mobileCommandQuery, setMobileCommandQuery] = useState<string>('');
  const [mobileExpandPanel, setMobileExpandPanel] = useState<
    'highlight' | 'color' | null
  >(null);

  const [outlineEntries, setOutlineEntries] = useState<OutlineNavEntry[]>([]);
  const [activeSectionId, setActiveSectionId] = useState<string>('');
  const [isDirty, setIsDirty] = useState<boolean>(false);
  const [editRevision, setEditRevision] = useState<number>(0);
  const [saveFeedback, setSaveFeedback] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [openPopover, setOpenPopover] = useState<ActivePopover>(null);
  const [customTextColor, setCustomTextColor] = useState<string>('#2563EB');
  const [currentHighlightColor, setCurrentHighlightColor] =
    useState<string>('#FEF08A');
  const [formatState, setFormatState] =
    useState<FormatActiveState>(INITIAL_FORMAT_STATE);
  const [wordCount, setWordCount] = useState<number>(0);
  const [detectedPageCount, setDetectedPageCount] = useState<number>(1);
  const [isRenderingDocx, setIsRenderingDocx] = useState<boolean>(false);
  const [isDocxPreviewRendered, setIsDocxPreviewRendered] =
    useState<boolean>(false);
  const [docxRenderFailed, setDocxRenderFailed] = useState<boolean>(false);

  const editorRef = useRef<HTMLDivElement | null>(null);
  const pageSurfaceRef = useRef<HTMLDivElement | null>(null);
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const ribbonRef = useRef<HTMLDivElement | null>(null);
  const isComposingRef = useRef<boolean>(false);
  const isApplyingRemotePatchRef = useRef<boolean>(false);
  const lastRemotePatchAtRef = useRef<number>(0);
  const typingPresenceTimerRef = useRef<number | null>(null);
  const broadcastThrottleTimerRef = useRef<number | null>(null);
  const lastBroadcastAtRef = useRef<number>(0);
  const pendingBroadcastHtmlRef = useRef<string | null>(null);
  const onPresenceUpdateRef = useRef(onPresenceUpdate);
  onPresenceUpdateRef.current = onPresenceUpdate;
  const onBroadcastPatchRef = useRef(onBroadcastPatch);
  onBroadcastPatchRef.current = onBroadcastPatch;
  const loadedDocIdRef = useRef<string | null>(null);
  const renderedRawBufferRef = useRef<ArrayBuffer | null>(null);
  const renderSeqRef = useRef<number>(0);
  const latestHtmlRef = useRef<string>('');
  const savedRangeRef = useRef<Range | null>(null);
  const saveFeedbackTimerRef = useRef<number | null>(null);
  const exportErrorTimerRef = useRef<number | null>(null);
  const documentRef = useRef<DocumentItem>(document);
  documentRef.current = document;

  const fontSize: FontSizeOption = FONT_SIZES[fontSizeIndex];
  const effectiveFontSize = Number(((fontSize * zoomPercent) / 100).toFixed(1));
  const themeStyle = PAPER_THEMES[paperTheme];
  const isDark = paperTheme === 'dark';

  const isNarrowViewport = containerWidth > 0 && containerWidth < 768;
  const fitScale = isNarrowViewport
    ? Math.min(1, Math.max(0.32, (containerWidth - 24) / 794))
    : 1;
  const isScaledPrintMode = pageLayout === 'print' && isNarrowViewport;
  const effectivePrintScale = isScaledPrintMode
    ? Math.min(1.5, Math.max(0.32, printZoom ?? fitScale))
    : 1;

  const handleSelectPageLayout = useCallback((nextLayout: WordPageLayout) => {
    setPageLayout(nextLayout);
    saveStoredPageLayout(nextLayout);
    if (nextLayout === 'mobile') {
      setPrintZoom(null);
    }
  }, []);

  const handleTogglePageLayout = useCallback(() => {
    setPageLayout((prev) => {
      const next: WordPageLayout = prev === 'print' ? 'mobile' : 'print';
      saveStoredPageLayout(next);
      if (next === 'mobile') {
        setPrintZoom(null);
      }
      return next;
    });
  }, []);

  const handleStepPrintZoom = useCallback(
    (delta: number) => {
      setPrintZoom((prev) => {
        const base = prev ?? fitScale;
        const next = Math.round((base + delta) * 100) / 100;
        return Math.min(1.5, Math.max(0.4, next));
      });
    },
    [fitScale]
  );

  // Scan structural headings directly from live DOM without recreating nodes (IME-safe)
  const syncOutlineFromDom = useCallback(() => {
    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }

    const pageSections = Array.from(
      editorEl.querySelectorAll('section.lite-docx')
    );
    const candidateNodes = Array.from(
      editorEl.querySelectorAll<HTMLElement>('h1, h2, h3, h4, p')
    );

    const usedIds = new Set<string>();
    const nextEntries: OutlineNavEntry[] = [];
    let counter = 1;

    for (const node of candidateNodes) {
      if (
        node.closest(
          'header, footer, article.lite-docx-header, article.lite-docx-footer'
        )
      ) {
        continue;
      }

      const tagLower = node.tagName.toLowerCase();
      const classAttr =
        typeof node.className === 'string' ? node.className.toLowerCase() : '';
      const isHeadingClass = /(?:^|\s|_)heading[-_ ]?[1-4]\b/.test(classAttr);
      const isExplicitHeadingTag =
        tagLower === 'h1' ||
        tagLower === 'h2' ||
        tagLower === 'h3' ||
        tagLower === 'h4' ||
        isHeadingClass;

      if (!isExplicitHeadingTag && node.closest('td, th, li')) {
        continue;
      }

      const text = (node.textContent ?? '').replace(/\s+/g, ' ').trim();
      if (!isLikelyStructuralOutlineTitle(text, isExplicitHeadingTag)) {
        continue;
      }

      let id = node.id.trim();
      if (!id || usedIds.has(id)) {
        while (usedIds.has(`heading-${counter}`)) {
          counter += 1;
        }
        id = `heading-${counter}`;
        node.id = id;
        counter += 1;
      }
      usedIds.add(id);

      const level = inferDomOutlineLevel(node, text);
      const enclosingPage = node.closest('section.lite-docx');
      const pageIdx =
        enclosingPage instanceof HTMLElement
          ? pageSections.indexOf(enclosingPage)
          : -1;
      const pageNumber =
        pageIdx >= 0
          ? pageIdx + 1
          : Math.max(1, Math.floor(nextEntries.length / 2) + 1);

      nextEntries.push({
        level,
        section: {
          id,
          romanIndex:
            level === 1
              ? toRomanNumeral(
                  nextEntries.filter((entry) => entry.level === 1).length + 1
                )
              : String(nextEntries.length + 1),
          title: text,
          summary: '',
          keyPoints: [],
          pageNumber,
        },
      });
    }

    setOutlineEntries(nextEntries);
    setActiveSectionId((prev) => {
      if (prev && nextEntries.some((entry) => entry.section.id === prev)) {
        return prev;
      }
      return nextEntries[0]?.section.id ?? '';
    });

    const clone = editorEl.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('style').forEach((s) => s.remove());
    const plainText = (clone.textContent ?? '').trim();
    const words = plainText ? plainText.split(/\s+/).filter(Boolean).length : 0;
    setWordCount(words);

    const pageBreaks = editorEl.querySelectorAll('[data-page-break="true"]').length;
    const totalPages =
      pageSections.length > 0
        ? pageSections.length
        : Math.max(1, Math.ceil(words / 350) + pageBreaks);
    setDetectedPageCount(totalPages);
  }, []);

  const captureCurrentSelection = useCallback(() => {
    const editorEl = editorRef.current;
    const selection = window.document.getSelection();
    if (!editorEl || !selection || selection.rangeCount === 0) {
      return;
    }
    const range = selection.getRangeAt(0);
    if (editorEl.contains(range.commonAncestorContainer)) {
      savedRangeRef.current = range.cloneRange();
    }
  }, []);

  const refreshFormatState = useCallback(() => {
    const editorEl = editorRef.current;
    const selection = window.document.getSelection();
    if (!editorEl || !selection || selection.rangeCount === 0) {
      return;
    }
    const anchorNode = selection.anchorNode;
    if (!anchorNode || !editorEl.contains(anchorNode)) {
      return;
    }

    let current: Node | null = anchorNode;
    let detectedBlock: BlockTag = 'p';
    while (current && current !== editorEl) {
      if (current instanceof HTMLElement) {
        const tag = current.tagName.toLowerCase();
        if (tag === 'h1' || tag === 'h2' || tag === 'h3' || tag === 'p') {
          detectedBlock = tag;
          break;
        }
      }
      current = current.parentNode;
    }

    let align: FormatActiveState['align'] = 'left';
    try {
      if (window.document.queryCommandState('justifyCenter')) {
        align = 'center';
      } else if (window.document.queryCommandState('justifyRight')) {
        align = 'right';
      } else if (window.document.queryCommandState('justifyFull')) {
        align = 'justify';
      }

      setFormatState({
        block: detectedBlock,
        bold: window.document.queryCommandState('bold'),
        italic: window.document.queryCommandState('italic'),
        underline: window.document.queryCommandState('underline'),
        strikeThrough: window.document.queryCommandState('strikeThrough'),
        align,
        unorderedList: window.document.queryCommandState('insertUnorderedList'),
        orderedList: window.document.queryCommandState('insertOrderedList'),
      });
    } catch {
      // Ignore queryCommandState exceptions in restricted environments
    }
  }, []);

  // Sync paperTheme when defaultPaperTheme preference changes in Settings
  useEffect(() => {
    setPaperTheme(defaultPaperTheme);
  }, [defaultPaperTheme]);

  // Initialize or upgrade editor DOM with docx-preview when rawArrayBuffer is available
  useEffect(() => {
    const isNewDoc = loadedDocIdRef.current !== document.id;
    const rawBuffer =
      document.rawArrayBuffer && document.rawArrayBuffer.byteLength > 0
        ? document.rawArrayBuffer
        : null;
    const hasNewRawBuffer =
      rawBuffer !== null && renderedRawBufferRef.current !== rawBuffer;

    if (!isNewDoc && !hasNewRawBuffer) {
      return;
    }

    if (isNewDoc) {
      loadedDocIdRef.current = document.id;
      renderedRawBufferRef.current = null;
      savedRangeRef.current = null;
      setIsDirty(false);
      setEditRevision(0);
      setSaveFeedback(false);
      setExportError(null);
      setOpenPopover(null);
      setIsMobileDrawerOpen(false);
      setIsMobileSheetExpanded(false);
      setIsMobileTabMenuOpen(false);
      setDocxRenderFailed(false);

      if (scrollContainerRef.current) {
        scrollContainerRef.current.scrollTop = 0;
      }
    }

    if (!isNewDoc && isDirty) {
      return;
    }

    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }

    if (rawBuffer) {
      renderedRawBufferRef.current = rawBuffer;
      const currentSeq = renderSeqRef.current + 1;
      renderSeqRef.current = currentSeq;

      setIsRenderingDocx(true);
      setIsDocxPreviewRendered(true);
      setDocxRenderFailed(false);
      editorEl.innerHTML = '';

      void renderAsync(rawBuffer.slice(0), editorEl, undefined, {
        className: 'lite-docx',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        useBase64URL: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
      })
        .then(() => {
          if (renderSeqRef.current !== currentSeq || !editorRef.current) {
            return;
          }
          latestHtmlRef.current = stripInjectedStylesFromHtml(
            editorRef.current,
            ''
          );
          syncOutlineFromDom();
        })
        .catch(() => {
          if (renderSeqRef.current !== currentSeq || !editorRef.current) {
            return;
          }
          setIsDocxPreviewRendered(false);
          setDocxRenderFailed(true);
          const fallbackHtml = buildInitialWordHtml(documentRef.current);
          editorRef.current.innerHTML = fallbackHtml;
          latestHtmlRef.current = fallbackHtml;
          syncOutlineFromDom();
        })
        .finally(() => {
          if (renderSeqRef.current === currentSeq) {
            setIsRenderingDocx(false);
          }
        });
      return;
    }

    renderSeqRef.current += 1;
    setIsRenderingDocx(false);
    setIsDocxPreviewRendered(false);
    const initialHtml = buildInitialWordHtml(document);
    latestHtmlRef.current = initialHtml;
    editorEl.innerHTML = initialHtml;
    syncOutlineFromDom();
  }, [document, isDirty, syncOutlineFromDom]);

  // Track selection changes inside editor for toolbar state & range restore
  useEffect(() => {
    const handleSelectionChange = () => {
      if (isComposingRef.current) {
        return;
      }
      captureCurrentSelection();
      refreshFormatState();
    };
    window.document.addEventListener('selectionchange', handleSelectionChange);
    return () => {
      window.document.removeEventListener(
        'selectionchange',
        handleSelectionChange
      );
    };
  }, [captureCurrentSelection, refreshFormatState]);

  // Close formatting popover on outside pointer down
  useEffect(() => {
    if (!openPopover) {
      return;
    }
    const handleOutsideClick = (event: MouseEvent) => {
      if (
        ribbonRef.current &&
        event.target instanceof Node &&
        !ribbonRef.current.contains(event.target)
      ) {
        setOpenPopover(null);
      }
    };
    window.document.addEventListener('mousedown', handleOutsideClick);
    return () => {
      window.document.removeEventListener('mousedown', handleOutsideClick);
    };
  }, [openPopover]);

  // Measure scrollContainerRef width & inner page height for proportional mobile Print Layout scaling
  useEffect(() => {
    const scrollEl = scrollContainerRef.current;
    const pageEl = pageSurfaceRef.current;
    if (!scrollEl || typeof ResizeObserver === 'undefined') {
      return;
    }

    const updateSizes = () => {
      setContainerWidth(scrollEl.clientWidth);
      if (pageEl) {
        setContentHeight(pageEl.offsetHeight);
      }
    };
    updateSizes();

    const observer = new ResizeObserver(() => {
      updateSizes();
    });
    observer.observe(scrollEl);
    if (pageEl) {
      observer.observe(pageEl);
    }
    return () => {
      observer.disconnect();
    };
  }, []);

  // Cleanup timers on unmount
  useEffect(() => {
    return () => {
      if (saveFeedbackTimerRef.current !== null) {
        window.clearTimeout(saveFeedbackTimerRef.current);
      }
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
      if (typingPresenceTimerRef.current !== null) {
        window.clearTimeout(typingPresenceTimerRef.current);
      }
      if (broadcastThrottleTimerRef.current !== null) {
        window.clearTimeout(broadcastThrottleTimerRef.current);
      }
    };
  }, []);

  const buildLiveDocument = useCallback((): DocumentItem => {
    const currentDoc = documentRef.current;
    const rawHtml = stripInjectedStylesFromHtml(
      editorRef.current,
      latestHtmlRef.current
    );
    const { htmlWithIds, outline, plainText } = extractOutlineFromHtml(rawHtml);
    const trimmedPreview =
      plainText.length > 260
        ? `${plainText.slice(0, 259).trimEnd()}…`
        : plainText;

    const liveOutline =
      outlineEntries.length > 0
        ? outlineEntries.map((entry) => entry.section)
        : outline.length > 0
          ? outline
          : currentDoc.outline;

    return {
      ...currentDoc,
      wordHtml: htmlWithIds,
      outline: liveOutline,
      pageCount: detectedPageCount,
      contentPreview: trimmedPreview || currentDoc.contentPreview,
    };
  }, [detectedPageCount, outlineEntries]);

  const handleSave = useCallback(() => {
    if (isReadOnly) {
      return;
    }
    const updatedDoc = buildLiveDocument();
    onSaveDocument(updatedDoc);
    setIsDirty(false);
    setSaveFeedback(true);

    if (saveFeedbackTimerRef.current !== null) {
      window.clearTimeout(saveFeedbackTimerRef.current);
    }
    saveFeedbackTimerRef.current = window.setTimeout(() => {
      setSaveFeedback(false);
      saveFeedbackTimerRef.current = null;
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
        thumbnail: createThumbnailDataUri(document.type, newTitle, '#2563EB'),
        originalFileName: `${newTitle}.${document.type}`,
        updatedAt: 'Vừa chỉnh sửa',
      };
      onSaveDocument(renamedDoc);
      setIsDirty(false);
      setSaveFeedback(true);

      if (saveFeedbackTimerRef.current !== null) {
        window.clearTimeout(saveFeedbackTimerRef.current);
      }
      saveFeedbackTimerRef.current = window.setTimeout(() => {
        setSaveFeedback(false);
        saveFeedbackTimerRef.current = null;
      }, 2200);
    },
    [buildLiveDocument, document.type, isReadOnly, onSaveDocument]
  );

  // Auto-save with 1800ms debounce (IME-safe; resets on every edit & cancels on manual save/unmount)
  useEffect(() => {
    if (isReadOnly || !autoSaveEnabled || !isDirty || isComposingRef.current) {
      return;
    }
    const timerId = window.setTimeout(() => {
      if (!isComposingRef.current) {
        handleSave();
      }
    }, 1800);
    return () => {
      window.clearTimeout(timerId);
    };
  }, [autoSaveEnabled, isDirty, isReadOnly, editRevision, handleSave]);

  // Keyboard shortcuts: Escape (close popover / mobile sheet / mobile drawer) & Ctrl/Cmd+S (save)
  useEffect(() => {
    if (isModalOpen) {
      return;
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        if (!isReadOnly) {
          handleSave();
        }
        return;
      }

      if (event.key === 'Escape') {
        if (openPopover !== null) {
          setOpenPopover(null);
          return;
        }
        if (isMobileTabMenuOpen) {
          setIsMobileTabMenuOpen(false);
          return;
        }
        if (isMobileSheetExpanded) {
          setIsMobileSheetExpanded(false);
          return;
        }
        if (isMobileDrawerOpen) {
          setIsMobileDrawerOpen(false);
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    isModalOpen,
    isMobileDrawerOpen,
    isMobileSheetExpanded,
    isMobileTabMenuOpen,
    isReadOnly,
    openPopover,
    handleSave,
  ]);

  const flushBroadcastHtml = useCallback((htmlToSend: string) => {
    const plain = (editorRef.current?.textContent ?? '')
      .replace(/\s+/g, ' ')
      .trim();
    const preview =
      plain.length > 260 ? `${plain.slice(0, 259).trimEnd()}…` : plain;
    lastBroadcastAtRef.current = Date.now();
    onBroadcastPatchRef.current?.({
      kind: 'word_html_patch',
      wordHtml: htmlToSend,
      contentPreview: preview || undefined,
    });
  }, []);

  const handleEditorMutation = useCallback(() => {
    if (!editorRef.current || isApplyingRemotePatchRef.current || isReadOnly) {
      return;
    }
    const nextCleanHtml = stripInjectedStylesFromHtml(
      editorRef.current,
      latestHtmlRef.current
    );
    latestHtmlRef.current = nextCleanHtml;
    setIsDirty(true);
    setEditRevision((prev) => prev + 1);
    setSaveFeedback(false);
    syncOutlineFromDom();
    refreshFormatState();

    // Update realtime typing presence & reset after 1800ms
    onPresenceUpdateRef.current?.({
      isTyping: true,
      cursorLabel: 'Đang chỉnh sửa văn bản',
    });
    if (typingPresenceTimerRef.current !== null) {
      window.clearTimeout(typingPresenceTimerRef.current);
    }
    typingPresenceTimerRef.current = window.setTimeout(() => {
      typingPresenceTimerRef.current = null;
      onPresenceUpdateRef.current?.({
        isTyping: false,
        cursorLabel: 'Đang xem tài liệu',
      });
    }, 1800);

    // Throttled realtime broadcast (~80ms)
    if (onBroadcastPatchRef.current) {
      const now = Date.now();
      const elapsed = now - lastBroadcastAtRef.current;
      if (elapsed >= 80 && broadcastThrottleTimerRef.current === null) {
        flushBroadcastHtml(nextCleanHtml);
      } else {
        pendingBroadcastHtmlRef.current = nextCleanHtml;
        if (broadcastThrottleTimerRef.current === null) {
          broadcastThrottleTimerRef.current = window.setTimeout(() => {
            broadcastThrottleTimerRef.current = null;
            const pending = pendingBroadcastHtmlRef.current;
            pendingBroadcastHtmlRef.current = null;
            if (pending !== null) {
              flushBroadcastHtml(pending);
            }
          }, Math.max(16, 80 - elapsed));
        }
      }
    }
  }, [flushBroadcastHtml, isReadOnly, syncOutlineFromDom, refreshFormatState]);

  // Apply incoming realtime Word HTML patch without broadcast loop or scroll jump
  useEffect(() => {
    if (
      !incomingRemotePatch ||
      incomingRemotePatch.kind !== 'word_html_patch'
    ) {
      return;
    }
    if (incomingRemotePatch.sentAt <= lastRemotePatchAtRef.current) {
      return;
    }
    lastRemotePatchAtRef.current = incomingRemotePatch.sentAt;

    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }
    const incomingHtml = incomingRemotePatch.wordHtml ?? '';
    if (incomingHtml === latestHtmlRef.current) {
      return;
    }

    const prevScrollTop = scrollContainerRef.current?.scrollTop ?? 0;
    isApplyingRemotePatchRef.current = true;
    try {
      let preservedStylePrefix = '';
      if (!incomingHtml.includes('<style')) {
        const existingStyles = Array.from(editorEl.querySelectorAll('style'))
          .map((node) => node.outerHTML)
          .join('');
        preservedStylePrefix = existingStyles;
      }
      editorEl.innerHTML = `${preservedStylePrefix}${incomingHtml}`;
      latestHtmlRef.current = incomingHtml;
      syncOutlineFromDom();
    } finally {
      isApplyingRemotePatchRef.current = false;
    }

    window.requestAnimationFrame(() => {
      if (scrollContainerRef.current) {
        scrollContainerRef.current.scrollTop = prevScrollTop;
      }
    });
  }, [incomingRemotePatch, syncOutlineFromDom]);

  const restoreEditorSelection = useCallback(() => {
    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }
    editorEl.focus({ preventScroll: true });

    const selection = window.document.getSelection();
    if (!selection) {
      return;
    }

    const saved = savedRangeRef.current;
    if (saved && editorEl.contains(saved.commonAncestorContainer)) {
      selection.removeAllRanges();
      selection.addRange(saved);
      return;
    }

    if (
      selection.rangeCount === 0 ||
      !editorEl.contains(selection.anchorNode)
    ) {
      const targetContainer =
        editorEl.querySelector('section.lite-docx') ?? editorEl;
      const endRange = window.document.createRange();
      endRange.selectNodeContents(targetContainer);
      endRange.collapse(false);
      selection.removeAllRanges();
      selection.addRange(endRange);
      savedRangeRef.current = endRange.cloneRange();
    }
  }, []);

  const executeFormatCommand = useCallback(
    (command: string, value?: string) => {
      if (isReadOnly) {
        return;
      }
      restoreEditorSelection();
      try {
        if (command === 'hiliteColor') {
          const applied = window.document.execCommand(
            'hiliteColor',
            false,
            value ?? 'transparent'
          );
          if (!applied) {
            window.document.execCommand(
              'backColor',
              false,
              value ?? 'transparent'
            );
          }
        } else {
          window.document.execCommand(command, false, value);
        }
      } catch {
        // Ignore unsupported execCommand errors gracefully
      }
      captureCurrentSelection();
      handleEditorMutation();
    },
    [isReadOnly, restoreEditorSelection, captureCurrentSelection, handleEditorMutation]
  );

  const handleApplyBlockTag = (tag: BlockTag) => {
    executeFormatCommand('formatBlock', `<${tag}>`);
    setOpenPopover(null);
  };

  const handleSelectFontFamily = (nextFamily: WordFontFamily) => {
    setFontFamily(nextFamily);
    const selection = window.document.getSelection();
    if (selection && !selection.isCollapsed) {
      executeFormatCommand('fontName', nextFamily);
    }
  };

  const handleInsertTable = (rows = 3, cols = 3) => {
    executeFormatCommand('insertHTML', buildQuickTableHtml(rows, cols));
    setOpenPopover(null);
  };

  const handleInsertHorizontalRule = () => {
    executeFormatCommand(
      'insertHTML',
      '<hr style="border:0;border-top:1.5px solid #cbd5e1;margin:16px 0;" /><p><br></p>'
    );
  };

  const handleInsertPageBreak = () => {
    executeFormatCommand(
      'insertHTML',
      '<div data-page-break="true" style="page-break-after:always;break-after:page;margin:24px 0;padding:6px 0;border-top:1px dashed #94a3b8;border-bottom:1px dashed #94a3b8;text-align:center;font-size:11px;color:#64748b;user-select:none;">— Ngắt trang —</div><p><br></p>'
    );
  };

  const handleInsertCalloutBox = () => {
    executeFormatCommand(
      'insertHTML',
      '<blockquote><strong>Ghi chú / Trích dẫn:</strong> Nhập nội dung điểm nhấn hoặc trích dẫn quan trọng tại đây...</blockquote><p><br></p>'
    );
  };

  const handleStepZoom = (direction: -1 | 1) => {
    setZoomPercent((prev) => {
      const idx = ZOOM_PERCENTS.indexOf(prev);
      const nextIdx = Math.max(
        0,
        Math.min(ZOOM_PERCENTS.length - 1, idx + direction)
      );
      return ZOOM_PERCENTS[nextIdx];
    });
  };

  const handlePreventMouseDown = (event: ReactMouseEvent<HTMLButtonElement>) => {
    event.preventDefault();
  };

  const handleExportDocx = async () => {
    if (isExporting) {
      return;
    }
    setIsExporting(true);
    setExportError(null);
    try {
      const currentHtml = stripInjectedStylesFromHtml(
        editorRef.current,
        latestHtmlRef.current
      );
      await exportWordToDocx(document.title, currentHtml);
    } catch {
      setExportError('Không thể xuất tệp .docx. Vui lòng thử lại.');
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

  const handleDecreaseFont = () => {
    setFontSizeIndex((prev) => (prev > 0 ? prev - 1 : prev));
  };

  const handleIncreaseFont = () => {
    setFontSizeIndex((prev) =>
      prev < FONT_SIZES.length - 1 ? prev + 1 : prev
    );
  };

  const handleJumpToSection = (sectionId: string) => {
    setActiveSectionId(sectionId);
    setIsMobileDrawerOpen(false);
    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }
    const target = Array.from(
      editorEl.querySelectorAll<HTMLElement>('[id]')
    ).find((node) => node.id === sectionId);
    target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const handleReadingScroll = () => {
    const container = scrollContainerRef.current;
    const editorEl = editorRef.current;
    if (!container || !editorEl || outlineEntries.length === 0) {
      return;
    }
    const containerRect = container.getBoundingClientRect();
    const headingMap = new Map<string, HTMLElement>();
    for (const node of Array.from(
      editorEl.querySelectorAll<HTMLElement>('[id]')
    )) {
      if (node.id) {
        headingMap.set(node.id, node);
      }
    }

    let currentId = outlineEntries[0].section.id;
    for (const entry of outlineEntries) {
      const node = headingMap.get(entry.section.id);
      if (!node) {
        continue;
      }
      const relativeTop = node.getBoundingClientRect().top - containerRect.top;
      if (relativeTop <= 140) {
        currentId = entry.section.id;
      } else {
        break;
      }
    }
    setActiveSectionId((prev) => (prev === currentId ? prev : currentId));
  };

  const quickCommands = useMemo(
    () => [
      {
        label: 'Tiêu đề 1 (H1)',
        action: () => handleApplyBlockTag('h1'),
      },
      {
        label: 'Tiêu đề 2 (H2)',
        action: () => handleApplyBlockTag('h2'),
      },
      {
        label: 'Văn bản thường (P)',
        action: () => handleApplyBlockTag('p'),
      },
      {
        label: 'Tô sáng vàng',
        action: () => {
          setCurrentHighlightColor('#FEF08A');
          executeFormatCommand('hiliteColor', '#FEF08A');
        },
      },
      {
        label: 'Chèn bảng 3×3',
        action: () => handleInsertTable(3, 3),
      },
      {
        label: 'Dấu ngắt trang',
        action: () => handleInsertPageBreak(),
      },
      {
        label: 'Khung ghi chú / trích dẫn',
        action: () => handleInsertCalloutBox(),
      },
      {
        label: 'Xóa định dạng',
        action: () => executeFormatCommand('removeFormat'),
      },
      {
        label: 'Mở Mục lục tự động',
        action: () => {
          setIsMobileSheetExpanded(false);
          setIsMobileDrawerOpen(true);
        },
      },
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [executeFormatCommand]
  );

  const filteredQuickCommands = useMemo(() => {
    const q = mobileCommandQuery.trim().toLowerCase();
    if (!q) {
      return quickCommands;
    }
    return quickCommands.filter((item) => item.label.toLowerCase().includes(q));
  }, [mobileCommandQuery, quickCommands]);

  const currentBlockLabel =
    BLOCK_STYLE_OPTIONS.find((item) => item.tag === formatState.block)?.label ??
    'Văn bản thường';

  const toolbarBtnClass = (active: boolean) =>
    `inline-flex items-center justify-center min-h-[30px] min-w-[30px] px-2 rounded-lg text-xs font-medium transition-colors shrink-0 ${
      active
        ? isDark
          ? 'bg-indigo-500/30 text-indigo-200 ring-1 ring-indigo-400/40'
          : 'bg-indigo-50 text-indigo-700 ring-1 ring-indigo-200'
        : isDark
          ? 'text-slate-300 hover:bg-white/[0.08] hover:text-white'
          : 'text-slate-700 hover:bg-black/[0.05] hover:text-slate-950'
    }`;

  const renderOutlineContent = (variant: 'desktop' | 'mobile') => (
    <div className="flex-1 min-h-0 flex flex-col justify-between overflow-hidden">
      <nav
        aria-label={
          variant === 'mobile'
            ? 'Mục lục tự động di động'
            : 'Mục lục tự động'
        }
        className="flex-1 min-h-0 overflow-y-auto p-3 space-y-1"
      >
        {outlineEntries.length > 0 ? (
          outlineEntries.map((entry) => {
            const { section, level } = entry;
            const isActive = section.id === activeSectionId;
            const indentClass =
              level === 1 ? 'pl-3' : level === 2 ? 'pl-4' : 'pl-6';

            return (
              <button
                key={`${variant}-${section.id}`}
                type="button"
                onClick={() => handleJumpToSection(section.id)}
                aria-current={isActive ? 'location' : undefined}
                className={`w-full text-left rounded-xl pr-2.5 py-2 transition-all flex items-start justify-between gap-2 ${indentClass} ${
                  isActive ? themeStyle.outlineActive : themeStyle.outlineIdle
                }`}
              >
                <span
                  className={`min-w-0 leading-snug line-clamp-2 ${
                    level === 1
                      ? 'text-xs font-semibold'
                      : level === 2
                        ? 'text-xs'
                        : 'text-[11.5px] opacity-90'
                  }`}
                >
                  {section.title}
                </span>
                <span
                  className={`text-[10px] font-mono shrink-0 pt-0.5 tabular-nums ${themeStyle.textMuted}`}
                >
                  Tr.{section.pageNumber}
                </span>
              </button>
            );
          })
        ) : (
          <div
            className={`rounded-xl border border-dashed p-3.5 text-xs leading-relaxed ${themeStyle.calloutBorder} ${themeStyle.textMuted}`}
          >
            Chưa có tiêu đề chương mục. Chọn <strong>Tiêu đề 1 / 2 / 3</strong>{' '}
            hoặc gõ tiêu đề chương (<code>CHƯƠNG 1</code>, <code>I.</code>,{' '}
            <code>1.1.</code>) để tạo mục lục tự động.
          </div>
        )}
      </nav>

      {/* Bottom Outline Summary — pb-safe separated from inner p-3 */}
      <div
        className={`shrink-0 border-t ${themeStyle.sidebarBorder} ${
          variant === 'mobile' ? 'pb-safe' : ''
        }`}
      >
        <div className="p-3 space-y-2.5">
          <div
            className={`rounded-xl border p-3 space-y-1.5 ${themeStyle.calloutBg} ${themeStyle.calloutBorder}`}
          >
            <div className="flex items-center justify-between gap-2 text-[11.5px]">
              <span className="font-semibold">Thống kê trực tiếp</span>
              <span
                className={`px-2 py-0.5 rounded-full text-[10.5px] font-bold tabular-nums ${
                  isDirty
                    ? isDark
                      ? 'bg-amber-500/20 text-amber-300'
                      : 'bg-amber-100 text-amber-800'
                    : isDark
                      ? 'bg-emerald-500/20 text-emerald-300'
                      : 'bg-emerald-100 text-emerald-800'
                }`}
              >
                {isDirty ? 'Chưa lưu' : 'Đã đồng bộ'}
              </span>
            </div>
            <p className={`text-[11px] tabular-nums ${themeStyle.textMuted}`}>
              {outlineEntries.length} chương mục ·{' '}
              {wordCount.toLocaleString('vi-VN')} từ · {detectedPageCount} trang
            </p>
          </div>

          <div className={`px-1 space-y-1 text-[11px] ${themeStyle.textMuted}`}>
            <div className="flex items-center gap-1.5 truncate">
              <FolderOpen className="h-3.5 w-3.5 shrink-0 opacity-80" />
              <span className="truncate">
                {document.folder ?? 'Tài liệu Word'}
              </span>
            </div>
            <p className="tabular-nums">
              {isDocxPreviewRendered
                ? 'Định dạng .docx gốc'
                : document.isUploaded
                  ? 'Tệp của bạn'
                  : 'Tài liệu Word'}
              {document.fileSize ? ` · ${document.fileSize}` : ''}
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  const hasCloudSource = Boolean(
    document.storagePath || document.r2ObjectKey || document.isSyncedToCloud
  );
  const canRetryCloudBuffer =
    !isHydratingCloudFile &&
    !document.rawArrayBuffer &&
    Boolean(onRetryCloudDownload) &&
    (hasCloudSource || docxRenderFailed);

  const activeTypingCollaborators = useMemo(
    () =>
      (collaborators ?? []).filter(
        (c) => c.sessionId !== selfSessionId && Boolean(c.isTyping)
      ),
    [collaborators, selfSessionId]
  );

  const fontStack = getFontStack(fontFamily);
  const scaledSheetMaxWidth =
    pageLayout === 'mobile'
      ? undefined
      : Math.round(((isDocxPreviewRendered ? 860 : 794) * zoomPercent) / 100);

  const hostCssVars: CSSProperties = {
    '--lite-doc-page-bg': themeStyle.pageHexBg,
    '--lite-doc-page-text': themeStyle.pageHexText,
    '--lite-doc-font-family': fontStack,
    '--lite-doc-line-height': String(lineHeight),
    fontFamily: fontStack,
    fontSize: `${effectiveFontSize}px`,
    lineHeight,
    ...(scaledSheetMaxWidth !== undefined
      ? { maxWidth: `${scaledSheetMaxWidth}px` }
      : {}),
  } as CSSProperties;

  return (
    <div
      className={`h-[100dvh] w-full flex flex-col overflow-hidden transition-colors duration-200 ${themeStyle.outerBg} ${themeStyle.textPrimary} ${themeStyle.scrollbarClass}`}
    >
      <style>{DOCX_PREVIEW_SCOPED_CSS}</style>

      {/* Floating Context Selection Mini-Toolbar */}
      <FloatingSelectionToolbar
        containerRef={editorRef}
        disabled={isReadOnly || isRenderingDocx || isModalOpen}
        activeState={{
          block: formatState.block,
          bold: formatState.bold,
          italic: formatState.italic,
          underline: formatState.underline,
          textColor: customTextColor,
          highlightColor: currentHighlightColor,
        }}
        onSelectBlock={(tag) => handleApplyBlockTag(tag)}
        onToggleBold={() => executeFormatCommand('bold')}
        onToggleItalic={() => executeFormatCommand('italic')}
        onToggleUnderline={() => executeFormatCommand('underline')}
        onSelectTextColor={(color) => {
          setCustomTextColor(color);
          executeFormatCommand('foreColor', color);
        }}
        onSelectHighlight={(color) => {
          setCurrentHighlightColor(color);
          executeFormatCommand('hiliteColor', color);
        }}
        onClearFormat={() => executeFormatCommand('removeFormat')}
        textColorSwatches={TEXT_COLOR_SWATCHES}
        highlightSwatches={HIGHLIGHT_SWATCHES}
      />

      {/* Shared Top Workspace Header */}
      <WorkspaceTopBar
        document={document}
        theme={isDark ? 'dark' : 'light'}
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument())}
        onRenameTitle={isReadOnly ? undefined : handleRenameTitle}
        collaborators={collaborators}
        selfSessionId={selfSessionId}
        collabPermission={collabPermission}
        leftExtra={
          <>
            {/* Mobile Outline Drawer Trigger */}
            <button
              type="button"
              onClick={() => setIsMobileDrawerOpen(true)}
              aria-label="Mở mục lục tự động"
              title="Mục lục tự động"
              className={`md:hidden inline-flex items-center justify-center min-h-[38px] min-w-[38px] rounded-xl border transition-colors ${
                isDark
                  ? 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                  : 'bg-[#FAF8F5] hover:bg-[#EFECE6] border-[#E7E2DA] text-slate-700'
              }`}
            >
              <List className="h-4 w-4" />
            </button>

            {/* Desktop Outline Sidebar Collapse/Expand Trigger */}
            <button
              type="button"
              onClick={() => setIsDesktopSidebarOpen((prev) => !prev)}
              aria-label={
                isDesktopSidebarOpen
                  ? 'Thu gọn cột Mục lục'
                  : 'Mở rộng cột Mục lục'
              }
              title={
                isDesktopSidebarOpen ? 'Thu gọn Mục lục' : 'Mở rộng Mục lục'
              }
              className={`hidden md:inline-flex items-center justify-center h-9 w-9 rounded-xl border transition-colors ${
                isDark
                  ? 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                  : 'bg-[#FAF8F5] hover:bg-[#EFECE6] border-[#E7E2DA] text-slate-700'
              }`}
            >
              {isDesktopSidebarOpen ? (
                <PanelLeftClose className="h-4 w-4" />
              ) : (
                <PanelLeftOpen className="h-4 w-4" />
              )}
            </button>
          </>
        }
        toolbar={
          <>
            {/* Quick Page Layout Toggle: Bố cục In <-> Xem di động */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={handleTogglePageLayout}
              aria-pressed={pageLayout === 'mobile'}
              aria-label={
                pageLayout === 'print'
                  ? 'Đang ở Bố cục In — bấm để chuyển sang Xem di động'
                  : 'Đang ở Xem di động — bấm để chuyển sang Bố cục In'
              }
              title={
                pageLayout === 'print'
                  ? 'Bố cục In (bấm để chuyển Xem di động)'
                  : 'Xem di động (bấm để chuyển Bố cục In)'
              }
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2 sm:px-2.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
                pageLayout === 'mobile'
                  ? isDark
                    ? 'bg-blue-500/20 border-blue-400/45 text-blue-200'
                    : 'bg-blue-50 border-blue-300 text-blue-700'
                  : isDark
                    ? 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                    : 'bg-[#FAF8F5] hover:bg-[#EFECE6] border-[#E5E0D8] text-slate-700'
              }`}
            >
              {pageLayout === 'print' ? (
                <ScanText className="h-3.5 w-3.5 shrink-0" />
              ) : (
                <Smartphone className="h-3.5 w-3.5 shrink-0" />
              )}
              <span className="text-[11px] sm:text-xs">
                {pageLayout === 'print' ? 'Bố cục In' : 'Xem di động'}
              </span>
            </button>

            {/* Save Changes Button (disabled in read-only share mode) */}
            {!isReadOnly && (
              <button
                type="button"
                onClick={handleSave}
                aria-label="Lưu thay đổi tài liệu Word"
                title="Lưu thay đổi (Ctrl+S)"
                className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 sm:px-3 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
                  saveFeedback
                    ? isDark
                      ? 'bg-emerald-500/20 border-emerald-400/50 text-emerald-200'
                      : 'bg-emerald-50 border-emerald-300 text-emerald-700'
                    : isDirty
                      ? 'bg-indigo-600 hover:bg-indigo-700 border-indigo-600 text-white shadow-sm'
                      : isDark
                        ? 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
                        : 'bg-[#FAF8F5] hover:bg-[#EFECE6] border-[#E5E0D8] text-slate-700'
                }`}
              >
                {saveFeedback ? (
                  <>
                    <Check className="h-3.5 w-3.5 shrink-0" />
                    <span className="hidden sm:inline">Đã lưu</span>
                  </>
                ) : (
                  <>
                    <span className="relative flex items-center justify-center">
                      <Save className="h-3.5 w-3.5 shrink-0" />
                      {isDirty && (
                        <span className="absolute -top-1 -right-1 h-2 w-2 rounded-full bg-amber-400 ring-1 ring-indigo-600" />
                      )}
                    </span>
                    <span className="hidden sm:inline">Lưu thay đổi</span>
                  </>
                )}
              </button>
            )}

            {/* Export Real .docx Button */}
            <button
              type="button"
              onClick={() => void handleExportDocx()}
              disabled={isExporting}
              aria-label="Tải xuống tệp Word .docx"
              title="Tải xuống tệp .docx"
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 sm:px-3 rounded-xl text-xs font-semibold border transition-colors whitespace-nowrap disabled:opacity-60 disabled:cursor-wait ${
                isDark
                  ? 'bg-blue-500/20 hover:bg-blue-500/30 border-blue-400/35 text-blue-200'
                  : 'bg-blue-50 hover:bg-blue-100 border-blue-200/80 text-blue-700'
              }`}
            >
              {isExporting ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
              ) : (
                <Download className="h-3.5 w-3.5 shrink-0" />
              )}
              <span className="hidden md:inline">
                {isExporting ? 'Đang xuất...' : 'Tải xuống .docx'}
              </span>
            </button>
          </>
        }
      />

      {/* Desktop 1-Row Ribbon Bar (hidden on mobile < md) */}
      <div
        ref={ribbonRef}
        className={`hidden md:block shrink-0 border-b px-3 sm:px-4 py-1.5 z-10 transition-colors duration-200 relative ${themeStyle.ribbonBg} ${themeStyle.ribbonBorder}`}
      >
        <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-0.5 min-w-0">
          {/* 1. Primary Section Switcher: Trang chủ | Chèn | Bố trí | Xem */}
          <div
            role="tablist"
            aria-label="Chọn phần công cụ Word"
            className={`inline-flex items-center gap-0.5 p-0.5 rounded-xl border shrink-0 ${
              isDark
                ? 'bg-[#1E2534] border-white/10'
                : 'bg-[#F5F3EE] border-[#E5E0D8]'
            }`}
          >
            {RIBBON_TABS.map((tab) => {
              const isSelected = activeRibbonTab === tab.key;
              return (
                <button
                  key={tab.key}
                  type="button"
                  role="tab"
                  aria-selected={isSelected}
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setActiveRibbonTab(tab.key)}
                  className={`min-h-[28px] px-2.5 rounded-lg text-xs font-semibold transition-all whitespace-nowrap ${
                    isSelected
                      ? isDark
                        ? 'bg-indigo-600 text-white shadow-sm'
                        : 'bg-white text-indigo-700 shadow-sm ring-1 ring-black/5'
                      : isDark
                        ? 'text-slate-400 hover:text-slate-200'
                        : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>

          <span
            aria-hidden="true"
            className={`h-5 w-px mx-0.5 shrink-0 ${
              isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
            }`}
          />

          {isReadOnly && (
            <span
              className={`shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-semibold border ${
                isDark
                  ? 'bg-amber-500/15 border-amber-400/30 text-amber-200'
                  : 'bg-amber-50 border-amber-200 text-amber-800'
              }`}
            >
              Chế độ chỉ xem
            </span>
          )}

          {/* 2. Core Quick Formatting Cluster (disabled when read-only) */}
          <fieldset
            disabled={isReadOnly}
            className={`flex items-center gap-1 min-w-0 border-0 p-0 m-0 ${
              isReadOnly ? 'opacity-45 pointer-events-none' : ''
            }`}
          >
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() =>
              setOpenPopover((prev) => (prev === 'block' ? null : 'block'))
            }
            aria-expanded={openPopover === 'block'}
            aria-label="Chọn kiểu đoạn hoặc tiêu đề"
            className={`shrink-0 inline-flex items-center justify-between gap-1.5 min-h-[30px] px-2.5 rounded-lg border text-xs font-medium transition-colors whitespace-nowrap ${
              isDark
                ? 'bg-[#1E2534] border-white/10 text-slate-200 hover:bg-white/10'
                : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800 hover:bg-[#F2EFEA]'
            }`}
          >
            <span>{currentBlockLabel}</span>
            <ChevronDown className="h-3.5 w-3.5 opacity-70" />
          </button>

          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('bold')}
            aria-pressed={formatState.bold}
            aria-label="In đậm"
            title="In đậm (Ctrl+B)"
            className={toolbarBtnClass(formatState.bold)}
          >
            <Bold className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('italic')}
            aria-pressed={formatState.italic}
            aria-label="In nghiêng"
            title="In nghiêng (Ctrl+I)"
            className={toolbarBtnClass(formatState.italic)}
          >
            <Italic className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('underline')}
            aria-pressed={formatState.underline}
            aria-label="Gạch chân"
            title="Gạch chân (Ctrl+U)"
            className={toolbarBtnClass(formatState.underline)}
          >
            <Underline className="h-3.5 w-3.5" />
          </button>

          {/* Text Color Popover Trigger */}
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() =>
              setOpenPopover((prev) => (prev === 'color' ? null : 'color'))
            }
            aria-expanded={openPopover === 'color'}
            aria-label="Chọn màu chữ"
            title="Màu chữ"
            className={toolbarBtnClass(openPopover === 'color')}
          >
            <Palette className="h-3.5 w-3.5" />
            <span
              aria-hidden="true"
              className="ml-1 h-2.5 w-2.5 rounded-full border border-black/20"
              style={{ backgroundColor: customTextColor }}
            />
          </button>

          {/* Highlight Color Popover Trigger */}
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() =>
              setOpenPopover((prev) =>
                prev === 'highlight' ? null : 'highlight'
              )
            }
            aria-expanded={openPopover === 'highlight'}
            aria-label="Chọn màu nền highlight"
            title="Tô sáng"
            className={toolbarBtnClass(openPopover === 'highlight')}
          >
            <Highlighter className="h-3.5 w-3.5 text-amber-500" />
          </button>

          {/* Alignment (4 buttons) */}
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('justifyLeft')}
            aria-pressed={formatState.align === 'left'}
            aria-label="Căn trái"
            title="Căn trái"
            className={toolbarBtnClass(formatState.align === 'left')}
          >
            <AlignLeft className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('justifyCenter')}
            aria-pressed={formatState.align === 'center'}
            aria-label="Căn giữa"
            title="Căn giữa"
            className={toolbarBtnClass(formatState.align === 'center')}
          >
            <AlignCenter className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('justifyRight')}
            aria-pressed={formatState.align === 'right'}
            aria-label="Căn phải"
            title="Căn phải"
            className={toolbarBtnClass(formatState.align === 'right')}
          >
            <AlignRight className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('justifyFull')}
            aria-pressed={formatState.align === 'justify'}
            aria-label="Căn đều hai bên"
            title="Căn đều"
            className={toolbarBtnClass(formatState.align === 'justify')}
          >
            <AlignJustify className="h-3.5 w-3.5" />
          </button>

          {/* Lists (Bullet & Numbered) */}
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('insertUnorderedList')}
            aria-pressed={formatState.unorderedList}
            aria-label="Danh sách dấu chấm"
            title="Danh sách chấm"
            className={toolbarBtnClass(formatState.unorderedList)}
          >
            <List className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => executeFormatCommand('insertOrderedList')}
            aria-pressed={formatState.orderedList}
            aria-label="Danh sách đánh số"
            title="Danh sách số"
            className={toolbarBtnClass(formatState.orderedList)}
          >
            <ListOrdered className="h-3.5 w-3.5" />
          </button>
          </fieldset>

          <span
            aria-hidden="true"
            className={`h-5 w-px mx-0.5 shrink-0 ${
              isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
            }`}
          />

          {/* 3. Contextual Tools by Active Ribbon Tab */}
          {activeRibbonTab === 'home' && (
            <>
              {/* Font Family: Times New Roman | Arial */}
              <div
                role="group"
                aria-label="Chọn phông chữ"
                className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ${
                  isDark
                    ? 'bg-[#1E2534] border-white/10'
                    : 'bg-[#F5F3EE] border-[#E5E0D8]'
                }`}
              >
                {FONT_FAMILIES.map((family) => {
                  const isSelected = fontFamily === family;
                  return (
                    <button
                      key={family}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleSelectFontFamily(family)}
                      aria-pressed={isSelected}
                      className={`min-h-[26px] px-2 rounded-md text-[11px] font-medium transition-colors whitespace-nowrap ${
                        isSelected
                          ? isDark
                            ? 'bg-indigo-500/30 text-white'
                            : 'bg-white text-slate-900 shadow-sm ring-1 ring-black/5'
                          : isDark
                            ? 'text-slate-400 hover:text-slate-200'
                            : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {family}
                    </button>
                  );
                })}
              </div>

              {/* Font Size Controls (12..24) */}
              <div
                role="group"
                aria-label="Điều chỉnh cỡ chữ"
                className={`inline-flex items-center rounded-lg border p-0.5 shrink-0 ${
                  isDark
                    ? 'bg-[#1E2534] border-white/10'
                    : 'bg-[#F5F3EE] border-[#E5E0D8]'
                }`}
              >
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleDecreaseFont}
                  disabled={fontSizeIndex === 0}
                  aria-label="Giảm cỡ chữ"
                  title="Giảm cỡ chữ"
                  className={`min-h-[26px] px-2 rounded-md text-[11.5px] font-semibold transition-colors disabled:opacity-35 disabled:cursor-not-allowed ${
                    isDark
                      ? 'text-slate-200 hover:bg-white/10'
                      : 'text-slate-700 hover:bg-white'
                  }`}
                >
                  A-
                </button>
                <span
                  className={`px-1.5 text-[11px] font-semibold tabular-nums select-none ${
                    isDark ? 'text-slate-300' : 'text-slate-600'
                  }`}
                >
                  {fontSize}px
                </span>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleIncreaseFont}
                  disabled={fontSizeIndex === FONT_SIZES.length - 1}
                  aria-label="Tăng cỡ chữ"
                  title="Tăng cỡ chữ"
                  className={`min-h-[26px] px-2 rounded-md text-[11.5px] font-semibold transition-colors disabled:opacity-35 disabled:cursor-not-allowed ${
                    isDark
                      ? 'text-slate-200 hover:bg-white/10'
                      : 'text-slate-700 hover:bg-white'
                  }`}
                >
                  A+
                </button>
              </div>

              {/* Strikethrough (ab) */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('strikeThrough')}
                aria-pressed={formatState.strikeThrough}
                aria-label="Gạch ngang chữ"
                title="Gạch ngang (ab)"
                className={toolbarBtnClass(formatState.strikeThrough)}
              >
                <Strikethrough className="h-3.5 w-3.5" />
              </button>

              {/* Clear Formatting */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('removeFormat')}
                aria-label="Xóa định dạng"
                title="Xóa định dạng"
                className={toolbarBtnClass(false)}
              >
                <RemoveFormatting className="h-3.5 w-3.5" />
              </button>
            </>
          )}

          {activeRibbonTab === 'insert' && (
            <>
              {(
                [
                  { rows: 2, cols: 2, label: 'Bảng 2×2' },
                  { rows: 3, cols: 3, label: 'Bảng 3×3' },
                  { rows: 4, cols: 4, label: 'Bảng 4×4' },
                ] as const
              ).map((tbl) => (
                <button
                  key={tbl.label}
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleInsertTable(tbl.rows, tbl.cols)}
                  className={`shrink-0 inline-flex items-center gap-1 min-h-[30px] px-2.5 rounded-lg text-xs font-medium transition-colors whitespace-nowrap ${
                    isDark
                      ? 'bg-[#1E2534] text-slate-200 hover:bg-white/10 border border-white/10'
                      : 'bg-[#FAF8F5] text-slate-700 hover:bg-[#EFECE6] border border-[#E5E0D8]'
                  }`}
                >
                  <Table className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
                  <span>{tbl.label}</span>
                </button>
              ))}

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleInsertHorizontalRule}
                className={`shrink-0 inline-flex items-center gap-1 min-h-[30px] px-2.5 rounded-lg text-xs font-medium transition-colors whitespace-nowrap ${
                  isDark
                    ? 'text-slate-200 hover:bg-white/[0.08]'
                    : 'text-slate-700 hover:bg-black/[0.05]'
                }`}
              >
                <Minus className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
                <span>Đường kẻ ngang</span>
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleInsertPageBreak}
                className={`shrink-0 inline-flex items-center gap-1 min-h-[30px] px-2.5 rounded-lg text-xs font-medium transition-colors whitespace-nowrap ${
                  isDark
                    ? 'text-slate-200 hover:bg-white/[0.08]'
                    : 'text-slate-700 hover:bg-black/[0.05]'
                }`}
              >
                <SplitSquareVertical className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
                <span>Ngắt trang</span>
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={handleInsertCalloutBox}
                className={`shrink-0 inline-flex items-center gap-1 min-h-[30px] px-2.5 rounded-lg text-xs font-medium transition-colors whitespace-nowrap ${
                  isDark
                    ? 'text-slate-200 hover:bg-white/[0.08]'
                    : 'text-slate-700 hover:bg-black/[0.05]'
                }`}
              >
                <Quote className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
                <span>Khung ghi chú</span>
              </button>
            </>
          )}

          {activeRibbonTab === 'layout' && (
            <>
              {/* Paper Theme Switcher */}
              <div
                role="group"
                aria-label="Chọn màu giấy trang A4"
                className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ${
                  isDark
                    ? 'bg-[#1E2534] border-white/10'
                    : 'bg-[#F5F3EE] border-[#E5E0D8]'
                }`}
              >
                {(['light', 'sepia', 'dark'] as const).map((mode) => {
                  const config = PAPER_THEMES[mode];
                  const isSelected = paperTheme === mode;
                  return (
                    <button
                      key={mode}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setPaperTheme(mode)}
                      aria-pressed={isSelected}
                      title={config.label}
                      className={`inline-flex items-center gap-1 min-h-[26px] px-2.5 rounded-md text-[11px] font-medium transition-all whitespace-nowrap ${
                        isSelected
                          ? isDark
                            ? 'bg-indigo-500/30 text-white shadow-sm'
                            : 'bg-white text-slate-900 shadow-sm ring-1 ring-black/5'
                          : isDark
                            ? 'text-slate-400 hover:text-slate-200'
                            : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {mode === 'light' ? (
                        <Sun className="h-3.5 w-3.5 text-amber-500 shrink-0" />
                      ) : mode === 'dark' ? (
                        <Moon className="h-3.5 w-3.5 text-indigo-400 shrink-0" />
                      ) : (
                        <span
                          aria-hidden="true"
                          className={`h-3 w-3 rounded-full border shrink-0 ${config.swatchClass}`}
                        />
                      )}
                      <span>{config.label}</span>
                    </button>
                  );
                })}
              </div>

              {/* Line Height Switcher: 1.15 | 1.5 | 2.0 */}
              <div
                role="group"
                aria-label="Chọn độ giãn dòng"
                className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ${
                  isDark
                    ? 'bg-[#1E2534] border-white/10'
                    : 'bg-[#F5F3EE] border-[#E5E0D8]'
                }`}
              >
                <span
                  className={`px-2 text-[11px] font-medium ${themeStyle.textMuted}`}
                >
                  Giãn dòng:
                </span>
                {LINE_HEIGHTS.map((lh) => {
                  const isSelected = lineHeight === lh;
                  return (
                    <button
                      key={lh}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => setLineHeight(lh)}
                      aria-pressed={isSelected}
                      className={`min-h-[26px] px-2 rounded-md text-[11px] font-semibold tabular-nums transition-colors ${
                        isSelected
                          ? isDark
                            ? 'bg-indigo-500/30 text-white'
                            : 'bg-white text-slate-900 shadow-sm ring-1 ring-black/5'
                          : isDark
                            ? 'text-slate-400 hover:text-slate-200'
                            : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {lh}
                    </button>
                  );
                })}
              </div>
            </>
          )}

          {activeRibbonTab === 'view' && (
            <>
              {/* 2 Visual Page Layout Cards: Bố cục In & Xem di động */}
              <div
                role="group"
                aria-label="Chọn bố cục trang"
                className="inline-flex items-center gap-1.5 shrink-0"
              >
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleSelectPageLayout('print')}
                  aria-pressed={pageLayout === 'print'}
                  className={`shrink-0 inline-flex items-center gap-1.5 min-h-[30px] px-2.5 rounded-lg text-xs font-semibold border transition-all whitespace-nowrap ${
                    pageLayout === 'print'
                      ? isDark
                        ? 'border-blue-400 bg-blue-500/20 text-blue-100 ring-2 ring-blue-500/25'
                        : 'border-blue-600 bg-blue-50/80 text-blue-700 ring-2 ring-blue-500/20'
                      : isDark
                        ? 'bg-[#1E2534] border-white/10 text-slate-300 hover:bg-white/10'
                        : 'bg-[#FAF8F5] border-[#E5E0D8] text-slate-700 hover:bg-[#EFECE6]'
                  }`}
                >
                  <ScanText className="h-3.5 w-3.5 shrink-0" />
                  <span>Bố cục In</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleSelectPageLayout('mobile')}
                  aria-pressed={pageLayout === 'mobile'}
                  className={`shrink-0 inline-flex items-center gap-1.5 min-h-[30px] px-2.5 rounded-lg text-xs font-semibold border transition-all whitespace-nowrap ${
                    pageLayout === 'mobile'
                      ? isDark
                        ? 'border-blue-400 bg-blue-500/20 text-blue-100 ring-2 ring-blue-500/25'
                        : 'border-blue-600 bg-blue-50/80 text-blue-700 ring-2 ring-blue-500/20'
                      : isDark
                        ? 'bg-[#1E2534] border-white/10 text-slate-300 hover:bg-white/10'
                        : 'bg-[#FAF8F5] border-[#E5E0D8] text-slate-700 hover:bg-[#EFECE6]'
                  }`}
                >
                  <Smartphone className="h-3.5 w-3.5 shrink-0" />
                  <span>Chế độ xem trên di động</span>
                </button>
              </div>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => setIsDesktopSidebarOpen((prev) => !prev)}
                aria-pressed={isDesktopSidebarOpen}
                className={`shrink-0 inline-flex items-center gap-1.5 min-h-[30px] px-2.5 rounded-lg text-xs font-medium border transition-colors whitespace-nowrap ${
                  isDesktopSidebarOpen
                    ? isDark
                      ? 'bg-indigo-500/25 border-indigo-400/40 text-indigo-200'
                      : 'bg-indigo-50 border-indigo-200 text-indigo-700'
                    : isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-300'
                      : 'bg-[#FAF8F5] border-[#E5E0D8] text-slate-700'
                }`}
              >
                <BookOpen className="h-3.5 w-3.5 shrink-0" />
                <span>Mục lục tự động ({outlineEntries.length})</span>
              </button>

              {/* Zoom Controls */}
              <div
                role="group"
                aria-label="Thu phóng trang giấy"
                className={`inline-flex items-center rounded-lg border p-0.5 shrink-0 ${
                  isDark
                    ? 'bg-[#1E2534] border-white/10'
                    : 'bg-[#F5F3EE] border-[#E5E0D8]'
                }`}
              >
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleStepZoom(-1)}
                  disabled={zoomPercent === ZOOM_PERCENTS[0]}
                  aria-label="Thu nhỏ"
                  title="Zoom -"
                  className="min-h-[26px] px-2 rounded-md text-xs font-semibold disabled:opacity-35"
                >
                  <ZoomOut className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setZoomPercent(100)}
                  title="Đặt lại 100%"
                  className="px-2 text-[11px] font-semibold tabular-nums"
                >
                  {zoomPercent}%
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleStepZoom(1)}
                  disabled={
                    zoomPercent === ZOOM_PERCENTS[ZOOM_PERCENTS.length - 1]
                  }
                  aria-label="Phóng to"
                  title="Zoom +"
                  className="min-h-[26px] px-2 rounded-md text-xs font-semibold disabled:opacity-35"
                >
                  <ZoomIn className="h-3.5 w-3.5" />
                </button>
              </div>

              <span
                className={`px-2.5 py-1 rounded-lg border text-[11px] font-medium tabular-nums whitespace-nowrap ${themeStyle.calloutBg} ${themeStyle.calloutBorder} ${themeStyle.textMuted}`}
              >
                {wordCount.toLocaleString('vi-VN')} từ · {detectedPageCount}{' '}
                trang
              </span>
            </>
          )}
        </div>

        {/* Floating Ribbon Popovers (Desktop) */}
        {openPopover === 'block' && (
          <div
            role="menu"
            className={`absolute left-48 top-full mt-1.5 w-48 rounded-xl border p-1.5 shadow-xl z-30 ${
              isDark
                ? 'bg-[#1E2430] border-white/15 text-slate-100'
                : 'bg-white border-[#E5E0D8] text-slate-900'
            }`}
          >
            {BLOCK_STYLE_OPTIONS.map((option) => (
              <button
                key={option.tag}
                type="button"
                role="menuitem"
                onMouseDown={handlePreventMouseDown}
                onClick={() => handleApplyBlockTag(option.tag)}
                className={`w-full text-left px-3 py-2 rounded-lg flex items-center justify-between gap-2 transition-colors ${
                  formatState.block === option.tag
                    ? isDark
                      ? 'bg-indigo-500/25 text-white'
                      : 'bg-indigo-50 text-indigo-700'
                    : isDark
                      ? 'hover:bg-white/[0.06]'
                      : 'hover:bg-slate-100'
                }`}
              >
                <span className={option.previewClass}>{option.label}</span>
                {formatState.block === option.tag && (
                  <Check className="h-3.5 w-3.5 shrink-0" />
                )}
              </button>
            ))}
          </div>
        )}

        {openPopover === 'color' && (
          <div
            className={`absolute left-64 top-full mt-1.5 w-60 max-w-[calc(100vw-1.5rem)] rounded-xl border p-3 shadow-xl z-30 space-y-2.5 ${
              isDark
                ? 'bg-[#1E2430] border-white/15 text-slate-100'
                : 'bg-white border-[#E5E0D8] text-slate-900'
            }`}
          >
            <p className={`text-[11px] font-semibold ${themeStyle.textMuted}`}>
              Chọn màu chữ
            </p>
            <div className="grid grid-cols-8 gap-1.5">
              {TEXT_COLOR_SWATCHES.map((swatch) => (
                <button
                  key={swatch.value}
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => {
                    setCustomTextColor(swatch.value);
                    executeFormatCommand('foreColor', swatch.value);
                    setOpenPopover(null);
                  }}
                  title={swatch.label}
                  aria-label={`Màu chữ ${swatch.label}`}
                  className="h-6 w-6 rounded-full border border-black/20 hover:scale-110 transition-transform"
                  style={{ backgroundColor: swatch.value }}
                />
              ))}
            </div>
            <label className="flex items-center justify-between gap-2 pt-2 border-t border-current/10 text-xs cursor-pointer">
              <span className={themeStyle.textMuted}>Màu tùy chỉnh:</span>
              <input
                type="color"
                value={customTextColor}
                onChange={(event) => {
                  const chosen = event.target.value;
                  setCustomTextColor(chosen);
                  executeFormatCommand('foreColor', chosen);
                }}
                className="h-6 w-10 rounded cursor-pointer border-0 bg-transparent p-0"
              />
            </label>
          </div>
        )}

        {openPopover === 'highlight' && (
          <div
            className={`absolute left-72 top-full mt-1.5 max-w-[calc(100vw-1.5rem)] rounded-xl border p-3 shadow-xl z-30 space-y-2 ${
              isDark
                ? 'bg-[#1E2430] border-white/15 text-slate-100'
                : 'bg-white border-[#E5E0D8] text-slate-900'
            }`}
          >
            <p className={`text-[11px] font-semibold ${themeStyle.textMuted}`}>
              Màu nền highlight
            </p>
            <div className="flex flex-wrap items-center gap-2">
              {HIGHLIGHT_SWATCHES.map((swatch) => (
                <button
                  key={swatch.value}
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => {
                    setCurrentHighlightColor(swatch.value);
                    executeFormatCommand('hiliteColor', swatch.value);
                    setOpenPopover(null);
                  }}
                  title={swatch.label}
                  aria-label={`Màu highlight ${swatch.label}`}
                  className={`h-7 px-2.5 rounded-lg border text-[11px] font-medium flex items-center justify-center transition-transform hover:scale-105 ${
                    swatch.value === 'transparent'
                      ? isDark
                        ? 'border-white/25 text-slate-200'
                        : 'border-slate-300 text-slate-700'
                      : 'border-black/15 text-slate-900'
                  } ${swatch.chipClass}`}
                >
                  {swatch.value === 'transparent' ? 'Xóa' : swatch.label}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Cloud Hydration / Retry Status Strip */}
      {(isHydratingCloudFile || isRenderingDocx || canRetryCloudBuffer) && (
        <div
          className={`shrink-0 px-4 py-2 border-b text-xs flex items-center justify-between gap-3 ${
            isDark
              ? 'bg-indigo-500/15 border-indigo-400/25 text-indigo-200'
              : 'bg-indigo-50/90 border-indigo-200/80 text-indigo-900'
          }`}
        >
          <div className="flex items-center gap-2 min-w-0">
            {isHydratingCloudFile || isRenderingDocx ? (
              <>
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0 text-indigo-500" />
                <span className="truncate">
                  {isHydratingCloudFile
                    ? 'Đang tải bản Word gốc từ đám mây...'
                    : 'Đang dựng trang Word chuẩn định dạng gốc...'}
                </span>
              </>
            ) : (
              <>
                <CloudDownload className="h-3.5 w-3.5 shrink-0 text-indigo-500" />
                <span className="truncate">
                  {docxRenderFailed
                    ? 'Không thể dựng bản gốc .docx — đang hiển thị bản văn bản dự phòng.'
                    : 'Đang hiển thị bản văn bản đã lưu. Bạn có thể tải lại tệp Word .docx gốc từ đám mây.'}
                </span>
              </>
            )}
          </div>

          {canRetryCloudBuffer && onRetryCloudDownload && (
            <button
              type="button"
              onClick={onRetryCloudDownload}
              className="shrink-0 inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-[11px] font-semibold transition-colors"
            >
              <RefreshCw className="h-3 w-3" />
              <span>Thử tải lại</span>
            </button>
          )}
        </div>
      )}

      {/* Export Error Notice */}
      {exportError && (
        <div
          className={`shrink-0 px-4 py-2 bg-rose-500/15 border-b border-rose-400/30 text-xs flex items-center justify-between gap-2 ${
            isDark ? 'text-rose-200' : 'text-rose-700'
          }`}
        >
          <span>{exportError}</span>
          <button
            type="button"
            onClick={() => setExportError(null)}
            aria-label="Đóng thông báo lỗi"
            className="p-1 rounded hover:bg-rose-500/10"
          >
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      )}

      {/* Main Split Body */}
      <div className="flex-1 min-h-0 flex relative overflow-hidden">
        {/* Mobile Drawer Backdrop (< md) */}
        <div
          aria-hidden="true"
          onClick={() => setIsMobileDrawerOpen(false)}
          className={`md:hidden fixed inset-0 z-30 bg-slate-950/55 backdrop-blur-[2px] transition-opacity duration-200 ${
            isMobileDrawerOpen
              ? 'opacity-100 pointer-events-auto'
              : 'opacity-0 pointer-events-none'
          }`}
        />

        {/* Mobile Auto-Outline Drawer (< md, z-40) */}
        <aside
          aria-label="Mục lục tài liệu di động"
          className={`md:hidden fixed inset-y-0 left-0 z-40 w-[272px] max-w-[85vw] border-r shadow-2xl flex flex-col transition-transform duration-200 ease-out ${
            themeStyle.sidebarBg
          } ${themeStyle.sidebarBorder} ${
            isMobileDrawerOpen
              ? 'translate-x-0 pointer-events-auto'
              : '-translate-x-full pointer-events-none'
          }`}
        >
          <div
            className={`h-14 shrink-0 px-4 border-b flex items-center justify-between gap-2 ${themeStyle.sidebarBorder}`}
          >
            <div className="flex items-center gap-2 min-w-0">
              <BookOpen className="h-4 w-4 text-indigo-500 shrink-0" />
              <span className="text-xs font-semibold truncate">
                Mục lục tự động
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsMobileDrawerOpen(false)}
              aria-label="Đóng mục lục"
              className={`inline-flex items-center justify-center min-h-[38px] min-w-[38px] rounded-xl transition-colors ${
                isDark
                  ? 'bg-white/5 hover:bg-white/15 text-slate-300'
                  : 'bg-black/5 hover:bg-black/10 text-slate-700'
              }`}
            >
              <X className="h-4 w-4" />
            </button>
          </div>
          {renderOutlineContent('mobile')}
        </aside>

        {/* Desktop Left Auto-Outline Sidebar (240px, collapsible on md+) */}
        {isDesktopSidebarOpen && (
          <aside
            aria-label="Mục lục tự động"
            className={`hidden md:flex w-[240px] shrink-0 border-r flex-col overflow-hidden transition-colors duration-200 ${themeStyle.sidebarBg} ${themeStyle.sidebarBorder}`}
          >
            <div
              className={`shrink-0 px-4 py-3 border-b flex items-center justify-between gap-2 ${themeStyle.sidebarBorder}`}
            >
              <div className="flex items-center gap-1.5 min-w-0">
                <Sparkles className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
                <span className="text-xs font-semibold truncate">
                  Mục lục tự động
                </span>
              </div>
              <span
                className={`text-[10px] font-semibold uppercase tracking-wider px-1.5 py-0.5 rounded tabular-nums ${
                  isDark
                    ? 'bg-white/10 text-slate-300'
                    : 'bg-black/[0.05] text-slate-500'
                }`}
              >
                {outlineEntries.length} mục
              </span>
            </div>
            {renderOutlineContent('desktop')}
          </aside>
        )}

        {/* Center Editable A4 Paper Canvas */}
        <main
          ref={scrollContainerRef}
          onScroll={handleReadingScroll}
          className={`flex-1 min-w-0 overflow-y-auto scrollbar-none px-2.5 pr-6 py-4 sm:px-6 sm:pr-8 sm:py-7 scroll-smooth ${
            isScaledPrintMode && effectivePrintScale > fitScale
              ? 'overflow-x-auto'
              : 'overflow-x-hidden'
          }`}
        >
          {/* Remote Collaborator Typing Badge */}
          {activeTypingCollaborators.length > 0 && (
            <div className="sticky top-0 z-20 flex justify-center pointer-events-none mb-2">
              <div
                className={`inline-flex items-center gap-2 px-3 py-1 rounded-full border shadow-md text-[11.5px] font-medium backdrop-blur-md ${
                  isDark
                    ? 'bg-[#1A2130]/95 border-white/15 text-slate-100'
                    : 'bg-white/95 border-slate-200 text-slate-800'
                }`}
              >
                {activeTypingCollaborators.slice(0, 2).map((collab) => (
                  <span
                    key={collab.sessionId}
                    className="inline-flex items-center gap-1.5"
                  >
                    <span
                      className="h-2 w-2 rounded-full animate-pulse shrink-0"
                      style={{ backgroundColor: collab.color }}
                    />
                    <span>
                      <strong>{collab.displayName}</strong> đang chỉnh sửa...
                    </span>
                  </span>
                ))}
                {activeTypingCollaborators.length > 2 && (
                  <span className="opacity-70">
                    +{activeTypingCollaborators.length - 2} người khác
                  </span>
                )}
              </div>
            </div>
          )}

          {/* Meta Header Pill above A4 Sheet(s) */}
          <div
            style={
              scaledSheetMaxWidth !== undefined
                ? { maxWidth: `${scaledSheetMaxWidth}px` }
                : undefined
            }
            className="mx-auto mb-3 px-2 flex flex-wrap items-center justify-between gap-2 text-[11.5px]"
          >
            <div className="flex items-center gap-2 min-w-0">
              <span
                className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full font-semibold ${
                  isDark
                    ? 'bg-blue-500/20 text-blue-300'
                    : 'bg-blue-50 text-blue-700 border border-blue-200/70'
                }`}
              >
                <BookOpen className="h-3 w-3" />
                <span>
                  {isReadOnly
                    ? 'Chỉ xem'
                    : isDocxPreviewRendered
                      ? 'Bản Word chuẩn định dạng gốc (.docx)'
                      : 'Soạn thảo trực tiếp A4'}
                </span>
              </span>
              <span className={`truncate ${themeStyle.textMuted}`}>
                {document.updatedAt}
              </span>
            </div>
            <span className={`tabular-nums ${themeStyle.textMuted}`}>
              {wordCount.toLocaleString('vi-VN')} từ · {fontFamily} {fontSize}px
              {zoomPercent !== 100 ? ` · ${zoomPercent}%` : ''}
            </span>
          </div>

          {/* Layout-Aware Page Host (Print Layout with proportional scaling on <768px OR Mobile Reflow View) */}
          {(() => {
            const scaledWidthPx = Math.round(794 * effectivePrintScale);
            const scaledHeightStyle: CSSProperties =
              isScaledPrintMode && contentHeight > 0
                ? {
                    width: `${scaledWidthPx}px`,
                    height: `${Math.ceil(contentHeight * effectivePrintScale)}px`,
                    margin: '0 auto',
                    position: 'relative',
                  }
                : {};
            const innerScaleStyle: CSSProperties = isScaledPrintMode
              ? {
                  ...hostCssVars,
                  width: '794px',
                  maxWidth: '794px',
                  marginLeft: `${(scaledWidthPx - 794) / 2}px`,
                  transform: `scale(${effectivePrintScale})`,
                  transformOrigin: 'top center',
                }
              : hostCssVars;

            return (
              <div style={scaledHeightStyle}>
                <div
                  ref={pageSurfaceRef}
                  style={innerScaleStyle}
                  className={`lite-word-host mx-auto transition-colors duration-200 ${
                    pageLayout === 'mobile'
                      ? 'word-layout-mobile'
                      : 'word-layout-print'
                  } ${
                    isDocxPreviewRendered
                      ? ''
                      : `lite-word-editor-surface rounded-xl sm:rounded-2xl border shadow-sm ${themeStyle.paperBg} ${themeStyle.paperBorder}`
                  }`}
                >
                  <div
                    ref={editorRef}
                    contentEditable={!isReadOnly && !isRenderingDocx}
                    suppressContentEditableWarning
                    spellCheck={false}
                    role="textbox"
                    aria-multiline="true"
                    aria-readonly={isReadOnly}
                    aria-label={`Nội dung soạn thảo tài liệu ${document.title}`}
                    onCompositionStart={() => {
                      isComposingRef.current = true;
                    }}
                    onCompositionEnd={() => {
                      isComposingRef.current = false;
                      handleEditorMutation();
                    }}
                    onInput={() => {
                      if (isComposingRef.current) {
                        return;
                      }
                      handleEditorMutation();
                    }}
                    onKeyUp={() => {
                      captureCurrentSelection();
                      refreshFormatState();
                    }}
                    onMouseUp={() => {
                      captureCurrentSelection();
                      refreshFormatState();
                    }}
                    style={{
                      fontFamily: fontStack,
                      fontSize: `${effectiveFontSize}px`,
                      lineHeight,
                    }}
                    className={
                      isDocxPreviewRendered
                        ? 'min-h-[680px] outline-none break-words'
                        : `min-h-[720px] px-6 sm:px-14 py-8 sm:py-12 outline-none break-words ${themeStyle.textPrimary} ${themeStyle.editorProse} [&_h1]:text-[1.65em] [&_h1]:font-bold [&_h1]:leading-snug [&_h1]:tracking-tight [&_h1]:mb-4 [&_h1]:mt-2 [&_h1]:scroll-mt-6 [&_h2]:text-[1.3em] [&_h2]:font-bold [&_h2]:leading-snug [&_h2]:mt-7 [&_h2]:mb-3 [&_h2]:pb-2 [&_h2]:border-b [&_h2]:scroll-mt-6 [&_h3]:text-[1.12em] [&_h3]:font-semibold [&_h3]:mt-5 [&_h3]:mb-2 [&_h3]:scroll-mt-6 [&_p]:mb-3.5 [&_ul]:list-disc [&_ul]:pl-6 [&_ul]:mb-4 [&_ul]:space-y-1.5 [&_ol]:list-decimal [&_ol]:pl-6 [&_ol]:mb-4 [&_ol]:space-y-1.5 [&_blockquote]:border-l-4 [&_blockquote]:rounded-r-xl [&_blockquote]:px-4 [&_blockquote]:py-3 [&_blockquote]:my-4 [&_table]:w-full [&_table]:border-collapse [&_table]:my-5 [&_table]:text-[0.94em] [&_th]:border [&_th]:px-3 [&_th]:py-2 [&_th]:text-left [&_th]:font-semibold [&_td]:border [&_td]:px-3 [&_td]:py-2 [&_td]:align-top [&_img]:max-w-full [&_img]:h-auto [&_img]:rounded-xl [&_img]:my-4`
                    }
                  />
                </div>
              </div>
            );
          })()}

          {/* Bottom spacer so mobile bottom toolbar & floating TOC never hide content */}
          <div className="h-28 md:h-6" aria-hidden="true" />
        </main>

        {/* Compact Mobile Print-Layout Zoom Pill (only in Bố cục In on narrow screens) */}
        {isScaledPrintMode && !isMobileDrawerOpen && !isMobileSheetExpanded && (
          <div
            role="group"
            aria-label="Thu phóng Bố cục In"
            className={`md:hidden fixed bottom-14 right-5 z-30 inline-flex items-center gap-0.5 p-1 rounded-full border shadow-lg backdrop-blur-md ${
              isDark
                ? 'bg-[#1E2430]/95 border-white/15 text-white'
                : 'bg-white/95 border-[#E2DDD4] text-slate-900'
            }`}
          >
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => handleStepPrintZoom(-0.1)}
              disabled={effectivePrintScale <= 0.42}
              aria-label="Thu nhỏ trang A4"
              title="Thu nhỏ"
              className="h-7 w-7 rounded-full inline-flex items-center justify-center hover:bg-black/5 disabled:opacity-35"
            >
              <ZoomOut className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => setPrintZoom(null)}
              title="Đặt lại vừa màn hình"
              className={`px-2 py-1 rounded-full text-[11px] font-semibold tabular-nums whitespace-nowrap ${
                printZoom === null
                  ? isDark
                    ? 'bg-blue-500/25 text-blue-200'
                    : 'bg-blue-50 text-blue-700'
                  : 'hover:bg-black/5'
              }`}
            >
              {printZoom === null ? 'Vừa màn hình · ' : ''}
              {Math.round(effectivePrintScale * 100)}%
            </button>
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => handleStepPrintZoom(0.1)}
              disabled={effectivePrintScale >= 1.48}
              aria-label="Phóng to trang A4"
              title="Phóng to"
              className="h-7 w-7 rounded-full inline-flex items-center justify-center hover:bg-black/5 disabled:opacity-35"
            >
              <ZoomIn className="h-3.5 w-3.5" />
            </button>
          </div>
        )}

        {/* Custom Tactile Vertical Scrollbar (▲ / Pill Thumb / ▼) */}
        <CustomVerticalScrollbar
          scrollContainerRef={scrollContainerRef}
          theme={isDark ? 'dark' : 'light'}
          bottomOffsetClass="bottom-12 md:bottom-0"
          getDragLabel={({ percent }) => {
            const totalPages = Math.max(1, detectedPageCount);
            const estPage = Math.min(
              totalPages,
              Math.max(1, Math.round((percent / 100) * (totalPages - 1)) + 1)
            );
            return `Trang ${estPage} / ${totalPages} · ${Math.round(percent)}%`;
          }}
        />

        {/* Floating Mobile Outline Trigger Button (< md, above h-12 bottom bar) */}
        {!isMobileDrawerOpen && !isMobileSheetExpanded && (
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở Mục lục tự động"
            className={`md:hidden fixed bottom-14 left-3 z-30 inline-flex items-center gap-2 min-h-[38px] px-3 rounded-full border shadow-lg text-xs font-semibold transition-all ${
              isDark
                ? 'bg-[#1E2430]/95 border-white/15 text-white hover:bg-[#273041]'
                : 'bg-white/95 border-[#E2DDD4] text-slate-900 hover:bg-white'
            }`}
          >
            <List className="h-3.5 w-3.5 text-indigo-500" />
            <span>Mục lục</span>
            {outlineEntries.length > 0 && (
              <span className="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-indigo-500 text-white text-[10px] font-bold tabular-nums">
                {outlineEntries.length}
              </span>
            )}
          </button>
        )}

        {/* Mobile Collapsed Bottom Toolbar (Word Mobile 1-Row Dock, md:hidden) */}
        {!isMobileSheetExpanded && !isReadOnly && (
          <div
            role="toolbar"
            aria-label="Thanh công cụ nhanh Word di động"
            className="fixed bottom-0 inset-x-0 z-40 bg-slate-100 border-t border-slate-200 h-12 flex items-center text-slate-800 md:hidden"
          >
            <div className="overflow-x-auto no-scrollbar flex items-center gap-1 px-2 flex-1 min-w-0">
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('bold')}
                aria-pressed={formatState.bold}
                aria-label="In đậm"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center text-sm font-bold shrink-0 transition-colors ${
                  formatState.bold
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <Bold className="h-4 w-4" />
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('italic')}
                aria-pressed={formatState.italic}
                aria-label="In nghiêng"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.italic
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <Italic className="h-4 w-4" />
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('underline')}
                aria-pressed={formatState.underline}
                aria-label="Gạch chân"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.underline
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <Underline className="h-4 w-4" />
              </button>

              {/* Highlight Pen */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() =>
                  executeFormatCommand('hiliteColor', currentHighlightColor)
                }
                aria-label="Bút tô sáng"
                className="h-9 min-w-[36px] px-2 rounded-lg inline-flex flex-col items-center justify-center shrink-0 text-slate-700 hover:bg-slate-200/80"
              >
                <Highlighter className="h-3.5 w-3.5 text-amber-600" />
                <span
                  aria-hidden="true"
                  className="mt-0.5 h-1 w-4 rounded-full border border-black/15"
                  style={{
                    backgroundColor:
                      currentHighlightColor === 'transparent'
                        ? '#FEF08A'
                        : currentHighlightColor,
                  }}
                />
              </button>

              {/* Font Color A with Color Bar Underneath */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() =>
                  executeFormatCommand('foreColor', customTextColor)
                }
                aria-label="Màu chữ A"
                className="h-9 min-w-[36px] px-2 rounded-lg inline-flex flex-col items-center justify-center shrink-0 text-slate-800 hover:bg-slate-200/80"
              >
                <span className="text-xs font-extrabold leading-none">A</span>
                <span
                  aria-hidden="true"
                  className="mt-0.5 h-1 w-4 rounded-full"
                  style={{ backgroundColor: customTextColor }}
                />
              </button>

              <span
                aria-hidden="true"
                className="h-5 w-px bg-slate-300 mx-0.5 shrink-0"
              />

              {/* Bullet & Numbered Lists */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('insertUnorderedList')}
                aria-pressed={formatState.unorderedList}
                aria-label="Danh sách chấm"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.unorderedList
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <List className="h-4 w-4" />
              </button>

              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('insertOrderedList')}
                aria-pressed={formatState.orderedList}
                aria-label="Danh sách số"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.orderedList
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <ListOrdered className="h-4 w-4" />
              </button>

              <span
                aria-hidden="true"
                className="h-5 w-px bg-slate-300 mx-0.5 shrink-0"
              />

              {/* Alignment Buttons */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('justifyLeft')}
                aria-pressed={formatState.align === 'left'}
                aria-label="Căn trái"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.align === 'left'
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <AlignLeft className="h-4 w-4" />
              </button>
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('justifyCenter')}
                aria-pressed={formatState.align === 'center'}
                aria-label="Căn giữa"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.align === 'center'
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <AlignCenter className="h-4 w-4" />
              </button>
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('justifyRight')}
                aria-pressed={formatState.align === 'right'}
                aria-label="Căn phải"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.align === 'right'
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <AlignRight className="h-4 w-4" />
              </button>
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => executeFormatCommand('justifyFull')}
                aria-pressed={formatState.align === 'justify'}
                aria-label="Căn đều"
                className={`h-9 min-w-[36px] px-2 rounded-lg inline-flex items-center justify-center shrink-0 transition-colors ${
                  formatState.align === 'justify'
                    ? 'bg-blue-100 text-blue-700 ring-1 ring-blue-300'
                    : 'text-slate-700 hover:bg-slate-200/80'
                }`}
              >
                <AlignJustify className="h-4 w-4" />
              </button>
            </div>

            {/* Expand Arrow Button ▲ */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => setIsMobileSheetExpanded(true)}
              aria-label="Mở rộng bảng công cụ Word"
              title="Mở rộng bảng công cụ"
              className="shrink-0 h-full px-3.5 border-l border-slate-300 inline-flex items-center justify-center text-slate-700 hover:bg-slate-200 transition-colors"
            >
              <ChevronUp className="h-4 w-4" />
            </button>
          </div>
        )}

        {/* Mobile Expanded Tool Sheet (Word Mobile Bottom Sheet, md:hidden, no blocking backdrop) */}
        {isMobileSheetExpanded && !isReadOnly && (
          <section
            aria-label="Bảng công cụ chi tiết Word di động"
            className="fixed bottom-0 inset-x-0 z-40 h-[44vh] bg-white border-t border-slate-200 rounded-t-2xl shadow-2xl flex flex-col text-slate-900 md:hidden"
          >
            {/* Sheet Header Row */}
            <div className="shrink-0 px-3.5 py-2 border-b border-slate-200 flex items-center justify-between gap-2 relative">
              {/* Section Switcher Dropdown Button: Trang chủ ⌄ */}
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => setIsMobileTabMenuOpen((prev) => !prev)}
                aria-expanded={isMobileTabMenuOpen}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-blue-50 hover:bg-blue-100/80 text-blue-700 text-sm font-bold transition-colors"
              >
                <span>{RIBBON_TAB_LABELS[activeRibbonTab]}</span>
                <ChevronDown
                  className={`h-4 w-4 transition-transform ${
                    isMobileTabMenuOpen ? 'rotate-180' : ''
                  }`}
                />
              </button>

              {/* Right Cluster: Lightbulb + Undo + Redo + Collapse ▼ */}
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setIsMobileCommandSearchOpen((prev) => !prev)
                  }
                  aria-pressed={isMobileCommandSearchOpen}
                  aria-label="Gợi ý và tìm lệnh nhanh"
                  title="Gợi ý lệnh"
                  className={`h-9 w-9 rounded-xl inline-flex items-center justify-center transition-colors ${
                    isMobileCommandSearchOpen
                      ? 'bg-amber-100 text-amber-700'
                      : 'text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  <Lightbulb className="h-4 w-4" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('undo')}
                  aria-label="Hoàn tác"
                  title="Hoàn tác"
                  className="h-9 w-9 rounded-xl inline-flex items-center justify-center text-slate-600 hover:bg-slate-100 transition-colors"
                >
                  <Undo2 className="h-4 w-4" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('redo')}
                  aria-label="Làm lại"
                  title="Làm lại"
                  className="h-9 w-9 rounded-xl inline-flex items-center justify-center text-slate-600 hover:bg-slate-100 transition-colors"
                >
                  <Redo2 className="h-4 w-4" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => {
                    setIsMobileTabMenuOpen(false);
                    setIsMobileSheetExpanded(false);
                  }}
                  aria-label="Thu gọn bảng công cụ"
                  title="Thu gọn"
                  className="h-9 w-9 rounded-xl inline-flex items-center justify-center text-slate-700 hover:bg-slate-100 transition-colors"
                >
                  <ChevronDown className="h-4 w-4" />
                </button>
              </div>

              {/* Floating Blue Section Menu (#1E40AF) */}
              {isMobileTabMenuOpen && (
                <div
                  role="menu"
                  className="absolute left-3.5 top-full mt-1 w-48 rounded-xl overflow-hidden shadow-2xl border border-blue-900/20 bg-[#1E40AF] text-white z-50 p-1.5 space-y-1"
                >
                  {RIBBON_TABS.map((tab) => {
                    const isSelected = activeRibbonTab === tab.key;
                    return (
                      <button
                        key={tab.key}
                        type="button"
                        role="menuitem"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => {
                          setActiveRibbonTab(tab.key);
                          setIsMobileTabMenuOpen(false);
                        }}
                        className={`w-full text-left px-3 py-2 rounded-lg text-xs flex items-center justify-between transition-colors ${
                          isSelected
                            ? 'bg-white text-[#1E40AF] font-bold shadow-sm'
                            : 'text-white hover:bg-white/15 font-medium'
                        }`}
                      >
                        <span>{tab.label}</span>
                        {isSelected && <Check className="h-3.5 w-3.5" />}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Quick Command Search Bar (toggled by Lightbulb) */}
            {isMobileCommandSearchOpen && (
              <div className="shrink-0 px-3.5 py-2 bg-slate-50 border-b border-slate-200 space-y-2">
                <input
                  type="text"
                  value={mobileCommandQuery}
                  onChange={(event) =>
                    setMobileCommandQuery(event.target.value)
                  }
                  placeholder="Tìm nhanh lệnh (vd: bảng, tiêu đề, ngắt trang...)"
                  className="w-full h-8 px-3 rounded-lg border border-slate-300 bg-white text-xs text-slate-900 outline-none focus:border-blue-600"
                />
                <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar pb-0.5">
                  {filteredQuickCommands.map((cmd) => (
                    <button
                      key={cmd.label}
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={cmd.action}
                      className="shrink-0 px-2.5 py-1 rounded-full bg-white border border-slate-300 text-[11px] font-medium text-slate-700 hover:bg-blue-50 hover:text-blue-700"
                    >
                      {cmd.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {/* Sheet Scrollable Content */}
            <div className="flex-1 min-h-0 overflow-y-auto px-3.5 py-3 space-y-3 pb-safe">
              {activeRibbonTab === 'home' && (
                <>
                  {/* Row 1: Font Family & Font Size */}
                  <div className="flex items-center gap-2">
                    <div className="flex-1 grid grid-cols-2 gap-1 p-1 rounded-xl bg-slate-100 border border-slate-200">
                      {FONT_FAMILIES.map((family) => {
                        const isSelected = fontFamily === family;
                        return (
                          <button
                            key={family}
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => handleSelectFontFamily(family)}
                            aria-pressed={isSelected}
                            className={`py-1.5 px-2 rounded-lg text-xs font-semibold truncate transition-all ${
                              isSelected
                                ? 'bg-white text-blue-700 shadow-sm'
                                : 'text-slate-600'
                            }`}
                          >
                            {family}
                          </button>
                        );
                      })}
                    </div>

                    <div className="inline-flex items-center gap-1 p-1 rounded-xl bg-slate-100 border border-slate-200 shrink-0">
                      {FONT_SIZES.map((sz, idx) => {
                        const isSelected = fontSizeIndex === idx;
                        return (
                          <button
                            key={sz}
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => setFontSizeIndex(idx)}
                            className={`h-7 min-w-[26px] px-1.5 rounded-lg text-xs font-bold tabular-nums transition-colors ${
                              isSelected
                                ? 'bg-blue-600 text-white'
                                : 'text-slate-700 hover:bg-white'
                            }`}
                          >
                            {sz}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {/* Row 2: 4 Large Buttons B, I, U, ab (strikeThrough) */}
                  <div className="grid grid-cols-4 gap-2">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('bold')}
                      aria-pressed={formatState.bold}
                      className={`h-10 rounded-xl border font-bold text-sm flex items-center justify-center transition-colors ${
                        formatState.bold
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}
                    >
                      B
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('italic')}
                      aria-pressed={formatState.italic}
                      className={`h-10 rounded-xl border italic font-serif text-sm flex items-center justify-center transition-colors ${
                        formatState.italic
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}
                    >
                      I
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('underline')}
                      aria-pressed={formatState.underline}
                      className={`h-10 rounded-xl border underline text-sm font-semibold flex items-center justify-center transition-colors ${
                        formatState.underline
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}
                    >
                      U
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('strikeThrough')}
                      aria-pressed={formatState.strikeThrough}
                      className={`h-10 rounded-xl border line-through text-xs font-semibold flex items-center justify-center transition-colors ${
                        formatState.strikeThrough
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}
                    >
                      ab
                    </button>
                  </div>

                  {/* Row 3: Tô sáng > */}
                  <div className="rounded-xl border border-slate-200 bg-slate-50/70 overflow-hidden">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        setMobileExpandPanel((prev) =>
                          prev === 'highlight' ? null : 'highlight'
                        )
                      }
                      className="w-full px-3.5 py-2.5 flex items-center justify-between text-xs font-semibold text-slate-800"
                    >
                      <span className="inline-flex items-center gap-2">
                        <Highlighter className="h-4 w-4 text-amber-500" />
                        <span>Tô sáng</span>
                      </span>
                      <ChevronRight
                        className={`h-4 w-4 text-slate-400 transition-transform ${
                          mobileExpandPanel === 'highlight' ? 'rotate-90' : ''
                        }`}
                      />
                    </button>
                    {mobileExpandPanel === 'highlight' && (
                      <div className="px-3.5 pb-3 pt-1 flex flex-wrap items-center gap-2 border-t border-slate-200/80">
                        {HIGHLIGHT_SWATCHES.map((swatch) => (
                          <button
                            key={swatch.value}
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => {
                              setCurrentHighlightColor(swatch.value);
                              executeFormatCommand('hiliteColor', swatch.value);
                            }}
                            className={`h-8 px-3 rounded-lg border text-xs font-medium ${swatch.chipClass} border-slate-300 text-slate-800`}
                          >
                            {swatch.label}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Row 4: Màu Phông > */}
                  <div className="rounded-xl border border-slate-200 bg-slate-50/70 overflow-hidden">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        setMobileExpandPanel((prev) =>
                          prev === 'color' ? null : 'color'
                        )
                      }
                      className="w-full px-3.5 py-2.5 flex items-center justify-between text-xs font-semibold text-slate-800"
                    >
                      <span className="inline-flex items-center gap-2">
                        <Palette className="h-4 w-4 text-blue-600" />
                        <span>Màu Phông</span>
                        <span
                          aria-hidden="true"
                          className="h-3 w-3 rounded-full border border-black/20"
                          style={{ backgroundColor: customTextColor }}
                        />
                      </span>
                      <ChevronRight
                        className={`h-4 w-4 text-slate-400 transition-transform ${
                          mobileExpandPanel === 'color' ? 'rotate-90' : ''
                        }`}
                      />
                    </button>
                    {mobileExpandPanel === 'color' && (
                      <div className="px-3.5 pb-3 pt-2 flex flex-wrap items-center gap-2.5 border-t border-slate-200/80">
                        {TEXT_COLOR_SWATCHES.map((swatch) => (
                          <button
                            key={swatch.value}
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => {
                              setCustomTextColor(swatch.value);
                              executeFormatCommand('foreColor', swatch.value);
                            }}
                            title={swatch.label}
                            className="h-7 w-7 rounded-full border border-black/20"
                            style={{ backgroundColor: swatch.value }}
                          />
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Row 5: Căn lề & Danh sách */}
                  <div className="grid grid-cols-6 gap-1.5">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('justifyLeft')}
                      aria-pressed={formatState.align === 'left'}
                      aria-label="Căn trái"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.align === 'left'
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <AlignLeft className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('justifyCenter')}
                      aria-pressed={formatState.align === 'center'}
                      aria-label="Căn giữa"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.align === 'center'
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <AlignCenter className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('justifyRight')}
                      aria-pressed={formatState.align === 'right'}
                      aria-label="Căn phải"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.align === 'right'
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <AlignRight className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('justifyFull')}
                      aria-pressed={formatState.align === 'justify'}
                      aria-label="Căn đều"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.align === 'justify'
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <AlignJustify className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() =>
                        executeFormatCommand('insertUnorderedList')
                      }
                      aria-pressed={formatState.unorderedList}
                      aria-label="Danh sách chấm"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.unorderedList
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <List className="h-4 w-4" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => executeFormatCommand('insertOrderedList')}
                      aria-pressed={formatState.orderedList}
                      aria-label="Danh sách số"
                      className={`h-9 rounded-xl border flex items-center justify-center ${
                        formatState.orderedList
                          ? 'bg-blue-50 border-blue-400 text-blue-700'
                          : 'bg-slate-50 border-slate-200 text-slate-700'
                      }`}
                    >
                      <ListOrdered className="h-4 w-4" />
                    </button>
                  </div>

                  {/* Row 6: Kiểu đoạn / Tiêu đề */}
                  <div className="grid grid-cols-4 gap-1.5">
                    {BLOCK_STYLE_OPTIONS.map((opt) => (
                      <button
                        key={opt.tag}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleApplyBlockTag(opt.tag)}
                        className={`py-2 px-2 rounded-xl border text-xs font-semibold truncate ${
                          formatState.block === opt.tag
                            ? 'bg-blue-50 border-blue-400 text-blue-700'
                            : 'bg-slate-50 border-slate-200 text-slate-700'
                        }`}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>

                  {/* Row 7: Xóa Định dạng */}
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => executeFormatCommand('removeFormat')}
                    className="w-full py-2.5 px-3.5 rounded-xl border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 flex items-center justify-center gap-2"
                  >
                    <RemoveFormatting className="h-4 w-4 text-rose-500" />
                    <span>Xóa Định dạng</span>
                  </button>
                </>
              )}

              {activeRibbonTab === 'insert' && (
                <div className="space-y-3">
                  <div className="grid grid-cols-3 gap-2">
                    {(
                      [
                        { rows: 2, cols: 2, label: 'Bảng 2×2' },
                        { rows: 3, cols: 3, label: 'Bảng 3×3' },
                        { rows: 4, cols: 4, label: 'Bảng 4×4' },
                      ] as const
                    ).map((tbl) => (
                      <button
                        key={tbl.label}
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleInsertTable(tbl.rows, tbl.cols)}
                        className="py-2.5 px-3 rounded-xl border border-slate-200 bg-slate-50 hover:bg-blue-50 text-xs font-semibold text-slate-800 flex items-center justify-center gap-1.5"
                      >
                        <Table className="h-4 w-4 text-blue-600" />
                        <span>{tbl.label}</span>
                      </button>
                    ))}
                  </div>

                  <div className="grid grid-cols-1 gap-2">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleInsertHorizontalRule}
                      className="py-2.5 px-3.5 rounded-xl border border-slate-200 bg-slate-50 text-xs font-semibold text-slate-800 flex items-center gap-2.5"
                    >
                      <Minus className="h-4 w-4 text-blue-600" />
                      <span>Chèn đường kẻ ngang</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleInsertPageBreak}
                      className="py-2.5 px-3.5 rounded-xl border border-slate-200 bg-slate-50 text-xs font-semibold text-slate-800 flex items-center gap-2.5"
                    >
                      <SplitSquareVertical className="h-4 w-4 text-blue-600" />
                      <span>Chèn dấu ngắt trang</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={handleInsertCalloutBox}
                      className="py-2.5 px-3.5 rounded-xl border border-slate-200 bg-slate-50 text-xs font-semibold text-slate-800 flex items-center gap-2.5"
                    >
                      <Quote className="h-4 w-4 text-blue-600" />
                      <span>Chèn khung ghi chú / trích dẫn</span>
                    </button>
                  </div>
                </div>
              )}

              {activeRibbonTab === 'layout' && (
                <div className="space-y-3">
                  <div className="space-y-1.5">
                    <p className="text-xs font-semibold text-slate-600">
                      Màu nền trang giấy
                    </p>
                    <div className="grid grid-cols-3 gap-2">
                      {(['light', 'sepia', 'dark'] as const).map((mode) => {
                        const config = PAPER_THEMES[mode];
                        const isSelected = paperTheme === mode;
                        return (
                          <button
                            key={mode}
                            type="button"
                            onMouseDown={handlePreventMouseDown}
                            onClick={() => setPaperTheme(mode)}
                            className={`py-2.5 px-2.5 rounded-xl border text-xs font-semibold flex items-center justify-center gap-1.5 ${
                              isSelected
                                ? 'bg-blue-50 border-blue-500 text-blue-700'
                                : 'bg-slate-50 border-slate-200 text-slate-700'
                            }`}
                          >
                            <span
                              className={`h-3.5 w-3.5 rounded-full border ${config.swatchClass}`}
                            />
                            <span>{config.label}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <p className="text-xs font-semibold text-slate-600">
                      Khoảng cách giãn dòng
                    </p>
                    <div className="grid grid-cols-3 gap-2">
                      {LINE_HEIGHTS.map((lh) => (
                        <button
                          key={lh}
                          type="button"
                          onMouseDown={handlePreventMouseDown}
                          onClick={() => setLineHeight(lh)}
                          className={`py-2.5 rounded-xl border text-xs font-bold tabular-nums ${
                            lineHeight === lh
                              ? 'bg-blue-50 border-blue-500 text-blue-700'
                              : 'bg-slate-50 border-slate-200 text-slate-700'
                          }`}
                        >
                          {lh}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}

              {activeRibbonTab === 'view' && (
                <div className="space-y-3">
                  {/* 2 Visual Page Layout Cards (Word Mobile style) */}
                  <div className="grid grid-cols-2 gap-2.5">
                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleSelectPageLayout('print')}
                      aria-pressed={pageLayout === 'print'}
                      className={`p-3 rounded-xl border text-left flex flex-col items-center justify-center gap-2 transition-all ${
                        pageLayout === 'print'
                          ? 'border-blue-600 bg-blue-50/80 ring-2 ring-blue-500/20 text-blue-700'
                          : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      <ScanText className="h-6 w-6 shrink-0" />
                      <span className="text-xs font-bold text-center leading-tight">
                        Bố cục In
                      </span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={handlePreventMouseDown}
                      onClick={() => handleSelectPageLayout('mobile')}
                      aria-pressed={pageLayout === 'mobile'}
                      className={`p-3 rounded-xl border text-left flex flex-col items-center justify-center gap-2 transition-all ${
                        pageLayout === 'mobile'
                          ? 'border-blue-600 bg-blue-50/80 ring-2 ring-blue-500/20 text-blue-700'
                          : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      <Smartphone className="h-6 w-6 shrink-0" />
                      <span className="text-xs font-bold text-center leading-tight">
                        Chế độ xem trên di động
                      </span>
                    </button>
                  </div>

                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => {
                      setIsMobileSheetExpanded(false);
                      setIsMobileDrawerOpen(true);
                    }}
                    className="w-full py-2.5 px-3.5 rounded-xl bg-blue-600 text-white text-xs font-semibold flex items-center justify-between"
                  >
                    <span className="inline-flex items-center gap-2">
                      <BookOpen className="h-4 w-4" />
                      <span>Mở Mục lục tự động (Auto-Outline)</span>
                    </span>
                    <span className="px-2 py-0.5 rounded-full bg-white/20 text-[11px] tabular-nums">
                      {outlineEntries.length} mục
                    </span>
                  </button>

                  <div className="flex items-center justify-between gap-2 p-2.5 rounded-xl border border-slate-200 bg-slate-50">
                    <span className="text-xs font-semibold text-slate-700">
                      Thu phóng trang
                    </span>
                    <div className="inline-flex items-center gap-1">
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleStepZoom(-1)}
                        disabled={zoomPercent === ZOOM_PERCENTS[0]}
                        className="h-8 px-2.5 rounded-lg bg-white border border-slate-300 text-xs font-bold disabled:opacity-40"
                      >
                        Zoom -
                      </button>
                      <span className="px-2 text-xs font-bold tabular-nums">
                        {zoomPercent}%
                      </span>
                      <button
                        type="button"
                        onMouseDown={handlePreventMouseDown}
                        onClick={() => handleStepZoom(1)}
                        disabled={
                          zoomPercent ===
                          ZOOM_PERCENTS[ZOOM_PERCENTS.length - 1]
                        }
                        className="h-8 px-2.5 rounded-lg bg-white border border-slate-300 text-xs font-bold disabled:opacity-40"
                      >
                        Zoom +
                      </button>
                    </div>
                  </div>

                  <div className="p-3 rounded-xl border border-slate-200 bg-slate-50 flex items-center justify-between text-xs text-slate-700">
                    <span>Thống kê tài liệu:</span>
                    <span className="font-bold tabular-nums">
                      {wordCount.toLocaleString('vi-VN')} từ ·{' '}
                      {detectedPageCount} trang
                    </span>
                  </div>
                </div>
              )}
            </div>
          </section>
        )}
      </div>
    </div>
  );
}

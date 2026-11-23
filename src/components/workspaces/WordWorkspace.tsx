import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent, CSSProperties, MouseEvent as ReactMouseEvent } from 'react';
import { renderAsync } from 'docx-preview';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  BookOpen,
  CaseSensitive,
  Check,
  ChevronDown,
  CloudDownload,
  Columns,
  Download,
  Eye,
  FileDown,
  FileSearch,
  FolderOpen,
  Highlighter,
  Image as ImageIcon,
  IndentDecrease,
  IndentIncrease,
  Italic,
  Link2,
  List,
  ListOrdered,
  Loader2,
  Lock,
  Maximize2,
  MessageSquarePlus,
  Minimize2,
  Minus,
  Moon,
  Paintbrush,
  Palette,
  PanelLeftClose,
  PanelLeftOpen,
  Pencil,
  Quote,
  Redo2,
  RefreshCw,
  RemoveFormatting,
  Ruler,
  Save,
  ScanText,
  Search,
  Shapes,
  Sigma,
  Smartphone,
  Sparkles,
  SplitSquareVertical,
  Square,
  Strikethrough,
  Subscript,
  Sun,
  Superscript,
  Table,
  Underline,
  Undo2,
  X,
} from 'lucide-react';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  WordCommentItem,
  WordPaperTheme,
} from '../../types/document';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import {
  extractOutlineFromHtml,
  isLikelyStructuralOutlineTitle,
} from '../../utils/fileParser';
import {
  exportOriginalDocumentBinary,
  exportWordToDocx,
} from '../../utils/fileExporter';
import { useWorkspaceSync } from '../../contexts/WorkspaceSyncContext';
import { useDocumentPinchZoom } from '../../hooks/useDocumentPinchZoom';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { FloatingSelectionToolbar } from './FloatingSelectionToolbar';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { OfficeStatusBar } from './OfficeStatusBar';
import { WorkspaceTopBar } from './WorkspaceTopBar';
import {
  WordCommentsPanel,
  WordFindReplaceBar,
  WordHorizontalRuler,
  WordInsertToolModals,
  WordTableContextBar,
  WordTableGridPicker,
  applyCaseTransformToText,
  cmToPx,
  computeWordDocumentStats,
  getWordPageDimensions,
  type WordCaseTransform,
  type WordColumnCount,
  type WordMarginPreset,
  type WordPageOrientation,
  type WordPaperSize,
  type WordToolModalKind,
  type WordUnderlineStyle,
} from './WordInsertTools';

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
type WordEditMode = 'original' | 'edit';
const WORD_PAGE_LAYOUT_STORAGE_KEY = 'lite_word_page_layout';

function resolveInitialWordEditMode(
  doc: DocumentItem,
  isReadOnly: boolean
): WordEditMode {
  if (isReadOnly) {
    return 'original';
  }
  const hasOriginalBinary = Boolean(
    doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0
  );
  if (hasOriginalBinary && !doc.hasLocalEdits) {
    return 'original';
  }
  return 'edit';
}

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

const FONT_SIZES = [10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36] as const;
type FontSizeOption = (typeof FONT_SIZES)[number];

const FONT_FAMILIES = [
  'Times New Roman',
  'Arial',
  'Calibri',
  'Georgia',
  'Roboto',
  'Courier New',
] as const;
type WordFontFamily = (typeof FONT_FAMILIES)[number];

const LINE_HEIGHTS = [1.0, 1.15, 1.5, 2.0] as const;
type WordLineHeight = (typeof LINE_HEIGHTS)[number];

const PARAGRAPH_SPACINGS = [
  { label: 'Gọn (4px)', valuePx: 4 },
  { label: 'Chuẩn (12px)', valuePx: 12 },
  { label: 'Thoáng (20px)', valuePx: 20 },
] as const;

type RibbonTab = 'home' | 'insert' | 'layout' | 'view';
type BlockTag = 'p' | 'h1' | 'h2' | 'h3';
type ActivePopover =
  | 'block'
  | 'color'
  | 'highlight'
  | 'underline'
  | 'case'
  | 'table_grid'
  | null;

const RIBBON_TABS: { key: RibbonTab; label: string }[] = [
  { key: 'home', label: 'Trang đầu' },
  { key: 'insert', label: 'Chèn' },
  { key: 'layout', label: 'Bố trí' },
  { key: 'view', label: 'Xem lại & Công cụ' },
];

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
  switch (family) {
    case 'Arial':
      return 'Arial, Helvetica, sans-serif';
    case 'Calibri':
      return "Calibri, 'Segoe UI', Arial, sans-serif";
    case 'Georgia':
      return "Georgia, 'Times New Roman', serif";
    case 'Roboto':
      return "'Roboto', 'Inter', Arial, sans-serif";
    case 'Courier New':
      return "'Courier New', Courier, monospace";
    default:
      return "'Times New Roman', Times, serif";
  }
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
  clone.querySelectorAll('mark[data-word-find="true"]').forEach((mark) => {
    const parent = mark.parentNode;
    if (!parent) return;
    while (mark.firstChild) {
      parent.insertBefore(mark.firstChild, mark);
    }
    parent.removeChild(mark);
  });
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
  const { onDownloadOriginalFile } = useWorkspaceSync();

  const [wordEditMode, setWordEditMode] = useState<WordEditMode>(() =>
    resolveInitialWordEditMode(document, isReadOnly)
  );
  const isEditingLocked = isReadOnly || wordEditMode === 'original';

  const [paperTheme, setPaperTheme] =
    useState<WordPaperTheme>(defaultPaperTheme);
  const [pageLayout, setPageLayout] = useState<WordPageLayout>(() =>
    readStoredPageLayout()
  );
  const [containerWidth, setContainerWidth] = useState<number>(0);
  const [contentHeight, setContentHeight] = useState<number>(0);
  const [fontFamily, setFontFamily] =
    useState<WordFontFamily>('Times New Roman');
  const [fontSizeIndex, setFontSizeIndex] = useState<number>(5); // 16px default
  const [lineHeight, setLineHeight] = useState<WordLineHeight>(1.5);
  const [paragraphSpacingPx, setParagraphSpacingPx] = useState<number>(12);
  const [zoomPercent, setZoomPercent] = useState<number>(100);
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTab>('home');

  const [paperSize, setPaperSize] = useState<WordPaperSize>('A4');
  const [pageOrientation, setPageOrientation] =
    useState<WordPageOrientation>('portrait');
  const [marginPreset, setMarginPreset] = useState<WordMarginPreset>('normal');
  const [marginLeftCm, setMarginLeftCm] = useState<number>(2.0);
  const [marginRightCm, setMarginRightCm] = useState<number>(2.0);
  const [marginTopCm, setMarginTopCm] = useState<number>(2.0);
  const [marginBottomCm, setMarginBottomCm] = useState<number>(2.0);
  const [firstLineIndentCm, setFirstLineIndentCm] = useState<number>(0);
  const [columnCount, setColumnCount] = useState<WordColumnCount>(1);
  const [pageBorderEnabled, setPageBorderEnabled] = useState<boolean>(false);
  const [showRuler, setShowRuler] = useState<boolean>(true);
  const [showHeaderFooter, setShowHeaderFooter] = useState<boolean>(false);
  const [headerText, setHeaderText] = useState<string>(document.title);
  const [footerText, setFooterText] = useState<string>('LITE Workspace');
  const [showPageNumbers, setShowPageNumbers] = useState<boolean>(true);

  const [isReadMode, setIsReadMode] = useState<boolean>(false);
  const [isCommentsOpen, setIsCommentsOpen] = useState<boolean>(false);
  const [wordComments, setWordComments] = useState<WordCommentItem[]>(
    () => document.wordComments ?? []
  );
  const [selectedQuote, setSelectedQuote] = useState<string>('');

  const [isFindOpen, setIsFindOpen] = useState<boolean>(false);
  const [findQuery, setFindQuery] = useState<string>('');
  const [replaceText, setReplaceText] = useState<string>('');
  const [findCaseSensitive, setFindCaseSensitive] = useState<boolean>(false);
  const [findMatchCount, setFindMatchCount] = useState<number>(0);
  const [activeFindMatchIndex, setActiveFindMatchIndex] = useState<number>(0);

  const [activeToolModal, setActiveToolModal] =
    useState<WordToolModalKind>(null);
  const [isFormatPainterActive, setIsFormatPainterActive] =
    useState<boolean>(false);
  const [activeTableCell, setActiveTableCell] =
    useState<HTMLTableCellElement | null>(null);
  const [selectedImageEl, setSelectedImageEl] =
    useState<HTMLImageElement | null>(null);

  const [isDesktopSidebarOpen, setIsDesktopSidebarOpen] =
    useState<boolean>(true);
  const [isMobileDrawerOpen, setIsMobileDrawerOpen] = useState<boolean>(false);

  const [outlineEntries, setOutlineEntries] = useState<OutlineNavEntry[]>([]);
  const [activeSectionId, setActiveSectionId] = useState<string>('');
  const [currentScrollPage, setCurrentScrollPage] = useState<number>(1);
  const [isDirty, setIsDirty] = useState<boolean>(false);
  const [editRevision, setEditRevision] = useState<number>(0);
  const [saveFeedback, setSaveFeedback] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [isDownloadingOriginal, setIsDownloadingOriginal] =
    useState<boolean>(false);
  const [exportNotice, setExportNotice] = useState<string | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);
  const [openPopover, setOpenPopover] = useState<ActivePopover>(null);
  const [customTextColor, setCustomTextColor] = useState<string>('#2563EB');
  const [currentHighlightColor, setCurrentHighlightColor] =
    useState<string>('#FEF08A');
  const [underlineStyle, setUnderlineStyle] =
    useState<WordUnderlineStyle>('solid');
  const [formatState, setFormatState] =
    useState<FormatActiveState>(INITIAL_FORMAT_STATE);
  const [wordCount, setWordCount] = useState<number>(0);
  const [charCount, setCharCount] = useState<number>(0);
  const [detectedPageCount, setDetectedPageCount] = useState<number>(1);
  const [isRenderingDocx, setIsRenderingDocx] = useState<boolean>(false);
  const [isDocxPreviewRendered, setIsDocxPreviewRendered] =
    useState<boolean>(false);
  const [docxRenderFailed, setDocxRenderFailed] = useState<boolean>(false);

  const editorRef = useRef<HTMLDivElement | null>(null);
  const pageSurfaceRef = useRef<HTMLDivElement | null>(null);
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const ribbonRef = useRef<HTMLDivElement | null>(null);
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const painterStyleRef = useRef<{
    fontFamily: string;
    fontSize: string;
    fontWeight: string;
    fontStyle: string;
    textDecorationLine: string;
    color: string;
    backgroundColor: string;
  } | null>(null);
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
  const exportNoticeTimerRef = useRef<number | null>(null);
  const hasLocalEditsSessionRef = useRef<boolean>(
    Boolean(document.hasLocalEdits)
  );
  const isReadOnlyRef = useRef<boolean>(isReadOnly);
  isReadOnlyRef.current = isReadOnly;
  const documentRef = useRef<DocumentItem>(document);
  documentRef.current = document;

  const fontSize: FontSizeOption = FONT_SIZES[fontSizeIndex] ?? 16;
  const themeStyle = PAPER_THEMES[paperTheme];
  const isDark = paperTheme === 'dark';

  const pageDimensions = useMemo(
    () => getWordPageDimensions(paperSize, pageOrientation),
    [paperSize, pageOrientation]
  );

  useDocumentPinchZoom(scrollContainerRef, zoomPercent, setZoomPercent, {
    min: 25,
    max: 200,
    disabled: isModalOpen || activeToolModal !== null,
  });

  const basePaperWidthPx = isDocxPreviewRendered
    ? Math.max(820, pageDimensions.widthPx)
    : pageDimensions.widthPx;

  const handleFitToWidth = useCallback(() => {
    if (containerWidth <= 0) {
      setZoomPercent(100);
      return;
    }
    const available = Math.max(220, containerWidth - 36);
    const pct = Math.max(
      25,
      Math.min(200, Math.round((available / basePaperWidthPx) * 100))
    );
    setZoomPercent(pct);
  }, [basePaperWidthPx, containerWidth]);

  const handleTogglePageLayout = useCallback(() => {
    setPageLayout((prev) => {
      const next: WordPageLayout = prev === 'print' ? 'mobile' : 'print';
      saveStoredPageLayout(next);
      return next;
    });
  }, []);

  const handleSelectMarginPreset = useCallback((preset: WordMarginPreset) => {
    setMarginPreset(preset);
    if (preset === 'normal') {
      setMarginLeftCm(2.0);
      setMarginRightCm(2.0);
      setMarginTopCm(2.0);
      setMarginBottomCm(2.0);
    } else if (preset === 'narrow') {
      setMarginLeftCm(1.2);
      setMarginRightCm(1.2);
      setMarginTopCm(1.2);
      setMarginBottomCm(1.2);
    } else if (preset === 'wide') {
      setMarginLeftCm(3.2);
      setMarginRightCm(3.2);
      setMarginTopCm(2.5);
      setMarginBottomCm(2.5);
    }
  }, []);

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
    setCharCount(plainText.replace(/\r?\n/g, '').length);

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
      const text = selection.toString().trim();
      if (text) {
        setSelectedQuote(text.slice(0, 180));
      }
      const anchorEl =
        range.commonAncestorContainer instanceof HTMLElement
          ? range.commonAncestorContainer
          : range.commonAncestorContainer.parentElement;
      const cell = anchorEl?.closest('td, th');
      setActiveTableCell(
        cell instanceof HTMLTableCellElement && editorEl.contains(cell)
          ? cell
          : null
      );
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

  // Lock editor to original view whenever read-only permission applies
  useEffect(() => {
    if (isReadOnly) {
      setWordEditMode('original');
    }
  }, [isReadOnly]);

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
      hasLocalEditsSessionRef.current = Boolean(document.hasLocalEdits);
      setWordEditMode(
        resolveInitialWordEditMode(document, isReadOnlyRef.current)
      );
      setIsDirty(false);
      setEditRevision(0);
      setSaveFeedback(false);
      setExportNotice(null);
      setExportError(null);
      setOpenPopover(null);
      setIsMobileDrawerOpen(false);
      setIsFindOpen(false);
      setIsFormatPainterActive(false);
      setActiveTableCell(null);
      setSelectedImageEl(null);
      setWordComments(document.wordComments ?? []);
      setHeaderText(document.title);
      setDocxRenderFailed(false);

      if (scrollContainerRef.current) {
        scrollContainerRef.current.scrollTop = 0;
      }
    } else if (
      hasNewRawBuffer &&
      !isDirty &&
      !document.hasLocalEdits &&
      !hasLocalEditsSessionRef.current
    ) {
      setWordEditMode('original');
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
      if (exportNoticeTimerRef.current !== null) {
        window.clearTimeout(exportNoticeTimerRef.current);
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

    const didEditLocally = Boolean(
      currentDoc.hasLocalEdits ||
        hasLocalEditsSessionRef.current ||
        isDirty ||
        editRevision > 0
    );

    return {
      ...currentDoc,
      wordHtml: htmlWithIds,
      outline: liveOutline,
      pageCount: detectedPageCount,
      contentPreview: trimmedPreview || currentDoc.contentPreview,
      wordComments: wordComments.length > 0 ? wordComments : undefined,
      hasLocalEdits: didEditLocally,
    };
  }, [detectedPageCount, editRevision, isDirty, outlineEntries, wordComments]);

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

  // Keyboard shortcuts: Escape, Ctrl/Cmd+S (save), Ctrl/Cmd+F / Ctrl/Cmd+H (find & replace)
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

      if (
        (event.ctrlKey || event.metaKey) &&
        (event.key.toLowerCase() === 'f' || event.key.toLowerCase() === 'h')
      ) {
        event.preventDefault();
        setIsFindOpen(true);
        return;
      }

      if (event.key === 'Escape') {
        if (activeToolModal !== null) {
          setActiveToolModal(null);
          return;
        }
        if (openPopover !== null) {
          setOpenPopover(null);
          return;
        }
        if (isFormatPainterActive) {
          setIsFormatPainterActive(false);
          return;
        }
        if (isFindOpen) {
          setIsFindOpen(false);
          return;
        }
        if (isReadMode) {
          setIsReadMode(false);
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
    activeToolModal,
    isFindOpen,
    isFormatPainterActive,
    isModalOpen,
    isMobileDrawerOpen,
    isReadMode,
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
    if (!editorRef.current || isApplyingRemotePatchRef.current || isEditingLocked) {
      return;
    }
    const nextCleanHtml = stripInjectedStylesFromHtml(
      editorRef.current,
      latestHtmlRef.current
    );
    latestHtmlRef.current = nextCleanHtml;
    hasLocalEditsSessionRef.current = true;
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
  }, [
    flushBroadcastHtml,
    isEditingLocked,
    syncOutlineFromDom,
    refreshFormatState,
  ]);

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
      if (isEditingLocked) {
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
    [
      isEditingLocked,
      restoreEditorSelection,
      captureCurrentSelection,
      handleEditorMutation,
    ]
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

  const handleApplyUnderlineStyle = (style: WordUnderlineStyle) => {
    setUnderlineStyle(style);
    setOpenPopover(null);
    if (isEditingLocked) return;
    restoreEditorSelection();
    const selection = window.document.getSelection();
    if (selection && !selection.isCollapsed && selection.rangeCount > 0) {
      const span = window.document.createElement('span');
      span.style.textDecorationLine = 'underline';
      span.style.textDecorationStyle = style;
      const range = selection.getRangeAt(0);
      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        selection.selectAllChildren(span);
      } catch {
        executeFormatCommand('underline');
      }
      captureCurrentSelection();
      handleEditorMutation();
    } else {
      executeFormatCommand('underline');
    }
  };

  const handleTransformCase = (mode: WordCaseTransform) => {
    setOpenPopover(null);
    if (isEditingLocked) return;
    restoreEditorSelection();
    const selection = window.document.getSelection();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
      showExportNotice('Hãy bôi đen đoạn văn bản cần đổi kiểu chữ HOA/thường.');
      return;
    }
    const text = selection.toString();
    const transformed = applyCaseTransformToText(text, mode);
    executeFormatCommand('insertText', transformed);
  };

  const handleToggleFormatPainter = () => {
    if (isEditingLocked) return;
    if (isFormatPainterActive) {
      setIsFormatPainterActive(false);
      painterStyleRef.current = null;
      return;
    }
    const selection = window.document.getSelection();
    const editorEl = editorRef.current;
    if (!editorEl || !selection || selection.rangeCount === 0) {
      showExportNotice('Đặt con trỏ vào đoạn chữ mẫu trước khi bấm Chổi định dạng.');
      return;
    }
    const node = selection.anchorNode;
    const el =
      node instanceof HTMLElement
        ? node
        : node?.parentElement && editorEl.contains(node.parentElement)
          ? node.parentElement
          : editorEl;
    const computed = window.getComputedStyle(el);
    painterStyleRef.current = {
      fontFamily: computed.fontFamily,
      fontSize: computed.fontSize,
      fontWeight: computed.fontWeight,
      fontStyle: computed.fontStyle,
      textDecorationLine: computed.textDecorationLine,
      color: computed.color,
      backgroundColor: computed.backgroundColor,
    };
    setIsFormatPainterActive(true);
  };

  const applyFormatPainterToSelection = useCallback(() => {
    if (!isFormatPainterActive || !painterStyleRef.current || isEditingLocked) {
      return;
    }
    const selection = window.document.getSelection();
    const editorEl = editorRef.current;
    if (
      !editorEl ||
      !selection ||
      selection.isCollapsed ||
      selection.rangeCount === 0
    ) {
      return;
    }
    const range = selection.getRangeAt(0);
    if (!editorEl.contains(range.commonAncestorContainer)) {
      return;
    }
    const styles = painterStyleRef.current;
    const span = window.document.createElement('span');
    span.style.fontFamily = styles.fontFamily;
    span.style.fontSize = styles.fontSize;
    span.style.fontWeight = styles.fontWeight;
    span.style.fontStyle = styles.fontStyle;
    span.style.textDecorationLine = styles.textDecorationLine;
    span.style.color = styles.color;
    if (
      styles.backgroundColor &&
      styles.backgroundColor !== 'rgba(0, 0, 0, 0)' &&
      styles.backgroundColor !== 'transparent'
    ) {
      span.style.backgroundColor = styles.backgroundColor;
    }
    try {
      span.appendChild(range.extractContents());
      range.insertNode(span);
      selection.removeAllRanges();
      const nextRange = window.document.createRange();
      nextRange.selectNodeContents(span);
      selection.addRange(nextRange);
      captureCurrentSelection();
      handleEditorMutation();
    } finally {
      setIsFormatPainterActive(false);
      painterStyleRef.current = null;
    }
  }, [
    captureCurrentSelection,
    handleEditorMutation,
    isEditingLocked,
    isFormatPainterActive,
  ]);

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

  const handleImageUploadChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || isEditingLocked) return;
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = typeof reader.result === 'string' ? reader.result : '';
      if (!dataUrl) return;
      const safeAlt = escapeHtml(file.name.replace(/\.[^.]+$/, ''));
      executeFormatCommand(
        'insertHTML',
        `<p style="text-align:center;"><img src="${dataUrl}" alt="${safeAlt}" style="max-width:75%;height:auto;border-radius:8px;display:inline-block;margin:10px auto;" /></p><p><br></p>`
      );
    };
    reader.readAsDataURL(file);
  };

  const handleUpdateSelectedImage = (
    patch: { align?: 'left' | 'center' | 'right'; widthPct?: number }
  ) => {
    if (!selectedImageEl || isEditingLocked) return;
    if (patch.widthPct !== undefined) {
      selectedImageEl.style.width = `${patch.widthPct}%`;
      selectedImageEl.style.maxWidth = '100%';
    }
    if (patch.align) {
      selectedImageEl.style.display = 'block';
      if (patch.align === 'left') {
        selectedImageEl.style.margin = '10px auto 10px 0';
      } else if (patch.align === 'right') {
        selectedImageEl.style.margin = '10px 0 10px auto';
      } else {
        selectedImageEl.style.margin = '10px auto';
      }
    }
    handleEditorMutation();
  };

  const handleTableAction = (
    action:
      | 'addRow'
      | 'addCol'
      | 'delRow'
      | 'delCol'
      | 'mergeRight'
      | 'splitCell'
      | 'toggleBorders'
      | 'fill'
      | 'deleteTable',
    fillColor?: string
  ) => {
    const cell = activeTableCell;
    if (!cell || isEditingLocked) return;
    const row = cell.closest('tr');
    const table = cell.closest('table');
    if (!row || !table) return;

    if (action === 'fill' && fillColor !== undefined) {
      cell.style.backgroundColor = fillColor === 'transparent' ? '' : fillColor;
      handleEditorMutation();
      return;
    }

    const cellIdx = Array.from(row.cells).indexOf(cell);

    switch (action) {
      case 'addRow': {
        const newRow = row.cloneNode(true) as HTMLTableRowElement;
        Array.from(newRow.cells).forEach((c) => {
          c.innerHTML = '&nbsp;';
        });
        row.parentNode?.insertBefore(newRow, row.nextSibling);
        break;
      }
      case 'addCol': {
        Array.from(table.rows).forEach((r, rIdx) => {
          const tag = rIdx === 0 && r.querySelector('th') ? 'th' : 'td';
          const newCell = window.document.createElement(tag);
          newCell.innerHTML = tag === 'th' ? 'Cột mới' : '&nbsp;';
          const refCell = r.cells[cellIdx + 1] ?? null;
          r.insertBefore(newCell, refCell);
        });
        break;
      }
      case 'delRow': {
        if (table.rows.length <= 1) {
          table.remove();
          setActiveTableCell(null);
        } else {
          row.remove();
          setActiveTableCell(null);
        }
        break;
      }
      case 'delCol': {
        if (row.cells.length <= 1) {
          table.remove();
          setActiveTableCell(null);
        } else {
          Array.from(table.rows).forEach((r) => {
            const target = r.cells[cellIdx];
            target?.remove();
          });
          setActiveTableCell(null);
        }
        break;
      }
      case 'mergeRight': {
        const nextCell = row.cells[cellIdx + 1];
        if (nextCell) {
          const curSpan = Number(cell.colSpan) || 1;
          const nextSpan = Number(nextCell.colSpan) || 1;
          cell.colSpan = curSpan + nextSpan;
          const nextText = (nextCell.innerHTML ?? '').trim();
          if (nextText && nextText !== '&nbsp;') {
            cell.innerHTML = `${cell.innerHTML} ${nextText}`;
          }
          nextCell.remove();
        }
        break;
      }
      case 'splitCell': {
        const curSpan = Number(cell.colSpan) || 1;
        if (curSpan > 1) {
          cell.colSpan = curSpan - 1;
          const extra = window.document.createElement(cell.tagName.toLowerCase());
          extra.innerHTML = '&nbsp;';
          row.insertBefore(extra, cell.nextSibling);
        }
        break;
      }
      case 'toggleBorders': {
        const currentlyHidden = table.dataset.noBorders === 'true';
        table.dataset.noBorders = currentlyHidden ? 'false' : 'true';
        Array.from(table.querySelectorAll<HTMLElement>('td, th')).forEach(
          (c) => {
            c.style.border = currentlyHidden ? '' : '1px dashed transparent';
          }
        );
        break;
      }
      case 'deleteTable': {
        table.remove();
        setActiveTableCell(null);
        break;
      }
    }

    handleEditorMutation();
  };

  const clearFindMarks = useCallback(() => {
    const editorEl = editorRef.current;
    if (!editorEl) return;
    const marks = Array.from(
      editorEl.querySelectorAll<HTMLElement>('mark[data-word-find="true"]')
    );
    marks.forEach((mark) => {
      const parent = mark.parentNode;
      if (!parent) return;
      while (mark.firstChild) {
        parent.insertBefore(mark.firstChild, mark);
      }
      parent.removeChild(mark);
      parent.normalize();
    });
  }, []);

  const runFindHighlights = useCallback(
    (query: string, caseSensitive: boolean, focusIdx = 0) => {
      clearFindMarks();
      const editorEl = editorRef.current;
      const trimmed = query;
      if (!editorEl || !trimmed) {
        setFindMatchCount(0);
        setActiveFindMatchIndex(0);
        return;
      }

      const walker = window.document.createTreeWalker(
        editorEl,
        NodeFilter.SHOW_TEXT,
        {
          acceptNode: (node) => {
            const parent = node.parentElement;
            if (!parent || parent.closest('style, script')) {
              return NodeFilter.FILTER_REJECT;
            }
            return node.nodeValue && node.nodeValue.length > 0
              ? NodeFilter.FILTER_ACCEPT
              : NodeFilter.FILTER_REJECT;
          },
        }
      );

      const textNodes: Text[] = [];
      let current = walker.nextNode();
      while (current) {
        if (current instanceof Text) {
          textNodes.push(current);
        }
        current = walker.nextNode();
      }

      const needle = caseSensitive ? trimmed : trimmed.toLocaleLowerCase('vi-VN');
      const createdMarks: HTMLElement[] = [];

      for (const textNode of textNodes) {
        const raw = textNode.nodeValue ?? '';
        const haystack = caseSensitive ? raw : raw.toLocaleLowerCase('vi-VN');
        if (!haystack.includes(needle)) continue;

        const frag = window.document.createDocumentFragment();
        let cursor = 0;
        let idx = haystack.indexOf(needle, cursor);
        while (idx !== -1) {
          if (idx > cursor) {
            frag.appendChild(window.document.createTextNode(raw.slice(cursor, idx)));
          }
          const mark = window.document.createElement('mark');
          mark.dataset.wordFind = 'true';
          mark.style.backgroundColor = '#FDE047';
          mark.style.color = '#0F172A';
          mark.style.borderRadius = '2px';
          mark.textContent = raw.slice(idx, idx + needle.length);
          frag.appendChild(mark);
          createdMarks.push(mark);
          cursor = idx + needle.length;
          idx = haystack.indexOf(needle, cursor);
        }
        if (cursor < raw.length) {
          frag.appendChild(window.document.createTextNode(raw.slice(cursor)));
        }
        textNode.parentNode?.replaceChild(frag, textNode);
      }

      setFindMatchCount(createdMarks.length);
      if (createdMarks.length > 0) {
        const clamped =
          ((focusIdx % createdMarks.length) + createdMarks.length) %
          createdMarks.length;
        setActiveFindMatchIndex(clamped);
        createdMarks.forEach((m, i) => {
          m.style.outline = i === clamped ? '2px solid #185ABD' : 'none';
          m.style.backgroundColor = i === clamped ? '#FB923C' : '#FDE047';
        });
        createdMarks[clamped].scrollIntoView({
          behavior: 'smooth',
          block: 'center',
        });
      } else {
        setActiveFindMatchIndex(0);
      }
    },
    [clearFindMarks]
  );

  useEffect(() => {
    if (!isFindOpen) {
      clearFindMarks();
      setFindMatchCount(0);
      setActiveFindMatchIndex(0);
      return;
    }
    runFindHighlights(findQuery, findCaseSensitive, 0);
  }, [clearFindMarks, findCaseSensitive, findQuery, isFindOpen, runFindHighlights]);

  const handleFindNavigate = (delta: -1 | 1) => {
    if (findMatchCount === 0) return;
    const nextIdx =
      (activeFindMatchIndex + delta + findMatchCount) % findMatchCount;
    runFindHighlights(findQuery, findCaseSensitive, nextIdx);
  };

  const handleReplaceOne = () => {
    if (isEditingLocked || findMatchCount === 0) return;
    const editorEl = editorRef.current;
    if (!editorEl) return;
    const marks = Array.from(
      editorEl.querySelectorAll<HTMLElement>('mark[data-word-find="true"]')
    );
    const target = marks[activeFindMatchIndex];
    if (!target || !target.parentNode) return;
    target.parentNode.replaceChild(
      window.document.createTextNode(replaceText),
      target
    );
    handleEditorMutation();
    runFindHighlights(findQuery, findCaseSensitive, activeFindMatchIndex);
  };

  const handleReplaceAll = () => {
    if (isEditingLocked || findMatchCount === 0) return;
    const editorEl = editorRef.current;
    if (!editorEl) return;
    const marks = Array.from(
      editorEl.querySelectorAll<HTMLElement>('mark[data-word-find="true"]')
    );
    marks.forEach((mark) => {
      mark.parentNode?.replaceChild(
        window.document.createTextNode(replaceText),
        mark
      );
    });
    handleEditorMutation();
    setFindMatchCount(0);
    setActiveFindMatchIndex(0);
  };

  const handleAddComment = (text: string) => {
    const newComment: WordCommentItem = {
      id: `cmt-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      quote: selectedQuote,
      text,
      author: 'Bạn',
      createdAt: new Date().toLocaleTimeString('vi-VN', {
        hour: '2-digit',
        minute: '2-digit',
      }),
    };
    setWordComments((prev) => [newComment, ...prev]);
    setIsDirty(true);
    setEditRevision((prev) => prev + 1);
  };

  const handleDeleteComment = (id: string) => {
    setWordComments((prev) => prev.filter((c) => c.id !== id));
    setIsDirty(true);
    setEditRevision((prev) => prev + 1);
  };

  const handleJumpToQuote = (quote: string) => {
    if (!quote) return;
    setIsFindOpen(true);
    setFindQuery(quote);
  };

  const handlePreventMouseDown = (event: ReactMouseEvent<HTMLButtonElement>) => {
    event.preventDefault();
  };

  const showExportError = (message: string) => {
    setExportError(message);
    if (exportErrorTimerRef.current !== null) {
      window.clearTimeout(exportErrorTimerRef.current);
    }
    exportErrorTimerRef.current = window.setTimeout(() => {
      setExportError(null);
      exportErrorTimerRef.current = null;
    }, 4000);
  };

  const showExportNotice = (message: string) => {
    setExportNotice(message);
    if (exportNoticeTimerRef.current !== null) {
      window.clearTimeout(exportNoticeTimerRef.current);
    }
    exportNoticeTimerRef.current = window.setTimeout(() => {
      setExportNotice(null);
      exportNoticeTimerRef.current = null;
    }, 5000);
  };

  const handleSwitchWordMode = (nextMode: WordEditMode) => {
    if (nextMode === 'edit' && isReadOnly) {
      return;
    }
    setWordEditMode(nextMode);
    if (nextMode === 'original') {
      setOpenPopover(null);
      setIsFormatPainterActive(false);
    }
  };

  const handleDownloadOriginal = async () => {
    if (isDownloadingOriginal || isExporting) {
      return;
    }
    setIsDownloadingOriginal(true);
    setExportError(null);
    setExportNotice(null);
    try {
      if (onDownloadOriginalFile) {
        await onDownloadOriginalFile();
        showExportNotice(
          'Tệp nhị phân gốc ban đầu đã được xử lý tải xuống thiết bị.'
        );
      } else {
        const result = await exportOriginalDocumentBinary(document);
        showExportNotice(result.message);
      }
    } catch (err: unknown) {
      showExportError(
        err instanceof Error
          ? err.message
          : 'Không thể tải xuống bản gốc. Vui lòng thử lại.'
      );
    } finally {
      setIsDownloadingOriginal(false);
    }
  };

  const handleExportDocx = async () => {
    if (isExporting || isDownloadingOriginal) {
      return;
    }
    setIsExporting(true);
    setExportError(null);
    setExportNotice(null);
    try {
      const currentHtml = stripInjectedStylesFromHtml(
        editorRef.current,
        latestHtmlRef.current
      );
      const result = await exportWordToDocx(document.title, currentHtml);
      showExportNotice(result.message);
    } catch {
      showExportError('Không thể xuất tệp .docx. Vui lòng thử lại.');
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

    const maxScroll = Math.max(1, container.scrollHeight - container.clientHeight);
    const scrollRatio = Math.max(0, Math.min(1, container.scrollTop / maxScroll));
    const estPage = Math.max(
      1,
      Math.min(
        Math.max(1, detectedPageCount),
        Math.round(scrollRatio * Math.max(0, detectedPageCount - 1)) + 1
      )
    );
    setCurrentScrollPage((prev) => (prev === estPage ? prev : estPage));
  };

  const docStats = useMemo(
    () => computeWordDocumentStats(editorRef.current, detectedPageCount),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [detectedPageCount, editRevision, wordCount, isDocxPreviewRendered]
  );

  const currentBlockLabel =
    BLOCK_STYLE_OPTIONS.find((item) => item.tag === formatState.block)?.label ??
    'Văn bản thường';

  const toolbarBtnClass = (active: boolean) =>
    `inline-flex items-center justify-center min-h-[30px] min-w-[30px] px-2 rounded-lg text-xs font-medium transition-colors shrink-0 ${
      active
        ? isDark
          ? 'bg-[#185ABD]/35 text-blue-100 ring-1 ring-blue-400/50'
          : 'bg-[#185ABD]/15 text-[#185ABD] ring-1 ring-[#185ABD]/35'
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
  const zoomScale = zoomPercent / 100;
  const pageFontSizePx =
    pageLayout === 'mobile'
      ? Number(((fontSize * zoomPercent) / 100).toFixed(1))
      : fontSize;

  const hostCssVars: CSSProperties = {
    '--lite-doc-page-bg': themeStyle.pageHexBg,
    '--lite-doc-page-text': themeStyle.pageHexText,
    '--lite-doc-font-family': fontStack,
    '--lite-doc-line-height': String(lineHeight),
    '--lite-doc-para-spacing': `${paragraphSpacingPx}px`,
    '--lite-doc-first-indent': `${cmToPx(firstLineIndentCm)}px`,
    fontFamily: fontStack,
    fontSize: `${pageFontSizePx}px`,
    lineHeight,
    ...(pageLayout === 'print'
      ? { width: `${basePaperWidthPx}px`, maxWidth: `${basePaperWidthPx}px` }
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
        disabled={isEditingLocked || isRenderingDocx || isModalOpen}
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

      {/* Shared Top Workspace Header (hidden in distraction-free Read Mode) */}
      {!isReadMode && (
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
            {/* Mode Switcher: Xem nguyên bản <-> Chỉnh sửa nhanh */}
            <div
              role="group"
              aria-label="Chế độ xem và chỉnh sửa Word"
              className={`shrink-0 inline-flex items-center p-0.5 rounded-xl border ${
                isDark
                  ? 'bg-white/[0.04] border-white/10'
                  : 'bg-[#F5F2EC] border-[#E5E0D8]'
              }`}
            >
              <button
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => handleSwitchWordMode('original')}
                aria-pressed={wordEditMode === 'original'}
                title="Xem nguyên bản (khóa sửa nhầm, giữ chuẩn định dạng gốc)"
                className={`inline-flex items-center gap-1 min-h-[30px] px-2 sm:px-2.5 rounded-[10px] text-[11px] sm:text-xs font-semibold transition-all whitespace-nowrap ${
                  wordEditMode === 'original'
                    ? isDark
                      ? 'bg-indigo-500/25 text-indigo-100 shadow-sm ring-1 ring-indigo-400/40'
                      : 'bg-white text-indigo-700 shadow-sm ring-1 ring-indigo-200/80'
                    : isDark
                      ? 'text-slate-300 hover:text-white'
                      : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Eye className="h-3.5 w-3.5 shrink-0" />
                <span>Xem nguyên bản</span>
              </button>
              <button
                type="button"
                disabled={isReadOnly}
                onMouseDown={handlePreventMouseDown}
                onClick={() => handleSwitchWordMode('edit')}
                aria-pressed={wordEditMode === 'edit'}
                title={
                  isReadOnly
                    ? 'Tài liệu đang ở chế độ chỉ xem'
                    : 'Chuyển sang Chỉnh sửa nhanh'
                }
                className={`inline-flex items-center gap-1 min-h-[30px] px-2 sm:px-2.5 rounded-[10px] text-[11px] sm:text-xs font-semibold transition-all whitespace-nowrap disabled:opacity-45 disabled:cursor-not-allowed ${
                  wordEditMode === 'edit' && !isReadOnly
                    ? isDark
                      ? 'bg-indigo-500/25 text-indigo-100 shadow-sm ring-1 ring-indigo-400/40'
                      : 'bg-white text-indigo-700 shadow-sm ring-1 ring-indigo-200/80'
                    : isDark
                      ? 'text-slate-300 hover:text-white'
                      : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Pencil className="h-3.5 w-3.5 shrink-0" />
                <span>Chỉnh sửa nhanh</span>
              </button>
            </div>

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

            {/* Download 100% Original File Button */}
            <button
              type="button"
              onClick={() => void handleDownloadOriginal()}
              disabled={isDownloadingOriginal || isExporting}
              aria-label="Tải xuống bản gốc"
              title="Tải xuống bản gốc ban đầu"
              className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-2.5 sm:px-3 rounded-xl text-xs font-semibold border transition-colors whitespace-nowrap disabled:opacity-60 disabled:cursor-wait ${
                isDark
                  ? 'bg-sky-500/15 hover:bg-sky-500/25 border-sky-400/35 text-sky-200'
                  : 'bg-sky-50 hover:bg-sky-100 border-sky-200/80 text-sky-700'
              }`}
            >
              {isDownloadingOriginal ? (
                <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
              ) : (
                <FileDown className="h-3.5 w-3.5 shrink-0" />
              )}
              <span className="hidden sm:inline">
                {isDownloadingOriginal ? 'Đang tải...' : 'Tải xuống bản gốc'}
              </span>
            </button>

            {/* Export Real .docx Button */}
            <button
              type="button"
              onClick={() => void handleExportDocx()}
              disabled={isExporting || isDownloadingOriginal}
              aria-label="Tải xuống tệp Word .docx"
              title="Tải xuống tệp .docx (bản chỉnh sửa)"
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
      )}

      {/* Hidden local image input for Insert -> Image */}
      <input
        ref={imageInputRef}
        type="file"
        accept="image/*"
        onChange={handleImageUploadChange}
        className="hidden"
      />

      {/* Unified Horizontal-Scrollable Microsoft Word Ribbon (#185ABD) */}
      {!isReadMode && (
        <div
          ref={ribbonRef}
          className={`shrink-0 border-b px-2.5 sm:px-4 py-1.5 z-20 transition-colors duration-200 relative ${themeStyle.ribbonBg} ${themeStyle.ribbonBorder}`}
        >
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-0.5 min-w-0">
            {/* 1. Primary 4-Tab Ribbon Switcher */}
            <div
              role="tablist"
              aria-label="Chọn phần công cụ Word"
              className={`inline-flex items-center gap-0.5 p-0.5 rounded-xl border shrink-0 ${
                isDark
                  ? 'bg-[#1E2534] border-white/10'
                  : 'bg-[#F1F5FA] border-[#D8E2F0]'
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
                        ? 'bg-[#185ABD] text-white shadow-sm'
                        : isDark
                          ? 'text-slate-300 hover:text-white'
                          : 'text-slate-700 hover:text-[#185ABD]'
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

            {isReadOnly ? (
              <span
                className={`shrink-0 inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-semibold border ${
                  isDark
                    ? 'bg-amber-500/15 border-amber-400/30 text-amber-200'
                    : 'bg-amber-50 border-amber-200 text-amber-800'
                }`}
              >
                Chế độ chỉ xem
              </span>
            ) : wordEditMode === 'original' ? (
              <div
                className={`shrink-0 inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-semibold border ${
                  isDark
                    ? 'bg-blue-500/15 border-blue-400/30 text-blue-200'
                    : 'bg-blue-50 border-blue-200 text-[#185ABD]'
                }`}
              >
                <Lock className="h-3 w-3 shrink-0" />
                <span>Xem nguyên bản (khóa sửa)</span>
                <button
                  type="button"
                  onClick={() => handleSwitchWordMode('edit')}
                  className="underline underline-offset-2 font-bold hover:opacity-80"
                >
                  Chỉnh sửa nhanh
                </button>
              </div>
            ) : null}

            {/* TAB 1: TRANG ĐẦU (HOME) */}
            {activeRibbonTab === 'home' && (
              <fieldset
                disabled={isEditingLocked}
                className={`flex items-center gap-1 min-w-0 border-0 p-0 m-0 shrink-0 ${
                  isEditingLocked ? 'opacity-45 pointer-events-none' : ''
                }`}
              >
                {/* Undo / Redo */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('undo')}
                  aria-label="Hoàn tác"
                  title="Hoàn tác (Ctrl+Z)"
                  className={toolbarBtnClass(false)}
                >
                  <Undo2 className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('redo')}
                  aria-label="Làm lại"
                  title="Làm lại (Ctrl+Y)"
                  className={toolbarBtnClass(false)}
                >
                  <Redo2 className="h-3.5 w-3.5" />
                </button>

                {/* Format Painter */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleToggleFormatPainter}
                  aria-pressed={isFormatPainterActive}
                  aria-label="Sao chép định dạng (Format Painter)"
                  title="Chổi sao chép định dạng (chọn chữ mẫu -> bấm chổi -> bôi chữ đích)"
                  className={toolbarBtnClass(isFormatPainterActive)}
                >
                  <Paintbrush className="h-3.5 w-3.5" />
                </button>

                {/* Block Style */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setOpenPopover((prev) => (prev === 'block' ? null : 'block'))
                  }
                  aria-expanded={openPopover === 'block'}
                  aria-label="Chọn kiểu đoạn hoặc tiêu đề"
                  className={`shrink-0 inline-flex items-center justify-between gap-1 min-h-[30px] px-2 rounded-lg border text-xs font-medium transition-colors whitespace-nowrap ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200 hover:bg-white/10'
                      : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800 hover:bg-[#F2EFEA]'
                  }`}
                >
                  <span>{currentBlockLabel}</span>
                  <ChevronDown className="h-3.5 w-3.5 opacity-70" />
                </button>

                {/* Font Family Selector */}
                <select
                  aria-label="Phông chữ"
                  value={fontFamily}
                  onChange={(e) =>
                    handleSelectFontFamily(e.target.value as WordFontFamily)
                  }
                  className={`h-[30px] px-2 rounded-lg border text-xs font-medium outline-none shrink-0 ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200'
                      : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800'
                  }`}
                >
                  {FONT_FAMILIES.map((family) => (
                    <option key={family} value={family}>
                      {family}
                    </option>
                  ))}
                </select>

                {/* Font Size Selector + A^ / Av */}
                <div
                  role="group"
                  aria-label="Cỡ chữ"
                  className={`inline-flex items-center rounded-lg border p-0.5 shrink-0 ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10'
                      : 'bg-[#F5F3EE] border-[#E5E0D8]'
                  }`}
                >
                  <select
                    aria-label="Chọn cỡ chữ"
                    value={fontSizeIndex}
                    onChange={(e) => setFontSizeIndex(Number(e.target.value))}
                    className={`h-[24px] px-1 rounded text-[11px] font-semibold tabular-nums bg-transparent outline-none ${
                      isDark ? 'text-slate-200' : 'text-slate-800'
                    }`}
                  >
                    {FONT_SIZES.map((sz, idx) => (
                      <option
                        key={sz}
                        value={idx}
                        className="text-slate-900 bg-white"
                      >
                        {sz}
                      </option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleIncreaseFont}
                    disabled={fontSizeIndex === FONT_SIZES.length - 1}
                    title="Tăng cỡ chữ (A^)"
                    className="min-h-[24px] px-1.5 rounded text-[11px] font-bold hover:bg-black/5 disabled:opacity-35"
                  >
                    A^
                  </button>
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={handleDecreaseFont}
                    disabled={fontSizeIndex === 0}
                    title="Giảm cỡ chữ (Av)"
                    className="min-h-[24px] px-1.5 rounded text-[11px] font-bold hover:bg-black/5 disabled:opacity-35"
                  >
                    Av
                  </button>
                </div>

                {/* B / I / U (+ style) / S / x2 / H2O / Aa */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('bold')}
                  aria-pressed={formatState.bold}
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
                  title="In nghiêng (Ctrl+I)"
                  className={toolbarBtnClass(formatState.italic)}
                >
                  <Italic className="h-3.5 w-3.5" />
                </button>
                <div className="inline-flex items-center shrink-0">
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() => handleApplyUnderlineStyle(underlineStyle)}
                    aria-pressed={formatState.underline}
                    title="Gạch chân (Ctrl+U)"
                    className={toolbarBtnClass(formatState.underline)}
                  >
                    <Underline className="h-3.5 w-3.5" />
                  </button>
                  <button
                    type="button"
                    onMouseDown={handlePreventMouseDown}
                    onClick={() =>
                      setOpenPopover((prev) =>
                        prev === 'underline' ? null : 'underline'
                      )
                    }
                    aria-label="Kiểu gạch chân"
                    title="Chọn kiểu gạch chân"
                    className="h-[30px] px-1 rounded hover:bg-black/5"
                  >
                    <ChevronDown className="h-3 w-3 opacity-75" />
                  </button>
                </div>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('strikeThrough')}
                  aria-pressed={formatState.strikeThrough}
                  title="Gạch ngang chữ"
                  className={toolbarBtnClass(formatState.strikeThrough)}
                >
                  <Strikethrough className="h-3.5 w-3.5" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('superscript')}
                  title="Chỉ số trên (x²)"
                  aria-label="Chỉ số trên"
                  className={toolbarBtnClass(false)}
                >
                  <Superscript className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('subscript')}
                  title="Chỉ số dưới (H₂O)"
                  aria-label="Chỉ số dưới"
                  className={toolbarBtnClass(false)}
                >
                  <Subscript className="h-3.5 w-3.5" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setOpenPopover((prev) => (prev === 'case' ? null : 'case'))
                  }
                  aria-expanded={openPopover === 'case'}
                  title="Đổi chữ HOA / thường (Aa)"
                  aria-label="Đổi chữ HOA hoặc thường"
                  className={toolbarBtnClass(openPopover === 'case')}
                >
                  <CaseSensitive className="h-3.5 w-3.5" />
                </button>

                {/* Color & Highlight & Clear */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setOpenPopover((prev) => (prev === 'color' ? null : 'color'))
                  }
                  aria-expanded={openPopover === 'color'}
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
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setOpenPopover((prev) =>
                      prev === 'highlight' ? null : 'highlight'
                    )
                  }
                  aria-expanded={openPopover === 'highlight'}
                  title="Tô sáng (Highlight)"
                  className={toolbarBtnClass(openPopover === 'highlight')}
                >
                  <Highlighter className="h-3.5 w-3.5 text-amber-500" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('removeFormat')}
                  title="Xóa định dạng"
                  aria-label="Xóa định dạng"
                  className={toolbarBtnClass(false)}
                >
                  <RemoveFormatting className="h-3.5 w-3.5" />
                </button>

                <span
                  aria-hidden="true"
                  className={`h-5 w-px mx-0.5 shrink-0 ${
                    isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
                  }`}
                />

                {/* Alignment + Lists + Indent */}
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('justifyLeft')}
                  aria-pressed={formatState.align === 'left'}
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
                  title="Căn đều"
                  className={toolbarBtnClass(formatState.align === 'justify')}
                >
                  <AlignJustify className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('insertUnorderedList')}
                  aria-pressed={formatState.unorderedList}
                  title="Danh sách dấu chấm"
                  className={toolbarBtnClass(formatState.unorderedList)}
                >
                  <List className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('insertOrderedList')}
                  aria-pressed={formatState.orderedList}
                  title="Danh sách đánh số"
                  className={toolbarBtnClass(formatState.orderedList)}
                >
                  <ListOrdered className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('outdent')}
                  title="Giảm thụt lề"
                  className={toolbarBtnClass(false)}
                >
                  <IndentDecrease className="h-3.5 w-3.5" />
                </button>
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => executeFormatCommand('indent')}
                  title="Tăng thụt lề"
                  className={toolbarBtnClass(false)}
                >
                  <IndentIncrease className="h-3.5 w-3.5" />
                </button>
              </fieldset>
            )}

            {/* TAB 2: CHÈN (INSERT) */}
            {activeRibbonTab === 'insert' && (
              <fieldset
                disabled={isEditingLocked}
                className={`flex items-center gap-1.5 min-w-0 border-0 p-0 m-0 shrink-0 ${
                  isEditingLocked ? 'opacity-45 pointer-events-none' : ''
                }`}
              >
                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() =>
                    setOpenPopover((prev) =>
                      prev === 'table_grid' ? null : 'table_grid'
                    )
                  }
                  className={`shrink-0 inline-flex items-center gap-1 min-h-[30px] px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200 hover:bg-white/10'
                      : 'bg-[#FAF8F5] border-[#E5E0D8] text-slate-800 hover:bg-blue-50'
                  }`}
                >
                  <Table className="h-3.5 w-3.5 text-[#185ABD] shrink-0" />
                  <span>Bảng</span>
                  <ChevronDown className="h-3 w-3 opacity-70" />
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => imageInputRef.current?.click()}
                  className={toolbarBtnClass(false)}
                >
                  <ImageIcon className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Ảnh</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setActiveToolModal('shape_textbox')}
                  className={toolbarBtnClass(false)}
                >
                  <Shapes className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Hình khối &amp; Hộp chữ</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setActiveToolModal('hyperlink')}
                  className={toolbarBtnClass(false)}
                >
                  <Link2 className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Liên kết</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setActiveToolModal('symbol_equation')}
                  className={toolbarBtnClass(false)}
                >
                  <Sigma className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Ký tự &amp; Công thức</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleInsertPageBreak}
                  className={toolbarBtnClass(false)}
                >
                  <SplitSquareVertical className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Ngắt trang</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleInsertHorizontalRule}
                  className={toolbarBtnClass(false)}
                >
                  <Minus className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Đường kẻ ngang</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={handleInsertCalloutBox}
                  className={toolbarBtnClass(false)}
                >
                  <Quote className="h-3.5 w-3.5 text-[#185ABD] mr-1 shrink-0" />
                  <span>Khung trích dẫn</span>
                </button>

                <button
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => setShowHeaderFooter((prev) => !prev)}
                  aria-pressed={showHeaderFooter}
                  className={toolbarBtnClass(showHeaderFooter)}
                >
                  <span>Header / Footer &amp; Số trang</span>
                </button>
              </fieldset>
            )}

            {/* TAB 3: BỐ TRÍ (LAYOUT) */}
            {activeRibbonTab === 'layout' && (
              <div className="flex items-center gap-1.5 shrink-0">
                {/* Paper Theme */}
                <div
                  role="group"
                  aria-label="Chọn màu giấy"
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
                        className={`inline-flex items-center gap-1 min-h-[26px] px-2 rounded-md text-[11px] font-medium whitespace-nowrap ${
                          isSelected
                            ? 'bg-[#185ABD] text-white shadow-sm'
                            : isDark
                              ? 'text-slate-300'
                              : 'text-slate-700'
                        }`}
                      >
                        {mode === 'light' ? (
                          <Sun className="h-3 w-3 text-amber-400 shrink-0" />
                        ) : mode === 'dark' ? (
                          <Moon className="h-3 w-3 shrink-0" />
                        ) : (
                          <span
                            className={`h-2.5 w-2.5 rounded-full border ${config.swatchClass}`}
                          />
                        )}
                        <span>{config.shortLabel}</span>
                      </button>
                    );
                  })}
                </div>

                {/* Paper Size: A4 | A5 | Letter */}
                <select
                  aria-label="Khổ giấy"
                  value={paperSize}
                  onChange={(e) => setPaperSize(e.target.value as WordPaperSize)}
                  className={`h-[30px] px-2 rounded-lg border text-xs font-medium outline-none ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200'
                      : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800'
                  }`}
                >
                  <option value="A4">Khổ A4</option>
                  <option value="A5">Khổ A5</option>
                  <option value="Letter">Letter</option>
                </select>

                {/* Orientation: Dọc / Ngang */}
                <button
                  type="button"
                  onClick={() =>
                    setPageOrientation((prev) =>
                      prev === 'portrait' ? 'landscape' : 'portrait'
                    )
                  }
                  className={toolbarBtnClass(pageOrientation === 'landscape')}
                >
                  <span>
                    {pageOrientation === 'portrait' ? 'Hướng Dọc' : 'Hướng Ngang'}
                  </span>
                </button>

                {/* Margin Presets */}
                <select
                  aria-label="Lề trang"
                  value={marginPreset}
                  onChange={(e) =>
                    handleSelectMarginPreset(e.target.value as WordMarginPreset)
                  }
                  className={`h-[30px] px-2 rounded-lg border text-xs font-medium outline-none ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200'
                      : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800'
                  }`}
                >
                  <option value="normal">Lề Thường (2cm)</option>
                  <option value="narrow">Lề Hẹp (1.2cm)</option>
                  <option value="wide">Lề Rộng (3.2cm)</option>
                  <option value="custom">Lề Tùy chỉnh</option>
                </select>

                {/* Columns 1 / 2 / 3 */}
                <div
                  role="group"
                  aria-label="Chia cột văn bản"
                  className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10'
                      : 'bg-[#F5F3EE] border-[#E5E0D8]'
                  }`}
                >
                  <Columns className="h-3.5 w-3.5 mx-1 text-[#185ABD]" />
                  {([1, 2, 3] as const).map((cols) => (
                    <button
                      key={cols}
                      type="button"
                      onClick={() => setColumnCount(cols)}
                      className={`min-h-[26px] px-2 rounded-md text-[11px] font-semibold ${
                        columnCount === cols
                          ? 'bg-[#185ABD] text-white'
                          : isDark
                            ? 'text-slate-300'
                            : 'text-slate-700'
                      }`}
                    >
                      {cols} cột
                    </button>
                  ))}
                </div>

                {/* Line Height: 1.0 | 1.15 | 1.5 | 2.0 */}
                <div
                  role="group"
                  aria-label="Giãn dòng"
                  className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10'
                      : 'bg-[#F5F3EE] border-[#E5E0D8]'
                  }`}
                >
                  <span className={`px-1.5 text-[11px] ${themeStyle.textMuted}`}>
                    Giãn dòng:
                  </span>
                  {LINE_HEIGHTS.map((lh) => (
                    <button
                      key={lh}
                      type="button"
                      onClick={() => setLineHeight(lh)}
                      className={`min-h-[26px] px-2 rounded-md text-[11px] font-semibold tabular-nums ${
                        lineHeight === lh
                          ? 'bg-[#185ABD] text-white'
                          : isDark
                            ? 'text-slate-300'
                            : 'text-slate-700'
                      }`}
                    >
                      {lh}
                    </button>
                  ))}
                </div>

                {/* Paragraph Spacing */}
                <select
                  aria-label="Khoảng cách đoạn"
                  value={paragraphSpacingPx}
                  onChange={(e) => setParagraphSpacingPx(Number(e.target.value))}
                  className={`h-[30px] px-2 rounded-lg border text-xs font-medium outline-none ${
                    isDark
                      ? 'bg-[#1E2534] border-white/10 text-slate-200'
                      : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800'
                  }`}
                >
                  {PARAGRAPH_SPACINGS.map((sp) => (
                    <option key={sp.valuePx} value={sp.valuePx}>
                      Đoạn: {sp.label}
                    </option>
                  ))}
                </select>

                {/* Page Border & Ruler toggles */}
                <button
                  type="button"
                  onClick={() => setPageBorderEnabled((prev) => !prev)}
                  aria-pressed={pageBorderEnabled}
                  className={toolbarBtnClass(pageBorderEnabled)}
                >
                  <Square className="h-3.5 w-3.5 mr-1" />
                  <span>Viền trang</span>
                </button>
                <button
                  type="button"
                  onClick={() => setShowRuler((prev) => !prev)}
                  aria-pressed={showRuler}
                  className={toolbarBtnClass(showRuler)}
                >
                  <Ruler className="h-3.5 w-3.5 mr-1" />
                  <span>Thước kẻ</span>
                </button>
              </div>
            )}

            {/* TAB 4: XEM LẠI & CÔNG CỤ (REVIEW & TOOLS) */}
            {activeRibbonTab === 'view' && (
              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  type="button"
                  onClick={() => setIsFindOpen((prev) => !prev)}
                  aria-pressed={isFindOpen}
                  className={toolbarBtnClass(isFindOpen)}
                >
                  <Search className="h-3.5 w-3.5 text-[#185ABD] mr-1" />
                  <span>Tìm &amp; Thay thế</span>
                </button>

                <button
                  type="button"
                  onClick={() => setActiveToolModal('word_count')}
                  className={toolbarBtnClass(activeToolModal === 'word_count')}
                >
                  <BookOpen className="h-3.5 w-3.5 text-[#185ABD] mr-1" />
                  <span>Đếm từ ({wordCount.toLocaleString('vi-VN')})</span>
                </button>

                <button
                  type="button"
                  onClick={() => setIsCommentsOpen((prev) => !prev)}
                  aria-pressed={isCommentsOpen}
                  className={toolbarBtnClass(isCommentsOpen)}
                >
                  <MessageSquarePlus className="h-3.5 w-3.5 text-[#185ABD] mr-1" />
                  <span>Bình luận lề ({wordComments.length})</span>
                </button>

                <button
                  type="button"
                  onClick={() => setIsReadMode(true)}
                  className={toolbarBtnClass(false)}
                >
                  <Maximize2 className="h-3.5 w-3.5 text-[#185ABD] mr-1" />
                  <span>Đọc tập trung</span>
                </button>

                <button
                  type="button"
                  onClick={() => setActiveToolModal('extract_summary')}
                  className={toolbarBtnClass(
                    activeToolModal === 'extract_summary'
                  )}
                >
                  <FileSearch className="h-3.5 w-3.5 text-[#185ABD] mr-1" />
                  <span>Trích xuất &amp; Tóm tắt</span>
                </button>

                <button
                  type="button"
                  onClick={() => setIsDesktopSidebarOpen((prev) => !prev)}
                  aria-pressed={isDesktopSidebarOpen}
                  className={toolbarBtnClass(isDesktopSidebarOpen)}
                >
                  <span>Mục lục ({outlineEntries.length})</span>
                </button>
              </div>
            )}
          </div>

          {/* Floating Ribbon Popovers */}
          {openPopover === 'block' && (
            <div
              role="menu"
              className={`absolute left-32 top-full mt-1.5 w-48 rounded-xl border p-1.5 shadow-xl z-30 ${
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
                      ? 'bg-[#185ABD] text-white'
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

          {openPopover === 'underline' && (
            <div
              role="menu"
              className={`absolute left-64 top-full mt-1.5 w-44 rounded-xl border p-1.5 shadow-xl z-30 ${
                isDark
                  ? 'bg-[#1E2430] border-white/15 text-slate-100'
                  : 'bg-white border-[#E5E0D8] text-slate-900'
              }`}
            >
              {(
                [
                  { style: 'solid', label: 'Gạch chân đơn' },
                  { style: 'double', label: 'Gạch chân đôi' },
                  { style: 'dashed', label: 'Gạch chân nét đứt' },
                  { style: 'wavy', label: 'Gạch chân gợn sóng' },
                ] as const
              ).map((u) => (
                <button
                  key={u.style}
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleApplyUnderlineStyle(u.style)}
                  className="w-full text-left px-3 py-1.5 rounded-lg text-xs hover:bg-[#185ABD]/10 flex items-center justify-between"
                >
                  <span
                    style={{
                      textDecorationLine: 'underline',
                      textDecorationStyle: u.style,
                    }}
                  >
                    {u.label}
                  </span>
                  {underlineStyle === u.style && (
                    <Check className="h-3.5 w-3.5 text-[#185ABD]" />
                  )}
                </button>
              ))}
            </div>
          )}

          {openPopover === 'case' && (
            <div
              role="menu"
              className={`absolute left-80 top-full mt-1.5 w-48 rounded-xl border p-1.5 shadow-xl z-30 ${
                isDark
                  ? 'bg-[#1E2430] border-white/15 text-slate-100'
                  : 'bg-white border-[#E5E0D8] text-slate-900'
              }`}
            >
              {(
                [
                  { mode: 'sentence', label: 'Chữ đầu câu hoa (Sentence)' },
                  { mode: 'lower', label: 'chữ thường (lowercase)' },
                  { mode: 'upper', label: 'CHỮ IN HOA (UPPERCASE)' },
                  { mode: 'title', label: 'Viết Hoa Mỗi Từ (Title)' },
                ] as const
              ).map((c) => (
                <button
                  key={c.mode}
                  type="button"
                  onMouseDown={handlePreventMouseDown}
                  onClick={() => handleTransformCase(c.mode)}
                  className="w-full text-left px-3 py-1.5 rounded-lg text-xs hover:bg-[#185ABD]/10"
                >
                  {c.label}
                </button>
              ))}
            </div>
          )}

          {openPopover === 'table_grid' && (
            <div className="absolute left-36 top-full mt-1.5 z-30">
              <WordTableGridPicker
                isDark={isDark}
                onSelect={(r, c) => handleInsertTable(r, c)}
              />
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
      )}

      {/* Contextual Table Toolbar (when cursor is inside a table) */}
      {!isReadMode && !isEditingLocked && activeTableCell && (
        <WordTableContextBar
          isDark={isDark}
          onAddRow={() => handleTableAction('addRow')}
          onAddCol={() => handleTableAction('addCol')}
          onDeleteRow={() => handleTableAction('delRow')}
          onDeleteCol={() => handleTableAction('delCol')}
          onMergeRight={() => handleTableAction('mergeRight')}
          onSplitCell={() => handleTableAction('splitCell')}
          onCellFill={(color) => handleTableAction('fill', color)}
          onToggleBorders={() => handleTableAction('toggleBorders')}
          onDeleteTable={() => handleTableAction('deleteTable')}
        />
      )}

      {/* Contextual Image Align & Resize Toolbar (when clicking an image) */}
      {!isReadMode && !isEditingLocked && selectedImageEl && (
        <div
          role="toolbar"
          aria-label="Công cụ chỉnh ảnh"
          className={`shrink-0 border-b px-3 py-1.5 flex items-center gap-1.5 overflow-x-auto no-scrollbar z-10 text-[11px] ${
            isDark
              ? 'bg-[#161E2E] border-blue-400/25 text-slate-100'
              : 'bg-[#EFF4FC] border-[#C9DAF5] text-slate-900'
          }`}
        >
          <span className="font-bold text-[#185ABD] shrink-0 mr-1">
            Công cụ Ảnh:
          </span>
          <button
            type="button"
            onClick={() => handleUpdateSelectedImage({ align: 'left' })}
            className="h-7 px-2 rounded-lg border border-current/15 bg-white/80 text-slate-800"
          >
            Căn trái
          </button>
          <button
            type="button"
            onClick={() => handleUpdateSelectedImage({ align: 'center' })}
            className="h-7 px-2 rounded-lg border border-current/15 bg-white/80 text-slate-800"
          >
            Căn giữa
          </button>
          <button
            type="button"
            onClick={() => handleUpdateSelectedImage({ align: 'right' })}
            className="h-7 px-2 rounded-lg border border-current/15 bg-white/80 text-slate-800"
          >
            Căn phải
          </button>
          {[25, 50, 75, 100].map((pct) => (
            <button
              key={pct}
              type="button"
              onClick={() => handleUpdateSelectedImage({ widthPct: pct })}
              className="h-7 px-2 rounded-lg border border-current/15 bg-white/80 text-slate-800 tabular-nums"
            >
              {pct}%
            </button>
          ))}
          <button
            type="button"
            onClick={() => setSelectedImageEl(null)}
            className="h-7 px-2 rounded-lg text-slate-500 hover:bg-black/5 ml-auto"
          >
            Xong
          </button>
        </div>
      )}

      {/* Find & Replace Bar */}
      {isFindOpen && (
        <WordFindReplaceBar
          isDark={isDark}
          query={findQuery}
          replaceText={replaceText}
          matchCount={findMatchCount}
          activeMatchIndex={activeFindMatchIndex}
          caseSensitive={findCaseSensitive}
          onChangeQuery={setFindQuery}
          onChangeReplaceText={setReplaceText}
          onToggleCaseSensitive={() => setFindCaseSensitive((prev) => !prev)}
          onFindPrev={() => handleFindNavigate(-1)}
          onFindNext={() => handleFindNavigate(1)}
          onReplaceOne={handleReplaceOne}
          onReplaceAll={handleReplaceAll}
          onClose={() => setIsFindOpen(false)}
        />
      )}

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

      {/* Export Success / Status Notice */}
      {exportNotice && (
        <div
          role="status"
          className={`shrink-0 px-4 py-2 bg-emerald-500/15 border-b border-emerald-400/30 text-xs flex items-center justify-between gap-2 ${
            isDark ? 'text-emerald-200' : 'text-emerald-800'
          }`}
        >
          <span className="flex items-center gap-2 min-w-0">
            <Check className="h-3.5 w-3.5 shrink-0" />
            <span className="break-words">{exportNotice}</span>
          </span>
          <button
            type="button"
            onClick={() => setExportNotice(null)}
            aria-label="Đóng thông báo xuất tệp"
            className="p-1 rounded hover:bg-emerald-500/10 shrink-0"
          >
            <X className="h-3.5 w-3.5" />
          </button>
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

        {/* Desktop Left Auto-Outline Sidebar (240px, hidden in Read Mode) */}
        {!isReadMode && isDesktopSidebarOpen && (
          <aside
            aria-label="Mục lục tự động"
            className={`hidden md:flex w-[240px] shrink-0 border-r flex-col overflow-hidden transition-colors duration-200 ${themeStyle.sidebarBg} ${themeStyle.sidebarBorder}`}
          >
            <div
              className={`shrink-0 px-4 py-3 border-b flex items-center justify-between gap-2 ${themeStyle.sidebarBorder}`}
            >
              <div className="flex items-center gap-1.5 min-w-0">
                <Sparkles className="h-3.5 w-3.5 text-[#185ABD] shrink-0" />
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

        {/* Center Editable Paper Canvas with Pinch-Zoom & Horizontal Ruler */}
        <main
          ref={scrollContainerRef}
          onScroll={handleReadingScroll}
          className={`flex-1 min-w-0 overflow-y-auto overflow-x-auto scrollbar-none px-2.5 pr-6 py-4 sm:px-6 sm:pr-8 sm:py-6 scroll-smooth`}
        >
          {/* Read Mode Floating Exit Bar */}
          {isReadMode && (
            <div className="sticky top-2 z-30 flex justify-end mb-3 pointer-events-none">
              <button
                type="button"
                onClick={() => setIsReadMode(false)}
                className="pointer-events-auto inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-[#185ABD] hover:bg-blue-700 text-white text-xs font-semibold shadow-lg transition-colors"
              >
                <Minimize2 className="h-3.5 w-3.5" />
                <span>Thoát chế độ đọc tập trung</span>
              </button>
            </div>
          )}

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

          {/* Horizontal Ruler (in Print Layout when enabled) */}
          {!isReadMode && pageLayout === 'print' && showRuler && (
            <WordHorizontalRuler
              pageWidthPx={Math.round(basePaperWidthPx * zoomScale)}
              pageWidthCm={pageDimensions.widthCm}
              marginLeftCm={marginLeftCm}
              marginRightCm={marginRightCm}
              firstLineIndentCm={firstLineIndentCm}
              isDark={isDark}
              disabled={isEditingLocked}
              onChangeMargins={({
                marginLeftCm: nextL,
                marginRightCm: nextR,
                firstLineIndentCm: nextI,
              }) => {
                setMarginPreset('custom');
                setMarginLeftCm(nextL);
                setMarginRightCm(nextR);
                setFirstLineIndentCm(nextI);
              }}
            />
          )}

          {/* Format Painter Active Banner */}
          {isFormatPainterActive && !isReadMode && (
            <div className="mx-auto max-w-md mb-2 px-3 py-1.5 rounded-full bg-[#185ABD] text-white text-[11px] font-semibold text-center shadow-md">
              Chổi sao chép định dạng đang bật — bôi đen đoạn chữ đích để áp dụng (Esc để hủy)
            </div>
          )}

          {/* Layout-Aware Scaled Page Host (25%-200% zoom only scales document) */}
          {(() => {
            const isPrint = pageLayout === 'print';
            const scaledWidthPx = Math.round(basePaperWidthPx * zoomScale);
            const outerStyle: CSSProperties =
              isPrint && zoomScale !== 1
                ? {
                    width: `${scaledWidthPx}px`,
                    ...(contentHeight > 0
                      ? { height: `${Math.ceil(contentHeight * zoomScale)}px` }
                      : {}),
                    margin: '0 auto',
                    position: 'relative',
                  }
                : {};
            const innerStyle: CSSProperties =
              isPrint && zoomScale !== 1
                ? {
                    ...hostCssVars,
                    marginLeft: `${(scaledWidthPx - basePaperWidthPx) / 2}px`,
                    transform: `scale(${zoomScale})`,
                    transformOrigin: 'top center',
                  }
                : hostCssVars;

            const paperPaddingStyle: CSSProperties =
              !isDocxPreviewRendered && isPrint
                ? {
                    paddingTop: `${cmToPx(marginTopCm)}px`,
                    paddingBottom: `${cmToPx(marginBottomCm)}px`,
                    paddingLeft: `${cmToPx(marginLeftCm)}px`,
                    paddingRight: `${cmToPx(marginRightCm)}px`,
                    minHeight: `${pageDimensions.minHeightPx}px`,
                    columnCount: columnCount > 1 ? columnCount : undefined,
                    columnGap: columnCount > 1 ? '28px' : undefined,
                  }
                : !isDocxPreviewRendered
                  ? {
                      columnCount: columnCount > 1 ? columnCount : undefined,
                      columnGap: columnCount > 1 ? '24px' : undefined,
                    }
                  : {};

            return (
              <div style={outerStyle}>
                <div
                  ref={pageSurfaceRef}
                  style={innerStyle}
                  className={`lite-word-host mx-auto transition-colors duration-200 ${
                    pageLayout === 'mobile'
                      ? 'word-layout-mobile'
                      : 'word-layout-print'
                  } ${
                    isDocxPreviewRendered
                      ? ''
                      : `lite-word-editor-surface rounded-xl sm:rounded-2xl border shadow-sm flex flex-col ${themeStyle.paperBg} ${themeStyle.paperBorder} ${
                          pageBorderEnabled
                            ? 'ring-2 ring-[#185ABD]/35 ring-offset-2 ring-offset-transparent'
                            : ''
                        }`
                  }`}
                >
                  {/* Optional Page Header */}
                  {showHeaderFooter && !isDocxPreviewRendered && (
                    <div
                      className={`px-6 pt-3 pb-2 border-b text-[11px] flex items-center justify-between gap-2 ${
                        isDark
                          ? 'border-white/10 text-slate-400'
                          : 'border-slate-200/80 text-slate-500'
                      }`}
                    >
                      <input
                        type="text"
                        value={headerText}
                        onChange={(e) => setHeaderText(e.target.value)}
                        disabled={isEditingLocked}
                        aria-label="Tiêu đề đầu trang (Header)"
                        className="bg-transparent outline-none flex-1 min-w-0 font-medium"
                      />
                      {showPageNumbers && (
                        <span className="tabular-nums shrink-0">
                          Trang {currentScrollPage} / {detectedPageCount}
                        </span>
                      )}
                    </div>
                  )}

                  <div
                    ref={editorRef}
                    contentEditable={!isEditingLocked && !isRenderingDocx}
                    suppressContentEditableWarning
                    spellCheck={false}
                    role="textbox"
                    aria-multiline="true"
                    aria-readonly={isEditingLocked}
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
                    onClick={(event) => {
                      const target = event.target;
                      if (
                        target instanceof HTMLImageElement &&
                        editorRef.current?.contains(target)
                      ) {
                        setSelectedImageEl(target);
                      } else {
                        setSelectedImageEl(null);
                      }
                    }}
                    onKeyUp={() => {
                      captureCurrentSelection();
                      refreshFormatState();
                    }}
                    onMouseUp={() => {
                      applyFormatPainterToSelection();
                      captureCurrentSelection();
                      refreshFormatState();
                    }}
                    style={{
                      fontFamily: fontStack,
                      fontSize: `${pageFontSizePx}px`,
                      lineHeight,
                      ...paperPaddingStyle,
                    }}
                    className={
                      isDocxPreviewRendered
                        ? 'min-h-[680px] outline-none break-words'
                        : `flex-1 min-h-[720px] ${
                            isPrint ? '' : 'px-5 sm:px-10 py-6 sm:py-10'
                          } outline-none break-words ${themeStyle.textPrimary} ${themeStyle.editorProse} [&_h1]:text-[1.65em] [&_h1]:font-bold [&_h1]:leading-snug [&_h1]:tracking-tight [&_h1]:mb-4 [&_h1]:mt-2 [&_h1]:scroll-mt-6 [&_h2]:text-[1.3em] [&_h2]:font-bold [&_h2]:leading-snug [&_h2]:mt-7 [&_h2]:mb-3 [&_h2]:pb-2 [&_h2]:border-b [&_h2]:scroll-mt-6 [&_h3]:text-[1.12em] [&_h3]:font-semibold [&_h3]:mt-5 [&_h3]:mb-2 [&_h3]:scroll-mt-6 [&_p]:mb-[var(--lite-doc-para-spacing,12px)] [&_p]:indent-[var(--lite-doc-first-indent,0px)] [&_ul]:list-disc [&_ul]:pl-6 [&_ul]:mb-4 [&_ul]:space-y-1.5 [&_ol]:list-decimal [&_ol]:pl-6 [&_ol]:mb-4 [&_ol]:space-y-1.5 [&_blockquote]:border-l-4 [&_blockquote]:rounded-r-xl [&_blockquote]:px-4 [&_blockquote]:py-3 [&_blockquote]:my-4 [&_table]:w-full [&_table]:border-collapse [&_table]:my-5 [&_table]:text-[0.94em] [&_th]:border [&_th]:px-3 [&_th]:py-2 [&_th]:text-left [&_th]:font-semibold [&_td]:border [&_td]:px-3 [&_td]:py-2 [&_td]:align-top [&_img]:max-w-full [&_img]:h-auto [&_img]:rounded-xl [&_img]:my-4 [&_img]:cursor-pointer`
                    }
                  />

                  {/* Optional Page Footer */}
                  {showHeaderFooter && !isDocxPreviewRendered && (
                    <div
                      className={`px-6 py-2.5 border-t text-[11px] flex items-center justify-between gap-2 ${
                        isDark
                          ? 'border-white/10 text-slate-400'
                          : 'border-slate-200/80 text-slate-500'
                      }`}
                    >
                      <input
                        type="text"
                        value={footerText}
                        onChange={(e) => setFooterText(e.target.value)}
                        disabled={isEditingLocked}
                        aria-label="Chân trang (Footer)"
                        className="bg-transparent outline-none flex-1 min-w-0"
                      />
                      <label className="inline-flex items-center gap-1 cursor-pointer shrink-0">
                        <input
                          type="checkbox"
                          checked={showPageNumbers}
                          onChange={(e) => setShowPageNumbers(e.target.checked)}
                        />
                        <span>Số trang</span>
                      </label>
                    </div>
                  )}
                </div>
              </div>
            );
          })()}

          <div className="h-8" aria-hidden="true" />
        </main>

        {/* Right Margin Comments Panel */}
        {isCommentsOpen && !isReadMode && (
          <WordCommentsPanel
            isDark={isDark}
            comments={wordComments}
            isReadOnly={isReadOnly}
            selectedQuote={selectedQuote}
            onAddComment={handleAddComment}
            onDeleteComment={handleDeleteComment}
            onJumpToQuote={handleJumpToQuote}
            onClose={() => setIsCommentsOpen(false)}
          />
        )}

        {/* Custom Tactile Vertical Scrollbar */}
        <CustomVerticalScrollbar
          scrollContainerRef={scrollContainerRef}
          theme={isDark ? 'dark' : 'light'}
          bottomOffsetClass="bottom-1"
          getDragLabel={({ percent }) => {
            const totalPages = Math.max(1, detectedPageCount);
            const estPage = Math.min(
              totalPages,
              Math.max(1, Math.round((percent / 100) * (totalPages - 1)) + 1)
            );
            return `Trang ${estPage} / ${totalPages} · ${Math.round(percent)}%`;
          }}
        />

        {/* Floating Mobile Outline Trigger Button (< md) */}
        {!isMobileDrawerOpen && !isReadMode && (
          <button
            type="button"
            onMouseDown={handlePreventMouseDown}
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở Mục lục tự động"
            className={`md:hidden absolute bottom-3 left-3 z-30 inline-flex items-center gap-1.5 min-h-[36px] px-3 rounded-full border shadow-lg text-xs font-semibold transition-all ${
              isDark
                ? 'bg-[#1E2430]/95 border-white/15 text-white hover:bg-[#273041]'
                : 'bg-white/95 border-[#E2DDD4] text-slate-900 hover:bg-white'
            }`}
          >
            <List className="h-3.5 w-3.5 text-[#185ABD]" />
            <span>Mục lục</span>
            {outlineEntries.length > 0 && (
              <span className="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-[#185ABD] text-white text-[10px] font-bold tabular-nums">
                {outlineEntries.length}
              </span>
            )}
          </button>
        )}
      </div>

      {/* Fixed Bottom Office Status Bar (#185ABD) */}
      <OfficeStatusBar
        variant="word"
        theme={isDark ? 'dark' : 'light'}
        zoomPercent={zoomPercent}
        onZoomChange={setZoomPercent}
        onFitToWidth={handleFitToWidth}
        pageCurrent={currentScrollPage}
        pageTotal={detectedPageCount}
        wordCount={docStats.words}
        charCount={charCount || docStats.charsWithSpaces}
        onOpenWordStats={() => setActiveToolModal('word_count')}
      />

      {/* Insert & Review Modals (Word Count, Hyperlink, Symbol/Equation, Shapes/TextBox, Summary) */}
      <WordInsertToolModals
        activeModal={activeToolModal}
        isDark={isDark}
        stats={docStats}
        extractedText={(editorRef.current?.innerText ?? '').trim()}
        outlineTitles={outlineEntries.map((entry) => entry.section.title)}
        onClose={() => setActiveToolModal(null)}
        onInsertLink={(label, url) => {
          const safeUrl = escapeHtml(url);
          const safeLabel = escapeHtml(label);
          executeFormatCommand(
            'insertHTML',
            `<a href="${safeUrl}" target="_blank" rel="noopener noreferrer" style="color:#185ABD;text-decoration:underline;">${safeLabel}</a>&nbsp;`
          );
        }}
        onInsertHtml={(html) => {
          executeFormatCommand('insertHTML', html);
        }}
      />
    </div>
  );
}

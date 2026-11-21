import { useCallback, useEffect, useRef, useState } from 'react';
import type { MouseEvent as ReactMouseEvent } from 'react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  BookOpen,
  Check,
  ChevronDown,
  Download,
  FolderOpen,
  Highlighter,
  Italic,
  List,
  ListOrdered,
  Loader2,
  Moon,
  Palette,
  PanelLeftClose,
  PanelLeftOpen,
  Save,
  Sparkles,
  Sun,
  Table,
  Underline,
  X,
} from 'lucide-react';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  WordPaperTheme,
} from '../../types/document';
import { extractOutlineFromHtml } from '../../utils/fileParser';
import { exportWordToDocx } from '../../utils/fileExporter';
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
}

const FONT_SIZES = [14, 16, 18, 20] as const;
type FontSizeOption = (typeof FONT_SIZES)[number];

type BlockTag = 'p' | 'h1' | 'h2' | 'h3';
type ActivePopover = 'block' | 'color' | 'highlight' | null;

interface OutlineNavEntry {
  section: OutlineSection;
  level: 1 | 2 | 3;
}

interface FormatActiveState {
  block: BlockTag;
  bold: boolean;
  italic: boolean;
  underline: boolean;
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

const BLOCK_STYLE_OPTIONS: { tag: BlockTag; label: string; previewClass: string }[] = [
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
];

const HIGHLIGHT_SWATCHES = [
  { label: 'Xóa nền', value: 'transparent', chipClass: 'bg-transparent border-dashed' },
  { label: 'Vàng', value: '#FEF08A', chipClass: 'bg-[#FEF08A]' },
  { label: 'Xanh lá', value: '#BBF7D0', chipClass: 'bg-[#BBF7D0]' },
  { label: 'Xanh dương', value: '#BFDBFE', chipClass: 'bg-[#BFDBFE]' },
  { label: 'Hồng', value: '#FBCFE8', chipClass: 'bg-[#FBCFE8]' },
];

const INITIAL_FORMAT_STATE: FormatActiveState = {
  block: 'p',
  bold: false,
  italic: false,
  underline: false,
  align: 'left',
  unorderedList: false,
  orderedList: false,
};

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

function buildInitialWordHtml(doc: DocumentItem): string {
  if (doc.wordHtml && doc.wordHtml.trim().length > 0) {
    return extractOutlineFromHtml(doc.wordHtml).htmlWithIds;
  }

  const parts: string[] = [];
  parts.push(`<h1>${escapeHtml(doc.title)}</h1>`);

  if (doc.contentPreview && doc.contentPreview.trim().length > 0) {
    parts.push(`<blockquote>${escapeHtml(doc.contentPreview.trim())}</blockquote>`);
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
    parts.push('<p>Bắt đầu soạn thảo nội dung tài liệu của bạn tại đây...</p>');
  }

  return extractOutlineFromHtml(parts.join('')).htmlWithIds;
}

const QUICK_TABLE_HTML =
  '<table><thead><tr><th>Cột 1</th><th>Cột 2</th><th>Cột 3</th></tr></thead><tbody><tr><td>Nội dung 1.1</td><td>Nội dung 1.2</td><td>Nội dung 1.3</td></tr><tr><td>Nội dung 2.1</td><td>Nội dung 2.2</td><td>Nội dung 2.3</td></tr></tbody></table><p><br></p>';

export function WordWorkspace({
  document,
  onBackHome,
  onOpenModal,
  onSaveDocument,
  isModalOpen = false,
  defaultPaperTheme = 'light',
  autoSaveEnabled = false,
}: WordWorkspaceProps) {
  const [paperTheme, setPaperTheme] = useState<WordPaperTheme>(defaultPaperTheme);
  const [fontSizeIndex, setFontSizeIndex] = useState<number>(1); // 16px default
  const [isDesktopSidebarOpen, setIsDesktopSidebarOpen] =
    useState<boolean>(true);
  const [isMobileDrawerOpen, setIsMobileDrawerOpen] = useState<boolean>(false);
  const [outlineEntries, setOutlineEntries] = useState<OutlineNavEntry[]>([]);
  const [activeSectionId, setActiveSectionId] = useState<string>('');
  const [isDirty, setIsDirty] = useState<boolean>(false);
  const [editRevision, setEditRevision] = useState<number>(0);
  const [saveFeedback, setSaveFeedback] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [openPopover, setOpenPopover] = useState<ActivePopover>(null);
  const [customTextColor, setCustomTextColor] = useState<string>('#2563EB');
  const [formatState, setFormatState] =
    useState<FormatActiveState>(INITIAL_FORMAT_STATE);
  const [wordCount, setWordCount] = useState<number>(0);

  const editorRef = useRef<HTMLDivElement | null>(null);
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const ribbonRef = useRef<HTMLDivElement | null>(null);
  const isComposingRef = useRef<boolean>(false);
  const loadedDocIdRef = useRef<string | null>(null);
  const latestHtmlRef = useRef<string>('');
  const savedRangeRef = useRef<Range | null>(null);
  const saveFeedbackTimerRef = useRef<number | null>(null);
  const exportErrorTimerRef = useRef<number | null>(null);
  const documentRef = useRef<DocumentItem>(document);
  documentRef.current = document;

  const fontSize: FontSizeOption = FONT_SIZES[fontSizeIndex];
  const themeStyle = PAPER_THEMES[paperTheme];
  const isDark = paperTheme === 'dark';

  // Scan headings directly from live DOM without recreating nodes (IME-safe)
  const syncOutlineFromDom = useCallback(() => {
    const editorEl = editorRef.current;
    if (!editorEl) {
      return;
    }

    const headingNodes = Array.from(editorEl.querySelectorAll('h1, h2, h3'));
    const usedIds = new Set<string>();
    const nextEntries: OutlineNavEntry[] = [];
    let counter = 1;

    for (let idx = 0; idx < headingNodes.length; idx += 1) {
      const heading = headingNodes[idx];
      const text = (heading.textContent ?? '').replace(/\s+/g, ' ').trim();
      if (!text) {
        continue;
      }

      let id = heading.id.trim();
      if (!id || usedIds.has(id)) {
        while (usedIds.has(`heading-${counter}`)) {
          counter += 1;
        }
        id = `heading-${counter}`;
        heading.id = id;
        counter += 1;
      }
      usedIds.add(id);

      const tagLower = heading.tagName.toLowerCase();
      const level: 1 | 2 | 3 =
        tagLower === 'h1' ? 1 : tagLower === 'h2' ? 2 : 3;

      nextEntries.push({
        level,
        section: {
          id,
          romanIndex:
            level === 1
              ? toRomanNumeral(nextEntries.filter((e) => e.level === 1).length + 1)
              : level === 2
                ? toRomanNumeral(nextEntries.length + 1)
                : String(nextEntries.length + 1),
          title: text,
          summary: '',
          keyPoints: [],
          pageNumber: Math.max(1, Math.floor(idx / 2) + 1),
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

    const plainText = (editorEl.textContent ?? '').trim();
    setWordCount(plainText ? plainText.split(/\s+/).filter(Boolean).length : 0);
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

  // Initialize contentEditable DOM strictly when document.id changes (never on save)
  useEffect(() => {
    if (loadedDocIdRef.current === document.id) {
      return;
    }
    loadedDocIdRef.current = document.id;
    const initialHtml = buildInitialWordHtml(document);
    latestHtmlRef.current = initialHtml;

    if (editorRef.current) {
      editorRef.current.innerHTML = initialHtml;
    }

    savedRangeRef.current = null;
    setIsDirty(false);
    setEditRevision(0);
    setSaveFeedback(false);
    setExportError(null);
    setOpenPopover(null);
    setIsMobileDrawerOpen(false);

    if (scrollContainerRef.current) {
      scrollContainerRef.current.scrollTop = 0;
    }

    syncOutlineFromDom();
  }, [document, syncOutlineFromDom]);

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

  // Cleanup timers on unmount
  useEffect(() => {
    return () => {
      if (saveFeedbackTimerRef.current !== null) {
        window.clearTimeout(saveFeedbackTimerRef.current);
      }
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
    };
  }, []);

  const buildLiveDocument = useCallback((): DocumentItem => {
    const currentDoc = documentRef.current;
    const rawHtml = editorRef.current?.innerHTML ?? latestHtmlRef.current;
    const { htmlWithIds, outline, plainText } = extractOutlineFromHtml(rawHtml);
    const trimmedPreview =
      plainText.length > 260
        ? `${plainText.slice(0, 259).trimEnd()}…`
        : plainText;

    return {
      ...currentDoc,
      wordHtml: htmlWithIds,
      outline: outline.length > 0 ? outline : currentDoc.outline,
      contentPreview: trimmedPreview || currentDoc.contentPreview,
    };
  }, []);

  const handleSave = useCallback(() => {
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
  }, [buildLiveDocument, onSaveDocument]);

  // Auto-save with 1800ms debounce (IME-safe; resets on every edit & cancels on manual save/unmount)
  useEffect(() => {
    if (!autoSaveEnabled || !isDirty || isComposingRef.current) {
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
  }, [autoSaveEnabled, isDirty, editRevision, handleSave]);

  // Keyboard shortcuts: Escape (close popover / mobile drawer) & Ctrl/Cmd+S (save)
  useEffect(() => {
    if (isModalOpen) {
      return;
    }
    const handleKeyDown = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        handleSave();
        return;
      }

      if (event.key === 'Escape') {
        if (openPopover !== null) {
          setOpenPopover(null);
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
  }, [isModalOpen, isMobileDrawerOpen, openPopover, handleSave]);

  const handleEditorMutation = useCallback(() => {
    if (!editorRef.current) {
      return;
    }
    latestHtmlRef.current = editorRef.current.innerHTML;
    setIsDirty(true);
    setEditRevision((prev) => prev + 1);
    setSaveFeedback(false);
    syncOutlineFromDom();
    refreshFormatState();
  }, [syncOutlineFromDom, refreshFormatState]);

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

    // If no active selection inside editor yet, place caret at end
    if (
      selection.rangeCount === 0 ||
      !editorEl.contains(selection.anchorNode)
    ) {
      const endRange = window.document.createRange();
      endRange.selectNodeContents(editorEl);
      endRange.collapse(false);
      selection.removeAllRanges();
      selection.addRange(endRange);
      savedRangeRef.current = endRange.cloneRange();
    }
  }, []);

  const executeFormatCommand = useCallback(
    (command: string, value?: string) => {
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
    [restoreEditorSelection, captureCurrentSelection, handleEditorMutation]
  );

  const handleApplyBlockTag = (tag: BlockTag) => {
    executeFormatCommand('formatBlock', `<${tag}>`);
    setOpenPopover(null);
  };

  const handleInsertTable = () => {
    executeFormatCommand('insertHTML', QUICK_TABLE_HTML);
    setOpenPopover(null);
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
      const currentHtml = editorRef.current?.innerHTML ?? latestHtmlRef.current;
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
    const target = Array.from(editorEl.querySelectorAll('h1, h2, h3')).find(
      (node) => node.id === sectionId
    );
    target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  const handleReadingScroll = () => {
    const container = scrollContainerRef.current;
    const editorEl = editorRef.current;
    if (!container || !editorEl || outlineEntries.length === 0) {
      return;
    }
    const containerRect = container.getBoundingClientRect();
    const headingMap = new Map<string, Element>();
    for (const node of Array.from(editorEl.querySelectorAll('h1, h2, h3'))) {
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
      if (relativeTop <= 130) {
        currentId = entry.section.id;
      } else {
        break;
      }
    }
    setActiveSectionId((prev) => (prev === currentId ? prev : currentId));
  };

  const currentBlockLabel =
    BLOCK_STYLE_OPTIONS.find((item) => item.tag === formatState.block)?.label ??
    'Văn bản thường';

  const toolbarBtnClass = (active: boolean) =>
    `inline-flex items-center justify-center min-h-[32px] min-w-[32px] px-2 rounded-lg text-xs font-medium transition-colors shrink-0 ${
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
                className={`w-full text-left rounded-xl pr-3 py-2 transition-all flex items-start justify-between gap-2 ${indentClass} ${
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
                  className={`text-[10px] font-mono uppercase shrink-0 pt-0.5 ${themeStyle.textMuted}`}
                >
                  H{level}
                </span>
              </button>
            );
          })
        ) : (
          <div
            className={`rounded-xl border border-dashed p-3.5 text-xs leading-relaxed ${themeStyle.calloutBorder} ${themeStyle.textMuted}`}
          >
            Chưa có tiêu đề nào. Chọn <strong>Tiêu đề 1 / 2 / 3</strong> trên
            thanh công cụ để tự động tạo mục lục tại đây.
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
              {outlineEntries.length} tiêu đề · {wordCount.toLocaleString('vi-VN')}{' '}
              từ
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
              {document.isUploaded ? 'Tệp của bạn' : 'Tài liệu mẫu'}
              {document.fileSize ? ` · ${document.fileSize}` : ''}
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div
      className={`h-[100dvh] w-full flex flex-col overflow-hidden transition-colors duration-200 ${themeStyle.outerBg} ${themeStyle.textPrimary} ${themeStyle.scrollbarClass}`}
    >
      {/* Shared Top Workspace Header */}
      <WorkspaceTopBar
        document={document}
        theme={isDark ? 'dark' : 'light'}
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument())}
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
            {/* Save Changes Button */}
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

      {/* Rich Word Formatting Ribbon Bar */}
      <div
        ref={ribbonRef}
        className={`shrink-0 border-b px-2.5 sm:px-4 py-1.5 z-10 transition-colors duration-200 relative ${themeStyle.ribbonBg} ${themeStyle.ribbonBorder}`}
      >
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-1 overflow-x-auto no-scrollbar py-0.5 min-w-0">
            {/* Heading / Paragraph Style Dropdown Trigger */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() =>
                setOpenPopover((prev) => (prev === 'block' ? null : 'block'))
              }
              aria-expanded={openPopover === 'block'}
              aria-label="Chọn kiểu đoạn hoặc tiêu đề"
              className={`shrink-0 inline-flex items-center justify-between gap-1.5 min-h-[32px] px-2.5 rounded-lg border text-xs font-medium transition-colors whitespace-nowrap ${
                isDark
                  ? 'bg-[#1E2534] border-white/10 text-slate-200 hover:bg-white/10'
                  : 'bg-[#FAF8F5] border-[#E4DFD6] text-slate-800 hover:bg-[#F2EFEA]'
              }`}
            >
              <span>{currentBlockLabel}</span>
              <ChevronDown className="h-3.5 w-3.5 opacity-70" />
            </button>

            <span
              aria-hidden="true"
              className={`h-5 w-px mx-1 shrink-0 ${
                isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
              }`}
            />

            {/* Bold / Italic / Underline */}
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

            <span
              aria-hidden="true"
              className={`h-5 w-px mx-1 shrink-0 ${
                isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
              }`}
            />

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
              title="Màu nền highlight"
              className={toolbarBtnClass(openPopover === 'highlight')}
            >
              <Highlighter className="h-3.5 w-3.5 text-amber-500" />
            </button>

            <span
              aria-hidden="true"
              className={`h-5 w-px mx-1 shrink-0 ${
                isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
              }`}
            />

            {/* Alignment: Left / Center / Right / Justify */}
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

            <span
              aria-hidden="true"
              className={`h-5 w-px mx-1 shrink-0 ${
                isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
              }`}
            />

            {/* Lists: Bullet & Numbered */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => executeFormatCommand('insertUnorderedList')}
              aria-pressed={formatState.unorderedList}
              aria-label="Danh sách dấu chấm"
              title="Danh sách chấm (Bullet)"
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
              title="Danh sách số (Numbered)"
              className={toolbarBtnClass(formatState.orderedList)}
            >
              <ListOrdered className="h-3.5 w-3.5" />
            </button>

            {/* Quick Insert 3x3 Table */}
            <button
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={handleInsertTable}
              aria-label="Chèn bảng 3x3 nhanh"
              title="Chèn bảng 3×3 tại vị trí con trỏ"
              className={`shrink-0 inline-flex items-center gap-1.5 min-h-[32px] px-2.5 rounded-lg text-xs font-medium transition-colors whitespace-nowrap ${
                isDark
                  ? 'text-slate-200 hover:bg-white/[0.08]'
                  : 'text-slate-700 hover:bg-black/[0.05]'
              }`}
            >
              <Table className="h-3.5 w-3.5 text-indigo-500 shrink-0" />
              <span className="hidden lg:inline">Chèn bảng</span>
            </button>

            <span
              aria-hidden="true"
              className={`h-5 w-px mx-1 shrink-0 ${
                isDark ? 'bg-white/10' : 'bg-[#E6E2DA]'
              }`}
            />

            {/* Font Size Controls (14px / 16px / 18px / 20px) */}
            <div
              role="group"
              aria-label="Điều chỉnh cỡ chữ trang A4"
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

            {/* 3 Paper Theme Switcher */}
            <div
              role="group"
              aria-label="Chọn màu giấy trang A4"
              className={`inline-flex items-center gap-0.5 p-0.5 rounded-lg border shrink-0 ml-0.5 ${
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
                    className={`inline-flex items-center gap-1 min-h-[26px] px-2 rounded-md text-[11px] font-medium transition-all whitespace-nowrap ${
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
                    <span className="hidden xl:inline">{config.shortLabel}</span>
                  </button>
                );
              })}
            </div>
          </div>
        </div>

        {/* Floating Ribbon Popovers (positioned cleanly outside overflow-x-auto) */}
        {openPopover === 'block' && (
          <div
            role="menu"
            className={`absolute left-3 top-full mt-1.5 w-48 rounded-xl border p-1.5 shadow-xl z-30 ${
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
            className={`absolute left-3 sm:left-48 top-full mt-1.5 w-60 max-w-[calc(100vw-1.5rem)] rounded-xl border p-3 shadow-xl z-30 space-y-2.5 ${
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
            className={`absolute left-3 sm:left-56 top-full mt-1.5 max-w-[calc(100vw-1.5rem)] rounded-xl border p-3 shadow-xl z-30 space-y-2 ${
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
          className="flex-1 min-w-0 overflow-y-auto px-3 py-4 sm:px-6 sm:py-8 scroll-smooth"
        >
          <div
            className={`max-w-[794px] mx-auto rounded-xl sm:rounded-2xl border shadow-sm transition-colors duration-200 ${themeStyle.paperBg} ${themeStyle.paperBorder}`}
          >
            {/* Subtle Document Meta Strip inside top of A4 */}
            <div
              className={`px-6 sm:px-12 pt-5 pb-3 border-b flex flex-wrap items-center justify-between gap-2 text-[11.5px] ${themeStyle.calloutBorder} ${themeStyle.textMuted}`}
            >
              <div className="flex items-center gap-2 min-w-0">
                <span
                  className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full font-semibold ${
                    isDark
                      ? 'bg-blue-500/20 text-blue-300'
                      : 'bg-blue-50 text-blue-700'
                  }`}
                >
                  <BookOpen className="h-3 w-3" />
                  <span>Soạn thảo trực tiếp</span>
                </span>
                <span className="truncate">{document.updatedAt}</span>
              </div>
              <span className="tabular-nums">
                {wordCount.toLocaleString('vi-VN')} từ · {fontSize}px
              </span>
            </div>

            {/* Direct ContentEditable A4 Surface (IME / Unikey Safe) */}
            <div
              ref={editorRef}
              contentEditable
              suppressContentEditableWarning
              spellCheck={false}
              role="textbox"
              aria-multiline="true"
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
              style={{ fontSize: `${fontSize}px` }}
              className={`min-h-[680px] px-6 sm:px-12 py-6 sm:py-10 outline-none leading-[1.8] break-words ${themeStyle.textPrimary} ${themeStyle.editorProse} [&_h1]:text-[1.7em] [&_h1]:font-bold [&_h1]:leading-snug [&_h1]:tracking-tight [&_h1]:mb-4 [&_h1]:mt-2 [&_h1]:scroll-mt-6 [&_h2]:text-[1.32em] [&_h2]:font-bold [&_h2]:leading-snug [&_h2]:mt-7 [&_h2]:mb-3 [&_h2]:pb-2 [&_h2]:border-b [&_h2]:scroll-mt-6 [&_h3]:text-[1.12em] [&_h3]:font-semibold [&_h3]:mt-5 [&_h3]:mb-2 [&_h3]:scroll-mt-6 [&_p]:mb-3.5 [&_ul]:list-disc [&_ul]:pl-6 [&_ul]:mb-4 [&_ul]:space-y-1.5 [&_ol]:list-decimal [&_ol]:pl-6 [&_ol]:mb-4 [&_ol]:space-y-1.5 [&_li]:leading-[1.75] [&_blockquote]:border-l-4 [&_blockquote]:rounded-r-xl [&_blockquote]:px-4 [&_blockquote]:py-3 [&_blockquote]:my-4 [&_table]:w-full [&_table]:border-collapse [&_table]:my-5 [&_table]:text-[0.92em] [&_th]:border [&_th]:px-3 [&_th]:py-2 [&_th]:text-left [&_th]:font-semibold [&_td]:border [&_td]:px-3 [&_td]:py-2 [&_td]:align-top [&_img]:max-w-full [&_img]:h-auto [&_img]:rounded-xl [&_img]:my-4`}
            />
          </div>

          {/* Extra spacing below A4 so mobile floating TOC button never overlaps last line */}
          <div className="h-16 md:h-6" aria-hidden="true" />
        </main>

        {/* Floating Mobile Outline Trigger Button (< md) */}
        {!isMobileDrawerOpen && (
          <button
            type="button"
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở Mục lục tự động"
            className={`md:hidden absolute bottom-[max(1rem,env(safe-area-inset-bottom))] left-4 z-20 inline-flex items-center gap-2 min-h-[42px] px-3.5 rounded-full border shadow-lg text-xs font-semibold transition-all ${
              isDark
                ? 'bg-[#1E2430]/95 border-white/15 text-white hover:bg-[#273041]'
                : 'bg-white/95 border-[#E2DDD4] text-slate-900 hover:bg-white'
            }`}
          >
            <List className="h-4 w-4 text-indigo-500" />
            <span>Mục lục</span>
            {outlineEntries.length > 0 && (
              <span className="inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-indigo-500 text-white text-[10px] font-bold tabular-nums">
                {outlineEntries.length}
              </span>
            )}
          </button>
        )}
      </div>
    </div>
  );
}

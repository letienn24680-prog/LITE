import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
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
  Copy,
  Crosshair,
  Download,
  Image,
  Italic,
  Layers,
  Loader2,
  Pause,
  Play,
  Plus,
  QrCode,
  RotateCcw,
  Save,
  StickyNote,
  Trash2,
  Type,
  X,
} from 'lucide-react';
import type {
  DocumentItem,
  ModalType,
  OutlineSection,
  SlidePreviewItem,
  SlideTextBlock,
} from '../../types/document';
import { exportSlidesToPptx } from '../../utils/fileExporter';
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
}

interface LaserPosition {
  x: number;
  y: number;
  visible: boolean;
}

const SLIDE_FONT_SIZES = [14, 18, 24, 32, 40] as const;

const SLIDE_COLOR_SWATCHES = [
  { label: 'Trắng', value: '#FFFFFF' },
  { label: 'Bạc sáng', value: '#CBD5E1' },
  { label: 'Vàng nắng', value: '#FDE047' },
  { label: 'Xanh trời', value: '#38BDF8' },
  { label: 'Xanh ngọc', value: '#4ADE80' },
  { label: 'Hồng san hô', value: '#FB7185' },
  { label: 'Đen than', value: '#1E293B' },
];

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
  if (!/^[0-9a-fA-F]{6}$/.test(cleaned)) {
    return false;
  }
  const r = parseInt(cleaned.slice(0, 2), 16);
  const g = parseInt(cleaned.slice(2, 4), 16);
  const b = parseInt(cleaned.slice(4, 6), 16);
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
      blocks: slide.blocks.map((block) => ({ ...block })),
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
  const titleBlock =
    blocks.find((block) => block.role === 'title' && block.text.trim()) ??
    blocks.find((block) => block.text.trim());
  const subtitleBlock = blocks.find(
    (block) => block.role === 'subtitle' && block.text.trim()
  );
  const bodyBullets = blocks
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

interface AutoGrowTextareaProps {
  value: string;
  onChange: (value: string) => void;
  onFocus: () => void;
  disabled: boolean;
  placeholder: string;
  ariaLabel: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  color: string;
  align: 'left' | 'center' | 'right';
  autoFocus?: boolean;
}

function AutoGrowTextarea({
  value,
  onChange,
  onFocus,
  disabled,
  placeholder,
  ariaLabel,
  fontSize,
  bold,
  italic,
  color,
  align,
  autoFocus = false,
}: AutoGrowTextareaProps) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useLayoutEffect(() => {
    const el = textareaRef.current;
    if (!el) {
      return;
    }
    el.style.height = '0px';
    el.style.height = `${Math.max(el.scrollHeight, 24)}px`;
  }, [value, fontSize, bold, italic, align]);

  useEffect(() => {
    const el = textareaRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const observer = new ResizeObserver(() => {
      el.style.height = '0px';
      el.style.height = `${Math.max(el.scrollHeight, 24)}px`;
    });
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, []);

  const scaledFontSize = `clamp(${Math.max(10, Math.round(fontSize * 0.42))}px, ${((fontSize / 960) * 100).toFixed(2)}cqw, ${fontSize}px)`;

  return (
    <textarea
      ref={textareaRef}
      value={value}
      onChange={(event) => onChange(event.target.value)}
      onFocus={onFocus}
      disabled={disabled}
      autoFocus={autoFocus}
      rows={1}
      placeholder={placeholder}
      aria-label={ariaLabel}
      style={{
        fontSize: scaledFontSize,
        color,
        textAlign: align,
        fontWeight: bold ? 700 : 400,
        fontStyle: italic ? 'italic' : 'normal',
      }}
      className="w-full bg-transparent border-0 outline-none resize-none overflow-hidden leading-[1.35] tracking-tight placeholder:opacity-45 disabled:pointer-events-none"
    />
  );
}

export function SlideWorkspace({
  document,
  onBackHome,
  onOpenModal,
  onSaveDocument,
  isModalOpen = false,
  autoSaveEnabled = false,
}: SlideWorkspaceProps) {
  const [slides, setSlides] = useState<SlidePreviewItem[]>(() =>
    initializeDeckSlides(document)
  );
  const [activeSlideIdx, setActiveSlideIdx] = useState<number>(0);
  const [selectedBlockId, setSelectedBlockId] = useState<string | null>(null);
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
  const [isNotesExpanded, setIsNotesExpanded] = useState<boolean>(true);
  const [isDirty, setIsDirty] = useState<boolean>(false);
  const [saveFeedback, setSaveFeedback] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [exportError, setExportError] = useState<string | null>(null);

  const loadedDocIdRef = useRef<string>(document.id);
  const uidRef = useRef<number>(1);
  const thumbnailRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const saveTimerRef = useRef<number | null>(null);
  const exportErrorTimerRef = useRef<number | null>(null);
  const documentRef = useRef<DocumentItem>(document);
  documentRef.current = document;

  // Reset per-document state strictly when opening another presentation (never on Save)
  useEffect(() => {
    if (loadedDocIdRef.current === document.id) {
      return;
    }
    loadedDocIdRef.current = document.id;
    setSlides(initializeDeckSlides(document));
    setActiveSlideIdx(0);
    setSelectedBlockId(null);
    setNewlyAddedBlockId(null);
    setIsLaserActive(false);
    setLaserPos({ x: 50, y: 50, visible: false });
    setElapsedSeconds(0);
    setIsTimerRunning(false);
    setIsMobileDrawerOpen(false);
    setIsDirty(false);
    setSaveFeedback(false);
    setExportError(null);
  }, [document]);

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
    }
  }, [isLaserActive]);

  // Cleanup timers on unmount
  useEffect(() => {
    return () => {
      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
      if (exportErrorTimerRef.current !== null) {
        window.clearTimeout(exportErrorTimerRef.current);
      }
    };
  }, []);

  const safeSlideIdx =
    slides.length > 0 ? Math.min(activeSlideIdx, slides.length - 1) : 0;
  const currentSlide = slides[safeSlideIdx];
  const currentBlocks = currentSlide.blocks ?? [];
  const selectedBlock =
    currentBlocks.find((block) => block.id === selectedBlockId) ?? null;
  const accentColor = currentSlide.accentColor || '#4F46E5';
  const hasLightSlideBg = isLightHexBackground(currentSlide.backgroundColor);

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
  }, [buildLiveDocument, onSaveDocument]);

  // Auto-save with 1800ms debounce when autoSaveEnabled is on (resets on every slide edit)
  useEffect(() => {
    if (!autoSaveEnabled || !isDirty) {
      return;
    }
    const timerId = window.setTimeout(() => {
      handleSave();
    }, 1800);
    return () => {
      window.clearTimeout(timerId);
    };
  }, [autoSaveEnabled, isDirty, slides, handleSave]);

  // Keyboard navigation + Ctrl/Cmd+S save (blocked while StageActionModal is open)
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (isModalOpen) {
        return;
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        handleSave();
        return;
      }

      if (event.key === 'Escape') {
        if (isMobileDrawerOpen) {
          setIsMobileDrawerOpen(false);
          return;
        }
        if (isLaserActive) {
          setIsLaserActive(false);
          return;
        }
        if (selectedBlockId !== null) {
          setSelectedBlockId(null);
        }
        return;
      }

      if (isEditableElement(event.target)) {
        return;
      }

      if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        event.preventDefault();
        setSelectedBlockId(null);
        setActiveSlideIdx((prev) =>
          prev < slides.length - 1 ? prev + 1 : prev
        );
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setSelectedBlockId(null);
        setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    slides.length,
    isMobileDrawerOpen,
    isLaserActive,
    isModalOpen,
    selectedBlockId,
    handleSave,
  ]);

  // Keep active slide thumbnail smoothly scrolled into view
  useEffect(() => {
    const node = thumbnailRefs.current[currentSlide.id];
    node?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }, [currentSlide.id]);

  const mutateSlideAt = (
    targetIdx: number,
    updater: (slide: SlidePreviewItem) => SlidePreviewItem
  ) => {
    setSlides((prev) =>
      prev.map((slide, idx) =>
        idx === targetIdx ? syncSlideSummary(updater(slide), idx) : slide
      )
    );
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleAddSlide = () => {
    uidRef.current += 1;
    const newSlideId = `${document.id}-custom-slide-${Date.now()}-${uidRef.current}`;
    const newTitleBlockId = `${newSlideId}-blk-title`;
    const newSlide: SlidePreviewItem = syncSlideSummary(
      {
        id: newSlideId,
        index: safeSlideIdx + 2,
        title: 'Tiêu đề slide mới',
        bullets: ['Nhập nội dung chính tại đây...'],
        accentColor,
        speakerNote: '',
        blocks: [
          {
            id: newTitleBlockId,
            text: 'Tiêu đề slide mới',
            fontSize: 32,
            bold: true,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'title',
          },
          {
            id: `${newSlideId}-blk-body`,
            text: '• Nhập ý chính đầu tiên\n• Nhập ý chính tiếp theo',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#E2E8F0',
            align: 'left',
            role: 'body',
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
    setIsDirty(true);
    setSaveFeedback(false);
  };

  const handleAddTextBlock = () => {
    uidRef.current += 1;
    const newBlockId = `${currentSlide.id}-blk-new-${Date.now()}-${uidRef.current}`;
    const defaultTextColor = hasLightSlideBg ? '#1E293B' : '#F1F5F9';
    const newBlock: SlideTextBlock = {
      id: newBlockId,
      text: 'Khung nội dung mới',
      fontSize: 18,
      bold: false,
      italic: false,
      color: defaultTextColor,
      align: 'left',
      role: 'body',
    };

    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: [...(slide.blocks ?? []), newBlock],
    }));
    setIsLaserActive(false);
    setSelectedBlockId(newBlockId);
    setNewlyAddedBlockId(newBlockId);
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

  const handleDeleteBlock = (blockId: string) => {
    mutateSlideAt(safeSlideIdx, (slide) => ({
      ...slide,
      blocks: (slide.blocks ?? []).filter((block) => block.id !== blockId),
    }));
    if (selectedBlockId === blockId) {
      setSelectedBlockId(null);
    }
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
      await exportSlidesToPptx(document.title, syncedSlides);
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
    setNewlyAddedBlockId(null);
    setActiveSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
  };

  const handleNextSlide = () => {
    setSelectedBlockId(null);
    setNewlyAddedBlockId(null);
    setActiveSlideIdx((prev) => (prev < slides.length - 1 ? prev + 1 : prev));
  };

  const handleSelectSlide = (idx: number) => {
    setSelectedBlockId(null);
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
    const x = Math.min(
      100,
      Math.max(0, ((event.clientX - rect.left) / rect.width) * 100)
    );
    const y = Math.min(
      100,
      Math.max(0, ((event.clientY - rect.top) / rect.height) * 100)
    );
    setLaserPos({ x, y, visible: true });
  };

  const handlePointerLeaveStage = () => {
    if (laserPos.visible) {
      setLaserPos((prev) => ({ ...prev, visible: false }));
    }
  };

  const renderSlideActionButtons = () => (
    <div className="shrink-0 p-2.5 border-b border-white/10 grid grid-cols-3 gap-1.5 bg-[#111620]">
      <button
        type="button"
        onClick={handleAddSlide}
        title="Thêm slide mới"
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-indigo-600/90 hover:bg-indigo-500 text-[11px] font-semibold text-white transition-colors whitespace-nowrap"
      >
        <Plus className="h-3.5 w-3.5 shrink-0" />
        <span>Thêm</span>
      </button>
      <button
        type="button"
        onClick={handleDuplicateSlide}
        title="Nhân bản slide đang chọn"
        className="inline-flex items-center justify-center gap-1 min-h-[34px] px-2 rounded-lg bg-white/[0.07] hover:bg-white/[0.14] border border-white/10 text-[11px] font-medium text-slate-200 transition-colors whitespace-nowrap"
      >
        <Copy className="h-3 w-3 shrink-0" />
        <span>Nhân bản</span>
      </button>
      <button
        type="button"
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
              <div
                className={`aspect-video w-full rounded-lg overflow-hidden border p-2 flex flex-col justify-between transition-colors ${
                  isActive
                    ? 'border-indigo-400/80 shadow-md shadow-indigo-950/50'
                    : 'border-white/10 group-hover:border-white/25'
                }`}
                style={
                  slide.backgroundColor
                    ? { backgroundColor: slide.backgroundColor }
                    : {
                        backgroundImage: `radial-gradient(circle at 86% 18%, ${slideAccent}36 0%, transparent 52%), linear-gradient(145deg, #171E2E 0%, #20293E 60%, #131926 100%)`,
                      }
                }
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

  const hasImages = Boolean(currentSlide.images && currentSlide.images.length > 0);
  const hasMediaColumn = hasImages || Boolean(currentSlide.metricHighlight);

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#181D27] text-slate-100 overflow-hidden">
      {/* Shared Top Workspace Header */}
      <WorkspaceTopBar
        document={document}
        theme="dark"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument())}
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
              <span className="hidden 2xl:inline">Điều khiển QR</span>
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

            {/* Primary Indigo CTA: Fullscreen Presentation with Live Edited Slides */}
            <button
              type="button"
              onClick={() => onOpenModal('present', buildLiveDocument())}
              aria-label="Trình chiếu toàn màn hình"
              title="Trình chiếu toàn màn hình (chiếu ngay nội dung mới nhất)"
              className="shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white shadow-sm shadow-indigo-950/60 transition-all whitespace-nowrap"
            >
              <Play className="h-3.5 w-3.5 fill-white text-white shrink-0" />
              <span className="hidden sm:inline xl:hidden">Trình chiếu</span>
              <span className="hidden xl:inline">
                Trình chiếu toàn màn hình
              </span>
            </button>
          </>
        }
      />

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
          <div className="flex-1 min-h-0 overflow-y-auto scrollbar-dark">
            {renderThumbnailList('desktop')}
          </div>
        </aside>

        {/* Center Stage Editor + Bottom Speaker Notes */}
        <main className="flex-1 min-w-0 flex flex-col bg-[#181D27] overflow-hidden relative">
          {/* Slide Text Block Editing Toolbar */}
          <div className="shrink-0 bg-[#141923]/95 border-b border-white/10 px-3 sm:px-5 py-1.5">
            <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-h-[34px]">
              {/* + Thêm khung chữ */}
              <button
                type="button"
                onClick={handleAddTextBlock}
                className="shrink-0 inline-flex items-center gap-1.5 min-h-[32px] px-2.5 rounded-lg bg-indigo-500/20 hover:bg-indigo-500/30 border border-indigo-400/35 text-xs font-semibold text-indigo-200 transition-colors whitespace-nowrap"
              >
                <Plus className="h-3.5 w-3.5 shrink-0" />
                <Type className="h-3.5 w-3.5 shrink-0" />
                <span>Thêm khung chữ</span>
              </button>

              {selectedBlock ? (
                <>
                  <span
                    aria-hidden="true"
                    className="h-5 w-px mx-1 bg-white/10 shrink-0"
                  />

                  {/* Font Size Picker: 14, 18, 24, 32, 40 */}
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

                  {/* Bold & Italic */}
                  <button
                    type="button"
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

                  <span
                    aria-hidden="true"
                    className="h-5 w-px mx-0.5 bg-white/10 shrink-0"
                  />

                  {/* Color Swatches + Custom Color Input */}
                  <div className="inline-flex items-center gap-1 px-1.5 py-1 rounded-lg bg-[#1C2331] border border-white/10 shrink-0">
                    {SLIDE_COLOR_SWATCHES.map((swatch) => {
                      const isColorActive =
                        (selectedBlock.color ?? '#FFFFFF').toLowerCase() ===
                        swatch.value.toLowerCase();
                      return (
                        <button
                          key={swatch.value}
                          type="button"
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
                    onClick={() => handleDeleteBlock(selectedBlock.id)}
                    aria-label="Xóa khung chữ đang chọn"
                    title="Xóa khung chữ"
                    className="shrink-0 inline-flex items-center gap-1 min-h-[32px] px-2.5 rounded-lg bg-rose-500/15 hover:bg-rose-500/25 border border-rose-400/35 text-xs font-medium text-rose-200 transition-colors whitespace-nowrap"
                  >
                    <Trash2 className="h-3.5 w-3.5 shrink-0" />
                    <span className="hidden lg:inline">Xóa khung</span>
                  </button>
                </>
              ) : (
                <p className="text-[11.5px] text-slate-400 truncate pl-1.5">
                  Bấm trực tiếp vào bất kỳ khung chữ nào trên slide 16:9 để gõ
                  sửa, đổi cỡ chữ, màu sắc và căn lề.
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
                  onClick={() => setIsLaserActive(false)}
                  className="ml-1 px-2 py-0.5 rounded-full bg-rose-500/25 hover:bg-rose-500/40 text-[11px] font-semibold text-white transition-colors shrink-0"
                >
                  Tắt
                </button>
              </div>
            </div>
          )}

          {/* Main 16:9 Editable Slide Stage Container */}
          <div className="flex-1 min-h-0 flex flex-col px-3 sm:px-6 lg:px-10 pt-3 sm:pt-4 pb-2 overflow-hidden">
            <div
              className="flex-1 min-h-0 w-full flex items-center justify-center relative"
              style={{ containerType: 'size' }}
            >
              <div
                onClick={() => {
                  if (!isLaserActive) {
                    setSelectedBlockId(null);
                  }
                }}
                onPointerMove={updateLaserFromPointer}
                onPointerDown={updateLaserFromPointer}
                onPointerLeave={handlePointerLeaveStage}
                style={{
                  width: 'min(100cqw, calc(100cqh * 16 / 9), 1080px)',
                }}
                className={`aspect-video relative rounded-2xl overflow-hidden border shadow-2xl shadow-black/55 transition-shadow duration-200 ${
                  isLaserActive
                    ? 'border-rose-400/55 ring-2 ring-rose-500/25 cursor-none touch-none select-none'
                    : 'border-white/15 ring-1 ring-white/[0.06]'
                }`}
              >
                {/* Inner Proportional Slide Canvas (containerType: inline-size for WYSIWYG font scale) */}
                <div
                  className="absolute inset-0 flex flex-col justify-between p-4 sm:p-7 md:p-9 lg:p-10"
                  style={{
                    containerType: 'inline-size',
                    ...(currentSlide.backgroundColor
                      ? { backgroundColor: currentSlide.backgroundColor }
                      : {
                          backgroundImage: `radial-gradient(circle at 86% 16%, ${accentColor}3A 0%, transparent 46%), radial-gradient(circle at 10% 90%, ${accentColor}1C 0%, transparent 40%), linear-gradient(145deg, #161D2E 0%, #1F293F 56%, #131926 100%)`,
                        }),
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

                  {/* Main Editable Blocks + Real Images / Metric Area */}
                  <div
                    className={`flex-1 min-h-0 overflow-y-auto scrollbar-dark pr-1 grid gap-3 sm:gap-5 items-center ${
                      hasMediaColumn && currentBlocks.length > 0
                        ? 'grid-cols-12'
                        : 'grid-cols-1'
                    }`}
                  >
                    {/* Editable SlideTextBlock Column */}
                    {currentBlocks.length > 0 && (
                      <div
                        className={`space-y-2 sm:space-y-3 ${
                          hasMediaColumn ? 'col-span-7 sm:col-span-8' : 'col-span-1'
                        }`}
                      >
                        {currentBlocks.map((block) => {
                          const isSelected =
                            !isLaserActive && selectedBlockId === block.id;
                          return (
                            <div
                              key={block.id}
                              onClick={(event) => {
                                if (isLaserActive) {
                                  return;
                                }
                                event.stopPropagation();
                                setSelectedBlockId(block.id);
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
                                fontSize={block.fontSize}
                                bold={Boolean(block.bold)}
                                italic={Boolean(block.italic)}
                                color={
                                  block.color ??
                                  (hasLightSlideBg ? '#0F172A' : '#FFFFFF')
                                }
                                align={block.align ?? 'left'}
                              />
                            </div>
                          );
                        })}
                      </div>
                    )}

                    {/* Real Extracted PPTX Images Gallery */}
                    {hasImages && currentSlide.images && (
                      <div
                        className={`${
                          currentBlocks.length > 0
                            ? 'col-span-5 sm:col-span-4'
                            : 'col-span-1 max-w-2xl mx-auto w-full'
                        } grid gap-2 ${
                          currentSlide.images.length > 1
                            ? 'grid-cols-2'
                            : 'grid-cols-1'
                        }`}
                      >
                        {currentSlide.images.map((imageItem, imgIdx) => (
                          <figure
                            key={imageItem.id}
                            className="rounded-xl overflow-hidden border border-white/15 bg-black/25 flex items-center justify-center p-1.5"
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
                        ))}
                      </div>
                    )}

                    {/* Metric Highlight Card (for mock decks without embedded images) */}
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

                    {/* Empty Slide Fallback Hint */}
                    {currentBlocks.length === 0 && !hasImages && (
                      <div className="col-span-1 flex flex-col items-center justify-center text-center py-8 opacity-75">
                        <Type className="h-7 w-7 mb-2 text-indigo-300" />
                        <p className="text-xs sm:text-sm font-medium">
                          Slide đang trống — bấm <strong>+ Thêm khung chữ</strong>{' '}
                          phía trên để bắt đầu soạn nội dung.
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

                {/* Live Glowing Red Laser Pointer Layer */}
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
                      className="cursor-pointer absolute left-2.5 sm:left-3.5 top-1/2 -translate-y-1/2 z-10 h-9 w-9 sm:h-10 sm:w-10 rounded-full bg-black/45 hover:bg-black/75 border border-white/15 text-white flex items-center justify-center backdrop-blur-sm transition disabled:opacity-25 disabled:cursor-not-allowed"
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
                      className="cursor-pointer absolute right-2.5 sm:right-3.5 top-1/2 -translate-y-1/2 z-10 h-9 w-9 sm:h-10 sm:w-10 rounded-full bg-black/45 hover:bg-black/75 border border-white/15 text-white flex items-center justify-center backdrop-blur-sm transition disabled:opacity-25 disabled:cursor-not-allowed"
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

          {/* Bottom Speaker Notes Dock (Collapsible + Synced to Slide State) */}
          <section
            aria-label="Ghi chú diễn giả"
            className="shrink-0 border-t border-white/10 bg-[#141923] pb-safe"
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
          </section>
        </main>
      </div>
    </div>
  );
}

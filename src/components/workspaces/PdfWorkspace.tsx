import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { ChangeEvent, KeyboardEvent as ReactKeyboardEvent } from 'react';
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  CloudDownload,
  Download,
  ExternalLink,
  Eye,
  FileText,
  Lightbulb,
  Loader2,
  PanelLeft,
  Presentation,
  Search,
  Sparkles,
  Upload,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {
  DocumentItem,
  ModalType,
  PdfPageItem,
} from '../../types/document';
import {
  loadPdfDocumentFromSource,
  renderPdfPageTextLayer,
  renderPdfPageToCanvas,
  renderPdfThumbnailDataUrl,
} from '../../utils/pdfEngine';
import type { PdfDocumentProxy } from '../../utils/pdfEngine';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { WorkspaceTopBar } from './WorkspaceTopBar';

export interface PdfWorkspaceProps {
  document: DocumentItem;
  onBackHome: () => void;
  onOpenModal: (
    modal: Exclude<ModalType, null>,
    overrideDoc?: DocumentItem
  ) => void;
  isModalOpen?: boolean;
  onReplacePdfFile?: (docId: string, file: File) => void;
  isHydratingCloudFile?: boolean;
  onRetryCloudDownload?: () => void;
}

const ZOOM_LEVELS = [75, 90, 100, 115, 130, 150] as const;
const DEFAULT_ZOOM_INDEX = 2; // 100%
const BASE_PAGE_WIDTH_PX = 760;
const THUMBNAIL_WIDTH_PX = 160;

type PdfLoadState = 'idle' | 'loading' | 'ready' | 'error';

function getPdfDownloadFilename(title: string): string {
  const cleaned = title.trim() || 'tai-lieu';
  return cleaned.toLowerCase().endsWith('.pdf') ? cleaned : `${cleaned}.pdf`;
}

function normalizeSearchText(text: string): string {
  return text.toLowerCase().replace(/\s+/g, ' ').trim();
}

function applyTextLayerSearchHighlight(
  container: HTMLElement | null,
  rawQuery: string
): void {
  if (!container) {
    return;
  }
  const query = normalizeSearchText(rawQuery);
  const spans = container.querySelectorAll<HTMLSpanElement>('span');
  for (const span of spans) {
    const content = normalizeSearchText(span.textContent ?? '');
    const isMatch = query.length > 0 && content.length > 0 && content.includes(query);
    if (isMatch) {
      span.style.backgroundColor = 'rgba(250, 204, 21, 0.42)';
      span.style.borderRadius = '2px';
    } else {
      span.style.backgroundColor = '';
      span.style.borderRadius = '';
    }
  }
}

function buildFallbackSamplePages(document: DocumentItem): PdfPageItem[] {
  const preview = document.contentPreview.trim();

  if (document.outline.length > 0) {
    return document.outline.map((section, idx) => {
      const paragraphs: string[] = [];
      if (idx === 0 && preview && preview !== section.summary.trim()) {
        paragraphs.push(preview);
      }
      if (section.summary.trim()) {
        paragraphs.push(section.summary.trim());
      }
      if (paragraphs.length === 0) {
        paragraphs.push(
          `Nội dung tổng hợp cho phần "${section.title}" trong tài liệu ${document.title}.`
        );
      }

      const sectionLabel = section.romanIndex
        ? `${section.romanIndex}. ${section.title}`
        : section.title;

      return {
        id: `${document.id}-sample-pdf-${section.id}`,
        pageNumber: idx + 1,
        sectionTitle: sectionLabel,
        heading: section.title,
        paragraphs,
        callout:
          idx === 0
            ? `Tài liệu thuộc nhóm "${document.folder || 'Thư viện PDF'}" • Cập nhật: ${document.updatedAt}.`
            : undefined,
        highlights:
          section.keyPoints.length > 0 ? section.keyPoints : undefined,
      };
    });
  }

  return [
    {
      id: `${document.id}-sample-pdf-1`,
      pageNumber: 1,
      sectionTitle: 'I. Tổng quan tài liệu',
      heading: document.title,
      paragraphs: [
        preview ||
          'Tài liệu PDF sẵn sàng để đọc và trình chiếu trong không gian làm việc tập trung LITE.',
      ],
      highlights: [
        `Định dạng: PDF${document.fileSize ? ` • Dung lượng ${document.fileSize}` : ''}`,
        `Cập nhật lần cuối: ${document.updatedAt}`,
      ],
    },
  ];
}

export function PdfWorkspace({
  document,
  onBackHome,
  onOpenModal,
  isModalOpen = false,
  onReplacePdfFile,
  isHydratingCloudFile = false,
  onRetryCloudDownload,
}: PdfWorkspaceProps) {
  const realPdfUrl = document.fileUrl ?? document.objectUrl ?? '';
  const hasCloudBacking = Boolean(document.storagePath || document.r2ObjectKey);
  const isRealPdfMode = Boolean(
    document.isUploaded || realPdfUrl || hasCloudBacking
  );

  const samplePages = useMemo<PdfPageItem[]>(() => {
    if (isRealPdfMode) {
      return [];
    }
    return document.pdfPages && document.pdfPages.length > 0
      ? document.pdfPages
      : buildFallbackSamplePages(document);
  }, [document, isRealPdfMode]);

  // Shared reader UI state
  const [currentPageIdx, setCurrentPageIdx] = useState(0);
  const [zoomIndex, setZoomIndex] = useState(DEFAULT_ZOOM_INDEX);
  const [isMobileDrawerOpen, setIsMobileDrawerOpen] = useState(false);
  const [isDesktopSidebarOpen, setIsDesktopSidebarOpen] = useState(true);
  const [replaceError, setReplaceError] = useState<string | null>(null);

  // Real PDF.js engine state
  const [pdfDoc, setPdfDoc] = useState<PdfDocumentProxy | null>(null);
  const [pdfNumPages, setPdfNumPages] = useState(0);
  const [pdfLoadState, setPdfLoadState] = useState<PdfLoadState>('idle');
  const [pdfLoadError, setPdfLoadError] = useState<string | null>(null);
  const [useBrowserFallback, setUseBrowserFallback] = useState(false);
  const [isRenderingPage, setIsRenderingPage] = useState(false);
  const [renderedSize, setRenderedSize] = useState<{
    width: number;
    height: number;
  } | null>(null);
  const [viewerWidthPx, setViewerWidthPx] = useState(BASE_PAGE_WIDTH_PX);
  const [thumbnails, setThumbnails] = useState<Record<number, string>>({});
  const [thumbPriorityPages, setThumbPriorityPages] = useState<number[]>([]);

  // In-PDF search state
  const [isSearchOpen, setIsSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isSearchingPdf, setIsSearchingPdf] = useState(false);
  const [searchMatchPages, setSearchMatchPages] = useState<number[]>([]);
  const [activeMatchIdx, setActiveMatchIdx] = useState(0);

  // Refs
  const viewerOuterRef = useRef<HTMLDivElement | null>(null);
  const viewerScrollRef = useRef<HTMLElement | null>(null);
  const activeThumbRef = useRef<HTMLButtonElement | null>(null);
  const replaceInputRef = useRef<HTMLInputElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const textLayerRef = useRef<HTMLDivElement | null>(null);
  const searchInputRef = useRef<HTMLInputElement | null>(null);

  const renderChainRef = useRef<Promise<void>>(Promise.resolve());
  const renderTokenRef = useRef(0);
  const activeCanvasCancelRef = useRef<(() => void) | null>(null);
  const activeTextCancelRef = useRef<(() => void) | null>(null);
  const thumbnailsCacheRef = useRef<Record<number, string>>({});
  const pageTextCacheRef = useRef<Map<number, string>>(new Map());
  const thumbObserverRef = useRef<IntersectionObserver | null>(null);
  const thumbNodeMapRef = useRef<Map<number, HTMLButtonElement>>(new Map());

  // Reset UI state when switching documents
  useEffect(() => {
    setCurrentPageIdx(0);
    setZoomIndex(DEFAULT_ZOOM_INDEX);
    setIsMobileDrawerOpen(false);
    setReplaceError(null);
    setUseBrowserFallback(false);
    setIsSearchOpen(false);
    setSearchQuery('');
    setSearchMatchPages([]);
    setActiveMatchIdx(0);
  }, [document.id]);

  // Measure outer non-scrolling reading pane width to fit Canvas cleanly on mobile/tablet without scrollbar loops
  useEffect(() => {
    const el = viewerOuterRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }

    const updateWidth = (width: number) => {
      if (width <= 0) {
        return;
      }
      setViewerWidthPx((prev) =>
        Math.abs(width - prev) >= 16 ? Math.round(width) : prev
      );
    };

    updateWidth(el.clientWidth);

    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        updateWidth(entry.contentRect.width);
      }
    });
    observer.observe(el);

    return () => {
      observer.disconnect();
    };
  }, [isRealPdfMode, realPdfUrl, isDesktopSidebarOpen]);

  // Load real PDF via PDF.js whenever realPdfUrl changes
  useEffect(() => {
    if (!isRealPdfMode || !realPdfUrl) {
      setPdfDoc(null);
      setPdfNumPages(0);
      setPdfLoadState('idle');
      setPdfLoadError(null);
      setRenderedSize(null);
      setThumbnails({});
      thumbnailsCacheRef.current = {};
      pageTextCacheRef.current.clear();
      return;
    }

    let cancelled = false;
    let loadedProxy: PdfDocumentProxy | null = null;

    setPdfDoc(null);
    setPdfNumPages(0);
    setPdfLoadState('loading');
    setPdfLoadError(null);
    setUseBrowserFallback(false);
    setRenderedSize(null);
    setThumbnails({});
    setThumbPriorityPages([]);
    thumbnailsCacheRef.current = {};
    pageTextCacheRef.current.clear();
    setSearchMatchPages([]);
    setActiveMatchIdx(0);

    loadPdfDocumentFromSource(realPdfUrl)
      .then((doc) => {
        if (cancelled) {
          void doc.destroy().catch(() => undefined);
          return;
        }
        loadedProxy = doc;
        setPdfDoc(doc);
        setPdfNumPages(doc.numPages);
        setPdfLoadState('ready');
        setCurrentPageIdx((prev) =>
          Math.max(0, Math.min(doc.numPages - 1, prev))
        );
      })
      .catch((error: unknown) => {
        if (cancelled) {
          return;
        }
        const message =
          error instanceof Error && error.message
            ? error.message
            : 'Không thể giải mã tệp PDF này bằng trình dựng tích hợp.';
        setPdfLoadError(message);
        setPdfLoadState('error');
      });

    return () => {
      cancelled = true;
      renderTokenRef.current += 1;
      activeCanvasCancelRef.current?.();
      activeCanvasCancelRef.current = null;
      activeTextCancelRef.current?.();
      activeTextCancelRef.current = null;
      if (loadedProxy) {
        void loadedProxy.destroy().catch(() => undefined);
      }
    };
  }, [isRealPdfMode, realPdfUrl]);

  const totalPages = isRealPdfMode ? pdfNumPages : samplePages.length;
  const safePageIdx =
    totalPages > 0 ? Math.min(currentPageIdx, totalPages - 1) : 0;
  const currentPageNumber = safePageIdx + 1;
  const currentSamplePage = !isRealPdfMode ? samplePages[safePageIdx] : undefined;
  const zoomPercent = ZOOM_LEVELS[zoomIndex];
  const zoomScale = zoomPercent / 100;

  const canGoPrev = safePageIdx > 0;
  const canGoNext = totalPages > 0 && safePageIdx < totalPages - 1;
  const canZoomOut = zoomIndex > 0;
  const canZoomIn = zoomIndex < ZOOM_LEVELS.length - 1;

  const goToPage = useCallback(
    (index: number) => {
      if (totalPages <= 0) {
        return;
      }
      const clamped = Math.max(0, Math.min(totalPages - 1, index));
      setCurrentPageIdx(clamped);
      setIsMobileDrawerOpen(false);
    },
    [totalPages]
  );

  const handlePrevPage = useCallback(() => {
    setCurrentPageIdx((prev) => Math.max(0, prev - 1));
  }, []);

  const handleNextPage = useCallback(() => {
    setCurrentPageIdx((prev) =>
      totalPages > 0 ? Math.min(totalPages - 1, prev + 1) : 0
    );
  }, [totalPages]);

  const handleZoomOut = () => {
    setZoomIndex((prev) => Math.max(0, prev - 1));
  };

  const handleZoomIn = () => {
    setZoomIndex((prev) => Math.min(ZOOM_LEVELS.length - 1, prev + 1));
  };

  const handleResetZoom = () => {
    setZoomIndex(DEFAULT_ZOOM_INDEX);
  };

  const handleTriggerReplaceFile = () => {
    setReplaceError(null);
    replaceInputRef.current?.click();
  };

  const handleReplaceFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) {
      return;
    }
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.pdf') && file.type !== 'application/pdf') {
      setReplaceError('Vui lòng chọn đúng tệp định dạng .pdf để mở lại.');
      return;
    }
    setReplaceError(null);
    onReplacePdfFile?.(document.id, file);
  };

  // Focus search input when opened
  useEffect(() => {
    if (isSearchOpen) {
      window.requestAnimationFrame(() => {
        searchInputRef.current?.focus();
        searchInputRef.current?.select();
      });
    }
  }, [isSearchOpen]);

  // Render active PDF.js page onto Canvas + synchronized selectable TextLayer
  useEffect(() => {
    if (
      !isRealPdfMode ||
      !pdfDoc ||
      pdfLoadState !== 'ready' ||
      useBrowserFallback ||
      totalPages <= 0
    ) {
      return;
    }

    const canvas = canvasRef.current;
    const textLayerEl = textLayerRef.current;
    if (!canvas || !textLayerEl) {
      return;
    }

    const token = ++renderTokenRef.current;
    activeCanvasCancelRef.current?.();
    activeCanvasCancelRef.current = null;
    activeTextCancelRef.current?.();
    activeTextCancelRef.current = null;

    const horizontalPadding = viewerWidthPx < 640 ? 24 : 56;
    const baseFitWidth = Math.min(
      BASE_PAGE_WIDTH_PX,
      Math.max(260, viewerWidthPx - horizontalPadding)
    );
    const targetWidthCss = Math.max(220, Math.round(baseFitWidth * zoomScale));

    setIsRenderingPage(true);

    renderChainRef.current = renderChainRef.current
      .catch(() => undefined)
      .then(async () => {
        if (token !== renderTokenRef.current) {
          return;
        }

        try {
          const page = await pdfDoc.getPage(currentPageNumber);
          if (token !== renderTokenRef.current) {
            return;
          }

          const baseViewport = page.getViewport({ scale: 1 });
          const renderResult = await renderPdfPageToCanvas({
            pdfDoc,
            pageNumber: currentPageNumber,
            canvas,
            targetWidthCss,
            maxDpr: 2,
          });
          activeCanvasCancelRef.current = renderResult.cancel;

          if (token !== renderTokenRef.current) {
            renderResult.cancel();
            return;
          }

          setRenderedSize({
            width: renderResult.widthCss,
            height: renderResult.heightCss,
          });
          setIsRenderingPage(false);

          const effectiveScale =
            baseViewport.width > 0
              ? renderResult.widthCss / baseViewport.width
              : zoomScale;

          textLayerEl.replaceChildren();
          textLayerEl.style.setProperty(
            '--scale-factor',
            String(effectiveScale)
          );
          textLayerEl.style.setProperty(
            '--total-scale-factor',
            String(effectiveScale)
          );

          const cancelTextLayer = await renderPdfPageTextLayer({
            pdfDoc,
            pageNumber: currentPageNumber,
            container: textLayerEl,
            viewportScale: effectiveScale,
          });
          activeTextCancelRef.current = cancelTextLayer;

          if (token !== renderTokenRef.current) {
            cancelTextLayer();
            return;
          }

          applyTextLayerSearchHighlight(textLayerEl, searchQuery);
        } catch (err: unknown) {
          if (token !== renderTokenRef.current) {
            return;
          }
          const errName = err instanceof Error ? err.name : '';
          if (errName !== 'RenderingCancelledException') {
            setIsRenderingPage(false);
          }
        }
      });

    return () => {
      if (renderTokenRef.current === token) {
        renderTokenRef.current += 1;
      }
      activeCanvasCancelRef.current?.();
      activeCanvasCancelRef.current = null;
      activeTextCancelRef.current?.();
      activeTextCancelRef.current = null;
    };
  }, [
    currentPageNumber,
    isRealPdfMode,
    pdfDoc,
    pdfLoadState,
    totalPages,
    useBrowserFallback,
    viewerWidthPx,
    zoomScale,
  ]);

  // Re-apply text highlight immediately when searchQuery changes on the already-rendered page
  useEffect(() => {
    if (!isRealPdfMode || useBrowserFallback) {
      return;
    }
    applyTextLayerSearchHighlight(textLayerRef.current, searchQuery);
  }, [isRealPdfMode, searchQuery, useBrowserFallback]);

  // Observe visible thumbnail slots for lazy PDF.js thumbnail generation
  useEffect(() => {
    if (
      !isRealPdfMode ||
      !pdfDoc ||
      pdfLoadState !== 'ready' ||
      useBrowserFallback ||
      (!isDesktopSidebarOpen && !isMobileDrawerOpen) ||
      typeof IntersectionObserver === 'undefined'
    ) {
      return;
    }

    const observer = new IntersectionObserver(
      (entries) => {
        const newlyVisible: number[] = [];
        for (const entry of entries) {
          if (!entry.isIntersecting) {
            continue;
          }
          const pageAttr = entry.target.getAttribute('data-pdf-thumb-page');
          const pageNum = pageAttr ? Number(pageAttr) : NaN;
          if (
            Number.isInteger(pageNum) &&
            pageNum >= 1 &&
            !thumbnailsCacheRef.current[pageNum]
          ) {
            newlyVisible.push(pageNum);
          }
        }
        if (newlyVisible.length > 0) {
          setThumbPriorityPages((prev) => {
            const merged = new Set(prev);
            let added = false;
            for (const p of newlyVisible) {
              if (!merged.has(p)) {
                merged.add(p);
                added = true;
              }
            }
            return added ? Array.from(merged) : prev;
          });
        }
      },
      { rootMargin: '180px 0px' }
    );

    thumbObserverRef.current = observer;
    for (const node of thumbNodeMapRef.current.values()) {
      observer.observe(node);
    }

    return () => {
      observer.disconnect();
      thumbObserverRef.current = null;
    };
  }, [
    isDesktopSidebarOpen,
    isMobileDrawerOpen,
    isRealPdfMode,
    pdfDoc,
    pdfLoadState,
    pdfNumPages,
    useBrowserFallback,
  ]);

  // Sequentially lazy-render thumbnails around currentPageNumber + observed visible thumbnails
  useEffect(() => {
    if (
      !isRealPdfMode ||
      !pdfDoc ||
      pdfLoadState !== 'ready' ||
      useBrowserFallback ||
      pdfNumPages <= 0 ||
      (!isDesktopSidebarOpen && !isMobileDrawerOpen)
    ) {
      return;
    }

    let cancelled = false;

    const queue: number[] = [];
    const seen = new Set<number>();
    const enqueue = (pageNum: number) => {
      if (
        pageNum >= 1 &&
        pageNum <= pdfNumPages &&
        !seen.has(pageNum) &&
        !thumbnailsCacheRef.current[pageNum]
      ) {
        seen.add(pageNum);
        queue.push(pageNum);
      }
    };

    // Prioritize current page and immediate neighbors first
    enqueue(currentPageNumber);
    for (let offset = 1; offset <= 5; offset += 1) {
      enqueue(currentPageNumber + offset);
      enqueue(currentPageNumber - offset);
    }
    for (const p of thumbPriorityPages) {
      enqueue(p);
    }

    if (queue.length === 0) {
      return;
    }

    const runQueue = async () => {
      for (const pageNum of queue) {
        if (cancelled) {
          return;
        }
        if (thumbnailsCacheRef.current[pageNum]) {
          continue;
        }
        try {
          const dataUrl = await renderPdfThumbnailDataUrl(
            pdfDoc,
            pageNum,
            THUMBNAIL_WIDTH_PX
          );
          if (cancelled || !dataUrl) {
            return;
          }
          thumbnailsCacheRef.current[pageNum] = dataUrl;
          setThumbnails((prev) =>
            prev[pageNum] === dataUrl ? prev : { ...prev, [pageNum]: dataUrl }
          );
        } catch {
          // Ignore single thumbnail failures gracefully
        }
      }
    };

    void runQueue();

    return () => {
      cancelled = true;
    };
  }, [
    currentPageNumber,
    isDesktopSidebarOpen,
    isMobileDrawerOpen,
    isRealPdfMode,
    pdfDoc,
    pdfLoadState,
    pdfNumPages,
    thumbPriorityPages,
    useBrowserFallback,
  ]);

  // Full-text keyword search across all pages of the real PDF (cached per page)
  useEffect(() => {
    if (!isRealPdfMode || !pdfDoc || pdfLoadState !== 'ready') {
      return;
    }
    const normalizedQuery = normalizeSearchText(searchQuery);
    if (!normalizedQuery) {
      setSearchMatchPages([]);
      setActiveMatchIdx(0);
      setIsSearchingPdf(false);
      return;
    }

    let cancelled = false;
    setIsSearchingPdf(true);

    const timer = window.setTimeout(() => {
      const scanPages = async () => {
        const matches: number[] = [];
        for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum += 1) {
          if (cancelled) {
            return;
          }
          let cachedText = pageTextCacheRef.current.get(pageNum);
          if (cachedText === undefined) {
            try {
              const page = await pdfDoc.getPage(pageNum);
              if (cancelled) {
                return;
              }
              const textContent = await page.getTextContent();
              if (cancelled) {
                return;
              }
              const joined = textContent.items
                .map((item) =>
                  'str' in item && typeof item.str === 'string' ? item.str : ''
                )
                .join(' ');
              cachedText = normalizeSearchText(joined);
              pageTextCacheRef.current.set(pageNum, cachedText);
            } catch {
              cachedText = '';
            }
          }
          if (cachedText.includes(normalizedQuery)) {
            matches.push(pageNum);
          }
        }

        if (cancelled) {
          return;
        }
        setSearchMatchPages(matches);
        setIsSearchingPdf(false);

        if (matches.length > 0) {
          const existingIdx = matches.indexOf(currentPageNumber);
          if (existingIdx !== -1) {
            setActiveMatchIdx(existingIdx);
          } else {
            setActiveMatchIdx(0);
            setCurrentPageIdx(matches[0] - 1);
          }
        } else {
          setActiveMatchIdx(0);
        }
      };

      void scanPages();
    }, 220);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [isRealPdfMode, pdfDoc, pdfLoadState, searchQuery]);

  const handleJumpMatch = (direction: 'prev' | 'next') => {
    if (searchMatchPages.length === 0) {
      return;
    }
    const nextIdx =
      direction === 'next'
        ? (activeMatchIdx + 1) % searchMatchPages.length
        : (activeMatchIdx - 1 + searchMatchPages.length) %
          searchMatchPages.length;
    setActiveMatchIdx(nextIdx);
    const targetPageNum = searchMatchPages[nextIdx];
    if (targetPageNum) {
      setCurrentPageIdx(targetPageNum - 1);
    }
  };

  const handleSearchInputKeyDown = (
    event: ReactKeyboardEvent<HTMLInputElement>
  ) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      setIsSearchOpen(false);
      setSearchQuery('');
      return;
    }
    if (event.key === 'Enter' && searchMatchPages.length > 0) {
      event.preventDefault();
      handleJumpMatch(event.shiftKey ? 'prev' : 'next');
    }
  };

  // Scroll reader to top and keep active thumbnail visible on page turn
  useEffect(() => {
    if (isRealPdfMode && (!realPdfUrl || useBrowserFallback)) {
      return;
    }
    viewerScrollRef.current?.scrollTo({ top: 0, behavior: 'smooth' });
    activeThumbRef.current?.scrollIntoView({
      block: 'nearest',
      behavior: 'smooth',
    });
  }, [isRealPdfMode, realPdfUrl, safePageIdx, useBrowserFallback]);

  // Global keyboard shortcuts (ArrowLeft / ArrowRight / PageUp / PageDown / Escape)
  useEffect(() => {
    if (isRealPdfMode && (!realPdfUrl || useBrowserFallback)) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (isModalOpen) {
        return;
      }

      const target = event.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === 'INPUT' ||
          target.tagName === 'TEXTAREA' ||
          target.tagName === 'SELECT' ||
          target.isContentEditable)
      ) {
        return;
      }

      if (event.key === 'Escape') {
        if (isMobileDrawerOpen) {
          event.preventDefault();
          setIsMobileDrawerOpen(false);
          return;
        }
        if (isSearchOpen) {
          event.preventDefault();
          setIsSearchOpen(false);
          setSearchQuery('');
          return;
        }
      }

      if (totalPages <= 1) {
        return;
      }

      if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setCurrentPageIdx((prev) => Math.max(0, prev - 1));
      } else if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        event.preventDefault();
        setCurrentPageIdx((prev) => Math.min(totalPages - 1, prev + 1));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [
    isMobileDrawerOpen,
    isModalOpen,
    isRealPdfMode,
    isSearchOpen,
    realPdfUrl,
    totalPages,
    useBrowserFallback,
  ]);

  const registerThumbNode = useCallback(
    (pageNum: number, node: HTMLButtonElement | null) => {
      const prev = thumbNodeMapRef.current.get(pageNum);
      if (prev && prev !== node) {
        thumbObserverRef.current?.unobserve(prev);
      }
      if (node) {
        thumbNodeMapRef.current.set(pageNum, node);
        thumbObserverRef.current?.observe(node);
      } else {
        thumbNodeMapRef.current.delete(pageNum);
      }
    },
    []
  );

  // =========================================================================
  // NHÁNH 1: TỆP PDF THẬT CHƯA CÓ URL (ĐANG TẢI TỪ ĐÁM MÂY HOẶC CẦN CHỌN LẠI)
  // =========================================================================
  if (isRealPdfMode && !realPdfUrl) {
    const canRetryCloud = Boolean(hasCloudBacking && onRetryCloudDownload);

    const missingPdfToolbar = (
      <div className="flex items-center gap-1.5 sm:gap-2">
        {isHydratingCloudFile ? (
          <span className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl bg-indigo-50 border border-indigo-200/80 text-indigo-700 text-xs font-semibold whitespace-nowrap shrink-0">
            <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
            <span className="hidden sm:inline">Đang tải từ đám mây...</span>
            <span className="sm:hidden">Đang tải...</span>
          </span>
        ) : (
          <>
            {canRetryCloud && (
              <button
                type="button"
                onClick={onRetryCloudDownload}
                className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm transition-colors whitespace-nowrap shrink-0"
              >
                <CloudDownload className="w-3.5 h-3.5 shrink-0" />
                <span className="hidden sm:inline">Tải lại PDF từ đám mây</span>
                <span className="sm:hidden">Tải đám mây</span>
              </button>
            )}

            {onReplacePdfFile && (
              <button
                type="button"
                onClick={handleTriggerReplaceFile}
                className={`inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl text-xs font-semibold transition-colors whitespace-nowrap shrink-0 ${
                  canRetryCloud
                    ? 'bg-white hover:bg-[#FAF8F5] border border-[#DDD7CC] text-slate-700 hover:text-pink-700'
                    : 'bg-pink-600 hover:bg-pink-700 text-white shadow-sm'
                }`}
              >
                <Upload className="w-3.5 h-3.5 shrink-0" />
                <span>{canRetryCloud ? 'Mở từ máy' : 'Mở lại tệp PDF'}</span>
              </button>
            )}
          </>
        )}
      </div>
    );

    return (
      <div className="h-[100dvh] w-full flex flex-col bg-[#E8E5DE] text-slate-900 overflow-hidden">
        <WorkspaceTopBar
          document={document}
          theme="light"
          onBackHome={onBackHome}
          onShare={() => onOpenModal('share', document)}
          toolbar={missingPdfToolbar}
        />

        <input
          ref={replaceInputRef}
          type="file"
          accept=".pdf,application/pdf"
          onChange={handleReplaceFileChange}
          className="hidden"
          aria-hidden="true"
        />

        <main className="flex-1 min-h-0 flex flex-col overflow-hidden">
          {isHydratingCloudFile ? (
            <div className="flex-1 min-h-0 overflow-auto p-4 sm:p-8 flex items-center justify-center">
              <div className="w-full max-w-md rounded-2xl bg-white border border-[#E2DDD5] shadow-md p-6 sm:p-8 text-center">
                <div className="w-14 h-14 mx-auto rounded-2xl bg-indigo-50 border border-indigo-200/70 text-indigo-600 flex items-center justify-center mb-4">
                  <Loader2 className="w-7 h-7 animate-spin" />
                </div>

                <h2 className="text-base sm:text-lg font-bold text-slate-900 mb-2">
                  Đang tải tệp PDF từ tài khoản đám mây...
                </h2>

                <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
                  Đang đồng bộ bản PDF gốc{' '}
                  <span className="font-semibold text-slate-800">
                    &ldquo;{document.title}&rdquo;
                  </span>{' '}
                  về thiết bị của bạn. Vui lòng đợi trong giây lát.
                </p>

                <div className="mt-5 pt-4 border-t border-slate-100 flex items-center justify-center gap-2 text-[11.5px] text-slate-400">
                  {document.fileSize && <span>{document.fileSize}</span>}
                  {document.fileSize && <span>•</span>}
                  <span>{document.updatedAt}</span>
                </div>
              </div>
            </div>
          ) : (
            <div className="flex-1 min-h-0 overflow-auto p-4 sm:p-8 flex items-center justify-center">
              <div className="w-full max-w-md rounded-2xl bg-white border border-[#E2DDD5] shadow-md p-6 sm:p-8 text-center">
                <div className="w-14 h-14 mx-auto rounded-2xl bg-pink-50 border border-pink-200/70 text-pink-600 flex items-center justify-center mb-4">
                  {canRetryCloud ? (
                    <CloudDownload className="w-7 h-7 text-indigo-600" />
                  ) : (
                    <FileText className="w-7 h-7" />
                  )}
                </div>

                <h2 className="text-base sm:text-lg font-bold text-slate-900 mb-2">
                  {canRetryCloud
                    ? 'Chưa tải được tệp PDF từ đám mây'
                    : 'Cần chọn lại tệp PDF từ máy'}
                </h2>

                <p className="text-xs sm:text-sm text-slate-600 leading-relaxed mb-5">
                  {canRetryCloud ? (
                    <>
                      Chưa thể nạp bản PDF gốc{' '}
                      <span className="font-semibold text-slate-800">
                        &ldquo;{document.title}&rdquo;
                      </span>{' '}
                      từ bộ nhớ đám mây. Bạn có thể thử tải lại ngay hoặc chọn
                      tệp trực tiếp từ thiết bị.
                    </>
                  ) : (
                    <>
                      Đường dẫn tạm của tệp{' '}
                      <span className="font-semibold text-slate-800">
                        &ldquo;{document.title}&rdquo;
                      </span>{' '}
                      đã được trình duyệt giải phóng sau khi tải lại trang. Chọn
                      lại tệp PDF gốc từ máy để tiếp tục xem ngay.
                    </>
                  )}
                </p>

                {replaceError && (
                  <p className="mb-4 px-3 py-2 rounded-xl bg-rose-50 border border-rose-200 text-xs font-medium text-rose-700">
                    {replaceError}
                  </p>
                )}

                <div className="space-y-2.5">
                  {canRetryCloud && (
                    <button
                      type="button"
                      onClick={onRetryCloudDownload}
                      className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 min-h-[44px] rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold shadow-sm transition-colors"
                    >
                      <CloudDownload className="w-4 h-4 shrink-0" />
                      <span>Tải lại PDF từ đám mây</span>
                    </button>
                  )}

                  {onReplacePdfFile ? (
                    <button
                      type="button"
                      onClick={handleTriggerReplaceFile}
                      className={`w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 min-h-[44px] rounded-xl text-sm font-semibold transition-colors ${
                        canRetryCloud
                          ? 'bg-white hover:bg-[#FAF8F5] border border-[#DDD7CC] text-slate-700 hover:text-pink-700'
                          : 'bg-pink-600 hover:bg-pink-700 text-white shadow-sm'
                      }`}
                    >
                      <Upload className="w-4 h-4 shrink-0" />
                      <span>Mở lại tệp PDF từ máy</span>
                    </button>
                  ) : (
                    !canRetryCloud && (
                      <button
                        type="button"
                        onClick={onBackHome}
                        className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 min-h-[44px] rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-sm font-semibold transition-colors"
                      >
                        <span>Quay lại Trang chủ để tải tệp</span>
                      </button>
                    )
                  )}
                </div>

                <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-center gap-2 text-[11.5px] text-slate-400">
                  {document.fileSize && <span>{document.fileSize}</span>}
                  {document.fileSize && <span>•</span>}
                  <span>{document.updatedAt}</span>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>
    );
  }

  // =========================================================================
  // NHÁNH 2: TRÌNH ĐỌC PDF.JS THẬT ĐA TRANG (CANVAS + TEXT LAYER + TÌM KIẾM)
  // =========================================================================
  if (isRealPdfMode && realPdfUrl) {
    const downloadName = getPdfDownloadFilename(document.title);
    const pageNumbers = Array.from({ length: pdfNumPages }, (_, idx) => idx + 1);
    const matchPageSet = new Set(searchMatchPages);

    const sidebarToggleLeftExtra =
      !useBrowserFallback && pdfLoadState === 'ready' && pdfNumPages > 0 ? (
        <>
          <button
            type="button"
            onClick={() => setIsMobileDrawerOpen(true)}
            aria-label="Mở danh sách trang PDF"
            title="Danh sách trang"
            className="md:hidden inline-flex items-center justify-center gap-1.5 min-h-[38px] px-2.5 rounded-xl bg-[#FAF8F5] hover:bg-[#F2EFEA] border border-[#E7E2DA] text-xs font-semibold text-slate-700 transition-colors shrink-0"
          >
            <PanelLeft className="w-4 h-4 text-pink-600 shrink-0" />
            <span className="hidden sm:inline">Trang</span>
          </button>
          <button
            type="button"
            onClick={() => setIsDesktopSidebarOpen((prev) => !prev)}
            aria-label={
              isDesktopSidebarOpen
                ? 'Ẩn cột thu nhỏ trang'
                : 'Hiện cột thu nhỏ trang'
            }
            aria-pressed={isDesktopSidebarOpen}
            title={
              isDesktopSidebarOpen
                ? 'Ẩn cột thu nhỏ trang'
                : 'Hiện cột thu nhỏ trang'
            }
            className={`hidden md:inline-flex items-center justify-center min-h-[36px] min-w-[36px] px-2 rounded-xl border text-xs font-medium transition-colors shrink-0 ${
              isDesktopSidebarOpen
                ? 'bg-pink-50 border-pink-200/80 text-pink-700 hover:bg-pink-100/70'
                : 'bg-[#FAF8F5] border-[#E7E2DA] text-slate-600 hover:bg-[#F2EFEA] hover:text-slate-900'
            }`}
          >
            <PanelLeft className="w-4 h-4 shrink-0" />
          </button>
        </>
      ) : undefined;

    const realPdfToolbar = (
      <div className="flex items-center gap-1.5 sm:gap-2">
        {!useBrowserFallback && pdfLoadState === 'ready' && pdfNumPages > 0 && (
          <>
            {/* Điều khiển số trang Trang X / N */}
            <div className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded-lg bg-[#F4F2ED] border border-[#E2DDD5] shrink-0">
              <button
                type="button"
                onClick={handlePrevPage}
                disabled={!canGoPrev}
                aria-label="Trang trước"
                className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-1.5 text-xs font-semibold text-slate-700 tabular-nums whitespace-nowrap">
                Trang {currentPageNumber} / {pdfNumPages}
              </span>
              <button
                type="button"
                onClick={handleNextPage}
                disabled={!canGoNext}
                aria-label="Trang sau"
                className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>

            {/* Điều khiển thu phóng Zoom */}
            <div className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded-lg bg-[#F4F2ED] border border-[#E2DDD5] shrink-0">
              <button
                type="button"
                onClick={handleZoomOut}
                disabled={!canZoomOut}
                aria-label="Thu nhỏ"
                title="Thu nhỏ"
                className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
              >
                <ZoomOut className="w-4 h-4" />
              </button>
              <button
                type="button"
                onClick={handleResetZoom}
                title="Bấm để đặt lại 100%"
                className={`px-1.5 py-0.5 rounded-md text-xs font-semibold tabular-nums whitespace-nowrap transition-colors ${
                  zoomPercent === 100
                    ? 'text-slate-700 hover:bg-white'
                    : 'text-pink-700 bg-white shadow-sm'
                }`}
              >
                {zoomPercent}%
              </button>
              <button
                type="button"
                onClick={handleZoomIn}
                disabled={!canZoomIn}
                aria-label="Phóng to"
                title="Phóng to"
                className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
              >
                <ZoomIn className="w-4 h-4" />
              </button>
            </div>

            {/* Nút bật/tắt Tìm trong PDF */}
            <button
              type="button"
              onClick={() => setIsSearchOpen((prev) => !prev)}
              aria-pressed={isSearchOpen}
              aria-label="Tìm trong PDF"
              title="Tìm từ khóa trong PDF"
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border text-xs font-semibold transition-colors whitespace-nowrap shrink-0 ${
                isSearchOpen
                  ? 'bg-pink-50 border-pink-300 text-pink-700'
                  : 'bg-[#FAF8F5] hover:bg-[#F2EFEA] border-[#E2DDD5] text-slate-700'
              }`}
            >
              <Search className="w-3.5 h-3.5 shrink-0" />
              <span className="hidden lg:inline">Tìm</span>
            </button>
          </>
        )}

        {useBrowserFallback && pdfLoadState !== 'error' && (
          <button
            type="button"
            onClick={() => setUseBrowserFallback(false)}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-[#FAF8F5] hover:bg-[#F2EFEA] border border-[#E2DDD5] text-xs font-semibold text-slate-700 transition-colors whitespace-nowrap shrink-0"
          >
            <Eye className="w-3.5 h-3.5 shrink-0" />
            <span className="hidden sm:inline">Trình đọc PDF.js</span>
          </button>
        )}

        {/* Tải xuống PDF */}
        <a
          href={realPdfUrl}
          download={downloadName}
          title="Tải xuống tệp PDF"
          className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg bg-white hover:bg-[#F8F6F2] border border-[#DDD7CC] text-xs font-semibold text-slate-700 hover:text-pink-700 transition-colors whitespace-nowrap shrink-0"
        >
          <Download className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden md:inline">Tải xuống PDF</span>
        </a>

        {/* Trình chiếu */}
        <button
          type="button"
          onClick={() => onOpenModal('present', document)}
          className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm transition-colors whitespace-nowrap shrink-0"
        >
          <Presentation className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden sm:inline">Trình chiếu</span>
        </button>
      </div>
    );

    const renderRealThumbnails = (isDrawer: boolean) => (
      <div className="p-3 space-y-3">
        {pageNumbers.map((pageNum) => {
          const isActive = pageNum === currentPageNumber;
          const thumbSrc = thumbnails[pageNum];
          const hasSearchHit = matchPageSet.has(pageNum);

          return (
            <button
              key={`real-pdf-thumb-${pageNum}`}
              ref={(node) => {
                registerThumbNode(pageNum, node);
                if (!isDrawer && isActive) {
                  activeThumbRef.current = node;
                }
              }}
              data-pdf-thumb-page={pageNum}
              type="button"
              onClick={() => goToPage(pageNum - 1)}
              className="w-full text-left group focus:outline-none"
            >
              <div
                className={`relative aspect-[3/4] w-full rounded-lg bg-white border transition-all overflow-hidden flex items-center justify-center ${
                  isActive
                    ? 'border-pink-600 ring-2 ring-pink-600/25 shadow-md'
                    : hasSearchHit
                      ? 'border-amber-400 ring-2 ring-amber-400/25'
                      : 'border-[#DDD7CC] group-hover:border-slate-400 group-hover:shadow-sm'
                }`}
              >
                {thumbSrc ? (
                  <img
                    src={thumbSrc}
                    alt={`Trang ${pageNum}`}
                    className="w-full h-full object-contain bg-white"
                    loading="lazy"
                  />
                ) : (
                  <div className="flex flex-col items-center justify-center gap-1.5 text-slate-300">
                    <FileText className="w-5 h-5" />
                    <span className="text-[10px] font-semibold text-slate-400 tabular-nums">
                      {pageNum}
                    </span>
                  </div>
                )}

                {hasSearchHit && (
                  <span className="absolute top-1.5 right-1.5 px-1.5 py-0.5 rounded bg-amber-400 text-amber-950 text-[9px] font-bold shadow-sm">
                    Khớp
                  </span>
                )}
              </div>

              <div className="mt-1.5 flex items-center justify-between px-0.5">
                <span
                  className={`text-[11.5px] tabular-nums ${
                    isActive
                      ? 'font-bold text-pink-700'
                      : 'font-medium text-slate-600 group-hover:text-slate-900'
                  }`}
                >
                  Trang {pageNum}
                </span>
              </div>
            </button>
          );
        })}
      </div>
    );

    return (
      <div className="h-[100dvh] w-full flex flex-col bg-[#E8E5DE] text-slate-900 overflow-hidden">
        <WorkspaceTopBar
          document={document}
          theme="light"
          onBackHome={onBackHome}
          onShare={() => onOpenModal('share', document)}
          leftExtra={sidebarToggleLeftExtra}
          toolbar={realPdfToolbar}
        />

        {/* Thanh tìm từ khóa trong PDF thật */}
        {isSearchOpen && !useBrowserFallback && pdfLoadState === 'ready' && (
          <div className="shrink-0 bg-white border-b border-[#E2DDD5] px-3 sm:px-4 py-2 flex flex-wrap items-center justify-between gap-2 z-10">
            <div className="flex items-center gap-2 flex-1 min-w-[200px] max-w-md relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                ref={searchInputRef}
                type="text"
                value={searchQuery}
                onChange={(event) => setSearchQuery(event.target.value)}
                onKeyDown={handleSearchInputKeyDown}
                placeholder="Tìm từ khóa trong các trang PDF..."
                aria-label="Tìm từ khóa trong PDF"
                className="w-full pl-8 pr-8 py-1.5 rounded-lg bg-[#FAF8F5] border border-[#DDD7CC] text-xs sm:text-[13px] text-slate-800 placeholder:text-slate-400 focus:outline-none focus:bg-white focus:border-pink-500 focus:ring-2 focus:ring-pink-500/15"
              />
              {searchQuery.length > 0 && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  aria-label="Xóa từ khóa tìm kiếm"
                  className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 rounded text-slate-400 hover:text-slate-700"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            <div className="flex items-center gap-1.5 text-xs">
              {isSearchingPdf ? (
                <span className="inline-flex items-center gap-1.5 text-slate-500 px-2">
                  <Loader2 className="w-3.5 h-3.5 animate-spin text-pink-600" />
                  Đang quét...
                </span>
              ) : searchQuery.trim().length > 0 ? (
                searchMatchPages.length > 0 ? (
                  <>
                    <span className="text-slate-600 font-medium tabular-nums px-1.5">
                      {activeMatchIdx + 1} / {searchMatchPages.length} trang khớp
                    </span>
                    <button
                      type="button"
                      onClick={() => handleJumpMatch('prev')}
                      aria-label="Trang khớp trước"
                      className="p-1 rounded-md border border-[#E2DDD5] bg-[#FAF8F5] hover:bg-white text-slate-700"
                    >
                      <ChevronUp className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => handleJumpMatch('next')}
                      aria-label="Trang khớp tiếp"
                      className="p-1 rounded-md border border-[#E2DDD5] bg-[#FAF8F5] hover:bg-white text-slate-700"
                    >
                      <ChevronDown className="w-3.5 h-3.5" />
                    </button>
                  </>
                ) : (
                  <span className="text-rose-600 font-medium px-1.5">
                    Không tìm thấy
                  </span>
                )
              ) : (
                <span className="hidden sm:inline text-slate-400 text-[11px]">
                  Nhấn Enter để nhảy tới trang khớp
                </span>
              )}

              <button
                type="button"
                onClick={() => {
                  setIsSearchOpen(false);
                  setSearchQuery('');
                }}
                aria-label="Đóng tìm kiếm"
                className="p-1 rounded-lg text-slate-400 hover:text-slate-800 hover:bg-slate-100 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}

        <div className="flex-1 min-h-0 flex relative overflow-hidden">
          {/* Cột thu nhỏ trang thật bên trái (Desktop md+) */}
          {!useBrowserFallback &&
            pdfLoadState === 'ready' &&
            pdfNumPages > 0 &&
            isDesktopSidebarOpen && (
              <aside
                aria-label="Danh sách trang PDF"
                className="hidden md:flex md:flex-col w-[200px] shrink-0 bg-[#F4F2ED] border-r border-[#E2DDD5] overflow-y-auto"
              >
                <div className="sticky top-0 z-10 px-3 py-2.5 bg-[#F4F2ED]/95 backdrop-blur-sm border-b border-[#E2DDD5] flex items-center justify-between">
                  <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
                    Thu nhỏ trang
                  </span>
                  <span className="text-[11px] font-semibold text-slate-600 bg-white px-2 py-0.5 rounded-full border border-[#E2DDD5] tabular-nums">
                    {pdfNumPages}
                  </span>
                </div>
                {renderRealThumbnails(false)}
              </aside>
            )}

          {/* Drawer thu nhỏ trang thật trên Mobile (< md) */}
          {!useBrowserFallback &&
            pdfLoadState === 'ready' &&
            pdfNumPages > 0 &&
            isMobileDrawerOpen && (
              <div className="fixed inset-0 z-40 flex md:hidden">
                <div
                  className="fixed inset-0 bg-slate-900/45 backdrop-blur-[1px]"
                  onClick={() => setIsMobileDrawerOpen(false)}
                  aria-hidden="true"
                />
                <aside
                  role="dialog"
                  aria-modal="true"
                  aria-label="Danh sách trang PDF"
                  className="relative z-10 w-[236px] max-w-[82vw] h-full bg-[#F4F2ED] border-r border-[#E2DDD5] shadow-2xl flex flex-col"
                >
                  <div className="px-3.5 py-3 border-b border-[#E2DDD5] flex items-center justify-between shrink-0">
                    <div className="flex items-center gap-2">
                      <FileText className="w-4 h-4 text-pink-600 shrink-0" />
                      <span className="text-xs font-bold text-slate-800">
                        Danh sách trang ({pdfNumPages})
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => setIsMobileDrawerOpen(false)}
                      aria-label="Đóng danh sách trang"
                      className="p-1 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-white transition-colors"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="flex-1 min-h-0 overflow-y-auto pb-safe">
                    {renderRealThumbnails(true)}
                  </div>
                </aside>
              </div>
            )}

          {/* Vùng hiển thị trung tâm */}
          <div
            ref={viewerOuterRef}
            className="flex-1 min-w-0 min-h-0 flex flex-col relative overflow-hidden"
          >
            {useBrowserFallback ? (
              <div className="flex-1 min-h-0 flex flex-col">
                <div className="shrink-0 px-3 sm:px-4 py-2 bg-[#F4F2ED] border-b border-[#DDD7CC] flex flex-wrap items-center justify-between gap-2 text-xs">
                  <span className="text-slate-600 font-medium">
                    Đang hiển thị bằng trình xem PDF mặc định của trình duyệt
                  </span>
                  <div className="flex items-center gap-2">
                    <a
                      href={realPdfUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-slate-50 border border-slate-200 text-slate-700 font-semibold transition-colors"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      Mở tab mới
                    </a>
                  </div>
                </div>
                <iframe
                  src={realPdfUrl}
                  title={document.title}
                  className="w-full flex-1 border-0 bg-white"
                />
              </div>
            ) : pdfLoadState === 'loading' || pdfLoadState === 'idle' ? (
              <div className="flex-1 min-h-0 flex items-center justify-center p-6">
                <div className="rounded-2xl bg-white/95 border border-[#E2DDD5] shadow-md px-7 py-6 flex flex-col items-center gap-3 text-center max-w-sm">
                  <div className="w-12 h-12 rounded-2xl bg-pink-50 border border-pink-200/70 text-pink-600 flex items-center justify-center">
                    <Loader2 className="w-6 h-6 animate-spin" />
                  </div>
                  <p className="text-sm font-bold text-slate-900">
                    Đang dựng trang PDF...
                  </p>
                  <p className="text-xs text-slate-500">
                    Đang nạp tài liệu &ldquo;{document.title}&rdquo; qua PDF.js
                    để hiển thị sắc nét và cho phép bôi đen sao chép chữ.
                  </p>
                </div>
              </div>
            ) : pdfLoadState === 'error' ? (
              <div className="flex-1 min-h-0 overflow-auto p-4 sm:p-8 flex items-center justify-center">
                <div className="w-full max-w-md rounded-2xl bg-white border border-[#E2DDD5] shadow-md p-6 sm:p-8 text-center">
                  <div className="w-14 h-14 mx-auto rounded-2xl bg-amber-50 border border-amber-200/80 text-amber-600 flex items-center justify-center mb-4">
                    <FileText className="w-7 h-7" />
                  </div>

                  <h2 className="text-base sm:text-lg font-bold text-slate-900 mb-2">
                    Không thể dựng trực tiếp tệp PDF này
                  </h2>

                  <p className="text-xs sm:text-sm text-slate-600 leading-relaxed mb-3">
                    Tệp PDF có thể đang được mã hóa mật khẩu hoặc dùng định dạng
                    nhúng đặc thù. Bạn vẫn có thể xem ngay bằng trình xem mặc
                    định của trình duyệt.
                  </p>

                  {pdfLoadError && (
                    <p className="mb-5 px-3 py-2 rounded-xl bg-slate-50 border border-slate-200 text-[11.5px] text-slate-500 break-words">
                      {pdfLoadError}
                    </p>
                  )}

                  <div className="space-y-2.5">
                    <button
                      type="button"
                      onClick={() => setUseBrowserFallback(true)}
                      className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 min-h-[44px] rounded-xl bg-pink-600 hover:bg-pink-700 text-white text-sm font-semibold shadow-sm transition-colors"
                    >
                      <Eye className="w-4 h-4 shrink-0" />
                      <span>Mở bằng trình xem trình duyệt</span>
                    </button>

                    <a
                      href={realPdfUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 min-h-[44px] rounded-xl bg-white hover:bg-[#FAF8F5] border border-[#DDD7CC] text-slate-700 text-sm font-semibold transition-colors"
                    >
                      <ExternalLink className="w-4 h-4 shrink-0" />
                      <span>Mở tệp PDF ở tab mới</span>
                    </a>
                  </div>
                </div>
              </div>
            ) : (
              <div className="relative flex-1 min-w-0 min-h-0 flex overflow-hidden">
                <main
                  ref={viewerScrollRef}
                  className="flex-1 min-w-0 min-h-0 overflow-auto scrollbar-none p-3 pr-5 sm:p-6 sm:pr-7 md:p-7 md:pr-8"
                >
                  <div className="w-fit min-w-full pb-8">
                    {/* Khung Canvas + Lớp văn bản chọn được 1:1 */}
                    <div
                      style={
                        renderedSize
                          ? {
                              width: `${renderedSize.width}px`,
                              height: `${renderedSize.height}px`,
                            }
                          : {
                              width: '100%',
                              maxWidth: `${BASE_PAGE_WIDTH_PX}px`,
                              minHeight: '68vh',
                            }
                      }
                      className="relative mx-auto bg-white shadow-md border border-slate-300/70 rounded-sm overflow-hidden"
                    >
                      <canvas
                        ref={canvasRef}
                        className="block"
                      />

                      <div
                        ref={textLayerRef}
                        className="textLayer"
                        aria-label={`Nội dung chữ trang ${currentPageNumber}`}
                      />

                      {isRenderingPage && (
                        <div className="absolute inset-0 z-10 bg-white/55 backdrop-blur-[1px] flex items-center justify-center pointer-events-none">
                          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-white shadow-sm border border-slate-200 text-xs font-semibold text-slate-600">
                            <Loader2 className="w-3.5 h-3.5 animate-spin text-pink-600" />
                            <span>Đang dựng trang {currentPageNumber}...</span>
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Thanh điều hướng trang nhanh dưới chân trang */}
                    <div className="mt-5 mx-auto w-full max-w-[760px] flex items-center justify-between gap-2 px-1 pb-safe">
                      <button
                        type="button"
                        onClick={handlePrevPage}
                        disabled={!canGoPrev}
                        className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white hover:bg-[#FAF9F6] border border-[#DCD6CB] text-xs font-semibold text-slate-700 shadow-sm disabled:opacity-40 disabled:hover:bg-white transition-colors"
                      >
                        <ChevronLeft className="w-4 h-4 shrink-0" />
                        <span>Trang trước</span>
                      </button>

                      <div className="flex flex-col items-center">
                        <span className="text-xs font-bold text-slate-700 tabular-nums">
                          Trang {currentPageNumber} / {pdfNumPages}
                        </span>
                        <span className="hidden sm:inline text-[11px] text-slate-500">
                          Bôi đen để sao chép chữ • Phím ← / → lật trang
                        </span>
                      </div>

                      <button
                        type="button"
                        onClick={handleNextPage}
                        disabled={!canGoNext}
                        className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white hover:bg-[#FAF9F6] border border-[#DCD6CB] text-xs font-semibold text-slate-700 shadow-sm disabled:opacity-40 disabled:hover:bg-white transition-colors"
                      >
                        <span>Trang sau</span>
                        <ChevronRight className="w-4 h-4 shrink-0" />
                      </button>
                    </div>
                  </div>
                </main>

                <CustomVerticalScrollbar
                  scrollContainerRef={viewerScrollRef}
                  theme="light"
                  getDragLabel={({ percent }) =>
                    `Trang ${currentPageNumber} / ${pdfNumPages} · ${Math.round(percent)}%`
                  }
                />
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  // =========================================================================
  // NHÁNH 3: CHẾ ĐỘ ĐỌC CẤU TRÚC ĐA TRANG CHO TÀI LIỆU PDF MẪU CÓ SẴN
  // =========================================================================
  const pageWidthPx = Math.round(BASE_PAGE_WIDTH_PX * zoomScale);
  const pageContainerStyle =
    zoomPercent <= 100
      ? {
          width: '100%',
          maxWidth: `${pageWidthPx}px`,
        }
      : {
          width: `${pageWidthPx}px`,
          minWidth: `${pageWidthPx}px`,
        };

  const bodyFontSizePx = Math.round(15 * zoomScale * 10) / 10;
  const headingFontSizePx = Math.round(25 * zoomScale * 10) / 10;

  const renderSampleThumbnailList = (isDrawer: boolean) => (
    <div className="p-3 space-y-3">
      {samplePages.map((page, idx) => {
        const isActive = idx === safePageIdx;
        return (
          <button
            key={page.id}
            ref={!isDrawer && isActive ? activeThumbRef : undefined}
            type="button"
            onClick={() => goToPage(idx)}
            className="w-full text-left group focus:outline-none"
          >
            <div
              className={`aspect-[3/4] w-full rounded-lg bg-white p-2.5 flex flex-col justify-between border transition-all overflow-hidden ${
                isActive
                  ? 'border-pink-600 ring-2 ring-pink-600/25 shadow-md'
                  : 'border-[#DDD7CC] group-hover:border-slate-400 group-hover:shadow-sm'
              }`}
            >
              <div>
                <div className="h-1 w-10 rounded-full bg-pink-600/80 mb-2" />
                <p className="text-[8.5px] font-bold uppercase tracking-wider text-pink-600/90 truncate mb-1">
                  {page.sectionTitle}
                </p>
                <p className="text-[10px] font-bold text-slate-800 leading-snug line-clamp-2 mb-2">
                  {page.heading}
                </p>
                <div className="space-y-1">
                  <div className="h-1 w-full rounded-full bg-slate-200" />
                  <div className="h-1 w-11/12 rounded-full bg-slate-200" />
                  <div className="h-1 w-4/5 rounded-full bg-slate-200/80" />
                </div>
                {page.callout && (
                  <div className="mt-2 h-3 w-full rounded bg-[#F7F3EC] border-l-2 border-l-pink-500" />
                )}
              </div>

              <div className="pt-1.5 border-t border-slate-100 flex items-center justify-between text-[8.5px] text-slate-400 tabular-nums">
                <span className="truncate max-w-[75%]">{document.title}</span>
                <span>{page.pageNumber}</span>
              </div>
            </div>

            <div className="mt-1.5 flex items-center justify-between px-0.5">
              <span
                className={`text-[11.5px] tabular-nums ${
                  isActive
                    ? 'font-bold text-pink-700'
                    : 'font-medium text-slate-600 group-hover:text-slate-900'
                }`}
              >
                Trang {page.pageNumber}
              </span>
              {page.highlights && page.highlights.length > 0 && (
                <span className="text-[10px] text-slate-400">
                  {page.highlights.length} ý chính
                </span>
              )}
            </div>
          </button>
        );
      })}
    </div>
  );

  const sampleLeftExtra = (
    <>
      <button
        type="button"
        onClick={() => setIsMobileDrawerOpen(true)}
        aria-label="Mở danh sách trang PDF"
        title="Danh sách trang"
        className="md:hidden inline-flex items-center justify-center gap-1.5 min-h-[38px] px-2.5 rounded-xl bg-[#FAF8F5] hover:bg-[#F2EFEA] border border-[#E7E2DA] text-xs font-semibold text-slate-700 transition-colors shrink-0"
      >
        <PanelLeft className="w-4 h-4 text-pink-600 shrink-0" />
        <span className="hidden sm:inline">Trang</span>
      </button>
      <button
        type="button"
        onClick={() => setIsDesktopSidebarOpen((prev) => !prev)}
        aria-label={
          isDesktopSidebarOpen
            ? 'Ẩn cột thu nhỏ trang'
            : 'Hiện cột thu nhỏ trang'
        }
        aria-pressed={isDesktopSidebarOpen}
        title={
          isDesktopSidebarOpen
            ? 'Ẩn cột thu nhỏ trang'
            : 'Hiện cột thu nhỏ trang'
        }
        className={`hidden md:inline-flex items-center justify-center min-h-[36px] min-w-[36px] px-2 rounded-xl border text-xs font-medium transition-colors shrink-0 ${
          isDesktopSidebarOpen
            ? 'bg-pink-50 border-pink-200/80 text-pink-700 hover:bg-pink-100/70'
            : 'bg-[#FAF8F5] border-[#E7E2DA] text-slate-600 hover:bg-[#F2EFEA] hover:text-slate-900'
        }`}
      >
        <PanelLeft className="w-4 h-4 shrink-0" />
      </button>
    </>
  );

  const structuredToolbar = (
    <div className="flex items-center gap-1.5 sm:gap-2">
      <div className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded-lg bg-[#F4F2ED] border border-[#E2DDD5] shrink-0">
        <button
          type="button"
          onClick={handlePrevPage}
          disabled={!canGoPrev}
          aria-label="Trang trước"
          className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
        >
          <ChevronLeft className="w-4 h-4" />
        </button>
        <span className="px-1.5 text-xs font-semibold text-slate-700 tabular-nums whitespace-nowrap">
          Trang {currentPageNumber} / {totalPages}
        </span>
        <button
          type="button"
          onClick={handleNextPage}
          disabled={!canGoNext}
          aria-label="Trang sau"
          className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
        >
          <ChevronRight className="w-4 h-4" />
        </button>
      </div>

      <div className="inline-flex items-center gap-0.5 px-1 py-0.5 rounded-lg bg-[#F4F2ED] border border-[#E2DDD5] shrink-0">
        <button
          type="button"
          onClick={handleZoomOut}
          disabled={!canZoomOut}
          aria-label="Thu nhỏ"
          title="Thu nhỏ"
          className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
        >
          <ZoomOut className="w-4 h-4" />
        </button>
        <button
          type="button"
          onClick={handleResetZoom}
          title="Bấm để đặt lại 100%"
          className={`px-1.5 py-0.5 rounded-md text-xs font-semibold tabular-nums whitespace-nowrap transition-colors ${
            zoomPercent === 100
              ? 'text-slate-700 hover:bg-white'
              : 'text-pink-700 bg-white shadow-sm'
          }`}
        >
          {zoomPercent}%
        </button>
        <button
          type="button"
          onClick={handleZoomIn}
          disabled={!canZoomIn}
          aria-label="Phóng to"
          title="Phóng to"
          className="p-1 rounded-md text-slate-700 hover:bg-white disabled:opacity-35 disabled:hover:bg-transparent transition-colors"
        >
          <ZoomIn className="w-4 h-4" />
        </button>
      </div>

      <button
        type="button"
        onClick={() => onOpenModal('present', document)}
        className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm transition-colors whitespace-nowrap shrink-0"
      >
        <Presentation className="w-3.5 h-3.5 shrink-0" />
        <span className="hidden sm:inline">Trình chiếu</span>
      </button>
    </div>
  );

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#E8E5DE] text-slate-900 overflow-hidden">
      <WorkspaceTopBar
        document={document}
        theme="light"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', document)}
        leftExtra={sampleLeftExtra}
        toolbar={structuredToolbar}
      />

      <div className="flex-1 min-h-0 flex relative overflow-hidden">
        {isDesktopSidebarOpen && (
          <aside
            aria-label="Danh sách trang PDF"
            className="hidden md:flex md:flex-col w-[200px] shrink-0 bg-[#F4F2ED] border-r border-[#E2DDD5] overflow-y-auto"
          >
            <div className="sticky top-0 z-10 px-3 py-2.5 bg-[#F4F2ED]/95 backdrop-blur-sm border-b border-[#E2DDD5] flex items-center justify-between">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
                Thu nhỏ trang
              </span>
              <span className="text-[11px] font-semibold text-slate-600 bg-white px-2 py-0.5 rounded-full border border-[#E2DDD5] tabular-nums">
                {totalPages}
              </span>
            </div>
            {renderSampleThumbnailList(false)}
          </aside>
        )}

        {isMobileDrawerOpen && (
          <div className="fixed inset-0 z-40 flex md:hidden">
            <div
              className="fixed inset-0 bg-slate-900/45 backdrop-blur-[1px]"
              onClick={() => setIsMobileDrawerOpen(false)}
              aria-hidden="true"
            />
            <aside
              role="dialog"
              aria-modal="true"
              aria-label="Danh sách trang PDF"
              className="relative z-10 w-[236px] max-w-[82vw] h-full bg-[#F4F2ED] border-r border-[#E2DDD5] shadow-2xl flex flex-col"
            >
              <div className="px-3.5 py-3 border-b border-[#E2DDD5] flex items-center justify-between shrink-0">
                <div className="flex items-center gap-2">
                  <FileText className="w-4 h-4 text-pink-600 shrink-0" />
                  <span className="text-xs font-bold text-slate-800">
                    Danh sách trang ({totalPages})
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => setIsMobileDrawerOpen(false)}
                  aria-label="Đóng danh sách trang"
                  className="p-1 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-white transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              <div className="flex-1 min-h-0 overflow-y-auto pb-safe">
                {renderSampleThumbnailList(true)}
              </div>
            </aside>
          </div>
        )}

        <div className="relative flex-1 min-w-0 min-h-0 flex overflow-hidden">
          <main
            ref={viewerScrollRef}
            className="flex-1 min-w-0 min-h-0 overflow-auto scrollbar-none p-3 pr-5 sm:p-6 sm:pr-7 md:p-8 md:pr-9"
          >
            {currentSamplePage && (
            <div className="w-fit min-w-full pb-8">
              <article
                style={pageContainerStyle}
                className="mx-auto bg-white shadow-md border border-slate-200/80 rounded-md overflow-hidden transition-[width,max-width] duration-150"
              >
                <div className="h-1.5 w-full bg-gradient-to-r from-pink-600 via-rose-500 to-pink-400" />

                <div className="px-5 py-6 sm:px-10 sm:py-9 md:px-12 md:py-11 flex flex-col min-h-[68vh] justify-between">
                  <div>
                    <header className="pb-3.5 mb-6 border-b border-slate-200/80 flex items-center justify-between gap-3 text-[11px] sm:text-xs text-slate-400">
                      <span className="font-medium text-slate-500 truncate">
                        {document.title}
                      </span>
                      <span className="shrink-0 font-semibold text-pink-700/80 tabular-nums">
                        Trang {currentSamplePage.pageNumber} / {totalPages}
                      </span>
                    </header>

                    <div className="mb-6">
                      <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-pink-50 border border-pink-200/70 text-pink-700 text-[11px] sm:text-xs font-bold tracking-wide mb-3">
                        <FileText className="w-3.5 h-3.5 shrink-0" />
                        {currentSamplePage.sectionTitle}
                      </span>
                      <h1
                        style={{ fontSize: `${headingFontSizePx}px` }}
                        className="font-bold text-slate-900 leading-snug tracking-tight"
                      >
                        {currentSamplePage.heading}
                      </h1>
                    </div>

                    <div
                      style={{ fontSize: `${bodyFontSizePx}px` }}
                      className="space-y-4 text-slate-700 leading-[1.78]"
                    >
                      {currentSamplePage.paragraphs.map((paragraph, pIdx) => (
                        <p key={`${currentSamplePage.id}-p-${pIdx}`}>
                          {paragraph}
                        </p>
                      ))}
                    </div>

                    {currentSamplePage.callout && (
                      <aside className="mt-7 rounded-r-xl rounded-l-sm bg-[#FBF8F3] border border-[#EAE2D6] border-l-4 border-l-pink-600 p-4 sm:p-5">
                        <div className="flex items-start gap-3">
                          <div className="w-7 h-7 rounded-lg bg-pink-100/80 text-pink-700 flex items-center justify-center shrink-0 mt-0.5">
                            <Lightbulb className="w-4 h-4" />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[11px] font-bold uppercase tracking-wider text-pink-700 mb-1">
                              Ghi chú trọng tâm
                            </p>
                            <p
                              style={{
                                fontSize: `${Math.max(13, bodyFontSizePx - 1)}px`,
                              }}
                              className="text-slate-700 leading-relaxed"
                            >
                              {currentSamplePage.callout}
                            </p>
                          </div>
                        </div>
                      </aside>
                    )}

                    {currentSamplePage.highlights &&
                      currentSamplePage.highlights.length > 0 && (
                        <section className="mt-7 rounded-xl bg-slate-50/90 border border-slate-200/80 p-4 sm:p-5">
                          <div className="flex items-center gap-2 mb-3">
                            <Sparkles className="w-4 h-4 text-pink-600 shrink-0" />
                            <h2 className="text-xs font-bold uppercase tracking-wider text-slate-600">
                              Điểm nhấn chính
                            </h2>
                          </div>
                          <ul
                            style={{
                              fontSize: `${Math.max(13, bodyFontSizePx - 0.5)}px`,
                            }}
                            className="space-y-2.5 text-slate-700"
                          >
                            {currentSamplePage.highlights.map((item, hIdx) => (
                              <li
                                key={`${currentSamplePage.id}-hl-${hIdx}`}
                                className="flex items-start gap-2.5 leading-relaxed"
                              >
                                <span
                                  className="mt-2 w-1.5 h-1.5 rounded-full bg-pink-600 shrink-0"
                                  aria-hidden="true"
                                />
                                <span>{item}</span>
                              </li>
                            ))}
                          </ul>
                        </section>
                      )}
                  </div>

                  <footer className="mt-10 pt-4 border-t border-slate-200/80 flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-400">
                    <div className="flex items-center gap-2 truncate">
                      <span>{document.folder || 'Thư viện LITE'}</span>
                      <span>•</span>
                      <span>{document.updatedAt}</span>
                      {document.fileSize && (
                        <>
                          <span>•</span>
                          <span>{document.fileSize}</span>
                        </>
                      )}
                    </div>
                    <span className="font-semibold text-slate-500 tabular-nums">
                      Trang {currentSamplePage.pageNumber} / {totalPages}
                    </span>
                  </footer>
                </div>
              </article>

              <div className="mt-5 mx-auto w-full max-w-[760px] flex items-center justify-between gap-2 px-1 pb-safe">
                <button
                  type="button"
                  onClick={handlePrevPage}
                  disabled={!canGoPrev}
                  className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white hover:bg-[#FAF9F6] border border-[#DCD6CB] text-xs font-semibold text-slate-700 shadow-sm disabled:opacity-40 disabled:hover:bg-white transition-colors"
                >
                  <ChevronLeft className="w-4 h-4 shrink-0" />
                  <span>Trang trước</span>
                </button>

                <div className="flex flex-col items-center">
                  <span className="text-xs font-bold text-slate-700 tabular-nums">
                    Trang {currentPageNumber} / {totalPages}
                  </span>
                  <span className="hidden sm:inline text-[11px] text-slate-500">
                    Dùng phím ← / → để lật trang
                  </span>
                </div>

                <button
                  type="button"
                  onClick={handleNextPage}
                  disabled={!canGoNext}
                  className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white hover:bg-[#FAF9F6] border border-[#DCD6CB] text-xs font-semibold text-slate-700 shadow-sm disabled:opacity-40 disabled:hover:bg-white transition-colors"
                >
                  <span>Trang sau</span>
                  <ChevronRight className="w-4 h-4 shrink-0" />
                </button>
              </div>
            </div>
          )}
        </main>

          <CustomVerticalScrollbar
            scrollContainerRef={viewerScrollRef}
            theme="light"
            getDragLabel={({ percent }) =>
              `Trang ${currentPageNumber} / ${totalPages} · ${Math.round(percent)}%`
            }
          />
        </div>
      </div>
    </div>
  );
}

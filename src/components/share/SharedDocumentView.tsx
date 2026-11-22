import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, FormEvent } from 'react';
import {
  AlertCircle,
  ArrowLeft,
  BookOpen,
  ChevronLeft,
  ChevronRight,
  Download,
  Eye,
  EyeOff,
  FileSpreadsheet,
  FileText,
  Home,
  KeyRound,
  Layers,
  Loader2,
  Lock,
  Presentation,
  RefreshCw,
  ShieldAlert,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {
  CellStyle,
  DocumentItem,
  DocumentType,
  SharedDocumentResolution,
  SlidePreviewItem,
  SlideTextBlock,
  SpreadsheetSheet,
} from '../../types/document';
import {
  resolvePublicShareLink,
  unlockPasswordProtectedShare,
} from '../../services/shareService';
import {
  columnIndexToLetter,
  evaluateSheetMatrix,
} from '../../utils/formulaEvaluator';
import {
  loadPdfDocumentFromSource,
  renderPdfPageToCanvas,
  type PdfDocumentProxy,
} from '../../utils/pdfEngine';
import {
  exportExcelToXlsx,
  exportSlidesToPptx,
  exportWordToDocx,
} from '../../utils/fileExporter';
import { PresenterStageView } from '../modals/PresenterStageView';

export interface SharedDocumentViewProps {
  shareToken: string;
  onNavigateHome: () => void;
}

const TYPE_BADGE_STYLES: Record<
  DocumentType,
  { label: string; className: string }
> = {
  pptx: {
    label: 'PPTX',
    className: 'bg-orange-50 text-orange-700 border-orange-200/80',
  },
  docx: {
    label: 'DOCX',
    className: 'bg-blue-50 text-blue-700 border-blue-200/80',
  },
  xlsx: {
    label: 'XLSX',
    className: 'bg-green-100 text-green-800 border-green-300/80',
  },
  pdf: {
    label: 'PDF',
    className: 'bg-pink-50 text-pink-700 border-pink-200/80',
  },
};

const FORBIDDEN_HTML_TAGS = new Set([
  'SCRIPT',
  'STYLE',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'SVG',
  'MATH',
  'LINK',
  'META',
  'BASE',
  'FORM',
  'INPUT',
  'TEXTAREA',
  'BUTTON',
]);

const PDF_ZOOM_LEVELS = [75, 90, 100, 115, 130, 150] as const;

const FORMULA_ERROR_CODES = new Set(['#CYCLE!', '#DIV/0!', '#ERROR!']);

const NUMBER_FORMATTER = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 4,
});

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isSafeHtmlUrl(rawUrl: string, allowDataImage: boolean): boolean {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return true;
  }
  const normalized = trimmed.replace(/[\u0000-\u0020]+/g, '').toLowerCase();
  if (
    normalized.startsWith('javascript:') ||
    normalized.startsWith('vbscript:') ||
    normalized.startsWith('data:text/html')
  ) {
    return false;
  }
  if (normalized.startsWith('data:')) {
    return allowDataImage && normalized.startsWith('data:image/');
  }
  return true;
}

interface WordOutlineHeading {
  id: string;
  label: string;
  level: 1 | 2 | 3;
}

function sanitizeAndIndexWordHtml(rawHtml: string): {
  html: string;
  headings: WordOutlineHeading[];
} {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(rawHtml || '', 'text/html');

  const removeNodes: Element[] = [];
  parsed.body.querySelectorAll('*').forEach((el) => {
    if (FORBIDDEN_HTML_TAGS.has(el.tagName.toUpperCase())) {
      removeNodes.push(el);
      return;
    }

    for (const attr of Array.from(el.attributes)) {
      const nameLower = attr.name.toLowerCase();
      if (nameLower.startsWith('on') || nameLower === 'contenteditable') {
        el.removeAttribute(attr.name);
        continue;
      }
      if (nameLower === 'src') {
        if (!isSafeHtmlUrl(attr.value, true)) {
          el.removeAttribute(attr.name);
        }
        continue;
      }
      if (nameLower === 'href' || nameLower === 'xlink:href') {
        if (!isSafeHtmlUrl(attr.value, false)) {
          el.removeAttribute(attr.name);
        }
      }
    }
  });

  removeNodes.forEach((node) => node.remove());

  const headings: WordOutlineHeading[] = [];
  let seq = 0;
  parsed.body.querySelectorAll('h1, h2, h3').forEach((el) => {
    const label = (el.textContent ?? '').replace(/\s+/g, ' ').trim();
    if (!label) {
      return;
    }
    seq += 1;
    const existingId = el.getAttribute('id')?.trim();
    const headingId = existingId || `shared-heading-${seq}`;
    el.setAttribute('id', headingId);

    const parsedLevel = Number.parseInt(el.tagName.slice(1), 10);
    const level: 1 | 2 | 3 =
      parsedLevel === 1 ? 1 : parsedLevel === 2 ? 2 : 3;
    headings.push({ id: headingId, label, level });
  });

  return {
    html: parsed.body.innerHTML,
    headings,
  };
}

function buildFallbackWordHtml(doc: DocumentItem): string {
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
    parts.push('<p>Tài liệu hiện chưa có nội dung văn bản.</p>');
  }

  return parts.join('');
}

function isLetterColumnHeaders(columns: string[]): boolean {
  if (columns.length === 0) {
    return false;
  }
  return columns.every(
    (col, idx) => col.trim().toUpperCase() === columnIndexToLetter(idx)
  );
}

interface NormalizedReadSheet {
  id: string;
  name: string;
  columnCount: number;
  evaluatedRows: (string | number)[][];
  cellStyles: Record<string, CellStyle>;
}

function normalizeSheetForReadView(
  sheet: SpreadsheetSheet,
  fallbackIndex: number
): NormalizedReadSheet {
  const rawColumns = Array.isArray(sheet.columns) ? sheet.columns : [];
  const rawRows = Array.isArray(sheet.rows) ? sheet.rows : [];
  const maxRowWidth = rawRows.reduce(
    (max, row) => Math.max(max, Array.isArray(row) ? row.length : 0),
    0
  );
  const colCount = Math.max(rawColumns.length, maxRowWidth, 1);

  const shouldPrependHeader =
    rawColumns.length > 0 &&
    !isLetterColumnHeaders(rawColumns) &&
    !(
      rawRows.length > 0 &&
      rawColumns.every(
        (col, idx) => String(rawRows[0]?.[idx] ?? '') === String(col)
      )
    );

  const nextStyles: Record<string, CellStyle> = {};
  if (sheet.cellStyles) {
    for (const [key, style] of Object.entries(sheet.cellStyles)) {
      const parts = key.split(':');
      if (parts.length === 2) {
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) {
          const targetRow = shouldPrependHeader ? r + 1 : r;
          nextStyles[`${targetRow}:${c}`] = { ...style };
        }
      }
    }
  }

  const normalizedRows: (string | number)[][] = [];
  if (shouldPrependHeader) {
    normalizedRows.push(
      Array.from({ length: colCount }, (_, idx) => rawColumns[idx] ?? '')
    );
    for (let c = 0; c < colCount; c += 1) {
      nextStyles[`0:${c}`] = { ...nextStyles[`0:${c}`], bold: true };
    }
  }

  for (const row of rawRows) {
    normalizedRows.push(
      Array.from({ length: colCount }, (_, colIdx) => {
        const val = Array.isArray(row) ? row[colIdx] : '';
        return val === undefined || val === null ? '' : val;
      })
    );
  }

  if (normalizedRows.length === 0) {
    normalizedRows.push(Array.from({ length: colCount }, () => ''));
  }

  return {
    id: sheet.id || `sheet-${fallbackIndex + 1}`,
    name: sheet.name || `Sheet ${fallbackIndex + 1}`,
    columnCount: colCount,
    evaluatedRows: evaluateSheetMatrix(normalizedRows),
    cellStyles: nextStyles,
  };
}

function formatCellValue(value: string | number): string {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? NUMBER_FORMATTER.format(value) : '';
  }
  return value;
}

function isLightHexColor(color?: string): boolean {
  if (!color) {
    return false;
  }
  const normalized = color.trim().replace(/^#/, '');
  let r = 0;
  let g = 0;
  let b = 0;
  if (normalized.length === 3) {
    r = parseInt(normalized[0] + normalized[0], 16);
    g = parseInt(normalized[1] + normalized[1], 16);
    b = parseInt(normalized[2] + normalized[2], 16);
  } else if (normalized.length === 6) {
    r = parseInt(normalized.slice(0, 2), 16);
    g = parseInt(normalized.slice(2, 4), 16);
    b = parseInt(normalized.slice(4, 6), 16);
  } else {
    return false;
  }
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return false;
  }
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.62;
}

function getResponsiveBlockFontSize(block: SlideTextBlock): string {
  const base = Math.max(13, Math.min(54, block.fontSize || 18));
  const minPx = Math.max(13, Math.round(base * 0.7));
  const vw = (base * 0.11).toFixed(2);
  return `clamp(${minPx}px, calc(10px + ${vw}vw), ${base}px)`;
}

function buildPairingCode(docId: string): string {
  let hash = 1357;
  for (let i = 0; i < docId.length; i += 1) {
    hash = (hash * 31 + docId.charCodeAt(i)) % 9000;
  }
  return `LITE-${1000 + hash}`;
}

function buildResolvedDocument(
  resolution: SharedDocumentResolution
): DocumentItem {
  const baseDoc = resolution.document;
  const fallbackType: DocumentType = resolution.fileType ?? 'docx';
  const fallbackTitle = resolution.title ?? 'Tài liệu được chia sẻ';

  if (baseDoc) {
    const resolvedUrl =
      resolution.downloadUrl ?? baseDoc.fileUrl ?? baseDoc.objectUrl;
    return {
      ...baseDoc,
      title: baseDoc.title || fallbackTitle,
      type: baseDoc.type || fallbackType,
      fileSize: baseDoc.fileSize ?? resolution.fileSize,
      originalFileName: baseDoc.originalFileName ?? resolution.fileName,
      fileUrl: resolvedUrl,
    };
  }

  return {
    id: `shared-${resolution.shareToken}`,
    title: fallbackTitle,
    type: fallbackType,
    updatedAt: 'Liên kết chia sẻ',
    isFavorite: false,
    thumbnail: '',
    slidesPreview: [],
    outline: [],
    contentPreview: '',
    fileSize: resolution.fileSize,
    originalFileName: resolution.fileName,
    fileUrl: resolution.downloadUrl,
    r2ObjectKey: resolution.r2ObjectKey,
    storagePath: resolution.storagePath,
    shareToken: resolution.shareToken,
    shareAllowDownload: resolution.allowDownload,
  };
}

function LiteLogoMark({ onClick }: { onClick: () => void }) {
  return (
    <button
      type="button"
      onClick={onClick}
      title="Mở LITE Workspace"
      aria-label="Mở LITE Workspace"
      className="group flex items-center gap-2 shrink-0 rounded-xl focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-indigo-500/40"
    >
      <span className="flex h-9 w-9 items-center justify-center rounded-xl bg-[#4F46E5] text-white shadow-sm shadow-indigo-600/25 transition-transform group-hover:scale-[1.03]">
        <svg
          width="18"
          height="18"
          viewBox="0 0 24 24"
          fill="none"
          aria-hidden="true"
        >
          <path
            d="M6.5 4.25H13.75L17.75 8.25V19.25C17.75 19.8023 17.3023 20.25 16.75 20.25H6.5C5.94772 20.25 5.5 19.8023 5.5 19.25V5.25C5.5 4.69772 5.94772 4.25 6.5 4.25Z"
            stroke="currentColor"
            strokeWidth="1.7"
            strokeLinejoin="round"
            className="opacity-45"
          />
          <path
            d="M13.5 4.5V8.5H17.5"
            stroke="currentColor"
            strokeWidth="1.6"
            strokeLinecap="round"
            strokeLinejoin="round"
            className="opacity-55"
          />
          <path
            d="M9.25 9.25V16.25H14.5"
            stroke="currentColor"
            strokeWidth="2.35"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </span>
      <span className="hidden sm:inline text-sm font-bold tracking-[0.08em] text-slate-900">
        LITE
      </span>
    </button>
  );
}

function getSafeSlideAspectRatio(
  slide: SlidePreviewItem | null | undefined,
  doc: DocumentItem
): number {
  const candidate = slide?.slideWidthRatio ?? doc.slideAspectRatio ?? 16 / 9;
  return Number.isFinite(candidate) && candidate > 0.3 && candidate < 4
    ? candidate
    : 16 / 9;
}

function isPositionedSlideItem(slide: SlidePreviewItem): boolean {
  if (slide.isMasterPositioned) {
    return true;
  }
  const hasPositionedBlock = (slide.blocks ?? []).some(
    (b) => typeof b.xPct === 'number' && typeof b.yPct === 'number'
  );
  const hasPositionedImage = (slide.images ?? []).some(
    (img) => typeof img.xPct === 'number' && typeof img.yPct === 'number'
  );
  return hasPositionedBlock || hasPositionedImage;
}

function inferPositionedLightBackground(slide: SlidePreviewItem): boolean {
  if (slide.backgroundColor) {
    return isLightHexColor(slide.backgroundColor);
  }
  if (slide.backgroundGradient || slide.backgroundImageSrc) {
    return false;
  }
  const coloredBlocks = (slide.blocks ?? []).filter((b) => Boolean(b.color));
  if (coloredBlocks.length === 0) {
    return true;
  }
  const lightTextCount = coloredBlocks.filter((b) =>
    isLightHexColor(b.color)
  ).length;
  return lightTextCount <= coloredBlocks.length / 2;
}

function buildSharedSlideBackground(
  slide: SlidePreviewItem,
  isPositioned: boolean,
  hasLightBg: boolean
): CSSProperties {
  const style: CSSProperties = {};
  if (slide.backgroundImageSrc) {
    style.backgroundImage = slide.backgroundGradient
      ? `url("${slide.backgroundImageSrc}"), ${slide.backgroundGradient}`
      : `url("${slide.backgroundImageSrc}")`;
    style.backgroundSize = 'cover';
    style.backgroundPosition = 'center';
    style.backgroundRepeat = 'no-repeat';
    style.backgroundColor =
      slide.backgroundColor ?? (isPositioned && hasLightBg ? '#FFFFFF' : '#151B28');
    return style;
  }

  if (slide.backgroundGradient) {
    style.backgroundImage = slide.backgroundGradient;
    style.backgroundColor =
      slide.backgroundColor ?? (isPositioned && hasLightBg ? '#FFFFFF' : '#151B28');
    return style;
  }

  if (slide.backgroundColor) {
    style.backgroundColor = slide.backgroundColor;
    return style;
  }

  if (isPositioned) {
    style.backgroundColor = hasLightBg ? '#FFFFFF' : '#151B28';
    return style;
  }

  style.backgroundColor = '#151B28';
  return style;
}

interface SharedSlideCanvasProps {
  doc: DocumentItem;
  slide: SlidePreviewItem;
  slideIndex: number;
  totalSlides: number;
  onPrevSlide: () => void;
  onNextSlide: () => void;
  onOpenPresenter: () => void;
}

function SharedSlideCanvas({
  doc,
  slide,
  slideIndex,
  totalSlides,
  onPrevSlide,
  onNextSlide,
  onOpenPresenter,
}: SharedSlideCanvasProps) {
  const frameRef = useRef<HTMLDivElement | null>(null);
  const [frameWidthPx, setFrameWidthPx] = useState<number>(860);

  useEffect(() => {
    const el = frameRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateWidth = () => {
      const width = el.getBoundingClientRect().width;
      if (width > 0) {
        setFrameWidthPx(width);
      }
    };
    updateWidth();
    const observer = new ResizeObserver(updateWidth);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, []);

  const aspectRatio = getSafeSlideAspectRatio(slide, doc);
  const isPositioned = isPositionedSlideItem(slide);
  const hasLightBg = isPositioned
    ? inferPositionedLightBackground(slide)
    : isLightHexColor(slide.backgroundColor);
  const backgroundStyle = buildSharedSlideBackground(
    slide,
    isPositioned,
    hasLightBg
  );
  const fontScale = frameWidthPx > 0 ? frameWidthPx / 960 : 0.9;
  const defaultTextColor = hasLightBg ? '#0F172A' : '#FFFFFF';

  const activeBlocks = (slide.blocks ?? []).filter(
    (b) =>
      b.kind === 'table' ||
      b.text.trim().length > 0 ||
      Boolean(b.fillColor) ||
      Boolean(b.borderColor)
  );
  const activeImages = slide.images ?? [];
  const hasRealContent =
    activeBlocks.length > 0 || activeImages.length > 0;

  return (
    <div className="rounded-2xl sm:rounded-3xl border border-[#DDD6CA] bg-white shadow-xl overflow-hidden flex flex-col">
      {/* Slide Surface */}
      <div
        ref={frameRef}
        className="w-full relative overflow-hidden"
        style={{
          aspectRatio: String(aspectRatio),
          ...backgroundStyle,
        }}
      >
        {isPositioned ? (
          <>
            {activeImages.map((img) => {
              const hasCoords =
                typeof img.xPct === 'number' && typeof img.yPct === 'number';
              return (
                <div
                  key={img.id}
                  className="overflow-hidden flex items-center justify-center"
                  style={
                    hasCoords
                      ? {
                          position: 'absolute',
                          left: `${img.xPct ?? 0}%`,
                          top: `${img.yPct ?? 0}%`,
                          width: `${img.wPct ?? 35}%`,
                          height: `${img.hPct ?? 35}%`,
                        }
                      : {
                          position: 'absolute',
                          right: '5%',
                          bottom: '8%',
                          width: '36%',
                          height: '48%',
                        }
                  }
                >
                  <img
                    src={img.src}
                    alt={img.alt ?? slide.title}
                    draggable={false}
                    className="w-full h-full object-contain select-none"
                  />
                </div>
              );
            })}

            {activeBlocks.map((block) => {
              const isTable =
                block.kind === 'table' &&
                Array.isArray(block.tableRows) &&
                block.tableRows.length > 0;
              const baseFontPt = block.fontSize || (isTable ? 12 : 18);
              const scaledFontPx = Math.max(8, baseFontPt * fontScale);
              const valignJustify =
                block.valign === 'middle'
                  ? 'center'
                  : block.valign === 'bottom'
                    ? 'flex-end'
                    : 'flex-start';
              const textColor = block.color ?? defaultTextColor;

              return (
                <div
                  key={block.id}
                  className="overflow-hidden flex flex-col"
                  style={{
                    position: 'absolute',
                    left: `${block.xPct ?? 5}%`,
                    top: `${block.yPct ?? 5}%`,
                    width: `${block.wPct ?? 90}%`,
                    height: `${block.hPct ?? 16}%`,
                    justifyContent: isTable ? 'flex-start' : valignJustify,
                    backgroundColor: block.fillColor ?? undefined,
                    border: block.borderColor
                      ? `1px solid ${block.borderColor}`
                      : undefined,
                    borderRadius:
                      block.fillColor || block.borderColor
                        ? '0.25rem'
                        : undefined,
                    padding: isTable ? 0 : '2px 6px',
                  }}
                >
                  {isTable && block.tableRows ? (
                    <table
                      className="w-full h-full border-collapse table-fixed"
                      style={{
                        fontSize: `${scaledFontPx.toFixed(2)}px`,
                        lineHeight: 1.22,
                        color: textColor,
                        fontFamily: block.fontFamily
                          ? `${block.fontFamily}, sans-serif`
                          : undefined,
                      }}
                    >
                      <tbody>
                        {block.tableRows.map((row, rIdx) => (
                          <tr
                            key={`${block.id}-tr-${rIdx}`}
                            style={{
                              backgroundColor:
                                rIdx === 0
                                  ? hasLightBg
                                    ? 'rgba(15, 23, 42, 0.07)'
                                    : 'rgba(255, 255, 255, 0.1)'
                                  : undefined,
                            }}
                          >
                            {row.map((cell, cIdx) => (
                              <td
                                key={`${block.id}-td-${rIdx}-${cIdx}`}
                                className="break-words align-middle overflow-hidden"
                                style={{
                                  border: `1px solid ${
                                    block.borderColor ??
                                    (hasLightBg
                                      ? 'rgba(15, 23, 42, 0.18)'
                                      : 'rgba(255, 255, 255, 0.22)')
                                  }`,
                                  padding: '3px 6px',
                                  fontWeight:
                                    rIdx === 0 || block.bold ? 600 : 400,
                                  textAlign: block.align ?? 'left',
                                }}
                              >
                                {cell}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : (
                    <p
                      className="w-full whitespace-pre-wrap break-words"
                      style={{
                        fontSize: `${scaledFontPx.toFixed(2)}px`,
                        lineHeight: 1.25,
                        fontWeight: block.bold ? 700 : 400,
                        fontStyle: block.italic ? 'italic' : 'normal',
                        fontFamily: block.fontFamily
                          ? `${block.fontFamily}, sans-serif`
                          : undefined,
                        color: textColor,
                        textAlign: block.align ?? 'left',
                      }}
                    >
                      {block.text}
                    </p>
                  )}
                </div>
              );
            })}
          </>
        ) : (
          <div className="w-full h-full p-5 sm:p-10 flex flex-col justify-center">
            {hasRealContent ? (
              <div
                className={`grid gap-6 items-center ${
                  activeImages.length > 0 && activeBlocks.length > 0
                    ? 'grid-cols-1 md:grid-cols-12'
                    : slide.metricHighlight && activeImages.length === 0
                      ? 'grid-cols-1 sm:grid-cols-12'
                      : 'grid-cols-1'
                }`}
              >
                {activeBlocks.length > 0 && (
                  <div
                    className={`space-y-3 ${
                      activeImages.length > 0
                        ? 'md:col-span-7'
                        : slide.metricHighlight
                          ? 'sm:col-span-8'
                          : 'col-span-1'
                    }`}
                  >
                    {activeBlocks.map((block) => (
                      <p
                        key={block.id}
                        className="whitespace-pre-wrap leading-snug break-words"
                        style={{
                          fontSize: getResponsiveBlockFontSize(block),
                          fontWeight: block.bold ? 700 : 400,
                          fontStyle: block.italic ? 'italic' : 'normal',
                          color: block.color ?? defaultTextColor,
                          textAlign: block.align ?? 'left',
                        }}
                      >
                        {block.text}
                      </p>
                    ))}
                  </div>
                )}

                {activeImages.length > 0 ? (
                  <div
                    className={`${
                      activeBlocks.length > 0 ? 'md:col-span-5' : 'col-span-1'
                    } grid gap-3 ${
                      activeImages.length > 1 ? 'grid-cols-2' : 'grid-cols-1'
                    }`}
                  >
                    {activeImages.map((img) => (
                      <div
                        key={img.id}
                        className="rounded-xl overflow-hidden bg-black/20 border border-white/10 p-1.5 flex items-center justify-center"
                      >
                        <img
                          src={img.src}
                          alt={img.alt ?? slide.title}
                          className="max-h-[340px] w-full object-contain rounded-lg"
                        />
                      </div>
                    ))}
                  </div>
                ) : (
                  slide.metricHighlight && (
                    <div
                      className="sm:col-span-4 rounded-2xl p-5 border flex flex-col justify-center sm:items-end sm:text-right"
                      style={{
                        backgroundColor: `${slide.accentColor}22`,
                        borderColor: `${slide.accentColor}66`,
                      }}
                    >
                      <div
                        className={`text-2xl sm:text-4xl font-bold tabular-nums ${
                          hasLightBg ? 'text-slate-900' : 'text-white'
                        }`}
                      >
                        {slide.metricHighlight.value}
                      </div>
                      <div
                        className={`mt-1.5 text-xs sm:text-sm ${
                          hasLightBg ? 'text-slate-700' : 'text-slate-200'
                        }`}
                      >
                        {slide.metricHighlight.label}
                      </div>
                    </div>
                  )
                )}
              </div>
            ) : (
              <div className="space-y-4">
                <h2
                  className={`text-xl sm:text-3xl font-bold leading-tight ${
                    hasLightBg ? 'text-slate-900' : 'text-white'
                  }`}
                >
                  {slide.title}
                </h2>
                {slide.subtitle && (
                  <p
                    className={`text-sm sm:text-base ${
                      hasLightBg ? 'text-slate-600' : 'text-slate-300'
                    }`}
                  >
                    {slide.subtitle}
                  </p>
                )}
                {slide.bullets.length > 0 && (
                  <ul className="space-y-2.5 pt-2">
                    {slide.bullets.map((bullet, idx) => (
                      <li
                        key={`${slide.id}-sh-b-${idx}`}
                        className={`flex items-start gap-2.5 text-sm leading-relaxed ${
                          hasLightBg ? 'text-slate-800' : 'text-slate-100'
                        }`}
                      >
                        <span
                          className="mt-2 h-2 w-2 rounded-full shrink-0"
                          style={{
                            backgroundColor: slide.accentColor,
                          }}
                        />
                        <span>{bullet}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Dedicated Control Bar Below Slide Surface */}
      <div className="px-4 sm:px-6 py-3 bg-[#151B28] border-t border-white/10 flex flex-wrap items-center justify-between gap-3 text-white">
        <div className="flex items-center gap-2">
          <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-200">
            <Layers className="h-3.5 w-3.5 text-indigo-400" />
            <span className="tabular-nums">
              Slide {slideIndex + 1} / {totalSlides}
            </span>
          </span>
        </div>

        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={onPrevSlide}
            disabled={slideIndex === 0}
            className="inline-flex items-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-semibold text-white disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <ChevronLeft className="h-4 w-4" />
            <span>Trước</span>
          </button>

          <button
            type="button"
            onClick={onNextSlide}
            disabled={slideIndex >= totalSlides - 1}
            className="inline-flex items-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-semibold text-white disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <span>Tiếp</span>
            <ChevronRight className="h-4 w-4" />
          </button>

          <button
            type="button"
            onClick={onOpenPresenter}
            className="inline-flex items-center gap-1.5 min-h-[36px] px-3 rounded-xl bg-[#4F46E5] hover:bg-indigo-600 text-xs font-semibold text-white transition-colors"
          >
            <Presentation className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Trình chiếu</span>
          </button>
        </div>
      </div>
    </div>
  );
}

export function SharedDocumentView({
  shareToken,
  onNavigateHome,
}: SharedDocumentViewProps) {
  const [resolution, setResolution] = useState<SharedDocumentResolution>({
    status: 'loading',
    shareToken,
  });

  // Password gate state
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [showPassword, setShowPassword] = useState<boolean>(false);
  const [passwordError, setPasswordError] = useState<string | null>(null);
  const [isUnlocking, setIsUnlocking] = useState<boolean>(false);

  // Viewer navigation states
  const [slideIdx, setSlideIdx] = useState<number>(0);
  const [activeSheetIdx, setActiveSheetIdx] = useState<number>(0);
  const [activeHeadingId, setActiveHeadingId] = useState<string>('');
  const [fallbackPdfPageIdx, setFallbackPdfPageIdx] = useState<number>(0);
  const [isDownloading, setIsDownloading] = useState<boolean>(false);
  const [isPresenterOpen, setIsPresenterOpen] = useState<boolean>(false);

  // Real PDF.js rendering states
  const [pdfDoc, setPdfDoc] = useState<PdfDocumentProxy | null>(null);
  const [pdfNumPages, setPdfNumPages] = useState<number>(0);
  const [pdfPage, setPdfPage] = useState<number>(1);
  const [pdfZoomPercent, setPdfZoomPercent] = useState<number>(100);
  const [isPdfLoading, setIsPdfLoading] = useState<boolean>(false);
  const [isPdfRenderingPage, setIsPdfRenderingPage] = useState<boolean>(false);
  const [pdfError, setPdfError] = useState<string | null>(null);
  const [pdfContainerWidth, setPdfContainerWidth] = useState<number>(0);

  const pdfCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const pdfViewportWrapRef = useRef<HTMLDivElement | null>(null);

  const loadShare = useCallback(async () => {
    setResolution({ status: 'loading', shareToken });
    setPasswordInput('');
    setShowPassword(false);
    setPasswordError(null);
    setIsUnlocking(false);
    setSlideIdx(0);
    setActiveSheetIdx(0);
    setActiveHeadingId('');
    setFallbackPdfPageIdx(0);
    setIsPresenterOpen(false);
    setPdfPage(1);
    setPdfZoomPercent(100);
    setPdfError(null);

    try {
      const result = await resolvePublicShareLink(shareToken);
      setResolution(result);
    } catch (error) {
      setResolution({
        status: 'error',
        shareToken,
        errorMessage:
          error instanceof Error
            ? error.message
            : 'Không thể tải dữ liệu chia sẻ.',
      });
    }
  }, [shareToken]);

  useEffect(() => {
    void loadShare();
  }, [loadShare]);

  const resolvedDoc = useMemo(() => {
    if (resolution.status !== 'ready') {
      return null;
    }
    return buildResolvedDocument(resolution);
  }, [resolution]);

  const pdfSourceUrl = useMemo(() => {
    if (!resolvedDoc || resolvedDoc.type !== 'pdf') {
      return '';
    }
    return (
      resolution.downloadUrl ??
      resolvedDoc.fileUrl ??
      resolvedDoc.objectUrl ??
      ''
    );
  }, [resolution.downloadUrl, resolvedDoc]);

  // Load PDF document via PDF.js when ready & URL is available
  useEffect(() => {
    if (resolution.status !== 'ready' || !pdfSourceUrl) {
      setPdfDoc(null);
      setPdfNumPages(0);
      setIsPdfLoading(false);
      return;
    }

    let cancelled = false;
    let loadedProxy: PdfDocumentProxy | null = null;
    setIsPdfLoading(true);
    setPdfError(null);

    void loadPdfDocumentFromSource(pdfSourceUrl)
      .then((proxy) => {
        if (cancelled) {
          void proxy.destroy();
          return;
        }
        loadedProxy = proxy;
        setPdfDoc(proxy);
        setPdfNumPages(proxy.numPages);
        setPdfPage(1);
      })
      .catch((error) => {
        if (cancelled) {
          return;
        }
        setPdfDoc(null);
        setPdfNumPages(0);
        setPdfError(
          error instanceof Error
            ? error.message
            : 'Không thể giải mã tệp PDF trên trình duyệt.'
        );
      })
      .finally(() => {
        if (!cancelled) {
          setIsPdfLoading(false);
        }
      });

    return () => {
      cancelled = true;
      if (loadedProxy) {
        void loadedProxy.destroy();
      }
    };
  }, [resolution.status, pdfSourceUrl]);

  // Measure PDF container width for responsive scaling
  useEffect(() => {
    const container = pdfViewportWrapRef.current;
    if (!container) {
      return;
    }

    const updateWidth = () => {
      setPdfContainerWidth(container.clientWidth);
    };
    updateWidth();

    if (typeof ResizeObserver === 'undefined') {
      window.addEventListener('resize', updateWidth);
      return () => window.removeEventListener('resize', updateWidth);
    }

    const observer = new ResizeObserver(updateWidth);
    observer.observe(container);
    return () => observer.disconnect();
  }, [resolution.status, pdfDoc]);

  // Render active PDF page onto canvas
  useEffect(() => {
    const canvas = pdfCanvasRef.current;
    if (!pdfDoc || !canvas || pdfNumPages <= 0) {
      return;
    }

    let cancelled = false;
    let cancelRender: (() => void) | null = null;
    const targetPage = Math.max(1, Math.min(pdfPage, pdfNumPages));
    const zoomScale = pdfZoomPercent / 100;
    const baseFitWidth =
      pdfContainerWidth > 0
        ? Math.max(260, Math.min(820, pdfContainerWidth - 48))
        : 760;

    setIsPdfRenderingPage(true);

    void renderPdfPageToCanvas({
      pdfDoc,
      pageNumber: targetPage,
      canvas,
      targetWidthCss: baseFitWidth,
      scale: zoomScale,
      maxDpr: 2,
    })
      .then((res) => {
        cancelRender = res.cancel;
        if (cancelled) {
          res.cancel();
        }
      })
      .catch(() => {
        if (!cancelled) {
          setPdfError('Không thể hiển thị trang PDF này.');
        }
      })
      .finally(() => {
        if (!cancelled) {
          setIsPdfRenderingPage(false);
        }
      });

    return () => {
      cancelled = true;
      if (cancelRender) {
        cancelRender();
      }
    };
  }, [pdfDoc, pdfNumPages, pdfPage, pdfZoomPercent, pdfContainerWidth]);

  // Keyboard navigation for PPTX slides
  const slidesCount = resolvedDoc?.slidesPreview.length ?? 0;
  useEffect(() => {
    if (
      resolution.status !== 'ready' ||
      resolvedDoc?.type !== 'pptx' ||
      isPresenterOpen ||
      slidesCount <= 1
    ) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (
        target?.tagName === 'INPUT' ||
        target?.tagName === 'TEXTAREA' ||
        target?.tagName === 'BUTTON'
      ) {
        return;
      }
      if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        event.preventDefault();
        setSlideIdx((prev) => (prev < slidesCount - 1 ? prev + 1 : prev));
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setSlideIdx((prev) => (prev > 0 ? prev - 1 : prev));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [resolution.status, resolvedDoc?.type, isPresenterOpen, slidesCount]);

  const wordContent = useMemo(() => {
    if (!resolvedDoc || resolvedDoc.type !== 'docx') {
      return null;
    }
    const rawHtml =
      resolvedDoc.wordHtml && resolvedDoc.wordHtml.trim().length > 0
        ? resolvedDoc.wordHtml
        : buildFallbackWordHtml(resolvedDoc);
    return sanitizeAndIndexWordHtml(rawHtml);
  }, [resolvedDoc]);

  const normalizedSheets = useMemo(() => {
    if (!resolvedDoc || resolvedDoc.type !== 'xlsx') {
      return [];
    }
    const sourceSheets = resolvedDoc.sheets ?? [];
    return sourceSheets.map((sheet, idx) =>
      normalizeSheetForReadView(sheet, idx)
    );
  }, [resolvedDoc]);

  const handlePasswordSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (isUnlocking) {
      return;
    }
    const trimmed = passwordInput.trim();
    if (!trimmed) {
      setPasswordError('Vui lòng nhập mật khẩu để mở khóa tài liệu.');
      return;
    }

    setIsUnlocking(true);
    setPasswordError(null);

    try {
      const unlocked = await unlockPasswordProtectedShare(shareToken, trimmed);
      if (unlocked.status === 'ready') {
        setResolution(unlocked);
      } else if (unlocked.status === 'password_required') {
        setPasswordError(
          unlocked.errorMessage ??
            'Mật khẩu chưa chính xác. Vui lòng kiểm tra lại.'
        );
      } else {
        setResolution(unlocked);
      }
    } catch (error) {
      setPasswordError(
        error instanceof Error
          ? error.message
          : 'Không thể xác thực mật khẩu. Vui lòng thử lại.'
      );
    } finally {
      setIsUnlocking(false);
    }
  };

  const handleDownload = async () => {
    if (!resolvedDoc || resolution.allowDownload !== true || isDownloading) {
      return;
    }

    if (resolution.downloadUrl) {
      const link = window.document.createElement('a');
      link.href = resolution.downloadUrl;
      link.download =
        resolution.fileName ||
        resolvedDoc.originalFileName ||
        resolvedDoc.title;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      window.document.body.appendChild(link);
      link.click();
      link.remove();
      return;
    }

    setIsDownloading(true);
    try {
      if (resolvedDoc.type === 'docx') {
        const htmlToExport =
          wordContent?.html ?? buildFallbackWordHtml(resolvedDoc);
        await exportWordToDocx(resolvedDoc.title, htmlToExport);
      } else if (resolvedDoc.type === 'xlsx') {
        const sheetsToExport: SpreadsheetSheet[] =
          resolvedDoc.sheets && resolvedDoc.sheets.length > 0
            ? resolvedDoc.sheets
            : [
                {
                  id: 'sheet-1',
                  name: 'Sheet 1',
                  columns: ['A', 'B', 'C'],
                  rows: [['', '', '']],
                },
              ];
        exportExcelToXlsx(resolvedDoc.title, sheetsToExport);
      } else if (resolvedDoc.type === 'pptx') {
        await exportSlidesToPptx(resolvedDoc.title, resolvedDoc.slidesPreview);
      }
    } finally {
      setIsDownloading(false);
    }
  };

  // 1. LOADING STATE
  if (resolution.status === 'loading') {
    return (
      <div className="min-h-[100dvh] bg-[#FAF8F5] text-slate-900 flex flex-col items-center justify-center p-4 pb-safe">
        <div className="w-full max-w-md rounded-3xl bg-white border border-[#E6E0D6] shadow-sm p-8 text-center space-y-4">
          <div className="mx-auto flex justify-center">
            <LiteLogoMark onClick={onNavigateHome} />
          </div>
          <div className="flex flex-col items-center gap-2 pt-2">
            <Loader2 className="h-7 w-7 animate-spin text-[#4F46E5]" />
            <h1 className="text-base font-semibold text-slate-900">
              Đang tải liên kết chia sẻ...
            </h1>
            <p className="text-xs text-slate-500 leading-relaxed">
              Đang xác thực quyền truy cập và chuẩn bị nội dung hiển thị.
            </p>
          </div>
        </div>
      </div>
    );
  }

  // 2. ACCESS BLOCKED / REVOKED / NOT FOUND / ERROR STATES
  if (
    resolution.status === 'private' ||
    resolution.status === 'revoked' ||
    resolution.status === 'not_found' ||
    resolution.status === 'error'
  ) {
    const isPrivateOrRevoked =
      resolution.status === 'private' || resolution.status === 'revoked';
    const heading = isPrivateOrRevoked
      ? 'Liên kết chia sẻ đã bị thu hồi hoặc đang ở chế độ Riêng tư'
      : resolution.status === 'not_found'
        ? 'Không tìm thấy liên kết chia sẻ'
        : 'Không thể mở liên kết chia sẻ';
    const description = isPrivateOrRevoked
      ? 'Chủ sở hữu đã chuyển tài liệu về Riêng tư hoặc thu hồi quyền truy cập công khai của liên kết này.'
      : resolution.status === 'not_found'
        ? 'Đường dẫn chia sẻ không tồn tại hoặc đã bị xóa khỏi hệ thống.'
        : (resolution.errorMessage ??
          'Đã xảy ra sự cố khi kết nối tới máy chủ. Vui lòng thử lại.');

    return (
      <div className="min-h-[100dvh] bg-[#FAF8F5] text-slate-900 flex flex-col items-center justify-center p-4 pb-safe">
        <div className="w-full max-w-md rounded-3xl bg-white border border-[#E6E0D6] shadow-lg p-6 sm:p-8 text-center space-y-5">
          <div className="mx-auto flex justify-center">
            <LiteLogoMark onClick={onNavigateHome} />
          </div>

          <div
            className={`mx-auto h-14 w-14 rounded-2xl border flex items-center justify-center ${
              isPrivateOrRevoked
                ? 'bg-amber-50 border-amber-200 text-amber-600'
                : 'bg-rose-50 border-rose-200 text-rose-600'
            }`}
          >
            {isPrivateOrRevoked ? (
              <ShieldAlert className="h-7 w-7" />
            ) : (
              <AlertCircle className="h-7 w-7" />
            )}
          </div>

          <div className="space-y-2">
            <h1 className="text-base sm:text-lg font-bold text-slate-900 leading-snug">
              {heading}
            </h1>
            <p className="text-xs sm:text-sm text-slate-600 leading-relaxed">
              {description}
            </p>
          </div>

          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-center gap-2.5 pt-1">
            {resolution.status === 'error' && (
              <button
                type="button"
                onClick={() => void loadShare()}
                className="min-h-[42px] px-4 rounded-xl border border-[#DDD6CA] bg-white hover:bg-[#FAF8F5] text-xs font-semibold text-slate-700 inline-flex items-center justify-center gap-1.5 transition-colors"
              >
                <RefreshCw className="h-3.5 w-3.5" />
                <span>Thử lại</span>
              </button>
            )}
            <button
              type="button"
              onClick={onNavigateHome}
              className="min-h-[42px] px-5 rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold shadow-sm shadow-indigo-600/20 inline-flex items-center justify-center gap-2 transition-colors"
            >
              <Home className="h-4 w-4" />
              <span>Về trang chủ LITE</span>
            </button>
          </div>
        </div>
      </div>
    );
  }

  // 3. PASSWORD GATE STATE
  if (resolution.status === 'password_required') {
    const badge = resolution.fileType
      ? TYPE_BADGE_STYLES[resolution.fileType]
      : null;

    return (
      <div className="min-h-[100dvh] bg-[#FAF8F5] text-slate-900 flex flex-col items-center justify-center p-4 pb-safe">
        <div className="w-full max-w-md rounded-3xl bg-white border border-[#E6E0D6] shadow-xl overflow-hidden">
          <div className="px-6 py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <LiteLogoMark onClick={onNavigateHome} />
            {badge && (
              <span
                className={`inline-flex items-center rounded-md border px-2 py-0.5 text-[10.5px] font-semibold ${badge.className}`}
              >
                {badge.label}
              </span>
            )}
          </div>

          <div className="p-6 space-y-5">
            <div className="flex items-start gap-3.5">
              <span className="h-11 w-11 rounded-2xl bg-indigo-50 border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                <KeyRound className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h1 className="text-base font-bold text-slate-900 leading-snug">
                  Tài liệu yêu cầu mật khẩu
                </h1>
                {resolution.title && (
                  <p className="text-xs font-medium text-slate-700 truncate mt-0.5">
                    {resolution.title}
                  </p>
                )}
                <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                  Nhập mật khẩu bảo vệ do người chia sẻ cung cấp để mở xem tài
                  liệu.
                </p>
              </div>
            </div>

            {passwordError && (
              <div
                role="alert"
                className="rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-2.5 flex items-start gap-2 text-xs text-rose-800"
              >
                <AlertCircle className="h-4 w-4 text-rose-600 shrink-0 mt-0.5" />
                <span>{passwordError}</span>
              </div>
            )}

            <form onSubmit={handlePasswordSubmit} className="space-y-4">
              <div className="space-y-1.5">
                <label
                  htmlFor="share-unlock-password"
                  className="block text-xs font-semibold text-slate-700"
                >
                  Mật khẩu mở khóa
                </label>
                <div className="relative">
                  <Lock
                    className="h-4 w-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none"
                    aria-hidden="true"
                  />
                  <input
                    id="share-unlock-password"
                    type={showPassword ? 'text' : 'password'}
                    value={passwordInput}
                    disabled={isUnlocking}
                    onChange={(event) => {
                      setPasswordInput(event.target.value);
                      if (passwordError) {
                        setPasswordError(null);
                      }
                    }}
                    placeholder="Nhập mật khẩu truy cập..."
                    autoFocus
                    className="w-full h-11 rounded-xl border border-[#DDD6CA] bg-[#FAF8F5]/60 pl-10 pr-11 text-sm text-slate-900 placeholder:text-slate-400 focus:bg-white focus:border-[#4F46E5] focus:outline-none focus:ring-2 focus:ring-indigo-500/15"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword((prev) => !prev)}
                    aria-label={
                      showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'
                    }
                    className="absolute right-1.5 top-1/2 -translate-y-1/2 h-8 w-8 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 flex items-center justify-center transition-colors"
                  >
                    {showPassword ? (
                      <EyeOff className="h-4 w-4" />
                    ) : (
                      <Eye className="h-4 w-4" />
                    )}
                  </button>
                </div>
              </div>

              <button
                type="submit"
                disabled={isUnlocking}
                className="w-full min-h-[44px] rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-sm font-semibold shadow-sm shadow-indigo-600/20 flex items-center justify-center gap-2 transition-colors disabled:opacity-60"
              >
                {isUnlocking ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Đang mở khóa...</span>
                  </>
                ) : (
                  <>
                    <KeyRound className="h-4 w-4" />
                    <span>Mở khóa tài liệu</span>
                  </>
                )}
              </button>
            </form>

            <div className="pt-1 text-center">
              <button
                type="button"
                onClick={onNavigateHome}
                className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-500 hover:text-slate-800 transition-colors"
              >
                <ArrowLeft className="h-3.5 w-3.5" />
                <span>Về trang chủ LITE</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // 4. READY STATE (PUBLIC BARRIER-FREE VIEWER)
  if (!resolvedDoc) {
    return null;
  }

  const badge = TYPE_BADGE_STYLES[resolvedDoc.type];
  const allowDownload = resolution.allowDownload === true;
  const viewCount = resolution.viewCount ?? 1;
  const slides: SlidePreviewItem[] = resolvedDoc.slidesPreview;
  const safeSlideIdx =
    slides.length > 0 ? Math.min(slideIdx, slides.length - 1) : 0;
  const activeSlide = slides[safeSlideIdx] ?? null;

  const pdfPages = resolvedDoc.pdfPages ?? [];
  const safeFallbackPdfIdx =
    pdfPages.length > 0 ? Math.min(fallbackPdfPageIdx, pdfPages.length - 1) : 0;
  const activeFallbackPdfPage = pdfPages[safeFallbackPdfIdx] ?? null;

  const canPresent =
    slides.length > 0 ||
    (resolvedDoc.type === 'pdf' &&
      (Boolean(pdfSourceUrl) || pdfPages.length > 0));

  const canDownloadNow =
    allowDownload &&
    (Boolean(resolution.downloadUrl) ||
      resolvedDoc.type === 'docx' ||
      resolvedDoc.type === 'xlsx' ||
      resolvedDoc.type === 'pptx');

  const safeSheetIdx =
    normalizedSheets.length > 0
      ? Math.min(activeSheetIdx, normalizedSheets.length - 1)
      : 0;
  const activeSheet = normalizedSheets[safeSheetIdx] ?? null;

  const handleJumpToHeading = (headingId: string) => {
    setActiveHeadingId(headingId);
    const target = window.document.getElementById(headingId);
    if (target) {
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  return (
    <div className="min-h-[100dvh] bg-[#FAF8F5] text-slate-900 flex flex-col">
      {/* Minimal Public Header */}
      <header className="sticky top-0 z-20 shrink-0 h-14 sm:h-16 bg-white/95 backdrop-blur-md border-b border-[#EAE6DF] px-3 sm:px-6 flex items-center justify-between gap-2.5">
        <div className="flex items-center gap-2.5 sm:gap-3 min-w-0">
          <LiteLogoMark onClick={onNavigateHome} />
          <span
            aria-hidden="true"
            className="hidden sm:block h-5 w-px bg-[#EAE6DF]"
          />
          <div className="min-w-0">
            <div className="flex items-center gap-2 min-w-0">
              <h1 className="text-xs sm:text-sm font-semibold text-slate-900 truncate">
                {resolvedDoc.title}
              </h1>
              <span
                className={`shrink-0 inline-flex items-center rounded border px-1.5 py-0 text-[10px] font-semibold ${badge.className}`}
              >
                {badge.label}
              </span>
            </div>
            <div className="flex items-center gap-2 text-[11px] text-slate-500 mt-0.5">
              <span className="inline-flex items-center gap-1 tabular-nums">
                <Eye className="h-3 w-3 text-slate-400" />
                <span>{viewCount.toLocaleString('vi-VN')} lượt xem</span>
              </span>
              {resolvedDoc.fileSize && (
                <>
                  <span>•</span>
                  <span className="truncate">{resolvedDoc.fileSize}</span>
                </>
              )}
            </div>
          </div>
        </div>

        <div className="shrink-0 flex items-center gap-1.5 sm:gap-2">
          {canPresent && (
            <button
              type="button"
              onClick={() => setIsPresenterOpen(true)}
              className="inline-flex items-center justify-center gap-1.5 min-h-[38px] px-2.5 sm:px-3.5 rounded-xl bg-[#4F46E5] hover:bg-indigo-700 text-white text-xs font-semibold shadow-sm shadow-indigo-600/20 transition-colors whitespace-nowrap"
            >
              <Presentation className="h-3.5 w-3.5 shrink-0" />
              <span className="hidden sm:inline">Trình chiếu</span>
            </button>
          )}

          {allowDownload ? (
            canDownloadNow && (
              <button
                type="button"
                onClick={() => void handleDownload()}
                disabled={isDownloading}
                className="inline-flex items-center justify-center gap-1.5 min-h-[38px] px-2.5 sm:px-3.5 rounded-xl border border-[#DDD6CA] bg-white hover:bg-[#FAF8F5] text-xs font-semibold text-slate-800 transition-colors disabled:opacity-60 whitespace-nowrap"
              >
                {isDownloading ? (
                  <Loader2 className="h-3.5 w-3.5 animate-spin shrink-0" />
                ) : (
                  <Download className="h-3.5 w-3.5 shrink-0 text-[#4F46E5]" />
                )}
                <span className="hidden sm:inline">Tải về</span>
              </button>
            )
          ) : (
            <span className="inline-flex items-center gap-1 min-h-[34px] px-2.5 rounded-xl bg-[#FAF8F5] border border-[#E6E0D6] text-[11px] font-medium text-slate-600 whitespace-nowrap">
              <Lock className="h-3 w-3 text-slate-400 shrink-0" />
              <span className="hidden md:inline">Chế độ chỉ xem</span>
              <span className="md:hidden">Chỉ xem</span>
            </span>
          )}

          <button
            type="button"
            onClick={onNavigateHome}
            className="inline-flex items-center justify-center gap-1.5 min-h-[38px] px-2.5 sm:px-3 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#F2EDE4] text-xs font-semibold text-slate-700 transition-colors whitespace-nowrap"
          >
            <Home className="h-3.5 w-3.5 shrink-0" />
            <span className="hidden sm:inline">Mở LITE</span>
          </button>
        </div>
      </header>

      {/* Main Content By Type */}
      <main className="flex-1 min-h-0 flex flex-col">
        {/* DOCX VIEW */}
        {resolvedDoc.type === 'docx' && wordContent && (
          <div className="flex-1 flex flex-col md:flex-row">
            {/* Mobile Heading Strip */}
            {wordContent.headings.length > 0 && (
              <div className="md:hidden shrink-0 border-b border-[#EAE5DC] bg-[#F5F1EA] px-3 py-2 flex items-center gap-2 overflow-x-auto no-scrollbar">
                {wordContent.headings.map((h) => (
                  <button
                    key={h.id}
                    type="button"
                    onClick={() => handleJumpToHeading(h.id)}
                    className={`shrink-0 px-3 py-1.5 rounded-xl text-xs whitespace-nowrap transition-colors ${
                      activeHeadingId === h.id
                        ? 'bg-white text-[#4F46E5] font-semibold shadow-sm'
                        : 'bg-white/70 text-slate-700 hover:bg-white'
                    }`}
                  >
                    {h.label}
                  </button>
                ))}
              </div>
            )}

            {/* Desktop Outline Sidebar */}
            {wordContent.headings.length > 0 && (
              <aside className="hidden md:flex w-[250px] shrink-0 border-r border-[#EAE5DC] bg-[#F5F1EA]/70 flex-col">
                <div className="px-4 py-3.5 border-b border-[#EAE5DC] flex items-center gap-2 text-xs font-semibold text-slate-700">
                  <BookOpen className="h-4 w-4 text-[#4F46E5]" />
                  <span>Mục lục tài liệu</span>
                </div>
                <nav className="flex-1 overflow-y-auto p-3 space-y-1">
                  {wordContent.headings.map((h) => (
                    <button
                      key={h.id}
                      type="button"
                      onClick={() => handleJumpToHeading(h.id)}
                      className={`w-full text-left rounded-xl px-3 py-2 text-xs leading-snug transition-colors ${
                        h.level === 2
                          ? 'pl-5'
                          : h.level === 3
                            ? 'pl-7 text-[11.5px]'
                            : 'font-semibold'
                      } ${
                        activeHeadingId === h.id
                          ? 'bg-white text-[#4F46E5] font-semibold shadow-sm ring-1 ring-indigo-500/20'
                          : 'text-slate-700 hover:bg-white/70'
                      }`}
                    >
                      {h.label}
                    </button>
                  ))}
                </nav>
              </aside>
            )}

            {/* A4 Read Sheet */}
            <div className="flex-1 min-w-0 px-3 py-5 sm:px-6 sm:py-8">
              <article className="max-w-[794px] mx-auto rounded-xl sm:rounded-2xl bg-white border border-[#E7E3DC] shadow-sm">
                <div className="px-6 sm:px-12 pt-5 pb-3 border-b border-[#ECE8E0] flex flex-wrap items-center justify-between gap-2 text-[11.5px] text-slate-500">
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full bg-blue-50 text-blue-700 font-semibold">
                    <FileText className="h-3 w-3" />
                    <span>Văn bản chia sẻ công khai</span>
                  </span>
                  {resolvedDoc.folder && <span>{resolvedDoc.folder}</span>}
                </div>
                <div
                  className="px-6 sm:px-12 py-6 sm:py-10 text-[15px] sm:text-[16px] text-[#1E293B] leading-[1.8] break-words [&_h1]:text-slate-900 [&_h1]:text-[1.7em] [&_h1]:font-bold [&_h1]:leading-snug [&_h1]:mb-4 [&_h1]:scroll-mt-20 [&_h2]:text-slate-900 [&_h2]:text-[1.32em] [&_h2]:font-bold [&_h2]:leading-snug [&_h2]:mt-7 [&_h2]:mb-3 [&_h2]:pb-2 [&_h2]:border-b [&_h2]:border-[#ECE8E0] [&_h2]:scroll-mt-20 [&_h3]:text-slate-800 [&_h3]:text-[1.12em] [&_h3]:font-semibold [&_h3]:mt-5 [&_h3]:mb-2 [&_h3]:scroll-mt-20 [&_p]:mb-3.5 [&_ul]:list-disc [&_ul]:pl-6 [&_ul]:mb-4 [&_ul]:space-y-1.5 [&_ol]:list-decimal [&_ol]:pl-6 [&_ol]:mb-4 [&_ol]:space-y-1.5 [&_li]:leading-[1.75] [&_blockquote]:border-l-4 [&_blockquote]:border-indigo-400 [&_blockquote]:bg-[#FAF8F5] [&_blockquote]:rounded-r-xl [&_blockquote]:px-4 [&_blockquote]:py-3 [&_blockquote]:my-4 [&_table]:w-full [&_table]:border-collapse [&_table]:my-5 [&_table]:text-[0.92em] [&_th]:bg-[#F8F7F4] [&_th]:border [&_th]:border-[#DDD7CE] [&_th]:px-3 [&_th]:py-2 [&_th]:text-left [&_th]:font-semibold [&_td]:border [&_td]:border-[#E5E0D8] [&_td]:px-3 [&_td]:py-2 [&_td]:align-top [&_img]:max-w-full [&_img]:h-auto [&_img]:rounded-xl [&_img]:my-4 [&_a]:text-indigo-600 [&_a]:underline"
                  dangerouslySetInnerHTML={{ __html: wordContent.html }}
                />
              </article>
            </div>
          </div>
        )}

        {/* PPTX VIEW */}
        {resolvedDoc.type === 'pptx' && (
          <div className="flex-1 max-w-5xl w-full mx-auto px-3 sm:px-6 py-5 sm:py-8 space-y-5">
            {activeSlide ? (
              <>
                <SharedSlideCanvas
                  doc={resolvedDoc}
                  slide={activeSlide}
                  slideIndex={safeSlideIdx}
                  totalSlides={slides.length}
                  onPrevSlide={() =>
                    setSlideIdx((prev) => Math.max(0, prev - 1))
                  }
                  onNextSlide={() =>
                    setSlideIdx((prev) =>
                      Math.min(slides.length - 1, prev + 1)
                    )
                  }
                  onOpenPresenter={() => setIsPresenterOpen(true)}
                />

                {/* Slide Thumbnails Strip */}
                {slides.length > 1 && (
                  <div className="flex items-center gap-2.5 overflow-x-auto pb-2 no-scrollbar">
                    {slides.map((slide, idx) => {
                      const isCurrent = idx === safeSlideIdx;
                      return (
                        <button
                          key={slide.id}
                          type="button"
                          onClick={() => setSlideIdx(idx)}
                          className={`shrink-0 w-40 rounded-2xl border p-3 text-left transition-all ${
                            isCurrent
                              ? 'border-[#4F46E5] bg-white shadow-md ring-2 ring-indigo-500/20'
                              : 'border-[#E6E0D6] bg-white/75 hover:bg-white'
                          }`}
                        >
                          <div className="flex items-center justify-between gap-1 text-[10.5px] font-semibold text-slate-500">
                            <span>Slide {idx + 1}</span>
                            <span
                              className="h-2 w-2 rounded-full"
                              style={{ backgroundColor: slide.accentColor }}
                            />
                          </div>
                          <p className="mt-1 text-xs font-semibold text-slate-900 truncate">
                            {slide.title || `Slide ${idx + 1}`}
                          </p>
                        </button>
                      );
                    })}
                  </div>
                )}

                {activeSlide.speakerNote && (
                  <div className="rounded-2xl border border-[#E6E0D6] bg-white p-4 space-y-1">
                    <span className="text-[11px] font-semibold uppercase tracking-wider text-slate-500">
                      Ghi chú diễn giả
                    </span>
                    <p className="text-xs sm:text-sm text-slate-700 leading-relaxed whitespace-pre-wrap">
                      {activeSlide.speakerNote}
                    </p>
                  </div>
                )}
              </>
            ) : (
              <div className="rounded-2xl border border-[#E6E0D6] bg-white p-8 text-center text-sm text-slate-500">
                Bài thuyết trình chưa có slide nào.
              </div>
            )}
          </div>
        )}

        {/* XLSX VIEW */}
        {resolvedDoc.type === 'xlsx' && (
          <div className="flex-1 max-w-6xl w-full mx-auto px-3 sm:px-6 py-5 sm:py-8 flex flex-col gap-4 min-h-0">
            {activeSheet ? (
              <>
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="inline-flex items-center gap-2 text-xs font-semibold text-slate-700">
                    <FileSpreadsheet className="h-4 w-4 text-emerald-600" />
                    <span>
                      {normalizedSheets.length} trang tính · Chế độ xem công
                      khai
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar">
                    {normalizedSheets.map((sheet, idx) => {
                      const isActive = idx === safeSheetIdx;
                      return (
                        <button
                          key={sheet.id}
                          type="button"
                          onClick={() => setActiveSheetIdx(idx)}
                          className={`shrink-0 min-h-[34px] px-3 py-1 rounded-xl text-xs font-semibold transition-all ${
                            isActive
                              ? 'bg-[#4F46E5] text-white shadow-sm'
                              : 'bg-white border border-[#E6E0D6] text-slate-700 hover:bg-[#F5F1EA]'
                          }`}
                        >
                          {sheet.name}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="flex-1 min-h-[320px] rounded-2xl border border-[#DDD6CA] bg-white shadow-sm overflow-auto">
                  <table className="w-full border-collapse text-xs">
                    <thead className="sticky top-0 z-10 bg-[#F5F1EA] text-slate-600">
                      <tr>
                        <th className="w-12 border-b border-r border-[#E2DCD2] px-2 py-2 text-center font-semibold">
                          #
                        </th>
                        {Array.from(
                          { length: activeSheet.columnCount },
                          (_, colIdx) => (
                            <th
                              key={colIdx}
                              className="min-w-[110px] border-b border-r border-[#E2DCD2] px-3 py-2 text-center font-semibold"
                            >
                              {columnIndexToLetter(colIdx)}
                            </th>
                          )
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {activeSheet.evaluatedRows.map((row, rowIdx) => (
                        <tr
                          key={rowIdx}
                          className="hover:bg-[#FAF8F5]/80 transition-colors"
                        >
                          <td className="border-b border-r border-[#EAE5DC] bg-[#FAF8F5] px-2 py-2 text-center font-medium text-slate-500 tabular-nums">
                            {rowIdx + 1}
                          </td>
                          {Array.from(
                            { length: activeSheet.columnCount },
                            (_, colIdx) => {
                              const rawVal = row[colIdx] ?? '';
                              const style =
                                activeSheet.cellStyles[`${rowIdx}:${colIdx}`];
                              const isFormulaErr =
                                typeof rawVal === 'string' &&
                                FORMULA_ERROR_CODES.has(rawVal);
                              const isNumeric = typeof rawVal === 'number';
                              const alignClass =
                                style?.align === 'center'
                                  ? 'text-center'
                                  : style?.align === 'right' ||
                                      (!style?.align && isNumeric)
                                    ? 'text-right'
                                    : 'text-left';

                              return (
                                <td
                                  key={colIdx}
                                  className={`border-b border-r border-[#EFECE6] px-3 py-2 whitespace-nowrap ${alignClass} ${
                                    style?.bold
                                      ? 'font-semibold text-slate-900'
                                      : 'text-slate-700'
                                  } ${
                                    isFormulaErr
                                      ? 'text-rose-600 font-semibold bg-rose-50/50'
                                      : ''
                                  } ${isNumeric ? 'tabular-nums' : ''}`}
                                >
                                  {formatCellValue(rawVal)}
                                </td>
                              );
                            }
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            ) : (
              <div className="rounded-2xl border border-[#E6E0D6] bg-white p-8 text-center text-sm text-slate-500">
                Bảng tính hiện chưa có dữ liệu.
              </div>
            )}
          </div>
        )}

        {/* PDF VIEW */}
        {resolvedDoc.type === 'pdf' && (
          <div className="flex-1 max-w-5xl w-full mx-auto px-3 sm:px-6 py-4 sm:py-6 flex flex-col gap-4">
            {pdfSourceUrl && !pdfError ? (
              <>
                {/* PDF Controls Bar */}
                <div className="rounded-2xl border border-[#E6E0D6] bg-white px-3.5 py-2.5 flex flex-wrap items-center justify-between gap-2 shadow-sm">
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => setPdfPage((prev) => Math.max(1, prev - 1))}
                      disabled={pdfPage <= 1 || isPdfLoading}
                      aria-label="Trang trước"
                      className="h-9 w-9 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#EFECE6] text-slate-700 flex items-center justify-center disabled:opacity-40"
                    >
                      <ChevronLeft className="h-4 w-4" />
                    </button>
                    <span className="px-2 text-xs font-semibold text-slate-800 tabular-nums">
                      Trang {pdfNumPages > 0 ? pdfPage : 0} / {pdfNumPages}
                    </span>
                    <button
                      type="button"
                      onClick={() =>
                        setPdfPage((prev) =>
                          pdfNumPages > 0 ? Math.min(pdfNumPages, prev + 1) : prev
                        )
                      }
                      disabled={
                        pdfNumPages === 0 ||
                        pdfPage >= pdfNumPages ||
                        isPdfLoading
                      }
                      aria-label="Trang tiếp"
                      className="h-9 w-9 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#EFECE6] text-slate-700 flex items-center justify-center disabled:opacity-40"
                    >
                      <ChevronRight className="h-4 w-4" />
                    </button>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        setPdfZoomPercent((prev) => {
                          const idx = PDF_ZOOM_LEVELS.indexOf(
                            prev as (typeof PDF_ZOOM_LEVELS)[number]
                          );
                          return idx > 0
                            ? PDF_ZOOM_LEVELS[idx - 1]
                            : PDF_ZOOM_LEVELS[0];
                        })
                      }
                      disabled={pdfZoomPercent <= PDF_ZOOM_LEVELS[0]}
                      aria-label="Thu nhỏ"
                      className="h-9 w-9 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#EFECE6] text-slate-700 flex items-center justify-center disabled:opacity-40"
                    >
                      <ZoomOut className="h-4 w-4" />
                    </button>
                    <span className="min-w-[46px] text-center text-xs font-semibold text-slate-700 tabular-nums">
                      {pdfZoomPercent}%
                    </span>
                    <button
                      type="button"
                      onClick={() =>
                        setPdfZoomPercent((prev) => {
                          const idx = PDF_ZOOM_LEVELS.indexOf(
                            prev as (typeof PDF_ZOOM_LEVELS)[number]
                          );
                          return idx >= 0 && idx < PDF_ZOOM_LEVELS.length - 1
                            ? PDF_ZOOM_LEVELS[idx + 1]
                            : PDF_ZOOM_LEVELS[PDF_ZOOM_LEVELS.length - 1];
                        })
                      }
                      disabled={
                        pdfZoomPercent >=
                        PDF_ZOOM_LEVELS[PDF_ZOOM_LEVELS.length - 1]
                      }
                      aria-label="Phóng to"
                      className="h-9 w-9 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#EFECE6] text-slate-700 flex items-center justify-center disabled:opacity-40"
                    >
                      <ZoomIn className="h-4 w-4" />
                    </button>
                  </div>
                </div>

                {/* PDF Canvas Viewport */}
                <div
                  ref={pdfViewportWrapRef}
                  className="flex-1 rounded-2xl border border-[#E6E0D6] bg-[#EFECE6]/70 p-3 sm:p-6 overflow-auto flex flex-col items-center justify-start min-h-[420px] relative"
                >
                  {(isPdfLoading || isPdfRenderingPage) && (
                    <div className="mb-3 inline-flex items-center gap-2 rounded-full bg-white/95 border border-[#E6E0D6] px-3.5 py-1.5 text-xs font-medium text-slate-700 shadow-sm">
                      <Loader2 className="h-3.5 w-3.5 animate-spin text-[#4F46E5]" />
                      <span>
                        {isPdfLoading
                          ? 'Đang nạp tệp PDF...'
                          : 'Đang dựng trang PDF...'}
                      </span>
                    </div>
                  )}

                  <canvas
                    ref={pdfCanvasRef}
                    className="rounded-lg shadow-lg bg-white shrink-0"
                  />
                </div>
              </>
            ) : activeFallbackPdfPage ? (
              <div className="space-y-4">
                <div className="rounded-2xl border border-[#E6E0D6] bg-white px-4 py-2.5 flex items-center justify-between gap-2">
                  <span className="text-xs font-semibold text-slate-700 tabular-nums">
                    Trang {safeFallbackPdfIdx + 1} / {pdfPages.length}
                  </span>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        setFallbackPdfPageIdx((prev) => Math.max(0, prev - 1))
                      }
                      disabled={safeFallbackPdfIdx === 0}
                      className="h-8 px-2.5 rounded-lg border border-[#E6E0D6] text-xs font-semibold text-slate-700 disabled:opacity-40"
                    >
                      Trước
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        setFallbackPdfPageIdx((prev) =>
                          Math.min(pdfPages.length - 1, prev + 1)
                        )
                      }
                      disabled={safeFallbackPdfIdx >= pdfPages.length - 1}
                      className="h-8 px-2.5 rounded-lg border border-[#E6E0D6] text-xs font-semibold text-slate-700 disabled:opacity-40"
                    >
                      Tiếp
                    </button>
                  </div>
                </div>

                <article className="max-w-3xl mx-auto rounded-2xl bg-white border border-[#E6E0D6] shadow-sm p-6 sm:p-10 space-y-4">
                  <span className="text-[11px] font-semibold uppercase tracking-wider text-[#4F46E5]">
                    {activeFallbackPdfPage.sectionTitle}
                  </span>
                  <h2 className="text-xl font-bold text-slate-900">
                    {activeFallbackPdfPage.heading}
                  </h2>
                  <div className="space-y-3 text-sm text-slate-700 leading-relaxed">
                    {activeFallbackPdfPage.paragraphs.map((p, idx) => (
                      <p key={idx}>{p}</p>
                    ))}
                  </div>
                  {activeFallbackPdfPage.callout && (
                    <div className="rounded-xl bg-[#FAF8F5] border border-[#E6E0D6] p-4 text-xs sm:text-sm text-slate-700">
                      {activeFallbackPdfPage.callout}
                    </div>
                  )}
                </article>
              </div>
            ) : (
              <div className="rounded-2xl border border-[#E6E0D6] bg-white p-8 text-center space-y-3">
                <AlertCircle className="h-8 w-8 text-amber-500 mx-auto" />
                <p className="text-sm font-semibold text-slate-900">
                  {pdfError ??
                    'Không có tệp PDF gốc để hiển thị trực tiếp trong bản xem thử này.'}
                </p>
                {resolvedDoc.outline.length > 0 && (
                  <div className="max-w-xl mx-auto text-left pt-3 space-y-2 border-t border-slate-100">
                    {resolvedDoc.outline.map((sec) => (
                      <div
                        key={sec.id}
                        className="rounded-xl bg-[#FAF8F5] border border-[#EAE5DC] p-3"
                      >
                        <p className="text-xs font-semibold text-slate-900">
                          {sec.romanIndex}. {sec.title}
                        </p>
                        {sec.summary && (
                          <p className="text-xs text-slate-600 mt-1">
                            {sec.summary}
                          </p>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </main>

      {/* Presentation Overlay via PresenterStageView */}
      {isPresenterOpen && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={`Trình chiếu ${resolvedDoc.title}`}
          className="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md pb-safe"
        >
          <div
            onClick={(event) => {
              if (event.target === event.currentTarget) {
                setIsPresenterOpen(false);
              }
            }}
            className="h-full w-full flex items-center justify-center p-2 sm:p-5"
          >
            <div className="w-full max-w-6xl h-full max-h-full">
              <PresenterStageView
                document={resolvedDoc}
                pairingCode={buildPairingCode(resolvedDoc.id)}
                onClose={() => setIsPresenterOpen(false)}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

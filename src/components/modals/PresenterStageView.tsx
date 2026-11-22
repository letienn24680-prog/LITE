import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react';
import {
  Check,
  ChevronLeft,
  ChevronRight,
  Clock,
  Columns2,
  Copy,
  Crosshair,
  Eraser,
  ExternalLink,
  FileText,
  Flashlight,
  Highlighter,
  Loader2,
  Maximize2,
  MessageSquareText,
  Minimize2,
  Monitor,
  MousePointer2,
  Pause,
  Pencil,
  Play,
  Presentation,
  QrCode,
  Radio,
  RotateCcw,
  Smartphone,
  Undo2,
  X,
} from 'lucide-react';
import {
  createPresenterRemoteSession,
  type RemoteCommandPayload,
  type RemoteTransportMode,
} from '../../services/realtimeRemote';
import type {
  AnnotationStroke,
  DocumentItem,
  NormalizedPoint,
  PdfPageItem,
  RemoteChannelState,
  SlidePreviewItem,
  SlideTextBlock,
  SpotlightState,
} from '../../types/document';
import {
  loadPdfDocumentFromSource,
  renderPdfPageToCanvas,
  renderPdfThumbnailDataUrl,
  type PdfDocumentProxy,
} from '../../utils/pdfEngine';
import { RealQrCode } from '../common/RealQrCode';

export interface PresenterStageViewProps {
  document: DocumentItem;
  pairingCode: string;
  onClose: () => void;
  onOpenQrModal?: () => void;
}

type StageToolMode = 'pointer' | 'pen' | 'highlighter' | 'laser' | 'spotlight';
type PresenterViewMode = 'audience' | 'presenter';

interface PenColorOption {
  label: string;
  hex: string;
}

const PEN_COLORS: readonly PenColorOption[] = [
  { label: 'Đỏ', hex: '#EF4444' },
  { label: 'Vàng', hex: '#F59E0B' },
  { label: 'Xanh lá', hex: '#22C55E' },
  { label: 'Xanh dương', hex: '#38BDF8' },
];

const HIGHLIGHTER_WIDTH_PX = 18;
const PEN_WIDTH_PX = 3.5;
const DEFAULT_SPOTLIGHT_RADIUS = 0.16;

function formatTimer(totalSeconds: number): string {
  const safe = Math.max(0, Math.floor(totalSeconds));
  const hours = Math.floor(safe / 3600);
  const minutes = Math.floor((safe % 3600) / 60);
  const seconds = safe % 60;
  return [hours, minutes, seconds]
    .map((unit) => String(unit).padStart(2, '0'))
    .join(':');
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

function getResponsiveBlockFontSize(
  block: SlideTextBlock,
  compact = false
): string {
  const rawBase = Math.max(13, Math.min(54, block.fontSize || 18));
  const base = compact ? Math.round(rawBase * 0.68) : rawBase;
  const minPx = Math.max(compact ? 10 : 13, Math.round(base * 0.72));
  const vw = (base * (compact ? 0.08 : 0.12)).toFixed(2);
  return `clamp(${minPx}px, calc(10px + ${vw}vw), ${base}px)`;
}

function buildRemotePairingUrl(roomCode: string): string {
  if (typeof window === 'undefined') {
    return `#remote/${encodeURIComponent(roomCode)}`;
  }
  return `${window.location.origin}${window.location.pathname}#remote/${encodeURIComponent(roomCode)}`;
}

function buildFallbackSlides(document: DocumentItem): SlidePreviewItem[] {
  if (document.slidesPreview.length > 0) {
    return document.slidesPreview;
  }

  if (document.pdfPages && document.pdfPages.length > 0) {
    return document.pdfPages.map((page: PdfPageItem, idx: number) => ({
      id: page.id,
      index: idx + 1,
      title: page.heading || page.sectionTitle || `Trang ${page.pageNumber}`,
      subtitle: page.sectionTitle,
      bullets:
        page.highlights && page.highlights.length > 0
          ? page.highlights
          : page.paragraphs.slice(0, 3),
      accentColor: '#DB2777',
      speakerNote: page.callout,
    }));
  }

  if (document.outline.length > 0) {
    return document.outline.map((section, idx) => ({
      id: `${document.id}-outline-slide-${section.id}`,
      index: idx + 1,
      title: `${section.romanIndex}. ${section.title}`,
      subtitle: `Trang ${section.pageNumber}`,
      bullets:
        section.keyPoints.length > 0
          ? section.keyPoints
          : section.summary
            ? [section.summary]
            : [],
      accentColor: '#4F46E5',
      speakerNote: section.summary,
    }));
  }

  return [
    {
      id: `${document.id}-cover-slide`,
      index: 1,
      title: document.title,
      subtitle: `${document.type.toUpperCase()}${document.fileSize ? ` · ${document.fileSize}` : ''}`,
      bullets: document.contentPreview ? [document.contentPreview] : [],
      accentColor: '#4F46E5',
    },
  ];
}

function clampUnit(value: number): number {
  if (value < 0) {
    return 0;
  }
  if (value > 1) {
    return 1;
  }
  return value;
}

function isInteractiveKeyboardTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  const tag = target.tagName;
  return (
    tag === 'INPUT' ||
    tag === 'TEXTAREA' ||
    tag === 'SELECT' ||
    tag === 'BUTTON' ||
    target.isContentEditable
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

function buildSlideCanvasBackground(
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

  style.backgroundImage = `radial-gradient(circle at 86% 18%, ${slide.accentColor}38 0%, transparent 45%), linear-gradient(145deg, #161D2E 0%, #1F293F 58%, #131926 100%)`;
  return style;
}

interface SlideCardCanvasProps {
  doc: DocumentItem;
  slide: SlidePreviewItem;
  slideIndex: number;
  totalSlides: number;
  compact?: boolean;
}

function SlideCardCanvas({
  doc,
  slide,
  slideIndex,
  totalSlides,
  compact = false,
}: SlideCardCanvasProps) {
  const cardRef = useRef<HTMLDivElement | null>(null);
  const [cardWidthPx, setCardWidthPx] = useState<number>(compact ? 260 : 880);

  useEffect(() => {
    const el = cardRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateWidth = () => {
      const width = el.getBoundingClientRect().width;
      if (width > 0) {
        setCardWidthPx(width);
      }
    };
    updateWidth();
    const observer = new ResizeObserver(updateWidth);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, []);

  const isPositioned = isPositionedSlideItem(slide);
  const fontScale = cardWidthPx > 0 ? cardWidthPx / 960 : compact ? 0.27 : 0.9;
  const hasCustomLightBg = isPositioned
    ? inferPositionedLightBackground(slide)
    : isLightHexColor(slide.backgroundColor);
  const backgroundStyle = buildSlideCanvasBackground(
    slide,
    isPositioned,
    hasCustomLightBg
  );
  const defaultTextColor = hasCustomLightBg ? '#0F172A' : '#FFFFFF';

  if (isPositioned) {
    const positionedBlocks = (slide.blocks ?? []).filter(
      (b) =>
        b.kind === 'table' ||
        b.text.trim().length > 0 ||
        Boolean(b.fillColor) ||
        Boolean(b.borderColor)
    );
    const positionedImages = slide.images ?? [];

    return (
      <div
        ref={cardRef}
        className={`w-full h-full relative overflow-hidden select-none ${
          compact ? 'rounded-xl' : 'rounded-2xl'
        }`}
        style={backgroundStyle}
      >
        {positionedImages.map((img) => {
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
                className="w-full h-full object-contain pointer-events-none select-none"
              />
            </div>
          );
        })}

        {positionedBlocks.map((block) => {
          const isTable =
            block.kind === 'table' &&
            Array.isArray(block.tableRows) &&
            block.tableRows.length > 0;
          const baseFontPt = block.fontSize || (isTable ? 12 : 18);
          const scaledFontPx = Math.max(
            compact ? 3.5 : 8,
            baseFontPt * fontScale
          );
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
                  block.fillColor || block.borderColor ? '0.25rem' : undefined,
                padding: isTable ? 0 : compact ? '1px 2px' : '2px 5px',
              }}
            >
              {isTable && block.tableRows ? (
                <table
                  className="w-full h-full border-collapse table-fixed"
                  style={{
                    fontSize: `${scaledFontPx.toFixed(2)}px`,
                    lineHeight: 1.2,
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
                              ? hasCustomLightBg
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
                                (hasCustomLightBg
                                  ? 'rgba(15, 23, 42, 0.18)'
                                  : 'rgba(255, 255, 255, 0.22)')
                              }`,
                              padding: compact ? '1px 2px' : '3px 6px',
                              fontWeight: rIdx === 0 || block.bold ? 600 : 400,
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
                    lineHeight: 1.24,
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

        {compact && (
          <span
            className={`absolute bottom-1.5 right-1.5 px-1.5 py-0.5 rounded-md text-[9px] font-semibold tabular-nums pointer-events-none ${
              hasCustomLightBg
                ? 'bg-slate-900/10 text-slate-800'
                : 'bg-black/45 text-white'
            }`}
          >
            {slideIndex + 1}/{totalSlides}
          </span>
        )}
      </div>
    );
  }

  const activeBlocks =
    slide.blocks?.filter((b) => b.text.trim().length > 0) ?? [];
  const activeImages = slide.images ?? [];
  const hasRealSlideContent =
    activeBlocks.length > 0 || activeImages.length > 0;

  return (
    <div
      ref={cardRef}
      className={`w-full h-full flex flex-col justify-between relative overflow-hidden select-none ${
        compact ? 'p-3.5 gap-2 rounded-xl' : 'p-5 sm:p-8 lg:p-10 gap-3.5 rounded-2xl'
      } ${hasCustomLightBg ? 'text-slate-900' : 'text-white'}`}
      style={backgroundStyle}
    >
      {/* Top Strip */}
      <div className="flex items-center justify-between gap-2 shrink-0">
        <div className="inline-flex items-center gap-2 min-w-0">
          <span
            className={`${compact ? 'h-1.5 w-4' : 'h-2 sm:h-2.5 w-6 sm:w-8'} rounded-full shrink-0`}
            style={{ backgroundColor: slide.accentColor }}
          />
          <span
            className={`font-semibold uppercase tracking-widest truncate ${
              compact ? 'text-[9px]' : 'text-[11px] sm:text-xs'
            } ${hasCustomLightBg ? 'text-slate-600' : 'text-slate-300'}`}
          >
            {doc.folder ?? doc.type.toUpperCase()}
          </span>
        </div>
        <span
          className={`shrink-0 rounded-full border font-semibold tabular-nums ${
            compact
              ? 'px-2 py-0.5 text-[9px]'
              : 'px-2.5 sm:px-3 py-1 text-[11px] sm:text-xs'
          } ${
            hasCustomLightBg
              ? 'bg-slate-900/10 border-slate-900/15 text-slate-800'
              : 'bg-black/40 border-white/10 text-white'
          }`}
        >
          Slide {slideIndex + 1} / {totalSlides}
        </span>
      </div>

      {/* Main Content */}
      {hasRealSlideContent ? (
        <div
          className={`my-auto grid items-center min-h-0 overflow-hidden ${
            compact ? 'gap-2.5' : 'gap-4 sm:gap-7'
          } ${
            activeBlocks.length > 0 && activeImages.length > 0
              ? 'grid-cols-12'
              : activeBlocks.length > 0 && slide.metricHighlight
                ? 'grid-cols-12'
                : 'grid-cols-1'
          }`}
        >
          {activeBlocks.length > 0 && (
            <div
              className={`${compact ? 'space-y-1.5' : 'space-y-2.5 sm:space-y-3.5'} ${
                activeImages.length > 0
                  ? 'col-span-7'
                  : slide.metricHighlight
                    ? 'col-span-8'
                    : 'col-span-1'
              }`}
            >
              {activeBlocks.map((block) => (
                <p
                  key={block.id}
                  className="leading-snug whitespace-pre-wrap break-words"
                  style={{
                    fontSize: getResponsiveBlockFontSize(block, compact),
                    fontWeight:
                      block.bold || block.role === 'title' ? 700 : 400,
                    fontStyle: block.italic ? 'italic' : 'normal',
                    color:
                      block.color ?? (hasCustomLightBg ? '#0F172A' : '#FFFFFF'),
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
                activeBlocks.length > 0 ? 'col-span-5' : 'col-span-1'
              } grid gap-2 ${
                activeImages.length > 1 ? 'grid-cols-2' : 'grid-cols-1'
              }`}
            >
              {activeImages.map((img) => (
                <div
                  key={img.id}
                  className={`rounded-xl overflow-hidden border flex items-center justify-center p-1 ${
                    hasCustomLightBg
                      ? 'bg-slate-900/5 border-slate-300/70'
                      : 'bg-black/25 border-white/10'
                  }`}
                >
                  <img
                    src={img.src}
                    alt={img.alt ?? slide.title}
                    className={`w-full object-contain rounded-lg ${
                      compact
                        ? 'max-h-20'
                        : activeBlocks.length > 0
                          ? 'max-h-[30vh]'
                          : 'max-h-[50vh]'
                    }`}
                  />
                </div>
              ))}
            </div>
          ) : (
            slide.metricHighlight && (
              <div
                className={`col-span-4 rounded-2xl border flex flex-col items-end justify-center text-right backdrop-blur-md ${
                  compact ? 'p-2.5' : 'p-4 sm:p-6'
                }`}
                style={{
                  backgroundColor: `${slide.accentColor}22`,
                  borderColor: `${slide.accentColor}66`,
                }}
              >
                <div
                  className={`font-bold tracking-tight tabular-nums ${
                    compact ? 'text-lg' : 'text-2xl sm:text-4xl'
                  } ${hasCustomLightBg ? 'text-slate-900' : 'text-white'}`}
                >
                  {slide.metricHighlight.value}
                </div>
                <div
                  className={`mt-1 leading-snug ${
                    compact ? 'text-[10px]' : 'text-xs sm:text-sm'
                  } ${hasCustomLightBg ? 'text-slate-700' : 'text-slate-200/90'}`}
                >
                  {slide.metricHighlight.label}
                </div>
              </div>
            )
          )}
        </div>
      ) : (
        <div
          className={`my-auto grid items-center min-h-0 overflow-hidden ${
            compact ? 'gap-3' : 'gap-5 sm:gap-8'
          } ${slide.metricHighlight ? 'grid-cols-12' : 'grid-cols-1'}`}
        >
          <div className={slide.metricHighlight ? 'col-span-8' : 'col-span-1'}>
            <h2
              className={`font-bold leading-tight tracking-tight ${
                compact ? 'text-sm line-clamp-2' : 'text-xl sm:text-3xl'
              } ${hasCustomLightBg ? 'text-slate-900' : 'text-white'}`}
            >
              {slide.title}
            </h2>
            {slide.subtitle && (
              <p
                className={`mt-1 ${
                  compact
                    ? 'text-[11px] line-clamp-1'
                    : 'sm:mt-2 text-xs sm:text-base'
                } ${hasCustomLightBg ? 'text-slate-600' : 'text-slate-300'}`}
              >
                {slide.subtitle}
              </p>
            )}
            {slide.bullets.length > 0 && (
              <ul
                className={
                  compact
                    ? 'mt-2 space-y-1'
                    : 'mt-3 sm:mt-5 space-y-1.5 sm:space-y-2.5'
                }
              >
                {(compact ? slide.bullets.slice(0, 3) : slide.bullets).map(
                  (bullet, idx) => (
                    <li
                      key={`${slide.id}-b-${idx}`}
                      className={`flex items-start gap-2 leading-relaxed ${
                        compact
                          ? 'text-[10px] line-clamp-1'
                          : 'text-xs sm:text-[15px]'
                      } ${hasCustomLightBg ? 'text-slate-800' : 'text-slate-100/95'}`}
                    >
                      <span
                        className={`rounded-full shrink-0 ${
                          compact
                            ? 'mt-1 h-1 w-1'
                            : 'mt-1.5 sm:mt-2 h-1.5 w-1.5 sm:h-2 sm:w-2'
                        }`}
                        style={{ backgroundColor: slide.accentColor }}
                      />
                      <span>{bullet}</span>
                    </li>
                  )
                )}
              </ul>
            )}
          </div>

          {slide.metricHighlight && (
            <div
              className={`col-span-4 rounded-2xl border flex flex-col items-end justify-center text-right backdrop-blur-md ${
                compact ? 'p-2.5' : 'p-4 sm:p-6'
              }`}
              style={{
                backgroundColor: `${slide.accentColor}22`,
                borderColor: `${slide.accentColor}66`,
              }}
            >
              <div
                className={`font-bold tracking-tight tabular-nums ${
                  compact ? 'text-lg' : 'text-2xl sm:text-4xl'
                } ${hasCustomLightBg ? 'text-slate-900' : 'text-white'}`}
              >
                {slide.metricHighlight.value}
              </div>
              <div
                className={`mt-1 leading-snug ${
                  compact ? 'text-[10px]' : 'text-xs sm:text-sm'
                } ${hasCustomLightBg ? 'text-slate-700' : 'text-slate-200/90'}`}
              >
                {slide.metricHighlight.label}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Bottom Footer */}
      <div
        className={`shrink-0 flex items-center justify-between gap-2 pt-2 border-t ${
          hasCustomLightBg ? 'border-slate-900/10' : 'border-white/10'
        }`}
      >
        <p
          className={`truncate ${
            compact ? 'text-[9px]' : 'text-[11px] sm:text-xs'
          } ${hasCustomLightBg ? 'text-slate-600' : 'text-slate-300/80'}`}
        >
          {doc.title}
        </p>
        <span
          className={`shrink-0 tabular-nums ${
            compact ? 'text-[9px]' : 'text-[11px]'
          } ${hasCustomLightBg ? 'text-slate-600' : 'text-slate-400'}`}
        >
          {slideIndex + 1}/{totalSlides}
        </span>
      </div>
    </div>
  );
}

export function PresenterStageView({
  document,
  pairingCode,
  onClose,
  onOpenQrModal,
}: PresenterStageViewProps) {
  const realPdfUrl =
    document.type === 'pdf'
      ? (document.fileUrl ?? document.objectUrl ?? '')
      : '';
  const isRealPdfMode =
    Boolean(realPdfUrl) && document.slidesPreview.length === 0;

  const fallbackSlides = useMemo(
    () => (isRealPdfMode ? [] : buildFallbackSlides(document)),
    [document, isRealPdfMode]
  );

  // Core navigation & mode state
  const [slideIdx, setSlideIdx] = useState<number>(0);
  const [viewMode, setViewMode] = useState<PresenterViewMode>('audience');
  const [isFullscreen, setIsFullscreen] = useState<boolean>(() =>
    typeof window !== 'undefined'
      ? Boolean(window.document.fullscreenElement)
      : false
  );
  const [showInlineQr, setShowInlineQr] = useState<boolean>(false);
  const [copiedRemoteUrl, setCopiedRemoteUrl] = useState<boolean>(false);

  // PDF.js engine state
  const [pdfDoc, setPdfDoc] = useState<PdfDocumentProxy | null>(null);
  const [pdfTotalPages, setPdfTotalPages] = useState<number>(0);
  const [pdfLoading, setPdfLoading] = useState<boolean>(false);
  const [pdfError, setPdfError] = useState<string | null>(null);
  const [renderedPdfSize, setRenderedPdfSize] = useState<{
    width: number;
    height: number;
  } | null>(null);
  const [nextPdfThumbUrl, setNextPdfThumbUrl] = useState<string | null>(null);

  // Timer state
  const [elapsedSeconds, setElapsedSeconds] = useState<number>(0);
  const [isTimerRunning, setIsTimerRunning] = useState<boolean>(false);

  // Interactive Annotation / Laser / Spotlight state
  const [activeTool, setActiveTool] = useState<StageToolMode>('pointer');
  const [penColor, setPenColor] = useState<string>(PEN_COLORS[0].hex);
  const [strokesBySlide, setStrokesBySlide] = useState<
    Record<number, AnnotationStroke[]>
  >({});
  const [activeStroke, setActiveStroke] = useState<AnnotationStroke | null>(
    null
  );

  const [localLaser, setLocalLaser] = useState<{
    active: boolean;
    x: number;
    y: number;
  }>({ active: false, x: 0.5, y: 0.5 });
  const [remoteLaser, setRemoteLaser] = useState<{
    active: boolean;
    x: number;
    y: number;
  }>({ active: false, x: 0.5, y: 0.5 });

  const [localSpotlight, setLocalSpotlight] = useState<SpotlightState>({
    active: false,
    x: 0.5,
    y: 0.5,
    radius: DEFAULT_SPOTLIGHT_RADIUS,
  });
  const [remoteSpotlight, setRemoteSpotlight] = useState<SpotlightState>({
    active: false,
    x: 0.5,
    y: 0.5,
    radius: DEFAULT_SPOTLIGHT_RADIUS,
  });

  // Remote session state
  const [remoteStatus, setRemoteStatus] = useState<{
    mode: RemoteTransportMode;
    connected: boolean;
    peerCount: number;
  }>({
    mode: 'local-broadcast',
    connected: false,
    peerCount: 0,
  });

  // Measured stage dimensions for PDF fit + circular spotlight radius
  const [stageViewportSize, setStageViewportSize] = useState<{
    width: number;
    height: number;
  }>({ width: 960, height: 540 });
  const [activeSurfaceSize, setActiveSurfaceSize] = useState<{
    width: number;
    height: number;
  }>({ width: 960, height: 540 });

  const stageViewportRef = useRef<HTMLDivElement | null>(null);
  const interactiveSurfaceRef = useRef<HTMLDivElement | null>(null);
  const pdfCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const remoteSessionRef = useRef<ReturnType<
    typeof createPresenterRemoteSession
  > | null>(null);
  const activeStrokeRef = useRef<AnnotationStroke | null>(null);
  const drawingPointerIdRef = useRef<number | null>(null);
  const remoteLaserTimeoutRef = useRef<number | null>(null);
  const copyTimerRef = useRef<number | null>(null);

  const totalSlides = isRealPdfMode
    ? Math.max(pdfTotalPages || document.pageCount || 1, 1)
    : Math.max(fallbackSlides.length, 1);
  const safeSlideIdx = Math.min(Math.max(0, slideIdx), totalSlides - 1);

  const currentSlide = !isRealPdfMode ? fallbackSlides[safeSlideIdx] ?? null : null;
  const nextSlide =
    !isRealPdfMode && safeSlideIdx + 1 < totalSlides
      ? fallbackSlides[safeSlideIdx + 1] ?? null
      : null;

  const currentSlideTitle = useMemo(() => {
    if (isRealPdfMode) {
      const matchingOutline = document.outline.find(
        (item) => item.pageNumber === safeSlideIdx + 1
      );
      return matchingOutline
        ? `${matchingOutline.romanIndex}. ${matchingOutline.title}`
        : `Trang ${safeSlideIdx + 1} / ${totalSlides}`;
    }
    return currentSlide?.title ?? `Slide ${safeSlideIdx + 1}`;
  }, [
    currentSlide?.title,
    document.outline,
    isRealPdfMode,
    safeSlideIdx,
    totalSlides,
  ]);

  const currentSpeakerNote = useMemo(() => {
    if (isRealPdfMode) {
      const pageNote = document.pdfPages?.[safeSlideIdx]?.callout;
      if (pageNote) {
        return pageNote;
      }
      const outlineMatch = document.outline.find(
        (item) => item.pageNumber === safeSlideIdx + 1
      );
      return outlineMatch?.summary ?? '';
    }
    return currentSlide?.speakerNote ?? '';
  }, [
    currentSlide?.speakerNote,
    document.outline,
    document.pdfPages,
    isRealPdfMode,
    safeSlideIdx,
  ]);

  const currentSlideStrokes = strokesBySlide[safeSlideIdx] ?? [];
  const normalizedRoomCode = pairingCode.trim().toUpperCase() || 'LITE-0000';
  const remoteUrl = useMemo(
    () => buildRemotePairingUrl(normalizedRoomCode),
    [normalizedRoomCode]
  );

  // Reset per-document state cleanly when document changes
  useEffect(() => {
    setSlideIdx(0);
    setStrokesBySlide({});
    setActiveStroke(null);
    activeStrokeRef.current = null;
    setElapsedSeconds(0);
    setIsTimerRunning(false);
    setActiveTool('pointer');
    setLocalLaser({ active: false, x: 0.5, y: 0.5 });
    setLocalSpotlight({
      active: false,
      x: 0.5,
      y: 0.5,
      radius: DEFAULT_SPOTLIGHT_RADIUS,
    });
  }, [document.id]);

  // Load real PDF document via PDF.js
  useEffect(() => {
    if (!isRealPdfMode || !realPdfUrl) {
      setPdfDoc(null);
      setPdfTotalPages(0);
      setPdfLoading(false);
      setPdfError(null);
      setRenderedPdfSize(null);
      return;
    }

    let cancelled = false;
    let loadedProxy: PdfDocumentProxy | null = null;

    setPdfLoading(true);
    setPdfError(null);

    loadPdfDocumentFromSource(realPdfUrl)
      .then((docProxy) => {
        if (cancelled) {
          void docProxy.destroy().catch(() => undefined);
          return;
        }
        loadedProxy = docProxy;
        setPdfDoc(docProxy);
        setPdfTotalPages(docProxy.numPages);
        setPdfLoading(false);
      })
      .catch(() => {
        if (!cancelled) {
          setPdfDoc(null);
          setPdfTotalPages(0);
          setPdfLoading(false);
          setPdfError(
            'Không thể tải tệp PDF này để trình chiếu. Hãy thử mở lại tệp từ thư viện.'
          );
        }
      });

    return () => {
      cancelled = true;
      setPdfDoc(null);
      if (loadedProxy) {
        void loadedProxy.destroy().catch(() => undefined);
      }
    };
  }, [isRealPdfMode, realPdfUrl]);

  // Measure outer stage viewport for responsive PDF fitting
  useEffect(() => {
    const el = stageViewportRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateSize = () => {
      const rect = el.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        setStageViewportSize({
          width: Math.floor(rect.width),
          height: Math.floor(rect.height),
        });
      }
    };
    updateSize();
    const observer = new ResizeObserver(updateSize);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, [viewMode, isFullscreen]);

  const currentSlideAspectRatio = getSafeSlideAspectRatio(currentSlide, document);
  const nextSlideAspectRatio = getSafeSlideAspectRatio(nextSlide, document);

  const fittedSlideBox = useMemo(() => {
    const availW = Math.max(220, stageViewportSize.width - 36);
    const availH = Math.max(150, stageViewportSize.height - 36);
    const maxW = isFullscreen ? availW : Math.min(availW, 1024);
    if (maxW / availH > currentSlideAspectRatio) {
      const height = Math.floor(availH);
      const width = Math.floor(availH * currentSlideAspectRatio);
      return { width, height };
    }
    const width = Math.floor(maxW);
    const height = Math.floor(maxW / currentSlideAspectRatio);
    return { width, height };
  }, [stageViewportSize, currentSlideAspectRatio, isFullscreen]);

  // Measure active interactive surface for circular spotlight proportions
  useEffect(() => {
    const el = interactiveSurfaceRef.current;
    if (!el || typeof ResizeObserver === 'undefined') {
      return;
    }
    const updateSurface = () => {
      const rect = el.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        setActiveSurfaceSize({
          width: Math.floor(rect.width),
          height: Math.floor(rect.height),
        });
      }
    };
    updateSurface();
    const observer = new ResizeObserver(updateSurface);
    observer.observe(el);
    return () => {
      observer.disconnect();
    };
  }, [viewMode, isRealPdfMode, renderedPdfSize, currentSlideAspectRatio, stageViewportSize]);

  // Render active PDF page onto canvas
  useEffect(() => {
    if (!isRealPdfMode || !pdfDoc) {
      return;
    }
    const canvas = pdfCanvasRef.current;
    if (!canvas) {
      return;
    }

    let cancelled = false;
    let cancelRender: (() => void) | null = null;
    const pageNumber = Math.min(Math.max(1, safeSlideIdx + 1), pdfDoc.numPages);

    const run = async () => {
      try {
        const page = await pdfDoc.getPage(pageNumber);
        if (cancelled) {
          return;
        }
        const baseViewport = page.getViewport({ scale: 1 });
        const aspect = baseViewport.width / Math.max(1, baseViewport.height);
        const availW = Math.max(180, stageViewportSize.width - 12);
        const availH = Math.max(140, stageViewportSize.height - 12);
        const targetWidthCss = Math.max(
          160,
          Math.floor(Math.min(availW, availH * aspect))
        );

        const result = await renderPdfPageToCanvas({
          pdfDoc,
          pageNumber,
          canvas,
          targetWidthCss,
          maxDpr: 2,
        });
        cancelRender = result.cancel;
        if (!cancelled) {
          setRenderedPdfSize({
            width: Math.round(result.widthCss),
            height: Math.round(result.heightCss),
          });
        }
      } catch {
        if (!cancelled) {
          setPdfError('Không thể hiển thị trang PDF này.');
        }
      }
    };

    void run();

    return () => {
      cancelled = true;
      cancelRender?.();
    };
  }, [
    isRealPdfMode,
    pdfDoc,
    safeSlideIdx,
    stageViewportSize.width,
    stageViewportSize.height,
  ]);

  // Render next PDF page thumbnail for Presenter View
  useEffect(() => {
    if (
      viewMode !== 'presenter' ||
      !isRealPdfMode ||
      !pdfDoc ||
      safeSlideIdx + 1 >= pdfDoc.numPages
    ) {
      setNextPdfThumbUrl(null);
      return;
    }

    let cancelled = false;
    const nextPageNum = safeSlideIdx + 2;

    renderPdfThumbnailDataUrl(pdfDoc, nextPageNum, 360)
      .then((dataUrl) => {
        if (!cancelled) {
          setNextPdfThumbUrl(dataUrl);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setNextPdfThumbUrl(null);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [viewMode, isRealPdfMode, pdfDoc, safeSlideIdx]);

  // Presentation Timer loop
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

  // Fullscreen sync listener
  useEffect(() => {
    const handleFullscreenChange = () => {
      setIsFullscreen(Boolean(window.document.fullscreenElement));
    };
    window.document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => {
      window.document.removeEventListener(
        'fullscreenchange',
        handleFullscreenChange
      );
    };
  }, []);

  // Build snapshot of current RemoteChannelState
  const buildChannelState = useCallback((): RemoteChannelState => {
    return {
      roomCode: normalizedRoomCode,
      documentId: document.id,
      documentTitle: document.title,
      slideIndex: safeSlideIdx,
      totalSlides,
      slideTitle: currentSlideTitle,
      speakerNote: currentSpeakerNote,
      elapsedSeconds,
      isTimerRunning,
      annotationsCount: currentSlideStrokes.length,
    };
  }, [
    normalizedRoomCode,
    document.id,
    document.title,
    safeSlideIdx,
    totalSlides,
    currentSlideTitle,
    currentSpeakerNote,
    elapsedSeconds,
    isTimerRunning,
    currentSlideStrokes.length,
  ]);

  const buildChannelStateRef = useRef(buildChannelState);
  buildChannelStateRef.current = buildChannelState;

  const totalSlidesRef = useRef(totalSlides);
  totalSlidesRef.current = totalSlides;

  const safeSlideIdxRef = useRef(safeSlideIdx);
  safeSlideIdxRef.current = safeSlideIdx;

  const handleRemoteCommandRef = useRef<(cmd: RemoteCommandPayload) => void>(
    () => {}
  );
  handleRemoteCommandRef.current = (cmd: RemoteCommandPayload) => {
    switch (cmd.action) {
      case 'next':
        setSlideIdx((prev) => Math.min(totalSlidesRef.current - 1, prev + 1));
        break;
      case 'prev':
        setSlideIdx((prev) => Math.max(0, prev - 1));
        break;
      case 'goto':
        if (typeof cmd.slideIndex === 'number') {
          const clamped = Math.max(
            0,
            Math.min(totalSlidesRef.current - 1, Math.floor(cmd.slideIndex))
          );
          setSlideIdx(clamped);
        }
        break;
      case 'laserMove':
        if (cmd.point) {
          setRemoteLaser({
            active: true,
            x: cmd.point.x,
            y: cmd.point.y,
          });
          if (remoteLaserTimeoutRef.current !== null) {
            window.clearTimeout(remoteLaserTimeoutRef.current);
          }
          remoteLaserTimeoutRef.current = window.setTimeout(() => {
            setRemoteLaser((prev) => ({ ...prev, active: false }));
            remoteLaserTimeoutRef.current = null;
          }, 8000);
        }
        break;
      case 'laserOff':
        if (remoteLaserTimeoutRef.current !== null) {
          window.clearTimeout(remoteLaserTimeoutRef.current);
          remoteLaserTimeoutRef.current = null;
        }
        setRemoteLaser((prev) => ({ ...prev, active: false }));
        break;
      case 'spotlightMove':
        if (cmd.point) {
          setRemoteSpotlight({
            active: true,
            x: cmd.point.x,
            y: cmd.point.y,
            radius: DEFAULT_SPOTLIGHT_RADIUS,
          });
        }
        break;
      case 'spotlightOff':
        setRemoteSpotlight((prev) => ({ ...prev, active: false }));
        break;
      case 'clearAnnotations': {
        const targetIdx = safeSlideIdxRef.current;
        setStrokesBySlide((prev) => {
          if (!prev[targetIdx] || prev[targetIdx].length === 0) {
            return prev;
          }
          const next = { ...prev };
          delete next[targetIdx];
          return next;
        });
        remoteSessionRef.current?.broadcastAnnotations(targetIdx, []);
        break;
      }
      case 'toggleTimer':
        setIsTimerRunning((prev) => !prev);
        break;
      case 'resetTimer':
        setIsTimerRunning(false);
        setElapsedSeconds(0);
        break;
      case 'requestState':
        remoteSessionRef.current?.broadcastState(buildChannelStateRef.current());
        break;
    }
  };

  // Establish bidirectional Realtime / Local Broadcast remote session
  useEffect(() => {
    const session = createPresenterRemoteSession(normalizedRoomCode, {
      onCommand: (cmd) => {
        handleRemoteCommandRef.current(cmd);
      },
      onStatusChange: (status) => {
        setRemoteStatus(status);
      },
    });
    remoteSessionRef.current = session;
    session.broadcastState(buildChannelStateRef.current());

    return () => {
      if (remoteLaserTimeoutRef.current !== null) {
        window.clearTimeout(remoteLaserTimeoutRef.current);
        remoteLaserTimeoutRef.current = null;
      }
      if (copyTimerRef.current !== null) {
        window.clearTimeout(copyTimerRef.current);
        copyTimerRef.current = null;
      }
      session.destroy();
      remoteSessionRef.current = null;
    };
  }, [normalizedRoomCode]);

  // Sync state whenever key presentation properties change (not every 1s tick)
  useEffect(() => {
    remoteSessionRef.current?.broadcastState(buildChannelState());
  }, [
    document.id,
    document.title,
    safeSlideIdx,
    totalSlides,
    currentSlideTitle,
    currentSpeakerNote,
    isTimerRunning,
    currentSlideStrokes.length,
    buildChannelState,
  ]);

  // Periodic heartbeat every 5s while timer is running to keep remote clock aligned
  useEffect(() => {
    if (!isTimerRunning) {
      return;
    }
    const heartbeat = window.setInterval(() => {
      remoteSessionRef.current?.broadcastState(buildChannelStateRef.current());
    }, 5000);
    return () => {
      window.clearInterval(heartbeat);
    };
  }, [isTimerRunning]);

  const handleCloseStage = useCallback(() => {
    if (window.document.fullscreenElement && window.document.exitFullscreen) {
      try {
        void window.document.exitFullscreen().catch(() => {});
      } catch {
        // Ignore fullscreen exit errors
      }
    }
    onClose();
  }, [onClose]);

  // Keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        if (showInlineQr) {
          setShowInlineQr(false);
          return;
        }
        if (!window.document.fullscreenElement) {
          handleCloseStage();
        }
        return;
      }

      if (isInteractiveKeyboardTarget(event.target)) {
        return;
      }

      if (
        event.key === 'ArrowRight' ||
        event.key === 'PageDown' ||
        event.key === ' '
      ) {
        event.preventDefault();
        setSlideIdx((prev) => Math.min(totalSlides - 1, prev + 1));
      } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault();
        setSlideIdx((prev) => Math.max(0, prev - 1));
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [handleCloseStage, showInlineQr, totalSlides]);

  const toggleBrowserFullscreen = () => {
    try {
      if (window.document.fullscreenElement) {
        void window.document.exitFullscreen?.().catch(() => {});
      } else {
        void window.document.documentElement.requestFullscreen?.().catch(() => {});
      }
    } catch {
      // Ignore environments where Fullscreen API is restricted
    }
  };

  const handleSelectTool = (tool: StageToolMode) => {
    setActiveTool(tool);
    if (tool !== 'laser') {
      setLocalLaser((prev) => ({ ...prev, active: false }));
    }
    if (tool === 'spotlight') {
      setLocalSpotlight((prev) => ({ ...prev, active: true }));
    } else {
      setLocalSpotlight((prev) => ({ ...prev, active: false }));
    }
  };

  const handleUndoStroke = () => {
    setStrokesBySlide((prev) => {
      const existing = prev[safeSlideIdx];
      if (!existing || existing.length === 0) {
        return prev;
      }
      const nextStrokes = existing.slice(0, -1);
      remoteSessionRef.current?.broadcastAnnotations(safeSlideIdx, nextStrokes);
      return {
        ...prev,
        [safeSlideIdx]: nextStrokes,
      };
    });
  };

  const handleClearSlideStrokes = () => {
    setStrokesBySlide((prev) => {
      if (!prev[safeSlideIdx] || prev[safeSlideIdx].length === 0) {
        return prev;
      }
      const next = { ...prev };
      delete next[safeSlideIdx];
      return next;
    });
    remoteSessionRef.current?.broadcastAnnotations(safeSlideIdx, []);
  };

  const getNormalizedPointFromEvent = (
    event: ReactPointerEvent<HTMLDivElement>
  ): NormalizedPoint | null => {
    const rect = interactiveSurfaceRef.current?.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      return null;
    }
    return {
      x: clampUnit((event.clientX - rect.left) / rect.width),
      y: clampUnit((event.clientY - rect.top) / rect.height),
    };
  };

  const handleStagePointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (activeTool === 'pointer') {
      return;
    }
    const point = getNormalizedPointFromEvent(event);
    if (!point) {
      return;
    }

    if (activeTool === 'laser') {
      setLocalLaser({ active: true, x: point.x, y: point.y });
      return;
    }

    if (activeTool === 'spotlight') {
      setLocalSpotlight((prev) => ({
        ...prev,
        active: true,
        x: point.x,
        y: point.y,
      }));
      return;
    }

    if (activeTool === 'pen' || activeTool === 'highlighter') {
      drawingPointerIdRef.current = event.pointerId;
      try {
        event.currentTarget.setPointerCapture(event.pointerId);
      } catch {
        // Ignore capture failure
      }
      const stroke: AnnotationStroke = {
        id: `stroke-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`,
        slideIndex: safeSlideIdx,
        color: penColor,
        width:
          activeTool === 'highlighter' ? HIGHLIGHTER_WIDTH_PX : PEN_WIDTH_PX,
        isHighlighter: activeTool === 'highlighter',
        points: [point, point],
      };
      activeStrokeRef.current = stroke;
      setActiveStroke(stroke);
    }
  };

  const handleStagePointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (activeTool === 'pointer') {
      return;
    }
    const point = getNormalizedPointFromEvent(event);
    if (!point) {
      return;
    }

    if (activeTool === 'laser') {
      setLocalLaser({ active: true, x: point.x, y: point.y });
      return;
    }

    if (activeTool === 'spotlight') {
      setLocalSpotlight((prev) => ({
        ...prev,
        active: true,
        x: point.x,
        y: point.y,
      }));
      return;
    }

    if (
      (activeTool === 'pen' || activeTool === 'highlighter') &&
      drawingPointerIdRef.current === event.pointerId &&
      activeStrokeRef.current
    ) {
      const current = activeStrokeRef.current;
      const lastPoint = current.points[current.points.length - 1];
      const dx = point.x - lastPoint.x;
      const dy = point.y - lastPoint.y;
      if (dx * dx + dy * dy < 0.000004) {
        return;
      }
      const updated: AnnotationStroke = {
        ...current,
        points: [...current.points, point],
      };
      activeStrokeRef.current = updated;
      setActiveStroke(updated);
    }
  };

  const commitActiveStroke = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (drawingPointerIdRef.current !== event.pointerId) {
      return;
    }
    drawingPointerIdRef.current = null;
    try {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) {
        event.currentTarget.releasePointerCapture(event.pointerId);
      }
    } catch {
      // Ignore release errors
    }

    const finished = activeStrokeRef.current;
    activeStrokeRef.current = null;
    setActiveStroke(null);

    if (!finished) {
      return;
    }

    setStrokesBySlide((prev) => {
      const existing = prev[finished.slideIndex] ?? [];
      const nextStrokes = [...existing, finished];
      remoteSessionRef.current?.broadcastAnnotations(
        finished.slideIndex,
        nextStrokes
      );
      return {
        ...prev,
        [finished.slideIndex]: nextStrokes,
      };
    });
  };

  const handleStagePointerLeave = () => {
    if (activeTool === 'laser') {
      setLocalLaser((prev) => ({ ...prev, active: false }));
    }
  };

  const handleQrButtonClick = () => {
    if (onOpenQrModal) {
      onOpenQrModal();
      return;
    }
    setShowInlineQr((prev) => !prev);
  };

  const handleCopyRemoteUrl = () => {
    if (navigator.clipboard?.writeText) {
      void navigator.clipboard.writeText(remoteUrl).catch(() => {});
    }
    setCopiedRemoteUrl(true);
    if (copyTimerRef.current !== null) {
      window.clearTimeout(copyTimerRef.current);
    }
    copyTimerRef.current = window.setTimeout(() => {
      setCopiedRemoteUrl(false);
      copyTimerRef.current = null;
    }, 1800);
  };

  const handleResetTimer = () => {
    setIsTimerRunning(false);
    setElapsedSeconds(0);
  };

  const displayedStrokes =
    activeStroke && activeStroke.slideIndex === safeSlideIdx
      ? [...currentSlideStrokes, activeStroke]
      : currentSlideStrokes;

  const effectiveLaser = remoteLaser.active ? remoteLaser : localLaser;
  const effectiveSpotlight = remoteSpotlight.active
    ? remoteSpotlight
    : localSpotlight;

  const spotlightRadiusPx = Math.max(
    54,
    Math.round(
      Math.min(activeSurfaceSize.width, activeSurfaceSize.height) *
        effectiveSpotlight.radius
    )
  );

  const stageCursorClass =
    activeTool === 'pen' || activeTool === 'highlighter'
      ? 'cursor-crosshair'
      : activeTool === 'laser' || activeTool === 'spotlight'
        ? 'cursor-none'
        : 'cursor-default';

  const hasColorPicker = activeTool === 'pen' || activeTool === 'highlighter';
  const transportLabel =
    remoteStatus.mode === 'supabase-realtime'
      ? 'Supabase Realtime'
      : 'Kết nối cục bộ';

  return (
    <div className="w-full h-full max-h-full rounded-2xl sm:rounded-3xl bg-[#0F141F] border border-white/15 shadow-2xl overflow-hidden text-slate-100 flex flex-col relative">
      {/* Top Stage Header */}
      <header className="shrink-0 px-3 sm:px-5 py-2.5 sm:py-3 border-b border-white/10 flex items-center justify-between gap-2 bg-[#151B28]">
        <div className="flex items-center gap-2.5 min-w-0">
          <span className="h-8 w-8 sm:h-9 sm:w-9 rounded-xl bg-indigo-500/20 border border-indigo-400/30 flex items-center justify-center text-indigo-300 shrink-0">
            <Presentation className="h-4 w-4 sm:h-[18px] sm:w-[18px]" />
          </span>
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <h3 className="text-xs sm:text-sm font-semibold text-white truncate">
                {document.title}
              </h3>
              <span className="shrink-0 px-2 py-0.5 rounded-md bg-white/10 text-[10px] font-semibold uppercase tracking-wide text-slate-200 tabular-nums">
                {isRealPdfMode ? 'Trang' : 'Slide'} {safeSlideIdx + 1} /{' '}
                {totalSlides}
              </span>
            </div>
            <p className="text-[11px] text-slate-400 truncate hidden sm:block">
              Phím ← / → hoặc Space để chuyển · Esc để thoát · Ký hiệu &amp; Laser
              đồng bộ tỷ lệ 0..1
            </p>
          </div>
        </div>

        <div className="shrink-0 flex items-center gap-1.5 sm:gap-2">
          {/* Compact Clock Pill in Audience Mode */}
          {viewMode === 'audience' && (
            <button
              type="button"
              onClick={() => setIsTimerRunning((prev) => !prev)}
              title={
                isTimerRunning
                  ? 'Tạm dừng đồng hồ thuyết trình'
                  : 'Bấm giờ thuyết trình'
              }
              className={`hidden md:inline-flex items-center gap-1.5 h-9 px-2.5 rounded-xl border text-xs font-mono font-semibold tabular-nums transition-colors ${
                isTimerRunning
                  ? 'bg-emerald-500/15 border-emerald-400/35 text-emerald-200'
                  : 'bg-white/[0.06] hover:bg-white/[0.12] border-white/10 text-slate-200'
              }`}
            >
              {isTimerRunning ? (
                <Pause className="h-3.5 w-3.5 text-emerald-300" />
              ) : (
                <Play className="h-3.5 w-3.5 text-indigo-300" />
              )}
              <span>{formatTimer(elapsedSeconds)}</span>
            </button>
          )}

          {/* Audience Stage <-> Presenter View Mode Switcher */}
          <button
            type="button"
            onClick={() =>
              setViewMode((prev) =>
                prev === 'audience' ? 'presenter' : 'audience'
              )
            }
            aria-pressed={viewMode === 'presenter'}
            className={`inline-flex items-center gap-1.5 h-9 px-2.5 sm:px-3 rounded-xl border text-xs font-medium transition-colors whitespace-nowrap ${
              viewMode === 'presenter'
                ? 'bg-indigo-600 border-indigo-500 text-white shadow-sm'
                : 'bg-white/[0.06] hover:bg-white/[0.14] border-white/10 text-slate-200'
            }`}
          >
            {viewMode === 'presenter' ? (
              <>
                <Monitor className="h-3.5 w-3.5 shrink-0" />
                <span className="hidden sm:inline">Sân khấu Khán giả</span>
              </>
            ) : (
              <>
                <Columns2 className="h-3.5 w-3.5 shrink-0 text-indigo-300" />
                <span className="hidden sm:inline">Presenter View</span>
              </>
            )}
          </button>

          {/* QR Remote Trigger */}
          <button
            type="button"
            onClick={handleQrButtonClick}
            className="inline-flex items-center gap-1.5 h-9 px-2.5 sm:px-3 rounded-xl bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 text-xs font-medium text-slate-200 transition-colors whitespace-nowrap"
          >
            <QrCode className="h-3.5 w-3.5 text-indigo-300 shrink-0" />
            <span className="hidden md:inline">QR Remote</span>
            {remoteStatus.peerCount > 0 && (
              <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 text-[10px] font-bold tabular-nums">
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-400 animate-pulse" />
                {remoteStatus.peerCount}
              </span>
            )}
          </button>

          {/* Browser Fullscreen Button */}
          <button
            type="button"
            onClick={toggleBrowserFullscreen}
            aria-label={
              isFullscreen ? 'Thoát toàn màn hình' : 'Toàn màn hình trình duyệt'
            }
            title={
              isFullscreen ? 'Thoát toàn màn hình' : 'Toàn màn hình trình duyệt'
            }
            className="hidden sm:inline-flex items-center justify-center h-9 w-9 rounded-xl bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 text-slate-200 transition-colors"
          >
            {isFullscreen ? (
              <Minimize2 className="h-4 w-4" />
            ) : (
              <Maximize2 className="h-4 w-4" />
            )}
          </button>

          {/* Close Presentation */}
          <button
            type="button"
            onClick={handleCloseStage}
            aria-label="Thoát chế độ trình chiếu"
            className="shrink-0 min-h-[36px] min-w-[36px] rounded-xl bg-white/5 hover:bg-white/15 border border-white/10 flex items-center justify-center text-slate-300 hover:text-white transition-colors"
          >
            <X className="h-4 w-4" />
          </button>
        </div>
      </header>

      {/* Interactive Annotation & Pointer Toolbar */}
      <div className="shrink-0 px-3 sm:px-5 py-2 border-b border-white/10 bg-[#121825] flex items-center justify-between gap-2 overflow-x-auto no-scrollbar">
        <div className="flex items-center gap-1.5 shrink-0">
          {(
            [
              { id: 'pointer', label: 'Con trỏ', icon: MousePointer2 },
              { id: 'pen', label: 'Bút vẽ', icon: Pencil },
              { id: 'highlighter', label: 'Dạ quang', icon: Highlighter },
              { id: 'laser', label: 'Tia Laser', icon: Crosshair },
              { id: 'spotlight', label: 'Spotlight', icon: Flashlight },
            ] as const
          ).map((tool) => {
            const ToolIcon = tool.icon;
            const isSelected = activeTool === tool.id;
            return (
              <button
                key={tool.id}
                type="button"
                onClick={() => handleSelectTool(tool.id)}
                aria-pressed={isSelected}
                className={`inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg border text-xs font-medium whitespace-nowrap transition-all ${
                  isSelected
                    ? 'bg-indigo-600 border-indigo-500 text-white shadow-sm'
                    : 'bg-white/[0.04] hover:bg-white/[0.1] border-white/10 text-slate-300'
                }`}
              >
                <ToolIcon className="h-3.5 w-3.5 shrink-0" />
                <span>{tool.label}</span>
              </button>
            );
          })}

          {/* 4 Color Swatches */}
          <div
            className={`flex items-center gap-1 pl-2 ml-1 border-l border-white/10 transition-opacity ${
              hasColorPicker ? 'opacity-100' : 'opacity-65'
            }`}
          >
            {PEN_COLORS.map((swatch) => {
              const selected = penColor === swatch.hex;
              return (
                <button
                  key={swatch.hex}
                  type="button"
                  onClick={() => {
                    setPenColor(swatch.hex);
                    if (activeTool !== 'pen' && activeTool !== 'highlighter') {
                      handleSelectTool('pen');
                    }
                  }}
                  title={`Màu ${swatch.label}`}
                  aria-label={`Chọn màu ${swatch.label}`}
                  className={`h-6 w-6 rounded-full flex items-center justify-center transition-transform ${
                    selected
                      ? 'ring-2 ring-white ring-offset-2 ring-offset-[#121825] scale-105'
                      : 'hover:scale-105 opacity-85 hover:opacity-100'
                  }`}
                  style={{ backgroundColor: swatch.hex }}
                />
              );
            })}
          </div>
        </div>

        <div className="flex items-center gap-1.5 shrink-0">
          {remoteSpotlight.active && (
            <button
              type="button"
              onClick={() =>
                setRemoteSpotlight((prev) => ({ ...prev, active: false }))
              }
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-amber-500/20 border border-amber-400/40 text-amber-200 text-xs font-medium whitespace-nowrap"
            >
              <Flashlight className="h-3.5 w-3.5" />
              <span>Tắt Spotlight từ xa</span>
            </button>
          )}

          <button
            type="button"
            onClick={handleUndoStroke}
            disabled={currentSlideStrokes.length === 0}
            className="inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.04] hover:bg-white/[0.1] border border-white/10 text-xs font-medium text-slate-200 whitespace-nowrap transition-colors disabled:opacity-35 disabled:cursor-not-allowed"
          >
            <Undo2 className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">Hoàn tác</span>
          </button>

          <button
            type="button"
            onClick={handleClearSlideStrokes}
            disabled={currentSlideStrokes.length === 0}
            className="inline-flex items-center gap-1.5 h-8 px-2.5 rounded-lg bg-white/[0.04] hover:bg-rose-500/20 border border-white/10 hover:border-rose-400/35 text-xs font-medium text-slate-200 hover:text-rose-200 whitespace-nowrap transition-colors disabled:opacity-35 disabled:cursor-not-allowed"
          >
            <Eraser className="h-3.5 w-3.5" />
            <span>Xóa nét slide này</span>
            {currentSlideStrokes.length > 0 && (
              <span className="px-1.5 py-0.5 rounded-full bg-white/10 text-[10px] tabular-nums">
                {currentSlideStrokes.length}
              </span>
            )}
          </button>
        </div>
      </div>

      {/* Main Body: Audience Stage OR 2-Column Presenter View */}
      <div
        className={`flex-1 min-h-0 overflow-hidden ${
          viewMode === 'presenter'
            ? 'grid grid-cols-1 lg:grid-cols-12'
            : 'flex flex-col'
        }`}
      >
        {/* Primary Stage Column */}
        <div
          className={`min-h-0 flex flex-col ${
            viewMode === 'presenter' ? 'lg:col-span-8 flex-1' : 'flex-1'
          }`}
        >
          <div
            ref={stageViewportRef}
            className="flex-1 min-h-0 p-3 sm:p-5 flex items-center justify-center overflow-hidden bg-[#0B0F18]"
          >
            {isRealPdfMode ? (
              pdfLoading && !pdfDoc ? (
                <div className="flex flex-col items-center gap-2.5 text-slate-300">
                  <Loader2 className="h-7 w-7 animate-spin text-indigo-400" />
                  <p className="text-xs font-medium">
                    Đang nạp tài liệu PDF để trình chiếu…
                  </p>
                </div>
              ) : pdfError ? (
                <div className="max-w-md rounded-2xl border border-rose-400/30 bg-rose-500/10 p-5 text-center space-y-2">
                  <FileText className="h-7 w-7 text-rose-300 mx-auto" />
                  <p className="text-sm font-semibold text-rose-100">
                    Không thể trình chiếu tệp PDF
                  </p>
                  <p className="text-xs text-rose-200/80">{pdfError}</p>
                </div>
              ) : (
                <div
                  ref={interactiveSurfaceRef}
                  onPointerDown={handleStagePointerDown}
                  onPointerMove={handleStagePointerMove}
                  onPointerUp={commitActiveStroke}
                  onPointerCancel={commitActiveStroke}
                  onPointerLeave={handleStagePointerLeave}
                  style={
                    renderedPdfSize
                      ? {
                          width: `${renderedPdfSize.width}px`,
                          height: `${renderedPdfSize.height}px`,
                        }
                      : undefined
                  }
                  className={`relative rounded-xl overflow-hidden shadow-2xl border border-white/15 bg-white select-none ${
                    activeTool !== 'pointer' ? 'touch-none' : ''
                  } ${stageCursorClass}`}
                >
                  <canvas ref={pdfCanvasRef} className="block" />

                  {/* Normalized 0..1 SVG Annotations */}
                  <svg
                    viewBox="0 0 1000 1000"
                    preserveAspectRatio="none"
                    className="absolute inset-0 w-full h-full pointer-events-none z-20"
                  >
                    {displayedStrokes.map((stroke) => (
                      <polyline
                        key={stroke.id}
                        fill="none"
                        stroke={stroke.color}
                        strokeWidth={stroke.width}
                        strokeOpacity={stroke.isHighlighter ? 0.38 : 0.96}
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        vectorEffect="non-scaling-stroke"
                        points={stroke.points
                          .map(
                            (p) =>
                              `${(p.x * 1000).toFixed(1)},${(p.y * 1000).toFixed(1)}`
                          )
                          .join(' ')}
                      />
                    ))}
                  </svg>

                  {/* Circular Spotlight Layer */}
                  {effectiveSpotlight.active && (
                    <div
                      className="absolute inset-0 pointer-events-none z-20 transition-[background] duration-75"
                      style={{
                        background: `radial-gradient(circle ${spotlightRadiusPx}px at ${(
                          effectiveSpotlight.x * 100
                        ).toFixed(2)}% ${(effectiveSpotlight.y * 100).toFixed(
                          2
                        )}%, rgba(255,255,255,0.03) 0%, rgba(255,255,255,0) 64%, rgba(5,8,16,0.84) 100%)`,
                      }}
                    >
                      <div
                        className="absolute -translate-x-1/2 -translate-y-1/2 rounded-full border border-amber-200/50 shadow-[0_0_24px_rgba(251,191,36,0.28)]"
                        style={{
                          left: `${effectiveSpotlight.x * 100}%`,
                          top: `${effectiveSpotlight.y * 100}%`,
                          width: `${spotlightRadiusPx * 2}px`,
                          height: `${spotlightRadiusPx * 2}px`,
                        }}
                      />
                    </div>
                  )}

                  {/* Laser Pointer Layer */}
                  {effectiveLaser.active && (
                    <div
                      className="absolute -translate-x-1/2 -translate-y-1/2 pointer-events-none z-30"
                      style={{
                        left: `${effectiveLaser.x * 100}%`,
                        top: `${effectiveLaser.y * 100}%`,
                      }}
                    >
                      <span className="block h-4 w-4 rounded-full bg-rose-500 ring-4 ring-rose-500/35 shadow-[0_0_20px_rgba(244,63,94,1)]" />
                      <span className="absolute inset-1 rounded-full bg-white/90" />
                    </div>
                  )}
                </div>
              )
            ) : currentSlide ? (
              <div
                ref={interactiveSurfaceRef}
                onPointerDown={handleStagePointerDown}
                onPointerMove={handleStagePointerMove}
                onPointerUp={commitActiveStroke}
                onPointerCancel={commitActiveStroke}
                onPointerLeave={handleStagePointerLeave}
                style={{
                  width: `${fittedSlideBox.width}px`,
                  height: `${fittedSlideBox.height}px`,
                }}
                className={`relative max-w-full max-h-full rounded-2xl shadow-2xl border border-white/15 overflow-hidden ${
                  activeTool !== 'pointer' ? 'touch-none' : ''
                } ${stageCursorClass}`}
              >
                <SlideCardCanvas
                  doc={document}
                  slide={currentSlide}
                  slideIndex={safeSlideIdx}
                  totalSlides={totalSlides}
                />

                {/* Normalized 0..1 SVG Annotations */}
                <svg
                  viewBox="0 0 1000 1000"
                  preserveAspectRatio="none"
                  className="absolute inset-0 w-full h-full pointer-events-none z-20"
                >
                  {displayedStrokes.map((stroke) => (
                    <polyline
                      key={stroke.id}
                      fill="none"
                      stroke={stroke.color}
                      strokeWidth={stroke.width}
                      strokeOpacity={stroke.isHighlighter ? 0.38 : 0.96}
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      vectorEffect="non-scaling-stroke"
                      points={stroke.points
                        .map(
                          (p) =>
                            `${(p.x * 1000).toFixed(1)},${(p.y * 1000).toFixed(1)}`
                        )
                        .join(' ')}
                    />
                  ))}
                </svg>

                {/* Circular Spotlight Layer */}
                {effectiveSpotlight.active && (
                  <div
                    className="absolute inset-0 pointer-events-none z-20"
                    style={{
                      background: `radial-gradient(circle ${spotlightRadiusPx}px at ${(
                        effectiveSpotlight.x * 100
                      ).toFixed(2)}% ${(effectiveSpotlight.y * 100).toFixed(
                        2
                      )}%, rgba(255,255,255,0.03) 0%, rgba(255,255,255,0) 64%, rgba(5,8,16,0.84) 100%)`,
                    }}
                  >
                    <div
                      className="absolute -translate-x-1/2 -translate-y-1/2 rounded-full border border-amber-200/50 shadow-[0_0_24px_rgba(251,191,36,0.28)]"
                      style={{
                        left: `${effectiveSpotlight.x * 100}%`,
                        top: `${effectiveSpotlight.y * 100}%`,
                        width: `${spotlightRadiusPx * 2}px`,
                        height: `${spotlightRadiusPx * 2}px`,
                      }}
                    />
                  </div>
                )}

                {/* Laser Pointer Layer */}
                {effectiveLaser.active && (
                  <div
                    className="absolute -translate-x-1/2 -translate-y-1/2 pointer-events-none z-30"
                    style={{
                      left: `${effectiveLaser.x * 100}%`,
                      top: `${effectiveLaser.y * 100}%`,
                    }}
                  >
                    <span className="block h-4 w-4 rounded-full bg-rose-500 ring-4 ring-rose-500/35 shadow-[0_0_20px_rgba(244,63,94,1)]" />
                    <span className="absolute inset-1 rounded-full bg-white/90" />
                  </div>
                )}
              </div>
            ) : (
              <div className="aspect-video w-full max-w-4xl rounded-2xl border border-white/10 bg-[#1B2231] flex items-center justify-center text-sm text-slate-400">
                Chưa có dữ liệu trang trình chiếu cho tài liệu này.
              </div>
            )}
          </div>

          {/* Bottom Navigation Bar */}
          <div className="shrink-0 px-3 sm:px-5 py-2.5 border-t border-white/10 bg-[#151B28] flex flex-wrap items-center justify-between gap-2">
            {/* Quick Slide Dots / Status */}
            <div className="flex items-center gap-2 min-w-0">
              {totalSlides <= 16 ? (
                <div className="hidden sm:flex items-center gap-1.5">
                  {Array.from({ length: totalSlides }).map((_, idx) => (
                    <button
                      key={`dot-${idx}`}
                      type="button"
                      onClick={() => setSlideIdx(idx)}
                      aria-label={`Đi tới trang ${idx + 1}`}
                      className={`h-2 rounded-full transition-all ${
                        idx === safeSlideIdx
                          ? 'w-6 bg-indigo-400'
                          : 'w-2 bg-white/25 hover:bg-white/50'
                      }`}
                    />
                  ))}
                </div>
              ) : (
                <span className="text-xs text-slate-300 tabular-nums">
                  {isRealPdfMode ? 'Trang' : 'Slide'} {safeSlideIdx + 1} /{' '}
                  {totalSlides}
                </span>
              )}

              {viewMode === 'audience' && currentSpeakerNote && (
                <button
                  type="button"
                  onClick={() => setViewMode('presenter')}
                  className="inline-flex items-center gap-1.5 h-9 px-3 rounded-xl bg-white/[0.06] hover:bg-white/[0.12] border border-white/10 text-xs font-medium text-indigo-200 transition-colors"
                >
                  <MessageSquareText className="h-3.5 w-3.5 shrink-0" />
                  <span>Xem ghi chú &amp; slide kế</span>
                </button>
              )}
            </div>

            <div className="flex items-center gap-2 ml-auto">
              <button
                type="button"
                onClick={() =>
                  setSlideIdx((prev) => (prev > 0 ? prev - 1 : prev))
                }
                disabled={safeSlideIdx === 0}
                className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-3.5 rounded-xl bg-white/10 hover:bg-white/15 border border-white/10 text-xs font-semibold text-white whitespace-nowrap transition disabled:opacity-35 disabled:cursor-not-allowed"
              >
                <ChevronLeft className="h-4 w-4" />
                <span>{isRealPdfMode ? 'Trang trước' : 'Slide trước'}</span>
              </button>
              <button
                type="button"
                onClick={() =>
                  setSlideIdx((prev) =>
                    prev < totalSlides - 1 ? prev + 1 : prev
                  )
                }
                disabled={safeSlideIdx >= totalSlides - 1}
                className="inline-flex items-center justify-center gap-1.5 min-h-[40px] px-4 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white whitespace-nowrap transition disabled:opacity-35 disabled:cursor-not-allowed"
              >
                <span>{isRealPdfMode ? 'Trang tiếp' : 'Slide tiếp'}</span>
                <ChevronRight className="h-4 w-4" />
              </button>
            </div>
          </div>
        </div>

        {/* Presenter View Right Panel (Next Slide + Timer + Speaker Notes + QR Remote Status) */}
        {viewMode === 'presenter' && (
          <aside className="lg:col-span-4 border-t lg:border-t-0 lg:border-l border-white/10 bg-[#141A27] p-3.5 sm:p-4 overflow-y-auto scrollbar-dark space-y-3.5 max-h-[44vh] lg:max-h-none">
            {/* Next Slide / Next Page Preview */}
            <div className="rounded-2xl bg-[#1A2233] border border-white/10 p-3 space-y-2">
              <div className="flex items-center justify-between text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                <span>
                  {isRealPdfMode ? 'Trang kế tiếp' : 'Slide kế tiếp'}
                </span>
                <span className="tabular-nums text-slate-300">
                  {safeSlideIdx + 1 < totalSlides
                    ? `${safeSlideIdx + 2} / ${totalSlides}`
                    : 'Cuối tài liệu'}
                </span>
              </div>

              {safeSlideIdx + 1 >= totalSlides ? (
                <div className="aspect-video rounded-xl border border-dashed border-white/15 bg-[#101623] flex items-center justify-center text-xs text-slate-400">
                  Đã đến {isRealPdfMode ? 'trang' : 'slide'} cuối cùng
                </div>
              ) : isRealPdfMode ? (
                <button
                  type="button"
                  onClick={() =>
                    setSlideIdx((prev) => Math.min(totalSlides - 1, prev + 1))
                  }
                  className="w-full aspect-video rounded-xl border border-white/15 bg-[#0D121D] overflow-hidden flex items-center justify-center hover:border-indigo-400/50 transition-colors"
                >
                  {nextPdfThumbUrl ? (
                    <img
                      src={nextPdfThumbUrl}
                      alt={`Xem trước trang ${safeSlideIdx + 2}`}
                      className="max-h-full max-w-full object-contain"
                    />
                  ) : (
                    <span className="text-xs text-slate-400">
                      Trang {safeSlideIdx + 2} (Nhấn để chuyển)
                    </span>
                  )}
                </button>
              ) : nextSlide ? (
                <button
                  type="button"
                  onClick={() =>
                    setSlideIdx((prev) => Math.min(totalSlides - 1, prev + 1))
                  }
                  style={{ aspectRatio: String(nextSlideAspectRatio) }}
                  className="w-full text-left rounded-xl overflow-hidden border border-white/15 hover:border-indigo-400/50 transition-colors"
                >
                  <SlideCardCanvas
                    doc={document}
                    slide={nextSlide}
                    slideIndex={safeSlideIdx + 1}
                    totalSlides={totalSlides}
                    compact
                  />
                </button>
              ) : null}
            </div>

            {/* Presentation Timer Card */}
            <div className="rounded-2xl bg-[#1A2233] border border-white/10 p-3.5 space-y-2.5">
              <div className="flex items-center justify-between">
                <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                  <Clock className="h-3.5 w-3.5 text-indigo-300" />
                  <span>Đồng hồ thuyết trình</span>
                </span>
                <span
                  className={`text-[11px] font-medium ${
                    isTimerRunning ? 'text-emerald-300' : 'text-slate-400'
                  }`}
                >
                  {isTimerRunning ? 'Đang chạy' : 'Tạm dừng'}
                </span>
              </div>

              <div className="flex items-center justify-between gap-2">
                <span className="font-mono text-2xl font-bold tracking-wider text-white tabular-nums">
                  {formatTimer(elapsedSeconds)}
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setIsTimerRunning((prev) => !prev)}
                    className={`inline-flex items-center gap-1.5 h-9 px-3 rounded-xl text-xs font-semibold transition-colors ${
                      isTimerRunning
                        ? 'bg-amber-500/20 hover:bg-amber-500/30 border border-amber-400/40 text-amber-200'
                        : 'bg-indigo-600 hover:bg-indigo-500 text-white'
                    }`}
                  >
                    {isTimerRunning ? (
                      <>
                        <Pause className="h-3.5 w-3.5" />
                        <span>Tạm dừng</span>
                      </>
                    ) : (
                      <>
                        <Play className="h-3.5 w-3.5" />
                        <span>Bắt đầu</span>
                      </>
                    )}
                  </button>
                  <button
                    type="button"
                    onClick={handleResetTimer}
                    aria-label="Đặt lại đồng hồ"
                    title="Đặt lại đồng hồ"
                    className="h-9 w-9 rounded-xl bg-white/[0.06] hover:bg-white/[0.14] border border-white/10 flex items-center justify-center text-slate-300 hover:text-white transition-colors"
                  >
                    <RotateCcw className="h-3.5 w-3.5" />
                  </button>
                </div>
              </div>
            </div>

            {/* Speaker Notes Card */}
            <div className="rounded-2xl bg-[#1A2233] border border-white/10 p-3.5 space-y-2">
              <div className="flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wider text-indigo-300">
                <MessageSquareText className="h-3.5 w-3.5" />
                <span>Ghi chú diễn giả</span>
              </div>
              <p className="text-xs sm:text-[13px] text-slate-200 leading-relaxed whitespace-pre-wrap">
                {currentSpeakerNote.trim()
                  ? currentSpeakerNote
                  : 'Slide này chưa có ghi chú diễn giả. Bạn có thể thêm ghi chú trong trình soạn thảo Slide.'}
              </p>
            </div>

            {/* QR Remote Connection Card */}
            <div className="rounded-2xl bg-[#1A2233] border border-white/10 p-3.5 space-y-2.5">
              <div className="flex items-center justify-between gap-2">
                <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                  <Smartphone className="h-3.5 w-3.5 text-indigo-300" />
                  <span>Điều khiển QR Remote</span>
                </span>
                <span className="font-mono text-xs font-bold tracking-wider text-indigo-200 bg-indigo-500/15 border border-indigo-400/30 px-2 py-0.5 rounded-md">
                  {normalizedRoomCode}
                </span>
              </div>

              <div className="flex items-center justify-between gap-2 text-xs">
                <div className="inline-flex items-center gap-1.5 text-slate-300">
                  <Radio
                    className={`h-3.5 w-3.5 ${
                      remoteStatus.peerCount > 0
                        ? 'text-emerald-400 animate-pulse'
                        : remoteStatus.connected
                          ? 'text-indigo-300'
                          : 'text-slate-500'
                    }`}
                  />
                  <span>
                    {remoteStatus.peerCount > 0
                      ? `${remoteStatus.peerCount} thiết bị đang điều khiển`
                      : 'Sẵn sàng ghép đôi thiết bị'}
                  </span>
                </div>
                <span className="text-[11px] text-slate-400">
                  {transportLabel}
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleQrButtonClick}
                  className="flex-1 inline-flex items-center justify-center gap-1.5 h-9 px-3 rounded-xl bg-white/10 hover:bg-white/15 border border-white/10 text-xs font-medium text-white transition-colors"
                >
                  <QrCode className="h-3.5 w-3.5 text-indigo-300" />
                  <span>Mở mã QR</span>
                </button>
                <a
                  href={remoteUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center justify-center gap-1 h-9 px-3 rounded-xl bg-indigo-600/80 hover:bg-indigo-500 text-xs font-medium text-white transition-colors"
                >
                  <ExternalLink className="h-3.5 w-3.5" />
                  <span>Mở Tab Remote</span>
                </a>
              </div>
            </div>
          </aside>
        )}
      </div>

      {/* Built-in Quick QR Popover when onOpenQrModal is not supplied */}
      {showInlineQr && (
        <div
          role="dialog"
          aria-label="Mã QR điều khiển từ xa"
          onClick={(e) => {
            if (e.target === e.currentTarget) {
              setShowInlineQr(false);
            }
          }}
          className="absolute inset-0 z-40 bg-slate-950/75 backdrop-blur-sm flex items-center justify-center p-4"
        >
          <div className="w-full max-w-sm rounded-2xl bg-[#1A202D] border border-white/15 p-5 text-center space-y-4 shadow-2xl">
            <div className="flex items-center justify-between">
              <div className="text-left">
                <h4 className="text-sm font-semibold text-white">
                  Quét QR để điều khiển từ xa
                </h4>
                <p className="text-[11px] text-slate-400">
                  Mã phòng: <strong className="text-indigo-300">{normalizedRoomCode}</strong>
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowInlineQr(false)}
                aria-label="Đóng mã QR"
                className="h-8 w-8 rounded-lg bg-white/5 hover:bg-white/15 flex items-center justify-center text-slate-300"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="inline-flex p-3 rounded-2xl bg-white shadow-xl">
              <RealQrCode
                value={remoteUrl}
                sizeClassName="h-44 w-44"
                ariaLabel={`Mã QR ghép đôi phòng ${normalizedRoomCode}`}
              />
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleCopyRemoteUrl}
                className="flex-1 inline-flex items-center justify-center gap-1.5 h-9 px-3 rounded-xl bg-white/10 hover:bg-white/15 text-xs font-medium text-white transition-colors"
              >
                {copiedRemoteUrl ? (
                  <>
                    <Check className="h-3.5 w-3.5 text-emerald-400" />
                    <span>Đã sao chép link</span>
                  </>
                ) : (
                  <>
                    <Copy className="h-3.5 w-3.5" />
                    <span>Chép link Remote</span>
                  </>
                )}
              </button>
              <a
                href={remoteUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center justify-center gap-1.5 h-9 px-3 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white transition-colors"
              >
                <ExternalLink className="h-3.5 w-3.5" />
                <span>Mở</span>
              </a>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

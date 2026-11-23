import { useEffect, useRef, useState } from 'react';
import {
  LayoutGrid,
  Maximize2,
  Minus,
  MonitorPlay,
  Plus,
  Presentation,
} from 'lucide-react';

export type OfficeStatusBarVariant = 'word' | 'excel' | 'powerpoint';

export interface OfficeStatusBarProps {
  variant: OfficeStatusBarVariant;
  theme?: 'light' | 'dark';
  zoomPercent: number;
  onZoomChange: (pct: number) => void;
  onFitToWidth?: () => void;
  pageCurrent?: number;
  pageTotal?: number;
  wordCount?: number;
  charCount?: number;
  onOpenWordStats?: () => void;
  excelStats?: {
    average: number | null;
    count: number;
    sum: number | null;
  };
  pptViewMode?: 'normal' | 'sorter';
  onChangePptViewMode?: (mode: 'normal' | 'sorter') => void;
  onStartSlideshow?: () => void;
}

const ZOOM_PRESETS = [50, 75, 100, 125, 150, 200] as const;

const VARIANT_ACCENTS: Record<
  OfficeStatusBarVariant,
  { hex: string; label: string; activeBgLight: string; activeBgDark: string }
> = {
  word: {
    hex: '#185ABD',
    label: 'Word',
    activeBgLight: 'bg-[#185ABD]/10 text-[#185ABD] border-[#185ABD]/30',
    activeBgDark: 'bg-[#185ABD]/25 text-blue-300 border-[#185ABD]/50',
  },
  excel: {
    hex: '#107C41',
    label: 'Excel',
    activeBgLight: 'bg-[#107C41]/10 text-[#107C41] border-[#107C41]/30',
    activeBgDark: 'bg-[#107C41]/25 text-emerald-300 border-[#107C41]/50',
  },
  powerpoint: {
    hex: '#C43E1C',
    label: 'PowerPoint',
    activeBgLight: 'bg-[#C43E1C]/10 text-[#C43E1C] border-[#C43E1C]/30',
    activeBgDark: 'bg-[#C43E1C]/25 text-orange-300 border-[#C43E1C]/50',
  },
};

const STAT_NUMBER_FORMATTER = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 2,
});

function formatStatNumber(val: number): string {
  return STAT_NUMBER_FORMATTER.format(val);
}

function clampZoomPercent(pct: number): number {
  if (!Number.isFinite(pct)) {
    return 100;
  }
  return Math.max(25, Math.min(200, Math.round(pct)));
}

export function OfficeStatusBar({
  variant,
  theme = 'light',
  zoomPercent,
  onZoomChange,
  onFitToWidth,
  pageCurrent = 1,
  pageTotal = 1,
  wordCount = 0,
  charCount = 0,
  onOpenWordStats,
  excelStats,
  pptViewMode = 'normal',
  onChangePptViewMode,
  onStartSlideshow,
}: OfficeStatusBarProps) {
  const [isZoomMenuOpen, setIsZoomMenuOpen] = useState(false);
  const zoomPopoverRef = useRef<HTMLDivElement | null>(null);

  const isDark = theme === 'dark';
  const accent = VARIANT_ACCENTS[variant];
  const clampedZoom = clampZoomPercent(zoomPercent);
  const safePageTotal = Math.max(1, pageTotal);
  const safePageCurrent = Math.max(1, Math.min(safePageTotal, pageCurrent));

  useEffect(() => {
    if (!isZoomMenuOpen) {
      return;
    }

    const handlePointerDown = (event: MouseEvent | TouchEvent) => {
      const target = event.target;
      if (
        target instanceof Node &&
        zoomPopoverRef.current &&
        !zoomPopoverRef.current.contains(target)
      ) {
        setIsZoomMenuOpen(false);
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setIsZoomMenuOpen(false);
      }
    };

    window.addEventListener('mousedown', handlePointerDown);
    window.addEventListener('touchstart', handlePointerDown);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousedown', handlePointerDown);
      window.removeEventListener('touchstart', handlePointerDown);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isZoomMenuOpen]);

  const handleStepZoom = (delta: number) => {
    onZoomChange(clampZoomPercent(clampedZoom + delta));
  };

  return (
    <footer
      aria-label={`Thanh trạng thái ${accent.label}`}
      className={`shrink-0 h-9 px-2.5 sm:px-4 border-t flex items-center justify-between gap-2 text-[11px] select-none z-20 ${
        isDark
          ? 'bg-[#111622] border-white/10 text-slate-300'
          : 'bg-[#F3F4F6] border-slate-200/90 text-slate-600'
      }`}
    >
      {/* Left: Format-specific stats & view mode controls */}
      <div className="min-w-0 flex-1 flex items-center gap-1.5 sm:gap-2.5 overflow-x-auto no-scrollbar">
        {variant === 'word' && (
          <>
            <span
              className={`inline-flex items-center px-2 py-0.5 rounded-md font-semibold whitespace-nowrap shrink-0 tabular-nums border ${
                isDark
                  ? 'bg-white/5 border-white/10 text-slate-200'
                  : 'bg-white border-slate-200/80 text-slate-700'
              }`}
            >
              Trang {safePageCurrent} / {safePageTotal}
            </span>

            {onOpenWordStats ? (
              <button
                type="button"
                onClick={onOpenWordStats}
                title="Xem thống kê chi tiết tài liệu"
                className={`inline-flex items-center gap-2 px-2 py-0.5 rounded-md font-medium whitespace-nowrap shrink-0 tabular-nums transition-colors ${
                  isDark
                    ? 'hover:bg-white/10 text-slate-300 hover:text-white'
                    : 'hover:bg-white text-slate-600 hover:text-[#185ABD]'
                }`}
              >
                <span>{formatStatNumber(wordCount)} từ</span>
                <span className="opacity-40" aria-hidden="true">
                  ·
                </span>
                <span className="hidden sm:inline">
                  {formatStatNumber(charCount)} ký tự
                </span>
                <span className="sm:hidden">
                  {formatStatNumber(charCount)} kt
                </span>
              </button>
            ) : (
              <span className="inline-flex items-center gap-2 whitespace-nowrap shrink-0 tabular-nums">
                <span>{formatStatNumber(wordCount)} từ</span>
                <span className="opacity-40" aria-hidden="true">
                  ·
                </span>
                <span className="hidden sm:inline">
                  {formatStatNumber(charCount)} ký tự
                </span>
                <span className="sm:hidden">
                  {formatStatNumber(charCount)} kt
                </span>
              </span>
            )}
          </>
        )}

        {variant === 'excel' && (
          <>
            <span
              className="inline-flex items-center gap-1.5 font-semibold whitespace-nowrap shrink-0"
              style={{ color: isDark ? '#34D399' : '#107C41' }}
            >
              <span
                className="w-1.5 h-1.5 rounded-full shrink-0"
                style={{ backgroundColor: '#107C41' }}
                aria-hidden="true"
              />
              <span className="hidden sm:inline">Sẵn sàng</span>
            </span>

            {excelStats && (
              <div className="flex items-center gap-1.5 sm:gap-2.5 whitespace-nowrap shrink-0 tabular-nums">
                {excelStats.average !== null && (
                  <span
                    className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded border ${
                      isDark
                        ? 'bg-white/5 border-white/10 text-slate-200'
                        : 'bg-white border-slate-200 text-slate-700'
                    }`}
                  >
                    <span className="opacity-70">Trung bình:</span>
                    <strong className="font-semibold">
                      {formatStatNumber(excelStats.average)}
                    </strong>
                  </span>
                )}

                <span
                  className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded border ${
                    isDark
                      ? 'bg-white/5 border-white/10 text-slate-200'
                      : 'bg-white border-slate-200 text-slate-700'
                  }`}
                >
                  <span className="opacity-70">Số lượng:</span>
                  <strong className="font-semibold">{excelStats.count}</strong>
                </span>

                {excelStats.sum !== null && (
                  <span
                    className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded border font-semibold ${
                      isDark ? accent.activeBgDark : accent.activeBgLight
                    }`}
                  >
                    <span>Tổng:</span>
                    <strong>{formatStatNumber(excelStats.sum)}</strong>
                  </span>
                )}
              </div>
            )}
          </>
        )}

        {variant === 'powerpoint' && (
          <>
            <span
              className={`inline-flex items-center px-2 py-0.5 rounded-md font-semibold whitespace-nowrap shrink-0 tabular-nums border ${
                isDark
                  ? 'bg-white/5 border-white/10 text-slate-200'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
            >
              Trang chiếu {safePageCurrent} / {safePageTotal}
            </span>

            <div className="inline-flex items-center gap-1 shrink-0">
              {onChangePptViewMode && (
                <>
                  <button
                    type="button"
                    onClick={() => onChangePptViewMode('normal')}
                    aria-pressed={pptViewMode === 'normal'}
                    title="Chế độ thiết kế bình thường"
                    className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-medium whitespace-nowrap transition-colors border ${
                      pptViewMode === 'normal'
                        ? isDark
                          ? accent.activeBgDark
                          : accent.activeBgLight
                        : isDark
                          ? 'border-transparent text-slate-400 hover:text-white hover:bg-white/10'
                          : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-white'
                    }`}
                  >
                    <Presentation className="w-3 h-3 shrink-0" />
                    <span className="hidden sm:inline">Bình thường</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => onChangePptViewMode('sorter')}
                    aria-pressed={pptViewMode === 'sorter'}
                    title="Sắp xếp trang chiếu (Slide Sorter)"
                    className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-medium whitespace-nowrap transition-colors border ${
                      pptViewMode === 'sorter'
                        ? isDark
                          ? accent.activeBgDark
                          : accent.activeBgLight
                        : isDark
                          ? 'border-transparent text-slate-400 hover:text-white hover:bg-white/10'
                          : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-white'
                    }`}
                  >
                    <LayoutGrid className="w-3 h-3 shrink-0" />
                    <span className="hidden md:inline">Sắp xếp trang chiếu</span>
                    <span className="sm:inline md:hidden hidden">Sắp xếp</span>
                  </button>
                </>
              )}

              {onStartSlideshow && (
                <button
                  type="button"
                  onClick={onStartSlideshow}
                  title="Trình chiếu toàn màn hình (F5)"
                  className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-semibold whitespace-nowrap transition-colors ${
                    isDark
                      ? 'bg-[#C43E1C] hover:bg-[#AD3517] text-white'
                      : 'bg-[#C43E1C] hover:bg-[#AD3517] text-white'
                  }`}
                >
                  <MonitorPlay className="w-3 h-3 shrink-0" />
                  <span className="hidden sm:inline">Trình chiếu F5</span>
                  <span className="sm:hidden">F5</span>
                </button>
              )}
            </div>
          </>
        )}
      </div>

      {/* Right: Zoom slider (25% - 200%) + Fit to width + Presets popover */}
      <div
        ref={zoomPopoverRef}
        className="shrink-0 flex items-center gap-1 sm:gap-1.5 relative"
      >
        {onFitToWidth && (
          <button
            type="button"
            onClick={onFitToWidth}
            title="Vừa chiều rộng trang / Fit to Screen"
            aria-label="Vừa chiều rộng trang"
            className={`hidden sm:inline-flex items-center justify-center w-6 h-6 rounded-md transition-colors ${
              isDark
                ? 'hover:bg-white/10 text-slate-300 hover:text-white'
                : 'hover:bg-white text-slate-600 hover:text-slate-900'
            }`}
          >
            <Maximize2 className="w-3 h-3" />
          </button>
        )}

        <button
          type="button"
          onClick={() => handleStepZoom(-10)}
          disabled={clampedZoom <= 25}
          aria-label="Thu nhỏ tài liệu"
          title="Thu nhỏ (-10%)"
          className={`inline-flex items-center justify-center w-6 h-6 rounded-md transition-colors disabled:opacity-35 ${
            isDark
              ? 'hover:bg-white/10 text-slate-300 hover:text-white'
              : 'hover:bg-white text-slate-600 hover:text-slate-900'
          }`}
        >
          <Minus className="w-3 h-3" />
        </button>

        <input
          type="range"
          min={25}
          max={200}
          step={5}
          value={clampedZoom}
          onChange={(event) => onZoomChange(clampZoomPercent(Number(event.target.value)))}
          aria-label="Thanh trượt thu phóng tài liệu"
          style={{ accentColor: accent.hex }}
          className="hidden sm:block w-20 md:w-24 h-1.5 cursor-pointer"
        />

        <button
          type="button"
          onClick={() => handleStepZoom(10)}
          disabled={clampedZoom >= 200}
          aria-label="Phóng to tài liệu"
          title="Phóng to (+10%)"
          className={`inline-flex items-center justify-center w-6 h-6 rounded-md transition-colors disabled:opacity-35 ${
            isDark
              ? 'hover:bg-white/10 text-slate-300 hover:text-white'
              : 'hover:bg-white text-slate-600 hover:text-slate-900'
          }`}
        >
          <Plus className="w-3 h-3" />
        </button>

        <button
          type="button"
          onClick={() => setIsZoomMenuOpen((prev) => !prev)}
          aria-expanded={isZoomMenuOpen}
          aria-haspopup="menu"
          title="Chọn nhanh mức thu phóng"
          className={`min-w-[44px] h-6 px-1.5 rounded-md font-semibold tabular-nums text-center transition-colors border ${
            isDark
              ? 'bg-white/5 border-white/10 text-slate-100 hover:bg-white/10'
              : 'bg-white border-slate-200/80 text-slate-800 hover:border-slate-300'
          }`}
        >
          {clampedZoom}%
        </button>

        {isZoomMenuOpen && (
          <div
            role="menu"
            aria-label="Mức thu phóng nhanh"
            className={`absolute bottom-full right-0 mb-1.5 w-48 rounded-xl border shadow-xl p-1.5 z-50 ${
              isDark
                ? 'bg-[#1B2230] border-white/15 text-slate-100'
                : 'bg-white border-slate-200 text-slate-800'
            }`}
          >
            <div className="grid grid-cols-3 gap-1">
              {ZOOM_PRESETS.map((preset) => {
                const isSelected = clampedZoom === preset;
                return (
                  <button
                    key={preset}
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      onZoomChange(preset);
                      setIsZoomMenuOpen(false);
                    }}
                    className={`px-2 py-1.5 rounded-lg text-[11px] font-semibold tabular-nums transition-colors ${
                      isSelected
                        ? 'text-white'
                        : isDark
                          ? 'hover:bg-white/10 text-slate-200'
                          : 'hover:bg-slate-100 text-slate-700'
                    }`}
                    style={
                      isSelected ? { backgroundColor: accent.hex } : undefined
                    }
                  >
                    {preset}%
                  </button>
                );
              })}
            </div>

            {onFitToWidth && (
              <button
                type="button"
                role="menuitem"
                onClick={() => {
                  onFitToWidth();
                  setIsZoomMenuOpen(false);
                }}
                className={`mt-1.5 w-full flex items-center justify-center gap-1.5 px-2.5 py-1.5 rounded-lg text-[11px] font-semibold border transition-colors ${
                  isDark
                    ? 'border-white/10 bg-white/5 hover:bg-white/10 text-slate-100'
                    : 'border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-800'
                }`}
              >
                <Maximize2 className="w-3 h-3" />
                <span>Vừa chiều rộng trang</span>
              </button>
            )}
          </div>
        )}
      </div>
    </footer>
  );
}

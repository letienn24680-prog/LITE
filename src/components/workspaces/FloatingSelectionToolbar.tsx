import { useCallback, useEffect, useRef, useState } from 'react';
import type { MouseEvent, ReactElement, RefObject } from 'react';
import {
  Bold,
  Highlighter,
  Italic,
  RemoveFormatting,
  Underline,
} from 'lucide-react';

export type FloatingBlockTag = 'p' | 'h1' | 'h2';

export interface FloatingColorSwatch {
  label: string;
  value: string;
  chipClass?: string;
}

export interface FloatingSelectionActiveState {
  block?: 'p' | 'h1' | 'h2' | 'h3';
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
  highlightColor?: string;
}

export interface FloatingSelectionToolbarProps {
  containerRef: RefObject<HTMLElement | null>;
  disabled?: boolean;
  activeState?: FloatingSelectionActiveState;
  onSelectBlock: (tag: FloatingBlockTag) => void;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleUnderline: () => void;
  onSelectTextColor: (color: string) => void;
  onSelectHighlight: (color: string) => void;
  onClearFormat: () => void;
  textColorSwatches?: readonly FloatingColorSwatch[];
  highlightSwatches?: readonly FloatingColorSwatch[];
}

const DEFAULT_TEXT_COLOR_SWATCHES: readonly FloatingColorSwatch[] = [
  { label: 'Mặc định', value: '#0F172A', chipClass: 'bg-slate-900 border-white/40' },
  { label: 'Trắng', value: '#FFFFFF', chipClass: 'bg-white border-slate-300' },
  { label: 'Xanh Indigo', value: '#4F46E5', chipClass: 'bg-indigo-600 border-indigo-300/60' },
  { label: 'Xanh dương', value: '#2563EB', chipClass: 'bg-blue-600 border-blue-300/60' },
  { label: 'Xanh lá', value: '#15803D', chipClass: 'bg-green-700 border-green-300/60' },
  { label: 'Cam đậm', value: '#EA580C', chipClass: 'bg-orange-600 border-orange-300/60' },
  { label: 'Đỏ Rose', value: '#E11D48', chipClass: 'bg-rose-600 border-rose-300/60' },
];

const DEFAULT_HIGHLIGHT_SWATCHES: readonly FloatingColorSwatch[] = [
  { label: 'Không tô nền', value: 'transparent', chipClass: 'bg-slate-800 border-white/35' },
  { label: 'Vàng sáng', value: '#FEF08A', chipClass: 'bg-[#FEF08A] border-amber-300' },
  { label: 'Xanh lá nhạt', value: '#BBF7D0', chipClass: 'bg-[#BBF7D0] border-emerald-300' },
  { label: 'Xanh dương nhạt', value: '#BFDBFE', chipClass: 'bg-[#BFDBFE] border-blue-300' },
  { label: 'Hồng nhạt', value: '#FBCFE8', chipClass: 'bg-[#FBCFE8] border-pink-300' },
];

interface ToolbarCoordinates {
  top: number;
  left: number;
  placement: 'above' | 'below';
}

function getSelectionAnchorRect(container: HTMLElement | null): DOMRect | null {
  if (!container || typeof window === 'undefined') {
    return null;
  }

  // 1. Check <textarea> or <input> selection inside container (e.g. SlideWorkspace)
  const activeEl = document.activeElement;
  if (
    (activeEl instanceof HTMLTextAreaElement || activeEl instanceof HTMLInputElement) &&
    container.contains(activeEl)
  ) {
    const start = activeEl.selectionStart ?? 0;
    const end = activeEl.selectionEnd ?? 0;
    if (end > start) {
      const selectedText = activeEl.value.slice(start, end).trim();
      if (selectedText.length > 0) {
        const inputRect = activeEl.getBoundingClientRect();
        if (
          (inputRect.width > 0 || inputRect.height > 0) &&
          inputRect.bottom > 0 &&
          inputRect.top < window.innerHeight
        ) {
          return inputRect;
        }
      }
    }
  }

  // 2. Check DOM Range selection inside contentEditable container (e.g. WordWorkspace)
  const sel = window.getSelection();
  if (!sel || sel.isCollapsed || sel.rangeCount === 0) {
    return null;
  }

  if (sel.toString().trim().length === 0) {
    return null;
  }

  const range = sel.getRangeAt(0);
  const ancestor = range.commonAncestorContainer;
  const ancestorElement =
    ancestor.nodeType === Node.TEXT_NODE ? ancestor.parentElement : (ancestor as Node);

  if (!ancestorElement || !container.contains(ancestorElement)) {
    return null;
  }

  let rect = range.getBoundingClientRect();
  if (rect.width === 0 && rect.height === 0) {
    const rects = range.getClientRects();
    if (rects.length > 0) {
      rect = rects[0];
    }
  }

  if (
    (rect.width === 0 && rect.height === 0) ||
    rect.bottom <= 0 ||
    rect.top >= window.innerHeight
  ) {
    return null;
  }

  return rect;
}

export function FloatingSelectionToolbar({
  containerRef,
  disabled = false,
  activeState,
  onSelectBlock,
  onToggleBold,
  onToggleItalic,
  onToggleUnderline,
  onSelectTextColor,
  onSelectHighlight,
  onClearFormat,
  textColorSwatches = DEFAULT_TEXT_COLOR_SWATCHES,
  highlightSwatches = DEFAULT_HIGHLIGHT_SWATCHES,
}: FloatingSelectionToolbarProps): ReactElement | null {
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const [coords, setCoords] = useState<ToolbarCoordinates | null>(null);
  const [openMenu, setOpenMenu] = useState<'textColor' | 'highlight' | null>(null);

  const updateToolbarPosition = useCallback(() => {
    if (disabled) {
      setCoords(null);
      setOpenMenu(null);
      return;
    }

    const rect = getSelectionAnchorRect(containerRef.current);
    if (!rect) {
      setCoords(null);
      setOpenMenu(null);
      return;
    }

    const viewportW = window.innerWidth;
    const measuredW = toolbarRef.current?.offsetWidth ?? 336;
    const effectiveHalfW = Math.min(measuredW, viewportW - 24) / 2;

    const rawCenterX = rect.left + rect.width / 2;
    const clampedLeft = Math.max(
      12 + effectiveHalfW,
      Math.min(viewportW - 12 - effectiveHalfW, rawCenterX)
    );

    const placement: 'above' | 'below' = rect.top < 80 ? 'below' : 'above';
    const top =
      placement === 'above'
        ? Math.max(12, rect.top - 10)
        : Math.min(window.innerHeight - 56, rect.bottom + 10);

    setCoords({
      top,
      left: clampedLeft,
      placement,
    });
  }, [containerRef, disabled]);

  useEffect(() => {
    if (disabled) {
      setCoords(null);
      setOpenMenu(null);
      return;
    }

    let rafId = 0;
    const scheduleUpdate = () => {
      window.cancelAnimationFrame(rafId);
      rafId = window.requestAnimationFrame(() => {
        updateToolbarPosition();
      });
    };

    document.addEventListener('selectionchange', scheduleUpdate);
    document.addEventListener('mouseup', scheduleUpdate);
    document.addEventListener('touchend', scheduleUpdate);
    document.addEventListener('keyup', scheduleUpdate);
    document.addEventListener('select', scheduleUpdate, true);
    window.addEventListener('scroll', scheduleUpdate, true);
    window.addEventListener('resize', scheduleUpdate);

    return () => {
      window.cancelAnimationFrame(rafId);
      document.removeEventListener('selectionchange', scheduleUpdate);
      document.removeEventListener('mouseup', scheduleUpdate);
      document.removeEventListener('touchend', scheduleUpdate);
      document.removeEventListener('keyup', scheduleUpdate);
      document.removeEventListener('select', scheduleUpdate, true);
      window.removeEventListener('scroll', scheduleUpdate, true);
      window.removeEventListener('resize', scheduleUpdate);
    };
  }, [disabled, updateToolbarPosition]);

  const handlePreventMouseDown = (event: MouseEvent<HTMLElement>) => {
    event.preventDefault();
    event.stopPropagation();
  };

  if (disabled || !coords) {
    return null;
  }

  const activeTextColor = activeState?.textColor || '#60A5FA';
  const activeHighlightColor =
    activeState?.highlightColor && activeState.highlightColor !== 'transparent'
      ? activeState.highlightColor
      : '#FEF08A';

  const openSubmenuBelow = coords.placement === 'below' || coords.top < 150;

  return (
    <div
      ref={toolbarRef}
      role="toolbar"
      aria-label="Thanh định dạng nhanh văn bản đang chọn"
      onMouseDown={handlePreventMouseDown}
      style={{
        top: `${coords.top}px`,
        left: `${coords.left}px`,
        transform:
          coords.placement === 'above'
            ? 'translate(-50%, -100%)'
            : 'translate(-50%, 0)',
      }}
      className="fixed z-50 max-w-[calc(100vw-24px)] select-none rounded-xl border border-white/15 bg-slate-900/95 px-1.5 py-1 text-white shadow-xl backdrop-blur-md"
    >
      <div className="flex items-center gap-0.5 overflow-x-auto no-scrollbar">
        {/* Group 1: Block headings H1, H2, P */}
        {(['h1', 'h2', 'p'] as const).map((tag) => {
          const isActive = activeState?.block === tag;
          const label = tag.toUpperCase();
          return (
            <button
              key={tag}
              type="button"
              onMouseDown={handlePreventMouseDown}
              onClick={() => {
                onSelectBlock(tag);
                setOpenMenu(null);
              }}
              aria-pressed={isActive}
              title={
                tag === 'h1'
                  ? 'Tiêu đề lớn (H1)'
                  : tag === 'h2'
                    ? 'Tiêu đề mục (H2)'
                    : 'Đoạn văn thường (P)'
              }
              className={`flex h-7 min-w-[28px] shrink-0 items-center justify-center rounded-lg px-1.5 text-[11px] font-bold tracking-tight transition-colors ${
                isActive
                  ? 'bg-indigo-600 text-white'
                  : 'text-slate-200 hover:bg-white/10 hover:text-white'
              }`}
            >
              {label}
            </button>
          );
        })}

        <span aria-hidden="true" className="mx-0.5 h-4 w-px shrink-0 bg-white/15" />

        {/* Group 2: Bold, Italic, Underline */}
        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() => {
            onToggleBold();
            setOpenMenu(null);
          }}
          aria-pressed={Boolean(activeState?.bold)}
          title="In đậm (Ctrl+B)"
          aria-label="In đậm"
          className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg transition-colors ${
            activeState?.bold
              ? 'bg-indigo-600 text-white'
              : 'text-slate-200 hover:bg-white/10 hover:text-white'
          }`}
        >
          <Bold className="h-3.5 w-3.5" strokeWidth={2.4} aria-hidden="true" />
        </button>

        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() => {
            onToggleItalic();
            setOpenMenu(null);
          }}
          aria-pressed={Boolean(activeState?.italic)}
          title="In nghiêng (Ctrl+I)"
          aria-label="In nghiêng"
          className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg transition-colors ${
            activeState?.italic
              ? 'bg-indigo-600 text-white'
              : 'text-slate-200 hover:bg-white/10 hover:text-white'
          }`}
        >
          <Italic className="h-3.5 w-3.5" strokeWidth={2.2} aria-hidden="true" />
        </button>

        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() => {
            onToggleUnderline();
            setOpenMenu(null);
          }}
          aria-pressed={Boolean(activeState?.underline)}
          title="Gạch chân (Ctrl+U)"
          aria-label="Gạch chân"
          className={`flex h-7 w-7 shrink-0 items-center justify-center rounded-lg transition-colors ${
            activeState?.underline
              ? 'bg-indigo-600 text-white'
              : 'text-slate-200 hover:bg-white/10 hover:text-white'
          }`}
        >
          <Underline className="h-3.5 w-3.5" strokeWidth={2.2} aria-hidden="true" />
        </button>

        <span aria-hidden="true" className="mx-0.5 h-4 w-px shrink-0 bg-white/15" />

        {/* Group 3: Text color A */}
        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() =>
            setOpenMenu((prev) => (prev === 'textColor' ? null : 'textColor'))
          }
          aria-expanded={openMenu === 'textColor'}
          title="Màu chữ"
          aria-label="Màu chữ"
          className={`flex h-7 w-7 shrink-0 flex-col items-center justify-center rounded-lg transition-colors ${
            openMenu === 'textColor'
              ? 'bg-white/20 text-white'
              : 'text-slate-200 hover:bg-white/10 hover:text-white'
          }`}
        >
          <span className="text-[11px] font-bold leading-none">A</span>
          <span
            aria-hidden="true"
            style={{ backgroundColor: activeTextColor }}
            className="mt-0.5 h-[2.5px] w-3.5 rounded-full"
          />
        </button>

        {/* Group 4: Highlighter */}
        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() =>
            setOpenMenu((prev) => (prev === 'highlight' ? null : 'highlight'))
          }
          aria-expanded={openMenu === 'highlight'}
          title="Tô sáng văn bản"
          aria-label="Tô sáng văn bản"
          className={`flex h-7 w-7 shrink-0 flex-col items-center justify-center rounded-lg transition-colors ${
            openMenu === 'highlight'
              ? 'bg-white/20 text-white'
              : 'text-slate-200 hover:bg-white/10 hover:text-white'
          }`}
        >
          <Highlighter className="h-3.5 w-3.5" aria-hidden="true" />
          <span
            aria-hidden="true"
            style={{ backgroundColor: activeHighlightColor }}
            className="mt-0.5 h-[2px] w-3.5 rounded-full"
          />
        </button>

        <span aria-hidden="true" className="mx-0.5 h-4 w-px shrink-0 bg-white/15" />

        {/* Group 5: Clear formatting */}
        <button
          type="button"
          onMouseDown={handlePreventMouseDown}
          onClick={() => {
            onClearFormat();
            setOpenMenu(null);
          }}
          title="Xóa định dạng"
          aria-label="Xóa định dạng"
          className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg text-slate-300 transition-colors hover:bg-rose-500/20 hover:text-rose-200"
        >
          <RemoveFormatting className="h-3.5 w-3.5" aria-hidden="true" />
        </button>
      </div>

      {/* Submenu: Text Color Swatches */}
      {openMenu === 'textColor' && (
        <div
          onMouseDown={handlePreventMouseDown}
          className={`absolute left-1/2 z-50 flex -translate-x-1/2 items-center gap-1.5 rounded-xl border border-white/15 bg-slate-900 px-2.5 py-2 shadow-2xl ${
            openSubmenuBelow ? 'top-full mt-1.5' : 'bottom-full mb-1.5'
          }`}
        >
          {textColorSwatches.map((swatch) => {
            const isCurrent =
              activeState?.textColor?.toLowerCase() === swatch.value.toLowerCase();
            return (
              <button
                key={`tc-${swatch.value}`}
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => {
                  onSelectTextColor(swatch.value);
                  setOpenMenu(null);
                }}
                title={`Màu chữ: ${swatch.label}`}
                aria-label={`Màu chữ: ${swatch.label}`}
                style={{ backgroundColor: swatch.value }}
                className={`h-5 w-5 shrink-0 rounded-full border transition-transform hover:scale-110 ${
                  swatch.chipClass ?? 'border-white/30'
                } ${isCurrent ? 'ring-2 ring-indigo-400 ring-offset-1 ring-offset-slate-900' : ''}`}
              />
            );
          })}
        </div>
      )}

      {/* Submenu: Highlight Swatches */}
      {openMenu === 'highlight' && (
        <div
          onMouseDown={handlePreventMouseDown}
          className={`absolute left-1/2 z-50 flex -translate-x-1/2 items-center gap-1.5 rounded-xl border border-white/15 bg-slate-900 px-2.5 py-2 shadow-2xl ${
            openSubmenuBelow ? 'top-full mt-1.5' : 'bottom-full mb-1.5'
          }`}
        >
          {highlightSwatches.map((swatch) => {
            const isTransparent = swatch.value === 'transparent';
            const isCurrent =
              activeState?.highlightColor?.toLowerCase() === swatch.value.toLowerCase();
            return (
              <button
                key={`hl-${swatch.value}`}
                type="button"
                onMouseDown={handlePreventMouseDown}
                onClick={() => {
                  onSelectHighlight(swatch.value);
                  setOpenMenu(null);
                }}
                title={`Tô sáng: ${swatch.label}`}
                aria-label={`Tô sáng: ${swatch.label}`}
                style={
                  isTransparent ? undefined : { backgroundColor: swatch.value }
                }
                className={`relative flex h-5 w-5 shrink-0 items-center justify-center rounded-full border transition-transform hover:scale-110 ${
                  swatch.chipClass ?? 'border-white/30'
                } ${isCurrent ? 'ring-2 ring-indigo-400 ring-offset-1 ring-offset-slate-900' : ''}`}
              >
                {isTransparent && (
                  <span
                    aria-hidden="true"
                    className="h-3.5 w-[1.5px] rotate-45 bg-rose-400"
                  />
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

import { useEffect, useRef, useState } from 'react';
import {
  ArrowLeftRight,
  BookOpen,
  Check,
  ChevronLeft,
  ChevronRight,
  Copy,
  FileSearch,
  Frame,
  Link2,
  MessageSquarePlus,
  Replace,
  Ruler,
  Search,
  Shapes,
  Sigma,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';

import type { WordCommentItem } from '../../types/document';
export type { WordCommentItem };

export type WordPaperSize = 'A4' | 'A5' | 'Letter';
export type WordPageOrientation = 'portrait' | 'landscape';
export type WordMarginPreset = 'normal' | 'narrow' | 'wide' | 'custom';
export type WordColumnCount = 1 | 2 | 3;
export type WordUnderlineStyle = 'solid' | 'double' | 'dashed' | 'wavy';
export type WordCaseTransform = 'upper' | 'lower' | 'title' | 'sentence';
export type WordShapeKind = 'rect' | 'roundRect' | 'circle' | 'arrow' | 'badge';

export interface WordDocumentStats {
  pages: number;
  words: number;
  charsWithSpaces: number;
  charsWithoutSpaces: number;
  paragraphs: number;
  lines: number;
  readingMinutes: number;
}

export interface WordPageDimensionsCm {
  widthCm: number;
  heightCm: number;
  widthPx: number;
  minHeightPx: number;
}

const PX_PER_CM = 37.7952755906;

export function getWordPageDimensions(
  paperSize: WordPaperSize,
  orientation: WordPageOrientation
): WordPageDimensionsCm {
  const base =
    paperSize === 'A5'
      ? { w: 14.8, h: 21.0 }
      : paperSize === 'Letter'
        ? { w: 21.59, h: 27.94 }
        : { w: 21.0, h: 29.7 };

  const widthCm = orientation === 'landscape' ? base.h : base.w;
  const heightCm = orientation === 'landscape' ? base.w : base.h;

  return {
    widthCm,
    heightCm,
    widthPx: Math.round(widthCm * PX_PER_CM),
    minHeightPx: Math.round(heightCm * PX_PER_CM),
  };
}

export function cmToPx(cm: number): number {
  return Math.round(cm * PX_PER_CM);
}

export function computeWordDocumentStats(
  editorEl: HTMLElement | null,
  detectedPageCount: number
): WordDocumentStats {
  if (!editorEl) {
    return {
      pages: Math.max(1, detectedPageCount),
      words: 0,
      charsWithSpaces: 0,
      charsWithoutSpaces: 0,
      paragraphs: 0,
      lines: 0,
      readingMinutes: 1,
    };
  }

  const clone = editorEl.cloneNode(true) as HTMLElement;
  clone.querySelectorAll('style').forEach((node) => node.remove());
  const rawText = clone.innerText ?? clone.textContent ?? '';
  const trimmed = rawText.trim();
  const words = trimmed ? trimmed.split(/\s+/).filter(Boolean).length : 0;
  const charsWithSpaces = rawText.replace(/\r?\n/g, '').length;
  const charsWithoutSpaces = rawText.replace(/\s+/g, '').length;
  const paragraphNodes = clone.querySelectorAll('p, h1, h2, h3, h4, li, blockquote');
  const paragraphs = Math.max(
    paragraphNodes.length,
    trimmed ? trimmed.split(/\n+/).filter((line) => line.trim().length > 0).length : 0
  );
  const lines = trimmed
    ? rawText
        .split(/\r?\n/)
        .filter((line) => line.trim().length > 0).length
    : 0;

  return {
    pages: Math.max(1, detectedPageCount),
    words,
    charsWithSpaces,
    charsWithoutSpaces,
    paragraphs,
    lines: Math.max(lines, Math.ceil(words / 12)),
    readingMinutes: Math.max(1, Math.ceil(words / 200)),
  };
}

export function applyCaseTransformToText(
  text: string,
  mode: WordCaseTransform
): string {
  switch (mode) {
    case 'upper':
      return text.toLocaleUpperCase('vi-VN');
    case 'lower':
      return text.toLocaleLowerCase('vi-VN');
    case 'title':
      return text
        .toLocaleLowerCase('vi-VN')
        .replace(/(^|\s+)(\S)/g, (_, space: string, ch: string) => {
          return `${space}${ch.toLocaleUpperCase('vi-VN')}`;
        });
    case 'sentence': {
      const lowered = text.toLocaleLowerCase('vi-VN');
      return lowered.replace(
        /(^\s*|[.!?]\s+)(\S)/g,
        (_, prefix: string, ch: string) =>
          `${prefix}${ch.toLocaleUpperCase('vi-VN')}`
      );
    }
  }
}

function escapeHtmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function buildShapeBlockHtml(
  kind: WordShapeKind,
  label: string,
  fillColor: string,
  strokeColor: string
): string {
  const safeLabel = escapeHtmlAttr(label.trim() || 'Hình khối');
  const safeFill = escapeHtmlAttr(fillColor);
  const safeStroke = escapeHtmlAttr(strokeColor);

  let svgBody = `<rect x="4" y="4" width="212" height="72" rx="4" fill="${safeFill}" stroke="${safeStroke}" stroke-width="2.5"/>`;
  if (kind === 'roundRect') {
    svgBody = `<rect x="4" y="4" width="212" height="72" rx="18" fill="${safeFill}" stroke="${safeStroke}" stroke-width="2.5"/>`;
  } else if (kind === 'circle') {
    svgBody = `<ellipse cx="110" cy="40" rx="102" ry="36" fill="${safeFill}" stroke="${safeStroke}" stroke-width="2.5"/>`;
  } else if (kind === 'arrow') {
    svgBody = `<polygon points="8,24 156,24 156,8 214,40 156,72 156,56 8,56" fill="${safeFill}" stroke="${safeStroke}" stroke-width="2.5"/>`;
  } else if (kind === 'badge') {
    svgBody = `<polygon points="110,4 134,26 212,26 150,46 172,76 110,58 48,76 70,46 8,26 86,26" fill="${safeFill}" stroke="${safeStroke}" stroke-width="2"/>`;
  }

  return `<div data-word-shape="${kind}" style="display:inline-flex;align-items:center;justify-content:center;position:relative;margin:10px 8px;vertical-align:middle;"><svg width="220" height="80" viewBox="0 0 220 80" xmlns="http://www.w3.org/2000/svg">${svgBody}<text x="110" y="45" text-anchor="middle" fill="#0F172A" font-size="13" font-weight="600" font-family="Arial, sans-serif">${safeLabel}</text></svg></div><p><br></p>`;
}

export function buildTextBoxHtml(
  title: string,
  body: string,
  accentColor: string
): string {
  const safeTitle = escapeHtmlAttr(title.trim() || 'Hộp văn bản');
  const safeBody = escapeHtmlAttr(
    body.trim() || 'Nhập nội dung nổi bật hoặc ghi chú quan trọng tại đây.'
  );
  const safeAccent = escapeHtmlAttr(accentColor);
  return `<div data-word-textbox="true" style="border:2px solid ${safeAccent};border-radius:10px;background:#F8FAFC;padding:12px 16px;margin:14px 0;color:#0F172A;"><div style="font-weight:700;color:${safeAccent};margin-bottom:4px;font-size:0.95em;">${safeTitle}</div><div style="font-size:0.95em;line-height:1.5;">${safeBody}</div></div><p><br></p>`;
}

export function buildEquationHtml(latexOrText: string): string {
  const safeEq = escapeHtmlAttr(latexOrText.trim());
  return `<span data-word-equation="true" style="display:inline-block;padding:3px 10px;margin:2px 4px;border-radius:6px;background:rgba(24,90,189,0.08);border:1px solid rgba(24,90,189,0.22);font-family:'Cambria Math','Times New Roman',serif;font-style:italic;font-size:1.04em;">${safeEq}</span>&nbsp;`;
}

export const WORD_SYMBOLS: string[] = [
  '©',
  '®',
  '™',
  '§',
  '¶',
  '°',
  '±',
  '×',
  '÷',
  '≠',
  '≈',
  '≤',
  '≥',
  '∞',
  '√',
  '∑',
  '∏',
  '∫',
  'Δ',
  'π',
  'α',
  'β',
  'θ',
  'Ω',
  '→',
  '←',
  '↔',
  '⇒',
  '✓',
  '✗',
  '★',
  '•',
  '—',
  '₫',
  '€',
  '£',
];

export const WORD_EQUATION_PRESETS: { label: string; preview: string }[] = [
  { label: 'Nghiệm bậc hai', preview: 'x = (−b ± √(b² − 4ac)) / 2a' },
  { label: 'Định lý Pytago', preview: 'a² + b² = c²' },
  { label: 'Diện tích hình tròn', preview: 'S = π · r²' },
  { label: 'Tổng dãy số', preview: 'Sₙ = ∑ᵢ₌₁ⁿ aᵢ' },
  { label: 'Tích phân xác định', preview: '∫ₐᵇ f(x) dx = F(b) − F(a)' },
  { label: 'Lãi kép', preview: 'A = P · (1 + r / n)ⁿᵗ' },
  { label: 'Định luật Ohm', preview: 'I = U / R' },
  { label: 'Công thức Einstein', preview: 'E = m · c²' },
];

export interface WordHorizontalRulerProps {
  pageWidthPx: number;
  pageWidthCm: number;
  marginLeftCm: number;
  marginRightCm: number;
  firstLineIndentCm: number;
  isDark: boolean;
  disabled?: boolean;
  onChangeMargins: (next: {
    marginLeftCm: number;
    marginRightCm: number;
    firstLineIndentCm: number;
  }) => void;
}

export function WordHorizontalRuler({
  pageWidthPx,
  pageWidthCm,
  marginLeftCm,
  marginRightCm,
  firstLineIndentCm,
  isDark,
  disabled = false,
  onChangeMargins,
}: WordHorizontalRulerProps) {
  const trackRef = useRef<HTMLDivElement | null>(null);
  const [draggingHandle, setDraggingHandle] = useState<
    'left' | 'right' | 'indent' | null
  >(null);

  const clampMargins = (
    nextLeft: number,
    nextRight: number,
    nextIndent: number
  ) => {
    const safeLeft = Math.max(0.5, Math.min(pageWidthCm * 0.38, nextLeft));
    const safeRight = Math.max(0.5, Math.min(pageWidthCm * 0.38, nextRight));
    const maxIndent = Math.max(0, pageWidthCm - safeLeft - safeRight - 3);
    const safeIndent = Math.max(0, Math.min(Math.min(3.5, maxIndent), nextIndent));
    return {
      marginLeftCm: Math.round(safeLeft * 10) / 10,
      marginRightCm: Math.round(safeRight * 10) / 10,
      firstLineIndentCm: Math.round(safeIndent * 10) / 10,
    };
  };

  useEffect(() => {
    if (!draggingHandle || disabled) return;

    const handlePointerMove = (event: PointerEvent) => {
      const track = trackRef.current;
      if (!track) return;
      const rect = track.getBoundingClientRect();
      if (rect.width <= 0) return;

      const ratio = Math.max(0, Math.min(1, (event.clientX - rect.left) / rect.width));
      const cmPos = ratio * pageWidthCm;

      if (draggingHandle === 'left') {
        onChangeMargins(clampMargins(cmPos, marginRightCm, firstLineIndentCm));
      } else if (draggingHandle === 'right') {
        onChangeMargins(
          clampMargins(marginLeftCm, pageWidthCm - cmPos, firstLineIndentCm)
        );
      } else {
        onChangeMargins(
          clampMargins(marginLeftCm, marginRightCm, cmPos - marginLeftCm)
        );
      }
    };

    const handlePointerUp = () => {
      setDraggingHandle(null);
    };

    window.addEventListener('pointermove', handlePointerMove);
    window.addEventListener('pointerup', handlePointerUp);
    return () => {
      window.removeEventListener('pointermove', handlePointerMove);
      window.removeEventListener('pointerup', handlePointerUp);
    };
  }, [
    disabled,
    draggingHandle,
    firstLineIndentCm,
    marginLeftCm,
    marginRightCm,
    onChangeMargins,
    pageWidthCm,
  ]);

  const leftPercent = Math.min(45, (marginLeftCm / pageWidthCm) * 100);
  const rightPercent = Math.min(45, (marginRightCm / pageWidthCm) * 100);
  const indentPercent = Math.min(
    90,
    ((marginLeftCm + firstLineIndentCm) / pageWidthCm) * 100
  );
  const totalWholeCm = Math.floor(pageWidthCm);

  return (
    <div
      style={{ width: `${pageWidthPx}px`, maxWidth: '100%' }}
      className="mx-auto mb-2 select-none"
    >
      <div
        ref={trackRef}
        aria-label="Thước kẻ ngang căn lề và thụt đầu dòng"
        className={`relative h-6 rounded-md border overflow-hidden ${
          isDark
            ? 'bg-[#171C27] border-white/15 text-slate-400'
            : 'bg-white border-[#DDD7CD] text-slate-500'
        }`}
      >
        {/* Shaded left & right margin zones */}
        <div
          aria-hidden="true"
          style={{ width: `${leftPercent}%` }}
          className={`absolute inset-y-0 left-0 ${
            isDark ? 'bg-slate-800/80' : 'bg-[#EAE6DE]'
          }`}
        />
        <div
          aria-hidden="true"
          style={{ width: `${rightPercent}%` }}
          className={`absolute inset-y-0 right-0 ${
            isDark ? 'bg-slate-800/80' : 'bg-[#EAE6DE]'
          }`}
        />

        {/* Centimeter ticks */}
        {Array.from({ length: totalWholeCm + 1 }, (_, cm) => {
          const pct = (cm / pageWidthCm) * 100;
          return (
            <div
              key={`cm-${cm}`}
              style={{ left: `${pct}%` }}
              className="absolute inset-y-0 pointer-events-none"
            >
              <div
                className={`w-px h-2 ${
                  isDark ? 'bg-slate-500/60' : 'bg-slate-400/70'
                }`}
              />
              {cm > 0 && cm < pageWidthCm - 0.3 && (
                <span className="absolute top-1.5 -translate-x-1/2 text-[9px] font-medium tabular-nums">
                  {cm}
                </span>
              )}
            </div>
          );
        })}

        {/* Left Margin Handle */}
        <button
          type="button"
          disabled={disabled}
          onPointerDown={(e) => {
            if (disabled) return;
            e.preventDefault();
            setDraggingHandle('left');
          }}
          title={`Lề trái: ${marginLeftCm.toFixed(1)} cm (kéo để chỉnh)`}
          aria-label={`Kéo chỉnh lề trái (${marginLeftCm.toFixed(1)} cm)`}
          style={{ left: `${leftPercent}%` }}
          className={`absolute bottom-0 -translate-x-1/2 w-3.5 h-3.5 rounded-t-sm border transition-colors ${
            draggingHandle === 'left'
              ? 'bg-[#185ABD] border-[#185ABD] scale-110'
              : isDark
                ? 'bg-slate-300 border-slate-500 hover:bg-blue-400'
                : 'bg-[#185ABD] border-white shadow-sm hover:bg-blue-700'
          } disabled:opacity-40 cursor-ew-resize`}
        />

        {/* First-line Indent Handle (top triangle) */}
        <button
          type="button"
          disabled={disabled}
          onPointerDown={(e) => {
            if (disabled) return;
            e.preventDefault();
            setDraggingHandle('indent');
          }}
          title={`Thụt lề đầu dòng: ${firstLineIndentCm.toFixed(1)} cm (kéo để chỉnh)`}
          aria-label={`Kéo chỉnh thụt lề đầu dòng (${firstLineIndentCm.toFixed(1)} cm)`}
          style={{ left: `${indentPercent}%` }}
          className={`absolute top-0 -translate-x-1/2 w-3.5 h-3 rounded-b-sm border transition-colors ${
            draggingHandle === 'indent'
              ? 'bg-amber-500 border-amber-600 scale-110'
              : isDark
                ? 'bg-amber-400 border-slate-700 hover:bg-amber-300'
                : 'bg-sky-600 border-white shadow-sm hover:bg-sky-500'
          } disabled:opacity-40 cursor-ew-resize`}
        />

        {/* Right Margin Handle */}
        <button
          type="button"
          disabled={disabled}
          onPointerDown={(e) => {
            if (disabled) return;
            e.preventDefault();
            setDraggingHandle('right');
          }}
          title={`Lề phải: ${marginRightCm.toFixed(1)} cm (kéo để chỉnh)`}
          aria-label={`Kéo chỉnh lề phải (${marginRightCm.toFixed(1)} cm)`}
          style={{ left: `${100 - rightPercent}%` }}
          className={`absolute bottom-0 -translate-x-1/2 w-3.5 h-3.5 rounded-t-sm border transition-colors ${
            draggingHandle === 'right'
              ? 'bg-[#185ABD] border-[#185ABD] scale-110'
              : isDark
                ? 'bg-slate-300 border-slate-500 hover:bg-blue-400'
                : 'bg-[#185ABD] border-white shadow-sm hover:bg-blue-700'
          } disabled:opacity-40 cursor-ew-resize`}
        />
      </div>

      <div
        className={`mt-1 flex items-center justify-between text-[10.5px] px-1 ${
          isDark ? 'text-slate-400' : 'text-slate-500'
        }`}
      >
        <span className="inline-flex items-center gap-1 tabular-nums">
          <Ruler className="h-3 w-3 text-[#185ABD]" />
          Lề trái {marginLeftCm.toFixed(1)}cm · Thụt đầu dòng{' '}
          {firstLineIndentCm.toFixed(1)}cm · Lề phải {marginRightCm.toFixed(1)}cm
        </span>
        {!disabled && (
          <button
            type="button"
            onClick={() =>
              onChangeMargins({
                marginLeftCm: 2.0,
                marginRightCm: 2.0,
                firstLineIndentCm: 0,
              })
            }
            className="text-[10.5px] font-medium text-[#185ABD] hover:underline"
          >
            Đặt lại chuẩn
          </button>
        )}
      </div>
    </div>
  );
}

export interface WordTableGridPickerProps {
  isDark: boolean;
  onSelect: (rows: number, cols: number) => void;
}

export function WordTableGridPicker({
  isDark,
  onSelect,
}: WordTableGridPickerProps) {
  const [hoverPos, setHoverPos] = useState<{ r: number; c: number }>({
    r: 3,
    c: 3,
  });

  return (
    <div
      className={`p-3 rounded-xl border shadow-xl w-[228px] space-y-2 ${
        isDark
          ? 'bg-[#1E2430] border-white/15 text-slate-100'
          : 'bg-white border-[#E5E0D8] text-slate-900'
      }`}
    >
      <div className="flex items-center justify-between text-xs font-semibold">
        <span>Chèn bảng</span>
        <span className="text-[#185ABD] tabular-nums">
          {hoverPos.r} × {hoverPos.c}
        </span>
      </div>
      <div className="grid grid-cols-8 gap-1">
        {Array.from({ length: 6 }, (_, rIdx) =>
          Array.from({ length: 8 }, (__, cIdx) => {
            const r = rIdx + 1;
            const c = cIdx + 1;
            const active = r <= hoverPos.r && c <= hoverPos.c;
            return (
              <button
                key={`grid-${r}-${c}`}
                type="button"
                aria-label={`Chèn bảng ${r} hàng ${c} cột`}
                onMouseEnter={() => setHoverPos({ r, c })}
                onFocus={() => setHoverPos({ r, c })}
                onClick={() => onSelect(r, c)}
                className={`h-5 w-5 rounded-[3px] border transition-colors ${
                  active
                    ? 'bg-[#185ABD]/25 border-[#185ABD]'
                    : isDark
                      ? 'bg-white/[0.04] border-white/15 hover:border-blue-400'
                      : 'bg-slate-50 border-slate-300 hover:border-blue-500'
                }`}
              />
            );
          })
        )}
      </div>
    </div>
  );
}

export interface WordFindReplaceBarProps {
  isDark: boolean;
  query: string;
  replaceText: string;
  matchCount: number;
  activeMatchIndex: number;
  caseSensitive: boolean;
  onChangeQuery: (val: string) => void;
  onChangeReplaceText: (val: string) => void;
  onToggleCaseSensitive: () => void;
  onFindPrev: () => void;
  onFindNext: () => void;
  onReplaceOne: () => void;
  onReplaceAll: () => void;
  onClose: () => void;
}

export function WordFindReplaceBar({
  isDark,
  query,
  replaceText,
  matchCount,
  activeMatchIndex,
  caseSensitive,
  onChangeQuery,
  onChangeReplaceText,
  onToggleCaseSensitive,
  onFindPrev,
  onFindNext,
  onReplaceOne,
  onReplaceAll,
  onClose,
}: WordFindReplaceBarProps) {
  return (
    <div
      role="region"
      aria-label="Tìm kiếm và thay thế văn bản"
      className={`shrink-0 border-b px-3 sm:px-4 py-2 flex flex-wrap items-center gap-2 z-20 ${
        isDark
          ? 'bg-[#171C27] border-white/10 text-slate-100'
          : 'bg-[#F6F8FC] border-[#DCE4F0] text-slate-900'
      }`}
    >
      <div className="flex items-center gap-1.5 flex-1 min-w-[200px] max-w-xs">
        <Search className="h-3.5 w-3.5 text-[#185ABD] shrink-0" />
        <input
          type="text"
          value={query}
          onChange={(e) => onChangeQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              if (e.shiftKey) {
                onFindPrev();
              } else {
                onFindNext();
              }
            } else if (e.key === 'Escape') {
              onClose();
            }
          }}
          placeholder="Tìm kiếm trong tài liệu..."
          className={`flex-1 min-w-0 h-8 px-2.5 rounded-lg border text-xs outline-none ${
            isDark
              ? 'bg-[#1E2534] border-white/15 text-white focus:border-blue-400'
              : 'bg-white border-slate-300 text-slate-900 focus:border-[#185ABD]'
          }`}
        />
        <span className="text-[11px] font-semibold tabular-nums opacity-75 shrink-0">
          {matchCount > 0 ? `${activeMatchIndex + 1}/${matchCount}` : '0/0'}
        </span>
      </div>

      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={onFindPrev}
          disabled={matchCount === 0}
          aria-label="Kết quả trước"
          title="Kết quả trước"
          className="h-8 w-8 rounded-lg border border-current/15 inline-flex items-center justify-center hover:bg-black/5 disabled:opacity-35"
        >
          <ChevronLeft className="h-3.5 w-3.5" />
        </button>
        <button
          type="button"
          onClick={onFindNext}
          disabled={matchCount === 0}
          aria-label="Kết quả tiếp theo"
          title="Kết quả tiếp theo"
          className="h-8 w-8 rounded-lg border border-current/15 inline-flex items-center justify-center hover:bg-black/5 disabled:opacity-35"
        >
          <ChevronRight className="h-3.5 w-3.5" />
        </button>
        <button
          type="button"
          onClick={onToggleCaseSensitive}
          aria-pressed={caseSensitive}
          title="Phân biệt chữ HOA / thường"
          className={`h-8 px-2 rounded-lg border text-[11px] font-bold transition-colors ${
            caseSensitive
              ? 'bg-[#185ABD] text-white border-[#185ABD]'
              : 'border-current/15 hover:bg-black/5'
          }`}
        >
          Aa
        </button>
      </div>

      <div className="flex items-center gap-1.5 flex-1 min-w-[200px] max-w-xs">
        <Replace className="h-3.5 w-3.5 text-[#185ABD] shrink-0" />
        <input
          type="text"
          value={replaceText}
          onChange={(e) => onChangeReplaceText(e.target.value)}
          placeholder="Thay thế bằng..."
          className={`flex-1 min-w-0 h-8 px-2.5 rounded-lg border text-xs outline-none ${
            isDark
              ? 'bg-[#1E2534] border-white/15 text-white focus:border-blue-400'
              : 'bg-white border-slate-300 text-slate-900 focus:border-[#185ABD]'
          }`}
        />
      </div>

      <div className="flex items-center gap-1.5">
        <button
          type="button"
          onClick={onReplaceOne}
          disabled={matchCount === 0}
          className="h-8 px-2.5 rounded-lg border border-current/15 text-xs font-medium hover:bg-black/5 disabled:opacity-40 whitespace-nowrap"
        >
          Thay thế
        </button>
        <button
          type="button"
          onClick={onReplaceAll}
          disabled={matchCount === 0}
          className="h-8 px-2.5 rounded-lg bg-[#185ABD] hover:bg-blue-700 text-white text-xs font-semibold disabled:opacity-40 whitespace-nowrap"
        >
          Thay tất cả
        </button>
        <button
          type="button"
          onClick={onClose}
          aria-label="Đóng thanh tìm kiếm"
          className="h-8 w-8 rounded-lg inline-flex items-center justify-center hover:bg-black/5"
        >
          <X className="h-4 w-4" />
        </button>
      </div>
    </div>
  );
}

export interface WordCommentsPanelProps {
  isDark: boolean;
  comments: WordCommentItem[];
  isReadOnly: boolean;
  selectedQuote: string;
  onAddComment: (text: string) => void;
  onDeleteComment: (id: string) => void;
  onJumpToQuote: (quote: string) => void;
  onClose: () => void;
}

export function WordCommentsPanel({
  isDark,
  comments,
  isReadOnly,
  selectedQuote,
  onAddComment,
  onDeleteComment,
  onJumpToQuote,
  onClose,
}: WordCommentsPanelProps) {
  const [draft, setDraft] = useState<string>('');

  const handleSubmit = () => {
    const trimmed = draft.trim();
    if (!trimmed) return;
    onAddComment(trimmed);
    setDraft('');
  };

  return (
    <aside
      aria-label="Bình luận lề tài liệu"
      className={`w-[290px] max-w-[88vw] shrink-0 border-l flex flex-col overflow-hidden z-20 ${
        isDark
          ? 'bg-[#161B26] border-white/10 text-slate-100'
          : 'bg-[#FAFBFD] border-[#E3E7EF] text-slate-900'
      }`}
    >
      <div
        className={`shrink-0 px-3.5 py-2.5 border-b flex items-center justify-between gap-2 ${
          isDark ? 'border-white/10' : 'border-[#E3E7EF]'
        }`}
      >
        <div className="flex items-center gap-2 min-w-0">
          <MessageSquarePlus className="h-4 w-4 text-[#185ABD] shrink-0" />
          <span className="text-xs font-bold truncate">
            Bình luận lề ({comments.length})
          </span>
        </div>
        <button
          type="button"
          onClick={onClose}
          aria-label="Đóng khung bình luận"
          className="h-7 w-7 rounded-lg inline-flex items-center justify-center hover:bg-black/5"
        >
          <X className="h-3.5 w-3.5" />
        </button>
      </div>

      {!isReadOnly && (
        <div
          className={`shrink-0 p-3 border-b space-y-2 ${
            isDark ? 'border-white/10' : 'border-[#E3E7EF]'
          }`}
        >
          {selectedQuote && (
            <div
              className={`px-2.5 py-1.5 rounded-lg border text-[11px] italic line-clamp-2 ${
                isDark
                  ? 'bg-amber-500/10 border-amber-400/30 text-amber-200'
                  : 'bg-amber-50 border-amber-200 text-amber-900'
              }`}
            >
              “{selectedQuote}”
            </div>
          )}
          <textarea
            rows={2}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            placeholder={
              selectedQuote
                ? 'Nhập bình luận cho đoạn đã chọn...'
                : 'Bôi đen chữ hoặc nhập bình luận chung...'
            }
            className={`w-full px-2.5 py-2 rounded-lg border text-xs resize-none outline-none ${
              isDark
                ? 'bg-[#1E2534] border-white/15 text-white focus:border-blue-400'
                : 'bg-white border-slate-300 text-slate-900 focus:border-[#185ABD]'
            }`}
          />
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!draft.trim()}
            className="w-full h-8 rounded-lg bg-[#185ABD] hover:bg-blue-700 text-white text-xs font-semibold disabled:opacity-40 transition-colors"
          >
            Thêm bình luận
          </button>
        </div>
      )}

      <div className="flex-1 min-h-0 overflow-y-auto p-3 space-y-2.5">
        {comments.length === 0 ? (
          <p className="text-xs opacity-65 text-center py-6">
            Chưa có bình luận nào. Bôi đen một đoạn văn bản rồi bấm Thêm bình luận.
          </p>
        ) : (
          comments.map((item) => (
            <div
              key={item.id}
              className={`rounded-xl border p-2.5 space-y-1.5 ${
                isDark
                  ? 'bg-[#1E2534] border-white/10'
                  : 'bg-white border-slate-200 shadow-xs'
              }`}
            >
              <div className="flex items-center justify-between gap-2">
                <span className="text-[11px] font-bold text-[#185ABD] truncate">
                  {item.author}
                </span>
                <div className="flex items-center gap-1 shrink-0">
                  <span className="text-[10px] opacity-60">{item.createdAt}</span>
                  {!isReadOnly && (
                    <button
                      type="button"
                      onClick={() => onDeleteComment(item.id)}
                      aria-label="Xóa bình luận"
                      className="p-1 rounded hover:bg-rose-500/15 text-rose-500"
                    >
                      <Trash2 className="h-3 w-3" />
                    </button>
                  )}
                </div>
              </div>
              {item.quote && (
                <button
                  type="button"
                  onClick={() => onJumpToQuote(item.quote)}
                  className={`w-full text-left px-2 py-1 rounded text-[11px] italic line-clamp-2 transition-colors ${
                    isDark
                      ? 'bg-white/5 hover:bg-white/10 text-slate-300'
                      : 'bg-slate-50 hover:bg-blue-50 text-slate-600'
                  }`}
                >
                  “{item.quote}”
                </button>
              )}
              <p className="text-xs leading-relaxed whitespace-pre-wrap">
                {item.text}
              </p>
            </div>
          ))
        )}
      </div>
    </aside>
  );
}

export type WordToolModalKind =
  | null
  | 'word_count'
  | 'hyperlink'
  | 'symbol_equation'
  | 'shape_textbox'
  | 'extract_summary';

export interface WordInsertToolModalsProps {
  activeModal: WordToolModalKind;
  isDark: boolean;
  stats: WordDocumentStats;
  extractedText: string;
  outlineTitles: string[];
  onClose: () => void;
  onInsertLink: (label: string, url: string) => void;
  onInsertHtml: (html: string) => void;
}

export function WordInsertToolModals({
  activeModal,
  isDark,
  stats,
  extractedText,
  outlineTitles,
  onClose,
  onInsertLink,
  onInsertHtml,
}: WordInsertToolModalsProps) {
  const [linkLabel, setLinkLabel] = useState<string>('');
  const [linkUrl, setLinkUrl] = useState<string>('https://');
  const [customEquation, setCustomEquation] = useState<string>('f(x) = a·x² + b·x + c');
  const [shapeKind, setShapeKind] = useState<WordShapeKind>('roundRect');
  const [shapeLabel, setShapeLabel] = useState<string>('Ý chính quan trọng');
  const [shapeFill, setShapeFill] = useState<string>('#DBEAFE');
  const [textBoxTitle, setTextBoxTitle] = useState<string>('Ghi chú nổi bật');
  const [textBoxBody, setTextBoxBody] = useState<string>(
    'Nhập nội dung điểm nhấn hoặc tóm tắt nhanh tại đây.'
  );
  const [copiedSummary, setCopiedSummary] = useState<boolean>(false);

  if (!activeModal) return null;

  const surfaceClass = isDark
    ? 'bg-[#1A202C] border-white/15 text-slate-100'
    : 'bg-white border-slate-200 text-slate-900';
  const inputClass = isDark
    ? 'bg-[#12161F] border-white/15 text-white focus:border-blue-400'
    : 'bg-slate-50 border-slate-300 text-slate-900 focus:border-[#185ABD]';

  const summaryBullets =
    outlineTitles.length > 0
      ? outlineTitles
      : extractedText
          .split(/[.!?]\s+/)
          .map((s) => s.trim())
          .filter((s) => s.length > 18)
          .slice(0, 6);

  return (
    <div
      role="dialog"
      aria-modal="true"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/55 backdrop-blur-xs flex items-center justify-center p-4"
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className={`w-full max-w-lg rounded-2xl border shadow-2xl overflow-hidden ${surfaceClass}`}
      >
        <div
          className={`px-4 py-3 border-b flex items-center justify-between ${
            isDark ? 'border-white/10' : 'border-slate-200'
          }`}
        >
          <div className="flex items-center gap-2 font-bold text-sm">
            {activeModal === 'word_count' && (
              <>
                <BookOpen className="h-4 w-4 text-[#185ABD]" />
                <span>Thống kê chi tiết tài liệu (Word Count)</span>
              </>
            )}
            {activeModal === 'hyperlink' && (
              <>
                <Link2 className="h-4 w-4 text-[#185ABD]" />
                <span>Chèn liên kết (Hyperlink)</span>
              </>
            )}
            {activeModal === 'symbol_equation' && (
              <>
                <Sigma className="h-4 w-4 text-[#185ABD]" />
                <span>Ký tự đặc biệt &amp; Công thức Toán</span>
              </>
            )}
            {activeModal === 'shape_textbox' && (
              <>
                <Shapes className="h-4 w-4 text-[#185ABD]" />
                <span>Chèn Hình khối &amp; Hộp văn bản</span>
              </>
            )}
            {activeModal === 'extract_summary' && (
              <>
                <FileSearch className="h-4 w-4 text-[#185ABD]" />
                <span>Trích xuất văn bản &amp; Tóm tắt nội dung</span>
              </>
            )}
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng hộp thoại"
            className="h-8 w-8 rounded-lg inline-flex items-center justify-center hover:bg-black/5"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="p-4 max-h-[75vh] overflow-y-auto space-y-4">
          {activeModal === 'word_count' && (
            <div className="grid grid-cols-2 gap-2.5 text-xs">
              {[
                { label: 'Số trang ước tính', value: stats.pages },
                { label: 'Tổng số từ', value: stats.words },
                { label: 'Ký tự (không dấu cách)', value: stats.charsWithoutSpaces },
                { label: 'Ký tự (có dấu cách)', value: stats.charsWithSpaces },
                { label: 'Số đoạn văn', value: stats.paragraphs },
                { label: 'Số dòng ước tính', value: stats.lines },
                { label: 'Thời gian đọc', value: `${stats.readingMinutes} phút` },
              ].map((row) => (
                <div
                  key={row.label}
                  className={`p-3 rounded-xl border flex items-center justify-between ${
                    isDark
                      ? 'bg-[#12161F] border-white/10'
                      : 'bg-slate-50 border-slate-200'
                  }`}
                >
                  <span className="opacity-75">{row.label}</span>
                  <span className="font-bold text-sm tabular-nums text-[#185ABD]">
                    {typeof row.value === 'number'
                      ? row.value.toLocaleString('vi-VN')
                      : row.value}
                  </span>
                </div>
              ))}
            </div>
          )}

          {activeModal === 'hyperlink' && (
            <div className="space-y-3 text-xs">
              <label className="block space-y-1">
                <span className="font-semibold">Văn bản hiển thị</span>
                <input
                  type="text"
                  value={linkLabel}
                  onChange={(e) => setLinkLabel(e.target.value)}
                  placeholder="Ví dụ: Tài liệu tham khảo"
                  className={`w-full h-9 px-3 rounded-xl border outline-none ${inputClass}`}
                />
              </label>
              <label className="block space-y-1">
                <span className="font-semibold">Địa chỉ URL</span>
                <input
                  type="url"
                  value={linkUrl}
                  onChange={(e) => setLinkUrl(e.target.value)}
                  placeholder="https://example.com"
                  className={`w-full h-9 px-3 rounded-xl border outline-none ${inputClass}`}
                />
              </label>
              <div className="flex justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={onClose}
                  className="h-9 px-3.5 rounded-xl border border-current/15 font-medium"
                >
                  Hủy
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const cleanUrl = linkUrl.trim();
                    if (!cleanUrl) return;
                    onInsertLink(linkLabel.trim() || cleanUrl, cleanUrl);
                    onClose();
                  }}
                  className="h-9 px-4 rounded-xl bg-[#185ABD] hover:bg-blue-700 text-white font-semibold"
                >
                  Chèn liên kết
                </button>
              </div>
            </div>
          )}

          {activeModal === 'symbol_equation' && (
            <div className="space-y-4 text-xs">
              <div className="space-y-2">
                <p className="font-bold text-[#185ABD]">Ký tự đặc biệt (bấm để chèn)</p>
                <div className="grid grid-cols-9 gap-1.5">
                  {WORD_SYMBOLS.map((sym) => (
                    <button
                      key={sym}
                      type="button"
                      onClick={() => {
                        onInsertHtml(`${sym} `);
                        onClose();
                      }}
                      className={`h-8 rounded-lg border font-semibold text-sm hover:border-[#185ABD] hover:text-[#185ABD] transition-colors ${
                        isDark
                          ? 'bg-[#12161F] border-white/10'
                          : 'bg-slate-50 border-slate-200'
                      }`}
                    >
                      {sym}
                    </button>
                  ))}
                </div>
              </div>

              <div className="space-y-2 pt-2 border-t border-current/10">
                <p className="font-bold text-[#185ABD]">Công thức Toán mẫu</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {WORD_EQUATION_PRESETS.map((eq) => (
                    <button
                      key={eq.label}
                      type="button"
                      onClick={() => {
                        onInsertHtml(buildEquationHtml(eq.preview));
                        onClose();
                      }}
                      className={`p-2 rounded-xl border text-left hover:border-[#185ABD] transition-colors ${
                        isDark
                          ? 'bg-[#12161F] border-white/10'
                          : 'bg-slate-50 border-slate-200'
                      }`}
                    >
                      <div className="text-[11px] opacity-70">{eq.label}</div>
                      <div className="font-serif italic font-semibold truncate">
                        {eq.preview}
                      </div>
                    </button>
                  ))}
                </div>

                <div className="flex items-center gap-2 pt-1">
                  <input
                    type="text"
                    value={customEquation}
                    onChange={(e) => setCustomEquation(e.target.value)}
                    placeholder="Nhập công thức tùy chỉnh..."
                    className={`flex-1 h-9 px-3 rounded-xl border outline-none ${inputClass}`}
                  />
                  <button
                    type="button"
                    onClick={() => {
                      if (!customEquation.trim()) return;
                      onInsertHtml(buildEquationHtml(customEquation));
                      onClose();
                    }}
                    className="h-9 px-3.5 rounded-xl bg-[#185ABD] text-white font-semibold shrink-0"
                  >
                    Chèn công thức
                  </button>
                </div>
              </div>
            </div>
          )}

          {activeModal === 'shape_textbox' && (
            <div className="space-y-4 text-xs">
              <div className="space-y-2">
                <p className="font-bold text-[#185ABD] flex items-center gap-1.5">
                  <Shapes className="h-3.5 w-3.5" />
                  <span>Hình khối (Shapes)</span>
                </p>
                <div className="grid grid-cols-5 gap-1.5">
                  {(
                    [
                      { k: 'rect', label: 'Chữ nhật' },
                      { k: 'roundRect', label: 'Bo góc' },
                      { k: 'circle', label: 'Elip' },
                      { k: 'arrow', label: 'Mũi tên' },
                      { k: 'badge', label: 'Ngôi sao' },
                    ] as const
                  ).map((item) => (
                    <button
                      key={item.k}
                      type="button"
                      onClick={() => setShapeKind(item.k)}
                      className={`py-2 px-1.5 rounded-xl border font-semibold transition-colors ${
                        shapeKind === item.k
                          ? 'bg-[#185ABD] text-white border-[#185ABD]'
                          : isDark
                            ? 'bg-[#12161F] border-white/10'
                            : 'bg-slate-50 border-slate-200'
                      }`}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    value={shapeLabel}
                    onChange={(e) => setShapeLabel(e.target.value)}
                    placeholder="Chữ bên trong hình khối..."
                    className={`flex-1 h-9 px-3 rounded-xl border outline-none ${inputClass}`}
                  />
                  <input
                    type="color"
                    value={shapeFill}
                    onChange={(e) => setShapeFill(e.target.value)}
                    title="Màu nền hình khối"
                    className="h-9 w-11 rounded-lg border border-current/15 cursor-pointer bg-transparent p-0.5"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      onInsertHtml(
                        buildShapeBlockHtml(shapeKind, shapeLabel, shapeFill, '#185ABD')
                      );
                      onClose();
                    }}
                    className="h-9 px-3 rounded-xl bg-[#185ABD] text-white font-semibold shrink-0"
                  >
                    Chèn hình
                  </button>
                </div>
              </div>

              <div className="space-y-2 pt-3 border-t border-current/10">
                <p className="font-bold text-[#185ABD] flex items-center gap-1.5">
                  <Frame className="h-3.5 w-3.5" />
                  <span>Hộp văn bản (Text Box)</span>
                </p>
                <input
                  type="text"
                  value={textBoxTitle}
                  onChange={(e) => setTextBoxTitle(e.target.value)}
                  placeholder="Tiêu đề hộp văn bản..."
                  className={`w-full h-9 px-3 rounded-xl border outline-none ${inputClass}`}
                />
                <textarea
                  rows={2}
                  value={textBoxBody}
                  onChange={(e) => setTextBoxBody(e.target.value)}
                  className={`w-full p-2.5 rounded-xl border outline-none resize-none ${inputClass}`}
                />
                <div className="flex justify-end">
                  <button
                    type="button"
                    onClick={() => {
                      onInsertHtml(
                        buildTextBoxHtml(textBoxTitle, textBoxBody, '#185ABD')
                      );
                      onClose();
                    }}
                    className="h-9 px-4 rounded-xl bg-[#185ABD] text-white font-semibold"
                  >
                    Chèn Hộp văn bản
                  </button>
                </div>
              </div>
            </div>
          )}

          {activeModal === 'extract_summary' && (
            <div className="space-y-3 text-xs">
              <div
                className={`p-3 rounded-xl border space-y-1.5 ${
                  isDark
                    ? 'bg-[#12161F] border-white/10'
                    : 'bg-blue-50/70 border-blue-200'
                }`}
              >
                <div className="flex items-center gap-1.5 font-bold text-[#185ABD]">
                  <Sparkles className="h-3.5 w-3.5" />
                  <span>Tóm tắt nhanh cấu trúc ({stats.words.toLocaleString('vi-VN')} từ)</span>
                </div>
                {summaryBullets.length > 0 ? (
                  <ul className="list-disc pl-4 space-y-1">
                    {summaryBullets.map((item, idx) => (
                      <li key={`sum-${idx}`}>{item}</li>
                    ))}
                  </ul>
                ) : (
                  <p className="opacity-75">Tài liệu chưa có đủ nội dung để tóm tắt.</p>
                )}
              </div>

              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="font-bold">Văn bản thuần trích xuất</span>
                  <button
                    type="button"
                    onClick={() => {
                      void navigator.clipboard?.writeText(extractedText).then(() => {
                        setCopiedSummary(true);
                        window.setTimeout(() => setCopiedSummary(false), 1800);
                      });
                    }}
                    className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[#185ABD] text-white font-semibold"
                  >
                    {copiedSummary ? (
                      <>
                        <Check className="h-3.5 w-3.5" />
                        <span>Đã sao chép</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-3.5 w-3.5" />
                        <span>Sao chép toàn bộ</span>
                      </>
                    )}
                  </button>
                </div>
                <textarea
                  readOnly
                  rows={7}
                  value={extractedText}
                  className={`w-full p-2.5 rounded-xl border outline-none resize-none ${inputClass}`}
                />
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export interface WordTableContextBarProps {
  isDark: boolean;
  onAddRow: () => void;
  onAddCol: () => void;
  onDeleteRow: () => void;
  onDeleteCol: () => void;
  onMergeRight: () => void;
  onSplitCell: () => void;
  onCellFill: (color: string) => void;
  onToggleBorders: () => void;
  onDeleteTable: () => void;
}

export function WordTableContextBar({
  isDark,
  onAddRow,
  onAddCol,
  onDeleteRow,
  onDeleteCol,
  onMergeRight,
  onSplitCell,
  onCellFill,
  onToggleBorders,
  onDeleteTable,
}: WordTableContextBarProps) {
  const btnCls = `h-7 px-2 rounded-lg border text-[11px] font-medium whitespace-nowrap transition-colors ${
    isDark
      ? 'bg-[#1E2534] border-white/15 text-slate-200 hover:bg-white/10'
      : 'bg-white border-slate-300 text-slate-800 hover:bg-blue-50 hover:border-[#185ABD]'
  }`;

  return (
    <div
      role="toolbar"
      aria-label="Công cụ chỉnh sửa bảng"
      className={`shrink-0 border-b px-3 py-1.5 flex items-center gap-1.5 overflow-x-auto no-scrollbar z-10 ${
        isDark
          ? 'bg-[#161E2E] border-blue-400/25 text-slate-100'
          : 'bg-[#EFF4FC] border-[#C9DAF5] text-slate-900'
      }`}
    >
      <span className="text-[11px] font-bold text-[#185ABD] shrink-0 mr-1 flex items-center gap-1">
        <ArrowLeftRight className="h-3.5 w-3.5" />
        Công cụ Bảng:
      </span>
      <button type="button" onClick={onAddRow} className={btnCls}>
        + Hàng
      </button>
      <button type="button" onClick={onAddCol} className={btnCls}>
        + Cột
      </button>
      <button type="button" onClick={onDeleteRow} className={btnCls}>
        − Hàng
      </button>
      <button type="button" onClick={onDeleteCol} className={btnCls}>
        − Cột
      </button>
      <button type="button" onClick={onMergeRight} className={btnCls}>
        Gộp ô phải
      </button>
      <button type="button" onClick={onSplitCell} className={btnCls}>
        Tách ô
      </button>
      <button type="button" onClick={onToggleBorders} className={btnCls}>
        Kẻ viền
      </button>
      <div className="inline-flex items-center gap-1 px-1.5 shrink-0">
        {['transparent', '#DBEAFE', '#FEF3C7', '#DCFCE7', '#FCE7F3'].map((bg) => (
          <button
            key={bg}
            type="button"
            onClick={() => onCellFill(bg)}
            title={bg === 'transparent' ? 'Xóa màu ô' : `Tô nền ô ${bg}`}
            className="h-5 w-5 rounded border border-slate-400/60"
            style={{
              backgroundColor: bg === 'transparent' ? '#FFFFFF' : bg,
            }}
          />
        ))}
      </div>
      <button
        type="button"
        onClick={onDeleteTable}
        className="h-7 px-2 rounded-lg border border-rose-300 bg-rose-50 text-rose-700 text-[11px] font-semibold whitespace-nowrap hover:bg-rose-100"
      >
        Xóa bảng
      </button>
    </div>
  );
}

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  ChangeEvent,
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
} from 'react';
import {
  AArrowDown,
  AArrowUp,
  AlertTriangle,
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDown,
  ArrowDownToLine,
  ArrowRightToLine,
  ArrowUp,
  BarChart3,
  Bold,
  Calculator,
  CaseSensitive,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  ClipboardPaste,
  Columns3,
  Copy,
  Download,
  Eraser,
  ExternalLink,
  Eye,
  EyeOff,
  FileDown,
  FileSpreadsheet,
  FileText,
  Filter,
  Grid3x3,
  Highlighter,
  History,
  ImagePlus,
  Info,
  Italic,
  LineChart,
  Link2,
  ListChecks,
  ListFilter,
  Loader2,
  Lock,
  Maximize2,
  Merge,
  MessageSquarePlus,
  MousePointer2,
  Paintbrush,
  PaintBucket,
  Palette,
  PenTool,
  Pencil,
  Percent,
  PieChart,
  Plus,
  Printer,
  RefreshCw,
  RotateCcw,
  Rows3,
  Save,
  Scissors,
  Search,
  Shapes,
  Share2,
  Shield,
  ShieldCheck,
  ShieldOff,
  Sigma,
  Snowflake,
  Sparkles,
  Split,
  Strikethrough,
  Table,
  Table2,
  Trash2,
  Type,
  Underline,
  Unlock,
  Users,
  Wand2,
  WrapText,
  X,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import type {
  CellStyle,
  DocumentItem,
  ExcelChartItem,
  ExcelConditionalFormat,
  ExcelDrawingStroke,
  ExcelFloatingObject,
  ExcelPageSetup,
  ModalType,
  SpreadsheetMergeRange,
  SpreadsheetSheet,
} from '../../types/document';
import { useWorkspaceSync } from '../../contexts/WorkspaceSyncContext';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import {
  columnIndexToLetter,
  evaluateSheetMatrix,
  formatCellRef,
  letterToColumnIndex,
} from '../../utils/formulaEvaluator';
import { exportExcelToXlsx } from '../../utils/fileExporter';
import { downloadDocumentAsPdf } from '../../utils/quickConverter';
import { saveFileUniversal } from '../../utils/saveFileUniversal';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { useDocumentPinchZoom } from '../../hooks/useDocumentPinchZoom';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { OfficeStatusBar } from './OfficeStatusBar';
import { WorkspaceTopBar } from './WorkspaceTopBar';
import { AIToolsDrawer } from '../ai/AIToolsDrawer';
import {
  adjustNumFmtDecimals,
  computeAutoFillValue,
  ConditionalFormatModal,
  DataValidationModal,
  evaluateConditionalFormats,
  ExcelChartsPanel,
  InsertChartModal,
  TextToColumnsModal,
} from './ExcelDataChartModal';
import {
  ExcelAiToolsModal,
  ExcelCellMetaModal,
  ExcelDrawingOverlay,
  ExcelFloatingObjectsLayer,
  ExcelFormulaLibraryModal,
  ONLYOFFICE_FORMULA_LIBRARY,
} from './ExcelOnlyOfficePanels';
import type {
  ExcelAiToolTab,
  ExcelCellMetaMode,
  ExcelDrawingTool,
  FormulaLibraryCategory,
} from './ExcelOnlyOfficePanels';

export interface ExcelWorkspaceProps {
  document: DocumentItem;
  onBackHome: () => void;
  onOpenModal: (
    modal: Exclude<ModalType, null>,
    overrideDoc?: DocumentItem
  ) => void;
  onSaveDocument: (updatedDoc: DocumentItem) => void;
  isModalOpen?: boolean;
  autoSaveEnabled?: boolean;
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

type RibbonTabId =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'layout'
  | 'formulas'
  | 'data'
  | 'collab'
  | 'protect'
  | 'view'
  | 'ai';

type OpenPopoverId =
  | 'fillColor'
  | 'textColor'
  | 'borders'
  | 'autoSum'
  | 'freeze'
  | 'case'
  | 'textDir'
  | 'fill'
  | 'clear'
  | 'shapes'
  | 'customSort'
  | 'commentsList'
  | null;

const ONLYOFFICE_RIBBON_TABS: ReadonlyArray<{
  id: RibbonTabId;
  label: string;
}> = [
  { id: 'file', label: 'Tệp' },
  { id: 'home', label: 'Trang chủ' },
  { id: 'insert', label: 'Chèn' },
  { id: 'draw', label: 'Vẽ' },
  { id: 'layout', label: 'Bố cục' },
  { id: 'formulas', label: 'Công thức' },
  { id: 'data', label: 'Dữ liệu' },
  { id: 'collab', label: 'Cộng tác' },
  { id: 'protect', label: 'Bảo vệ' },
  { id: 'view', label: 'Xem' },
  { id: 'ai', label: 'AI Tools' },
];

const FONT_FAMILY_OPTIONS: ReadonlyArray<{ label: string; value: string }> = [
  { label: 'Mặc định', value: '' },
  { label: 'Arial', value: 'Arial, sans-serif' },
  { label: 'Calibri', value: 'Calibri, sans-serif' },
  { label: 'Times New Roman', value: '"Times New Roman", serif' },
  { label: 'Georgia', value: 'Georgia, serif' },
  { label: 'Verdana', value: 'Verdana, sans-serif' },
  { label: 'Courier New', value: '"Courier New", monospace' },
];

const DRAW_COLOR_SWATCHES = [
  '#107C41',
  '#DC2626',
  '#2563EB',
  '#D97706',
  '#0F172A',
  '#7C3AED',
  '#FACC15',
] as const;

const DEFAULT_PAGE_SETUP: ExcelPageSetup = {
  margins: 'normal',
  orientation: 'portrait',
  paperSize: 'A4',
  printGridlines: true,
};

function isSafeHttpUrl(url: string | undefined): boolean {
  if (!url) return false;
  const trimmed = url.trim().toLowerCase();
  return trimmed.startsWith('https://') || trimmed.startsWith('http://');
}

function isCellInLockedRanges(
  rowIdx: number,
  colIdx: number,
  lockedRanges: string[] | undefined
): boolean {
  if (!lockedRanges || lockedRanges.length === 0) return false;
  for (const rangeStr of lockedRanges) {
    const parsed = parseCellOrRangeInput(rangeStr);
    if (!parsed) continue;
    const minR = Math.min(parsed.startRow, parsed.endRow);
    const maxR = Math.max(parsed.startRow, parsed.endRow);
    const minC = Math.min(parsed.startCol, parsed.endCol);
    const maxC = Math.max(parsed.startCol, parsed.endCol);
    if (rowIdx >= minR && rowIdx <= maxR && colIdx >= minC && colIdx <= maxC) {
      return true;
    }
  }
  return false;
}

interface CellCoord {
  row: number;
  col: number;
}

interface CellSelection {
  start: CellCoord;
  end: CellCoord;
}

interface NormalizedRange {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
  anchorRow: number;
  anchorCol: number;
}

interface SortConfig {
  colIndex: number;
  direction: 'asc' | 'desc';
}

interface IndexedRow {
  originalIndex: number;
  rawCells: (string | number)[];
  evaluatedCells: (string | number)[];
}

interface EditingCellState {
  originalRow: number;
  col: number;
  source: 'cell' | 'formula';
}

interface FillDragState {
  source: NormalizedRange;
  targetEnd: CellCoord;
  axis: 'vertical' | 'horizontal' | null;
}

interface ContextMenuState {
  x: number;
  y: number;
  targetType: 'cell' | 'row' | 'col';
  origRow: number;
  visRow: number;
  colIdx: number;
}

interface SheetTabMenuState {
  sheetIdx: number;
  x: number;
  y: number;
}

interface ColumnFilterPopoverState {
  colIdx: number;
  x: number;
  y: number;
}

const NUMBER_FORMATTER = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 2,
});

const DEFAULT_NEW_SHEET_ROWS = 12;
const DEFAULT_NEW_SHEET_COLS = 7;
const DEFAULT_COL_WIDTH = 136;
const DEFAULT_ROW_HEIGHT = 36;
const BASE_HEADER_HEIGHT = 32;
const BASE_ROW_HEADER_WIDTH = 48;

const FILL_COLOR_SWATCHES = [
  '#FFFFFF',
  '#F8FAFC',
  '#FEF9C3',
  '#DCFCE7',
  '#DBEAFE',
  '#FCE7F3',
  '#FEE2E2',
  '#FFEDD5',
  '#E0E7FF',
  '#F3E8FF',
  '#CCFBF1',
  '#E2E8F0',
  '#FDE047',
  '#86EFAC',
  '#93C5FD',
  '#FCA5A5',
] as const;

const TEXT_COLOR_SWATCHES = [
  '#0F172A',
  '#475569',
  '#107C41',
  '#15803D',
  '#1D4ED8',
  '#B91C1C',
  '#B45309',
  '#7C3AED',
  '#0E7490',
  '#BE185D',
  '#FFFFFF',
] as const;

const SHEET_TAB_COLORS = [
  '#107C41',
  '#2563EB',
  '#D97706',
  '#DC2626',
  '#7C3AED',
  '#0891B2',
  '#DB2777',
  '#475569',
] as const;

const FONT_SIZE_OPTIONS = [10, 11, 12, 13, 14, 16, 18, 20, 24] as const;

const NUMBER_FORMAT_PRESETS: ReadonlyArray<{
  id: string;
  label: string;
  numFmt: string;
}> = [
  { id: 'general', label: 'Chung', numFmt: 'General' },
  { id: 'number', label: 'Số (1.234,56)', numFmt: '#,##0.00' },
  { id: 'currency-vnd', label: 'Tiền tệ (₫)', numFmt: '#,##0 "₫"' },
  { id: 'percent', label: 'Phần trăm (%)', numFmt: '0.0%' },
  { id: 'date', label: 'Ngày tháng (dd/mm/yyyy)', numFmt: 'dd/mm/yyyy' },
];

function makeCellKey(rowIdx: number, colIdx: number): string {
  return `${rowIdx}:${colIdx}`;
}

function createLetterColumns(count: number): string[] {
  const safeCount = Math.max(1, count);
  return Array.from({ length: safeCount }, (_, idx) =>
    columnIndexToLetter(idx)
  );
}

function createBlankSheet(id: string, name: string): SpreadsheetSheet {
  return {
    id,
    name,
    columns: createLetterColumns(DEFAULT_NEW_SHEET_COLS),
    rows: Array.from({ length: DEFAULT_NEW_SHEET_ROWS }, () =>
      Array.from({ length: DEFAULT_NEW_SHEET_COLS }, () => '')
    ),
    cellStyles: {},
  };
}

function isLetterColumnHeaders(columns: string[]): boolean {
  if (columns.length === 0) {
    return false;
  }
  return columns.every(
    (col, idx) => col.trim().toUpperCase() === columnIndexToLetter(idx)
  );
}

function normalizeSheetForEditor(
  sheet: SpreadsheetSheet,
  fallbackIdx: number
): SpreadsheetSheet {
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

  const rowShift = shouldPrependHeader ? 1 : 0;

  const nextStyles: Record<string, CellStyle> = {};
  if (sheet.cellStyles) {
    for (const [key, style] of Object.entries(sheet.cellStyles)) {
      const parts = key.split(':');
      if (parts.length === 2) {
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) {
          nextStyles[makeCellKey(r + rowShift, c)] = { ...style };
        }
      }
    }
  }

  const normalizedRows: (string | number)[][] = [];

  if (shouldPrependHeader) {
    const headerRow: (string | number)[] = Array.from(
      { length: colCount },
      (_, colIdx) => rawColumns[colIdx] ?? ''
    );
    normalizedRows.push(headerRow);
    for (let colIdx = 0; colIdx < colCount; colIdx += 1) {
      nextStyles[makeCellKey(0, colIdx)] = {
        ...nextStyles[makeCellKey(0, colIdx)],
        bold: true,
      };
    }
  }

  for (const row of rawRows) {
    const padded: (string | number)[] = Array.from(
      { length: colCount },
      (_, colIdx) => {
        const val = Array.isArray(row) ? row[colIdx] : '';
        return val === undefined || val === null ? '' : val;
      }
    );
    normalizedRows.push(padded);
  }

  if (normalizedRows.length === 0) {
    normalizedRows.push(Array.from({ length: colCount }, () => ''));
  }

  const normalizedMerges: SpreadsheetMergeRange[] | undefined =
    Array.isArray(sheet.merges) && sheet.merges.length > 0
      ? sheet.merges
          .map((m) => ({
            s: {
              r: Math.max(0, m.s.r + rowShift),
              c: Math.max(0, m.s.c),
            },
            e: {
              r: Math.max(0, m.e.r + rowShift),
              c: Math.max(0, m.e.c),
            },
          }))
          .filter(
            (m) =>
              m.e.r >= m.s.r &&
              m.e.c >= m.s.c &&
              m.s.r < normalizedRows.length &&
              m.s.c < colCount
          )
      : undefined;

  const normalizedColWidths: number[] | undefined =
    Array.isArray(sheet.colWidths) && sheet.colWidths.length > 0
      ? Array.from({ length: colCount }, (_, idx) => {
          const w = sheet.colWidths?.[idx];
          return typeof w === 'number' && Number.isFinite(w) && w > 24
            ? Math.round(w)
            : DEFAULT_COL_WIDTH;
        })
      : undefined;

  const normalizedRowHeights: number[] | undefined =
    Array.isArray(sheet.rowHeights) && sheet.rowHeights.length > 0
      ? shouldPrependHeader
        ? [DEFAULT_ROW_HEIGHT, ...sheet.rowHeights]
        : [...sheet.rowHeights]
      : undefined;

  let normalizedValidations: Record<string, string[]> | undefined;
  if (sheet.dataValidations) {
    const mapped: Record<string, string[]> = {};
    for (const [key, opts] of Object.entries(sheet.dataValidations)) {
      if (!Array.isArray(opts) || opts.length === 0) continue;
      const parts = key.split(':');
      if (parts.length === 2) {
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) {
          mapped[makeCellKey(r + rowShift, c)] = [...opts];
        }
      }
    }
    if (Object.keys(mapped).length > 0) {
      normalizedValidations = mapped;
    }
  }

  const normalizedConditionalFormats: ExcelConditionalFormat[] | undefined =
    Array.isArray(sheet.conditionalFormats) && sheet.conditionalFormats.length > 0
      ? sheet.conditionalFormats.map((cf) => ({
          ...cf,
          range: {
            minRow: Math.max(0, cf.range.minRow + rowShift),
            maxRow: Math.max(0, cf.range.maxRow + rowShift),
            minCol: Math.max(0, cf.range.minCol),
            maxCol: Math.max(0, cf.range.maxCol),
          },
        }))
      : undefined;

  const normalizedHiddenRows: number[] | undefined =
    Array.isArray(sheet.hiddenRows) && sheet.hiddenRows.length > 0
      ? sheet.hiddenRows
          .map((r) => r + rowShift)
          .filter((r) => r >= 0 && r < normalizedRows.length)
      : undefined;

  const normalizedHiddenCols: number[] | undefined =
    Array.isArray(sheet.hiddenCols) && sheet.hiddenCols.length > 0
      ? sheet.hiddenCols.filter((c) => c >= 0 && c < colCount)
      : undefined;

  const normalizedCharts: ExcelChartItem[] | undefined =
    Array.isArray(sheet.charts) && sheet.charts.length > 0
      ? sheet.charts.map((ch) => ({
          ...ch,
          categories: [...ch.categories],
          series: ch.series.map((s) => ({ ...s, values: [...s.values] })),
        }))
      : undefined;

  let normalizedComments: Record<string, string> | undefined;
  if (sheet.cellComments) {
    const mapped: Record<string, string> = {};
    for (const [key, text] of Object.entries(sheet.cellComments)) {
      if (typeof text !== 'string' || !text.trim()) continue;
      const parts = key.split(':');
      if (parts.length === 2) {
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) {
          mapped[makeCellKey(r + rowShift, c)] = text;
        }
      }
    }
    if (Object.keys(mapped).length > 0) {
      normalizedComments = mapped;
    }
  }

  return {
    id: sheet.id || `sheet-${fallbackIdx + 1}`,
    name: sheet.name || `Sheet ${fallbackIdx + 1}`,
    columns: createLetterColumns(colCount),
    rows: normalizedRows,
    cellStyles: nextStyles,
    merges: normalizedMerges,
    colWidths: normalizedColWidths,
    rowHeights: normalizedRowHeights,
    frozenRows: sheet.frozenRows,
    frozenCols: sheet.frozenCols,
    tabColor: sheet.tabColor,
    hiddenRows: normalizedHiddenRows,
    hiddenCols: normalizedHiddenCols,
    autoFilterEnabled: sheet.autoFilterEnabled,
    dataValidations: normalizedValidations,
    conditionalFormats: normalizedConditionalFormats,
    charts: normalizedCharts,
    cellComments: normalizedComments,
    drawings: Array.isArray(sheet.drawings)
      ? sheet.drawings.map((d) => ({
          ...d,
          points: d.points.map((p) => ({ ...p })),
        }))
      : undefined,
    floatingObjects: Array.isArray(sheet.floatingObjects)
      ? sheet.floatingObjects.map((o) => ({ ...o }))
      : undefined,
    pageSetup: sheet.pageSetup ? { ...sheet.pageSetup } : undefined,
    isProtected: sheet.isProtected,
    lockedRanges: Array.isArray(sheet.lockedRanges)
      ? [...sheet.lockedRanges]
      : undefined,
    showGridlines: sheet.showGridlines,
  };
}

function initializeSheets(document: DocumentItem): SpreadsheetSheet[] {
  if (document.sheets && document.sheets.length > 0) {
    return document.sheets.map((sheet, idx) =>
      normalizeSheetForEditor(sheet, idx)
    );
  }
  return [createBlankSheet(`${document.id}-sheet-1`, 'Sheet 1')];
}

function parseNumericValue(value: string | number | undefined): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  const withoutPercent = trimmed.endsWith('%')
    ? trimmed.slice(0, -1).trim()
    : trimmed;
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(withoutPercent)) {
    return null;
  }
  const parsed = Number(withoutPercent);
  return Number.isFinite(parsed) ? parsed : null;
}

function isFormulaText(value: string | number | undefined): boolean {
  return typeof value === 'string' && value.trim().startsWith('=');
}

function isFormulaError(value: string | number | undefined): boolean {
  return (
    typeof value === 'string' &&
    /^#(DIV\/0!|CYCLE!|ERROR!|REF!|NAME\?|VALUE!|N\/A)$/i.test(value.trim())
  );
}

function formatNumber(value: number): string {
  return NUMBER_FORMATTER.format(value);
}

function formatExcelSerialDate(serial: number, includeTime: boolean): string {
  const wholeDays = Math.floor(serial);
  const utcMs = Date.UTC(1899, 11, 30) + wholeDays * 86_400_000;
  const date = new Date(utcMs);
  if (Number.isNaN(date.getTime())) {
    return formatNumber(serial);
  }
  const dd = String(date.getUTCDate()).padStart(2, '0');
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = date.getUTCFullYear();
  if (!includeTime) {
    return `${dd}/${mm}/${yyyy}`;
  }
  const frac = Math.max(0, serial - wholeDays);
  const totalMinutes = Math.round(frac * 24 * 60);
  const hh = String(Math.floor(totalMinutes / 60) % 24).padStart(2, '0');
  const mi = String(totalMinutes % 60).padStart(2, '0');
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

function formatCellDisplay(
  value: string | number | undefined,
  numFmt?: string
): string {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value !== 'number') {
    return value;
  }
  if (!Number.isFinite(value)) {
    return '';
  }
  const fmt = numFmt?.trim();
  if (!fmt || fmt === 'General' || fmt === '@') {
    return formatNumber(value);
  }

  const firstSection = fmt.split(';')[0] ?? fmt;
  const currencyTagMatch = /\[\$([^\]-]+)(?:-[^\]]+)?\]/.exec(firstSection);
  const stripped = firstSection
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .trim();

  if (stripped.includes('%')) {
    const decMatch = /\.([0#]+)/.exec(stripped);
    const decimals = decMatch ? Math.min(4, decMatch[1].length) : 0;
    const formattedPct = new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value * 100);
    return `${formattedPct}%`;
  }

  if (
    /[dy]/i.test(stripped) &&
    !/[0#]/.test(stripped) &&
    value >= 1 &&
    value <= 2_958_465
  ) {
    const hasTime = /h|s/i.test(stripped);
    return formatExcelSerialDate(value, hasTime);
  }

  const currencySymbol =
    currencyTagMatch?.[1] ??
    (/([₫$€£¥])/.exec(firstSection)?.[1] ??
      (/\bVND\b/i.test(firstSection) ? 'VND' : null));

  const decMatch = /\.([0#]+)/.exec(stripped);
  const decimals = decMatch ? Math.min(4, decMatch[1].length) : 0;
  const hasExplicitNumericPattern = /[0#]/.test(stripped);

  if (currencySymbol) {
    const formattedNum = new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value);
    const symbolIdx = firstSection.indexOf(currencySymbol);
    const digitIdx = firstSection.search(/[0#]/);
    if (symbolIdx >= 0 && digitIdx >= 0 && symbolIdx < digitIdx) {
      return `${currencySymbol}${formattedNum}`;
    }
    return `${formattedNum} ${currencySymbol}`;
  }

  if (hasExplicitNumericPattern) {
    return new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value);
  }

  return formatNumber(value);
}

function detectFormatPresetValue(numFmt: string | undefined): string {
  const fmt = (numFmt ?? '').trim();
  if (!fmt || fmt === 'General' || fmt === '@') return 'General';
  if (fmt.includes('%')) return '0.0%';
  if (fmt.includes('₫') || /\bVND\b/i.test(fmt) || fmt.includes('$')) {
    return '#,##0 "₫"';
  }
  if (/[dy]/i.test(fmt) && !/[0#]/.test(fmt)) return 'dd/mm/yyyy';
  return '#,##0.00';
}

function normalizeCommittedValue(draft: string): string | number {
  if (draft === '') {
    return '';
  }
  const trimmed = draft.trim();
  if (trimmed.startsWith('=')) {
    return trimmed;
  }
  if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) {
    const parsed = Number(trimmed);
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return draft;
}

// --- Coordinate Remapping Helpers for Row & Column Inserts / Deletes ---

function remapRecordByRowInsert<T>(
  record: Record<string, T> | undefined,
  atRow: number
): Record<string, T> | undefined {
  if (!record) return undefined;
  const next: Record<string, T> = {};
  for (const [key, val] of Object.entries(record)) {
    const [rStr, cStr] = key.split(':');
    const r = Number(rStr);
    const c = Number(cStr);
    if (!Number.isInteger(r) || !Number.isInteger(c)) continue;
    next[makeCellKey(r >= atRow ? r + 1 : r, c)] = val;
  }
  return next;
}

function remapRecordByColInsert<T>(
  record: Record<string, T> | undefined,
  atCol: number
): Record<string, T> | undefined {
  if (!record) return undefined;
  const next: Record<string, T> = {};
  for (const [key, val] of Object.entries(record)) {
    const [rStr, cStr] = key.split(':');
    const r = Number(rStr);
    const c = Number(cStr);
    if (!Number.isInteger(r) || !Number.isInteger(c)) continue;
    next[makeCellKey(r, c >= atCol ? c + 1 : c)] = val;
  }
  return next;
}

function remapRecordAfterRowDelete<T>(
  record: Record<string, T> | undefined,
  deletedRowIndices: Set<number>
): Record<string, T> | undefined {
  if (!record) return undefined;
  const sortedDeleted = Array.from(deletedRowIndices).sort((a, b) => a - b);
  const next: Record<string, T> = {};
  for (const [key, val] of Object.entries(record)) {
    const [rStr, cStr] = key.split(':');
    const rowIdx = Number(rStr);
    const colIdx = Number(cStr);
    if (!Number.isInteger(rowIdx) || !Number.isInteger(colIdx)) continue;
    if (deletedRowIndices.has(rowIdx)) continue;
    let shift = 0;
    for (const delIdx of sortedDeleted) {
      if (delIdx < rowIdx) shift += 1;
      else break;
    }
    next[makeCellKey(rowIdx - shift, colIdx)] = val;
  }
  return next;
}

function remapRecordAfterColDelete<T>(
  record: Record<string, T> | undefined,
  deletedColIndices: Set<number>
): Record<string, T> | undefined {
  if (!record) return undefined;
  const sortedDeleted = Array.from(deletedColIndices).sort((a, b) => a - b);
  const next: Record<string, T> = {};
  for (const [key, val] of Object.entries(record)) {
    const [rStr, cStr] = key.split(':');
    const rowIdx = Number(rStr);
    const colIdx = Number(cStr);
    if (!Number.isInteger(rowIdx) || !Number.isInteger(colIdx)) continue;
    if (deletedColIndices.has(colIdx)) continue;
    let shift = 0;
    for (const delIdx of sortedDeleted) {
      if (delIdx < colIdx) shift += 1;
      else break;
    }
    next[makeCellKey(rowIdx, colIdx - shift)] = val;
  }
  return next;
}

function remapMergesAfterRowDelete(
  merges: SpreadsheetMergeRange[] | undefined,
  deletedRowIndices: Set<number>
): SpreadsheetMergeRange[] | undefined {
  if (!merges || merges.length === 0) return undefined;
  const sortedDeleted = Array.from(deletedRowIndices).sort((a, b) => a - b);
  const countDeletedBefore = (rowIdx: number): number => {
    let shift = 0;
    for (const delIdx of sortedDeleted) {
      if (delIdx < rowIdx) shift += 1;
      else break;
    }
    return shift;
  };
  const remapped: SpreadsheetMergeRange[] = [];
  for (const m of merges) {
    let survivingStart: number | null = null;
    let survivingEnd: number | null = null;
    for (let r = m.s.r; r <= m.e.r; r += 1) {
      if (!deletedRowIndices.has(r)) {
        const shifted = r - countDeletedBefore(r);
        if (survivingStart === null) survivingStart = shifted;
        survivingEnd = shifted;
      }
    }
    if (
      survivingStart !== null &&
      survivingEnd !== null &&
      (survivingEnd > survivingStart || m.e.c > m.s.c)
    ) {
      remapped.push({
        s: { r: survivingStart, c: m.s.c },
        e: { r: survivingEnd, c: m.e.c },
      });
    }
  }
  return remapped.length > 0 ? remapped : undefined;
}

function remapMergesAfterColDelete(
  merges: SpreadsheetMergeRange[] | undefined,
  deletedColIndices: Set<number>
): SpreadsheetMergeRange[] | undefined {
  if (!merges || merges.length === 0) return undefined;
  const sortedDeleted = Array.from(deletedColIndices).sort((a, b) => a - b);
  const countDeletedBefore = (colIdx: number): number => {
    let shift = 0;
    for (const delIdx of sortedDeleted) {
      if (delIdx < colIdx) shift += 1;
      else break;
    }
    return shift;
  };
  const remapped: SpreadsheetMergeRange[] = [];
  for (const m of merges) {
    let survivingStart: number | null = null;
    let survivingEnd: number | null = null;
    for (let c = m.s.c; c <= m.e.c; c += 1) {
      if (!deletedColIndices.has(c)) {
        const shifted = c - countDeletedBefore(c);
        if (survivingStart === null) survivingStart = shifted;
        survivingEnd = shifted;
      }
    }
    if (
      survivingStart !== null &&
      survivingEnd !== null &&
      (survivingEnd > survivingStart || m.e.r > m.s.r)
    ) {
      remapped.push({
        s: { r: m.s.r, c: survivingStart },
        e: { r: m.e.r, c: survivingEnd },
      });
    }
  }
  return remapped.length > 0 ? remapped : undefined;
}

function remapMergesAfterRowInsert(
  merges: SpreadsheetMergeRange[] | undefined,
  atRow: number
): SpreadsheetMergeRange[] | undefined {
  if (!merges || merges.length === 0) return undefined;
  return merges.map((m) => ({
    s: { r: m.s.r >= atRow ? m.s.r + 1 : m.s.r, c: m.s.c },
    e: { r: m.e.r >= atRow ? m.e.r + 1 : m.e.r, c: m.e.c },
  }));
}

function remapMergesAfterColInsert(
  merges: SpreadsheetMergeRange[] | undefined,
  atCol: number
): SpreadsheetMergeRange[] | undefined {
  if (!merges || merges.length === 0) return undefined;
  return merges.map((m) => ({
    s: { r: m.s.r, c: m.s.c >= atCol ? m.s.c + 1 : m.s.c },
    e: { r: m.e.r, c: m.e.c >= atCol ? m.e.c + 1 : m.e.c },
  }));
}

function remapIndicesAfterDelete(
  indices: number[] | undefined,
  deletedSet: Set<number>
): number[] | undefined {
  if (!indices || indices.length === 0) return undefined;
  const sortedDel = Array.from(deletedSet).sort((a, b) => a - b);
  const next: number[] = [];
  for (const idx of indices) {
    if (deletedSet.has(idx)) continue;
    let shift = 0;
    for (const d of sortedDel) {
      if (d < idx) shift += 1;
      else break;
    }
    next.push(idx - shift);
  }
  return next.length > 0 ? next : undefined;
}

function remapIndicesAfterInsert(
  indices: number[] | undefined,
  atIdx: number
): number[] | undefined {
  if (!indices || indices.length === 0) return undefined;
  return indices.map((idx) => (idx >= atIdx ? idx + 1 : idx));
}

function parseCellOrRangeInput(
  raw: string
): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
  const cleaned = raw.trim().toUpperCase().replace(/\$/g, '');
  if (!cleaned) return null;
  const parts = cleaned.split(':');
  if (parts.length > 2) return null;

  const parseOne = (token: string): { row: number; col: number } | null => {
    const m = /^([A-Z]{1,3})(\d+)$/.exec(token.trim());
    if (!m) return null;
    const col = letterToColumnIndex(m[1]);
    const rowNum = Number.parseInt(m[2], 10);
    if (col < 0 || !Number.isFinite(rowNum) || rowNum < 1) return null;
    return { row: rowNum - 1, col };
  };

  const p1 = parseOne(parts[0]);
  if (!p1) return null;
  if (parts.length === 1) {
    return {
      startRow: p1.row,
      startCol: p1.col,
      endRow: p1.row,
      endCol: p1.col,
    };
  }
  const p2 = parseOne(parts[1]);
  if (!p2) return null;
  return {
    startRow: p1.row,
    startCol: p1.col,
    endRow: p2.row,
    endCol: p2.col,
  };
}

export function ExcelWorkspace({
  document,
  onBackHome,
  onOpenModal,
  onSaveDocument,
  isModalOpen = false,
  autoSaveEnabled = false,
  collaborators,
  selfSessionId,
  collabPermission = 'owner',
  onPresenceUpdate,
  onBroadcastPatch,
  incomingRemotePatch = null,
}: ExcelWorkspaceProps) {
  const workspaceSync = useWorkspaceSync();
  const isReadOnly = collabPermission === 'view';
  const [viewMode, setViewMode] = useState<'original' | 'edit'>(() =>
    collabPermission === 'view' ? 'original' : 'edit'
  );
  const viewModeRef = useRef<'original' | 'edit'>(viewMode);
  viewModeRef.current = viewMode;
  const isEditingLocked = isReadOnly || viewMode === 'original';

  const [sheets, setSheets] = useState<SpreadsheetSheet[]>(() =>
    initializeSheets(document)
  );
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTabId>('home');
  const [isRibbonCollapsed, setIsRibbonCollapsed] = useState(false);
  const [openPopover, setOpenPopover] = useState<OpenPopoverId>(null);
  const [showFormulaBar, setShowFormulaBar] = useState(true);
  const [recalcVersion, setRecalcVersion] = useState(0);

  // ONLYOFFICE Draw, Floating Objects, Clipboard, Format Painter & Modals
  const [activeDrawTool, setActiveDrawTool] = useState<ExcelDrawingTool>(null);
  const [drawColor, setDrawColor] = useState<string>('#107C41');
  const [drawStrokeWidth, setDrawStrokeWidth] = useState<number>(3);
  const [selectedFloatingObjId, setSelectedFloatingObjId] = useState<
    string | null
  >(null);
  const [clipboardData, setClipboardData] = useState<{
    values: (string | number)[][];
    styles: (CellStyle | undefined)[][];
  } | null>(null);
  const [formatPainterStyle, setFormatPainterStyle] =
    useState<CellStyle | null>(null);
  const [formulaModalCategory, setFormulaModalCategory] =
    useState<FormulaLibraryCategory | null>(null);
  const [aiModalTab, setAiModalTab] = useState<ExcelAiToolTab | null>(null);
  const [cellMetaModal, setCellMetaModal] = useState<{
    mode: ExcelCellMetaMode;
    origRow: number;
    colIdx: number;
  } | null>(null);
  const [isFileInfoOpen, setIsFileInfoOpen] = useState(false);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [isDownloadingOriginal, setIsDownloadingOriginal] = useState(false);
  const [customSortColIdx, setCustomSortColIdx] = useState(0);
  const [customSortDir, setCustomSortDir] = useState<'asc' | 'desc'>('asc');
  const imageInputRef = useRef<HTMLInputElement | null>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [columnFilters, setColumnFilters] = useState<
    Record<number, Set<string>>
  >({});
  const [filterPopover, setFilterPopover] =
    useState<ColumnFilterPopoverState | null>(null);
  const [filterSearchText, setFilterSearchText] = useState('');

  const [selection, setSelection] = useState<CellSelection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 },
  });
  const [editingCell, setEditingCell] = useState<EditingCellState | null>(null);
  const [editDraft, setEditDraft] = useState('');

  const [isNameBoxFocused, setIsNameBoxFocused] = useState(false);
  const [nameBoxDraft, setNameBoxDraft] = useState('A1');

  const [zoomPercent, setZoomPercent] = useState<number>(100);
  const [fillDrag, setFillDrag] = useState<FillDragState | null>(null);
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null);
  const [sheetTabMenu, setSheetTabMenu] = useState<SheetTabMenuState | null>(
    null
  );
  const [renamingSheetIdx, setRenamingSheetIdx] = useState<number | null>(null);
  const [renameSheetDraft, setRenameSheetDraft] = useState('');
  const [draggingSheetIdx, setDraggingSheetIdx] = useState<number | null>(null);
  const [openValidationKey, setOpenValidationKey] = useState<string | null>(
    null
  );

  // Modals
  const [chartModalType, setChartModalType] = useState<
    ExcelChartItem['type'] | null
  >(null);
  const [isCondFormatModalOpen, setIsCondFormatModalOpen] = useState(false);
  const [isValidationModalOpen, setIsValidationModalOpen] = useState(false);
  const [isTextToColsModalOpen, setIsTextToColsModalOpen] = useState(false);

  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [justSaved, setJustSaved] = useState(false);
  const [feedbackBanner, setFeedbackBanner] = useState<string | null>(null);
  const [naturalCanvasSize, setNaturalCanvasSize] = useState<{
    width: number;
    height: number;
  }>({ width: 0, height: 0 });

  const sheetsRef = useRef<SpreadsheetSheet[]>(sheets);
  sheetsRef.current = sheets;
  const didUserEditRef = useRef<boolean>(Boolean(document.hasLocalEdits));
  const lastAppliedPatchAtRef = useRef<number>(0);
  const isDraggingRef = useRef(false);
  const fillDragRef = useRef<FillDragState | null>(null);
  fillDragRef.current = fillDrag;
  const cancelBlurCommitRef = useRef(false);
  const saveTimerRef = useRef<number | null>(null);
  const feedbackTimerRef = useRef<number | null>(null);
  const gridContainerRef = useRef<HTMLDivElement | null>(null);
  const sheetCanvasRef = useRef<HTMLDivElement | null>(null);
  const cellInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);

  // Pinch-to-zoom & Ctrl+Wheel on Grid Viewport ONLY (never scaling Ribbon/StatusBar)
  useDocumentPinchZoom(gridContainerRef, zoomPercent, setZoomPercent, {
    min: 10,
    max: 300,
    disabled:
      isModalOpen ||
      chartModalType !== null ||
      isCondFormatModalOpen ||
      isValidationModalOpen ||
      isTextToColsModalOpen ||
      formulaModalCategory !== null ||
      aiModalTab !== null ||
      cellMetaModal !== null ||
      isFileInfoOpen ||
      activeDrawTool !== null,
  });

  useEffect(() => {
    const initialized = initializeSheets(document);
    sheetsRef.current = initialized;
    didUserEditRef.current = Boolean(document.hasLocalEdits);
    lastAppliedPatchAtRef.current = 0;
    setSheets(initialized);
    setActiveSheetIdx(0);
    setSearchQuery('');
    setSortConfig(null);
    setColumnFilters({});
    setFilterPopover(null);
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
    setEditingCell(null);
    setEditDraft('');
    setHasUnsavedChanges(false);
    setJustSaved(false);
    setFeedbackBanner(null);
    setOpenPopover(null);
    setContextMenu(null);
    setSheetTabMenu(null);
    setRenamingSheetIdx(null);
    setOpenValidationKey(null);
    setZoomPercent(100);
    const initialMode: 'original' | 'edit' =
      collabPermission === 'view' ? 'original' : 'edit';
    viewModeRef.current = initialMode;
    setViewMode(initialMode);
  }, [document.id, collabPermission]);

  useEffect(() => {
    return () => {
      if (feedbackTimerRef.current !== null) {
        window.clearTimeout(feedbackTimerRef.current);
      }
      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
    };
  }, []);

  // Close floating popovers on Escape or outside mousedown
  useEffect(() => {
    const handleOutside = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (target?.closest('[data-excel-popover="true"]')) {
        return;
      }
      setOpenPopover(null);
      setContextMenu(null);
      setSheetTabMenu(null);
      setFilterPopover(null);
      setOpenValidationKey(null);
    };
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setOpenPopover(null);
        setContextMenu(null);
        setSheetTabMenu(null);
        setFilterPopover(null);
        setOpenValidationKey(null);
      }
    };
    window.addEventListener('mousedown', handleOutside);
    window.addEventListener('keydown', handleKey);
    return () => {
      window.removeEventListener('mousedown', handleOutside);
      window.removeEventListener('keydown', handleKey);
    };
  }, []);

  // Apply incoming realtime excel_cell_patch without disrupting local selection
  useEffect(() => {
    if (
      !incomingRemotePatch ||
      incomingRemotePatch.kind !== 'excel_cell_patch'
    ) {
      return;
    }
    if (
      selfSessionId &&
      incomingRemotePatch.senderSessionId === selfSessionId
    ) {
      return;
    }
    if (incomingRemotePatch.sentAt <= lastAppliedPatchAtRef.current) {
      return;
    }
    lastAppliedPatchAtRef.current = incomingRemotePatch.sentAt;

    setSheets((prev) => {
      let next: SpreadsheetSheet[];
      if (
        Array.isArray(incomingRemotePatch.sheets) &&
        incomingRemotePatch.sheets.length > 0
      ) {
        next = incomingRemotePatch.sheets.map((sheet, idx) =>
          normalizeSheetForEditor(sheet, idx)
        );
      } else {
        const targetIdx = prev.findIndex(
          (sheet) => sheet.id === incomingRemotePatch.sheetId
        );
        if (targetIdx === -1) {
          return prev;
        }
        next = prev.map((sheet, idx) => {
          if (idx !== targetIdx) {
            return sheet;
          }
          const { rowIndex, colIndex, value } = incomingRemotePatch;
          if (rowIndex < 0 || colIndex < 0) {
            return sheet;
          }
          const neededCols = Math.max(sheet.columns.length, colIndex + 1);
          const neededRows = Math.max(sheet.rows.length, rowIndex + 1);
          const nextRows = Array.from({ length: neededRows }, (_, r) => {
            const existingRow = sheet.rows[r] ?? [];
            const padded = Array.from(
              { length: neededCols },
              (__, c) => existingRow[c] ?? ''
            );
            if (r === rowIndex) {
              padded[colIndex] = value;
            }
            return padded;
          });
          return {
            ...sheet,
            columns: createLetterColumns(neededCols),
            rows: nextRows,
          };
        });
      }
      sheetsRef.current = next;
      return next;
    });
  }, [incomingRemotePatch, selfSessionId]);

  useEffect(() => {
    if (editingCell?.source === 'cell' && cellInputRef.current) {
      const input = cellInputRef.current;
      input.focus();
      const len = input.value.length;
      input.setSelectionRange(len, len);
    }
  }, [editingCell]);

  const safeSheetIdx = Math.min(activeSheetIdx, Math.max(sheets.length - 1, 0));
  const activeSheet = sheets[safeSheetIdx];
  const columnCount = activeSheet.columns.length;
  const zoomScale = zoomPercent / 100;
  const isSheetProtected = Boolean(activeSheet.isProtected);
  const isEditBlocked = isEditingLocked || isSheetProtected;
  const showGridlines = activeSheet.showGridlines !== false;
  const activePageSetup: ExcelPageSetup = {
    ...DEFAULT_PAGE_SETUP,
    ...(activeSheet.pageSetup ?? {}),
  };

  const hiddenRowSet = useMemo(
    () => new Set<number>(activeSheet.hiddenRows ?? []),
    [activeSheet.hiddenRows]
  );

  const hiddenColSet = useMemo(
    () => new Set<number>(activeSheet.hiddenCols ?? []),
    [activeSheet.hiddenCols]
  );

  const visibleColIndices = useMemo<number[]>(() => {
    const list: number[] = [];
    for (let c = 0; c < columnCount; c += 1) {
      if (!hiddenColSet.has(c)) {
        list.push(c);
      }
    }
    if (list.length === 0 && columnCount > 0) {
      list.push(0);
    }
    return list;
  }, [columnCount, hiddenColSet]);

  const evaluatedMatrix = useMemo<(string | number)[][]>(() => {
    void recalcVersion;
    return evaluateSheetMatrix(activeSheet.rows);
  }, [activeSheet.rows, recalcVersion]);

  const conditionalStyleMap = useMemo(
    () =>
      evaluateConditionalFormats(
        activeSheet.conditionalFormats,
        evaluatedMatrix
      ),
    [activeSheet.conditionalFormats, evaluatedMatrix]
  );

  const hasHeaderRow = useMemo(() => {
    if (activeSheet.rows.length <= 1) {
      return false;
    }
    const firstRow = activeSheet.rows[0];
    const styles = activeSheet.cellStyles ?? {};
    let nonEmptyFirst = 0;
    let boldFirst = 0;
    let numericFirst = 0;

    for (let c = 0; c < columnCount; c += 1) {
      const val = firstRow?.[c];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        nonEmptyFirst += 1;
        if (styles[makeCellKey(0, c)]?.bold) {
          boldFirst += 1;
        }
        if (parseNumericValue(evaluatedMatrix[0]?.[c]) !== null) {
          numericFirst += 1;
        }
      }
    }

    if (nonEmptyFirst === 0) return false;
    if (boldFirst === nonEmptyFirst) return true;
    if (numericFirst === 0) {
      for (let r = 1; r < evaluatedMatrix.length; r += 1) {
        for (let c = 0; c < columnCount; c += 1) {
          if (parseNumericValue(evaluatedMatrix[r]?.[c]) !== null) {
            return true;
          }
        }
      }
    }
    return false;
  }, [activeSheet.cellStyles, activeSheet.rows, columnCount, evaluatedMatrix]);

  const numericColumnIndices = useMemo<number[]>(() => {
    const startRow = hasHeaderRow ? 1 : 0;
    const indices: number[] = [];
    for (const colIdx of visibleColIndices) {
      let nonEmpty = 0;
      let numericCount = 0;
      for (let r = startRow; r < evaluatedMatrix.length; r += 1) {
        if (hiddenRowSet.has(r)) continue;
        const val = evaluatedMatrix[r]?.[colIdx];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          nonEmpty += 1;
          if (parseNumericValue(val) !== null) {
            numericCount += 1;
          }
        }
      }
      if (nonEmpty > 0 && numericCount / nonEmpty >= 0.5) {
        indices.push(colIdx);
      }
    }
    return indices;
  }, [evaluatedMatrix, hasHeaderRow, hiddenRowSet, visibleColIndices]);

  const activeColumnFilterEntries = useMemo(
    () =>
      Object.entries(columnFilters)
        .map(([k, set]) => [Number(k), set] as const)
        .filter(([colIdx, set]) => colIdx < columnCount && set.size > 0),
    [columnCount, columnFilters]
  );

  const hasColumnFilters = activeColumnFilterEntries.length > 0;
  const isFiltered = searchQuery.trim().length > 0 || hasColumnFilters;
  const isSorted = sortConfig !== null;
  const hasActiveTransforms = isFiltered || isSorted;

  const visibleRows = useMemo<IndexedRow[]>(() => {
    const normalizedQuery = searchQuery.trim().toLowerCase();
    const indexed: IndexedRow[] = [];
    for (let idx = 0; idx < activeSheet.rows.length; idx += 1) {
      if (hiddenRowSet.has(idx)) continue;
      indexed.push({
        originalIndex: idx,
        rawCells: activeSheet.rows[idx],
        evaluatedCells: evaluatedMatrix[idx] ?? activeSheet.rows[idx],
      });
    }

    const matchesQuery = (row: IndexedRow): boolean => {
      if (!normalizedQuery) return true;
      for (let colIdx = 0; colIdx < columnCount; colIdx += 1) {
        const raw = row.rawCells[colIdx];
        const evaluated = row.evaluatedCells[colIdx];
        if (
          raw !== undefined &&
          raw !== null &&
          String(raw).toLowerCase().includes(normalizedQuery)
        ) {
          return true;
        }
        if (
          evaluated !== undefined &&
          evaluated !== null &&
          String(evaluated).toLowerCase().includes(normalizedQuery)
        ) {
          return true;
        }
        const num = parseNumericValue(evaluated);
        if (
          num !== null &&
          formatNumber(num).toLowerCase().includes(normalizedQuery)
        ) {
          return true;
        }
      }
      return false;
    };

    const matchesColFilters = (row: IndexedRow): boolean => {
      for (const [colIdx, allowedSet] of activeColumnFilterEntries) {
        const style =
          activeSheet.cellStyles?.[makeCellKey(row.originalIndex, colIdx)];
        const label =
          formatCellDisplay(row.evaluatedCells[colIdx], style?.numFmt).trim() ||
          '(Trống)';
        if (!allowedSet.has(label)) {
          return false;
        }
      }
      return true;
    };

    let filtered: IndexedRow[];
    if (normalizedQuery || hasColumnFilters) {
      if (hasHeaderRow && indexed.length > 1 && indexed[0].originalIndex === 0) {
        const dataMatches = indexed
          .slice(1)
          .filter((r) => matchesQuery(r) && matchesColFilters(r));
        filtered = [indexed[0], ...dataMatches];
      } else {
        filtered = indexed.filter(
          (r) => matchesQuery(r) && matchesColFilters(r)
        );
      }
    } else {
      filtered = indexed;
    }

    if (!sortConfig || sortConfig.colIndex >= columnCount) {
      return filtered;
    }

    const { colIndex, direction } = sortConfig;
    const factor = direction === 'asc' ? 1 : -1;

    const pinnedHeader =
      hasHeaderRow && filtered.length > 1 && filtered[0].originalIndex === 0
        ? filtered[0]
        : null;
    const rowsToSort = pinnedHeader ? filtered.slice(1) : filtered;

    const sorted = [...rowsToSort].sort((rowA, rowB) => {
      const valA = rowA.evaluatedCells[colIndex];
      const valB = rowB.evaluatedCells[colIndex];
      const emptyA =
        valA === undefined || valA === null || String(valA).trim() === '';
      const emptyB =
        valB === undefined || valB === null || String(valB).trim() === '';

      if (emptyA && emptyB) return rowA.originalIndex - rowB.originalIndex;
      if (emptyA) return 1;
      if (emptyB) return -1;

      const numA = parseNumericValue(valA);
      const numB = parseNumericValue(valB);
      let diff = 0;

      if (numA !== null && numB !== null) {
        diff = numA - numB;
      } else {
        diff = String(valA).localeCompare(String(valB), 'vi', {
          numeric: true,
          sensitivity: 'base',
        });
      }

      if (diff !== 0) return diff * factor;
      return rowA.originalIndex - rowB.originalIndex;
    });

    return pinnedHeader ? [pinnedHeader, ...sorted] : sorted;
  }, [
    activeColumnFilterEntries,
    activeSheet.cellStyles,
    activeSheet.rows,
    columnCount,
    evaluatedMatrix,
    hasColumnFilters,
    hasHeaderRow,
    hiddenRowSet,
    searchQuery,
    sortConfig,
  ]);

  const normalizedRange = useMemo<NormalizedRange | null>(() => {
    if (visibleRows.length === 0 || columnCount === 0) {
      return null;
    }
    const maxRowIdx = visibleRows.length - 1;
    const maxColIdx = columnCount - 1;

    const r1 = Math.max(0, Math.min(selection.start.row, maxRowIdx));
    const c1 = Math.max(0, Math.min(selection.start.col, maxColIdx));
    const r2 = Math.max(0, Math.min(selection.end.row, maxRowIdx));
    const c2 = Math.max(0, Math.min(selection.end.col, maxColIdx));

    return {
      minRow: Math.min(r1, r2),
      maxRow: Math.max(r1, r2),
      minCol: Math.min(c1, c2),
      maxCol: Math.max(c1, c2),
      anchorRow: r1,
      anchorCol: c1,
    };
  }, [columnCount, selection, visibleRows.length]);

  const anchorOriginalRow =
    normalizedRange !== null
      ? (visibleRows[normalizedRange.anchorRow]?.originalIndex ?? null)
      : null;
  const anchorCol = normalizedRange !== null ? normalizedRange.anchorCol : null;

  const anchorRawValue =
    anchorOriginalRow !== null && anchorCol !== null
      ? activeSheet.rows[anchorOriginalRow]?.[anchorCol]
      : undefined;

  const anchorEvaluatedValue =
    anchorOriginalRow !== null && anchorCol !== null
      ? evaluatedMatrix[anchorOriginalRow]?.[anchorCol]
      : undefined;

  const anchorStyle: CellStyle =
    anchorOriginalRow !== null && anchorCol !== null
      ? (activeSheet.cellStyles?.[makeCellKey(anchorOriginalRow, anchorCol)] ??
        {})
      : {};

  const anchorEffectiveAlign: 'left' | 'center' | 'right' | 'justify' =
    anchorStyle.align ??
    (parseNumericValue(anchorEvaluatedValue) !== null ? 'right' : 'left');
  const anchorEffectiveValign: 'top' | 'middle' | 'bottom' =
    anchorStyle.valign ?? 'middle';

  const rangeLabel = useMemo(() => {
    if (!normalizedRange) {
      return '—';
    }
    const startOrigRow =
      visibleRows[normalizedRange.minRow]?.originalIndex ??
      normalizedRange.minRow;
    const endOrigRow =
      visibleRows[normalizedRange.maxRow]?.originalIndex ??
      normalizedRange.maxRow;

    const startRef = formatCellRef(startOrigRow, normalizedRange.minCol);
    if (
      normalizedRange.minRow === normalizedRange.maxRow &&
      normalizedRange.minCol === normalizedRange.maxCol
    ) {
      return startRef;
    }
    const endRef = formatCellRef(endOrigRow, normalizedRange.maxCol);
    return `${startRef}:${endRef}`;
  }, [normalizedRange, visibleRows]);

  useEffect(() => {
    if (!isNameBoxFocused) {
      setNameBoxDraft(rangeLabel);
    }
  }, [isNameBoxFocused, rangeLabel]);

  const selectionOriginalBounds = useMemo(() => {
    if (!normalizedRange) {
      return { minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 };
    }
    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = 0;
    for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
      const orig = visibleRows[r]?.originalIndex;
      if (orig !== undefined) {
        minRow = Math.min(minRow, orig);
        maxRow = Math.max(maxRow, orig);
      }
    }
    if (!Number.isFinite(minRow)) minRow = 0;
    return {
      minRow,
      maxRow,
      minCol: normalizedRange.minCol,
      maxCol: normalizedRange.maxCol,
    };
  }, [normalizedRange, visibleRows]);

  const selectionStats = useMemo(() => {
    if (!normalizedRange) {
      return {
        totalCells: 0,
        nonEmptyCount: 0,
        numericCount: 0,
        sum: 0,
        avg: null as number | null,
      };
    }

    let totalCells = 0;
    let nonEmptyCount = 0;
    let numericCount = 0;
    let sum = 0;

    for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
      const rowData = visibleRows[r];
      if (!rowData) continue;
      for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
        if (hiddenColSet.has(c)) continue;
        totalCells += 1;
        const val = rowData.evaluatedCells[c];
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          nonEmptyCount += 1;
        }
        const parsed = parseNumericValue(val);
        if (parsed !== null) {
          sum += parsed;
          numericCount += 1;
        }
      }
    }

    return {
      totalCells,
      nonEmptyCount,
      numericCount,
      sum,
      avg: numericCount > 0 ? sum / numericCount : null,
    };
  }, [hiddenColSet, normalizedRange, visibleRows]);

  const selectedOriginalRowSet = useMemo<Set<number>>(() => {
    const set = new Set<number>();
    if (!normalizedRange) return set;
    for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
      const orig = visibleRows[r]?.originalIndex;
      if (orig !== undefined) set.add(orig);
    }
    return set;
  }, [normalizedRange, visibleRows]);

  const canDeleteSelectedRows =
    selectedOriginalRowSet.size > 0 &&
    activeSheet.rows.length > 1 &&
    activeSheet.rows.length - selectedOriginalRowSet.size >= 1;

  const isSelectionMerged = useMemo(() => {
    if (!normalizedRange || !activeSheet.merges?.length) return false;
    const { minRow, maxRow, minCol, maxCol } = selectionOriginalBounds;
    return activeSheet.merges.some(
      (m) =>
        m.s.r <= maxRow && m.e.r >= minRow && m.s.c <= maxCol && m.e.c >= minCol
    );
  }, [activeSheet.merges, normalizedRange, selectionOriginalBounds]);

  const buildLiveDocument = useCallback(
    (nextSheets: SpreadsheetSheet[]): DocumentItem => {
      const sheetNames = nextSheets.map((s) => s.name).join(', ');
      return {
        ...document,
        sheets: nextSheets,
        pageCount: nextSheets.length,
        hasLocalEdits: didUserEditRef.current
          ? true
          : Boolean(document.hasLocalEdits),
        contentPreview: `Bảng tính gồm ${nextSheets.length} trang tính (${sheetNames}).`,
      };
    },
    [document]
  );

  // Merge lookup map (active when rows are unfiltered/unsorted and no hidden cols/rows interfere)
  const { mergeOriginMap, mergeCoveredSet } = useMemo(() => {
    const originMap = new Map<string, { rowSpan: number; colSpan: number }>();
    const coveredSet = new Set<string>();
    if (
      hasActiveTransforms ||
      hiddenRowSet.size > 0 ||
      hiddenColSet.size > 0 ||
      !activeSheet.merges?.length
    ) {
      return { mergeOriginMap: originMap, mergeCoveredSet: coveredSet };
    }
    const maxRow = activeSheet.rows.length - 1;
    const maxCol = columnCount - 1;
    for (const m of activeSheet.merges) {
      const sR = Math.max(0, Math.min(maxRow, m.s.r));
      const eR = Math.max(sR, Math.min(maxRow, m.e.r));
      const sC = Math.max(0, Math.min(maxCol, m.s.c));
      const eC = Math.max(sC, Math.min(maxCol, m.e.c));
      if (eR === sR && eC === sC) continue;
      const originKey = makeCellKey(sR, sC);
      if (coveredSet.has(originKey)) continue;
      originMap.set(originKey, {
        rowSpan: eR - sR + 1,
        colSpan: eC - sC + 1,
      });
      for (let r = sR; r <= eR; r += 1) {
        for (let c = sC; c <= eC; c += 1) {
          if (r === sR && c === sC) continue;
          coveredSet.add(makeCellKey(r, c));
        }
      }
    }
    return { mergeOriginMap: originMap, mergeCoveredSet: coveredSet };
  }, [
    activeSheet.merges,
    activeSheet.rows.length,
    columnCount,
    hasActiveTransforms,
    hiddenColSet.size,
    hiddenRowSet.size,
  ]);

  // Precompute Freeze Panes sticky offsets at natural 100% scale
  const headerHeightPx = BASE_HEADER_HEIGHT;
  const rowHeaderWidthPx = BASE_ROW_HEADER_WIDTH;
  const frozenRowsCount = !hasActiveTransforms
    ? Math.max(0, Math.min(activeSheet.frozenRows ?? 0, visibleRows.length))
    : 0;
  const frozenColsCount = Math.max(
    0,
    Math.min(activeSheet.frozenCols ?? 0, visibleColIndices.length)
  );

  const frozenRowTopOffsets = useMemo(() => {
    const map = new Map<number, number>();
    let acc = headerHeightPx;
    for (let i = 0; i < frozenRowsCount; i += 1) {
      map.set(i, acc);
      const origR = visibleRows[i]?.originalIndex ?? i;
      const rawH = activeSheet.rowHeights?.[origR];
      const h =
        rawH && rawH > 16 ? Math.round(rawH) : DEFAULT_ROW_HEIGHT;
      acc += h;
    }
    return map;
  }, [activeSheet.rowHeights, frozenRowsCount, headerHeightPx, visibleRows]);

  const frozenColLeftOffsets = useMemo(() => {
    const map = new Map<number, number>();
    let acc = rowHeaderWidthPx;
    for (let i = 0; i < frozenColsCount; i += 1) {
      const colIdx = visibleColIndices[i];
      if (colIdx === undefined) break;
      map.set(colIdx, acc);
      const rawW = activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH;
      acc += Math.round(rawW);
    }
    return map;
  }, [
    activeSheet.colWidths,
    frozenColsCount,
    rowHeaderWidthPx,
    visibleColIndices,
  ]);

  const estimatedNaturalWidth = useMemo(
    () =>
      BASE_ROW_HEADER_WIDTH +
      visibleColIndices.reduce(
        (sum, c) => sum + Math.round(activeSheet.colWidths?.[c] ?? DEFAULT_COL_WIDTH),
        0
      ),
    [activeSheet.colWidths, visibleColIndices]
  );

  useEffect(() => {
    const el = sheetCanvasRef.current;
    if (!el) return;
    const measure = () => {
      const w = Math.max(el.offsetWidth, el.scrollWidth);
      const h = Math.max(el.offsetHeight, el.scrollHeight);
      setNaturalCanvasSize((prev) =>
        prev.width === w && prev.height === h ? prev : { width: w, height: h }
      );
    };
    measure();
    if (typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [
    activeSheet.id,
    activeSheet.rows.length,
    activeSheet.columns.length,
    activeSheet.colWidths,
    activeSheet.rowHeights,
    activeSheet.charts?.length,
    visibleColIndices.length,
    visibleRows.length,
  ]);

  const effectiveNaturalWidth =
    naturalCanvasSize.width > 0 ? naturalCanvasSize.width : estimatedNaturalWidth;
  const effectiveNaturalHeight = naturalCanvasSize.height;
  const scaledCanvasWidth =
    zoomScale === 1
      ? undefined
      : Math.max(1, Math.ceil(effectiveNaturalWidth * zoomScale));
  const scaledCanvasHeight =
    zoomScale === 1 || effectiveNaturalHeight <= 0
      ? undefined
      : Math.max(1, Math.ceil(effectiveNaturalHeight * zoomScale));

  const showTransientFeedback = useCallback((msg: string) => {
    setFeedbackBanner(msg);
    if (feedbackTimerRef.current !== null) {
      window.clearTimeout(feedbackTimerRef.current);
    }
    feedbackTimerRef.current = window.setTimeout(() => {
      setFeedbackBanner(null);
      feedbackTimerRef.current = null;
    }, 4500);
  }, []);

  const ensureEditMode = useCallback(
    (targetOrigRow?: number, targetCol?: number): boolean => {
      if (isReadOnly) return false;
      if (activeSheet.isProtected) {
        showTransientFeedback(
          'Trang tính đang được bảo vệ. Vào thẻ Bảo vệ để mở khóa trước khi sửa.'
        );
        return false;
      }
      if (
        targetOrigRow !== undefined &&
        targetCol !== undefined &&
        isCellInLockedRanges(targetOrigRow, targetCol, activeSheet.lockedRanges)
      ) {
        showTransientFeedback(
          `Ô ${formatCellRef(targetOrigRow, targetCol)} nằm trong vùng đã khóa (Bảo vệ → Mở khóa vùng ô).`
        );
        return false;
      }
      if (viewModeRef.current !== 'edit') {
        viewModeRef.current = 'edit';
        setViewMode('edit');
      }
      return true;
    },
    [
      activeSheet.isProtected,
      activeSheet.lockedRanges,
      isReadOnly,
      showTransientFeedback,
    ]
  );

  const broadcastExcelPatch = useCallback(
    (
      sheetId: string,
      rowIndex: number,
      colIndex: number,
      value: string | number,
      nextSheets: SpreadsheetSheet[]
    ) => {
      if (isReadOnly || !onBroadcastPatch) return;
      onBroadcastPatch({
        kind: 'excel_cell_patch',
        sheetId,
        rowIndex,
        colIndex,
        value,
        sheets: nextSheets,
      });
    },
    [isReadOnly, onBroadcastPatch]
  );

  // Presence updates
  useEffect(() => {
    if (!onPresenceUpdate || anchorOriginalRow === null || anchorCol === null) {
      return;
    }
    const cellCoord = formatCellRef(anchorOriginalRow, anchorCol);
    const isTyping =
      !isReadOnly &&
      editingCell !== null &&
      editingCell.originalRow === anchorOriginalRow &&
      editingCell.col === anchorCol;
    onPresenceUpdate({
      cellCoord,
      sheetId: activeSheet.id,
      cursorLabel: `Đang ở ô ${cellCoord}`,
      isTyping,
    });
  }, [
    activeSheet.id,
    anchorCol,
    anchorOriginalRow,
    editingCell,
    isReadOnly,
    onPresenceUpdate,
  ]);

  const peerPresenceByCellCoord = useMemo(() => {
    const map = new Map<string, CollaboratorPresence[]>();
    if (!collaborators || collaborators.length === 0) return map;
    for (const peer of collaborators) {
      if (peer.sessionId === selfSessionId || !peer.cellCoord) continue;
      if (peer.sheetId && peer.sheetId !== activeSheet.id) continue;
      const coordKey = peer.cellCoord.toUpperCase();
      const existing = map.get(coordKey);
      if (existing) existing.push(peer);
      else map.set(coordKey, [peer]);
    }
    return map;
  }, [activeSheet.id, collaborators, selfSessionId]);

  const updateActiveSheet = useCallback(
    (
      updater: (sheet: SpreadsheetSheet) => SpreadsheetSheet,
      patchMeta?: {
        rowIndex: number;
        colIndex: number;
        value: string | number;
      },
      options?: { bypassProtection?: boolean }
    ) => {
      if (isReadOnly || viewModeRef.current === 'original') return;
      const prevSheets = sheetsRef.current;
      const currentSheet = prevSheets[safeSheetIdx];
      if (!currentSheet) return;
      if (currentSheet.isProtected && !options?.bypassProtection) {
        showTransientFeedback(
          'Trang tính đang được bảo vệ. Vào thẻ Bảo vệ để mở khóa.'
        );
        return;
      }
      const updatedSheet = updater(currentSheet);
      if (updatedSheet === currentSheet) return;
      didUserEditRef.current = true;
      const nextSheets = prevSheets.map((sheet, idx) =>
        idx === safeSheetIdx ? updatedSheet : sheet
      );
      sheetsRef.current = nextSheets;
      setSheets(nextSheets);
      setHasUnsavedChanges(true);
      setJustSaved(false);

      broadcastExcelPatch(
        updatedSheet.id,
        patchMeta?.rowIndex ?? 0,
        patchMeta?.colIndex ?? 0,
        patchMeta?.value ?? updatedSheet.rows[0]?.[0] ?? '',
        nextSheets
      );
    },
    [broadcastExcelPatch, isReadOnly, safeSheetIdx, showTransientFeedback]
  );

  const commitCellValue = useCallback(
    (originalRow: number, col: number, draft: string) => {
      if (isReadOnly || viewModeRef.current === 'original') return;
      if (
        activeSheet.isProtected ||
        isCellInLockedRanges(originalRow, col, activeSheet.lockedRanges)
      ) {
        return;
      }
      const normalized = normalizeCommittedValue(draft);
      const existing = activeSheet.rows[originalRow]?.[col] ?? '';
      if (existing === normalized) return;
      updateActiveSheet(
        (sheet) => {
          if (originalRow < 0 || originalRow >= sheet.rows.length) return sheet;
          const nextRows = sheet.rows.map((row, rIdx) => {
            if (rIdx !== originalRow) return row;
            const copy = [...row];
            copy[col] = normalized;
            return copy;
          });
          return { ...sheet, rows: nextRows };
        },
        { rowIndex: originalRow, colIndex: col, value: normalized }
      );
    },
    [
      activeSheet.isProtected,
      activeSheet.lockedRanges,
      activeSheet.rows,
      isReadOnly,
      updateActiveSheet,
    ]
  );

  // Fill Handle Completion
  const applyFillHandleDrag = useCallback(
    (drag: FillDragState) => {
      if (isEditingLocked || hasActiveTransforms || !drag.axis) return;
      const { source, targetEnd, axis } = drag;
      const hasVerticalExtension =
        axis === 'vertical' &&
        (targetEnd.row > source.maxRow || targetEnd.row < source.minRow);
      const hasHorizontalExtension =
        axis === 'horizontal' &&
        (targetEnd.col > source.maxCol || targetEnd.col < source.minCol);

      if (!hasVerticalExtension && !hasHorizontalExtension) return;

      updateActiveSheet((sheet) => {
        const nextRows = sheet.rows.map((r) => [...r]);

        if (hasVerticalExtension) {
          const srcLen = source.maxRow - source.minRow + 1;
          const goingDown = targetEnd.row > source.maxRow;
          const startTargetRow = goingDown ? source.maxRow + 1 : targetEnd.row;
          const endTargetRow = goingDown ? targetEnd.row : source.minRow - 1;

          for (let c = source.minCol; c <= source.maxCol; c += 1) {
            const sourceVals: (string | number)[] = [];
            for (let r = source.minRow; r <= source.maxRow; r += 1) {
              const origR = visibleRows[r]?.originalIndex;
              sourceVals.push(origR !== undefined ? (sheet.rows[origR]?.[c] ?? '') : '');
            }
            for (let r = startTargetRow; r <= endTargetRow; r += 1) {
              const targetOrigR = visibleRows[r]?.originalIndex;
              if (targetOrigR === undefined || !nextRows[targetOrigR]) continue;
              const offsetFromEnd = goingDown
                ? r - source.maxRow
                : r - source.minRow;
              const cyclicIdx =
                (((r - source.minRow) % srcLen) + srcLen) % srcLen;
              const srcOrigR =
                visibleRows[source.minRow + cyclicIdx]?.originalIndex ??
                source.minRow + cyclicIdx;
              nextRows[targetOrigR][c] = computeAutoFillValue(
                sourceVals,
                offsetFromEnd,
                targetOrigR - srcOrigR,
                0,
                cyclicIdx
              );
            }
          }
        } else if (hasHorizontalExtension) {
          const srcLen = source.maxCol - source.minCol + 1;
          const goingRight = targetEnd.col > source.maxCol;
          const startTargetCol = goingRight ? source.maxCol + 1 : targetEnd.col;
          const endTargetCol = goingRight ? targetEnd.col : source.minCol - 1;

          for (let r = source.minRow; r <= source.maxRow; r += 1) {
            const origR = visibleRows[r]?.originalIndex;
            if (origR === undefined || !nextRows[origR]) continue;
            const sourceVals: (string | number)[] = [];
            for (let c = source.minCol; c <= source.maxCol; c += 1) {
              sourceVals.push(sheet.rows[origR]?.[c] ?? '');
            }
            for (let c = startTargetCol; c <= endTargetCol; c += 1) {
              const offsetFromEnd = goingRight
                ? c - source.maxCol
                : c - source.minCol;
              const cyclicIdx =
                (((c - source.minCol) % srcLen) + srcLen) % srcLen;
              const srcCol = source.minCol + cyclicIdx;
              nextRows[origR][c] = computeAutoFillValue(
                sourceVals,
                offsetFromEnd,
                0,
                c - srcCol,
                cyclicIdx
              );
            }
          }
        }

        return { ...sheet, rows: nextRows };
      });

      // Expand selection to include filled cells
      setSelection({
        start: {
          row: Math.min(source.minRow, targetEnd.row),
          col: Math.min(source.minCol, targetEnd.col),
        },
        end: {
          row: Math.max(source.maxRow, targetEnd.row),
          col: Math.max(source.maxCol, targetEnd.col),
        },
      });
    },
    [hasActiveTransforms, isEditingLocked, updateActiveSheet, visibleRows]
  );

  useEffect(() => {
    const handleMouseUp = () => {
      isDraggingRef.current = false;
      if (fillDragRef.current) {
        const drag = fillDragRef.current;
        fillDragRef.current = null;
        setFillDrag(null);
        applyFillHandleDrag(drag);
      }
    };
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [applyFillHandleDrag]);

  const startCellEditing = useCallback(
    (
      originalRow: number,
      col: number,
      source: 'cell' | 'formula',
      initialOverride?: string
    ) => {
      if (!ensureEditMode(originalRow, col)) return;
      cancelBlurCommitRef.current = false;
      const rawVal = activeSheet.rows[originalRow]?.[col];
      setEditingCell({ originalRow, col, source });
      setEditDraft(
        initialOverride !== undefined
          ? initialOverride
          : rawVal === undefined || rawVal === null
            ? ''
            : String(rawVal)
      );
    },
    [activeSheet.rows, ensureEditMode]
  );

  const cancelEditing = useCallback(() => {
    cancelBlurCommitRef.current = true;
    setEditingCell(null);
    setEditDraft('');
  }, []);

  const finishEditing = useCallback(
    (moveAfter?: 'down' | 'right' | 'left') => {
      if (!editingCell) return;
      const { originalRow, col } = editingCell;
      commitCellValue(originalRow, col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');

      if (moveAfter && normalizedRange) {
        const maxRow = Math.max(visibleRows.length - 1, 0);
        let nextRow = normalizedRange.anchorRow;
        let nextCol = normalizedRange.anchorCol;

        if (moveAfter === 'down') {
          nextRow = Math.min(maxRow, nextRow + 1);
        } else if (moveAfter === 'right') {
          const pos = visibleColIndices.indexOf(nextCol);
          if (pos !== -1 && pos + 1 < visibleColIndices.length) {
            nextCol = visibleColIndices[pos + 1];
          }
        } else if (moveAfter === 'left') {
          const pos = visibleColIndices.indexOf(nextCol);
          if (pos > 0) {
            nextCol = visibleColIndices[pos - 1];
          }
        }

        setSelection({
          start: { row: nextRow, col: nextCol },
          end: { row: nextRow, col: nextCol },
        });
      }

      window.requestAnimationFrame(() => {
        gridContainerRef.current?.focus({ preventScroll: true });
      });
    },
    [
      commitCellValue,
      editDraft,
      editingCell,
      normalizedRange,
      visibleColIndices,
      visibleRows.length,
    ]
  );

  const handleSave = useCallback(() => {
    if (isReadOnly) return;
    let latestSheets = sheets;
    if (editingCell) {
      const { originalRow, col } = editingCell;
      const normalized = normalizeCommittedValue(editDraft);
      const existingVal = sheets[safeSheetIdx]?.rows[originalRow]?.[col] ?? '';
      latestSheets = sheets.map((sheet, idx) => {
        if (idx !== safeSheetIdx) return sheet;
        const nextRows = sheet.rows.map((row, rIdx) => {
          if (rIdx !== originalRow) return row;
          const copy = [...row];
          copy[col] = normalized;
          return copy;
        });
        return { ...sheet, rows: nextRows };
      });
      sheetsRef.current = latestSheets;
      setSheets(latestSheets);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
      if (existingVal !== normalized && latestSheets[safeSheetIdx]) {
        didUserEditRef.current = true;
        broadcastExcelPatch(
          latestSheets[safeSheetIdx].id,
          originalRow,
          col,
          normalized,
          latestSheets
        );
      }
    }

    onSaveDocument(buildLiveDocument(latestSheets));
    setHasUnsavedChanges(false);
    setJustSaved(true);

    if (saveTimerRef.current !== null) {
      window.clearTimeout(saveTimerRef.current);
    }
    saveTimerRef.current = window.setTimeout(() => {
      setJustSaved(false);
      saveTimerRef.current = null;
    }, 2400);
  }, [
    broadcastExcelPatch,
    buildLiveDocument,
    editDraft,
    editingCell,
    isReadOnly,
    onSaveDocument,
    safeSheetIdx,
    sheets,
  ]);

  const handleRenameTitle = useCallback(
    (newTitle: string) => {
      if (isReadOnly) return;
      let latestSheets = sheets;
      if (editingCell) {
        const { originalRow, col } = editingCell;
        const normalized = normalizeCommittedValue(editDraft);
        latestSheets = sheets.map((sheet, idx) => {
          if (idx !== safeSheetIdx) return sheet;
          const nextRows = sheet.rows.map((row, rIdx) => {
            if (rIdx !== originalRow) return row;
            const copy = [...row];
            copy[col] = normalized;
            return copy;
          });
          return { ...sheet, rows: nextRows };
        });
        setSheets(latestSheets);
        cancelBlurCommitRef.current = true;
        setEditingCell(null);
        setEditDraft('');
      }

      const liveDoc = buildLiveDocument(latestSheets);
      const renamedDoc: DocumentItem = {
        ...liveDoc,
        title: newTitle,
        thumbnail: createThumbnailDataUri(document.type, newTitle, '#107C41'),
        originalFileName: `${newTitle}.${document.type}`,
        updatedAt: 'Vừa chỉnh sửa',
      };
      onSaveDocument(renamedDoc);
      setHasUnsavedChanges(false);
      setJustSaved(true);

      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
      saveTimerRef.current = window.setTimeout(() => {
        setJustSaved(false);
        saveTimerRef.current = null;
      }, 2400);
    },
    [
      buildLiveDocument,
      document.type,
      editDraft,
      editingCell,
      isReadOnly,
      onSaveDocument,
      safeSheetIdx,
      sheets,
    ]
  );

  const handlePrintSheetRef = useRef<() => void>(() => window.print());

  useEffect(() => {
    const handleWindowKeyDown = (event: globalThis.KeyboardEvent) => {
      if (isModalOpen) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        if (!isReadOnly) {
          handleSave();
        }
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        handlePrintSheetRef.current();
      }
    };
    window.addEventListener('keydown', handleWindowKeyDown);
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown);
    };
  }, [handleSave, isModalOpen, isReadOnly]);

  useEffect(() => {
    if (
      isReadOnly ||
      !autoSaveEnabled ||
      !hasUnsavedChanges ||
      editingCell !== null
    ) {
      return;
    }
    const timer = window.setTimeout(() => {
      handleSave();
    }, 1800);
    return () => {
      window.clearTimeout(timer);
    };
  }, [autoSaveEnabled, editingCell, handleSave, hasUnsavedChanges, isReadOnly]);

  const handleExportXlsx = () => {
    let exportSheets = sheets;
    if (editingCell) {
      const { originalRow, col } = editingCell;
      const normalized = normalizeCommittedValue(editDraft);
      exportSheets = sheets.map((sheet, idx) => {
        if (idx !== safeSheetIdx) return sheet;
        const nextRows = sheet.rows.map((row, rIdx) => {
          if (rIdx !== originalRow) return row;
          const copy = [...row];
          copy[col] = normalized;
          return copy;
        });
        return { ...sheet, rows: nextRows };
      });
    }
    exportExcelToXlsx(document.title, exportSheets);
  };

  // --- Styling Operations ---
  const applyStyleToSelection = useCallback(
    (updater: (prev: CellStyle, r: number, c: number) => CellStyle) => {
      if (isEditingLocked || !normalizedRange) return;
      updateActiveSheet((sheet) => {
        const nextStyles: Record<string, CellStyle> = {
          ...(sheet.cellStyles ?? {}),
        };
        for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
          const origRow = visibleRows[r]?.originalIndex;
          if (origRow === undefined) continue;
          for (
            let c = normalizedRange.minCol;
            c <= normalizedRange.maxCol;
            c += 1
          ) {
            if (hiddenColSet.has(c)) continue;
            const key = makeCellKey(origRow, c);
            nextStyles[key] = updater(nextStyles[key] ?? {}, r, c);
          }
        }
        return { ...sheet, cellStyles: nextStyles };
      });
    },
    [hiddenColSet, isEditingLocked, normalizedRange, updateActiveSheet, visibleRows]
  );

  const handleToggleBold = () => {
    const target = !anchorStyle.bold;
    applyStyleToSelection((prev) => ({ ...prev, bold: target }));
  };

  const handleToggleItalic = () => {
    const target = !anchorStyle.italic;
    applyStyleToSelection((prev) => ({ ...prev, italic: target }));
  };

  const handleToggleUnderline = () => {
    const target = !anchorStyle.underline;
    applyStyleToSelection((prev) => ({ ...prev, underline: target }));
  };

  const handleToggleStrike = () => {
    const target = !anchorStyle.strike;
    applyStyleToSelection((prev) => ({ ...prev, strike: target }));
  };

  const handleToggleWrapText = () => {
    const target = !anchorStyle.wrapText;
    applyStyleToSelection((prev) => ({ ...prev, wrapText: target }));
  };

  const handleSetAlign = (align: 'left' | 'center' | 'right' | 'justify') => {
    applyStyleToSelection((prev) => ({ ...prev, align }));
  };

  const handleSetValign = (valign: 'top' | 'middle' | 'bottom') => {
    applyStyleToSelection((prev) => ({ ...prev, valign }));
  };

  const handleSetFontFamily = (fontFamily: string) => {
    applyStyleToSelection((prev) => ({
      ...prev,
      fontFamily: fontFamily || undefined,
    }));
  };

  const handleSetFontSize = (size: number) => {
    const clamped = Math.max(8, Math.min(36, Math.round(size)));
    applyStyleToSelection((prev) => ({ ...prev, fontSize: clamped }));
  };

  const handleStepFontSize = (delta: number) => {
    const current = Math.round(anchorStyle.fontSize ?? 13);
    handleSetFontSize(current + delta);
  };

  const handleSetTextRotation = (
    textRotation: 'horizontal' | 'vertical' | 'up45' | 'down45'
  ) => {
    applyStyleToSelection((prev) => ({
      ...prev,
      textRotation: textRotation === 'horizontal' ? undefined : textRotation,
    }));
    setOpenPopover(null);
  };

  const handleTransformTextCase = (mode: 'upper' | 'lower' | 'proper') => {
    setOpenPopover(null);
    if (isEditBlocked || !normalizedRange) return;
    updateActiveSheet((sheet) => {
      const nextRows = sheet.rows.map((r) => [...r]);
      let changed = false;
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origR = visibleRows[r]?.originalIndex;
        if (origR === undefined || !nextRows[origR]) continue;
        for (
          let c = normalizedRange.minCol;
          c <= normalizedRange.maxCol;
          c += 1
        ) {
          if (
            hiddenColSet.has(c) ||
            isCellInLockedRanges(origR, c, sheet.lockedRanges)
          ) {
            continue;
          }
          const raw = nextRows[origR][c];
          if (typeof raw !== 'string' || !raw.trim() || raw.trim().startsWith('=')) {
            continue;
          }
          const converted =
            mode === 'upper'
              ? raw.toLocaleUpperCase('vi-VN')
              : mode === 'lower'
                ? raw.toLocaleLowerCase('vi-VN')
                : raw
                    .toLocaleLowerCase('vi-VN')
                    .replace(/(^|\s)([^\s])/g, (_, sep: string, ch: string) => {
                      return `${sep}${ch.toLocaleUpperCase('vi-VN')}`;
                    });
          if (converted !== raw) {
            nextRows[origR][c] = converted;
            changed = true;
          }
        }
      }
      return changed ? { ...sheet, rows: nextRows } : sheet;
    });
  };

  const handleSetFillColor = (color: string | undefined) => {
    applyStyleToSelection((prev) => ({ ...prev, bgColor: color }));
    setOpenPopover(null);
  };

  const handleSetTextColor = (color: string | undefined) => {
    applyStyleToSelection((prev) => ({ ...prev, textColor: color }));
    setOpenPopover(null);
  };

  const handleApplyBorderPreset = (
    preset: 'all' | 'outside' | 'thick-box' | 'bottom' | 'none'
  ) => {
    if (!normalizedRange) return;
    const { minRow, maxRow, minCol, maxCol } = normalizedRange;
    applyStyleToSelection((prev, visR, colIdx) => {
      if (preset === 'none') {
        return {
          ...prev,
          borderTop: undefined,
          borderRight: undefined,
          borderBottom: undefined,
          borderLeft: undefined,
        };
      }
      if (preset === 'all') {
        const thin = '1px solid #475569';
        return {
          ...prev,
          borderTop: thin,
          borderRight: thin,
          borderBottom: thin,
          borderLeft: thin,
        };
      }
      if (preset === 'outside') {
        const edge = '1px solid #334155';
        return {
          ...prev,
          borderTop: visR === minRow ? edge : prev.borderTop,
          borderBottom: visR === maxRow ? edge : prev.borderBottom,
          borderLeft: colIdx === minCol ? edge : prev.borderLeft,
          borderRight: colIdx === maxCol ? edge : prev.borderRight,
        };
      }
      if (preset === 'bottom') {
        if (visR === maxRow) {
          return { ...prev, borderBottom: '1.5px solid #334155' };
        }
        return prev;
      }
      // thick-box
      const thick = '2px solid #0F172A';
      return {
        ...prev,
        borderTop: visR === minRow ? thick : prev.borderTop,
        borderBottom: visR === maxRow ? thick : prev.borderBottom,
        borderLeft: colIdx === minCol ? thick : prev.borderLeft,
        borderRight: colIdx === maxCol ? thick : prev.borderRight,
      };
    });
    setOpenPopover(null);
  };

  // --- ONLYOFFICE File, Clipboard, Fill, Clear, Table, Insert, Draw, Layout, Protect, AI Handlers ---
  const handleSaveCopy = () => {
    const copyTitle = `${document.title || 'Bang-tinh'} (Bản sao)`;
    exportExcelToXlsx(copyTitle, sheetsRef.current);
    showTransientFeedback(`Đã tải xuống bản sao "${copyTitle}.xlsx".`);
  };

  const handleDownloadOriginalFile = async () => {
    if (isDownloadingOriginal) return;
    setIsDownloadingOriginal(true);
    try {
      if (workspaceSync.onDownloadOriginalFile) {
        await workspaceSync.onDownloadOriginalFile();
        showTransientFeedback('Đã tải xuống bản gốc .xlsx.');
        return;
      }
      const fallbackName =
        document.originalFileName || `${document.title || 'bang-tinh'}.xlsx`;
      const fallbackMime =
        document.originalMimeType ||
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      if (document.rawArrayBuffer && document.rawArrayBuffer.byteLength > 0) {
        const blob = new Blob([document.rawArrayBuffer.slice(0)], {
          type: fallbackMime,
        });
        const res = await saveFileUniversal(blob, fallbackName, {
          mimeType: fallbackMime,
          shareTitle: document.title,
        });
        showTransientFeedback(res.message);
        return;
      }
      exportExcelToXlsx(document.title, sheetsRef.current);
      showTransientFeedback(`Đã tải xuống "${fallbackName}".`);
    } catch (err: unknown) {
      showTransientFeedback(
        err instanceof Error ? err.message : 'Không thể tải bản gốc lúc này.'
      );
    } finally {
      setIsDownloadingOriginal(false);
    }
  };

  const handleExportPdf = async () => {
    if (isExportingPdf) return;
    setIsExportingPdf(true);
    try {
      const result = await downloadDocumentAsPdf(buildLiveDocument(sheetsRef.current));
      showTransientFeedback(result.message || 'Đã xuất tệp PDF.');
    } catch (err: unknown) {
      showTransientFeedback(
        err instanceof Error ? err.message : 'Không thể xuất PDF lúc này.'
      );
    } finally {
      setIsExportingPdf(false);
    }
  };

  const handlePrintSheet = () => {
    const marginCss =
      activePageSetup.margins === 'narrow'
        ? '8mm'
        : activePageSetup.margins === 'wide'
          ? '24mm'
          : '14mm';
    const styleId = 'lite-excel-print-page-style';
    let styleEl = window.document.getElementById(styleId);
    if (!styleEl) {
      styleEl = window.document.createElement('style');
      styleEl.id = styleId;
      window.document.head.appendChild(styleEl);
    }
    styleEl.textContent = `@media print { @page { size: ${activePageSetup.paperSize} ${activePageSetup.orientation}; margin: ${marginCss}; } ${
      activePageSetup.printGridlines === false
        ? 'table td, table th { border-color: transparent !important; }'
        : ''
    } }`;
    window.print();
  };
  handlePrintSheetRef.current = handlePrintSheet;

  const handleCopyCells = useCallback(() => {
    if (!normalizedRange) return;
    const values: (string | number)[][] = [];
    const styles: (CellStyle | undefined)[][] = [];
    for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
      const origR = visibleRows[r]?.originalIndex;
      if (origR === undefined) continue;
      const rowVals: (string | number)[] = [];
      const rowStyles: (CellStyle | undefined)[] = [];
      for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
        if (hiddenColSet.has(c)) continue;
        rowVals.push(activeSheet.rows[origR]?.[c] ?? '');
        const st = activeSheet.cellStyles?.[makeCellKey(origR, c)];
        rowStyles.push(st ? { ...st } : undefined);
      }
      values.push(rowVals);
      styles.push(rowStyles);
    }
    setClipboardData({ values, styles });
    const tsv = values.map((row) => row.map((v) => String(v)).join('\t')).join('\n');
    if (navigator.clipboard?.writeText) {
      void navigator.clipboard.writeText(tsv).catch(() => undefined);
    }
    showTransientFeedback(`Đã sao chép vùng ${rangeLabel}.`);
  }, [
    activeSheet.cellStyles,
    activeSheet.rows,
    hiddenColSet,
    normalizedRange,
    rangeLabel,
    showTransientFeedback,
    visibleRows,
  ]);

  const handlePasteCells = useCallback(async () => {
    if (isEditBlocked || anchorOriginalRow === null || anchorCol === null) return;
    let matrixVals: (string | number)[][] | null = null;
    let matrixStyles: (CellStyle | undefined)[][] | null = null;

    if (clipboardData && clipboardData.values.length > 0) {
      matrixVals = clipboardData.values;
      matrixStyles = clipboardData.styles;
    } else if (navigator.clipboard?.readText) {
      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          matrixVals = text
            .replace(/\r\n/g, '\n')
            .split('\n')
            .filter((line, idx, arr) => !(idx === arr.length - 1 && line === ''))
            .map((line) =>
              line.split('\t').map((cell) => normalizeCommittedValue(cell))
            );
        }
      } catch {
        // Clipboard API permission unavailable
      }
    }

    if (!matrixVals || matrixVals.length === 0) {
      showTransientFeedback('Bộ nhớ tạm đang trống — hãy sao chép ô trước khi dán.');
      return;
    }

    const startR = anchorOriginalRow;
    const startC = anchorCol;
    const pasteVals = matrixVals;
    const pasteStyles = matrixStyles;

    updateActiveSheet((sheet) => {
      const neededRows = Math.max(sheet.rows.length, startR + pasteVals.length);
      const maxPasteW = pasteVals.reduce((m, row) => Math.max(m, row.length), 0);
      const neededCols = Math.max(sheet.columns.length, startC + maxPasteW);
      const nextRows = Array.from({ length: neededRows }, (_, r) => {
        const existing = sheet.rows[r] ?? [];
        return Array.from({ length: neededCols }, (__, c) => existing[c] ?? '');
      });
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
      };

      for (let dr = 0; dr < pasteVals.length; dr += 1) {
        const targetR = startR + dr;
        for (let dc = 0; dc < pasteVals[dr].length; dc += 1) {
          const targetC = startC + dc;
          if (isCellInLockedRanges(targetR, targetC, sheet.lockedRanges)) {
            continue;
          }
          nextRows[targetR][targetC] = pasteVals[dr][dc];
          const copiedStyle = pasteStyles?.[dr]?.[dc];
          if (copiedStyle) {
            nextStyles[makeCellKey(targetR, targetC)] = { ...copiedStyle };
          }
        }
      }

      return {
        ...sheet,
        columns: createLetterColumns(neededCols),
        rows: nextRows,
        cellStyles: nextStyles,
      };
    });
  }, [
    anchorCol,
    anchorOriginalRow,
    clipboardData,
    isEditBlocked,
    showTransientFeedback,
    updateActiveSheet,
  ]);

  const handleToggleFormatPainter = () => {
    if (isEditBlocked) return;
    if (formatPainterStyle) {
      setFormatPainterStyle(null);
      return;
    }
    setFormatPainterStyle({ ...anchorStyle });
    showTransientFeedback(
      'Chổi sao chép định dạng đang bật — bấm vào ô đích để áp dụng định dạng.'
    );
  };

  const handleFillDirection = (direction: 'down' | 'right') => {
    setOpenPopover(null);
    if (isEditBlocked || !normalizedRange || hasActiveTransforms) return;
    const { minRow, maxRow, minCol, maxCol } = normalizedRange;
    if (direction === 'down' && maxRow <= minRow) {
      showTransientFeedback('Chọn ít nhất 2 dòng để Điền xuống (Fill Down).');
      return;
    }
    if (direction === 'right' && maxCol <= minCol) {
      showTransientFeedback('Chọn ít nhất 2 cột để Điền sang phải (Fill Right).');
      return;
    }

    updateActiveSheet((sheet) => {
      const nextRows = sheet.rows.map((r) => [...r]);
      if (direction === 'down') {
        const topOrigR = visibleRows[minRow]?.originalIndex;
        if (topOrigR === undefined) return sheet;
        for (let c = minCol; c <= maxCol; c += 1) {
          const srcVal = sheet.rows[topOrigR]?.[c] ?? '';
          for (let r = minRow + 1; r <= maxRow; r += 1) {
            const origR = visibleRows[r]?.originalIndex;
            if (
              origR === undefined ||
              !nextRows[origR] ||
              isCellInLockedRanges(origR, c, sheet.lockedRanges)
            ) {
              continue;
            }
            nextRows[origR][c] = computeAutoFillValue(
              [srcVal],
              r - minRow,
              origR - topOrigR,
              0,
              0
            );
          }
        }
      } else {
        for (let r = minRow; r <= maxRow; r += 1) {
          const origR = visibleRows[r]?.originalIndex;
          if (origR === undefined || !nextRows[origR]) continue;
          const srcVal = sheet.rows[origR]?.[minCol] ?? '';
          for (let c = minCol + 1; c <= maxCol; c += 1) {
            if (isCellInLockedRanges(origR, c, sheet.lockedRanges)) continue;
            nextRows[origR][c] = computeAutoFillValue(
              [srcVal],
              c - minCol,
              0,
              c - minCol,
              0
            );
          }
        }
      }
      return { ...sheet, rows: nextRows };
    });
  };

  const handleClearAction = (mode: 'contents' | 'formats' | 'all') => {
    setOpenPopover(null);
    if (isEditBlocked || !normalizedRange) return;
    updateActiveSheet((sheet) => {
      const nextRows =
        mode === 'formats' ? sheet.rows : sheet.rows.map((row) => [...row]);
      const nextStyles =
        mode === 'contents'
          ? sheet.cellStyles
          : { ...(sheet.cellStyles ?? {}) };

      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origR = visibleRows[r]?.originalIndex;
        if (origR === undefined) continue;
        for (
          let c = normalizedRange.minCol;
          c <= normalizedRange.maxCol;
          c += 1
        ) {
          if (
            hiddenColSet.has(c) ||
            isCellInLockedRanges(origR, c, sheet.lockedRanges)
          ) {
            continue;
          }
          if ((mode === 'contents' || mode === 'all') && nextRows[origR]) {
            nextRows[origR][c] = '';
          }
          if ((mode === 'formats' || mode === 'all') && nextStyles) {
            delete nextStyles[makeCellKey(origR, c)];
          }
        }
      }
      return {
        ...sheet,
        rows: nextRows,
        cellStyles: nextStyles,
      };
    });
  };

  const handleFormatAsTable = () => {
    if (isEditBlocked) return;
    const isSingle =
      !normalizedRange ||
      (normalizedRange.minRow === normalizedRange.maxRow &&
        normalizedRange.minCol === normalizedRange.maxCol);

    const minR = isSingle ? 0 : selectionOriginalBounds.minRow;
    const maxR = isSingle
      ? Math.max(0, activeSheet.rows.length - 1)
      : selectionOriginalBounds.maxRow;
    const minC = isSingle ? 0 : selectionOriginalBounds.minCol;
    const maxC = isSingle
      ? Math.max(0, columnCount - 1)
      : selectionOriginalBounds.maxCol;

    updateActiveSheet((sheet) => {
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
      };
      for (let r = minR; r <= maxR; r += 1) {
        for (let c = minC; c <= maxC; c += 1) {
          const key = makeCellKey(r, c);
          if (r === minR) {
            nextStyles[key] = {
              ...(nextStyles[key] ?? {}),
              bgColor: '#107C41',
              textColor: '#FFFFFF',
              bold: true,
              borderTop: '1px solid #0B5E31',
              borderBottom: '2px solid #0B5E31',
              borderLeft: '1px solid #0B5E31',
              borderRight: '1px solid #0B5E31',
            };
          } else {
            const isZebra = (r - minR) % 2 === 0;
            nextStyles[key] = {
              ...(nextStyles[key] ?? {}),
              bgColor: isZebra ? '#F0FDF4' : '#FFFFFF',
              borderTop: '1px solid #D1FAE5',
              borderBottom: '1px solid #D1FAE5',
              borderLeft: '1px solid #E2E8F0',
              borderRight: '1px solid #E2E8F0',
            };
          }
        }
      }
      return {
        ...sheet,
        cellStyles: nextStyles,
        autoFilterEnabled: true,
      };
    });
    showTransientFeedback('Đã áp dụng định dạng bảng Excel (Header xanh & sọc ngựa vằn + Bộ lọc).');
  };

  const handleInsertImageFile = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || isEditBlocked) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result !== 'string') return;
      const newObj: ExcelFloatingObject = {
        id: `img-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        kind: 'image',
        x: 80 + ((activeSheet.floatingObjects?.length ?? 0) % 5) * 28,
        y: 64 + ((activeSheet.floatingObjects?.length ?? 0) % 5) * 24,
        width: 220,
        height: 150,
        src: reader.result,
        text: file.name,
      };
      updateActiveSheet((sheet) => ({
        ...sheet,
        floatingObjects: [...(sheet.floatingObjects ?? []), newObj],
      }));
      setSelectedFloatingObjId(newObj.id);
      showTransientFeedback(`Đã chèn ảnh "${file.name}" vào trang tính.`);
    };
    reader.readAsDataURL(file);
  };

  const handleInsertShape = (
    shapeKind: NonNullable<ExcelFloatingObject['shapeKind']>
  ) => {
    setOpenPopover(null);
    if (isEditBlocked) return;
    const offset = ((activeSheet.floatingObjects?.length ?? 0) % 6) * 24;
    const newObj: ExcelFloatingObject = {
      id: `shape-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      kind: 'shape',
      shapeKind,
      x: 96 + offset,
      y: 72 + offset,
      width: 150,
      height: 86,
      text: shapeKind === 'callout' ? 'Ghi chú nhanh' : '',
      fillColor: '#ECFDF5',
      borderColor: '#107C41',
    };
    updateActiveSheet((sheet) => ({
      ...sheet,
      floatingObjects: [...(sheet.floatingObjects ?? []), newObj],
    }));
    setSelectedFloatingObjId(newObj.id);
  };

  const handleInsertTextBox = () => {
    if (isEditBlocked) return;
    const offset = ((activeSheet.floatingObjects?.length ?? 0) % 6) * 24;
    const newObj: ExcelFloatingObject = {
      id: `textbox-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      kind: 'textbox',
      x: 110 + offset,
      y: 80 + offset,
      width: 210,
      height: 88,
      text: 'Hộp văn bản — bấm đúp để sửa',
      fillColor: '#FFFFFF',
      borderColor: '#107C41',
    };
    updateActiveSheet((sheet) => ({
      ...sheet,
      floatingObjects: [...(sheet.floatingObjects ?? []), newObj],
    }));
    setSelectedFloatingObjId(newObj.id);
  };

  const handleUpdateFloatingObject = (
    id: string,
    patch: Partial<ExcelFloatingObject>
  ) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      floatingObjects: (sheet.floatingObjects ?? []).map((o) =>
        o.id === id ? { ...o, ...patch } : o
      ),
    }));
  };

  const handleDeleteFloatingObject = (id: string) => {
    updateActiveSheet((sheet) => {
      const next = (sheet.floatingObjects ?? []).filter((o) => o.id !== id);
      return { ...sheet, floatingObjects: next.length > 0 ? next : undefined };
    });
    setSelectedFloatingObjId((prev) => (prev === id ? null : prev));
  };

  const handleOpenCellMeta = (mode: ExcelCellMetaMode) => {
    if (isEditBlocked || anchorOriginalRow === null || anchorCol === null) return;
    setCellMetaModal({ mode, origRow: anchorOriginalRow, colIdx: anchorCol });
  };

  const handleSaveCellMeta = (value: string) => {
    if (!cellMetaModal) return;
    const key = makeCellKey(cellMetaModal.origRow, cellMetaModal.colIdx);
    if (cellMetaModal.mode === 'comment') {
      updateActiveSheet((sheet) => ({
        ...sheet,
        cellComments: { ...(sheet.cellComments ?? {}), [key]: value },
      }));
      showTransientFeedback(
        `Đã lưu ghi chú cho ô ${formatCellRef(cellMetaModal.origRow, cellMetaModal.colIdx)}.`
      );
    } else {
      const normalizedUrl =
        /^https?:\/\//i.test(value) ? value : `https://${value}`;
      updateActiveSheet((sheet) => ({
        ...sheet,
        cellStyles: {
          ...(sheet.cellStyles ?? {}),
          [key]: {
            ...(sheet.cellStyles?.[key] ?? {}),
            hyperlink: normalizedUrl,
          },
        },
      }));
      showTransientFeedback(
        `Đã gắn liên kết cho ô ${formatCellRef(cellMetaModal.origRow, cellMetaModal.colIdx)}.`
      );
    }
  };

  const handleRemoveCellMeta = () => {
    if (!cellMetaModal) return;
    const key = makeCellKey(cellMetaModal.origRow, cellMetaModal.colIdx);
    if (cellMetaModal.mode === 'comment') {
      updateActiveSheet((sheet) => {
        if (!sheet.cellComments) return sheet;
        const next = { ...sheet.cellComments };
        delete next[key];
        return {
          ...sheet,
          cellComments: Object.keys(next).length > 0 ? next : undefined,
        };
      });
    } else {
      updateActiveSheet((sheet) => {
        if (!sheet.cellStyles?.[key]) return sheet;
        const nextStyle = { ...sheet.cellStyles[key] };
        delete nextStyle.hyperlink;
        return {
          ...sheet,
          cellStyles: { ...sheet.cellStyles, [key]: nextStyle },
        };
      });
    }
  };

  const handleAddDrawingStroke = (stroke: ExcelDrawingStroke) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      drawings: [...(sheet.drawings ?? []), stroke],
    }));
  };

  const handleEraseDrawingStroke = (strokeId: string) => {
    updateActiveSheet((sheet) => {
      const next = (sheet.drawings ?? []).filter((s) => s.id !== strokeId);
      return { ...sheet, drawings: next.length > 0 ? next : undefined };
    });
  };

  const handleClearDrawings = () => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      drawings: undefined,
    }));
  };

  const handleUpdatePageSetup = (patch: Partial<ExcelPageSetup>) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      pageSetup: {
        ...DEFAULT_PAGE_SETUP,
        ...(sheet.pageSetup ?? {}),
        ...patch,
      },
    }));
  };

  const handleRecalculateSheet = () => {
    setRecalcVersion((v) => v + 1);
    showTransientFeedback('Đã tính toán lại toàn bộ công thức trên trang tính.');
  };

  const handleToggleProtectSheet = () => {
    if (isReadOnly) return;
    if (viewModeRef.current === 'original') {
      viewModeRef.current = 'edit';
      setViewMode('edit');
    }
    const nextProtected = !isSheetProtected;
    if (nextProtected) {
      setActiveDrawTool(null);
      if (editingCell) cancelEditing();
    }
    updateActiveSheet(
      (sheet) => ({
        ...sheet,
        isProtected: nextProtected,
      }),
      undefined,
      { bypassProtection: true }
    );
    showTransientFeedback(
      nextProtected
        ? 'Đã bật Bảo vệ trang tính — mọi thao tác sửa đã được khóa.'
        : 'Đã mở khóa trang tính — bạn có thể tiếp tục chỉnh sửa.'
    );
  };

  const handleToggleLockSelectedRange = () => {
    if (isReadOnly || !normalizedRange || rangeLabel === '—') return;
    if (viewModeRef.current === 'original') {
      viewModeRef.current = 'edit';
      setViewMode('edit');
    }
    const existing = activeSheet.lockedRanges ?? [];
    const alreadyLocked = existing.includes(rangeLabel);
    updateActiveSheet(
      (sheet) => {
        const curr = sheet.lockedRanges ?? [];
        const next = alreadyLocked
          ? curr.filter((r) => r !== rangeLabel)
          : [...curr, rangeLabel];
        return {
          ...sheet,
          lockedRanges: next.length > 0 ? next : undefined,
        };
      },
      undefined,
      { bypassProtection: true }
    );
    showTransientFeedback(
      alreadyLocked
        ? `Đã mở khóa vùng ${rangeLabel}.`
        : `Đã khóa vùng ${rangeLabel} khỏi chỉnh sửa.`
    );
  };

  const handleUnlockSpecificRange = (targetRange: string) => {
    if (isReadOnly) return;
    updateActiveSheet(
      (sheet) => {
        const next = (sheet.lockedRanges ?? []).filter((r) => r !== targetRange);
        return {
          ...sheet,
          lockedRanges: next.length > 0 ? next : undefined,
        };
      },
      undefined,
      { bypassProtection: true }
    );
  };

  const handleToggleShowGridlines = () => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      showGridlines: sheet.showGridlines === false ? true : false,
    }));
  };

  const handleNavigateToCell = useCallback(
    (origRow: number, colIdx: number) => {
      const clampedRow = Math.max(
        0,
        Math.min(activeSheet.rows.length - 1, origRow)
      );
      const clampedCol = Math.max(0, Math.min(columnCount - 1, colIdx));
      let visRow = visibleRows.findIndex((r) => r.originalIndex === clampedRow);
      if (visRow === -1) {
        setSearchQuery('');
        setColumnFilters({});
        setSortConfig(null);
        visRow = clampedRow;
      }
      setSelection({
        start: { row: visRow, col: clampedCol },
        end: { row: visRow, col: clampedCol },
      });
      window.requestAnimationFrame(() => {
        const targetEl = gridContainerRef.current?.querySelector(
          `[data-excel-cell="${clampedRow}:${clampedCol}"]`
        );
        if (targetEl instanceof HTMLElement) {
          targetEl.scrollIntoView({ block: 'center', inline: 'center' });
        }
      });
    },
    [activeSheet.rows.length, columnCount, visibleRows]
  );

  const handleAiInsertFormula = (formula: string) => {
    if (anchorOriginalRow === null || anchorCol === null) return;
    if (!ensureEditMode(anchorOriginalRow, anchorCol)) return;
    if (formula.endsWith('(')) {
      handleQuickInsertFormula(formula);
      return;
    }
    commitCellValue(anchorOriginalRow, anchorCol, formula);
    showTransientFeedback(
      `Đã chèn công thức ${formula} vào ô ${formatCellRef(anchorOriginalRow, anchorCol)}.`
    );
  };

  const handleAiCreateSummarySheet = (
    title: string,
    summaryRows: (string | number)[][]
  ) => {
    if (isReadOnly) return;
    const colCount = Math.max(
      2,
      summaryRows.reduce((m, r) => Math.max(m, r.length), 2)
    );
    const paddedRows = summaryRows.map((row) =>
      Array.from({ length: colCount }, (_, i) => row[i] ?? '')
    );
    const styles: Record<string, CellStyle> = {};
    for (let c = 0; c < colCount; c += 1) {
      styles[makeCellKey(0, c)] = {
        bold: true,
        bgColor: '#107C41',
        textColor: '#FFFFFF',
      };
    }
    const summarySheet: SpreadsheetSheet = {
      id: `sheet-ai-${Date.now()}`,
      name: title.slice(0, 28) || 'Tóm tắt AI',
      columns: createLetterColumns(colCount),
      rows: paddedRows,
      cellStyles: styles,
      tabColor: '#107C41',
    };
    const nextSheets = [...sheetsRef.current, summarySheet];
    commitSheetsChange(nextSheets, nextSheets.length - 1);
    showTransientFeedback(`Đã tạo trang tính "${summarySheet.name}".`);
  };

  const handleHighlightAuditCells = (
    cells: Array<{ row: number; col: number }>,
    bgColor: string
  ) => {
    if (isEditBlocked || cells.length === 0) return;
    updateActiveSheet((sheet) => {
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
      };
      for (const cell of cells) {
        const k = makeCellKey(cell.row, cell.col);
        nextStyles[k] = { ...(nextStyles[k] ?? {}), bgColor };
      }
      return { ...sheet, cellStyles: nextStyles };
    });
    showTransientFeedback(`Đã tô màu cảnh báo cho ${cells.length} ô.`);
  };

  const handleSetNumberFormat = (numFmt: string) => {
    applyStyleToSelection((prev) => ({
      ...prev,
      numFmt: numFmt === 'General' ? undefined : numFmt,
    }));
  };

  const handleAdjustDecimals = (delta: 1 | -1) => {
    applyStyleToSelection((prev) => ({
      ...prev,
      numFmt: adjustNumFmtDecimals(prev.numFmt, delta),
    }));
  };

  // --- Merge & Center / Unmerge ---
  const handleMergeAndCenter = () => {
    if (isEditingLocked || !normalizedRange) return;
    if (hasActiveTransforms) {
      showTransientFeedback('Vui lòng tắt bộ lọc/sắp xếp trước khi gộp ô.');
      return;
    }
    const { minRow, maxRow, minCol, maxCol } = selectionOriginalBounds;
    if (minRow === maxRow && minCol === maxCol) {
      showTransientFeedback('Hãy chọn ít nhất 2 ô để gộp & căn giữa.');
      return;
    }
    updateActiveSheet((sheet) => {
      const nextRows = sheet.rows.map((r) => [...r]);
      let keptValue: string | number = nextRows[minRow]?.[minCol] ?? '';
      for (let r = minRow; r <= maxRow; r += 1) {
        for (let c = minCol; c <= maxCol; c += 1) {
          if (r === minRow && c === minCol) continue;
          const v = nextRows[r]?.[c];
          if (
            (keptValue === '' || keptValue === undefined) &&
            v !== undefined &&
            v !== ''
          ) {
            keptValue = v;
          }
          if (nextRows[r]) {
            nextRows[r][c] = '';
          }
        }
      }
      if (nextRows[minRow]) {
        nextRows[minRow][minCol] = keptValue;
      }

      const survivingMerges = (sheet.merges ?? []).filter(
        (m) =>
          m.e.r < minRow || m.s.r > maxRow || m.e.c < minCol || m.s.c > maxCol
      );
      survivingMerges.push({
        s: { r: minRow, c: minCol },
        e: { r: maxRow, c: maxCol },
      });

      const originKey = makeCellKey(minRow, minCol);
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
        [originKey]: {
          ...(sheet.cellStyles?.[originKey] ?? {}),
          align: 'center',
          valign: 'middle',
        },
      };

      return {
        ...sheet,
        rows: nextRows,
        merges: survivingMerges,
        cellStyles: nextStyles,
      };
    });
  };

  const handleUnmerge = () => {
    if (isEditingLocked || !normalizedRange) return;
    const { minRow, maxRow, minCol, maxCol } = selectionOriginalBounds;
    updateActiveSheet((sheet) => {
      if (!sheet.merges?.length) return sheet;
      const nextMerges = sheet.merges.filter(
        (m) =>
          m.e.r < minRow || m.s.r > maxRow || m.e.c < minCol || m.s.c > maxCol
      );
      return {
        ...sheet,
        merges: nextMerges.length > 0 ? nextMerges : undefined,
      };
    });
  };

  // --- AutoSum (Σ) ---
  const handleAutoSum = (fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => {
    setOpenPopover(null);
    if (isEditingLocked || !normalizedRange) return;

    const { minRow, maxRow, minCol, maxCol } = selectionOriginalBounds;
    const isSingle = minRow === maxRow && minCol === maxCol;

    // Multi-cell range: write formula below (if multi-row) or to right (if single-row)
    if (!isSingle && !hasActiveTransforms) {
      if (maxRow > minRow) {
        const targetRow = maxRow + 1;
        updateActiveSheet((sheet) => {
          const nextRows = sheet.rows.map((r) => [...r]);
          while (nextRows.length <= targetRow) {
            nextRows.push(Array.from({ length: sheet.columns.length }, () => ''));
          }
          for (let c = minCol; c <= maxCol; c += 1) {
            const refStart = formatCellRef(minRow, c);
            const refEnd = formatCellRef(maxRow, c);
            nextRows[targetRow][c] = `=${fn}(${refStart}:${refEnd})`;
          }
          const nextStyles = { ...(sheet.cellStyles ?? {}) };
          for (let c = minCol; c <= maxCol; c += 1) {
            const k = makeCellKey(targetRow, c);
            nextStyles[k] = { ...nextStyles[k], bold: true };
          }
          return { ...sheet, rows: nextRows, cellStyles: nextStyles };
        });
        setSelection({
          start: { row: targetRow, col: minCol },
          end: { row: targetRow, col: maxCol },
        });
        return;
      }

      // Single-row multi-col selection -> place formula in next column
      const targetCol = maxCol + 1;
      updateActiveSheet((sheet) => {
        const needNewCol = targetCol >= sheet.columns.length;
        const nextCols = needNewCol
          ? createLetterColumns(targetCol + 1)
          : sheet.columns;
        const nextRows = sheet.rows.map((row, rIdx) => {
          const padded = needNewCol ? [...row, ''] : [...row];
          if (rIdx === minRow) {
            padded[targetCol] = `=${fn}(${formatCellRef(minRow, minCol)}:${formatCellRef(minRow, maxCol)})`;
          }
          return padded;
        });
        const k = makeCellKey(minRow, targetCol);
        return {
          ...sheet,
          columns: nextCols,
          rows: nextRows,
          colWidths:
            needNewCol && sheet.colWidths
              ? [...sheet.colWidths, DEFAULT_COL_WIDTH]
              : sheet.colWidths,
          cellStyles: {
            ...(sheet.cellStyles ?? {}),
            [k]: { ...(sheet.cellStyles?.[k] ?? {}), bold: true },
          },
        };
      });
      setSelection({
        start: { row: normalizedRange.minRow, col: targetCol },
        end: { row: normalizedRange.minRow, col: targetCol },
      });
      return;
    }

    // Single cell: detect contiguous numeric cells above or left
    if (anchorOriginalRow !== null && anchorCol !== null) {
      let topRow = anchorOriginalRow - 1;
      while (
        topRow >= 0 &&
        parseNumericValue(evaluatedMatrix[topRow]?.[anchorCol]) !== null
      ) {
        topRow -= 1;
      }
      if (topRow < anchorOriginalRow - 1) {
        const formula = `=${fn}(${formatCellRef(topRow + 1, anchorCol)}:${formatCellRef(anchorOriginalRow - 1, anchorCol)})`;
        commitCellValue(anchorOriginalRow, anchorCol, formula);
        return;
      }

      let leftCol = anchorCol - 1;
      while (
        leftCol >= 0 &&
        parseNumericValue(evaluatedMatrix[anchorOriginalRow]?.[leftCol]) !== null
      ) {
        leftCol -= 1;
      }
      if (leftCol < anchorCol - 1) {
        const formula = `=${fn}(${formatCellRef(anchorOriginalRow, leftCol + 1)}:${formatCellRef(anchorOriginalRow, anchorCol - 1)})`;
        commitCellValue(anchorOriginalRow, anchorCol, formula);
        return;
      }

      handleQuickInsertFormula(`=${fn}(`);
    }
  };

  // --- Freeze Panes ---
  const handleSetFreezePanes = (frozenRows: number, frozenCols: number) => {
    setOpenPopover(null);
    if (isEditingLocked) {
      showTransientFeedback('Chuyển sang "Chỉnh sửa nhanh" để thay đổi Cố định khung.');
      return;
    }
    updateActiveSheet((sheet) => ({
      ...sheet,
      frozenRows,
      frozenCols,
    }));
  };

  // --- AutoFilter Toggle ---
  const handleToggleAutoFilter = () => {
    if (isEditingLocked) {
      // Allow clearing local column filters in view mode
      if (hasColumnFilters) {
        setColumnFilters({});
      }
      return;
    }
    const nextEnabled = !activeSheet.autoFilterEnabled;
    updateActiveSheet((sheet) => ({
      ...sheet,
      autoFilterEnabled: nextEnabled,
    }));
    if (!nextEnabled) {
      setColumnFilters({});
      setFilterPopover(null);
    }
  };

  // --- Remove Duplicates ---
  const handleRemoveDuplicates = () => {
    if (isEditingLocked) return;
    if (activeSheet.rows.length <= 1) {
      showTransientFeedback('Bảng tính chưa đủ dòng để kiểm tra trùng lặp.');
      return;
    }
    const startRow = hasHeaderRow ? 1 : 0;
    const colsToCheck =
      normalizedRange && normalizedRange.maxCol > normalizedRange.minCol
        ? Array.from(
            { length: normalizedRange.maxCol - normalizedRange.minCol + 1 },
            (_, i) => normalizedRange.minCol + i
          )
        : Array.from({ length: columnCount }, (_, i) => i);

    const seen = new Set<string>();
    const dupRows = new Set<number>();

    for (let r = startRow; r < activeSheet.rows.length; r += 1) {
      const row = evaluatedMatrix[r] ?? activeSheet.rows[r];
      const key = colsToCheck
        .map((c) => String(row?.[c] ?? '').trim().toLowerCase())
        .join('\u0001');
      // Don't drop completely blank rows unless all checked cols are blank duplicates
      const isAllBlank = colsToCheck.every(
        (c) => String(row?.[c] ?? '').trim() === ''
      );
      if (isAllBlank) continue;
      if (seen.has(key)) {
        dupRows.add(r);
      } else {
        seen.add(key);
      }
    }

    if (dupRows.size === 0) {
      showTransientFeedback('Không tìm thấy dòng trùng lặp nào.');
      return;
    }

    updateActiveSheet((sheet) => {
      const remaining = sheet.rows.filter((_, idx) => !dupRows.has(idx));
      if (remaining.length === 0) return sheet;
      return {
        ...sheet,
        rows: remaining,
        cellStyles: remapRecordAfterRowDelete(sheet.cellStyles, dupRows) ?? {},
        merges: remapMergesAfterRowDelete(sheet.merges, dupRows),
        dataValidations: remapRecordAfterRowDelete(
          sheet.dataValidations,
          dupRows
        ),
        rowHeights: sheet.rowHeights
          ? sheet.rowHeights.filter((_, idx) => !dupRows.has(idx))
          : undefined,
        hiddenRows: remapIndicesAfterDelete(sheet.hiddenRows, dupRows),
      };
    });
    showTransientFeedback(
      `Đã xóa ${dupRows.size} dòng trùng lặp, giữ lại ${activeSheet.rows.length - dupRows.size} dòng duy nhất.`
    );
  };

  // --- Text to Columns ---
  const handleApplyTextToColumns = (delimiter: string) => {
    if (isEditingLocked || anchorCol === null || !delimiter) return;
    const targetCol = anchorCol;
    const startRow = hasHeaderRow ? 1 : 0;

    updateActiveSheet((sheet) => {
      let maxParts = 1;
      const splitMap = new Map<number, string[]>();
      for (let r = startRow; r < sheet.rows.length; r += 1) {
        const raw = String(sheet.rows[r]?.[targetCol] ?? '');
        if (!raw.includes(delimiter)) continue;
        const parts = raw.split(delimiter).map((p) => p.trim());
        if (parts.length > 1) {
          splitMap.set(r, parts);
          maxParts = Math.max(maxParts, parts.length);
        }
      }
      if (splitMap.size === 0) return sheet;

      const neededCols = Math.max(sheet.columns.length, targetCol + maxParts);
      const addedCols = neededCols - sheet.columns.length;
      const nextRows = sheet.rows.map((row, rIdx) => {
        const padded = Array.from(
          { length: neededCols },
          (_, c) => row[c] ?? ''
        );
        const parts = splitMap.get(rIdx);
        if (parts) {
          parts.forEach((part, offset) => {
            padded[targetCol + offset] = normalizeCommittedValue(part);
          });
        }
        return padded;
      });

      return {
        ...sheet,
        columns: createLetterColumns(neededCols),
        rows: nextRows,
        colWidths: sheet.colWidths
          ? [
              ...sheet.colWidths,
              ...Array.from({ length: addedCols }, () => DEFAULT_COL_WIDTH),
            ]
          : undefined,
      };
    });
    showTransientFeedback(
      `Đã tách cột ${columnIndexToLetter(targetCol)} thành công.`
    );
  };

  // --- Data Validation ---
  const handleApplyDataValidation = (options: string[]) => {
    if (isEditingLocked || !normalizedRange) return;
    updateActiveSheet((sheet) => {
      const nextVal: Record<string, string[]> = {
        ...(sheet.dataValidations ?? {}),
      };
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origR = visibleRows[r]?.originalIndex;
        if (origR === undefined) continue;
        for (
          let c = normalizedRange.minCol;
          c <= normalizedRange.maxCol;
          c += 1
        ) {
          nextVal[makeCellKey(origR, c)] = [...options];
        }
      }
      return { ...sheet, dataValidations: nextVal };
    });
  };

  const handleClearDataValidation = () => {
    if (isEditingLocked || !normalizedRange) return;
    updateActiveSheet((sheet) => {
      if (!sheet.dataValidations) return sheet;
      const nextVal: Record<string, string[]> = { ...sheet.dataValidations };
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origR = visibleRows[r]?.originalIndex;
        if (origR === undefined) continue;
        for (
          let c = normalizedRange.minCol;
          c <= normalizedRange.maxCol;
          c += 1
        ) {
          delete nextVal[makeCellKey(origR, c)];
        }
      }
      return {
        ...sheet,
        dataValidations: Object.keys(nextVal).length > 0 ? nextVal : undefined,
      };
    });
  };

  // --- Conditional Formatting ---
  const handleAddConditionalFormat = (rule: ExcelConditionalFormat) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      conditionalFormats: [...(sheet.conditionalFormats ?? []), rule],
    }));
  };

  const handleDeleteConditionalFormat = (ruleId: string) => {
    updateActiveSheet((sheet) => {
      const next = (sheet.conditionalFormats ?? []).filter(
        (r) => r.id !== ruleId
      );
      return {
        ...sheet,
        conditionalFormats: next.length > 0 ? next : undefined,
      };
    });
  };

  // --- Charts ---
  const handleInsertChart = (chart: ExcelChartItem) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      charts: [...(sheet.charts ?? []), chart],
    }));
    showTransientFeedback(`Đã chèn "${chart.title}" vào trang tính.`);
  };

  const handleChangeChartType = (
    chartId: string,
    nextType: ExcelChartItem['type']
  ) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      charts: (sheet.charts ?? []).map((c) =>
        c.id === chartId ? { ...c, type: nextType } : c
      ),
    }));
  };

  const handleDeleteChart = (chartId: string) => {
    updateActiveSheet((sheet) => {
      const next = (sheet.charts ?? []).filter((c) => c.id !== chartId);
      return { ...sheet, charts: next.length > 0 ? next : undefined };
    });
  };

  // --- Structural Row / Column Operations (Ribbon + Context Menu) ---
  const handleInsertRowAt = (atRow: number) => {
    if (isEditingLocked) return;
    if (editingCell) cancelEditing();
    const safeAt = Math.max(0, Math.min(activeSheet.rows.length, atRow));
    updateActiveSheet(
      (sheet) => {
        const blankRow = Array.from({ length: sheet.columns.length }, () => '');
        const nextRows = [
          ...sheet.rows.slice(0, safeAt),
          blankRow,
          ...sheet.rows.slice(safeAt),
        ];
        const nextRowHeights = sheet.rowHeights
          ? [
              ...sheet.rowHeights.slice(0, safeAt),
              DEFAULT_ROW_HEIGHT,
              ...sheet.rowHeights.slice(safeAt),
            ]
          : undefined;
        return {
          ...sheet,
          rows: nextRows,
          cellStyles: remapRecordByRowInsert(sheet.cellStyles, safeAt) ?? {},
          cellComments: remapRecordByRowInsert(sheet.cellComments, safeAt),
          merges: remapMergesAfterRowInsert(sheet.merges, safeAt),
          dataValidations: remapRecordByRowInsert(
            sheet.dataValidations,
            safeAt
          ),
          rowHeights: nextRowHeights,
          hiddenRows: remapIndicesAfterInsert(sheet.hiddenRows, safeAt),
        };
      },
      { rowIndex: safeAt, colIndex: 0, value: '' }
    );
  };

  const handleInsertColumnAt = (atCol: number) => {
    if (isEditBlocked) return;
    if (editingCell) cancelEditing();
    const safeAt = Math.max(0, Math.min(columnCount, atCol));
    updateActiveSheet(
      (sheet) => {
        const nextRows = sheet.rows.map((row) => [
          ...row.slice(0, safeAt),
          '',
          ...row.slice(safeAt),
        ]);
        const nextColWidths = sheet.colWidths
          ? [
              ...sheet.colWidths.slice(0, safeAt),
              DEFAULT_COL_WIDTH,
              ...sheet.colWidths.slice(safeAt),
            ]
          : undefined;
        return {
          ...sheet,
          columns: createLetterColumns(sheet.columns.length + 1),
          rows: nextRows,
          cellStyles: remapRecordByColInsert(sheet.cellStyles, safeAt) ?? {},
          cellComments: remapRecordByColInsert(sheet.cellComments, safeAt),
          merges: remapMergesAfterColInsert(sheet.merges, safeAt),
          dataValidations: remapRecordByColInsert(
            sheet.dataValidations,
            safeAt
          ),
          colWidths: nextColWidths,
          hiddenCols: remapIndicesAfterInsert(sheet.hiddenCols, safeAt),
        };
      },
      { rowIndex: 0, colIndex: safeAt, value: '' }
    );
  };

  const handleDeleteRowIndices = (toDelete: Set<number>) => {
    if (
      isEditBlocked ||
      toDelete.size === 0 ||
      activeSheet.rows.length - toDelete.size < 1
    ) {
      return;
    }
    if (editingCell) cancelEditing();
    updateActiveSheet((sheet) => {
      const remainingRows = sheet.rows.filter((_, rIdx) => !toDelete.has(rIdx));
      if (remainingRows.length === 0) return sheet;
      return {
        ...sheet,
        rows: remainingRows,
        cellStyles: remapRecordAfterRowDelete(sheet.cellStyles, toDelete) ?? {},
        cellComments: remapRecordAfterRowDelete(sheet.cellComments, toDelete),
        merges: remapMergesAfterRowDelete(sheet.merges, toDelete),
        dataValidations: remapRecordAfterRowDelete(
          sheet.dataValidations,
          toDelete
        ),
        rowHeights: sheet.rowHeights
          ? sheet.rowHeights.filter((_, rIdx) => !toDelete.has(rIdx))
          : undefined,
        hiddenRows: remapIndicesAfterDelete(sheet.hiddenRows, toDelete),
      };
    });
    const nextRow = Math.max(0, (normalizedRange?.minRow ?? 0) - 1);
    setSelection({
      start: { row: nextRow, col: normalizedRange?.minCol ?? 0 },
      end: { row: nextRow, col: normalizedRange?.minCol ?? 0 },
    });
  };

  const handleDeleteColumnIndex = (colIdx: number) => {
    if (isEditBlocked || columnCount <= 1) return;
    if (editingCell) cancelEditing();
    const toDelete = new Set<number>([colIdx]);
    updateActiveSheet((sheet) => {
      if (sheet.columns.length <= 1) return sheet;
      const nextRows = sheet.rows.map((row) =>
        row.filter((_, c) => c !== colIdx)
      );
      return {
        ...sheet,
        columns: createLetterColumns(sheet.columns.length - 1),
        rows: nextRows,
        cellStyles: remapRecordAfterColDelete(sheet.cellStyles, toDelete) ?? {},
        cellComments: remapRecordAfterColDelete(sheet.cellComments, toDelete),
        merges: remapMergesAfterColDelete(sheet.merges, toDelete),
        dataValidations: remapRecordAfterColDelete(
          sheet.dataValidations,
          toDelete
        ),
        colWidths: sheet.colWidths
          ? sheet.colWidths.filter((_, c) => c !== colIdx)
          : undefined,
        hiddenCols: remapIndicesAfterDelete(sheet.hiddenCols, toDelete),
      };
    });
    setColumnFilters({});
    setSortConfig((prev) =>
      prev && prev.colIndex === colIdx
        ? null
        : prev && prev.colIndex > colIdx
          ? { ...prev, colIndex: prev.colIndex - 1 }
          : prev
    );
    const nextCol = Math.max(0, Math.min(columnCount - 2, colIdx));
    setSelection({
      start: { row: normalizedRange?.minRow ?? 0, col: nextCol },
      end: { row: normalizedRange?.minRow ?? 0, col: nextCol },
    });
  };

  const handleHideRow = (origRow: number) => {
    if (isEditingLocked) return;
    if (visibleRows.length <= 1) {
      showTransientFeedback('Cần giữ lại ít nhất 1 dòng hiển thị.');
      return;
    }
    updateActiveSheet((sheet) => {
      const next = new Set(sheet.hiddenRows ?? []);
      next.add(origRow);
      return { ...sheet, hiddenRows: Array.from(next).sort((a, b) => a - b) };
    });
  };

  const handleHideColumn = (colIdx: number) => {
    if (isEditingLocked) return;
    if (visibleColIndices.length <= 1) {
      showTransientFeedback('Cần giữ lại ít nhất 1 cột hiển thị.');
      return;
    }
    updateActiveSheet((sheet) => {
      const next = new Set(sheet.hiddenCols ?? []);
      next.add(colIdx);
      return { ...sheet, hiddenCols: Array.from(next).sort((a, b) => a - b) };
    });
  };

  const handleUnhideAll = () => {
    if (isEditingLocked) return;
    updateActiveSheet((sheet) => ({
      ...sheet,
      hiddenRows: undefined,
      hiddenCols: undefined,
    }));
  };

  const handleAddRow = () => {
    if (isEditingLocked) return;
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    const newRowIndex = activeSheet.rows.length;
    handleInsertRowAt(newRowIndex);
    if (searchQuery.trim()) setSearchQuery('');
    if (sortConfig !== null) setSortConfig(null);
    setSelection({
      start: { row: newRowIndex, col: visibleColIndices[0] ?? 0 },
      end: { row: newRowIndex, col: visibleColIndices[0] ?? 0 },
    });
  };

  const handleAddColumn = () => {
    if (isEditingLocked) return;
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    const newColIndex = columnCount;
    handleInsertColumnAt(newColIndex);
    setSelection({
      start: { row: 0, col: newColIndex },
      end: { row: 0, col: newColIndex },
    });
  };

  const handleDeleteSelectedRows = () => {
    handleDeleteRowIndices(new Set(selectedOriginalRowSet));
  };

  // --- Sheet Tabs Management (Add, Rename, Duplicate, Color, Reorder, Delete) ---
  const commitSheetsChange = (
    nextSheets: SpreadsheetSheet[],
    nextActiveIdx: number
  ) => {
    didUserEditRef.current = true;
    sheetsRef.current = nextSheets;
    setSheets(nextSheets);
    setActiveSheetIdx(
      Math.max(0, Math.min(nextSheets.length - 1, nextActiveIdx))
    );
    setHasUnsavedChanges(true);
    setJustSaved(false);
    const active = nextSheets[Math.max(0, Math.min(nextSheets.length - 1, nextActiveIdx))];
    if (active) {
      broadcastExcelPatch(active.id, 0, 0, active.rows[0]?.[0] ?? '', nextSheets);
    }
  };

  const handleSelectSheet = (index: number) => {
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    setActiveSheetIdx(index);
    setActiveDrawTool(null);
    setSelectedFloatingObjId(null);
    setSearchQuery('');
    setSortConfig(null);
    setColumnFilters({});
    setFilterPopover(null);
    setOpenValidationKey(null);
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  };

  const handleAddSheet = () => {
    if (isEditingLocked) return;
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    const currentSheets = sheetsRef.current;
    const existingNames = new Set(currentSheets.map((s) => s.name.toLowerCase()));
    let nextNum = currentSheets.length + 1;
    while (existingNames.has(`sheet ${nextNum}`)) {
      nextNum += 1;
    }
    const newSheet = createBlankSheet(
      `sheet-${Date.now()}-${nextNum}`,
      `Sheet ${nextNum}`
    );
    commitSheetsChange([...currentSheets, newSheet], currentSheets.length);
    setSearchQuery('');
    setSortConfig(null);
    setColumnFilters({});
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  };

  const handleDuplicateSheet = (idx: number) => {
    if (isEditingLocked) return;
    const source = sheetsRef.current[idx];
    if (!source) return;
    const copy: SpreadsheetSheet = {
      ...source,
      id: `sheet-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      name: `${source.name} (Bản sao)`,
      columns: [...source.columns],
      rows: source.rows.map((r) => [...r]),
      cellStyles: source.cellStyles
        ? Object.fromEntries(
            Object.entries(source.cellStyles).map(([k, v]) => [k, { ...v }])
          )
        : {},
      merges: source.merges?.map((m) => ({ s: { ...m.s }, e: { ...m.e } })),
      colWidths: source.colWidths ? [...source.colWidths] : undefined,
      rowHeights: source.rowHeights ? [...source.rowHeights] : undefined,
      hiddenRows: source.hiddenRows ? [...source.hiddenRows] : undefined,
      hiddenCols: source.hiddenCols ? [...source.hiddenCols] : undefined,
      dataValidations: source.dataValidations
        ? Object.fromEntries(
            Object.entries(source.dataValidations).map(([k, v]) => [k, [...v]])
          )
        : undefined,
      conditionalFormats: source.conditionalFormats?.map((cf) => ({
        ...cf,
        id: `cf-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        range: { ...cf.range },
      })),
      charts: source.charts?.map((ch) => ({
        ...ch,
        id: `chart-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        categories: [...ch.categories],
        series: ch.series.map((s) => ({ ...s, values: [...s.values] })),
      })),
      cellComments: source.cellComments ? { ...source.cellComments } : undefined,
      drawings: source.drawings?.map((d) => ({
        ...d,
        id: `draw-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        points: d.points.map((p) => ({ ...p })),
      })),
      floatingObjects: source.floatingObjects?.map((o) => ({
        ...o,
        id: `obj-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      })),
      pageSetup: source.pageSetup ? { ...source.pageSetup } : undefined,
      lockedRanges: source.lockedRanges ? [...source.lockedRanges] : undefined,
    };
    const nextSheets = [
      ...sheetsRef.current.slice(0, idx + 1),
      copy,
      ...sheetsRef.current.slice(idx + 1),
    ];
    commitSheetsChange(nextSheets, idx + 1);
    setSheetTabMenu(null);
  };

  const handleSetSheetTabColor = (idx: number, color: string | undefined) => {
    if (isEditingLocked) return;
    const nextSheets = sheetsRef.current.map((s, i) =>
      i === idx ? { ...s, tabColor: color } : s
    );
    commitSheetsChange(nextSheets, safeSheetIdx);
    setSheetTabMenu(null);
  };

  const handleReorderSheet = (fromIdx: number, toIdx: number) => {
    if (
      isEditingLocked ||
      fromIdx === toIdx ||
      fromIdx < 0 ||
      toIdx < 0 ||
      fromIdx >= sheetsRef.current.length ||
      toIdx >= sheetsRef.current.length
    ) {
      return;
    }
    const nextSheets = [...sheetsRef.current];
    const [moved] = nextSheets.splice(fromIdx, 1);
    nextSheets.splice(toIdx, 0, moved);
    let nextActive = safeSheetIdx;
    if (safeSheetIdx === fromIdx) {
      nextActive = toIdx;
    } else if (fromIdx < safeSheetIdx && toIdx >= safeSheetIdx) {
      nextActive = safeSheetIdx - 1;
    } else if (fromIdx > safeSheetIdx && toIdx <= safeSheetIdx) {
      nextActive = safeSheetIdx + 1;
    }
    commitSheetsChange(nextSheets, nextActive);
    setSheetTabMenu(null);
  };

  const handleDeleteSheet = (idx: number) => {
    if (isEditingLocked || sheetsRef.current.length <= 1) return;
    const nextSheets = sheetsRef.current.filter((_, i) => i !== idx);
    const nextActive =
      safeSheetIdx >= nextSheets.length
        ? nextSheets.length - 1
        : safeSheetIdx > idx
          ? safeSheetIdx - 1
          : safeSheetIdx;
    commitSheetsChange(nextSheets, nextActive);
    setSheetTabMenu(null);
  };

  const startRenameSheet = (idx: number) => {
    if (isEditingLocked) return;
    const target = sheetsRef.current[idx];
    if (!target) return;
    setRenamingSheetIdx(idx);
    setRenameSheetDraft(target.name);
    setSheetTabMenu(null);
  };

  const commitRenameSheet = () => {
    if (renamingSheetIdx === null) return;
    const trimmed = renameSheetDraft.trim();
    if (trimmed) {
      const nextSheets = sheetsRef.current.map((s, i) =>
        i === renamingSheetIdx ? { ...s, name: trimmed } : s
      );
      commitSheetsChange(nextSheets, safeSheetIdx);
    }
    setRenamingSheetIdx(null);
  };

  // --- Filters & Sorting ---
  const handleResetFiltersAndSort = () => {
    setSearchQuery('');
    setSortConfig(null);
    setColumnFilters({});
    setFilterPopover(null);
  };

  const handleSetSort = (colIdx: number, direction: 'asc' | 'desc' | null) => {
    setSortConfig(direction ? { colIndex: colIdx, direction } : null);
  };

  const filterPopoverDistinctValues = useMemo<string[]>(() => {
    if (!filterPopover) return [];
    const colIdx = filterPopover.colIdx;
    const startRow = hasHeaderRow ? 1 : 0;
    const values = new Set<string>();
    for (let r = startRow; r < activeSheet.rows.length; r += 1) {
      if (hiddenRowSet.has(r)) continue;
      const style = activeSheet.cellStyles?.[makeCellKey(r, colIdx)];
      const formatted =
        formatCellDisplay(evaluatedMatrix[r]?.[colIdx], style?.numFmt).trim() ||
        '(Trống)';
      values.add(formatted);
    }
    return Array.from(values).sort((a, b) =>
      a.localeCompare(b, 'vi', { numeric: true, sensitivity: 'base' })
    );
  }, [
    activeSheet.cellStyles,
    activeSheet.rows.length,
    evaluatedMatrix,
    filterPopover,
    hasHeaderRow,
    hiddenRowSet,
  ]);

  const handleToggleColumnFilterValue = (colIdx: number, valLabel: string) => {
    setColumnFilters((prev) => {
      const allVals = filterPopoverDistinctValues;
      const currentAllowed = prev[colIdx]
        ? new Set(prev[colIdx])
        : new Set(allVals);
      if (currentAllowed.has(valLabel)) {
        currentAllowed.delete(valLabel);
      } else {
        currentAllowed.add(valLabel);
      }
      const next = { ...prev };
      if (currentAllowed.size === 0 || currentAllowed.size === allVals.length) {
        delete next[colIdx];
      } else {
        next[colIdx] = currentAllowed;
      }
      return next;
    });
  };

  // --- Selection & Grid Navigation ---
  const handleSelectColumn = (colIdx: number, extendRange: boolean) => {
    if (visibleRows.length === 0) return;
    const lastRow = visibleRows.length - 1;
    if (extendRange && normalizedRange) {
      setSelection({
        start: { row: 0, col: normalizedRange.anchorCol },
        end: { row: lastRow, col: colIdx },
      });
      return;
    }
    setSelection({
      start: { row: 0, col: colIdx },
      end: { row: lastRow, col: colIdx },
    });
  };

  const handleSelectRow = (rowIdx: number, extendRange: boolean) => {
    if (columnCount === 0) return;
    const lastCol = visibleColIndices[visibleColIndices.length - 1] ?? 0;
    const firstCol = visibleColIndices[0] ?? 0;
    if (extendRange && normalizedRange) {
      setSelection({
        start: { row: normalizedRange.anchorRow, col: firstCol },
        end: { row: rowIdx, col: lastCol },
      });
      return;
    }
    setSelection({
      start: { row: rowIdx, col: firstCol },
      end: { row: rowIdx, col: lastCol },
    });
  };

  const handleSelectAllVisible = () => {
    if (visibleRows.length === 0 || visibleColIndices.length === 0) return;
    setSelection({
      start: { row: 0, col: visibleColIndices[0] },
      end: {
        row: visibleRows.length - 1,
        col: visibleColIndices[visibleColIndices.length - 1],
      },
    });
  };

  const handleQuickSelectNumericColumn = () => {
    if (numericColumnIndices.length === 0 || visibleRows.length === 0) return;
    const startVisRow =
      hasHeaderRow &&
      visibleRows.length > 1 &&
      visibleRows[0].originalIndex === 0
        ? 1
        : 0;
    const lastRow = visibleRows.length - 1;
    let nextCol = numericColumnIndices[0];

    if (
      normalizedRange &&
      normalizedRange.minRow === startVisRow &&
      normalizedRange.maxRow === lastRow &&
      normalizedRange.minCol === normalizedRange.maxCol
    ) {
      const currentPos = numericColumnIndices.indexOf(normalizedRange.minCol);
      if (currentPos !== -1) {
        nextCol =
          numericColumnIndices[(currentPos + 1) % numericColumnIndices.length];
      }
    }

    setSelection({
      start: { row: startVisRow, col: nextCol },
      end: { row: lastRow, col: nextCol },
    });
  };

  const handleFitToWidth = useCallback(() => {
    const container = gridContainerRef.current;
    if (!container) {
      setZoomPercent(100);
      return;
    }
    const availableW = Math.max(160, container.clientWidth - 20);
    const naturalW =
      sheetCanvasRef.current && sheetCanvasRef.current.offsetWidth > 0
        ? sheetCanvasRef.current.offsetWidth
        : effectiveNaturalWidth;
    if (naturalW <= 0) {
      setZoomPercent(100);
      return;
    }
    const fitPct = Math.floor((availableW / naturalW) * 100);
    setZoomPercent(Math.max(10, Math.min(300, fitPct)));
  }, [effectiveNaturalWidth]);

  // Name Box jump-to-cell
  const handleNameBoxCommit = () => {
    setIsNameBoxFocused(false);
    const parsed = parseCellOrRangeInput(nameBoxDraft);
    if (!parsed) {
      setNameBoxDraft(rangeLabel);
      return;
    }
    const clampedStartRow = Math.max(
      0,
      Math.min(activeSheet.rows.length - 1, parsed.startRow)
    );
    const clampedEndRow = Math.max(
      0,
      Math.min(activeSheet.rows.length - 1, parsed.endRow)
    );
    const clampedStartCol = Math.max(
      0,
      Math.min(columnCount - 1, parsed.startCol)
    );
    const clampedEndCol = Math.max(0, Math.min(columnCount - 1, parsed.endCol));

    // Find matching visRow index; if filtered out, clear filters so user reaches cell
    let visStart = visibleRows.findIndex(
      (r) => r.originalIndex === clampedStartRow
    );
    let visEnd = visibleRows.findIndex(
      (r) => r.originalIndex === clampedEndRow
    );
    if (visStart === -1 || visEnd === -1) {
      setSearchQuery('');
      setColumnFilters({});
      setSortConfig(null);
      visStart = clampedStartRow;
      visEnd = clampedEndRow;
    }

    setSelection({
      start: { row: visStart, col: clampedStartCol },
      end: { row: visEnd, col: clampedEndCol },
    });

    window.requestAnimationFrame(() => {
      const targetEl = gridContainerRef.current?.querySelector(
        `[data-excel-cell="${clampedStartRow}:${clampedStartCol}"]`
      );
      if (targetEl instanceof HTMLElement) {
        targetEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
      gridContainerRef.current?.focus({ preventScroll: true });
    });
  };

  const handleCellMouseDown = (
    event: ReactMouseEvent<HTMLTableCellElement>,
    visRowIdx: number,
    colIdx: number
  ) => {
    if (event.button !== 0) return;
    setOpenValidationKey(null);
    setSelectedFloatingObjId(null);
    const origRow = visibleRows[visRowIdx]?.originalIndex;
    if (
      editingCell &&
      editingCell.source === 'cell' &&
      editingCell.originalRow === origRow &&
      editingCell.col === colIdx
    ) {
      return;
    }
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }

    event.preventDefault();
    gridContainerRef.current?.focus({ preventScroll: true });

    // Format Painter: apply copied style to clicked cell immediately
    if (formatPainterStyle && origRow !== undefined && !isEditBlocked) {
      const painterCopy = { ...formatPainterStyle };
      setFormatPainterStyle(null);
      setSelection({
        start: { row: visRowIdx, col: colIdx },
        end: { row: visRowIdx, col: colIdx },
      });
      if (!isCellInLockedRanges(origRow, colIdx, activeSheet.lockedRanges)) {
        updateActiveSheet((sheet) => ({
          ...sheet,
          cellStyles: {
            ...(sheet.cellStyles ?? {}),
            [makeCellKey(origRow, colIdx)]: painterCopy,
          },
        }));
      }
      return;
    }

    if (event.shiftKey) {
      setSelection((prev) => ({
        start: normalizedRange
          ? { row: normalizedRange.anchorRow, col: normalizedRange.anchorCol }
          : prev.start,
        end: { row: visRowIdx, col: colIdx },
      }));
      return;
    }

    isDraggingRef.current = true;
    setSelection({
      start: { row: visRowIdx, col: colIdx },
      end: { row: visRowIdx, col: colIdx },
    });
  };

  const handleCellMouseEnter = (visRowIdx: number, colIdx: number) => {
    if (fillDragRef.current) {
      const src = fillDragRef.current.source;
      const distV =
        visRowIdx > src.maxRow
          ? visRowIdx - src.maxRow
          : visRowIdx < src.minRow
            ? src.minRow - visRowIdx
            : 0;
      const distH =
        colIdx > src.maxCol
          ? colIdx - src.maxCol
          : colIdx < src.minCol
            ? src.minCol - colIdx
            : 0;

      if (distV === 0 && distH === 0) {
        const next: FillDragState = {
          source: src,
          targetEnd: { row: src.maxRow, col: src.maxCol },
          axis: null,
        };
        fillDragRef.current = next;
        setFillDrag(next);
        return;
      }

      const axis: 'vertical' | 'horizontal' =
        distV >= distH ? 'vertical' : 'horizontal';
      const next: FillDragState = {
        source: src,
        targetEnd:
          axis === 'vertical'
            ? { row: visRowIdx, col: src.maxCol }
            : { row: src.maxRow, col: colIdx },
        axis,
      };
      fillDragRef.current = next;
      setFillDrag(next);
      return;
    }

    if (!isDraggingRef.current) return;
    setSelection((prev) => ({
      start: prev.start,
      end: { row: visRowIdx, col: colIdx },
    }));
  };

  const handleCellDoubleClick = (visRowIdx: number, colIdx: number) => {
    const origRow = visibleRows[visRowIdx]?.originalIndex;
    if (origRow === undefined) return;
    if (!ensureEditMode(origRow, colIdx)) return;
    setSelection({
      start: { row: visRowIdx, col: colIdx },
      end: { row: visRowIdx, col: colIdx },
    });
    startCellEditing(origRow, colIdx, 'cell');
  };

  const handleOpenContextMenu = (
    event: ReactMouseEvent,
    targetType: 'cell' | 'row' | 'col',
    origRow: number,
    visRow: number,
    colIdx: number
  ) => {
    event.preventDefault();
    event.stopPropagation();
    const clampedX = Math.min(event.clientX, window.innerWidth - 230);
    const clampedY = Math.min(event.clientY, window.innerHeight - 310);
    setContextMenu({
      x: Math.max(8, clampedX),
      y: Math.max(8, clampedY),
      targetType,
      origRow,
      visRow,
      colIdx,
    });
  };

  const handleClearSelectedCells = () => {
    if (isReadOnly || viewModeRef.current === 'original' || !normalizedRange) {
      return;
    }
    updateActiveSheet((sheet) => {
      const nextRows = sheet.rows.map((row) => [...row]);
      let changed = false;
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origRow = visibleRows[r]?.originalIndex;
        if (origRow === undefined || !nextRows[origRow]) continue;
        for (
          let c = normalizedRange.minCol;
          c <= normalizedRange.maxCol;
          c += 1
        ) {
          if (hiddenColSet.has(c)) continue;
          if (nextRows[origRow][c] !== '') {
            nextRows[origRow][c] = '';
            changed = true;
          }
        }
      }
      return changed ? { ...sheet, rows: nextRows } : sheet;
    });
  };

  const handleGridKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (isModalOpen || editingCell !== null) return;
    if (visibleRows.length === 0 || columnCount === 0 || !normalizedRange) return;

    if ((event.ctrlKey || event.metaKey) && !event.altKey) {
      const lowerKey = event.key.toLowerCase();
      if (lowerKey === 'c') {
        event.preventDefault();
        handleCopyCells();
        return;
      }
      if (lowerKey === 'x') {
        event.preventDefault();
        handleCopyCells();
        if (ensureEditMode(anchorOriginalRow ?? undefined, anchorCol ?? undefined)) {
          handleClearSelectedCells();
        }
        return;
      }
      if (lowerKey === 'v') {
        event.preventDefault();
        if (ensureEditMode(anchorOriginalRow ?? undefined, anchorCol ?? undefined)) {
          void handlePasteCells();
        }
        return;
      }
    }

    if (event.key === 'Enter' || event.key === 'F2') {
      if (
        anchorOriginalRow === null ||
        anchorCol === null ||
        !ensureEditMode(anchorOriginalRow, anchorCol)
      ) {
        return;
      }
      event.preventDefault();
      startCellEditing(anchorOriginalRow, anchorCol, 'cell');
      return;
    }

    if (event.key === 'Backspace' || event.key === 'Delete') {
      if (
        !ensureEditMode(
          anchorOriginalRow ?? undefined,
          anchorCol ?? undefined
        )
      ) {
        return;
      }
      event.preventDefault();
      handleClearSelectedCells();
      return;
    }

    if (
      !isReadOnly &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      event.key.length === 1 &&
      event.key !== ' '
    ) {
      if (
        anchorOriginalRow !== null &&
        anchorCol !== null &&
        ensureEditMode(anchorOriginalRow, anchorCol)
      ) {
        event.preventDefault();
        startCellEditing(anchorOriginalRow, anchorCol, 'cell', event.key);
      }
      return;
    }

    let deltaRow = 0;
    let deltaColStep = 0;
    if (event.key === 'ArrowUp') deltaRow = -1;
    else if (event.key === 'ArrowDown') deltaRow = 1;
    else if (event.key === 'ArrowLeft') deltaColStep = -1;
    else if (event.key === 'ArrowRight') deltaColStep = 1;
    else if (event.key === 'Tab') deltaColStep = event.shiftKey ? -1 : 1;
    else return;

    event.preventDefault();
    const maxRow = visibleRows.length - 1;

    setSelection((prev) => {
      const clampedStartRow = Math.max(0, Math.min(maxRow, prev.start.row));
      const clampedEndRow = Math.max(0, Math.min(maxRow, prev.end.row));
      const targetRow = Math.max(0, Math.min(maxRow, clampedEndRow + deltaRow));

      let targetCol = prev.end.col;
      if (deltaColStep !== 0 && visibleColIndices.length > 0) {
        const curPos = visibleColIndices.indexOf(prev.end.col);
        const safePos = curPos === -1 ? 0 : curPos;
        const nextPos = Math.max(
          0,
          Math.min(visibleColIndices.length - 1, safePos + deltaColStep)
        );
        targetCol = visibleColIndices[nextPos];
      }

      if (event.shiftKey && event.key !== 'Tab') {
        return {
          start: { row: clampedStartRow, col: prev.start.col },
          end: { row: targetRow, col: targetCol },
        };
      }
      return {
        start: { row: targetRow, col: targetCol },
        end: { row: targetRow, col: targetCol },
      };
    });
  };

  // --- Formula Bar + Live 38-Function Autocomplete ---
  const handleQuickInsertFormula = (formulaPrefix: string) => {
    setOpenPopover(null);
    if (isEditingLocked || anchorOriginalRow === null || anchorCol === null) {
      return;
    }
    startCellEditing(anchorOriginalRow, anchorCol, 'formula', formulaPrefix);
    window.requestAnimationFrame(() => {
      if (formulaInputRef.current) {
        formulaInputRef.current.focus();
        const len = formulaInputRef.current.value.length;
        formulaInputRef.current.setSelectionRange(len, len);
      }
    });
  };

  const formulaSuggestions = useMemo(() => {
    if (!editingCell || !editDraft.trimStart().startsWith('=')) {
      return [];
    }
    const m = /(?:^|[=+\-*/^(,;&<> ])([A-Za-z]{1,12})$/.exec(editDraft);
    if (!m) {
      return [];
    }
    const prefix = m[1].toUpperCase();
    return ONLYOFFICE_FORMULA_LIBRARY.filter((fn) =>
      fn.name.startsWith(prefix)
    ).slice(0, 7);
  }, [editDraft, editingCell]);

  const handleApplyFormulaSuggestion = (fnName: string) => {
    if (!editingCell) return;
    const nextDraft = editDraft.replace(/([A-Za-z]{1,12})$/, `${fnName}(`);
    setEditDraft(nextDraft);
    window.requestAnimationFrame(() => {
      const targetInput =
        editingCell.source === 'formula'
          ? formulaInputRef.current
          : cellInputRef.current;
      if (targetInput) {
        targetInput.focus();
        const len = nextDraft.length;
        targetInput.setSelectionRange(len, len);
      }
    });
  };

  const handleFormulaFocus = () => {
    if (
      anchorOriginalRow === null ||
      anchorCol === null ||
      !ensureEditMode(anchorOriginalRow, anchorCol)
    ) {
      return;
    }
    if (
      editingCell?.source === 'formula' &&
      editingCell.originalRow === anchorOriginalRow &&
      editingCell.col === anchorCol
    ) {
      return;
    }
    if (
      editingCell?.source === 'cell' &&
      editingCell.originalRow === anchorOriginalRow &&
      editingCell.col === anchorCol
    ) {
      startCellEditing(anchorOriginalRow, anchorCol, 'formula', editDraft);
      return;
    }
    startCellEditing(anchorOriginalRow, anchorCol, 'formula');
  };

  const handleFormulaChange = (event: ChangeEvent<HTMLInputElement>) => {
    if (
      !ensureEditMode(
        anchorOriginalRow ?? undefined,
        anchorCol ?? undefined
      )
    ) {
      return;
    }
    const nextValue = event.target.value;
    if (!editingCell && anchorOriginalRow !== null && anchorCol !== null) {
      startCellEditing(anchorOriginalRow, anchorCol, 'formula', nextValue);
      return;
    }
    setEditDraft(nextValue);
  };

  const handleFormulaKeyDown = (event: ReactKeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Tab' && formulaSuggestions.length > 0) {
      event.preventDefault();
      handleApplyFormulaSuggestion(formulaSuggestions[0].name);
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      finishEditing('down');
      formulaInputRef.current?.blur();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      cancelEditing();
      formulaInputRef.current?.blur();
      gridContainerRef.current?.focus({ preventScroll: true });
    }
  };

  const handleFormulaBlur = () => {
    if (cancelBlurCommitRef.current) {
      cancelBlurCommitRef.current = false;
      return;
    }
    if (editingCell?.source === 'formula') {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      setEditingCell(null);
      setEditDraft('');
    }
  };

  const handleInlineInputKeyDown = (
    event: ReactKeyboardEvent<HTMLInputElement>
  ) => {
    if (event.key === 'Tab' && formulaSuggestions.length > 0) {
      event.preventDefault();
      handleApplyFormulaSuggestion(formulaSuggestions[0].name);
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      finishEditing('down');
    } else if (event.key === 'Tab') {
      event.preventDefault();
      finishEditing(event.shiftKey ? 'left' : 'right');
    } else if (event.key === 'Escape') {
      event.preventDefault();
      cancelEditing();
      gridContainerRef.current?.focus({ preventScroll: true });
    }
  };

  const handleInlineInputBlur = () => {
    if (cancelBlurCommitRef.current) {
      cancelBlurCommitRef.current = false;
      return;
    }
    if (editingCell?.source === 'cell') {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      setEditingCell(null);
      setEditDraft('');
    }
  };

  const formulaInputValue =
    editingCell &&
    anchorOriginalRow === editingCell.originalRow &&
    anchorCol === editingCell.col
      ? editDraft
      : anchorRawValue === undefined || anchorRawValue === null
        ? ''
        : String(anchorRawValue);

  // Fill preview helper bounds
  const fillPreviewBounds = useMemo(() => {
    if (!fillDrag || !fillDrag.axis) return null;
    const { source, targetEnd, axis } = fillDrag;
    if (axis === 'vertical') {
      if (targetEnd.row > source.maxRow) {
        return {
          minRow: source.maxRow + 1,
          maxRow: targetEnd.row,
          minCol: source.minCol,
          maxCol: source.maxCol,
        };
      }
      if (targetEnd.row < source.minRow) {
        return {
          minRow: targetEnd.row,
          maxRow: source.minRow - 1,
          minCol: source.minCol,
          maxCol: source.maxCol,
        };
      }
    } else {
      if (targetEnd.col > source.maxCol) {
        return {
          minRow: source.minRow,
          maxRow: source.maxRow,
          minCol: source.maxCol + 1,
          maxCol: targetEnd.col,
        };
      }
      if (targetEnd.col < source.minCol) {
        return {
          minRow: source.minRow,
          maxRow: source.maxRow,
          minCol: targetEnd.col,
          maxCol: source.minCol - 1,
        };
      }
    }
    return null;
  }, [fillDrag]);

  const hasOriginalBinary = Boolean(
    document.rawArrayBuffer ||
      document.isUploaded ||
      document.storagePath ||
      document.r2ObjectKey
  );

  const handleSwitchViewMode = useCallback(
    (nextMode: 'original' | 'edit') => {
      if (nextMode === 'edit' && isReadOnly) return;
      if (nextMode === 'original' && editingCell) {
        commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
        cancelBlurCommitRef.current = true;
        setEditingCell(null);
        setEditDraft('');
      }
      viewModeRef.current = nextMode;
      setViewMode(nextMode);
    },
    [commitCellValue, editDraft, editingCell, isReadOnly]
  );

  const anchorValidationOptions =
    anchorOriginalRow !== null && anchorCol !== null
      ? (activeSheet.dataValidations?.[
          makeCellKey(anchorOriginalRow, anchorCol)
        ] ?? [])
      : [];

  const textToColsSamples = useMemo(() => {
    if (!isTextToColsModalOpen || anchorCol === null) return [];
    const startR = hasHeaderRow ? 1 : 0;
    return activeSheet.rows
      .slice(startR, startR + 10)
      .map((r) => String(r[anchorCol] ?? ''));
  }, [activeSheet.rows, anchorCol, hasHeaderRow, isTextToColsModalOpen]);

  const hiddenCount = hiddenRowSet.size + hiddenColSet.size;
  const activeFreezeLabel =
    (activeSheet.frozenRows ?? 0) > 0 && (activeSheet.frozenCols ?? 0) > 0
      ? 'Cố định dòng & cột đầu'
      : (activeSheet.frozenRows ?? 0) > 0
        ? 'Cố định dòng đầu'
        : (activeSheet.frozenCols ?? 0) > 0
          ? 'Cố định cột đầu'
          : null;

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#F3F4F6] text-slate-900 overflow-hidden">
      <WorkspaceTopBar
        document={document}
        theme="light"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument(sheets))}
        onRenameTitle={isReadOnly ? undefined : handleRenameTitle}
        collaborators={collaborators}
        selfSessionId={selfSessionId}
        collabPermission={collabPermission}
        modeSwitch={{
          visible: hasOriginalBinary,
          mode: viewMode,
          onChange: handleSwitchViewMode,
          editDisabled: isReadOnly,
          originalTitle: 'Xem nguyên bản (Giống Microsoft Excel)',
          editTitle: 'Chỉnh sửa nhanh ô, công thức và định dạng',
        }}
        onSave={isReadOnly ? undefined : handleSave}
        isDirty={hasUnsavedChanges}
        isSaved={justSaved}
        onDownloadEdited={handleExportXlsx}
      />

      {feedbackBanner && (
        <div className="shrink-0 px-3 sm:px-4 py-1.5 bg-[#ECFDF5] border-b border-[#107C41]/20 text-[#107C41] text-xs font-medium flex items-center justify-between gap-2">
          <span className="truncate">{feedbackBanner}</span>
          <button
            type="button"
            onClick={() => setFeedbackBanner(null)}
            aria-label="Đóng thông báo"
            className="p-0.5 rounded hover:bg-[#107C41]/10 shrink-0"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Hidden image uploader for Insert -> Image */}
      <input
        ref={imageInputRef}
        type="file"
        accept="image/*"
        onChange={handleInsertImageFile}
        className="hidden"
      />

      {/* Microsoft Excel / ONLYOFFICE 11-Tab Fluent Ribbon (#107C41) */}
      <div
        data-excel-popover="true"
        className="shrink-0 bg-[#F8FAFC] border-b border-slate-200/80 relative z-30"
      >
        {/* 11-Tab Strip + Collapse Button */}
        <div className="flex items-center justify-between gap-2 px-2.5 sm:px-4 pt-1">
          <div
            role="tablist"
            aria-label="Thanh công cụ 11 thẻ Excel"
            className="flex items-center gap-0.5 overflow-x-auto no-scrollbar"
          >
            {ONLYOFFICE_RIBBON_TABS.map((tab) => {
              const isActive = activeRibbonTab === tab.id;
              return (
                <button
                  key={tab.id}
                  type="button"
                  role="tab"
                  aria-selected={isActive}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => {
                    setActiveRibbonTab(tab.id);
                    setIsRibbonCollapsed(false);
                    setOpenPopover(null);
                    if (tab.id !== 'draw' && activeDrawTool !== null) {
                      setActiveDrawTool(null);
                    }
                  }}
                  className={`px-2.5 sm:px-3 py-1.5 text-xs whitespace-nowrap border-b-2 transition-colors rounded-t-md ${
                    isActive
                      ? 'border-[#107C41] text-[#107C41] font-semibold bg-white/90'
                      : 'border-transparent text-slate-600 font-medium hover:text-slate-900 hover:bg-slate-200/50'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>

          <div className="flex items-center gap-1.5 shrink-0 pb-0.5">
            {isSheetProtected && (
              <span className="hidden sm:inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-semibold">
                <Lock className="w-3 h-3" />
                Đang bảo vệ
              </span>
            )}
            {hasActiveTransforms && (
              <button
                type="button"
                onClick={handleResetFiltersAndSort}
                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-50 hover:bg-amber-100 border border-amber-200/80 text-amber-800 text-[11px] font-semibold whitespace-nowrap shrink-0 transition-colors"
                title="Đặt lại toàn bộ bộ lọc và sắp xếp"
              >
                <RotateCcw className="w-3 h-3 shrink-0" />
                <span className="hidden sm:inline">Đặt lại lọc</span>
              </button>
            )}
            {isFiltered && (
              <span className="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-[#ECFDF5] border border-[#107C41]/25 text-[#107C41] text-[11px] font-semibold">
                <Filter className="w-3 h-3" />
                {visibleRows.length}/{activeSheet.rows.length} dòng
              </span>
            )}
            {activeFreezeLabel && (
              <span className="hidden lg:inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-sky-50 border border-sky-200/80 text-sky-700 text-[11px] font-semibold">
                <Snowflake className="w-3 h-3" />
                {activeFreezeLabel}
              </span>
            )}
            {hiddenCount > 0 && !isEditBlocked && (
              <button
                type="button"
                onClick={handleUnhideAll}
                className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-800 text-[11px] font-semibold transition-colors"
                title="Hiện lại toàn bộ dòng và cột đang ẩn"
              >
                <Eye className="w-3 h-3" />
                <span>Hiện {hiddenCount} ẩn</span>
              </button>
            )}
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => setIsRibbonCollapsed((v) => !v)}
              title={
                isRibbonCollapsed
                  ? 'Mở thanh công cụ Ribbon'
                  : 'Thu gọn thanh công cụ Ribbon'
              }
              aria-label={isRibbonCollapsed ? 'Mở Ribbon' : 'Thu gọn Ribbon'}
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg text-slate-500 hover:text-[#107C41] hover:bg-slate-200/70 transition-colors"
            >
              {isRibbonCollapsed ? (
                <ChevronDown className="w-4 h-4" />
              ) : (
                <ChevronUp className="w-4 h-4" />
              )}
            </button>
          </div>
        </div>

        {/* Floating Fluent Command Card */}
        {!isRibbonCollapsed && (
          <div className="rounded-xl bg-white shadow-[0_1px_3px_rgba(15,23,42,0.06)] border border-slate-200/80 mx-2 my-1.5 px-2.5 py-1.5 overflow-x-auto no-scrollbar">
            {/* TAB 1: TỆP (FILE) */}
            {activeRibbonTab === 'file' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  disabled={isReadOnly}
                  onClick={handleSave}
                  className="h-8 px-3 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>Lưu</span>
                </button>
                <button
                  type="button"
                  onClick={handleSaveCopy}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 transition-colors"
                >
                  <Copy className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Lưu bản sao</span>
                </button>
                <span className="h-6 w-px bg-slate-200/80 mx-1 shrink-0" />
                <button
                  type="button"
                  disabled={isDownloadingOriginal}
                  onClick={() => void handleDownloadOriginalFile()}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  {isDownloadingOriginal ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-[#107C41]" />
                  ) : (
                    <Download className="w-3.5 h-3.5 text-[#107C41]" />
                  )}
                  <span>Tải bản gốc .xlsx</span>
                </button>
                <button
                  type="button"
                  onClick={handleExportXlsx}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 transition-colors"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Tải bản sửa .xlsx</span>
                </button>
                <button
                  type="button"
                  disabled={isExportingPdf}
                  onClick={() => void handleExportPdf()}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  {isExportingPdf ? (
                    <Loader2 className="w-3.5 h-3.5 animate-spin text-rose-600" />
                  ) : (
                    <FileDown className="w-3.5 h-3.5 text-rose-600" />
                  )}
                  <span>Xuất PDF</span>
                </button>
                <span className="h-6 w-px bg-slate-200/80 mx-1 shrink-0" />
                <button
                  type="button"
                  onClick={handlePrintSheet}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 transition-colors"
                >
                  <Printer className="w-3.5 h-3.5 text-slate-600" />
                  <span>In trang tính</span>
                </button>
                <button
                  type="button"
                  onClick={() => setIsFileInfoOpen(true)}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 transition-colors"
                >
                  <Info className="w-3.5 h-3.5 text-slate-600" />
                  <span>Thông tin tệp</span>
                </button>
              </div>
            )}

            {/* TAB 2: TRANG CHỦ (HOME — 2-row compact ONLYOFFICE grouped Ribbon) */}
            {activeRibbonTab === 'home' && (
              <div className="flex items-stretch gap-2 w-max">
                {/* Group 1: Clipboard */}
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    disabled={isEditBlocked || !normalizedRange}
                    onClick={() => void handlePasteCells()}
                    title="Dán (Ctrl+V)"
                    className="h-[58px] px-2.5 rounded-lg border border-slate-200/80 bg-slate-50/70 hover:bg-[#ECFDF5] hover:border-[#107C41]/40 text-slate-700 hover:text-[#107C41] flex flex-col items-center justify-center gap-0.5 text-[11px] font-semibold disabled:opacity-40 transition-colors"
                  >
                    <ClipboardPaste className="w-4 h-4 text-[#107C41]" />
                    <span>Dán</span>
                  </button>
                  <div className="flex flex-col gap-1">
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        disabled={!normalizedRange}
                        onClick={handleCopyCells}
                        title="Sao chép (Ctrl+C)"
                        className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                      >
                        <Copy className="w-3.5 h-3.5" />
                        <span>Chép</span>
                      </button>
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        disabled={isEditBlocked || !normalizedRange}
                        onClick={() => {
                          handleCopyCells();
                          handleClearSelectedCells();
                        }}
                        title="Cắt (Ctrl+X)"
                        className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                      >
                        <Scissors className="w-3.5 h-3.5" />
                        <span>Cắt</span>
                      </button>
                    </div>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={handleToggleFormatPainter}
                      title="Chổi sao chép định dạng (Format Painter)"
                      className={`h-7 px-2 rounded-md border text-[11px] font-semibold inline-flex items-center justify-center gap-1 disabled:opacity-40 transition-colors ${
                        formatPainterStyle
                          ? 'bg-[#107C41] border-[#107C41] text-white'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]'
                      }`}
                    >
                      <Paintbrush className="w-3.5 h-3.5" />
                      <span>Sao chép ĐD</span>
                    </button>
                  </div>
                </div>

                <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                {/* Group 2: Font (2 rows) */}
                <div className="flex flex-col gap-1 shrink-0">
                  <div className="flex items-center gap-1">
                    <select
                      aria-label="Phông chữ"
                      title="Phông chữ"
                      disabled={isEditBlocked || !normalizedRange}
                      value={anchorStyle.fontFamily ?? ''}
                      onChange={(e) => handleSetFontFamily(e.target.value)}
                      className="h-7 w-28 px-1.5 rounded-md border border-slate-200/80 bg-white text-xs font-medium text-slate-700 focus:outline-none focus:border-[#107C41] disabled:opacity-40"
                    >
                      {FONT_FAMILY_OPTIONS.map((opt) => (
                        <option key={opt.label} value={opt.value}>
                          {opt.label}
                        </option>
                      ))}
                    </select>
                    <select
                      aria-label="Cỡ chữ ô"
                      title="Cỡ chữ"
                      disabled={isEditBlocked || !normalizedRange}
                      value={Math.round(anchorStyle.fontSize ?? 13)}
                      onChange={(e) => handleSetFontSize(Number(e.target.value))}
                      className="h-7 px-1.5 rounded-md border border-slate-200/80 bg-white text-xs font-semibold text-slate-700 focus:outline-none focus:border-[#107C41] disabled:opacity-40"
                    >
                      {FONT_SIZE_OPTIONS.map((sz) => (
                        <option key={sz} value={sz}>
                          {sz}
                        </option>
                      ))}
                    </select>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleStepFontSize(1)}
                      title="Tăng cỡ chữ (A^)"
                      className="w-7 h-7 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] inline-flex items-center justify-center text-slate-700 disabled:opacity-40"
                    >
                      <AArrowUp className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleStepFontSize(-1)}
                      title="Giảm cỡ chữ (Av)"
                      className="w-7 h-7 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] inline-flex items-center justify-center text-slate-700 disabled:opacity-40"
                    >
                      <AArrowDown className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) => (p === 'case' ? null : 'case'))
                      }
                      title="Đổi kiểu chữ HOA / thường / Đầu từ (Aa)"
                      className={`h-7 px-1.5 rounded-md border inline-flex items-center gap-0.5 text-xs font-semibold disabled:opacity-40 ${
                        openPopover === 'case'
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]'
                      }`}
                    >
                      <CaseSensitive className="w-3.5 h-3.5" />
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </div>

                  <div className="flex items-center gap-1">
                    <div className="inline-flex items-center gap-0.5 rounded-md bg-slate-50 border border-slate-200/80 p-0.5">
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={handleToggleBold}
                        disabled={isEditBlocked || !normalizedRange}
                        aria-pressed={anchorStyle.bold === true}
                        title="In đậm (B)"
                        className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                          anchorStyle.bold
                            ? 'bg-[#ECFDF5] text-[#107C41] font-bold'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        <Bold className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={handleToggleItalic}
                        disabled={isEditBlocked || !normalizedRange}
                        aria-pressed={anchorStyle.italic === true}
                        title="In nghiêng (I)"
                        className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                          anchorStyle.italic
                            ? 'bg-[#ECFDF5] text-[#107C41]'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        <Italic className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={handleToggleUnderline}
                        disabled={isEditBlocked || !normalizedRange}
                        aria-pressed={anchorStyle.underline === true}
                        title="Gạch chân (U)"
                        className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                          anchorStyle.underline
                            ? 'bg-[#ECFDF5] text-[#107C41]'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        <Underline className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={handleToggleStrike}
                        disabled={isEditBlocked || !normalizedRange}
                        aria-pressed={anchorStyle.strike === true}
                        title="Gạch ngang chữ (S)"
                        className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                          anchorStyle.strike
                            ? 'bg-[#ECFDF5] text-[#107C41]'
                            : 'text-slate-700 hover:bg-white'
                        }`}
                      >
                        <Strikethrough className="w-3.5 h-3.5" />
                      </button>
                    </div>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) =>
                          p === 'textColor' ? null : 'textColor'
                        )
                      }
                      title="Màu chữ"
                      className={`h-7 px-1.5 rounded-md border inline-flex items-center gap-0.5 disabled:opacity-40 ${
                        openPopover === 'textColor'
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <span className="flex flex-col items-center">
                        <Type className="w-3.5 h-3.5" />
                        <span
                          className="w-3.5 h-1 rounded-full -mt-0.5"
                          style={{
                            backgroundColor: anchorStyle.textColor ?? '#0F172A',
                          }}
                        />
                      </span>
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) =>
                          p === 'fillColor' ? null : 'fillColor'
                        )
                      }
                      title="Màu nền ô"
                      className={`h-7 px-1.5 rounded-md border inline-flex items-center gap-0.5 disabled:opacity-40 ${
                        openPopover === 'fillColor'
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <span className="flex flex-col items-center">
                        <PaintBucket className="w-3.5 h-3.5" />
                        <span
                          className="w-3.5 h-1 rounded-full -mt-0.5 border border-black/15"
                          style={{
                            backgroundColor: anchorStyle.bgColor ?? '#FEF9C3',
                          }}
                        />
                      </span>
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) =>
                          p === 'borders' ? null : 'borders'
                        )
                      }
                      title="Kẻ viền ô (Borders)"
                      className={`h-7 px-1.5 rounded-md border inline-flex items-center gap-0.5 text-xs font-medium disabled:opacity-40 ${
                        openPopover === 'borders'
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <Grid3x3 className="w-3.5 h-3.5" />
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </div>
                </div>

                <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                {/* Group 3: Alignment (2 rows) */}
                <div className="flex flex-col gap-1 shrink-0">
                  <div className="flex items-center gap-1">
                    <div
                      role="group"
                      aria-label="Căn lề dọc"
                      className="inline-flex items-center gap-0.5 rounded-md bg-slate-50 border border-slate-200/80 p-0.5"
                    >
                      {(
                        [
                          {
                            id: 'top',
                            icon: AlignVerticalJustifyStart,
                            label: 'Căn trên',
                          },
                          {
                            id: 'middle',
                            icon: AlignVerticalJustifyCenter,
                            label: 'Căn giữa dọc',
                          },
                          {
                            id: 'bottom',
                            icon: AlignVerticalJustifyEnd,
                            label: 'Căn dưới',
                          },
                        ] as const
                      ).map(({ id, icon: Icon, label }) => (
                        <button
                          key={id}
                          type="button"
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={() => handleSetValign(id)}
                          disabled={isEditBlocked || !normalizedRange}
                          title={label}
                          aria-label={label}
                          className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                            anchorEffectiveValign === id
                              ? 'bg-[#ECFDF5] text-[#107C41]'
                              : 'text-slate-600 hover:bg-white'
                          }`}
                        >
                          <Icon className="w-3.5 h-3.5" />
                        </button>
                      ))}
                    </div>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={handleToggleWrapText}
                      disabled={isEditBlocked || !normalizedRange}
                      aria-pressed={anchorStyle.wrapText === true}
                      title="Tự xuống dòng (Wrap Text)"
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-medium disabled:opacity-40 ${
                        anchorStyle.wrapText
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#107C41] font-semibold'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <WrapText className="w-3.5 h-3.5" />
                      <span>Xuống dòng</span>
                    </button>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) =>
                          p === 'textDir' ? null : 'textDir'
                        )
                      }
                      title="Hướng chữ trong ô"
                      className={`h-7 px-1.5 rounded-md border inline-flex items-center gap-0.5 text-[11px] font-semibold disabled:opacity-40 ${
                        anchorStyle.textRotation || openPopover === 'textDir'
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <span>ab↗</span>
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </div>

                  <div className="flex items-center gap-1">
                    <div
                      role="group"
                      aria-label="Căn lề ngang"
                      className="inline-flex items-center gap-0.5 rounded-md bg-slate-50 border border-slate-200/80 p-0.5"
                    >
                      {(
                        [
                          { id: 'left', icon: AlignLeft, label: 'Căn trái' },
                          { id: 'center', icon: AlignCenter, label: 'Căn giữa' },
                          { id: 'right', icon: AlignRight, label: 'Căn phải' },
                          { id: 'justify', icon: AlignJustify, label: 'Căn đều' },
                        ] as const
                      ).map(({ id, icon: Icon, label }) => (
                        <button
                          key={id}
                          type="button"
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={() => handleSetAlign(id)}
                          disabled={isEditBlocked || !normalizedRange}
                          title={label}
                          aria-label={label}
                          className={`w-6 h-6 rounded inline-flex items-center justify-center disabled:opacity-40 ${
                            anchorEffectiveAlign === id
                              ? 'bg-[#ECFDF5] text-[#107C41]'
                              : 'text-slate-600 hover:bg-white'
                          }`}
                        >
                          <Icon className="w-3.5 h-3.5" />
                        </button>
                      ))}
                    </div>

                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={
                        isSelectionMerged ? handleUnmerge : handleMergeAndCenter
                      }
                      disabled={isEditBlocked || !normalizedRange}
                      title={
                        isSelectionMerged
                          ? 'Bỏ gộp ô (Unmerge)'
                          : 'Gộp & căn giữa (Merge & Center)'
                      }
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-medium disabled:opacity-40 ${
                        isSelectionMerged
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#107C41] font-semibold'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      {isSelectionMerged ? (
                        <>
                          <Split className="w-3.5 h-3.5" />
                          <span>Bỏ gộp</span>
                        </>
                      ) : (
                        <>
                          <Merge className="w-3.5 h-3.5" />
                          <span>Gộp &amp; Căn giữa</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                {/* Group 4: Number Format (2 rows) */}
                <div className="flex flex-col gap-1 shrink-0">
                  <select
                    aria-label="Định dạng số"
                    title="Định dạng số"
                    disabled={isEditBlocked || !normalizedRange}
                    value={detectFormatPresetValue(anchorStyle.numFmt)}
                    onChange={(e) => handleSetNumberFormat(e.target.value)}
                    className="h-7 px-2 rounded-md border border-slate-200/80 bg-white text-xs font-medium text-slate-700 focus:outline-none focus:border-[#107C41] disabled:opacity-40"
                  >
                    {NUMBER_FORMAT_PRESETS.map((fmt) => (
                      <option key={fmt.id} value={fmt.numFmt}>
                        {fmt.label}
                      </option>
                    ))}
                  </select>
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleSetNumberFormat('#,##0 "₫"')}
                      title="Định dạng tiền tệ VNĐ (₫)"
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-bold text-slate-700 disabled:opacity-40"
                    >
                      ₫
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleSetNumberFormat('0.0%')}
                      title="Phần trăm (%)"
                      className="h-7 px-1.5 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-slate-700 inline-flex items-center justify-center disabled:opacity-40"
                    >
                      <Percent className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleSetNumberFormat('#,##0')}
                      title="Dấu phân cách hàng nghìn (#,##0)"
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-bold text-slate-700 disabled:opacity-40"
                    >
                      ,
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleAdjustDecimals(1)}
                      title="Tăng chữ số thập phân (.00)"
                      className="h-7 px-1.5 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-bold text-slate-700 disabled:opacity-40"
                    >
                      .00
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => handleAdjustDecimals(-1)}
                      title="Giảm chữ số thập phân (.0)"
                      className="h-7 px-1.5 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-bold text-slate-700 disabled:opacity-40"
                    >
                      .0
                    </button>
                  </div>
                </div>

                <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                {/* Group 5: Quick Formulas, Fill, Clear & Sort (2 rows) */}
                <div className="flex flex-col gap-1 shrink-0">
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) =>
                          p === 'autoSum' ? null : 'autoSum'
                        )
                      }
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-semibold disabled:opacity-40 ${
                        openPopover === 'autoSum'
                          ? 'bg-[#107C41] border-[#107C41] text-white'
                          : 'bg-[#ECFDF5] border-[#107C41]/30 text-[#107C41] hover:bg-[#D1FAE5]'
                      }`}
                    >
                      <Sigma className="w-3.5 h-3.5" />
                      <span>AutoSum</span>
                      <ChevronDown className="w-3 h-3" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) => (p === 'fill' ? null : 'fill'))
                      }
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <ArrowDownToLine className="w-3 h-3 text-[#107C41]" />
                      <span>Điền</span>
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() =>
                        setOpenPopover((p) => (p === 'clear' ? null : 'clear'))
                      }
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-rose-50 hover:text-rose-600 text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Eraser className="w-3 h-3" />
                      <span>Xóa</span>
                      <ChevronDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </div>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={anchorCol === null}
                      onClick={() =>
                        anchorCol !== null && handleSetSort(anchorCol, 'asc')
                      }
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-semibold disabled:opacity-40 ${
                        sortConfig?.colIndex === anchorCol &&
                        sortConfig?.direction === 'asc'
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <ArrowUp className="w-3 h-3 text-[#107C41]" />
                      <span>A→Z</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={anchorCol === null}
                      onClick={() =>
                        anchorCol !== null && handleSetSort(anchorCol, 'desc')
                      }
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-semibold disabled:opacity-40 ${
                        sortConfig?.colIndex === anchorCol &&
                        sortConfig?.direction === 'desc'
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <ArrowDown className="w-3 h-3 text-[#107C41]" />
                      <span>Z→A</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={handleToggleAutoFilter}
                      className={`h-7 px-2 rounded-md border inline-flex items-center gap-1 text-[11px] font-semibold ${
                        activeSheet.autoFilterEnabled || hasColumnFilters
                          ? 'bg-[#107C41] border-[#107C41] text-white'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                      }`}
                    >
                      <ListFilter className="w-3 h-3" />
                      <span>Lọc phễu</span>
                    </button>
                  </div>
                </div>

                <span className="w-px bg-slate-200/80 my-1 shrink-0" />

                {/* Group 6: Cells, Format as Table & Conditional Formatting (2 rows) */}
                <div className="flex flex-col gap-1 shrink-0">
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked}
                      onClick={handleAddRow}
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Plus className="w-3 h-3 text-[#107C41]" />
                      <span>Hàng</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked}
                      onClick={handleAddColumn}
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Plus className="w-3 h-3 text-[#107C41]" />
                      <span>Cột</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked}
                      onClick={handleFormatAsTable}
                      title="Định dạng thành bảng (Header xanh #107C41 + sọc ngựa vằn + Bộ lọc)"
                      className="h-7 px-2 rounded-md border border-[#107C41]/30 bg-[#ECFDF5] hover:bg-[#D1FAE5] text-[11px] font-semibold text-[#107C41] inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Table className="w-3.5 h-3.5" />
                      <span>Format as Table</span>
                    </button>
                  </div>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !canDeleteSelectedRows}
                      onClick={handleDeleteSelectedRows}
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-rose-50 hover:text-rose-600 text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Trash2 className="w-3 h-3" />
                      <span>Xóa hàng</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={
                        isEditBlocked || anchorCol === null || columnCount <= 1
                      }
                      onClick={() =>
                        anchorCol !== null && handleDeleteColumnIndex(anchorCol)
                      }
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-rose-50 hover:text-rose-600 text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Trash2 className="w-3 h-3" />
                      <span>Xóa cột</span>
                    </button>
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditBlocked || !normalizedRange}
                      onClick={() => setIsCondFormatModalOpen(true)}
                      className="h-7 px-2 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-[11px] font-medium text-slate-700 inline-flex items-center gap-1 disabled:opacity-40"
                    >
                      <Sparkles className="w-3.5 h-3.5 text-[#107C41]" />
                      <span>Định dạng ĐK</span>
                      {(activeSheet.conditionalFormats?.length ?? 0) > 0 && (
                        <span className="px-1 rounded-full bg-[#107C41] text-white text-[10px] font-bold">
                          {activeSheet.conditionalFormats?.length}
                        </span>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* TAB 3: CHÈN (INSERT) */}
            {activeRibbonTab === 'insert' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={handleFormatAsTable}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <Table className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Bảng (Table)</span>
                </button>
                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setChartModalType('column')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <BarChart3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Biểu đồ Cột</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setChartModalType('line')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <LineChart className="w-3.5 h-3.5 text-[#2563EB]" />
                  <span>Biểu đồ Đường</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setChartModalType('pie')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <PieChart className="w-3.5 h-3.5 text-[#D97706]" />
                  <span>Biểu đồ Tròn</span>
                </button>
                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => imageInputRef.current?.click()}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <ImagePlus className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Ảnh</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'shapes' ? null : 'shapes'))
                  }
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold disabled:opacity-40 transition-colors ${
                    openPopover === 'shapes'
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]'
                  }`}
                >
                  <Shapes className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Hình khối</span>
                  <ChevronDown className="w-3 h-3 text-slate-400" />
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={handleInsertTextBox}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <FileText className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Hộp văn bản</span>
                </button>
                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />
                <button
                  type="button"
                  disabled={isEditBlocked || anchorOriginalRow === null}
                  onClick={() => handleOpenCellMeta('hyperlink')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <Link2 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Liên kết</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked || anchorOriginalRow === null}
                  onClick={() => handleOpenCellMeta('comment')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <MessageSquarePlus className="w-3.5 h-3.5 text-rose-500" />
                  <span>Ghi chú ô</span>
                </button>
              </div>
            )}

            {/* TAB 4: VẼ (DRAW) */}
            {activeRibbonTab === 'draw' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  onClick={() => setActiveDrawTool(null)}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors ${
                    activeDrawTool === null
                      ? 'bg-[#107C41] border-[#107C41] text-white'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <MousePointer2 className="w-3.5 h-3.5" />
                  <span>Chọn ô (Tắt vẽ)</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setActiveDrawTool('pen')}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold disabled:opacity-40 transition-colors ${
                    activeDrawTool === 'pen'
                      ? 'bg-[#ECFDF5] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <PenTool className="w-3.5 h-3.5" />
                  <span>Bút vẽ (Pen)</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setActiveDrawTool('highlighter')}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold disabled:opacity-40 transition-colors ${
                    activeDrawTool === 'highlighter'
                      ? 'bg-[#ECFDF5] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <Highlighter className="w-3.5 h-3.5" />
                  <span>Bút tô sáng</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() => setActiveDrawTool('eraser')}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold disabled:opacity-40 transition-colors ${
                    activeDrawTool === 'eraser'
                      ? 'bg-rose-50 border-rose-400 text-rose-700'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-rose-50'
                  }`}
                >
                  <Eraser className="w-3.5 h-3.5" />
                  <span>Cục tẩy</span>
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <div className="flex items-center gap-1">
                  {DRAW_COLOR_SWATCHES.map((hex) => (
                    <button
                      key={hex}
                      type="button"
                      disabled={isEditBlocked}
                      onClick={() => setDrawColor(hex)}
                      style={{ backgroundColor: hex }}
                      title={`Màu mực ${hex}`}
                      className={`w-5 h-5 rounded-full border transition-transform ${
                        drawColor === hex
                          ? 'ring-2 ring-[#107C41] ring-offset-1 scale-110 border-white'
                          : 'border-black/20 hover:scale-110'
                      }`}
                    />
                  ))}
                </div>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <div className="inline-flex items-center rounded-lg bg-slate-50 border border-slate-200/80 p-0.5">
                  {(
                    [
                      { w: 2, label: 'Mảnh' },
                      { w: 4, label: 'Vừa' },
                      { w: 7, label: 'Đậm' },
                    ] as const
                  ).map((item) => (
                    <button
                      key={item.w}
                      type="button"
                      onClick={() => setDrawStrokeWidth(item.w)}
                      className={`px-2 py-1 rounded text-[11px] font-semibold ${
                        drawStrokeWidth === item.w
                          ? 'bg-white text-[#107C41] shadow-sm'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  disabled={
                    isEditBlocked || (activeSheet.drawings?.length ?? 0) === 0
                  }
                  onClick={handleClearDrawings}
                  className="h-8 px-2.5 rounded-lg border border-rose-200 bg-rose-50 hover:bg-rose-100 text-xs font-semibold text-rose-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>Xóa nét vẽ ({activeSheet.drawings?.length ?? 0})</span>
                </button>
              </div>
            )}

            {/* TAB 5: BỐ CỤC (PAGE LAYOUT) */}
            {activeRibbonTab === 'layout' && (
              <div className="flex items-center gap-2 w-max">
                <div className="inline-flex items-center gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">
                    Lề in:
                  </span>
                  {(
                    [
                      { id: 'normal', label: 'Thường' },
                      { id: 'narrow', label: 'Hẹp' },
                      { id: 'wide', label: 'Rộng' },
                    ] as const
                  ).map((m) => (
                    <button
                      key={m.id}
                      type="button"
                      disabled={isEditBlocked}
                      onClick={() => handleUpdatePageSetup({ margins: m.id })}
                      className={`h-7 px-2 rounded-md border text-xs font-semibold disabled:opacity-40 ${
                        activePageSetup.margins === m.id
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {m.label}
                    </button>
                  ))}
                </div>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <div className="inline-flex items-center gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">
                    Hướng giấy:
                  </span>
                  {(
                    [
                      { id: 'portrait', label: 'Dọc' },
                      { id: 'landscape', label: 'Ngang' },
                    ] as const
                  ).map((o) => (
                    <button
                      key={o.id}
                      type="button"
                      disabled={isEditBlocked}
                      onClick={() =>
                        handleUpdatePageSetup({ orientation: o.id })
                      }
                      className={`h-7 px-2.5 rounded-md border text-xs font-semibold disabled:opacity-40 ${
                        activePageSetup.orientation === o.id
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {o.label}
                    </button>
                  ))}
                </div>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <div className="inline-flex items-center gap-1">
                  <span className="text-[11px] font-semibold text-slate-500">
                    Khổ giấy:
                  </span>
                  {(['A4', 'A5', 'Letter'] as const).map((sz) => (
                    <button
                      key={sz}
                      type="button"
                      disabled={isEditBlocked}
                      onClick={() => handleUpdatePageSetup({ paperSize: sz })}
                      className={`h-7 px-2 rounded-md border text-xs font-semibold disabled:opacity-40 ${
                        activePageSetup.paperSize === sz
                          ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                          : 'bg-white border-slate-200/80 text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {sz}
                    </button>
                  ))}
                </div>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <button
                  type="button"
                  disabled={isEditBlocked || !normalizedRange}
                  onClick={() => {
                    handleUpdatePageSetup({ printArea: rangeLabel });
                    showTransientFeedback(`Đã đặt vùng in: ${rangeLabel}.`);
                  }}
                  className="h-7 px-2.5 rounded-md border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 disabled:opacity-40"
                >
                  Đặt vùng in ({rangeLabel})
                </button>
                {activePageSetup.printArea && (
                  <button
                    type="button"
                    disabled={isEditBlocked}
                    onClick={() => {
                      handleUpdatePageSetup({ printArea: undefined });
                      showTransientFeedback('Đã xóa vùng in tùy chỉnh.');
                    }}
                    className="h-7 px-2 rounded-md border border-rose-200 bg-rose-50 text-xs font-semibold text-rose-700 disabled:opacity-40"
                  >
                    Xóa vùng in ({activePageSetup.printArea})
                  </button>
                )}

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() =>
                    handleUpdatePageSetup({
                      printGridlines: activePageSetup.printGridlines === false,
                    })
                  }
                  className={`h-7 px-2.5 rounded-md border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 ${
                    activePageSetup.printGridlines !== false
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-600'
                  }`}
                >
                  <Grid3x3 className="w-3.5 h-3.5" />
                  <span>In đường lưới</span>
                </button>

                <button
                  type="button"
                  onClick={handlePrintSheet}
                  className="h-7 px-2.5 rounded-md bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold inline-flex items-center gap-1.5"
                >
                  <Printer className="w-3.5 h-3.5" />
                  <span>In ngay</span>
                </button>
              </div>
            )}

            {/* TAB 6: CÔNG THỨC (FORMULAS) */}
            {activeRibbonTab === 'formulas' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  disabled={isEditBlocked || !normalizedRange}
                  onClick={() => setFormulaModalCategory('all')}
                  className="h-8 px-3 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <Calculator className="w-3.5 h-3.5" />
                  <span>Chèn hàm (fx)</span>
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                {(
                  [
                    { cat: 'autosum', label: 'AutoSum (Σ)' },
                    { cat: 'financial', label: 'Tài chính (PMT/FV/PV)' },
                    { cat: 'logic', label: 'Logic' },
                    { cat: 'text', label: 'Văn bản' },
                    { cat: 'datetime', label: 'Ngày & Giờ' },
                    { cat: 'lookup', label: 'Tra cứu' },
                    { cat: 'math', label: 'Toán học' },
                  ] as const
                ).map(({ cat, label }) => (
                  <button
                    key={cat}
                    type="button"
                    disabled={isEditBlocked || !normalizedRange}
                    onClick={() => setFormulaModalCategory(cat)}
                    className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] disabled:opacity-40 transition-colors"
                  >
                    {label}
                  </button>
                ))}

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  onClick={handleRecalculateSheet}
                  className="h-8 px-2.5 rounded-lg border border-[#107C41]/30 bg-[#ECFDF5] hover:bg-[#D1FAE5] text-xs font-semibold text-[#107C41] inline-flex items-center gap-1.5 transition-colors"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Tính toán lại trang tính</span>
                </button>
              </div>
            )}

            {/* TAB 7: DỮ LIỆU (DATA) */}
            {activeRibbonTab === 'data' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  disabled={anchorCol === null}
                  onClick={() =>
                    anchorCol !== null && handleSetSort(anchorCol, 'asc')
                  }
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-40 ${
                    sortConfig?.colIndex === anchorCol &&
                    sortConfig?.direction === 'asc'
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <ArrowUp className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>A → Z</span>
                </button>
                <button
                  type="button"
                  disabled={anchorCol === null}
                  onClick={() =>
                    anchorCol !== null && handleSetSort(anchorCol, 'desc')
                  }
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-40 ${
                    sortConfig?.colIndex === anchorCol &&
                    sortConfig?.direction === 'desc'
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <ArrowDown className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Z → A</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setCustomSortColIdx(anchorCol ?? 0);
                    setOpenPopover((p) =>
                      p === 'customSort' ? null : 'customSort'
                    );
                  }}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold ${
                    openPopover === 'customSort'
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <span>Sắp xếp tùy chỉnh</span>
                  <ChevronDown className="w-3 h-3 text-slate-400" />
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  onClick={handleToggleAutoFilter}
                  className={`h-8 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors ${
                    activeSheet.autoFilterEnabled || hasColumnFilters
                      ? 'bg-[#107C41] hover:bg-[#0B5E31] border-[#107C41] text-white'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <ListFilter className="w-3.5 h-3.5" />
                  <span>Bộ lọc tự động</span>
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={handleRemoveDuplicates}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40"
                >
                  <Rows3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Xóa trùng lặp</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked || anchorCol === null}
                  onClick={() => setIsTextToColsModalOpen(true)}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40"
                >
                  <Columns3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Tách cột</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked || !normalizedRange}
                  onClick={() => setIsValidationModalOpen(true)}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40"
                >
                  <ListChecks className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Xác thực dữ liệu</span>
                </button>

                {numericColumnIndices.length > 0 && visibleRows.length > 0 && (
                  <>
                    <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />
                    <button
                      type="button"
                      onClick={handleQuickSelectNumericColumn}
                      className="h-8 px-2.5 rounded-lg bg-[#ECFDF5] hover:bg-[#D1FAE5] border border-[#107C41]/25 text-[#107C41] text-xs font-semibold inline-flex items-center gap-1"
                    >
                      <Sigma className="w-3.5 h-3.5" />
                      <span>Chọn cột số</span>
                    </button>
                  </>
                )}
              </div>
            )}

            {/* TAB 8: CỘNG TÁC (COLLABORATION) */}
            {activeRibbonTab === 'collab' && (
              <div className="flex items-center gap-2 w-max">
                <button
                  type="button"
                  onClick={() => onOpenModal('share', buildLiveDocument(sheets))}
                  className="h-8 px-3 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold inline-flex items-center gap-1.5 transition-colors"
                >
                  <Share2 className="w-3.5 h-3.5" />
                  <span>Chia sẻ tài liệu</span>
                </button>
                <button
                  type="button"
                  disabled={!workspaceSync.onOpenVersionHistory}
                  onClick={() => workspaceSync.onOpenVersionHistory?.()}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors"
                >
                  <History className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Lịch sử phiên bản</span>
                </button>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <button
                  type="button"
                  disabled={isEditBlocked || anchorOriginalRow === null}
                  onClick={() => handleOpenCellMeta('comment')}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 disabled:opacity-40"
                >
                  <MessageSquarePlus className="w-3.5 h-3.5 text-rose-500" />
                  <span>Ghi chú ô ({rangeLabel})</span>
                </button>
                <button
                  type="button"
                  onClick={() =>
                    setOpenPopover((p) =>
                      p === 'commentsList' ? null : 'commentsList'
                    )
                  }
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-slate-50 hover:bg-[#ECFDF5] text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5"
                >
                  <span>
                    Danh sách ghi chú (
                    {Object.keys(activeSheet.cellComments ?? {}).length})
                  </span>
                  <ChevronDown className="w-3 h-3 text-slate-400" />
                </button>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200/80 text-xs text-slate-600">
                  <Users className="w-3.5 h-3.5 text-[#107C41]" />
                  <span className="font-semibold">
                    {collaborators && collaborators.length > 0
                      ? `${collaborators.length} người trực tuyến`
                      : 'Chỉ mình bạn trực tuyến'}
                  </span>
                  {collaborators && collaborators.length > 0 && (
                    <div className="flex items-center -space-x-1.5">
                      {collaborators.slice(0, 4).map((peer) => (
                        <span
                          key={peer.sessionId}
                          style={{ backgroundColor: peer.color }}
                          title={`${peer.displayName}${peer.cellCoord ? ` (${peer.cellCoord})` : ''}`}
                          className="w-5 h-5 rounded-full text-[10px] font-bold text-white inline-flex items-center justify-center ring-2 ring-white"
                        >
                          {peer.displayName.slice(0, 1).toUpperCase()}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* TAB 9: BẢO VỆ (PROTECT) */}
            {activeRibbonTab === 'protect' && (
              <div className="flex items-center gap-2 w-max">
                <button
                  type="button"
                  disabled={isReadOnly}
                  onClick={handleToggleProtectSheet}
                  className={`h-8 px-3 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors ${
                    isSheetProtected
                      ? 'bg-amber-500 hover:bg-amber-600 border-amber-500 text-white'
                      : 'bg-[#107C41] hover:bg-[#0B5E31] border-[#107C41] text-white'
                  }`}
                >
                  {isSheetProtected ? (
                    <>
                      <ShieldOff className="w-3.5 h-3.5" />
                      <span>Mở khóa trang tính</span>
                    </>
                  ) : (
                    <>
                      <ShieldCheck className="w-3.5 h-3.5" />
                      <span>Bảo vệ trang tính (Protect Sheet)</span>
                    </>
                  )}
                </button>

                <span className="h-6 w-px bg-slate-200/80 shrink-0" />

                <button
                  type="button"
                  disabled={isReadOnly || !normalizedRange}
                  onClick={handleToggleLockSelectedRange}
                  className={`h-8 px-2.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 transition-colors ${
                    activeSheet.lockedRanges?.includes(rangeLabel)
                      ? 'bg-amber-50 border-amber-300 text-amber-800 hover:bg-amber-100'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]'
                  }`}
                >
                  {activeSheet.lockedRanges?.includes(rangeLabel) ? (
                    <>
                      <Unlock className="w-3.5 h-3.5" />
                      <span>Mở khóa vùng {rangeLabel}</span>
                    </>
                  ) : (
                    <>
                      <Lock className="w-3.5 h-3.5 text-[#107C41]" />
                      <span>Khóa vùng chọn ({rangeLabel})</span>
                    </>
                  )}
                </button>

                {activeSheet.lockedRanges &&
                  activeSheet.lockedRanges.length > 0 && (
                    <div className="flex items-center gap-1">
                      <span className="text-[11px] font-semibold text-slate-500">
                        Vùng đã khóa:
                      </span>
                      {activeSheet.lockedRanges.map((rStr) => (
                        <span
                          key={rStr}
                          className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-amber-50 border border-amber-200 text-amber-800 text-[11px] font-semibold"
                        >
                          <Shield className="w-3 h-3" />
                          <span>{rStr}</span>
                          {!isReadOnly && (
                            <button
                              type="button"
                              onClick={() => handleUnlockSpecificRange(rStr)}
                              title={`Mở khóa ${rStr}`}
                              className="hover:text-rose-600"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          )}
                        </span>
                      ))}
                    </div>
                  )}
              </div>
            )}

            {/* TAB 10: XEM (VIEW) */}
            {activeRibbonTab === 'view' && (
              <div className="flex items-center gap-1.5 w-max">
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() =>
                    handleSetFreezePanes(1, activeSheet.frozenCols ?? 0)
                  }
                  className={`h-8 px-2.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 ${
                    (activeSheet.frozenRows ?? 0) > 0
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <Snowflake className="w-3.5 h-3.5" />
                  <span>Cố định dòng đầu</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={() =>
                    handleSetFreezePanes(activeSheet.frozenRows ?? 0, 1)
                  }
                  className={`h-8 px-2.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 ${
                    (activeSheet.frozenCols ?? 0) > 0
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-700 hover:bg-[#ECFDF5]'
                  }`}
                >
                  <Snowflake className="w-3.5 h-3.5" />
                  <span>Cố định cột đầu</span>
                </button>
                <button
                  type="button"
                  disabled={isEditBlocked || !activeFreezeLabel}
                  onClick={() => handleSetFreezePanes(0, 0)}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-rose-50 hover:text-rose-600 text-xs font-semibold text-slate-700 disabled:opacity-40"
                >
                  Bỏ cố định
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  disabled={isEditBlocked}
                  onClick={handleToggleShowGridlines}
                  className={`h-8 px-2.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 disabled:opacity-40 ${
                    showGridlines
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-600'
                  }`}
                >
                  <Grid3x3 className="w-3.5 h-3.5" />
                  <span>Đường lưới</span>
                </button>
                <button
                  type="button"
                  onClick={() => setShowFormulaBar((v) => !v)}
                  className={`h-8 px-2.5 rounded-lg border text-xs font-semibold inline-flex items-center gap-1.5 ${
                    showFormulaBar
                      ? 'bg-[#ECFDF5] border-[#107C41]/50 text-[#107C41]'
                      : 'bg-white border-slate-200/80 text-slate-600'
                  }`}
                >
                  <span className="font-serif italic font-bold">fx</span>
                  <span>Thanh công thức</span>
                </button>

                <span className="h-6 w-px bg-slate-200/80 mx-0.5 shrink-0" />

                <button
                  type="button"
                  onClick={() =>
                    setZoomPercent((z) => Math.max(10, z - 10))
                  }
                  title="Thu nhỏ (-10%)"
                  className="h-8 px-2 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-slate-700 inline-flex items-center"
                >
                  <ZoomOut className="w-3.5 h-3.5" />
                </button>
                <span className="px-2 text-xs font-bold text-[#107C41] tabular-nums">
                  {zoomPercent}%
                </span>
                <button
                  type="button"
                  onClick={() =>
                    setZoomPercent((z) => Math.min(300, z + 10))
                  }
                  title="Phóng to (+10%)"
                  className="h-8 px-2 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-slate-700 inline-flex items-center"
                >
                  <ZoomIn className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={() => setZoomPercent(100)}
                  className="h-8 px-2.5 rounded-lg border border-slate-200/80 bg-white hover:bg-[#ECFDF5] hover:text-[#107C41] text-xs font-semibold text-slate-700"
                >
                  100%
                </button>
                <button
                  type="button"
                  onClick={handleFitToWidth}
                  className="h-8 px-2.5 rounded-lg border border-[#107C41]/35 bg-[#ECFDF5] hover:bg-[#D1FAE5] text-xs font-semibold text-[#107C41] inline-flex items-center gap-1"
                >
                  <Maximize2 className="w-3.5 h-3.5" />
                  <span>Vừa màn hình (Fit)</span>
                </button>
              </div>
            )}

            {/* TAB 11: AI TOOLS */}
            {activeRibbonTab === 'ai' && (
              <div className="flex items-center gap-2 w-max">
                <button
                  type="button"
                  onClick={() => setAiModalTab('summary')}
                  className="h-8 px-3 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold inline-flex items-center gap-1.5 shadow-sm transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Tóm tắt số liệu</span>
                </button>
                <button
                  type="button"
                  onClick={() => setAiModalTab('formula')}
                  className="h-8 px-3 rounded-lg border border-[#107C41]/35 bg-[#ECFDF5] hover:bg-[#D1FAE5] text-[#107C41] text-xs font-semibold inline-flex items-center gap-1.5 transition-colors"
                >
                  <Wand2 className="w-3.5 h-3.5" />
                  <span>Tạo công thức từ tiếng Việt</span>
                </button>
                <button
                  type="button"
                  onClick={() => setAiModalTab('audit')}
                  className="h-8 px-3 rounded-lg border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-semibold inline-flex items-center gap-1.5 transition-colors"
                >
                  <AlertTriangle className="w-3.5 h-3.5" />
                  <span>Phát hiện ô lỗi / trùng lặp</span>
                </button>
              </div>
            )}
          </div>
        )}

        {/* Ribbon Dropdown Popovers */}
        {openPopover === 'fillColor' && (
          <div className="absolute left-3 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-3 w-56">
            <p className="text-[11px] font-bold uppercase tracking-wide text-slate-400 mb-2">
              Màu nền ô
            </p>
            <div className="grid grid-cols-8 gap-1.5 mb-2.5">
              {FILL_COLOR_SWATCHES.map((hex) => (
                <button
                  key={hex}
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => handleSetFillColor(hex)}
                  style={{ backgroundColor: hex }}
                  title={hex}
                  className="w-5 h-5 rounded border border-slate-300 hover:scale-110 transition-transform"
                />
              ))}
            </div>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFillColor(undefined)}
              className="w-full py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-600 hover:bg-slate-50"
            >
              Không tô màu (Mặc định)
            </button>
          </div>
        )}

        {openPopover === 'textColor' && (
          <div className="absolute left-16 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-3 w-56">
            <p className="text-[11px] font-bold uppercase tracking-wide text-slate-400 mb-2">
              Màu chữ
            </p>
            <div className="flex flex-wrap gap-1.5 mb-2.5">
              {TEXT_COLOR_SWATCHES.map((hex) => (
                <button
                  key={hex}
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => handleSetTextColor(hex)}
                  style={{ backgroundColor: hex }}
                  title={hex}
                  className="w-6 h-6 rounded-full border border-slate-300 hover:scale-110 transition-transform"
                />
              ))}
            </div>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetTextColor(undefined)}
              className="w-full py-1.5 rounded-lg border border-slate-200 text-xs font-semibold text-slate-600 hover:bg-slate-50"
            >
              Màu chữ tự động
            </button>
          </div>
        )}

        {openPopover === 'borders' && (
          <div className="absolute left-28 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-56">
            {(
              [
                { id: 'all', label: 'Tất cả viền (All Borders)' },
                { id: 'outside', label: 'Viền ngoài (Outside Borders)' },
                { id: 'thick-box', label: 'Khung ngoài đậm (Thick Box)' },
                { id: 'bottom', label: 'Viền dưới (Bottom Border)' },
                { id: 'none', label: 'Không viền (No Border)' },
              ] as const
            ).map((item) => (
              <button
                key={item.id}
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleApplyBorderPreset(item.id)}
                className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41] transition-colors"
              >
                {item.label}
              </button>
            ))}
          </div>
        )}

        {openPopover === 'case' && (
          <div className="absolute left-32 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-52">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleTransformTextCase('upper')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              CHỮ IN HOA (UPPER)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleTransformTextCase('lower')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              chữ thường (lower)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleTransformTextCase('proper')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Viết Hoa Đầu Từ (Proper)
            </button>
          </div>
        )}

        {openPopover === 'textDir' && (
          <div className="absolute left-52 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-52">
            {(
              [
                { id: 'horizontal', label: 'Ngang (Mặc định)' },
                { id: 'vertical', label: 'Chữ dọc (Vertical)' },
                { id: 'up45', label: 'Nghiêng lên 45°' },
                { id: 'down45', label: 'Nghiêng xuống 45°' },
              ] as const
            ).map((item) => (
              <button
                key={item.id}
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleSetTextRotation(item.id)}
                className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
              >
                {item.label}
              </button>
            ))}
          </div>
        )}

        {openPopover === 'autoSum' && (
          <div className="absolute left-64 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-52">
            {(
              [
                { fn: 'SUM', label: 'Tổng (SUM)' },
                { fn: 'AVERAGE', label: 'Trung bình (AVERAGE)' },
                { fn: 'COUNT', label: 'Đếm số (COUNT)' },
                { fn: 'MAX', label: 'Lớn nhất (MAX)' },
                { fn: 'MIN', label: 'Nhỏ nhất (MIN)' },
              ] as const
            ).map(({ fn, label }) => (
              <button
                key={fn}
                type="button"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => handleAutoSum(fn)}
                className="w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41] transition-colors"
              >
                <span>{label}</span>
                <span className="text-[10px] font-mono text-slate-400">
                  ={fn}
                </span>
              </button>
            ))}
          </div>
        )}

        {openPopover === 'fill' && (
          <div className="absolute left-72 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-52">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleFillDirection('down')}
              className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              <ArrowDownToLine className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Điền xuống dưới (Fill Down)</span>
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleFillDirection('right')}
              className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              <ArrowRightToLine className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Điền sang phải (Fill Right)</span>
            </button>
          </div>
        )}

        {openPopover === 'clear' && (
          <div className="absolute left-80 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-48">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleClearAction('contents')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Xóa nội dung
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleClearAction('formats')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Xóa định dạng
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleClearAction('all')}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-semibold text-rose-600 hover:bg-rose-50"
            >
              Xóa tất cả
            </button>
          </div>
        )}

        {openPopover === 'shapes' && (
          <div className="absolute left-40 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-2 w-60 grid grid-cols-2 gap-1.5">
            {(
              [
                { kind: 'rect', label: 'Hình chữ nhật' },
                { kind: 'roundRect', label: 'Bo góc' },
                { kind: 'ellipse', label: 'Hình tròn / Elip' },
                { kind: 'arrowRight', label: 'Mũi tên phải' },
                { kind: 'star', label: 'Ngôi sao' },
                { kind: 'callout', label: 'Hộp thoại' },
              ] as const
            ).map(({ kind, label }) => (
              <button
                key={kind}
                type="button"
                onClick={() => handleInsertShape(kind)}
                className="px-2.5 py-2 rounded-lg border border-slate-200/80 bg-slate-50 hover:bg-[#ECFDF5] hover:border-[#107C41]/40 hover:text-[#107C41] text-xs font-semibold text-slate-700 text-left transition-colors"
              >
                {label}
              </button>
            ))}
          </div>
        )}

        {openPopover === 'customSort' && (
          <div className="absolute left-24 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-3 w-64 space-y-2.5">
            <p className="text-xs font-bold text-slate-800">
              Sắp xếp tùy chỉnh
            </p>
            <label className="block space-y-1">
              <span className="text-[11px] font-semibold text-slate-500">
                Chọn cột sắp xếp
              </span>
              <select
                value={customSortColIdx}
                onChange={(e) => setCustomSortColIdx(Number(e.target.value))}
                className="w-full h-8 px-2 rounded-lg border border-slate-200 text-xs font-medium"
              >
                {activeSheet.columns.map((_, idx) => (
                  <option key={idx} value={idx}>
                    Cột {columnIndexToLetter(idx)}
                    {hasHeaderRow && activeSheet.rows[0]?.[idx]
                      ? ` — ${String(activeSheet.rows[0][idx]).slice(0, 18)}`
                      : ''}
                  </option>
                ))}
              </select>
            </label>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setCustomSortDir('asc')}
                className={`flex-1 py-1.5 rounded-lg border text-xs font-semibold ${
                  customSortDir === 'asc'
                    ? 'bg-[#ECFDF5] border-[#107C41] text-[#107C41]'
                    : 'bg-white border-slate-200 text-slate-600'
                }`}
              >
                Tăng dần (A→Z)
              </button>
              <button
                type="button"
                onClick={() => setCustomSortDir('desc')}
                className={`flex-1 py-1.5 rounded-lg border text-xs font-semibold ${
                  customSortDir === 'desc'
                    ? 'bg-[#ECFDF5] border-[#107C41] text-[#107C41]'
                    : 'bg-white border-slate-200 text-slate-600'
                }`}
              >
                Giảm dần (Z→A)
              </button>
            </div>
            <button
              type="button"
              onClick={() => {
                handleSetSort(customSortColIdx, customSortDir);
                setOpenPopover(null);
              }}
              className="w-full py-1.5 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold"
            >
              Áp dụng sắp xếp
            </button>
          </div>
        )}

        {openPopover === 'commentsList' && (
          <div className="absolute left-40 top-full mt-1 z-50 bg-white rounded-xl shadow-2xl border border-slate-200 p-2.5 w-72 max-h-64 overflow-y-auto space-y-1.5">
            <p className="text-[11px] font-bold uppercase tracking-wide text-slate-400 px-1">
              Ghi chú trên &ldquo;{activeSheet.name}&rdquo;
            </p>
            {Object.entries(activeSheet.cellComments ?? {}).length === 0 ? (
              <p className="text-xs text-slate-400 px-1 py-3 text-center">
                Chưa có ghi chú ô nào trên trang tính này.
              </p>
            ) : (
              Object.entries(activeSheet.cellComments ?? {}).map(
                ([cellKey, commentText]) => {
                  const [rStr, cStr] = cellKey.split(':');
                  const r = Number(rStr);
                  const c = Number(cStr);
                  return (
                    <button
                      key={cellKey}
                      type="button"
                      onClick={() => {
                        handleNavigateToCell(r, c);
                        setOpenPopover(null);
                      }}
                      className="w-full text-left p-2 rounded-lg border border-slate-100 hover:bg-[#ECFDF5] transition-colors"
                    >
                      <span className="text-[11px] font-mono font-bold text-rose-600">
                        {formatCellRef(r, c)}
                      </span>
                      <p className="text-xs text-slate-700 line-clamp-2 mt-0.5">
                        {commentText}
                      </p>
                    </button>
                  );
                }
              )
            )}
          </div>
        )}

        {openPopover === 'freeze' && (
          <div className="absolute left-3 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-60">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(1, activeSheet.frozenCols ?? 0)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Cố định dòng đầu tiên (Freeze Top Row)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(activeSheet.frozenRows ?? 0, 1)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Cố định cột đầu tiên (Freeze First Column)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(1, 1)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#ECFDF5] hover:text-[#107C41]"
            >
              Cố định cả dòng đầu &amp; cột đầu
            </button>
            <div className="my-1 border-t border-slate-100" />
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(0, 0)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-semibold text-rose-600 hover:bg-rose-50"
            >
              Bỏ cố định khung (Unfreeze)
            </button>
          </div>
        )}
      </div>

      {/* Protected Sheet Warning Banner */}
      {isSheetProtected && (
        <div className="shrink-0 px-3 sm:px-4 py-1.5 bg-amber-50 border-b border-amber-200 text-amber-900 text-xs flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <ShieldCheck className="w-4 h-4 text-amber-600 shrink-0" />
            <span className="truncate font-medium">
              Trang tính &ldquo;{activeSheet.name}&rdquo; đang được bảo vệ — mọi thao tác sửa ô đã bị khóa.
            </span>
          </div>
          {!isReadOnly && (
            <button
              type="button"
              onClick={handleToggleProtectSheet}
              className="px-2.5 py-1 rounded-md bg-amber-600 hover:bg-amber-700 text-white text-[11px] font-semibold whitespace-nowrap shrink-0"
            >
              Mở khóa ngay
            </button>
          )}
        </div>
      )}

      {/* Active Drawing Mode Helper Banner */}
      {activeDrawTool !== null && !isEditBlocked && (
        <div className="shrink-0 px-3 sm:px-4 py-1 bg-[#ECFDF5] border-b border-[#107C41]/20 text-[#107C41] text-xs flex items-center justify-between gap-2">
          <span className="font-medium truncate">
            Đang bật công cụ{' '}
            <strong>
              {activeDrawTool === 'pen'
                ? 'Bút vẽ'
                : activeDrawTool === 'highlighter'
                  ? 'Bút tô sáng'
                  : 'Cục tẩy nét vẽ'}
            </strong>{' '}
            — vẽ trực tiếp lên bảng tính.
          </span>
          <button
            type="button"
            onClick={() => setActiveDrawTool(null)}
            className="px-2.5 py-0.5 rounded-md bg-[#107C41] hover:bg-[#0B5E31] text-white text-[11px] font-semibold whitespace-nowrap shrink-0"
          >
            Tắt vẽ &amp; Chọn ô
          </button>
        </div>
      )}

      {/* Name Box (A1) + Formula Bar (fx) + Search Bar */}
      {showFormulaBar && (
      <div className="shrink-0 bg-white border-b border-slate-200/80 px-2.5 sm:px-4 py-1.5 relative z-20">
        <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
          {/* Editable Name Box */}
          <input
            type="text"
            value={isNameBoxFocused ? nameBoxDraft : rangeLabel}
            onFocus={(e) => {
              setIsNameBoxFocused(true);
              setNameBoxDraft(rangeLabel);
              e.currentTarget.select();
            }}
            onChange={(e) => setNameBoxDraft(e.target.value)}
            onBlur={handleNameBoxCommit}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                handleNameBoxCommit();
                e.currentTarget.blur();
              } else if (e.key === 'Escape') {
                setIsNameBoxFocused(false);
                setNameBoxDraft(rangeLabel);
                e.currentTarget.blur();
              }
            }}
            aria-label="Hộp tên ô (Name Box)"
            title="Nhập tọa độ ô (vd: C12 hoặc B2:D8) và nhấn Enter để nhảy nhanh"
            className="w-[78px] sm:w-[92px] px-2 py-1 rounded-lg bg-slate-50 border border-slate-200 focus:bg-white focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15 focus:outline-none text-xs font-semibold text-[#107C41] text-center tabular-nums shrink-0 transition-all"
          />

          {/* Formula Bar (fx) + Autocomplete Dropdown */}
          <div className="relative flex-1 min-w-[180px]">
            <div className="flex items-center gap-2 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200 focus-within:bg-white focus-within:border-[#107C41] focus-within:ring-2 focus-within:ring-[#107C41]/15 transition-all">
              <span
                className="font-serif italic font-bold text-[#107C41]/75 select-none shrink-0 pr-2 border-r border-slate-200 text-sm"
                aria-hidden="true"
              >
                fx
              </span>
              <input
                ref={formulaInputRef}
                type="text"
                value={formulaInputValue}
                readOnly={isReadOnly}
                disabled={anchorOriginalRow === null || anchorCol === null}
                onFocus={handleFormulaFocus}
                onChange={handleFormulaChange}
                onKeyDown={handleFormulaKeyDown}
                onBlur={handleFormulaBlur}
                placeholder={
                  isReadOnly
                    ? 'Chế độ chỉ xem — chọn ô để xem giá trị hoặc công thức'
                    : viewMode === 'original'
                      ? 'Đang xem bản gốc — bấm vào đây hoặc bấm đúp ô để chỉnh sửa ngay...'
                      : 'Nhập giá trị hoặc gõ = để gợi ý 40 hàm Excel (SUM, IF, VLOOKUP, INDEX...)...'
                }
                aria-label="Thanh công thức fx"
                className="w-full py-0.5 bg-transparent text-xs sm:text-[13px] font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none disabled:opacity-50"
              />
              {isFormulaText(anchorRawValue) && (
                <span className="hidden md:inline-flex items-center px-2 py-0.5 rounded-md bg-[#ECFDF5] border border-[#107C41]/25 text-[11px] font-semibold text-[#107C41] tabular-nums shrink-0 whitespace-nowrap">
                  = {formatCellDisplay(anchorEvaluatedValue, anchorStyle.numFmt)}
                </span>
              )}
            </div>

            {/* Live Formula Autocomplete Popover */}
            {formulaSuggestions.length > 0 && (
              <div className="absolute left-0 right-0 sm:right-auto sm:min-w-[360px] top-full mt-1 z-50 bg-white rounded-xl shadow-2xl border border-slate-200 py-1 overflow-hidden">
                <div className="px-2.5 py-1 bg-[#F8FAFC] border-b border-slate-100 flex items-center justify-between text-[10.5px] font-semibold text-slate-500">
                  <span>Gợi ý hàm Excel (nhấn Tab hoặc chọn)</span>
                  <span className="text-[#107C41]">
                    {formulaSuggestions.length} kết quả
                  </span>
                </div>
                {formulaSuggestions.map((item, idx) => (
                  <button
                    key={item.name}
                    type="button"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      handleApplyFormulaSuggestion(item.name);
                    }}
                    className={`w-full text-left px-3 py-1.5 flex flex-col hover:bg-[#ECFDF5] transition-colors ${
                      idx === 0 ? 'bg-[#ECFDF5]/60' : ''
                    }`}
                  >
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-xs font-bold font-mono text-[#107C41]">
                        {item.name}
                      </span>
                      <span className="text-[10.5px] font-mono text-slate-400 truncate">
                        {item.syntax}
                      </span>
                    </div>
                    <span className="text-[11px] text-slate-600 truncate">
                      {item.description}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>

          {/* Realtime Search Filter */}
          <div className="relative w-full sm:w-52 lg:w-60 shrink-0">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
              placeholder="Tìm & lọc nhanh..."
              aria-label="Lọc dữ liệu trong bảng tính"
              className="w-full pl-8 pr-7 py-1 rounded-lg bg-slate-50 border border-slate-200 text-xs sm:text-[13px] text-slate-800 placeholder:text-slate-400 focus:outline-none focus:bg-white focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15 transition-all"
            />
            {searchQuery.length > 0 && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                aria-label="Xóa từ khóa lọc"
                className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 rounded-md text-slate-400 hover:text-slate-700 hover:bg-slate-200/70 transition-colors"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>
      )}

      {/* Main Grid Viewport (Zoomed 10%-300% via CSS transform scale + sizer, with CustomVerticalScrollbar) */}
      <div className="relative flex-1 min-h-0 flex overflow-hidden">
        <div
          ref={gridContainerRef}
          tabIndex={0}
          onKeyDown={handleGridKeyDown}
          className="w-full h-full overflow-auto scrollbar-none pr-4 bg-white select-none focus:outline-none"
        >
          <div
            style={
              scaledCanvasWidth !== undefined || scaledCanvasHeight !== undefined
                ? {
                    width:
                      scaledCanvasWidth !== undefined
                        ? `${scaledCanvasWidth}px`
                        : undefined,
                    height:
                      scaledCanvasHeight !== undefined
                        ? `${scaledCanvasHeight}px`
                        : undefined,
                    position: 'relative',
                  }
                : { width: 'max-content', minWidth: '100%' }
            }
          >
            <div
              ref={sheetCanvasRef}
              className="relative"
              style={{
                width: 'max-content',
                ...(zoomScale === 1
                  ? {}
                  : {
                      position: 'absolute',
                      top: 0,
                      left: 0,
                      transform: `scale(${zoomScale})`,
                      transformOrigin: 'top left',
                    }),
              }}
            >
              <table
                style={{
                  fontSize: '13px',
                }}
                className="w-max border-separate border-spacing-0"
              >
                <colgroup>
                  <col
                    style={{
                      width: `${rowHeaderWidthPx}px`,
                      minWidth: `${rowHeaderWidthPx}px`,
                    }}
                  />
                  {visibleColIndices.map((colIdx) => {
                    const colW = Math.round(
                      activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH
                    );
                    return (
                      <col
                        key={`colgroup-${colIdx}`}
                        style={{
                          width: `${colW}px`,
                          minWidth: `${colW}px`,
                        }}
                      />
                    );
                  })}
                </colgroup>

                <thead>
                  <tr style={{ height: `${headerHeightPx}px` }}>
                    <th
                      scope="col"
                      style={{
                        width: `${rowHeaderWidthPx}px`,
                        minWidth: `${rowHeaderWidthPx}px`,
                        height: `${headerHeightPx}px`,
                      }}
                      className="sticky left-0 top-0 z-30 p-0 bg-[#F8FAFC] border-b border-r border-[#E2E8F0]"
                    >
                      <button
                        type="button"
                        onClick={handleSelectAllVisible}
                        title="Chọn toàn bộ bảng tính đang hiển thị"
                        aria-label="Chọn toàn bộ bảng tính"
                        className="w-full h-full flex items-center justify-center text-slate-400 hover:text-[#107C41] hover:bg-[#ECFDF5] transition-colors"
                      >
                        <Table2 className="w-3.5 h-3.5" />
                      </button>
                    </th>

                    {visibleColIndices.map((colIdx, visibleColPos) => {
                      const letter = columnIndexToLetter(colIdx);
                      const colW = Math.round(
                        activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH
                      );
                      const isColSelected =
                        normalizedRange !== null &&
                        colIdx >= normalizedRange.minCol &&
                        colIdx <= normalizedRange.maxCol;
                      const isSortedCol = sortConfig?.colIndex === colIdx;
                      const sortDir = isSortedCol ? sortConfig.direction : null;
                      const isColFiltered =
                        (columnFilters[colIdx]?.size ?? 0) > 0;
                      const isFrozenCol = visibleColPos < frozenColsCount;
                      const isLastFrozenCol =
                        frozenColsCount > 0 &&
                        visibleColPos === frozenColsCount - 1;
                      const frozenLeft = frozenColLeftOffsets.get(colIdx);

                      const thStyle: CSSProperties = {
                        width: `${colW}px`,
                        minWidth: `${colW}px`,
                        height: `${headerHeightPx}px`,
                        ...(isFrozenCol && frozenLeft !== undefined
                          ? {
                              left: `${frozenLeft}px`,
                              zIndex: 26,
                            }
                          : {}),
                        ...(isLastFrozenCol
                          ? { boxShadow: '2px 0 0 0 rgba(16,124,65,0.35)' }
                          : {}),
                      };

                      return (
                        <th
                          key={`col-${colIdx}`}
                          scope="col"
                          style={thStyle}
                          onContextMenu={(e) =>
                            handleOpenContextMenu(e, 'col', 0, 0, colIdx)
                          }
                          aria-sort={
                            sortDir === 'asc'
                              ? 'ascending'
                              : sortDir === 'desc'
                                ? 'descending'
                                : 'none'
                          }
                          className={`sticky top-0 z-20 p-0 border-b border-r border-[#E2E8F0] transition-colors ${
                            isColSelected
                              ? 'bg-[#ECFDF5] text-[#107C41]'
                              : isSortedCol || isColFiltered
                                ? 'bg-[#ECFDF5]/80 text-[#0B5E31]'
                                : 'bg-[#F8FAFC] text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          <div
                            className={`w-full h-full flex items-center justify-between pl-2 pr-1 ${
                              isColSelected ? 'border-b-2 border-b-[#107C41]' : ''
                            }`}
                          >
                            <button
                              type="button"
                              onClick={(event) =>
                                handleSelectColumn(colIdx, event.shiftKey)
                              }
                              title={`Bấm để chọn cả cột ${letter} (nhấn chuột phải để chèn/xóa/ẩn cột)`}
                              className="flex-1 h-full text-center text-[11.5px] font-semibold tracking-wider focus:outline-none truncate"
                            >
                              {letter}
                            </button>

                            <button
                              type="button"
                              data-excel-popover="true"
                              onMouseDown={(e) => e.stopPropagation()}
                              onClick={(event) => {
                                event.stopPropagation();
                                const rect =
                                  event.currentTarget.getBoundingClientRect();
                                setFilterSearchText('');
                                setFilterPopover((prev) =>
                                  prev?.colIdx === colIdx
                                    ? null
                                    : {
                                        colIdx,
                                        x: Math.max(
                                          8,
                                          Math.min(
                                            rect.left - 180,
                                            window.innerWidth - 256
                                          )
                                        ),
                                        y: Math.min(
                                          rect.bottom + 4,
                                          window.innerHeight - 330
                                        ),
                                      }
                                );
                              }}
                              aria-label={`Sắp xếp & lọc cột ${letter}`}
                              title={`Sắp xếp & lọc giá trị cột ${letter}`}
                              className={`p-1 rounded transition-colors shrink-0 ${
                                isColFiltered || isSortedCol
                                  ? 'text-[#107C41] bg-white shadow-sm'
                                  : activeSheet.autoFilterEnabled
                                    ? 'text-[#107C41]/80 bg-white/70 hover:bg-white'
                                    : 'text-slate-400 hover:text-slate-700 hover:bg-white/60'
                              }`}
                            >
                              {isColFiltered ? (
                                <Filter className="w-3 h-3" />
                              ) : sortDir === 'asc' ? (
                                <ArrowUp className="w-3 h-3" />
                              ) : sortDir === 'desc' ? (
                                <ArrowDown className="w-3 h-3" />
                              ) : (
                                <ChevronDown className="w-3 h-3" />
                              )}
                            </button>
                          </div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>

                <tbody>
                  {visibleRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={visibleColIndices.length + 1}
                        className="py-16 px-4 text-center bg-white"
                      >
                        <div className="max-w-sm mx-auto flex flex-col items-center gap-2.5">
                          <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-400">
                            <Search className="w-5 h-5" />
                          </div>
                          <p className="text-sm font-semibold text-slate-700">
                            Không tìm thấy dòng dữ liệu nào phù hợp
                          </p>
                          <button
                            type="button"
                            onClick={handleResetFiltersAndSort}
                            className="mt-1 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#107C41] hover:bg-[#0D6535] text-white text-xs font-semibold transition-colors"
                          >
                            <RotateCcw className="w-3.5 h-3.5" />
                            Đặt lại bộ lọc
                          </button>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    visibleRows.map((rowItem, visRowIdx) => {
                      const origRowIdx = rowItem.originalIndex;
                      const isRowSelected =
                        normalizedRange !== null &&
                        visRowIdx >= normalizedRange.minRow &&
                        visRowIdx <= normalizedRange.maxRow;
                      const isPinnedHeaderVisual =
                        hasHeaderRow && origRowIdx === 0;

                      const rawRowH = activeSheet.rowHeights?.[origRowIdx];
                      const rowH =
                        rawRowH && rawRowH > 16
                          ? Math.round(rawRowH)
                          : DEFAULT_ROW_HEIGHT;
                      const isFrozenRow = visRowIdx < frozenRowsCount;
                      const isLastFrozenRow =
                        frozenRowsCount > 0 &&
                        visRowIdx === frozenRowsCount - 1;
                      const frozenTop = frozenRowTopOffsets.get(visRowIdx);

                      return (
                        <tr
                          key={`${activeSheet.id}-orig-row-${origRowIdx}`}
                          style={{ height: `${rowH}px` }}
                          className="group transition-colors"
                        >
                          <td
                            onClick={(event) =>
                              handleSelectRow(visRowIdx, event.shiftKey)
                            }
                            onContextMenu={(e) =>
                              handleOpenContextMenu(
                                e,
                                'row',
                                origRowIdx,
                                visRowIdx,
                                0
                              )
                            }
                            style={{
                              width: `${rowHeaderWidthPx}px`,
                              minWidth: `${rowHeaderWidthPx}px`,
                              ...(isFrozenRow && frozenTop !== undefined
                                ? {
                                    position: 'sticky',
                                    top: `${frozenTop}px`,
                                    zIndex: 24,
                                  }
                                : {}),
                              ...(isLastFrozenRow
                                ? { boxShadow: '0 2px 0 0 rgba(16,124,65,0.35)' }
                                : {}),
                            }}
                            title={`Dòng ${origRowIdx + 1} — bấm chọn cả dòng (chuột phải để chèn/xóa/ẩn dòng)`}
                            className={`sticky left-0 z-10 px-1.5 border-b border-r border-[#E2E8F0] text-center text-xs tabular-nums cursor-pointer transition-colors ${
                              isRowSelected
                                ? 'bg-[#ECFDF5] text-[#107C41] font-bold'
                                : isPinnedHeaderVisual
                                  ? 'bg-[#F8FAFC] text-slate-700 font-semibold'
                                  : 'bg-[#F8FAFC] text-slate-600 font-semibold group-hover:bg-slate-100'
                            }`}
                          >
                            {origRowIdx + 1}
                          </td>

                          {visibleColIndices.map((colIdx, visibleColPos) => {
                            const cellKey = makeCellKey(origRowIdx, colIdx);
                            if (mergeCoveredSet.has(cellKey)) {
                              return null;
                            }
                            const mergeSpan = mergeOriginMap.get(cellKey);

                            const rawValue = rowItem.rawCells[colIdx];
                            const evaluatedValue =
                              rowItem.evaluatedCells[colIdx];
                            const numericVal = parseNumericValue(evaluatedValue);
                            const isNumeric = numericVal !== null;
                            const isFormula = isFormulaText(rawValue);
                            const isError = isFormulaError(evaluatedValue);

                            const style = activeSheet.cellStyles?.[cellKey];
                            const condStyle = conditionalStyleMap.get(cellKey);
                            const validationOpts =
                              activeSheet.dataValidations?.[cellKey];
                            const hasValidation = Boolean(
                              validationOpts && validationOpts.length > 0
                            );

                            const displayValue = formatCellDisplay(
                              evaluatedValue,
                              style?.numFmt
                            );

                            const effectiveAlign =
                              style?.align ?? (isNumeric ? 'right' : 'left');
                            const alignClass =
                              effectiveAlign === 'center'
                                ? 'text-center'
                                : effectiveAlign === 'right'
                                  ? 'text-right'
                                  : effectiveAlign === 'justify'
                                    ? 'text-justify'
                                    : 'text-left';

                            const cellComment =
                              activeSheet.cellComments?.[cellKey];
                            const cellHyperlink = isSafeHttpUrl(style?.hyperlink)
                              ? style?.hyperlink?.trim()
                              : undefined;
                            const isThisCellLocked = isCellInLockedRanges(
                              origRowIdx,
                              colIdx,
                              activeSheet.lockedRanges
                            );

                            const isSelected =
                              normalizedRange !== null &&
                              visRowIdx >= normalizedRange.minRow &&
                              visRowIdx <= normalizedRange.maxRow &&
                              colIdx >= normalizedRange.minCol &&
                              colIdx <= normalizedRange.maxCol;

                            const isSingleSelection =
                              normalizedRange !== null &&
                              normalizedRange.minRow === normalizedRange.maxRow &&
                              normalizedRange.minCol === normalizedRange.maxCol;

                            const isAnchor =
                              normalizedRange !== null &&
                              visRowIdx === normalizedRange.anchorRow &&
                              colIdx === normalizedRange.anchorCol;

                            const isFillHandleCorner =
                              !isEditBlocked &&
                              !hasActiveTransforms &&
                              !editingCell &&
                              normalizedRange !== null &&
                              visRowIdx === normalizedRange.maxRow &&
                              colIdx === normalizedRange.maxCol;

                            const isInFillPreview =
                              fillPreviewBounds !== null &&
                              visRowIdx >= fillPreviewBounds.minRow &&
                              visRowIdx <= fillPreviewBounds.maxRow &&
                              colIdx >= fillPreviewBounds.minCol &&
                              colIdx <= fillPreviewBounds.maxCol;

                            const isEditingThisCell =
                              editingCell !== null &&
                              editingCell.source === 'cell' &&
                              editingCell.originalRow === origRowIdx &&
                              editingCell.col === colIdx;

                            const cellCoord = formatCellRef(origRowIdx, colIdx);
                            const cellPeers =
                              peerPresenceByCellCoord.get(cellCoord);
                            const primaryPeer = cellPeers?.[0];
                            const peerBadgeText = primaryPeer
                              ? cellPeers && cellPeers.length > 1
                                ? `${primaryPeer.displayName} +${cellPeers.length - 1}`
                                : primaryPeer.displayName
                              : null;
                            const peerBadgePositionClass =
                              visRowIdx === 0
                                ? visibleRows.length > 1
                                  ? 'top-full left-0'
                                  : 'top-0 right-0'
                                : 'bottom-full left-0';

                            const shadows: string[] = [];
                            if (
                              isSelected &&
                              !isSingleSelection &&
                              !isEditingThisCell &&
                              normalizedRange
                            ) {
                              if (visRowIdx === normalizedRange.minRow) {
                                shadows.push('inset 0 2px 0 0 #107C41');
                              }
                              if (visRowIdx === normalizedRange.maxRow) {
                                shadows.push('inset 0 -2px 0 0 #107C41');
                              }
                              if (colIdx === normalizedRange.minCol) {
                                shadows.push('inset 2px 0 0 0 #107C41');
                              }
                              if (colIdx === normalizedRange.maxCol) {
                                shadows.push('inset -2px 0 0 0 #107C41');
                              }
                            }

                            const isFrozenCol = visibleColPos < frozenColsCount;
                            const isLastFrozenCol =
                              frozenColsCount > 0 &&
                              visibleColPos === frozenColsCount - 1;
                            const frozenLeft = frozenColLeftOffsets.get(colIdx);

                            if (isLastFrozenRow) {
                              shadows.push('0 2px 0 0 rgba(16,124,65,0.35)');
                            }
                            if (isLastFrozenCol) {
                              shadows.push('2px 0 0 0 rgba(16,124,65,0.35)');
                            }

                            const effectiveBg =
                              condStyle?.bgColor ?? style?.bgColor;
                            const effectiveTextColor =
                              condStyle?.textColor ?? style?.textColor;
                            const hasCustomBg = Boolean(effectiveBg);

                            const cellInlineStyle: CSSProperties = {};
                            if (shadows.length > 0) {
                              cellInlineStyle.boxShadow = shadows.join(', ');
                            }
                            if (!isEditingThisCell && effectiveBg) {
                              cellInlineStyle.backgroundColor = effectiveBg;
                            }
                            if (!isError && effectiveTextColor) {
                              cellInlineStyle.color = effectiveTextColor;
                            }
                            if (
                              typeof style?.fontSize === 'number' &&
                              Number.isFinite(style.fontSize) &&
                              style.fontSize > 0
                            ) {
                              cellInlineStyle.fontSize = `${Math.max(
                                9,
                                Math.min(48, Math.round(style.fontSize))
                              )}px`;
                            }
                            if (style?.fontFamily) {
                              cellInlineStyle.fontFamily = style.fontFamily;
                            }
                            if (style?.italic) {
                              cellInlineStyle.fontStyle = 'italic';
                            }
                            if (style?.underline && style?.strike) {
                              cellInlineStyle.textDecoration =
                                'underline line-through';
                            } else if (style?.underline) {
                              cellInlineStyle.textDecoration = 'underline';
                            } else if (style?.strike) {
                              cellInlineStyle.textDecoration = 'line-through';
                            }
                            cellInlineStyle.verticalAlign =
                              style?.valign ?? 'middle';
                            if (style?.borderTop) {
                              cellInlineStyle.borderTop = style.borderTop;
                            }
                            if (style?.borderRight) {
                              cellInlineStyle.borderRight = style.borderRight;
                            }
                            if (style?.borderBottom) {
                              cellInlineStyle.borderBottom = style.borderBottom;
                            }
                            if (style?.borderLeft) {
                              cellInlineStyle.borderLeft = style.borderLeft;
                            }

                            const colW = Math.round(
                              activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH
                            );
                            if (!mergeSpan) {
                              cellInlineStyle.width = `${colW}px`;
                              cellInlineStyle.maxWidth = `${colW}px`;
                            }

                            if (isFrozenRow || isFrozenCol) {
                              cellInlineStyle.position = 'sticky';
                              if (isFrozenRow && frozenTop !== undefined) {
                                cellInlineStyle.top = `${frozenTop}px`;
                              }
                              if (isFrozenCol && frozenLeft !== undefined) {
                                cellInlineStyle.left = `${frozenLeft}px`;
                              }
                              cellInlineStyle.zIndex =
                                isFrozenRow && isFrozenCol
                                  ? 22
                                  : isFrozenRow
                                    ? 15
                                    : 12;
                            }

                            return (
                              <td
                                key={`${origRowIdx}-${colIdx}`}
                                data-excel-cell={cellKey}
                                rowSpan={mergeSpan?.rowSpan}
                                colSpan={mergeSpan?.colSpan}
                                onMouseDown={(event) =>
                                  handleCellMouseDown(event, visRowIdx, colIdx)
                                }
                                onMouseEnter={() =>
                                  handleCellMouseEnter(visRowIdx, colIdx)
                                }
                                onDoubleClick={() =>
                                  handleCellDoubleClick(visRowIdx, colIdx)
                                }
                                onContextMenu={(e) =>
                                  handleOpenContextMenu(
                                    e,
                                    'cell',
                                    origRowIdx,
                                    visRowIdx,
                                    colIdx
                                  )
                                }
                                title={
                                  [
                                    isFormula
                                      ? `Công thức: ${String(rawValue)}`
                                      : null,
                                    cellComment
                                      ? `Ghi chú: ${cellComment}`
                                      : null,
                                    cellHyperlink
                                      ? `Liên kết: ${cellHyperlink}`
                                      : null,
                                    isThisCellLocked
                                      ? 'Ô nằm trong vùng đã khóa'
                                      : null,
                                  ]
                                    .filter(Boolean)
                                    .join('\n') || undefined
                                }
                                style={cellInlineStyle}
                                className={`relative border-b border-r ${
                                  showGridlines
                                    ? 'border-[#E2E8F0]'
                                    : 'border-transparent'
                                } transition-colors ${
                                  style?.wrapText
                                    ? 'whitespace-pre-wrap break-words py-1 leading-snug'
                                    : 'whitespace-nowrap overflow-visible'
                                } ${
                                  isEditingThisCell
                                    ? 'p-0 bg-white'
                                    : 'px-2.5 cursor-cell'
                                } ${alignClass} ${
                                  isError
                                    ? 'text-rose-600 font-semibold'
                                    : isNumeric
                                      ? 'tabular-nums text-slate-900'
                                      : 'text-slate-700'
                                } ${
                                  style?.bold
                                    ? 'font-bold'
                                    : isNumeric
                                      ? 'font-medium'
                                      : 'font-normal'
                                } ${
                                  isEditingThisCell
                                    ? 'ring-2 ring-inset ring-[#107C41] z-20'
                                    : isSelected
                                      ? isSingleSelection || isAnchor
                                        ? hasCustomBg
                                          ? 'ring-2 ring-inset ring-[#107C41]'
                                          : 'bg-[#ECFDF5] ring-2 ring-inset ring-[#107C41]'
                                        : hasCustomBg
                                          ? ''
                                          : 'bg-[#ECFDF5]/75'
                                      : isInFillPreview
                                        ? 'bg-[#107C41]/10 outline outline-1 outline-dashed outline-[#107C41]'
                                        : hasCustomBg
                                          ? ''
                                          : isPinnedHeaderVisual
                                            ? 'bg-[#F8FAFC] group-hover:bg-slate-100/70'
                                            : 'bg-white group-hover:bg-slate-50/60'
                                }`}
                              >
                                {isSelected &&
                                  hasCustomBg &&
                                  !isEditingThisCell && (
                                    <span
                                      aria-hidden="true"
                                      className="absolute inset-0 bg-[#107C41]/15 pointer-events-none"
                                    />
                                  )}

                                {primaryPeer && (
                                  <>
                                    <span
                                      aria-hidden="true"
                                      style={{
                                        boxShadow: `inset 0 0 0 2px ${primaryPeer.color}`,
                                      }}
                                      className="absolute inset-0 pointer-events-none z-10"
                                    />
                                    {peerBadgeText && (
                                      <span
                                        style={{
                                          backgroundColor: primaryPeer.color,
                                        }}
                                        className={`absolute ${peerBadgePositionClass} z-20 pointer-events-none px-1.5 py-0.5 rounded-sm shadow-sm text-[10px] font-semibold leading-none text-white whitespace-nowrap max-w-[140px] truncate`}
                                      >
                                        {peerBadgeText}
                                        {primaryPeer.isTyping ? '…' : ''}
                                      </span>
                                    )}
                                  </>
                                )}

                                {isEditingThisCell ? (
                                  <input
                                    ref={cellInputRef}
                                    type="text"
                                    value={editDraft}
                                    onChange={(event) =>
                                      setEditDraft(event.target.value)
                                    }
                                    onKeyDown={handleInlineInputKeyDown}
                                    onBlur={handleInlineInputBlur}
                                    aria-label={`Chỉnh sửa ô ${cellCoord}`}
                                    style={{ height: `${rowH}px` }}
                                    className={`w-full px-2.5 bg-white text-slate-900 focus:outline-none select-text ${alignClass} ${
                                      style?.bold ? 'font-bold' : 'font-normal'
                                    }`}
                                  />
                                ) : (
                                  <div className="flex items-center justify-between gap-1 w-full min-w-0">
                                    <span
                                      style={
                                        style?.textRotation === 'vertical'
                                          ? { writingMode: 'vertical-rl' }
                                          : style?.textRotation === 'up45'
                                            ? {
                                                display: 'inline-block',
                                                transform: 'rotate(-35deg)',
                                                transformOrigin: 'center',
                                              }
                                            : style?.textRotation === 'down45'
                                              ? {
                                                  display: 'inline-block',
                                                  transform: 'rotate(35deg)',
                                                  transformOrigin: 'center',
                                                }
                                              : undefined
                                      }
                                      className={`flex-1 min-w-0 ${
                                        style?.wrapText ? '' : 'truncate'
                                      } ${
                                        cellHyperlink
                                          ? 'text-[#107C41] underline decoration-[#107C41]/60'
                                          : ''
                                      }`}
                                    >
                                      {displayValue}
                                    </span>

                                    {cellHyperlink && (
                                      <a
                                        href={cellHyperlink}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        onMouseDown={(e) => e.stopPropagation()}
                                        onClick={(e) => e.stopPropagation()}
                                        title={`Mở liên kết: ${cellHyperlink}`}
                                        className="text-[#107C41] hover:text-[#0B5E31] shrink-0"
                                      >
                                        <ExternalLink className="w-3 h-3" />
                                      </a>
                                    )}

                                    {/* Data Validation In-Cell Dropdown Trigger */}
                                    {hasValidation &&
                                      isAnchor &&
                                      !isEditBlocked && (
                                        <button
                                          type="button"
                                          data-excel-popover="true"
                                          onMouseDown={(e) => e.stopPropagation()}
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setOpenValidationKey((prev) =>
                                              prev === cellKey ? null : cellKey
                                            );
                                          }}
                                          aria-label="Chọn giá trị từ danh sách"
                                          title="Chọn giá trị hợp lệ"
                                          className="w-5 h-5 rounded bg-[#107C41] text-white inline-flex items-center justify-center shrink-0 hover:bg-[#0B5E31] shadow-sm"
                                        >
                                          <ChevronDown className="w-3 h-3" />
                                        </button>
                                      )}
                                  </div>
                                )}

                                {/* Data Validation Popover List */}
                                {hasValidation &&
                                  openValidationKey === cellKey &&
                                  !isEditBlocked && (
                                    <div
                                      data-excel-popover="true"
                                      onMouseDown={(e) => e.stopPropagation()}
                                      className="absolute left-0 top-full mt-0.5 z-50 min-w-full w-48 max-h-44 overflow-y-auto bg-white rounded-lg shadow-xl border border-[#B7DFC7] py-1 text-left"
                                    >
                                      {validationOpts?.map((opt) => (
                                        <button
                                          key={opt}
                                          type="button"
                                          onClick={() => {
                                            commitCellValue(
                                              origRowIdx,
                                              colIdx,
                                              opt
                                            );
                                            setOpenValidationKey(null);
                                          }}
                                          className={`w-full text-left px-3 py-1.5 text-xs hover:bg-[#ECFDF5] hover:text-[#107C41] transition-colors ${
                                            String(evaluatedValue ?? '') === opt
                                              ? 'bg-[#ECFDF5] text-[#107C41] font-bold'
                                              : 'text-slate-700'
                                          }`}
                                        >
                                          {opt}
                                        </button>
                                      ))}
                                    </div>
                                  )}

                                {/* Comment red corner triangle (takes visual priority) or Formula emerald corner triangle */}
                                {cellComment && !isEditingThisCell ? (
                                  <span
                                    aria-hidden="true"
                                    title={`Ghi chú: ${cellComment}`}
                                    className="absolute top-0 right-0 w-0 h-0 border-t-[7px] border-t-rose-500 border-l-[7px] border-l-transparent pointer-events-none"
                                  />
                                ) : (
                                  isFormula &&
                                  !isEditingThisCell && (
                                    <span
                                      aria-hidden="true"
                                      className="absolute top-0 right-0 w-0 h-0 border-t-[6px] border-t-[#107C41] border-l-[6px] border-l-transparent pointer-events-none"
                                    />
                                  )
                                )}

                                {isThisCellLocked && !isEditingThisCell && (
                                  <span
                                    aria-hidden="true"
                                    className="absolute bottom-0.5 left-0.5 w-1.5 h-1.5 rounded-full bg-amber-400/90 pointer-events-none"
                                  />
                                )}

                                {/* Fill Handle (Bottom-Right Square on Selection) */}
                                {isFillHandleCorner && (
                                  <span
                                    role="presentation"
                                    title="Kéo để tự động điền chuỗi số hoặc công thức"
                                    onMouseDown={(e) => {
                                      e.stopPropagation();
                                      e.preventDefault();
                                      if (!normalizedRange) return;
                                      const initial: FillDragState = {
                                        source: normalizedRange,
                                        targetEnd: {
                                          row: normalizedRange.maxRow,
                                          col: normalizedRange.maxCol,
                                        },
                                        axis: null,
                                      };
                                      fillDragRef.current = initial;
                                      setFillDrag(initial);
                                    }}
                                    className="absolute -bottom-[5px] -right-[5px] w-3 h-3 bg-[#107C41] border-2 border-white rounded-[2px] cursor-crosshair z-20 shadow-sm hover:scale-125 transition-transform"
                                  />
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>

              {/* Embedded Interactive SVG Charts */}
              {activeSheet.charts && activeSheet.charts.length > 0 && (
                <ExcelChartsPanel
                  charts={activeSheet.charts}
                  isEditingLocked={isEditBlocked}
                  onChangeChartType={handleChangeChartType}
                  onDeleteChart={handleDeleteChart}
                />
              )}

              {/* Floating Images, Shapes & Text Boxes */}
              <ExcelFloatingObjectsLayer
                objects={activeSheet.floatingObjects ?? []}
                selectedObjectId={selectedFloatingObjId}
                isProtected={isSheetProtected}
                isReadOnly={isEditingLocked}
                zoomScale={zoomScale}
                onSelectObject={setSelectedFloatingObjId}
                onUpdateObject={handleUpdateFloatingObject}
                onDeleteObject={handleDeleteFloatingObject}
              />

              {/* Freehand Drawing Overlay (pointer-events-none when activeTool === null!) */}
              <ExcelDrawingOverlay
                strokes={activeSheet.drawings ?? []}
                activeTool={isEditBlocked ? null : activeDrawTool}
                activeColor={drawColor}
                strokeWidth={drawStrokeWidth}
                zoomScale={zoomScale}
                onAddStroke={handleAddDrawingStroke}
                onEraseStroke={handleEraseDrawingStroke}
              />
            </div>
          </div>
        </div>

        <CustomVerticalScrollbar
          scrollContainerRef={gridContainerRef}
          theme="light"
          getDragLabel={({ percent }) => `${Math.round(percent)}%`}
        />
      </div>

      {/* Sheet Tabs Bar (Fluent Office 365): Add (+), Rename, Duplicate, Color, Drag-Reorder, Delete */}
      <div className="shrink-0 border-t border-slate-200/80 bg-[#F3F4F6] px-2.5 sm:px-3.5 pt-1.5 pb-1 flex items-center justify-between gap-2">
        <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-w-0">
          <div
            role="tablist"
            aria-label="Danh sách trang tính"
            className="flex items-center gap-1 shrink-0"
          >
            {sheets.map((sheet, idx) => {
              const isActive = idx === safeSheetIdx;
              const isRenaming = renamingSheetIdx === idx;

              return (
                <div
                  key={sheet.id}
                  draggable={!isEditingLocked && !isRenaming}
                  onDragStart={() => setDraggingSheetIdx(idx)}
                  onDragOver={(e) => {
                    if (draggingSheetIdx !== null && draggingSheetIdx !== idx) {
                      e.preventDefault();
                    }
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (draggingSheetIdx !== null && draggingSheetIdx !== idx) {
                      handleReorderSheet(draggingSheetIdx, idx);
                    }
                    setDraggingSheetIdx(null);
                  }}
                  onDragEnd={() => setDraggingSheetIdx(null)}
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setSheetTabMenu({
                      sheetIdx: idx,
                      x: Math.max(8, Math.min(e.clientX, window.innerWidth - 220)),
                      y: Math.max(8, e.clientY - 235),
                    });
                  }}
                  style={
                    sheet.tabColor
                      ? {
                          borderBottomColor: sheet.tabColor,
                          borderBottomWidth: '2px',
                        }
                      : undefined
                  }
                  className={`inline-flex items-center gap-1 rounded-t-lg text-xs whitespace-nowrap shrink-0 transition-all ${
                    isActive
                      ? 'bg-white text-[#107C41] font-semibold border-t border-x border-slate-200/80 border-b-2 border-b-[#107C41] shadow-[0_1px_2px_rgba(15,23,42,0.06)]'
                      : 'bg-transparent text-slate-600 font-medium border border-transparent hover:bg-white/75 hover:text-slate-900'
                  }`}
                >
                  {isRenaming ? (
                    <input
                      type="text"
                      autoFocus
                      value={renameSheetDraft}
                      onChange={(e) => setRenameSheetDraft(e.target.value)}
                      onBlur={commitRenameSheet}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') commitRenameSheet();
                        else if (e.key === 'Escape') setRenamingSheetIdx(null);
                      }}
                      className="w-28 px-2 py-1 text-xs font-semibold text-slate-900 bg-white rounded focus:outline-none focus:ring-1 focus:ring-[#107C41]"
                    />
                  ) : (
                    <button
                      type="button"
                      role="tab"
                      aria-selected={isActive}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleSelectSheet(idx)}
                      onDoubleClick={() => startRenameSheet(idx)}
                      title="Bấm đúp để đổi tên, kéo thả để sắp xếp lại hoặc chuột phải để tùy chọn"
                      className="inline-flex items-center gap-1.5 pl-2.5 pr-1.5 py-1 focus:outline-none"
                    >
                      {sheet.tabColor && (
                        <span
                          className="w-2 h-2 rounded-full shrink-0"
                          style={{ backgroundColor: sheet.tabColor }}
                        />
                      )}
                      <span>{sheet.name}</span>
                    </button>
                  )}

                  {!isEditingLocked && !isRenaming && (
                    <button
                      type="button"
                      data-excel-popover="true"
                      onMouseDown={(e) => e.stopPropagation()}
                      onClick={(e) => {
                        e.stopPropagation();
                        const rect = e.currentTarget.getBoundingClientRect();
                        setSheetTabMenu((prev) =>
                          prev?.sheetIdx === idx
                            ? null
                            : {
                                sheetIdx: idx,
                                x: Math.max(
                                  8,
                                  Math.min(rect.left, window.innerWidth - 220)
                                ),
                                y: Math.max(8, rect.top - 235),
                              }
                        );
                      }}
                      aria-label={`Tùy chọn trang tính ${sheet.name}`}
                      className="pr-1.5 pl-0.5 py-1 text-slate-400 hover:text-[#107C41] transition-colors"
                    >
                      <ChevronDown className="w-3 h-3" />
                    </button>
                  )}
                </div>
              );
            })}
          </div>

          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={handleAddSheet}
            disabled={isEditingLocked}
            title="Thêm trang tính mới (+)"
            aria-label="Thêm trang tính mới"
            className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-white hover:bg-[#ECFDF5] border border-slate-200/90 hover:border-[#107C41]/40 text-slate-600 hover:text-[#107C41] shadow-[0_1px_2px_rgba(15,23,42,0.04)] disabled:opacity-40 shrink-0 transition-all"
          >
            <Plus className="w-4 h-4" />
          </button>
        </div>

        <span className="hidden md:inline text-[11px] text-slate-400 whitespace-nowrap shrink-0">
          Chuột phải ô/dòng/cột để thêm menu • Kéo góc xanh lá để AutoFill
        </span>
      </div>

      {/* Bottom Edge: OfficeStatusBar (variant="excel") with Average / Count / Sum & 10%-300% Zoom */}
      <OfficeStatusBar
        variant="excel"
        theme="light"
        zoomPercent={zoomPercent}
        onZoomChange={setZoomPercent}
        onFitToWidth={handleFitToWidth}
        excelStats={{
          average: selectionStats.avg,
          count: selectionStats.nonEmptyCount,
          sum: selectionStats.numericCount > 0 ? selectionStats.sum : null,
        }}
      />

      {/* Column Header Filter & Sort Popover */}
      {filterPopover && (
        <div
          data-excel-popover="true"
          style={{ left: `${filterPopover.x}px`, top: `${filterPopover.y}px` }}
          className="fixed z-50 w-60 bg-white rounded-xl shadow-2xl border border-slate-200 overflow-hidden"
        >
          <div className="p-2 border-b border-slate-100 bg-[#F4F8F5] space-y-1">
            <div className="flex items-center justify-between px-1">
              <span className="text-[11px] font-bold text-[#107C41]">
                Cột {columnIndexToLetter(filterPopover.colIdx)}
              </span>
              <button
                type="button"
                onClick={() => setFilterPopover(null)}
                className="p-0.5 rounded text-slate-400 hover:text-slate-700"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
            <button
              type="button"
              onClick={() => {
                handleSetSort(filterPopover.colIdx, 'asc');
                setFilterPopover(null);
              }}
              className="w-full flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs font-medium text-slate-700 hover:bg-white hover:text-[#107C41]"
            >
              <ArrowUp className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Sắp xếp A → Z (Tăng dần)</span>
            </button>
            <button
              type="button"
              onClick={() => {
                handleSetSort(filterPopover.colIdx, 'desc');
                setFilterPopover(null);
              }}
              className="w-full flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs font-medium text-slate-700 hover:bg-white hover:text-[#107C41]"
            >
              <ArrowDown className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Sắp xếp Z → A (Giảm dần)</span>
            </button>
          </div>

          <div className="p-2 space-y-2">
            <input
              type="text"
              value={filterSearchText}
              onChange={(e) => setFilterSearchText(e.target.value)}
              placeholder="Tìm giá trị lọc..."
              className="w-full px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs focus:outline-none focus:border-[#107C41]"
            />

            <div className="max-h-40 overflow-y-auto space-y-0.5 pr-1">
              {filterPopoverDistinctValues
                .filter((val) =>
                  val.toLowerCase().includes(filterSearchText.trim().toLowerCase())
                )
                .map((val) => {
                  const activeSet = columnFilters[filterPopover.colIdx];
                  const isChecked = !activeSet || activeSet.has(val);
                  return (
                    <label
                      key={val}
                      className="flex items-center gap-2 px-2 py-1 rounded-md hover:bg-slate-50 text-xs text-slate-700 cursor-pointer"
                    >
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() =>
                          handleToggleColumnFilterValue(
                            filterPopover.colIdx,
                            val
                          )
                        }
                        className="rounded border-slate-300 text-[#107C41] focus:ring-[#107C41]"
                      />
                      <span className="truncate">{val}</span>
                    </label>
                  );
                })}
            </div>

            <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-100">
              <button
                type="button"
                onClick={() => {
                  setColumnFilters((prev) => {
                    const next = { ...prev };
                    delete next[filterPopover.colIdx];
                    return next;
                  });
                  if (sortConfig?.colIndex === filterPopover.colIdx) {
                    setSortConfig(null);
                  }
                }}
                className="text-[11px] font-semibold text-rose-600 hover:underline"
              >
                Xóa lọc cột này
              </button>
              <button
                type="button"
                onClick={() => setFilterPopover(null)}
                className="px-3 py-1 rounded-lg bg-[#107C41] text-white text-xs font-semibold"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Right-Click Context Menu (Rows, Columns, Cells) */}
      {contextMenu && (
        <div
          data-excel-popover="true"
          style={{ left: `${contextMenu.x}px`, top: `${contextMenu.y}px` }}
          className="fixed z-50 w-56 bg-white rounded-xl shadow-2xl border border-slate-200 py-1.5 text-xs"
        >
          <div className="px-3 py-1 text-[10.5px] font-bold uppercase tracking-wide text-slate-400 border-b border-slate-100 mb-1">
            {contextMenu.targetType === 'row'
              ? `Dòng ${contextMenu.origRow + 1}`
              : contextMenu.targetType === 'col'
                ? `Cột ${columnIndexToLetter(contextMenu.colIdx)}`
                : `Ô ${formatCellRef(contextMenu.origRow, contextMenu.colIdx)}`}
          </div>

          {contextMenu.targetType !== 'col' && (
            <>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleInsertRowAt(contextMenu.origRow);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Chèn dòng phía trên</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleInsertRowAt(contextMenu.origRow + 1);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Chèn dòng phía dưới</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked || activeSheet.rows.length <= 1}
                onClick={() => {
                  handleDeleteRowIndices(new Set([contextMenu.origRow]));
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 text-rose-600 hover:bg-rose-50 disabled:opacity-40 flex items-center gap-2"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Xóa dòng {contextMenu.origRow + 1}</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleHideRow(contextMenu.origRow);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-slate-100 disabled:opacity-40 flex items-center gap-2"
              >
                <EyeOff className="w-3.5 h-3.5" />
                <span>Ẩn dòng {contextMenu.origRow + 1}</span>
              </button>
            </>
          )}

          {contextMenu.targetType === 'cell' && (
            <div className="my-1 border-t border-slate-100" />
          )}

          {contextMenu.targetType !== 'row' && (
            <>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleInsertColumnAt(contextMenu.colIdx);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Chèn cột bên trái</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleInsertColumnAt(contextMenu.colIdx + 1);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Chèn cột bên phải</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked || columnCount <= 1}
                onClick={() => {
                  handleDeleteColumnIndex(contextMenu.colIdx);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 text-rose-600 hover:bg-rose-50 disabled:opacity-40 flex items-center gap-2"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Xóa cột {columnIndexToLetter(contextMenu.colIdx)}</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleHideColumn(contextMenu.colIdx);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-slate-100 disabled:opacity-40 flex items-center gap-2"
              >
                <EyeOff className="w-3.5 h-3.5" />
                <span>Ẩn cột {columnIndexToLetter(contextMenu.colIdx)}</span>
              </button>
            </>
          )}

          {hiddenCount > 0 && !isEditingLocked && (
            <>
              <div className="my-1 border-t border-slate-100" />
              <button
                type="button"
                onClick={() => {
                  handleUnhideAll();
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 text-[#107C41] font-semibold hover:bg-[#E6F4EC] flex items-center gap-2"
              >
                <Eye className="w-3.5 h-3.5" />
                <span>Hiện tất cả dòng/cột ẩn</span>
              </button>
            </>
          )}

          {contextMenu.targetType === 'cell' && (
            <>
              <div className="my-1 border-t border-slate-100" />
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  if (isSelectionMerged) handleUnmerge();
                  else handleMergeAndCenter();
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                {isSelectionMerged ? (
                  <Split className="w-3.5 h-3.5" />
                ) : (
                  <Merge className="w-3.5 h-3.5" />
                )}
                <span>
                  {isSelectionMerged ? 'Bỏ gộp ô' : 'Gộp & Căn giữa vùng chọn'}
                </span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  setIsValidationModalOpen(true);
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
              >
                <ListChecks className="w-3.5 h-3.5" />
                <span>Xác thực dữ liệu (Dropdown)</span>
              </button>
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => {
                  handleClearSelectedCells();
                  setContextMenu(null);
                }}
                className="w-full text-left px-3 py-1.5 hover:bg-slate-100 disabled:opacity-40 flex items-center gap-2"
              >
                <Eraser className="w-3.5 h-3.5" />
                <span>Xóa nội dung ô chọn</span>
              </button>
            </>
          )}
        </div>
      )}

      {/* Sheet Tab Context Menu */}
      {sheetTabMenu && sheets[sheetTabMenu.sheetIdx] && (
        <div
          data-excel-popover="true"
          style={{ left: `${sheetTabMenu.x}px`, top: `${sheetTabMenu.y}px` }}
          className="fixed z-50 w-52 bg-white rounded-xl shadow-2xl border border-slate-200 py-1.5 text-xs"
        >
          <button
            type="button"
            disabled={isEditingLocked}
            onClick={() => startRenameSheet(sheetTabMenu.sheetIdx)}
            className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
          >
            <Pencil className="w-3.5 h-3.5" />
            <span>Đổi tên trang tính</span>
          </button>
          <button
            type="button"
            disabled={isEditingLocked}
            onClick={() => handleDuplicateSheet(sheetTabMenu.sheetIdx)}
            className="w-full text-left px-3 py-1.5 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40 flex items-center gap-2"
          >
            <Copy className="w-3.5 h-3.5" />
            <span>Nhân bản trang tính</span>
          </button>
          <div className="my-1 border-t border-slate-100" />
          <div className="px-3 py-1.5 space-y-1.5">
            <div className="flex items-center gap-1.5 text-[11px] font-semibold text-slate-500">
              <Palette className="w-3.5 h-3.5" />
              <span>Màu thẻ</span>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {SHEET_TAB_COLORS.map((hex) => (
                <button
                  key={hex}
                  type="button"
                  disabled={isEditingLocked}
                  onClick={() =>
                    handleSetSheetTabColor(sheetTabMenu.sheetIdx, hex)
                  }
                  style={{ backgroundColor: hex }}
                  title={hex}
                  className="w-4 h-4 rounded-full border border-black/15 hover:scale-110 transition-transform"
                />
              ))}
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() =>
                  handleSetSheetTabColor(sheetTabMenu.sheetIdx, undefined)
                }
                title="Bỏ màu thẻ"
                className="px-1.5 h-4 rounded text-[10px] border border-slate-300 text-slate-500 hover:bg-slate-100"
              >
                ×
              </button>
            </div>
          </div>
          <div className="my-1 border-t border-slate-100" />
          <div className="flex items-center gap-1 px-2 py-1">
            <button
              type="button"
              disabled={isEditingLocked || sheetTabMenu.sheetIdx <= 0}
              onClick={() =>
                handleReorderSheet(
                  sheetTabMenu.sheetIdx,
                  sheetTabMenu.sheetIdx - 1
                )
              }
              className="flex-1 inline-flex items-center justify-center gap-1 py-1 rounded-lg bg-slate-100 hover:bg-[#E6F4EC] text-slate-700 hover:text-[#107C41] disabled:opacity-35 text-[11px] font-semibold"
            >
              <ChevronLeft className="w-3 h-3" />
              <span>Trái</span>
            </button>
            <button
              type="button"
              disabled={
                isEditingLocked || sheetTabMenu.sheetIdx >= sheets.length - 1
              }
              onClick={() =>
                handleReorderSheet(
                  sheetTabMenu.sheetIdx,
                  sheetTabMenu.sheetIdx + 1
                )
              }
              className="flex-1 inline-flex items-center justify-center gap-1 py-1 rounded-lg bg-slate-100 hover:bg-[#E6F4EC] text-slate-700 hover:text-[#107C41] disabled:opacity-35 text-[11px] font-semibold"
            >
              <span>Phải</span>
              <ChevronRight className="w-3 h-3" />
            </button>
          </div>
          {sheets.length > 1 && (
            <>
              <div className="my-1 border-t border-slate-100" />
              <button
                type="button"
                disabled={isEditingLocked}
                onClick={() => handleDeleteSheet(sheetTabMenu.sheetIdx)}
                className="w-full text-left px-3 py-1.5 text-rose-600 hover:bg-rose-50 disabled:opacity-40 flex items-center gap-2 font-semibold"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Xóa trang tính</span>
              </button>
            </>
          )}
        </div>
      )}

      {/* Modals: Insert Chart / Conditional Formatting / Data Validation / Text to Columns */}
      {chartModalType !== null && (
        <InsertChartModal
          initialType={chartModalType}
          evaluatedMatrix={evaluatedMatrix}
          selectionRange={selectionOriginalBounds}
          rangeLabel={rangeLabel}
          onClose={() => setChartModalType(null)}
          onInsert={handleInsertChart}
        />
      )}

      {isCondFormatModalOpen && (
        <ConditionalFormatModal
          rules={activeSheet.conditionalFormats ?? []}
          selectionRange={selectionOriginalBounds}
          rangeLabel={rangeLabel}
          onClose={() => setIsCondFormatModalOpen(false)}
          onAddRule={handleAddConditionalFormat}
          onDeleteRule={handleDeleteConditionalFormat}
        />
      )}

      {isValidationModalOpen && (
        <DataValidationModal
          rangeLabel={rangeLabel}
          initialOptions={anchorValidationOptions}
          onClose={() => setIsValidationModalOpen(false)}
          onApply={handleApplyDataValidation}
          onClear={handleClearDataValidation}
        />
      )}

      {isTextToColsModalOpen && anchorCol !== null && (
        <TextToColumnsModal
          columnLetter={columnIndexToLetter(anchorCol)}
          sampleValues={textToColsSamples}
          onClose={() => setIsTextToColsModalOpen(false)}
          onSplit={handleApplyTextToColumns}
        />
      )}

      {/* ONLYOFFICE Formula Library Modal (43 hàm gồm PMT/FV/PV) */}
      <ExcelFormulaLibraryModal
        isOpen={formulaModalCategory !== null}
        initialCategory={formulaModalCategory ?? 'all'}
        activeCellRef={rangeLabel}
        onClose={() => setFormulaModalCategory(null)}
        onInsertFormula={(snippet) => {
          if (snippet.endsWith('()')) {
            handleAiInsertFormula(snippet);
          } else {
            handleQuickInsertFormula(snippet);
          }
        }}
      />

      {/* ONLYOFFICE AI Tools Drawer (Công thức tiếng Việt & Tóm tắt số liệu) */}
      <AIToolsDrawer
        isOpen={aiModalTab === 'summary' || aiModalTab === 'formula'}
        mode="excel"
        initialAction={aiModalTab ?? 'formula'}
        isReadOnly={isSheetProtected}
        documentTitle={document.title}
        excelSheetName={activeSheet.name}
        excelColumns={activeSheet.columns}
        excelRows={activeSheet.rows}
        excelActiveCellRef={rangeLabel}
        onExcelInsertFormula={(formula) => {
          handleAiInsertFormula(formula);
          setAiModalTab(null);
        }}
        onExcelOpenAudit={() => setAiModalTab('audit')}
        onClose={() => setAiModalTab(null)}
      />

      {/* ONLYOFFICE AI Audit Modal (Phát hiện lỗi & trùng lặp) */}
      <ExcelAiToolsModal
        isOpen={aiModalTab === 'audit'}
        initialTab="audit"
        sheetName={activeSheet.name}
        rows={activeSheet.rows}
        evaluatedMatrix={evaluatedMatrix}
        activeCellRef={rangeLabel}
        isProtected={isSheetProtected}
        onClose={() => setAiModalTab(null)}
        onInsertFormula={handleAiInsertFormula}
        onCreateSummarySheet={handleAiCreateSummarySheet}
        onNavigateToCell={handleNavigateToCell}
        onHighlightAuditCells={handleHighlightAuditCells}
      />

      {/* ONLYOFFICE Cell Comment / Hyperlink Modal */}
      <ExcelCellMetaModal
        isOpen={cellMetaModal !== null}
        mode={cellMetaModal?.mode ?? 'comment'}
        cellRef={
          cellMetaModal
            ? formatCellRef(cellMetaModal.origRow, cellMetaModal.colIdx)
            : rangeLabel
        }
        initialValue={
          cellMetaModal
            ? cellMetaModal.mode === 'comment'
              ? (activeSheet.cellComments?.[
                  makeCellKey(cellMetaModal.origRow, cellMetaModal.colIdx)
                ] ?? '')
              : (activeSheet.cellStyles?.[
                  makeCellKey(cellMetaModal.origRow, cellMetaModal.colIdx)
                ]?.hyperlink ?? '')
            : ''
        }
        onClose={() => setCellMetaModal(null)}
        onSave={handleSaveCellMeta}
        onRemove={handleRemoveCellMeta}
      />

      {/* File Info Modal (Tab Tệp -> Thông tin tệp) */}
      {isFileInfoOpen && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => setIsFileInfoOpen(false)}
          className="fixed inset-0 z-50 bg-slate-950/50 backdrop-blur-[1px] flex items-center justify-center p-4"
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="w-full max-w-md rounded-2xl bg-white border border-slate-200 shadow-2xl overflow-hidden"
          >
            <div className="px-4 py-3 bg-[#F8FAFC] border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="w-8 h-8 rounded-lg bg-[#ECFDF5] text-[#107C41] inline-flex items-center justify-center">
                  <FileSpreadsheet className="w-4 h-4" />
                </span>
                <h3 className="text-sm font-bold text-slate-900">
                  Thông tin tệp bảng tính
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setIsFileInfoOpen(false)}
                aria-label="Đóng thông tin tệp"
                className="p-1 rounded-lg text-slate-400 hover:bg-slate-200/70 hover:text-slate-700"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4 space-y-2.5 text-xs">
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Tên tệp</span>
                <span className="font-semibold text-slate-900 truncate">
                  {document.title}
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Định dạng</span>
                <span className="font-semibold uppercase text-[#107C41]">
                  .{document.type}
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Số trang tính</span>
                <span className="font-semibold text-slate-900">
                  {sheets.length} sheet
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Kích thước trang hiện tại</span>
                <span className="font-semibold text-slate-900">
                  {activeSheet.rows.length} dòng × {columnCount} cột
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Dung lượng</span>
                <span className="font-semibold text-slate-900">
                  {document.fileSize || 'Đang tính'}
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1 border-b border-slate-100">
                <span className="text-slate-500">Cập nhật gần nhất</span>
                <span className="font-semibold text-slate-900">
                  {document.updatedAt || 'Vừa mở'}
                </span>
              </div>
              <div className="flex justify-between gap-3 py-1">
                <span className="text-slate-500">Trạng thái bảo vệ</span>
                <span
                  className={`font-semibold ${
                    isSheetProtected ? 'text-amber-700' : 'text-[#107C41]'
                  }`}
                >
                  {isSheetProtected
                    ? 'Đang khóa trang tính'
                    : 'Đang mở chỉnh sửa'}
                </span>
              </div>
            </div>
            <div className="px-4 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setIsFileInfoOpen(false)}
                className="px-4 py-1.5 rounded-lg bg-[#107C41] hover:bg-[#0B5E31] text-white text-xs font-semibold"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

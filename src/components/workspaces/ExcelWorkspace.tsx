import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  ChangeEvent,
  CSSProperties,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
} from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDown,
  ArrowUp,
  BarChart3,
  Bold,
  Calculator,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Columns3,
  Copy,
  Download,
  Eraser,
  Eye,
  EyeOff,
  FileSpreadsheet,
  Filter,
  Grid3x3,
  Italic,
  LineChart,
  ListChecks,
  ListFilter,
  Loader2,
  Merge,
  PaintBucket,
  Palette,
  Pencil,
  PieChart,
  Plus,
  RotateCcw,
  Rows3,
  Save,
  Search,
  Sigma,
  Snowflake,
  Sparkles,
  Split,
  Table2,
  Trash2,
  Type,
  Underline,
  WrapText,
  X,
} from 'lucide-react';
import type {
  CellStyle,
  DocumentItem,
  ExcelChartItem,
  ExcelConditionalFormat,
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
import { saveFileUniversal } from '../../utils/saveFileUniversal';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { useDocumentPinchZoom } from '../../hooks/useDocumentPinchZoom';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { OfficeStatusBar } from './OfficeStatusBar';
import { WorkspaceTopBar } from './WorkspaceTopBar';
import {
  adjustNumFmtDecimals,
  computeAutoFillValue,
  ConditionalFormatModal,
  DataValidationModal,
  evaluateConditionalFormats,
  EXCEL_FORMULA_CATALOG,
  EXCEL_FORMULA_CATEGORY_LABELS,
  ExcelChartsPanel,
  InsertChartModal,
  TextToColumnsModal,
} from './ExcelDataChartModal';
import type { ExcelFormulaCategory } from './ExcelDataChartModal';

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

type RibbonTabId = 'home' | 'insert-formula' | 'data-view';

type OpenPopoverId =
  | 'fillColor'
  | 'textColor'
  | 'borders'
  | 'autoSum'
  | 'formulas'
  | 'freeze'
  | null;

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
    collabPermission === 'view' ||
    (Boolean(document.isUploaded) && !document.hasLocalEdits)
      ? 'original'
      : 'edit'
  );
  const isEditingLocked = isReadOnly || viewMode === 'original';

  const [sheets, setSheets] = useState<SpreadsheetSheet[]>(() =>
    initializeSheets(document)
  );
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [activeRibbonTab, setActiveRibbonTab] = useState<RibbonTabId>('home');
  const [isRibbonCollapsed, setIsRibbonCollapsed] = useState(false);
  const [openPopover, setOpenPopover] = useState<OpenPopoverId>(null);
  const [formulaBrowserCategory, setFormulaBrowserCategory] =
    useState<ExcelFormulaCategory>('math');

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
  const [isDownloadingOriginal, setIsDownloadingOriginal] = useState(false);
  const [feedbackBanner, setFeedbackBanner] = useState<string | null>(null);

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
  const cellInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);

  // Pinch-to-zoom & Ctrl+Wheel on Grid Viewport ONLY (never scaling Ribbon/StatusBar)
  useDocumentPinchZoom(gridContainerRef, zoomPercent, setZoomPercent, {
    min: 25,
    max: 200,
    disabled:
      isModalOpen ||
      chartModalType !== null ||
      isCondFormatModalOpen ||
      isValidationModalOpen ||
      isTextToColsModalOpen,
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
    setViewMode(
      collabPermission === 'view' ||
        (Boolean(document.isUploaded) && !document.hasLocalEdits)
        ? 'original'
        : 'edit'
    );
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

  const evaluatedMatrix = useMemo<(string | number)[][]>(
    () => evaluateSheetMatrix(activeSheet.rows),
    [activeSheet.rows]
  );

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

  const anchorEffectiveAlign: 'left' | 'center' | 'right' =
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

  // Precompute Freeze Panes sticky offsets (px scaled by zoomScale)
  const headerHeightPx = Math.max(22, Math.round(BASE_HEADER_HEIGHT * zoomScale));
  const rowHeaderWidthPx = Math.max(
    32,
    Math.round(BASE_ROW_HEADER_WIDTH * zoomScale)
  );
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
      const h = Math.max(
        20,
        Math.round((rawH && rawH > 16 ? rawH : DEFAULT_ROW_HEIGHT) * zoomScale)
      );
      acc += h;
    }
    return map;
  }, [
    activeSheet.rowHeights,
    frozenRowsCount,
    headerHeightPx,
    visibleRows,
    zoomScale,
  ]);

  const frozenColLeftOffsets = useMemo(() => {
    const map = new Map<number, number>();
    let acc = rowHeaderWidthPx;
    for (let i = 0; i < frozenColsCount; i += 1) {
      const colIdx = visibleColIndices[i];
      if (colIdx === undefined) break;
      map.set(colIdx, acc);
      const rawW = activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH;
      acc += Math.max(48, Math.round(rawW * zoomScale));
    }
    return map;
  }, [
    activeSheet.colWidths,
    frozenColsCount,
    rowHeaderWidthPx,
    visibleColIndices,
    zoomScale,
  ]);

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

  const handleDownloadOriginalBinary = useCallback(async () => {
    if (isDownloadingOriginal) return;
    setIsDownloadingOriginal(true);
    try {
      if (workspaceSync.onDownloadOriginalFile) {
        await workspaceSync.onDownloadOriginalFile();
        showTransientFeedback('Đã tải xuống bản gốc.');
        return;
      }
      const fallbackFileName =
        document.originalFileName ||
        (document.title.toLowerCase().endsWith('.xlsx') ||
        document.title.toLowerCase().endsWith('.xls') ||
        document.title.toLowerCase().endsWith('.csv')
          ? document.title
          : `${document.title || 'bang-tinh'}.xlsx`);
      const fallbackMime =
        document.originalMimeType ||
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

      if (document.rawArrayBuffer && document.rawArrayBuffer.byteLength > 0) {
        const blob = new Blob([document.rawArrayBuffer.slice(0)], {
          type: fallbackMime,
        });
        const res = await saveFileUniversal(blob, fallbackFileName, {
          mimeType: fallbackMime,
          shareTitle: document.title,
        });
        showTransientFeedback(res.message);
        return;
      }
      if (document.sourceFile) {
        const res = await saveFileUniversal(
          document.sourceFile,
          document.sourceFile.name || fallbackFileName,
          { mimeType: document.sourceFile.type || fallbackMime }
        );
        showTransientFeedback(res.message);
        return;
      }
      const url = document.fileUrl ?? document.objectUrl;
      if (url) {
        const resp = await fetch(url);
        if (resp.ok) {
          const blob = await resp.blob();
          const res = await saveFileUniversal(blob, fallbackFileName, {
            mimeType: blob.type || fallbackMime,
          });
          showTransientFeedback(res.message);
          return;
        }
      }
      exportExcelToXlsx(document.title, sheetsRef.current);
      showTransientFeedback(`Đã xuất "${fallbackFileName}".`);
    } catch (err: unknown) {
      showTransientFeedback(
        err instanceof Error
          ? err.message
          : 'Không thể tải xuống bản gốc lúc này.'
      );
    } finally {
      setIsDownloadingOriginal(false);
    }
  }, [
    document.fileUrl,
    document.objectUrl,
    document.originalFileName,
    document.originalMimeType,
    document.rawArrayBuffer,
    document.sourceFile,
    document.title,
    isDownloadingOriginal,
    showTransientFeedback,
    workspaceSync,
  ]);

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
      }
    ) => {
      if (isEditingLocked) return;
      const prevSheets = sheetsRef.current;
      const currentSheet = prevSheets[safeSheetIdx];
      if (!currentSheet) return;
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
    [broadcastExcelPatch, isEditingLocked, safeSheetIdx]
  );

  const commitCellValue = useCallback(
    (originalRow: number, col: number, draft: string) => {
      if (isEditingLocked) return;
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
    [activeSheet.rows, isEditingLocked, updateActiveSheet]
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
      if (isEditingLocked) return;
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
    [activeSheet.rows, isEditingLocked]
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

  useEffect(() => {
    const handleWindowKeyDown = (event: globalThis.KeyboardEvent) => {
      if (isModalOpen || isReadOnly) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        handleSave();
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

  const handleToggleWrapText = () => {
    const target = !anchorStyle.wrapText;
    applyStyleToSelection((prev) => ({ ...prev, wrapText: target }));
  };

  const handleSetAlign = (align: 'left' | 'center' | 'right') => {
    applyStyleToSelection((prev) => ({ ...prev, align }));
  };

  const handleSetValign = (valign: 'top' | 'middle' | 'bottom') => {
    applyStyleToSelection((prev) => ({ ...prev, valign }));
  };

  const handleSetFontSize = (size: number) => {
    applyStyleToSelection((prev) => ({ ...prev, fontSize: size }));
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
    preset: 'all' | 'thick-box' | 'bottom' | 'none'
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
    if (isEditingLocked) return;
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
      isEditingLocked ||
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
    if (isEditingLocked || columnCount <= 1) return;
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
    const availableW = Math.max(240, container.clientWidth - 24);
    const totalUnscaledW =
      BASE_ROW_HEADER_WIDTH +
      visibleColIndices.reduce(
        (sum, c) => sum + (activeSheet.colWidths?.[c] ?? DEFAULT_COL_WIDTH),
        0
      );
    if (totalUnscaledW <= 0) {
      setZoomPercent(100);
      return;
    }
    const idealPct = Math.round((availableW / totalUnscaledW) * 100);
    setZoomPercent(Math.max(25, Math.min(200, Math.min(100, idealPct))));
  }, [activeSheet.colWidths, visibleColIndices]);

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
    if (isEditingLocked) return;
    const origRow = visibleRows[visRowIdx]?.originalIndex;
    if (origRow === undefined) return;
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
    if (isEditingLocked || !normalizedRange) return;
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

    if (event.key === 'Enter' || event.key === 'F2') {
      if (isEditingLocked) return;
      event.preventDefault();
      if (anchorOriginalRow !== null && anchorCol !== null) {
        startCellEditing(anchorOriginalRow, anchorCol, 'cell');
      }
      return;
    }

    if (event.key === 'Backspace' || event.key === 'Delete') {
      if (isEditingLocked) return;
      event.preventDefault();
      handleClearSelectedCells();
      return;
    }

    if (
      !isEditingLocked &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey &&
      event.key.length === 1 &&
      event.key !== ' '
    ) {
      if (anchorOriginalRow !== null && anchorCol !== null) {
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
    return EXCEL_FORMULA_CATALOG.filter((fn) =>
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
    if (isEditingLocked || anchorOriginalRow === null || anchorCol === null) {
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
    if (isEditingLocked) return;
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

  const topBarToolbar = (
    <div className="flex items-center gap-1.5 sm:gap-2">
      {/* Chuyển đổi Xem nguyên bản ↔ Chỉnh sửa nhanh */}
      <div
        role="group"
        aria-label="Chế độ xem bảng tính"
        className="inline-flex items-center p-0.5 rounded-xl bg-[#EDF3EF] border border-[#D4E2D9] shrink-0"
      >
        <button
          type="button"
          onClick={() => {
            if (editingCell) {
              commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
              cancelBlurCommitRef.current = true;
              setEditingCell(null);
              setEditDraft('');
            }
            setViewMode('original');
          }}
          title="Xem nguyên bản (Giống Microsoft Excel)"
          className={`inline-flex items-center gap-1 px-2 sm:px-2.5 py-1 rounded-lg text-[11.5px] font-semibold transition-all whitespace-nowrap ${
            viewMode === 'original'
              ? 'bg-white text-[#107C41] shadow-sm'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Eye className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden lg:inline">Xem nguyên bản</span>
          <span className="lg:hidden">Gốc</span>
        </button>

        {!isReadOnly && (
          <button
            type="button"
            onClick={() => setViewMode('edit')}
            title="Chỉnh sửa nhanh ô, công thức và định dạng"
            className={`inline-flex items-center gap-1 px-2 sm:px-2.5 py-1 rounded-lg text-[11.5px] font-semibold transition-all whitespace-nowrap ${
              viewMode === 'edit'
                ? 'bg-white text-[#107C41] shadow-sm'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Pencil className="w-3.5 h-3.5 shrink-0" />
            <span className="hidden lg:inline">Chỉnh sửa nhanh</span>
            <span className="lg:hidden">Sửa</span>
          </button>
        )}
      </div>

      {hasActiveTransforms && (
        <button
          type="button"
          onClick={handleResetFiltersAndSort}
          className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200/80 text-xs font-medium transition-colors whitespace-nowrap shrink-0"
          title="Đặt lại toàn bộ bộ lọc và sắp xếp"
        >
          <RotateCcw className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden md:inline">Đặt lại lọc</span>
        </button>
      )}

      {!isReadOnly && (
        <button
          type="button"
          onClick={handleSave}
          title="Lưu thay đổi (Ctrl+S)"
          className={`inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap shrink-0 ${
            justSaved
              ? 'bg-emerald-50 border-emerald-300 text-[#107C41]'
              : hasUnsavedChanges
                ? 'bg-[#107C41] hover:bg-[#0D6535] border-[#107C41] text-white shadow-sm'
                : 'bg-[#F5F8F6] hover:bg-[#E7F0EA] border-[#D4E2D9] text-slate-700'
          }`}
        >
          {justSaved ? (
            <>
              <Check className="w-3.5 h-3.5 shrink-0" />
              <span>Đã lưu</span>
            </>
          ) : (
            <>
              <Save className="w-3.5 h-3.5 shrink-0" />
              <span className="hidden sm:inline">Lưu thay đổi</span>
              <span className="sm:hidden">Lưu</span>
              {hasUnsavedChanges && (
                <span
                  className="w-1.5 h-1.5 rounded-full bg-amber-300 shrink-0"
                  aria-hidden="true"
                />
              )}
            </>
          )}
        </button>
      )}

      <button
        type="button"
        onClick={() => void handleDownloadOriginalBinary()}
        disabled={isDownloadingOriginal}
        title="Tải xuống bản gốc ban đầu 100%"
        className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl bg-white hover:bg-[#E6F4EC] border border-[#D4E2D9] text-slate-700 hover:text-[#107C41] text-xs font-semibold transition-colors whitespace-nowrap shrink-0 disabled:opacity-50"
      >
        {isDownloadingOriginal ? (
          <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
        ) : (
          <Download className="w-3.5 h-3.5 shrink-0" />
        )}
        <span className="hidden xl:inline">Tải xuống bản gốc</span>
        <span className="xl:hidden">Bản gốc</span>
      </button>

      <button
        type="button"
        onClick={handleExportXlsx}
        title="Xuất bảng tính hiện tại ra tệp .xlsx"
        className="inline-flex items-center gap-1.5 px-2 sm:px-2.5 py-1.5 rounded-xl bg-[#F6FAF7] hover:bg-[#E6F4EC] border border-[#D4E2D9] text-slate-700 hover:text-[#107C41] text-xs font-semibold transition-colors whitespace-nowrap shrink-0"
      >
        <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
        <span className="hidden md:inline">.xlsx</span>
      </button>
    </div>
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
    <div className="h-[100dvh] w-full flex flex-col bg-[#F3F6F4] text-slate-900 overflow-hidden">
      <WorkspaceTopBar
        document={document}
        theme="light"
        onBackHome={onBackHome}
        onShare={() => onOpenModal('share', buildLiveDocument(sheets))}
        onRenameTitle={isReadOnly ? undefined : handleRenameTitle}
        collaborators={collaborators}
        selfSessionId={selfSessionId}
        collabPermission={collabPermission}
        toolbar={topBarToolbar}
      />

      {feedbackBanner && (
        <div className="shrink-0 px-3 sm:px-4 py-1.5 bg-[#E6F4EC] border-b border-[#B7DFC7] text-[#107C41] text-xs font-medium flex items-center justify-between gap-2">
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

      {/* Microsoft Excel #107C41 Ribbon (Tabs + Scrollable Command Strip — Responsive Desktop & Mobile Android) */}
      <div
        data-excel-popover="true"
        className="shrink-0 bg-[#F3F6F4] border-b border-[#D8E2DC] relative z-30"
      >
        {/* Tab Strip */}
        <div className="flex items-center justify-between gap-2 px-2.5 sm:px-4 border-b border-[#E1EAE4]">
          <div
            role="tablist"
            aria-label="Ribbon Microsoft Excel"
            className="flex items-center gap-0.5 overflow-x-auto no-scrollbar"
          >
            {(
              [
                { id: 'home', label: 'Trang đầu' },
                { id: 'insert-formula', label: 'Chèn & Công thức' },
                { id: 'data-view', label: 'Dữ liệu & Xem' },
              ] as const
            ).map((tab) => {
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
                  }}
                  className={`px-3 py-1.5 text-xs whitespace-nowrap border-b-2 transition-colors ${
                    isActive
                      ? 'border-[#107C41] text-[#107C41] font-bold bg-white/80 rounded-t-md'
                      : 'border-transparent text-slate-600 font-medium hover:text-slate-900 hover:bg-white/40'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            {isFiltered && (
              <span className="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded bg-[#E6F4EC] border border-[#B7DFC7] text-[#107C41] text-[11px] font-semibold">
                <Filter className="w-3 h-3" />
                {visibleRows.length}/{activeSheet.rows.length} dòng
              </span>
            )}
            {activeFreezeLabel && (
              <span className="hidden lg:inline-flex items-center gap-1 px-2 py-0.5 rounded bg-sky-50 border border-sky-200 text-sky-700 text-[11px] font-semibold">
                <Snowflake className="w-3 h-3" />
                {activeFreezeLabel}
              </span>
            )}
            {hiddenCount > 0 && !isEditingLocked && (
              <button
                type="button"
                onClick={handleUnhideAll}
                className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-800 text-[11px] font-semibold"
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
              title={isRibbonCollapsed ? 'Mở thanh công cụ Ribbon' : 'Thu gọn thanh công cụ Ribbon'}
              aria-label={isRibbonCollapsed ? 'Mở Ribbon' : 'Thu gọn Ribbon'}
              className="p-1 rounded-md text-slate-500 hover:text-slate-800 hover:bg-white/80 transition-colors"
            >
              {isRibbonCollapsed ? (
                <ChevronDown className="w-3.5 h-3.5" />
              ) : (
                <ChevronUp className="w-3.5 h-3.5" />
              )}
            </button>
          </div>
        </div>

        {/* Scrollable Ribbon Command Row */}
        {!isRibbonCollapsed && (
          <div className="bg-white px-2.5 sm:px-3.5 py-1.5 overflow-x-auto no-scrollbar">
            {activeRibbonTab === 'home' && (
              <div className="flex items-center gap-1.5 w-max">
                {/* Font Style: B / I / U / Size / Text Color / Fill Color */}
                <div className="inline-flex items-center gap-0.5 rounded-lg bg-[#F5F8F6] border border-[#DCE5DF] p-0.5 shrink-0">
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={handleToggleBold}
                    disabled={isEditingLocked || !normalizedRange}
                    aria-pressed={anchorStyle.bold === true}
                    title="In đậm (Bold)"
                    className={`w-7 h-7 rounded-md inline-flex items-center justify-center transition-colors disabled:opacity-40 ${
                      anchorStyle.bold
                        ? 'bg-[#107C41] text-white'
                        : 'text-slate-700 hover:bg-white'
                    }`}
                  >
                    <Bold className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={handleToggleItalic}
                    disabled={isEditingLocked || !normalizedRange}
                    aria-pressed={anchorStyle.italic === true}
                    title="In nghiêng (Italic)"
                    className={`w-7 h-7 rounded-md inline-flex items-center justify-center transition-colors disabled:opacity-40 ${
                      anchorStyle.italic
                        ? 'bg-[#107C41] text-white'
                        : 'text-slate-700 hover:bg-white'
                    }`}
                  >
                    <Italic className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={handleToggleUnderline}
                    disabled={isEditingLocked || !normalizedRange}
                    aria-pressed={anchorStyle.underline === true}
                    title="Gạch chân (Underline)"
                    className={`w-7 h-7 rounded-md inline-flex items-center justify-center transition-colors disabled:opacity-40 ${
                      anchorStyle.underline
                        ? 'bg-[#107C41] text-white'
                        : 'text-slate-700 hover:bg-white'
                    }`}
                  >
                    <Underline className="w-3.5 h-3.5" />
                  </button>
                </div>

                <select
                  aria-label="Cỡ chữ ô"
                  title="Cỡ chữ"
                  disabled={isEditingLocked || !normalizedRange}
                  value={Math.round(anchorStyle.fontSize ?? 13)}
                  onChange={(e) => handleSetFontSize(Number(e.target.value))}
                  className="h-7 px-1.5 rounded-lg border border-[#DCE5DF] bg-white text-xs font-semibold text-slate-700 focus:outline-none focus:border-[#107C41] disabled:opacity-40 shrink-0"
                >
                  {FONT_SIZE_OPTIONS.map((sz) => (
                    <option key={sz} value={sz}>
                      {sz}
                    </option>
                  ))}
                </select>

                {/* Fill Color Button */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'fillColor' ? null : 'fillColor'))
                  }
                  title="Màu nền ô (Fill Color)"
                  className={`h-7 px-2 rounded-lg border inline-flex items-center gap-1 text-xs font-medium transition-colors disabled:opacity-40 shrink-0 ${
                    openPopover === 'fillColor'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
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

                {/* Text Color Button */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'textColor' ? null : 'textColor'))
                  }
                  title="Màu chữ (Font Color)"
                  className={`h-7 px-2 rounded-lg border inline-flex items-center gap-1 text-xs font-medium transition-colors disabled:opacity-40 shrink-0 ${
                    openPopover === 'textColor'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
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

                {/* Borders Dropdown */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'borders' ? null : 'borders'))
                  }
                  title="Kẻ khung viền ô (Borders)"
                  className={`h-7 px-2 rounded-lg border inline-flex items-center gap-1 text-xs font-medium transition-colors disabled:opacity-40 shrink-0 ${
                    openPopover === 'borders'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <Grid3x3 className="w-3.5 h-3.5" />
                  <span className="hidden xl:inline">Viền</span>
                  <ChevronDown className="w-3 h-3 text-slate-400" />
                </button>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* 9-Position Alignment (Horizontal + Vertical) */}
                <div
                  role="group"
                  aria-label="Căn lề ngang"
                  className="inline-flex items-center gap-0.5 rounded-lg bg-[#F5F8F6] border border-[#DCE5DF] p-0.5 shrink-0"
                >
                  {(
                    [
                      { id: 'left', icon: AlignLeft, label: 'Căn trái' },
                      { id: 'center', icon: AlignCenter, label: 'Căn giữa ngang' },
                      { id: 'right', icon: AlignRight, label: 'Căn phải' },
                    ] as const
                  ).map(({ id, icon: Icon, label }) => (
                    <button
                      key={id}
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => handleSetAlign(id)}
                      disabled={isEditingLocked || !normalizedRange}
                      title={label}
                      aria-label={label}
                      className={`w-6 h-6 rounded inline-flex items-center justify-center transition-colors disabled:opacity-40 ${
                        anchorEffectiveAlign === id
                          ? 'bg-[#107C41] text-white'
                          : 'text-slate-600 hover:bg-white'
                      }`}
                    >
                      <Icon className="w-3.5 h-3.5" />
                    </button>
                  ))}
                </div>

                <div
                  role="group"
                  aria-label="Căn lề dọc"
                  className="inline-flex items-center gap-0.5 rounded-lg bg-[#F5F8F6] border border-[#DCE5DF] p-0.5 shrink-0"
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
                      disabled={isEditingLocked || !normalizedRange}
                      title={label}
                      aria-label={label}
                      className={`w-6 h-6 rounded inline-flex items-center justify-center transition-colors disabled:opacity-40 ${
                        anchorEffectiveValign === id
                          ? 'bg-[#107C41] text-white'
                          : 'text-slate-600 hover:bg-white'
                      }`}
                    >
                      <Icon className="w-3.5 h-3.5" />
                    </button>
                  ))}
                </div>

                {/* Wrap Text */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleToggleWrapText}
                  disabled={isEditingLocked || !normalizedRange}
                  aria-pressed={anchorStyle.wrapText === true}
                  title="Tự động xuống dòng trong ô (Wrap Text)"
                  className={`h-7 px-2 rounded-lg border inline-flex items-center gap-1 text-xs font-medium transition-colors disabled:opacity-40 shrink-0 ${
                    anchorStyle.wrapText
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41] font-semibold'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <WrapText className="w-3.5 h-3.5" />
                  <span className="hidden lg:inline">Xuống dòng</span>
                </button>

                {/* Merge & Center / Unmerge */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={
                    isSelectionMerged ? handleUnmerge : handleMergeAndCenter
                  }
                  disabled={isEditingLocked || !normalizedRange}
                  title={
                    isSelectionMerged
                      ? 'Bỏ gộp các ô đang chọn (Unmerge)'
                      : 'Gộp và căn giữa vùng chọn (Merge & Center)'
                  }
                  className={`h-7 px-2 rounded-lg border inline-flex items-center gap-1 text-xs font-medium transition-colors disabled:opacity-40 shrink-0 ${
                    isSelectionMerged
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41] font-semibold'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
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
                      <span>Gộp ô</span>
                    </>
                  )}
                </button>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* Number Format & Decimals */}
                <select
                  aria-label="Định dạng số"
                  title="Định dạng số (Chung, Số, Tiền tệ VNĐ, Phần trăm, Ngày tháng)"
                  disabled={isEditingLocked || !normalizedRange}
                  value={detectFormatPresetValue(anchorStyle.numFmt)}
                  onChange={(e) => handleSetNumberFormat(e.target.value)}
                  className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-white text-xs font-medium text-slate-700 focus:outline-none focus:border-[#107C41] disabled:opacity-40 shrink-0"
                >
                  {NUMBER_FORMAT_PRESETS.map((fmt) => (
                    <option key={fmt.id} value={fmt.numFmt}>
                      {fmt.label}
                    </option>
                  ))}
                </select>

                <div className="inline-flex items-center rounded-lg border border-[#DCE5DF] bg-white overflow-hidden shrink-0">
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => handleAdjustDecimals(1)}
                    disabled={isEditingLocked || !normalizedRange}
                    title="Tăng chữ số thập phân (.00)"
                    className="h-7 px-2 text-[11px] font-bold text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41] border-r border-[#DCE5DF] disabled:opacity-40"
                  >
                    .00
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => handleAdjustDecimals(-1)}
                    disabled={isEditingLocked || !normalizedRange}
                    title="Giảm chữ số thập phân (.0)"
                    className="h-7 px-2 text-[11px] font-bold text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41] disabled:opacity-40"
                  >
                    .0
                  </button>
                </div>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* Conditional Formatting */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => setIsCondFormatModalOpen(true)}
                  disabled={isEditingLocked || !normalizedRange}
                  title="Định dạng có điều kiện (Lớn hơn, Nhỏ hơn, Bằng, Trùng lặp)"
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-medium text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0 transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Định dạng ĐK</span>
                  {(activeSheet.conditionalFormats?.length ?? 0) > 0 && (
                    <span className="px-1.5 py-0.2 rounded-full bg-[#107C41] text-white text-[10px] font-bold">
                      {activeSheet.conditionalFormats?.length}
                    </span>
                  )}
                </button>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* Cells Structure: Add Row / Add Col / Delete Row / Clear */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleAddRow}
                  disabled={isEditingLocked}
                  title="Thêm dòng mới vào cuối bảng"
                  className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] text-xs font-medium text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1 disabled:opacity-40 shrink-0"
                >
                  <Plus className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Dòng</span>
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleAddColumn}
                  disabled={isEditingLocked}
                  title="Thêm cột mới vào cuối bảng"
                  className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] text-xs font-medium text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1 disabled:opacity-40 shrink-0"
                >
                  <Plus className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Cột</span>
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleDeleteSelectedRows}
                  disabled={isEditingLocked || !canDeleteSelectedRows}
                  title="Xóa các dòng đang chọn"
                  className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-white hover:bg-rose-50 hover:border-rose-200 text-xs font-medium text-slate-700 hover:text-rose-600 inline-flex items-center gap-1 disabled:opacity-40 shrink-0"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span className="hidden lg:inline">Xóa dòng</span>
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleClearSelectedCells}
                  disabled={isEditingLocked || !normalizedRange}
                  title="Xóa nội dung các ô đang chọn"
                  className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-white hover:bg-slate-100 text-xs font-medium text-slate-600 inline-flex items-center gap-1 disabled:opacity-40 shrink-0"
                >
                  <Eraser className="w-3.5 h-3.5" />
                </button>
              </div>
            )}

            {activeRibbonTab === 'insert-formula' && (
              <div className="flex items-center gap-1.5 w-max">
                {/* AutoSum (Σ) */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'autoSum' ? null : 'autoSum'))
                  }
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors disabled:opacity-40 shrink-0 ${
                    openPopover === 'autoSum'
                      ? 'bg-[#107C41] border-[#107C41] text-white'
                      : 'bg-[#E6F4EC] border-[#B7DFC7] text-[#107C41] hover:bg-[#D5ECDF]'
                  }`}
                >
                  <Sigma className="w-3.5 h-3.5" />
                  <span>AutoSum (Σ)</span>
                  <ChevronDown className="w-3 h-3" />
                </button>

                {/* Formula Library (38+ functions) */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'formulas' ? null : 'formulas'))
                  }
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors disabled:opacity-40 shrink-0 ${
                    openPopover === 'formulas'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <Calculator className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Thư viện hàm ({EXCEL_FORMULA_CATALOG.length})</span>
                  <ChevronDown className="w-3 h-3 text-slate-400" />
                </button>

                {/* Quick popular formula shortcuts */}
                {(['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNTIF'] as const).map(
                  (fn) => (
                    <button
                      key={fn}
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      disabled={isEditingLocked || !normalizedRange}
                      onClick={() => handleQuickInsertFormula(`=${fn}(`)}
                      className="h-7 px-2 rounded-lg border border-[#DCE5DF] bg-[#F8FAF8] hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-[11.5px] font-bold text-[#107C41] disabled:opacity-40 shrink-0"
                    >
                      ={fn}()
                    </button>
                  )
                )}

                <span className="h-5 w-px bg-[#DCE5DF] mx-1 shrink-0" />

                {/* Insert Charts: Column / Line / Pie */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked}
                  onClick={() => setChartModalType('column')}
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <BarChart3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Biểu đồ Cột</span>
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked}
                  onClick={() => setChartModalType('line')}
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <LineChart className="w-3.5 h-3.5 text-[#2563EB]" />
                  <span>Biểu đồ Đường</span>
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked}
                  onClick={() => setChartModalType('pie')}
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <PieChart className="w-3.5 h-3.5 text-[#D97706]" />
                  <span>Biểu đồ Tròn</span>
                </button>

                {numericColumnIndices.length > 0 && visibleRows.length > 0 && (
                  <>
                    <span className="h-5 w-px bg-[#DCE5DF] mx-1 shrink-0" />
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={handleQuickSelectNumericColumn}
                      className="h-7 px-2.5 rounded-lg bg-[#107C41]/10 hover:bg-[#107C41]/15 border border-[#107C41]/25 text-[#107C41] text-xs font-semibold inline-flex items-center gap-1 shrink-0"
                    >
                      <Sigma className="w-3.5 h-3.5" />
                      <span>Chọn cột số</span>
                    </button>
                  </>
                )}
              </div>
            )}

            {activeRibbonTab === 'data-view' && (
              <div className="flex items-center gap-1.5 w-max">
                {/* Freeze Panes */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked}
                  onClick={() =>
                    setOpenPopover((p) => (p === 'freeze' ? null : 'freeze'))
                  }
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors disabled:opacity-40 shrink-0 ${
                    activeFreezeLabel || openPopover === 'freeze'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <Snowflake className="w-3.5 h-3.5" />
                  <span>Cố định khung</span>
                  <ChevronDown className="w-3 h-3" />
                </button>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* Sort & AutoFilter */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={anchorCol === null}
                  onClick={() =>
                    anchorCol !== null && handleSetSort(anchorCol, 'asc')
                  }
                  title="Sắp xếp cột hiện tại tăng dần (A → Z)"
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-40 shrink-0 ${
                    sortConfig?.colIndex === anchorCol &&
                    sortConfig?.direction === 'asc'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <ArrowUp className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>A → Z</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={anchorCol === null}
                  onClick={() =>
                    anchorCol !== null && handleSetSort(anchorCol, 'desc')
                  }
                  title="Sắp xếp cột hiện tại giảm dần (Z → A)"
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1 text-xs font-semibold disabled:opacity-40 shrink-0 ${
                    sortConfig?.colIndex === anchorCol &&
                    sortConfig?.direction === 'desc'
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <ArrowDown className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Z → A</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={handleToggleAutoFilter}
                  title="Bật/tắt phễu lọc trực tiếp trên tiêu đề cột"
                  className={`h-7 px-2.5 rounded-lg border inline-flex items-center gap-1.5 text-xs font-semibold transition-colors shrink-0 ${
                    activeSheet.autoFilterEnabled || hasColumnFilters
                      ? 'bg-[#107C41] border-[#107C41] text-white'
                      : 'bg-white border-[#DCE5DF] text-slate-700 hover:bg-[#F5F8F6]'
                  }`}
                >
                  <ListFilter className="w-3.5 h-3.5" />
                  <span>Bộ lọc cột</span>
                </button>

                <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />

                {/* Data Tools */}
                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked}
                  onClick={handleRemoveDuplicates}
                  title="Xóa các dòng có giá trị trùng lặp"
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <Rows3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Xóa trùng lặp</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || anchorCol === null}
                  onClick={() => setIsTextToColsModalOpen(true)}
                  title="Tách dữ liệu văn bản trong cột hiện tại sang nhiều cột"
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <Columns3 className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Tách cột (Text to Columns)</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  disabled={isEditingLocked || !normalizedRange}
                  onClick={() => setIsValidationModalOpen(true)}
                  title="Tạo danh sách thả xuống chọn nhanh trong ô (Data Validation)"
                  className="h-7 px-2.5 rounded-lg border border-[#DCE5DF] bg-white hover:bg-[#E6F4EC] hover:border-[#107C41]/40 text-xs font-semibold text-slate-700 hover:text-[#107C41] inline-flex items-center gap-1.5 disabled:opacity-40 shrink-0"
                >
                  <ListChecks className="w-3.5 h-3.5 text-[#107C41]" />
                  <span>Xác thực dữ liệu</span>
                </button>

                {hiddenCount > 0 && !isEditingLocked && (
                  <>
                    <span className="h-5 w-px bg-[#DCE5DF] mx-0.5 shrink-0" />
                    <button
                      type="button"
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={handleUnhideAll}
                      className="h-7 px-2.5 rounded-lg border border-amber-300 bg-amber-50 hover:bg-amber-100 text-xs font-semibold text-amber-800 inline-flex items-center gap-1.5 shrink-0"
                    >
                      <Eye className="w-3.5 h-3.5" />
                      <span>Hiện tất cả dòng/cột ẩn ({hiddenCount})</span>
                    </button>
                  </>
                )}
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
                className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41] transition-colors"
              >
                {item.label}
              </button>
            ))}
          </div>
        )}

        {openPopover === 'autoSum' && (
          <div className="absolute left-3 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-52">
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
                className="w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41] transition-colors"
              >
                <span>{label}</span>
                <span className="text-[10px] font-mono text-slate-400">
                  ={fn}
                </span>
              </button>
            ))}
          </div>
        )}

        {openPopover === 'formulas' && (
          <div className="absolute left-3 sm:left-24 top-full mt-1 z-50 bg-white rounded-xl shadow-2xl border border-slate-200 w-[min(92vw,440px)] overflow-hidden">
            <div className="flex items-center gap-1 p-2 border-b border-slate-100 bg-slate-50 overflow-x-auto no-scrollbar">
              {(
                Object.keys(
                  EXCEL_FORMULA_CATEGORY_LABELS
                ) as ExcelFormulaCategory[]
              ).map((cat) => (
                <button
                  key={cat}
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => setFormulaBrowserCategory(cat)}
                  className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold whitespace-nowrap transition-colors ${
                    formulaBrowserCategory === cat
                      ? 'bg-[#107C41] text-white'
                      : 'text-slate-600 hover:bg-white'
                  }`}
                >
                  {EXCEL_FORMULA_CATEGORY_LABELS[cat]}
                </button>
              ))}
            </div>
            <div className="max-h-64 overflow-y-auto p-1.5 divide-y divide-slate-100">
              {EXCEL_FORMULA_CATALOG.filter(
                (f) => f.category === formulaBrowserCategory
              ).map((item) => (
                <button
                  key={item.name}
                  type="button"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => handleQuickInsertFormula(`=${item.name}(`)}
                  className="w-full text-left px-2.5 py-2 hover:bg-[#E6F4EC] rounded-lg transition-colors"
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold text-[#107C41] font-mono">
                      {item.name}
                    </span>
                    <span className="text-[10.5px] font-mono text-slate-400 truncate">
                      {item.syntax}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-600 mt-0.5">
                    {item.description}
                  </p>
                </button>
              ))}
            </div>
          </div>
        )}

        {openPopover === 'freeze' && (
          <div className="absolute left-3 top-full mt-1 z-50 bg-white rounded-xl shadow-xl border border-slate-200 p-1.5 w-60">
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(1, activeSheet.frozenCols ?? 0)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41]"
            >
              Cố định dòng đầu tiên (Freeze Top Row)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(activeSheet.frozenRows ?? 0, 1)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41]"
            >
              Cố định cột đầu tiên (Freeze First Column)
            </button>
            <button
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => handleSetFreezePanes(1, 1)}
              className="w-full text-left px-3 py-2 rounded-lg text-xs font-medium text-slate-700 hover:bg-[#E6F4EC] hover:text-[#107C41]"
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

      {/* Name Box (A1) + Formula Bar (fx) + Search Bar */}
      <div className="shrink-0 bg-white border-b border-[#DCE5DF] px-2.5 sm:px-4 py-1.5 relative z-20">
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
            className="w-[78px] sm:w-[92px] px-2 py-1 rounded-lg bg-[#F4F8F5] border border-[#D4E2D9] focus:bg-white focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15 focus:outline-none text-xs font-bold text-[#107C41] text-center tabular-nums shrink-0 transition-all"
          />

          {/* Formula Bar (fx) + Autocomplete Dropdown */}
          <div className="relative flex-1 min-w-[180px]">
            <div className="flex items-center gap-2 px-2.5 py-1 rounded-lg bg-[#FAFBF9] border border-[#DCE5DF] focus-within:bg-white focus-within:border-[#107C41] focus-within:ring-2 focus-within:ring-[#107C41]/15 transition-all">
              <span
                className="font-serif italic font-bold text-[#107C41]/70 select-none shrink-0 pr-2 border-r border-[#DCE5DF] text-sm"
                aria-hidden="true"
              >
                fx
              </span>
              <input
                ref={formulaInputRef}
                type="text"
                value={formulaInputValue}
                readOnly={isEditingLocked}
                disabled={anchorOriginalRow === null || anchorCol === null}
                onFocus={handleFormulaFocus}
                onChange={handleFormulaChange}
                onKeyDown={handleFormulaKeyDown}
                onBlur={handleFormulaBlur}
                placeholder={
                  isReadOnly
                    ? 'Chế độ chỉ xem — chọn ô để xem giá trị hoặc công thức'
                    : viewMode === 'original'
                      ? 'Chế độ Xem nguyên bản — bấm "Chỉnh sửa nhanh" để sửa ô hoặc công thức'
                      : 'Nhập giá trị hoặc gõ = để gợi ý 40 hàm Excel (SUM, IF, VLOOKUP, INDEX...)...'
                }
                aria-label="Thanh công thức fx"
                className="w-full py-0.5 bg-transparent text-xs sm:text-[13px] font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none disabled:opacity-50"
              />
              {isFormulaText(anchorRawValue) && (
                <span className="hidden md:inline-flex items-center px-2 py-0.5 rounded bg-[#E6F4EC] border border-[#B7DFC7] text-[11px] font-semibold text-[#107C41] tabular-nums shrink-0 whitespace-nowrap">
                  = {formatCellDisplay(anchorEvaluatedValue, anchorStyle.numFmt)}
                </span>
              )}
            </div>

            {/* Live Formula Autocomplete Popover */}
            {formulaSuggestions.length > 0 && (
              <div className="absolute left-0 right-0 sm:right-auto sm:min-w-[360px] top-full mt-1 z-50 bg-white rounded-xl shadow-2xl border border-[#C6DFCF] py-1 overflow-hidden">
                <div className="px-2.5 py-1 bg-[#F3F8F5] border-b border-slate-100 flex items-center justify-between text-[10.5px] font-semibold text-slate-500">
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
                    className={`w-full text-left px-3 py-1.5 flex flex-col hover:bg-[#E6F4EC] transition-colors ${
                      idx === 0 ? 'bg-[#F4FAF6]' : ''
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
              className="w-full pl-8 pr-7 py-1 rounded-lg bg-[#F5F8F6] border border-[#DCE5DF] text-xs sm:text-[13px] text-slate-800 placeholder:text-slate-400 focus:outline-none focus:bg-white focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15 transition-all"
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

      {/* Main Grid Viewport (Zoomed 25%-200% via pinch / StatusBar, with CustomVerticalScrollbar) */}
      <div className="relative flex-1 min-h-0 flex overflow-hidden">
        <div
          ref={gridContainerRef}
          tabIndex={0}
          onKeyDown={handleGridKeyDown}
          className="w-full h-full overflow-auto scrollbar-none pr-4 bg-white select-none focus:outline-none"
        >
          <table
            style={{
              fontSize: `${Math.max(8, Math.round(13 * zoomScale * 10) / 10)}px`,
            }}
            className="w-max min-w-full border-separate border-spacing-0"
          >
            <colgroup>
              <col
                style={{
                  width: `${rowHeaderWidthPx}px`,
                  minWidth: `${rowHeaderWidthPx}px`,
                }}
              />
              {visibleColIndices.map((colIdx) => {
                const rawW = activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH;
                const scaledW = Math.max(48, Math.round(rawW * zoomScale));
                return (
                  <col
                    key={`colgroup-${colIdx}`}
                    style={{
                      width: `${scaledW}px`,
                      minWidth: `${scaledW}px`,
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
                  className="sticky left-0 top-0 z-30 p-0 bg-[#EBF1ED] border-b border-r border-[#D0DCD4]"
                >
                  <button
                    type="button"
                    onClick={handleSelectAllVisible}
                    title="Chọn toàn bộ bảng tính đang hiển thị"
                    aria-label="Chọn toàn bộ bảng tính"
                    className="w-full h-full flex items-center justify-center text-slate-400 hover:text-[#107C41] hover:bg-[#DFECE4] transition-colors"
                  >
                    <Table2 className="w-3.5 h-3.5" />
                  </button>
                </th>

                {visibleColIndices.map((colIdx, visibleColPos) => {
                  const letter = columnIndexToLetter(colIdx);
                  const rawW =
                    activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH;
                  const scaledW = Math.max(48, Math.round(rawW * zoomScale));
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
                    width: `${scaledW}px`,
                    minWidth: `${scaledW}px`,
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
                      className={`sticky top-0 z-20 p-0 border-b border-r border-[#D0DCD4] transition-colors ${
                        isColSelected
                          ? 'bg-[#D6EFE0] text-[#107C41]'
                          : isSortedCol || isColFiltered
                            ? 'bg-[#E6F4EC] text-[#0E6234]'
                            : 'bg-[#F0F5F2] text-slate-600 hover:bg-[#E4EEE8]'
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
                          className="flex-1 h-full text-center text-[11.5px] font-bold tracking-wider focus:outline-none truncate"
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
                  const isPinnedHeaderVisual = hasHeaderRow && origRowIdx === 0;

                  const rawRowH = activeSheet.rowHeights?.[origRowIdx];
                  const scaledRowH = Math.max(
                    22,
                    Math.round(
                      (rawRowH && rawRowH > 16 ? rawRowH : DEFAULT_ROW_HEIGHT) *
                        zoomScale
                    )
                  );
                  const isFrozenRow = visRowIdx < frozenRowsCount;
                  const isLastFrozenRow =
                    frozenRowsCount > 0 && visRowIdx === frozenRowsCount - 1;
                  const frozenTop = frozenRowTopOffsets.get(visRowIdx);

                  return (
                    <tr
                      key={`${activeSheet.id}-orig-row-${origRowIdx}`}
                      style={{ height: `${scaledRowH}px` }}
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
                        className={`sticky left-0 z-10 px-1.5 border-b border-r border-[#D6E0D9] text-center text-xs tabular-nums cursor-pointer transition-colors ${
                          isRowSelected
                            ? 'bg-[#D6EFE0] text-[#107C41] font-bold'
                            : isPinnedHeaderVisual
                              ? 'bg-[#EAF1EC] text-slate-600 font-semibold'
                              : 'bg-[#F5F8F6] text-slate-500 font-medium group-hover:bg-[#EBF2ED]'
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
                        const evaluatedValue = rowItem.evaluatedCells[colIdx];
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
                              : 'text-left';

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
                          !isEditingLocked &&
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
                        const cellPeers = peerPresenceByCellCoord.get(cellCoord);
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
                            8,
                            Math.min(
                              64,
                              Math.round(style.fontSize * zoomScale * 10) / 10
                            )
                          )}px`;
                        }
                        if (style?.fontFamily) {
                          cellInlineStyle.fontFamily = style.fontFamily;
                        }
                        if (style?.italic) {
                          cellInlineStyle.fontStyle = 'italic';
                        }
                        if (style?.underline) {
                          cellInlineStyle.textDecoration = 'underline';
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

                        const rawW =
                          activeSheet.colWidths?.[colIdx] ?? DEFAULT_COL_WIDTH;
                        const scaledW = Math.max(
                          48,
                          Math.round(rawW * zoomScale)
                        );
                        if (!mergeSpan) {
                          cellInlineStyle.width = `${scaledW}px`;
                          cellInlineStyle.maxWidth = `${scaledW}px`;
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
                              isFormula
                                ? `Công thức: ${String(rawValue)}`
                                : undefined
                            }
                            style={cellInlineStyle}
                            className={`relative border-b border-r border-[#E4EBE6] transition-colors ${
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
                                      : 'bg-[#E6F4EC] ring-2 ring-inset ring-[#107C41]'
                                    : hasCustomBg
                                      ? ''
                                      : 'bg-[#E6F4EC]/85'
                                  : isInFillPreview
                                    ? 'bg-[#107C41]/10 outline outline-1 outline-dashed outline-[#107C41]'
                                    : hasCustomBg
                                      ? ''
                                      : isPinnedHeaderVisual
                                        ? 'bg-[#F3F8F5] group-hover:bg-[#EBF3EE]'
                                        : 'bg-white group-hover:bg-[#F9FCFA]'
                            }`}
                          >
                            {isSelected && hasCustomBg && !isEditingThisCell && (
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
                                style={{ height: `${scaledRowH}px` }}
                                className={`w-full px-2.5 bg-white text-slate-900 focus:outline-none select-text ${alignClass} ${
                                  style?.bold ? 'font-bold' : 'font-normal'
                                }`}
                              />
                            ) : (
                              <div className="flex items-center justify-between gap-1 w-full min-w-0">
                                <span
                                  className={`flex-1 min-w-0 ${
                                    style?.wrapText ? '' : 'truncate'
                                  }`}
                                >
                                  {displayValue}
                                </span>

                                {/* Data Validation In-Cell Dropdown Trigger */}
                                {hasValidation &&
                                  isAnchor &&
                                  !isEditingLocked && (
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
                                      className="w-5 h-5 rounded bg-[#107C41] text-white inline-flex items-center justify-center shrink-0 hover:bg-[#0D6535] shadow-sm"
                                    >
                                      <ChevronDown className="w-3 h-3" />
                                    </button>
                                  )}
                              </div>
                            )}

                            {/* Data Validation Popover List */}
                            {hasValidation &&
                              openValidationKey === cellKey &&
                              !isEditingLocked && (
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
                                        commitCellValue(origRowIdx, colIdx, opt);
                                        setOpenValidationKey(null);
                                      }}
                                      className={`w-full text-left px-3 py-1.5 text-xs hover:bg-[#E6F4EC] hover:text-[#107C41] transition-colors ${
                                        String(evaluatedValue ?? '') === opt
                                          ? 'bg-[#F0F9F3] text-[#107C41] font-bold'
                                          : 'text-slate-700'
                                      }`}
                                    >
                                      {opt}
                                    </button>
                                  ))}
                                </div>
                              )}

                            {/* Formula corner triangle */}
                            {isFormula && !isEditingThisCell && (
                              <span
                                aria-hidden="true"
                                className="absolute top-0 right-0 w-0 h-0 border-t-[6px] border-t-[#107C41] border-l-[6px] border-l-transparent pointer-events-none"
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
                                className="absolute -bottom-[4px] -right-[4px] w-2.5 h-2.5 bg-[#107C41] border border-white cursor-crosshair z-20 shadow-sm hover:scale-125 transition-transform"
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
              isEditingLocked={isEditingLocked}
              onChangeChartType={handleChangeChartType}
              onDeleteChart={handleDeleteChart}
            />
          )}
        </div>

        <CustomVerticalScrollbar
          scrollContainerRef={gridContainerRef}
          theme="light"
          getDragLabel={({ percent }) => `${Math.round(percent)}%`}
        />
      </div>

      {/* Sheet Tabs Bar: Add (+), Rename, Duplicate, Color, Drag-Reorder, Delete */}
      <div className="shrink-0 border-t border-[#D6E0D9] bg-[#EFF4F1] px-2.5 sm:px-3.5 py-1.5 flex items-center justify-between gap-2">
        <div className="flex items-center gap-1 overflow-x-auto no-scrollbar min-w-0">
          <div
            role="tablist"
            aria-label="Danh sách trang tính"
            className="flex items-center gap-1 shrink-0"
          >
            {sheets.map((sheet, idx) => {
              const isActive = idx === safeSheetIdx;
              const isRenaming = renamingSheetIdx === idx;
              const accentColor = sheet.tabColor ?? (isActive ? '#107C41' : undefined);

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
                    accentColor
                      ? { borderBottomColor: accentColor, borderBottomWidth: '3px' }
                      : undefined
                  }
                  className={`inline-flex items-center gap-1 rounded-t-lg rounded-b-sm text-xs whitespace-nowrap shrink-0 transition-all border ${
                    isActive
                      ? 'bg-white text-[#107C41] font-bold border-[#C7D8CD] shadow-sm'
                      : 'bg-[#E5ECE8]/70 text-slate-600 font-medium border-transparent hover:bg-white/80 hover:text-slate-900'
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
            className="inline-flex items-center justify-center w-7 h-7 rounded-lg bg-white/80 hover:bg-white border border-dashed border-[#AEC3B5] hover:border-[#107C41] text-[#107C41] disabled:opacity-40 shrink-0 transition-colors"
          >
            <Plus className="w-4 h-4" />
          </button>
        </div>

        <span className="hidden md:inline text-[11px] text-slate-400 whitespace-nowrap shrink-0">
          Chuột phải ô/dòng/cột để thêm menu • Kéo góc xanh lá để AutoFill
        </span>
      </div>

      {/* Bottom Edge: OfficeStatusBar (variant="excel") with Average / Count / Sum & 25%-200% Zoom */}
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
    </div>
  );
}

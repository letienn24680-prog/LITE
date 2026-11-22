import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  ChangeEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent as ReactMouseEvent,
} from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDown,
  ArrowUp,
  ArrowUpDown,
  Bold,
  Calculator,
  Check,
  ChevronDown,
  ChevronUp,
  Download,
  FileSpreadsheet,
  Filter,
  Plus,
  RotateCcw,
  Save,
  Search,
  Sigma,
  Table2,
  Trash2,
  X,
} from 'lucide-react';
import type {
  CellStyle,
  DocumentItem,
  ModalType,
  SpreadsheetSheet,
} from '../../types/document';
import { createThumbnailDataUri } from '../../data/mockDocuments';
import {
  columnIndexToLetter,
  evaluateSheetMatrix,
  formatCellRef,
} from '../../utils/formulaEvaluator';
import { exportExcelToXlsx } from '../../utils/fileExporter';
import type {
  CollaboratorPresence,
  DocRealtimePatch,
} from '../../services/realtimeCollab';
import { CustomVerticalScrollbar } from '../ui/CustomVerticalScrollbar';
import { WorkspaceTopBar } from './WorkspaceTopBar';

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

const NUMBER_FORMATTER = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 2,
});

const DEFAULT_NEW_SHEET_ROWS = 10;
const DEFAULT_NEW_SHEET_COLS = 6;

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

  const nextStyles: Record<string, CellStyle> = {};
  if (sheet.cellStyles) {
    for (const [key, style] of Object.entries(sheet.cellStyles)) {
      const parts = key.split(':');
      if (parts.length === 2) {
        const r = Number(parts[0]);
        const c = Number(parts[1]);
        if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) {
          const targetRow = shouldPrependHeader ? r + 1 : r;
          nextStyles[makeCellKey(targetRow, c)] = { ...style };
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

  return {
    id: sheet.id || `sheet-${fallbackIdx + 1}`,
    name: sheet.name || `Sheet ${fallbackIdx + 1}`,
    columns: createLetterColumns(colCount),
    rows: normalizedRows,
    cellStyles: nextStyles,
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

function formatCellDisplay(value: string | number | undefined): string {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? formatNumber(value) : '';
  }
  return value;
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

function remapCellStylesAfterRowDelete(
  styles: Record<string, CellStyle> | undefined,
  deletedRowIndices: Set<number>
): Record<string, CellStyle> {
  if (!styles) {
    return {};
  }
  const sortedDeleted = Array.from(deletedRowIndices).sort((a, b) => a - b);
  const next: Record<string, CellStyle> = {};

  for (const [key, style] of Object.entries(styles)) {
    const parts = key.split(':');
    if (parts.length !== 2) {
      continue;
    }
    const rowIdx = Number(parts[0]);
    const colIdx = Number(parts[1]);
    if (!Number.isInteger(rowIdx) || !Number.isInteger(colIdx)) {
      continue;
    }
    if (deletedRowIndices.has(rowIdx)) {
      continue;
    }
    let shift = 0;
    for (const delIdx of sortedDeleted) {
      if (delIdx < rowIdx) {
        shift += 1;
      } else {
        break;
      }
    }
    next[makeCellKey(rowIdx - shift, colIdx)] = { ...style };
  }

  return next;
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
  const isReadOnly = collabPermission === 'view';

  const [sheets, setSheets] = useState<SpreadsheetSheet[]>(() =>
    initializeSheets(document)
  );
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [searchQuery, setSearchQuery] = useState('');
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [selection, setSelection] = useState<CellSelection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 },
  });
  const [editingCell, setEditingCell] = useState<EditingCellState | null>(null);
  const [editDraft, setEditDraft] = useState('');
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false);
  const [justSaved, setJustSaved] = useState(false);
  const [isMobileToolbarExpanded, setIsMobileToolbarExpanded] = useState(false);

  const sheetsRef = useRef<SpreadsheetSheet[]>(sheets);
  sheetsRef.current = sheets;
  const lastAppliedPatchAtRef = useRef<number>(0);
  const isDraggingRef = useRef(false);
  const cancelBlurCommitRef = useRef(false);
  const saveTimerRef = useRef<number | null>(null);
  const gridContainerRef = useRef<HTMLDivElement | null>(null);
  const cellInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    const initialized = initializeSheets(document);
    sheetsRef.current = initialized;
    lastAppliedPatchAtRef.current = 0;
    setSheets(initialized);
    setActiveSheetIdx(0);
    setSearchQuery('');
    setSortConfig(null);
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
    setEditingCell(null);
    setEditDraft('');
    setHasUnsavedChanges(false);
    setJustSaved(false);
    setIsMobileToolbarExpanded(false);
  }, [document.id]);

  // Apply incoming realtime excel_cell_patch without disrupting local selection or active editing on other cells
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
    const handleMouseUp = () => {
      isDraggingRef.current = false;
    };
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, []);

  useEffect(() => {
    return () => {
      if (saveTimerRef.current !== null) {
        window.clearTimeout(saveTimerRef.current);
      }
    };
  }, []);

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

  const evaluatedMatrix = useMemo<(string | number)[][]>(
    () => evaluateSheetMatrix(activeSheet.rows),
    [activeSheet.rows]
  );

  // Detect whether Row 1 (index 0) acts as a header row (all bold or pure text header above numeric data)
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

    if (nonEmptyFirst === 0) {
      return false;
    }
    if (boldFirst === nonEmptyFirst) {
      return true;
    }
    if (numericFirst === 0) {
      // Check whether any subsequent row has numeric values
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

    for (let colIdx = 0; colIdx < columnCount; colIdx += 1) {
      let nonEmpty = 0;
      let numericCount = 0;
      for (let r = startRow; r < evaluatedMatrix.length; r += 1) {
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
  }, [columnCount, evaluatedMatrix, hasHeaderRow]);

  const visibleRows = useMemo<IndexedRow[]>(() => {
    const normalizedQuery = searchQuery.trim().toLowerCase();
    const indexed: IndexedRow[] = activeSheet.rows.map((rawCells, idx) => ({
      originalIndex: idx,
      rawCells,
      evaluatedCells: evaluatedMatrix[idx] ?? rawCells,
    }));

    const matchesQuery = (row: IndexedRow): boolean => {
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
        if (num !== null && formatNumber(num).toLowerCase().includes(normalizedQuery)) {
          return true;
        }
      }
      return false;
    };

    let filtered: IndexedRow[];
    if (normalizedQuery) {
      if (hasHeaderRow && indexed.length > 1) {
        const dataMatches = indexed.slice(1).filter(matchesQuery);
        if (dataMatches.length > 0) {
          filtered = [indexed[0], ...dataMatches];
        } else {
          filtered = matchesQuery(indexed[0]) ? [indexed[0]] : [];
        }
      } else {
        filtered = indexed.filter(matchesQuery);
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

      if (emptyA && emptyB) {
        return rowA.originalIndex - rowB.originalIndex;
      }
      if (emptyA) {
        return 1;
      }
      if (emptyB) {
        return -1;
      }

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

      if (diff !== 0) {
        return diff * factor;
      }
      return rowA.originalIndex - rowB.originalIndex;
    });

    return pinnedHeader ? [pinnedHeader, ...sorted] : sorted;
  }, [
    activeSheet.rows,
    columnCount,
    evaluatedMatrix,
    hasHeaderRow,
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
      if (!rowData) {
        continue;
      }
      for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
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
  }, [normalizedRange, visibleRows]);

  const selectedOriginalRowSet = useMemo<Set<number>>(() => {
    const set = new Set<number>();
    if (!normalizedRange) {
      return set;
    }
    for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
      const orig = visibleRows[r]?.originalIndex;
      if (orig !== undefined) {
        set.add(orig);
      }
    }
    return set;
  }, [normalizedRange, visibleRows]);

  const canDeleteSelectedRows =
    selectedOriginalRowSet.size > 0 &&
    activeSheet.rows.length > 1 &&
    activeSheet.rows.length - selectedOriginalRowSet.size >= 1;

  const isFiltered = searchQuery.trim().length > 0;
  const isSorted = sortConfig !== null;
  const hasActiveTransforms = isFiltered || isSorted;

  const buildLiveDocument = useCallback(
    (nextSheets: SpreadsheetSheet[]): DocumentItem => {
      const sheetNames = nextSheets.map((s) => s.name).join(', ');
      return {
        ...document,
        sheets: nextSheets,
        pageCount: nextSheets.length,
        contentPreview: `Bảng tính gồm ${nextSheets.length} trang tính (${sheetNames}).`,
      };
    },
    [document]
  );

  const broadcastExcelPatch = useCallback(
    (
      sheetId: string,
      rowIndex: number,
      colIndex: number,
      value: string | number,
      nextSheets: SpreadsheetSheet[]
    ) => {
      if (isReadOnly || !onBroadcastPatch) {
        return;
      }
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

  // Broadcast active cell coordinate & typing status via Presence
  useEffect(() => {
    if (
      !onPresenceUpdate ||
      anchorOriginalRow === null ||
      anchorCol === null
    ) {
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
    if (!collaborators || collaborators.length === 0) {
      return map;
    }
    for (const peer of collaborators) {
      if (peer.sessionId === selfSessionId || !peer.cellCoord) {
        continue;
      }
      if (peer.sheetId && peer.sheetId !== activeSheet.id) {
        continue;
      }
      const coordKey = peer.cellCoord.toUpperCase();
      const existing = map.get(coordKey);
      if (existing) {
        existing.push(peer);
      } else {
        map.set(coordKey, [peer]);
      }
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
      if (isReadOnly) {
        return;
      }
      const prevSheets = sheetsRef.current;
      const currentSheet = prevSheets[safeSheetIdx];
      if (!currentSheet) {
        return;
      }
      const updatedSheet = updater(currentSheet);
      if (updatedSheet === currentSheet) {
        return;
      }
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
    [broadcastExcelPatch, isReadOnly, safeSheetIdx]
  );

  const commitCellValue = useCallback(
    (originalRow: number, col: number, draft: string) => {
      if (isReadOnly) {
        return;
      }
      const normalized = normalizeCommittedValue(draft);
      const existing = activeSheet.rows[originalRow]?.[col] ?? '';
      if (existing === normalized) {
        return;
      }
      updateActiveSheet(
        (sheet) => {
          if (originalRow < 0 || originalRow >= sheet.rows.length) {
            return sheet;
          }
          const nextRows = sheet.rows.map((row, rIdx) => {
            if (rIdx !== originalRow) {
              return row;
            }
            const copy = [...row];
            copy[col] = normalized;
            return copy;
          });
          return { ...sheet, rows: nextRows };
        },
        { rowIndex: originalRow, colIndex: col, value: normalized }
      );
    },
    [activeSheet.rows, isReadOnly, updateActiveSheet]
  );

  const startCellEditing = useCallback(
    (
      originalRow: number,
      col: number,
      source: 'cell' | 'formula',
      initialOverride?: string
    ) => {
      if (isReadOnly) {
        return;
      }
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
    [activeSheet.rows, isReadOnly]
  );

  const cancelEditing = useCallback(() => {
    cancelBlurCommitRef.current = true;
    setEditingCell(null);
    setEditDraft('');
  }, []);

  const finishEditing = useCallback(
    (moveAfter?: 'down' | 'right' | 'left') => {
      if (!editingCell) {
        return;
      }
      const { originalRow, col } = editingCell;
      commitCellValue(originalRow, col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');

      if (moveAfter && normalizedRange) {
        const maxRow = Math.max(visibleRows.length - 1, 0);
        const maxCol = Math.max(columnCount - 1, 0);
        let nextRow = normalizedRange.anchorRow;
        let nextCol = normalizedRange.anchorCol;

        if (moveAfter === 'down') {
          nextRow = Math.min(maxRow, nextRow + 1);
        } else if (moveAfter === 'right') {
          nextCol = Math.min(maxCol, nextCol + 1);
        } else if (moveAfter === 'left') {
          nextCol = Math.max(0, nextCol - 1);
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
      columnCount,
      commitCellValue,
      editDraft,
      editingCell,
      normalizedRange,
      visibleRows.length,
    ]
  );

  const handleSave = useCallback(() => {
    if (isReadOnly) {
      return;
    }
    let latestSheets = sheets;
    if (editingCell) {
      const { originalRow, col } = editingCell;
      const normalized = normalizeCommittedValue(editDraft);
      const existingVal = sheets[safeSheetIdx]?.rows[originalRow]?.[col] ?? '';
      latestSheets = sheets.map((sheet, idx) => {
        if (idx !== safeSheetIdx) {
          return sheet;
        }
        const nextRows = sheet.rows.map((row, rIdx) => {
          if (rIdx !== originalRow) {
            return row;
          }
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
      if (isReadOnly) {
        return;
      }
      let latestSheets = sheets;
      if (editingCell) {
        const { originalRow, col } = editingCell;
        const normalized = normalizeCommittedValue(editDraft);
        latestSheets = sheets.map((sheet, idx) => {
          if (idx !== safeSheetIdx) {
            return sheet;
          }
          const nextRows = sheet.rows.map((row, rIdx) => {
            if (rIdx !== originalRow) {
              return row;
            }
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
        thumbnail: createThumbnailDataUri(document.type, newTitle, '#15803D'),
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
      if (isModalOpen || isReadOnly) {
        return;
      }
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
        if (idx !== safeSheetIdx) {
          return sheet;
        }
        const nextRows = sheet.rows.map((row, rIdx) => {
          if (rIdx !== originalRow) {
            return row;
          }
          const copy = [...row];
          copy[col] = normalized;
          return copy;
        });
        return { ...sheet, rows: nextRows };
      });
    }
    exportExcelToXlsx(document.title, exportSheets);
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
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
  };

  const handleAddSheet = () => {
    if (isReadOnly) {
      return;
    }
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
    const nextIndex = currentSheets.length;
    const nextSheets = [...currentSheets, newSheet];
    sheetsRef.current = nextSheets;
    setSheets(nextSheets);
    setActiveSheetIdx(nextIndex);
    setSearchQuery('');
    setSortConfig(null);
    setSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
    setHasUnsavedChanges(true);
    setJustSaved(false);
    broadcastExcelPatch(newSheet.id, 0, 0, '', nextSheets);
  };

  const handleAddRow = () => {
    if (isReadOnly) {
      return;
    }
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    const newRowIndex = activeSheet.rows.length;
    updateActiveSheet(
      (sheet) => ({
        ...sheet,
        rows: [
          ...sheet.rows,
          Array.from({ length: sheet.columns.length }, () => ''),
        ],
      }),
      { rowIndex: newRowIndex, colIndex: 0, value: '' }
    );
    if (searchQuery.trim()) {
      setSearchQuery('');
    }
    if (sortConfig !== null) {
      setSortConfig(null);
    }
    setSelection({
      start: { row: newRowIndex, col: 0 },
      end: { row: newRowIndex, col: 0 },
    });
  };

  const handleAddColumn = () => {
    if (isReadOnly) {
      return;
    }
    if (editingCell) {
      commitCellValue(editingCell.originalRow, editingCell.col, editDraft);
      cancelBlurCommitRef.current = true;
      setEditingCell(null);
      setEditDraft('');
    }
    const newColIndex = columnCount;
    updateActiveSheet(
      (sheet) => ({
        ...sheet,
        columns: createLetterColumns(sheet.columns.length + 1),
        rows: sheet.rows.map((row) => [...row, '']),
      }),
      { rowIndex: 0, colIndex: newColIndex, value: '' }
    );
    setSelection({
      start: { row: 0, col: newColIndex },
      end: { row: 0, col: newColIndex },
    });
  };

  const handleDeleteSelectedRows = () => {
    if (isReadOnly || !canDeleteSelectedRows) {
      return;
    }
    if (editingCell) {
      cancelEditing();
    }
    const toDelete = new Set(selectedOriginalRowSet);
    updateActiveSheet((sheet) => {
      const remainingRows = sheet.rows.filter((_, rIdx) => !toDelete.has(rIdx));
      if (remainingRows.length === 0) {
        return sheet;
      }
      return {
        ...sheet,
        rows: remainingRows,
        cellStyles: remapCellStylesAfterRowDelete(sheet.cellStyles, toDelete),
      };
    });
    const nextRow = Math.max(0, (normalizedRange?.minRow ?? 0) - 1);
    setSelection({
      start: { row: nextRow, col: normalizedRange?.minCol ?? 0 },
      end: { row: nextRow, col: normalizedRange?.minCol ?? 0 },
    });
  };

  const handleToggleBold = () => {
    if (isReadOnly || !normalizedRange) {
      return;
    }
    const targetBold = !anchorStyle.bold;
    updateActiveSheet((sheet) => {
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
      };
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origRow = visibleRows[r]?.originalIndex;
        if (origRow === undefined) {
          continue;
        }
        for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
          const key = makeCellKey(origRow, c);
          nextStyles[key] = { ...nextStyles[key], bold: targetBold };
        }
      }
      return { ...sheet, cellStyles: nextStyles };
    });
  };

  const handleSetAlign = (align: 'left' | 'center' | 'right') => {
    if (isReadOnly || !normalizedRange) {
      return;
    }
    updateActiveSheet((sheet) => {
      const nextStyles: Record<string, CellStyle> = {
        ...(sheet.cellStyles ?? {}),
      };
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origRow = visibleRows[r]?.originalIndex;
        if (origRow === undefined) {
          continue;
        }
        for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
          const key = makeCellKey(origRow, c);
          nextStyles[key] = { ...nextStyles[key], align };
        }
      }
      return { ...sheet, cellStyles: nextStyles };
    });
  };

  const handleResetFiltersAndSort = () => {
    setSearchQuery('');
    setSortConfig(null);
  };

  const handleToggleSort = (colIdx: number) => {
    setSortConfig((prev) => {
      if (!prev || prev.colIndex !== colIdx) {
        return { colIndex: colIdx, direction: 'asc' };
      }
      if (prev.direction === 'asc') {
        return { colIndex: colIdx, direction: 'desc' };
      }
      return null;
    });
  };

  const handleSelectColumn = (colIdx: number, extendRange: boolean) => {
    if (visibleRows.length === 0) {
      return;
    }
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
    if (columnCount === 0) {
      return;
    }
    const lastCol = columnCount - 1;
    if (extendRange && normalizedRange) {
      setSelection({
        start: { row: normalizedRange.anchorRow, col: 0 },
        end: { row: rowIdx, col: lastCol },
      });
      return;
    }
    setSelection({
      start: { row: rowIdx, col: 0 },
      end: { row: rowIdx, col: lastCol },
    });
  };

  const handleSelectAllVisible = () => {
    if (visibleRows.length === 0 || columnCount === 0) {
      return;
    }
    setSelection({
      start: { row: 0, col: 0 },
      end: { row: visibleRows.length - 1, col: columnCount - 1 },
    });
  };

  const handleQuickSelectNumericColumn = () => {
    if (numericColumnIndices.length === 0 || visibleRows.length === 0) {
      return;
    }
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

  const handleCellMouseDown = (
    event: ReactMouseEvent<HTMLTableCellElement>,
    visRowIdx: number,
    colIdx: number
  ) => {
    if (event.button !== 0) {
      return;
    }
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
    if (!isDraggingRef.current) {
      return;
    }
    setSelection((prev) => ({
      start: prev.start,
      end: { row: visRowIdx, col: colIdx },
    }));
  };

  const handleCellDoubleClick = (visRowIdx: number, colIdx: number) => {
    if (isReadOnly) {
      return;
    }
    const origRow = visibleRows[visRowIdx]?.originalIndex;
    if (origRow === undefined) {
      return;
    }
    setSelection({
      start: { row: visRowIdx, col: colIdx },
      end: { row: visRowIdx, col: colIdx },
    });
    startCellEditing(origRow, colIdx, 'cell');
  };

  const handleClearSelectedCells = () => {
    if (isReadOnly || !normalizedRange) {
      return;
    }
    updateActiveSheet((sheet) => {
      const nextRows = sheet.rows.map((row) => [...row]);
      let changed = false;
      for (let r = normalizedRange.minRow; r <= normalizedRange.maxRow; r += 1) {
        const origRow = visibleRows[r]?.originalIndex;
        if (origRow === undefined || !nextRows[origRow]) {
          continue;
        }
        for (let c = normalizedRange.minCol; c <= normalizedRange.maxCol; c += 1) {
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
    if (isModalOpen || editingCell !== null) {
      return;
    }
    if (visibleRows.length === 0 || columnCount === 0 || !normalizedRange) {
      return;
    }

    if (event.key === 'Enter' || event.key === 'F2') {
      if (isReadOnly) {
        return;
      }
      event.preventDefault();
      if (anchorOriginalRow !== null && anchorCol !== null) {
        startCellEditing(anchorOriginalRow, anchorCol, 'cell');
      }
      return;
    }

    if (event.key === 'Backspace' || event.key === 'Delete') {
      if (isReadOnly) {
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
      if (anchorOriginalRow !== null && anchorCol !== null) {
        event.preventDefault();
        startCellEditing(anchorOriginalRow, anchorCol, 'cell', event.key);
      }
      return;
    }

    let deltaRow = 0;
    let deltaCol = 0;
    if (event.key === 'ArrowUp') {
      deltaRow = -1;
    } else if (event.key === 'ArrowDown') {
      deltaRow = 1;
    } else if (event.key === 'ArrowLeft') {
      deltaCol = -1;
    } else if (event.key === 'ArrowRight') {
      deltaCol = 1;
    } else if (event.key === 'Tab') {
      deltaCol = event.shiftKey ? -1 : 1;
    } else {
      return;
    }

    event.preventDefault();
    const maxRow = visibleRows.length - 1;
    const maxCol = columnCount - 1;

    setSelection((prev) => {
      const clampedStartRow = Math.max(0, Math.min(maxRow, prev.start.row));
      const clampedStartCol = Math.max(0, Math.min(maxCol, prev.start.col));
      const clampedEndRow = Math.max(0, Math.min(maxRow, prev.end.row));
      const clampedEndCol = Math.max(0, Math.min(maxCol, prev.end.col));
      const targetRow = Math.max(0, Math.min(maxRow, clampedEndRow + deltaRow));
      const targetCol = Math.max(0, Math.min(maxCol, clampedEndCol + deltaCol));

      if (event.shiftKey && event.key !== 'Tab') {
        return {
          start: { row: clampedStartRow, col: clampedStartCol },
          end: { row: targetRow, col: targetCol },
        };
      }
      return {
        start: { row: targetRow, col: targetCol },
        end: { row: targetRow, col: targetCol },
      };
    });
  };

  const handleQuickInsertFormula = (formulaPrefix: string) => {
    if (isReadOnly || anchorOriginalRow === null || anchorCol === null) {
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

  const handleFormulaFocus = () => {
    if (isReadOnly || anchorOriginalRow === null || anchorCol === null) {
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
    if (isReadOnly) {
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

  const topBarToolbar = (
    <div className="flex items-center gap-1.5 sm:gap-2">
      {hasActiveTransforms && (
        <button
          type="button"
          onClick={handleResetFiltersAndSort}
          className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200/80 text-xs font-medium transition-colors whitespace-nowrap shrink-0"
          title="Đặt lại bộ lọc và sắp xếp"
        >
          <RotateCcw className="w-3.5 h-3.5 shrink-0" />
          <span className="hidden md:inline">Đặt lại</span>
        </button>
      )}

      {!isReadOnly && (
        <button
          type="button"
          onClick={handleSave}
          title="Lưu thay đổi (Ctrl+S)"
          className={`inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap shrink-0 ${
            justSaved
              ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
              : hasUnsavedChanges
                ? 'bg-[#15803D] hover:bg-[#166534] border-[#15803D] text-white shadow-sm'
                : 'bg-[#F5F7F5] hover:bg-[#E9EFE9] border-[#D8E1D8] text-slate-700'
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
        onClick={handleExportXlsx}
        title="Tải bảng tính xuống tệp .xlsx"
        className="inline-flex items-center gap-1.5 px-2.5 sm:px-3 py-1.5 rounded-xl bg-white hover:bg-[#F3F6F3] border border-[#D8E1D8] text-slate-700 hover:text-[#15803D] text-xs font-semibold transition-colors whitespace-nowrap shrink-0"
      >
        <Download className="w-3.5 h-3.5 shrink-0" />
        <span className="hidden md:inline">Tải xuống .xlsx</span>
        <span className="md:hidden">.xlsx</span>
      </button>
    </div>
  );

  return (
    <div className="h-[100dvh] w-full flex flex-col bg-[#F5F7F5] text-slate-900 overflow-hidden">
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

      {/* Thanh công cụ chỉnh sửa cấu trúc & định dạng ô (Desktop md+) */}
      <div className="hidden md:block shrink-0 bg-[#F8FAF8] border-b border-[#E3E7E3] px-3 sm:px-4 py-1.5">
        <div className="flex items-center justify-between gap-2 overflow-x-auto no-scrollbar">
          <div className="flex items-center gap-1.5 shrink-0">
            {/* Thêm / Xóa dòng & cột */}
            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddRow}
              disabled={isReadOnly}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-[#ECFDF3] border border-[#DCE3DC] hover:border-[#15803D]/40 text-xs font-medium text-slate-700 hover:text-[#15803D] disabled:opacity-40 disabled:pointer-events-none transition-colors whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5 shrink-0" />
              <span>Thêm dòng</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddColumn}
              disabled={isReadOnly}
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-[#ECFDF3] border border-[#DCE3DC] hover:border-[#15803D]/40 text-xs font-medium text-slate-700 hover:text-[#15803D] disabled:opacity-40 disabled:pointer-events-none transition-colors whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5 shrink-0" />
              <span>Thêm cột</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleDeleteSelectedRows}
              disabled={isReadOnly || !canDeleteSelectedRows}
              title={
                isReadOnly
                  ? 'Chế độ chỉ xem'
                  : canDeleteSelectedRows
                    ? 'Xóa các dòng đang chọn'
                    : 'Cần giữ lại ít nhất 1 dòng trong bảng tính'
              }
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-rose-50 border border-[#DCE3DC] hover:border-rose-200 text-xs font-medium text-slate-700 hover:text-rose-600 disabled:opacity-40 disabled:hover:bg-white disabled:hover:text-slate-700 disabled:hover:border-[#DCE3DC] transition-colors whitespace-nowrap shrink-0"
            >
              <Trash2 className="w-3.5 h-3.5 shrink-0" />
              <span>Xóa dòng chọn</span>
            </button>

            <span className="h-4 w-px bg-[#D7DFD7] mx-0.5 shrink-0" aria-hidden="true" />

            {/* Định dạng ô: In đậm & Căn lề */}
            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleToggleBold}
              disabled={isReadOnly || !normalizedRange}
              aria-pressed={anchorStyle.bold === true}
              title="In đậm ô/vùng đang chọn"
              className={`inline-flex items-center justify-center w-7 h-7 rounded-lg border text-xs transition-colors shrink-0 disabled:opacity-40 ${
                anchorStyle.bold
                  ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                  : 'bg-white border-[#DCE3DC] text-slate-700 hover:bg-[#EEF4EE]'
              }`}
            >
              <Bold className="w-3.5 h-3.5" />
            </button>

            <div
              role="group"
              aria-label="Căn lề ô"
              className="inline-flex items-center rounded-lg bg-white border border-[#DCE3DC] p-0.5 shrink-0"
            >
              <button
                type="button"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => handleSetAlign('left')}
                disabled={isReadOnly || !normalizedRange}
                aria-label="Căn trái"
                title="Căn trái"
                className={`p-1 rounded-md transition-colors disabled:opacity-40 ${
                  anchorEffectiveAlign === 'left'
                    ? 'bg-[#DCFCE7] text-[#15803D]'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <AlignLeft className="w-3.5 h-3.5" />
              </button>
              <button
                type="button"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => handleSetAlign('center')}
                disabled={isReadOnly || !normalizedRange}
                aria-label="Căn giữa"
                title="Căn giữa"
                className={`p-1 rounded-md transition-colors disabled:opacity-40 ${
                  anchorEffectiveAlign === 'center'
                    ? 'bg-[#DCFCE7] text-[#15803D]'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <AlignCenter className="w-3.5 h-3.5" />
              </button>
              <button
                type="button"
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => handleSetAlign('right')}
                disabled={isReadOnly || !normalizedRange}
                aria-label="Căn phải"
                title="Căn phải"
                className={`p-1 rounded-md transition-colors disabled:opacity-40 ${
                  anchorEffectiveAlign === 'right'
                    ? 'bg-[#DCFCE7] text-[#15803D]'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <AlignRight className="w-3.5 h-3.5" />
              </button>
            </div>

            {numericColumnIndices.length > 0 && visibleRows.length > 0 && (
              <>
                <span className="h-4 w-px bg-[#D7DFD7] mx-0.5 shrink-0" aria-hidden="true" />
                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleQuickSelectNumericColumn}
                  title="Chọn nhanh cột số liệu để xem SUM / AVG"
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-[#15803D]/10 hover:bg-[#15803D]/15 border border-[#15803D]/25 text-[#15803D] text-xs font-semibold transition-colors whitespace-nowrap shrink-0"
                >
                  <Sigma className="w-3.5 h-3.5 shrink-0" />
                  <span>Chọn cột số</span>
                </button>
              </>
            )}
          </div>

          <div className="hidden lg:flex items-center gap-2 text-[11px] text-slate-500 shrink-0 whitespace-nowrap">
            {isFiltered && (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-50 border border-emerald-200 text-[#15803D] font-medium">
                <Filter className="w-3 h-3" />
                {visibleRows.length}/{activeSheet.rows.length} dòng
              </span>
            )}
            <span>
              Bấm đúp ô hoặc gõ vào <strong className="font-semibold text-slate-700">fx</strong> (hỗ trợ <code className="text-[#15803D]">=SUM(...)</code>, <code className="text-[#15803D]">=AVERAGE(...)</code>)
            </span>
          </div>
        </div>
      </div>

      {/* Thanh công thức (fx) & Lọc dữ liệu */}
      <div className="shrink-0 bg-white border-b border-[#E3E7E3] px-3 sm:px-4 py-2">
        <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
          <div
            className="min-w-[74px] sm:min-w-[88px] px-2.5 py-1.5 rounded-lg bg-[#F5F7F5] border border-[#DCE3DC] text-xs font-bold text-[#15803D] text-center tabular-nums select-all shrink-0"
            title="Tọa độ ô hoặc vùng đang chọn"
          >
            {rangeLabel}
          </div>

          <div className="flex-1 min-w-[180px] flex items-center gap-2 px-2.5 py-1 rounded-lg bg-[#FAFBF9] border border-[#DCE3DC] focus-within:bg-white focus-within:border-[#15803D] focus-within:ring-2 focus-within:ring-[#15803D]/15 transition-all">
            <span
              className="font-serif italic font-bold text-slate-400 select-none shrink-0 pr-2 border-r border-[#E3E7E3] text-sm"
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
                  : 'Nhập giá trị hoặc công thức (vd: =SUM(D2:D9))...'
              }
              aria-label="Thanh công thức fx"
              className="w-full py-0.5 bg-transparent text-xs sm:text-[13px] font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none disabled:opacity-50"
            />
            {isFormulaText(anchorRawValue) && (
              <span className="hidden md:inline-flex items-center px-2 py-0.5 rounded bg-emerald-50 border border-emerald-200/80 text-[11px] font-semibold text-[#15803D] tabular-nums shrink-0 whitespace-nowrap">
                = {formatCellDisplay(anchorEvaluatedValue)}
              </span>
            )}
          </div>

          <div className="relative w-full sm:w-56 lg:w-64 shrink-0">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
              placeholder="Lọc dữ liệu trong bảng tính..."
              aria-label="Lọc dữ liệu trong bảng tính"
              className="w-full pl-8 pr-8 py-1.5 rounded-lg bg-[#F7F9F7] border border-[#DCE3DC] text-xs sm:text-[13px] text-slate-800 placeholder:text-slate-400 focus:outline-none focus:bg-white focus:border-[#15803D] focus:ring-2 focus:ring-[#15803D]/15 transition-all"
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

      {/* Lưới bảng tính ở giữa */}
      <div className="relative flex-1 min-h-0 flex overflow-hidden">
        <div
          ref={gridContainerRef}
          tabIndex={0}
          onKeyDown={handleGridKeyDown}
          className="w-full h-full overflow-auto scrollbar-none pr-4 bg-white select-none focus:outline-none"
        >
        <table className="w-max min-w-full border-separate border-spacing-0 text-[13px]">
          <thead>
            <tr className="h-8">
              <th
                scope="col"
                className="sticky left-0 top-0 z-30 w-11 sm:w-12 p-0 bg-[#EDF1ED] border-b border-r border-[#D4DCD4]"
              >
                <button
                  type="button"
                  onClick={handleSelectAllVisible}
                  title="Chọn toàn bộ bảng tính đang hiển thị"
                  aria-label="Chọn toàn bộ bảng tính"
                  className="w-full h-8 flex items-center justify-center text-slate-400 hover:text-[#15803D] hover:bg-[#E2EBE2] transition-colors"
                >
                  <Table2 className="w-3.5 h-3.5" />
                </button>
              </th>

              {activeSheet.columns.map((_, colIdx) => {
                const letter = columnIndexToLetter(colIdx);
                const isColSelected =
                  normalizedRange !== null &&
                  colIdx >= normalizedRange.minCol &&
                  colIdx <= normalizedRange.maxCol;
                const isSortedCol = sortConfig?.colIndex === colIdx;
                const sortDir = isSortedCol ? sortConfig.direction : null;

                return (
                  <th
                    key={`col-${colIdx}`}
                    scope="col"
                    aria-sort={
                      sortDir === 'asc'
                        ? 'ascending'
                        : sortDir === 'desc'
                          ? 'descending'
                          : 'none'
                    }
                    className={`sticky top-0 z-20 p-0 border-b border-r border-[#D4DCD4] min-w-[126px] sm:min-w-[148px] transition-colors ${
                      isColSelected
                        ? 'bg-[#DCFCE7] text-[#15803D]'
                        : isSortedCol
                          ? 'bg-[#ECFDF3] text-[#14532D]'
                          : 'bg-[#F1F5F1] text-slate-600 hover:bg-[#E6EFE6]'
                    }`}
                  >
                    <div
                      className={`w-full h-8 flex items-center justify-between pl-2.5 pr-1 ${
                        isColSelected ? 'border-b-2 border-b-[#15803D]' : ''
                      }`}
                    >
                      <button
                        type="button"
                        onClick={(event) =>
                          handleSelectColumn(colIdx, event.shiftKey)
                        }
                        title={`Bấm để chọn cả cột ${letter} (giữ Shift để chọn nhiều cột)`}
                        className="flex-1 h-full text-center text-[11.5px] font-bold tracking-wider focus:outline-none"
                      >
                        {letter}
                      </button>

                      <button
                        type="button"
                        onClick={(event) => {
                          event.stopPropagation();
                          handleToggleSort(colIdx);
                        }}
                        aria-label={`Sắp xếp cột ${letter}`}
                        title={`Sắp xếp cột ${letter} (tăng / giảm / mặc định)`}
                        className={`p-1 rounded transition-colors shrink-0 ${
                          isSortedCol
                            ? 'text-[#15803D] bg-white/80 shadow-sm'
                            : 'text-slate-400 hover:text-slate-700 hover:bg-white/60'
                        }`}
                      >
                        {sortDir === 'asc' ? (
                          <ArrowUp className="w-3 h-3" />
                        ) : sortDir === 'desc' ? (
                          <ArrowDown className="w-3 h-3" />
                        ) : (
                          <ArrowUpDown className="w-3 h-3" />
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
                  colSpan={columnCount + 1}
                  className="py-16 px-4 text-center bg-white"
                >
                  <div className="max-w-sm mx-auto flex flex-col items-center gap-2.5">
                    <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-400">
                      <Search className="w-5 h-5" />
                    </div>
                    <p className="text-sm font-semibold text-slate-700">
                      Không tìm thấy dòng dữ liệu nào phù hợp
                    </p>
                    <p className="text-xs text-slate-500">
                      Từ khóa &ldquo;{searchQuery.trim()}&rdquo; không khớp với
                      ô nào trong &ldquo;{activeSheet.name}&rdquo;.
                    </p>
                    <button
                      type="button"
                      onClick={handleResetFiltersAndSort}
                      className="mt-1 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#15803D] hover:bg-[#166534] text-white text-xs font-semibold transition-colors"
                    >
                      <RotateCcw className="w-3.5 h-3.5" />
                      Xóa bộ lọc
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

                return (
                  <tr
                    key={`${activeSheet.id}-orig-row-${origRowIdx}`}
                    className="group transition-colors"
                  >
                    <td
                      onClick={(event) =>
                        handleSelectRow(visRowIdx, event.shiftKey)
                      }
                      title={`Dòng ${origRowIdx + 1} — bấm chọn cả dòng (giữ Shift để chọn nhiều dòng)`}
                      className={`sticky left-0 z-10 w-11 sm:w-12 h-9 px-2 border-b border-r border-[#DDE3DD] text-center text-xs tabular-nums cursor-pointer transition-colors ${
                        isRowSelected
                          ? 'bg-[#DCFCE7] text-[#15803D] font-bold'
                          : isPinnedHeaderVisual
                            ? 'bg-[#EEF3EE] text-slate-600 font-semibold'
                            : 'bg-[#F6F8F6] text-slate-500 font-medium group-hover:bg-[#EDF2ED]'
                      }`}
                    >
                      {origRowIdx + 1}
                    </td>

                    {activeSheet.columns.map((_, colIdx) => {
                      const rawValue = rowItem.rawCells[colIdx];
                      const evaluatedValue = rowItem.evaluatedCells[colIdx];
                      const numericVal = parseNumericValue(evaluatedValue);
                      const isNumeric = numericVal !== null;
                      const isFormula = isFormulaText(rawValue);
                      const isError = isFormulaError(evaluatedValue);
                      const displayValue = formatCellDisplay(evaluatedValue);

                      const style =
                        activeSheet.cellStyles?.[
                          makeCellKey(origRowIdx, colIdx)
                        ];
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

                      let perimeterShadow: string | undefined;
                      if (
                        isSelected &&
                        !isSingleSelection &&
                        !isEditingThisCell &&
                        normalizedRange
                      ) {
                        const shadows: string[] = [];
                        if (visRowIdx === normalizedRange.minRow) {
                          shadows.push('inset 0 2px 0 0 #16a34a');
                        }
                        if (visRowIdx === normalizedRange.maxRow) {
                          shadows.push('inset 0 -2px 0 0 #16a34a');
                        }
                        if (colIdx === normalizedRange.minCol) {
                          shadows.push('inset 2px 0 0 0 #16a34a');
                        }
                        if (colIdx === normalizedRange.maxCol) {
                          shadows.push('inset -2px 0 0 0 #16a34a');
                        }
                        if (shadows.length > 0) {
                          perimeterShadow = shadows.join(', ');
                        }
                      }

                      return (
                        <td
                          key={`${origRowIdx}-${colIdx}`}
                          onMouseDown={(event) =>
                            handleCellMouseDown(event, visRowIdx, colIdx)
                          }
                          onMouseEnter={() =>
                            handleCellMouseEnter(visRowIdx, colIdx)
                          }
                          onDoubleClick={() =>
                            handleCellDoubleClick(visRowIdx, colIdx)
                          }
                          title={
                            isFormula
                              ? `Công thức: ${String(rawValue)}`
                              : undefined
                          }
                          style={
                            perimeterShadow
                              ? { boxShadow: perimeterShadow }
                              : undefined
                          }
                          className={`relative h-9 border-b border-r border-[#E6EBE6] whitespace-nowrap transition-colors ${
                            isEditingThisCell ? 'p-0 bg-white' : 'px-3 cursor-cell'
                          } ${alignClass} ${
                            isError
                              ? 'text-rose-600 font-semibold'
                              : isNumeric
                                ? 'tabular-nums text-slate-900'
                                : 'text-slate-700'
                          } ${
                            style?.bold
                              ? 'font-bold text-slate-900'
                              : isNumeric
                                ? 'font-medium'
                                : 'font-normal'
                          } ${
                            isEditingThisCell
                              ? 'ring-2 ring-inset ring-green-600 z-10'
                              : isSelected
                                ? isSingleSelection || isAnchor
                                  ? 'bg-green-50 ring-2 ring-inset ring-green-600'
                                  : 'bg-green-50/90'
                                : isPinnedHeaderVisual
                                  ? 'bg-[#F6F9F6] group-hover:bg-[#EEF4EE]'
                                  : 'bg-white group-hover:bg-[#FAFDF9]'
                          }`}
                        >
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
                              className={`w-full h-9 px-3 bg-white text-[13px] text-slate-900 focus:outline-none select-text ${alignClass} ${
                                style?.bold ? 'font-bold' : 'font-normal'
                              }`}
                            />
                          ) : (
                            <>
                              {displayValue}
                              {isFormula && (
                                <span
                                  aria-hidden="true"
                                  className="absolute top-0 right-0 w-0 h-0 border-t-[6px] border-t-emerald-500 border-l-[6px] border-l-transparent pointer-events-none"
                                />
                              )}
                            </>
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
        </div>

        <CustomVerticalScrollbar
          scrollContainerRef={gridContainerRef}
          theme="light"
          getDragLabel={({ percent }) => `${Math.round(percent)}%`}
        />
      </div>

      {/* Thanh đáy cố định: Sheet tabs + Thêm Sheet + Tính nhanh SUM / AVG / COUNT */}
      <div className="shrink-0 border-t border-[#DDE3DD] bg-[#F3F6F3] pb-14 md:pb-safe">
        <div className="px-3 sm:px-4 py-2 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
          {/* Trái: Tabs chuyển trang tính + nút Thêm Sheet */}
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar min-w-0">
            <div
              role="tablist"
              aria-label="Danh sách trang tính"
              className="flex items-center gap-1.5 shrink-0"
            >
              {sheets.map((sheet, idx) => {
                const isActive = idx === safeSheetIdx;
                return (
                  <button
                    key={sheet.id}
                    type="button"
                    role="tab"
                    aria-selected={isActive}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => handleSelectSheet(idx)}
                    className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs whitespace-nowrap shrink-0 transition-all ${
                      isActive
                        ? 'bg-white text-[#15803D] font-bold border border-[#15803D] shadow-sm'
                        : 'text-slate-600 font-medium border border-transparent hover:bg-white/70 hover:text-slate-900'
                    }`}
                  >
                    <FileSpreadsheet
                      className={`w-3.5 h-3.5 shrink-0 ${
                        isActive ? 'text-[#15803D]' : 'text-slate-400'
                      }`}
                    />
                    <span>{sheet.name}</span>
                  </button>
                );
              })}
            </div>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddSheet}
              disabled={isReadOnly}
              title={
                isReadOnly
                  ? 'Chế độ chỉ xem'
                  : 'Thêm trang tính mới (10 dòng × 6 cột)'
              }
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-white/80 hover:bg-white border border-dashed border-[#B8C7B8] hover:border-[#15803D] text-xs font-semibold text-slate-600 hover:text-[#15803D] disabled:opacity-40 disabled:pointer-events-none whitespace-nowrap shrink-0 transition-colors"
            >
              <Plus className="w-3.5 h-3.5 shrink-0" />
              <span>Thêm Sheet</span>
            </button>
          </div>

          {/* Phải: Thanh tính nhanh SUM / AVG / COUNT */}
          <div className="flex items-center gap-1.5 sm:gap-2 overflow-x-auto no-scrollbar shrink-0 text-xs">
            {selectionStats.numericCount > 0 ? (
              <>
                <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-[#DCFCE7] border border-[#BBF7D0] text-[#15803D] whitespace-nowrap shrink-0 tabular-nums">
                  <span className="text-[10.5px] font-bold uppercase tracking-wide opacity-80">
                    SUM
                  </span>
                  <span className="font-bold">
                    {formatNumber(selectionStats.sum)}
                  </span>
                </div>

                <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-[#DCE3DC] text-slate-700 whitespace-nowrap shrink-0 tabular-nums">
                  <span className="text-[10.5px] font-bold uppercase tracking-wide text-slate-400">
                    AVG
                  </span>
                  <span className="font-semibold text-slate-900">
                    {selectionStats.avg !== null
                      ? formatNumber(selectionStats.avg)
                      : '—'}
                  </span>
                </div>

                <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-[#DCE3DC] text-slate-700 whitespace-nowrap shrink-0 tabular-nums">
                  <span className="text-[10.5px] font-bold uppercase tracking-wide text-slate-400">
                    COUNT
                  </span>
                  <span className="font-semibold text-slate-900">
                    {selectionStats.nonEmptyCount}
                  </span>
                </div>
              </>
            ) : (
              <>
                <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white border border-[#DCE3DC] text-slate-700 whitespace-nowrap shrink-0 tabular-nums">
                  <span className="text-[10.5px] font-bold uppercase tracking-wide text-slate-400">
                    COUNT
                  </span>
                  <span className="font-semibold text-slate-900">
                    {selectionStats.nonEmptyCount}
                  </span>
                </div>

                {numericColumnIndices.length > 0 && visibleRows.length > 0 && (
                  <button
                    type="button"
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={handleQuickSelectNumericColumn}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-white hover:bg-[#ECFDF3] border border-dashed border-[#15803D]/40 text-[#15803D] font-medium whitespace-nowrap shrink-0 transition-colors"
                  >
                    <Calculator className="w-3.5 h-3.5 shrink-0" />
                    <span>
                      Chọn cột {columnIndexToLetter(numericColumnIndices[0])} để
                      xem SUM / AVG
                    </span>
                  </button>
                )}
              </>
            )}
          </div>
        </div>
      </div>

      {/* Thanh công cụ Mobile (< md): 1 hàng thu gọn ở đáy hoặc Bảng công cụ chi tiết mở rộng */}
      {!isMobileToolbarExpanded ? (
        <div className="fixed bottom-0 inset-x-0 z-40 bg-slate-100 border-t border-slate-200 h-12 flex items-center text-slate-800 md:hidden">
          <div className="overflow-x-auto no-scrollbar flex items-center gap-1 px-2 flex-1 min-w-0">
            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleToggleBold}
              disabled={isReadOnly || !normalizedRange}
              aria-pressed={anchorStyle.bold === true}
              title="In đậm ô"
              className={`inline-flex items-center justify-center h-8 w-8 rounded-lg border text-xs font-bold shrink-0 transition-colors disabled:opacity-40 ${
                anchorStyle.bold
                  ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
            >
              <Bold className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => handleSetAlign('left')}
              disabled={isReadOnly || !normalizedRange}
              aria-label="Căn trái"
              title="Căn trái"
              className={`inline-flex items-center justify-center h-8 w-8 rounded-lg border shrink-0 transition-colors disabled:opacity-40 ${
                anchorEffectiveAlign === 'left'
                  ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => handleSetAlign('center')}
              disabled={isReadOnly || !normalizedRange}
              aria-label="Căn giữa"
              title="Căn giữa"
              className={`inline-flex items-center justify-center h-8 w-8 rounded-lg border shrink-0 transition-colors disabled:opacity-40 ${
                anchorEffectiveAlign === 'center'
                  ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => handleSetAlign('right')}
              disabled={isReadOnly || !normalizedRange}
              aria-label="Căn phải"
              title="Căn phải"
              className={`inline-flex items-center justify-center h-8 w-8 rounded-lg border shrink-0 transition-colors disabled:opacity-40 ${
                anchorEffectiveAlign === 'right'
                  ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                  : 'bg-white border-slate-200 text-slate-700'
              }`}
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddRow}
              disabled={isReadOnly}
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white border border-slate-200 text-xs font-medium text-slate-700 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5 text-[#15803D] shrink-0" />
              <span>Dòng</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddColumn}
              disabled={isReadOnly}
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white border border-slate-200 text-xs font-medium text-slate-700 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5 text-[#15803D] shrink-0" />
              <span>Cột</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleDeleteSelectedRows}
              disabled={isReadOnly || !canDeleteSelectedRows}
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white border border-slate-200 text-xs font-medium text-rose-600 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Trash2 className="w-3.5 h-3.5 shrink-0" />
              <span>Xóa dòng</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleAddSheet}
              disabled={isReadOnly}
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-white border border-slate-200 text-xs font-medium text-slate-700 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Plus className="w-3.5 h-3.5 text-[#15803D] shrink-0" />
              <span>Sheet</span>
            </button>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={handleQuickSelectNumericColumn}
              disabled={
                numericColumnIndices.length === 0 || visibleRows.length === 0
              }
              className="inline-flex items-center gap-1 h-8 px-2.5 rounded-lg bg-[#DCFCE7] border border-[#15803D]/30 text-xs font-semibold text-[#15803D] disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Sigma className="w-3.5 h-3.5 shrink-0" />
              <span>Chọn cột số</span>
            </button>
          </div>

          <button
            type="button"
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => setIsMobileToolbarExpanded(true)}
            aria-label="Mở rộng bảng công cụ chi tiết"
            title="Mở rộng bảng công cụ chi tiết"
            className="h-full px-3.5 border-l border-slate-300 flex items-center justify-center text-slate-700 hover:bg-slate-200/70 active:bg-slate-200 shrink-0 transition-colors"
          >
            <ChevronUp className="w-4 h-4" />
          </button>
        </div>
      ) : (
        <div className="fixed bottom-0 inset-x-0 z-40 h-[44vh] bg-white border-t border-slate-200 rounded-t-2xl shadow-2xl flex flex-col text-slate-900 md:hidden">
          {/* Hàng tiêu đề bảng */}
          <div className="shrink-0 flex items-center justify-between gap-2 px-3.5 py-2.5 border-b border-slate-200 bg-slate-50 rounded-t-2xl">
            <div className="flex items-center gap-2 min-w-0">
              <span className="inline-flex items-center px-2 py-0.5 rounded-md bg-[#DCFCE7] border border-[#15803D]/25 text-xs font-bold text-[#15803D] tabular-nums shrink-0">
                {rangeLabel}
              </span>
              <span className="text-xs font-semibold text-slate-700 truncate">
                {activeSheet.name}
              </span>
            </div>

            <button
              type="button"
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => setIsMobileToolbarExpanded(false)}
              aria-label="Thu gọn bảng công cụ"
              title="Thu gọn bảng công cụ"
              className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-slate-200/80 hover:bg-slate-300/70 text-xs font-semibold text-slate-700 shrink-0 transition-colors"
            >
              <span>Thu gọn</span>
              <ChevronDown className="w-4 h-4" />
            </button>
          </div>

          {/* Nội dung bảng công cụ mở rộng */}
          <div className="flex-1 min-h-0 overflow-y-auto px-3.5 py-3 space-y-3 pb-safe">
            {/* Hàng 1 — Định dạng ô */}
            <div className="space-y-1.5">
              <p className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                Định dạng ô
              </p>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleToggleBold}
                  disabled={isReadOnly || !normalizedRange}
                  aria-pressed={anchorStyle.bold === true}
                  className={`inline-flex items-center gap-1.5 h-9 px-3 rounded-xl border text-xs font-semibold transition-colors disabled:opacity-40 ${
                    anchorStyle.bold
                      ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D]'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  <Bold className="w-3.5 h-3.5" />
                  <span>In đậm</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => handleSetAlign('left')}
                  disabled={isReadOnly || !normalizedRange}
                  className={`inline-flex items-center gap-1 h-9 px-2.5 rounded-xl border text-xs font-medium transition-colors disabled:opacity-40 ${
                    anchorEffectiveAlign === 'left'
                      ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D] font-semibold'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  <AlignLeft className="w-3.5 h-3.5" />
                  <span>Căn trái</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => handleSetAlign('center')}
                  disabled={isReadOnly || !normalizedRange}
                  className={`inline-flex items-center gap-1 h-9 px-2.5 rounded-xl border text-xs font-medium transition-colors disabled:opacity-40 ${
                    anchorEffectiveAlign === 'center'
                      ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D] font-semibold'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  <AlignCenter className="w-3.5 h-3.5" />
                  <span>Căn giữa</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => handleSetAlign('right')}
                  disabled={isReadOnly || !normalizedRange}
                  className={`inline-flex items-center gap-1 h-9 px-2.5 rounded-xl border text-xs font-medium transition-colors disabled:opacity-40 ${
                    anchorEffectiveAlign === 'right'
                      ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D] font-semibold'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  <AlignRight className="w-3.5 h-3.5" />
                  <span>Căn phải</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleClearSelectedCells}
                  disabled={isReadOnly || !normalizedRange}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-rose-50 border border-rose-200 text-xs font-semibold text-rose-700 disabled:opacity-40 transition-colors"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>Xóa nội dung ô</span>
                </button>
              </div>
            </div>

            {/* Hàng 2 — Cấu trúc Dòng & Cột */}
            <div className="space-y-1.5">
              <p className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                Cấu trúc Dòng &amp; Cột
              </p>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleAddRow}
                  disabled={isReadOnly}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-semibold text-slate-700 disabled:opacity-40"
                >
                  <Plus className="w-3.5 h-3.5 text-[#15803D]" />
                  <span>Thêm dòng</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleAddColumn}
                  disabled={isReadOnly}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-semibold text-slate-700 disabled:opacity-40"
                >
                  <Plus className="w-3.5 h-3.5 text-[#15803D]" />
                  <span>Thêm cột</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleDeleteSelectedRows}
                  disabled={isReadOnly || !canDeleteSelectedRows}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-rose-50 border border-rose-200 text-xs font-semibold text-rose-700 disabled:opacity-40"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>Xóa dòng chọn</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleSelectAllVisible}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-semibold text-slate-700"
                >
                  <Table2 className="w-3.5 h-3.5 text-[#15803D]" />
                  <span>Chọn toàn bộ bảng</span>
                </button>
              </div>
            </div>

            {/* Hàng 3 — Công thức & Thống kê nhanh */}
            <div className="space-y-1.5">
              <p className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                Công thức &amp; Thống kê nhanh
              </p>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => handleQuickInsertFormula('=SUM(')}
                  disabled={
                    isReadOnly || anchorOriginalRow === null || anchorCol === null
                  }
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-emerald-50 border border-emerald-200 text-xs font-bold text-[#15803D] disabled:opacity-40"
                >
                  <span>=SUM()</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => handleQuickInsertFormula('=AVERAGE(')}
                  disabled={
                    isReadOnly || anchorOriginalRow === null || anchorCol === null
                  }
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-emerald-50 border border-emerald-200 text-xs font-bold text-[#15803D] disabled:opacity-40"
                >
                  <span>=AVERAGE()</span>
                </button>

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleQuickSelectNumericColumn}
                  disabled={
                    numericColumnIndices.length === 0 ||
                    visibleRows.length === 0
                  }
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-[#DCFCE7] border border-[#15803D]/30 text-xs font-semibold text-[#15803D] disabled:opacity-40"
                >
                  <Sigma className="w-3.5 h-3.5" />
                  <span>Chọn cột số</span>
                </button>

                <span className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-slate-100 border border-slate-200 text-xs font-semibold text-slate-700 tabular-nums">
                  <strong className="text-[#15803D]">SUM:</strong>{' '}
                  {formatNumber(selectionStats.sum)}
                </span>

                <span className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-slate-100 border border-slate-200 text-xs font-semibold text-slate-700 tabular-nums">
                  <strong className="text-slate-500">AVG:</strong>{' '}
                  {selectionStats.avg !== null
                    ? formatNumber(selectionStats.avg)
                    : '—'}
                </span>

                <span className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-slate-100 border border-slate-200 text-xs font-semibold text-slate-700 tabular-nums">
                  <strong className="text-slate-500">COUNT:</strong>{' '}
                  {selectionStats.nonEmptyCount}
                </span>
              </div>
            </div>

            {/* Hàng 4 — Trang tính (Sheets) */}
            <div className="space-y-1.5">
              <p className="text-[11px] font-semibold uppercase tracking-wider text-slate-400">
                Trang tính (Sheets)
              </p>
              <div className="flex flex-wrap items-center gap-1.5">
                {sheets.map((sheet, idx) => {
                  const isActive = idx === safeSheetIdx;
                  return (
                    <button
                      key={`mobile-sheet-${sheet.id}`}
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => handleSelectSheet(idx)}
                      className={`inline-flex items-center gap-1.5 h-9 px-3 rounded-xl border text-xs transition-colors ${
                        isActive
                          ? 'bg-[#DCFCE7] border-[#15803D] text-[#15803D] font-bold'
                          : 'bg-slate-50 border-slate-200 text-slate-700 font-medium'
                      }`}
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
                      <span>{sheet.name}</span>
                    </button>
                  );
                })}

                <button
                  type="button"
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={handleAddSheet}
                  disabled={isReadOnly}
                  className="inline-flex items-center gap-1 h-9 px-3 rounded-xl bg-white border border-dashed border-[#15803D]/50 text-xs font-semibold text-[#15803D] disabled:opacity-40"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>Thêm Sheet</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

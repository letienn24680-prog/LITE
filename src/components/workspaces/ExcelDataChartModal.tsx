import { useMemo, useState } from 'react';
import {
  BarChart3,
  Columns3,
  LineChart,
  ListChecks,
  PieChart,
  Plus,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react';
import type {
  ExcelChartItem,
  ExcelChartSeries,
  ExcelConditionalFormat,
} from '../../types/document';
import {
  columnIndexToLetter,
  letterToColumnIndex,
} from '../../utils/formulaEvaluator';

export type ExcelFormulaCategory =
  | 'math'
  | 'logic'
  | 'text'
  | 'date'
  | 'lookup';

export interface ExcelFormulaMeta {
  name: string;
  syntax: string;
  category: ExcelFormulaCategory;
  description: string;
}

export const EXCEL_FORMULA_CATEGORY_LABELS: Record<
  ExcelFormulaCategory,
  string
> = {
  math: 'Toán & Thống kê',
  logic: 'Logic',
  text: 'Văn bản',
  date: 'Ngày & Giờ',
  lookup: 'Tra cứu',
};

export const EXCEL_FORMULA_CATALOG: readonly ExcelFormulaMeta[] = [
  // Toán & Thống kê
  {
    name: 'SUM',
    syntax: 'SUM(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Tính tổng các giá trị số trong vùng chọn',
  },
  {
    name: 'AVERAGE',
    syntax: 'AVERAGE(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Tính giá trị trung bình cộng của các số',
  },
  {
    name: 'COUNT',
    syntax: 'COUNT(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Đếm số ô chứa giá trị số',
  },
  {
    name: 'COUNTA',
    syntax: 'COUNTA(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Đếm số ô không trống trong vùng',
  },
  {
    name: 'COUNTBLANK',
    syntax: 'COUNTBLANK(vùng)',
    category: 'math',
    description: 'Đếm số ô trống trong vùng',
  },
  {
    name: 'COUNTIF',
    syntax: 'COUNTIF(vùng, điều_kiện)',
    category: 'math',
    description: 'Đếm số ô thỏa mãn điều kiện (vd: ">10")',
  },
  {
    name: 'SUMIF',
    syntax: 'SUMIF(vùng_đk, điều_kiện, [vùng_tổng])',
    category: 'math',
    description: 'Tính tổng các ô thỏa mãn điều kiện',
  },
  {
    name: 'AVERAGEIF',
    syntax: 'AVERAGEIF(vùng_đk, điều_kiện, [vùng_tb])',
    category: 'math',
    description: 'Tính trung bình các ô thỏa mãn điều kiện',
  },
  {
    name: 'MAX',
    syntax: 'MAX(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Tìm giá trị lớn nhất',
  },
  {
    name: 'MIN',
    syntax: 'MIN(vùng_1, [vùng_2], ...)',
    category: 'math',
    description: 'Tìm giá trị nhỏ nhất',
  },
  {
    name: 'ROUND',
    syntax: 'ROUND(số, số_chữ_số)',
    category: 'math',
    description: 'Làm tròn số tới số chữ số thập phân chỉ định',
  },
  {
    name: 'ROUNDUP',
    syntax: 'ROUNDUP(số, số_chữ_số)',
    category: 'math',
    description: 'Làm tròn lên (xa số 0)',
  },
  {
    name: 'ROUNDDOWN',
    syntax: 'ROUNDDOWN(số, số_chữ_số)',
    category: 'math',
    description: 'Làm tròn xuống (gần số 0)',
  },
  {
    name: 'INT',
    syntax: 'INT(số)',
    category: 'math',
    description: 'Làm tròn xuống số nguyên gần nhất',
  },
  {
    name: 'ABS',
    syntax: 'ABS(số)',
    category: 'math',
    description: 'Trả về giá trị tuyệt đối của một số',
  },

  // Logic
  {
    name: 'IF',
    syntax: 'IF(điều_kiện, nếu_đúng, [nếu_sai])',
    category: 'logic',
    description: 'Kiểm tra điều kiện và trả về giá trị tương ứng',
  },
  {
    name: 'IFS',
    syntax: 'IFS(đk_1, gt_1, đk_2, gt_2, ...)',
    category: 'logic',
    description: 'Kiểm tra nhiều điều kiện theo thứ tự',
  },
  {
    name: 'AND',
    syntax: 'AND(đk_1, [đk_2], ...)',
    category: 'logic',
    description: 'Trả về TRUE nếu mọi điều kiện đều đúng',
  },
  {
    name: 'OR',
    syntax: 'OR(đk_1, [đk_2], ...)',
    category: 'logic',
    description: 'Trả về TRUE nếu có ít nhất 1 điều kiện đúng',
  },
  {
    name: 'NOT',
    syntax: 'NOT(điều_kiện)',
    category: 'logic',
    description: 'Đảo ngược giá trị logic TRUE / FALSE',
  },

  // Văn bản
  {
    name: 'CONCAT',
    syntax: 'CONCAT(chuỗi_1, [chuỗi_2], ...)',
    category: 'text',
    description: 'Ghép nhiều chuỗi hoặc vùng văn bản thành một',
  },
  {
    name: 'CONCATENATE',
    syntax: 'CONCATENATE(chuỗi_1, chuỗi_2, ...)',
    category: 'text',
    description: 'Nối nhiều chuỗi văn bản lại với nhau',
  },
  {
    name: 'TEXTJOIN',
    syntax: 'TEXTJOIN(dấu_ngăn, bỏ_trống, vùng_1, ...)',
    category: 'text',
    description: 'Ghép văn bản có dấu ngăn cách tùy chọn',
  },
  {
    name: 'LEFT',
    syntax: 'LEFT(chuỗi, [số_ký_tự])',
    category: 'text',
    description: 'Lấy các ký tự bên trái của chuỗi',
  },
  {
    name: 'RIGHT',
    syntax: 'RIGHT(chuỗi, [số_ký_tự])',
    category: 'text',
    description: 'Lấy các ký tự bên phải của chuỗi',
  },
  {
    name: 'MID',
    syntax: 'MID(chuỗi, vị_trí_đầu, số_ký_tự)',
    category: 'text',
    description: 'Cắt chuỗi con từ vị trí chỉ định',
  },
  {
    name: 'LEN',
    syntax: 'LEN(chuỗi)',
    category: 'text',
    description: 'Đếm số ký tự trong chuỗi',
  },
  {
    name: 'TRIM',
    syntax: 'TRIM(chuỗi)',
    category: 'text',
    description: 'Xóa khoảng trắng thừa đầu/cuối và giữa từ',
  },
  {
    name: 'UPPER',
    syntax: 'UPPER(chuỗi)',
    category: 'text',
    description: 'Chuyển toàn bộ chuỗi sang CHỮ IN HOA',
  },
  {
    name: 'LOWER',
    syntax: 'LOWER(chuỗi)',
    category: 'text',
    description: 'Chuyển toàn bộ chuỗi sang chữ thường',
  },
  {
    name: 'PROPER',
    syntax: 'PROPER(chuỗi)',
    category: 'text',
    description: 'Viết Hoa Chữ Cái Đầu Mỗi Từ',
  },

  // Ngày & Giờ
  {
    name: 'TODAY',
    syntax: 'TODAY()',
    category: 'date',
    description: 'Trả về ngày hiện tại',
  },
  {
    name: 'NOW',
    syntax: 'NOW()',
    category: 'date',
    description: 'Trả về ngày và giờ hiện tại',
  },
  {
    name: 'DAY',
    syntax: 'DAY(ngày)',
    category: 'date',
    description: 'Lấy số ngày trong tháng (1–31)',
  },
  {
    name: 'MONTH',
    syntax: 'MONTH(ngày)',
    category: 'date',
    description: 'Lấy số tháng trong năm (1–12)',
  },
  {
    name: 'YEAR',
    syntax: 'YEAR(ngày)',
    category: 'date',
    description: 'Lấy số năm từ giá trị ngày',
  },

  // Tra cứu
  {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(giá_trị, bảng, cột_trả_về, [khớp_gần])',
    category: 'lookup',
    description: 'Tra cứu theo cột đầu tiên của bảng dữ liệu',
  },
  {
    name: 'HLOOKUP',
    syntax: 'HLOOKUP(giá_trị, bảng, dòng_trả_về, [khớp_gần])',
    category: 'lookup',
    description: 'Tra cứu theo dòng đầu tiên của bảng dữ liệu',
  },
  {
    name: 'INDEX',
    syntax: 'INDEX(vùng, số_dòng, [số_cột])',
    category: 'lookup',
    description: 'Trả về giá trị tại vị trí dòng, cột trong vùng',
  },
  {
    name: 'MATCH',
    syntax: 'MATCH(giá_trị, vùng, [kiểu_khớp])',
    category: 'lookup',
    description: 'Tìm vị trí tương đối của giá trị trong vùng',
  },
];

export const EXCEL_CHART_PALETTE = [
  '#107C41',
  '#2563EB',
  '#D97706',
  '#DC2626',
  '#7C3AED',
  '#0891B2',
  '#DB2777',
  '#4F46E5',
] as const;

export const CONDITIONAL_FORMAT_PRESETS: ReadonlyArray<{
  label: string;
  bgColor: string;
  textColor: string;
}> = [
  { label: 'Đỏ nhạt — Chữ đỏ đậm', bgColor: '#FEE2E2', textColor: '#991B1B' },
  { label: 'Vàng nhạt — Chữ nâu đậm', bgColor: '#FEF3C7', textColor: '#92400E' },
  { label: 'Xanh lá — Chữ xanh đậm', bgColor: '#DCFCE7', textColor: '#166534' },
  { label: 'Xanh dương — Chữ xanh đậm', bgColor: '#DBEAFE', textColor: '#1E40AF' },
  { label: 'Tím nhạt — Chữ tím đậm', bgColor: '#F3E8FF', textColor: '#6B21A8' },
];

/**
 * Shift relative cell references (like A1, B2:D4) inside an Excel formula
 * while preserving absolute references ($A$1, $A1, A$1) and string literals.
 */
export function shiftFormulaReferences(
  formula: string,
  deltaRow: number,
  deltaCol: number
): string {
  if (!formula.trim().startsWith('=')) {
    return formula;
  }
  let result = '';
  let i = 0;
  const len = formula.length;

  while (i < len) {
    const ch = formula[i];
    // Preserve quoted string literals untouched
    if (ch === '"') {
      let end = i + 1;
      while (end < len) {
        if (formula[end] === '"') {
          if (end + 1 < len && formula[end + 1] === '"') {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += formula.slice(i, end);
      i = end;
      continue;
    }

    // Match potential cell reference starting with optional $ + letters + optional $ + digits
    const remaining = formula.slice(i);
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)\b/.exec(remaining);
    if (match) {
      const prevChar = i > 0 ? formula[i - 1] : '';
      const nextChar =
        i + match[0].length < len ? formula[i + match[0].length] : '';
      // Avoid matching function names or mid-identifiers
      const isValidBoundary =
        !/[A-Za-z0-9_]/.test(prevChar) && nextChar !== '(';
      if (isValidBoundary) {
        const colAbs = match[1] === '$';
        const colIdx = letterToColumnIndex(match[2]);
        const rowAbs = match[3] === '$';
        const rowNum = Number.parseInt(match[4], 10);
        if (colIdx >= 0 && Number.isFinite(rowNum) && rowNum >= 1) {
          const nextCol = colAbs ? colIdx : Math.max(0, colIdx + deltaCol);
          const nextRow = rowAbs ? rowNum : Math.max(1, rowNum + deltaRow);
          result += `${colAbs ? '$' : ''}${columnIndexToLetter(nextCol)}${
            rowAbs ? '$' : ''
          }${nextRow}`;
          i += match[0].length;
          continue;
        }
      }
    }

    result += ch;
    i += 1;
  }

  return result;
}

function parseStrictNumeric(val: string | number | undefined): number | null {
  if (typeof val === 'number') {
    return Number.isFinite(val) ? val : null;
  }
  if (typeof val !== 'string') {
    return null;
  }
  const trimmed = val.trim();
  if (!trimmed) {
    return null;
  }
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) {
    return null;
  }
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : null;
}

/**
 * Compute auto-filled cell value for Fill Handle drag.
 * - Formulas: shifts relative references from the corresponding cyclic source cell.
 * - Numeric series (2+ numbers): continues arithmetic progression.
 * - Single number: increments by step (+1 / -1 per step).
 * - Text with trailing number (e.g. "Quý 1", "SP-01"): increments number cleanly.
 * - Otherwise repeats source value cyclically.
 */
export function computeAutoFillValue(
  sourceValues: (string | number)[],
  targetOffsetFromEnd: number,
  deltaRowFromSourceCell: number,
  deltaColFromSourceCell: number,
  cyclicSourceIndex: number
): string | number {
  const baseVal = sourceValues[cyclicSourceIndex] ?? '';
  if (typeof baseVal === 'string' && baseVal.trim().startsWith('=')) {
    return shiftFormulaReferences(
      baseVal,
      deltaRowFromSourceCell,
      deltaColFromSourceCell
    );
  }

  // Check if all source values are numeric
  const numericList = sourceValues.map((v) => parseStrictNumeric(v));
  const allNumeric =
    sourceValues.length > 0 && numericList.every((n) => n !== null);

  if (allNumeric) {
    const nums = numericList as number[];
    if (nums.length >= 2) {
      const step = nums[nums.length - 1] - nums[nums.length - 2];
      const base =
        targetOffsetFromEnd > 0
          ? nums[nums.length - 1] + step * targetOffsetFromEnd
          : nums[0] + step * targetOffsetFromEnd;
      return Math.round(base * 1_000_000) / 1_000_000;
    }
    const single = nums[0];
    return single + targetOffsetFromEnd;
  }

  // Trailing integer pattern in single text cell (e.g. "Tháng 1" -> "Tháng 2")
  if (sourceValues.length === 1 && typeof baseVal === 'string') {
    const m = /^(.*?)(\d+)(\s*)$/.exec(baseVal);
    if (m) {
      const numStr = m[2];
      const parsed = Number.parseInt(numStr, 10);
      if (Number.isFinite(parsed)) {
        const nextNum = Math.max(0, parsed + targetOffsetFromEnd);
        const padded =
          numStr.length > 1 && numStr.startsWith('0')
            ? String(nextNum).padStart(numStr.length, '0')
            : String(nextNum);
        return `${m[1]}${padded}${m[3]}`;
      }
    }
  }

  return baseVal;
}

/**
 * Increase or decrease decimal places on an Excel numFmt string.
 */
export function adjustNumFmtDecimals(
  currentFmt: string | undefined,
  delta: 1 | -1
): string {
  const fmt = (currentFmt ?? '').trim();

  if (!fmt || fmt === 'General' || fmt === '@') {
    return delta > 0 ? '#,##0.0' : '#,##0';
  }

  // Percentage
  if (fmt.includes('%')) {
    const match = /\.([0#]+)/.exec(fmt);
    const currentDec = match ? match[1].length : 0;
    const nextDec = Math.max(0, Math.min(4, currentDec + delta));
    return nextDec > 0 ? `0.${'0'.repeat(nextDec)}%` : '0%';
  }

  // Currency VNĐ ₫
  if (fmt.includes('₫') || /\bVND\b/i.test(fmt)) {
    const match = /\.([0#]+)/.exec(fmt);
    const currentDec = match ? match[1].length : 0;
    const nextDec = Math.max(0, Math.min(4, currentDec + delta));
    return nextDec > 0 ? `#,##0.${'0'.repeat(nextDec)} "₫"` : '#,##0 "₫"';
  }

  // Currency $
  if (fmt.includes('$')) {
    const match = /\.([0#]+)/.exec(fmt);
    const currentDec = match ? match[1].length : 0;
    const nextDec = Math.max(0, Math.min(4, currentDec + delta));
    return nextDec > 0 ? `$#,##0.${'0'.repeat(nextDec)}` : '$#,##0';
  }

  // Generic numeric format
  const match = /\.([0#]+)/.exec(fmt);
  const currentDec = match ? match[1].length : 0;
  const nextDec = Math.max(0, Math.min(4, currentDec + delta));
  return nextDec > 0 ? `#,##0.${'0'.repeat(nextDec)}` : '#,##0';
}

/**
 * Evaluate all conditional formatting rules against the evaluated sheet matrix.
 */
export function evaluateConditionalFormats(
  rules: ExcelConditionalFormat[] | undefined,
  evaluatedMatrix: (string | number)[][]
): Map<string, { bgColor: string; textColor?: string }> {
  const map = new Map<string, { bgColor: string; textColor?: string }>();
  if (!rules || rules.length === 0 || evaluatedMatrix.length === 0) {
    return map;
  }

  for (const rule of rules) {
    const minRow = Math.max(0, rule.range.minRow);
    const maxRow = Math.min(evaluatedMatrix.length - 1, rule.range.maxRow);
    if (minRow > maxRow) {
      continue;
    }

    if (rule.rule === 'duplicate') {
      const freq = new Map<string, number>();
      for (let r = minRow; r <= maxRow; r += 1) {
        const row = evaluatedMatrix[r];
        if (!row) continue;
        const maxCol = Math.min(row.length - 1, rule.range.maxCol);
        for (let c = Math.max(0, rule.range.minCol); c <= maxCol; c += 1) {
          const val = row[c];
          if (val === undefined || val === null) continue;
          const norm = String(val).trim().toLowerCase();
          if (!norm) continue;
          freq.set(norm, (freq.get(norm) ?? 0) + 1);
        }
      }
      for (let r = minRow; r <= maxRow; r += 1) {
        const row = evaluatedMatrix[r];
        if (!row) continue;
        const maxCol = Math.min(row.length - 1, rule.range.maxCol);
        for (let c = Math.max(0, rule.range.minCol); c <= maxCol; c += 1) {
          const val = row[c];
          if (val === undefined || val === null) continue;
          const norm = String(val).trim().toLowerCase();
          if (norm && (freq.get(norm) ?? 0) > 1) {
            map.set(`${r}:${c}`, {
              bgColor: rule.bgColor,
              textColor: rule.textColor,
            });
          }
        }
      }
      continue;
    }

    const rawTarget = (rule.value ?? '').trim();
    const numTarget = parseStrictNumeric(rawTarget);

    for (let r = minRow; r <= maxRow; r += 1) {
      const row = evaluatedMatrix[r];
      if (!row) continue;
      const maxCol = Math.min(row.length - 1, rule.range.maxCol);
      for (let c = Math.max(0, rule.range.minCol); c <= maxCol; c += 1) {
        const val = row[c];
        if (val === undefined || val === null || String(val).trim() === '') {
          continue;
        }
        const cellNum = parseStrictNumeric(val);
        let matched = false;

        if (rule.rule === 'gt') {
          if (cellNum !== null && numTarget !== null) {
            matched = cellNum > numTarget;
          }
        } else if (rule.rule === 'lt') {
          if (cellNum !== null && numTarget !== null) {
            matched = cellNum < numTarget;
          }
        } else if (rule.rule === 'eq') {
          if (cellNum !== null && numTarget !== null) {
            matched = cellNum === numTarget;
          } else {
            matched =
              String(val).trim().toLowerCase() === rawTarget.toLowerCase();
          }
        }

        if (matched) {
          map.set(`${r}:${c}`, {
            bgColor: rule.bgColor,
            textColor: rule.textColor,
          });
        }
      }
    }
  }

  return map;
}

const COMPACT_FMT = new Intl.NumberFormat('vi-VN', {
  notation: 'compact',
  maximumFractionDigits: 1,
});

const FULL_FMT = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 2,
});

function extractChartDataFromRange(
  evaluatedMatrix: (string | number)[][],
  range: { minRow: number; maxRow: number; minCol: number; maxCol: number }
): { categories: string[]; series: ExcelChartSeries[] } {
  const rowCount = evaluatedMatrix.length;
  const colCount = evaluatedMatrix[0]?.length ?? 0;
  if (rowCount === 0 || colCount === 0) {
    return { categories: [], series: [] };
  }

  // If user selected only 1 cell, expand to the non-empty block of the sheet (up to 20 rows x 8 cols)
  let { minRow, maxRow, minCol, maxCol } = range;
  if (minRow === maxRow && minCol === maxCol) {
    minRow = 0;
    minCol = 0;
    maxRow = Math.min(rowCount - 1, 19);
    maxCol = Math.min(colCount - 1, 7);
    // Trim completely empty trailing rows
    while (
      maxRow > minRow &&
      (evaluatedMatrix[maxRow] ?? []).every(
        (v) => v === undefined || v === null || String(v).trim() === ''
      )
    ) {
      maxRow -= 1;
    }
  }

  const sliceRows: (string | number)[][] = [];
  for (let r = minRow; r <= Math.min(maxRow, rowCount - 1); r += 1) {
    const rowSlice: (string | number)[] = [];
    for (let c = minCol; c <= Math.min(maxCol, colCount - 1); c += 1) {
      rowSlice.push(evaluatedMatrix[r]?.[c] ?? '');
    }
    sliceRows.push(rowSlice);
  }

  if (sliceRows.length === 0) {
    return { categories: [], series: [] };
  }

  const width = sliceRows[0].length;
  // Detect whether first row in selection is a text header row
  const firstRowHasNumeric = sliceRows[0].some(
    (v) => parseStrictNumeric(v) !== null
  );
  const hasHeader = sliceRows.length >= 2 && !firstRowHasNumeric;
  const headerRow = hasHeader ? sliceRows[0] : null;
  const dataRows = hasHeader ? sliceRows.slice(1) : sliceRows;

  // Detect whether first column is category labels (non-numeric)
  let firstColNonNumericCount = 0;
  for (const row of dataRows) {
    if (
      String(row[0] ?? '').trim() !== '' &&
      parseStrictNumeric(row[0]) === null
    ) {
      firstColNonNumericCount += 1;
    }
  }
  const hasCategoryCol = width >= 2 && firstColNonNumericCount > 0;
  const dataStartCol = hasCategoryCol ? 1 : 0;

  const categories = dataRows.map((row, idx) => {
    if (hasCategoryCol) {
      const label = String(row[0] ?? '').trim();
      return label || `Mục ${idx + 1}`;
    }
    return `Dòng ${minRow + (hasHeader ? 1 : 0) + idx + 1}`;
  });

  const series: ExcelChartSeries[] = [];
  for (let c = dataStartCol; c < width; c += 1) {
    const vals = dataRows.map((row) => parseStrictNumeric(row[c]) ?? 0);
    const hasAnyNonZeroOrNumeric = dataRows.some(
      (row) => parseStrictNumeric(row[c]) !== null
    );
    if (!hasAnyNonZeroOrNumeric) {
      continue;
    }
    const colHeader = headerRow?.[c];
    const name =
      colHeader !== undefined && String(colHeader).trim() !== ''
        ? String(colHeader).trim()
        : `Cột ${columnIndexToLetter(minCol + c)}`;
    series.push({
      name,
      values: vals,
      color: EXCEL_CHART_PALETTE[series.length % EXCEL_CHART_PALETTE.length],
    });
  }

  return { categories, series };
}

interface ExcelChartSvgProps {
  chart: ExcelChartItem;
  compact?: boolean;
}

export function ExcelChartSvg({ chart, compact = false }: ExcelChartSvgProps) {
  const width = compact ? 380 : 520;
  const height = compact ? 220 : 270;

  if (chart.categories.length === 0 || chart.series.length === 0) {
    return (
      <div className="h-[200px] flex items-center justify-center text-xs text-slate-400 bg-slate-50 rounded-lg border border-dashed border-slate-200">
        Vùng dữ liệu chưa có số liệu hợp lệ để vẽ biểu đồ.
      </div>
    );
  }

  if (chart.type === 'pie') {
    const primary = chart.series[0];
    const items = chart.categories
      .map((label, idx) => ({
        label,
        value: Math.max(0, primary.values[idx] ?? 0),
        color: EXCEL_CHART_PALETTE[idx % EXCEL_CHART_PALETTE.length],
      }))
      .filter((item) => item.value > 0);

    const total = items.reduce((acc, item) => acc + item.value, 0);
    if (total <= 0) {
      return (
        <div className="h-[200px] flex items-center justify-center text-xs text-slate-400 bg-slate-50 rounded-lg border border-dashed border-slate-200">
          Biểu đồ tròn cần ít nhất 1 giá trị dương.
        </div>
      );
    }

    const cx = compact ? 105 : 135;
    const cy = height / 2;
    const radius = compact ? 76 : 96;
    let cumulativeAngle = -Math.PI / 2;

    const slices = items.map((item) => {
      const fraction = item.value / total;
      const angle = fraction * Math.PI * 2;
      const startAngle = cumulativeAngle;
      const endAngle = cumulativeAngle + angle;
      cumulativeAngle = endAngle;

      const x1 = cx + radius * Math.cos(startAngle);
      const y1 = cy + radius * Math.sin(startAngle);
      const x2 = cx + radius * Math.cos(endAngle);
      const y2 = cy + radius * Math.sin(endAngle);
      const largeArcFlag = angle > Math.PI ? 1 : 0;

      const d =
        items.length === 1
          ? `M ${cx} ${cy - radius} A ${radius} ${radius} 0 1 1 ${cx - 0.01} ${cy - radius} Z`
          : `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;

      return {
        ...item,
        fraction,
        d,
      };
    });

    return (
      <div className="flex flex-col sm:flex-row items-center gap-3">
        <svg
          viewBox={`0 0 ${compact ? 210 : 270} ${height}`}
          className="w-full max-w-[240px] h-auto shrink-0"
          role="img"
          aria-label={chart.title}
        >
          {slices.map((slice, idx) => (
            <path
              key={`slice-${idx}`}
              d={slice.d}
              fill={slice.color}
              stroke="#ffffff"
              strokeWidth={2}
            >
              <title>{`${slice.label}: ${FULL_FMT.format(slice.value)} (${Math.round(slice.fraction * 100)}%)`}</title>
            </path>
          ))}
        </svg>

        <div className="flex-1 min-w-0 w-full space-y-1.5 max-h-[210px] overflow-y-auto pr-1">
          {slices.map((slice, idx) => (
            <div
              key={`pie-legend-${idx}`}
              className="flex items-center justify-between gap-2 text-xs"
            >
              <div className="flex items-center gap-1.5 min-w-0">
                <span
                  className="w-2.5 h-2.5 rounded-sm shrink-0"
                  style={{ backgroundColor: slice.color }}
                />
                <span className="truncate text-slate-700 font-medium">
                  {slice.label}
                </span>
              </div>
              <span className="tabular-nums text-slate-500 shrink-0 font-semibold">
                {Math.round(slice.fraction * 100)}% ({COMPACT_FMT.format(slice.value)})
              </span>
            </div>
          ))}
        </div>
      </div>
    );
  }

  // Column & Line charts
  const padLeft = 48;
  const padRight = 16;
  const padTop = 18;
  const padBottom = 42;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const allValues = chart.series.flatMap((s) => s.values);
  const rawMax = Math.max(0, ...allValues);
  const rawMin = Math.min(0, ...allValues);
  const maxVal = rawMax === rawMin ? rawMax + 10 : rawMax * 1.1;
  const minVal = rawMin < 0 ? rawMin * 1.1 : 0;
  const span = Math.max(1e-6, maxVal - minVal);

  const yOf = (val: number) =>
    padTop + plotH - ((val - minVal) / span) * plotH;
  const zeroY = yOf(0);
  const ticks = 4;
  const catCount = chart.categories.length;
  const groupW = plotW / Math.max(1, catCount);

  return (
    <div className="w-full space-y-2">
      <svg
        viewBox={`0 0 ${width} ${height}`}
        className="w-full h-auto"
        role="img"
        aria-label={chart.title}
      >
        {/* Horizontal Gridlines & Y labels */}
        {Array.from({ length: ticks + 1 }, (_, i) => {
          const tVal = minVal + (span * i) / ticks;
          const y = yOf(tVal);
          return (
            <g key={`grid-${i}`}>
              <line
                x1={padLeft}
                y1={y}
                x2={width - padRight}
                y2={y}
                stroke={i === 0 ? '#CBD5E1' : '#E2E8F0'}
                strokeDasharray={i === 0 ? undefined : '3 3'}
                strokeWidth={1}
              />
              <text
                x={padLeft - 6}
                y={y + 3.5}
                textAnchor="end"
                fontSize={10}
                fill="#64748B"
              >
                {COMPACT_FMT.format(tVal)}
              </text>
            </g>
          );
        })}

        {/* Zero baseline */}
        <line
          x1={padLeft}
          y1={zeroY}
          x2={width - padRight}
          y2={zeroY}
          stroke="#94A3B8"
          strokeWidth={1}
        />

        {/* Bars or Lines */}
        {chart.type === 'column' ? (
          chart.categories.map((cat, catIdx) => {
            const innerPad = Math.max(4, groupW * 0.16);
            const availW = Math.max(6, groupW - innerPad * 2);
            const barW = Math.max(
              4,
              Math.min(36, availW / Math.max(1, chart.series.length) - 2)
            );
            const totalBarsW =
              chart.series.length * barW + (chart.series.length - 1) * 2;
            const startX =
              padLeft + catIdx * groupW + (groupW - totalBarsW) / 2;

            return (
              <g key={`bar-group-${catIdx}`}>
                {chart.series.map((s, sIdx) => {
                  const val = s.values[catIdx] ?? 0;
                  const yVal = yOf(val);
                  const barTop = Math.min(yVal, zeroY);
                  const barHeight = Math.max(1.5, Math.abs(zeroY - yVal));
                  const x = startX + sIdx * (barW + 2);
                  return (
                    <rect
                      key={`bar-${catIdx}-${sIdx}`}
                      x={x}
                      y={barTop}
                      width={barW}
                      height={barHeight}
                      rx={2.5}
                      fill={s.color}
                      opacity={0.92}
                    >
                      <title>{`${cat} — ${s.name}: ${FULL_FMT.format(val)}`}</title>
                    </rect>
                  );
                })}
              </g>
            );
          })
        ) : (
          chart.series.map((s, sIdx) => {
            const points = chart.categories.map((_, catIdx) => {
              const x = padLeft + catIdx * groupW + groupW / 2;
              const y = yOf(s.values[catIdx] ?? 0);
              return { x, y, val: s.values[catIdx] ?? 0 };
            });
            const polylineStr = points.map((p) => `${p.x},${p.y}`).join(' ');
            return (
              <g key={`line-series-${sIdx}`}>
                <polyline
                  fill="none"
                  stroke={s.color}
                  strokeWidth={2.5}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  points={polylineStr}
                />
                {points.map((p, pIdx) => (
                  <circle
                    key={`pt-${sIdx}-${pIdx}`}
                    cx={p.x}
                    cy={p.y}
                    r={3.5}
                    fill="#ffffff"
                    stroke={s.color}
                    strokeWidth={2}
                  >
                    <title>{`${chart.categories[pIdx]} — ${s.name}: ${FULL_FMT.format(p.val)}`}</title>
                  </circle>
                ))}
              </g>
            );
          })
        )}

        {/* X Axis Labels */}
        {chart.categories.map((cat, catIdx) => {
          const x = padLeft + catIdx * groupW + groupW / 2;
          const shortLabel = cat.length > 10 ? `${cat.slice(0, 9)}…` : cat;
          return (
            <text
              key={`xlabel-${catIdx}`}
              x={x}
              y={height - 14}
              textAnchor="middle"
              fontSize={10}
              fill="#475569"
            >
              {shortLabel}
            </text>
          );
        })}
      </svg>

      {/* Legend */}
      <div className="flex flex-wrap items-center justify-center gap-3 pt-0.5">
        {chart.series.map((s, idx) => (
          <div
            key={`legend-${idx}`}
            className="inline-flex items-center gap-1.5 text-[11px] font-medium text-slate-600"
          >
            <span
              className="w-2.5 h-2.5 rounded-sm shrink-0"
              style={{ backgroundColor: s.color }}
            />
            <span>{s.name}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

export interface ExcelChartsPanelProps {
  charts: ExcelChartItem[];
  isEditingLocked: boolean;
  onChangeChartType: (chartId: string, nextType: ExcelChartItem['type']) => void;
  onDeleteChart: (chartId: string) => void;
}

export function ExcelChartsPanel({
  charts,
  isEditingLocked,
  onChangeChartType,
  onDeleteChart,
}: ExcelChartsPanelProps) {
  if (charts.length === 0) {
    return null;
  }

  return (
    <div className="mt-4 mx-3 mb-6">
      <div className="flex items-center gap-2 mb-2.5">
        <BarChart3 className="w-4 h-4 text-[#107C41]" />
        <h3 className="text-xs font-bold uppercase tracking-wide text-slate-600">
          Biểu đồ trong trang tính ({charts.length})
        </h3>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3.5 max-w-5xl">
        {charts.map((chart) => (
          <div
            key={chart.id}
            className="bg-white rounded-xl border border-[#D4DFD7] shadow-sm p-3.5 flex flex-col gap-2.5"
          >
            <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-2">
              <h4 className="text-[13px] font-bold text-slate-800 truncate">
                {chart.title}
              </h4>

              {!isEditingLocked && (
                <div className="flex items-center gap-1 shrink-0">
                  <div className="inline-flex items-center rounded-lg bg-slate-100 p-0.5">
                    {(
                      [
                        { type: 'column', icon: BarChart3, label: 'Cột' },
                        { type: 'line', icon: LineChart, label: 'Đường' },
                        { type: 'pie', icon: PieChart, label: 'Tròn' },
                      ] as const
                    ).map(({ type, icon: Icon, label }) => (
                      <button
                        key={type}
                        type="button"
                        onClick={() => onChangeChartType(chart.id, type)}
                        title={`Đổi sang biểu đồ ${label}`}
                        className={`p-1 rounded-md transition-colors ${
                          chart.type === type
                            ? 'bg-white text-[#107C41] shadow-sm'
                            : 'text-slate-500 hover:text-slate-800'
                        }`}
                      >
                        <Icon className="w-3.5 h-3.5" />
                      </button>
                    ))}
                  </div>

                  <button
                    type="button"
                    onClick={() => onDeleteChart(chart.id)}
                    title="Xóa biểu đồ"
                    aria-label="Xóa biểu đồ"
                    className="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}
            </div>

            <ExcelChartSvg chart={chart} compact />
          </div>
        ))}
      </div>
    </div>
  );
}

export interface InsertChartModalProps {
  initialType: ExcelChartItem['type'];
  evaluatedMatrix: (string | number)[][];
  selectionRange: {
    minRow: number;
    maxRow: number;
    minCol: number;
    maxCol: number;
  };
  rangeLabel: string;
  onClose: () => void;
  onInsert: (chart: ExcelChartItem) => void;
}

export function InsertChartModal({
  initialType,
  evaluatedMatrix,
  selectionRange,
  rangeLabel,
  onClose,
  onInsert,
}: InsertChartModalProps) {
  const [chartType, setChartType] =
    useState<ExcelChartItem['type']>(initialType);
  const [title, setTitle] = useState<string>(
    `Biểu đồ dữ liệu (${rangeLabel})`
  );

  const extracted = useMemo(
    () => extractChartDataFromRange(evaluatedMatrix, selectionRange),
    [evaluatedMatrix, selectionRange]
  );

  const previewChart: ExcelChartItem = useMemo(
    () => ({
      id: 'preview',
      type: chartType,
      title: title.trim() || 'Biểu đồ không tiêu đề',
      categories: extracted.categories,
      series: extracted.series,
    }),
    [chartType, extracted.categories, extracted.series, title]
  );

  const canInsert =
    extracted.categories.length > 0 && extracted.series.length > 0;

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-[1px] flex items-center justify-center p-3 sm:p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) {
          onClose();
        }
      }}
    >
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-xl max-h-[90dvh] flex flex-col overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-[#F3F8F5] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-[#107C41] text-white flex items-center justify-center">
              <BarChart3 className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Chèn biểu đồ từ vùng dữ liệu
              </h3>
              <p className="text-[11px] text-slate-500">
                Nguồn vùng chọn: <strong className="text-[#107C41]">{rangeLabel}</strong>
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-200/70"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 overflow-y-auto space-y-4">
          <div className="grid grid-cols-3 gap-2">
            {(
              [
                { type: 'column', label: 'Biểu đồ Cột', icon: BarChart3 },
                { type: 'line', label: 'Biểu đồ Đường', icon: LineChart },
                { type: 'pie', label: 'Biểu đồ Tròn', icon: PieChart },
              ] as const
            ).map(({ type, label, icon: Icon }) => (
              <button
                key={type}
                type="button"
                onClick={() => setChartType(type)}
                className={`flex flex-col items-center gap-1.5 p-2.5 rounded-xl border text-xs font-semibold transition-all ${
                  chartType === type
                    ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41] shadow-sm'
                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                }`}
              >
                <Icon className="w-5 h-5" />
                <span>{label}</span>
              </button>
            ))}
          </div>

          <label className="block space-y-1">
            <span className="text-xs font-semibold text-slate-700">
              Tiêu đề biểu đồ
            </span>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="Nhập tiêu đề biểu đồ..."
              className="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs sm:text-sm focus:outline-none focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15"
            />
          </label>

          <div className="rounded-xl border border-slate-200 bg-slate-50/70 p-3">
            <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 mb-2">
              Xem trước trực quan
            </p>
            <ExcelChartSvg chart={previewChart} />
          </div>
        </div>

        <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-2 rounded-lg border border-slate-300 bg-white text-xs font-semibold text-slate-700 hover:bg-slate-100"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={!canInsert}
            onClick={() => {
              onInsert({
                ...previewChart,
                id: `chart-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
              });
              onClose();
            }}
            className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg bg-[#107C41] hover:bg-[#0D6535] text-white text-xs font-semibold disabled:opacity-40 transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Chèn biểu đồ</span>
          </button>
        </div>
      </div>
    </div>
  );
}

export interface ConditionalFormatModalProps {
  rules: ExcelConditionalFormat[];
  selectionRange: {
    minRow: number;
    maxRow: number;
    minCol: number;
    maxCol: number;
  };
  rangeLabel: string;
  onClose: () => void;
  onAddRule: (rule: ExcelConditionalFormat) => void;
  onDeleteRule: (ruleId: string) => void;
}

export function ConditionalFormatModal({
  rules,
  selectionRange,
  rangeLabel,
  onClose,
  onAddRule,
  onDeleteRule,
}: ConditionalFormatModalProps) {
  const [ruleType, setRuleType] =
    useState<ExcelConditionalFormat['rule']>('gt');
  const [thresholdValue, setThresholdValue] = useState('0');
  const [presetIdx, setPresetIdx] = useState(0);

  const selectedPreset =
    CONDITIONAL_FORMAT_PRESETS[presetIdx] ?? CONDITIONAL_FORMAT_PRESETS[0];

  const handleAdd = () => {
    if (ruleType !== 'duplicate' && thresholdValue.trim() === '') {
      return;
    }
    onAddRule({
      id: `cf-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      range: { ...selectionRange },
      rule: ruleType,
      value: ruleType === 'duplicate' ? undefined : thresholdValue.trim(),
      bgColor: selectedPreset.bgColor,
      textColor: selectedPreset.textColor,
    });
  };

  const describeRule = (r: ExcelConditionalFormat): string => {
    const startRef = `${columnIndexToLetter(r.range.minCol)}${r.range.minRow + 1}`;
    const endRef = `${columnIndexToLetter(r.range.maxCol)}${r.range.maxRow + 1}`;
    const refText = startRef === endRef ? startRef : `${startRef}:${endRef}`;
    switch (r.rule) {
      case 'gt':
        return `${refText}: Lớn hơn ${r.value ?? ''}`;
      case 'lt':
        return `${refText}: Nhỏ hơn ${r.value ?? ''}`;
      case 'eq':
        return `${refText}: Bằng "${r.value ?? ''}"`;
      case 'duplicate':
        return `${refText}: Giá trị trùng lặp`;
    }
  };

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-[1px] flex items-center justify-center p-3 sm:p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) {
          onClose();
        }
      }}
    >
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md max-h-[90dvh] flex flex-col overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-[#F3F8F5] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-[#107C41] text-white flex items-center justify-center">
              <Sparkles className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Định dạng có điều kiện
              </h3>
              <p className="text-[11px] text-slate-500">
                Áp dụng cho vùng: <strong className="text-[#107C41]">{rangeLabel}</strong>
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-200/70"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 overflow-y-auto space-y-4">
          <div className="space-y-3 rounded-xl border border-slate-200 bg-slate-50/70 p-3">
            <label className="block space-y-1">
              <span className="text-xs font-semibold text-slate-700">
                Quy tắc tô màu ô
              </span>
              <select
                value={ruleType}
                onChange={(e) =>
                  setRuleType(e.target.value as ExcelConditionalFormat['rule'])
                }
                className="w-full px-3 py-2 rounded-lg border border-slate-300 bg-white text-xs sm:text-sm focus:outline-none focus:border-[#107C41]"
              >
                <option value="gt">Lớn hơn (&gt;)</option>
                <option value="lt">Nhỏ hơn (&lt;)</option>
                <option value="eq">Bằng (=)</option>
                <option value="duplicate">Giá trị trùng lặp (Duplicate)</option>
              </select>
            </label>

            {ruleType !== 'duplicate' && (
              <label className="block space-y-1">
                <span className="text-xs font-semibold text-slate-700">
                  Giá trị so sánh
                </span>
                <input
                  type="text"
                  value={thresholdValue}
                  onChange={(e) => setThresholdValue(e.target.value)}
                  placeholder="Ví dụ: 100 hoặc Đạt"
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 bg-white text-xs sm:text-sm focus:outline-none focus:border-[#107C41]"
                />
              </label>
            )}

            <div className="space-y-1.5">
              <span className="text-xs font-semibold text-slate-700">
                Kiểu màu nổi bật
              </span>
              <div className="grid grid-cols-1 gap-1.5">
                {CONDITIONAL_FORMAT_PRESETS.map((preset, idx) => (
                  <button
                    key={preset.label}
                    type="button"
                    onClick={() => setPresetIdx(idx)}
                    style={{
                      backgroundColor: preset.bgColor,
                      color: preset.textColor,
                    }}
                    className={`w-full px-3 py-1.5 rounded-lg text-left text-xs font-semibold border transition-all ${
                      presetIdx === idx
                        ? 'border-[#107C41] ring-2 ring-[#107C41]/25'
                        : 'border-black/10 opacity-85 hover:opacity-100'
                    }`}
                  >
                    {preset.label}
                  </button>
                ))}
              </div>
            </div>

            <button
              type="button"
              onClick={handleAdd}
              className="w-full inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-[#107C41] hover:bg-[#0D6535] text-white text-xs font-semibold transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Thêm quy tắc cho vùng {rangeLabel}</span>
            </button>
          </div>

          <div className="space-y-2">
            <p className="text-xs font-bold text-slate-600">
              Quy tắc đang áp dụng ({rules.length})
            </p>
            {rules.length === 0 ? (
              <p className="text-xs text-slate-400 italic">
                Chưa có quy tắc định dạng có điều kiện nào trên trang tính này.
              </p>
            ) : (
              <div className="space-y-1.5 max-h-40 overflow-y-auto">
                {rules.map((r) => (
                  <div
                    key={r.id}
                    className="flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white text-xs"
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <span
                        className="w-4 h-4 rounded border border-black/15 shrink-0"
                        style={{ backgroundColor: r.bgColor }}
                      />
                      <span className="truncate font-medium text-slate-700">
                        {describeRule(r)}
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => onDeleteRule(r.id)}
                      aria-label="Xóa quy tắc"
                      className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 shrink-0"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        <div className="px-4 py-2.5 border-t border-slate-200 bg-slate-50 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold"
          >
            Xong
          </button>
        </div>
      </div>
    </div>
  );
}

export interface DataValidationModalProps {
  rangeLabel: string;
  initialOptions: string[];
  onClose: () => void;
  onApply: (options: string[]) => void;
  onClear: () => void;
}

const VALIDATION_PRESETS: ReadonlyArray<{ label: string; values: string[] }> = [
  {
    label: 'Trạng thái công việc',
    values: ['Chưa bắt đầu', 'Đang thực hiện', 'Hoàn thành', 'Tạm hoãn'],
  },
  { label: 'Kết quả Đạt / Không đạt', values: ['Đạt', 'Không đạt', 'Chờ duyệt'] },
  { label: 'Mức độ ưu tiên', values: ['Cao', 'Trung bình', 'Thấp'] },
];

export function DataValidationModal({
  rangeLabel,
  initialOptions,
  onClose,
  onApply,
  onClear,
}: DataValidationModalProps) {
  const [rawText, setRawText] = useState(initialOptions.join(', '));

  const parsedOptions = useMemo(
    () =>
      Array.from(
        new Set(
          rawText
            .split(/[,;\n]+/)
            .map((s) => s.trim())
            .filter(Boolean)
        )
      ),
    [rawText]
  );

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-[1px] flex items-center justify-center p-3 sm:p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) {
          onClose();
        }
      }}
    >
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-[#F3F8F5] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-[#107C41] text-white flex items-center justify-center">
              <ListChecks className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Xác thực dữ liệu (Danh sách chọn)
              </h3>
              <p className="text-[11px] text-slate-500">
                Tạo dropdown chọn nhanh cho vùng{' '}
                <strong className="text-[#107C41]">{rangeLabel}</strong>
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-200/70"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 space-y-3.5">
          <label className="block space-y-1">
            <span className="text-xs font-semibold text-slate-700">
              Danh sách giá trị (cách nhau bằng dấu phẩy hoặc xuống dòng)
            </span>
            <textarea
              rows={3}
              value={rawText}
              onChange={(e) => setRawText(e.target.value)}
              placeholder="Ví dụ: Hoàn thành, Đang làm, Chờ duyệt"
              className="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs sm:text-sm focus:outline-none focus:border-[#107C41] focus:ring-2 focus:ring-[#107C41]/15"
            />
          </label>

          <div className="space-y-1.5">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">
              Mẫu nhanh
            </span>
            <div className="flex flex-wrap gap-1.5">
              {VALIDATION_PRESETS.map((preset) => (
                <button
                  key={preset.label}
                  type="button"
                  onClick={() => setRawText(preset.values.join(', '))}
                  className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-[#E6F4EC] text-slate-700 hover:text-[#107C41] text-[11px] font-semibold transition-colors"
                >
                  {preset.label}
                </button>
              ))}
            </div>
          </div>

          {parsedOptions.length > 0 && (
            <div className="flex flex-wrap gap-1 pt-1">
              {parsedOptions.map((opt) => (
                <span
                  key={opt}
                  className="inline-flex items-center px-2 py-0.5 rounded-md bg-[#E6F4EC] border border-[#107C41]/25 text-[11px] font-semibold text-[#107C41]"
                >
                  {opt}
                </span>
              ))}
            </div>
          )}
        </div>

        <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between gap-2">
          <button
            type="button"
            onClick={() => {
              onClear();
              onClose();
            }}
            className="px-3 py-1.5 rounded-lg border border-rose-200 bg-rose-50 hover:bg-rose-100 text-xs font-semibold text-rose-700"
          >
            Xóa xác thực
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-xs font-semibold text-slate-700"
            >
              Hủy
            </button>
            <button
              type="button"
              disabled={parsedOptions.length === 0}
              onClick={() => {
                onApply(parsedOptions);
                onClose();
              }}
              className="px-4 py-1.5 rounded-lg bg-[#107C41] hover:bg-[#0D6535] text-white text-xs font-semibold disabled:opacity-40"
            >
              Áp dụng ({parsedOptions.length})
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export interface TextToColumnsModalProps {
  columnLetter: string;
  sampleValues: string[];
  onClose: () => void;
  onSplit: (delimiter: string) => void;
}

export function TextToColumnsModal({
  columnLetter,
  sampleValues,
  onClose,
  onSplit,
}: TextToColumnsModalProps) {
  const [delimiterMode, setDelimiterMode] = useState<
    'comma' | 'semicolon' | 'space' | 'tab' | 'pipe' | 'custom'
  >('comma');
  const [customDelim, setCustomDelim] = useState('-');

  const activeDelimiter =
    delimiterMode === 'comma'
      ? ','
      : delimiterMode === 'semicolon'
        ? ';'
        : delimiterMode === 'space'
          ? ' '
          : delimiterMode === 'tab'
            ? '\t'
            : delimiterMode === 'pipe'
              ? '|'
              : customDelim;

  const previewRows = useMemo(() => {
    const samples = sampleValues.filter((s) => s.trim() !== '').slice(0, 5);
    if (!activeDelimiter) {
      return samples.map((s) => [s]);
    }
    return samples.map((s) => s.split(activeDelimiter).map((p) => p.trim()));
  }, [activeDelimiter, sampleValues]);

  const maxCols = Math.max(
    1,
    previewRows.reduce((m, r) => Math.max(m, r.length), 1)
  );

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/45 backdrop-blur-[1px] flex items-center justify-center p-3 sm:p-4"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) {
          onClose();
        }
      }}
    >
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-200 bg-[#F3F8F5] flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <div className="w-7 h-7 rounded-lg bg-[#107C41] text-white flex items-center justify-center">
              <Columns3 className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Tách văn bản thành cột (Text to Columns)
              </h3>
              <p className="text-[11px] text-slate-500">
                Tách dữ liệu tại cột <strong className="text-[#107C41]">{columnLetter}</strong> sang các cột liền kề
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-200/70"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 space-y-4">
          <div className="space-y-1.5">
            <span className="text-xs font-semibold text-slate-700">
              Chọn ký tự phân tách
            </span>
            <div className="grid grid-cols-3 gap-1.5">
              {(
                [
                  { id: 'comma', label: 'Dấu phẩy (,)' },
                  { id: 'semicolon', label: 'Chấm phẩy (;)' },
                  { id: 'space', label: 'Khoảng trắng' },
                  { id: 'tab', label: 'Tab' },
                  { id: 'pipe', label: 'Gạch dọc (|)' },
                  { id: 'custom', label: 'Ký tự khác' },
                ] as const
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setDelimiterMode(item.id)}
                  className={`px-2.5 py-2 rounded-lg border text-xs font-semibold transition-colors ${
                    delimiterMode === item.id
                      ? 'bg-[#E6F4EC] border-[#107C41] text-[#107C41]'
                      : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                  }`}
                >
                  {item.label}
                </button>
              ))}
            </div>

            {delimiterMode === 'custom' && (
              <input
                type="text"
                value={customDelim}
                onChange={(e) => setCustomDelim(e.target.value)}
                placeholder="Nhập ký tự phân tách (vd: - hoặc /)"
                className="mt-2 w-full px-3 py-1.5 rounded-lg border border-slate-300 text-xs focus:outline-none focus:border-[#107C41]"
              />
            )}
          </div>

          <div className="space-y-1.5">
            <span className="text-xs font-semibold text-slate-700">
              Xem trước kết quả tách ({maxCols} cột)
            </span>
            <div className="rounded-xl border border-slate-200 overflow-x-auto bg-slate-50">
              {previewRows.length === 0 ? (
                <p className="p-3 text-xs text-slate-400">
                  Cột {columnLetter} không có văn bản để tách.
                </p>
              ) : (
                <table className="w-full text-xs border-collapse">
                  <tbody>
                    {previewRows.map((cells, rIdx) => (
                      <tr key={`prev-${rIdx}`} className="border-b border-slate-200 last:border-b-0">
                        {Array.from({ length: maxCols }, (_, cIdx) => (
                          <td
                            key={`prev-${rIdx}-${cIdx}`}
                            className="px-2.5 py-1.5 border-r border-slate-200 last:border-r-0 bg-white text-slate-700 whitespace-nowrap"
                          >
                            {cells[cIdx] ?? ''}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>

        <div className="px-4 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 rounded-lg border border-slate-300 bg-white text-xs font-semibold text-slate-700"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={!activeDelimiter || previewRows.length === 0}
            onClick={() => {
              onSplit(activeDelimiter);
              onClose();
            }}
            className="px-4 py-1.5 rounded-lg bg-[#107C41] hover:bg-[#0D6535] text-white text-xs font-semibold disabled:opacity-40"
          >
            Tách thành cột
          </button>
        </div>
      </div>
    </div>
  );
}

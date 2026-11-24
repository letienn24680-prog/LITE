export function columnIndexToLetter(index: number): string {
  if (!Number.isFinite(index) || index < 0) {
    return 'A';
  }
  let n = Math.floor(index);
  let result = '';
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

export function letterToColumnIndex(letters: string): number {
  const clean = letters.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(clean)) {
    return -1;
  }
  let index = 0;
  for (let i = 0; i < clean.length; i++) {
    index = index * 26 + (clean.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function formatCellRef(rowIdx: number, colIdx: number): string {
  const safeRow = Math.max(0, Math.floor(rowIdx));
  return `${columnIndexToLetter(colIdx)}${safeRow + 1}`;
}

type FormulaErrorCode = '#CYCLE!' | '#DIV/0!' | '#ERROR!' | '#N/A';

class FormulaError extends Error {
  readonly code: FormulaErrorCode;

  constructor(code: FormulaErrorCode) {
    super(code);
    this.code = code;
  }
}

const FORMULA_ERRORS = new Set<string>(['#CYCLE!', '#DIV/0!', '#ERROR!', '#N/A']);

type EvalValue = string | number | boolean;

type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'IDENT'
  | 'CELL'
  | 'PLUS'
  | 'MINUS'
  | 'STAR'
  | 'SLASH'
  | 'CARET'
  | 'AMP'
  | 'EQ'
  | 'NEQ'
  | 'GTE'
  | 'LTE'
  | 'GT'
  | 'LT'
  | 'LPAREN'
  | 'RPAREN'
  | 'COLON'
  | 'COMMA'
  | 'EOF';

interface Token {
  type: TokenType;
  value: string;
}

const CELL_REF_REGEX = /^\$?([A-Za-z]+)\$?(\d+)$/;

function parseCellCoords(ref: string): { row: number; col: number } | null {
  const match = CELL_REF_REGEX.exec(ref);
  if (!match) {
    return null;
  }
  const col = letterToColumnIndex(match[1]);
  const rowNum = Number.parseInt(match[2], 10);
  if (col < 0 || !Number.isFinite(rowNum) || rowNum < 1) {
    return null;
  }
  return { row: rowNum - 1, col };
}

function tokenize(expression: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const len = expression.length;

  while (i < len) {
    const ch = expression[i];

    if (/\s/.test(ch)) {
      i++;
      continue;
    }

    // String literal "..." (supports Excel escaped double quotes "")
    if (ch === '"') {
      i++;
      let str = '';
      let closed = false;
      while (i < len) {
        if (expression[i] === '"') {
          if (i + 1 < len && expression[i + 1] === '"') {
            str += '"';
            i += 2;
            continue;
          }
          i++;
          closed = true;
          break;
        }
        str += expression[i];
        i++;
      }
      if (!closed) {
        throw new FormulaError('#ERROR!');
      }
      tokens.push({ type: 'STRING', value: str });
      continue;
    }

    // Two-character comparison operators first
    const twoChar = expression.slice(i, i + 2);
    if (twoChar === '<>' || twoChar === '!=') {
      tokens.push({ type: 'NEQ', value: twoChar });
      i += 2;
      continue;
    }
    if (twoChar === '>=') {
      tokens.push({ type: 'GTE', value: twoChar });
      i += 2;
      continue;
    }
    if (twoChar === '<=') {
      tokens.push({ type: 'LTE', value: twoChar });
      i += 2;
      continue;
    }

    if (ch === '=') {
      tokens.push({ type: 'EQ', value: ch });
      i++;
      continue;
    }
    if (ch === '>') {
      tokens.push({ type: 'GT', value: ch });
      i++;
      continue;
    }
    if (ch === '<') {
      tokens.push({ type: 'LT', value: ch });
      i++;
      continue;
    }
    if (ch === '&') {
      tokens.push({ type: 'AMP', value: ch });
      i++;
      continue;
    }
    if (ch === '+') {
      tokens.push({ type: 'PLUS', value: ch });
      i++;
      continue;
    }
    if (ch === '-') {
      tokens.push({ type: 'MINUS', value: ch });
      i++;
      continue;
    }
    if (ch === '*') {
      tokens.push({ type: 'STAR', value: ch });
      i++;
      continue;
    }
    if (ch === '/') {
      tokens.push({ type: 'SLASH', value: ch });
      i++;
      continue;
    }
    if (ch === '^') {
      tokens.push({ type: 'CARET', value: ch });
      i++;
      continue;
    }
    if (ch === '(') {
      tokens.push({ type: 'LPAREN', value: ch });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'RPAREN', value: ch });
      i++;
      continue;
    }
    if (ch === ':') {
      tokens.push({ type: 'COLON', value: ch });
      i++;
      continue;
    }
    if (ch === ',' || ch === ';') {
      tokens.push({ type: 'COMMA', value: ch });
      i++;
      continue;
    }

    // Number literal (integer, decimal, optional exponent)
    if (/[0-9]/.test(ch) || (ch === '.' && i + 1 < len && /[0-9]/.test(expression[i + 1]))) {
      const start = i;
      let hasDot = false;
      while (i < len) {
        const c = expression[i];
        if (/[0-9]/.test(c)) {
          i++;
        } else if (c === '.' && !hasDot) {
          hasDot = true;
          i++;
        } else {
          break;
        }
      }
      if (i < len && (expression[i] === 'e' || expression[i] === 'E')) {
        let expIdx = i + 1;
        if (expIdx < len && (expression[expIdx] === '+' || expression[expIdx] === '-')) {
          expIdx++;
        }
        if (expIdx < len && /[0-9]/.test(expression[expIdx])) {
          i = expIdx;
          while (i < len && /[0-9]/.test(expression[i])) {
            i++;
          }
        }
      }
      tokens.push({ type: 'NUMBER', value: expression.slice(start, i) });
      continue;
    }

    // Identifier, boolean literal, or cell reference (supports optional $ like $A$1)
    if (/[A-Za-z_$]/.test(ch)) {
      const start = i;
      while (i < len && /[A-Za-z0-9_$.]/.test(expression[i])) {
        i++;
      }
      const raw = expression.slice(start, i);
      const upper = raw.toUpperCase();
      // Peek ahead skipping spaces to check if this is a function call
      let lookahead = i;
      while (lookahead < len && /\s/.test(expression[lookahead])) {
        lookahead++;
      }
      if (lookahead < len && expression[lookahead] === '(') {
        tokens.push({ type: 'IDENT', value: upper });
      } else if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({ type: 'BOOLEAN', value: upper });
      } else if (CELL_REF_REGEX.test(raw)) {
        tokens.push({ type: 'CELL', value: upper });
      } else {
        throw new FormulaError('#ERROR!');
      }
      continue;
    }

    throw new FormulaError('#ERROR!');
  }

  tokens.push({ type: 'EOF', value: '' });
  return tokens;
}

function parseNumericLiteral(value: EvalValue): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === 'boolean') {
    return value ? 1 : 0;
  }
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }
  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    const parsed = Number(trimmed);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function normalizeNumber(value: number): number {
  if (!Number.isFinite(value)) {
    throw new FormulaError('#ERROR!');
  }
  if (Object.is(value, -0) || value === 0) {
    return 0;
  }
  if (Number.isInteger(value)) {
    return value;
  }
  return Number(value.toPrecision(12));
}

function coerceToNumber(val: EvalValue): number {
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) {
      throw new FormulaError('#ERROR!');
    }
    return val;
  }
  if (typeof val === 'boolean') {
    return val ? 1 : 0;
  }
  const trimmed = val.trim();
  if (!trimmed) {
    return 0;
  }
  if (FORMULA_ERRORS.has(trimmed)) {
    throw new FormulaError(trimmed as FormulaErrorCode);
  }
  const num = parseNumericLiteral(trimmed);
  if (num === null) {
    throw new FormulaError('#ERROR!');
  }
  return num;
}

function coerceToBoolean(val: EvalValue): boolean {
  if (typeof val === 'boolean') {
    return val;
  }
  if (typeof val === 'number') {
    return val !== 0;
  }
  const trimmed = val.trim();
  if (FORMULA_ERRORS.has(trimmed)) {
    throw new FormulaError(trimmed as FormulaErrorCode);
  }
  if (!trimmed) {
    return false;
  }
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') {
    return true;
  }
  if (upper === 'FALSE') {
    return false;
  }
  const num = parseNumericLiteral(trimmed);
  if (num !== null) {
    return num !== 0;
  }
  throw new FormulaError('#ERROR!');
}

function coerceToString(val: EvalValue): string {
  if (typeof val === 'string') {
    if (FORMULA_ERRORS.has(val.trim())) {
      throw new FormulaError(val.trim() as FormulaErrorCode);
    }
    return val;
  }
  if (typeof val === 'boolean') {
    return val ? 'TRUE' : 'FALSE';
  }
  return String(normalizeNumber(val));
}

function finalizeOutputValue(val: EvalValue): string | number {
  if (typeof val === 'number') {
    return normalizeNumber(val);
  }
  if (typeof val === 'boolean') {
    return val ? 'TRUE' : 'FALSE';
  }
  if (FORMULA_ERRORS.has(val.trim())) {
    throw new FormulaError(val.trim() as FormulaErrorCode);
  }
  return val;
}

function compareValues(left: EvalValue, right: EvalValue, op: TokenType): boolean {
  if (typeof left === 'string' && FORMULA_ERRORS.has(left.trim())) {
    throw new FormulaError(left.trim() as FormulaErrorCode);
  }
  if (typeof right === 'string' && FORMULA_ERRORS.has(right.trim())) {
    throw new FormulaError(right.trim() as FormulaErrorCode);
  }

  const lNum =
    typeof left === 'number'
      ? left
      : typeof left === 'string' && left.trim() !== ''
      ? parseNumericLiteral(left)
      : null;
  const rNum =
    typeof right === 'number'
      ? right
      : typeof right === 'string' && right.trim() !== ''
      ? parseNumericLiteral(right)
      : null;

  const effectiveLNum =
    lNum !== null ? lNum : left === '' && rNum !== null ? 0 : null;
  const effectiveRNum =
    rNum !== null ? rNum : right === '' && lNum !== null ? 0 : null;

  let cmp: number;
  if (typeof left === 'boolean' && typeof right === 'boolean') {
    cmp = Number(left) - Number(right);
  } else if (effectiveLNum !== null && effectiveRNum !== null) {
    cmp = effectiveLNum - effectiveRNum;
  } else {
    const lStr = coerceToString(left).toLowerCase();
    const rStr = coerceToString(right).toLowerCase();
    cmp = lStr < rStr ? -1 : lStr > rStr ? 1 : 0;
  }

  switch (op) {
    case 'EQ':
      return cmp === 0;
    case 'NEQ':
      return cmp !== 0;
    case 'GT':
      return cmp > 0;
    case 'GTE':
      return cmp >= 0;
    case 'LT':
      return cmp < 0;
    case 'LTE':
      return cmp <= 0;
    default:
      throw new FormulaError('#ERROR!');
  }
}

function excelWildcardToRegex(pattern: string): RegExp {
  let regexStr = '^';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) {
      const next = pattern[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        regexStr += `\\${next}`;
        i++;
        continue;
      }
    }
    if (ch === '*') {
      regexStr += '.*';
    } else if (ch === '?') {
      regexStr += '.';
    } else {
      regexStr += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
  }
  regexStr += '$';
  return new RegExp(regexStr, 'i');
}

function hasExcelWildcard(pattern: string): boolean {
  for (let i = 0; i < pattern.length; i++) {
    if (pattern[i] === '~' && i + 1 < pattern.length) {
      i++;
      continue;
    }
    if (pattern[i] === '*' || pattern[i] === '?') {
      return true;
    }
  }
  return false;
}

function matchesExcelCriteria(cellVal: string | number, criteria: EvalValue): boolean {
  if (typeof cellVal === 'string' && FORMULA_ERRORS.has(cellVal.trim())) {
    return false;
  }

  if (typeof criteria === 'number') {
    const cNum = parseNumericLiteral(cellVal);
    return cNum !== null && cNum === criteria;
  }

  if (typeof criteria === 'boolean') {
    const target = criteria ? 'TRUE' : 'FALSE';
    return String(cellVal).trim().toUpperCase() === target;
  }

  const rawCrit = criteria;
  if (rawCrit === '') {
    return cellVal === '';
  }

  const opMatch = /^(<>|!=|>=|<=|>|<|=)([\s\S]*)$/.exec(rawCrit);
  if (opMatch) {
    const op = opMatch[1];
    const rhs = opMatch[2];

    if (rhs === '') {
      if (op === '=') return cellVal === '';
      if (op === '<>' || op === '!=') return cellVal !== '';
      return false;
    }

    const rhsNum = parseNumericLiteral(rhs);
    const cellNum = cellVal === '' ? null : parseNumericLiteral(cellVal);

    if (op === '>' || op === '>=' || op === '<' || op === '<=') {
      if (rhsNum !== null) {
        if (cellNum === null) {
          return false;
        }
        if (op === '>') return cellNum > rhsNum;
        if (op === '>=') return cellNum >= rhsNum;
        if (op === '<') return cellNum < rhsNum;
        return cellNum <= rhsNum;
      }
      if (cellVal === '') {
        return false;
      }
      const lStr = String(cellVal).toLowerCase();
      const rStr = rhs.toLowerCase();
      if (op === '>') return lStr > rStr;
      if (op === '>=') return lStr >= rStr;
      if (op === '<') return lStr < rStr;
      return lStr <= rStr;
    }

    // '=' or '<>' / '!='
    let isEq: boolean;
    if (rhsNum !== null && cellNum !== null) {
      isEq = cellNum === rhsNum;
    } else if (hasExcelWildcard(rhs)) {
      isEq = excelWildcardToRegex(rhs).test(String(cellVal));
    } else {
      isEq = String(cellVal).trim().toLowerCase() === rhs.trim().toLowerCase();
    }
    return op === '=' ? isEq : !isEq;
  }

  // Bare criteria without operator prefix
  const critNum = parseNumericLiteral(rawCrit);
  const cellNum = cellVal === '' ? null : parseNumericLiteral(cellVal);
  if (critNum !== null && cellNum !== null) {
    return cellNum === critNum;
  }
  if (hasExcelWildcard(rawCrit)) {
    return excelWildcardToRegex(rawCrit).test(String(cellVal));
  }
  return String(cellVal).trim().toLowerCase() === rawCrit.trim().toLowerCase();
}

function valuesMatchForLookup(candidate: string | number, lookup: EvalValue): boolean {
  if (typeof candidate === 'string' && FORMULA_ERRORS.has(candidate.trim())) {
    return false;
  }
  const cNum = parseNumericLiteral(candidate);
  const lNum = parseNumericLiteral(lookup);
  if (cNum !== null && lNum !== null && typeof lookup !== 'boolean') {
    return cNum === lNum;
  }
  const lStr = coerceToString(lookup).trim();
  const cStr = String(candidate).trim();
  if (hasExcelWildcard(lStr)) {
    return excelWildcardToRegex(lStr).test(cStr);
  }
  return cStr.toLowerCase() === lStr.toLowerCase();
}

function compareForApproxLookup(candidate: string | number, lookup: EvalValue): number | null {
  if (candidate === '' || (typeof candidate === 'string' && FORMULA_ERRORS.has(candidate.trim()))) {
    return null;
  }
  const cNum = parseNumericLiteral(candidate);
  const lNum = parseNumericLiteral(lookup);
  if (cNum !== null && lNum !== null && typeof lookup !== 'boolean') {
    return cNum - lNum;
  }
  if (cNum === null && lNum === null) {
    const cStr = String(candidate).trim().toLowerCase();
    const lStr = coerceToString(lookup).trim().toLowerCase();
    return cStr < lStr ? -1 : cStr > lStr ? 1 : 0;
  }
  return null;
}

function parseExcelDateInput(val: EvalValue): { year: number; month: number; day: number } | null {
  if (typeof val === 'boolean') {
    return null;
  }

  // Excel serial date (days since 1899-12-30 UTC; valid 1..2958465 covers 1900..9999)
  const serial = parseNumericLiteral(val);
  if (serial !== null && serial >= 1 && serial <= 2958465) {
    const wholeDays = Math.floor(serial);
    const utcMs = Date.UTC(1899, 11, 30) + wholeDays * 86_400_000;
    const d = new Date(utcMs);
    if (!Number.isNaN(d.getTime())) {
      return {
        year: d.getUTCFullYear(),
        month: d.getUTCMonth() + 1,
        day: d.getUTCDate(),
      };
    }
  }

  const str = String(val).trim();
  if (!str) {
    return null;
  }

  // dd/mm/yyyy or dd-mm-yyyy (optional time HH:mm[:ss])
  const dmyMatch = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/.exec(str);
  if (dmyMatch) {
    const day = Number.parseInt(dmyMatch[1], 10);
    const month = Number.parseInt(dmyMatch[2], 10);
    const year = Number.parseInt(dmyMatch[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const check = new Date(Date.UTC(year, month - 1, day));
      if (
        check.getUTCFullYear() === year &&
        check.getUTCMonth() === month - 1 &&
        check.getUTCDate() === day
      ) {
        return { year, month, day };
      }
    }
  }

  // ISO yyyy-mm-dd or yyyy/mm/dd
  const ymdMatch = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[T\s]\d{1,2}:\d{2}(?::\d{2})?.*)?$/.exec(
    str
  );
  if (ymdMatch) {
    const year = Number.parseInt(ymdMatch[1], 10);
    const month = Number.parseInt(ymdMatch[2], 10);
    const day = Number.parseInt(ymdMatch[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const check = new Date(Date.UTC(year, month - 1, day));
      if (
        check.getUTCFullYear() === year &&
        check.getUTCMonth() === month - 1 &&
        check.getUTCDate() === day
      ) {
        return { year, month, day };
      }
    }
  }

  return null;
}

function formatDateDdMmYyyy(date: Date, includeTime: boolean): string {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  if (!includeTime) {
    return `${dd}/${mm}/${yyyy}`;
  }
  const hh = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return `${dd}/${mm}/${yyyy} ${hh}:${min}`;
}

function excelRoundAwayFromZero(num: number, digits: number, mode: 'round' | 'up' | 'down'): number {
  const d = Math.trunc(digits);
  const factor = 10 ** d;
  if (!Number.isFinite(factor) || factor === 0) {
    throw new FormulaError('#ERROR!');
  }
  const sign = num < 0 ? -1 : 1;
  const scaled = Math.abs(num) * factor;
  let roundedAbs: number;
  if (mode === 'up') {
    roundedAbs = Math.ceil(scaled - 1e-12) / factor;
  } else if (mode === 'down') {
    roundedAbs = Math.floor(scaled + 1e-12) / factor;
  } else {
    roundedAbs = Math.floor(scaled + 0.5 + 1e-12) / factor;
  }
  return normalizeNumber(sign * roundedAbs);
}

function toProperCaseVietnamese(text: string): string {
  const lower = text.toLocaleLowerCase('vi-VN');
  return lower.replace(
    /(^|[^\p{L}\p{N}])(\p{L})/gu,
    (_match, sep: string, letter: string) => `${sep}${letter.toLocaleUpperCase('vi-VN')}`
  );
}

const MAX_SCAN_CELLS = 50_000;

class FormulaParser {
  private readonly tokens: Token[];
  private pos = 0;
  private readonly rows: (string | number)[][];
  private readonly visiting: Set<string>;
  private readonly cache: Map<string, string | number>;

  constructor(
    tokens: Token[],
    rows: (string | number)[][],
    visiting: Set<string>,
    cache: Map<string, string | number>
  ) {
    this.tokens = tokens;
    this.rows = rows;
    this.visiting = visiting;
    this.cache = cache;
  }

  private peek(offset = 0): Token {
    return this.tokens[this.pos + offset] ?? { type: 'EOF', value: '' };
  }

  private consume(): Token {
    const token = this.peek();
    this.pos++;
    return token;
  }

  parse(): string | number {
    // Direct single-cell reference "=A1": preserve text or numeric value cleanly (empty -> 0)
    if (this.peek().type === 'CELL' && this.peek(1).type === 'EOF') {
      const cellToken = this.consume();
      const coords = parseCellCoords(cellToken.value);
      if (!coords) {
        throw new FormulaError('#ERROR!');
      }
      const val = evaluateCellInternal(
        this.rows,
        coords.row,
        coords.col,
        this.visiting,
        this.cache
      );
      if (typeof val === 'string' && FORMULA_ERRORS.has(val)) {
        throw new FormulaError(val as FormulaErrorCode);
      }
      if (val === '') {
        return 0;
      }
      const maybeNum = parseNumericLiteral(val);
      return maybeNum !== null ? normalizeNumber(maybeNum) : val;
    }

    const result = this.parseComparison();
    if (this.peek().type !== 'EOF') {
      throw new FormulaError('#ERROR!');
    }
    return finalizeOutputValue(result);
  }

  private parseComparison(): EvalValue {
    let left = this.parseConcat();
    while (
      this.peek().type === 'EQ' ||
      this.peek().type === 'NEQ' ||
      this.peek().type === 'GTE' ||
      this.peek().type === 'LTE' ||
      this.peek().type === 'GT' ||
      this.peek().type === 'LT'
    ) {
      const op = this.consume().type;
      const right = this.parseConcat();
      left = compareValues(left, right, op);
    }
    return left;
  }

  private parseConcat(): EvalValue {
    let left = this.parseAddSub();
    while (this.peek().type === 'AMP') {
      this.consume();
      const right = this.parseAddSub();
      left = `${coerceToString(left)}${coerceToString(right)}`;
    }
    return left;
  }

  private parseAddSub(): EvalValue {
    let left = this.parseMulDiv();
    while (this.peek().type === 'PLUS' || this.peek().type === 'MINUS') {
      const op = this.consume().type;
      const right = this.parseMulDiv();
      const lNum = coerceToNumber(left);
      const rNum = coerceToNumber(right);
      left = op === 'PLUS' ? lNum + rNum : lNum - rNum;
    }
    return left;
  }

  private parseMulDiv(): EvalValue {
    let left = this.parsePower();
    while (this.peek().type === 'STAR' || this.peek().type === 'SLASH') {
      const op = this.consume().type;
      const right = this.parsePower();
      const lNum = coerceToNumber(left);
      const rNum = coerceToNumber(right);
      if (op === 'SLASH') {
        if (rNum === 0) {
          throw new FormulaError('#DIV/0!');
        }
        left = lNum / rNum;
      } else {
        left = lNum * rNum;
      }
    }
    return left;
  }

  private parsePower(): EvalValue {
    let left = this.parseUnary();
    while (this.peek().type === 'CARET') {
      this.consume();
      const right = this.parseUnary();
      left = coerceToNumber(left) ** coerceToNumber(right);
    }
    return left;
  }

  private parseUnary(): EvalValue {
    if (this.peek().type === 'PLUS') {
      this.consume();
      return coerceToNumber(this.parseUnary());
    }
    if (this.peek().type === 'MINUS') {
      this.consume();
      return -coerceToNumber(this.parseUnary());
    }
    return this.parsePrimary();
  }

  private parsePrimary(): EvalValue {
    const token = this.peek();

    if (token.type === 'NUMBER') {
      this.consume();
      const parsed = Number(token.value);
      if (!Number.isFinite(parsed)) {
        throw new FormulaError('#ERROR!');
      }
      return parsed;
    }

    if (token.type === 'STRING') {
      this.consume();
      return token.value;
    }

    if (token.type === 'BOOLEAN') {
      this.consume();
      return token.value === 'TRUE';
    }

    if (token.type === 'LPAREN') {
      this.consume();
      const value = this.parseComparison();
      if (this.peek().type !== 'RPAREN') {
        throw new FormulaError('#ERROR!');
      }
      this.consume();
      return value;
    }

    if (token.type === 'IDENT') {
      return this.parseFunctionCall();
    }

    if (token.type === 'CELL') {
      this.consume();
      const coords = parseCellCoords(token.value);
      if (!coords) {
        throw new FormulaError('#ERROR!');
      }
      const cellVal = evaluateCellInternal(
        this.rows,
        coords.row,
        coords.col,
        this.visiting,
        this.cache
      );
      if (typeof cellVal === 'string' && FORMULA_ERRORS.has(cellVal.trim())) {
        throw new FormulaError(cellVal.trim() as FormulaErrorCode);
      }
      return cellVal;
    }

    throw new FormulaError('#ERROR!');
  }

  private sliceArgumentTokenGroups(): Token[][] {
    const groups: Token[][] = [];
    if (this.peek().type === 'RPAREN') {
      this.consume();
      return groups;
    }

    let current: Token[] = [];
    let depth = 0;

    while (this.peek().type !== 'EOF') {
      const tok = this.consume();
      if (tok.type === 'LPAREN') {
        depth++;
        current.push(tok);
      } else if (tok.type === 'RPAREN') {
        if (depth === 0) {
          groups.push(current);
          return groups;
        }
        depth--;
        current.push(tok);
      } else if (tok.type === 'COMMA' && depth === 0) {
        groups.push(current);
        current = [];
      } else {
        current.push(tok);
      }
    }

    throw new FormulaError('#ERROR!');
  }

  private evaluateSubTokens(subTokens: Token[]): EvalValue {
    if (subTokens.length === 0) {
      return '';
    }
    const subParser = new FormulaParser(
      [...subTokens, { type: 'EOF', value: '' }],
      this.rows,
      this.visiting,
      this.cache
    );
    const value = subParser.parseComparison();
    if (subParser.peek().type !== 'EOF') {
      throw new FormulaError('#ERROR!');
    }
    return value;
  }

  private isRangeTokenGroup(group: Token[]): boolean {
    return (
      group.length === 3 &&
      group[0].type === 'CELL' &&
      group[1].type === 'COLON' &&
      group[2].type === 'CELL'
    );
  }

  private parseRangeBounds(
    startRef: string,
    endRef: string
  ): { minRow: number; maxRow: number; minCol: number; maxCol: number } {
    const start = parseCellCoords(startRef);
    const end = parseCellCoords(endRef);
    if (!start || !end) {
      throw new FormulaError('#ERROR!');
    }

    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);

    for (const activeKey of this.visiting) {
      const [vr, vc] = activeKey.split(':').map(Number);
      if (vr >= minRow && vr <= maxRow && vc >= minCol && vc <= maxCol) {
        throw new FormulaError('#CYCLE!');
      }
    }

    return { minRow, maxRow, minCol, maxCol };
  }

  private parseRangeOrCellBounds(
    group: Token[]
  ): { minRow: number; maxRow: number; minCol: number; maxCol: number } | null {
    if (this.isRangeTokenGroup(group)) {
      return this.parseRangeBounds(group[0].value, group[2].value);
    }
    if (group.length === 1 && group[0].type === 'CELL') {
      const coords = parseCellCoords(group[0].value);
      if (!coords) {
        throw new FormulaError('#ERROR!');
      }
      return this.parseRangeBounds(group[0].value, group[0].value);
    }
    return null;
  }

  private collectStringsFromArgGroups(groups: Token[][], ignoreEmpty: boolean): string[] {
    const out: string[] = [];
    for (const group of groups) {
      if (group.length === 0) {
        if (!ignoreEmpty) out.push('');
        continue;
      }
      if (this.isRangeTokenGroup(group)) {
        const { minRow, maxRow, minCol, maxCol } = this.parseRangeBounds(
          group[0].value,
          group[2].value
        );
        const boundedMaxRow = Math.min(maxRow, Math.max(0, this.rows.length - 1));
        const maxCols = this.rows.reduce((m, r) => Math.max(m, r.length), 0);
        const boundedMaxCol = Math.min(maxCol, Math.max(0, maxCols - 1));
        for (let r = minRow; r <= boundedMaxRow; r++) {
          for (let c = minCol; c <= boundedMaxCol; c++) {
            const val = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
            if (typeof val === 'string' && FORMULA_ERRORS.has(val.trim())) {
              throw new FormulaError(val.trim() as FormulaErrorCode);
            }
            const str = val === '' ? '' : coerceToString(val);
            if (!ignoreEmpty || str !== '') {
              out.push(str);
            }
          }
        }
        continue;
      }
      const val = this.evaluateSubTokens(group);
      const str = val === '' ? '' : coerceToString(val);
      if (!ignoreEmpty || str !== '') {
        out.push(str);
      }
    }
    return out;
  }

  private parseFunctionCall(): EvalValue {
    const fnName = this.consume().value;
    if (this.peek().type !== 'LPAREN') {
      throw new FormulaError('#ERROR!');
    }
    this.consume(); // '('

    // --- LOGIC FUNCTIONS ---

    if (fnName === 'IF') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 2 || argGroups.length > 3) {
        throw new FormulaError('#ERROR!');
      }
      const isTruthy = coerceToBoolean(this.evaluateSubTokens(argGroups[0]));
      if (isTruthy) {
        return argGroups[1].length > 0 ? this.evaluateSubTokens(argGroups[1]) : 0;
      }
      if (argGroups.length === 3) {
        return argGroups[2].length > 0 ? this.evaluateSubTokens(argGroups[2]) : 0;
      }
      return false;
    }

    if (fnName === 'IFS') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 2 || argGroups.length % 2 !== 0) {
        throw new FormulaError('#ERROR!');
      }
      for (let i = 0; i < argGroups.length; i += 2) {
        const cond = coerceToBoolean(this.evaluateSubTokens(argGroups[i]));
        if (cond) {
          return argGroups[i + 1].length > 0 ? this.evaluateSubTokens(argGroups[i + 1]) : 0;
        }
      }
      throw new FormulaError('#N/A');
    }

    if (fnName === 'AND' || fnName === 'OR') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length === 0) {
        throw new FormulaError('#ERROR!');
      }
      let result = fnName === 'AND';
      let evaluatedAny = false;
      for (const group of argGroups) {
        if (this.isRangeTokenGroup(group)) {
          const { minRow, maxRow, minCol, maxCol } = this.parseRangeBounds(
            group[0].value,
            group[2].value
          );
          const boundedMaxRow = Math.min(maxRow, Math.max(0, this.rows.length - 1));
          const maxCols = this.rows.reduce((m, r) => Math.max(m, r.length), 0);
          const boundedMaxCol = Math.min(maxCol, Math.max(0, maxCols - 1));
          for (let r = minRow; r <= boundedMaxRow; r++) {
            for (let c = minCol; c <= boundedMaxCol; c++) {
              const raw = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
              if (raw === '') continue;
              const b = coerceToBoolean(raw);
              evaluatedAny = true;
              result = fnName === 'AND' ? result && b : result || b;
            }
          }
          continue;
        }
        const val = coerceToBoolean(this.evaluateSubTokens(group));
        evaluatedAny = true;
        result = fnName === 'AND' ? result && val : result || val;
      }
      if (!evaluatedAny) {
        throw new FormulaError('#ERROR!');
      }
      return result;
    }

    if (fnName === 'NOT') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return !coerceToBoolean(this.evaluateSubTokens(argGroups[0]));
    }

    // --- MATH ROUNDING / ABS / INT ---

    if (fnName === 'ROUND' || fnName === 'ROUNDUP' || fnName === 'ROUNDDOWN') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 1 || argGroups.length > 2) {
        throw new FormulaError('#ERROR!');
      }
      const num = coerceToNumber(this.evaluateSubTokens(argGroups[0]));
      const digits =
        argGroups.length === 2 && argGroups[1].length > 0
          ? coerceToNumber(this.evaluateSubTokens(argGroups[1]))
          : 0;
      const mode = fnName === 'ROUNDUP' ? 'up' : fnName === 'ROUNDDOWN' ? 'down' : 'round';
      return excelRoundAwayFromZero(num, digits, mode);
    }

    if (fnName === 'INT') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[0])));
    }

    if (fnName === 'ABS') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return Math.abs(coerceToNumber(this.evaluateSubTokens(argGroups[0])));
    }

    // --- CONDITIONAL & BLANK COUNTS / SUMS / AVERAGES ---

    if (fnName === 'COUNTA') {
      const argGroups = this.sliceArgumentTokenGroups();
      let count = 0;
      for (const group of argGroups) {
        if (group.length === 0) continue;
        if (this.isRangeTokenGroup(group)) {
          const { minRow, maxRow, minCol, maxCol } = this.parseRangeBounds(
            group[0].value,
            group[2].value
          );
          const boundedMaxRow = Math.min(maxRow, Math.max(0, this.rows.length - 1));
          const maxCols = this.rows.reduce((m, r) => Math.max(m, r.length), 0);
          const boundedMaxCol = Math.min(maxCol, Math.max(0, maxCols - 1));
          for (let r = minRow; r <= boundedMaxRow; r++) {
            for (let c = minCol; c <= boundedMaxCol; c++) {
              const val = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
              if (typeof val === 'number' || (typeof val === 'string' && val.trim() !== '')) {
                count++;
              }
            }
          }
          continue;
        }
        if (group.length === 1 && group[0].type === 'CELL') {
          const coords = parseCellCoords(group[0].value);
          if (!coords) throw new FormulaError('#ERROR!');
          const val = evaluateCellInternal(
            this.rows,
            coords.row,
            coords.col,
            this.visiting,
            this.cache
          );
          if (typeof val === 'number' || (typeof val === 'string' && val.trim() !== '')) {
            count++;
          }
          continue;
        }
        const evaluated = this.evaluateSubTokens(group);
        if (
          typeof evaluated === 'number' ||
          typeof evaluated === 'boolean' ||
          (typeof evaluated === 'string' && evaluated.trim() !== '')
        ) {
          count++;
        }
      }
      return count;
    }

    if (fnName === 'COUNTBLANK') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length === 0) {
        throw new FormulaError('#ERROR!');
      }
      let blanks = 0;
      for (const group of argGroups) {
        const bounds = this.parseRangeOrCellBounds(group);
        if (!bounds) {
          const v = this.evaluateSubTokens(group);
          if (v === '') blanks++;
          continue;
        }
        const totalCells =
          (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1);
        if (totalCells > MAX_SCAN_CELLS) {
          throw new FormulaError('#ERROR!');
        }
        for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
          for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
            const val = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
            if (val === '') {
              blanks++;
            }
          }
        }
      }
      return blanks;
    }

    if (fnName === 'COUNTIF') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 2) {
        throw new FormulaError('#ERROR!');
      }
      const bounds = this.parseRangeOrCellBounds(argGroups[0]);
      if (!bounds) {
        throw new FormulaError('#ERROR!');
      }
      const criteria = this.evaluateSubTokens(argGroups[1]);
      const totalCells =
        (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1);
      const boundedMaxRow =
        criteria === '' && totalCells <= MAX_SCAN_CELLS
          ? bounds.maxRow
          : Math.min(bounds.maxRow, Math.max(0, this.rows.length - 1));
      const maxCols = this.rows.reduce((m, r) => Math.max(m, r.length), 0);
      const boundedMaxCol =
        criteria === '' && totalCells <= MAX_SCAN_CELLS
          ? bounds.maxCol
          : Math.min(bounds.maxCol, Math.max(0, maxCols - 1));

      let count = 0;
      for (let r = bounds.minRow; r <= boundedMaxRow; r++) {
        for (let c = bounds.minCol; c <= boundedMaxCol; c++) {
          const val = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
          if (matchesExcelCriteria(val, criteria)) {
            count++;
          }
        }
      }
      return count;
    }

    if (fnName === 'SUMIF' || fnName === 'AVERAGEIF') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 2 || argGroups.length > 3) {
        throw new FormulaError('#ERROR!');
      }
      const critBounds = this.parseRangeOrCellBounds(argGroups[0]);
      if (!critBounds) {
        throw new FormulaError('#ERROR!');
      }
      const criteria = this.evaluateSubTokens(argGroups[1]);
      const sumBounds =
        argGroups.length === 3 && argGroups[2].length > 0
          ? this.parseRangeOrCellBounds(argGroups[2])
          : critBounds;
      if (!sumBounds) {
        throw new FormulaError('#ERROR!');
      }

      const rowSpan = critBounds.maxRow - critBounds.minRow;
      const colSpan = critBounds.maxCol - critBounds.minCol;
      const maxDataRow = Math.max(0, this.rows.length - 1);
      const maxDataCol = Math.max(0, this.rows.reduce((m, r) => Math.max(m, r.length), 0) - 1);

      // Guard cycles on target sum range if different from critBounds
      if (sumBounds !== critBounds) {
        for (const activeKey of this.visiting) {
          const [vr, vc] = activeKey.split(':').map(Number);
          if (
            vr >= sumBounds.minRow &&
            vr <= sumBounds.minRow + rowSpan &&
            vc >= sumBounds.minCol &&
            vc <= sumBounds.minCol + colSpan
          ) {
            throw new FormulaError('#CYCLE!');
          }
        }
      }

      const boundedRowSpan = Math.min(
        rowSpan,
        Math.max(maxDataRow - critBounds.minRow, maxDataRow - sumBounds.minRow, 0)
      );
      const boundedColSpan = Math.min(
        colSpan,
        Math.max(maxDataCol - critBounds.minCol, maxDataCol - sumBounds.minCol, 0)
      );

      let sum = 0;
      let matchedNumericCount = 0;
      for (let dr = 0; dr <= boundedRowSpan; dr++) {
        for (let dc = 0; dc <= boundedColSpan; dc++) {
          const critVal = evaluateCellInternal(
            this.rows,
            critBounds.minRow + dr,
            critBounds.minCol + dc,
            this.visiting,
            this.cache
          );
          if (!matchesExcelCriteria(critVal, criteria)) {
            continue;
          }
          const targetVal = evaluateCellInternal(
            this.rows,
            sumBounds.minRow + dr,
            sumBounds.minCol + dc,
            this.visiting,
            this.cache
          );
          if (typeof targetVal === 'string' && FORMULA_ERRORS.has(targetVal.trim())) {
            throw new FormulaError(targetVal.trim() as FormulaErrorCode);
          }
          const num = parseNumericLiteral(targetVal);
          if (num !== null) {
            sum += num;
            matchedNumericCount++;
          }
        }
      }

      if (fnName === 'AVERAGEIF') {
        if (matchedNumericCount === 0) {
          throw new FormulaError('#DIV/0!');
        }
        return sum / matchedNumericCount;
      }
      return sum;
    }

    // --- STRING FUNCTIONS ---

    if (fnName === 'CONCATENATE') {
      const argGroups = this.sliceArgumentTokenGroups();
      return argGroups.map((g) => coerceToString(this.evaluateSubTokens(g))).join('');
    }

    if (fnName === 'CONCAT') {
      const argGroups = this.sliceArgumentTokenGroups();
      return this.collectStringsFromArgGroups(argGroups, false).join('');
    }

    if (fnName === 'TEXTJOIN') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 3) {
        throw new FormulaError('#ERROR!');
      }
      const delimiter = coerceToString(this.evaluateSubTokens(argGroups[0]));
      const ignoreEmpty = coerceToBoolean(this.evaluateSubTokens(argGroups[1]));
      const parts = this.collectStringsFromArgGroups(argGroups.slice(2), ignoreEmpty);
      return parts.join(delimiter);
    }

    if (fnName === 'LEFT' || fnName === 'RIGHT') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 1 || argGroups.length > 2) {
        throw new FormulaError('#ERROR!');
      }
      const text = coerceToString(this.evaluateSubTokens(argGroups[0]));
      const count =
        argGroups.length === 2 && argGroups[1].length > 0
          ? Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[1])))
          : 1;
      if (count < 0) {
        throw new FormulaError('#ERROR!');
      }
      if (count === 0) {
        return '';
      }
      return fnName === 'LEFT' ? text.slice(0, count) : text.slice(Math.max(0, text.length - count));
    }

    if (fnName === 'MID') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 3) {
        throw new FormulaError('#ERROR!');
      }
      const text = coerceToString(this.evaluateSubTokens(argGroups[0]));
      const startNum = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[1])));
      const numChars = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[2])));
      if (startNum < 1 || numChars < 0) {
        throw new FormulaError('#ERROR!');
      }
      return text.slice(startNum - 1, startNum - 1 + numChars);
    }

    if (fnName === 'LEN') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return coerceToString(this.evaluateSubTokens(argGroups[0])).length;
    }

    if (fnName === 'TRIM') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return coerceToString(this.evaluateSubTokens(argGroups[0]))
        .trim()
        .replace(/ {2,}/g, ' ');
    }

    if (fnName === 'UPPER') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return coerceToString(this.evaluateSubTokens(argGroups[0])).toLocaleUpperCase('vi-VN');
    }

    if (fnName === 'LOWER') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return coerceToString(this.evaluateSubTokens(argGroups[0])).toLocaleLowerCase('vi-VN');
    }

    if (fnName === 'PROPER') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      return toProperCaseVietnamese(coerceToString(this.evaluateSubTokens(argGroups[0])));
    }

    // --- DATE & TIME FUNCTIONS ---

    if (fnName === 'TODAY' || fnName === 'NOW') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length > 0) {
        throw new FormulaError('#ERROR!');
      }
      return formatDateDdMmYyyy(new Date(), fnName === 'NOW');
    }

    if (fnName === 'DAY' || fnName === 'MONTH' || fnName === 'YEAR') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length !== 1) {
        throw new FormulaError('#ERROR!');
      }
      const parsedDate = parseExcelDateInput(this.evaluateSubTokens(argGroups[0]));
      if (!parsedDate) {
        throw new FormulaError('#ERROR!');
      }
      if (fnName === 'DAY') return parsedDate.day;
      if (fnName === 'MONTH') return parsedDate.month;
      return parsedDate.year;
    }

    // --- FINANCIAL FUNCTIONS ---

    if (fnName === 'PMT' || fnName === 'FV' || fnName === 'PV') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 3 || argGroups.length > 5) {
        throw new FormulaError('#ERROR!');
      }
      const rate = coerceToNumber(this.evaluateSubTokens(argGroups[0]));
      const nper = coerceToNumber(this.evaluateSubTokens(argGroups[1]));
      const arg3 = coerceToNumber(this.evaluateSubTokens(argGroups[2]));
      const arg4 =
        argGroups.length >= 4 && argGroups[3].length > 0
          ? coerceToNumber(this.evaluateSubTokens(argGroups[3]))
          : 0;
      const rawType =
        argGroups.length === 5 && argGroups[4].length > 0
          ? coerceToNumber(this.evaluateSubTokens(argGroups[4]))
          : 0;
      const payType = rawType !== 0 ? 1 : 0;

      if (fnName === 'PMT') {
        if (nper === 0) {
          throw new FormulaError('#DIV/0!');
        }
        const pv = arg3;
        const fv = arg4;
        if (rate === 0) {
          return normalizeNumber(-(pv + fv) / nper);
        }
        const pvif = (1 + rate) ** nper;
        const denom = (1 + rate * payType) * (pvif - 1);
        if (!Number.isFinite(pvif) || !Number.isFinite(denom) || denom === 0) {
          throw new FormulaError('#DIV/0!');
        }
        return normalizeNumber((-rate * (pv * pvif + fv)) / denom);
      }

      if (fnName === 'FV') {
        const pmt = arg3;
        const pv = arg4;
        if (rate === 0) {
          return normalizeNumber(-(pv + pmt * nper));
        }
        const pvif = (1 + rate) ** nper;
        if (!Number.isFinite(pvif)) {
          throw new FormulaError('#ERROR!');
        }
        return normalizeNumber(
          -(pv * pvif + pmt * (1 + rate * payType) * ((pvif - 1) / rate))
        );
      }

      // PV(rate, nper, pmt, [fv], [type])
      const pmt = arg3;
      const fv = arg4;
      if (rate === 0) {
        return normalizeNumber(-(fv + pmt * nper));
      }
      const pvif = (1 + rate) ** nper;
      if (!Number.isFinite(pvif) || pvif === 0) {
        throw new FormulaError('#DIV/0!');
      }
      return normalizeNumber(
        -(fv + pmt * (1 + rate * payType) * ((pvif - 1) / rate)) / pvif
      );
    }

    // --- LOOKUP FUNCTIONS ---

    if (fnName === 'VLOOKUP') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 3 || argGroups.length > 4) {
        throw new FormulaError('#ERROR!');
      }
      const lookupVal = this.evaluateSubTokens(argGroups[0]);
      const bounds = this.parseRangeOrCellBounds(argGroups[1]);
      if (!bounds) {
        throw new FormulaError('#ERROR!');
      }

      const colIndex = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[2])));
      const rangeWidth = bounds.maxCol - bounds.minCol + 1;
      if (colIndex < 1 || colIndex > rangeWidth) {
        throw new FormulaError('#ERROR!');
      }

      let rangeLookup = true;
      if (argGroups.length === 4 && argGroups[3].length > 0) {
        rangeLookup = coerceToBoolean(this.evaluateSubTokens(argGroups[3]));
      }

      const boundedMaxRow = Math.min(bounds.maxRow, Math.max(0, this.rows.length - 1));
      const targetCol = bounds.minCol + colIndex - 1;

      if (!rangeLookup) {
        for (let r = bounds.minRow; r <= boundedMaxRow; r++) {
          const firstColVal = evaluateCellInternal(
            this.rows,
            r,
            bounds.minCol,
            this.visiting,
            this.cache
          );
          if (valuesMatchForLookup(firstColVal, lookupVal)) {
            const resultVal = evaluateCellInternal(
              this.rows,
              r,
              targetCol,
              this.visiting,
              this.cache
            );
            if (typeof resultVal === 'string' && FORMULA_ERRORS.has(resultVal.trim())) {
              throw new FormulaError(resultVal.trim() as FormulaErrorCode);
            }
            return resultVal === '' ? 0 : resultVal;
          }
        }
        throw new FormulaError('#N/A');
      }

      let bestRow = -1;
      for (let r = bounds.minRow; r <= boundedMaxRow; r++) {
        const firstColVal = evaluateCellInternal(
          this.rows,
          r,
          bounds.minCol,
          this.visiting,
          this.cache
        );
        const diff = compareForApproxLookup(firstColVal, lookupVal);
        if (diff === null) continue;
        if (diff === 0) {
          bestRow = r;
          break;
        }
        if (diff < 0) {
          bestRow = r;
        } else if (bestRow !== -1) {
          break;
        }
      }

      if (bestRow === -1) {
        throw new FormulaError('#N/A');
      }

      const resultVal = evaluateCellInternal(
        this.rows,
        bestRow,
        targetCol,
        this.visiting,
        this.cache
      );
      if (typeof resultVal === 'string' && FORMULA_ERRORS.has(resultVal.trim())) {
        throw new FormulaError(resultVal.trim() as FormulaErrorCode);
      }
      return resultVal === '' ? 0 : resultVal;
    }

    if (fnName === 'HLOOKUP') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 3 || argGroups.length > 4) {
        throw new FormulaError('#ERROR!');
      }
      const lookupVal = this.evaluateSubTokens(argGroups[0]);
      const bounds = this.parseRangeOrCellBounds(argGroups[1]);
      if (!bounds) {
        throw new FormulaError('#ERROR!');
      }

      const rowIndex = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[2])));
      const rangeHeight = bounds.maxRow - bounds.minRow + 1;
      if (rowIndex < 1 || rowIndex > rangeHeight) {
        throw new FormulaError('#ERROR!');
      }

      let rangeLookup = true;
      if (argGroups.length === 4 && argGroups[3].length > 0) {
        rangeLookup = coerceToBoolean(this.evaluateSubTokens(argGroups[3]));
      }

      const maxCols = this.rows.reduce((m, r) => Math.max(m, r.length), 0);
      const boundedMaxCol = Math.min(bounds.maxCol, Math.max(0, maxCols - 1));
      const targetRow = bounds.minRow + rowIndex - 1;

      if (!rangeLookup) {
        for (let c = bounds.minCol; c <= boundedMaxCol; c++) {
          const firstRowVal = evaluateCellInternal(
            this.rows,
            bounds.minRow,
            c,
            this.visiting,
            this.cache
          );
          if (valuesMatchForLookup(firstRowVal, lookupVal)) {
            const resultVal = evaluateCellInternal(
              this.rows,
              targetRow,
              c,
              this.visiting,
              this.cache
            );
            if (typeof resultVal === 'string' && FORMULA_ERRORS.has(resultVal.trim())) {
              throw new FormulaError(resultVal.trim() as FormulaErrorCode);
            }
            return resultVal === '' ? 0 : resultVal;
          }
        }
        throw new FormulaError('#N/A');
      }

      let bestCol = -1;
      for (let c = bounds.minCol; c <= boundedMaxCol; c++) {
        const firstRowVal = evaluateCellInternal(
          this.rows,
          bounds.minRow,
          c,
          this.visiting,
          this.cache
        );
        const diff = compareForApproxLookup(firstRowVal, lookupVal);
        if (diff === null) continue;
        if (diff === 0) {
          bestCol = c;
          break;
        }
        if (diff < 0) {
          bestCol = c;
        } else if (bestCol !== -1) {
          break;
        }
      }

      if (bestCol === -1) {
        throw new FormulaError('#N/A');
      }

      const resultVal = evaluateCellInternal(
        this.rows,
        targetRow,
        bestCol,
        this.visiting,
        this.cache
      );
      if (typeof resultVal === 'string' && FORMULA_ERRORS.has(resultVal.trim())) {
        throw new FormulaError(resultVal.trim() as FormulaErrorCode);
      }
      return resultVal === '' ? 0 : resultVal;
    }

    if (fnName === 'MATCH') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 2 || argGroups.length > 3) {
        throw new FormulaError('#ERROR!');
      }
      const lookupVal = this.evaluateSubTokens(argGroups[0]);
      const bounds = this.parseRangeOrCellBounds(argGroups[1]);
      if (!bounds) {
        throw new FormulaError('#ERROR!');
      }
      const matchType =
        argGroups.length === 3 && argGroups[2].length > 0
          ? Math.trunc(coerceToNumber(this.evaluateSubTokens(argGroups[2])))
          : 1;

      const rowCount = bounds.maxRow - bounds.minRow + 1;
      const colCount = bounds.maxCol - bounds.minCol + 1;
      if (rowCount > 1 && colCount > 1) {
        throw new FormulaError('#N/A');
      }

      const values: Array<string | number> = [];
      if (rowCount === 1) {
        for (let c = bounds.minCol; c <= bounds.maxCol; c++) {
          values.push(evaluateCellInternal(this.rows, bounds.minRow, c, this.visiting, this.cache));
        }
      } else {
        for (let r = bounds.minRow; r <= bounds.maxRow; r++) {
          values.push(evaluateCellInternal(this.rows, r, bounds.minCol, this.visiting, this.cache));
        }
      }

      if (matchType === 0) {
        for (let idx = 0; idx < values.length; idx++) {
          if (valuesMatchForLookup(values[idx], lookupVal)) {
            return idx + 1;
          }
        }
        throw new FormulaError('#N/A');
      }

      if (matchType > 0) {
        let bestIdx = -1;
        for (let idx = 0; idx < values.length; idx++) {
          const diff = compareForApproxLookup(values[idx], lookupVal);
          if (diff === null) continue;
          if (diff <= 0) {
            bestIdx = idx;
          } else {
            break;
          }
        }
        if (bestIdx === -1) throw new FormulaError('#N/A');
        return bestIdx + 1;
      }

      // matchType < 0: descending order, smallest value >= lookupVal
      let bestIdx = -1;
      for (let idx = 0; idx < values.length; idx++) {
        const diff = compareForApproxLookup(values[idx], lookupVal);
        if (diff === null) continue;
        if (diff >= 0) {
          bestIdx = idx;
        } else {
          break;
        }
      }
      if (bestIdx === -1) throw new FormulaError('#N/A');
      return bestIdx + 1;
    }

    if (fnName === 'INDEX') {
      const argGroups = this.sliceArgumentTokenGroups();
      if (argGroups.length < 2 || argGroups.length > 3) {
        throw new FormulaError('#ERROR!');
      }
      const bounds = this.parseRangeOrCellBounds(argGroups[0]);
      if (!bounds) {
        throw new FormulaError('#ERROR!');
      }
      const rowCount = bounds.maxRow - bounds.minRow + 1;
      const colCount = bounds.maxCol - bounds.minCol + 1;

      const arg1 = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[1])));
      let targetRowOffset: number;
      let targetColOffset: number;

      if (argGroups.length === 2 || argGroups[2].length === 0) {
        if (rowCount === 1 && colCount > 1) {
          targetRowOffset = 0;
          targetColOffset = arg1 - 1;
        } else {
          targetRowOffset = arg1 - 1;
          targetColOffset = 0;
        }
      } else {
        const arg2 = Math.floor(coerceToNumber(this.evaluateSubTokens(argGroups[2])));
        targetRowOffset = arg1 - 1;
        targetColOffset = arg2 - 1;
      }

      if (
        targetRowOffset < 0 ||
        targetRowOffset >= rowCount ||
        targetColOffset < 0 ||
        targetColOffset >= colCount
      ) {
        throw new FormulaError('#ERROR!');
      }

      const resultVal = evaluateCellInternal(
        this.rows,
        bounds.minRow + targetRowOffset,
        bounds.minCol + targetColOffset,
        this.visiting,
        this.cache
      );
      if (typeof resultVal === 'string' && FORMULA_ERRORS.has(resultVal.trim())) {
        throw new FormulaError(resultVal.trim() as FormulaErrorCode);
      }
      return resultVal === '' ? 0 : resultVal;
    }

    // --- NUMERIC AGGREGATE FUNCTIONS: SUM, AVERAGE/AVG, COUNT, MIN, MAX ---
    const numbers: number[] = [];

    if (this.peek().type !== 'RPAREN') {
      while (true) {
        this.collectNumericArgument(numbers);
        if (this.peek().type === 'COMMA') {
          this.consume();
          continue;
        }
        break;
      }
    }

    if (this.peek().type !== 'RPAREN') {
      throw new FormulaError('#ERROR!');
    }
    this.consume(); // ')'

    switch (fnName) {
      case 'SUM':
        return numbers.reduce((acc, n) => acc + n, 0);
      case 'AVERAGE':
      case 'AVG':
        if (numbers.length === 0) {
          throw new FormulaError('#DIV/0!');
        }
        return numbers.reduce((acc, n) => acc + n, 0) / numbers.length;
      case 'COUNT':
        return numbers.length;
      case 'MIN':
        return numbers.length === 0 ? 0 : Math.min(...numbers);
      case 'MAX':
        return numbers.length === 0 ? 0 : Math.max(...numbers);
      default:
        throw new FormulaError('#ERROR!');
    }
  }

  private collectNumericArgument(target: number[]): void {
    // Check for range argument: CELL : CELL
    if (
      this.peek().type === 'CELL' &&
      this.peek(1).type === 'COLON' &&
      this.peek(2).type === 'CELL'
    ) {
      const startTok = this.consume();
      this.consume(); // ':'
      const endTok = this.consume();

      const { minRow, maxRow, minCol, maxCol } = this.parseRangeBounds(
        startTok.value,
        endTok.value
      );

      const boundedMaxRow = Math.min(maxRow, Math.max(0, this.rows.length - 1));
      const maxPopulatedCols = this.rows.reduce((max, row) => Math.max(max, row.length), 0);
      const boundedMaxCol = Math.min(maxCol, Math.max(0, maxPopulatedCols - 1));

      for (let r = minRow; r <= boundedMaxRow; r++) {
        for (let c = minCol; c <= boundedMaxCol; c++) {
          const val = evaluateCellInternal(this.rows, r, c, this.visiting, this.cache);
          if (typeof val === 'string' && FORMULA_ERRORS.has(val)) {
            throw new FormulaError(val as FormulaErrorCode);
          }
          const num = parseNumericLiteral(val);
          if (num !== null) {
            target.push(num);
          }
        }
      }
      return;
    }

    // Standalone single cell argument: ignore blank/non-numeric cells just like Excel
    if (
      this.peek().type === 'CELL' &&
      (this.peek(1).type === 'COMMA' || this.peek(1).type === 'RPAREN')
    ) {
      const cellTok = this.consume();
      const coords = parseCellCoords(cellTok.value);
      if (!coords) {
        throw new FormulaError('#ERROR!');
      }
      const val = evaluateCellInternal(
        this.rows,
        coords.row,
        coords.col,
        this.visiting,
        this.cache
      );
      if (typeof val === 'string' && FORMULA_ERRORS.has(val)) {
        throw new FormulaError(val as FormulaErrorCode);
      }
      const num = parseNumericLiteral(val);
      if (num !== null) {
        target.push(num);
      }
      return;
    }

    // General numeric expression argument (e.g. A1+B1, 42, SUM(C1:C3)*2)
    const computed = coerceToNumber(this.parseComparison());
    if (!Number.isFinite(computed)) {
      throw new FormulaError('#ERROR!');
    }
    target.push(computed);
  }
}

function evaluateCellInternal(
  rows: (string | number)[][],
  rowIdx: number,
  colIdx: number,
  visiting: Set<string>,
  cache: Map<string, string | number>
): string | number {
  if (rowIdx < 0 || colIdx < 0 || rowIdx >= rows.length) {
    return '';
  }
  const row = rows[rowIdx];
  if (!row || colIdx >= row.length) {
    return '';
  }

  const key = `${rowIdx}:${colIdx}`;
  if (cache.has(key)) {
    return cache.get(key)!;
  }

  const raw = row[colIdx];
  if (typeof raw === 'number') {
    cache.set(key, raw);
    return raw;
  }

  if (typeof raw !== 'string') {
    return '';
  }

  const trimmed = raw.trim();
  if (!trimmed.startsWith('=')) {
    cache.set(key, raw);
    return raw;
  }

  if (visiting.has(key)) {
    return '#CYCLE!';
  }

  const expression = trimmed.slice(1).trim();
  if (!expression) {
    cache.set(key, '#ERROR!');
    return '#ERROR!';
  }

  visiting.add(key);
  try {
    const tokens = tokenize(expression);
    const parser = new FormulaParser(tokens, rows, visiting, cache);
    const result = parser.parse();
    cache.set(key, result);
    return result;
  } catch (err) {
    const code: FormulaErrorCode = err instanceof FormulaError ? err.code : '#ERROR!';
    cache.set(key, code);
    return code;
  } finally {
    visiting.delete(key);
  }
}

export function evaluateCellValue(
  rows: (string | number)[][],
  rowIdx: number,
  colIdx: number
): string | number {
  const visiting = new Set<string>();
  const cache = new Map<string, string | number>();
  return evaluateCellInternal(rows, rowIdx, colIdx, visiting, cache);
}

export function evaluateSheetMatrix(rows: (string | number)[][]): (string | number)[][] {
  const cache = new Map<string, string | number>();
  return rows.map((row, rowIdx) =>
    row.map((_, colIdx) => {
      const visiting = new Set<string>();
      return evaluateCellInternal(rows, rowIdx, colIdx, visiting, cache);
    })
  );
}

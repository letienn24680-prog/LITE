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

type FormulaErrorCode = '#CYCLE!' | '#DIV/0!' | '#ERROR!';

class FormulaError extends Error {
  readonly code: FormulaErrorCode;

  constructor(code: FormulaErrorCode) {
    super(code);
    this.code = code;
  }
}

const FORMULA_ERRORS = new Set<string>(['#CYCLE!', '#DIV/0!', '#ERROR!']);

type TokenType =
  | 'NUMBER'
  | 'IDENT'
  | 'CELL'
  | 'PLUS'
  | 'MINUS'
  | 'STAR'
  | 'SLASH'
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
      let start = i;
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

    // Identifier or cell reference (supports optional $ for absolute refs like $A$1)
    if (/[A-Za-z_$]/.test(ch)) {
      const start = i;
      while (i < len && /[A-Za-z0-9_$]/.test(expression[i])) {
        i++;
      }
      const raw = expression.slice(start, i);
      // Peek ahead skipping spaces to check if this is a function call
      let lookahead = i;
      while (lookahead < len && /\s/.test(expression[lookahead])) {
        lookahead++;
      }
      if (lookahead < len && expression[lookahead] === '(') {
        tokens.push({ type: 'IDENT', value: raw.toUpperCase() });
      } else if (CELL_REF_REGEX.test(raw)) {
        tokens.push({ type: 'CELL', value: raw.toUpperCase() });
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

function parseNumericLiteral(value: string | number): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
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

function coerceToNumber(val: string | number): number {
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) {
      throw new FormulaError('#ERROR!');
    }
    return val;
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
    // Allow direct single-cell reference formulas like "=A1" to preserve text values
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

    const result = this.parseAddSub();
    if (this.peek().type !== 'EOF') {
      throw new FormulaError('#ERROR!');
    }
    return normalizeNumber(result);
  }

  private parseAddSub(): number {
    let left = this.parseMulDiv();
    while (this.peek().type === 'PLUS' || this.peek().type === 'MINUS') {
      const op = this.consume().type;
      const right = this.parseMulDiv();
      left = op === 'PLUS' ? left + right : left - right;
    }
    return left;
  }

  private parseMulDiv(): number {
    let left = this.parseUnary();
    while (this.peek().type === 'STAR' || this.peek().type === 'SLASH') {
      const op = this.consume().type;
      const right = this.parseUnary();
      if (op === 'SLASH') {
        if (right === 0) {
          throw new FormulaError('#DIV/0!');
        }
        left = left / right;
      } else {
        left = left * right;
      }
    }
    return left;
  }

  private parseUnary(): number {
    if (this.peek().type === 'PLUS') {
      this.consume();
      return this.parseUnary();
    }
    if (this.peek().type === 'MINUS') {
      this.consume();
      return -this.parseUnary();
    }
    return this.parsePrimary();
  }

  private parsePrimary(): number {
    const token = this.peek();

    if (token.type === 'NUMBER') {
      this.consume();
      const parsed = Number(token.value);
      if (!Number.isFinite(parsed)) {
        throw new FormulaError('#ERROR!');
      }
      return parsed;
    }

    if (token.type === 'LPAREN') {
      this.consume();
      const value = this.parseAddSub();
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
      return coerceToNumber(cellVal);
    }

    throw new FormulaError('#ERROR!');
  }

  private parseFunctionCall(): number {
    const fnName = this.consume().value;
    if (this.peek().type !== 'LPAREN') {
      throw new FormulaError('#ERROR!');
    }
    this.consume(); // '('

    const numbers: number[] = [];

    if (this.peek().type !== 'RPAREN') {
      while (true) {
        this.collectFunctionArgument(numbers);
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

  private collectFunctionArgument(target: number[]): void {
    // Check for range argument: CELL : CELL
    if (
      this.peek().type === 'CELL' &&
      this.peek(1).type === 'COLON' &&
      this.peek(2).type === 'CELL'
    ) {
      const startTok = this.consume();
      this.consume(); // ':'
      const endTok = this.consume();

      const start = parseCellCoords(startTok.value);
      const end = parseCellCoords(endTok.value);
      if (!start || !end) {
        throw new FormulaError('#ERROR!');
      }

      const minRow = Math.min(start.row, end.row);
      const maxRow = Math.max(start.row, end.row);
      const minCol = Math.min(start.col, end.col);
      const maxCol = Math.max(start.col, end.col);

      // Guard against cyclic self-reference even if the cell is outside populated bounds
      for (const activeKey of this.visiting) {
        const [vr, vc] = activeKey.split(':').map(Number);
        if (vr >= minRow && vr <= maxRow && vc >= minCol && vc <= maxCol) {
          throw new FormulaError('#CYCLE!');
        }
      }

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
    const computed = this.parseAddSub();
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

import JSZip from 'jszip';
import mammoth from 'mammoth/mammoth.browser';
import * as XLSX from 'xlsx';
import { createThumbnailDataUri } from '../data/mockDocuments';
import type {
  DocumentItem,
  DocumentType,
  OutlineSection,
  SlideImageItem,
  SlidePreviewItem,
  SlideTextBlock,
  SpreadsheetSheet,
} from '../types/document';
import { columnIndexToLetter } from './formulaEvaluator';
import { inspectPdfMetadata } from './pdfEngine';

export interface ParseFileResult {
  document?: DocumentItem;
  error?: string;
  createdObjectUrls: string[];
}

const IMAGE_MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
};

const BLOCK_TAGS = new Set([
  'ADDRESS',
  'ARTICLE',
  'ASIDE',
  'BLOCKQUOTE',
  'DD',
  'DIV',
  'DL',
  'DT',
  'FIGCAPTION',
  'FIGURE',
  'FOOTER',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'HEADER',
  'HR',
  'LI',
  'MAIN',
  'NAV',
  'OL',
  'P',
  'PRE',
  'SECTION',
  'TABLE',
  'TD',
  'TH',
  'TR',
  'UL',
]);

const FORBIDDEN_HTML_TAGS = new Set([
  'SCRIPT',
  'STYLE',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'FORM',
  'LINK',
  'META',
  'BASE',
  'NOSCRIPT',
]);

function toRomanNumeral(num: number): string {
  if (!Number.isFinite(num) || num <= 0) {
    return 'I';
  }
  const map: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.floor(num);
  let result = '';
  for (const [value, symbol] of map) {
    while (remaining >= value) {
      result += symbol;
      remaining -= value;
    }
  }
  return result;
}

function formatFileSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return '0 B';
  }
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function stripFileExtension(fileName: string): string {
  const trimmed = fileName.trim();
  const dotIdx = trimmed.lastIndexOf('.');
  if (dotIdx <= 0) {
    return trimmed || 'Tài liệu mới';
  }
  return trimmed.slice(0, dotIdx).trim() || 'Tài liệu mới';
}

function getFileExtension(fileName: string): string {
  const trimmed = fileName.trim();
  const dotIdx = trimmed.lastIndexOf('.');
  if (dotIdx < 0) {
    return '';
  }
  return trimmed.slice(dotIdx).toLowerCase();
}

function truncatePreview(text: string, maxLength = 260): string {
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 1).trimEnd()}…`;
}

function createDocumentId(type: DocumentType): string {
  const rand = Math.random().toString(36).slice(2, 8);
  return `uploaded-${type}-${Date.now()}-${rand}`;
}

function extractReadableText(root: Node): string {
  const chunks: string[] = [];

  const walk = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const value = node.nodeValue ?? '';
      if (value) {
        chunks.push(value);
      }
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const el = node as Element;
    const tag = el.tagName.toUpperCase();
    if (tag === 'BR') {
      chunks.push('\n');
      return;
    }
    if (FORBIDDEN_HTML_TAGS.has(tag)) {
      return;
    }

    const isBlock = BLOCK_TAGS.has(tag);
    if (isBlock && chunks.length > 0) {
      chunks.push('\n');
    }
    for (let i = 0; i < el.childNodes.length; i++) {
      walk(el.childNodes[i]);
    }
    if (isBlock) {
      chunks.push('\n');
    }
  };

  walk(root);

  return chunks
    .join('')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function isSafeUrl(rawUrl: string, allowDataImage: boolean): boolean {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return true;
  }
  const normalized = trimmed.replace(/[\u0000-\u0020]+/g, '').toLowerCase();
  if (
    normalized.startsWith('javascript:') ||
    normalized.startsWith('vbscript:') ||
    normalized.startsWith('data:text/html')
  ) {
    return false;
  }
  if (normalized.startsWith('data:')) {
    return allowDataImage && normalized.startsWith('data:image/');
  }
  return true;
}

function sanitizeWordHtml(rawHtml: string): string {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(rawHtml || '', 'text/html');

  const removeNodes: Element[] = [];
  const allElements = parsed.body.querySelectorAll('*');
  allElements.forEach((el) => {
    if (FORBIDDEN_HTML_TAGS.has(el.tagName.toUpperCase())) {
      removeNodes.push(el);
      return;
    }

    const attrs = Array.from(el.attributes);
    for (const attr of attrs) {
      const nameLower = attr.name.toLowerCase();
      if (nameLower.startsWith('on')) {
        el.removeAttribute(attr.name);
        continue;
      }
      if (nameLower === 'src') {
        if (!isSafeUrl(attr.value, true)) {
          el.removeAttribute(attr.name);
        }
        continue;
      }
      if (nameLower === 'href' || nameLower === 'xlink:href') {
        if (!isSafeUrl(attr.value, false)) {
          el.removeAttribute(attr.name);
        }
      }
    }
  });

  removeNodes.forEach((node) => node.remove());
  return parsed.body.innerHTML;
}

export function extractOutlineFromHtml(html: string): {
  htmlWithIds: string;
  outline: OutlineSection[];
  plainText: string;
} {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(html || '', 'text/html');
  const headingElements = Array.from(parsed.body.querySelectorAll('h1, h2, h3'));

  const validHeadings = headingElements.filter(
    (el) => (el.textContent ?? '').replace(/\s+/g, ' ').trim().length > 0
  );

  const minLevel =
    validHeadings.length > 0
      ? Math.min(...validHeadings.map((el) => Number.parseInt(el.tagName.slice(1), 10)))
      : 1;

  const outline: OutlineSection[] = [];
  let topCount = 0;
  let subCount = 0;
  let subSubCount = 0;
  let headingSeq = 0;

  // Walk top-level siblings to gather real summary/keyPoints following each heading
  const allNodes = Array.from(parsed.body.querySelectorAll('h1, h2, h3, p, li'));

  for (let i = 0; i < validHeadings.length; i++) {
    const headingEl = validHeadings[i];
    headingSeq++;
    const id = `heading-${headingSeq}`;
    headingEl.setAttribute('id', id);

    const level = Number.parseInt(headingEl.tagName.slice(1), 10);
    let romanIndex: string;
    if (level <= minLevel) {
      topCount++;
      subCount = 0;
      subSubCount = 0;
      romanIndex = toRomanNumeral(topCount);
    } else if (level === minLevel + 1) {
      if (topCount === 0) {
        topCount = 1;
      }
      subCount++;
      subSubCount = 0;
      romanIndex = `${topCount}.${subCount}`;
    } else {
      if (topCount === 0) {
        topCount = 1;
      }
      if (subCount === 0) {
        subCount = 1;
      }
      subSubCount++;
      romanIndex = `${topCount}.${subCount}.${subSubCount}`;
    }

    const title = (headingEl.textContent ?? '').replace(/\s+/g, ' ').trim();

    // Collect real summary and bullet points between this heading and the next heading
    const startIdx = allNodes.indexOf(headingEl);
    const nextHeadingEl = validHeadings[i + 1];
    const endIdx = nextHeadingEl ? allNodes.indexOf(nextHeadingEl) : allNodes.length;

    let summary = '';
    const keyPoints: string[] = [];

    if (startIdx >= 0) {
      for (let j = startIdx + 1; j < endIdx; j++) {
        const candidate = allNodes[j];
        const tag = candidate.tagName.toUpperCase();
        const text = (candidate.textContent ?? '').replace(/\s+/g, ' ').trim();
        if (!text) {
          continue;
        }
        if (tag === 'LI' && keyPoints.length < 3) {
          keyPoints.push(truncatePreview(text, 140));
        } else if (tag === 'P' && !summary) {
          summary = truncatePreview(text, 180);
        }
      }
    }

    outline.push({
      id,
      romanIndex,
      title,
      summary,
      keyPoints,
      pageNumber: Math.max(1, Math.floor(i / 3) + 1),
    });
  }

  const plainText = extractReadableText(parsed.body);

  return {
    htmlWithIds: parsed.body.innerHTML,
    outline,
    plainText,
  };
}

// --- XML & ZIP helpers for PPTX ---

function getElementsByLocalName(root: Document | Element, localName: string): Element[] {
  const target = localName.toLowerCase();
  const matched: Element[] = [];
  const all = root.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName.toLowerCase() === target) {
      matched.push(el);
    }
  }
  return matched;
}

function getDirectChildrenByLocalName(parent: Element, localName: string): Element[] {
  const target = localName.toLowerCase();
  const result: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === Node.ELEMENT_NODE) {
      const el = child as Element;
      if (el.localName.toLowerCase() === target) {
        result.push(el);
      }
    }
  }
  return result;
}

function findZipFileInsensitive(zip: JSZip, targetPath: string): JSZip.JSZipObject | null {
  const normalized = targetPath.replace(/^\/+/, '');
  const exact = zip.file(normalized);
  if (exact) {
    return exact;
  }
  const lower = normalized.toLowerCase();
  for (const key of Object.keys(zip.files)) {
    if (key.toLowerCase() === lower) {
      return zip.files[key];
    }
  }
  return null;
}

function resolveZipPath(baseDir: string, relativeTarget: string): string {
  const cleanTarget = relativeTarget.replace(/\\/g, '/').trim();
  if (cleanTarget.startsWith('/')) {
    return cleanTarget.replace(/^\/+/, '');
  }
  const parts = baseDir.replace(/\\/g, '/').split('/').filter(Boolean);
  const relParts = cleanTarget.split('/').filter(Boolean);
  for (const segment of relParts) {
    if (segment === '.') {
      continue;
    }
    if (segment === '..') {
      parts.pop();
    } else {
      parts.push(segment);
    }
  }
  return parts.join('/');
}

function extractParagraphText(pElement: Element): string {
  const parts: string[] = [];

  const walk = (node: Element): void => {
    for (let i = 0; i < node.childNodes.length; i++) {
      const child = node.childNodes[i];
      if (child.nodeType !== Node.ELEMENT_NODE) {
        continue;
      }
      const el = child as Element;
      const local = el.localName.toLowerCase();
      if (local === 't') {
        parts.push(el.textContent ?? '');
      } else if (local === 'br') {
        parts.push('\n');
      } else if (local === 'tab') {
        parts.push(' ');
      } else if (local === 'r' || local === 'fld') {
        walk(el);
      }
    }
  };

  walk(pElement);
  return parts.join('').replace(/\r\n/g, '\n').trim();
}

function extractFirstRunStyle(txBody: Element): {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: 'left' | 'center' | 'right';
} {
  const style: {
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
    align?: 'left' | 'center' | 'right';
  } = {};

  const pPrList = getElementsByLocalName(txBody, 'pPr');
  for (const pPr of pPrList) {
    const algn = pPr.getAttribute('algn');
    if (algn === 'ctr') {
      style.align = 'center';
      break;
    }
    if (algn === 'r') {
      style.align = 'right';
      break;
    }
    if (algn === 'l' || algn === 'just') {
      style.align = 'left';
      break;
    }
  }

  const rPrList = getElementsByLocalName(txBody, 'rPr');
  for (const rPr of rPrList) {
    if (style.fontSize === undefined) {
      const szAttr = rPr.getAttribute('sz');
      if (szAttr) {
        const hundredths = Number.parseInt(szAttr, 10);
        if (Number.isFinite(hundredths) && hundredths > 0) {
          style.fontSize = Math.min(48, Math.max(12, Math.round(hundredths / 100)));
        }
      }
    }
    if (style.bold === undefined) {
      const bAttr = rPr.getAttribute('b');
      if (bAttr === '1' || bAttr === 'true') {
        style.bold = true;
      }
    }
    if (style.italic === undefined) {
      const iAttr = rPr.getAttribute('i');
      if (iAttr === '1' || iAttr === 'true') {
        style.italic = true;
      }
    }
    if (style.color === undefined) {
      const solidFill = getDirectChildrenByLocalName(rPr, 'solidFill')[0];
      if (solidFill) {
        const srgb = getElementsByLocalName(solidFill, 'srgbClr')[0];
        const hex = srgb?.getAttribute('val')?.trim();
        if (hex && /^[0-9a-fA-F]{6}$/.test(hex)) {
          style.color = `#${hex.toUpperCase()}`;
        }
      }
    }
  }

  return style;
}

function extractSlideBackgroundColor(xmlDoc: Document): string | undefined {
  const bgList = getElementsByLocalName(xmlDoc, 'bg');
  for (const bg of bgList) {
    const solidFill = getElementsByLocalName(bg, 'solidFill')[0];
    if (!solidFill) {
      continue;
    }
    const srgb = getElementsByLocalName(solidFill, 'srgbClr')[0];
    const val = srgb?.getAttribute('val')?.trim();
    if (val && /^[0-9a-fA-F]{6}$/.test(val)) {
      return `#${val.toUpperCase()}`;
    }
  }
  return undefined;
}

function getPlaceholderType(spElement: Element): string {
  const ph = getElementsByLocalName(spElement, 'ph')[0];
  return (ph?.getAttribute('type') ?? '').trim();
}

function extractNotesText(notesXml: string): string {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(notesXml, 'application/xml');
  const shapes = getElementsByLocalName(xmlDoc, 'sp');
  const ignoredTypes = new Set(['sldNum', 'sldImg', 'hdr', 'ftr', 'dt']);
  const lines: string[] = [];

  for (const sp of shapes) {
    const phType = getPlaceholderType(sp);
    if (phType && ignoredTypes.has(phType)) {
      continue;
    }
    const txBodies = getDirectChildrenByLocalName(sp, 'txBody');
    for (const txBody of txBodies) {
      const paragraphs = getDirectChildrenByLocalName(txBody, 'p');
      for (const p of paragraphs) {
        const text = extractParagraphText(p);
        if (text) {
          lines.push(text);
        }
      }
    }
  }

  return lines.join('\n').trim();
}

async function parseDocxFile(file: File): Promise<ParseFileResult> {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const sanitizedHtml = sanitizeWordHtml(result.value ?? '');
  const { htmlWithIds, outline, plainText } = extractOutlineFromHtml(sanitizedHtml);

  const title = stripFileExtension(file.name);
  const fileSize = formatFileSize(file.size);
  const wordHtml = htmlWithIds.trim() ? htmlWithIds : '<p></p>';
  const contentPreview = plainText
    ? truncatePreview(plainText, 260)
    : `Tệp Word thật tải lên từ thiết bị (${fileSize}).`;

  const document: DocumentItem = {
    id: createDocumentId('docx'),
    title,
    type: 'docx',
    updatedAt: 'Vừa tải lên',
    isFavorite: false,
    isShared: false,
    folder: 'Tải lên gần đây',
    fileSize,
    pageCount: Math.max(1, Math.ceil(plainText.length / 1800)),
    thumbnail: createThumbnailDataUri('docx', title, '#2563EB'),
    slidesPreview: [],
    outline,
    contentPreview,
    wordHtml,
    isUploaded: true,
    sourceFile: file,
  };

  return {
    document,
    createdObjectUrls: [],
  };
}

async function parseExcelFile(file: File, ext: string): Promise<ParseFileResult> {
  const arrayBuffer = await file.arrayBuffer();
  let workbook: XLSX.WorkBook;

  if (ext === '.csv') {
    const utf8Text = new TextDecoder('utf-8').decode(arrayBuffer);
    workbook = XLSX.read(utf8Text, { type: 'string' });
  } else {
    workbook = XLSX.read(arrayBuffer, { type: 'array' });
  }

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    return {
      error: 'Không tìm thấy trang tính hợp lệ trong tệp bảng tính này.',
      createdObjectUrls: [],
    };
  }

  const sheets: SpreadsheetSheet[] = workbook.SheetNames.map((sheetName, sheetIdx) => {
    const ws = workbook.Sheets[sheetName];
    let rawMatrix: unknown[][] = [];

    if (ws) {
      const refStr = ws['!ref'];
      if (refStr) {
        const decodedRef = XLSX.utils.decode_range(refStr);
        rawMatrix = XLSX.utils.sheet_to_json<unknown[]>(ws, {
          header: 1,
          defval: '',
          raw: true,
          range: {
            s: { r: 0, c: 0 },
            e: { r: decodedRef.e.r, c: decodedRef.e.c },
          },
        });
      } else {
        rawMatrix = XLSX.utils.sheet_to_json<unknown[]>(ws, {
          header: 1,
          defval: '',
          raw: true,
        });
      }
    }

    const rows: (string | number)[][] = rawMatrix.map((row) =>
      (Array.isArray(row) ? row : []).map((cell) => {
        if (typeof cell === 'number') {
          return Number.isFinite(cell) ? cell : '';
        }
        if (typeof cell === 'boolean') {
          return cell ? 'TRUE' : 'FALSE';
        }
        if (cell === null || cell === undefined) {
          return '';
        }
        return String(cell);
      })
    );

    // Overlay real formulas directly at exact [r][c] coordinates
    if (ws) {
      for (const cellKey of Object.keys(ws)) {
        if (cellKey.startsWith('!')) {
          continue;
        }
        const cellObj = ws[cellKey] as XLSX.CellObject | undefined;
        if (!cellObj || typeof cellObj.f !== 'string' || !cellObj.f.trim()) {
          continue;
        }
        const coords = XLSX.utils.decode_cell(cellKey);
        if (coords.r < 0 || coords.c < 0) {
          continue;
        }
        while (rows.length <= coords.r) {
          rows.push([]);
        }
        while (rows[coords.r].length <= coords.c) {
          rows[coords.r].push('');
        }
        const formulaText = cellObj.f.trim();
        rows[coords.r][coords.c] = formulaText.startsWith('=') ? formulaText : `=${formulaText}`;
      }
    }

    const maxCols = Math.max(6, ...rows.map((r) => r.length));
    if (rows.length === 0) {
      rows.push(Array.from({ length: maxCols }, () => ''));
    } else {
      for (let r = 0; r < rows.length; r++) {
        while (rows[r].length < maxCols) {
          rows[r].push('');
        }
      }
    }

    const columns = Array.from({ length: maxCols }, (_, colIdx) => columnIndexToLetter(colIdx));

    return {
      id: `uploaded-sheet-${sheetIdx + 1}-${Date.now()}`,
      name: sheetName || `Sheet ${sheetIdx + 1}`,
      columns,
      rows,
    };
  });

  const title = stripFileExtension(file.name);
  const fileSize = formatFileSize(file.size);
  const sheetSummary = sheets.map((s) => `${s.name} (${s.rows.length} dòng)`).join(', ');
  const contentPreview = truncatePreview(
    `Bảng tính gồm ${sheets.length} trang tính: ${sheetSummary}.`,
    260
  );

  const document: DocumentItem = {
    id: createDocumentId('xlsx'),
    title,
    type: 'xlsx',
    updatedAt: 'Vừa tải lên',
    isFavorite: false,
    isShared: false,
    folder: 'Tải lên gần đây',
    fileSize,
    pageCount: sheets.length,
    thumbnail: createThumbnailDataUri('xlsx', title, '#15803D'),
    slidesPreview: [],
    outline: [],
    sheets,
    contentPreview,
    isUploaded: true,
    sourceFile: file,
  };

  return {
    document,
    createdObjectUrls: [],
  };
}

function getRelationshipIdAttr(el: Element): string {
  const direct = el.getAttribute('r:id');
  if (direct) {
    return direct.trim();
  }
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    if (
      attr.localName === 'id' &&
      (attr.namespaceURI ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ||
        attr.name.endsWith(':id'))
    ) {
      return attr.value.trim();
    }
  }
  return '';
}

async function resolveOrderedPptxSlideEntries(
  zip: JSZip,
  xmlParser: DOMParser
): Promise<Array<{ path: string; num: number }>> {
  const ordered: Array<{ path: string; num: number }> = [];
  const seenPathsLower = new Set<string>();

  const presFile = findZipFileInsensitive(zip, 'ppt/presentation.xml');
  const presRelsFile = findZipFileInsensitive(zip, 'ppt/_rels/presentation.xml.rels');

  if (presFile && presRelsFile) {
    try {
      const [presXml, presRelsXml] = await Promise.all([
        presFile.async('string'),
        presRelsFile.async('string'),
      ]);

      const relsDoc = xmlParser.parseFromString(presRelsXml, 'application/xml');
      const relTargetById = new Map<string, string>();
      for (const rel of getElementsByLocalName(relsDoc, 'Relationship')) {
        const id = (rel.getAttribute('Id') ?? '').trim();
        const target = (rel.getAttribute('Target') ?? '').trim();
        const targetMode = (rel.getAttribute('TargetMode') ?? '').trim();
        if (id && target && targetMode.toLowerCase() !== 'external') {
          relTargetById.set(id, target);
        }
      }

      const presDoc = xmlParser.parseFromString(presXml, 'application/xml');
      const sldIdElements = getElementsByLocalName(presDoc, 'sldId');

      for (const sldIdEl of sldIdElements) {
        const rId = getRelationshipIdAttr(sldIdEl);
        if (!rId) {
          continue;
        }
        const target = relTargetById.get(rId);
        if (!target) {
          continue;
        }
        const resolvedPath = resolveZipPath('ppt', target);
        const zipObj = findZipFileInsensitive(zip, resolvedPath);
        if (!zipObj || zipObj.dir) {
          continue;
        }
        const canonicalPath = zipObj.name;
        const lowerKey = canonicalPath.toLowerCase();
        if (seenPathsLower.has(lowerKey)) {
          continue;
        }
        seenPathsLower.add(lowerKey);
        const numMatch = /slide(\d+)\.xml$/i.exec(canonicalPath);
        const num = numMatch ? Number.parseInt(numMatch[1], 10) : ordered.length + 1;
        ordered.push({ path: canonicalPath, num });
      }
    } catch {
      // Fallback cleanly to numeric slide scan below
    }
  }

  // Fallback / append any remaining slides outside sldIdLst sorted numerically
  const fallbackEntries: Array<{ path: string; num: number }> = [];
  for (const filePath of Object.keys(zip.files)) {
    const match = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(filePath);
    if (match && !zip.files[filePath].dir && !seenPathsLower.has(filePath.toLowerCase())) {
      fallbackEntries.push({
        path: filePath,
        num: Number.parseInt(match[1], 10),
      });
    }
  }
  fallbackEntries.sort((a, b) => a.num - b.num);
  ordered.push(...fallbackEntries);

  return ordered;
}

async function parsePptxFile(file: File): Promise<ParseFileResult> {
  const createdObjectUrls: string[] = [];

  try {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const xmlParser = new DOMParser();

    const slideEntries = await resolveOrderedPptxSlideEntries(zip, xmlParser);

    if (slideEntries.length === 0) {
      return {
        error: 'Không tìm thấy trang slide hợp lệ bên trong tệp .pptx này.',
        createdObjectUrls: [],
      };
    }

    const slidesPreview: SlidePreviewItem[] = [];
    const ignoredShapePh = new Set(['sldNum', 'dt', 'ftr']);

    for (let i = 0; i < slideEntries.length; i++) {
      const { path: slidePath, num: slideNum } = slideEntries[i];
      const slideFile = zip.file(slidePath);
      if (!slideFile) {
        continue;
      }

      const xmlText = await slideFile.async('string');
      const xmlDoc = xmlParser.parseFromString(xmlText, 'application/xml');
      const explicitBackgroundColor = extractSlideBackgroundColor(xmlDoc);

      const blocks: SlideTextBlock[] = [];
      let blockCounter = 0;

      // 1. Parse all text shapes (sp)
      const shapes = getElementsByLocalName(xmlDoc, 'sp');
      for (const sp of shapes) {
        const phType = getPlaceholderType(sp);
        if (phType && ignoredShapePh.has(phType)) {
          continue;
        }

        const txBodies = getDirectChildrenByLocalName(sp, 'txBody');
        const paragraphsText: string[] = [];
        let runStyle: ReturnType<typeof extractFirstRunStyle> = {};

        for (const txBody of txBodies) {
          if (Object.keys(runStyle).length === 0) {
            runStyle = extractFirstRunStyle(txBody);
          }
          const paragraphs = getDirectChildrenByLocalName(txBody, 'p');
          for (const p of paragraphs) {
            const pText = extractParagraphText(p);
            if (pText) {
              paragraphsText.push(pText);
            }
          }
        }

        const combinedText = paragraphsText.join('\n').trim();
        if (!combinedText) {
          continue;
        }

        let role: 'title' | 'subtitle' | 'body' = 'body';
        let defaultFontSize = 16;
        let defaultBold = false;

        if (phType === 'title' || phType === 'ctrTitle') {
          role = 'title';
          defaultFontSize = 32;
          defaultBold = true;
        } else if (phType === 'subTitle') {
          role = 'subtitle';
          defaultFontSize = 20;
        }

        blockCounter++;
        blocks.push({
          id: `slide-${slideNum}-block-${blockCounter}`,
          text: combinedText,
          fontSize: runStyle.fontSize ?? defaultFontSize,
          bold: runStyle.bold ?? defaultBold,
          italic: runStyle.italic,
          color: runStyle.color,
          align: runStyle.align ?? (phType === 'ctrTitle' ? 'center' : 'left'),
          role,
        });
      }

      // 2. Also extract text from embedded PowerPoint tables (a:tbl) if present
      const tables = getElementsByLocalName(xmlDoc, 'tbl');
      for (const tbl of tables) {
        const rows = getDirectChildrenByLocalName(tbl, 'tr');
        const tableLines: string[] = [];
        for (const tr of rows) {
          const cells = getDirectChildrenByLocalName(tr, 'tc');
          const cellTexts = cells
            .map((tc) => {
              const paras = getElementsByLocalName(tc, 'p');
              return paras
                .map((p) => extractParagraphText(p))
                .filter(Boolean)
                .join(' ');
            })
            .filter(Boolean);
          if (cellTexts.length > 0) {
            tableLines.push(cellTexts.join(' | '));
          }
        }
        const tableCombined = tableLines.join('\n').trim();
        if (tableCombined) {
          blockCounter++;
          blocks.push({
            id: `slide-${slideNum}-block-${blockCounter}`,
            text: tableCombined,
            fontSize: 15,
            role: 'body',
            align: 'left',
          });
        }
      }

      // 3. Parse slide relationships for real images & speaker notes
      const images: SlideImageItem[] = [];
      let notesPath: string | null = null;
      const seenImagePaths = new Set<string>();

      const relsFile = findZipFileInsensitive(zip, `ppt/slides/_rels/slide${slideNum}.xml.rels`);
      if (relsFile) {
        const relsXml = await relsFile.async('string');
        const relsDoc = xmlParser.parseFromString(relsXml, 'application/xml');
        const relationships = getElementsByLocalName(relsDoc, 'Relationship');

        for (const rel of relationships) {
          const relType = rel.getAttribute('Type') ?? '';
          const relTarget = rel.getAttribute('Target') ?? '';
          const targetMode = rel.getAttribute('TargetMode') ?? '';
          if (!relTarget || targetMode.toLowerCase() === 'external') {
            continue;
          }

          const resolvedPath = resolveZipPath('ppt/slides', relTarget);

          if (relType.includes('/image')) {
            const pathKey = resolvedPath.toLowerCase();
            if (seenImagePaths.has(pathKey)) {
              continue;
            }
            const extMatch = /\.([a-z0-9]+)$/i.exec(resolvedPath);
            const imgExt = extMatch ? extMatch[1].toLowerCase() : '';
            const mime = IMAGE_MIME_BY_EXT[imgExt];
            if (!mime) {
              continue;
            }
            const imgZipObj = findZipFileInsensitive(zip, resolvedPath);
            if (!imgZipObj) {
              continue;
            }
            seenImagePaths.add(pathKey);
            const imgBuffer = await imgZipObj.async('arraybuffer');
            const blob = new Blob([imgBuffer], { type: mime });
            const blobUrl = URL.createObjectURL(blob);
            createdObjectUrls.push(blobUrl);
            const fileName = resolvedPath.split('/').pop() ?? `image-${images.length + 1}`;
            images.push({
              id: `slide-${slideNum}-img-${images.length + 1}`,
              src: blobUrl,
              alt: `Ảnh slide ${i + 1} (${fileName})`,
            });
          } else if (relType.includes('/notesSlide') && !notesPath) {
            notesPath = resolvedPath;
          }
        }
      }

      // 4. Read real speaker notes from rel target or fallback path
      let speakerNote: string | undefined;
      const notesZipObj =
        (notesPath ? findZipFileInsensitive(zip, notesPath) : null) ??
        findZipFileInsensitive(zip, `ppt/notesSlides/notesSlide${slideNum}.xml`);
      if (notesZipObj) {
        const notesXml = await notesZipObj.async('string');
        const extractedNote = extractNotesText(notesXml);
        if (extractedNote) {
          speakerNote = extractedNote;
        }
      }

      // 5. Derive clean title, subtitle, and bullets from real blocks
      const titleBlock = blocks.find((b) => b.role === 'title');
      const subtitleBlock = blocks.find((b) => b.role === 'subtitle');

      let slideTitle = '';
      if (titleBlock) {
        slideTitle = titleBlock.text.split('\n')[0].trim();
      } else if (blocks.length > 0) {
        slideTitle = blocks[0].text.split('\n')[0].trim();
      }
      if (!slideTitle) {
        slideTitle = `Slide ${i + 1}`;
      }

      const subtitle = subtitleBlock ? subtitleBlock.text.replace(/\n+/g, ' ').trim() : undefined;

      const bullets: string[] = [];
      for (const block of blocks) {
        if (block === titleBlock || block === subtitleBlock) {
          continue;
        }
        const lines = block.text
          .split('\n')
          .map((line) => line.trim())
          .filter(Boolean);
        for (const line of lines) {
          if (!titleBlock && block === blocks[0] && line === slideTitle) {
            continue;
          }
          bullets.push(line);
        }
      }

      // Infer white background when slide inherits default light Slide Master (no explicit <p:bg>)
      // while text blocks carry dark colors, so dark text never sinks into the dark stage.
      let backgroundColor = explicitBackgroundColor;
      if (!backgroundColor) {
        const luminances: number[] = [];
        for (const block of blocks) {
          const hexMatch = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.exec(block.color ?? '');
          if (hexMatch) {
            const rawHex = hexMatch[1];
            const hex =
              rawHex.length === 3
                ? `${rawHex[0]}${rawHex[0]}${rawHex[1]}${rawHex[1]}${rawHex[2]}${rawHex[2]}`
                : rawHex;
            const r = Number.parseInt(hex.slice(0, 2), 16) / 255;
            const g = Number.parseInt(hex.slice(2, 4), 16) / 255;
            const b = Number.parseInt(hex.slice(4, 6), 16) / 255;
            luminances.push(0.2126 * r + 0.7152 * g + 0.0722 * b);
          }
        }
        if (luminances.length > 0) {
          const avgLuminance =
            luminances.reduce((sum, val) => sum + val, 0) / luminances.length;
          if (avgLuminance < 0.45) {
            backgroundColor = '#FFFFFF';
          }
        }
      }

      slidesPreview.push({
        id: `uploaded-slide-${slideNum}-${Date.now()}`,
        index: i + 1,
        title: slideTitle,
        subtitle,
        bullets,
        blocks,
        images,
        speakerNote,
        accentColor: '#4F46E5',
        backgroundColor,
      });
    }

    const outline: OutlineSection[] = slidesPreview.map((slide) => ({
      id: slide.id,
      romanIndex: toRomanNumeral(slide.index),
      title: slide.title,
      summary: slide.subtitle ?? slide.bullets[0] ?? '',
      keyPoints: slide.bullets.slice(0, 3),
      pageNumber: slide.index,
    }));

    const firstSlide = slidesPreview[0];
    const firstSlideTextParts = firstSlide
      ? [firstSlide.title, firstSlide.subtitle ?? '', ...firstSlide.bullets].filter(Boolean)
      : [];
    const title = stripFileExtension(file.name);
    const fileSize = formatFileSize(file.size);
    const contentPreview =
      firstSlideTextParts.length > 0
        ? truncatePreview(firstSlideTextParts.join(' · '), 260)
        : `Tệp trình chiếu gồm ${slidesPreview.length} slide (${fileSize}).`;

    const document: DocumentItem = {
      id: createDocumentId('pptx'),
      title,
      type: 'pptx',
      updatedAt: 'Vừa tải lên',
      isFavorite: false,
      isShared: false,
      folder: 'Tải lên gần đây',
      fileSize,
      pageCount: slidesPreview.length,
      thumbnail: createThumbnailDataUri('pptx', title, '#4F46E5'),
      slidesPreview,
      outline,
      contentPreview,
      isUploaded: true,
      sourceFile: file,
    };

    return {
      document,
      createdObjectUrls,
    };
  } catch (err) {
    for (const url of createdObjectUrls) {
      URL.revokeObjectURL(url);
    }
    throw err;
  }
}

async function parsePdfFile(file: File): Promise<ParseFileResult> {
  const fileUrl = URL.createObjectURL(file);
  const title = stripFileExtension(file.name);
  const fileSize = formatFileSize(file.size);

  let pageCount: number | undefined;
  let outline: OutlineSection[] = [];
  let contentPreview = `Tệp PDF thật tải lên từ thiết bị (${fileSize}).`;

  try {
    const meta = await inspectPdfMetadata(file);
    if (meta.pageCount > 0) {
      pageCount = meta.pageCount;
    }
    if (meta.outline.length > 0) {
      outline = meta.outline;
    }
    if (meta.contentPreview) {
      contentPreview = meta.contentPreview;
    } else if (meta.pageCount > 0) {
      contentPreview = `Tệp PDF thật gồm ${meta.pageCount} trang tải lên từ thiết bị (${fileSize}).`;
    }
  } catch {
    // Keep safe fallback metadata if PDF is scanned/encrypted or fails quick inspection
  }

  const document: DocumentItem = {
    id: createDocumentId('pdf'),
    title,
    type: 'pdf',
    updatedAt: 'Vừa tải lên',
    isFavorite: false,
    isShared: false,
    folder: 'Tải lên gần đây',
    fileSize,
    pageCount,
    thumbnail: createThumbnailDataUri('pdf', title, '#DB2777'),
    fileUrl,
    objectUrl: fileUrl,
    pdfPages: [],
    outline,
    slidesPreview: [],
    contentPreview,
    isUploaded: true,
    sourceFile: file,
  };

  return {
    document,
    createdObjectUrls: [fileUrl],
  };
}

export async function parseUploadedFile(file: File): Promise<ParseFileResult> {
  const ext = getFileExtension(file.name);

  if (ext === '.doc') {
    return {
      error:
        'Định dạng .doc cũ chưa hỗ trợ đọc trực tiếp trên trình duyệt. Vui lòng mở bằng Word / Google Docs và chọn "Lưu dưới dạng (.docx)" rồi tải lại.',
      createdObjectUrls: [],
    };
  }

  try {
    if (ext === '.docx') {
      return await parseDocxFile(file);
    }
    if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      return await parseExcelFile(file, ext);
    }
    if (ext === '.pptx') {
      return await parsePptxFile(file);
    }
    if (ext === '.pdf') {
      return await parsePdfFile(file);
    }

    return {
      error: `Định dạng tệp "${ext || file.name}" chưa được hỗ trợ. Vui lòng chọn tệp PPTX, DOCX, XLSX, XLS, CSV hoặc PDF.`,
      createdObjectUrls: [],
    };
  } catch (err) {
    const detail = err instanceof Error && err.message ? ` (${err.message})` : '';
    return {
      error: `Không thể đọc nội dung tệp "${file.name}". Tệp có thể bị hỏng hoặc đang được đặt mật khẩu bảo vệ${detail}.`,
      createdObjectUrls: [],
    };
  }
}

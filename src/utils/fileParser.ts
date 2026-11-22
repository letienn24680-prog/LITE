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

const COVER_CONTAINS_BLACKLIST = [
  'BỘ GIÁO DỤC',
  'TRƯỜNG ĐẠI HỌC',
  'HỌC VIỆN',
  'CỘNG HÒA XÃ HỘI',
  'ĐỘC LẬP - TỰ DO',
  'ĐỘC LẬP – TỰ DO',
  'HỌ VÀ TÊN',
  'SINH VIÊN',
  'HỌC VIÊN',
  'GIẢNG VIÊN',
  'NGƯỜI HƯỚNG DẪN',
  'CÁN BỘ HƯỚNG DẪN',
  'GVHD',
  'MÃ SỐ',
  'MSSV',
  'MSHV',
  'CHUYÊN NGÀNH',
  'NGÀNH HỌC',
  'NĂM HỌC',
  'NIÊN KHÓA',
  'HỆ ĐÀO TẠO',
  'BỘ MÔN',
];

const COVER_PREFIX_BLACKLIST = [
  'TRƯỜNG ',
  'TRƯỜNG:',
  'KHOA ',
  'KHOA:',
  'VIỆN ',
  'VIỆN:',
  'LỚP ',
  'LỚP:',
  'KHÓA ',
  'KHÓA:',
  'ĐỀ TÀI',
  'TÊN ĐỀ TÀI',
  'TIỂU LUẬN',
];

const COVER_DOC_KIND_KEYWORDS = ['ĐỒ ÁN', 'LUẬN VĂN', 'KHÓA LUẬN', 'BÁO CÁO THỰC TẬP'];

const STRUCTURAL_START_KEYWORDS = [
  'LỜI CAM ĐOAN',
  'LỜI CẢM ƠN',
  'TÓM TẮT',
  'MỤC LỤC',
  'DANH MỤC',
  'MỞ ĐẦU',
  'LỜI MỞ ĐẦU',
  'CHƯƠNG ',
  'CHƯƠNG:',
  'CHAPTER ',
  'PHẦN ',
  'PHẦN:',
  'PART ',
  'KẾT LUẬN',
  'KIẾN NGHỊ',
  'TÀI LIỆU THAM KHẢO',
  'PHỤ LỤC',
];

const LOCATION_YEAR_REGEX =
  /^(?:HÀ NỘI|TP\.?\s*HỒ CHÍ MINH|THÀNH PHỐ HỒ CHÍ MINH|ĐÀ NẴNG|HẢI PHÒNG|CẦN THƠ|HUẾ|THÁNG\s+\d{1,2})(?:[\s,\-–/]+.*?(?:19|20)\d{2}|\s*)$/i;

const YEAR_FOOTER_REGEX = /,\s*(?:THÁNG\s+\d{1,2}\s+)?NĂM\s+(?:19|20)\d{2}\s*$/i;
const DOTTED_TOC_LINE_REGEX = /\.{4,}\s*\d*\s*$/;
const ROMAN_HEADING_REGEX =
  /^(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV)\.\s+\S/;
const NUMBERED_DOT_HEADING_REGEX = /^\d+(?:\.\d+){0,3}\.\s+\S/;
const MULTI_LEVEL_NUMBER_HEADING_REGEX = /^\d+\.\d+(?:\.\d+){0,2}\s+\S/;

function matchesStructuralOutlinePattern(clean: string, upper: string): boolean {
  for (const kw of STRUCTURAL_START_KEYWORDS) {
    if (upper === kw.trim() || upper.startsWith(kw)) {
      return true;
    }
  }
  if (ROMAN_HEADING_REGEX.test(clean)) {
    return true;
  }
  if (NUMBERED_DOT_HEADING_REGEX.test(clean) || MULTI_LEVEL_NUMBER_HEADING_REGEX.test(clean)) {
    // Reject long sentence-like paragraphs that happen to start with a number
    if (clean.length <= 120 && !/[,;]$/.test(clean)) {
      return true;
    }
  }
  return false;
}

export function isLikelyStructuralOutlineTitle(
  text: string,
  isExplicitHeadingTag: boolean
): boolean {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (clean.length < 2 || clean.length > 140) {
    return false;
  }

  // Reject lines without any letter (pure numbers, dates, separator lines)
  if (!/[A-Za-zÀ-ỹ]/u.test(clean)) {
    return false;
  }

  // Reject printed table-of-contents entries containing dotted leaders ".... 12"
  if (DOTTED_TOC_LINE_REGEX.test(clean)) {
    return false;
  }

  const upper = clean.toUpperCase();

  // Reject administrative / cover-page lines even if bolded or styled as Heading
  for (const banned of COVER_CONTAINS_BLACKLIST) {
    if (upper.includes(banned)) {
      return false;
    }
  }

  for (const prefix of COVER_PREFIX_BLACKLIST) {
    if (upper === prefix.trim() || upper.startsWith(prefix)) {
      return false;
    }
  }

  if (LOCATION_YEAR_REGEX.test(clean) || YEAR_FOOTER_REGEX.test(clean)) {
    return false;
  }

  const isStructuralMatch = matchesStructuralOutlinePattern(clean, upper);

  // Reject cover-page document type headers (e.g. "ĐỒ ÁN TỐT NGHIỆP", "LUẬN VĂN THẠC SĨ")
  // unless prefixed by a true chapter/section number
  if (!isStructuralMatch) {
    for (const kindKw of COVER_DOC_KIND_KEYWORDS) {
      if (upper === kindKw || upper.startsWith(`${kindKw} `) || (upper.includes(kindKw) && clean.length <= 90)) {
        return false;
      }
    }
  }

  if (isStructuralMatch) {
    return true;
  }

  return isExplicitHeadingTag;
}

function inferStructuralHeadingLevel(el: Element, title: string): number {
  const tag = el.tagName.toUpperCase();
  if (tag === 'H1') {
    return 1;
  }
  if (tag === 'H2') {
    return 2;
  }
  if (tag === 'H3') {
    return 3;
  }

  const numbered = /^(\d+(?:\.\d+){0,3})\.?\s+/.exec(title);
  if (numbered) {
    const depth = numbered[1].split('.').length;
    return Math.min(3, Math.max(1, depth));
  }
  return 1;
}

export function extractOutlineFromHtml(html: string): {
  htmlWithIds: string;
  outline: OutlineSection[];
  plainText: string;
} {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(html || '', 'text/html');

  const explicitHeadingElements = Array.from(parsed.body.querySelectorAll('h1, h2, h3')).filter(
    (el) => {
      const text = (el.textContent ?? '').replace(/\s+/g, ' ').trim();
      return isLikelyStructuralOutlineTitle(text, true);
    }
  );

  let validHeadings: Element[];
  if (explicitHeadingElements.length >= 2) {
    validHeadings = explicitHeadingElements;
  } else {
    // Fallback: include structural <p> headings (e.g. CHƯƠNG 1, 1.1., LỜI MỞ ĐẦU) outside tables/lists
    const candidates = Array.from(parsed.body.querySelectorAll('h1, h2, h3, p'));
    validHeadings = candidates.filter((el) => {
      const tag = el.tagName.toUpperCase();
      const isExplicit = tag === 'H1' || tag === 'H2' || tag === 'H3';
      if (!isExplicit && el.closest('td, th, li')) {
        return false;
      }
      const text = (el.textContent ?? '').replace(/\s+/g, ' ').trim();
      return isLikelyStructuralOutlineTitle(text, isExplicit);
    });
  }

  const levels = validHeadings.map((el) => {
    const title = (el.textContent ?? '').replace(/\s+/g, ' ').trim();
    return inferStructuralHeadingLevel(el, title);
  });
  const minLevel = levels.length > 0 ? Math.min(...levels) : 1;

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

    const level = levels[i];
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

interface PptxThemePalette {
  colors: Record<string, string>;
  majorFont?: string;
  minorFont?: string;
}

const DEFAULT_PPTX_THEME_COLORS: Record<string, string> = {
  dk1: '#000000',
  lt1: '#FFFFFF',
  dk2: '#1F2937',
  lt2: '#F1F5F9',
  accent1: '#4F46E5',
  accent2: '#EA580C',
  accent3: '#64748B',
  accent4: '#F59E0B',
  accent5: '#0EA5E9',
  accent6: '#10B981',
  hlink: '#2563EB',
  folHlink: '#7C3AED',
  tx1: '#000000',
  bg1: '#FFFFFF',
  tx2: '#1F2937',
  bg2: '#F1F5F9',
};

interface EmuBox {
  x: number;
  y: number;
  cx: number;
  cy: number;
}

interface GroupTransform {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
}

const IDENTITY_GROUP_TRANSFORM: GroupTransform = {
  offsetX: 0,
  offsetY: 0,
  scaleX: 1,
  scaleY: 1,
};

function clampByte(n: number): number {
  return Math.min(255, Math.max(0, Math.round(n)));
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${[clampByte(r), clampByte(g), clampByte(b)]
    .map((v) => v.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()}`;
}

function hexToRgb(hex: string): [number, number, number] | null {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex.trim());
  if (!m) {
    return null;
  }
  const h = m[1];
  return [
    Number.parseInt(h.slice(0, 2), 16),
    Number.parseInt(h.slice(2, 4), 16),
    Number.parseInt(h.slice(4, 6), 16),
  ];
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) {
    return [0, 0, l];
  }
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  if (max === rn) {
    h = (gn - bn) / d + (gn < bn ? 6 : 0);
  } else if (max === gn) {
    h = (bn - rn) / d + 2;
  } else {
    h = (rn - gn) / d + 4;
  }
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) {
    const v = clampByte(l * 255);
    return [v, v, v];
  }
  const hue2rgb = (p: number, q: number, tIn: number): number => {
    let t = tIn;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [
    clampByte(hue2rgb(p, q, h + 1 / 3) * 255),
    clampByte(hue2rgb(p, q, h) * 255),
    clampByte(hue2rgb(p, q, h - 1 / 3) * 255),
  ];
}

function applyDrawingColorModifiers(baseHex: string, colorEl: Element): string {
  const rgb = hexToRgb(baseHex);
  if (!rgb) {
    return baseHex;
  }
  let [r, g, b] = rgb;
  let lumMod: number | null = null;
  let lumOff: number | null = null;

  for (let i = 0; i < colorEl.childNodes.length; i++) {
    const child = colorEl.childNodes[i];
    if (child.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const modEl = child as Element;
    const local = modEl.localName.toLowerCase();
    const rawVal = Number.parseInt(modEl.getAttribute('val') ?? '', 10);
    if (!Number.isFinite(rawVal)) {
      continue;
    }
    const factor = rawVal / 100000;
    if (local === 'tint') {
      r = clampByte(r * factor + 255 * (1 - factor));
      g = clampByte(g * factor + 255 * (1 - factor));
      b = clampByte(b * factor + 255 * (1 - factor));
    } else if (local === 'shade') {
      r = clampByte(r * factor);
      g = clampByte(g * factor);
      b = clampByte(b * factor);
    } else if (local === 'lummod') {
      lumMod = factor;
    } else if (local === 'lumoff') {
      lumOff = factor;
    }
  }

  if (lumMod !== null || lumOff !== null) {
    const [h, s, l] = rgbToHsl(r, g, b);
    const nextL = Math.min(1, Math.max(0, l * (lumMod ?? 1) + (lumOff ?? 0)));
    [r, g, b] = hslToRgb(h, s, nextL);
  }

  return rgbToHex(r, g, b);
}

function resolveDrawingColor(
  container: Element | undefined,
  themeColors: Record<string, string>
): string | undefined {
  if (!container) {
    return undefined;
  }

  const srgb = getDirectChildrenByLocalName(container, 'srgbClr')[0];
  if (srgb) {
    const val = srgb.getAttribute('val')?.trim();
    if (val && /^[0-9a-fA-F]{6}$/.test(val)) {
      return applyDrawingColorModifiers(`#${val.toUpperCase()}`, srgb);
    }
  }

  const scheme = getDirectChildrenByLocalName(container, 'schemeClr')[0];
  if (scheme) {
    const key = (scheme.getAttribute('val') ?? '').trim();
    const mapped = themeColors[key] ?? themeColors[key.toLowerCase()];
    if (mapped) {
      return applyDrawingColorModifiers(mapped, scheme);
    }
  }

  const sys = getDirectChildrenByLocalName(container, 'sysClr')[0];
  if (sys) {
    const lastClr = sys.getAttribute('lastClr')?.trim();
    if (lastClr && /^[0-9a-fA-F]{6}$/.test(lastClr)) {
      return applyDrawingColorModifiers(`#${lastClr.toUpperCase()}`, sys);
    }
  }

  return undefined;
}

async function parsePptxThemePalette(
  zip: JSZip,
  xmlParser: DOMParser
): Promise<PptxThemePalette> {
  const colors: Record<string, string> = { ...DEFAULT_PPTX_THEME_COLORS };
  let majorFont: string | undefined;
  let minorFont: string | undefined;

  const themeFile = findZipFileInsensitive(zip, 'ppt/theme/theme1.xml');
  if (!themeFile) {
    return { colors, majorFont, minorFont };
  }

  try {
    const themeXml = await themeFile.async('string');
    const themeDoc = xmlParser.parseFromString(themeXml, 'application/xml');

    const clrScheme = getElementsByLocalName(themeDoc, 'clrScheme')[0];
    if (clrScheme) {
      const slotNames = [
        'dk1',
        'lt1',
        'dk2',
        'lt2',
        'accent1',
        'accent2',
        'accent3',
        'accent4',
        'accent5',
        'accent6',
        'hlink',
        'folHlink',
      ];
      for (const slot of slotNames) {
        const slotEl = getDirectChildrenByLocalName(clrScheme, slot)[0];
        if (!slotEl) {
          continue;
        }
        const resolved = resolveDrawingColor(slotEl, colors);
        if (resolved) {
          colors[slot] = resolved;
          colors[slot.toLowerCase()] = resolved;
        }
      }
      // Map OpenXML presentation theme aliases
      colors.tx1 = colors.dk1;
      colors.bg1 = colors.lt1;
      colors.tx2 = colors.dk2;
      colors.bg2 = colors.lt2;
    }

    const fontScheme = getElementsByLocalName(themeDoc, 'fontScheme')[0];
    if (fontScheme) {
      const majorEl = getDirectChildrenByLocalName(fontScheme, 'majorFont')[0];
      const minorEl = getDirectChildrenByLocalName(fontScheme, 'minorFont')[0];
      const majorLatin = majorEl ? getDirectChildrenByLocalName(majorEl, 'latin')[0] : undefined;
      const minorLatin = minorEl ? getDirectChildrenByLocalName(minorEl, 'latin')[0] : undefined;
      const mj = majorLatin?.getAttribute('typeface')?.trim();
      const mn = minorLatin?.getAttribute('typeface')?.trim();
      if (mj) majorFont = mj;
      if (mn) minorFont = mn;
    }
  } catch {
    // Keep default palette on malformed theme XML
  }

  return { colors, majorFont, minorFont };
}

function withSansFallback(fontName: string | undefined): string {
  const base = (fontName ?? "").trim();
  const fallback = "'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif";
  if (!base) return fallback;
  if (/serif|mono|times|georgia|cambria|consolas|courier/i.test(base)) return base;
  return `"${base.replace(/"/g, "")}", ${fallback}`;
}

function resolveThemeFontName(
  rawTypeface: string | null | undefined,
  theme: PptxThemePalette
): string | undefined {
  const tf = (rawTypeface ?? '').trim();
  if (!tf) {
    return withSansFallback(theme.minorFont);
  }
  if (tf.startsWith('+mj-')) {
    return withSansFallback(theme.majorFont);
  }
  if (tf.startsWith('+mn-')) {
    return withSansFallback(theme.minorFont);
  }
  return withSansFallback(tf);
}

function getRelationshipAttr(el: Element, attrLocalName: string): string {
  const direct = el.getAttribute(`r:${attrLocalName}`);
  if (direct) {
    return direct.trim();
  }
  const attrs = el.attributes;
  const targetLower = attrLocalName.toLowerCase();
  for (let i = 0; i < attrs.length; i++) {
    const attr = attrs[i];
    if (
      attr.localName.toLowerCase() === targetLower &&
      (attr.namespaceURI ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ||
        attr.name.toLowerCase().endsWith(`:${targetLower}`))
    ) {
      return attr.value.trim();
    }
  }
  return '';
}

function extractEmuBoxFromXfrm(xfrmEl: Element | undefined): EmuBox | null {
  if (!xfrmEl) {
    return null;
  }
  const off = getDirectChildrenByLocalName(xfrmEl, 'off')[0];
  const ext = getDirectChildrenByLocalName(xfrmEl, 'ext')[0];
  if (!off || !ext) {
    return null;
  }
  const x = Number.parseInt(off.getAttribute('x') ?? '', 10);
  const y = Number.parseInt(off.getAttribute('y') ?? '', 10);
  const cx = Number.parseInt(ext.getAttribute('cx') ?? '', 10);
  const cy = Number.parseInt(ext.getAttribute('cy') ?? '', 10);
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(cx) ||
    !Number.isFinite(cy) ||
    cx <= 0 ||
    cy <= 0
  ) {
    return null;
  }
  return { x, y, cx, cy };
}

function composeGroupTransform(
  parent: GroupTransform,
  grpXfrmEl: Element | undefined
): GroupTransform {
  if (!grpXfrmEl) {
    return parent;
  }
  const outer = extractEmuBoxFromXfrm(grpXfrmEl);
  if (!outer) {
    return parent;
  }
  const chOff = getDirectChildrenByLocalName(grpXfrmEl, 'chOff')[0];
  const chExt = getDirectChildrenByLocalName(grpXfrmEl, 'chExt')[0];
  const chX = Number.parseInt(chOff?.getAttribute('x') ?? '0', 10) || 0;
  const chY = Number.parseInt(chOff?.getAttribute('y') ?? '0', 10) || 0;
  const chCx = Number.parseInt(chExt?.getAttribute('cx') ?? '0', 10) || outer.cx;
  const chCy = Number.parseInt(chExt?.getAttribute('cy') ?? '0', 10) || outer.cy;

  const localScaleX = chCx > 0 ? outer.cx / chCx : 1;
  const localScaleY = chCy > 0 ? outer.cy / chCy : 1;

  return {
    offsetX: parent.offsetX + (outer.x - chX * localScaleX) * parent.scaleX,
    offsetY: parent.offsetY + (outer.y - chY * localScaleY) * parent.scaleY,
    scaleX: parent.scaleX * localScaleX,
    scaleY: parent.scaleY * localScaleY,
  };
}

function mapBoxToPct(
  localBox: EmuBox,
  groupTransform: GroupTransform,
  slideCx: number,
  slideCy: number
): { xPct: number; yPct: number; wPct: number; hPct: number } {
  const absX = groupTransform.offsetX + localBox.x * groupTransform.scaleX;
  const absY = groupTransform.offsetY + localBox.y * groupTransform.scaleY;
  const absCx = localBox.cx * groupTransform.scaleX;
  const absCy = localBox.cy * groupTransform.scaleY;

  const round2 = (v: number): number => Math.round(v * 100) / 100;
  const xPct = round2(Math.min(99, Math.max(0, (absX / slideCx) * 100)));
  const yPct = round2(Math.min(99, Math.max(0, (absY / slideCy) * 100)));
  const wPct = round2(Math.min(100 - xPct, Math.max(1, (absCx / slideCx) * 100)));
  const hPct = round2(Math.min(100 - yPct, Math.max(1, (absCy / slideCy) * 100)));
  return { xPct, yPct, wPct, hPct };
}

function getDefaultPlaceholderPct(
  phType: string,
  fallbackIndex: number
): { xPct: number; yPct: number; wPct: number; hPct: number } {
  if (phType === 'title' || phType === 'ctrTitle') {
    return { xPct: 6, yPct: 6, wPct: 88, hPct: 16 };
  }
  if (phType === 'subTitle') {
    return { xPct: 8, yPct: 24, wPct: 84, hPct: 14 };
  }
  const top = Math.min(72, 24 + fallbackIndex * 16);
  return { xPct: 6, yPct: top, wPct: 88, hPct: Math.max(18, 92 - top) };
}

function extractFirstRunStyle(
  txBody: Element,
  theme: PptxThemePalette
): {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: 'left' | 'center' | 'right';
  fontFamily?: string;
  valign?: 'top' | 'middle' | 'bottom';
} {
  const style: {
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
    align?: 'left' | 'center' | 'right';
    fontFamily?: string;
    valign?: 'top' | 'middle' | 'bottom';
  } = {};

  const bodyPr = getDirectChildrenByLocalName(txBody, 'bodyPr')[0];
  const anchor = bodyPr?.getAttribute('anchor')?.trim();
  if (anchor === 't') {
    style.valign = 'top';
  } else if (anchor === 'ctr') {
    style.valign = 'middle';
  } else if (anchor === 'b') {
    style.valign = 'bottom';
  }

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

  // Inspect both run properties (<a:rPr>) and default run properties (<a:defRPr>)
  const propEls = [
    ...getElementsByLocalName(txBody, 'rPr'),
    ...getElementsByLocalName(txBody, 'defRPr'),
  ];

  for (const rPr of propEls) {
    if (style.fontSize === undefined) {
      const szAttr = rPr.getAttribute('sz');
      if (szAttr) {
        const hundredths = Number.parseInt(szAttr, 10);
        if (Number.isFinite(hundredths) && hundredths > 0) {
          style.fontSize = Math.min(96, Math.max(8, Math.round(hundredths / 100)));
        }
      }
    }
    if (style.bold === undefined) {
      const bAttr = rPr.getAttribute('b');
      if (bAttr === '1' || bAttr === 'true') {
        style.bold = true;
      } else if (bAttr === '0' || bAttr === 'false') {
        style.bold = false;
      }
    }
    if (style.italic === undefined) {
      const iAttr = rPr.getAttribute('i');
      if (iAttr === '1' || iAttr === 'true') {
        style.italic = true;
      } else if (iAttr === '0' || iAttr === 'false') {
        style.italic = false;
      }
    }
    if (style.color === undefined) {
      const solidFill = getDirectChildrenByLocalName(rPr, 'solidFill')[0];
      const resolved = resolveDrawingColor(solidFill, theme.colors);
      if (resolved) {
        style.color = resolved;
      }
    }
    if (style.fontFamily === undefined) {
      const latin = getDirectChildrenByLocalName(rPr, 'latin')[0];
      const ea = getDirectChildrenByLocalName(rPr, 'ea')[0];
      const family =
        resolveThemeFontName(latin?.getAttribute('typeface'), theme) ??
        resolveThemeFontName(ea?.getAttribute('typeface'), theme);
      if (family) {
        style.fontFamily = family;
      }
    }
  }

  return style;
}

function extractGradientFromElement(
  gradFillEl: Element,
  themeColors: Record<string, string>
): { gradientCss?: string; firstColor?: string } {
  const gsLst = getDirectChildrenByLocalName(gradFillEl, 'gsLst')[0];
  if (!gsLst) {
    return {};
  }
  const stops: Array<{ posPct: number; color: string }> = [];
  for (const gs of getDirectChildrenByLocalName(gsLst, 'gs')) {
    const rawPos = Number.parseInt(gs.getAttribute('pos') ?? '0', 10);
    const posPct = Number.isFinite(rawPos)
      ? Math.min(100, Math.max(0, Math.round(rawPos / 1000)))
      : 0;
    const color = resolveDrawingColor(gs, themeColors);
    if (color) {
      stops.push({ posPct, color });
    }
  }
  if (stops.length === 0) {
    return {};
  }
  stops.sort((a, b) => a.posPct - b.posPct);
  if (stops.length === 1) {
    return { firstColor: stops[0].color };
  }

  const lin = getDirectChildrenByLocalName(gradFillEl, 'lin')[0];
  const rawAng = Number.parseInt(lin?.getAttribute('ang') ?? '', 10);
  const cssDeg = Number.isFinite(rawAng) ? Math.round((rawAng / 60000 + 90) % 360) : 135;
  const stopsCss = stops.map((s) => `${s.color} ${s.posPct}%`).join(', ');
  return {
    gradientCss: `linear-gradient(${cssDeg}deg, ${stopsCss})`,
    firstColor: stops[0].color,
  };
}

function extractShapeFillAndBorder(
  spPr: Element | undefined,
  themeColors: Record<string, string>
): { fillColor?: string; borderColor?: string; fillBlipEmbedId?: string } {
  if (!spPr) {
    return {};
  }
  let fillColor: string | undefined;
  let fillBlipEmbedId: string | undefined;

  const hasNoFill = getDirectChildrenByLocalName(spPr, 'noFill').length > 0;
  if (!hasNoFill) {
    const solidFill = getDirectChildrenByLocalName(spPr, 'solidFill')[0];
    if (solidFill) {
      fillColor = resolveDrawingColor(solidFill, themeColors);
    } else {
      const gradFill = getDirectChildrenByLocalName(spPr, 'gradFill')[0];
      if (gradFill) {
        fillColor = extractGradientFromElement(gradFill, themeColors).firstColor;
      } else {
        const blipFill = getDirectChildrenByLocalName(spPr, 'blipFill')[0];
        if (blipFill) {
          const blip = getElementsByLocalName(blipFill, 'blip')[0];
          const embedId = blip ? getRelationshipAttr(blip, 'embed') : '';
          if (embedId) {
            fillBlipEmbedId = embedId;
          }
        }
      }
    }
  }

  let borderColor: string | undefined;
  const ln = getDirectChildrenByLocalName(spPr, 'ln')[0];
  if (ln && getDirectChildrenByLocalName(ln, 'noFill').length === 0) {
    const lnSolid = getDirectChildrenByLocalName(ln, 'solidFill')[0];
    borderColor = resolveDrawingColor(lnSolid, themeColors);
  }

  return { fillColor, borderColor, fillBlipEmbedId };
}

function extractSlideBackgroundInfo(
  xmlDoc: Document,
  themeColors: Record<string, string>
): {
  backgroundColor?: string;
  backgroundGradient?: string;
  backgroundBlipEmbedId?: string;
} {
  const bg = getElementsByLocalName(xmlDoc, 'bg')[0];
  if (!bg) {
    return {};
  }

  const bgPr = getDirectChildrenByLocalName(bg, 'bgPr')[0];
  if (bgPr) {
    const blipFill = getDirectChildrenByLocalName(bgPr, 'blipFill')[0];
    if (blipFill) {
      const blip = getElementsByLocalName(blipFill, 'blip')[0];
      const embedId = blip ? getRelationshipAttr(blip, 'embed') : '';
      if (embedId) {
        return { backgroundBlipEmbedId: embedId };
      }
    }

    const gradFill = getDirectChildrenByLocalName(bgPr, 'gradFill')[0];
    if (gradFill) {
      const grad = extractGradientFromElement(gradFill, themeColors);
      return {
        backgroundColor: grad.firstColor,
        backgroundGradient: grad.gradientCss,
      };
    }

    const solidFill = getDirectChildrenByLocalName(bgPr, 'solidFill')[0];
    if (solidFill) {
      return {
        backgroundColor: resolveDrawingColor(solidFill, themeColors),
      };
    }
  }

  const bgRef = getDirectChildrenByLocalName(bg, 'bgRef')[0];
  if (bgRef) {
    return {
      backgroundColor: resolveDrawingColor(bgRef, themeColors),
    };
  }

  return {};
}

function getPlaceholderMeta(spElement: Element): { type: string; idx: string } {
  const ph = getElementsByLocalName(spElement, 'ph')[0];
  return {
    type: (ph?.getAttribute('type') ?? '').trim(),
    idx: (ph?.getAttribute('idx') ?? '').trim(),
  };
}

function getPlaceholderType(spElement: Element): string {
  return getPlaceholderMeta(spElement).type;
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
  const rawBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer: rawBuffer.slice(0) });
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
    rawArrayBuffer: rawBuffer.slice(0),
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

const DEFAULT_SLIDE_CX = 12192000; // 13.333 in (16:9 widescreen)
const DEFAULT_SLIDE_CY = 6858000; // 7.5 in

interface PptxRelRecord {
  id: string;
  type: string;
  resolvedPath: string;
}

async function parseRelationshipsFile(
  zip: JSZip,
  xmlParser: DOMParser,
  relsPath: string,
  baseDir: string
): Promise<Map<string, PptxRelRecord>> {
  const map = new Map<string, PptxRelRecord>();
  const relsFile = findZipFileInsensitive(zip, relsPath);
  if (!relsFile) {
    return map;
  }
  try {
    const relsXml = await relsFile.async('string');
    const relsDoc = xmlParser.parseFromString(relsXml, 'application/xml');
    for (const rel of getElementsByLocalName(relsDoc, 'Relationship')) {
      const id = (rel.getAttribute('Id') ?? '').trim();
      const type = (rel.getAttribute('Type') ?? '').trim();
      const target = (rel.getAttribute('Target') ?? '').trim();
      const targetMode = (rel.getAttribute('TargetMode') ?? '').trim();
      if (!id || !target || targetMode.toLowerCase() === 'external') {
        continue;
      }
      map.set(id, {
        id,
        type,
        resolvedPath: resolveZipPath(baseDir, target),
      });
    }
  } catch {
    // Ignore malformed .rels safely
  }
  return map;
}

async function extractLayoutPlaceholderBoxes(
  zip: JSZip,
  xmlParser: DOMParser,
  layoutPath: string
): Promise<Map<string, EmuBox>> {
  const boxes = new Map<string, EmuBox>();
  const layoutFile = findZipFileInsensitive(zip, layoutPath);
  if (!layoutFile) {
    return boxes;
  }
  try {
    const xml = await layoutFile.async('string');
    const doc = xmlParser.parseFromString(xml, 'application/xml');
    for (const sp of getElementsByLocalName(doc, 'sp')) {
      const ph = getPlaceholderMeta(sp);
      if (!ph.type && !ph.idx) {
        continue;
      }
      const spPr = getDirectChildrenByLocalName(sp, 'spPr')[0];
      const xfrm = spPr ? getDirectChildrenByLocalName(spPr, 'xfrm')[0] : undefined;
      const box = extractEmuBoxFromXfrm(xfrm);
      if (!box) {
        continue;
      }
      if (ph.type && ph.idx) {
        boxes.set(`${ph.type}:${ph.idx}`, box);
      }
      if (ph.type && !boxes.has(ph.type)) {
        boxes.set(ph.type, box);
      }
      if (ph.idx && !boxes.has(`idx:${ph.idx}`)) {
        boxes.set(`idx:${ph.idx}`, box);
      }
    }
  } catch {
    // Ignore malformed layout XML
  }
  return boxes;
}

async function parsePresentationMetaAndOrder(
  zip: JSZip,
  xmlParser: DOMParser
): Promise<{
  slideEntries: Array<{ path: string; num: number }>;
  slideCx: number;
  slideCy: number;
}> {
  const ordered: Array<{ path: string; num: number }> = [];
  const seenPathsLower = new Set<string>();
  let slideCx = DEFAULT_SLIDE_CX;
  let slideCy = DEFAULT_SLIDE_CY;

  const presFile = findZipFileInsensitive(zip, 'ppt/presentation.xml');
  const presRels = await parseRelationshipsFile(
    zip,
    xmlParser,
    'ppt/_rels/presentation.xml.rels',
    'ppt'
  );

  if (presFile) {
    try {
      const presXml = await presFile.async('string');
      const presDoc = xmlParser.parseFromString(presXml, 'application/xml');

      const sldSz = getElementsByLocalName(presDoc, 'sldSz')[0];
      if (sldSz) {
        const parsedCx = Number.parseInt(sldSz.getAttribute('cx') ?? '', 10);
        const parsedCy = Number.parseInt(sldSz.getAttribute('cy') ?? '', 10);
        if (Number.isFinite(parsedCx) && parsedCx > 0 && Number.isFinite(parsedCy) && parsedCy > 0) {
          slideCx = parsedCx;
          slideCy = parsedCy;
        }
      }

      if (presRels.size > 0) {
        const sldIdElements = getElementsByLocalName(presDoc, 'sldId');
        for (const sldIdEl of sldIdElements) {
          const rId = getRelationshipAttr(sldIdEl, 'id');
          if (!rId) {
            continue;
          }
          const rel = presRels.get(rId);
          if (!rel) {
            continue;
          }
          const zipObj = findZipFileInsensitive(zip, rel.resolvedPath);
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

  return {
    slideEntries: ordered,
    slideCx,
    slideCy,
  };
}

export async function parsePptxArrayBuffer(
  arrayBuffer: ArrayBuffer,
  options?: { docIdPrefix?: string }
): Promise<{
  slidesPreview: SlidePreviewItem[];
  slideAspectRatio: number;
  outline: OutlineSection[];
  contentPreview: string;
  createdObjectUrls: string[];
}> {
  const createdObjectUrls: string[] = [];

  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const xmlParser = new DOMParser();

    const [{ slideEntries, slideCx, slideCy }, theme] = await Promise.all([
      parsePresentationMetaAndOrder(zip, xmlParser),
      parsePptxThemePalette(zip, xmlParser),
    ]);

    if (slideEntries.length === 0) {
      throw new Error('Không tìm thấy trang slide hợp lệ bên trong tệp .pptx này.');
    }

    const slideAspectRatio = slideCx / slideCy;
    const slidesPreview: SlidePreviewItem[] = [];
    const ignoredShapePh = new Set(['sldNum', 'dt', 'ftr']);
    const idPrefix = options?.docIdPrefix ? `${options.docIdPrefix}-` : 'uploaded-';
    const layoutPlaceholderCache = new Map<string, Map<string, EmuBox>>();
    const mediaBlobCache = new Map<string, string>();

    const getMediaBlobUrl = async (resolvedPath: string): Promise<string | null> => {
      const lowerKey = resolvedPath.toLowerCase();
      const cached = mediaBlobCache.get(lowerKey);
      if (cached) {
        return cached;
      }
      const extMatch = /\.([a-z0-9]+)$/i.exec(resolvedPath);
      const imgExt = extMatch ? extMatch[1].toLowerCase() : '';
      const mime = IMAGE_MIME_BY_EXT[imgExt];
      if (!mime) {
        return null;
      }
      const imgZipObj = findZipFileInsensitive(zip, resolvedPath);
      if (!imgZipObj || imgZipObj.dir) {
        return null;
      }
      const imgBuffer = await imgZipObj.async('arraybuffer');
      const blob = new Blob([imgBuffer], { type: mime });
      const blobUrl = URL.createObjectURL(blob);
      createdObjectUrls.push(blobUrl);
      mediaBlobCache.set(lowerKey, blobUrl);
      return blobUrl;
    };

    for (let i = 0; i < slideEntries.length; i++) {
      const { path: slidePath, num: slideNum } = slideEntries[i];
      const slideFile = zip.file(slidePath);
      if (!slideFile) {
        continue;
      }

      const lastSlash = slidePath.lastIndexOf('/');
      const slideDir = lastSlash >= 0 ? slidePath.slice(0, lastSlash) : '';
      const slideBaseName = lastSlash >= 0 ? slidePath.slice(lastSlash + 1) : slidePath;
      const relsPath = slideDir
        ? `${slideDir}/_rels/${slideBaseName}.rels`
        : `_rels/${slideBaseName}.rels`;

      const [xmlText, slideRels] = await Promise.all([
        slideFile.async('string'),
        parseRelationshipsFile(zip, xmlParser, relsPath, slideDir || 'ppt/slides'),
      ]);

      const xmlDoc = xmlParser.parseFromString(xmlText, 'application/xml');

      // Find linked slideLayout (for inherited placeholder geometry) and notesSlide
      let layoutPath: string | null = null;
      let notesPath: string | null = null;
      for (const rel of slideRels.values()) {
        if (rel.type.includes('/slideLayout') && !layoutPath) {
          layoutPath = rel.resolvedPath;
        } else if (rel.type.includes('/notesSlide') && !notesPath) {
          notesPath = rel.resolvedPath;
        }
      }

      let layoutPlaceholders = new Map<string, EmuBox>();
      if (layoutPath) {
        const layoutKey = layoutPath.toLowerCase();
        const cachedLayout = layoutPlaceholderCache.get(layoutKey);
        if (cachedLayout) {
          layoutPlaceholders = cachedLayout;
        } else {
          layoutPlaceholders = await extractLayoutPlaceholderBoxes(zip, xmlParser, layoutPath);
          layoutPlaceholderCache.set(layoutKey, layoutPlaceholders);
        }
      }

      // 1. Parse slide background (<p:bg>) with inheritance from slideLayout & slideMaster
      let layoutDoc: Document | null = null;
      let layoutRels = new Map<string, PptxRelRecord>();
      let masterDoc: Document | null = null;
      let masterRels = new Map<string, PptxRelRecord>();

      if (layoutPath) {
        const lFile = findZipFileInsensitive(zip, layoutPath);
        if (lFile) {
          layoutDoc = xmlParser.parseFromString(await lFile.async("string"), "application/xml");
          const lSlash = layoutPath.lastIndexOf("/");
          const lDir = lSlash >= 0 ? layoutPath.slice(0, lSlash) : "";
          const lBase = lSlash >= 0 ? layoutPath.slice(lSlash + 1) : layoutPath;
          layoutRels = await parseRelationshipsFile(zip, xmlParser, lDir ? `${lDir}/_rels/${lBase}.rels` : `_rels/${lBase}.rels`, lDir || "ppt/slideLayouts");
          for (const rel of layoutRels.values()) {
            if (rel.type.includes("/slideMaster")) {
              const mFile = findZipFileInsensitive(zip, rel.resolvedPath);
              if (mFile) {
                masterDoc = xmlParser.parseFromString(await mFile.async("string"), "application/xml");
                const mSlash = rel.resolvedPath.lastIndexOf("/");
                const mDir = mSlash >= 0 ? rel.resolvedPath.slice(0, mSlash) : "";
                const mBase = mSlash >= 0 ? rel.resolvedPath.slice(mSlash + 1) : rel.resolvedPath;
                masterRels = await parseRelationshipsFile(zip, xmlParser, mDir ? `${mDir}/_rels/${mBase}.rels` : `_rels/${mBase}.rels`, mDir || "ppt/slideMasters");
              }
              break;
            }
          }
        }
      }

      const bgSources: Array<{ doc: Document; rels: Map<string, PptxRelRecord> }> = [
        { doc: xmlDoc, rels: slideRels },
        ...(layoutDoc ? [{ doc: layoutDoc, rels: layoutRels }] : []),
        ...(masterDoc ? [{ doc: masterDoc, rels: masterRels }] : []),
      ];

      let backgroundColor: string | undefined;
      let backgroundGradient: string | undefined;
      let backgroundImageSrc: string | undefined;
      const backgroundImagePaths = new Set<string>();

      for (const srcEntry of bgSources) {
        const info = extractSlideBackgroundInfo(srcEntry.doc, theme.colors);
        if (info.backgroundBlipEmbedId) {
          const bgRel = srcEntry.rels.get(info.backgroundBlipEmbedId);
          if (bgRel) {
            const bgBlob = await getMediaBlobUrl(bgRel.resolvedPath);
            if (bgBlob) {
              backgroundImageSrc = bgBlob;
              backgroundImagePaths.add(bgRel.resolvedPath.toLowerCase());
              break;
            }
          }
        }
        if (info.backgroundGradient || info.backgroundColor) {
          backgroundColor = info.backgroundColor;
          backgroundGradient = info.backgroundGradient;
          break;
        }
      }

      const blocks: SlideTextBlock[] = [];
      const images: SlideImageItem[] = [];
      const positionedImagePaths = new Set<string>();
      let blockCounter = 0;

      // 2. Recursively traverse <p:spTree> preserving Z-order & nested <p:grpSp> coordinates
      const walkShapeContainer = async (
        containerEl: Element,
        groupTransform: GroupTransform
      ): Promise<void> => {
        for (let cIdx = 0; cIdx < containerEl.childNodes.length; cIdx++) {
          const child = containerEl.childNodes[cIdx];
          if (child.nodeType !== Node.ELEMENT_NODE) {
            continue;
          }
          const el = child as Element;
          const local = el.localName.toLowerCase();

          if (local === 'grpsp') {
            const grpSpPr = getDirectChildrenByLocalName(el, 'grpSpPr')[0];
            const grpXfrm = grpSpPr ? getDirectChildrenByLocalName(grpSpPr, 'xfrm')[0] : undefined;
            const nextTransform = composeGroupTransform(groupTransform, grpXfrm);
            await walkShapeContainer(el, nextTransform);
            continue;
          }

          if (local === 'sp') {
            const phMeta = getPlaceholderMeta(el);
            const phType = phMeta.type;
            if (phType && ignoredShapePh.has(phType)) {
              continue;
            }

            const spPr = getDirectChildrenByLocalName(el, 'spPr')[0];
            const xfrm = spPr ? getDirectChildrenByLocalName(spPr, 'xfrm')[0] : undefined;
            const localBox =
              extractEmuBoxFromXfrm(xfrm) ??
              (phType && phMeta.idx ? layoutPlaceholders.get(`${phType}:${phMeta.idx}`) : undefined) ??
              (phType ? layoutPlaceholders.get(phType) : undefined) ??
              (phMeta.idx ? layoutPlaceholders.get(`idx:${phMeta.idx}`) : undefined) ??
              null;

            const coords = localBox
              ? mapBoxToPct(localBox, groupTransform, slideCx, slideCy)
              : getDefaultPlaceholderPct(phType, blockCounter);

            const { fillColor, borderColor, fillBlipEmbedId } = extractShapeFillAndBorder(
              spPr,
              theme.colors
            );

            // Picture-filled shape (e.g. image cropped into shape frame)
            if (fillBlipEmbedId) {
              const rel = slideRels.get(fillBlipEmbedId);
              if (rel && !backgroundImagePaths.has(rel.resolvedPath.toLowerCase())) {
                const blobUrl = await getMediaBlobUrl(rel.resolvedPath);
                if (blobUrl) {
                  positionedImagePaths.add(rel.resolvedPath.toLowerCase());
                  const fileName = rel.resolvedPath.split('/').pop() ?? `image-${images.length + 1}`;
                  images.push({
                    id: `${idPrefix}slide-${slideNum}-img-${images.length + 1}`,
                    src: blobUrl,
                    alt: `Ảnh slide ${i + 1} (${fileName})`,
                    xPct: coords.xPct,
                    yPct: coords.yPct,
                    wPct: coords.wPct,
                    hPct: coords.hPct,
                  });
                }
              }
            }

            const txBodies = getDirectChildrenByLocalName(el, 'txBody');
            const paragraphsText: string[] = [];
            let runStyle: ReturnType<typeof extractFirstRunStyle> = {};

            for (const txBody of txBodies) {
              if (Object.keys(runStyle).length === 0) {
                runStyle = extractFirstRunStyle(txBody, theme);
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
              if (!phType && !phMeta.idx && (fillColor || borderColor) && coords.wPct < 99.5 && coords.hPct < 99.5) {
                blockCounter++;
                blocks.push({
                  id: `${idPrefix}slide-${slideNum}-shape-${blockCounter}`,
                  kind: "text",
                  text: "",
                  fontSize: 12,
                  fillColor,
                  borderColor,
                  role: "body",
                  xPct: coords.xPct,
                  yPct: coords.yPct,
                  wPct: coords.wPct,
                  hPct: coords.hPct,
                });
              }
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
              id: `${idPrefix}slide-${slideNum}-block-${blockCounter}`,
              kind: 'text',
              text: combinedText,
              fontSize: runStyle.fontSize ?? defaultFontSize,
              bold: runStyle.bold ?? defaultBold,
              italic: runStyle.italic,
              color: runStyle.color,
              align: runStyle.align ?? (phType === 'ctrTitle' ? 'center' : 'left'),
              valign: runStyle.valign ?? 'top',
              fontFamily: runStyle.fontFamily,
              fillColor,
              borderColor,
              role,
              xPct: coords.xPct,
              yPct: coords.yPct,
              wPct: coords.wPct,
              hPct: coords.hPct,
            });
            continue;
          }

          if (local === 'pic') {
            const blipFill = getDirectChildrenByLocalName(el, 'blipFill')[0];
            const blip = blipFill ? getElementsByLocalName(blipFill, 'blip')[0] : undefined;
            const embedId = blip ? getRelationshipAttr(blip, 'embed') : '';
            if (!embedId) {
              continue;
            }
            const rel = slideRels.get(embedId);
            if (!rel || backgroundImagePaths.has(rel.resolvedPath.toLowerCase())) {
              continue;
            }
            const blobUrl = await getMediaBlobUrl(rel.resolvedPath);
            if (!blobUrl) {
              continue;
            }

            const spPr = getDirectChildrenByLocalName(el, 'spPr')[0];
            const xfrm = spPr ? getDirectChildrenByLocalName(spPr, 'xfrm')[0] : undefined;
            const localBox = extractEmuBoxFromXfrm(xfrm);
            const coords = localBox
              ? mapBoxToPct(localBox, groupTransform, slideCx, slideCy)
              : undefined;

            const cNvPr = getElementsByLocalName(el, 'cNvPr')[0];
            const descr =
              (cNvPr?.getAttribute('descr') ?? cNvPr?.getAttribute('name') ?? '').trim();
            const fileName = rel.resolvedPath.split('/').pop() ?? `image-${images.length + 1}`;
            positionedImagePaths.add(rel.resolvedPath.toLowerCase());

            images.push({
              id: `${idPrefix}slide-${slideNum}-img-${images.length + 1}`,
              src: blobUrl,
              alt: descr || `Ảnh slide ${i + 1} (${fileName})`,
              xPct: coords?.xPct,
              yPct: coords?.yPct,
              wPct: coords?.wPct,
              hPct: coords?.hPct,
            });
            continue;
          }

          if (local === 'graphicframe') {
            const tbl = getElementsByLocalName(el, 'tbl')[0];
            if (!tbl) {
              continue;
            }
            const xfrm = getDirectChildrenByLocalName(el, 'xfrm')[0];
            const localBox = extractEmuBoxFromXfrm(xfrm);
            const coords = localBox
              ? mapBoxToPct(localBox, groupTransform, slideCx, slideCy)
              : { xPct: 6, yPct: 22, wPct: 88, hPct: 62 };

            const trElements = getDirectChildrenByLocalName(tbl, 'tr');
            const tableRows: string[][] = [];
            for (const tr of trElements) {
              const cells = getDirectChildrenByLocalName(tr, 'tc');
              if (cells.length === 0) {
                continue;
              }
              const rowValues = cells.map((tc) => {
                const paras = getElementsByLocalName(tc, 'p');
                return paras
                  .map((p) => extractParagraphText(p))
                  .filter(Boolean)
                  .join(' ');
              });
              tableRows.push(rowValues);
            }

            const hasAnyCellContent = tableRows.some((r) => r.some((cell) => cell.trim().length > 0));
            if (!hasAnyCellContent) {
              continue;
            }

            const tableTextFallback = tableRows
              .map((row) => row.join(' | '))
              .filter((line) => line.replace(/\|/g, '').trim().length > 0)
              .join('\n')
              .trim();

            blockCounter++;
            blocks.push({
              id: `${idPrefix}slide-${slideNum}-block-${blockCounter}`,
              kind: 'table',
              tableRows,
              text: tableTextFallback,
              fontSize: 13,
              role: 'body',
              align: 'left',
              valign: 'middle',
              xPct: coords.xPct,
              yPct: coords.yPct,
              wPct: coords.wPct,
              hPct: coords.hPct,
            });
          }
        }
      };

      const walkDecorOnly = async (docToWalk: Document | null, relsMap: Map<string, PptxRelRecord>) => {
        if (!docToWalk) return;
        const tree = getElementsByLocalName(docToWalk, "spTree")[0];
        if (!tree) return;
        for (let cIdx = 0; cIdx < tree.childNodes.length; cIdx++) {
          const child = tree.childNodes[cIdx];
          if (child.nodeType !== Node.ELEMENT_NODE) continue;
          const el = child as Element;
          const local = el.localName.toLowerCase();
          if (local === "sp") {
            const ph = getPlaceholderMeta(el);
            if (ph.type || ph.idx) continue;
            const spPr = getDirectChildrenByLocalName(el, "spPr")[0];
            const xfrm = spPr ? getDirectChildrenByLocalName(spPr, "xfrm")[0] : undefined;
            const box = extractEmuBoxFromXfrm(xfrm);
            if (!box) continue;
            const coords = mapBoxToPct(box, IDENTITY_GROUP_TRANSFORM, slideCx, slideCy);
            const { fillColor, borderColor, fillBlipEmbedId } = extractShapeFillAndBorder(spPr, theme.colors);
            if (fillBlipEmbedId) {
              const rel = relsMap.get(fillBlipEmbedId);
              if (rel) {
                const blobUrl = await getMediaBlobUrl(rel.resolvedPath);
                if (blobUrl) {
                  images.push({ id: `${idPrefix}slide-${slideNum}-mimg-${images.length + 1}`, src: blobUrl, alt: "Layout", xPct: coords.xPct, yPct: coords.yPct, wPct: coords.wPct, hPct: coords.hPct });
                }
              }
            } else if (fillColor || borderColor) {
              blockCounter++;
              blocks.push({ id: `${idPrefix}slide-${slideNum}-mshape-${blockCounter}`, kind: "text", text: "", fontSize: 12, fillColor, borderColor, role: "body", xPct: coords.xPct, yPct: coords.yPct, wPct: coords.wPct, hPct: coords.hPct });
            }
          } else if (local === "pic") {
            const blipFill = getDirectChildrenByLocalName(el, "blipFill")[0];
            const blip = blipFill ? getElementsByLocalName(blipFill, "blip")[0] : undefined;
            const embedId = blip ? getRelationshipAttr(blip, "embed") : "";
            const rel = embedId ? relsMap.get(embedId) : undefined;
            if (!rel) continue;
            const blobUrl = await getMediaBlobUrl(rel.resolvedPath);
            if (!blobUrl) continue;
            const spPr = getDirectChildrenByLocalName(el, "spPr")[0];
            const xfrm = spPr ? getDirectChildrenByLocalName(spPr, "xfrm")[0] : undefined;
            const box = extractEmuBoxFromXfrm(xfrm);
            if (!box) continue;
            const coords = mapBoxToPct(box, IDENTITY_GROUP_TRANSFORM, slideCx, slideCy);
            images.push({ id: `${idPrefix}slide-${slideNum}-mpic-${images.length + 1}`, src: blobUrl, alt: "Master", xPct: coords.xPct, yPct: coords.yPct, wPct: coords.wPct, hPct: coords.hPct });
          }
        }
      };

      if (xmlDoc.documentElement.getAttribute("showMasterSp") !== "0") {
        if (!layoutDoc || layoutDoc.documentElement.getAttribute("showMasterSp") !== "0") {
          await walkDecorOnly(masterDoc, masterRels);
        }
        await walkDecorOnly(layoutDoc, layoutRels);
      }

      const spTree = getElementsByLocalName(xmlDoc, 'spTree')[0];
      if (spTree) {
        await walkShapeContainer(spTree, IDENTITY_GROUP_TRANSFORM);
      }

      // Fallback: pick up any unreferenced image rels not already placed via <p:pic>/<p:sp>/<p:bg>
      for (const rel of slideRels.values()) {
        if (!rel.type.includes('/image')) {
          continue;
        }
        const pathKey = rel.resolvedPath.toLowerCase();
        if (backgroundImagePaths.has(pathKey) || positionedImagePaths.has(pathKey)) {
          continue;
        }
        const blobUrl = await getMediaBlobUrl(rel.resolvedPath);
        if (!blobUrl) {
          continue;
        }
        positionedImagePaths.add(pathKey);
        const fileName = rel.resolvedPath.split('/').pop() ?? `image-${images.length + 1}`;
        images.push({
          id: `${idPrefix}slide-${slideNum}-img-${images.length + 1}`,
          src: blobUrl,
          alt: `Ảnh slide ${i + 1} (${fileName})`,
        });
      }

      // 3. Read real speaker notes from rel target or fallback path
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

      // 4. Derive clean title, subtitle, and bullets from real blocks
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
      // while text blocks carry dark colors, so dark text never sinks into a dark stage.
      if (!backgroundColor && !backgroundGradient && !backgroundImageSrc) {
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
        id: `${idPrefix}slide-${slideNum}-${Date.now()}`,
        index: i + 1,
        title: slideTitle,
        subtitle,
        bullets,
        blocks,
        images,
        speakerNote,
        accentColor: theme.colors.accent1 || '#4F46E5',
        backgroundColor,
        backgroundGradient,
        backgroundImageSrc,
        slideWidthRatio: slideAspectRatio,
        isMasterPositioned: true,
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
    const contentPreview =
      firstSlideTextParts.length > 0
        ? truncatePreview(firstSlideTextParts.join(' · '), 260)
        : `Tệp trình chiếu gồm ${slidesPreview.length} slide.`;

    return {
      slidesPreview,
      slideAspectRatio,
      outline,
      contentPreview,
      createdObjectUrls,
    };
  } catch (err) {
    for (const url of createdObjectUrls) {
      URL.revokeObjectURL(url);
    }
    throw err;
  }
}

async function parsePptxFile(file: File): Promise<ParseFileResult> {
  const rawBuffer = await file.arrayBuffer();
  try {
    const parsed = await parsePptxArrayBuffer(rawBuffer.slice(0));
    const title = stripFileExtension(file.name);
    const fileSize = formatFileSize(file.size);
    const contentPreview =
      parsed.contentPreview || `Tệp trình chiếu gồm ${parsed.slidesPreview.length} slide (${fileSize}).`;

    const document: DocumentItem = {
      id: createDocumentId('pptx'),
      title,
      type: 'pptx',
      updatedAt: 'Vừa tải lên',
      isFavorite: false,
      isShared: false,
      folder: 'Tải lên gần đây',
      fileSize,
      pageCount: parsed.slidesPreview.length,
      thumbnail: createThumbnailDataUri('pptx', title, '#4F46E5'),
      slidesPreview: parsed.slidesPreview,
      outline: parsed.outline,
      contentPreview,
      rawArrayBuffer: rawBuffer.slice(0),
      slideAspectRatio: parsed.slideAspectRatio,
      isUploaded: true,
      sourceFile: file,
    };

    return {
      document,
      createdObjectUrls: parsed.createdObjectUrls,
    };
  } catch (err) {
    if (
      err instanceof Error &&
      err.message.includes('Không tìm thấy trang slide hợp lệ')
    ) {
      return {
        error: err.message,
        createdObjectUrls: [],
      };
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

import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import * as XLSX from 'xlsx';
import type {
  DocumentItem,
  SlidePreviewItem,
  SlideShapeKind,
  SlideTextBlock,
  SpreadsheetSheet,
} from '../types/document';
import { evaluateCellValue } from './formulaEvaluator';
import { saveFileUniversal, type SaveFileResult } from './saveFileUniversal';

function sanitizeFileName(title: string, fallback: string, stripExtRegex?: RegExp): string {
  let cleaned = (title || '').trim();
  if (stripExtRegex) {
    cleaned = cleaned.replace(stripExtRegex, '').trim();
  }
  cleaned = cleaned
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^\.+|\.+$/g, '')
    .trim();

  if (!cleaned) {
    return fallback;
  }
  return cleaned.length > 100 ? cleaned.slice(0, 100).trim() : cleaned;
}

async function triggerBlobDownload(
  blob: Blob,
  fileName: string,
  mimeType?: string
): Promise<SaveFileResult> {
  return saveFileUniversal(blob, fileName, {
    mimeType: mimeType || blob.type || 'application/octet-stream',
    shareTitle: fileName,
  });
}

// ============================================================================
// 1. EXPORT WORD (.DOCX) — Genuine OpenXML Package via JSZip
// ============================================================================

const NAMED_CSS_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'DC2626',
  blue: '2563EB',
  green: '16A34A',
  yellow: 'EAB308',
  orange: 'EA580C',
  purple: '9333EA',
  pink: 'DB2777',
  gray: '64748B',
  grey: '64748B',
  navy: '1E3A8A',
  teal: '0D9488',
  cyan: '06B6D4',
  magenta: 'C026D3',
  brown: '92400E',
};

function escapeXmlText(value: string): string {
  // Strip invalid XML 1.0 control characters before escaping entities
  const safe = value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');
  return safe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function cssColorToHex6(raw: string | null | undefined): string | undefined {
  if (!raw) {
    return undefined;
  }
  const value = raw.trim().toLowerCase();
  if (
    !value ||
    value === 'transparent' ||
    value === 'inherit' ||
    value === 'initial' ||
    value === 'unset' ||
    value === 'none' ||
    value === 'auto'
  ) {
    return undefined;
  }

  if (NAMED_CSS_COLORS[value]) {
    return NAMED_CSS_COLORS[value];
  }

  const hexMatch = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value);
  if (hexMatch) {
    const hex = hexMatch[1].toUpperCase();
    if (hex.length === 3) {
      return `${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
    }
    return hex;
  }

  const rgbMatch =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([0-9.]+))?\s*\)$/i.exec(value);
  if (rgbMatch) {
    if (rgbMatch[4] !== undefined && Number.parseFloat(rgbMatch[4]) === 0) {
      return undefined;
    }
    const r = Math.min(255, Math.max(0, Number.parseInt(rgbMatch[1], 10)));
    const g = Math.min(255, Math.max(0, Number.parseInt(rgbMatch[2], 10)));
    const b = Math.min(255, Math.max(0, Number.parseInt(rgbMatch[3], 10)));
    return [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  return undefined;
}

type DocxAlign = 'left' | 'center' | 'right' | 'both';

interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  colorHex?: string;
  bgHex?: string;
}

interface ParagraphOptions {
  styleId?: 'Heading1' | 'Heading2' | 'Heading3';
  align?: DocxAlign;
  listType?: 'bullet' | 'number';
  listLevel?: number;
  isQuote?: boolean;
}

const DOCX_BLOCK_TAGS = new Set([
  'ADDRESS',
  'ARTICLE',
  'ASIDE',
  'BLOCKQUOTE',
  'DIV',
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
  'OL',
  'P',
  'PRE',
  'SECTION',
  'TABLE',
  'UL',
]);

function extractAlignment(el: HTMLElement): DocxAlign | undefined {
  if (el.tagName.toUpperCase() === 'CENTER') {
    return 'center';
  }
  const rawAlign = (el.style.textAlign || el.getAttribute('align') || '').trim().toLowerCase();
  if (rawAlign === 'center') {
    return 'center';
  }
  if (rawAlign === 'right' || rawAlign === 'end') {
    return 'right';
  }
  if (rawAlign === 'justify') {
    return 'both';
  }
  if (rawAlign === 'left' || rawAlign === 'start') {
    return 'left';
  }
  return undefined;
}

function isTransparentCssColor(raw: string): boolean {
  const value = raw.trim().toLowerCase();
  if (value === 'transparent' || value === 'none' || value === 'initial' || value === 'unset') {
    return true;
  }
  const rgbaComma =
    /^rgba\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*([0-9.]+)%?\s*\)$/.exec(value);
  if (rgbaComma && Number.parseFloat(rgbaComma[1]) === 0) {
    return true;
  }
  const rgbSlash =
    /^rgba?\(\s*\d{1,3}\s+\d{1,3}\s+\d{1,3}\s*\/\s*([0-9.]+)%?\s*\)$/.exec(value);
  if (rgbSlash && Number.parseFloat(rgbSlash[1]) === 0) {
    return true;
  }
  if (/^#[0-9a-f]{6}00$/.test(value) || /^#[0-9a-f]{3}0$/.test(value)) {
    return true;
  }
  return false;
}

interface EmbeddedDocxImage {
  rId: string;
  docPrId: number;
  fileName: string;
  ext: 'png' | 'jpeg' | 'gif';
  mime: string;
  bytes: Uint8Array;
}

interface DocxExportContext {
  images: EmbeddedDocxImage[];
  nextRelId: number;
  nextDocPrId: number;
}

function decodeBase64ToBytes(base64: string): Uint8Array | null {
  try {
    const clean = base64.replace(/\s+/g, '');
    if (!clean) {
      return null;
    }
    const binary = atob(clean);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes.length > 0 ? bytes : null;
  } catch {
    return null;
  }
}

function readImageDimensions(
  bytes: Uint8Array,
  ext: 'png' | 'jpeg' | 'gif'
): { width: number; height: number } {
  const fallback = { width: 480, height: 320 };

  if (ext === 'png') {
    if (
      bytes.length >= 24 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47
    ) {
      const width =
        ((bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19]) >>> 0;
      const height =
        ((bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23]) >>> 0;
      if (width > 0 && height > 0) {
        return { width, height };
      }
    }
    return fallback;
  }

  if (ext === 'gif') {
    if (bytes.length >= 10 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
      const width = bytes[6] | (bytes[7] << 8);
      const height = bytes[8] | (bytes[9] << 8);
      if (width > 0 && height > 0) {
        return { width, height };
      }
    }
    return fallback;
  }

  // JPEG SOF0..SOF15 scan
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 3 < bytes.length) {
      while (offset < bytes.length && bytes[offset] !== 0xff) {
        offset++;
      }
      while (offset < bytes.length && bytes[offset] === 0xff) {
        offset++;
      }
      if (offset >= bytes.length) {
        break;
      }
      const marker = bytes[offset++];
      if (marker === 0xd9 || marker === 0xda) {
        break;
      }
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) {
        continue;
      }
      if (offset + 1 >= bytes.length) {
        break;
      }
      const segLen = (bytes[offset] << 8) | bytes[offset + 1];
      if (segLen < 2 || offset + segLen > bytes.length) {
        break;
      }
      const isSof =
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc;
      if (isSof && offset + 6 < bytes.length) {
        const height = (bytes[offset + 3] << 8) | bytes[offset + 4];
        const width = (bytes[offset + 5] << 8) | bytes[offset + 6];
        if (width > 0 && height > 0) {
          return { width, height };
        }
      }
      offset += segLen;
    }
  }

  return fallback;
}

function buildImageDrawingRunXml(el: HTMLElement, ctx: DocxExportContext): string | null {
  const src = (el.getAttribute('src') ?? '').trim();
  const match = /^data:image\/(png|jpeg|jpg|gif);base64,([A-Za-z0-9+/=\s]+)$/i.exec(src);
  if (!match) {
    return null;
  }

  const rawSubtype = match[1].toLowerCase();
  const ext: 'png' | 'jpeg' | 'gif' =
    rawSubtype === 'jpg' || rawSubtype === 'jpeg'
      ? 'jpeg'
      : rawSubtype === 'gif'
      ? 'gif'
      : 'png';
  const mime = ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/png';

  const bytes = decodeBase64ToBytes(match[2]);
  if (!bytes) {
    return null;
  }

  const dims = readImageDimensions(bytes, ext);
  const emuPerPx = 9525; // 914400 EMU/in / 96 DPI
  const maxWidthEmu = 5486400; // 6.0 inches (fits comfortably within A4 margins)
  const maxHeightEmu = 6858000; // 7.5 inches

  let cx = Math.max(emuPerPx, Math.round(dims.width * emuPerPx));
  let cy = Math.max(emuPerPx, Math.round(dims.height * emuPerPx));

  const scale = Math.min(1, maxWidthEmu / cx, maxHeightEmu / cy);
  if (scale < 1) {
    cx = Math.max(emuPerPx, Math.round(cx * scale));
    cy = Math.max(emuPerPx, Math.round(cy * scale));
  }

  const imageIndex = ctx.images.length + 1;
  const rId = `rId${ctx.nextRelId++}`;
  const docPrId = ctx.nextDocPrId++;
  const fileName = `image${imageIndex}.${ext}`;
  const altText = escapeXmlText((el.getAttribute('alt') ?? '').trim() || `Image ${imageIndex}`);

  ctx.images.push({
    rId,
    docPrId,
    fileName,
    ext,
    mime,
    bytes,
  });

  return (
    '<w:r>' +
    '<w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${docPrId}" name="Picture ${docPrId}" descr="${altText}"/>` +
    '<wp:cNvGraphicFramePr>' +
    '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>' +
    '</wp:cNvGraphicFramePr>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:nvPicPr>' +
    `<pic:cNvPr id="${docPrId}" name="${fileName}" descr="${altText}"/>` +
    '<pic:cNvPicPr/>' +
    '</pic:nvPicPr>' +
    '<pic:blipFill>' +
    `<a:blip r:embed="${rId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch>' +
    '</pic:blipFill>' +
    '<pic:spPr>' +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</pic:spPr>' +
    '</pic:pic>' +
    '</a:graphicData>' +
    '</a:graphic>' +
    '</wp:inline>' +
    '</w:drawing>' +
    '</w:r>'
  );
}

function mergeElementRunFormat(el: HTMLElement, parent: RunFormat): RunFormat {
  const next: RunFormat = { ...parent };
  const tag = el.tagName.toUpperCase();

  if (tag === 'STRONG' || tag === 'B' || tag === 'TH') {
    next.bold = true;
  } else if (tag === 'EM' || tag === 'I') {
    next.italic = true;
  } else if (tag === 'U') {
    next.underline = true;
  } else if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') {
    next.strike = true;
  } else if (tag === 'MARK') {
    next.bgHex = 'FEF08A';
  } else if (tag === 'FONT') {
    const fontColor = cssColorToHex6(el.getAttribute('color'));
    if (fontColor) {
      next.colorHex = fontColor;
    }
  }

  const fontWeight = (el.style.fontWeight || '').trim().toLowerCase();
  if (fontWeight === 'bold' || fontWeight === 'bolder' || Number.parseInt(fontWeight, 10) >= 600) {
    next.bold = true;
  } else if (fontWeight === 'normal' || (fontWeight !== '' && Number.parseInt(fontWeight, 10) < 500)) {
    next.bold = false;
  }

  const fontStyle = (el.style.fontStyle || '').trim().toLowerCase();
  if (fontStyle === 'italic' || fontStyle === 'oblique') {
    next.italic = true;
  } else if (fontStyle === 'normal') {
    next.italic = false;
  }

  const textDecoration = (el.style.textDecoration || el.style.textDecorationLine || '')
    .trim()
    .toLowerCase();
  if (textDecoration.includes('underline')) {
    next.underline = true;
  }
  if (textDecoration.includes('line-through')) {
    next.strike = true;
  }

  const inlineColor = cssColorToHex6(el.style.color);
  if (inlineColor) {
    next.colorHex = inlineColor;
  }

  const rawBg = (el.style.backgroundColor || '').trim();
  if (rawBg) {
    if (isTransparentCssColor(rawBg)) {
      delete next.bgHex;
    } else {
      const inlineBg = cssColorToHex6(rawBg);
      if (inlineBg) {
        // Ignore pure white background artifacts from copy/paste or cleared highlights
        if (inlineBg !== 'FFFFFF') {
          next.bgHex = inlineBg;
        } else {
          delete next.bgHex;
        }
      }
    }
  }

  return next;
}

function buildRunPropertiesXml(format: RunFormat): string {
  const parts: string[] = [];
  if (format.bold) {
    parts.push('<w:b/>');
  }
  if (format.italic) {
    parts.push('<w:i/>');
  }
  if (format.strike) {
    parts.push('<w:strike/>');
  }
  if (format.colorHex) {
    parts.push(`<w:color w:val="${format.colorHex}"/>`);
  }
  if (format.underline) {
    parts.push('<w:u w:val="single"/>');
  }
  if (format.bgHex) {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${format.bgHex}"/>`);
  }
  return parts.length > 0 ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

function collectInlineRunsXml(
  nodes: Node[],
  format: RunFormat,
  ctx: DocxExportContext
): string[] {
  const runs: string[] = [];

  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      const rawText = (node.nodeValue ?? '').replace(/\r\n/g, '\n');
      if (!rawText) {
        continue;
      }
      // Collapse pure indentation newlines between block tags unless meaningful text exists
      const normalized = rawText.replace(/\n+/g, ' ');
      if (!normalized) {
        continue;
      }
      const rPr = buildRunPropertiesXml(format);
      runs.push(`<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(normalized)}</w:t></w:r>`);
      continue;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }

    const el = node as HTMLElement;
    const tag = el.tagName.toUpperCase();
    if (tag === 'SCRIPT' || tag === 'STYLE') {
      continue;
    }
    if (tag === 'BR') {
      runs.push('<w:r><w:br/></w:r>');
      continue;
    }
    if (tag === 'IMG') {
      const drawingRun = buildImageDrawingRunXml(el, ctx);
      if (drawingRun) {
        runs.push(drawingRun);
      }
      continue;
    }

    const nextFormat = mergeElementRunFormat(el, format);
    runs.push(...collectInlineRunsXml(Array.from(el.childNodes), nextFormat, ctx));
  }

  return runs;
}

function buildParagraphXml(runsXml: string[], options: ParagraphOptions = {}): string {
  const pPrParts: string[] = [];

  if (options.styleId) {
    pPrParts.push(`<w:pStyle w:val="${options.styleId}"/>`);
  }

  if (options.listType) {
    const numId = options.listType === 'number' ? '2' : '1';
    const ilvl = Math.min(3, Math.max(0, options.listLevel ?? 0));
    pPrParts.push(
      `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`
    );
  }

  if (options.isQuote) {
    pPrParts.push(
      '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="94A3B8"/></w:pBdr>'
    );
    pPrParts.push('<w:ind w:left="480"/>');
  }

  if (options.align) {
    pPrParts.push(`<w:jc w:val="${options.align}"/>`);
  }

  const pPr = pPrParts.length > 0 ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : '';
  return `<w:p>${pPr}${runsXml.join('')}</w:p>`;
}

function hasDirectBlockChildren(el: HTMLElement): boolean {
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i];
    if (child.nodeType === Node.ELEMENT_NODE) {
      const tag = (child as HTMLElement).tagName.toUpperCase();
      if (DOCX_BLOCK_TAGS.has(tag)) {
        return true;
      }
    }
  }
  return false;
}

function isWhitespaceOnlyNodes(nodes: Node[]): boolean {
  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if ((node.nodeValue ?? '').trim().length > 0) {
        return false;
      }
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      return false;
    }
  }
  return true;
}

function convertTableToDocxXml(
  tableEl: HTMLTableElement,
  inheritedFormat: RunFormat,
  ctx: DocxExportContext
): string {
  const trElements: HTMLTableRowElement[] = [];

  const gatherRows = (container: Element): void => {
    for (let i = 0; i < container.childNodes.length; i++) {
      const child = container.childNodes[i];
      if (child.nodeType !== Node.ELEMENT_NODE) {
        continue;
      }
      const el = child as HTMLElement;
      const tag = el.tagName.toUpperCase();
      if (tag === 'TR') {
        trElements.push(el as HTMLTableRowElement);
      } else if (tag === 'THEAD' || tag === 'TBODY' || tag === 'TFOOT') {
        gatherRows(el);
      }
    }
  };
  gatherRows(tableEl);

  if (trElements.length === 0) {
    return '';
  }

  let maxCols = 1;
  for (const tr of trElements) {
    let colCount = 0;
    for (let i = 0; i < tr.childNodes.length; i++) {
      const child = tr.childNodes[i];
      if (child.nodeType === Node.ELEMENT_NODE) {
        const tag = (child as HTMLElement).tagName.toUpperCase();
        if (tag === 'TD' || tag === 'TH') {
          const span = Number.parseInt((child as HTMLElement).getAttribute('colspan') ?? '1', 10);
          colCount += Number.isFinite(span) && span > 0 ? span : 1;
        }
      }
    }
    if (colCount > maxCols) {
      maxCols = colCount;
    }
  }

  const totalWidthTwips = 9360;
  const colWidth = Math.max(600, Math.floor(totalWidthTwips / maxCols));
  const tblGridCols = Array.from(
    { length: maxCols },
    () => `<w:gridCol w:w="${colWidth}"/>`
  ).join('');

  const rowsXml: string[] = [];
  for (const tr of trElements) {
    const cellElements = Array.from(tr.childNodes).filter((n): n is HTMLTableCellElement => {
      if (n.nodeType !== Node.ELEMENT_NODE) {
        return false;
      }
      const tag = (n as HTMLElement).tagName.toUpperCase();
      return tag === 'TD' || tag === 'TH';
    });

    if (cellElements.length === 0) {
      continue;
    }

    const cellsXml: string[] = [];
    for (const cell of cellElements) {
      const isHeader = cell.tagName.toUpperCase() === 'TH';
      const colspan = Number.parseInt(cell.getAttribute('colspan') ?? '1', 10);
      const validSpan = Number.isFinite(colspan) && colspan > 1 ? colspan : 1;
      const cellFormat = mergeElementRunFormat(cell, {
        ...inheritedFormat,
        bold: isHeader ? true : inheritedFormat.bold,
      });
      const cellAlign = extractAlignment(cell);

      const cellBlocks: string[] = [];
      if (hasDirectBlockChildren(cell)) {
        cellBlocks.push(
          ...convertContainerNodesToDocxXml(
            Array.from(cell.childNodes),
            ctx,
            cellFormat,
            cellAlign
          )
        );
      } else {
        const runs = collectInlineRunsXml(Array.from(cell.childNodes), cellFormat, ctx);
        cellBlocks.push(buildParagraphXml(runs, { align: cellAlign }));
      }

      // OOXML requires every table cell to contain at least one paragraph
      if (cellBlocks.length === 0) {
        cellBlocks.push('<w:p/>');
      }

      const tcPrParts: string[] = [
        `<w:tcW w:w="${colWidth * validSpan}" w:type="dxa"/>`,
      ];
      if (validSpan > 1) {
        tcPrParts.push(`<w:gridSpan w:val="${validSpan}"/>`);
      }
      if (isHeader) {
        tcPrParts.push('<w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>');
      }
      tcPrParts.push('<w:vAlign w:val="center"/>');

      cellsXml.push(`<w:tc><w:tcPr>${tcPrParts.join('')}</w:tcPr>${cellBlocks.join('')}</w:tc>`);
    }

    rowsXml.push(`<w:tr>${cellsXml.join('')}</w:tr>`);
  }

  if (rowsXml.length === 0) {
    return '';
  }

  return (
    '<w:tbl>' +
    '<w:tblPr>' +
    `<w:tblW w:w="${totalWidthTwips}" w:type="dxa"/>` +
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>' +
    '</w:tblBorders>' +
    '<w:tblCellMar>' +
    '<w:top w:w="80" w:type="dxa"/>' +
    '<w:left w:w="120" w:type="dxa"/>' +
    '<w:bottom w:w="80" w:type="dxa"/>' +
    '<w:right w:w="120" w:type="dxa"/>' +
    '</w:tblCellMar>' +
    '</w:tblPr>' +
    `<w:tblGrid>${tblGridCols}</w:tblGrid>` +
    rowsXml.join('') +
    '</w:tbl>'
  );
}

function convertListToDocxXml(
  listEl: HTMLElement,
  listType: 'bullet' | 'number',
  level: number,
  inheritedFormat: RunFormat,
  ctx: DocxExportContext
): string[] {
  const blocks: string[] = [];

  for (let i = 0; i < listEl.childNodes.length; i++) {
    const child = listEl.childNodes[i];
    if (child.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }
    const el = child as HTMLElement;
    const tag = el.tagName.toUpperCase();

    if (tag === 'UL' || tag === 'OL') {
      blocks.push(
        ...convertListToDocxXml(
          el,
          tag === 'OL' ? 'number' : 'bullet',
          Math.min(3, level + 1),
          inheritedFormat,
          ctx
        )
      );
      continue;
    }

    if (tag !== 'LI') {
      continue;
    }

    const liFormat = mergeElementRunFormat(el, inheritedFormat);
    const liAlign = extractAlignment(el);

    const inlineBuffer: Node[] = [];
    const nestedBlocks: string[] = [];

    const flushInlineAsListItem = (): void => {
      if (inlineBuffer.length === 0) {
        return;
      }
      if (!isWhitespaceOnlyNodes(inlineBuffer)) {
        const runs = collectInlineRunsXml(inlineBuffer, liFormat, ctx);
        blocks.push(
          buildParagraphXml(runs, {
            align: liAlign,
            listType,
            listLevel: level,
          })
        );
      }
      inlineBuffer.length = 0;
    };

    for (let j = 0; j < el.childNodes.length; j++) {
      const liChild = el.childNodes[j];
      if (liChild.nodeType === Node.ELEMENT_NODE) {
        const childEl = liChild as HTMLElement;
        const childTag = childEl.tagName.toUpperCase();
        if (childTag === 'UL' || childTag === 'OL') {
          flushInlineAsListItem();
          nestedBlocks.push(
            ...convertListToDocxXml(
              childEl,
              childTag === 'OL' ? 'number' : 'bullet',
              Math.min(3, level + 1),
              liFormat,
              ctx
            )
          );
          continue;
        }
        if (childTag === 'P' || childTag === 'DIV') {
          flushInlineAsListItem();
          const pRuns = collectInlineRunsXml(
            Array.from(childEl.childNodes),
            mergeElementRunFormat(childEl, liFormat),
            ctx
          );
          blocks.push(
            buildParagraphXml(pRuns, {
              align: extractAlignment(childEl) ?? liAlign,
              listType,
              listLevel: level,
            })
          );
          continue;
        }
      }
      inlineBuffer.push(liChild);
    }

    flushInlineAsListItem();
    blocks.push(...nestedBlocks);
  }

  return blocks;
}

function convertContainerNodesToDocxXml(
  nodes: Node[],
  ctx: DocxExportContext,
  inheritedFormat: RunFormat = {},
  inheritedAlign?: DocxAlign,
  isQuote = false
): string[] {
  const blocks: string[] = [];
  const inlineBuffer: Node[] = [];

  const flushInlineBuffer = (): void => {
    if (inlineBuffer.length === 0) {
      return;
    }
    if (!isWhitespaceOnlyNodes(inlineBuffer)) {
      const runs = collectInlineRunsXml(inlineBuffer, inheritedFormat, ctx);
      if (runs.length > 0) {
        blocks.push(buildParagraphXml(runs, { align: inheritedAlign, isQuote }));
      }
    }
    inlineBuffer.length = 0;
  };

  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      inlineBuffer.push(node);
      continue;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      continue;
    }

    const el = node as HTMLElement;
    const tag = el.tagName.toUpperCase();
    if (tag === 'SCRIPT' || tag === 'STYLE') {
      continue;
    }

    if (!DOCX_BLOCK_TAGS.has(tag)) {
      inlineBuffer.push(el);
      continue;
    }

    flushInlineBuffer();

    const nextFormat = mergeElementRunFormat(el, inheritedFormat);
    const align = extractAlignment(el) ?? inheritedAlign;

    if (tag === 'H1' || tag === 'H2' || tag === 'H3' || tag === 'H4' || tag === 'H5' || tag === 'H6') {
      const styleId: 'Heading1' | 'Heading2' | 'Heading3' =
        tag === 'H1' ? 'Heading1' : tag === 'H2' ? 'Heading2' : 'Heading3';
      const runs = collectInlineRunsXml(Array.from(el.childNodes), nextFormat, ctx);
      blocks.push(buildParagraphXml(runs, { styleId, align }));
      continue;
    }

    if (tag === 'UL' || tag === 'OL') {
      blocks.push(
        ...convertListToDocxXml(el, tag === 'OL' ? 'number' : 'bullet', 0, nextFormat, ctx)
      );
      continue;
    }

    if (tag === 'TABLE') {
      const tblXml = convertTableToDocxXml(el as HTMLTableElement, nextFormat, ctx);
      if (tblXml) {
        blocks.push(tblXml);
        blocks.push('<w:p/>');
      }
      continue;
    }

    if (tag === 'HR') {
      blocks.push(
        '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr></w:pPr></w:p>'
      );
      continue;
    }

    if (tag === 'BLOCKQUOTE') {
      if (hasDirectBlockChildren(el)) {
        blocks.push(
          ...convertContainerNodesToDocxXml(
            Array.from(el.childNodes),
            ctx,
            nextFormat,
            align,
            true
          )
        );
      } else {
        const runs = collectInlineRunsXml(Array.from(el.childNodes), nextFormat, ctx);
        blocks.push(buildParagraphXml(runs, { align, isQuote: true }));
      }
      continue;
    }

    // Generic block container (P, DIV, SECTION, ARTICLE, PRE, etc.)
    if (hasDirectBlockChildren(el)) {
      blocks.push(
        ...convertContainerNodesToDocxXml(
          Array.from(el.childNodes),
          ctx,
          nextFormat,
          align,
          isQuote
        )
      );
    } else {
      const runs = collectInlineRunsXml(Array.from(el.childNodes), nextFormat, ctx);
      blocks.push(buildParagraphXml(runs, { align, isQuote }));
    }
  }

  flushInlineBuffer();
  return blocks;
}

function buildDocxContentTypesXml(images: EmbeddedDocxImage[]): string {
  const mimeByExt = new Map<string, string>();
  for (const img of images) {
    if (!mimeByExt.has(img.ext)) {
      mimeByExt.set(img.ext, img.mime);
    }
  }
  const extraDefaults = Array.from(mimeByExt.entries())
    .map(([ext, mime]) => `  <Default Extension="${ext}" ContentType="${mime}"/>`)
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>${extraDefaults ? `\n${extraDefaults}` : ''}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;
}

const DOCX_ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function buildDocxDocumentRelsXml(images: EmbeddedDocxImage[]): string {
  const imageRels = images
    .map(
      (img) =>
        `  <Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.fileName}"/>`
    )
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>${imageRels ? `\n${imageRels}` : ''}
</Relationships>`;
}

const DOCX_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="1E293B"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="260" w:after="120"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="0F172A"/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="200" w:after="100"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="1E3A8A"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="160" w:after="80"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:color w:val="334155"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>`;

const DOCX_NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="◦"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="▪"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="3">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="3">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%4."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`;

const DOCX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const XLSX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const PPTX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const PDF_MIME_TYPE = 'application/pdf';

async function buildWordDocxBlob(htmlContent: string): Promise<Blob> {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(htmlContent || '', 'text/html');
  const ctx: DocxExportContext = {
    images: [],
    nextRelId: 3, // rId1 = styles.xml, rId2 = numbering.xml
    nextDocPrId: 1,
  };
  const bodyBlocks = convertContainerNodesToDocxXml(Array.from(parsed.body.childNodes), ctx);

  if (bodyBlocks.length === 0) {
    bodyBlocks.push('<w:p/>');
  }

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document' +
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:body>' +
    bodyBlocks.join('') +
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>' +
    '</w:body>' +
    '</w:document>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', buildDocxContentTypesXml(ctx.images));
  zip.file('_rels/.rels', DOCX_ROOT_RELS_XML);
  zip.file('word/_rels/document.xml.rels', buildDocxDocumentRelsXml(ctx.images));
  zip.file('word/styles.xml', DOCX_STYLES_XML);
  zip.file('word/numbering.xml', DOCX_NUMBERING_XML);
  zip.file('word/document.xml', documentXml);

  for (const img of ctx.images) {
    zip.file(`word/media/${img.fileName}`, img.bytes);
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME_TYPE,
  });
}

export async function exportWordToDocx(
  title: string,
  htmlContent: string
): Promise<SaveFileResult> {
  const blob = await buildWordDocxBlob(htmlContent);
  const safeTitle = sanitizeFileName(title, 'tai-lieu', /\.docx$/i);
  return triggerBlobDownload(blob, `${safeTitle}.docx`, DOCX_MIME_TYPE);
}

// ============================================================================
// 2. EXPORT EXCEL (.XLSX) — Real SheetJS Workbook with Live Formulas
// ============================================================================

function buildUniqueSheetName(rawName: string, index: number, usedLower: Set<string>): string {
  const cleaned = (rawName || '')
    .replace(/[:\\/?*\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^'+|'+$/g, '')
    .trim();

  const base = (cleaned || `Sheet ${index + 1}`).slice(0, 31).trim() || `Sheet ${index + 1}`;
  if (!usedLower.has(base.toLowerCase())) {
    usedLower.add(base.toLowerCase());
    return base;
  }

  let suffixNum = 2;
  while (suffixNum < 1000) {
    const suffix = ` (${suffixNum})`;
    const candidate = `${base.slice(0, 31 - suffix.length).trimEnd()}${suffix}`;
    if (!usedLower.has(candidate.toLowerCase())) {
      usedLower.add(candidate.toLowerCase());
      return candidate;
    }
    suffixNum++;
  }

  const fallback = `Sheet_${index + 1}`;
  usedLower.add(fallback.toLowerCase());
  return fallback;
}

function buildExcelWorkbook(sheets: SpreadsheetSheet[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const usedSheetNames = new Set<string>();

  const sourceSheets: SpreadsheetSheet[] =
    sheets.length > 0
      ? sheets
      : [
          {
            id: 'sheet-empty',
            name: 'Sheet 1',
            columns: ['A', 'B', 'C', 'D', 'E', 'F'],
            rows: [['', '', '', '', '', '']],
          },
        ];

  sourceSheets.forEach((sheet, sheetIdx) => {
    const safeRows = sheet.rows.map((row) =>
      row.map((cell) => {
        if (typeof cell === 'number') {
          return Number.isFinite(cell) ? cell : '';
        }
        return cell ?? '';
      })
    );

    const ws = XLSX.utils.aoa_to_sheet(safeRows);

    // Convert '=' strings into real SheetJS formula cells with accurate cached result type
    for (let r = 0; r < safeRows.length; r++) {
      const row = safeRows[r];
      for (let c = 0; c < row.length; c++) {
        const raw = row[c];
        if (typeof raw === 'string') {
          const trimmed = raw.trim();
          if (trimmed.startsWith('=') && trimmed.length > 1) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const formulaText = trimmed.slice(1).trim();
            const computed = evaluateCellValue(sheet.rows, r, c);
            let formulaCell: XLSX.CellObject;
            if (typeof computed === 'number' && Number.isFinite(computed)) {
              formulaCell = {
                t: 'n',
                f: formulaText,
                v: computed,
              };
            } else if (typeof computed === 'string' && !computed.startsWith('#')) {
              formulaCell = {
                t: 's',
                f: formulaText,
                v: computed,
              };
            } else {
              formulaCell = {
                t: 'n',
                f: formulaText,
              };
            }
            ws[addr] = formulaCell;
          }
        }
      }
    }

    // Preserve merged cell ranges
    if (Array.isArray(sheet.merges) && sheet.merges.length > 0) {
      ws['!merges'] = sheet.merges.map((m) => ({
        s: { r: m.s.r, c: m.s.c },
        e: { r: m.e.r, c: m.e.c },
      }));
    }

    // Preserve explicit column widths (px) & hidden columns
    const hiddenColSet = new Set<number>(
      Array.isArray(sheet.hiddenCols)
        ? sheet.hiddenCols.filter(
            (c): c is number => typeof c === 'number' && Number.isInteger(c) && c >= 0
          )
        : []
    );
    const maxHiddenCol =
      hiddenColSet.size > 0 ? Math.max(...Array.from(hiddenColSet)) + 1 : 0;
    const colCount = Math.max(
      sheet.columns.length,
      sheet.colWidths?.length ?? 0,
      maxHiddenCol,
      ...safeRows.map((row) => row.length),
      1
    );
    const colWidths: XLSX.ColInfo[] = [];
    for (let c = 0; c < colCount; c++) {
      const explicitPx = sheet.colWidths?.[c];
      const isHidden = hiddenColSet.has(c);
      let colInfo: XLSX.ColInfo;
      if (typeof explicitPx === 'number' && Number.isFinite(explicitPx) && explicitPx > 0) {
        colInfo = { wpx: Math.round(explicitPx) };
      } else {
        let maxLen = 10;
        for (let r = 0; r < safeRows.length; r++) {
          const val = safeRows[r]?.[c];
          if (val !== undefined && val !== null) {
            const len = String(val).length;
            if (len > maxLen) {
              maxLen = len;
            }
          }
        }
        colInfo = { wch: Math.min(40, maxLen + 3) };
      }
      if (isHidden) {
        colInfo.hidden = true;
      }
      colWidths.push(colInfo);
    }
    ws['!cols'] = colWidths;

    // Preserve explicit row heights (px) & hidden rows
    const hiddenRowSet = new Set<number>(
      Array.isArray(sheet.hiddenRows)
        ? sheet.hiddenRows.filter(
            (r): r is number => typeof r === 'number' && Number.isInteger(r) && r >= 0
          )
        : []
    );
    const maxHiddenRow =
      hiddenRowSet.size > 0 ? Math.max(...Array.from(hiddenRowSet)) + 1 : 0;
    const rowMetaCount = Math.max(sheet.rowHeights?.length ?? 0, maxHiddenRow);
    if (rowMetaCount > 0) {
      const rowMeta: XLSX.RowInfo[] = [];
      for (let r = 0; r < rowMetaCount; r++) {
        const hPx = sheet.rowHeights?.[r];
        const info: XLSX.RowInfo = {};
        if (typeof hPx === 'number' && Number.isFinite(hPx) && hPx > 0) {
          info.hpx = Math.round(hPx);
        }
        if (hiddenRowSet.has(r)) {
          info.hidden = true;
        }
        rowMeta.push(info);
      }
      ws['!rows'] = rowMeta;
    }

    const uniqueName = buildUniqueSheetName(sheet.name, sheetIdx, usedSheetNames);
    XLSX.utils.book_append_sheet(wb, ws, uniqueName);
  });

  return wb;
}

function buildExcelBlob(sheets: SpreadsheetSheet[]): Blob {
  const wb = buildExcelWorkbook(sheets);
  const output = XLSX.write(wb, {
    type: 'array',
    bookType: 'xlsx',
  }) as ArrayBuffer | Uint8Array;
  const bytes =
    output instanceof Uint8Array ? new Uint8Array(output) : new Uint8Array(output);
  return new Blob([bytes], { type: XLSX_MIME_TYPE });
}

export async function exportExcelToXlsx(
  title: string,
  sheets: SpreadsheetSheet[]
): Promise<SaveFileResult> {
  const blob = buildExcelBlob(sheets);
  const safeTitle = sanitizeFileName(title, 'bang-tinh', /\.xlsx$/i);
  return triggerBlobDownload(blob, `${safeTitle}.xlsx`, XLSX_MIME_TYPE);
}

// ============================================================================
// 3. EXPORT POWERPOINT (.PPTX) — Real Deck via PptxGenJS
// ============================================================================

function normalizePptxHex(color: string | undefined, fallback: string): string {
  const parsed = cssColorToHex6(color);
  return parsed ?? fallback;
}

function getRelativeLuminance(hex6: string): number {
  const r = Number.parseInt(hex6.slice(0, 2), 16) / 255;
  const g = Number.parseInt(hex6.slice(2, 4), 16) / 255;
  const b = Number.parseInt(hex6.slice(4, 6), 16) / 255;
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

async function imageUrlToDataUrl(src: string): Promise<string | null> {
  if (!src) {
    return null;
  }
  if (src.startsWith('data:image/')) {
    return src;
  }
  try {
    const response = await fetch(src);
    if (!response.ok) {
      return null;
    }
    const blob = await response.blob();
    return await new Promise<string | null>((resolve) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        resolve(typeof reader.result === 'string' ? reader.result : null);
      };
      reader.onerror = () => resolve(null);
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

function isShapeBlock(block: SlideTextBlock): boolean {
  return block.kind === 'shape' || Boolean(block.shapeKind);
}

function getEffectiveSlideBlocks(slide: SlidePreviewItem): SlideTextBlock[] {
  if (slide.blocks && slide.blocks.length > 0) {
    return slide.blocks.filter(
      (b) =>
        (typeof b.text === 'string' && b.text.trim().length > 0) ||
        isShapeBlock(b) ||
        (b.kind === 'table' && Array.isArray(b.tableRows) && b.tableRows.length > 0) ||
        (b.kind === 'chart' && Boolean(b.chartData && b.chartData.series.length > 0)) ||
        (b.kind === 'media' && Boolean(b.mediaSrc))
    );
  }

  const fallback: SlideTextBlock[] = [];
  if (slide.title?.trim()) {
    fallback.push({
      id: `${slide.id}-title`,
      text: slide.title.trim(),
      fontSize: 30,
      bold: true,
      align: 'left',
      role: 'title',
    });
  }
  if (slide.subtitle?.trim()) {
    fallback.push({
      id: `${slide.id}-subtitle`,
      text: slide.subtitle.trim(),
      fontSize: 18,
      align: 'left',
      role: 'subtitle',
    });
  }
  if (slide.bullets && slide.bullets.length > 0) {
    const cleanBullets = slide.bullets.map((line) => line.trim()).filter(Boolean);
    if (cleanBullets.length > 0) {
      fallback.push({
        id: `${slide.id}-body`,
        text: cleanBullets.map((line) => `• ${line}`).join('\n'),
        fontSize: 15,
        align: 'left',
        role: 'body',
      });
    }
  }
  return fallback;
}

function mapShapeKindToPptxShape(
  pptx: PptxGenJS,
  shapeKind?: SlideShapeKind
): PptxGenJS.SHAPE_NAME {
  switch (shapeKind) {
    case 'rect':
      return pptx.ShapeType.rect;
    case 'roundRect':
      return pptx.ShapeType.roundRect;
    case 'ellipse':
      return pptx.ShapeType.ellipse;
    case 'triangle':
      return pptx.ShapeType.triangle;
    case 'arrowRight':
      return pptx.ShapeType.rightArrow;
    case 'arrowLeft':
      return pptx.ShapeType.leftArrow;
    case 'arrowUp':
      return pptx.ShapeType.upArrow;
    case 'arrowDown':
      return pptx.ShapeType.downArrow;
    case 'star':
      return pptx.ShapeType.star5;
    case 'callout':
      return pptx.ShapeType.wedgeRoundRectCallout;
    case 'line':
      return pptx.ShapeType.line;
    default:
      return pptx.ShapeType.rect;
  }
}

function buildSafeHyperlink(url?: string): PptxGenJS.HyperlinkProps | undefined {
  if (!url) {
    return undefined;
  }
  const trimmed = url.trim();
  if (/^https?:\/\//i.test(trimmed) || /^mailto:/i.test(trimmed)) {
    return { url: trimmed };
  }
  return undefined;
}

function addShapeBlockToPptxSlide(
  pptx: PptxGenJS,
  slideObj: PptxGenJS.Slide,
  block: SlideTextBlock,
  box: { x: number; y: number; w: number; h: number },
  defaultTextColor: string,
  accentHex: string
): void {
  const shapeKind = block.shapeKind ?? 'rect';
  const shapeName = mapShapeKindToPptxShape(pptx, shapeKind);
  const fillHex = block.fillColor ? cssColorToHex6(block.fillColor) : undefined;
  const borderHex = block.borderColor ? cssColorToHex6(block.borderColor) : undefined;
  const lineWidth =
    typeof block.borderWidthPx === 'number' &&
    Number.isFinite(block.borderWidthPx) &&
    block.borderWidthPx > 0
      ? Math.max(0.5, Math.min(12, block.borderWidthPx * 0.75))
      : undefined;
  const rotate =
    typeof block.rotationDeg === 'number' && Number.isFinite(block.rotationDeg)
      ? Math.max(-360, Math.min(360, block.rotationDeg))
      : undefined;
  const hyperlink = buildSafeHyperlink(block.linkUrl);

  try {
    if (shapeKind === 'line') {
      slideObj.addShape(pptx.ShapeType.line, {
        x: box.x,
        y: box.y + box.h / 2,
        w: box.w,
        h: 0,
        line: {
          color: borderHex || fillHex || accentHex || '475569',
          width: lineWidth ?? 2,
        },
        ...(rotate !== undefined ? { rotate } : {}),
        ...(hyperlink ? { hyperlink } : {}),
      });
      return;
    }

    const effectiveFillHex = fillHex ?? (!borderHex ? accentHex : undefined);
    const shapeFill: PptxGenJS.ShapeFillProps | undefined = effectiveFillHex
      ? { color: effectiveFillHex }
      : undefined;
    const shapeLine: PptxGenJS.ShapeLineProps = borderHex
      ? { color: borderHex, width: lineWidth ?? 1.5 }
      : { type: 'none' };
    const shapeText = typeof block.text === 'string' ? block.text.trim() : '';

    if (shapeText.length > 0) {
      const autoContrastText = effectiveFillHex
        ? getRelativeLuminance(effectiveFillHex) < 0.45
          ? 'FFFFFF'
          : '0F172A'
        : defaultTextColor;
      const textColor = normalizePptxHex(block.color, autoContrastText);
      slideObj.addText(block.text, {
        shape: shapeName,
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
        fill: shapeFill,
        line: shapeLine,
        fontFace: block.fontFamily || 'Arial',
        fontSize: Math.max(8, Math.min(72, Math.round(block.fontSize || 14))),
        bold: Boolean(block.bold),
        italic: Boolean(block.italic),
        color: textColor,
        align: block.align ?? 'center',
        valign: block.valign ?? 'middle',
        wrap: true,
        margin: 0.06,
        ...(shapeKind === 'roundRect' ? { rectRadius: 0.16 } : {}),
        ...(rotate !== undefined ? { rotate } : {}),
        ...(hyperlink ? { hyperlink } : {}),
      });
    } else {
      slideObj.addShape(shapeName, {
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
        fill: shapeFill,
        line: shapeLine,
        ...(shapeKind === 'roundRect' ? { rectRadius: 0.16 } : {}),
        ...(rotate !== undefined ? { rotate } : {}),
        ...(hyperlink ? { hyperlink } : {}),
      });
    }
  } catch {
    // Fallback cleanly to basic rectangle so export never crashes PptxGenJS
    try {
      slideObj.addShape(pptx.ShapeType.rect, {
        x: box.x,
        y: box.y,
        w: Math.max(0.2, box.w),
        h: Math.max(0.1, box.h),
        fill: { color: fillHex ?? accentHex },
        line: borderHex ? { color: borderHex, width: 1 } : { type: 'none' },
      });
    } catch {
      // Ignore unrecoverable shape error
    }
  }
}

function addChartBlockToPptxSlide(
  pptx: PptxGenJS,
  slideObj: PptxGenJS.Slide,
  block: SlideTextBlock,
  box: { x: number; y: number; w: number; h: number },
  accentHex: string
): void {
  const chart = block.chartData;
  if (!chart || !Array.isArray(chart.series) || chart.series.length === 0) {
    return;
  }
  try {
    const categories =
      Array.isArray(chart.categories) && chart.categories.length > 0
        ? chart.categories.map((c) => String(c))
        : chart.series[0].values.map((_, idx) => `Mục ${idx + 1}`);
    const safeSeries = chart.series.map((s, idx) => ({
      name: s.name || `Chuỗi ${idx + 1}`,
      labels: categories,
      values: categories.map((_, valIdx) => {
        const n = s.values[valIdx];
        return typeof n === 'number' && Number.isFinite(n) ? n : 0;
      }),
    }));
    const chartColors = chart.series.map((s) => normalizePptxHex(s.color, accentHex));
    const pptxChartType =
      chart.chartType === 'line'
        ? pptx.ChartType.line
        : chart.chartType === 'pie'
          ? pptx.ChartType.pie
          : pptx.ChartType.bar;

    slideObj.addChart(pptxChartType, safeSeries, {
      x: box.x,
      y: box.y,
      w: Math.max(1.2, box.w),
      h: Math.max(1.0, box.h),
      ...(chart.title ? { title: chart.title, showTitle: true } : { showTitle: false }),
      showLegend: true,
      legendPos: 'b',
      ...(chart.chartType === 'column' ? { barDir: 'col' } : {}),
      chartColors,
    });
  } catch {
    // Fallback cleanly to chart title placeholder if chart data is malformed
    try {
      slideObj.addText(chart.title || block.text || 'Biểu đồ', {
        shape: pptx.ShapeType.roundRect,
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
        fill: { color: 'F8FAFC' },
        line: { color: 'CBD5E1', width: 1 },
        color: '334155',
        fontSize: 13,
        align: 'center',
        valign: 'middle',
      });
    } catch {
      // Ignore fallback error
    }
  }
}

function addMediaBlockPlaceholderToPptxSlide(
  pptx: PptxGenJS,
  slideObj: PptxGenJS.Slide,
  block: SlideTextBlock,
  box: { x: number; y: number; w: number; h: number }
): void {
  const label =
    (typeof block.text === 'string' && block.text.trim()) ||
    (block.mediaKind === 'audio' ? 'Âm thanh nhúng' : 'Video nhúng');
  const hyperlink = buildSafeHyperlink(block.mediaSrc || block.linkUrl);
  try {
    slideObj.addText(label, {
      shape: pptx.ShapeType.roundRect,
      x: box.x,
      y: box.y,
      w: box.w,
      h: box.h,
      fill: { color: '1E293B' },
      line: { color: '475569', width: 1 },
      color: 'F8FAFC',
      fontSize: 13,
      bold: true,
      align: 'center',
      valign: 'middle',
      wrap: true,
      ...(hyperlink ? { hyperlink } : {}),
    });
  } catch {
    // Ignore media placeholder failure
  }
}

function resolveSlideBackgroundHex(
  slide: SlidePreviewItem,
  blocks: SlideTextBlock[]
): string {
  if (slide.backgroundColor) {
    return normalizePptxHex(slide.backgroundColor, '181D27');
  }
  // Infer light vs dark background from explicit block colors so text always remains readable
  const explicitLuminances: number[] = [];
  for (const block of blocks) {
    const hex = cssColorToHex6(block.color);
    if (hex) {
      explicitLuminances.push(getRelativeLuminance(hex));
    }
  }
  if (explicitLuminances.length > 0) {
    const avgLum =
      explicitLuminances.reduce((acc, l) => acc + l, 0) / explicitLuminances.length;
    return avgLum > 0.65 ? '181D27' : 'FFFFFF';
  }
  return '181D27';
}

function estimateBlockHeightInches(text: string, fontSizePt: number, boxWidthInches: number): number {
  const safeFont = Math.min(48, Math.max(10, fontSizePt || 16));
  // Average character width ~ 0.52 * fontSize in points; 72 pt = 1 inch
  const charsPerLine = Math.max(14, Math.floor((boxWidthInches * 72) / (safeFont * 0.54)));
  const rawLines = text.split('\n');
  let totalLines = 0;
  for (const line of rawLines) {
    totalLines += Math.max(1, Math.ceil(line.length / charsPerLine));
  }
  const lineHeightInches = (safeFont / 72) * 1.38;
  return Math.max(0.42, totalLines * lineHeightInches + 0.16);
}

function addTableBlockToPptxSlide(
  slideObj: PptxGenJS.Slide,
  block: SlideTextBlock,
  box: { x: number; y: number; w: number; h: number }
): void {
  const rawRows = block.tableRows ?? [];
  if (rawRows.length === 0) {
    return;
  }
  const maxCols = Math.max(1, ...rawRows.map((row) => row.length));
  const textColor = normalizePptxHex(block.color, '0F172A');
  const fontFace = block.fontFamily || 'Arial';
  const fontSize = Math.max(9, Math.min(24, block.fontSize || 12));

  const tableRows: PptxGenJS.TableRow[] = rawRows.map((row, rowIdx) => {
    const padded = Array.from({ length: maxCols }, (_, colIdx) => row[colIdx] ?? '');
    return padded.map((cellText) => ({
      text: cellText,
      options: {
        fontFace,
        fontSize,
        bold: rowIdx === 0,
        color: textColor,
        fill: {
          color: rowIdx === 0 ? 'E2E8F0' : rowIdx % 2 === 1 ? 'FFFFFF' : 'F8FAFC',
        },
        border: { type: 'solid', pt: 0.75, color: 'CBD5E1' },
        valign: 'middle',
      },
    }));
  });

  slideObj.addTable(tableRows, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
  });
}

async function buildPptxPresentation(
  title: string,
  slides: SlidePreviewItem[],
  slideAspectRatio?: number
): Promise<PptxGenJS> {
  const pptx = new PptxGenJS();
  const effectiveRatio = slideAspectRatio ?? slides[0]?.slideWidthRatio ?? 1.77;
  const isStandard4x3 = effectiveRatio < 1.5;
  pptx.layout = isStandard4x3 ? 'LAYOUT_4x3' : 'LAYOUT_WIDE';
  const slideW = isStandard4x3 ? 10 : 13.333;
  const slideH = 7.5;
  pptx.title = title || 'Bài thuyết trình';
  pptx.author = 'LITE Workspace';

  const sourceSlides: SlidePreviewItem[] =
    slides.length > 0
      ? slides
      : [
          {
            id: 'empty-slide-1',
            index: 1,
            title: title || 'Slide 1',
            bullets: [],
            accentColor: '#4F46E5',
          },
        ];

  for (const slide of sourceSlides) {
    const slideObj = pptx.addSlide();
    if (slide.hidden === true) {
      slideObj.hidden = true;
    }
    const blocks = getEffectiveSlideBlocks(slide);
    const bgHex = resolveSlideBackgroundHex(slide, blocks);
    const isDarkBg = getRelativeLuminance(bgHex) < 0.45;
    const defaultTextColor = isDarkBg ? 'F8FAFC' : '0F172A';
    const defaultMutedColor = isDarkBg ? 'CBD5E1' : '475569';
    const accentHex = normalizePptxHex(slide.accentColor, '4F46E5');

    // Slide background: image if present, else hex color
    let appliedBgImage = false;
    if (slide.backgroundImageSrc) {
      try {
        const bgData = await imageUrlToDataUrl(slide.backgroundImageSrc);
        if (bgData) {
          slideObj.background = { data: bgData };
          appliedBgImage = true;
        }
      } catch {
        // Fallback cleanly to solid background color
      }
    }
    if (!appliedBgImage) {
      slideObj.background = { color: bgHex };
    }

    // Resolve embedded slide images cleanly while keeping original coordinate metadata
    const resolvedImages: Array<{
      data: string;
      alt?: string;
      xPct?: number;
      yPct?: number;
      wPct?: number;
      hPct?: number;
    }> = [];
    if (slide.images && slide.images.length > 0) {
      for (const img of slide.images) {
        try {
          const dataUrl = await imageUrlToDataUrl(img.src);
          if (dataUrl) {
            resolvedImages.push({
              data: dataUrl,
              alt: img.alt,
              xPct: img.xPct,
              yPct: img.yPct,
              wPct: img.wPct,
              hPct: img.hPct,
            });
          }
        } catch {
          // Ignore single unreadable image without failing export
        }
      }
    }

    const positionedBlocks = blocks
      .filter((b) => typeof b.xPct === 'number' && typeof b.yPct === 'number')
      .map((block, index) => ({ block, index }))
      .sort((a, b) => {
        const zDiff = (a.block.zIndex ?? 0) - (b.block.zIndex ?? 0);
        return zDiff !== 0 ? zDiff : a.index - b.index;
      })
      .map((item) => item.block);
    const unpositionedBlocks = blocks.filter(
      (b) => typeof b.xPct !== 'number' || typeof b.yPct !== 'number'
    );
    const positionedImages = resolvedImages.filter(
      (img) => typeof img.xPct === 'number' && typeof img.yPct === 'number'
    );
    const unpositionedImages = resolvedImages.filter(
      (img) => typeof img.xPct !== 'number' || typeof img.yPct !== 'number'
    );

    const hasPositionedContent =
      Boolean(slide.isMasterPositioned) ||
      positionedBlocks.length > 0 ||
      positionedImages.length > 0;

    // Decorative accent bar only for auto-layout slides without master positioning or bg image
    if (!hasPositionedContent && !appliedBgImage) {
      slideObj.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: slideW,
        h: 0.1,
        fill: { color: accentHex },
        line: { type: 'none' },
      });
    }

    // 1. Export explicitly positioned images (exact % -> inch mapping)
    for (const img of positionedImages) {
      const x = ((img.xPct ?? 0) / 100) * slideW;
      const y = ((img.yPct ?? 0) / 100) * slideH;
      const w = Math.max(0.3, ((img.wPct ?? 30) / 100) * slideW);
      const h = Math.max(0.3, ((img.hPct ?? 30) / 100) * slideH);
      try {
        slideObj.addImage({
          data: img.data,
          x,
          y,
          w,
          h,
          sizing: { type: 'contain', w, h },
          altText: img.alt,
        });
      } catch {
        // Skip invalid image safely
      }
    }

    // 2. Export explicitly positioned text, shape, table, chart & media blocks (exact % -> inch mapping)
    for (const block of positionedBlocks) {
      const x = ((block.xPct ?? 0) / 100) * slideW;
      const y = ((block.yPct ?? 0) / 100) * slideH;
      const w = Math.max(0.4, ((block.wPct ?? 40) / 100) * slideW);
      const h = Math.max(
        block.shapeKind === 'line' ? 0.05 : 0.25,
        ((block.hPct ?? 12) / 100) * slideH
      );

      if (block.kind === 'table' && block.tableRows && block.tableRows.length > 0) {
        addTableBlockToPptxSlide(slideObj, block, { x, y, w, h });
        continue;
      }

      if (block.kind === 'chart' && block.chartData && block.chartData.series.length > 0) {
        addChartBlockToPptxSlide(pptx, slideObj, block, { x, y, w, h }, accentHex);
        continue;
      }

      if (block.kind === 'media') {
        addMediaBlockPlaceholderToPptxSlide(pptx, slideObj, block, { x, y, w, h });
        continue;
      }

      if (isShapeBlock(block)) {
        addShapeBlockToPptxSlide(
          pptx,
          slideObj,
          block,
          { x, y, w, h },
          defaultTextColor,
          accentHex
        );
        continue;
      }

      const fallbackColor =
        block.role === 'subtitle' ? defaultMutedColor : defaultTextColor;
      const textColor = normalizePptxHex(block.color, fallbackColor);
      const fillHex = block.fillColor ? cssColorToHex6(block.fillColor) : undefined;
      const borderHex = block.borderColor ? cssColorToHex6(block.borderColor) : undefined;
      const rotate =
        typeof block.rotationDeg === 'number' && Number.isFinite(block.rotationDeg)
          ? Math.max(-360, Math.min(360, block.rotationDeg))
          : undefined;
      const hyperlink = buildSafeHyperlink(block.linkUrl);

      slideObj.addText(block.text ?? '', {
        x,
        y,
        w,
        h,
        fontFace: block.fontFamily || 'Arial',
        fontSize: Math.max(8, Math.min(96, Math.round(block.fontSize || 16))),
        bold: Boolean(block.bold),
        italic: Boolean(block.italic),
        color: textColor,
        align: block.align ?? 'left',
        valign: block.valign ?? 'top',
        fill: fillHex ? { color: fillHex } : undefined,
        line: borderHex ? { color: borderHex, width: 1 } : undefined,
        wrap: true,
        margin: 0.04,
        ...(rotate !== undefined ? { rotate } : {}),
        ...(hyperlink ? { hyperlink } : {}),
      });
    }

    // 3. Auto-layout fallback for unpositioned blocks & images
    const hasAutoImages = unpositionedImages.length > 0;
    const hasAutoBlocks = unpositionedBlocks.length > 0;

    if (hasAutoBlocks) {
      const textX = slideW * 0.055;
      const textW = hasAutoImages ? slideW * 0.555 : slideW * 0.89;
      const startY = slideH * 0.075;
      const maxAvailableH = slideH * 0.86;
      const gap = 0.16;

      const desiredHeights = unpositionedBlocks.map((block) => {
        if (block.kind === 'table' && block.tableRows) {
          return Math.max(1.2, Math.min(3.5, block.tableRows.length * 0.42));
        }
        if (block.kind === 'chart') {
          return 2.6;
        }
        if (block.kind === 'media') {
          return 1.4;
        }
        if (isShapeBlock(block)) {
          return block.shapeKind === 'line'
            ? 0.24
            : Math.max(
                0.75,
                block.text?.trim()
                  ? estimateBlockHeightInches(block.text, block.fontSize, textW)
                  : 1.0
              );
        }
        return estimateBlockHeightInches(block.text ?? '', block.fontSize, textW);
      });
      const totalDesiredH =
        desiredHeights.reduce((sum, h) => sum + h, 0) +
        gap * Math.max(0, unpositionedBlocks.length - 1);
      const scale = totalDesiredH > maxAvailableH ? maxAvailableH / totalDesiredH : 1;

      let cursorY = startY;
      unpositionedBlocks.forEach((block, idx) => {
        const scaledH = Math.max(
          block.shapeKind === 'line' ? 0.15 : 0.35,
          desiredHeights[idx] * scale
        );
        if (block.kind === 'table' && block.tableRows && block.tableRows.length > 0) {
          addTableBlockToPptxSlide(slideObj, block, {
            x: textX,
            y: cursorY,
            w: textW,
            h: scaledH,
          });
        } else if (
          block.kind === 'chart' &&
          block.chartData &&
          block.chartData.series.length > 0
        ) {
          addChartBlockToPptxSlide(
            pptx,
            slideObj,
            block,
            { x: textX, y: cursorY, w: textW, h: scaledH },
            accentHex
          );
        } else if (block.kind === 'media') {
          addMediaBlockPlaceholderToPptxSlide(pptx, slideObj, block, {
            x: textX,
            y: cursorY,
            w: textW,
            h: scaledH,
          });
        } else if (isShapeBlock(block)) {
          addShapeBlockToPptxSlide(
            pptx,
            slideObj,
            block,
            { x: textX, y: cursorY, w: textW, h: scaledH },
            defaultTextColor,
            accentHex
          );
        } else {
          const scaledFont = Math.max(
            10,
            Math.round((block.fontSize || 16) * (scale < 0.88 ? Math.max(0.72, scale) : 1))
          );
          const fallbackColor =
            block.role === 'subtitle' ? defaultMutedColor : defaultTextColor;
          const textColor = normalizePptxHex(block.color, fallbackColor);
          const fillHex = block.fillColor ? cssColorToHex6(block.fillColor) : undefined;
          const borderHex = block.borderColor ? cssColorToHex6(block.borderColor) : undefined;
          const rotate =
            typeof block.rotationDeg === 'number' && Number.isFinite(block.rotationDeg)
              ? Math.max(-360, Math.min(360, block.rotationDeg))
              : undefined;
          const hyperlink = buildSafeHyperlink(block.linkUrl);

          slideObj.addText(block.text ?? '', {
            x: textX,
            y: cursorY,
            w: textW,
            h: scaledH,
            fontFace: block.fontFamily || 'Arial',
            fontSize: scaledFont,
            bold: Boolean(block.bold),
            italic: Boolean(block.italic),
            color: textColor,
            align: block.align ?? 'left',
            valign: block.valign ?? 'top',
            fill: fillHex ? { color: fillHex } : undefined,
            line: borderHex ? { color: borderHex, width: 1 } : undefined,
            wrap: true,
            margin: 0.04,
            ...(rotate !== undefined ? { rotate } : {}),
            ...(hyperlink ? { hyperlink } : {}),
          });
        }

        cursorY += scaledH + gap * scale;
      });
    }

    if (hasAutoImages) {
      const imgAreaX = hasAutoBlocks ? slideW * 0.635 : slideW * 0.08;
      const imgAreaY = slideH * 0.09;
      const imgAreaW = hasAutoBlocks ? slideW * 0.315 : slideW * 0.84;
      const imgAreaH = slideH * 0.82;
      const count = unpositionedImages.length;

      if (count === 1) {
        try {
          slideObj.addImage({
            data: unpositionedImages[0].data,
            x: imgAreaX,
            y: imgAreaY,
            w: imgAreaW,
            h: imgAreaH,
            sizing: { type: 'contain', w: imgAreaW, h: imgAreaH },
            altText: unpositionedImages[0].alt,
          });
        } catch {
          // Skip invalid image format safely
        }
      } else {
        const cols = hasAutoBlocks ? 1 : Math.min(2, count);
        const rows = Math.ceil(count / cols);
        const cellGap = 0.18;
        const cellW = (imgAreaW - cellGap * (cols - 1)) / cols;
        const cellH = (imgAreaH - cellGap * (rows - 1)) / rows;

        unpositionedImages.forEach((img, idx) => {
          const r = Math.floor(idx / cols);
          const c = idx % cols;
          try {
            slideObj.addImage({
              data: img.data,
              x: imgAreaX + c * (cellW + cellGap),
              y: imgAreaY + r * (cellH + cellGap),
              w: cellW,
              h: cellH,
              sizing: { type: 'contain', w: cellW, h: cellH },
              altText: img.alt,
            });
          } catch {
            // Skip invalid image format safely
          }
        });
      }
    }

    // 4. Speaker notes
    if (slide.speakerNote && slide.speakerNote.trim()) {
      slideObj.addNotes(slide.speakerNote.trim());
    }
  }

  return pptx;
}

async function buildPptxBlob(
  title: string,
  slides: SlidePreviewItem[],
  slideAspectRatio?: number
): Promise<Blob> {
  const pptx = await buildPptxPresentation(title, slides, slideAspectRatio);
  const rawOut = await pptx.write({ outputType: 'blob' });
  if (rawOut instanceof Blob) {
    return rawOut.type === PPTX_MIME_TYPE
      ? rawOut
      : new Blob([rawOut], { type: PPTX_MIME_TYPE });
  }
  if (rawOut instanceof Uint8Array) {
    return new Blob([new Uint8Array(rawOut)], { type: PPTX_MIME_TYPE });
  }
  return new Blob([rawOut], { type: PPTX_MIME_TYPE });
}

export async function exportSlidesToPptx(
  title: string,
  slides: SlidePreviewItem[],
  slideAspectRatio?: number
): Promise<SaveFileResult> {
  const blob = await buildPptxBlob(title, slides, slideAspectRatio);
  const safeTitle = sanitizeFileName(title, 'bai-thuyet-trinh', /\.pptx$/i);
  return triggerBlobDownload(blob, `${safeTitle}.pptx`, PPTX_MIME_TYPE);
}

// ============================================================================
// 4. BUILD UPLOAD PAYLOAD — Unified Blob Builder for Cloud / Google Drive
// ============================================================================

function buildFallbackWordHtmlFromDoc(doc: DocumentItem): string {
  const parts: string[] = [];
  if (doc.title.trim()) {
    parts.push(`<h1>${escapeXmlText(doc.title.trim())}</h1>`);
  }
  if (doc.contentPreview.trim()) {
    parts.push(`<p>${escapeXmlText(doc.contentPreview.trim())}</p>`);
  }
  for (const section of doc.outline) {
    const heading = section.romanIndex
      ? `${section.romanIndex}. ${section.title}`
      : section.title;
    parts.push(`<h2>${escapeXmlText(heading)}</h2>`);
    const customParas = doc.wordParagraphs?.[section.id];
    if (customParas && customParas.length > 0) {
      for (const para of customParas) {
        if (para.trim()) {
          parts.push(`<p>${escapeXmlText(para.trim())}</p>`);
        }
      }
    } else if (section.summary.trim()) {
      parts.push(`<p>${escapeXmlText(section.summary.trim())}</p>`);
    }
    if (section.keyPoints.length > 0) {
      const items = section.keyPoints
        .filter((pt) => pt.trim().length > 0)
        .map((pt) => `<li>${escapeXmlText(pt.trim())}</li>`)
        .join('');
      if (items) {
        parts.push(`<ul>${items}</ul>`);
      }
    }
  }
  return parts.join('') || '<p></p>';
}

export async function buildDocumentUploadPayload(
  doc: DocumentItem
): Promise<{ blob: Blob; fileName: string; mimeType: string }> {
  if (doc.type === 'docx') {
    const safeTitle = sanitizeFileName(doc.title, 'tai-lieu', /\.docx$/i);
    const fileName = `${safeTitle}.docx`;

    if (!doc.hasLocalEdits && doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
      const mimeType = doc.originalMimeType || DOCX_MIME_TYPE;
      const blob = new Blob([doc.rawArrayBuffer], { type: mimeType });
      return { blob, fileName, mimeType };
    }
    if (doc.wordHtml && doc.wordHtml.trim().length > 0) {
      const blob = await buildWordDocxBlob(doc.wordHtml);
      return { blob, fileName, mimeType: DOCX_MIME_TYPE };
    }
    if (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
      const blob = new Blob([doc.rawArrayBuffer], { type: DOCX_MIME_TYPE });
      return { blob, fileName, mimeType: DOCX_MIME_TYPE };
    }
    const fallbackHtml = buildFallbackWordHtmlFromDoc(doc);
    const blob = await buildWordDocxBlob(fallbackHtml);
    return { blob, fileName, mimeType: DOCX_MIME_TYPE };
  }

  if (doc.type === 'xlsx') {
    const safeTitle = sanitizeFileName(doc.title, 'bang-tinh', /\.xlsx$/i);
    const fileName = `${safeTitle}.xlsx`;

    if (!doc.hasLocalEdits && doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
      const mimeType = doc.originalMimeType || XLSX_MIME_TYPE;
      const blob = new Blob([doc.rawArrayBuffer], { type: mimeType });
      return { blob, fileName, mimeType };
    }
    const blob = buildExcelBlob(doc.sheets ?? []);
    return { blob, fileName, mimeType: XLSX_MIME_TYPE };
  }

  if (doc.type === 'pptx') {
    const safeTitle = sanitizeFileName(doc.title, 'bai-thuyet-trinh', /\.pptx$/i);
    const fileName = `${safeTitle}.pptx`;

    if (!doc.hasLocalEdits && doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
      const mimeType = doc.originalMimeType || PPTX_MIME_TYPE;
      const blob = new Blob([doc.rawArrayBuffer], { type: mimeType });
      return { blob, fileName, mimeType };
    }
    const blob = await buildPptxBlob(
      doc.title,
      doc.slidesPreview ?? [],
      doc.slideAspectRatio
    );
    return { blob, fileName, mimeType: PPTX_MIME_TYPE };
  }

  // PDF
  const safeTitle = sanitizeFileName(doc.title, 'tai-lieu', /\.pdf$/i);
  const fileName = `${safeTitle}.pdf`;

  if (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
    const blob = new Blob([doc.rawArrayBuffer], { type: PDF_MIME_TYPE });
    return { blob, fileName, mimeType: PDF_MIME_TYPE };
  }

  if (doc.sourceFile) {
    const blob =
      doc.sourceFile.type === PDF_MIME_TYPE
        ? doc.sourceFile
        : new Blob([doc.sourceFile], { type: PDF_MIME_TYPE });
    return { blob, fileName, mimeType: PDF_MIME_TYPE };
  }

  const pdfUrl = doc.fileUrl ?? doc.objectUrl;
  if (pdfUrl) {
    const response = await fetch(pdfUrl);
    if (!response.ok) {
      throw new Error(
        `Không thể đọc dữ liệu tệp PDF gốc (HTTP ${response.status}).`
      );
    }
    const fetchedBlob = await response.blob();
    const blob =
      fetchedBlob.type === PDF_MIME_TYPE
        ? fetchedBlob
        : new Blob([fetchedBlob], { type: PDF_MIME_TYPE });
    return { blob, fileName, mimeType: PDF_MIME_TYPE };
  }

  throw new Error(
    'Tài liệu PDF mẫu chưa có tệp PDF nhị phân gốc để tải lên Google Drive. Vui lòng mở tệp .pdf thật từ máy của bạn.'
  );
}

function getDefaultMimeTypeForDocType(type: DocumentItem['type']): string {
  switch (type) {
    case 'docx':
      return DOCX_MIME_TYPE;
    case 'xlsx':
      return XLSX_MIME_TYPE;
    case 'pptx':
      return PPTX_MIME_TYPE;
    case 'pdf':
      return PDF_MIME_TYPE;
  }
}

function resolveOriginalDownloadFileName(doc: DocumentItem): string {
  const defaultExt = `.${doc.type}`;
  if (doc.originalFileName && doc.originalFileName.trim().length > 0) {
    const cleanedOriginal = sanitizeFileName(doc.originalFileName, `tai-lieu${defaultExt}`);
    return /\.[a-z0-9]{2,5}$/i.test(cleanedOriginal)
      ? cleanedOriginal
      : `${cleanedOriginal}${defaultExt}`;
  }
  const safeBase = sanitizeFileName(
    doc.title,
    'tai-lieu',
    new RegExp(`\\.${doc.type}$`, 'i')
  );
  return `${safeBase}${defaultExt}`;
}

export async function exportOriginalDocumentBinary(
  doc: DocumentItem,
  binaryBuffer?: ArrayBuffer | null
): Promise<SaveFileResult> {
  const mimeType =
    doc.originalMimeType ||
    doc.sourceFile?.type ||
    getDefaultMimeTypeForDocType(doc.type);
  const fileName = resolveOriginalDownloadFileName(doc);

  if (binaryBuffer && binaryBuffer.byteLength > 0) {
    const blob = new Blob([binaryBuffer], { type: mimeType });
    return saveFileUniversal(blob, fileName, { mimeType, shareTitle: doc.title });
  }

  if (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
    const blob = new Blob([doc.rawArrayBuffer], { type: mimeType });
    return saveFileUniversal(blob, fileName, { mimeType, shareTitle: doc.title });
  }

  if (doc.sourceFile && doc.sourceFile.size > 0) {
    const blob =
      doc.sourceFile.type === mimeType
        ? doc.sourceFile
        : new Blob([doc.sourceFile], { type: mimeType });
    return saveFileUniversal(blob, fileName, { mimeType, shareTitle: doc.title });
  }

  if (doc.type === 'pdf') {
    const pdfUrl = doc.fileUrl ?? doc.objectUrl;
    if (pdfUrl) {
      const response = await fetch(pdfUrl);
      if (response.ok) {
        const fetchedBlob = await response.blob();
        const blob =
          fetchedBlob.type === PDF_MIME_TYPE
            ? fetchedBlob
            : new Blob([fetchedBlob], { type: PDF_MIME_TYPE });
        return saveFileUniversal(blob, fileName, {
          mimeType: PDF_MIME_TYPE,
          shareTitle: doc.title,
        });
      }
    }
  }

  // Fallback for sample/newly authored documents without a stored original binary
  const payload = await buildDocumentUploadPayload(doc);
  return saveFileUniversal(payload.blob, payload.fileName, {
    mimeType: payload.mimeType,
    shareTitle: doc.title,
  });
}

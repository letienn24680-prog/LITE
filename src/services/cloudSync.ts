import { createThumbnailDataUri } from '../data/mockDocuments';
import { isSupabaseConfigured, supabase } from '../lib/supabase';
import {
  deleteFileFromR2,
  getR2PresignedDownloadUrl,
  isR2ApiConfigured,
  uploadFileViaR2PresignedUrl,
} from './r2Storage';
import type {
  CellStyle,
  CloudUserSummary,
  DocumentFolder,
  DocumentItem,
  DocumentType,
  ExcelChartItem,
  ExcelChartSeries,
  ExcelConditionalFormat,
  ExcelDrawingStroke,
  ExcelFloatingObject,
  ExcelPageSetup,
  OutlineSection,
  PdfPageItem,
  SlideAnimationConfig,
  SlideChartData,
  SlideChartSeries,
  SlideCommentItem,
  SlideDrawingStroke,
  SlideImageItem,
  SlidePreviewItem,
  SlideShapeKind,
  SlideTextBlock,
  SlideTransitionConfig,
  SpreadsheetMergeRange,
  SpreadsheetSheet,
  WordCommentItem,
  WordDrawingStroke,
} from '../types/document';

const STORAGE_BUCKET = 'user-files';
const DOCUMENTS_TABLE = 'documents';
const FOLDERS_TABLE = 'folders';
const MAX_INLINE_DATA_URI_CHARS_WITH_BINARY = 96_000;
const MAX_INLINE_DATA_URI_CHARS_NO_BINARY = 350_000;
const MAX_SLIDE_INLINE_IMAGE_BUDGET_CHARS = 900_000;
const MAX_JSONB_STRING_CHARS = 65_536;
const MAX_DRAWING_STROKES = 400;
const MAX_POINTS_PER_STROKE = 2000;

const CLOUD_MIME_BY_TYPE: Record<DocumentType, string> = {
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pdf: 'application/pdf',
};

const CLOUD_EXTENSION_BY_TYPE: Record<DocumentType, string> = {
  pptx: 'pptx',
  docx: 'docx',
  xlsx: 'xlsx',
  pdf: 'pdf',
};

const DEFAULT_ACCENT_BY_TYPE: Record<DocumentType, string> = {
  pptx: '#4F46E5',
  docx: '#2563EB',
  xlsx: '#15803D',
  pdf: '#DB2777',
};

export interface CloudParsedContent {
  wordHtml?: string;
  slidesPreview: SlidePreviewItem[];
  slideAspectRatio?: number;
  sheets?: SpreadsheetSheet[];
  pdfPages?: PdfPageItem[];
  outline: OutlineSection[];
  wordParagraphs?: Record<string, string[]>;
  wordComments?: WordCommentItem[];
  wordDrawings?: WordDrawingStroke[];
  wordProtected?: boolean;
  slideProtected?: boolean;
  folderId?: string | null;
  deletedAt?: number | null;
}

interface CloudDocumentRow {
  id: string;
  user_id: string;
  title: string;
  file_type: string;
  file_name: string | null;
  file_size: string | null;
  file_size_bytes: number | null;
  storage_path: string | null;
  r2_object_key: string | null;
  folder: string | null;
  is_favorite: boolean;
  is_shared: boolean;
  page_count: number | null;
  content_preview: string | null;
  parsed_content: unknown;
  created_at: string | null;
  updated_at: string | null;
  last_accessed_at: string | null;
}

function getConfiguredSupabaseClient() {
  if (!isSupabaseConfigured || !supabase) {
    throw new Error(
      'Chế độ đám mây chưa được cấu hình. Vui lòng thiết lập VITE_SUPABASE_URL và VITE_SUPABASE_ANON_KEY.'
    );
  }
  return supabase;
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isValidDocumentType(value: unknown): value is DocumentType {
  return value === 'pptx' || value === 'docx' || value === 'xlsx' || value === 'pdf';
}

function formatByteSize(bytes: number): string {
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

function parseFormattedFileSize(fileSize?: string): number | null {
  if (!fileSize) {
    return null;
  }
  const match = fileSize.trim().match(/^([\d.]+)\s*(B|KB|MB|GB)$/i);
  if (!match) {
    return null;
  }
  const amount = Number.parseFloat(match[1]);
  if (!Number.isFinite(amount) || amount < 0) {
    return null;
  }
  const unit = match[2].toUpperCase();
  if (unit === 'B') {
    return Math.round(amount);
  }
  if (unit === 'KB') {
    return Math.round(amount * 1024);
  }
  if (unit === 'MB') {
    return Math.round(amount * 1024 * 1024);
  }
  if (unit === 'GB') {
    return Math.round(amount * 1024 * 1024 * 1024);
  }
  return null;
}

function isValidIsoTimestamp(value?: string | null): value is string {
  if (!value) {
    return false;
  }
  const timestamp = Date.parse(value);
  return !Number.isNaN(timestamp);
}

function resolveMimeTypeForFileName(
  fileName: string | undefined,
  explicitMime: string | undefined,
  fallbackDocType?: DocumentType
): string | undefined {
  const cleanExplicit = explicitMime?.trim();
  if (cleanExplicit && cleanExplicit.toLowerCase() !== 'application/octet-stream') {
    return cleanExplicit;
  }
  const lowerName = (fileName ?? '').trim().toLowerCase();
  if (lowerName.endsWith('.pdf')) {
    return CLOUD_MIME_BY_TYPE.pdf;
  }
  if (lowerName.endsWith('.docx') || lowerName.endsWith('.doc')) {
    return CLOUD_MIME_BY_TYPE.docx;
  }
  if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls')) {
    return CLOUD_MIME_BY_TYPE.xlsx;
  }
  if (lowerName.endsWith('.pptx') || lowerName.endsWith('.ppt')) {
    return CLOUD_MIME_BY_TYPE.pptx;
  }
  return fallbackDocType ? CLOUD_MIME_BY_TYPE[fallbackDocType] : undefined;
}

function buildOriginalFileForCloud(doc: DocumentItem): File | null {
  if (!(doc.rawArrayBuffer instanceof ArrayBuffer) || doc.rawArrayBuffer.byteLength === 0) {
    return null;
  }
  const ext = CLOUD_EXTENSION_BY_TYPE[doc.type];
  const rawOriginalName = doc.originalFileName?.trim();
  const baseName = rawOriginalName || doc.title.trim() || 'document';
  const hasKnownExt = /\.(docx|xlsx|pptx|pdf)$/i.test(baseName);
  const fileName = hasKnownExt ? baseName : `${baseName}.${ext}`;
  const mimeType =
    resolveMimeTypeForFileName(fileName, doc.originalMimeType, doc.type) ??
    CLOUD_MIME_BY_TYPE[doc.type];

  try {
    return new File([doc.rawArrayBuffer.slice(0)], fileName, { type: mimeType });
  } catch {
    return null;
  }
}

function sanitizeJsonSafeValue(
  value: unknown,
  maxStringLength = MAX_JSONB_STRING_CHARS,
  depth = 0
): unknown {
  if (depth > 12 || value === undefined || value === null) {
    return value === null ? null : undefined;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : undefined;
  }
  if (typeof value === 'string') {
    if (value.startsWith('blob:') || value.length > maxStringLength) {
      return undefined;
    }
    return value;
  }
  if (
    value instanceof ArrayBuffer ||
    ArrayBuffer.isView(value) ||
    (typeof Blob !== 'undefined' && value instanceof Blob)
  ) {
    return undefined;
  }
  if (Array.isArray(value)) {
    const cleanedItems: unknown[] = [];
    for (const item of value) {
      const cleaned = sanitizeJsonSafeValue(item, maxStringLength, depth + 1);
      if (cleaned !== undefined) {
        cleanedItems.push(cleaned);
      }
    }
    return cleanedItems;
  }
  if (isPlainRecord(value)) {
    const cleanedRecord: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      if (key === 'rawArrayBuffer' || key === 'sourceFile' || key === 'objectUrl') {
        continue;
      }
      const cleaned = sanitizeJsonSafeValue(entry, maxStringLength, depth + 1);
      if (cleaned !== undefined) {
        cleanedRecord[key] = cleaned;
      }
    }
    return cleanedRecord;
  }
  return undefined;
}

function sanitizeSlidesForCloud(
  slides: SlidePreviewItem[],
  options?: { hasBinaryBackup?: boolean; stripOversizedDataUris?: boolean }
): SlidePreviewItem[] {
  const shouldStripOversizedData = options?.stripOversizedDataUris !== false;
  const maxInlineDataUriLength = options?.hasBinaryBackup
    ? MAX_INLINE_DATA_URI_CHARS_WITH_BINARY
    : MAX_INLINE_DATA_URI_CHARS_NO_BINARY;
  let remainingInlineBudget = MAX_SLIDE_INLINE_IMAGE_BUDGET_CHARS;

  const isPersistableImageSrc = (src: unknown): src is string => {
    if (typeof src !== 'string') {
      return false;
    }
    const trimmed = src.trim();
    if (!trimmed || trimmed.startsWith('blob:')) {
      return false;
    }
    if (trimmed.startsWith('data:')) {
      if (!shouldStripOversizedData) {
        return true;
      }
      if (
        trimmed.length > maxInlineDataUriLength ||
        trimmed.length > remainingInlineBudget
      ) {
        return false;
      }
      remainingInlineBudget -= trimmed.length;
    }
    return true;
  };

  return slides
    .filter((slide): slide is SlidePreviewItem => isPlainRecord(slide))
    .map((slide) => {
      const cleanedBase = sanitizeJsonSafeValue(slide, MAX_JSONB_STRING_CHARS, 0);
      const nextSlide: SlidePreviewItem = isPlainRecord(cleanedBase)
        ? (cleanedBase as unknown as SlidePreviewItem)
        : { ...slide };

      if (typeof slide.hidden === 'boolean') {
        nextSlide.hidden = slide.hidden;
      } else {
        delete nextSlide.hidden;
      }

      if (isPlainRecord(slide.transition)) {
        const rawType = slide.transition.type;
        if (
          rawType === 'none' ||
          rawType === 'fade' ||
          rawType === 'push' ||
          rawType === 'wipe' ||
          rawType === 'split' ||
          rawType === 'zoom'
        ) {
          const cleanTransition: SlideTransitionConfig = { type: rawType };
          if (
            typeof slide.transition.durationMs === 'number' &&
            Number.isFinite(slide.transition.durationMs) &&
            slide.transition.durationMs >= 0
          ) {
            cleanTransition.durationMs = slide.transition.durationMs;
          }
          nextSlide.transition = cleanTransition;
        } else {
          delete nextSlide.transition;
        }
      }

      if (Array.isArray(slide.blocks)) {
        const sanitizedBlocks: SlideTextBlock[] = [];
        for (const block of slide.blocks) {
          if (!isPlainRecord(block)) {
            continue;
          }
          const cleanedBlockRaw = sanitizeJsonSafeValue(
            block,
            MAX_JSONB_STRING_CHARS,
            1
          );
          if (!isPlainRecord(cleanedBlockRaw)) {
            continue;
          }
          const safeBlock = cleanedBlockRaw as unknown as SlideTextBlock;

          if (typeof block.shapeKind === 'string') {
            safeBlock.shapeKind = block.shapeKind as SlideShapeKind;
          }
          if (isPlainRecord(block.chartData)) {
            const chartType = block.chartData.chartType;
            if (
              (chartType === 'column' || chartType === 'line' || chartType === 'pie') &&
              Array.isArray(block.chartData.categories) &&
              Array.isArray(block.chartData.series)
            ) {
              const safeSeries: SlideChartSeries[] = [];
              for (const s of block.chartData.series) {
                if (!isPlainRecord(s) || !Array.isArray(s.values)) {
                  continue;
                }
                const seriesItem: SlideChartSeries = {
                  name: typeof s.name === 'string' ? s.name : '',
                  values: s.values.filter(
                    (val): val is number => typeof val === 'number' && Number.isFinite(val)
                  ),
                };
                if (typeof s.color === 'string') {
                  seriesItem.color = s.color;
                }
                safeSeries.push(seriesItem);
              }
              const safeChart: SlideChartData = {
                chartType,
                categories: block.chartData.categories.map((c) => String(c ?? '')),
                series: safeSeries,
              };
              if (typeof block.chartData.title === 'string') {
                safeChart.title = block.chartData.title;
              }
              safeBlock.chartData = safeChart;
            } else {
              delete safeBlock.chartData;
            }
          }
          if (isPlainRecord(block.animation)) {
            const animType = block.animation.type;
            if (
              animType === 'none' ||
              animType === 'appear' ||
              animType === 'fadeIn' ||
              animType === 'flyIn' ||
              animType === 'zoomIn' ||
              animType === 'pulse'
            ) {
              const safeAnim: SlideAnimationConfig = { type: animType };
              if (
                typeof block.animation.order === 'number' &&
                Number.isFinite(block.animation.order)
              ) {
                safeAnim.order = block.animation.order;
              }
              safeBlock.animation = safeAnim;
            } else {
              delete safeBlock.animation;
            }
          }
          if (typeof block.underline === 'boolean') {
            safeBlock.underline = block.underline;
          } else {
            delete safeBlock.underline;
          }
          if (typeof block.strikethrough === 'boolean') {
            safeBlock.strikethrough = block.strikethrough;
          } else {
            delete safeBlock.strikethrough;
          }
          if (typeof block.highlightColor === 'string' && block.highlightColor.trim()) {
            safeBlock.highlightColor = block.highlightColor.trim().slice(0, 64);
          } else {
            delete safeBlock.highlightColor;
          }
          if (
            block.textScript === 'normal' ||
            block.textScript === 'super' ||
            block.textScript === 'sub'
          ) {
            safeBlock.textScript = block.textScript;
          } else {
            delete safeBlock.textScript;
          }
          if (
            typeof block.lineHeight === 'number' &&
            Number.isFinite(block.lineHeight) &&
            block.lineHeight > 0
          ) {
            safeBlock.lineHeight = block.lineHeight;
          } else {
            delete safeBlock.lineHeight;
          }
          if (block.textColumns === 1 || block.textColumns === 2) {
            safeBlock.textColumns = block.textColumns;
          } else {
            delete safeBlock.textColumns;
          }
          if (
            block.align === 'left' ||
            block.align === 'center' ||
            block.align === 'right' ||
            block.align === 'justify'
          ) {
            safeBlock.align = block.align;
          } else {
            delete safeBlock.align;
          }
          if (typeof block.groupId === 'string' && block.groupId.trim()) {
            safeBlock.groupId = block.groupId;
          }
          if (typeof block.zIndex === 'number' && Number.isFinite(block.zIndex)) {
            safeBlock.zIndex = block.zIndex;
          }
          if (typeof block.linkUrl === 'string' && !block.linkUrl.startsWith('blob:')) {
            safeBlock.linkUrl = block.linkUrl;
          }
          if (block.mediaKind === 'video' || block.mediaKind === 'audio') {
            safeBlock.mediaKind = block.mediaKind;
          }
          if (
            typeof block.mediaSrc === 'string' &&
            !block.mediaSrc.startsWith('blob:') &&
            !(block.mediaSrc.startsWith('data:') && block.mediaSrc.length > maxInlineDataUriLength)
          ) {
            safeBlock.mediaSrc = block.mediaSrc;
          } else if (block.mediaSrc !== undefined) {
            delete safeBlock.mediaSrc;
          }

          sanitizedBlocks.push(safeBlock);
        }
        nextSlide.blocks = sanitizedBlocks;
      }

      if (slide.backgroundImageSrc !== undefined) {
        if (isPersistableImageSrc(slide.backgroundImageSrc)) {
          nextSlide.backgroundImageSrc = slide.backgroundImageSrc;
        } else {
          delete nextSlide.backgroundImageSrc;
        }
      }

      if (Array.isArray(slide.images)) {
        const persistableImages: SlideImageItem[] = [];
        for (const img of slide.images) {
          if (!isPlainRecord(img) || !isPersistableImageSrc(img.src)) {
            continue;
          }
          const cleanedImg = sanitizeJsonSafeValue(img, MAX_JSONB_STRING_CHARS, 1);
          if (isPlainRecord(cleanedImg)) {
            const safeImg = cleanedImg as unknown as SlideImageItem;
            safeImg.src = img.src;
            persistableImages.push(safeImg);
          }
        }
        if (persistableImages.length > 0) {
          nextSlide.images = persistableImages;
        } else {
          delete nextSlide.images;
        }
      }

      const cleanSlideDrawings: SlideDrawingStroke[] | undefined =
        sanitizeDrawingStrokes(slide.drawings, 'sdraw', '#C43E1C');
      if (cleanSlideDrawings) {
        nextSlide.drawings = cleanSlideDrawings;
      } else {
        delete nextSlide.drawings;
      }

      if (Array.isArray(slide.comments)) {
        const cleanComments: SlideCommentItem[] = [];
        for (const item of slide.comments) {
          if (cleanComments.length >= 200) {
            break;
          }
          if (!isPlainRecord(item) || typeof item.text !== 'string') {
            continue;
          }
          const text = item.text.trim();
          if (!text) {
            continue;
          }
          cleanComments.push({
            id:
              typeof item.id === 'string' && item.id.trim()
                ? item.id.trim().slice(0, 120)
                : `scmt-${cleanComments.length + 1}`,
            text: text.slice(0, 4096),
            author:
              typeof item.author === 'string' && item.author.trim()
                ? item.author.trim().slice(0, 120)
                : 'Người dùng',
            createdAt:
              typeof item.createdAt === 'string' && item.createdAt.trim()
                ? item.createdAt.trim().slice(0, 64)
                : new Date().toISOString(),
          });
        }
        if (cleanComments.length > 0) {
          nextSlide.comments = cleanComments;
        } else {
          delete nextSlide.comments;
        }
      } else {
        delete nextSlide.comments;
      }

      if (typeof slide.footerText === 'string') {
        nextSlide.footerText = slide.footerText.trim().slice(0, 300);
      } else {
        delete nextSlide.footerText;
      }

      if (typeof slide.showSlideNumber === 'boolean') {
        nextSlide.showSlideNumber = slide.showSlideNumber;
      } else {
        delete nextSlide.showSlideNumber;
      }

      return nextSlide;
    });
}

function sanitizeDrawingStrokes(
  raw: unknown,
  idPrefix: string,
  fallbackColor: string
): WordDrawingStroke[] | undefined {
  if (!Array.isArray(raw)) {
    return undefined;
  }
  const cleanStrokes: WordDrawingStroke[] = [];
  for (const stroke of raw) {
    if (cleanStrokes.length >= MAX_DRAWING_STROKES) {
      break;
    }
    if (!isPlainRecord(stroke)) {
      continue;
    }
    const tool = stroke.tool;
    if (tool !== 'pen' && tool !== 'highlighter') {
      continue;
    }
    if (!Array.isArray(stroke.points)) {
      continue;
    }
    const points: Array<{ x: number; y: number }> = [];
    for (const pt of stroke.points) {
      if (points.length >= MAX_POINTS_PER_STROKE) {
        break;
      }
      if (
        isPlainRecord(pt) &&
        typeof pt.x === 'number' &&
        typeof pt.y === 'number' &&
        Number.isFinite(pt.x) &&
        Number.isFinite(pt.y)
      ) {
        points.push({ x: pt.x, y: pt.y });
      }
    }
    if (points.length === 0) {
      continue;
    }
    cleanStrokes.push({
      id:
        typeof stroke.id === 'string' && stroke.id.trim()
          ? stroke.id.trim()
          : `${idPrefix}-${cleanStrokes.length + 1}`,
      tool,
      color:
        typeof stroke.color === 'string' && stroke.color.trim()
          ? stroke.color.trim().slice(0, 64)
          : fallbackColor,
      width:
        typeof stroke.width === 'number' &&
        Number.isFinite(stroke.width) &&
        stroke.width > 0
          ? stroke.width
          : 2,
      points,
    });
  }
  return cleanStrokes.length > 0 ? cleanStrokes : undefined;
}

function sanitizeWordDrawingsForCloud(
  raw: unknown
): WordDrawingStroke[] | undefined {
  return sanitizeDrawingStrokes(raw, 'wdraw', '#185ABD');
}

function sanitizeCellStyleForCloud(rawStyle: unknown): CellStyle | null {
  if (!isPlainRecord(rawStyle)) {
    return null;
  }
  const style: CellStyle = {};
  if (typeof rawStyle.bold === 'boolean') style.bold = rawStyle.bold;
  if (typeof rawStyle.italic === 'boolean') style.italic = rawStyle.italic;
  if (typeof rawStyle.underline === 'boolean') style.underline = rawStyle.underline;
  if (typeof rawStyle.strike === 'boolean') style.strike = rawStyle.strike;
  if (typeof rawStyle.fontSize === 'number' && Number.isFinite(rawStyle.fontSize) && rawStyle.fontSize > 0) {
    style.fontSize = rawStyle.fontSize;
  }
  if (typeof rawStyle.fontFamily === 'string' && rawStyle.fontFamily.trim()) {
    style.fontFamily = rawStyle.fontFamily.trim().slice(0, 120);
  }
  if (typeof rawStyle.textColor === 'string' && rawStyle.textColor.trim()) {
    style.textColor = rawStyle.textColor.trim().slice(0, 64);
  }
  if (typeof rawStyle.bgColor === 'string' && rawStyle.bgColor.trim()) {
    style.bgColor = rawStyle.bgColor.trim().slice(0, 64);
  }
  if (
    rawStyle.align === 'left' ||
    rawStyle.align === 'center' ||
    rawStyle.align === 'right' ||
    rawStyle.align === 'justify'
  ) {
    style.align = rawStyle.align;
  }
  if (
    rawStyle.valign === 'top' ||
    rawStyle.valign === 'middle' ||
    rawStyle.valign === 'bottom'
  ) {
    style.valign = rawStyle.valign;
  }
  if (typeof rawStyle.wrapText === 'boolean') style.wrapText = rawStyle.wrapText;
  if (typeof rawStyle.numFmt === 'string' && rawStyle.numFmt.trim()) {
    style.numFmt = rawStyle.numFmt.slice(0, 120);
  }
  if (
    rawStyle.textRotation === 'horizontal' ||
    rawStyle.textRotation === 'vertical' ||
    rawStyle.textRotation === 'up45' ||
    rawStyle.textRotation === 'down45'
  ) {
    style.textRotation = rawStyle.textRotation;
  }
  if (typeof rawStyle.hyperlink === 'string') {
    const trimmedLink = rawStyle.hyperlink.trim();
    if (
      trimmedLink &&
      !/^(?:javascript|vbscript|blob):/i.test(trimmedLink) &&
      trimmedLink.length <= 2048
    ) {
      style.hyperlink = trimmedLink;
    }
  }
  if (typeof rawStyle.borderTop === 'string' && rawStyle.borderTop.trim()) {
    style.borderTop = rawStyle.borderTop.trim().slice(0, 120);
  }
  if (typeof rawStyle.borderRight === 'string' && rawStyle.borderRight.trim()) {
    style.borderRight = rawStyle.borderRight.trim().slice(0, 120);
  }
  if (typeof rawStyle.borderBottom === 'string' && rawStyle.borderBottom.trim()) {
    style.borderBottom = rawStyle.borderBottom.trim().slice(0, 120);
  }
  if (typeof rawStyle.borderLeft === 'string' && rawStyle.borderLeft.trim()) {
    style.borderLeft = rawStyle.borderLeft.trim().slice(0, 120);
  }
  return Object.keys(style).length > 0 ? style : null;
}

function sanitizeSheetsForCloud(
  rawSheets: unknown,
  options?: { hasBinaryBackup?: boolean; stripOversizedDataUris?: boolean }
): SpreadsheetSheet[] | undefined {
  if (!Array.isArray(rawSheets)) {
    return undefined;
  }
  const shouldStripOversizedData = options?.stripOversizedDataUris !== false;
  const maxInlineImageLength = options?.hasBinaryBackup
    ? MAX_INLINE_DATA_URI_CHARS_WITH_BINARY
    : MAX_INLINE_DATA_URI_CHARS_NO_BINARY;

  const cleanedSheets: SpreadsheetSheet[] = [];
  rawSheets.forEach((entry, index) => {
    if (!isPlainRecord(entry)) {
      return;
    }

    const columns = Array.isArray(entry.columns)
      ? entry.columns.map((col) => (typeof col === 'string' ? col : String(col ?? '')))
      : [];

    const rows: SpreadsheetSheet['rows'] = Array.isArray(entry.rows)
      ? entry.rows.map((row) =>
          Array.isArray(row)
            ? row.map((cell) => {
                if (typeof cell === 'string') {
                  if (cell.startsWith('blob:')) {
                    return '' as unknown as SpreadsheetSheet['rows'][number][number];
                  }
                  const trimmed =
                    cell.length > MAX_JSONB_STRING_CHARS
                      ? cell.slice(0, MAX_JSONB_STRING_CHARS)
                      : cell;
                  return trimmed as unknown as SpreadsheetSheet['rows'][number][number];
                }
                if (typeof cell === 'number' && Number.isFinite(cell)) {
                  return cell as unknown as SpreadsheetSheet['rows'][number][number];
                }
                if (typeof cell === 'boolean') {
                  return String(cell) as unknown as SpreadsheetSheet['rows'][number][number];
                }
                return '' as unknown as SpreadsheetSheet['rows'][number][number];
              })
            : []
        )
      : [];

    const cleanedSheet: SpreadsheetSheet = {
      id:
        typeof entry.id === 'string' && entry.id.trim()
          ? entry.id.trim()
          : `sheet-${index + 1}`,
      name:
        typeof entry.name === 'string' && entry.name.trim()
          ? entry.name.trim()
          : `Trang tính ${index + 1}`,
      columns,
      rows,
    };

    if (isPlainRecord(entry.cellStyles)) {
      const cleanedCellStyles: Record<string, CellStyle> = {};
      for (const [cellKey, rawStyle] of Object.entries(entry.cellStyles)) {
        const sanitizedStyle = sanitizeCellStyleForCloud(rawStyle);
        if (sanitizedStyle) {
          cleanedCellStyles[cellKey] = sanitizedStyle;
        }
      }
      if (Object.keys(cleanedCellStyles).length > 0) {
        cleanedSheet.cellStyles = cleanedCellStyles;
      }
    }

    if (Array.isArray(entry.merges)) {
      const cleanedMerges = entry.merges
        .filter((merge): merge is Record<string, unknown> => isPlainRecord(merge))
        .map((merge) => sanitizeJsonSafeValue(merge, 256, 1))
        .filter((merge): merge is Record<string, unknown> => isPlainRecord(merge));
      if (cleanedMerges.length > 0) {
        cleanedSheet.merges = cleanedMerges as unknown as SpreadsheetMergeRange[];
      }
    }

    if (Array.isArray(entry.colWidths)) {
      const widths = entry.colWidths.filter(
        (width): width is number =>
          typeof width === 'number' && Number.isFinite(width) && width >= 0
      );
      if (widths.length > 0) {
        cleanedSheet.colWidths = widths;
      }
    }

    if (Array.isArray(entry.rowHeights)) {
      const heights = entry.rowHeights.filter(
        (height): height is number =>
          typeof height === 'number' && Number.isFinite(height) && height >= 0
      );
      if (heights.length > 0) {
        cleanedSheet.rowHeights = heights;
      }
    }

    if (
      typeof entry.frozenRows === 'number' &&
      Number.isFinite(entry.frozenRows) &&
      entry.frozenRows >= 0
    ) {
      cleanedSheet.frozenRows = Math.floor(entry.frozenRows);
    }

    if (
      typeof entry.frozenCols === 'number' &&
      Number.isFinite(entry.frozenCols) &&
      entry.frozenCols >= 0
    ) {
      cleanedSheet.frozenCols = Math.floor(entry.frozenCols);
    }

    if (typeof entry.tabColor === 'string' && entry.tabColor.trim()) {
      cleanedSheet.tabColor = entry.tabColor.trim().slice(0, 64);
    }

    if (Array.isArray(entry.hiddenRows)) {
      const hiddenRows = entry.hiddenRows.filter(
        (idx): idx is number =>
          typeof idx === 'number' && Number.isInteger(idx) && idx >= 0
      );
      if (hiddenRows.length > 0) {
        cleanedSheet.hiddenRows = hiddenRows;
      }
    }

    if (Array.isArray(entry.hiddenCols)) {
      const hiddenCols = entry.hiddenCols.filter(
        (idx): idx is number =>
          typeof idx === 'number' && Number.isInteger(idx) && idx >= 0
      );
      if (hiddenCols.length > 0) {
        cleanedSheet.hiddenCols = hiddenCols;
      }
    }

    if (typeof entry.autoFilterEnabled === 'boolean') {
      cleanedSheet.autoFilterEnabled = entry.autoFilterEnabled;
    }

    if (isPlainRecord(entry.dataValidations)) {
      const cleanValidations: Record<string, string[]> = {};
      let hasValidation = false;
      for (const [cellKey, optionsList] of Object.entries(entry.dataValidations)) {
        if (Array.isArray(optionsList)) {
          cleanValidations[cellKey] = optionsList
            .filter((item): item is string => typeof item === 'string')
            .slice(0, 500);
          hasValidation = true;
        }
      }
      if (hasValidation) {
        cleanedSheet.dataValidations = cleanValidations;
      }
    }

    if (Array.isArray(entry.conditionalFormats)) {
      const cleanRules: ExcelConditionalFormat[] = [];
      for (const item of entry.conditionalFormats) {
        if (!isPlainRecord(item) || !isPlainRecord(item.range)) {
          continue;
        }
        const rule = item.rule;
        if (
          rule !== 'gt' &&
          rule !== 'lt' &&
          rule !== 'eq' &&
          rule !== 'duplicate'
        ) {
          continue;
        }
        const { minRow, maxRow, minCol, maxCol } = item.range;
        if (
          typeof minRow !== 'number' ||
          typeof maxRow !== 'number' ||
          typeof minCol !== 'number' ||
          typeof maxCol !== 'number' ||
          !Number.isFinite(minRow) ||
          !Number.isFinite(maxRow) ||
          !Number.isFinite(minCol) ||
          !Number.isFinite(maxCol)
        ) {
          continue;
        }
        const formattedRule: ExcelConditionalFormat = {
          id:
            typeof item.id === 'string' && item.id.trim()
              ? item.id.trim()
              : `cf-${cleanRules.length + 1}`,
          range: { minRow, maxRow, minCol, maxCol },
          rule,
          bgColor: typeof item.bgColor === 'string' ? item.bgColor : '#FEF3C7',
        };
        if (typeof item.value === 'string') {
          formattedRule.value = item.value;
        }
        if (typeof item.textColor === 'string') {
          formattedRule.textColor = item.textColor;
        }
        cleanRules.push(formattedRule);
      }
      if (cleanRules.length > 0) {
        cleanedSheet.conditionalFormats = cleanRules;
      }
    }

    if (Array.isArray(entry.charts)) {
      const cleanCharts: ExcelChartItem[] = [];
      for (const chart of entry.charts) {
        if (!isPlainRecord(chart)) {
          continue;
        }
        const chartType = chart.type;
        if (chartType !== 'column' && chartType !== 'line' && chartType !== 'pie') {
          continue;
        }
        const categories = Array.isArray(chart.categories)
          ? chart.categories.map((cat) => String(cat ?? ''))
          : [];
        const cleanSeries: ExcelChartSeries[] = [];
        if (Array.isArray(chart.series)) {
          for (const ser of chart.series) {
            if (!isPlainRecord(ser) || !Array.isArray(ser.values)) {
              continue;
            }
            cleanSeries.push({
              name: typeof ser.name === 'string' ? ser.name : '',
              values: ser.values.filter(
                (v): v is number => typeof v === 'number' && Number.isFinite(v)
              ),
              color: typeof ser.color === 'string' ? ser.color : '#107C41',
            });
          }
        }
        cleanCharts.push({
          id:
            typeof chart.id === 'string' && chart.id.trim()
              ? chart.id.trim()
              : `chart-${cleanCharts.length + 1}`,
          type: chartType,
          title: typeof chart.title === 'string' ? chart.title : 'Biểu đồ',
          categories,
          series: cleanSeries,
        });
      }
      if (cleanCharts.length > 0) {
        cleanedSheet.charts = cleanCharts;
      }
    }

    if (isPlainRecord(entry.cellComments)) {
      const cleanComments: Record<string, string> = {};
      for (const [cellKey, commentVal] of Object.entries(entry.cellComments)) {
        if (typeof commentVal === 'string' && commentVal.trim()) {
          cleanComments[cellKey] = commentVal.trim().slice(0, 4096);
        }
      }
      if (Object.keys(cleanComments).length > 0) {
        cleanedSheet.cellComments = cleanComments;
      }
    }

    const cleanDrawings: ExcelDrawingStroke[] | undefined = sanitizeDrawingStrokes(
      entry.drawings,
      'draw',
      '#EF4444'
    );
    if (cleanDrawings) {
      cleanedSheet.drawings = cleanDrawings;
    }

    if (Array.isArray(entry.floatingObjects)) {
      const cleanFloating: ExcelFloatingObject[] = [];
      for (const obj of entry.floatingObjects) {
        if (!isPlainRecord(obj)) {
          continue;
        }
        const kind = obj.kind;
        if (kind !== 'image' && kind !== 'shape' && kind !== 'textbox') {
          continue;
        }
        const x =
          typeof obj.x === 'number' && Number.isFinite(obj.x) ? obj.x : 0;
        const y =
          typeof obj.y === 'number' && Number.isFinite(obj.y) ? obj.y : 0;
        const width =
          typeof obj.width === 'number' &&
          Number.isFinite(obj.width) &&
          obj.width > 0
            ? obj.width
            : 160;
        const height =
          typeof obj.height === 'number' &&
          Number.isFinite(obj.height) &&
          obj.height > 0
            ? obj.height
            : 100;

        let safeSrc: string | undefined;
        if (typeof obj.src === 'string') {
          const trimmedSrc = obj.src.trim();
          const isOversizedData =
            shouldStripOversizedData &&
            trimmedSrc.startsWith('data:') &&
            trimmedSrc.length > maxInlineImageLength;
          if (trimmedSrc && !trimmedSrc.startsWith('blob:') && !isOversizedData) {
            safeSrc = trimmedSrc;
          }
        }
        if (kind === 'image' && !safeSrc) {
          continue;
        }

        const floatingItem: ExcelFloatingObject = {
          id:
            typeof obj.id === 'string' && obj.id.trim()
              ? obj.id.trim()
              : `obj-${cleanFloating.length + 1}`,
          kind,
          x,
          y,
          width,
          height,
        };
        if (safeSrc) {
          floatingItem.src = safeSrc;
        }
        if (typeof obj.text === 'string') {
          floatingItem.text = obj.text.slice(0, MAX_JSONB_STRING_CHARS);
        }
        if (
          obj.shapeKind === 'rect' ||
          obj.shapeKind === 'roundRect' ||
          obj.shapeKind === 'ellipse' ||
          obj.shapeKind === 'arrowRight' ||
          obj.shapeKind === 'star' ||
          obj.shapeKind === 'callout'
        ) {
          floatingItem.shapeKind = obj.shapeKind;
        }
        if (typeof obj.fillColor === 'string' && obj.fillColor.trim()) {
          floatingItem.fillColor = obj.fillColor.trim().slice(0, 64);
        }
        if (typeof obj.borderColor === 'string' && obj.borderColor.trim()) {
          floatingItem.borderColor = obj.borderColor.trim().slice(0, 64);
        }
        cleanFloating.push(floatingItem);
      }
      if (cleanFloating.length > 0) {
        cleanedSheet.floatingObjects = cleanFloating;
      }
    }

    if (isPlainRecord(entry.pageSetup)) {
      const rawSetup = entry.pageSetup;
      const margins =
        rawSetup.margins === 'normal' ||
        rawSetup.margins === 'narrow' ||
        rawSetup.margins === 'wide'
          ? rawSetup.margins
          : 'normal';
      const orientation =
        rawSetup.orientation === 'portrait' || rawSetup.orientation === 'landscape'
          ? rawSetup.orientation
          : 'portrait';
      const paperSize =
        rawSetup.paperSize === 'A4' ||
        rawSetup.paperSize === 'A5' ||
        rawSetup.paperSize === 'Letter'
          ? rawSetup.paperSize
          : 'A4';
      const cleanPageSetup: ExcelPageSetup = {
        margins,
        orientation,
        paperSize,
      };
      if (typeof rawSetup.printArea === 'string' && rawSetup.printArea.trim()) {
        cleanPageSetup.printArea = rawSetup.printArea.trim().slice(0, 64);
      }
      if (typeof rawSetup.printGridlines === 'boolean') {
        cleanPageSetup.printGridlines = rawSetup.printGridlines;
      }
      cleanedSheet.pageSetup = cleanPageSetup;
    }

    if (typeof entry.isProtected === 'boolean') {
      cleanedSheet.isProtected = entry.isProtected;
    }

    if (Array.isArray(entry.lockedRanges)) {
      const cleanLocked = entry.lockedRanges
        .filter((range): range is string => typeof range === 'string')
        .map((range) => range.trim())
        .filter((range) => range.length > 0 && range.length <= 64)
        .slice(0, 500);
      if (cleanLocked.length > 0) {
        cleanedSheet.lockedRanges = cleanLocked;
      }
    }

    if (typeof entry.showGridlines === 'boolean') {
      cleanedSheet.showGridlines = entry.showGridlines;
    }

    cleanedSheets.push(cleanedSheet);
  });

  return cleanedSheets;
}

function sanitizeOutlineForCloud(rawOutline: unknown): OutlineSection[] {
  if (!Array.isArray(rawOutline)) {
    return [];
  }
  const result: OutlineSection[] = [];
  for (const section of rawOutline) {
    if (!isPlainRecord(section)) {
      continue;
    }
    const cleaned = sanitizeJsonSafeValue(section, MAX_JSONB_STRING_CHARS, 0);
    if (isPlainRecord(cleaned)) {
      result.push(cleaned as unknown as OutlineSection);
    }
  }
  return result;
}

function sanitizePdfPagesForCloud(
  rawPdfPages: unknown,
  hasBinaryBackup?: boolean
): PdfPageItem[] | undefined {
  if (!Array.isArray(rawPdfPages)) {
    return undefined;
  }
  const maxFieldLength = hasBinaryBackup
    ? MAX_INLINE_DATA_URI_CHARS_WITH_BINARY
    : MAX_INLINE_DATA_URI_CHARS_NO_BINARY;
  const pages: PdfPageItem[] = [];
  for (const page of rawPdfPages) {
    if (!isPlainRecord(page)) {
      continue;
    }
    const cleaned = sanitizeJsonSafeValue(page, maxFieldLength, 0);
    if (isPlainRecord(cleaned)) {
      pages.push(cleaned as unknown as PdfPageItem);
    }
  }
  return pages;
}

function sanitizeWordCommentsForCloud(
  rawComments: unknown
): WordCommentItem[] | undefined {
  if (!Array.isArray(rawComments)) {
    return undefined;
  }
  const comments: WordCommentItem[] = [];
  rawComments.forEach((entry, index) => {
    if (!isPlainRecord(entry) || typeof entry.text !== 'string') {
      return;
    }
    const text = entry.text.trim();
    if (!text) {
      return;
    }
    comments.push({
      id:
        typeof entry.id === 'string' && entry.id.trim()
          ? entry.id.trim()
          : `comment-${index + 1}`,
      quote:
        typeof entry.quote === 'string'
          ? entry.quote.slice(0, 2000)
          : '',
      text: text.slice(0, MAX_JSONB_STRING_CHARS),
      author:
        typeof entry.author === 'string' && entry.author.trim()
          ? entry.author.trim().slice(0, 120)
          : 'Người dùng',
      createdAt:
        typeof entry.createdAt === 'string' && entry.createdAt.trim()
          ? entry.createdAt.trim()
          : new Date().toISOString(),
    });
  });
  return comments.length > 0 ? comments : undefined;
}

function normalizeWordParagraphs(
  value: unknown
): Record<string, string[]> | undefined {
  if (!isPlainRecord(value)) {
    return undefined;
  }
  const result: Record<string, string[]> = {};
  let hasEntry = false;
  for (const [key, entry] of Object.entries(value)) {
    if (Array.isArray(entry)) {
      result[key] = entry.filter((item): item is string => typeof item === 'string');
      hasEntry = true;
    }
  }
  return hasEntry ? result : undefined;
}

function normalizeParsedContent(
  raw: unknown,
  options?: { hasBinaryBackup?: boolean }
): CloudParsedContent {
  if (!isPlainRecord(raw)) {
    return {
      slidesPreview: [],
      outline: [],
    };
  }

  const rawSlides = Array.isArray(raw.slidesPreview)
    ? (raw.slidesPreview as SlidePreviewItem[])
    : [];

  const normalized: CloudParsedContent = {
    slidesPreview: sanitizeSlidesForCloud(rawSlides, {
      hasBinaryBackup: options?.hasBinaryBackup,
      stripOversizedDataUris: false,
    }),
    outline: sanitizeOutlineForCloud(raw.outline),
  };

  if (typeof raw.wordHtml === 'string') {
    normalized.wordHtml = raw.wordHtml;
  }
  if (
    typeof raw.slideAspectRatio === 'number' &&
    Number.isFinite(raw.slideAspectRatio) &&
    raw.slideAspectRatio > 0
  ) {
    normalized.slideAspectRatio = raw.slideAspectRatio;
  }
  const normalizedSheets = sanitizeSheetsForCloud(raw.sheets, {
    hasBinaryBackup: options?.hasBinaryBackup,
    stripOversizedDataUris: false,
  });
  if (normalizedSheets) {
    normalized.sheets = normalizedSheets;
  }
  if (Array.isArray(raw.pdfPages)) {
    normalized.pdfPages = sanitizePdfPagesForCloud(raw.pdfPages, false);
  }
  const normalizedParagraphs = normalizeWordParagraphs(raw.wordParagraphs);
  if (normalizedParagraphs) {
    normalized.wordParagraphs = normalizedParagraphs;
  }
  const normalizedComments = sanitizeWordCommentsForCloud(raw.wordComments);
  if (normalizedComments) {
    normalized.wordComments = normalizedComments;
  }
  const normalizedWordDrawings = sanitizeWordDrawingsForCloud(raw.wordDrawings);
  if (normalizedWordDrawings) {
    normalized.wordDrawings = normalizedWordDrawings;
  }
  if (typeof raw.wordProtected === 'boolean') {
    normalized.wordProtected = raw.wordProtected;
  }
  if (typeof raw.slideProtected === 'boolean') {
    normalized.slideProtected = raw.slideProtected;
  }
  if (typeof raw.folderId === 'string' && raw.folderId.trim()) {
    normalized.folderId = raw.folderId.trim();
  } else if (raw.folderId === null) {
    normalized.folderId = null;
  }
  if (typeof raw.deletedAt === 'number' && Number.isFinite(raw.deletedAt)) {
    normalized.deletedAt = raw.deletedAt;
  } else if (raw.deletedAt === null) {
    normalized.deletedAt = null;
  }

  return normalized;
}

function sanitizeStorageFileName(originalName: string): string {
  const trimmed = originalName.trim();
  const lastDotIdx = trimmed.lastIndexOf('.');
  const rawBase = lastDotIdx > 0 ? trimmed.slice(0, lastDotIdx) : trimmed;
  const rawExt = lastDotIdx > 0 ? trimmed.slice(lastDotIdx + 1) : '';

  const asciiBase = rawBase
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/-{2,}/g, '-')
    .replace(/^[-_]+|[-_]+$/g, '')
    .slice(0, 100);

  const safeBase = asciiBase || 'document';
  const safeExt = rawExt
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase()
    .slice(0, 12);

  return safeExt ? `${safeBase}.${safeExt}` : safeBase;
}

function sanitizePathSegment(segment: string, fallback: string): string {
  const cleaned = segment
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, '_')
    .replace(/\.{2,}/g, '_')
    .slice(0, 120);
  return cleaned || fallback;
}

function mapCloudRowToDocumentItem(row: CloudDocumentRow): DocumentItem {
  const docType: DocumentType = isValidDocumentType(row.file_type)
    ? row.file_type
    : 'docx';
  const hasBinaryBackup = Boolean(row.storage_path || row.r2_object_key);
  const parsed = normalizeParsedContent(row.parsed_content, { hasBinaryBackup });
  const title = row.title?.trim() || 'Tài liệu chưa đặt tên';
  const accentFromSlide = parsed.slidesPreview[0]?.accentColor;
  const accentColor =
    docType === 'pptx' &&
    typeof accentFromSlide === 'string' &&
    /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(accentFromSlide.trim())
      ? accentFromSlide.trim()
      : DEFAULT_ACCENT_BY_TYPE[docType];

  const timestampForDisplay = row.updated_at ?? row.last_accessed_at ?? row.created_at;
  const computedFileSize =
    row.file_size ??
    (typeof row.file_size_bytes === 'number'
      ? formatByteSize(row.file_size_bytes)
      : undefined);

  const doc: DocumentItem = {
    id: row.id,
    title,
    type: docType,
    updatedAt: formatRelativeVietnameseTime(timestampForDisplay),
    isFavorite: Boolean(row.is_favorite),
    isShared: Boolean(row.is_shared),
    thumbnail: createThumbnailDataUri(docType, title, accentColor),
    slidesPreview: parsed.slidesPreview,
    slideAspectRatio: parsed.slideAspectRatio,
    outline: parsed.outline,
    contentPreview: row.content_preview ?? '',
    pageCount: typeof row.page_count === 'number' ? row.page_count : undefined,
    fileSize: computedFileSize,
    folder: row.folder ?? undefined,
    folderId: parsed.folderId ?? undefined,
    deletedAt: parsed.deletedAt ?? undefined,
    sheets: parsed.sheets,
    pdfPages: parsed.pdfPages,
    wordParagraphs: parsed.wordParagraphs,
    wordHtml: parsed.wordHtml,
    wordComments: parsed.wordComments,
    wordDrawings: parsed.wordDrawings,
    wordProtected: parsed.wordProtected,
    slideProtected: parsed.slideProtected,
    isUploaded: true,
    lastSavedAt: row.updated_at ?? undefined,
    lastAccessedAt: row.last_accessed_at ?? undefined,
    storagePath: row.storage_path ?? undefined,
    r2ObjectKey: row.r2_object_key ?? undefined,
    originalFileName: row.file_name ?? undefined,
    hasLocalEdits: false,
    isSyncedToCloud: true,
    userId: row.user_id,
  };

  return doc;
}

/**
 * Translates common Supabase Auth error messages into clear Vietnamese feedback.
 */
export function mapAuthErrorToVietnamese(message: string): string {
  const normalized = message.trim().toLowerCase();
  if (!normalized) {
    return 'Đã xảy ra lỗi xác thực. Vui lòng thử lại.';
  }

  if (
    normalized.includes('invalid login credentials') ||
    normalized.includes('invalid_credentials') ||
    normalized.includes('invalid email or password')
  ) {
    return 'Email hoặc mật khẩu chưa chính xác. Vui lòng kiểm tra lại.';
  }

  if (normalized.includes('email not confirmed')) {
    return 'Email chưa được xác nhận. Vui lòng kiểm tra hộp thư (kể cả mục Spam) và bấm liên kết xác thực.';
  }

  if (
    normalized.includes('user already registered') ||
    normalized.includes('already been registered') ||
    normalized.includes('email already exists') ||
    normalized.includes('already exists')
  ) {
    return 'Email này đã được đăng ký. Vui lòng chuyển sang tab Đăng nhập.';
  }

  if (
    normalized.includes('password should be at least') ||
    normalized.includes('weak password') ||
    normalized.includes('at least 6 characters')
  ) {
    return 'Mật khẩu quá ngắn. Vui lòng nhập tối thiểu 6 ký tự.';
  }

  if (
    normalized.includes('unable to validate email address') ||
    normalized.includes('invalid email') ||
    normalized.includes('email address') && normalized.includes('invalid')
  ) {
    return 'Địa chỉ email không hợp lệ. Vui lòng kiểm tra lại định dạng.';
  }

  if (
    normalized.includes('rate limit') ||
    normalized.includes('too many requests') ||
    normalized.includes('only request this after') ||
    normalized.includes('over_email_send_rate_limit') ||
    normalized.includes('over_request_rate_limit')
  ) {
    return 'Bạn thao tác quá nhanh. Vui lòng đợi ít phút rồi thử lại.';
  }

  if (
    normalized.includes('failed to fetch') ||
    normalized.includes('networkerror') ||
    normalized.includes('network request failed') ||
    normalized.includes('load failed') ||
    normalized.includes('timeout') ||
    normalized.includes('network')
  ) {
    return 'Không thể kết nối tới máy chủ. Vui lòng kiểm tra kết nối mạng và thử lại.';
  }

  if (normalized.includes('signups not allowed')) {
    return 'Hệ thống hiện tạm khóa đăng ký tài khoản mới.';
  }

  if (
    normalized.includes('auth session missing') ||
    normalized.includes('refresh token') ||
    normalized.includes('jwt expired')
  ) {
    return 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.';
  }

  return `Không thể hoàn tất xác thực (${message.trim()}). Vui lòng thử lại.`;
}

/**
 * Builds a concise user summary from a Supabase Auth user object.
 */
export function buildCloudUserSummary(user: {
  id: string;
  email?: string | null;
  user_metadata?: Record<string, unknown>;
}): CloudUserSummary {
  const email = (user.email ?? '').trim();
  const metadata = user.user_metadata;

  let displayName = '';
  if (metadata) {
    const candidateKeys = ['full_name', 'name', 'display_name', 'username'];
    for (const key of candidateKeys) {
      const candidate = metadata[key];
      if (typeof candidate === 'string' && candidate.trim().length > 0) {
        displayName = candidate.trim();
        break;
      }
    }
  }

  if (!displayName) {
    const emailPrefix = email.includes('@') ? email.split('@')[0].trim() : email;
    displayName = emailPrefix || 'Người dùng LITE';
  }

  return {
    id: user.id,
    email,
    displayName,
  };
}

/**
 * Formats an ISO timestamp as relative Vietnamese time:
 * 'Vừa xong', 'X phút trước', 'Hôm nay, HH:MM', or 'dd/MM/yyyy HH:MM'.
 */
export function formatRelativeVietnameseTime(isoString?: string | null): string {
  if (!isoString) {
    return 'Chưa đồng bộ';
  }
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) {
    return 'Chưa đồng bộ';
  }

  const now = new Date();
  const diffMs = now.getTime() - date.getTime();

  if (diffMs < 60_000 && diffMs >= -60_000) {
    return 'Vừa xong';
  }

  if (diffMs >= 60_000 && diffMs < 3_600_000) {
    const minutes = Math.floor(diffMs / 60_000);
    return `${minutes} phút trước`;
  }

  const isSameDay =
    date.getFullYear() === now.getFullYear() &&
    date.getMonth() === now.getMonth() &&
    date.getDate() === now.getDate();

  if (isSameDay) {
    const timeFormatted = new Intl.DateTimeFormat('vi-VN', {
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    }).format(date);
    return `Hôm nay, ${timeFormatted}`;
  }

  const parts = new Intl.DateTimeFormat('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(date);

  const partMap = new Map<string, string>();
  for (const part of parts) {
    partMap.set(part.type, part.value);
  }

  const day = partMap.get('day') ?? '01';
  const month = partMap.get('month') ?? '01';
  const year = partMap.get('year') ?? String(date.getFullYear());
  const hour = partMap.get('hour') ?? '00';
  const minute = partMap.get('minute') ?? '00';

  return `${day}/${month}/${year} ${hour}:${minute}`;
}

/**
 * Fetches all cloud documents belonging to the user ordered by most recently accessed.
 */
export async function fetchUserDocuments(userId: string): Promise<DocumentItem[]> {
  const client = getConfiguredSupabaseClient();
  const { data, error } = await client
    .from(DOCUMENTS_TABLE)
    .select('*')
    .eq('user_id', userId)
    .order('last_accessed_at', { ascending: false });

  if (error) {
    throw new Error(
      `Không thể tải danh sách tài liệu từ đám mây: ${error.message}`
    );
  }

  const rows = Array.isArray(data) ? (data as CloudDocumentRow[]) : [];
  return rows.map(mapCloudRowToDocumentItem);
}

/**
 * Sanitizes file name into ASCII-safe Storage key and uploads file to 'user-files' bucket
 * at {user_id}/{document_id}/{file_name}.
 */
export async function uploadDocumentFileToStorage(
  userId: string,
  docId: string,
  file: File
): Promise<{ storagePath: string; fileName: string }> {
  const client = getConfiguredSupabaseClient();
  const safeUserId = sanitizePathSegment(userId, 'user');
  const safeDocId = sanitizePathSegment(docId, 'doc');
  const safeFileName = sanitizeStorageFileName(file.name);
  const storagePath = `${safeUserId}/${safeDocId}/${safeFileName}`;

  const contentType = resolveMimeTypeForFileName(file.name, file.type);

  const { error } = await client.storage
    .from(STORAGE_BUCKET)
    .upload(storagePath, file, {
      upsert: true,
      contentType: contentType || undefined,
    });

  if (error) {
    throw new Error(
      `Không thể tải tệp "${file.name}" lên kho lưu trữ đám mây: ${error.message}`
    );
  }

  return {
    storagePath,
    fileName: safeFileName,
  };
}

/**
 * Downloads a file from Supabase Storage 'user-files' bucket and returns a browser Object URL + Blob.
 */
export async function getCloudFileObjectUrl(
  storagePath: string
): Promise<{ objectUrl: string; blob: Blob }> {
  if (!storagePath.trim()) {
    throw new Error('Đường dẫn tệp trên đám mây không hợp lệ.');
  }
  const client = getConfiguredSupabaseClient();
  const { data, error } = await client.storage
    .from(STORAGE_BUCKET)
    .download(storagePath);

  if (error || !data) {
    throw new Error(
      `Không thể tải tệp gốc từ đám mây${error?.message ? `: ${error.message}` : '.'}`
    );
  }

  const objectUrl = URL.createObjectURL(data);
  return {
    objectUrl,
    blob: data,
  };
}

async function readBlobAsArrayBuffer(blob: Blob): Promise<ArrayBuffer> {
  if (typeof blob.arrayBuffer === 'function') {
    return await blob.arrayBuffer();
  }
  return await new Response(blob).arrayBuffer();
}

async function fetchArrayBufferFromSignedOrRemoteUrl(
  url: string
): Promise<ArrayBuffer> {
  const response = await fetch(url, { cache: 'no-store' });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  const buffer = await response.arrayBuffer();
  if (buffer.byteLength === 0) {
    throw new Error('Tệp tải về có kích thước 0 byte.');
  }
  return buffer;
}

/**
 * Downloads a file from Supabase Storage 'user-files' bucket (or R2 fallback) directly as an ArrayBuffer
 * (used to hydrate .docx / .xlsx / .pptx / .pdf raw binary across devices when rawArrayBuffer is not in memory).
 */
export async function downloadCloudFileArrayBuffer(
  storagePath: string,
  options?: { r2ObjectKey?: string; accessToken?: string | null }
): Promise<ArrayBuffer> {
  const cleanPath = storagePath.trim();
  const cleanR2Key = options?.r2ObjectKey?.trim() ?? '';

  if (!cleanPath && !cleanR2Key) {
    throw new Error('Đường dẫn tệp trên đám mây không hợp lệ.');
  }

  let storageFailureMessage = '';

  if (cleanPath) {
    const client = getConfiguredSupabaseClient();
    const { data, error } = await client.storage
      .from(STORAGE_BUCKET)
      .download(cleanPath);

    if (!error && data) {
      const buffer = await readBlobAsArrayBuffer(data);
      if (buffer.byteLength > 0) {
        return buffer;
      }
      storageFailureMessage = 'Dữ liệu tệp tải về bị rỗng.';
    } else if (error?.message) {
      storageFailureMessage = error.message;
    }

    // Fallback to signed URL fetch in case direct authenticated blob download hits MIME/WebView quirks
    try {
      const { data: signedData, error: signedError } = await client.storage
        .from(STORAGE_BUCKET)
        .createSignedUrl(cleanPath, 120);
      if (!signedError && signedData?.signedUrl) {
        return await fetchArrayBufferFromSignedOrRemoteUrl(signedData.signedUrl);
      }
      if (signedError?.message && !storageFailureMessage) {
        storageFailureMessage = signedError.message;
      }
    } catch (signedFetchError) {
      if (!storageFailureMessage && signedFetchError instanceof Error) {
        storageFailureMessage = signedFetchError.message;
      }
    }
  }

  if (cleanR2Key && isR2ApiConfigured) {
    let resolvedToken = options?.accessToken ?? null;
    if (!resolvedToken && isSupabaseConfigured && supabase) {
      try {
        const sessionRes = await supabase.auth.getSession();
        resolvedToken = sessionRes.data.session?.access_token ?? null;
      } catch {
        resolvedToken = null;
      }
    }
    const { downloadUrl } = await getR2PresignedDownloadUrl({
      objectKey: cleanR2Key,
      accessToken: resolvedToken,
    });
    return await fetchArrayBufferFromSignedOrRemoteUrl(downloadUrl);
  }

  throw new Error(
    `Không thể tải dữ liệu tệp gốc từ đám mây${storageFailureMessage ? `: ${storageFailureMessage}` : '.'}`
  );
}

/**
 * Downloads a cloud file from Supabase Storage and triggers browser file download.
 */
export async function downloadCloudFileToDevice(
  storagePath: string,
  downloadFileName: string
): Promise<void> {
  const { objectUrl } = await getCloudFileObjectUrl(storagePath);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = downloadFileName.trim() || 'tai-lieu-lite';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);

  try {
    anchor.click();
  } finally {
    document.body.removeChild(anchor);
    window.setTimeout(() => {
      URL.revokeObjectURL(objectUrl);
    }, 1000);
  }
}

/**
 * Uploads the source file if needed and upserts the document record into public.documents.
 * Prioritizes preserving the intact original binary (rawArrayBuffer / sourceFile) when the file
 * has not been edited locally, storing it in R2 or Supabase Storage 'user-files' without bloating JSONB.
 */
export async function upsertCloudDocument(
  userId: string,
  doc: DocumentItem,
  options?: {
    file?: File | null;
    touchUpdatedAt?: boolean;
    forceUploadFile?: boolean;
    accessToken?: string | null;
  }
): Promise<DocumentItem> {
  const client = getConfiguredSupabaseClient();
  const explicitFile =
    options?.file instanceof File && options.file.size > 0 ? options.file : null;
  const docSourceFile =
    doc.sourceFile instanceof File && doc.sourceFile.size > 0 ? doc.sourceFile : null;
  const reconstructedOriginalFile = buildOriginalFileForCloud(doc);

  // When unedited locally, never replace an intact original file with a simplified export
  const binaryFileToUpload = doc.hasLocalEdits
    ? (explicitFile ?? docSourceFile ?? reconstructedOriginalFile)
    : (docSourceFile ?? reconstructedOriginalFile ?? explicitFile);

  let nextStoragePath = doc.storagePath;
  let nextR2ObjectKey = doc.r2ObjectKey;
  let nextOriginalFileName = doc.originalFileName ?? binaryFileToUpload?.name;

  const hasCloudBinaryAlready = Boolean(nextStoragePath || nextR2ObjectKey);
  const shouldUploadBinary =
    Boolean(binaryFileToUpload) &&
    (!hasCloudBinaryAlready || Boolean(options?.forceUploadFile));

  if (shouldUploadBinary && binaryFileToUpload) {
    let uploadedViaR2 = false;

    if (isR2ApiConfigured && (!nextStoragePath || Boolean(options?.forceUploadFile))) {
      let resolvedToken = options?.accessToken ?? null;
      if (!resolvedToken) {
        try {
          const sessionRes = await client.auth.getSession();
          resolvedToken = sessionRes.data.session?.access_token ?? null;
        } catch {
          resolvedToken = null;
        }
      }
      if (resolvedToken) {
        try {
          const r2Result = await uploadFileViaR2PresignedUrl({
            file: binaryFileToUpload,
            documentId: doc.id,
            accessToken: resolvedToken,
          });
          nextR2ObjectKey = r2Result.objectKey;
          uploadedViaR2 = true;
        } catch {
          // Fall back to Supabase Storage 'user-files' bucket below
        }
      }
    }

    if (!uploadedViaR2) {
      const uploaded = await uploadDocumentFileToStorage(
        userId,
        doc.id,
        binaryFileToUpload
      );
      nextStoragePath = uploaded.storagePath;
      if (options?.forceUploadFile || !nextOriginalFileName) {
        nextOriginalFileName = binaryFileToUpload.name;
      }
    } else if (options?.forceUploadFile || !nextOriginalFileName) {
      nextOriginalFileName = binaryFileToUpload.name;
    }
  }

  const hasBinaryBackup = Boolean(nextStoragePath || nextR2ObjectKey);
  const sanitizedSlides = sanitizeSlidesForCloud(doc.slidesPreview ?? [], {
    hasBinaryBackup,
    stripOversizedDataUris: true,
  });
  const sanitizedOutline = sanitizeOutlineForCloud(doc.outline ?? []);
  const sanitizedSheets = sanitizeSheetsForCloud(doc.sheets, {
    hasBinaryBackup,
    stripOversizedDataUris: true,
  });
  const sanitizedPdfPages = sanitizePdfPagesForCloud(doc.pdfPages, hasBinaryBackup);

  const parsedContent: CloudParsedContent = {
    slidesPreview: sanitizedSlides,
    outline: sanitizedOutline,
  };
  if (typeof doc.wordHtml === 'string') {
    parsedContent.wordHtml = doc.wordHtml;
  }
  if (
    typeof doc.slideAspectRatio === 'number' &&
    Number.isFinite(doc.slideAspectRatio) &&
    doc.slideAspectRatio > 0
  ) {
    parsedContent.slideAspectRatio = doc.slideAspectRatio;
  }
  if (sanitizedSheets) {
    parsedContent.sheets = sanitizedSheets;
  }
  if (sanitizedPdfPages) {
    parsedContent.pdfPages = sanitizedPdfPages;
  }
  const normalizedParagraphs = normalizeWordParagraphs(doc.wordParagraphs);
  if (normalizedParagraphs) {
    parsedContent.wordParagraphs = normalizedParagraphs;
  }
  const sanitizedWordComments = sanitizeWordCommentsForCloud(doc.wordComments);
  if (sanitizedWordComments) {
    parsedContent.wordComments = sanitizedWordComments;
  }
  const sanitizedWordDrawings = sanitizeWordDrawingsForCloud(doc.wordDrawings);
  if (sanitizedWordDrawings) {
    parsedContent.wordDrawings = sanitizedWordDrawings;
  }
  if (typeof doc.wordProtected === 'boolean') {
    parsedContent.wordProtected = doc.wordProtected;
  }
  if (typeof doc.slideProtected === 'boolean') {
    parsedContent.slideProtected = doc.slideProtected;
  }
  if (doc.folderId !== undefined) {
    parsedContent.folderId = doc.folderId ?? null;
  }
  if (doc.deletedAt !== undefined) {
    parsedContent.deletedAt = doc.deletedAt ?? null;
  }

  const touchUpdatedAt = options?.touchUpdatedAt !== false;
  const nowIso = new Date().toISOString();
  const updatedAtIso =
    touchUpdatedAt || !isValidIsoTimestamp(doc.lastSavedAt)
      ? nowIso
      : doc.lastSavedAt;
  const lastAccessedIso =
    touchUpdatedAt || !isValidIsoTimestamp(doc.lastAccessedAt)
      ? nowIso
      : doc.lastAccessedAt;

  const knownByteSize =
    binaryFileToUpload?.size ??
    (doc.rawArrayBuffer instanceof ArrayBuffer ? doc.rawArrayBuffer.byteLength : undefined);
  const computedFileSize =
    doc.fileSize ??
    (typeof knownByteSize === 'number' && knownByteSize > 0
      ? formatByteSize(knownByteSize)
      : undefined);
  const fileSizeBytes =
    typeof knownByteSize === 'number' && knownByteSize > 0
      ? knownByteSize
      : parseFormattedFileSize(computedFileSize);

  const rowPayload = {
    id: doc.id,
    user_id: userId,
    title: doc.title.trim() || 'Tài liệu mới',
    file_type: doc.type,
    file_name: nextOriginalFileName ?? null,
    file_size: computedFileSize ?? null,
    file_size_bytes: fileSizeBytes,
    storage_path: nextStoragePath ?? null,
    r2_object_key: nextR2ObjectKey ?? null,
    folder: doc.folder ?? null,
    is_favorite: Boolean(doc.isFavorite),
    is_shared: Boolean(doc.isShared),
    page_count: typeof doc.pageCount === 'number' ? doc.pageCount : null,
    content_preview: doc.contentPreview ?? '',
    parsed_content: parsedContent,
    updated_at: updatedAtIso,
    last_accessed_at: lastAccessedIso,
  };

  const { error } = await client
    .from(DOCUMENTS_TABLE)
    .upsert(rowPayload, { onConflict: 'user_id,id' });

  if (error) {
    throw new Error(
      `Không thể đồng bộ tài liệu "${doc.title}" lên đám mây: ${error.message}`
    );
  }

  return {
    ...doc,
    fileSize: computedFileSize,
    storagePath: nextStoragePath,
    r2ObjectKey: nextR2ObjectKey,
    originalFileName: nextOriginalFileName,
    hasLocalEdits: doc.hasLocalEdits ?? false,
    lastSavedAt: updatedAtIso,
    lastAccessedAt: lastAccessedIso,
    updatedAt: touchUpdatedAt
      ? formatRelativeVietnameseTime(updatedAtIso)
      : doc.updatedAt,
    isSyncedToCloud: true,
    isUploaded: true,
    userId,
  };
}

/**
 * Updates last_accessed_at for a synced document when the user opens it.
 */
export async function touchDocumentLastAccessed(
  userId: string,
  docId: string
): Promise<void> {
  const client = getConfiguredSupabaseClient();
  const { error } = await client
    .from(DOCUMENTS_TABLE)
    .update({ last_accessed_at: new Date().toISOString() })
    .eq('user_id', userId)
    .eq('id', docId);

  if (error) {
    throw new Error(
      `Không thể cập nhật thời gian truy cập gần nhất: ${error.message}`
    );
  }
}

/**
 * Sequentially syncs local documents to the user's Supabase account.
 * Reconstructs original binaries from rawArrayBuffer when sourceFile is absent
 * and preserves unedited uploads across R2 and Supabase Storage.
 */
export async function syncLocalDocumentsToCloud(
  userId: string,
  docs: DocumentItem[]
): Promise<{ syncedDocs: DocumentItem[]; failedCount: number }> {
  const syncedDocs: DocumentItem[] = [];
  let failedCount = 0;

  let accessToken: string | null = null;
  if (isR2ApiConfigured && isSupabaseConfigured && supabase) {
    try {
      const sessionRes = await supabase.auth.getSession();
      accessToken = sessionRes.data.session?.access_token ?? null;
    } catch {
      accessToken = null;
    }
  }

  for (const doc of docs) {
    try {
      const originalFile = doc.sourceFile ?? buildOriginalFileForCloud(doc);
      const syncedDoc = await upsertCloudDocument(userId, doc, {
        file: originalFile,
        touchUpdatedAt: false,
        accessToken,
      });
      syncedDocs.push(syncedDoc);
    } catch {
      failedCount += 1;
    }
  }

  return {
    syncedDocs,
    failedCount,
  };
}

/**
 * Deletes a document record from Supabase documents, along with best-effort cleanup
 * of its share links, Supabase Storage object, and Cloudflare R2 object.
 */
export async function deleteCloudDocument(
  userId: string,
  doc: Pick<DocumentItem, 'id' | 'storagePath' | 'r2ObjectKey'>,
  options?: { accessToken?: string | null }
): Promise<void> {
  const client = getConfiguredSupabaseClient();

  try {
    await client
      .from('document_shares')
      .delete()
      .eq('user_id', userId)
      .eq('document_id', doc.id);
  } catch {
    // Ignore if document_shares table is not yet created
  }

  const { error } = await client
    .from(DOCUMENTS_TABLE)
    .delete()
    .eq('user_id', userId)
    .eq('id', doc.id);

  if (error) {
    throw new Error(`Không thể xóa tài liệu trên đám mây: ${error.message}`);
  }

  if (doc.storagePath && doc.storagePath.trim()) {
    try {
      await client.storage
        .from(STORAGE_BUCKET)
        .remove([doc.storagePath.trim()]);
    } catch {
      // Best-effort Supabase Storage cleanup
    }
  }

  if (
    doc.r2ObjectKey &&
    doc.r2ObjectKey.trim() &&
    isR2ApiConfigured &&
    options?.accessToken
  ) {
    try {
      await deleteFileFromR2({
        objectKey: doc.r2ObjectKey,
        accessToken: options.accessToken,
      });
    } catch {
      // Best-effort Cloudflare R2 cleanup
    }
  }
}

interface CloudFolderRow {
  id: string;
  user_id: string;
  name: string;
  color: string | null;
  created_at: string | null;
  updated_at: string | null;
}

export async function fetchUserFolders(
  userId: string
): Promise<DocumentFolder[] | null> {
  if (!isSupabaseConfigured || !supabase) {
    return null;
  }
  try {
    const { data, error } = await supabase
      .from(FOLDERS_TABLE)
      .select('id, user_id, name, color, created_at, updated_at')
      .eq('user_id', userId)
      .order('created_at', { ascending: true });

    if (error || !Array.isArray(data)) {
      return null;
    }

    return (data as CloudFolderRow[]).map((row) => ({
      id: row.id,
      name: row.name?.trim() || 'Thư mục mới',
      color: row.color?.trim() || '#4F46E5',
      userId: row.user_id,
      ownerScope: row.user_id,
      createdAt: row.created_at ?? new Date().toISOString(),
      updatedAt: row.updated_at ?? undefined,
    }));
  } catch {
    return null;
  }
}

export async function upsertCloudFolder(
  userId: string,
  folder: DocumentFolder
): Promise<void> {
  if (!isSupabaseConfigured || !supabase) {
    return;
  }
  try {
    await supabase.from(FOLDERS_TABLE).upsert(
      {
        id: folder.id,
        user_id: userId,
        name: folder.name.trim() || 'Thư mục mới',
        color: folder.color || '#4F46E5',
        created_at: folder.createdAt || new Date().toISOString(),
        updated_at: new Date().toISOString(),
      },
      { onConflict: 'user_id,id' }
    );
  } catch {
    // Graceful when folders table migration has not yet been applied
  }
}

export async function deleteCloudFolder(
  userId: string,
  folderId: string
): Promise<void> {
  if (!isSupabaseConfigured || !supabase) {
    return;
  }
  try {
    await supabase
      .from(FOLDERS_TABLE)
      .delete()
      .eq('user_id', userId)
      .eq('id', folderId);
  } catch {
    // Graceful when folders table migration has not yet been applied
  }
}

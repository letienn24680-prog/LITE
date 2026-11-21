import {
  GlobalWorkerOptions,
  RenderingCancelledException,
  TextLayer,
  getDocument,
  type PDFDocumentProxy,
  type PDFPageProxy,
  type RenderTask,
} from 'pdfjs-dist/legacy/build/pdf.mjs';
import pdfWorkerUrl from 'pdfjs-dist/legacy/build/pdf.worker.min.mjs?url';
import type { OutlineSection } from '../types/document';

GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export type PdfDocumentProxy = PDFDocumentProxy & {
  destroy: () => Promise<void>;
};
export type PdfPageProxy = PDFPageProxy;

export interface PdfMetadataSummary {
  pageCount: number;
  contentPreview: string;
  outline: OutlineSection[];
}

const MAX_CANVAS_PIXELS = 16_777_216; // 4096 x 4096 safety cap for iOS Safari
const TEXT_LAYER_STYLE_ATTR = 'data-lite-pdf-text-layer';

const activeCanvasTasks = new WeakMap<HTMLCanvasElement, RenderTask>();
const activeTextLayers = new WeakMap<HTMLElement, TextLayer>();

interface PdfOutlineItemRaw {
  title: string;
  dest: string | unknown[] | null;
  items?: PdfOutlineItemRaw[];
}

function isRenderCancelledError(err: unknown): boolean {
  if (err instanceof RenderingCancelledException) {
    return true;
  }
  if (err instanceof Error) {
    return (
      err.name === 'RenderingCancelledException' ||
      err.name === 'AbortException' ||
      err.message.toLowerCase().includes('rendering cancelled')
    );
  }
  return false;
}

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

function truncateText(text: string, maxLength = 260): string {
  const normalized = text.replace(/\s+/g, ' ').trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 1).trimEnd()}…`;
}

function ensurePdfTextLayerStyles(): void {
  if (typeof document === 'undefined') {
    return;
  }
  if (document.head.querySelector(`style[${TEXT_LAYER_STYLE_ATTR}="true"]`)) {
    return;
  }
  const styleEl = document.createElement('style');
  styleEl.setAttribute(TEXT_LAYER_STYLE_ATTR, 'true');
  styleEl.textContent = `
.textLayer {
  position: absolute;
  text-align: initial;
  inset: 0;
  overflow: clip;
  opacity: 1;
  line-height: 1;
  letter-spacing: normal;
  word-spacing: normal;
  -webkit-text-size-adjust: none;
  text-size-adjust: none;
  forced-color-adjust: none;
  transform-origin: 0 0;
  caret-color: CanvasText;
  z-index: 2;
  --min-font-size: 1;
  --text-scale-factor: calc(var(--total-scale-factor, 1) * var(--min-font-size));
  --min-font-size-inv: calc(1 / var(--min-font-size));
}
.textLayer span,
.textLayer br {
  color: transparent;
  position: absolute;
  white-space: pre;
  cursor: text;
  transform-origin: 0% 0%;
  -webkit-user-select: text;
  user-select: text;
}
.textLayer > :not(.markedContent),
.textLayer .markedContent span:not(.markedContent) {
  z-index: 1;
  --font-height: 0;
  font-size: calc(var(--text-scale-factor) * var(--font-height));
  --scale-x: 1;
  --rotate: 0deg;
  transform: rotate(var(--rotate)) scaleX(var(--scale-x)) scale(var(--min-font-size-inv));
}
.textLayer .markedContent {
  display: contents;
}
.textLayer ::selection {
  background: rgba(79, 70, 229, 0.28);
  color: transparent;
}
`;
  document.head.appendChild(styleEl);
}

function attachDestroyCompat(
  rawDoc: PDFDocumentProxy,
  destroyFn: () => Promise<void>
): PdfDocumentProxy {
  const doc = rawDoc as PdfDocumentProxy;
  if (typeof doc.destroy !== 'function') {
    Object.defineProperty(doc, 'destroy', {
      value: () => destroyFn(),
      configurable: true,
      writable: true,
    });
  }
  return doc;
}

export async function loadPdfDocumentFromSource(
  source: string | ArrayBuffer | Uint8Array
): Promise<PdfDocumentProxy> {
  if (typeof source === 'string') {
    const loadingTask = getDocument({
      url: source,
      withCredentials: false,
    });
    const rawDoc = await loadingTask.promise;
    return attachDestroyCompat(rawDoc, () => loadingTask.destroy());
  }

  // Always pass a detached-safe copy so worker transfer never invalidates caller memory
  const copiedBytes =
    source instanceof ArrayBuffer ? new Uint8Array(source.slice(0)) : source.slice();

  const loadingTask = getDocument({
    data: copiedBytes,
  });
  const rawDoc = await loadingTask.promise;
  return attachDestroyCompat(rawDoc, () => loadingTask.destroy());
}

async function resolveOutlineDestPageNumber(
  pdfDoc: PdfDocumentProxy,
  dest: string | unknown[] | null
): Promise<number> {
  if (!dest) {
    return 1;
  }

  try {
    const explicitDest =
      typeof dest === 'string' ? await pdfDoc.getDestination(dest) : dest;
    if (!Array.isArray(explicitDest) || explicitDest.length === 0) {
      return 1;
    }

    const pageRef = explicitDest[0];
    if (typeof pageRef === 'number' && Number.isFinite(pageRef)) {
      return Math.min(pdfDoc.numPages, Math.max(1, Math.floor(pageRef) + 1));
    }
    if (
      pageRef &&
      typeof pageRef === 'object' &&
      'num' in pageRef &&
      'gen' in pageRef
    ) {
      const pageIdx = await pdfDoc.getPageIndex(
        pageRef as Parameters<PDFDocumentProxy['getPageIndex']>[0]
      );
      return Math.min(pdfDoc.numPages, Math.max(1, pageIdx + 1));
    }
  } catch {
    return 1;
  }

  return 1;
}

async function extractPdfOutlineSections(pdfDoc: PdfDocumentProxy): Promise<OutlineSection[]> {
  const rawOutline = (await pdfDoc.getOutline().catch(() => null)) as
    | PdfOutlineItemRaw[]
    | null;
  if (!Array.isArray(rawOutline) || rawOutline.length === 0) {
    return [];
  }

  const sections: OutlineSection[] = [];
  const maxItems = 60;
  let topCount = 0;

  for (const item of rawOutline) {
    if (sections.length >= maxItems) {
      break;
    }
    const title = (item.title ?? '').replace(/\s+/g, ' ').trim();
    if (!title) {
      continue;
    }

    topCount++;
    const pageNumber = await resolveOutlineDestPageNumber(pdfDoc, item.dest);
    const childTitles = Array.isArray(item.items)
      ? item.items
          .map((child) => (child.title ?? '').replace(/\s+/g, ' ').trim())
          .filter(Boolean)
      : [];

    sections.push({
      id: `pdf-outline-${sections.length + 1}`,
      romanIndex: toRomanNumeral(topCount),
      title,
      summary: childTitles.length > 0 ? childTitles.slice(0, 2).join(' · ') : `Trang ${pageNumber}`,
      keyPoints: childTitles.slice(0, 3),
      pageNumber,
    });

    if (Array.isArray(item.items) && item.items.length > 0) {
      let subCount = 0;
      for (const child of item.items) {
        if (sections.length >= maxItems) {
          break;
        }
        const subTitle = (child.title ?? '').replace(/\s+/g, ' ').trim();
        if (!subTitle) {
          continue;
        }
        subCount++;
        const subPage = await resolveOutlineDestPageNumber(pdfDoc, child.dest);
        sections.push({
          id: `pdf-outline-${sections.length + 1}`,
          romanIndex: `${topCount}.${subCount}`,
          title: subTitle,
          summary: `Trang ${subPage}`,
          keyPoints: [],
          pageNumber: subPage,
        });
      }
    }
  }

  return sections;
}

export async function inspectPdfMetadata(file: File): Promise<PdfMetadataSummary> {
  const arrayBuffer = await file.arrayBuffer();
  const pdfDoc = await loadPdfDocumentFromSource(arrayBuffer);

  try {
    const pageCount = pdfDoc.numPages;
    const textParts: string[] = [];
    const samplePageLimit = Math.min(pageCount, 2);

    for (let pageNum = 1; pageNum <= samplePageLimit; pageNum++) {
      const page = await pdfDoc.getPage(pageNum);
      try {
        const textContent = await page.getTextContent();
        const pageChunks: string[] = [];
        for (const item of textContent.items) {
          if ('str' in item && typeof item.str === 'string' && item.str.trim().length > 0) {
            pageChunks.push(item.str);
          }
        }
        if (pageChunks.length > 0) {
          textParts.push(pageChunks.join(' '));
        }
      } finally {
        page.cleanup();
      }
      if (textParts.join(' ').length >= 300) {
        break;
      }
    }

    const outline = await extractPdfOutlineSections(pdfDoc);
    const contentPreview = truncateText(textParts.join(' '), 260);

    return {
      pageCount,
      contentPreview,
      outline,
    };
  } finally {
    await pdfDoc.destroy().catch(() => undefined);
  }
}

export async function renderPdfPageToCanvas(options: {
  pdfDoc: PdfDocumentProxy;
  pageNumber: number;
  canvas: HTMLCanvasElement;
  targetWidthCss?: number;
  scale?: number;
  maxDpr?: number;
}): Promise<{ widthCss: number; heightCss: number; cancel: () => void }> {
  const { pdfDoc, canvas } = options;

  // Cancel any in-flight render task on this canvas before starting a new one
  const previousTask = activeCanvasTasks.get(canvas);
  if (previousTask) {
    try {
      previousTask.cancel();
    } catch {
      // Ignore already finished task
    }
    activeCanvasTasks.delete(canvas);
  }

  const safePageNum = Math.min(Math.max(1, Math.floor(options.pageNumber)), pdfDoc.numPages);
  const page = await pdfDoc.getPage(safePageNum);
  const baseViewport = page.getViewport({ scale: 1 });

  const zoomMultiplier =
    typeof options.scale === 'number' && Number.isFinite(options.scale) && options.scale > 0
      ? options.scale
      : 1;
  const cssScale =
    typeof options.targetWidthCss === 'number' &&
    Number.isFinite(options.targetWidthCss) &&
    options.targetWidthCss > 0
      ? (options.targetWidthCss / Math.max(1, baseViewport.width)) * zoomMultiplier
      : zoomMultiplier;

  const viewport = page.getViewport({ scale: cssScale });
  const rawDpr = typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1;
  let effectiveDpr = Math.min(rawDpr, options.maxDpr ?? 2);

  // Guard against iOS Safari Canvas GPU memory limit (4096 x 4096 pixels)
  const projectedPixels = viewport.width * effectiveDpr * (viewport.height * effectiveDpr);
  if (projectedPixels > MAX_CANVAS_PIXELS) {
    effectiveDpr *= Math.sqrt(MAX_CANVAS_PIXELS / projectedPixels);
  }

  const pixelWidth = Math.max(1, Math.floor(viewport.width * effectiveDpr));
  const pixelHeight = Math.max(1, Math.floor(viewport.height * effectiveDpr));

  canvas.width = pixelWidth;
  canvas.height = pixelHeight;
  canvas.style.width = `${viewport.width}px`;
  canvas.style.height = `${viewport.height}px`;

  const transform =
    effectiveDpr !== 1 ? [effectiveDpr, 0, 0, effectiveDpr, 0, 0] : undefined;

  const renderTask = page.render({
    canvas,
    viewport,
    transform,
  });

  activeCanvasTasks.set(canvas, renderTask);

  const cancel = (): void => {
    try {
      renderTask.cancel();
    } catch {
      // Ignore cancellation errors on completed tasks
    }
    if (activeCanvasTasks.get(canvas) === renderTask) {
      activeCanvasTasks.delete(canvas);
    }
  };

  try {
    await renderTask.promise;
  } catch (err) {
    if (!isRenderCancelledError(err)) {
      throw err;
    }
  } finally {
    if (activeCanvasTasks.get(canvas) === renderTask) {
      activeCanvasTasks.delete(canvas);
    }
  }

  return {
    widthCss: viewport.width,
    heightCss: viewport.height,
    cancel,
  };
}

export async function renderPdfPageTextLayer(options: {
  pdfDoc: PdfDocumentProxy;
  pageNumber: number;
  container: HTMLElement;
  viewportScale: number;
}): Promise<() => void> {
  const { pdfDoc, container } = options;

  ensurePdfTextLayerStyles();

  const previousLayer = activeTextLayers.get(container);
  if (previousLayer) {
    try {
      previousLayer.cancel();
    } catch {
      // Ignore cancellation error
    }
    activeTextLayers.delete(container);
  }

  container.replaceChildren();

  const safePageNum = Math.min(Math.max(1, Math.floor(options.pageNumber)), pdfDoc.numPages);
  const safeScale =
    Number.isFinite(options.viewportScale) && options.viewportScale > 0
      ? options.viewportScale
      : 1;

  const page = await pdfDoc.getPage(safePageNum);
  const viewport = page.getViewport({ scale: safeScale });

  container.classList.add('textLayer');
  container.style.setProperty('--scale-factor', String(viewport.scale));
  container.style.setProperty('--total-scale-factor', String(viewport.scale));
  container.style.setProperty('--scale-round-x', '1px');
  container.style.setProperty('--scale-round-y', '1px');
  container.style.width = `${viewport.width}px`;
  container.style.height = `${viewport.height}px`;

  const textContent = await page.getTextContent();
  const textLayer = new TextLayer({
    textContentSource: textContent,
    container,
    viewport,
  });

  activeTextLayers.set(container, textLayer);

  let cancelled = false;
  const cancel = (): void => {
    if (cancelled) {
      return;
    }
    cancelled = true;
    try {
      textLayer.cancel();
    } catch {
      // Ignore cancellation errors
    }
    if (activeTextLayers.get(container) === textLayer) {
      activeTextLayers.delete(container);
    }
    container.replaceChildren();
  };

  try {
    await textLayer.render();
  } catch (err) {
    if (!isRenderCancelledError(err)) {
      throw err;
    }
  }

  return cancel;
}

export async function renderPdfThumbnailDataUrl(
  pdfDoc: PdfDocumentProxy,
  pageNumber: number,
  widthPx = 160
): Promise<string> {
  const safePageNum = Math.min(Math.max(1, Math.floor(pageNumber)), pdfDoc.numPages);
  const page = await pdfDoc.getPage(safePageNum);
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = Math.max(0.1, widthPx / Math.max(1, baseViewport.width));
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(viewport.width));
  canvas.height = Math.max(1, Math.round(viewport.height));

  try {
    const renderTask = page.render({
      canvas,
      viewport,
    });
    await renderTask.promise;

    const webp = canvas.toDataURL('image/webp', 0.82);
    if (webp.startsWith('data:image/webp')) {
      return webp;
    }
    return canvas.toDataURL('image/png');
  } finally {
    // Release offscreen canvas GPU backing store immediately
    canvas.width = 0;
    canvas.height = 0;
    page.cleanup();
  }
}

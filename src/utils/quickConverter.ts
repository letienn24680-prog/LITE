import html2canvas from 'html2canvas';
import { PDFDocument } from 'pdf-lib';
import type {
  CellStyle,
  DocumentItem,
  SlidePreviewItem,
  SpreadsheetSheet,
} from '../types/document';
import { columnIndexToLetter, evaluateSheetMatrix } from './formulaEvaluator';
import { saveFileUniversal } from './saveFileUniversal';
import type { SaveFileResult } from './saveFileUniversal';

const A4_PORTRAIT_WIDTH_PT = 595.28;
const A4_PORTRAIT_HEIGHT_PT = 841.89;
const A4_LANDSCAPE_WIDTH_PT = 841.89;
const A4_LANDSCAPE_HEIGHT_PT = 595.28;

export interface PdfDownloadResult {
  blob: Blob;
  fileName: string;
  method: SaveFileResult['method'];
  savedPath?: string;
  message: string;
}

export interface MergePdfResult {
  blob: Blob;
  fileName: string;
  totalPages: number;
  method?: SaveFileResult['method'];
  savedPath?: string;
  message: string;
}

export interface SplitPdfResult {
  blob: Blob;
  fileName: string;
  pageCount: number;
  totalSourcePages: number;
  method?: SaveFileResult['method'];
  savedPath?: string;
  message: string;
}

function sanitizePdfFileName(title: string, fallback: string): string {
  const cleaned = (title || '')
    .replace(/\.(docx?|xlsx?|csv|pptx?|pdf)$/i, '')
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^\.+|\.+$/g, '')
    .trim();
  const base = cleaned || fallback;
  const truncated = base.length > 100 ? base.slice(0, 100).trim() : base;
  return `${truncated}.pdf`;
}

/**
 * Universal blob download helper (replaces legacy 1.5s revokeObjectURL approach).
 * Safe on Web (60s URL retain + fallback), Electron IPC, and Capacitor Android.
 */
export async function triggerBlobDownload(
  blob: Blob,
  fileName: string
): Promise<SaveFileResult> {
  return saveFileUniversal(blob, fileName, {
    mimeType: blob.type || 'application/pdf',
    shareTitle: fileName,
  });
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function sanitizeHtmlFragment(rawHtml: string): string {
  const parser = new DOMParser();
  const parsed = parser.parseFromString(rawHtml || '', 'text/html');
  parsed.body
    .querySelectorAll('script, iframe, object, embed, link')
    .forEach((node) => node.remove());

  const allElements = parsed.body.querySelectorAll('*');
  allElements.forEach((el) => {
    for (const attr of Array.from(el.attributes)) {
      const name = attr.name.toLowerCase();
      const val = attr.value.trim().toLowerCase();
      if (
        name.startsWith('on') ||
        ((name === 'href' || name === 'src') && val.startsWith('javascript:'))
      ) {
        el.removeAttribute(attr.name);
      }
    }
  });
  return parsed.body.innerHTML;
}

function buildFallbackWordHtml(doc: DocumentItem): string {
  const parts: string[] = [`<h1>${escapeHtml(doc.title)}</h1>`];
  if (doc.contentPreview) {
    parts.push(`<blockquote>${escapeHtml(doc.contentPreview)}</blockquote>`);
  }
  for (const section of doc.outline) {
    parts.push(
      `<h2>${escapeHtml(section.romanIndex)}. ${escapeHtml(section.title)}</h2>`
    );
    const paragraphs =
      doc.wordParagraphs?.[section.id] ??
      (section.summary ? [section.summary] : []);
    for (const p of paragraphs) {
      parts.push(`<p>${escapeHtml(p)}</p>`);
    }
    if (section.keyPoints.length > 0) {
      parts.push(
        `<ul>${section.keyPoints.map((kp) => `<li>${escapeHtml(kp)}</li>`).join('')}</ul>`
      );
    }
  }
  return parts.join('');
}

/**
 * Mounts an offscreen capture host at viewport origin (0, 0) wrapped in a 1x1 hidden veil.
 * Keeping the target at (0, 0) with opacity: 1 prevents blank rasterization on Android WebView
 * and Chromium builds that skip off-viewport layers at left: -12000px.
 */
function createOffscreenHost(
  widthPx: number,
  extraCss = ''
): { wrapper: HTMLDivElement; host: HTMLDivElement } {
  const wrapper = document.createElement('div');
  wrapper.style.position = 'fixed';
  wrapper.style.left = '0';
  wrapper.style.top = '0';
  wrapper.style.width = '1px';
  wrapper.style.height = '1px';
  wrapper.style.overflow = 'hidden';
  wrapper.style.opacity = '0.001';
  wrapper.style.pointerEvents = 'none';
  wrapper.style.zIndex = '-9999';

  const host = document.createElement('div');
  host.style.position = 'relative';
  host.style.width = `${widthPx}px`;
  host.style.backgroundColor = '#ffffff';
  host.style.color = '#1e293b';
  host.style.opacity = '1';

  if (extraCss) {
    const styleEl = document.createElement('style');
    styleEl.textContent = extraCss;
    host.appendChild(styleEl);
  }

  wrapper.appendChild(host);
  document.body.appendChild(wrapper);
  return { wrapper, host };
}

function preloadImageUrl(src: string | undefined): Promise<void> {
  if (!src) {
    return Promise.resolve();
  }
  return new Promise<void>((resolve) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    const done = () => {
      img.onload = null;
      img.onerror = null;
      resolve();
    };
    img.onload = done;
    img.onerror = done;
    img.src = src;
    window.setTimeout(done, 2500);
  });
}

async function waitForPaintAndAssets(
  root: HTMLElement,
  extraImageUrls: string[] = []
): Promise<void> {
  if (document.fonts?.ready) {
    try {
      await document.fonts.ready;
    } catch {
      // Ignore font readiness errors
    }
  }
  const images = Array.from(root.querySelectorAll('img'));
  await Promise.all([
    ...images.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete && img.naturalWidth > 0) {
            resolve();
            return;
          }
          const done = () => {
            img.removeEventListener('load', done);
            img.removeEventListener('error', done);
            resolve();
          };
          img.addEventListener('load', done);
          img.addEventListener('error', done);
          window.setTimeout(done, 2500);
        })
    ),
    ...extraImageUrls.map((url) => preloadImageUrl(url)),
  ]);

  await new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => resolve());
    });
  });
}

function computeSafeCaptureScale(
  targetEl: HTMLElement,
  desiredScale = 2
): number {
  const h = Math.max(1, targetEl.scrollHeight || targetEl.offsetHeight);
  const w = Math.max(1, targetEl.scrollWidth || targetEl.offsetWidth);
  // Keep canvas within safe WebView GPU limits (<= 14,000 px tall and <= 40M pixels)
  const maxHeightScale = 14000 / h;
  const maxAreaScale = Math.sqrt(36_000_000 / (w * h));
  return Math.max(1, Math.min(desiredScale, maxHeightScale, maxAreaScale));
}

async function appendCanvasSlicedToPdf(
  pdfDoc: PDFDocument,
  canvas: HTMLCanvasElement,
  pageWidthPt: number,
  pageHeightPt: number,
  marginPt = 0
): Promise<void> {
  if (canvas.width <= 0 || canvas.height <= 0) {
    throw new Error('Khung hình xuất PDF bị rỗng. Vui lòng thử lại.');
  }

  const contentWidthPt = pageWidthPt - marginPt * 2;
  const contentHeightPt = pageHeightPt - marginPt * 2;
  const pxPerPt = canvas.width / contentWidthPt;
  const sliceHeightPx = Math.max(1, Math.floor(contentHeightPt * pxPerPt));

  let offsetY = 0;
  while (offsetY < canvas.height) {
    const currentSliceHeight = Math.min(
      sliceHeightPx,
      canvas.height - offsetY
    );
    const sliceCanvas = document.createElement('canvas');
    sliceCanvas.width = canvas.width;
    sliceCanvas.height = currentSliceHeight;
    const ctx = sliceCanvas.getContext('2d');
    if (!ctx) {
      throw new Error('Không thể khởi tạo bộ đệm hình ảnh để xuất PDF.');
    }
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
    ctx.drawImage(
      canvas,
      0,
      offsetY,
      canvas.width,
      currentSliceHeight,
      0,
      0,
      canvas.width,
      currentSliceHeight
    );

    const dataUrl = sliceCanvas.toDataURL('image/jpeg', 0.92);
    const embeddedImage = await pdfDoc.embedJpg(dataUrl);
    const page = pdfDoc.addPage([pageWidthPt, pageHeightPt]);
    const drawHeightPt = currentSliceHeight / pxPerPt;

    page.drawImage(embeddedImage, {
      x: marginPt,
      y: pageHeightPt - marginPt - drawHeightPt,
      width: contentWidthPt,
      height: drawHeightPt,
    });

    // Release temporary slice canvas memory immediately
    sliceCanvas.width = 0;
    sliceCanvas.height = 0;

    offsetY += currentSliceHeight;
  }
}

async function renderWordDocumentToPdf(
  pdfDoc: PDFDocument,
  doc: DocumentItem
): Promise<void> {
  const htmlBody = doc.wordHtml?.trim()
    ? sanitizeHtmlFragment(doc.wordHtml)
    : buildFallbackWordHtml(doc);

  const printCss = `
    .lite-pdf-word {
      width: 794px;
      box-sizing: border-box;
      padding: 52px 58px;
      background: #ffffff;
      color: #1e293b;
      font-family: 'Times New Roman', 'Be Vietnam Pro', 'Inter', serif;
      font-size: 16px;
      line-height: 1.7;
      word-break: break-word;
    }
    .lite-pdf-word h1 { font-size: 26px; font-weight: 700; margin: 0 0 14px 0; color: #0f172a; line-height: 1.3; }
    .lite-pdf-word h2 { font-size: 20px; font-weight: 700; margin: 22px 0 10px 0; padding-bottom: 6px; border-bottom: 1px solid #e2e8f0; color: #0f172a; line-height: 1.35; }
    .lite-pdf-word h3 { font-size: 17px; font-weight: 700; margin: 16px 0 8px 0; color: #1e293b; }
    .lite-pdf-word p { margin: 0 0 11px 0; }
    .lite-pdf-word ul, .lite-pdf-word ol { margin: 0 0 12px 0; padding-left: 24px; }
    .lite-pdf-word li { margin-bottom: 5px; }
    .lite-pdf-word blockquote { margin: 12px 0; padding: 10px 14px; border-left: 4px solid #4f46e5; background: #f8fafc; color: #334155; border-radius: 0 8px 8px 0; }
    .lite-pdf-word table { width: 100%; border-collapse: collapse; margin: 14px 0; font-size: 14px; }
    .lite-pdf-word th, .lite-pdf-word td { border: 1px solid #cbd5e1; padding: 7px 10px; text-align: left; vertical-align: top; }
    .lite-pdf-word th { background: #f1f5f9; font-weight: 700; }
    .lite-pdf-word img { max-width: 100%; height: auto; margin: 10px 0; border-radius: 6px; }
  `;

  const { wrapper, host } = createOffscreenHost(794, printCss);
  try {
    const bodyContainer = document.createElement('div');
    bodyContainer.className = 'lite-pdf-word';
    bodyContainer.innerHTML = htmlBody;
    host.appendChild(bodyContainer);

    await waitForPaintAndAssets(bodyContainer);

    const captureScale = computeSafeCaptureScale(bodyContainer, 2);
    const canvas = await html2canvas(bodyContainer, {
      scale: captureScale,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false,
      windowWidth: 794,
      scrollX: 0,
      scrollY: 0,
    });

    await appendCanvasSlicedToPdf(
      pdfDoc,
      canvas,
      A4_PORTRAIT_WIDTH_PT,
      A4_PORTRAIT_HEIGHT_PT,
      0
    );
  } finally {
    wrapper.remove();
  }
}

function formatExcelPdfCellValue(
  val: string | number | undefined,
  numFmt?: string
): string {
  if (val === undefined || val === null) {
    return '';
  }
  if (typeof val !== 'number') {
    return String(val);
  }
  if (!Number.isFinite(val)) {
    return '';
  }
  const fmt = numFmt?.trim();
  if (!fmt || fmt === 'General' || fmt === '@') {
    return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(
      val
    );
  }
  const first = fmt.split(';')[0] ?? fmt;
  const currencyTag = /\[\$([^\]-]+)(?:-[^\]]+)?\]/.exec(first);
  const stripped = first
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .trim();

  if (stripped.includes('%')) {
    const decMatch = /\.([0#]+)/.exec(stripped);
    const dec = decMatch ? Math.min(4, decMatch[1].length) : 0;
    return `${new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: dec,
      maximumFractionDigits: dec,
    }).format(val * 100)}%`;
  }

  if (/[dy]/i.test(stripped) && !/[0#]/.test(stripped) && val >= 1 && val <= 2_958_465) {
    const utcMs = Date.UTC(1899, 11, 30) + Math.floor(val) * 86_400_000;
    const d = new Date(utcMs);
    if (!Number.isNaN(d.getTime())) {
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      return `${dd}/${mm}/${d.getUTCFullYear()}`;
    }
  }

  const currencySym =
    currencyTag?.[1] ??
    /([₫$€£¥])/.exec(first)?.[1] ??
    (/\bVND\b/i.test(first) ? 'VND' : null);
  const decMatch = /\.([0#]+)/.exec(stripped);
  const dec = decMatch ? Math.min(4, decMatch[1].length) : 0;
  if (currencySym) {
    const numStr = new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: dec,
      maximumFractionDigits: dec,
    }).format(val);
    const symIdx = first.indexOf(currencySym);
    const digitIdx = first.search(/[0#]/);
    return symIdx >= 0 && digitIdx >= 0 && symIdx < digitIdx
      ? `${currencySym}${numStr}`
      : `${numStr} ${currencySym}`;
  }
  if (/[0#]/.test(stripped)) {
    return new Intl.NumberFormat('vi-VN', {
      minimumFractionDigits: dec,
      maximumFractionDigits: dec,
    }).format(val);
  }
  return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 2 }).format(
    val
  );
}

function buildExcelCellCss(
  style: CellStyle | undefined,
  isNum: boolean
): string {
  const parts: string[] = [];
  const align = style?.align ?? (isNum ? 'right' : 'left');
  parts.push(`text-align:${align};`);
  if (style?.valign) {
    parts.push(`vertical-align:${style.valign};`);
  } else {
    parts.push('vertical-align:middle;');
  }
  if (style?.bold) {
    parts.push('font-weight:700;');
  }
  if (style?.italic) {
    parts.push('font-style:italic;');
  }
  if (style?.underline) {
    parts.push('text-decoration:underline;');
  }
  if (style?.bgColor) {
    parts.push(`background-color:${style.bgColor};`);
  }
  if (style?.textColor) {
    parts.push(`color:${style.textColor};`);
  }
  if (
    typeof style?.fontSize === 'number' &&
    Number.isFinite(style.fontSize) &&
    style.fontSize > 0
  ) {
    parts.push(
      `font-size:${Math.max(9, Math.min(36, Math.round(style.fontSize)))}px;`
    );
  }
  if (style?.fontFamily) {
    parts.push(`font-family:${style.fontFamily};`);
  }
  if (style?.wrapText) {
    parts.push('white-space:pre-wrap;word-break:break-word;');
  }
  if (style?.borderTop) {
    parts.push(`border-top:${style.borderTop};`);
  }
  if (style?.borderRight) {
    parts.push(`border-right:${style.borderRight};`);
  }
  if (style?.borderBottom) {
    parts.push(`border-bottom:${style.borderBottom};`);
  }
  if (style?.borderLeft) {
    parts.push(`border-left:${style.borderLeft};`);
  }
  return parts.join('');
}

async function renderExcelSheetsToPdf(
  pdfDoc: PDFDocument,
  doc: DocumentItem
): Promise<void> {
  const sheets: SpreadsheetSheet[] =
    doc.sheets && doc.sheets.length > 0
      ? doc.sheets
      : [
          {
            id: 'sheet-1',
            name: 'Sheet 1',
            columns: ['A', 'B', 'C'],
            rows: [[doc.title, doc.contentPreview || 'Bảng tính trống', '']],
          },
        ];

  for (const sheet of sheets) {
    const colCount = Math.max(
      sheet.columns.length,
      ...sheet.rows.map((r) => r.length),
      1
    );
    const totalColWidthPx =
      sheet.colWidths && sheet.colWidths.length > 0
        ? sheet.colWidths.reduce((sum, w) => sum + (w > 20 ? w : 100), 36)
        : colCount * 120 + 36;
    const useLandscape = colCount > 5 || totalColWidthPx > 720;
    const hostWidthPx = useLandscape ? 1122 : 794;
    const pageW = useLandscape ? A4_LANDSCAPE_WIDTH_PT : A4_PORTRAIT_WIDTH_PT;
    const pageH = useLandscape ? A4_LANDSCAPE_HEIGHT_PT : A4_PORTRAIT_HEIGHT_PT;

    const evaluated = evaluateSheetMatrix(sheet.rows);

    // Build merge lookup respecting sheet.merges
    const mergeOrigins = new Map<string, { rowSpan: number; colSpan: number }>();
    const mergeCovered = new Set<string>();
    if (sheet.merges && sheet.merges.length > 0) {
      const maxRow = Math.max(0, evaluated.length - 1);
      const maxCol = colCount - 1;
      for (const m of sheet.merges) {
        const sR = Math.max(0, Math.min(maxRow, m.s.r));
        const eR = Math.max(sR, Math.min(maxRow, m.e.r));
        const sC = Math.max(0, Math.min(maxCol, m.s.c));
        const eC = Math.max(sC, Math.min(maxCol, m.e.c));
        if (eR === sR && eC === sC) {
          continue;
        }
        const originKey = `${sR}:${sC}`;
        if (mergeCovered.has(originKey)) {
          continue;
        }
        mergeOrigins.set(originKey, {
          rowSpan: eR - sR + 1,
          colSpan: eC - sC + 1,
        });
        for (let r = sR; r <= eR; r += 1) {
          for (let c = sC; c <= eC; c += 1) {
            if (r === sR && c === sC) {
              continue;
            }
            mergeCovered.add(`${r}:${c}`);
          }
        }
      }
    }

    const colgroupHtml = (() => {
      const cols: string[] = ['<col style="width:36px;" />'];
      const availableW = hostWidthPx - 80 - 36;
      const rawWidths = Array.from({ length: colCount }, (_, idx) => {
        const w = sheet.colWidths?.[idx];
        return typeof w === 'number' && Number.isFinite(w) && w > 20 ? w : 110;
      });
      const sumW = rawWidths.reduce((acc, w) => acc + w, 0) || 1;
      for (const w of rawWidths) {
        const pct = ((w / sumW) * 100).toFixed(2);
        const scaledPx = Math.max(42, Math.round((w / sumW) * availableW));
        cols.push(`<col style="width:${pct}%;min-width:${scaledPx}px;" />`);
      }
      return `<colgroup>${cols.join('')}</colgroup>`;
    })();

    const printCss = `
      .lite-pdf-sheet {
        width: ${hostWidthPx}px;
        box-sizing: border-box;
        padding: 36px 40px;
        background: #ffffff;
        color: #0f172a;
        font-family: 'Be Vietnam Pro', 'Inter', system-ui, sans-serif;
      }
      .lite-pdf-sheet-head {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        margin-bottom: 14px;
        padding-bottom: 10px;
        border-bottom: 2px solid #15803d;
      }
      .lite-pdf-sheet-title { font-size: 18px; font-weight: 700; color: #0f172a; margin: 0; }
      .lite-pdf-sheet-sub { font-size: 12px; font-weight: 600; color: #15803d; }
      .lite-pdf-sheet table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12px; }
      .lite-pdf-sheet th, .lite-pdf-sheet td { border: 1px solid #cbd5e1; padding: 5px 7px; word-break: break-word; overflow-wrap: break-word; }
      .lite-pdf-sheet thead th { background: #f1f5f9; color: #475569; font-weight: 700; text-align: center; font-size: 11px; }
      .lite-pdf-sheet tbody td.row-num { background: #f8fafc; color: #64748b; text-align: center; font-weight: 600; width: 36px; }
    `;

    const { wrapper, host } = createOffscreenHost(hostWidthPx, printCss);
    try {
      const container = document.createElement('div');
      container.className = 'lite-pdf-sheet';

      const headerRowHtml = Array.from({ length: colCount }, (_, cIdx) => {
        const letter = sheet.columns[cIdx] ?? columnIndexToLetter(cIdx);
        return `<th>${escapeHtml(letter)}</th>`;
      }).join('');

      const rowsHtml = evaluated
        .map((row, rIdx) => {
          const rowHeight = sheet.rowHeights?.[rIdx];
          const trStyle =
            typeof rowHeight === 'number' && rowHeight > 16
              ? ` style="height:${Math.round(rowHeight)}px;"`
              : '';
          const cells = Array.from({ length: colCount }, (__, cIdx) => {
            const key = `${rIdx}:${cIdx}`;
            if (mergeCovered.has(key)) {
              return '';
            }
            const span = mergeOrigins.get(key);
            const spanAttrs = span
              ? `${span.rowSpan > 1 ? ` rowspan="${span.rowSpan}"` : ''}${
                  span.colSpan > 1 ? ` colspan="${span.colSpan}"` : ''
                }`
              : '';
            const val = row[cIdx];
            const style = sheet.cellStyles?.[key];
            const isNum = typeof val === 'number';
            const cellCss = buildExcelCellCss(style, isNum);
            const formatted = formatExcelPdfCellValue(val, style?.numFmt);
            return `<td${spanAttrs} style="${cellCss}">${escapeHtml(formatted)}</td>`;
          }).join('');
          return `<tr${trStyle}><td class="row-num">${rIdx + 1}</td>${cells}</tr>`;
        })
        .join('');

      container.innerHTML = `
        <div class="lite-pdf-sheet-head">
          <h1 class="lite-pdf-sheet-title">${escapeHtml(doc.title)}</h1>
          <span class="lite-pdf-sheet-sub">${escapeHtml(sheet.name)}</span>
        </div>
        <table>
          ${colgroupHtml}
          <thead><tr><th>#</th>${headerRowHtml}</tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      `;
      host.appendChild(container);

      await waitForPaintAndAssets(container);

      const captureScale = computeSafeCaptureScale(container, 2);
      const canvas = await html2canvas(container, {
        scale: captureScale,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false,
        windowWidth: hostWidthPx,
        scrollX: 0,
        scrollY: 0,
      });

      await appendCanvasSlicedToPdf(pdfDoc, canvas, pageW, pageH, 0);
    } finally {
      wrapper.remove();
    }
  }
}

function renderRichParagraphsMarkup(
  block: NonNullable<SlidePreviewItem['blocks']>[number],
  scaleRatio: number
): string {
  if (!block.richParagraphs || block.richParagraphs.length === 0) {
    return escapeHtml(block.text);
  }
  return block.richParagraphs
    .map((para) => {
      const pAlign = para.align ?? block.align ?? 'left';
      const bullet = para.bullet ? `${escapeHtml(para.bullet)} ` : '';
      const indentPx = para.level && para.level > 0 ? para.level * 18 : 0;
      const runsHtml = para.runs
        .map((run) => {
          const fs = Math.max(
            10,
            Math.round((run.fontSize || block.fontSize || 16) * scaleRatio)
          );
          const fw = run.bold ?? block.bold ? 700 : 400;
          const fst = run.italic ?? block.italic ? 'italic' : 'normal';
          const td = run.underline ? 'underline' : 'none';
          const col = run.color ?? block.color ?? '#FFFFFF';
          const ff = run.fontFamily ?? block.fontFamily ?? "'Be Vietnam Pro','Inter',sans-serif";
          return `<span style="font-size:${fs}px;font-weight:${fw};font-style:${fst};text-decoration:${td};color:${col};font-family:${ff};">${escapeHtml(run.text)}</span>`;
        })
        .join('');
      return `<div style="text-align:${pAlign};padding-left:${indentPx}px;margin-bottom:4px;line-height:1.35;">${bullet}${runsHtml}</div>`;
    })
    .join('');
}

function buildSlideHtmlMarkup(
  slide: SlidePreviewItem,
  doc: DocumentItem,
  widthPx: number,
  heightPx: number
): string {
  const accent = slide.accentColor || '#4F46E5';
  const bgStyle = slide.backgroundImageSrc
    ? `background-color:${slide.backgroundColor || '#ffffff'};background-image:url('${slide.backgroundImageSrc}');background-size:cover;background-position:center;`
    : slide.backgroundGradient
      ? `background-image:${slide.backgroundGradient};`
      : slide.backgroundColor
        ? `background-color:${slide.backgroundColor};`
        : `background:radial-gradient(circle at 86% 16%, ${accent}38 0%, transparent 48%), linear-gradient(145deg, #161D2E 0%, #1F293F 58%, #131926 100%);`;

  const hasPositionedContent =
    Boolean(slide.isMasterPositioned) ||
    Boolean(slide.blocks?.some((b) => b.xPct !== undefined && b.yPct !== undefined)) ||
    Boolean(slide.images?.some((img) => img.xPct !== undefined && img.yPct !== undefined));

  const scaleRatio = widthPx / 960;

  if (hasPositionedContent) {
    const imagesHtml = (slide.images ?? [])
      .map((img) => {
        const x = img.xPct ?? 55;
        const y = img.yPct ?? 18;
        const w = img.wPct ?? 36;
        const h = img.hPct ?? 52;
        const rot = img.rotationDeg ? `transform:rotate(${img.rotationDeg}deg);` : '';
        return `<img src="${escapeHtml(img.src)}" alt="${escapeHtml(img.alt ?? '')}" style="position:absolute;left:${x}%;top:${y}%;width:${w}%;height:${h}%;object-fit:contain;border-radius:${img.borderRadius ?? '6px'};${rot}" />`;
      })
      .join('');

    const blocksHtml = (slide.blocks ?? [])
      .map((block) => {
        const x = block.xPct ?? 6;
        const y = block.yPct ?? 12;
        const w = block.wPct ?? 88;
        const h = block.hPct ?? 20;
        const justifyV =
          block.valign === 'middle'
            ? 'center'
            : block.valign === 'bottom'
              ? 'flex-end'
              : 'flex-start';

        if (block.kind === 'table' && block.tableRows && block.tableRows.length > 0) {
          const rowsMarkup = block.tableRows
            .map((row, rIdx) => {
              const cells = row
                .map((cell, cIdx) => {
                  const cs = block.tableCellStyles?.[rIdx]?.[cIdx];
                  const cellBg = cs?.fillColor ? `background:${cs.fillColor};` : '';
                  const cellColor = cs?.textColor
                    ? `color:${cs.textColor};`
                    : `color:${block.color ?? '#0F172A'};`;
                  const cellBold = cs?.bold ? 'font-weight:700;' : '';
                  const cellAlign = cs?.align ? `text-align:${cs.align};` : '';
                  return `<td style="border:1px solid rgba(148,163,184,0.45);padding:6px 8px;${cellBg}${cellColor}${cellBold}${cellAlign}">${escapeHtml(cell)}</td>`;
                })
                .join('');
              return `<tr>${cells}</tr>`;
            })
            .join('');
          return `<div style="position:absolute;left:${x}%;top:${y}%;width:${w}%;min-height:${h}%;overflow:visible;"><table style="width:100%;border-collapse:collapse;font-size:${Math.round((block.fontSize || 14) * scaleRatio)}px;">${rowsMarkup}</table></div>`;
        }

        const scaledFont = Math.max(
          11,
          Math.round((block.fontSize || 18) * scaleRatio)
        );
        const borderCss = block.borderColor
          ? `border:${block.borderWidthPx ?? 1}px solid ${block.borderColor};`
          : '';
        const fillCss = block.fillGradient
          ? `background:${block.fillGradient};`
          : block.fillColor
            ? `background-color:${block.fillColor};`
            : '';
        const rotCss = block.rotationDeg
          ? `transform:rotate(${block.rotationDeg}deg);`
          : '';
        const innerContent = renderRichParagraphsMarkup(block, scaleRatio);

        return `<div style="position:absolute;left:${x}%;top:${y}%;width:${w}%;min-height:${h}%;display:flex;flex-direction:column;justify-content:${justifyV};padding:6px 10px;box-sizing:border-box;overflow:visible;border-radius:${block.borderRadius ?? '6px'};${fillCss}${borderCss}${rotCss}font-family:${block.fontFamily ?? "'Be Vietnam Pro','Inter',sans-serif"};font-size:${scaledFont}px;font-weight:${block.bold ? 700 : 400};font-style:${block.italic ? 'italic' : 'normal'};color:${block.color ?? '#FFFFFF'};text-align:${block.align ?? 'left'};white-space:pre-wrap;word-break:break-word;line-height:1.35;">${innerContent}</div>`;
      })
      .join('');

    return `<div style="position:relative;width:${widthPx}px;height:${heightPx}px;overflow:hidden;${bgStyle}">${imagesHtml}${blocksHtml}</div>`;
  }

  // Fallback structured slide card
  const bulletsMarkup = slide.bullets
    .map(
      (b) =>
        `<li style="margin-bottom:12px;display:flex;align-items:flex-start;gap:12px;"><span style="width:9px;height:9px;border-radius:9999px;background:${accent};margin-top:10px;flex-shrink:0;"></span><span>${escapeHtml(b)}</span></li>`
    )
    .join('');

  return `
    <div style="position:relative;width:${widthPx}px;height:${heightPx}px;overflow:hidden;padding:56px 64px;box-sizing:border-box;display:flex;flex-direction:column;justify-content:space-between;color:#ffffff;font-family:'Be Vietnam Pro','Inter',sans-serif;${bgStyle}">
      <div style="display:flex;align-items:center;justify-content:space-between;">
        <span style="display:inline-block;width:44px;height:6px;border-radius:9999px;background:${accent};"></span>
        <span style="font-size:15px;font-weight:600;color:#cbd5e1;">Slide ${slide.index} · ${escapeHtml(doc.title)}</span>
      </div>
      <div style="margin:auto 0;">
        <h1 style="margin:0 0 14px 0;font-size:40px;font-weight:700;line-height:1.25;">${escapeHtml(slide.title)}</h1>
        ${slide.subtitle ? `<p style="margin:0 0 24px 0;font-size:21px;color:#cbd5e1;">${escapeHtml(slide.subtitle)}</p>` : ''}
        <ul style="list-style:none;padding:0;margin:0;font-size:22px;line-height:1.5;color:#f1f5f9;">${bulletsMarkup}</ul>
      </div>
      <div style="height:4px;background:rgba(255,255,255,0.12);border-radius:9999px;"></div>
    </div>
  `;
}

async function renderSlidesToPdf(
  pdfDoc: PDFDocument,
  doc: DocumentItem
): Promise<void> {
  const slides: SlidePreviewItem[] =
    doc.slidesPreview.length > 0
      ? doc.slidesPreview
      : [
          {
            id: `${doc.id}-s1`,
            index: 1,
            title: doc.title,
            bullets: [doc.contentPreview || 'Nội dung trình chiếu'],
            accentColor: '#4F46E5',
          },
        ];

  const baseWidthPx = 1280;
  for (const slide of slides) {
    const rawRatio = slide.slideWidthRatio ?? doc.slideAspectRatio ?? 16 / 9;
    const ratio =
      Number.isFinite(rawRatio) && rawRatio >= 1.1 && rawRatio <= 2.4
        ? rawRatio
        : 16 / 9;
    const heightPx = Math.max(480, Math.round(baseWidthPx / ratio));
    const pageW = A4_LANDSCAPE_WIDTH_PT;
    const pageH = Math.round((A4_LANDSCAPE_WIDTH_PT / ratio) * 100) / 100;

    const { wrapper, host } = createOffscreenHost(baseWidthPx);
    try {
      const slideEl = document.createElement('div');
      slideEl.innerHTML = buildSlideHtmlMarkup(
        slide,
        doc,
        baseWidthPx,
        heightPx
      );
      const targetNode = (slideEl.firstElementChild as HTMLElement) ?? slideEl;
      host.appendChild(targetNode);

      const bgUrls = slide.backgroundImageSrc ? [slide.backgroundImageSrc] : [];
      await waitForPaintAndAssets(targetNode, bgUrls);

      const canvas = await html2canvas(targetNode, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false,
        windowWidth: baseWidthPx,
        windowHeight: heightPx,
        scrollX: 0,
        scrollY: 0,
      });

      if (canvas.width <= 0 || canvas.height <= 0) {
        throw new Error(`Không thể chụp nội dung Slide ${slide.index}.`);
      }

      const jpgDataUrl = canvas.toDataURL('image/jpeg', 0.93);
      const embedded = await pdfDoc.embedJpg(jpgDataUrl);
      const pdfPage = pdfDoc.addPage([pageW, pageH]);
      pdfPage.drawImage(embedded, {
        x: 0,
        y: 0,
        width: pageW,
        height: pageH,
      });
    } finally {
      wrapper.remove();
    }
  }
}

export async function exportDocumentToPdf(
  doc: DocumentItem
): Promise<{ blob: Blob; fileName: string }> {
  const fileName = sanitizePdfFileName(doc.title, 'tai-lieu');

  if (doc.type === 'pdf') {
    if (doc.sourceFile) {
      const buffer = await doc.sourceFile.arrayBuffer();
      return {
        blob: new Blob([buffer], { type: 'application/pdf' }),
        fileName,
      };
    }
    if (doc.rawArrayBuffer && doc.rawArrayBuffer.byteLength > 0) {
      return {
        blob: new Blob([doc.rawArrayBuffer.slice(0)], {
          type: 'application/pdf',
        }),
        fileName,
      };
    }
    const existingUrl = doc.fileUrl ?? doc.objectUrl;
    if (existingUrl) {
      const response = await fetch(existingUrl);
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        return {
          blob: new Blob([buffer], { type: 'application/pdf' }),
          fileName,
        };
      }
    }
    throw new Error(
      'Tệp này đã là định dạng PDF. Hãy dùng tab Tách PDF hoặc Gộp PDF để xử lý.'
    );
  }

  const pdfDoc = await PDFDocument.create();
  if (doc.type === 'docx') {
    await renderWordDocumentToPdf(pdfDoc, doc);
  } else if (doc.type === 'xlsx') {
    await renderExcelSheetsToPdf(pdfDoc, doc);
  } else if (doc.type === 'pptx') {
    await renderSlidesToPdf(pdfDoc, doc);
  }

  const pdfBytes = await pdfDoc.save();
  const blob = new Blob([new Uint8Array(pdfBytes)], {
    type: 'application/pdf',
  });
  return { blob, fileName };
}

export async function downloadDocumentAsPdf(
  doc: DocumentItem
): Promise<PdfDownloadResult> {
  const result = await exportDocumentToPdf(doc);
  const saveRes = await saveFileUniversal(result.blob, result.fileName, {
    mimeType: 'application/pdf',
    shareTitle: doc.title || result.fileName,
  });
  return {
    blob: result.blob,
    fileName: result.fileName,
    method: saveRes.method,
    savedPath: saveRes.savedPath,
    message: saveRes.message,
  };
}

export async function inspectPdfPageCount(file: File): Promise<number> {
  const buffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
  return pdfDoc.getPageCount();
}

export async function mergePdfFiles(
  files: File[],
  outputTitle = 'Tai-lieu-gop',
  options?: { download?: boolean }
): Promise<MergePdfResult> {
  if (files.length < 2) {
    throw new Error('Vui lòng chọn ít nhất 2 tệp PDF để gộp.');
  }

  const mergedPdf = await PDFDocument.create();
  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const srcPdf = await PDFDocument.load(buffer, { ignoreEncryption: true });
    const indices = srcPdf.getPageIndices();
    const copiedPages = await mergedPdf.copyPages(srcPdf, indices);
    copiedPages.forEach((page) => mergedPdf.addPage(page));
  }

  const pdfBytes = await mergedPdf.save();
  const blob = new Blob([new Uint8Array(pdfBytes)], {
    type: 'application/pdf',
  });
  const fileName = sanitizePdfFileName(outputTitle, 'Tai-lieu-gop');
  const totalPages = mergedPdf.getPageCount();

  if (options?.download) {
    const saveRes = await saveFileUniversal(blob, fileName, {
      mimeType: 'application/pdf',
      shareTitle: fileName,
    });
    return {
      blob,
      fileName,
      totalPages,
      method: saveRes.method,
      savedPath: saveRes.savedPath,
      message: saveRes.message,
    };
  }

  return {
    blob,
    fileName,
    totalPages,
    message: `Đã gộp ${totalPages} trang thành tệp "${fileName}".`,
  };
}

export function parsePageRangeExpression(
  rangeInput: string,
  totalPages: number
): number[] {
  const cleaned = rangeInput.trim();
  if (!cleaned) {
    throw new Error('Vui lòng nhập khoảng trang cần tách (ví dụ: 1-3, 5).');
  }
  if (totalPages <= 0) {
    throw new Error('Tệp PDF không có trang hợp lệ.');
  }

  const parts = cleaned
    .split(',')
    .map((p) => p.trim())
    .filter(Boolean);

  if (parts.length === 0) {
    throw new Error('Khoảng trang không hợp lệ.');
  }

  const zeroBasedSet = new Set<number>();
  const orderedZeroBased: number[] = [];

  for (const part of parts) {
    const rangeMatch = /^(\d+)\s*-\s*(\d+)$/.exec(part);
    if (rangeMatch) {
      const start = Number.parseInt(rangeMatch[1], 10);
      const end = Number.parseInt(rangeMatch[2], 10);
      if (start < 1 || end < 1 || start > totalPages || end > totalPages) {
        throw new Error(
          `Khoảng trang "${part}" vượt quá giới hạn (1–${totalPages}).`
        );
      }
      const step = start <= end ? 1 : -1;
      for (
        let page = start;
        step > 0 ? page <= end : page >= end;
        page += step
      ) {
        const zeroIdx = page - 1;
        if (!zeroBasedSet.has(zeroIdx)) {
          zeroBasedSet.add(zeroIdx);
          orderedZeroBased.push(zeroIdx);
        }
      }
      continue;
    }

    if (/^\d+$/.test(part)) {
      const pageNum = Number.parseInt(part, 10);
      if (pageNum < 1 || pageNum > totalPages) {
        throw new Error(
          `Số trang "${pageNum}" nằm ngoài phạm vi (1–${totalPages}).`
        );
      }
      const zeroIdx = pageNum - 1;
      if (!zeroBasedSet.has(zeroIdx)) {
        zeroBasedSet.add(zeroIdx);
        orderedZeroBased.push(zeroIdx);
      }
      continue;
    }

    throw new Error(
      `Cú pháp "${part}" chưa đúng. Chỉ nhập số và dấu gạch ngang (vd: 1-3, 5).`
    );
  }

  if (orderedZeroBased.length === 0) {
    throw new Error('Không tìm thấy trang nào hợp lệ để tách.');
  }

  return orderedZeroBased;
}

export async function splitPdfFileByRanges(
  file: File,
  rangeInput: string,
  options?: { download?: boolean }
): Promise<SplitPdfResult> {
  const buffer = await file.arrayBuffer();
  const srcPdf = await PDFDocument.load(buffer, { ignoreEncryption: true });
  const totalSourcePages = srcPdf.getPageCount();
  const zeroBasedIndices = parsePageRangeExpression(
    rangeInput,
    totalSourcePages
  );

  const outPdf = await PDFDocument.create();
  const copiedPages = await outPdf.copyPages(srcPdf, zeroBasedIndices);
  copiedPages.forEach((page) => outPdf.addPage(page));

  const pdfBytes = await outPdf.save();
  const blob = new Blob([new Uint8Array(pdfBytes)], {
    type: 'application/pdf',
  });
  const baseName = file.name.replace(/\.pdf$/i, '') || 'tai-lieu';
  const fileName = sanitizePdfFileName(`${baseName}-tach-trang`, 'tai-lieu-tach');
  const pageCount = outPdf.getPageCount();

  if (options?.download) {
    const saveRes = await saveFileUniversal(blob, fileName, {
      mimeType: 'application/pdf',
      shareTitle: fileName,
    });
    return {
      blob,
      fileName,
      pageCount,
      totalSourcePages,
      method: saveRes.method,
      savedPath: saveRes.savedPath,
      message: saveRes.message,
    };
  }

  return {
    blob,
    fileName,
    pageCount,
    totalSourcePages,
    message: `Đã tách ${pageCount}/${totalSourcePages} trang thành tệp "${fileName}".`,
  };
}

import html2canvas from 'html2canvas';
import { PDFDocument } from 'pdf-lib';
import type { DocumentItem, SlidePreviewItem, SpreadsheetSheet } from '../types/document';
import { evaluateSheetMatrix } from './formulaEvaluator';

const A4_PORTRAIT_WIDTH_PT = 595.28;
const A4_PORTRAIT_HEIGHT_PT = 841.89;
const A4_LANDSCAPE_WIDTH_PT = 841.89;
const A4_LANDSCAPE_HEIGHT_PT = 595.28;

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

export function triggerBlobDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1500);
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
    const paragraphs = doc.wordParagraphs?.[section.id] ?? (section.summary ? [section.summary] : []);
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

function createOffscreenHost(widthPx: number, extraCss = ''): HTMLDivElement {
  const host = document.createElement('div');
  host.style.position = 'fixed';
  host.style.left = '-12000px';
  host.style.top = '0';
  host.style.width = `${widthPx}px`;
  host.style.backgroundColor = '#ffffff';
  host.style.color = '#1e293b';
  host.style.zIndex = '-1';
  host.style.pointerEvents = 'none';
  if (extraCss) {
    const styleEl = document.createElement('style');
    styleEl.textContent = extraCss;
    host.appendChild(styleEl);
  }
  document.body.appendChild(host);
  return host;
}

async function waitForImagesAndFonts(root: HTMLElement): Promise<void> {
  if (document.fonts?.ready) {
    try {
      await document.fonts.ready;
    } catch {
      // Ignore font readiness errors
    }
  }
  const images = Array.from(root.querySelectorAll('img'));
  await Promise.all(
    images.map(
      (img) =>
        new Promise<void>((resolve) => {
          if (img.complete) {
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
    )
  );
}

async function appendCanvasSlicedToPdf(
  pdfDoc: PDFDocument,
  canvas: HTMLCanvasElement,
  pageWidthPt: number,
  pageHeightPt: number,
  marginPt = 0
): Promise<void> {
  const contentWidthPt = pageWidthPt - marginPt * 2;
  const contentHeightPt = pageHeightPt - marginPt * 2;
  const pxPerPt = canvas.width / contentWidthPt;
  const sliceHeightPx = Math.max(1, Math.floor(contentHeightPt * pxPerPt));

  let offsetY = 0;
  while (offsetY < canvas.height) {
    const currentSliceHeight = Math.min(sliceHeightPx, canvas.height - offsetY);
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

  const host = createOffscreenHost(794, printCss);
  try {
    const bodyContainer = document.createElement('div');
    bodyContainer.className = 'lite-pdf-word';
    bodyContainer.innerHTML = htmlBody;
    host.appendChild(bodyContainer);

    await waitForImagesAndFonts(bodyContainer);

    const canvas = await html2canvas(bodyContainer, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false,
      windowWidth: 794,
    });

    await appendCanvasSlicedToPdf(
      pdfDoc,
      canvas,
      A4_PORTRAIT_WIDTH_PT,
      A4_PORTRAIT_HEIGHT_PT,
      0
    );
  } finally {
    host.remove();
  }
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
    const useLandscape = colCount > 5;
    const hostWidthPx = useLandscape ? 1122 : 794;
    const pageW = useLandscape ? A4_LANDSCAPE_WIDTH_PT : A4_PORTRAIT_WIDTH_PT;
    const pageH = useLandscape ? A4_LANDSCAPE_HEIGHT_PT : A4_PORTRAIT_HEIGHT_PT;

    const evaluated = evaluateSheetMatrix(sheet.rows);
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
      .lite-pdf-sheet table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
      .lite-pdf-sheet th, .lite-pdf-sheet td { border: 1px solid #cbd5e1; padding: 6px 8px; word-break: break-word; }
      .lite-pdf-sheet thead th { background: #f1f5f9; color: #475569; font-weight: 700; text-align: center; font-size: 11px; }
      .lite-pdf-sheet tbody td.row-num { background: #f8fafc; color: #64748b; text-align: center; font-weight: 600; width: 34px; }
    `;

    const host = createOffscreenHost(hostWidthPx, printCss);
    try {
      const container = document.createElement('div');
      container.className = 'lite-pdf-sheet';

      const headerRowHtml = Array.from({ length: colCount }, (_, cIdx) => {
        const letter = sheet.columns[cIdx] ?? String.fromCharCode(65 + cIdx);
        return `<th>${escapeHtml(letter)}</th>`;
      }).join('');

      const rowsHtml = evaluated
        .map((row, rIdx) => {
          const cells = Array.from({ length: colCount }, (__, cIdx) => {
            const val = row[cIdx];
            const style = sheet.cellStyles?.[`${rIdx}:${cIdx}`];
            const isNum = typeof val === 'number';
            const align = style?.align ?? (isNum ? 'right' : 'left');
            const weight = style?.bold ? 'font-weight:700;' : '';
            const formatted =
              val === undefined || val === null
                ? ''
                : isNum
                  ? new Intl.NumberFormat('vi-VN', {
                      maximumFractionDigits: 2,
                    }).format(val)
                  : String(val);
            return `<td style="text-align:${align};${weight}">${escapeHtml(formatted)}</td>`;
          }).join('');
          return `<tr><td class="row-num">${rIdx + 1}</td>${cells}</tr>`;
        })
        .join('');

      container.innerHTML = `
        <div class="lite-pdf-sheet-head">
          <h1 class="lite-pdf-sheet-title">${escapeHtml(doc.title)}</h1>
          <span class="lite-pdf-sheet-sub">${escapeHtml(sheet.name)}</span>
        </div>
        <table>
          <thead><tr><th>#</th>${headerRowHtml}</tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      `;
      host.appendChild(container);

      await waitForImagesAndFonts(container);

      const canvas = await html2canvas(container, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false,
        windowWidth: hostWidthPx,
      });

      await appendCanvasSlicedToPdf(pdfDoc, canvas, pageW, pageH, 0);
    } finally {
      host.remove();
    }
  }
}

function buildSlideHtmlMarkup(
  slide: SlidePreviewItem,
  doc: DocumentItem,
  widthPx: number,
  heightPx: number
): string {
  const accent = slide.accentColor || '#4F46E5';
  const bgStyle = slide.backgroundImageSrc
    ? `background-image:url('${slide.backgroundImageSrc}');background-size:cover;background-position:center;`
    : slide.backgroundGradient
      ? `background-image:${slide.backgroundGradient};`
      : slide.backgroundColor
        ? `background-color:${slide.backgroundColor};`
        : `background:radial-gradient(circle at 86% 16%, ${accent}38 0%, transparent 48%), linear-gradient(145deg, #161D2E 0%, #1F293F 58%, #131926 100%);`;

  const hasPositionedContent =
    Boolean(slide.isMasterPositioned) ||
    Boolean(slide.blocks?.some((b) => b.xPct !== undefined && b.yPct !== undefined)) ||
    Boolean(slide.images?.some((img) => img.xPct !== undefined && img.yPct !== undefined));

  if (hasPositionedContent) {
    const imagesHtml = (slide.images ?? [])
      .map((img) => {
        const x = img.xPct ?? 55;
        const y = img.yPct ?? 18;
        const w = img.wPct ?? 36;
        const h = img.hPct ?? 52;
        return `<img src="${escapeHtml(img.src)}" alt="${escapeHtml(img.alt ?? '')}" style="position:absolute;left:${x}%;top:${y}%;width:${w}%;height:${h}%;object-fit:contain;border-radius:${img.borderRadius ?? '6px'};" />`;
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
          return `<div style="position:absolute;left:${x}%;top:${y}%;width:${w}%;height:${h}%;overflow:hidden;"><table style="width:100%;height:100%;border-collapse:collapse;font-size:${Math.round((block.fontSize || 14) * (widthPx / 960))}px;">${rowsMarkup}</table></div>`;
        }

        const scaledFont = Math.round((block.fontSize || 18) * (widthPx / 960));
        const borderCss = block.borderColor
          ? `border:${block.borderWidthPx ?? 1}px solid ${block.borderColor};`
          : '';
        const fillCss = block.fillGradient
          ? `background:${block.fillGradient};`
          : block.fillColor
            ? `background-color:${block.fillColor};`
            : '';

        return `<div style="position:absolute;left:${x}%;top:${y}%;width:${w}%;height:${h}%;display:flex;flex-direction:column;justify-content:${justifyV};padding:6px 10px;box-sizing:border-box;overflow:hidden;border-radius:${block.borderRadius ?? '6px'};${fillCss}${borderCss}font-family:${block.fontFamily ?? "'Be Vietnam Pro','Inter',sans-serif"};font-size:${scaledFont}px;font-weight:${block.bold ? 700 : 400};font-style:${block.italic ? 'italic' : 'normal'};color:${block.color ?? '#FFFFFF'};text-align:${block.align ?? 'left'};white-space:pre-wrap;line-height:1.3;">${escapeHtml(block.text)}</div>`;
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
    const ratio = slide.slideWidthRatio ?? doc.slideAspectRatio ?? 16 / 9;
    const heightPx = Math.max(480, Math.round(baseWidthPx / ratio));
    const pageW = 960;
    const pageH = Math.round(pageW / ratio);

    const host = createOffscreenHost(baseWidthPx);
    try {
      const slideEl = document.createElement('div');
      slideEl.innerHTML = buildSlideHtmlMarkup(slide, doc, baseWidthPx, heightPx);
      const targetNode = (slideEl.firstElementChild as HTMLElement) ?? slideEl;
      host.appendChild(targetNode);

      await waitForImagesAndFonts(targetNode);

      const canvas = await html2canvas(targetNode, {
        scale: 1.75,
        useCORS: true,
        backgroundColor: '#161D2E',
        logging: false,
        windowWidth: baseWidthPx,
      });

      const jpgDataUrl = canvas.toDataURL('image/jpeg', 0.92);
      const embedded = await pdfDoc.embedJpg(jpgDataUrl);
      const pdfPage = pdfDoc.addPage([pageW, pageH]);
      pdfPage.drawImage(embedded, {
        x: 0,
        y: 0,
        width: pageW,
        height: pageH,
      });
    } finally {
      host.remove();
    }
  }
}

export async function exportDocumentToPdf(
  doc: DocumentItem
): Promise<{ blob: Blob; fileName: string }> {
  const fileName = sanitizePdfFileName(doc.title, 'tai-lieu');

  if (doc.type === 'pdf') {
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
): Promise<{ blob: Blob; fileName: string }> {
  const result = await exportDocumentToPdf(doc);
  triggerBlobDownload(result.blob, result.fileName);
  return result;
}

export async function inspectPdfPageCount(file: File): Promise<number> {
  const buffer = await file.arrayBuffer();
  const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
  return pdfDoc.getPageCount();
}

export async function mergePdfFiles(
  files: File[],
  outputTitle = 'Tai-lieu-gop'
): Promise<{ blob: Blob; fileName: string; totalPages: number }> {
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
  return {
    blob,
    fileName: sanitizePdfFileName(outputTitle, 'Tai-lieu-gop'),
    totalPages: mergedPdf.getPageCount(),
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
      for (let page = start; step > 0 ? page <= end : page >= end; page += step) {
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
  rangeInput: string
): Promise<{
  blob: Blob;
  fileName: string;
  pageCount: number;
  totalSourcePages: number;
}> {
  const buffer = await file.arrayBuffer();
  const srcPdf = await PDFDocument.load(buffer, { ignoreEncryption: true });
  const totalSourcePages = srcPdf.getPageCount();
  const zeroBasedIndices = parsePageRangeExpression(rangeInput, totalSourcePages);

  const outPdf = await PDFDocument.create();
  const copiedPages = await outPdf.copyPages(srcPdf, zeroBasedIndices);
  copiedPages.forEach((page) => outPdf.addPage(page));

  const pdfBytes = await outPdf.save();
  const blob = new Blob([new Uint8Array(pdfBytes)], {
    type: 'application/pdf',
  });
  const baseName = file.name.replace(/\.pdf$/i, '') || 'tai-lieu';

  return {
    blob,
    fileName: sanitizePdfFileName(`${baseName}-tach-trang`, 'tai-lieu-tach'),
    pageCount: outPdf.getPageCount(),
    totalSourcePages,
  };
}

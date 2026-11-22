import type {
  CellStyle,
  DocumentItem,
  DocumentType,
  OutlineSection,
  SlidePreviewItem,
  SpreadsheetSheet,
} from '../types/document';

const DEFAULT_ACCENTS: Record<DocumentType, string> = {
  pptx: '#4F46E5',
  docx: '#2563EB',
  xlsx: '#15803D',
  pdf: '#DB2777',
};

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function truncateText(text: string, maxLength: number): string {
  const normalized = text.trim();
  if (normalized.length <= maxLength) {
    return normalized;
  }
  return `${normalized.slice(0, maxLength - 1).trimEnd()}…`;
}

function sanitizeHexColor(hex: string | undefined, fallback: string): string {
  if (!hex) {
    return fallback;
  }
  const trimmed = hex.trim();
  return /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(trimmed) ? trimmed : fallback;
}

export function createThumbnailDataUri(
  type: DocumentType,
  title: string,
  accentHex?: string
): string {
  const accent = sanitizeHexColor(accentHex, DEFAULT_ACCENTS[type]);
  const safeTitle = escapeXml(truncateText(title || 'Tài liệu mới', 30));
  const fontStack = "Be Vietnam Pro, Inter, system-ui, -apple-system, sans-serif";

  let svg: string;

  if (type === 'pptx') {
    svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 400" width="640" height="400">
  <defs>
    <linearGradient id="slideBg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#1A2030"/>
      <stop offset="100%" stop-color="#273147"/>
    </linearGradient>
    <linearGradient id="accentGlow" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="${accent}" stop-opacity="0.34"/>
      <stop offset="100%" stop-color="${accent}" stop-opacity="0.06"/>
    </linearGradient>
  </defs>
  <rect width="640" height="400" fill="#F3F0EA"/>
  <rect x="24" y="24" width="592" height="352" rx="18" fill="url(#slideBg)"/>
  <circle cx="548" cy="84" r="118" fill="url(#accentGlow)"/>
  <rect x="24" y="24" width="592" height="8" rx="4" fill="${accent}"/>
  <rect x="52" y="58" width="74" height="24" rx="12" fill="${accent}" fill-opacity="0.24"/>
  <text x="89" y="74" text-anchor="middle" font-family="${fontStack}" font-size="11" font-weight="700" fill="#E0E7FF">SLIDE DECK</text>
  <text x="52" y="122" font-family="${fontStack}" font-size="25" font-weight="700" fill="#FFFFFF">${safeTitle}</text>
  <rect x="52" y="140" width="132" height="4" rx="2" fill="${accent}"/>
  <g fill="#CBD5E1" fill-opacity="0.88">
    <circle cx="60" cy="186" r="4" fill="${accent}"/>
    <rect x="74" y="180" width="212" height="10" rx="5" fill="#E2E8F0" fill-opacity="0.85"/>
    <rect x="74" y="198" width="164" height="7" rx="3.5" fill="#94A3B8" fill-opacity="0.55"/>
    <circle cx="60" cy="236" r="4" fill="${accent}"/>
    <rect x="74" y="230" width="196" height="10" rx="5" fill="#E2E8F0" fill-opacity="0.85"/>
    <rect x="74" y="248" width="142" height="7" rx="3.5" fill="#94A3B8" fill-opacity="0.55"/>
    <circle cx="60" cy="286" r="4" fill="${accent}"/>
    <rect x="74" y="280" width="224" height="10" rx="5" fill="#E2E8F0" fill-opacity="0.85"/>
    <rect x="74" y="298" width="128" height="7" rx="3.5" fill="#94A3B8" fill-opacity="0.55"/>
  </g>
  <rect x="358" y="156" width="226" height="174" rx="14" fill="#FFFFFF" fill-opacity="0.07" stroke="#FFFFFF" stroke-opacity="0.14"/>
  <text x="380" y="186" font-family="${fontStack}" font-size="11" font-weight="600" fill="#94A3B8">HIỆU SUẤT TIẾN ĐỘ</text>
  <text x="380" y="220" font-family="${fontStack}" font-size="28" font-weight="700" fill="#FFFFFF">+38%</text>
  <rect x="380" y="274" width="26" height="36" rx="5" fill="${accent}" fill-opacity="0.45"/>
  <rect x="418" y="256" width="26" height="54" rx="5" fill="${accent}" fill-opacity="0.65"/>
  <rect x="456" y="242" width="26" height="68" rx="5" fill="${accent}" fill-opacity="0.85"/>
  <rect x="494" y="224" width="26" height="86" rx="5" fill="${accent}"/>
  <rect x="532" y="246" width="26" height="64" rx="5" fill="#38BDF8" fill-opacity="0.85"/>
  <rect x="52" y="340" width="64" height="6" rx="3" fill="#FFFFFF" fill-opacity="0.22"/>
  <circle cx="566" cy="343" r="4" fill="${accent}"/>
  <circle cx="580" cy="343" r="4" fill="#FFFFFF" fill-opacity="0.25"/>
</svg>`;
  } else if (type === 'docx') {
    svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 400" width="640" height="400">
  <rect width="640" height="400" fill="#EEF3FA"/>
  <rect x="64" y="22" width="512" height="356" rx="16" fill="#FFFFFF" stroke="#DCE5F2" stroke-width="1.5"/>
  <rect x="64" y="22" width="512" height="72" rx="16" fill="${accent}" fill-opacity="0.09"/>
  <rect x="64" y="78" width="512" height="16" fill="${accent}" fill-opacity="0.09"/>
  <rect x="96" y="42" width="56" height="22" rx="6" fill="${accent}"/>
  <text x="124" y="57" text-anchor="middle" font-family="${fontStack}" font-size="10.5" font-weight="700" fill="#FFFFFF">DOCX</text>
  <text x="166" y="59" font-family="${fontStack}" font-size="20" font-weight="700" fill="#0F172A">${safeTitle}</text>
  <rect x="96" y="114" width="144" height="11" rx="5.5" fill="${accent}" fill-opacity="0.78"/>
  <rect x="96" y="138" width="448" height="8" rx="4" fill="#CBD5E1"/>
  <rect x="96" y="156" width="424" height="8" rx="4" fill="#CBD5E1"/>
  <rect x="96" y="174" width="356" height="8" rx="4" fill="#E2E8F0"/>
  <rect x="96" y="202" width="448" height="72" rx="12" fill="#F8FAFC" stroke="#E2E8F0"/>
  <rect x="96" y="202" width="5" height="72" rx="2.5" fill="${accent}"/>
  <rect x="118" y="220" width="160" height="9" rx="4.5" fill="#334155"/>
  <rect x="118" y="239" width="388" height="7" rx="3.5" fill="#94A3B8"/>
  <rect x="118" y="254" width="318" height="7" rx="3.5" fill="#CBD5E1"/>
  <rect x="96" y="294" width="212" height="58" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <rect x="114" y="310" width="94" height="8" rx="4" fill="${accent}" fill-opacity="0.65"/>
  <rect x="114" y="328" width="166" height="7" rx="3.5" fill="#CBD5E1"/>
  <rect x="332" y="294" width="212" height="58" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <rect x="350" y="310" width="108" height="8" rx="4" fill="${accent}" fill-opacity="0.65"/>
  <rect x="350" y="328" width="164" height="7" rx="3.5" fill="#CBD5E1"/>
</svg>`;
  } else if (type === 'xlsx') {
    svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 400" width="640" height="400">
  <rect width="640" height="400" fill="#EDF5F0"/>
  <rect x="32" y="22" width="576" height="356" rx="16" fill="#FFFFFF" stroke="#D3E4D9" stroke-width="1.5"/>
  <rect x="32" y="22" width="576" height="58" rx="16" fill="${accent}"/>
  <rect x="32" y="62" width="576" height="18" fill="${accent}"/>
  <rect x="54" y="40" width="54" height="22" rx="6" fill="#FFFFFF" fill-opacity="0.22"/>
  <text x="81" y="55" text-anchor="middle" font-family="${fontStack}" font-size="10.5" font-weight="700" fill="#FFFFFF">XLSX</text>
  <text x="122" y="57" font-family="${fontStack}" font-size="18.5" font-weight="700" fill="#FFFFFF">${safeTitle}</text>
  <rect x="32" y="80" width="576" height="30" fill="#F7FAF8"/>
  <text x="54" y="99" font-family="${fontStack}" font-size="11" font-weight="700" fill="${accent}">fx</text>
  <rect x="80" y="88" width="248" height="14" rx="4" fill="#FFFFFF" stroke="#DCE8E0"/>
  <text x="90" y="99" font-family="${fontStack}" font-size="10" font-weight="600" fill="#475569">=SUM(D2:D9)</text>
  <rect x="466" y="88" width="120" height="14" rx="4" fill="${accent}" fill-opacity="0.12"/>
  <text x="526" y="99" text-anchor="middle" font-family="${fontStack}" font-size="10" font-weight="700" fill="${accent}">SUM 2.250 tr</text>
  <g font-family="${fontStack}" font-size="10.5" font-weight="700" fill="#475569">
    <rect x="32" y="110" width="576" height="28" fill="#EFF5F1"/>
    <text x="50" y="128" text-anchor="middle" fill="#64748B">#</text>
    <text x="84" y="128">A · Hạng mục</text>
    <text x="278" y="128">B · Ngân sách</text>
    <text x="404" y="128">C · Thực chi</text>
    <text x="514" y="128">D · Tiến độ</text>
  </g>
  <g stroke="#E2ECE5" stroke-width="1">
    <line x1="32" y1="110" x2="608" y2="110"/>
    <line x1="32" y1="138" x2="608" y2="138"/>
    <line x1="32" y1="174" x2="608" y2="174"/>
    <line x1="32" y1="210" x2="608" y2="210"/>
    <line x1="32" y1="246" x2="608" y2="246"/>
    <line x1="32" y1="282" x2="608" y2="282"/>
    <line x1="32" y1="318" x2="608" y2="318"/>
    <line x1="68" y1="110" x2="68" y2="338"/>
    <line x1="262" y1="110" x2="262" y2="338"/>
    <line x1="388" y1="110" x2="388" y2="338"/>
    <line x1="498" y1="110" x2="498" y2="338"/>
  </g>
  <rect x="32" y="174" width="576" height="36" fill="#F8FBF9"/>
  <rect x="32" y="246" width="576" height="36" fill="#F8FBF9"/>
  <rect x="262" y="174" width="126" height="72" fill="${accent}" fill-opacity="0.1" stroke="${accent}" stroke-width="2"/>
  <g font-family="${fontStack}" font-size="10.5" fill="#64748B">
    <text x="50" y="160" text-anchor="middle">1</text>
    <text x="50" y="196" text-anchor="middle">2</text>
    <text x="50" y="232" text-anchor="middle">3</text>
    <text x="50" y="268" text-anchor="middle">4</text>
    <text x="50" y="304" text-anchor="middle">5</text>
  </g>
  <g>
    <rect x="84" y="151" width="142" height="9" rx="4.5" fill="#334155" fill-opacity="0.78"/>
    <rect x="84" y="187" width="124" height="9" rx="4.5" fill="#334155" fill-opacity="0.78"/>
    <rect x="84" y="223" width="152" height="9" rx="4.5" fill="#334155" fill-opacity="0.78"/>
    <rect x="84" y="259" width="132" height="9" rx="4.5" fill="#334155" fill-opacity="0.78"/>
    <rect x="84" y="295" width="116" height="9" rx="4.5" fill="#334155" fill-opacity="0.78"/>
  </g>
  <g font-family="${fontStack}" font-size="11" font-weight="600" fill="#0F172A" text-anchor="end">
    <text x="372" y="160">420,0</text>
    <text x="372" y="196" fill="${accent}">280,0</text>
    <text x="372" y="232" fill="${accent}">360,0</text>
    <text x="372" y="268">240,0</text>
    <text x="372" y="304">310,0</text>
    <text x="482" y="160" fill="#475569">368,5</text>
    <text x="482" y="196" fill="#475569">265,0</text>
    <text x="482" y="232" fill="#475569">294,2</text>
    <text x="482" y="268" fill="#475569">228,8</text>
    <text x="482" y="304" fill="#475569">217,0</text>
  </g>
  <g>
    <rect x="514" y="151" width="74" height="8" rx="4" fill="#E2ECE5"/>
    <rect x="514" y="151" width="65" height="8" rx="4" fill="${accent}"/>
    <rect x="514" y="187" width="74" height="8" rx="4" fill="#E2ECE5"/>
    <rect x="514" y="187" width="74" height="8" rx="4" fill="${accent}"/>
    <rect x="514" y="223" width="74" height="8" rx="4" fill="#E2ECE5"/>
    <rect x="514" y="223" width="60" height="8" rx="4" fill="${accent}" fill-opacity="0.82"/>
    <rect x="514" y="259" width="74" height="8" rx="4" fill="#E2ECE5"/>
    <rect x="514" y="259" width="74" height="8" rx="4" fill="${accent}"/>
    <rect x="514" y="295" width="74" height="8" rx="4" fill="#E2ECE5"/>
    <rect x="514" y="295" width="52" height="8" rx="4" fill="${accent}" fill-opacity="0.72"/>
  </g>
  <rect x="32" y="338" width="576" height="40" rx="16" fill="#F2F7F4"/>
  <rect x="32" y="338" width="576" height="16" fill="#F2F7F4"/>
  <rect x="48" y="346" width="112" height="24" rx="6" fill="#FFFFFF" stroke="${accent}" stroke-width="1.2"/>
  <text x="104" y="362" text-anchor="middle" font-family="${fontStack}" font-size="10.5" font-weight="700" fill="${accent}">Sheet 1</text>
  <rect x="168" y="346" width="102" height="24" rx="6" fill="#E6EFE9"/>
  <text x="219" y="362" text-anchor="middle" font-family="${fontStack}" font-size="10.5" font-weight="600" fill="#475569">Sheet 2</text>
  <text x="588" y="362" text-anchor="end" font-family="${fontStack}" font-size="10.5" font-weight="600" fill="#475569">AVG 81,4%</text>
</svg>`;
  } else {
    svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 400" width="640" height="400">
  <rect width="640" height="400" fill="#F8EFF3"/>
  <rect x="52" y="22" width="536" height="356" rx="16" fill="#FFFFFF" stroke="#EED9E3" stroke-width="1.5"/>
  <rect x="52" y="22" width="536" height="10" rx="5" fill="${accent}"/>
  <rect x="82" y="52" width="48" height="22" rx="6" fill="${accent}" fill-opacity="0.14"/>
  <text x="106" y="67" text-anchor="middle" font-family="${fontStack}" font-size="10.5" font-weight="700" fill="${accent}">PDF</text>
  <text x="142" y="69" font-family="${fontStack}" font-size="20" font-weight="700" fill="#1E293B">${safeTitle}</text>
  <rect x="82" y="94" width="476" height="1" fill="#F1E4EA"/>
  <rect x="82" y="114" width="286" height="154" rx="12" fill="#FDF8FA" stroke="#F3DCE7"/>
  <text x="102" y="138" font-family="${fontStack}" font-size="11" font-weight="600" fill="#64748B">PHÂN BỔ CHỦ ĐỀ CỐT LÕI</text>
  <path d="M104 236 L154 202 L204 214 L254 166 L304 178 L346 146 L346 246 L104 246 Z" fill="${accent}" fill-opacity="0.14"/>
  <path d="M104 236 L154 202 L204 214 L254 166 L304 178 L346 146" fill="none" stroke="${accent}" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="154" cy="202" r="4.5" fill="${accent}"/>
  <circle cx="254" cy="166" r="4.5" fill="${accent}"/>
  <circle cx="346" cy="146" r="4.5" fill="${accent}"/>
  <rect x="386" y="114" width="172" height="70" rx="12" fill="#FAF8F5" stroke="#ECE6DC"/>
  <text x="404" y="140" font-family="${fontStack}" font-size="10.5" font-weight="600" fill="#64748B">CHƯƠNG MỤC</text>
  <text x="404" y="168" font-family="${fontStack}" font-size="22" font-weight="700" fill="#0F172A">04 phần</text>
  <rect x="386" y="198" width="172" height="70" rx="12" fill="#FAF8F5" stroke="#ECE6DC"/>
  <text x="404" y="224" font-family="${fontStack}" font-size="10.5" font-weight="600" fill="#64748B">ĐỘ PHỦ TÓM TẮT</text>
  <rect x="404" y="242" width="136" height="8" rx="4" fill="#E2E8F0"/>
  <rect x="404" y="242" width="108" height="8" rx="4" fill="${accent}"/>
  <rect x="82" y="290" width="476" height="9" rx="4.5" fill="#CBD5E1"/>
  <rect x="82" y="310" width="438" height="9" rx="4.5" fill="#CBD5E1"/>
  <rect x="82" y="330" width="332" height="9" rx="4.5" fill="#E2E8F0"/>
</svg>`;
  }

  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
}

export const LEGACY_MOCK_DOCUMENT_IDS: readonly string[] = [
  'doc-baocao-2026',
  'doc-thietke-sanpham',
  'doc-budget-2026',
  'doc-tonghop-kienthuc',
  'doc-kehoach-nhom',
] as const;

export const mockDocuments: DocumentItem[] = [];

export type EditableDocumentType = 'docx' | 'xlsx' | 'pptx';
export type NewDocumentTemplateMode = 'blank' | 'starter';

export interface CreateEditableDocumentParams {
  type: EditableDocumentType;
  title?: string;
  templateMode?: NewDocumentTemplateMode;
  folder?: string;
}

const DEFAULT_NEW_TITLES: Record<EditableDocumentType, string> = {
  docx: 'Tài liệu Word mới',
  xlsx: 'Bảng tính Excel mới',
  pptx: 'Bài thuyết trình mới',
};

export function createNewEditableDocument(
  params: CreateEditableDocumentParams
): DocumentItem {
  const { type } = params;
  const templateMode: NewDocumentTemplateMode = params.templateMode ?? 'blank';
  const rawTitle = (params.title ?? '').trim().replace(/\.(docx|xlsx|pptx)$/i, '').trim();
  const cleanTitle = rawTitle || DEFAULT_NEW_TITLES[type];
  const safeHtmlTitle = escapeXml(cleanTitle);
  const accent = DEFAULT_ACCENTS[type];
  const id = `${type}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  const folder = params.folder?.trim() || 'Tài liệu của tôi';
  const nowIso = new Date().toISOString();

  const baseDoc: Omit<
    DocumentItem,
    'slidesPreview' | 'outline' | 'contentPreview' | 'pageCount'
  > = {
    id,
    title: cleanTitle,
    type,
    updatedAt: 'Vừa tạo',
    lastSavedAt: nowIso,
    lastAccessedAt: nowIso,
    isFavorite: false,
    isShared: false,
    isUploaded: true,
    folder,
    fileSize: '12 KB',
    originalFileName: `${cleanTitle}.${type}`,
    thumbnail: createThumbnailDataUri(type, cleanTitle, accent),
  };

  if (type === 'docx') {
    if (templateMode === 'starter') {
      const wordHtml = [
        `<h1 id="heading-1">${safeHtmlTitle}</h1>`,
        '<h2 id="sec-1">I. Mục tiêu &amp; Tổng quan</h2>',
        '<p>Tài liệu này tóm tắt mục tiêu trọng tâm, phạm vi công việc và các nguyên tắc phối hợp cốt lõi. Bạn có thể nhấp trực tiếp vào bất kỳ đoạn văn bản nào trên trang A4 để chỉnh sửa, định dạng hoặc chèn thêm nội dung.</p>',
        '<h2 id="sec-2">II. Nội dung chi tiết</h2>',
        '<p>Bảng tổng hợp các hạng mục ưu tiên và trạng thái thực hiện hiện tại:</p>',
        '<table><thead><tr><th>Hạng mục công việc</th><th>Đầu mối phụ trách</th><th>Mốc hoàn thành</th></tr></thead><tbody><tr><td>Khởi tạo khung tài liệu &amp; tiêu chuẩn định dạng</td><td>Nhóm Sản phẩm</td><td>Tuần 1</td></tr><tr><td>Hoàn thiện nội dung chuyên môn &amp; số liệu</td><td>Nhóm Kỹ thuật</td><td>Tuần 2</td></tr><tr><td>Rà soát nghiệm thu &amp; xuất bản chính thức</td><td>Ban Điều phối</td><td>Tuần 3</td></tr></tbody></table>',
        '<h2 id="sec-3">III. Kế hoạch tiếp theo</h2>',
        '<ul><li>Cập nhật số liệu thực tế vào bảng theo dõi hàng tuần.</li><li>Thống nhất các điểm nhấn quan trọng trước buổi họp tổng kết.</li><li>Lưu thay đổi hoặc xuất tệp .docx để chia sẻ cho các bên liên quan.</li></ul>',
      ].join('');

      const outline: OutlineSection[] = [
        {
          id: 'sec-1',
          romanIndex: 'I',
          title: 'I. Mục tiêu & Tổng quan',
          summary:
            'Tóm tắt mục tiêu trọng tâm, phạm vi công việc và nguyên tắc phối hợp cốt lõi.',
          keyPoints: [
            'Xác định mục tiêu & phạm vi công việc',
            'Soạn thảo trực tiếp trên trang A4',
          ],
          pageNumber: 1,
        },
        {
          id: 'sec-2',
          romanIndex: 'II',
          title: 'II. Nội dung chi tiết',
          summary:
            'Bảng tổng hợp các hạng mục ưu tiên, đầu mối phụ trách và mốc hoàn thành.',
          keyPoints: [
            'Khởi tạo khung tài liệu & tiêu chuẩn định dạng',
            'Hoàn thiện nội dung chuyên môn & số liệu',
            'Rà soát nghiệm thu & xuất bản chính thức',
          ],
          pageNumber: 1,
        },
        {
          id: 'sec-3',
          romanIndex: 'III',
          title: 'III. Kế hoạch tiếp theo',
          summary: 'Danh sách các đầu việc cần triển khai và mốc theo dõi định kỳ.',
          keyPoints: [
            'Cập nhật số liệu thực tế hàng tuần',
            'Thống nhất điểm nhấn trước buổi họp tổng kết',
            'Lưu thay đổi hoặc xuất tệp .docx',
          ],
          pageNumber: 1,
        },
      ];

      return {
        ...baseDoc,
        pageCount: 1,
        wordHtml,
        outline,
        slidesPreview: [],
        contentPreview:
          'Tài liệu khởi tạo theo bố cục gợi ý gồm 3 phần: Mục tiêu & Tổng quan, Nội dung chi tiết kèm bảng 3 cột và Kế hoạch tiếp theo.',
      };
    }

    const blankHtml = `<h1 id="heading-1">${safeHtmlTitle}</h1><p>Bắt đầu soạn thảo văn bản của bạn tại đây...</p>`;
    return {
      ...baseDoc,
      pageCount: 1,
      wordHtml: blankHtml,
      outline: [
        {
          id: 'heading-1',
          romanIndex: 'I',
          title: cleanTitle,
          summary: 'Bắt đầu soạn thảo văn bản của bạn tại đây...',
          keyPoints: ['Soạn thảo trực tiếp trên trang giấy A4'],
          pageNumber: 1,
        },
      ],
      slidesPreview: [],
      contentPreview: 'Bắt đầu soạn thảo văn bản của bạn tại đây...',
    };
  }

  if (type === 'xlsx') {
    const columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];

    if (templateMode === 'starter') {
      const starterRows: (string | number)[][] = [
        ['Hạng mục', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ghi chú', '', '', ''],
        ['Thiết bị trình chiếu phòng họp', 4, 12500000, '=B2*C2', 'Đã bàn giao', '', '', ''],
        ['Bản quyền phần mềm cộng tác', 25, 1800000, '=B3*C3', 'Gia hạn năm', '', '', ''],
        ['Chi phí hạ tầng đám mây & lưu trữ', 12, 3500000, '=B4*C4', 'Theo tháng', '', '', ''],
        ['Tổng cộng', '=SUM(B2:B4)', '', '=SUM(D2:D4)', 'Tự động tính tổng', '', '', ''],
      ];
      for (let r = 0; r < 15; r += 1) {
        starterRows.push(['', '', '', '', '', '', '', '']);
      }

      const cellStyles: Record<string, CellStyle> = {
        '0:0': { bold: true, align: 'left' },
        '0:1': { bold: true, align: 'right' },
        '0:2': { bold: true, align: 'right' },
        '0:3': { bold: true, align: 'right' },
        '0:4': { bold: true, align: 'left' },
        '4:0': { bold: true, align: 'left' },
        '4:1': { bold: true, align: 'right' },
        '4:3': { bold: true, align: 'right' },
      };

      const sheets: SpreadsheetSheet[] = [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          columns,
          rows: starterRows,
          cellStyles,
        },
      ];

      return {
        ...baseDoc,
        pageCount: 1,
        sheets,
        outline: [],
        slidesPreview: [],
        contentPreview:
          'Bảng tính mẫu có sẵn công thức Thành tiền (=B2*C2) và dòng Tổng cộng (=SUM(D2:D4)) để chỉnh sửa số liệu ngay.',
      };
    }

    const blankRows: (string | number)[][] = Array.from({ length: 20 }, () =>
      Array.from({ length: 8 }, () => '')
    );

    return {
      ...baseDoc,
      pageCount: 1,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          columns,
          rows: blankRows,
          cellStyles: {},
        },
      ],
      outline: [],
      slidesPreview: [],
      contentPreview:
        'Bảng tính Excel trắng gồm 20 dòng × 8 cột, sẵn sàng nhập số liệu và công thức.',
    };
  }

  // type === 'pptx'
  if (templateMode === 'starter') {
    const slidesPreview: SlidePreviewItem[] = [
      {
        id: `${id}-slide-1`,
        index: 1,
        title: cleanTitle,
        subtitle: 'Bối cảnh, mục tiêu cốt lõi & định hướng triển khai',
        bullets: [
          'Tóm tắt thông điệp quan trọng nhất trong 1 câu ngắn gọn',
          'Người trình bày • Ngày cập nhật mới nhất',
        ],
        accentColor: '#4F46E5',
        speakerNote:
          'Giới thiệu ngắn gọn mục tiêu buổi trình bày và cấu trúc 3 phần chính.',
        blocks: [
          {
            id: `${id}-s1-title`,
            text: cleanTitle,
            fontSize: 32,
            bold: true,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'title',
          },
          {
            id: `${id}-s1-subtitle`,
            text: 'Bối cảnh, mục tiêu cốt lõi & định hướng triển khai',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#CBD5E1',
            align: 'left',
            role: 'subtitle',
          },
          {
            id: `${id}-s1-body-1`,
            text: '• Tóm tắt thông điệp quan trọng nhất trong 1 câu ngắn gọn\n• Người trình bày • Ngày cập nhật mới nhất',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'body',
          },
        ],
      },
      {
        id: `${id}-slide-2`,
        index: 2,
        title: 'Nội dung trọng tâm & Giải pháp',
        subtitle: 'Ba trụ cột chính cần tập trung nguồn lực',
        bullets: [
          'Trụ cột 1: Chuẩn hóa quy trình và dữ liệu đầu vào',
          'Trụ cột 2: Tối ưu trải nghiệm cộng tác và tốc độ xử lý',
          'Trụ cột 3: Đo lường kết quả thực tế qua các chỉ số rõ ràng',
        ],
        metricHighlight: {
          label: 'Mục tiêu hiệu suất',
          value: '+35%',
        },
        accentColor: '#0EA5E9',
        speakerNote:
          'Đi sâu vào 3 luận điểm chính và dẫn chứng số liệu minh họa.',
        blocks: [
          {
            id: `${id}-s2-title`,
            text: 'Nội dung trọng tâm & Giải pháp',
            fontSize: 32,
            bold: true,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'title',
          },
          {
            id: `${id}-s2-subtitle`,
            text: 'Ba trụ cột chính cần tập trung nguồn lực',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#CBD5E1',
            align: 'left',
            role: 'subtitle',
          },
          {
            id: `${id}-s2-body-1`,
            text: '• Trụ cột 1: Chuẩn hóa quy trình và dữ liệu đầu vào\n• Trụ cột 2: Tối ưu trải nghiệm cộng tác và tốc độ xử lý\n• Trụ cột 3: Đo lường kết quả thực tế qua các chỉ số rõ ràng',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'body',
          },
        ],
      },
      {
        id: `${id}-slide-3`,
        index: 3,
        title: 'Tổng kết & Kế hoạch hành động',
        subtitle: 'Các bước tiếp theo và mốc thời gian cam kết',
        bullets: [
          'Chốt phạm vi ưu tiên và phân công đầu mối phụ trách',
          'Triển khai thử nghiệm và thu thập phản hồi thực tế',
          'Đánh giá nghiệm thu và mở rộng quy mô vận hành',
        ],
        metricHighlight: {
          label: 'Mốc hoàn thành',
          value: '100%',
        },
        accentColor: '#10B981',
        speakerNote:
          'Tổng kết lại các quyết định cần thông qua và mở phần thảo luận.',
        blocks: [
          {
            id: `${id}-s3-title`,
            text: 'Tổng kết & Kế hoạch hành động',
            fontSize: 32,
            bold: true,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'title',
          },
          {
            id: `${id}-s3-subtitle`,
            text: 'Các bước tiếp theo và mốc thời gian cam kết',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#CBD5E1',
            align: 'left',
            role: 'subtitle',
          },
          {
            id: `${id}-s3-body-1`,
            text: '• Chốt phạm vi ưu tiên và phân công đầu mối phụ trách\n• Triển khai thử nghiệm và thu thập phản hồi thực tế\n• Đánh giá nghiệm thu và mở rộng quy mô vận hành',
            fontSize: 18,
            bold: false,
            italic: false,
            color: '#FFFFFF',
            align: 'left',
            role: 'body',
          },
        ],
      },
    ];

    const outline: OutlineSection[] = slidesPreview.map((slide, idx) => ({
      id: `${id}-outline-${idx + 1}`,
      romanIndex: ['I', 'II', 'III'][idx] ?? String(idx + 1),
      title: slide.title,
      summary: slide.subtitle ?? slide.title,
      keyPoints: slide.bullets,
      pageNumber: idx + 1,
    }));

    return {
      ...baseDoc,
      pageCount: 3,
      slidesPreview,
      outline,
      contentPreview:
        'Bộ bài thuyết trình khởi tạo gồm 3 slide 16:9: Trang bìa, Nội dung trọng tâm và Tổng kết & Hành động.',
    };
  }

  const blankSlide: SlidePreviewItem = {
    id: `${id}-slide-1`,
    index: 1,
    title: cleanTitle,
    subtitle: 'Nhấp đúp hoặc chọn khung chữ để chỉnh sửa nội dung',
    bullets: ['Thêm điểm nhấn chính của bạn tại đây'],
    accentColor: '#4F46E5',
    speakerNote: 'Ghi chú diễn giả cho slide mở đầu.',
    blocks: [
      {
        id: `${id}-s1-title`,
        text: cleanTitle,
        fontSize: 32,
        bold: true,
        italic: false,
        color: '#FFFFFF',
        align: 'left',
        role: 'title',
      },
      {
        id: `${id}-s1-subtitle`,
        text: 'Nhấp đúp hoặc chọn khung chữ để chỉnh sửa nội dung',
        fontSize: 18,
        bold: false,
        italic: false,
        color: '#CBD5E1',
        align: 'left',
        role: 'subtitle',
      },
      {
        id: `${id}-s1-body-1`,
        text: '• Thêm điểm nhấn chính của bạn tại đây',
        fontSize: 18,
        bold: false,
        italic: false,
        color: '#FFFFFF',
        align: 'left',
        role: 'body',
      },
    ],
  };

  return {
    ...baseDoc,
    pageCount: 1,
    slidesPreview: [blankSlide],
    outline: [
      {
        id: `${id}-outline-1`,
        romanIndex: 'I',
        title: cleanTitle,
        summary: 'Nhấp đúp hoặc chọn khung chữ để chỉnh sửa nội dung',
        keyPoints: ['Thêm điểm nhấn chính của bạn tại đây'],
        pageNumber: 1,
      },
    ],
    contentPreview:
      'Bài thuyết trình 16:9 mới — chọn khung chữ để soạn thảo hoặc thêm slide mới.',
  };
}

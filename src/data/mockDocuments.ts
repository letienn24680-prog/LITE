import type { DocumentItem, DocumentType } from '../types/document';

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

export const mockDocuments: DocumentItem[] = [
  {
    id: 'doc-baocao-2026',
    title: 'Báo cáo dự án 2026',
    type: 'pptx',
    updatedAt: 'Hôm nay, 09:40',
    isFavorite: true,
    isShared: true,
    isUploaded: false,
    folder: 'Dự án 2026',
    pageCount: 18,
    fileSize: '4.2 MB',
    thumbnail: createThumbnailDataUri('pptx', 'Báo cáo dự án 2026', '#4F46E5'),
    slidesPreview: [
      {
        id: 'bc-slide-1',
        index: 1,
        title: 'Tổng quan chiến lược Dự án 2026',
        subtitle: 'Chuyển đổi quy trình vận hành & chuẩn hóa dữ liệu tài liệu',
        bullets: [
          'Hợp nhất 3 kho lưu trữ phân tán về một không gian làm việc duy nhất',
          'Chuẩn hóa trải nghiệm trình chiếu và đọc tài liệu trực tiếp trên trình duyệt',
          'Đặt mục tiêu giảm 40% thời gian chuẩn bị báo cáo định kỳ toàn khối',
        ],
        metricHighlight: {
          label: 'Tốc độ hoàn thành mục tiêu',
          value: '+38%',
        },
        accentColor: '#4F46E5',
        speakerNote:
          'Nhấn mạnh bối cảnh năm 2026 khi các nhóm làm việc kết hợp cần truy cập tài liệu tức thì không phụ thuộc thiết bị.',
      },
      {
        id: 'bc-slide-2',
        index: 2,
        title: 'Thực trạng hạ tầng & điểm nghẽn hiện tại',
        subtitle: 'Khảo sát nội bộ trên 6 phòng ban trong 2 quý gần nhất',
        bullets: [
          'Tài liệu PPTX và DOCX thường lỗi phông chữ khi mở chéo giữa Windows và macOS',
          'Chi phí băng thông truyền tải file gốc dung lượng lớn tăng 24% mỗi quý',
          'Thiếu cơ chế đồng bộ ghi chú và mục lục nhanh cho người đọc trên di động',
        ],
        metricHighlight: {
          label: 'Độ ổn định dịch vụ',
          value: '99.94%',
        },
        accentColor: '#6366F1',
        speakerNote:
          'Chỉ rõ 3 điểm đau lớn nhất dẫn đến quyết định xây dựng tầng hiển thị chuẩn hóa.',
      },
      {
        id: 'bc-slide-3',
        index: 3,
        title: 'Kiến trúc giải pháp tối ưu chi phí',
        subtitle: 'Xử lý tiền kết xuất nhẹ kết hợp lưu trữ đối tượng phân tầng',
        bullets: [
          'Tự động sinh bản xem trước nhẹ giúp tải trang đầu dưới 600ms',
          'Tích hợp Smart Reader tách mục lục I–IV giúp tra cứu nhanh không cần cuộn dài',
          'Điều khiển trình chiếu qua mã QR hỗ trợ diễn giả linh hoạt tại phòng họp',
        ],
        metricHighlight: {
          label: 'Tiết kiệm băng thông',
          value: '-42%',
        },
        accentColor: '#0EA5E9',
        speakerNote:
          'Dẫn chứng mô hình triển khai gọn nhẹ giúp dễ mở rộng khi số lượng tài liệu tăng gấp 5 lần.',
      },
      {
        id: 'bc-slide-4',
        index: 4,
        title: 'Lộ trình triển khai & cam kết đầu ra',
        subtitle: 'Kế hoạch 4 giai đoạn từ thí điểm đến vận hành toàn diện',
        bullets: [
          'Quý I–II: Hoàn thiện lõi lưu trữ, trình chiếu mượt và đọc thông minh',
          'Quý III: Mở rộng chia sẻ liên nhóm kèm phân quyền theo thư mục',
          'Quý IV: Đo lường hiệu quả sử dụng thực tế và tối ưu tự động hóa báo cáo',
        ],
        metricHighlight: {
          label: 'Người dùng mục tiêu',
          value: '120K',
        },
        accentColor: '#10B981',
        speakerNote:
          'Chốt các mốc nghiệm thu rõ ràng cho ban điều hành trước khi chuyển sang phần hỏi đáp.',
      },
    ],
    outline: [
      {
        id: 'bc-outline-1',
        romanIndex: 'I',
        title: 'Đặt vấn đề',
        summary:
          'Xác định nhu cầu cấp thiết về một không gian lưu trữ và trình chiếu tài liệu đồng nhất, loại bỏ sự cố vỡ bố cục khi cộng tác đa nền tảng.',
        keyPoints: [
          'Tập trung hóa tài liệu PPTX, DOCX, XLSX, PDF từ nhiều nguồn rải rác',
          'Đảm bảo mở xem và trình chiếu tức thì không cần cài phần mềm văn phòng',
          'Tối ưu trải nghiệm cho cả buổi họp trực tiếp lẫn làm việc từ xa',
        ],
        pageNumber: 1,
      },
      {
        id: 'bc-outline-2',
        romanIndex: 'II',
        title: 'Thực trạng',
        summary:
          'Phân tích số liệu vận hành 6 tháng qua cho thấy thời gian tìm kiếm, chuyển đổi định dạng và xử lý lỗi hiển thị chiếm tỷ trọng đáng kể.',
        keyPoints: [
          '31% phiên họp bị gián đoạn 3–5 phút do lỗi font hoặc lệch khung slide',
          'Dung lượng tệp đính kèm trùng lặp làm tăng chi phí lưu trữ ngoài dự kiến',
          'Người đọc gặp khó khăn khi theo dõi tài liệu dài trên màn hình nhỏ',
        ],
        pageNumber: 5,
      },
      {
        id: 'bc-outline-3',
        romanIndex: 'III',
        title: 'Giải pháp',
        summary:
          'Đề xuất kiến trúc 2 tầng tách bạch giữa Trang chủ quản lý toàn màn hình và 4 không gian làm việc chuyên biệt theo từng định dạng tệp.',
        keyPoints: [
          'Tách biệt không gian trình chiếu Slide, đọc A4 Word, bảng tính Excel và xem trang PDF',
          'Hỗ trợ chế độ đọc giấy kem ấm bảo vệ mắt và chế độ tối linh hoạt',
          'Cung cấp bộ điều khiển QR, bút Laser cùng đồng hồ canh giờ cho người thuyết trình',
        ],
        pageNumber: 10,
      },
      {
        id: 'bc-outline-4',
        romanIndex: 'IV',
        title: 'Kết luận',
        summary:
          'Tổng hợp lợi ích định lượng, lộ trình triển khai theo từng quý và bộ chỉ số đánh giá mức độ thành công của dự án trong năm 2026.',
        keyPoints: [
          'Giảm 42% chi phí băng thông truyền tải và rút ngắn thời gian chuẩn bị họp',
          'Khởi chạy bản lõi trong Quý II để thu thập phản hồi thực tế từ nhóm tiên phong',
          'Sẵn sàng mở rộng phân quyền chia sẻ và tương tác trực tiếp ở giai đoạn tiếp theo',
        ],
        pageNumber: 15,
      },
    ],
    contentPreview:
      'Báo cáo tổng hợp tiến độ và định hướng Dự án 2026 tập trung vào việc chuẩn hóa toàn bộ vòng đời tài liệu từ khâu tải lên, sắp xếp theo thư mục đến trình chiếu và đọc thông minh trực tuyến. Nhờ áp dụng cấu trúc hiển thị nhất quán, nhóm vận hành rút ngắn đáng kể thời gian chuẩn bị trước mỗi phiên họp đồng thời kiểm soát hiệu quả chi phí hạ tầng.',
  },
  {
    id: 'doc-thietke-sanpham',
    title: 'Thiết kế sản phẩm',
    type: 'docx',
    updatedAt: 'Hôm qua, 16:25',
    isFavorite: true,
    isShared: true,
    isUploaded: false,
    folder: 'Sản phẩm',
    pageCount: 12,
    fileSize: '1.8 MB',
    thumbnail: createThumbnailDataUri('docx', 'Thiết kế sản phẩm', '#2563EB'),
    slidesPreview: [
      {
        id: 'tk-slide-1',
        index: 1,
        title: 'Ngôn ngữ thiết kế tối giản LITE',
        subtitle: 'Cân bằng giữa không gian giấy ấm áp và sân khấu trình chiếu tương phản',
        bullets: [
          'Bảng màu nền kem ấm #FAF8F5 giúp đọc dài không gây mỏi mắt',
          'Không gian chuyên biệt cho từng định dạng tạo tập trung tuyệt đối vào nội dung',
          'Nhịp khoảng trắng rộng rãi ưu tiên thao tác một chạm rõ ràng',
        ],
        metricHighlight: {
          label: 'Điểm hài lòng UX',
          value: '4.9/5',
        },
        accentColor: '#2563EB',
        speakerNote:
          'Giải thích lý do phân tách sắc độ sáng – tối giữa vùng quản lý thư viện và không gian trình chiếu chuyên sâu.',
      },
      {
        id: 'tk-slide-2',
        index: 2,
        title: 'Kiến trúc điều hướng 2 tầng rõ ràng',
        subtitle: 'Trang chủ tìm tệp cực nhanh, bấm mở vào ngay không gian chuyên dụng',
        bullets: [
          'Trang chủ toàn màn hình hiển thị lưới thẻ 4 cột thoáng đãng, tìm lọc tức thì',
          'Slide, Word, Excel và PDF có bộ công cụ vượt trội riêng biệt không lẫn lộn',
          'Thanh tiêu đề thống nhất 56px giúp quay về Trang chủ và chia sẻ trong 1 chạm',
        ],
        metricHighlight: {
          label: 'Số bước mở tài liệu',
          value: '1 chạm',
        },
        accentColor: '#0284C7',
      },
      {
        id: 'tk-slide-3',
        index: 3,
        title: 'Tiêu chuẩn chữ & khả năng tiếp cận',
        subtitle: 'Tối ưu hiển thị dấu tiếng Việt rõ nét ở mọi kích thước',
        bullets: [
          'Ưu tiên phông chữ Be Vietnam Pro kết hợp Inter cho chỉ số và bảng biểu',
          'Đảm bảo độ tương phản chữ tối thiểu đạt chuẩn WCAG AA trên cả 3 chế độ giấy',
          'Hỗ trợ điều hướng bàn phím mượt mà cho trình chiếu và đóng nhanh bằng Esc',
        ],
        metricHighlight: {
          label: 'Chuẩn tương phản',
          value: 'WCAG AA',
        },
        accentColor: '#4F46E5',
      },
    ],
    outline: [
      {
        id: 'tk-outline-1',
        romanIndex: 'I',
        title: 'Mục tiêu trải nghiệm',
        summary:
          'Định hình nguyên tắc thiết kế lấy sự tập trung làm trung tâm, giúp người dùng tìm thấy và mở tài liệu chỉ trong vài giây.',
        keyPoints: [
          'Loại bỏ các thanh công cụ rườm rà không cần thiết trên màn hình chính',
          'Phân cấp thị giác rõ giữa tiêu đề, thẻ định dạng và nút hành động chính',
          'Giữ tốc độ phản hồi giao diện mượt mà trên cả laptop phổ thông lẫn di động',
        ],
        pageNumber: 1,
      },
      {
        id: 'tk-outline-2',
        romanIndex: 'II',
        title: 'Hệ thống thành phần',
        summary:
          'Quy chuẩn hóa thẻ tài liệu, bộ lọc dạng viên thuốc, thanh kéo thả và nhãn định dạng PPTX, DOCX, XLSX, PDF nhất quán.',
        keyPoints: [
          'Màu nhãn định dạng dễ phân biệt nhanh: cam PPTX, xanh dương DOCX, xanh lá XLSX, hồng PDF',
          'Thanh Header 56px dùng chung cho cả 4 không gian giúp giữ định hướng nhất quán',
          'Tái sử dụng trọn bộ biểu tượng đường nét đồng nhất từ lucide-react',
        ],
        pageNumber: 4,
      },
      {
        id: 'tk-outline-3',
        romanIndex: 'III',
        title: 'Luồng tương tác chính',
        summary:
          'Mô tả chi tiết chuỗi thao tác từ kéo thả tải tệp, lọc tìm kiếm theo thời gian thực đến chuyển đổi sang 4 không gian làm việc chuyên sâu.',
        keyPoints: [
          'Tệp mới tải lên lập tức xuất hiện đầu danh sách và sẵn sàng mở bằng một chạm',
          'Không gian Word hỗ trợ đổi 3 nền giấy, tăng giảm cỡ chữ và bôi đậm ghi chú trực tiếp',
          'Không gian Slide tích hợp bút Laser, điều khiển QR và ghi chú diễn giả liền mạch',
        ],
        pageNumber: 7,
      },
      {
        id: 'tk-outline-4',
        romanIndex: 'IV',
        title: 'Tiêu chuẩn kiểm định',
        summary:
          'Thiết lập danh mục kiểm thử giao diện về độ sắc nét chữ tiếng Việt, trạng thái rỗng và tính nhất quán trên mọi kích thước màn hình.',
        keyPoints: [
          'Kiểm tra không tràn chữ với tiêu đề dài trên cả thẻ thư viện và thanh công cụ',
          'Đảm bảo chuyển chế độ màu giấy giữ nguyên vị trí đọc và ghi chú đã đánh dấu',
          'Xác minh thông báo lỗi lịch sự khi người dùng thả tệp sai định dạng hỗ trợ',
        ],
        pageNumber: 10,
      },
    ],
    wordParagraphs: {
      'tk-outline-1': [
        'Mục tiêu cốt lõi của ngôn ngữ thiết kế LITE là trả lại sự tĩnh tại và tập trung cho người làm việc với tài liệu. Thay vì dàn trải hàng chục biểu tượng định dạng phức tạp như các bộ ứng dụng văn phòng truyền thống, giao diện ưu tiên không gian thở, tương phản sắc nét và các điểm chạm vừa tầm mắt.',
        'Ở tầng Trang chủ, người dùng chỉ cần tập trung vào hai hành động tự nhiên nhất: tìm kiếm hoặc lọc nhanh tài liệu cần dùng, sau đó mở thẳng vào không gian xử lý chuyên biệt. Nhờ tách bạch rõ nhiệm vụ giữa khu vực quản lý và khu vực đọc/trình chiếu, tải nhận thức giảm đi đáng kể ngay cả khi thư viện mở rộng lên hàng trăm tệp.',
        'Mọi chuyển đổi trạng thái như tải tệp, đổi bộ lọc hay chuyển trang tài liệu đều được thiết kế phản hồi tức thì dưới 100 mili-giây, tạo cảm giác nhẹ nhàng và tin cậy trên cả màn hình máy tính lẫn điện thoại.',
      ],
      'tk-outline-2': [
        'Hệ thống nhận diện định dạng sử dụng bốn sắc độ có chủ đích để mắt người nhận diện tức thì mà không cần đọc kỹ phần đuôi tệp: cam ấm cho Slide (PPTX), xanh dương chuẩn mực cho Văn bản (DOCX), xanh lá đậm cho Bảng tính (XLSX) và hồng đỏ nổi bật cho Tài liệu xuất bản (PDF).',
        'Toàn bộ thanh điều hướng đầu trang trong bốn không gian làm việc đều kế thừa chung một khung chiều cao 56px, đặt nút quay về Trang chủ và tên tài liệu ở bên trái, cụm công cụ chuyên môn ở giữa và nút Chia sẻ ở góc phải. Sự nhất quán này giúp người dùng chuyển qua lại giữa các loại tệp mà không phải học lại cách thao tác.',
      ],
      'tk-outline-3': [
        'Đối với văn bản dài dạng DOCX, trải nghiệm đọc được nâng tầm bằng khổ trang A4 căn giữa chuẩn tỷ lệ in ấn kết hợp cột Mục lục tự động bên trái. Người đọc có thể chuyển đổi linh hoạt giữa nền trắng tinh khiết ban ngày, giấy kem ấm dịu mắt buổi chiều và nền tối bảo vệ thị lực vào ban đêm.',
        'Tính năng bôi đậm ghi chú trực tiếp trên đoạn văn biến tài liệu tĩnh thành không gian làm việc sống động. Chỉ cần bật chế độ đánh dấu và chạm vào đoạn văn quan trọng, hệ thống lập tức tô sáng và mở ô ghi chú ngắn ngay tại chỗ, hỗ trợ rà soát hợp đồng, tài liệu đặc tả hay giáo trình cực kỳ thuận tiện.',
        'Trên màn hình di động, cột mục lục tự động thu gọn thành ngăn trượt mở nhanh theo yêu cầu, đồng thời giữ nguyên cỡ chữ và lề đọc thoải mái để người dùng không bao giờ phải phóng to thu nhỏ thủ công.',
      ],
      'tk-outline-4': [
        'Trước mỗi đợt phát hành, toàn bộ thành phần giao diện phải vượt qua bộ tiêu chí kiểm định gồm hiển thị dấu tiếng Việt không cắt ngọn ở mọi cỡ chữ từ 14px đến 20px, vùng chạm tối thiểu 44px trên màn hình cảm ứng và khả năng vận hành trọn vẹn bằng bàn phím.',
        'Ngoài các kịch bản thuận lợi, quy trình kiểm thử đặc biệt chú trọng các trạng thái biên: tệp mới tải lên chưa có mục lục chi tiết, thư mục tìm kiếm không có kết quả khớp, hoặc chuyển đổi kích thước cửa sổ đột ngột giữa desktop và tablet. Mọi tình huống đều phải giữ bố cục ổn định, không vỡ khung và có hướng dẫn thao tác rõ ràng.',
      ],
    },
    contentPreview:
      'Tài liệu đặc tả thiết kế sản phẩm xác lập chuẩn mực kiến trúc 2 tầng, hệ thống màu sắc dịu mắt và ngôn ngữ tương tác tối giản cho nền tảng LITE. Mọi thành phần từ ô tìm kiếm, bộ lọc viên thuốc đến không gian đọc khổ A4 đều được cân chỉnh khoảng trắng và độ tương phản nhằm mang lại cảm giác gọn gàng, tin cậy ngay từ lần sử dụng đầu tiên.',
  },
  {
    id: 'doc-budget-2026',
    title: 'Bảng theo dõi tiến độ & Ngân sách 2026',
    type: 'xlsx',
    updatedAt: 'Hôm nay, 09:40',
    isFavorite: true,
    isShared: true,
    isUploaded: false,
    folder: 'Tài chính & Tiến độ',
    pageCount: 2,
    fileSize: '1.8 MB',
    thumbnail: createThumbnailDataUri(
      'xlsx',
      'Bảng theo dõi tiến độ & Ngân sách 2026',
      '#15803D'
    ),
    slidesPreview: [
      {
        id: 'bg-slide-1',
        index: 1,
        title: 'Tổng quan giải ngân Ngân sách 2026',
        subtitle: 'Theo dõi đồng thời 8 hạng mục trọng điểm và phân bổ theo 4 quý',
        bullets: [
          'Tổng ngân sách phê duyệt 2.250 triệu đồng cho 8 phân hệ sản phẩm cốt lõi',
          'Tỷ lệ giải ngân thực tế đạt 79,9% với mức tiết kiệm trung bình 8,4% so với dự toán',
          'Hai phân hệ Thiết kế giao diện và Smart Reader đã hoàn thành 100% đúng hạn',
        ],
        metricHighlight: {
          label: 'Tiến độ trung bình',
          value: '81,4%',
        },
        accentColor: '#15803D',
        speakerNote:
          'Mở trực tiếp Sheet 1 để đối chiếu chi tiết chênh lệch Ngân sách và Thực chi từng hạng mục.',
      },
      {
        id: 'bg-slide-2',
        index: 2,
        title: 'Hiệu quả kiểm soát chi phí theo hạng mục',
        subtitle: 'Tập trung nguồn lực vào hạ tầng lõi và trải nghiệm người dùng',
        bullets: [
          'Hạ tầng lưu trữ & CDN chiếm 18,7% ngân sách nhưng giúp giảm 42% phí băng thông',
          'Nhóm bảng tính XLSX và trình xem PDF đang duy trì biên chi phí an toàn dưới kế hoạch',
          'Quỹ bảo mật & phân quyền chia sẻ giữ sẵn 54,6% dự toán cho chặng kiểm định Quý III–IV',
        ],
        metricHighlight: {
          label: 'Tổng thực chi',
          value: '1.798,4 tr',
        },
        accentColor: '#16A34A',
      },
      {
        id: 'bg-slide-3',
        index: 3,
        title: 'Cân đối phân bổ nguồn lực theo 4 Quý',
        subtitle: 'Tổng hợp ngân sách 6 bộ phận chủ chốt trên Sheet 2',
        bullets: [
          'Tổng ngân sách vận hành 6 bộ phận cả năm đạt 2.454 triệu đồng',
          'Đỉnh đầu tư tập trung vào Quý II và Quý III phục vụ hoàn thiện 4 không gian chuyên biệt',
          'Duy trì quỹ dự phòng & đào tạo 165 triệu đồng để ứng biến linh hoạt với yêu cầu mới',
        ],
        metricHighlight: {
          label: 'Tổng phân bổ 4 quý',
          value: '2.454 tr',
        },
        accentColor: '#059669',
      },
    ],
    outline: [
      {
        id: 'bg-outline-1',
        romanIndex: 'I',
        title: 'Cấu trúc ngân sách tổng thể',
        summary:
          'Tổng hợp kế hoạch tài chính năm 2026 chia theo 8 hạng mục triển khai và 6 bộ phận chức năng nhằm minh bạch hóa dòng tiền.',
        keyPoints: [
          'Theo dõi song song cột Ngân sách dự kiến, Thực chi và Tiến độ hoàn thành (%)',
          'Chuẩn hóa đơn vị tính triệu đồng (tr) giúp tính tổng SUM và trung bình AVG chính xác',
          'Phân công rõ nhân sự phụ trách cho từng hạng mục công việc',
        ],
        pageNumber: 1,
      },
      {
        id: 'bg-outline-2',
        romanIndex: 'II',
        title: 'Tiến độ 8 hạng mục trọng điểm',
        summary:
          'Đánh giá chi tiết tình trạng triển khai từng mô-đun từ hạ tầng, thiết kế đến bốn không gian làm việc chuyên sâu.',
        keyPoints: [
          '2 hạng mục đã hoàn thành 100%, 1 hạng mục đang kiểm thử đạt 92%',
          '4 hạng mục đang triển khai đúng nhịp độ từ 70% đến 88%',
          'Hạng mục bảo mật & phân quyền chuẩn bị tăng tốc ở chặng tiếp theo',
        ],
        pageNumber: 1,
      },
      {
        id: 'bg-outline-3',
        romanIndex: 'III',
        title: 'Phân bổ chi phí theo 4 Quý',
        summary:
          'Bảng phân kỳ ngân sách trên Sheet 2 giúp ban điều hành kiểm soát dòng tiền theo từng giai đoạn phát triển sản phẩm.',
        keyPoints: [
          'Kỹ thuật Nền tảng và Vận hành & Hạ tầng chiếm hơn 55% tổng ngân sách năm',
          'Tăng dần ngân sách Tăng trưởng & Cộng đồng từ 48 tr Quý I lên 94,5 tr Quý IV',
          'Giữ ổn định chi phí Kiểm thử & Chất lượng xuyên suốt cả bốn quý',
        ],
        pageNumber: 2,
      },
      {
        id: 'bg-outline-4',
        romanIndex: 'IV',
        title: 'Cảnh báo & khuyến nghị điều hành',
        summary:
          'Đưa ra ngưỡng kiểm soát chênh lệch chi phí và khuyến nghị rà soát định kỳ cuối mỗi tháng trên bảng tính chia sẻ chung.',
        keyPoints: [
          'Dùng bộ lọc trạng thái và sắp xếp cột Thực chi để phát hiện nhanh hạng mục chạm trần',
          'Rà soát lại hạng mục Bảo mật trước tuần đầu Quý III để đảm bảo tiến độ bàn giao',
          'Cập nhật số liệu thực chi định kỳ vào thứ Sáu hàng tuần ngay trên kho dữ liệu chung',
        ],
        pageNumber: 2,
      },
    ],
    sheets: [
      {
        id: 'sheet-progress-budget',
        name: 'Sheet 1: Tiến độ & Ngân sách',
        columns: [
          'Hạng mục',
          'Phụ trách',
          'Trạng thái',
          'Ngân sách (tr)',
          'Thực chi (tr)',
          'Tiến độ (%)',
        ],
        rows: [
          ['Hạ tầng lưu trữ & CDN', 'Minh Khôi', 'Đang triển khai', 420, 368.5, 88],
          ['Thiết kế giao diện LITE', 'Thu Hà', 'Hoàn thành', 280, 265, 100],
          ['Lõi trình chiếu Slide & QR', 'Quốc Anh', 'Đang triển khai', 360, 294.2, 82],
          ['Không gian đọc Smart Reader', 'Lan Phương', 'Hoàn thành', 240, 228.8, 100],
          ['Bảng tính & Phân tích XLSX', 'Đức Huy', 'Đang triển khai', 310, 217, 70],
          ['Trình xem & Chú thích PDF', 'Ngọc Mai', 'Kiểm thử', 195, 172.5, 92],
          ['Bảo mật & Phân quyền chia sẻ', 'Hoàng Nam', 'Lên kế hoạch', 260, 118, 45],
          ['Truyền thông & Ra mắt thử nghiệm', 'Bích Ngọc', 'Đang triển khai', 185, 134.4, 74],
        ],
      },
      {
        id: 'sheet-quarterly-allocation',
        name: 'Sheet 2: Phân bổ theo Quý',
        columns: [
          'Bộ phận',
          'Quý 1 (tr)',
          'Quý 2 (tr)',
          'Quý 3 (tr)',
          'Quý 4 (tr)',
          'Tổng năm (tr)',
        ],
        rows: [
          ['Kỹ thuật Nền tảng', 185, 220, 210, 195, 810],
          ['Thiết kế Sản phẩm', 95, 110, 88, 72, 365],
          ['Vận hành & Hạ tầng', 120, 132.5, 140, 147.5, 540],
          ['Kiểm thử & Chất lượng', 62, 74, 80, 68, 284],
          ['Tăng trưởng & Cộng đồng', 48, 65, 82.5, 94.5, 290],
          ['Dự phòng & Đào tạo', 35, 40, 42, 48, 165],
        ],
      },
    ],
    contentPreview:
      'Bảng theo dõi tiến độ & Ngân sách 2026 tổng hợp toàn bộ dự toán, thực chi và phần trăm hoàn thành của 8 hạng mục trọng điểm cùng bảng phân bổ ngân sách theo 4 quý của 6 bộ phận. Người dùng có thể lọc dòng trực tiếp, sắp xếp theo từng cột số liệu và bôi chọn vùng ô để tính nhanh SUM, AVG ngay trên trình duyệt.',
  },
  {
    id: 'doc-tonghop-kienthuc',
    title: 'Tổng hợp kiến thức',
    type: 'pdf',
    updatedAt: '2 ngày trước',
    isFavorite: false,
    isShared: false,
    isUploaded: false,
    folder: 'Tài liệu tham khảo',
    pageCount: 6,
    fileSize: '6.5 MB',
    thumbnail: createThumbnailDataUri('pdf', 'Tổng hợp kiến thức', '#DB2777'),
    slidesPreview: [
      {
        id: 'kt-slide-1',
        index: 1,
        title: 'Bản đồ kiến thức hệ thống tài liệu số',
        subtitle: 'Cấu trúc hóa 4 trụ cột từ lưu trữ, trích xuất đến hiển thị tối ưu',
        bullets: [
          'Phân biệt định dạng tĩnh PDF với định dạng có cấu trúc PPTX, DOCX và XLSX',
          'Nguyên tắc trích xuất mục lục phân cấp hỗ trợ tra cứu tức thời',
          'Chuẩn nén hình ảnh xem trước giữ độ sắc nét ở mức dung lượng tối thiểu',
        ],
        metricHighlight: {
          label: 'Chủ đề trọng tâm',
          value: '6 trang',
        },
        accentColor: '#DB2777',
      },
      {
        id: 'kt-slide-2',
        index: 2,
        title: 'Phương pháp đọc chủ động & ghi nhớ nhanh',
        subtitle: 'Ứng dụng mô hình quét cấu trúc I–IV trước khi đọc chi tiết',
        bullets: [
          'Đọc lướt dàn ý 4 phần để nắm khung lập luận trong vòng 90 giây',
          'Đối chiếu các gạch đầu dòng cốt lõi với số trang gốc để kiểm chứng',
          'Tùy chỉnh thu phóng vừa chiều rộng hoặc chuyển chế độ đọc tối ban đêm',
        ],
        metricHighlight: {
          label: 'Tăng tốc độ đọc',
          value: '2.1x',
        },
        accentColor: '#E11D48',
      },
      {
        id: 'kt-slide-3',
        index: 3,
        title: 'Danh mục đối chiếu & thực hành chuẩn',
        subtitle: 'Tổng hợp mẫu biểu và quy tắc đặt tên tệp khoa học',
        bullets: [
          'Đặt tên tài liệu theo cấu trúc chủ đề + mốc thời gian để tìm kiếm chính xác',
          'Gắn dấu sao yêu thích cho tối đa 5 tài liệu tra cứu thường xuyên mỗi tuần',
          'Phân loại theo thư mục chức năng giúp kiểm soát phiên bản rõ ràng',
        ],
        metricHighlight: {
          label: 'Tỷ lệ tra cứu trúng',
          value: '96%',
        },
        accentColor: '#9333EA',
      },
    ],
    outline: [
      {
        id: 'kt-outline-1',
        romanIndex: 'I',
        title: 'Khái niệm nền tảng',
        summary:
          'Hệ thống hóa các khái niệm cơ bản về cấu trúc tài liệu điện tử, tiêu chuẩn mã hóa chữ tiếng Việt và cơ chế hiển thị trên web hiện đại.',
        keyPoints: [
          'So sánh ưu nhược điểm giữa PDF cố định trang và tài liệu dòng chảy DOCX',
          'Vai trò của siêu dữ liệu (metadata) trong tìm kiếm và phân nhóm nhanh',
          'Nguyên tắc bảo toàn bố cục khi chuyển đổi tài liệu giữa các thiết bị',
        ],
        pageNumber: 1,
      },
      {
        id: 'kt-outline-2',
        romanIndex: 'II',
        title: 'Phương pháp phân tích',
        summary:
          'Trình bày kỹ thuật tách ý chính, lập bảng tóm tắt điều hành và nhận diện các chỉ số quan trọng trong báo cáo chuyên sâu.',
        keyPoints: [
          'Áp dụng quy tắc 3 ý trọng tâm cho mỗi phần nội dung dài trên 5 trang',
          'Liên kết trực tiếp đoạn trích tóm tắt với vị trí trang tương ứng',
          'Sử dụng ký hiệu mục lục La Mã giúp tạo cấu trúc đọc mạch lạc',
        ],
        pageNumber: 3,
      },
      {
        id: 'kt-outline-3',
        romanIndex: 'III',
        title: 'Tình huống ứng dụng',
        summary:
          'Minh họa 3 kịch bản thực tế: ôn tập chuyên môn, rà soát hồ sơ dự án và chuẩn bị tài liệu trao đổi nhanh với đối tác.',
        keyPoints: [
          'Rút ngắn 50% thời gian đọc hiểu tài liệu kỹ thuật trước buổi họp rà soát',
          'Dùng chế độ mở rộng toàn màn hình và xoay trang khi đối chiếu bảng biểu lớn',
          'Kết hợp đánh dấu yêu thích để gom nhanh bộ tài liệu cần ôn trong ngày',
        ],
        pageNumber: 4,
      },
      {
        id: 'kt-outline-4',
        romanIndex: 'IV',
        title: 'Tài liệu đối chiếu',
        summary:
          'Bảng thuật ngữ viết tắt, danh mục tiêu chuẩn tham khảo và hướng dẫn cập nhật định kỳ kho tri thức cá nhân.',
        keyPoints: [
          'Tổng hợp 18 thuật ngữ cốt lõi thường gặp trong quản trị tài liệu số',
          'Quy ước lưu trữ phiên bản giúp tránh nhầm lẫn bản nháp và bản phát hành',
          'Gợi ý lịch rà soát kho tài liệu định kỳ vào cuối mỗi tháng',
        ],
        pageNumber: 6,
      },
    ],
    pdfPages: [
      {
        id: 'kt-page-1',
        pageNumber: 1,
        sectionTitle: 'I. Khái niệm nền tảng',
        heading: 'Cấu trúc tài liệu số & phân loại định dạng phổ biến',
        paragraphs: [
          'Trong môi trường làm việc hiện đại, một tổ chức thường vận hành song song bốn nhóm định dạng tài liệu chính: bài trình chiếu Slide (PPTX), văn bản biên tập (DOCX), bảng tính phân tích (XLSX) và tài liệu cố định bố cục (PDF). Mỗi định dạng được tối ưu cho một mục đích truyền đạt riêng biệt.',
          'Khác với văn bản dòng chảy cho phép tự động ngắt dòng theo kích thước màn hình, định dạng PDF khóa chặt tọa độ từng ký tự, đồ thị và bảng biểu. Nhờ đó tài liệu giữ nguyên tính pháp lý và thẩm mỹ khi phát hành báo cáo tài chính, hợp đồng hoặc tài liệu đào tạo.',
        ],
        callout:
          'Nguyên tắc chọn định dạng: Dùng DOCX/XLSX khi cần biên tập liên tục; xuất PDF khi cần cố định bố cục để phát hành hoặc đối chiếu lâu dài.',
        highlights: [
          'PDF bảo toàn 100% bố cục, font chữ và tỷ lệ bảng biểu trên mọi hệ điều hành.',
          'Siêu dữ liệu (tiêu đề, thư mục, thẻ loại) quyết định tốc độ tìm kiếm trong kho tài liệu lớn.',
        ],
      },
      {
        id: 'kt-page-2',
        pageNumber: 2,
        sectionTitle: 'I. Khái niệm nền tảng',
        heading: 'Chuẩn hóa hiển thị tiếng Việt & tối ưu tốc độ tải trang',
        paragraphs: [
          'Một trong những nguyên nhân phổ biến gây khó chịu khi đọc tài liệu điện tử là lỗi thiếu phông chữ hỗ trợ dấu tiếng Việt hoặc tệp nặng hàng chục megabyte tải chậm trên mạng di động. Việc nhúng chuẩn bảng mã UTF-8 và phông chữ hiện đại như Be Vietnam Pro giúp văn bản luôn sắc nét ở mọi mức thu phóng.',
          'Bên cạnh yếu tố hiển thị, kỹ thuật phân trang theo nhu cầu (lazy rendering) và sinh ảnh thu nhỏ trang nhẹ cho phép người đọc nhảy ngay tới trang 15 hay trang 20 mà không phải chờ tải toàn bộ tập tin gốc.',
        ],
        callout:
          'Tốc độ mở trang đầu dưới 1 giây giúp tăng gấp đôi tần suất tra cứu tài liệu nội bộ trên thiết bị di động.',
        highlights: [
          'Chuẩn mã hóa UTF-8 ngăn triệt để hiện tượng vỡ dấu tiếng Việt khi mở chéo thiết bị.',
          'Thanh trang thu nhỏ bên trái giúp định vị nhanh các biểu đồ và phụ lục dài.',
        ],
      },
      {
        id: 'kt-page-3',
        pageNumber: 3,
        sectionTitle: 'II. Phương pháp phân tích',
        heading: 'Kỹ thuật đọc quét cấu trúc I–IV & chắt lọc ý cốt lõi',
        paragraphs: [
          'Khi tiếp cận một báo cáo dày từ 20 trang trở lên, đọc tuyến tính từ trang đầu đến trang cuối thường tiêu tốn nhiều thời gian mà khó giữ bức tranh tổng thể. Phương pháp đọc chủ động bắt đầu bằng việc quét nhanh 4 tầng mục lục La Mã để hình dung mạch lập luận của tác giả.',
          'Sau khi nắm khung sườn, người đọc áp dụng quy tắc ba ý trọng tâm cho mỗi chương: xác định luận điểm chính, trích xuất con số chứng minh và ghi chú câu hỏi cần kiểm chứng. Cách tiếp cận này giúp rút ngắn hơn một nửa thời gian chuẩn bị trước các cuộc họp chuyên môn.',
        ],
        callout:
          'Đọc lướt dàn ý trong 90 giây đầu tiên giúp não bộ tạo sẵn “ngăn chứa” thông tin trước khi đi sâu vào số liệu chi tiết.',
        highlights: [
          'Quy tắc 3 ý trọng tâm giữ bản tóm tắt luôn súc tích và dễ hành động.',
          'Gắn số trang gốc cạnh mỗi ghi chú giúp kiểm chứng lại nguồn chỉ trong 1 lần bấm.',
        ],
      },
      {
        id: 'kt-page-4',
        pageNumber: 4,
        sectionTitle: 'III. Tình huống ứng dụng',
        heading: 'Kịch bản rà soát hồ sơ dự án & tài liệu kỹ thuật',
        paragraphs: [
          'Trong giai đoạn nghiệm thu hoặc thẩm định hồ sơ, các thành viên thường cần đối chiếu chéo giữa bản mô tả thiết kế, bảng dự toán và tài liệu PDF đã ký duyệt. Sử dụng trình xem PDF có hỗ trợ tìm kiếm từ khóa tức thì trong tài liệu giúp định vị chính xác điều khoản cần kiểm tra.',
          'Đối với các trang chứa sơ đồ kiến trúc ngang hoặc bảng thông số nhiều cột, thao tác xoay trang 90 độ kết hợp thu phóng linh hoạt từ 75% đến 150% giúp quan sát trọn vẹn chi tiết mà không cần tải thêm phần mềm đọc chuyên dụng.',
        ],
        callout:
          'Kết hợp tìm kiếm từ khóa trên thanh công cụ với phím điều hướng trang giúp kiểm tra chéo hồ sơ nhanh hơn 3 lần.',
        highlights: [
          'Xoay trang và căn vừa chiều rộng đặc biệt hữu ích với bản vẽ sơ đồ và phụ lục ngang.',
          'Chế độ đọc ban đêm giảm chói mắt rõ rệt khi rà soát văn bản trong phòng tối.',
        ],
      },
      {
        id: 'kt-page-5',
        pageNumber: 5,
        sectionTitle: 'III. Tình huống ứng dụng',
        heading: 'Tổ chức kho tri thức nhóm & quy ước đặt tên tệp khoa học',
        paragraphs: [
          'Kho tài liệu chỉ thực sự phát huy giá trị khi bất kỳ thành viên mới nào cũng tìm đúng bản phát hành cuối cùng trong vòng 15 giây. Quy ước đặt tên nhất quán theo công thức [Chủ đề] + [Giai đoạn/Năm] giúp loại bỏ hoàn toàn các tên tệp mơ hồ như “baocao_final_v2_sua”.',
          'Bên cạnh việc chia thư mục theo chức năng như Dự án, Sản phẩm, Tài chính & Tiến độ hay Vận hành, mỗi cá nhân nên duy trì danh sách Yêu thích giới hạn dưới 5 tài liệu cốt lõi của tuần hiện tại để giữ góc làm việc luôn gọn nhẹ.',
        ],
        callout:
          'Một quy ước đặt tên đơn giản nhưng nhất quán có giá trị hơn hàng chục tầng thư mục lồng nhau phức tạp.',
        highlights: [
          'Phân loại theo tối đa 2 cấp thư mục kết hợp bộ lọc định dạng giúp truy xuất tức thì.',
          'Đánh dấu Yêu thích cho bộ tài liệu đang dùng trong tuần để truy cập ngay tại Trang chủ.',
        ],
      },
      {
        id: 'kt-page-6',
        pageNumber: 6,
        sectionTitle: 'IV. Tài liệu đối chiếu',
        heading: 'Bảng thuật ngữ chuẩn & lịch rà soát kho tài liệu định kỳ',
        paragraphs: [
          'Trang cuối cùng tổng hợp danh mục thuật ngữ thường dùng trong quản trị tài liệu số bao gồm: Smart Reader (trình đọc tách mục lục tự động), Live Stage (sân khấu trình chiếu trực tuyến), Metadata (siêu dữ liệu mô tả tệp) và Safe Area (vùng hiển thị an toàn trên thiết bị di động).',
          'Để kho tri thức không bị phình to bởi các bản nháp lỗi thời, nhóm vận hành khuyến nghị thực hiện phiên rà soát 15 phút vào ngày làm việc cuối cùng mỗi tháng: lưu trữ các tài liệu đã kết thúc chu kỳ, cập nhật chỉ số mới và kiểm tra lại quyền chia sẻ liên kết công khai.',
        ],
        callout:
          'Lịch bảo trì đề xuất: 15 phút cuối tháng để dọn bản nháp, cập nhật quyền chia sẻ và đồng bộ tài liệu trọng tâm.',
        highlights: [
          'Rà soát quyền chia sẻ định kỳ giúp bảo vệ tài liệu nội bộ nhạy cảm.',
          'Lưu trữ gọn tài liệu cũ giữ hiệu năng tìm kiếm và lọc luôn nhanh như ngày đầu.',
        ],
      },
    ],
    contentPreview:
      'Cuốn sổ tay tổng hợp kiến thức biên soạn ngắn gọn các nguyên lý tổ chức thông tin, kỹ thuật đọc hiểu tài liệu dài và quy chuẩn quản trị kho lưu trữ cá nhân. Cấu trúc 6 trang phân chia theo 4 phần rõ ràng kèm khung nhấn mạnh và điểm cốt lõi giúp người đọc nhanh chóng nắm bắt trọng tâm mà không bị quá tải bởi văn bản thuần.',
  },
  {
    id: 'doc-kehoach-nhom',
    title: 'Kế hoạch nhóm',
    type: 'pptx',
    updatedAt: '18/11/2026',
    isFavorite: false,
    isShared: true,
    isUploaded: false,
    folder: 'Vận hành',
    pageCount: 14,
    fileSize: '3.1 MB',
    thumbnail: createThumbnailDataUri('pptx', 'Kế hoạch nhóm', '#EA580C'),
    slidesPreview: [
      {
        id: 'kh-slide-1',
        index: 1,
        title: 'Mục tiêu ưu tiên Quý IV — Nhóm Sản phẩm',
        subtitle: 'Đồng bộ nhịp độ phát triển và bàn giao đúng hạn 100% hạng mục lõi',
        bullets: [
          'Hoàn thiện giao diện Trang chủ toàn màn hình và 4 không gian làm việc chuyên biệt',
          'Giữ chu kỳ phát hành ổn định mỗi 2 tuần kèm phiên trình bày nội bộ',
          'Nâng tỷ lệ tự động hóa kiểm thử luồng chính lên trên 85%',
        ],
        metricHighlight: {
          label: 'Cam kết đúng hạn',
          value: '95%',
        },
        accentColor: '#EA580C',
        speakerNote:
          'Thống nhất 3 ưu tiên cao nhất để toàn nhóm tránh dàn trải nguồn lực vào các yêu cầu phụ.',
      },
      {
        id: 'kh-slide-2',
        index: 2,
        title: 'Phân bổ nguồn lực & vai trò chủ chốt',
        subtitle: 'Cơ chế phối hợp gọn nhẹ giữa Thiết kế, Frontend và Kiểm thử',
        bullets: [
          'Nhóm Thiết kế chốt hệ thống thẻ, bảng màu và kiểm tra sắc độ đọc 3 chế độ giấy',
          'Nhóm Frontend chia tách rõ tầng dữ liệu, thư viện và 4 không gian định dạng riêng',
          'Nhóm Vận hành tổng hợp phản hồi người dùng thử nghiệm mỗi thứ Sáu',
        ],
        metricHighlight: {
          label: 'Thành viên tham gia',
          value: '8 người',
        },
        accentColor: '#F59E0B',
      },
      {
        id: 'kh-slide-3',
        index: 3,
        title: 'Lịch trình 4 Sprint trọng điểm',
        subtitle: 'Các mốc bàn giao nối tiếp rõ đầu ra và tiêu chí nghiệm thu',
        bullets: [
          'Sprint 1: Khởi tạo khung ứng dụng, kiểu dữ liệu và dữ liệu mẫu chuẩn 4 định dạng',
          'Sprint 2: Hoàn thiện Trang chủ tìm kiếm, lọc đa chiều và kéo thả tải tệp cục bộ',
          'Sprint 3–4: Hoàn thiện Slide, Word, Excel, PDF và chia sẻ liên kết linh hoạt',
        ],
        metricHighlight: {
          label: 'Thời lượng mỗi Sprint',
          value: '2 tuần',
        },
        accentColor: '#4F46E5',
      },
      {
        id: 'kh-slide-4',
        index: 4,
        title: 'Chỉ số đo lường & quản trị rủi ro',
        subtitle: 'Theo dõi sức khỏe tiến độ minh bạch qua bảng tổng hợp tuần',
        bullets: [
          'Đo thời gian tải khung làm việc đầu tiên dưới 1 giây trên mạng phổ thông',
          'Duy trì 0 lỗi TypeScript và 0 lỗi vỡ bố cục trên cả desktop lẫn di động',
          'Dự phòng 15% quỹ thời gian mỗi Sprint để xử lý phản hồi phát sinh',
        ],
        metricHighlight: {
          label: 'Ngân sách dự phòng',
          value: '15%',
        },
        accentColor: '#059669',
      },
    ],
    outline: [
      {
        id: 'kh-outline-1',
        romanIndex: 'I',
        title: 'Mục tiêu quý',
        summary:
          'Chốt danh sách mục tiêu đo lường được cho toàn nhóm trong chu kỳ hiện tại, bám sát định hướng ra mắt phiên bản ổn định.',
        keyPoints: [
          'Ưu tiên trải nghiệm mượt mà ở các thao tác cốt lõi thay vì nhồi tính năng',
          'Mọi hạng mục bàn giao đều có tiêu chí hoàn thành định lượng rõ ràng',
          'Thống nhất buổi trình chiếu đánh giá tiến độ vào chiều thứ Năm hàng tuần',
        ],
        pageNumber: 1,
      },
      {
        id: 'kh-outline-2',
        romanIndex: 'II',
        title: 'Phân bổ nguồn lực',
        summary:
          'Xác định rõ đầu mối chịu trách nhiệm từng phân hệ giúp các thành viên phối hợp song song mà không chồng chéo công việc.',
        keyPoints: [
          'Phân tách sở hữu tệp và hợp đồng dữ liệu minh bạch ngay từ đầu Sprint',
          'Thiết lập kênh trao đổi ngắn gọn xử lý vướng mắc kỹ thuật trong 30 phút',
          'Cân bằng khối lượng công việc giữa phát triển tính năng và tối ưu chất lượng',
        ],
        pageNumber: 4,
      },
      {
        id: 'kh-outline-3',
        romanIndex: 'III',
        title: 'Lịch trình Sprint',
        summary:
          'Phân rã lộ trình thành 4 chặng hai tuần với mốc kiểm thử tích hợp rõ ràng trước khi chuyển sang giai đoạn mở rộng.',
        keyPoints: [
          'Hoàn thành bản dựng khung giao diện và dữ liệu mẫu ngay trong chặng đầu',
          'Dành riêng 2 ngày cuối mỗi Sprint cho rà soát tương tác và đánh bóng UI',
          'Cập nhật trực tiếp trạng thái tài liệu kế hoạch vào thư mục Vận hành chung',
        ],
        pageNumber: 8,
      },
      {
        id: 'kh-outline-4',
        romanIndex: 'IV',
        title: 'Chỉ số đo lường',
        summary:
          'Định nghĩa bộ KPI theo dõi hiệu suất nhóm, chất lượng mã nguồn và mức độ hài lòng của người dùng trong giai đoạn dùng thử.',
        keyPoints: [
          'Tỷ lệ hoàn thành hạng mục đúng hạn đạt tối thiểu 95% qua các Sprint',
          'Thời gian phản hồi khi chuyển tài liệu hoặc lọc tìm kiếm luôn tức thì',
          'Thu thập ít nhất 20 lượt góp ý có cấu trúc trước khi đóng gói phiên bản mới',
        ],
        pageNumber: 12,
      },
    ],
    contentPreview:
      'Kế hoạch nhóm trình bày chi tiết mục tiêu quý, phân bổ trách nhiệm từng thành viên và lịch trình 4 Sprint trọng điểm của khối vận hành. Bộ chỉ số theo dõi minh bạch cùng quỹ thời gian dự phòng giúp cả nhóm duy trì nhịp bàn giao ổn định và chủ động thích ứng với các thay đổi phát sinh.',
  },
];

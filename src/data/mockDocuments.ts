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

export const LEGACY_MOCK_DOCUMENT_IDS: readonly string[] = [
  'doc-baocao-2026',
  'doc-thietke-sanpham',
  'doc-budget-2026',
  'doc-tonghop-kienthuc',
  'doc-kehoach-nhom',
] as const;

export const mockDocuments: DocumentItem[] = [];

import { useState } from 'react';
import type { ChangeEvent } from 'react';
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  BarChart3,
  Circle,
  Link2,
  MessageSquare,
  Minus,
  Music,
  PieChart,
  Square,
  Star,
  Triangle,
  TrendingUp,
  Video,
  X,
} from 'lucide-react';
import type { SlideTextBlock } from '../../types/document';

export type PptShapeKind =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'arrowRight'
  | 'arrowLeft'
  | 'arrowUp'
  | 'arrowDown'
  | 'star'
  | 'callout'
  | 'line';

export type PptSlideLayoutId =
  | 'title_slide'
  | 'title_content'
  | 'two_column'
  | 'comparison'
  | 'title_only'
  | 'blank';

export interface PptChartData {
  chartType: 'column' | 'line' | 'pie';
  title?: string;
  categories: string[];
  series: { name: string; values: number[]; color?: string }[];
}

export interface PptColorTheme {
  id: string;
  name: string;
  bgColor: string;
  bgGradient?: string;
  accentColor: string;
  titleColor: string;
  bodyColor: string;
  isLight: boolean;
  swatches: [string, string, string];
}

export const PPT_COLOR_THEMES: readonly PptColorTheme[] = [
  {
    id: 'ppt-office-dark',
    name: 'Đêm sâu Indigo',
    bgColor: '#161D2E',
    accentColor: '#4F46E5',
    titleColor: '#FFFFFF',
    bodyColor: '#E2E8F0',
    isLight: false,
    swatches: ['#161D2E', '#4F46E5', '#38BDF8'],
  },
  {
    id: 'ppt-crimson-office',
    name: 'PowerPoint Đỏ Cam',
    bgColor: '#FFFFFF',
    accentColor: '#C43E1C',
    titleColor: '#1E293B',
    bodyColor: '#334155',
    isLight: true,
    swatches: ['#FFFFFF', '#C43E1C', '#F97316'],
  },
  {
    id: 'ppt-classic-blue',
    name: 'Xanh Công Sở',
    bgColor: '#F8FAFC',
    accentColor: '#185ABD',
    titleColor: '#0F172A',
    bodyColor: '#334155',
    isLight: true,
    swatches: ['#F8FAFC', '#185ABD', '#0284C7'],
  },
  {
    id: 'ppt-warm-paper',
    name: 'Giấy Kem Ấm',
    bgColor: '#FAF8F5',
    accentColor: '#B45309',
    titleColor: '#1C1917',
    bodyColor: '#44403C',
    isLight: true,
    swatches: ['#FAF8F5', '#B45309', '#D97706'],
  },
  {
    id: 'ppt-emerald-night',
    name: 'Ngọc Lục Bảo',
    bgColor: '#06281E',
    bgGradient: 'linear-gradient(135deg, #06281E 0%, #0F3D2E 100%)',
    accentColor: '#10B981',
    titleColor: '#ECFDF5',
    bodyColor: '#A7F3D0',
    isLight: false,
    swatches: ['#06281E', '#10B981', '#34D399'],
  },
  {
    id: 'ppt-deep-ocean',
    name: 'Biển Sâu',
    bgColor: '#0F172A',
    bgGradient: 'linear-gradient(145deg, #0F172A 0%, #1E293B 60%, #0C4A6E 100%)',
    accentColor: '#38BDF8',
    titleColor: '#F8FAFC',
    bodyColor: '#CBD5E1',
    isLight: false,
    swatches: ['#0F172A', '#38BDF8', '#0EA5E9'],
  },
  {
    id: 'ppt-sunset-plum',
    name: 'Tím Hoàng Hôn',
    bgColor: '#1E1B4B',
    bgGradient: 'linear-gradient(140deg, #1E1B4B 0%, #4C1D95 58%, #831843 100%)',
    accentColor: '#F472B6',
    titleColor: '#FDF2F8',
    bodyColor: '#E9D5FF',
    isLight: false,
    swatches: ['#1E1B4B', '#F472B6', '#A855F7'],
  },
  {
    id: 'ppt-charcoal-gold',
    name: 'Than Chì Vàng',
    bgColor: '#18181B',
    accentColor: '#FACC15',
    titleColor: '#FAFAFA',
    bodyColor: '#D4D4D8',
    isLight: false,
    swatches: ['#18181B', '#FACC15', '#EAB308'],
  },
];

export interface PptLayoutOption {
  id: PptSlideLayoutId;
  label: string;
  desc: string;
}

export const PPT_SLIDE_LAYOUTS: readonly PptLayoutOption[] = [
  { id: 'title_slide', label: 'Trang tiêu đề', desc: 'Tiêu đề lớn giữa trang + phụ đề' },
  { id: 'title_content', label: 'Tiêu đề & Nội dung', desc: 'Tiêu đề trên + khung nội dung lớn' },
  { id: 'two_column', label: 'Hai cột', desc: 'Tiêu đề + 2 cột nội dung song song' },
  { id: 'comparison', label: 'So sánh', desc: 'Tiêu đề + 2 nhóm tiêu đề phụ & nội dung' },
  { id: 'title_only', label: 'Chỉ có tiêu đề', desc: 'Tiêu đề phía trên + khoảng trống tự do' },
  { id: 'blank', label: 'Trang trống', desc: 'Trang trắng sạch không có khung sẵn' },
];

export function createBlocksForSlideLayout(
  layoutId: PptSlideLayoutId,
  slideId: string,
  colors: { titleColor: string; bodyColor: string; accentColor: string }
): SlideTextBlock[] {
  const stamp = Date.now();
  if (layoutId === 'blank') {
    return [];
  }
  if (layoutId === 'title_slide') {
    return [
      {
        id: `${slideId}-lay-title-${stamp}`,
        text: 'Tiêu đề bài thuyết trình',
        fontSize: 40,
        bold: true,
        color: colors.titleColor,
        align: 'center',
        valign: 'bottom',
        role: 'title',
        xPct: 8,
        yPct: 24,
        wPct: 84,
        hPct: 24,
      },
      {
        id: `${slideId}-lay-sub-${stamp}`,
        text: 'Phụ đề hoặc tên người trình bày · LITE Workspace',
        fontSize: 20,
        bold: false,
        color: colors.bodyColor,
        align: 'center',
        valign: 'top',
        role: 'subtitle',
        xPct: 14,
        yPct: 52,
        wPct: 72,
        hPct: 16,
      },
    ];
  }
  if (layoutId === 'title_only') {
    return [
      {
        id: `${slideId}-lay-title-${stamp}`,
        text: 'Tiêu đề trang chiếu',
        fontSize: 32,
        bold: true,
        color: colors.titleColor,
        align: 'left',
        valign: 'middle',
        role: 'title',
        xPct: 6,
        yPct: 7,
        wPct: 88,
        hPct: 15,
      },
    ];
  }
  if (layoutId === 'two_column') {
    return [
      {
        id: `${slideId}-lay-title-${stamp}`,
        text: 'Tiêu đề hai cột',
        fontSize: 30,
        bold: true,
        color: colors.titleColor,
        align: 'left',
        valign: 'middle',
        role: 'title',
        xPct: 6,
        yPct: 7,
        wPct: 88,
        hPct: 14,
      },
      {
        id: `${slideId}-lay-col1-${stamp}`,
        text: '• Nội dung cột trái 1\n• Nội dung cột trái 2\n• Nội dung cột trái 3',
        fontSize: 18,
        bold: false,
        color: colors.bodyColor,
        align: 'left',
        valign: 'top',
        role: 'body',
        borderColor: `${colors.accentColor}55`,
        borderWidthPx: 1,
        borderRadius: '8px',
        xPct: 6,
        yPct: 25,
        wPct: 42,
        hPct: 62,
      },
      {
        id: `${slideId}-lay-col2-${stamp}`,
        text: '• Nội dung cột phải 1\n• Nội dung cột phải 2\n• Nội dung cột phải 3',
        fontSize: 18,
        bold: false,
        color: colors.bodyColor,
        align: 'left',
        valign: 'top',
        role: 'body',
        borderColor: `${colors.accentColor}55`,
        borderWidthPx: 1,
        borderRadius: '8px',
        xPct: 52,
        yPct: 25,
        wPct: 42,
        hPct: 62,
      },
    ];
  }
  if (layoutId === 'comparison') {
    return [
      {
        id: `${slideId}-lay-title-${stamp}`,
        text: 'So sánh phương án',
        fontSize: 30,
        bold: true,
        color: colors.titleColor,
        align: 'left',
        valign: 'middle',
        role: 'title',
        xPct: 6,
        yPct: 6,
        wPct: 88,
        hPct: 13,
      },
      {
        id: `${slideId}-lay-h1-${stamp}`,
        text: 'Phương án A',
        fontSize: 20,
        bold: true,
        color: '#FFFFFF',
        fillColor: colors.accentColor,
        borderRadius: '8px',
        align: 'center',
        valign: 'middle',
        role: 'subtitle',
        xPct: 6,
        yPct: 22,
        wPct: 42,
        hPct: 10,
      },
      {
        id: `${slideId}-lay-b1-${stamp}`,
        text: '• Ưu điểm nổi bật\n• Chi phí triển khai\n• Thời gian hoàn thành',
        fontSize: 16,
        bold: false,
        color: colors.bodyColor,
        borderColor: `${colors.accentColor}55`,
        borderWidthPx: 1,
        borderRadius: '8px',
        align: 'left',
        valign: 'top',
        role: 'body',
        xPct: 6,
        yPct: 34,
        wPct: 42,
        hPct: 54,
      },
      {
        id: `${slideId}-lay-h2-${stamp}`,
        text: 'Phương án B',
        fontSize: 20,
        bold: true,
        color: '#FFFFFF',
        fillColor: colors.accentColor,
        borderRadius: '8px',
        align: 'center',
        valign: 'middle',
        role: 'subtitle',
        xPct: 52,
        yPct: 22,
        wPct: 42,
        hPct: 10,
      },
      {
        id: `${slideId}-lay-b2-${stamp}`,
        text: '• Ưu điểm nổi bật\n• Chi phí triển khai\n• Thời gian hoàn thành',
        fontSize: 16,
        bold: false,
        color: colors.bodyColor,
        borderColor: `${colors.accentColor}55`,
        borderWidthPx: 1,
        borderRadius: '8px',
        align: 'left',
        valign: 'top',
        role: 'body',
        xPct: 52,
        yPct: 34,
        wPct: 42,
        hPct: 54,
      },
    ];
  }
  // title_content
  return [
    {
      id: `${slideId}-lay-title-${stamp}`,
      text: 'Tiêu đề trang chiếu',
      fontSize: 32,
      bold: true,
      color: colors.titleColor,
      align: 'left',
      valign: 'middle',
      role: 'title',
      xPct: 6,
      yPct: 8,
      wPct: 88,
      hPct: 16,
    },
    {
      id: `${slideId}-lay-body-${stamp}`,
      text: '• Nhập ý chính đầu tiên\n• Nhập ý chính tiếp theo\n• Bấm vào để chỉnh sửa nội dung',
      fontSize: 18,
      bold: false,
      color: colors.bodyColor,
      align: 'left',
      valign: 'top',
      role: 'body',
      xPct: 6,
      yPct: 28,
      wPct: 88,
      hPct: 58,
    },
  ];
}

export const PPT_SHAPE_OPTIONS: ReadonlyArray<{
  kind: PptShapeKind;
  label: string;
  Icon: typeof Square;
}> = [
  { kind: 'rect', label: 'Chữ nhật', Icon: Square },
  { kind: 'roundRect', label: 'Bo góc', Icon: Square },
  { kind: 'ellipse', label: 'Hình tròn', Icon: Circle },
  { kind: 'triangle', label: 'Tam giác', Icon: Triangle },
  { kind: 'arrowRight', label: 'Mũi tên phải', Icon: ArrowRight },
  { kind: 'arrowLeft', label: 'Mũi tên trái', Icon: ArrowLeft },
  { kind: 'arrowUp', label: 'Mũi tên lên', Icon: ArrowUp },
  { kind: 'arrowDown', label: 'Mũi tên xuống', Icon: ArrowDown },
  { kind: 'star', label: 'Ngôi sao', Icon: Star },
  { kind: 'callout', label: 'Chú thích', Icon: MessageSquare },
  { kind: 'line', label: 'Đường kẻ', Icon: Minus },
];

export function SlideShapeSvg({
  shapeKind,
  fillColor,
  borderColor,
  borderWidthPx = 2,
}: {
  shapeKind: PptShapeKind;
  fillColor?: string;
  borderColor?: string;
  borderWidthPx?: number;
}) {
  const fill = fillColor || 'rgba(196, 62, 28, 0.22)';
  const stroke = borderColor || '#F97316';
  const sw = Math.max(1, borderWidthPx);

  return (
    <svg
      viewBox="0 0 100 100"
      preserveAspectRatio="none"
      aria-hidden="true"
      className="absolute inset-0 w-full h-full pointer-events-none"
    >
      {shapeKind === 'rect' && (
        <rect
          x="2"
          y="2"
          width="96"
          height="96"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'roundRect' && (
        <rect
          x="2"
          y="2"
          width="96"
          height="96"
          rx="14"
          ry="14"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'ellipse' && (
        <ellipse
          cx="50"
          cy="50"
          rx="48"
          ry="48"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'triangle' && (
        <polygon
          points="50,3 97,97 3,97"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'arrowRight' && (
        <polygon
          points="3,30 62,30 62,10 97,50 62,90 62,70 3,70"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'arrowLeft' && (
        <polygon
          points="97,30 38,30 38,10 3,50 38,90 38,70 97,70"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'arrowUp' && (
        <polygon
          points="30,97 30,38 10,38 50,3 90,38 70,38 70,97"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'arrowDown' && (
        <polygon
          points="30,3 30,62 10,62 50,97 90,62 70,62 70,3"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'star' && (
        <polygon
          points="50,3 63,35 97,38 71,61 79,95 50,77 21,95 29,61 3,38 37,35"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'callout' && (
        <path
          d="M 6,6 H 94 A 4,4 0 0 1 98,10 V 72 A 4,4 0 0 1 94,76 H 38 L 20,96 L 24,76 H 6 A 4,4 0 0 1 2,72 V 10 A 4,4 0 0 1 6,6 Z"
          fill={fill}
          stroke={stroke}
          strokeWidth={sw}
          vectorEffect="non-scaling-stroke"
        />
      )}
      {shapeKind === 'line' && (
        <line
          x1="2"
          y1="50"
          x2="98"
          y2="50"
          stroke={stroke}
          strokeWidth={Math.max(2, sw + 1)}
          strokeLinecap="round"
          vectorEffect="non-scaling-stroke"
        />
      )}
    </svg>
  );
}

const CHART_PALETTE = ['#C43E1C', '#2563EB', '#16A34A', '#D97706', '#9333EA'];

export function createDefaultSlideChartData(
  chartType: 'column' | 'line' | 'pie'
): PptChartData {
  return {
    chartType,
    title:
      chartType === 'pie'
        ? 'Cơ cấu tỷ trọng (%)'
        : chartType === 'line'
          ? 'Xu hướng tăng trưởng theo quý'
          : 'Doanh thu theo quý (triệu VNĐ)',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series:
      chartType === 'pie'
        ? [{ name: 'Tỷ trọng', values: [35, 28, 22, 15], color: '#C43E1C' }]
        : [
            { name: '2025', values: [42, 58, 65, 84], color: '#C43E1C' },
            { name: '2026', values: [54, 72, 81, 98], color: '#2563EB' },
          ],
  };
}

export function SlideChartSvg({
  chartData,
  hasLightBg,
  fontScale,
}: {
  chartData: PptChartData;
  hasLightBg: boolean;
  fontScale: number;
}) {
  const categories = chartData.categories.length > 0 ? chartData.categories : ['A', 'B', 'C'];
  const series = chartData.series.length > 0 ? chartData.series : [{ name: 'Dữ liệu', values: [10, 20, 30], color: '#C43E1C' }];
  const titlePx = Math.max(7, Math.round(14 * fontScale));
  const labelPx = Math.max(5.5, Math.round(10 * fontScale));
  const textColor = hasLightBg ? '#1E293B' : '#F1F5F9';
  const mutedColor = hasLightBg ? '#64748B' : '#94A3B8';
  const gridColor = hasLightBg ? 'rgba(15,23,42,0.12)' : 'rgba(255,255,255,0.14)';

  const allValues = series.flatMap((s) => s.values);
  const maxVal = Math.max(1, ...allValues, 10);

  return (
    <div
      className={`w-full h-full flex flex-col justify-between p-2 rounded-lg border overflow-hidden select-none ${
        hasLightBg
          ? 'bg-white/90 border-slate-200'
          : 'bg-slate-900/70 border-white/15'
      }`}
    >
      {chartData.title && (
        <p
          className="font-bold text-center truncate shrink-0"
          style={{ fontSize: `${titlePx}px`, color: textColor }}
        >
          {chartData.title}
        </p>
      )}

      <div className="flex-1 min-h-0 w-full flex items-center justify-center py-1">
        {chartData.chartType === 'pie' ? (
          (() => {
            const vals = series[0]?.values ?? [30, 40, 30];
            const total = Math.max(1, vals.reduce((a, b) => a + Math.max(0, b), 0));
            let cumAngle = -Math.PI / 2;
            const cx = 60;
            const cy = 50;
            const r = 40;
            return (
              <svg viewBox="0 0 180 100" className="w-full h-full">
                {vals.map((val, idx) => {
                  const safeVal = Math.max(0, val);
                  const sliceAngle = (safeVal / total) * Math.PI * 2;
                  const x1 = cx + r * Math.cos(cumAngle);
                  const y1 = cy + r * Math.sin(cumAngle);
                  const endAngle = cumAngle + sliceAngle;
                  const x2 = cx + r * Math.cos(endAngle);
                  const y2 = cy + r * Math.sin(endAngle);
                  const largeArc = sliceAngle > Math.PI ? 1 : 0;
                  const pathD = `M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
                  cumAngle = endAngle;
                  const sliceColor = CHART_PALETTE[idx % CHART_PALETTE.length];
                  return (
                    <g key={idx}>
                      <path d={pathD} fill={sliceColor} stroke={hasLightBg ? '#fff' : '#0f172a'} strokeWidth="1.5" />
                      <rect x="112" y={18 + idx * 17} width="8" height="8" rx="2" fill={sliceColor} />
                      <text x="124" y={25 + idx * 17} fontSize="8" fill={textColor}>
                        {categories[idx] ?? `Mục ${idx + 1}`} ({Math.round((safeVal / total) * 100)}%)
                      </text>
                    </g>
                  );
                })}
              </svg>
            );
          })()
        ) : chartData.chartType === 'line' ? (
          <svg viewBox="0 0 240 120" className="w-full h-full">
            {[0, 0.25, 0.5, 0.75, 1].map((t, i) => {
              const y = 10 + (1 - t) * 86;
              return (
                <line
                  key={i}
                  x1="28"
                  y1={y}
                  x2="230"
                  y2={y}
                  stroke={gridColor}
                  strokeWidth="1"
                />
              );
            })}
            {categories.map((cat, idx) => {
              const x =
                categories.length === 1
                  ? 129
                  : 38 + (idx / (categories.length - 1)) * 182;
              return (
                <text
                  key={idx}
                  x={x}
                  y="112"
                  textAnchor="middle"
                  fontSize="8.5"
                  fill={mutedColor}
                >
                  {cat}
                </text>
              );
            })}
            {series.map((s, sIdx) => {
              const color = s.color || CHART_PALETTE[sIdx % CHART_PALETTE.length];
              const pts = categories.map((_, idx) => {
                const x =
                  categories.length === 1
                    ? 129
                    : 38 + (idx / (categories.length - 1)) * 182;
                const val = s.values[idx] ?? 0;
                const y = 96 - (Math.max(0, val) / maxVal) * 82;
                return { x, y, val };
              });
              const poly = pts.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
              return (
                <g key={sIdx}>
                  <polyline
                    fill="none"
                    stroke={color}
                    strokeWidth="2.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    points={poly}
                  />
                  {pts.map((p, pIdx) => (
                    <circle key={pIdx} cx={p.x} cy={p.y} r="3.2" fill={color} />
                  ))}
                </g>
              );
            })}
          </svg>
        ) : (
          <svg viewBox="0 0 240 120" className="w-full h-full">
            {[0, 0.25, 0.5, 0.75, 1].map((t, i) => {
              const y = 10 + (1 - t) * 86;
              return (
                <line
                  key={i}
                  x1="24"
                  y1={y}
                  x2="232"
                  y2={y}
                  stroke={gridColor}
                  strokeWidth="1"
                />
              );
            })}
            {categories.map((cat, idx) => {
              const groupWidth = 200 / categories.length;
              const groupX = 28 + idx * groupWidth;
              const barGap = 3;
              const barW = Math.max(
                6,
                (groupWidth - barGap * (series.length + 1)) / series.length
              );
              return (
                <g key={idx}>
                  {series.map((s, sIdx) => {
                    const val = Math.max(0, s.values[idx] ?? 0);
                    const barH = (val / maxVal) * 82;
                    const x = groupX + barGap + sIdx * (barW + barGap);
                    const y = 96 - barH;
                    const color =
                      s.color || CHART_PALETTE[sIdx % CHART_PALETTE.length];
                    return (
                      <rect
                        key={sIdx}
                        x={x}
                        y={y}
                        width={barW}
                        height={Math.max(2, barH)}
                        rx="2"
                        fill={color}
                      />
                    );
                  })}
                  <text
                    x={groupX + groupWidth / 2}
                    y="112"
                    textAnchor="middle"
                    fontSize="8.5"
                    fill={mutedColor}
                  >
                    {cat}
                  </text>
                </g>
              );
            })}
          </svg>
        )}
      </div>

      {chartData.chartType !== 'pie' && (
        <div className="shrink-0 flex flex-wrap items-center justify-center gap-2.5 pt-0.5">
          {series.map((s, idx) => (
            <span
              key={idx}
              className="inline-flex items-center gap-1"
              style={{ fontSize: `${labelPx}px`, color: mutedColor }}
            >
              <span
                className="h-2 w-2 rounded-sm shrink-0"
                style={{
                  backgroundColor:
                    s.color || CHART_PALETTE[idx % CHART_PALETTE.length],
                }}
              />
              <span>{s.name}</span>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

export function SlideChartInsertModal({
  isOpen,
  onClose,
  onInsertChart,
}: {
  isOpen: boolean;
  onClose: () => void;
  onInsertChart: (chartData: PptChartData) => void;
}) {
  const [chartType, setChartType] = useState<'column' | 'line' | 'pie'>('column');
  const [title, setTitle] = useState<string>('Biểu đồ tăng trưởng theo quý');
  const [categoriesText, setCategoriesText] = useState<string>('Q1, Q2, Q3, Q4');
  const [seriesName, setSeriesName] = useState<string>('Số liệu');
  const [valuesText, setValuesText] = useState<string>('45, 68, 82, 96');

  if (!isOpen) return null;

  const handleConfirm = () => {
    const categories = categoriesText
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    const values = valuesText
      .split(',')
      .map((s) => Number.parseFloat(s.trim()))
      .map((n) => (Number.isFinite(n) ? n : 0));

    const safeCats = categories.length > 0 ? categories : ['Q1', 'Q2', 'Q3', 'Q4'];
    const safeVals = safeCats.map((_, idx) => values[idx] ?? (idx + 1) * 20);

    onInsertChart({
      chartType,
      title: title.trim() || undefined,
      categories: safeCats,
      series: [
        {
          name: seriesName.trim() || 'Dữ liệu',
          values: safeVals,
          color: '#C43E1C',
        },
      ],
    });
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Chèn biểu đồ vào Slide"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm flex items-center justify-center p-4"
    >
      <div className="w-full max-w-md rounded-2xl bg-[#181D27] border border-white/15 shadow-2xl overflow-hidden text-slate-100">
        <div className="px-4 py-3 border-b border-white/10 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <BarChart3 className="h-4 w-4 text-[#F97316]" />
            <h3 className="text-sm font-semibold">Chèn Biểu đồ PowerPoint</h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg hover:bg-white/10 text-slate-400 hover:text-white"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="p-4 space-y-3 text-xs">
          <div>
            <span className="block text-slate-400 mb-1.5 font-medium">Loại biểu đồ</span>
            <div className="grid grid-cols-3 gap-2">
              {(
                [
                  { id: 'column', label: 'Cột (Column)', Icon: BarChart3 },
                  { id: 'line', label: 'Đường (Line)', Icon: TrendingUp },
                  { id: 'pie', label: 'Tròn (Pie)', Icon: PieChart },
                ] as const
              ).map(({ id, label, Icon }) => (
                <button
                  key={id}
                  type="button"
                  onClick={() => setChartType(id)}
                  className={`flex flex-col items-center gap-1 p-2.5 rounded-xl border transition-colors ${
                    chartType === id
                      ? 'bg-[#C43E1C]/25 border-[#C43E1C] text-white'
                      : 'bg-white/[0.04] border-white/10 text-slate-300 hover:bg-white/10'
                  }`}
                >
                  <Icon className="h-4 w-4" />
                  <span>{label}</span>
                </button>
              ))}
            </div>
          </div>

          <div>
            <label className="block text-slate-400 mb-1 font-medium">Tiêu đề biểu đồ</label>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
            />
          </div>

          <div>
            <label className="block text-slate-400 mb-1 font-medium">
              Nhãn danh mục (ngăn cách bởi dấu phẩy)
            </label>
            <input
              type="text"
              value={categoriesText}
              onChange={(e) => setCategoriesText(e.target.value)}
              placeholder="Q1, Q2, Q3, Q4"
              className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
            />
          </div>

          <div className="grid grid-cols-2 gap-2.5">
            <div>
              <label className="block text-slate-400 mb-1 font-medium">Tên chuỗi số liệu</label>
              <input
                type="text"
                value={seriesName}
                onChange={(e) => setSeriesName(e.target.value)}
                className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
              />
            </div>
            <div>
              <label className="block text-slate-400 mb-1 font-medium">Giá trị (dấu phẩy)</label>
              <input
                type="text"
                value={valuesText}
                onChange={(e) => setValuesText(e.target.value)}
                placeholder="45, 68, 82, 96"
                className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
              />
            </div>
          </div>
        </div>

        <div className="px-4 py-3 border-t border-white/10 bg-[#141923] flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 rounded-xl bg-white/10 hover:bg-white/15 text-xs font-medium text-slate-200"
          >
            Hủy
          </button>
          <button
            type="button"
            onClick={handleConfirm}
            className="px-4 py-1.5 rounded-xl bg-[#C43E1C] hover:bg-[#a83316] text-xs font-semibold text-white"
          >
            Chèn biểu đồ
          </button>
        </div>
      </div>
    </div>
  );
}

export function SlideLinkMediaModal({
  mode,
  onClose,
  onInsertLink,
  onInsertMedia,
}: {
  mode: 'link' | 'video' | 'audio' | null;
  onClose: () => void;
  onInsertLink: (text: string, url: string) => void;
  onInsertMedia: (kind: 'video' | 'audio', src: string, label: string) => void;
}) {
  const [urlInput, setUrlInput] = useState<string>('');
  const [labelInput, setLabelInput] = useState<string>('');

  if (!mode) return null;

  const handleFilePick = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || mode === 'link') return;
    const blobUrl = URL.createObjectURL(file);
    onInsertMedia(mode, blobUrl, file.name);
    setUrlInput('');
    setLabelInput('');
    onClose();
  };

  const handleSubmit = () => {
    const cleanUrl = urlInput.trim();
    if (!cleanUrl) return;
    if (mode === 'link') {
      const normalized = /^https?:\/\//i.test(cleanUrl)
        ? cleanUrl
        : `https://${cleanUrl}`;
      onInsertLink(labelInput.trim() || normalized, normalized);
    } else {
      onInsertMedia(mode, cleanUrl, labelInput.trim() || (mode === 'video' ? 'Video nhúng' : 'Âm thanh nhúng'));
    }
    setUrlInput('');
    setLabelInput('');
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Chèn liên kết hoặc đa phương tiện"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm flex items-center justify-center p-4"
    >
      <div className="w-full max-w-md rounded-2xl bg-[#181D27] border border-white/15 shadow-2xl overflow-hidden text-slate-100">
        <div className="px-4 py-3 border-b border-white/10 flex items-center justify-between">
          <div className="flex items-center gap-2">
            {mode === 'link' ? (
              <Link2 className="h-4 w-4 text-sky-400" />
            ) : mode === 'video' ? (
              <Video className="h-4 w-4 text-[#F97316]" />
            ) : (
              <Music className="h-4 w-4 text-emerald-400" />
            )}
            <h3 className="text-sm font-semibold">
              {mode === 'link'
                ? 'Chèn Liên kết (Hyperlink)'
                : mode === 'video'
                  ? 'Chèn Video vào Slide'
                  : 'Chèn Âm thanh vào Slide'}
            </h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Đóng"
            className="p-1.5 rounded-lg hover:bg-white/10 text-slate-400 hover:text-white"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <div className="p-4 space-y-3 text-xs">
          <div>
            <label className="block text-slate-400 mb-1 font-medium">
              {mode === 'link' ? 'Văn bản hiển thị' : 'Tên mô tả (tùy chọn)'}
            </label>
            <input
              type="text"
              value={labelInput}
              onChange={(e) => setLabelInput(e.target.value)}
              placeholder={
                mode === 'link'
                  ? 'Ví dụ: Xem báo cáo chi tiết'
                  : 'Ví dụ: Video giới thiệu sản phẩm'
              }
              className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
            />
          </div>

          <div>
            <label className="block text-slate-400 mb-1 font-medium">
              {mode === 'link' ? 'Đường dẫn URL' : 'URL Video / Âm thanh'}
            </label>
            <input
              type="url"
              value={urlInput}
              onChange={(e) => setUrlInput(e.target.value)}
              placeholder={
                mode === 'link'
                  ? 'https://example.com'
                  : 'https://.../media.mp4'
              }
              className="w-full px-3 py-2 rounded-xl bg-[#111620] border border-white/15 text-white focus:border-[#C43E1C] outline-none"
            />
          </div>

          {mode !== 'link' && (
            <div className="pt-1">
              <label className="block text-slate-400 mb-1 font-medium">
                Hoặc chọn tệp từ thiết bị
              </label>
              <input
                type="file"
                accept={mode === 'video' ? 'video/*' : 'audio/*'}
                onChange={handleFilePick}
                className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-white/10 file:text-white hover:file:bg-white/20"
              />
            </div>
          )}
        </div>

        <div className="px-4 py-3 border-t border-white/10 bg-[#141923] flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 rounded-xl bg-white/10 hover:bg-white/15 text-xs font-medium text-slate-200"
          >
            Hủy
          </button>
          <button
            type="button"
            disabled={!urlInput.trim()}
            onClick={handleSubmit}
            className="px-4 py-1.5 rounded-xl bg-[#C43E1C] hover:bg-[#a83316] disabled:opacity-40 text-xs font-semibold text-white"
          >
            Chèn vào Slide
          </button>
        </div>
      </div>
    </div>
  );
}

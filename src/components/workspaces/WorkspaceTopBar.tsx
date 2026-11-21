import type { ReactNode } from 'react';
import {
  ArrowLeft,
  BookOpen,
  FileSpreadsheet,
  FileText,
  Presentation,
  Share2,
} from 'lucide-react';
import type { DocumentItem, DocumentType } from '../../types/document';

export interface WorkspaceTopBarProps {
  document: DocumentItem;
  theme?: 'light' | 'dark';
  onBackHome: () => void;
  onShare: () => void;
  leftExtra?: ReactNode;
  toolbar?: ReactNode;
}

interface FormatVisualConfig {
  label: string;
  Icon: typeof Presentation;
  iconBoxLight: string;
  iconBoxDark: string;
  badgeLight: string;
  badgeDark: string;
}

const FORMAT_VISUALS: Record<DocumentType, FormatVisualConfig> = {
  pptx: {
    label: 'PPTX',
    Icon: Presentation,
    iconBoxLight: 'bg-orange-50 text-orange-600 border-orange-200/80',
    iconBoxDark: 'bg-orange-500/15 text-orange-300 border-orange-400/25',
    badgeLight: 'bg-orange-50 text-orange-700 border-orange-200/80',
    badgeDark: 'bg-orange-500/15 text-orange-300 border-orange-400/25',
  },
  docx: {
    label: 'DOCX',
    Icon: FileText,
    iconBoxLight: 'bg-blue-50 text-blue-600 border-blue-200/80',
    iconBoxDark: 'bg-blue-500/15 text-blue-300 border-blue-400/25',
    badgeLight: 'bg-blue-50 text-blue-700 border-blue-200/80',
    badgeDark: 'bg-blue-500/15 text-blue-300 border-blue-400/25',
  },
  xlsx: {
    label: 'XLSX',
    Icon: FileSpreadsheet,
    iconBoxLight: 'bg-green-50 text-green-700 border-green-200/80',
    iconBoxDark: 'bg-emerald-500/15 text-emerald-300 border-emerald-400/25',
    badgeLight: 'bg-green-50 text-green-800 border-green-200/80',
    badgeDark: 'bg-emerald-500/15 text-emerald-300 border-emerald-400/25',
  },
  pdf: {
    label: 'PDF',
    Icon: BookOpen,
    iconBoxLight: 'bg-pink-50 text-pink-600 border-pink-200/80',
    iconBoxDark: 'bg-pink-500/15 text-pink-300 border-pink-400/25',
    badgeLight: 'bg-pink-50 text-pink-700 border-pink-200/80',
    badgeDark: 'bg-pink-500/15 text-pink-300 border-pink-400/25',
  },
};

export function WorkspaceTopBar({
  document,
  theme = 'light',
  onBackHome,
  onShare,
  leftExtra,
  toolbar,
}: WorkspaceTopBarProps) {
  const isDark = theme === 'dark';
  const visual = FORMAT_VISUALS[document.type];
  const FormatIcon = visual.Icon;

  return (
    <header
      className={`h-14 shrink-0 w-full px-2.5 sm:px-4 flex items-center justify-between gap-2 sm:gap-3 border-b transition-colors duration-200 z-20 ${
        isDark
          ? 'bg-[#141923] border-white/10 text-slate-100'
          : 'bg-[#FFFFFF] border-[#EAE6DF] text-slate-900'
      }`}
    >
      {/* Left Cluster: Back Home + Optional Drawer Trigger + Format Icon + File Title + Badge */}
      <div className="flex-1 min-w-0 flex items-center gap-1.5 sm:gap-2">
        <button
          type="button"
          onClick={onBackHome}
          aria-label="Quay lại Trang chủ"
          title="Quay lại Trang chủ"
          className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[38px] min-w-[38px] sm:min-w-0 px-2 sm:px-2.5 rounded-xl text-xs font-medium border transition-colors ${
            isDark
              ? 'bg-white/[0.05] hover:bg-white/[0.12] border-white/10 text-slate-200 hover:text-white'
              : 'bg-[#FAF8F5] hover:bg-[#F2EFEA] border-[#E7E2DA] text-slate-700 hover:text-slate-950'
          }`}
        >
          <ArrowLeft className="h-4 w-4 shrink-0" />
          <span className="hidden sm:inline whitespace-nowrap">Trang chủ</span>
        </button>

        {leftExtra && (
          <div className="shrink-0 flex items-center">{leftExtra}</div>
        )}

        <span
          aria-hidden="true"
          className={`hidden sm:block h-5 w-px shrink-0 ${
            isDark ? 'bg-white/10' : 'bg-[#EAE6DF]'
          }`}
        />

        <div className="min-w-0 flex items-center gap-2">
          <span
            aria-hidden="true"
            className={`hidden sm:flex shrink-0 h-8 w-8 rounded-lg border items-center justify-center ${
              isDark ? visual.iconBoxDark : visual.iconBoxLight
            }`}
          >
            <FormatIcon className="h-4 w-4" />
          </span>

          <h1
            title={document.title}
            className={`truncate text-[13.5px] sm:text-sm font-semibold tracking-tight ${
              isDark ? 'text-white' : 'text-slate-900'
            }`}
          >
            {document.title}
          </h1>

          <span
            className={`hidden sm:inline-flex shrink-0 items-center px-1.5 py-0.5 rounded-md border text-[10px] font-bold tracking-wider ${
              isDark ? visual.badgeDark : visual.badgeLight
            }`}
          >
            {visual.label}
          </span>
        </div>
      </div>

      {/* Right Cluster: Format-Specific Toolbar + Share Button */}
      <div className="flex items-center gap-1.5 sm:gap-2 min-w-0 max-w-[58%] sm:max-w-[68%]">
        {toolbar && (
          <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-0.5 min-w-0">
            {toolbar}
          </div>
        )}

        <button
          type="button"
          onClick={onShare}
          aria-label="Chia sẻ tài liệu"
          title="Chia sẻ tài liệu"
          className={`shrink-0 inline-flex items-center justify-center gap-1.5 min-h-[38px] sm:min-h-[36px] px-2.5 sm:px-3 rounded-xl text-xs font-semibold border transition-all whitespace-nowrap ${
            isDark
              ? 'bg-white/[0.1] hover:bg-white/[0.18] border-white/15 text-white'
              : 'bg-indigo-50 hover:bg-indigo-100 border-indigo-200/80 text-indigo-700'
          }`}
        >
          <Share2 className="h-3.5 w-3.5 shrink-0" />
          <span className="hidden sm:inline">Chia sẻ</span>
        </button>
      </div>
    </header>
  );
}

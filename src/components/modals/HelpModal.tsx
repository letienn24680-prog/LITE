import { useEffect, useState } from 'react';
import {
  BookOpen,
  Cloud,
  Database,
  FileSpreadsheet,
  FileText,
  HelpCircle,
  Keyboard,
  Laptop,
  Presentation,
  QrCode,
  Smartphone,
  Sparkles,
  Upload,
  X,
} from 'lucide-react';

export interface HelpModalProps {
  isOpen: boolean;
  onClose: () => void;
}

type HelpSectionFilter = 'all' | 'upload' | 'shortcuts' | 'qr' | 'sync';

interface HelpFilterTab {
  key: HelpSectionFilter;
  label: string;
}

const HELP_TABS: readonly HelpFilterTab[] = [
  { key: 'all', label: 'Tất cả' },
  { key: 'upload', label: 'Tải lên & Tạo' },
  { key: 'shortcuts', label: 'Phím tắt' },
  { key: 'qr', label: 'Quét QR' },
  { key: 'sync', label: 'Đồng bộ' },
];

export function HelpModal({ isOpen, onClose }: HelpModalProps) {
  const [activeTab, setActiveTab] = useState<HelpSectionFilter>('all');

  useEffect(() => {
    if (!isOpen) {
      return;
    }
    setActiveTab('all');
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [isOpen, onClose]);

  if (!isOpen) {
    return null;
  }

  const showSection = (key: Exclude<HelpSectionFilter, 'all'>) =>
    activeTab === 'all' || activeTab === key;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="help-modal-title"
      className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm pb-safe"
    >
      <div
        onClick={(event) => {
          if (event.target === event.currentTarget) {
            onClose();
          }
        }}
        className="h-full w-full flex items-center justify-center p-3 sm:p-6"
      >
        <div className="w-full max-w-3xl max-h-full rounded-2xl sm:rounded-3xl bg-white text-slate-900 border border-[#E6E0D6] shadow-2xl overflow-hidden flex flex-col">
          {/* Header */}
          <div className="shrink-0 px-4 sm:px-6 py-3.5 sm:py-4 bg-[#FAF8F5] border-b border-[#EAE5DC] flex items-center justify-between gap-3">
            <div className="flex items-center gap-3 min-w-0">
              <span className="h-9 w-9 shrink-0 rounded-xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-[#4F46E5]">
                <HelpCircle className="h-4 w-4" />
              </span>
              <div className="min-w-0">
                <h2
                  id="help-modal-title"
                  className="text-sm font-semibold text-slate-900 truncate"
                >
                  Hướng dẫn sử dụng LITE Workspace
                </h2>
                <p className="text-xs text-slate-500 truncate mt-0.5">
                  Tài liệu, phím tắt chỉnh sửa, điều khiển QR và đồng bộ đa
                  thiết bị
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng trợ giúp"
              className="shrink-0 min-h-[40px] min-w-[40px] sm:min-h-[36px] sm:min-w-[36px] rounded-xl bg-white hover:bg-slate-100 border border-[#E6E0D6] text-slate-600 hover:text-slate-900 flex items-center justify-center transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* Filter Pills */}
          <div className="shrink-0 px-4 sm:px-6 py-2.5 bg-white border-b border-[#F0ECE4] flex items-center gap-1.5 overflow-x-auto no-scrollbar">
            {HELP_TABS.map((tab) => {
              const isSelected = activeTab === tab.key;
              return (
                <button
                  key={tab.key}
                  type="button"
                  onClick={() => setActiveTab(tab.key)}
                  aria-pressed={isSelected}
                  className={`shrink-0 min-h-[34px] px-3 py-1 rounded-xl text-xs font-semibold transition-all whitespace-nowrap ${
                    isSelected
                      ? 'bg-[#4F46E5] text-white shadow-sm shadow-indigo-600/20'
                      : 'bg-[#FAF8F5] text-slate-600 hover:text-slate-900 hover:bg-[#F2EDE4] border border-[#E6E0D6]'
                  }`}
                >
                  {tab.label}
                </button>
              );
            })}
          </div>

          {/* Content */}
          <div className="flex-1 min-h-0 overflow-y-auto">
            <div className="p-4 sm:p-6 space-y-5">
              {/* 1. Upload & Create */}
              {showSection('upload') && (
                <section className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 sm:p-5 space-y-3">
                  <div className="flex items-center gap-2.5">
                    <span className="h-8 w-8 rounded-xl bg-indigo-50 border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                      <Upload className="h-4 w-4" />
                    </span>
                    <div>
                      <h3 className="text-sm font-semibold text-slate-900">
                        1. Tải lên &amp; Tạo mới tệp
                      </h3>
                      <p className="text-xs text-slate-500">
                        Xử lý tài liệu thật 100% ngay trên trình duyệt
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs text-slate-700 leading-relaxed">
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-1.5">
                      <p className="font-semibold text-slate-900">
                        Kéo thả hoặc bấm &ldquo;+ Tải lên&rdquo;
                      </p>
                      <p>
                        Hỗ trợ trực tiếp các định dạng:{' '}
                        <span className="font-mono font-semibold text-slate-900">
                          .pptx, .docx, .xlsx, .xls, .csv, .pdf
                        </span>
                        . Giải mã nội dung thật 100% trên trình duyệt, không gửi
                        lên máy chủ bên thứ ba.
                      </p>
                    </div>

                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-1.5">
                      <p className="font-semibold text-slate-900">
                        Chỉnh sửa &amp; Tạo mới trực tiếp
                      </p>
                      <p>
                        Tạo slide mới hoặc nhân bản slide trong{' '}
                        <strong>Slide</strong>, thêm sheet mới trong{' '}
                        <strong>Excel</strong>, soạn thảo &amp; xuất tệp{' '}
                        <strong>Word</strong> định dạng chuẩn ngay trên trình
                        duyệt.
                      </p>
                    </div>
                  </div>
                </section>
              )}

              {/* 2. Keyboard Shortcuts */}
              {showSection('shortcuts') && (
                <section className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 sm:p-5 space-y-3.5">
                  <div className="flex items-center gap-2.5">
                    <span className="h-8 w-8 rounded-xl bg-indigo-50 border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                      <Keyboard className="h-4 w-4" />
                    </span>
                    <div>
                      <h3 className="text-sm font-semibold text-slate-900">
                        2. Phím tắt chỉnh sửa nhanh
                      </h3>
                      <p className="text-xs text-slate-500">
                        Tăng tốc thao tác trong Word, Excel và Slide
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    {/* Word */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-2.5">
                      <div className="flex items-center gap-2 text-xs font-semibold text-blue-700">
                        <FileText className="h-4 w-4 shrink-0" />
                        <span>Word (.docx)</span>
                      </div>
                      <ul className="space-y-2 text-xs text-slate-700">
                        <li className="flex items-start justify-between gap-2">
                          <span>Đậm / Nghiêng / Gạch chân</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Ctrl/Cmd + B/I/U
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Lưu tài liệu</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Ctrl/Cmd + S
                          </kbd>
                        </li>
                        <li className="text-[11.5px] text-slate-500 leading-relaxed pt-1 border-t border-slate-100">
                          Tự động cập nhật mục lục từ{' '}
                          <strong className="text-slate-700">
                            Tiêu đề 1–3
                          </strong>
                          .
                        </li>
                      </ul>
                    </div>

                    {/* Excel */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-2.5">
                      <div className="flex items-center gap-2 text-xs font-semibold text-emerald-700">
                        <FileSpreadsheet className="h-4 w-4 shrink-0" />
                        <span>Excel (.xlsx)</span>
                      </div>
                      <ul className="space-y-2 text-xs text-slate-700">
                        <li className="flex items-start justify-between gap-2">
                          <span>Sửa ô trực tiếp</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Đúp / Enter / F2
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Sang ô kế tiếp</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Tab
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Chọn vùng SUM / AVG</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Shift + Click
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Nhập công thức &amp; Lưu</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            =SUM(A1:A5) · Ctrl/Cmd+S
                          </kbd>
                        </li>
                      </ul>
                    </div>

                    {/* Slide */}
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-2.5">
                      <div className="flex items-center gap-2 text-xs font-semibold text-orange-700">
                        <Presentation className="h-4 w-4 shrink-0" />
                        <span>Slide (.pptx)</span>
                      </div>
                      <ul className="space-y-2 text-xs text-slate-700">
                        <li className="flex items-start justify-between gap-2">
                          <span>Chuyển slide nhanh</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            ← / →
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Chiếu điểm đỏ</span>
                          <span className="shrink-0 rounded-md border border-indigo-200 bg-indigo-50 px-1.5 py-0.5 text-[10.5px] font-semibold text-[#4F46E5]">
                            Nút Bút Laser
                          </span>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Tắt Laser / Đóng khung</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Esc
                          </kbd>
                        </li>
                        <li className="flex items-start justify-between gap-2">
                          <span>Lưu bài thuyết trình</span>
                          <kbd className="shrink-0 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-mono text-[10.5px] font-semibold text-slate-800">
                            Ctrl/Cmd + S
                          </kbd>
                        </li>
                      </ul>
                    </div>
                  </div>
                </section>
              )}

              {/* 3. QR Remote Control */}
              {showSection('qr') && (
                <section className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 sm:p-5 space-y-3">
                  <div className="flex items-center gap-2.5">
                    <span className="h-8 w-8 rounded-xl bg-indigo-50 border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                      <QrCode className="h-4 w-4" />
                    </span>
                    <div>
                      <h3 className="text-sm font-semibold text-slate-900">
                        3. Quét QR điều khiển từ xa
                      </h3>
                      <p className="text-xs text-slate-500">
                        Biến điện thoại thành remote chuyển slide khi thuyết
                        trình
                      </p>
                    </div>
                  </div>

                  <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div className="space-y-1.5 text-xs text-slate-700 leading-relaxed">
                      <p>
                        Khi đang mở tệp trình chiếu trong{' '}
                        <strong className="text-slate-900">
                          SlideWorkspace
                        </strong>
                        , bấm nút{' '}
                        <span className="inline-flex items-center gap-1 rounded-md border border-[#E2DCD2] bg-[#FAF8F5] px-1.5 py-0.5 font-semibold text-slate-900">
                          <QrCode className="h-3 w-3 text-[#4F46E5]" />
                          Điều khiển QR
                        </span>{' '}
                        trên thanh công cụ (hoặc trong chế độ Trình chiếu).
                      </p>
                      <p className="text-slate-500">
                        Mở camera điện thoại quét mã QR hoặc nhập mã ghép đôi
                        phòng dạng{' '}
                        <code className="font-mono font-semibold text-[#4F46E5] bg-indigo-50 px-1.5 py-0.5 rounded">
                          LITE-XXXX
                        </code>{' '}
                        để điều khiển chuyển trang từ xa.
                      </p>
                    </div>
                    <div className="shrink-0 flex items-center gap-2 rounded-xl bg-indigo-50/70 border border-indigo-100 px-3 py-2 text-xs font-semibold text-[#4F46E5]">
                      <Smartphone className="h-4 w-4" />
                      <span>LITE-XXXX</span>
                    </div>
                  </div>
                </section>
              )}

              {/* 4. Multi-device Cloud Sync vs Local Mode */}
              {showSection('sync') && (
                <section className="rounded-2xl border border-[#EAE5DC] bg-[#FAF8F5] p-4 sm:p-5 space-y-3">
                  <div className="flex items-center gap-2.5">
                    <span className="h-8 w-8 rounded-xl bg-indigo-50 border border-indigo-100 text-[#4F46E5] flex items-center justify-center shrink-0">
                      <Cloud className="h-4 w-4" />
                    </span>
                    <div>
                      <h3 className="text-sm font-semibold text-slate-900">
                        4. Cơ chế đồng bộ đa thiết bị
                      </h3>
                      <p className="text-xs text-slate-500">
                        So sánh Chế độ cục bộ và Tài khoản đám mây Supabase
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div className="rounded-xl bg-white border border-[#EAE5DC] p-3.5 space-y-2">
                      <div className="flex items-center gap-2 text-xs font-semibold text-slate-900">
                        <Database className="h-4 w-4 text-slate-500 shrink-0" />
                        <span>Chế độ cục bộ (Local Mode)</span>
                      </div>
                      <ul className="space-y-1.5 text-xs text-slate-600 leading-relaxed">
                        <li className="flex items-start gap-2">
                          <BookOpen className="h-3.5 w-3.5 text-slate-400 shrink-0 mt-0.5" />
                          <span>
                            Lưu dữ liệu vào{' '}
                            <code className="font-mono text-slate-800">
                              localStorage
                            </code>{' '}
                            trên duy nhất 1 máy / 1 trình duyệt.
                          </span>
                        </li>
                        <li className="flex items-start gap-2">
                          <Sparkles className="h-3.5 w-3.5 text-slate-400 shrink-0 mt-0.5" />
                          <span>
                            Dùng ngay lập tức không cần đăng nhập; phù hợp thao
                            tác nhanh trên một thiết bị.
                          </span>
                        </li>
                      </ul>
                    </div>

                    <div className="rounded-xl bg-white border border-indigo-200 p-3.5 space-y-2">
                      <div className="flex items-center gap-2 text-xs font-semibold text-[#4F46E5]">
                        <Cloud className="h-4 w-4 shrink-0" />
                        <span>Đăng nhập tài khoản Supabase</span>
                      </div>
                      <ul className="space-y-1.5 text-xs text-slate-700 leading-relaxed">
                        <li className="flex items-start gap-2">
                          <Laptop className="h-3.5 w-3.5 text-[#4F46E5] shrink-0 mt-0.5" />
                          <span>
                            Tải tệp gốc lên Storage{' '}
                            <code className="font-mono font-semibold text-slate-900">
                              user-files
                            </code>{' '}
                            an toàn theo từng tài khoản.
                          </span>
                        </li>
                        <li className="flex items-start gap-2">
                          <Smartphone className="h-3.5 w-3.5 text-[#4F46E5] shrink-0 mt-0.5" />
                          <span>
                            Đồng bộ nội dung đã chỉnh sửa (
                            <code className="font-mono text-slate-900">
                              parsed_content
                            </code>{' '}
                            +{' '}
                            <code className="font-mono text-slate-900">
                              last_accessed_at
                            </code>
                            ) xuyên suốt giữa máy tính &amp; điện thoại.
                          </span>
                        </li>
                      </ul>
                    </div>
                  </div>
                </section>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

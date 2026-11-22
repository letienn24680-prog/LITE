import { useState } from 'react';
import {
  ArrowLeft,
  CheckCircle2,
  FileText,
  FolderLock,
  HardDrive,
  Scale,
  ShieldCheck,
} from 'lucide-react';

export type LegalPageKind = 'privacy' | 'terms';

export interface LegalPagesViewProps {
  page: LegalPageKind;
  onNavigatePage?: (page: LegalPageKind) => void;
  onBackToApp: () => void;
}

export function LegalPagesView({
  page,
  onNavigatePage,
  onBackToApp,
}: LegalPagesViewProps) {
  const [internalPage, setInternalPage] = useState<LegalPageKind>(page);
  const activePage = onNavigatePage ? page : internalPage;

  const handleSwitchPage = (target: LegalPageKind) => {
    if (onNavigatePage) {
      onNavigatePage(target);
      return;
    }
    setInternalPage(target);
    if (typeof window !== 'undefined') {
      window.history.pushState({}, '', `/${target}`);
      window.dispatchEvent(new PopStateEvent('popstate'));
    }
  };

  return (
    <div className="min-h-[100dvh] w-full bg-[#FAF8F5] text-slate-900 flex flex-col">
      {/* Sticky Header */}
      <header className="sticky top-0 z-20 shrink-0 border-b border-[#EAE6DF] bg-white/95 backdrop-blur-md">
        <div className="mx-auto flex max-w-4xl flex-wrap items-center justify-between gap-3 px-4 py-3 sm:px-6">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={onBackToApp}
              className="inline-flex items-center gap-1.5 min-h-[38px] px-3 rounded-xl border border-[#E6E0D6] bg-[#FAF8F5] hover:bg-[#F2EDE4] text-xs font-semibold text-slate-700 hover:text-slate-950 transition-colors"
            >
              <ArrowLeft className="h-4 w-4 shrink-0" />
              <span>Quay lại LITE</span>
            </button>

            <div className="flex items-center gap-2 min-w-0">
              <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-indigo-600 text-xs font-bold text-white shadow-sm shadow-indigo-600/20">
                L
              </span>
              <span className="text-sm font-bold tracking-tight text-slate-900 truncate">
                LITE Workspace
              </span>
            </div>
          </div>

          <nav
            aria-label="Trang pháp lý"
            className="inline-flex items-center gap-1 rounded-xl border border-[#EAE5DC] bg-[#FAF8F5] p-1"
          >
            <button
              type="button"
              onClick={() => handleSwitchPage('privacy')}
              aria-current={activePage === 'privacy' ? 'page' : undefined}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all ${
                activePage === 'privacy'
                  ? 'bg-white text-indigo-700 shadow-sm ring-1 ring-[#E6E0D6]'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <ShieldCheck className="h-3.5 w-3.5 shrink-0" />
              <span>Chính sách bảo mật</span>
            </button>
            <button
              type="button"
              onClick={() => handleSwitchPage('terms')}
              aria-current={activePage === 'terms' ? 'page' : undefined}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-semibold transition-all ${
                activePage === 'terms'
                  ? 'bg-white text-indigo-700 shadow-sm ring-1 ring-[#E6E0D6]'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Scale className="h-3.5 w-3.5 shrink-0" />
              <span>Điều khoản sử dụng</span>
            </button>
          </nav>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 px-4 py-6 sm:px-6 sm:py-10">
        <article className="mx-auto max-w-3xl rounded-3xl border border-[#E6E0D6] bg-white p-6 sm:p-10 shadow-sm space-y-6">
          {activePage === 'privacy' ? (
            <>
              <header className="border-b border-[#F0ECE4] pb-5 space-y-2">
                <div className="inline-flex items-center gap-1.5 rounded-full bg-indigo-50 border border-indigo-100 px-3 py-1 text-[11px] font-semibold text-indigo-700">
                  <ShieldCheck className="h-3.5 w-3.5" />
                  <span>Cập nhật mới nhất · Năm 2026</span>
                </div>
                <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900">
                  Chính sách bảo mật — LITE Workspace
                </h1>
                <p className="text-sm text-slate-600 leading-relaxed">
                  LITE Workspace được thiết kế theo nguyên tắc tối giản, ưu tiên
                  xử lý ngay trên trình duyệt/thiết bị của bạn và minh bạch tuyệt
                  đối về quyền truy cập dữ liệu.
                </p>
              </header>

              {/* Highlighted Google Drive Transparency Commitment */}
              <section className="rounded-2xl border border-indigo-200 bg-indigo-50/60 p-4 sm:p-5 flex items-start gap-3.5">
                <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-indigo-600 text-white">
                  <HardDrive className="h-4 w-4" />
                </span>
                <div className="space-y-1.5 text-sm leading-relaxed text-slate-800">
                  <h2 className="font-semibold text-indigo-950">
                    Cam kết minh bạch về quyền truy cập Google Drive
                  </h2>
                  <p className="text-[13.5px] text-slate-700">
                    <strong>LITE Workspace</strong> chỉ truy cập các tệp do người
                    dùng <strong>chủ động mở</strong> hoặc{' '}
                    <strong>tạo trong thư mục <code>LITE Workspace</code></strong>{' '}
                    trên Google Drive cá nhân, hoàn toàn{' '}
                    <strong>
                      không lưu trữ hay thu thập dữ liệu Drive của người dùng lên
                      máy chủ trung gian
                    </strong>
                    . Mọi thao tác đọc, chỉnh sửa và xuất tệp diễn ra trực tiếp
                    giữa trình duyệt/ứng dụng của bạn và API chính thức của
                    Google Drive.
                  </p>
                </div>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900 flex items-center gap-2">
                  <FolderLock className="h-4 w-4 text-indigo-600 shrink-0" />
                  <span>1. Dữ liệu tài khoản &amp; phạm vi sử dụng</span>
                </h2>
                <ul className="space-y-2 text-sm text-slate-700 leading-relaxed list-disc pl-5">
                  <li>
                    Khi đăng nhập bằng Email hoặc Google OAuth, hệ thống chỉ sử
                    dụng mã định danh tài khoản và địa chỉ email để phân tách
                    không gian lưu trữ riêng tư của bạn.
                  </li>
                  <li>
                    Mật khẩu và phiên xác thực được quản lý trực tiếp bởi hạ
                    tầng xác thực Supabase Auth; LITE Workspace không lưu trữ mật
                    khẩu dạng văn bản thuần.
                  </li>
                  <li>
                    Bản nháp ngoại tuyến, lịch sử phiên bản cục bộ (IndexedDB)
                    và mã thông báo Drive (<code>localStorage</code>) được cô lập
                    riêng biệt theo từng tài khoản (<code>userId</code>) và tự
                    động dọn sạch khi bạn đăng xuất.
                  </li>
                </ul>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900 flex items-center gap-2">
                  <FileText className="h-4 w-4 text-indigo-600 shrink-0" />
                  <span>2. Xử lý tài liệu &amp; chuyển đổi PDF cục bộ</span>
                </h2>
                <p className="text-sm text-slate-700 leading-relaxed">
                  Các chức năng đọc xem trước <code>.docx</code>,{' '}
                  <code>.xlsx</code>, <code>.pptx</code>, <code>.pdf</code>, xuất
                  sang PDF, gộp PDF và tách PDF đều vận hành{' '}
                  <strong>100% trên thiết bị của người dùng</strong>. Tệp không
                  bị gửi sang bất kỳ dịch vụ chuyển đổi bên thứ ba nào.
                </p>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900 flex items-center gap-2">
                  <CheckCircle2 className="h-4 w-4 text-indigo-600 shrink-0" />
                  <span>3. Chia sẻ liên kết &amp; quyền kiểm soát</span>
                </h2>
                <ul className="space-y-2 text-sm text-slate-700 leading-relaxed list-disc pl-5">
                  <li>
                    Tài liệu mặc định ở chế độ{' '}
                    <strong>Hạn chế (Chỉ mình tôi)</strong>. Chỉ khi bạn chủ
                    động bật liên kết chia sẻ (<code>Chỉ xem</code> hoặc{' '}
                    <code>Được phép chỉnh sửa</code>), người có liên kết mới
                    truy cập được nội dung tương ứng.
                  </li>
                  <li>
                    Bạn có thể thu hồi quyền truy cập liên kết chia sẻ, xóa tài
                    liệu trên đám mây, hoặc ngắt liên kết Google Drive trong phần
                    quản lý tài khoản Google bất kỳ lúc nào.
                  </li>
                </ul>
              </section>
            </>
          ) : (
            <>
              <header className="border-b border-[#F0ECE4] pb-5 space-y-2">
                <div className="inline-flex items-center gap-1.5 rounded-full bg-indigo-50 border border-indigo-100 px-3 py-1 text-[11px] font-semibold text-indigo-700">
                  <Scale className="h-3.5 w-3.5" />
                  <span>Cập nhật mới nhất · Năm 2026</span>
                </div>
                <h1 className="text-2xl sm:text-3xl font-bold tracking-tight text-slate-900">
                  Điều khoản sử dụng — LITE Workspace
                </h1>
                <p className="text-sm text-slate-600 leading-relaxed">
                  Khi truy cập và sử dụng ứng dụng web, ứng dụng máy tính hoặc
                  ứng dụng di động LITE Workspace, bạn đồng ý tuân thủ các điều
                  khoản cơ bản dưới đây.
                </p>
              </header>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900">
                  1. Phạm vi dịch vụ &amp; Quyền sử dụng
                </h2>
                <p className="text-sm text-slate-700 leading-relaxed">
                  LITE Workspace cung cấp không gian lưu trữ, mở xem, soạn thảo,
                  trình chiếu và chuyển đổi tài liệu văn phòng (<code>DOCX</code>
                  , <code>XLSX</code>, <code>PPTX</code>, <code>PDF</code>) hỗ
                  trợ làm việc ngoại tuyến, đồng bộ đám mây và cộng tác thời gian
                  thực.
                </p>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900">
                  2. Quyền sở hữu nội dung của người dùng
                </h2>
                <p className="text-sm text-slate-700 leading-relaxed">
                  Bạn giữ toàn bộ quyền sở hữu trí tuệ đối với mọi tệp, văn bản,
                  bảng tính, slide và dữ liệu do bạn tải lên hoặc tạo mới trên
                  LITE Workspace. Chúng tôi không yêu cầu bất kỳ quyền khai thác
                  thương mại nào đối với tài liệu của bạn.
                </p>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900">
                  3. Trách nhiệm khi chia sẻ &amp; cộng tác
                </h2>
                <ul className="space-y-2 text-sm text-slate-700 leading-relaxed list-disc pl-5">
                  <li>
                    Bạn chịu trách nhiệm quản lý quyền truy cập khi gửi liên kết
                    chia sẻ (<code>?share=...&amp;mode=view|edit</code> hoặc liên
                    kết mở nhanh trong ứng dụng <code>liteapp://open</code>).
                  </li>
                  <li>
                    Không sử dụng nền tảng để lưu trữ hoặc phát tán nội dung vi
                    phạm pháp luật, mã độc hoặc xâm phạm quyền riêng tư của bên
                    thứ ba.
                  </li>
                </ul>
              </section>

              <section className="space-y-2.5">
                <h2 className="text-base font-semibold text-slate-900">
                  4. Sao lưu &amp; giới hạn trách nhiệm
                </h2>
                <p className="text-sm text-slate-700 leading-relaxed">
                  LITE Workspace tích hợp lưu nháp cục bộ, hàng đợi đồng bộ khi
                  mất mạng và lịch sử 10 phiên bản gần nhất nhằm hạn chế mất mát
                  dữ liệu. Tuy nhiên, người dùng vẫn nên sao lưu định kỳ các tài
                  liệu quan trọng về thiết bị hoặc Google Drive cá nhân.
                </p>
              </section>
            </>
          )}

          {/* Card Footer */}
          <footer className="pt-5 border-t border-[#F0ECE4] flex flex-wrap items-center justify-between gap-3">
            <button
              type="button"
              onClick={() =>
                handleSwitchPage(activePage === 'privacy' ? 'terms' : 'privacy')
              }
              className="text-xs font-semibold text-indigo-600 hover:text-indigo-700 underline underline-offset-2"
            >
              {activePage === 'privacy'
                ? 'Xem Điều khoản sử dụng →'
                : '← Xem Chính sách bảo mật'}
            </button>

            <button
              type="button"
              onClick={onBackToApp}
              className="inline-flex items-center gap-1.5 min-h-[40px] px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-xs font-semibold text-white transition-colors"
            >
              <ArrowLeft className="h-3.5 w-3.5" />
              <span>Quay lại LITE Workspace</span>
            </button>
          </footer>
        </article>
      </main>
    </div>
  );
}

export default LegalPagesView;

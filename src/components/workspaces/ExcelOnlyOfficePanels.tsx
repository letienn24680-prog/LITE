import { useEffect, useMemo, useRef, useState } from 'react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import {
  AlertTriangle,
  Check,
  Copy,
  FilePlus2,
  FunctionSquare,
  Link2,
  MessageSquarePlus,
  PaintBucket,
  Search,
  Sparkles,
  Trash2,
  Wand2,
  X,
} from 'lucide-react';
import type {
  ExcelDrawingStroke,
  ExcelFloatingObject,
} from '../../types/document';
import {
  columnIndexToLetter,
  formatCellRef,
} from '../../utils/formulaEvaluator';

/* ============================================================================
 * 1. ExcelFormulaLibraryModal
 * ========================================================================== */

export type FormulaLibraryCategory =
  | 'all'
  | 'autosum'
  | 'financial'
  | 'logic'
  | 'text'
  | 'datetime'
  | 'lookup'
  | 'math';

export interface FormulaLibraryItem {
  name: string;
  category: Exclude<FormulaLibraryCategory, 'all'>;
  syntax: string;
  example: string;
  snippet: string;
  description: string;
}

const FORMULA_CATEGORY_TABS: ReadonlyArray<{
  key: FormulaLibraryCategory;
  label: string;
}> = [
  { key: 'all', label: 'Tất cả' },
  { key: 'autosum', label: 'AutoSum' },
  { key: 'financial', label: 'Tài chính' },
  { key: 'logic', label: 'Logic' },
  { key: 'text', label: 'Văn bản' },
  { key: 'datetime', label: 'Ngày & Giờ' },
  { key: 'lookup', label: 'Tra cứu' },
  { key: 'math', label: 'Toán học' },
];

export const ONLYOFFICE_FORMULA_LIBRARY: readonly FormulaLibraryItem[] = [
  // AutoSum
  {
    name: 'SUM',
    category: 'autosum',
    syntax: 'SUM(số_1, [số_2], ...)',
    example: '=SUM(B2:B10)',
    snippet: '=SUM(',
    description: 'Tính tổng toàn bộ giá trị số trong ô hoặc vùng được chọn.',
  },
  {
    name: 'AVERAGE',
    category: 'autosum',
    syntax: 'AVERAGE(số_1, [số_2], ...)',
    example: '=AVERAGE(C2:C10)',
    snippet: '=AVERAGE(',
    description: 'Tính giá trị trung bình cộng của các giá trị số.',
  },
  {
    name: 'COUNT',
    category: 'autosum',
    syntax: 'COUNT(vùng_1, [vùng_2], ...)',
    example: '=COUNT(A2:A20)',
    snippet: '=COUNT(',
    description: 'Đếm số ô có chứa dữ liệu dạng số trong vùng.',
  },
  {
    name: 'COUNTA',
    category: 'autosum',
    syntax: 'COUNTA(vùng_1, [vùng_2], ...)',
    example: '=COUNTA(A2:A20)',
    snippet: '=COUNTA(',
    description: 'Đếm toàn bộ các ô không trống (gồm cả chữ và số).',
  },
  {
    name: 'MAX',
    category: 'autosum',
    syntax: 'MAX(vùng_1, [vùng_2], ...)',
    example: '=MAX(D2:D20)',
    snippet: '=MAX(',
    description: 'Tìm giá trị số lớn nhất trong vùng dữ liệu.',
  },
  {
    name: 'MIN',
    category: 'autosum',
    syntax: 'MIN(vùng_1, [vùng_2], ...)',
    example: '=MIN(D2:D20)',
    snippet: '=MIN(',
    description: 'Tìm giá trị số nhỏ nhất trong vùng dữ liệu.',
  },

  // Tài chính
  {
    name: 'PMT',
    category: 'financial',
    syntax: 'PMT(lãi_suất, số_kỳ, giá_trị_hiện_tại, [fv], [type])',
    example: '=PMT(0.08/12, 12, -100000000)',
    snippet: '=PMT(',
    description: 'Tính số tiền phải trả định kỳ cho khoản vay với lãi suất cố định.',
  },
  {
    name: 'FV',
    category: 'financial',
    syntax: 'FV(lãi_suất, số_kỳ, khoản_nộp, [pv], [type])',
    example: '=FV(0.06/12, 24, -2000000)',
    snippet: '=FV(',
    description: 'Tính giá trị tương lai của khoản đầu tư tích lũy định kỳ.',
  },
  {
    name: 'PV',
    category: 'financial',
    syntax: 'PV(lãi_suất, số_kỳ, khoản_nộp, [fv], [type])',
    example: '=PV(0.07/12, 12, -5000000)',
    snippet: '=PV(',
    description: 'Tính giá trị hiện tại của dòng tiền trả định kỳ trong tương lai.',
  },

  // Logic
  {
    name: 'IF',
    category: 'logic',
    syntax: 'IF(điều_kiện, giá_trị_đúng, [giá_trị_sai])',
    example: '=IF(D2>=5, "Đạt", "Trượt")',
    snippet: '=IF(',
    description: 'Kiểm tra biểu thức điều kiện và trả kết quả tương ứng.',
  },
  {
    name: 'IFS',
    category: 'logic',
    syntax: 'IFS(đk_1, kq_1, đk_2, kq_2, ...)',
    example: '=IFS(D2>=8, "Giỏi", D2>=6.5, "Khá", TRUE, "Trung bình")',
    snippet: '=IFS(',
    description: 'Kiểm tra tuần tự nhiều điều kiện lồng nhau gọn gàng.',
  },
  {
    name: 'AND',
    category: 'logic',
    syntax: 'AND(đk_1, [đk_2], ...)',
    example: '=AND(B2>0, C2>=5)',
    snippet: '=AND(',
    description: 'Trả về TRUE khi tất cả điều kiện đầu vào đều đúng.',
  },
  {
    name: 'OR',
    category: 'logic',
    syntax: 'OR(đk_1, [đk_2], ...)',
    example: '=OR(B2="VIP", C2>=100)',
    snippet: '=OR(',
    description: 'Trả về TRUE khi có ít nhất một điều kiện đúng.',
  },
  {
    name: 'NOT',
    category: 'logic',
    syntax: 'NOT(điều_kiện)',
    example: '=NOT(B2="Hủy")',
    snippet: '=NOT(',
    description: 'Đảo ngược giá trị logic TRUE thành FALSE và ngược lại.',
  },

  // Văn bản
  {
    name: 'CONCAT',
    category: 'text',
    syntax: 'CONCAT(chuỗi_1, [chuỗi_2], ...)',
    example: '=CONCAT(A2, " - ", B2)',
    snippet: '=CONCAT(',
    description: 'Nối nhiều chuỗi hoặc vùng ô văn bản lại thành một chuỗi.',
  },
  {
    name: 'CONCATENATE',
    category: 'text',
    syntax: 'CONCATENATE(chuỗi_1, chuỗi_2, ...)',
    example: '=CONCATENATE(A2, " ", B2)',
    snippet: '=CONCATENATE(',
    description: 'Ghép nối các đoạn văn bản theo đúng thứ tự liệt kê.',
  },
  {
    name: 'TEXTJOIN',
    category: 'text',
    syntax: 'TEXTJOIN(ký_tự_ngăn, bỏ_ô_trống, vùng_1, ...)',
    example: '=TEXTJOIN(", ", TRUE, A2:A5)',
    snippet: '=TEXTJOIN(',
    description: 'Ghép nhiều ô bằng ký tự phân tách và tùy chọn bỏ qua ô trống.',
  },
  {
    name: 'LEFT',
    category: 'text',
    syntax: 'LEFT(văn_bản, [số_ký_tự])',
    example: '=LEFT(A2, 4)',
    snippet: '=LEFT(',
    description: 'Trích xuất các ký tự tính từ bên trái của chuỗi.',
  },
  {
    name: 'RIGHT',
    category: 'text',
    syntax: 'RIGHT(văn_bản, [số_ký_tự])',
    example: '=RIGHT(A2, 3)',
    snippet: '=RIGHT(',
    description: 'Trích xuất các ký tự tính từ bên phải của chuỗi.',
  },
  {
    name: 'MID',
    category: 'text',
    syntax: 'MID(văn_bản, vị_trí_bắt_đầu, số_ký_tự)',
    example: '=MID(A2, 2, 4)',
    snippet: '=MID(',
    description: 'Cắt đoạn văn bản con bắt đầu tại vị trí chỉ định.',
  },
  {
    name: 'LEN',
    category: 'text',
    syntax: 'LEN(văn_bản)',
    example: '=LEN(A2)',
    snippet: '=LEN(',
    description: 'Đếm tổng số ký tự (bao gồm cả khoảng trắng) trong ô.',
  },
  {
    name: 'TRIM',
    category: 'text',
    syntax: 'TRIM(văn_bản)',
    example: '=TRIM(A2)',
    snippet: '=TRIM(',
    description: 'Loại bỏ khoảng trắng thừa ở đầu, cuối và giữa các từ.',
  },
  {
    name: 'UPPER',
    category: 'text',
    syntax: 'UPPER(văn_bản)',
    example: '=UPPER(A2)',
    snippet: '=UPPER(',
    description: 'Chuyển toàn bộ văn bản thành CHỮ IN HOA.',
  },
  {
    name: 'LOWER',
    category: 'text',
    syntax: 'LOWER(văn_bản)',
    example: '=LOWER(A2)',
    snippet: '=LOWER(',
    description: 'Chuyển toàn bộ văn bản thành chữ thường.',
  },
  {
    name: 'PROPER',
    category: 'text',
    syntax: 'PROPER(văn_bản)',
    example: '=PROPER(A2)',
    snippet: '=PROPER(',
    description: 'Viết Hoa Chữ Cái Đầu Tiên Của Mỗi Từ.',
  },

  // Ngày & Giờ
  {
    name: 'TODAY',
    category: 'datetime',
    syntax: 'TODAY()',
    example: '=TODAY()',
    snippet: '=TODAY()',
    description: 'Trả về ngày hiện tại theo đồng hồ hệ thống.',
  },
  {
    name: 'NOW',
    category: 'datetime',
    syntax: 'NOW()',
    example: '=NOW()',
    snippet: '=NOW()',
    description: 'Trả về ngày và giờ hiện tại theo đồng hồ hệ thống.',
  },
  {
    name: 'DAY',
    category: 'datetime',
    syntax: 'DAY(ngày)',
    example: '=DAY(TODAY())',
    snippet: '=DAY(',
    description: 'Lấy giá trị ngày trong tháng (1–31) từ ngày cho trước.',
  },
  {
    name: 'MONTH',
    category: 'datetime',
    syntax: 'MONTH(ngày)',
    example: '=MONTH(TODAY())',
    snippet: '=MONTH(',
    description: 'Lấy giá trị tháng trong năm (1–12) từ ngày cho trước.',
  },
  {
    name: 'YEAR',
    category: 'datetime',
    syntax: 'YEAR(ngày)',
    example: '=YEAR(TODAY())',
    snippet: '=YEAR(',
    description: 'Lấy giá trị năm (4 chữ số) từ ngày cho trước.',
  },

  // Tra cứu
  {
    name: 'VLOOKUP',
    category: 'lookup',
    syntax: 'VLOOKUP(giá_trị_tìm, vùng_bảng, cột_lấy, [khớp_gần])',
    example: '=VLOOKUP(E2, A2:C10, 3, FALSE)',
    snippet: '=VLOOKUP(',
    description: 'Tra cứu theo cột đầu tiên của bảng và trả giá trị ở cột chỉ định.',
  },
  {
    name: 'HLOOKUP',
    category: 'lookup',
    syntax: 'HLOOKUP(giá_trị_tìm, vùng_bảng, dòng_lấy, [khớp_gần])',
    example: '=HLOOKUP("Q1", A1:E4, 3, FALSE)',
    snippet: '=HLOOKUP(',
    description: 'Tra cứu theo dòng đầu tiên của bảng và trả giá trị ở dòng chỉ định.',
  },
  {
    name: 'INDEX',
    category: 'lookup',
    syntax: 'INDEX(vùng_dữ_liệu, số_dòng, [số_cột])',
    example: '=INDEX(A2:C10, 2, 3)',
    snippet: '=INDEX(',
    description: 'Trả về giá trị tại tọa độ dòng và cột bên trong vùng.',
  },
  {
    name: 'MATCH',
    category: 'lookup',
    syntax: 'MATCH(giá_trị_tìm, vùng_tìm, [kiểu_khớp])',
    example: '=MATCH("Hà Nội", A2:A20, 0)',
    snippet: '=MATCH(',
    description: 'Tìm vị trí thứ tự (bắt đầu từ 1) của giá trị trong vùng.',
  },

  // Toán học
  {
    name: 'COUNTBLANK',
    category: 'math',
    syntax: 'COUNTBLANK(vùng)',
    example: '=COUNTBLANK(A2:C10)',
    snippet: '=COUNTBLANK(',
    description: 'Đếm số lượng ô đang để trống trong một vùng.',
  },
  {
    name: 'COUNTIF',
    category: 'math',
    syntax: 'COUNTIF(vùng, điều_kiện)',
    example: '=COUNTIF(B2:B20, ">10")',
    snippet: '=COUNTIF(',
    description: 'Đếm số ô thỏa mãn điều kiện (vd: ">10", "=Đạt", "*VIP*").',
  },
  {
    name: 'SUMIF',
    category: 'math',
    syntax: 'SUMIF(vùng_điều_kiện, điều_kiện, [vùng_tính_tổng])',
    example: '=SUMIF(A2:A20, "Đã thu", B2:B20)',
    snippet: '=SUMIF(',
    description: 'Cộng các ô tương ứng thỏa mãn điều kiện đặt ra.',
  },
  {
    name: 'AVERAGEIF',
    category: 'math',
    syntax: 'AVERAGEIF(vùng_điều_kiện, điều_kiện, [vùng_tính_tb])',
    example: '=AVERAGEIF(A2:A20, "Nhóm A", C2:C20)',
    snippet: '=AVERAGEIF(',
    description: 'Tính trung bình cộng các ô thỏa mãn điều kiện.',
  },
  {
    name: 'ROUND',
    category: 'math',
    syntax: 'ROUND(số, số_lẻ)',
    example: '=ROUND(B2/3, 2)',
    snippet: '=ROUND(',
    description: 'Làm tròn số tới số chữ số thập phân mong muốn.',
  },
  {
    name: 'ROUNDUP',
    category: 'math',
    syntax: 'ROUNDUP(số, số_lẻ)',
    example: '=ROUNDUP(B2, 0)',
    snippet: '=ROUNDUP(',
    description: 'Làm tròn số theo hướng ra xa số 0.',
  },
  {
    name: 'ROUNDDOWN',
    category: 'math',
    syntax: 'ROUNDDOWN(số, số_lẻ)',
    example: '=ROUNDDOWN(B2, 0)',
    snippet: '=ROUNDDOWN(',
    description: 'Làm tròn số theo hướng tiến gần số 0.',
  },
  {
    name: 'INT',
    category: 'math',
    syntax: 'INT(số)',
    example: '=INT(B2)',
    snippet: '=INT(',
    description: 'Làm tròn xuống số nguyên gần nhất nhỏ hơn hoặc bằng giá trị.',
  },
  {
    name: 'ABS',
    category: 'math',
    syntax: 'ABS(số)',
    example: '=ABS(B2-C2)',
    snippet: '=ABS(',
    description: 'Trả về giá trị tuyệt đối (không âm) của một số.',
  },
];

export interface ExcelFormulaLibraryModalProps {
  isOpen: boolean;
  initialCategory?: FormulaLibraryCategory;
  activeCellRef?: string;
  onClose: () => void;
  onInsertFormula: (formulaSnippet: string) => void;
}

export function ExcelFormulaLibraryModal({
  isOpen,
  initialCategory = 'all',
  activeCellRef = 'A1',
  onClose,
  onInsertFormula,
}: ExcelFormulaLibraryModalProps) {
  const [category, setCategory] =
    useState<FormulaLibraryCategory>(initialCategory);
  const [query, setQuery] = useState('');
  const [selectedName, setSelectedName] = useState<string>('SUM');

  useEffect(() => {
    if (isOpen) {
      setCategory(initialCategory);
      setQuery('');
    }
  }, [initialCategory, isOpen]);

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

  const filteredFormulas = useMemo(() => {
    const q = query.trim().toLowerCase();
    return ONLYOFFICE_FORMULA_LIBRARY.filter((item) => {
      const matchesCat = category === 'all' || item.category === category;
      if (!matchesCat) {
        return false;
      }
      if (!q) {
        return true;
      }
      return (
        item.name.toLowerCase().includes(q) ||
        item.description.toLowerCase().includes(q) ||
        item.syntax.toLowerCase().includes(q)
      );
    });
  }, [category, query]);

  const selectedFormula = useMemo(() => {
    return (
      filteredFormulas.find((item) => item.name === selectedName) ??
      filteredFormulas[0] ??
      null
    );
  }, [filteredFormulas, selectedName]);

  if (!isOpen) {
    return null;
  }

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="excel-formula-library-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/55 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-5">
        <div
          onClick={(event) => event.stopPropagation()}
          className="flex max-h-full w-full max-w-3xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white text-slate-900 shadow-2xl"
        >
          {/* Header */}
          <header className="flex shrink-0 items-center justify-between gap-3 border-b border-slate-200 bg-[#F8FAFC] px-4 py-3 sm:px-5">
            <div className="flex items-center gap-2.5 min-w-0">
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-[#ECFDF5] text-[#107C41] border border-[#107C41]/20">
                <FunctionSquare className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h2
                  id="excel-formula-library-title"
                  className="truncate text-sm sm:text-base font-bold text-slate-900"
                >
                  Thư viện hàm Excel ({ONLYOFFICE_FORMULA_LIBRARY.length} hàm)
                </h2>
                <p className="truncate text-xs text-slate-500">
                  Ô đang chọn: <strong className="text-[#107C41]">{activeCellRef}</strong>
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng thư viện hàm"
              className="flex h-8 w-8 items-center justify-center rounded-xl text-slate-400 hover:bg-slate-200/70 hover:text-slate-700 transition-colors"
            >
              <X className="h-4 w-4" />
            </button>
          </header>

          {/* Search & Category Tabs */}
          <div className="shrink-0 space-y-2.5 border-b border-slate-200 bg-white px-4 py-3 sm:px-5">
            <div className="relative flex items-center">
              <Search className="pointer-events-none absolute left-3 h-4 w-4 text-slate-400" />
              <input
                type="text"
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="Tìm tên hàm hoặc mô tả (vd: SUM, VLOOKUP, tài chính, đếm)..."
                aria-label="Tìm kiếm hàm Excel"
                className="h-9 w-full rounded-xl border border-slate-200 bg-slate-50 pl-9 pr-8 text-xs sm:text-sm text-slate-900 placeholder:text-slate-400 focus:border-[#107C41] focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#107C41]/15"
              />
              {query.length > 0 && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  aria-label="Xóa từ khóa tìm hàm"
                  className="absolute right-2 rounded-lg p-1 text-slate-400 hover:bg-slate-200/60 hover:text-slate-700"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </div>

            <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar pb-0.5">
              {FORMULA_CATEGORY_TABS.map((tab) => {
                const isActive = category === tab.key;
                return (
                  <button
                    key={tab.key}
                    type="button"
                    onClick={() => setCategory(tab.key)}
                    className={`shrink-0 whitespace-nowrap rounded-lg px-2.5 py-1 text-xs font-semibold transition-colors border ${
                      isActive
                        ? 'bg-[#ECFDF5] border-[#107C41] text-[#107C41]'
                        : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    {tab.label}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Body: List + Detail */}
          <div className="min-h-0 flex-1 grid grid-cols-1 md:grid-cols-12 overflow-hidden">
            <div className="md:col-span-6 min-h-0 max-h-[42vh] md:max-h-none overflow-y-auto border-b md:border-b-0 md:border-r border-slate-200 p-2.5 space-y-1">
              {filteredFormulas.length === 0 ? (
                <div className="py-10 text-center text-xs text-slate-500">
                  Không tìm thấy hàm phù hợp với &ldquo;{query}&rdquo;.
                </div>
              ) : (
                filteredFormulas.map((item) => {
                  const isSelected = selectedFormula?.name === item.name;
                  return (
                    <button
                      key={item.name}
                      type="button"
                      onClick={() => setSelectedName(item.name)}
                      onDoubleClick={() => {
                        onInsertFormula(item.snippet);
                        onClose();
                      }}
                      className={`w-full flex items-start justify-between gap-2 rounded-xl px-3 py-2 text-left transition-colors border ${
                        isSelected
                          ? 'bg-[#ECFDF5] border-[#107C41]/40 text-[#0B5E31]'
                          : 'bg-white border-transparent hover:bg-slate-50 text-slate-800'
                      }`}
                    >
                      <div className="min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="font-mono text-xs font-bold text-[#107C41]">
                            {item.name}
                          </span>
                        </div>
                        <p className="mt-0.5 line-clamp-1 text-[11px] text-slate-500">
                          {item.description}
                        </p>
                      </div>
                    </button>
                  );
                })
              )}
            </div>

            <div className="md:col-span-6 flex flex-col justify-between bg-[#F8FAFC] p-4 sm:p-5 overflow-y-auto">
              {selectedFormula ? (
                <div className="space-y-3.5">
                  <div>
                    <span className="inline-flex items-center rounded-md bg-[#ECFDF5] px-2 py-0.5 text-[10.5px] font-bold text-[#107C41] border border-[#107C41]/20">
                      {
                        FORMULA_CATEGORY_TABS.find(
                          (t) => t.key === selectedFormula.category
                        )?.label
                      }
                    </span>
                    <h3 className="mt-2 font-mono text-lg font-bold text-slate-900">
                      {selectedFormula.name}
                    </h3>
                    <p className="mt-1 text-xs leading-relaxed text-slate-600">
                      {selectedFormula.description}
                    </p>
                  </div>

                  <div className="rounded-xl border border-slate-200 bg-white p-3 space-y-2">
                    <div>
                      <p className="text-[10.5px] font-semibold uppercase tracking-wider text-slate-400">
                        Cú pháp
                      </p>
                      <code className="mt-0.5 block font-mono text-xs font-semibold text-slate-800 break-all">
                        ={selectedFormula.syntax}
                      </code>
                    </div>
                    <div className="border-t border-slate-100 pt-2">
                      <p className="text-[10.5px] font-semibold uppercase tracking-wider text-slate-400">
                        Ví dụ thực tế
                      </p>
                      <code className="mt-0.5 block font-mono text-xs font-semibold text-[#107C41] break-all">
                        {selectedFormula.example}
                      </code>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="py-10 text-center text-xs text-slate-400">
                  Chọn một hàm bên trái để xem chi tiết cú pháp.
                </div>
              )}

              <div className="mt-5 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={onClose}
                  className="h-9 rounded-xl border border-slate-200 bg-white px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
                >
                  Đóng
                </button>
                <button
                  type="button"
                  disabled={!selectedFormula}
                  onClick={() => {
                    if (!selectedFormula) return;
                    onInsertFormula(selectedFormula.snippet);
                    onClose();
                  }}
                  className="inline-flex h-9 items-center gap-1.5 rounded-xl bg-[#107C41] px-4 text-xs font-semibold text-white shadow-sm hover:bg-[#0B5E31] disabled:opacity-50"
                >
                  <FunctionSquare className="h-3.5 w-3.5" />
                  <span>Chèn vào ô đang chọn ({activeCellRef})</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
 * 2. ExcelAiToolsModal (100% client-side summary, NL formula, audit)
 * ========================================================================== */

export type ExcelAiToolTab = 'summary' | 'formula' | 'audit';

export interface ExcelAiToolsModalProps {
  isOpen: boolean;
  initialTab?: ExcelAiToolTab;
  sheetName: string;
  rows: (string | number)[][];
  evaluatedMatrix: (string | number)[][];
  activeCellRef?: string;
  isProtected?: boolean;
  onClose: () => void;
  onInsertFormula: (formula: string) => void;
  onCreateSummarySheet?: (
    title: string,
    summaryRows: (string | number)[][]
  ) => void;
  onNavigateToCell?: (rowIdx: number, colIdx: number) => void;
  onHighlightAuditCells?: (
    cells: Array<{ row: number; col: number }>,
    bgColor: string
  ) => void;
}

interface ColumnNumericSummary {
  colIdx: number;
  letter: string;
  header: string;
  count: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
  firstDataRow1Based: number;
  lastDataRow1Based: number;
}

const STAT_NUM_FMT = new Intl.NumberFormat('vi-VN', {
  maximumFractionDigits: 2,
});

function parseAiNumericCell(value: string | number | undefined): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value !== 'string') {
    return null;
  }
  const trimmed = value.trim();
  if (!trimmed || trimmed.startsWith('#')) {
    return null;
  }
  const cleaned = trimmed
    .replace(/[%₫,\s]/g, '')
    .replace(/tr$/i, '');
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(cleaned)) {
    return null;
  }
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function isErrorCellValue(val: string | number | undefined): boolean {
  return (
    typeof val === 'string' &&
    /^#(DIV\/0!|CYCLE!|ERROR!|REF!|NAME\?|VALUE!|N\/A)$/i.test(val.trim())
  );
}

function buildFormulaFromVietnamesePrompt(
  prompt: string,
  colSummaries: ColumnNumericSummary[],
  totalRows: number
): { formula: string; explanation: string } | null {
  const raw = prompt.trim();
  if (!raw) {
    return null;
  }
  const lower = raw.toLowerCase();

  // Extract explicit A1:B10 range if present
  const rangeMatch = raw.match(/\b([A-Za-z]+\d+)\s*:\s*([A-Za-z]+\d+)\b/);
  let detectedRange = rangeMatch
    ? `${rangeMatch[1].toUpperCase()}:${rangeMatch[2].toUpperCase()}`
    : null;

  // Detect "cột X từ dòng N đến M" or "từ XN đến XM"
  if (!detectedRange) {
    const colRangeMatch = lower.match(
      /cột\s+([a-z]+)\s+(?:từ\s+)?(?:dòng|hàng)\s+(\d+)\s+(?:đến|tới|-)\s+(?:dòng|hàng\s+)?(\d+)/i
    );
    if (colRangeMatch) {
      const colL = colRangeMatch[1].toUpperCase();
      detectedRange = `${colL}${colRangeMatch[2]}:${colL}${colRangeMatch[3]}`;
    }
  }

  if (!detectedRange) {
    const colOnlyMatch = lower.match(/cột\s+([a-z]+)/i);
    if (colOnlyMatch) {
      const colL = colOnlyMatch[1].toUpperCase();
      const matchedCol = colSummaries.find((c) => c.letter === colL);
      const startRow = matchedCol ? matchedCol.firstDataRow1Based : 2;
      const endRow = Math.max(startRow, matchedCol ? matchedCol.lastDataRow1Based : totalRows);
      detectedRange = `${colL}${startRow}:${colL}${endRow}`;
    }
  }

  const fallbackCol = colSummaries[0];
  const defaultRange = fallbackCol
    ? `${fallbackCol.letter}${fallbackCol.firstDataRow1Based}:${fallbackCol.letter}${fallbackCol.lastDataRow1Based}`
    : `A2:A${Math.max(2, totalRows)}`;
  const targetRange = detectedRange ?? defaultRange;

  // VLOOKUP: "tìm E2 trong A2:C10 lấy cột 3"
  if (lower.includes('tìm') || lower.includes('tra cứu') || lower.includes('vlookup')) {
    const cellMatches = raw.match(/\b([A-Za-z]+\d+)\b/g) ?? [];
    const lookupCell = cellMatches[0]?.toUpperCase() || 'A2';
    const colIdxMatch = lower.match(/(?:lấy\s+)?cột\s+(\d+)/i);
    const colNum = colIdxMatch ? colIdxMatch[1] : '2';
    const tableRange = detectedRange ?? `A2:C${Math.max(2, totalRows)}`;
    return {
      formula: `=VLOOKUP(${lookupCell}, ${tableRange}, ${colNum}, FALSE)`,
      explanation: `Tra cứu chính xác giá trị ở ô ${lookupCell} trong bảng ${tableRange} và lấy cột thứ ${colNum}.`,
    };
  }

  // IF: "nếu D2 >= 5 thì Đạt ngược lại Trượt"
  const ifMatch = raw.match(
    /nếu\s+(.+?)\s+thì\s+(.+?)(?:\s+(?:ngược lại|còn không|nếu không)\s+(.+))?$/i
  );
  if (ifMatch && !lower.includes('đếm nếu') && !lower.includes('tổng nếu')) {
    const condRaw = ifMatch[1].trim().replace(/([a-z]+\d+)/gi, (m) => m.toUpperCase());
    const formatArg = (arg: string | undefined, fallback: string) => {
      if (!arg) return fallback;
      const t = arg.trim().replace(/^["']|["']$/g, '');
      if (/^[+-]?\d+(?:\.\d+)?$/.test(t) || /^(TRUE|FALSE)$/i.test(t) || /^[A-Za-z]+\d+$/.test(t)) {
        return t.toUpperCase();
      }
      return `"${t}"`;
    };
    const truePart = formatArg(ifMatch[2], '"Đạt"');
    const falsePart = formatArg(ifMatch[3], '""');
    return {
      formula: `=IF(${condRaw}, ${truePart}, ${falsePart})`,
      explanation: `Kiểm tra điều kiện (${condRaw}), trả ${truePart} nếu đúng và ${falsePart} nếu sai.`,
    };
  }

  // COUNTIF / SUMIF / AVERAGEIF
  const condOpMatch = raw.match(/(>=|<=|<>|!=|>|<|=)\s*([0-9.,]+|"[^"]+"|[^\s,]+)/);
  if ((lower.includes('đếm') && lower.includes('nếu')) || lower.includes('countif')) {
    const criteria = condOpMatch
      ? `"${condOpMatch[1] === '!=' ? '<>' : condOpMatch[1]}${condOpMatch[2].replace(/^"|"$/g, '')}"`
      : '">0"';
    return {
      formula: `=COUNTIF(${targetRange}, ${criteria})`,
      explanation: `Đếm các ô trong ${targetRange} thỏa điều kiện ${criteria}.`,
    };
  }

  if ((lower.includes('tổng') && lower.includes('nếu')) || lower.includes('sumif')) {
    const criteria = condOpMatch
      ? `"${condOpMatch[1] === '!=' ? '<>' : condOpMatch[1]}${condOpMatch[2].replace(/^"|"$/g, '')}"`
      : '">0"';
    return {
      formula: `=SUMIF(${targetRange}, ${criteria})`,
      explanation: `Tính tổng các ô trong ${targetRange} thỏa điều kiện ${criteria}.`,
    };
  }

  if ((lower.includes('trung bình') && lower.includes('nếu')) || lower.includes('averageif')) {
    const criteria = condOpMatch
      ? `"${condOpMatch[1] === '!=' ? '<>' : condOpMatch[1]}${condOpMatch[2].replace(/^"|"$/g, '')}"`
      : '">0"';
    return {
      formula: `=AVERAGEIF(${targetRange}, ${criteria})`,
      explanation: `Tính trung bình các ô trong ${targetRange} thỏa điều kiện ${criteria}.`,
    };
  }

  if (lower.includes('tổng') || lower.includes('cộng') || lower.includes('sum')) {
    return {
      formula: `=SUM(${targetRange})`,
      explanation: `Tính tổng các giá trị số trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('trung bình') || lower.includes('bình quân') || lower.includes('avg')) {
    return {
      formula: `=AVERAGE(${targetRange})`,
      explanation: `Tính trung bình cộng trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('trống') && lower.includes('đếm')) {
    return {
      formula: `=COUNTBLANK(${targetRange})`,
      explanation: `Đếm số ô đang bỏ trống trong vùng ${targetRange}.`,
    };
  }

  if (
    lower.includes('khác rỗng') ||
    lower.includes('không trống') ||
    lower.includes('counta')
  ) {
    return {
      formula: `=COUNTA(${targetRange})`,
      explanation: `Đếm số ô có dữ liệu (khác rỗng) trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('đếm') || lower.includes('bao nhiêu')) {
    return {
      formula: `=COUNT(${targetRange})`,
      explanation: `Đếm số ô chứa giá trị số trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('lớn nhất') || lower.includes('cao nhất') || lower.includes('max')) {
    return {
      formula: `=MAX(${targetRange})`,
      explanation: `Tìm giá trị lớn nhất trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('nhỏ nhất') || lower.includes('thấp nhất') || lower.includes('min')) {
    return {
      formula: `=MIN(${targetRange})`,
      explanation: `Tìm giá trị nhỏ nhất trong vùng ${targetRange}.`,
    };
  }

  if (lower.includes('hôm nay') || lower.includes('ngày hiện tại')) {
    return {
      formula: '=TODAY()',
      explanation: 'Lấy ngày tháng năm hiện tại.',
    };
  }

  if (lower.includes('trả góp') || lower.includes('khoản vay') || lower.includes('pmt')) {
    return {
      formula: '=PMT(0.08/12, 12, -100000000)',
      explanation: 'Mẫu tính tiền trả góp hàng tháng (lãi suất 8%/năm trong 12 tháng cho khoản vay 100tr).',
    };
  }

  if (lower.includes('viết hoa') || lower.includes('in hoa')) {
    const cellRef = raw.match(/\b([A-Za-z]+\d+)\b/)?.[1]?.toUpperCase() || 'A2';
    return {
      formula: `=UPPER(${cellRef})`,
      explanation: `Chuyển nội dung ô ${cellRef} sang chữ IN HOA.`,
    };
  }

  return {
    formula: `=SUM(${targetRange})`,
    explanation: `Gợi ý tính tổng cho vùng dữ liệu ${targetRange}. Bạn có thể viết rõ hơn như "trung bình C2:C10" hoặc "nếu D2 >= 5 thì Đạt ngược lại Trượt".`,
  };
}

export function ExcelAiToolsModal({
  isOpen,
  initialTab = 'summary',
  sheetName,
  rows,
  evaluatedMatrix,
  activeCellRef = 'A1',
  isProtected = false,
  onClose,
  onInsertFormula,
  onCreateSummarySheet,
  onNavigateToCell,
  onHighlightAuditCells,
}: ExcelAiToolsModalProps) {
  const [activeTab, setActiveTab] = useState<ExcelAiToolTab>(initialTab);
  const [nlInput, setNlInput] = useState('');
  const [copiedSummary, setCopiedSummary] = useState(false);

  useEffect(() => {
    if (isOpen) {
      setActiveTab(initialTab);
      setCopiedSummary(false);
    }
  }, [initialTab, isOpen]);

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

  const analysis = useMemo(() => {
    const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);
    const colSummaries: ColumnNumericSummary[] = [];
    const errorCells: Array<{
      row: number;
      col: number;
      ref: string;
      code: string;
      raw: string;
    }> = [];
    const valueOccurrences = new Map<
      string,
      Array<{ row: number; col: number; ref: string; display: string }>
    >();

    let nonEmptyCells = 0;
    let totalCells = rows.length * colCount;

    for (let c = 0; c < colCount; c += 1) {
      const headerRaw = String(evaluatedMatrix[0]?.[c] ?? rows[0]?.[c] ?? '').trim();
      const header = headerRaw || `Cột ${columnIndexToLetter(c)}`;
      let numCount = 0;
      let sum = 0;
      let min = Infinity;
      let max = -Infinity;
      let firstDataRow1Based = 2;
      let lastDataRow1Based = Math.max(2, rows.length);

      for (let r = 0; r < rows.length; r += 1) {
        const evalVal = evaluatedMatrix[r]?.[c] ?? rows[r]?.[c] ?? '';
        const rawVal = rows[r]?.[c] ?? '';
        const strEval = String(evalVal).trim();

        if (strEval !== '') {
          nonEmptyCells += 1;
        }

        if (isErrorCellValue(evalVal)) {
          errorCells.push({
            row: r,
            col: c,
            ref: formatCellRef(r, c),
            code: strEval,
            raw: String(rawVal),
          });
        }

        if (r > 0 && strEval !== '' && !isErrorCellValue(evalVal)) {
          const normalizedKey = strEval.toLowerCase();
          // Skip pure numeric values to highlight meaningful text duplicates (IDs, names, etc.)
          if (parseAiNumericCell(evalVal) === null) {
            const list = valueOccurrences.get(normalizedKey) ?? [];
            list.push({
              row: r,
              col: c,
              ref: formatCellRef(r, c),
              display: strEval,
            });
            valueOccurrences.set(normalizedKey, list);
          }
        }

        if (r > 0) {
          const parsed = parseAiNumericCell(evalVal);
          if (parsed !== null) {
            if (numCount === 0) {
              firstDataRow1Based = r + 1;
            }
            lastDataRow1Based = r + 1;
            numCount += 1;
            sum += parsed;
            if (parsed < min) min = parsed;
            if (parsed > max) max = parsed;
          }
        }
      }

      if (numCount > 0) {
        colSummaries.push({
          colIdx: c,
          letter: columnIndexToLetter(c),
          header,
          count: numCount,
          sum,
          avg: sum / numCount,
          min,
          max,
          firstDataRow1Based,
          lastDataRow1Based,
        });
      }
    }

    const duplicateGroups = Array.from(valueOccurrences.values()).filter(
      (list) => list.length > 1
    );

    const topCol = colSummaries.reduce<ColumnNumericSummary | null>(
      (best, cur) => (!best || cur.sum > best.sum ? cur : best),
      null
    );

    return {
      colCount,
      totalRows: rows.length,
      totalCells,
      nonEmptyCells,
      emptyCells: Math.max(0, totalCells - nonEmptyCells),
      colSummaries,
      errorCells,
      duplicateGroups,
      topCol,
    };
  }, [evaluatedMatrix, rows]);

  const formulaSuggestion = useMemo(
    () =>
      buildFormulaFromVietnamesePrompt(
        nlInput,
        analysis.colSummaries,
        analysis.totalRows
      ),
    [analysis.colSummaries, analysis.totalRows, nlInput]
  );

  if (!isOpen) {
    return null;
  }

  const summaryPlainText = [
    `Tóm tắt trang tính "${sheetName}":`,
    `- Quy mô: ${analysis.totalRows} dòng × ${analysis.colCount} cột (${analysis.nonEmptyCells} ô có dữ liệu, ${analysis.emptyCells} ô trống).`,
    ...analysis.colSummaries.map(
      (col) =>
        `- ${col.header} (${col.letter}): Tổng = ${STAT_NUM_FMT.format(col.sum)}, TB = ${STAT_NUM_FMT.format(col.avg)}, Min = ${STAT_NUM_FMT.format(col.min)}, Max = ${STAT_NUM_FMT.format(col.max)} (${col.count} mục).`
    ),
    analysis.topCol
      ? `- Nhận xét: Cột "${analysis.topCol.header}" có tổng giá trị cao nhất (${STAT_NUM_FMT.format(analysis.topCol.sum)}).`
      : '- Chưa phát hiện cột dữ liệu số để tổng hợp.',
  ].join('\n');

  const handleCopySummary = async () => {
    try {
      await navigator.clipboard.writeText(summaryPlainText);
      setCopiedSummary(true);
      window.setTimeout(() => setCopiedSummary(false), 2000);
    } catch {
      // Ignore clipboard failure
    }
  };

  const handleInsertSummarySheet = () => {
    if (!onCreateSummarySheet) {
      return;
    }
    const summaryRows: (string | number)[][] = [
      ['Cột', 'Tên tiêu đề', 'Số mục', 'Tổng cộng', 'Trung bình', 'Nhỏ nhất', 'Lớn nhất'],
      ...analysis.colSummaries.map((c) => [
        c.letter,
        c.header,
        c.count,
        Number(c.sum.toFixed(2)),
        Number(c.avg.toFixed(2)),
        Number(c.min.toFixed(2)),
        Number(c.max.toFixed(2)),
      ]),
    ];
    if (summaryRows.length === 1) {
      summaryRows.push(['—', 'Không có cột số liệu', 0, 0, 0, 0, 0]);
    }
    onCreateSummarySheet(`Tóm tắt - ${sheetName}`.slice(0, 28), summaryRows);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="excel-ai-modal-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/55 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-3 sm:p-5">
        <div
          onClick={(event) => event.stopPropagation()}
          className="flex max-h-full w-full max-w-3xl flex-col overflow-hidden rounded-2xl border border-slate-200 bg-white text-slate-900 shadow-2xl"
        >
          {/* Header */}
          <header className="flex shrink-0 items-center justify-between gap-3 border-b border-slate-200 bg-[#F8FAFC] px-4 py-3 sm:px-5">
            <div className="flex items-center gap-2.5 min-w-0">
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-[#ECFDF5] text-[#107C41] border border-[#107C41]/20">
                <Sparkles className="h-5 w-5" />
              </span>
              <div className="min-w-0">
                <h2
                  id="excel-ai-modal-title"
                  className="truncate text-sm sm:text-base font-bold text-slate-900"
                >
                  AI Tools — Trợ lý bảng tính thông minh
                </h2>
                <p className="truncate text-xs text-slate-500">
                  Phân tích cục bộ 100% trên trình duyệt · Trang tính: {sheetName}
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng AI Tools"
              className="flex h-8 w-8 items-center justify-center rounded-xl text-slate-400 hover:bg-slate-200/70 hover:text-slate-700"
            >
              <X className="h-4 w-4" />
            </button>
          </header>

          {/* Tabs */}
          <div className="shrink-0 flex items-center gap-1.5 border-b border-slate-200 bg-white px-4 pt-2.5 sm:px-5 overflow-x-auto no-scrollbar">
            {[
              { key: 'summary' as const, label: 'Tóm tắt số liệu' },
              { key: 'formula' as const, label: 'Tạo công thức từ tiếng Việt' },
              {
                key: 'audit' as const,
                label: `Phát hiện lỗi / trùng (${analysis.errorCells.length + analysis.duplicateGroups.length})`,
              },
            ].map((tab) => {
              const isActive = activeTab === tab.key;
              return (
                <button
                  key={tab.key}
                  type="button"
                  onClick={() => setActiveTab(tab.key)}
                  className={`relative shrink-0 whitespace-nowrap px-3 py-2 text-xs font-semibold transition-colors ${
                    isActive
                      ? 'text-[#107C41]'
                      : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  {tab.label}
                  {isActive && (
                    <span className="absolute inset-x-2 bottom-0 h-[2px] rounded-full bg-[#107C41]" />
                  )}
                </button>
              );
            })}
          </div>

          {/* Content */}
          <div className="min-h-0 flex-1 overflow-y-auto p-4 sm:p-5 bg-[#F8FAFC]">
            {activeTab === 'summary' && (
              <div className="space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                  <div className="rounded-xl border border-slate-200 bg-white p-3">
                    <p className="text-[11px] text-slate-500">Tổng dòng</p>
                    <p className="mt-1 text-lg font-bold text-slate-900 tabular-nums">
                      {analysis.totalRows}
                    </p>
                  </div>
                  <div className="rounded-xl border border-slate-200 bg-white p-3">
                    <p className="text-[11px] text-slate-500">Tổng cột</p>
                    <p className="mt-1 text-lg font-bold text-slate-900 tabular-nums">
                      {analysis.colCount}
                    </p>
                  </div>
                  <div className="rounded-xl border border-slate-200 bg-white p-3">
                    <p className="text-[11px] text-slate-500">Ô có dữ liệu</p>
                    <p className="mt-1 text-lg font-bold text-[#107C41] tabular-nums">
                      {analysis.nonEmptyCells}
                    </p>
                  </div>
                  <div className="rounded-xl border border-slate-200 bg-white p-3">
                    <p className="text-[11px] text-slate-500">Ô trống</p>
                    <p className="mt-1 text-lg font-bold text-slate-600 tabular-nums">
                      {analysis.emptyCells}
                    </p>
                  </div>
                </div>

                {analysis.topCol && (
                  <div className="rounded-xl border border-[#107C41]/25 bg-[#ECFDF5] px-4 py-3 text-xs text-[#0B5E31]">
                    <strong>Nhận xét nhanh:</strong> Cột{' '}
                    <strong>
                      {analysis.topCol.header} ({analysis.topCol.letter})
                    </strong>{' '}
                    có tổng giá trị lớn nhất đạt{' '}
                    <strong>{STAT_NUM_FMT.format(analysis.topCol.sum)}</strong>{' '}
                    (trung bình{' '}
                    <strong>{STAT_NUM_FMT.format(analysis.topCol.avg)}</strong>{' '}
                    trên {analysis.topCol.count} mục).
                  </div>
                )}

                {analysis.colSummaries.length === 0 ? (
                  <div className="rounded-xl border border-dashed border-slate-300 bg-white p-8 text-center text-xs text-slate-500">
                    Chưa tìm thấy cột chứa số liệu để tổng hợp thống kê.
                  </div>
                ) : (
                  <div className="overflow-x-auto rounded-xl border border-slate-200 bg-white">
                    <table className="w-full text-left text-xs">
                      <thead className="border-b border-slate-200 bg-slate-50 text-[11px] font-semibold text-slate-600">
                        <tr>
                          <th className="px-3 py-2">Cột</th>
                          <th className="px-3 py-2 text-right">Số mục</th>
                          <th className="px-3 py-2 text-right">Tổng</th>
                          <th className="px-3 py-2 text-right">Trung bình</th>
                          <th className="px-3 py-2 text-right">Nhỏ nhất</th>
                          <th className="px-3 py-2 text-right">Lớn nhất</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 tabular-nums">
                        {analysis.colSummaries.map((col) => (
                          <tr key={col.colIdx} className="hover:bg-slate-50/70">
                            <td className="px-3 py-2 font-semibold text-slate-800">
                              <span className="mr-1.5 inline-block rounded bg-[#ECFDF5] px-1.5 py-0.5 font-mono text-[10px] text-[#107C41]">
                                {col.letter}
                              </span>
                              {col.header}
                            </td>
                            <td className="px-3 py-2 text-right text-slate-600">
                              {col.count}
                            </td>
                            <td className="px-3 py-2 text-right font-bold text-[#107C41]">
                              {STAT_NUM_FMT.format(col.sum)}
                            </td>
                            <td className="px-3 py-2 text-right text-slate-700">
                              {STAT_NUM_FMT.format(col.avg)}
                            </td>
                            <td className="px-3 py-2 text-right text-slate-600">
                              {STAT_NUM_FMT.format(col.min)}
                            </td>
                            <td className="px-3 py-2 text-right text-slate-700">
                              {STAT_NUM_FMT.format(col.max)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}

                <div className="flex flex-wrap items-center justify-end gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => void handleCopySummary()}
                    className="inline-flex h-9 items-center gap-1.5 rounded-xl border border-slate-200 bg-white px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
                  >
                    {copiedSummary ? (
                      <Check className="h-3.5 w-3.5 text-[#107C41]" />
                    ) : (
                      <Copy className="h-3.5 w-3.5" />
                    )}
                    <span>{copiedSummary ? 'Đã sao chép' : 'Sao chép tóm tắt'}</span>
                  </button>

                  {onCreateSummarySheet && (
                    <button
                      type="button"
                      onClick={handleInsertSummarySheet}
                      className="inline-flex h-9 items-center gap-1.5 rounded-xl bg-[#107C41] px-4 text-xs font-semibold text-white hover:bg-[#0B5E31]"
                    >
                      <FilePlus2 className="h-3.5 w-3.5" />
                      <span>Chèn vào trang tính mới</span>
                    </button>
                  )}
                </div>
              </div>
            )}

            {activeTab === 'formula' && (
              <div className="space-y-4">
                <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
                  <label
                    htmlFor="ai-nl-formula-input"
                    className="block text-xs font-semibold text-slate-800"
                  >
                    Mô tả phép tính bằng tiếng Việt tự nhiên (chèn vào ô{' '}
                    <span className="text-[#107C41]">{activeCellRef}</span>):
                  </label>
                  <textarea
                    id="ai-nl-formula-input"
                    rows={2}
                    value={nlInput}
                    onChange={(event) => setNlInput(event.target.value)}
                    placeholder="Ví dụ: tổng cột B từ dòng 2 đến 10, hoặc: nếu D2 >= 5 thì Đạt ngược lại Trượt"
                    className="w-full rounded-xl border border-slate-200 bg-slate-50 p-3 text-xs sm:text-sm text-slate-900 placeholder:text-slate-400 focus:border-[#107C41] focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#107C41]/15"
                  />

                  <div className="flex flex-wrap gap-1.5">
                    {[
                      'Tổng cột B từ dòng 2 đến 10',
                      'Trung bình C2:C20',
                      'Nếu D2 >= 5 thì Đạt ngược lại Trượt',
                      'Đếm ô khác rỗng A2:A50',
                      'Tìm E2 trong A2:C10 lấy cột 3',
                      'Đếm nếu B2:B20 > 10',
                    ].map((sample) => (
                      <button
                        key={sample}
                        type="button"
                        onClick={() => setNlInput(sample)}
                        className="rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:border-[#107C41]/40 hover:bg-[#ECFDF5] hover:text-[#107C41] transition-colors"
                      >
                        {sample}
                      </button>
                    ))}
                  </div>
                </div>

                {formulaSuggestion ? (
                  <div className="rounded-xl border border-[#107C41]/30 bg-white p-4 space-y-3 shadow-sm">
                    <div className="flex items-center justify-between gap-2">
                      <span className="inline-flex items-center gap-1.5 text-xs font-bold text-[#107C41]">
                        <Wand2 className="h-4 w-4" />
                        Công thức được tạo
                      </span>
                      <button
                        type="button"
                        disabled={isProtected}
                        onClick={() => {
                          onInsertFormula(formulaSuggestion.formula);
                          onClose();
                        }}
                        className="inline-flex h-8 items-center gap-1.5 rounded-lg bg-[#107C41] px-3 text-xs font-semibold text-white hover:bg-[#0B5E31] disabled:opacity-50"
                      >
                        <span>Chèn vào {activeCellRef}</span>
                      </button>
                    </div>
                    <code className="block rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 font-mono text-sm font-bold text-slate-900">
                      {formulaSuggestion.formula}
                    </code>
                    <p className="text-xs text-slate-600">
                      {formulaSuggestion.explanation}
                    </p>
                  </div>
                ) : (
                  <div className="rounded-xl border border-dashed border-slate-300 bg-white p-6 text-center text-xs text-slate-500">
                    Nhập yêu cầu hoặc chọn một mẫu gợi ý phía trên để sinh công thức Excel tức thì.
                  </div>
                )}
              </div>
            )}

            {activeTab === 'audit' && (
              <div className="space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-slate-200 bg-white p-3.5">
                  <div className="text-xs text-slate-700">
                    Phát hiện{' '}
                    <strong className="text-rose-600">
                      {analysis.errorCells.length} ô lỗi công thức
                    </strong>{' '}
                    và{' '}
                    <strong className="text-amber-600">
                      {analysis.duplicateGroups.length} nhóm giá trị trùng lặp
                    </strong>
                    .
                  </div>

                  {onHighlightAuditCells &&
                    (analysis.errorCells.length > 0 ||
                      analysis.duplicateGroups.length > 0) && (
                      <button
                        type="button"
                        disabled={isProtected}
                        onClick={() => {
                          const targetCells = [
                            ...analysis.errorCells.map((e) => ({
                              row: e.row,
                              col: e.col,
                            })),
                            ...analysis.duplicateGroups.flatMap((group) =>
                              group.map((item) => ({
                                row: item.row,
                                col: item.col,
                              }))
                            ),
                          ];
                          onHighlightAuditCells(targetCells, '#FEF3C7');
                        }}
                        className="inline-flex h-8 items-center gap-1.5 rounded-lg bg-amber-500 px-3 text-xs font-semibold text-white hover:bg-amber-600 disabled:opacity-50"
                      >
                        <PaintBucket className="h-3.5 w-3.5" />
                        <span>Tô màu cảnh báo</span>
                      </button>
                    )}
                </div>

                {analysis.errorCells.length === 0 &&
                analysis.duplicateGroups.length === 0 ? (
                  <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-6 text-center text-xs font-medium text-emerald-800">
                    Tuyệt vời! Trang tính hiện không có ô lỗi công thức hay văn bản trùng lặp nào.
                  </div>
                ) : (
                  <div className="space-y-3">
                    {analysis.errorCells.length > 0 && (
                      <div className="rounded-xl border border-rose-200 bg-white p-3.5 space-y-2">
                        <h3 className="flex items-center gap-1.5 text-xs font-bold text-rose-700">
                          <AlertTriangle className="h-4 w-4" />
                          Ô lỗi công thức ({analysis.errorCells.length})
                        </h3>
                        <ul className="divide-y divide-slate-100">
                          {analysis.errorCells.map((err) => (
                            <li
                              key={`${err.row}-${err.col}`}
                              className="flex items-center justify-between gap-2 py-2 text-xs"
                            >
                              <div className="min-w-0">
                                <span className="font-mono font-bold text-rose-600 mr-2">
                                  {err.ref}
                                </span>
                                <span className="rounded bg-rose-50 px-1.5 py-0.5 font-mono text-[11px] font-semibold text-rose-700 mr-2">
                                  {err.code}
                                </span>
                                <span className="truncate text-slate-500 font-mono">
                                  {err.raw}
                                </span>
                              </div>
                              {onNavigateToCell && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    onNavigateToCell(err.row, err.col);
                                    onClose();
                                  }}
                                  className="shrink-0 rounded-lg border border-slate-200 px-2.5 py-1 text-[11px] font-semibold text-slate-700 hover:bg-slate-50"
                                >
                                  Đi tới ô
                                </button>
                              )}
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}

                    {analysis.duplicateGroups.length > 0 && (
                      <div className="rounded-xl border border-amber-200 bg-white p-3.5 space-y-2">
                        <h3 className="flex items-center gap-1.5 text-xs font-bold text-amber-800">
                          <AlertTriangle className="h-4 w-4" />
                          Giá trị trùng lặp ({analysis.duplicateGroups.length} nhóm)
                        </h3>
                        <ul className="divide-y divide-slate-100">
                          {analysis.duplicateGroups.map((group, idx) => (
                            <li
                              key={`dup-${idx}`}
                              className="flex flex-wrap items-center justify-between gap-2 py-2 text-xs"
                            >
                              <div className="min-w-0">
                                <span className="font-semibold text-slate-900 mr-2">
                                  &ldquo;{group[0].display}&rdquo;
                                </span>
                                <span className="text-slate-500">
                                  ({group.length} lần tại{' '}
                                  {group.map((g) => g.ref).join(', ')})
                                </span>
                              </div>
                              {onNavigateToCell && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    onNavigateToCell(group[0].row, group[0].col);
                                    onClose();
                                  }}
                                  className="shrink-0 rounded-lg border border-slate-200 px-2.5 py-1 text-[11px] font-semibold text-slate-700 hover:bg-slate-50"
                                >
                                  Đi tới ô
                                </button>
                              )}
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
 * 3. ExcelDrawingOverlay (pointer-events-none when activeTool === null)
 * ========================================================================== */

export type ExcelDrawingTool = 'pen' | 'highlighter' | 'eraser' | null;

export interface ExcelDrawingOverlayProps {
  strokes: ExcelDrawingStroke[];
  activeTool: ExcelDrawingTool;
  activeColor: string;
  strokeWidth: number;
  zoomScale?: number;
  onAddStroke: (stroke: ExcelDrawingStroke) => void;
  onEraseStroke: (strokeId: string) => void;
}

function pointsToSvgPoints(points: Array<{ x: number; y: number }>): string {
  return points.map((p) => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
}

export function ExcelDrawingOverlay({
  strokes,
  activeTool,
  activeColor,
  strokeWidth,
  zoomScale = 1,
  onAddStroke,
  onEraseStroke,
}: ExcelDrawingOverlayProps) {
  const svgRef = useRef<SVGSVGElement | null>(null);
  const [draftPoints, setDraftPoints] = useState<Array<{ x: number; y: number }>>(
    []
  );
  const isDrawingRef = useRef(false);
  const isErasingRef = useRef(false);
  const safeScale = zoomScale > 0 ? zoomScale : 1;

  const getLocalPoint = (
    event: ReactPointerEvent<SVGSVGElement>
  ): { x: number; y: number } => {
    const rect = svgRef.current?.getBoundingClientRect();
    if (!rect) {
      return { x: 0, y: 0 };
    }
    return {
      x: (event.clientX - rect.left) / safeScale,
      y: (event.clientY - rect.top) / safeScale,
    };
  };

  const handlePointerDown = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (activeTool === 'eraser') {
      isErasingRef.current = true;
      return;
    }
    if (activeTool !== 'pen' && activeTool !== 'highlighter') {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    isDrawingRef.current = true;
    setDraftPoints([getLocalPoint(event)]);
  };

  const handlePointerMove = (event: ReactPointerEvent<SVGSVGElement>) => {
    if (!isDrawingRef.current) {
      return;
    }
    event.preventDefault();
    const pt = getLocalPoint(event);
    setDraftPoints((prev) => {
      const last = prev[prev.length - 1];
      if (last && Math.hypot(pt.x - last.x, pt.y - last.y) < 1.5) {
        return prev;
      }
      return [...prev, pt];
    });
  };

  const commitDraftStroke = () => {
    if (!isDrawingRef.current) {
      return;
    }
    isDrawingRef.current = false;
    if (
      (activeTool === 'pen' || activeTool === 'highlighter') &&
      draftPoints.length > 0
    ) {
      const normalizedPts =
        draftPoints.length === 1
          ? [draftPoints[0], { x: draftPoints[0].x + 0.5, y: draftPoints[0].y + 0.5 }]
          : draftPoints;
      onAddStroke({
        id: `draw-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        tool: activeTool,
        color: activeColor,
        width: activeTool === 'highlighter' ? Math.max(10, strokeWidth * 3) : strokeWidth,
        points: normalizedPts,
      });
    }
    setDraftPoints([]);
  };

  const isDrawMode = activeTool === 'pen' || activeTool === 'highlighter';
  const isEraseMode = activeTool === 'eraser';

  if (strokes.length === 0 && !isDrawMode) {
    return null;
  }

  return (
    <svg
      ref={svgRef}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={() => {
        isErasingRef.current = false;
        commitDraftStroke();
      }}
      onPointerCancel={() => {
        isErasingRef.current = false;
        commitDraftStroke();
      }}
      className={`absolute inset-0 h-full w-full z-30 ${
        isDrawMode
          ? 'pointer-events-auto cursor-crosshair touch-none'
          : 'pointer-events-none'
      }`}
    >
      {strokes.map((stroke) => (
        <polyline
          key={stroke.id}
          points={pointsToSvgPoints(stroke.points)}
          fill="none"
          stroke={stroke.color}
          strokeWidth={stroke.width}
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeOpacity={stroke.tool === 'highlighter' ? 0.35 : 0.92}
          onPointerDown={(event) => {
            if (isEraseMode) {
              event.stopPropagation();
              isErasingRef.current = true;
              onEraseStroke(stroke.id);
            }
          }}
          onPointerEnter={() => {
            if (isEraseMode && isErasingRef.current) {
              onEraseStroke(stroke.id);
            }
          }}
          className={
            isEraseMode
              ? 'pointer-events-auto cursor-pointer hover:opacity-50'
              : 'pointer-events-none'
          }
        />
      ))}

      {draftPoints.length > 0 && isDrawMode && (
        <polyline
          points={pointsToSvgPoints(draftPoints)}
          fill="none"
          stroke={activeColor}
          strokeWidth={
            activeTool === 'highlighter'
              ? Math.max(10, strokeWidth * 3)
              : strokeWidth
          }
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeOpacity={activeTool === 'highlighter' ? 0.35 : 0.92}
          className="pointer-events-none"
        />
      )}
    </svg>
  );
}

/* ============================================================================
 * 4. ExcelFloatingObjectsLayer
 * ========================================================================== */

export interface ExcelFloatingObjectsLayerProps {
  objects: ExcelFloatingObject[];
  selectedObjectId: string | null;
  isProtected?: boolean;
  isReadOnly?: boolean;
  zoomScale?: number;
  onSelectObject: (id: string | null) => void;
  onUpdateObject: (
    id: string,
    patch: Partial<ExcelFloatingObject>
  ) => void;
  onDeleteObject: (id: string) => void;
}

function renderFloatingShapeSvg(obj: ExcelFloatingObject) {
  const fill = obj.fillColor ?? '#ECFDF5';
  const stroke = obj.borderColor ?? '#107C41';
  const shape = obj.shapeKind ?? 'rect';

  return (
    <svg
      viewBox="0 0 100 100"
      preserveAspectRatio="none"
      className="h-full w-full"
      aria-hidden="true"
    >
      {shape === 'ellipse' ? (
        <ellipse
          cx="50"
          cy="50"
          rx="47"
          ry="47"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
        />
      ) : shape === 'roundRect' ? (
        <rect
          x="3"
          y="3"
          width="94"
          height="94"
          rx="16"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
        />
      ) : shape === 'arrowRight' ? (
        <polygon
          points="4,30 62,30 62,12 96,50 62,88 62,70 4,70"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
          strokeLinejoin="round"
        />
      ) : shape === 'star' ? (
        <polygon
          points="50,5 63,35 95,38 70,60 78,93 50,76 22,93 30,60 5,38 37,35"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
          strokeLinejoin="round"
        />
      ) : shape === 'callout' ? (
        <path
          d="M6,6 H94 V74 H42 L24,94 L26,74 H6 Z"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
          strokeLinejoin="round"
        />
      ) : (
        <rect
          x="3"
          y="3"
          width="94"
          height="94"
          fill={fill}
          stroke={stroke}
          strokeWidth="3"
        />
      )}
    </svg>
  );
}

export function ExcelFloatingObjectsLayer({
  objects,
  selectedObjectId,
  isProtected = false,
  isReadOnly = false,
  zoomScale = 1,
  onSelectObject,
  onUpdateObject,
  onDeleteObject,
}: ExcelFloatingObjectsLayerProps) {
  const [editingId, setEditingId] = useState<string | null>(null);
  const canEdit = !isProtected && !isReadOnly;
  const safeScale = zoomScale > 0 ? zoomScale : 1;

  if (objects.length === 0) {
    return null;
  }

  const startDragObject = (
    event: ReactPointerEvent<HTMLDivElement>,
    obj: ExcelFloatingObject
  ) => {
    if (!canEdit || editingId === obj.id) {
      return;
    }
    if ((event.target as HTMLElement).closest('button, textarea, input')) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    onSelectObject(obj.id);

    const startX = event.clientX;
    const startY = event.clientY;
    const origX = obj.x;
    const origY = obj.y;

    const handleMove = (moveEvent: PointerEvent) => {
      const dx = (moveEvent.clientX - startX) / safeScale;
      const dy = (moveEvent.clientY - startY) / safeScale;
      onUpdateObject(obj.id, {
        x: Math.max(0, Math.round(origX + dx)),
        y: Math.max(0, Math.round(origY + dy)),
      });
    };

    const handleUp = () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);
    };

    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
  };

  return (
    <div className="pointer-events-none absolute inset-0 z-20">
      {objects.map((obj) => {
        const isSelected = selectedObjectId === obj.id;
        const isEditing = editingId === obj.id;

        return (
          <div
            key={obj.id}
            onPointerDown={(event) => startDragObject(event, obj)}
            onDoubleClick={(event) => {
              if (!canEdit || obj.kind === 'image') {
                return;
              }
              event.stopPropagation();
              onSelectObject(obj.id);
              setEditingId(obj.id);
            }}
            style={{
              left: obj.x,
              top: obj.y,
              width: obj.width,
              height: obj.height,
            }}
            className={`pointer-events-auto absolute rounded-lg transition-shadow ${
              canEdit ? 'cursor-move' : 'cursor-default'
            } ${
              isSelected
                ? 'ring-2 ring-[#107C41] shadow-lg'
                : 'hover:ring-1 hover:ring-[#107C41]/50'
            }`}
          >
            {obj.kind === 'image' && obj.src ? (
              <img
                src={obj.src}
                alt={obj.text || 'Ảnh chèn trong bảng tính'}
                draggable={false}
                className="h-full w-full rounded-lg object-contain bg-white/90 border border-slate-200"
              />
            ) : obj.kind === 'shape' ? (
              <div className="relative h-full w-full">
                {renderFloatingShapeSvg(obj)}
                <div className="absolute inset-0 flex items-center justify-center p-2">
                  {isEditing ? (
                    <textarea
                      autoFocus
                      value={obj.text ?? ''}
                      onChange={(event) =>
                        onUpdateObject(obj.id, { text: event.target.value })
                      }
                      onBlur={() => setEditingId(null)}
                      className="h-full w-full resize-none bg-transparent text-center text-xs font-semibold text-slate-900 focus:outline-none"
                    />
                  ) : (
                    <span className="text-center text-xs font-semibold text-slate-800 break-words">
                      {obj.text ?? ''}
                    </span>
                  )}
                </div>
              </div>
            ) : (
              <div
                style={{
                  backgroundColor: obj.fillColor ?? '#FFFFFF',
                  borderColor: obj.borderColor ?? '#107C41',
                }}
                className="h-full w-full rounded-lg border p-2 shadow-sm"
              >
                {isEditing ? (
                  <textarea
                    autoFocus
                    value={obj.text ?? ''}
                    onChange={(event) =>
                      onUpdateObject(obj.id, { text: event.target.value })
                    }
                    onBlur={() => setEditingId(null)}
                    placeholder="Nhập nội dung..."
                    className="h-full w-full resize-none bg-transparent text-xs text-slate-900 focus:outline-none"
                  />
                ) : (
                  <div className="h-full w-full overflow-hidden whitespace-pre-wrap text-xs text-slate-800">
                    {obj.text || 'Nhấp đúp để nhập chữ'}
                  </div>
                )}
              </div>
            )}

            {isSelected && canEdit && (
              <button
                type="button"
                onClick={(event) => {
                  event.stopPropagation();
                  onDeleteObject(obj.id);
                }}
                title="Xóa đối tượng"
                aria-label="Xóa đối tượng"
                className="absolute -right-2.5 -top-2.5 flex h-6 w-6 items-center justify-center rounded-full bg-rose-600 text-white shadow-md hover:bg-rose-700"
              >
                <Trash2 className="h-3 w-3" />
              </button>
            )}
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================================
 * 5. ExcelCellMetaModal (Comment or Hyperlink editor)
 * ========================================================================== */

export type ExcelCellMetaMode = 'comment' | 'hyperlink';

export interface ExcelCellMetaModalProps {
  isOpen: boolean;
  mode: ExcelCellMetaMode;
  cellRef: string;
  initialValue: string;
  onClose: () => void;
  onSave: (value: string) => void;
  onRemove?: () => void;
}

export function ExcelCellMetaModal({
  isOpen,
  mode,
  cellRef,
  initialValue,
  onClose,
  onSave,
  onRemove,
}: ExcelCellMetaModalProps) {
  const [draft, setDraft] = useState(initialValue);

  useEffect(() => {
    if (isOpen) {
      setDraft(initialValue);
    }
  }, [initialValue, isOpen]);

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

  const isComment = mode === 'comment';

  const handleConfirm = () => {
    const trimmed = draft.trim();
    if (!trimmed) {
      onRemove?.();
      onClose();
      return;
    }
    onSave(trimmed);
    onClose();
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="excel-cell-meta-title"
      onClick={onClose}
      className="fixed inset-0 z-50 bg-slate-950/55 backdrop-blur-[2px] pb-safe"
    >
      <div className="flex h-full w-full items-center justify-center p-4">
        <div
          onClick={(event) => event.stopPropagation()}
          className="w-full max-w-md rounded-2xl border border-slate-200 bg-white p-5 text-slate-900 shadow-2xl"
        >
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2.5">
              <span className="flex h-9 w-9 items-center justify-center rounded-xl bg-[#ECFDF5] text-[#107C41]">
                {isComment ? (
                  <MessageSquarePlus className="h-4 w-4" />
                ) : (
                  <Link2 className="h-4 w-4" />
                )}
              </span>
              <div>
                <h2
                  id="excel-cell-meta-title"
                  className="text-sm font-bold text-slate-900"
                >
                  {isComment ? 'Ghi chú ô' : 'Siêu liên kết ô'}{' '}
                  <span className="text-[#107C41]">{cellRef}</span>
                </h2>
                <p className="text-[11px] text-slate-500">
                  {isComment
                    ? 'Hiển thị tam giác đỏ ở góc trên phải của ô.'
                    : 'Gắn đường dẫn web (https://...) mở khi bấm vào ô.'}
                </p>
              </div>
            </div>

            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng"
              className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 hover:text-slate-700"
            >
              <X className="h-4 w-4" />
            </button>
          </div>

          <div className="mt-4">
            {isComment ? (
              <textarea
                autoFocus
                rows={4}
                value={draft}
                onChange={(event) => setDraft(event.target.value)}
                placeholder="Nhập nội dung ghi chú cho ô này..."
                className="w-full rounded-xl border border-slate-200 bg-slate-50 p-3 text-xs sm:text-sm text-slate-900 focus:border-[#107C41] focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#107C41]/15"
              />
            ) : (
              <input
                autoFocus
                type="url"
                value={draft}
                onChange={(event) => setDraft(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    handleConfirm();
                  }
                }}
                placeholder="https://example.com"
                className="h-10 w-full rounded-xl border border-slate-200 bg-slate-50 px-3 text-xs sm:text-sm text-slate-900 focus:border-[#107C41] focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#107C41]/15"
              />
            )}
          </div>

          <div className="mt-4 flex items-center justify-between gap-2">
            {onRemove && initialValue.trim().length > 0 ? (
              <button
                type="button"
                onClick={() => {
                  onRemove();
                  onClose();
                }}
                className="inline-flex h-9 items-center gap-1 rounded-xl border border-rose-200 bg-rose-50 px-3 text-xs font-semibold text-rose-700 hover:bg-rose-100"
              >
                <Trash2 className="h-3.5 w-3.5" />
                <span>Xóa</span>
              </button>
            ) : (
              <span />
            )}

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onClose}
                className="h-9 rounded-xl border border-slate-200 bg-white px-3.5 text-xs font-semibold text-slate-700 hover:bg-slate-50"
              >
                Hủy
              </button>
              <button
                type="button"
                onClick={handleConfirm}
                className="h-9 rounded-xl bg-[#107C41] px-4 text-xs font-semibold text-white hover:bg-[#0B5E31]"
              >
                Lưu
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

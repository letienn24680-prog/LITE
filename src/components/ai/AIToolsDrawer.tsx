import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  ArrowRight,
  Check,
  Copy,
  FilePlus2,
  KeyRound,
  Layers,
  Loader2,
  PenLine,
  Replace,
  Sparkles,
  SpellCheck2,
  Table2,
  Wand2,
  X,
} from 'lucide-react';
import {
  generateExcelFormulaWithAi,
  generateSlideDeckWithAi,
  getConfiguredGeminiApiKey,
  isAiEndpointAvailable,
  runWordAiAssistant,
  saveConfiguredGeminiApiKey,
  type AiExecutionSource,
  type AiGeneratedSlideDraft,
  type WordAiMode,
} from '../../services/aiAssistant';

export type AiDrawerMode = 'word' | 'excel' | 'pptx';

export interface AIToolsDrawerProps {
  isOpen: boolean;
  mode: AiDrawerMode;
  initialAction?: string;
  isDark?: boolean;
  isReadOnly?: boolean;
  documentTitle: string;
  // Word context
  wordPlainText?: string;
  wordSelectedText?: string;
  onWordInsertContent?: (payload: { html: string; text: string }) => void;
  onWordReplaceSelection?: (payload: { html: string; text: string }) => void;
  // Excel context
  excelSheetName?: string;
  excelColumns?: string[];
  excelRows?: (string | number)[][];
  excelActiveCellRef?: string;
  onExcelInsertFormula?: (formula: string) => void;
  onExcelOpenAudit?: () => void;
  // PowerPoint context
  pptxSlideCount?: number;
  onPptxAppendSlides?: (slides: AiGeneratedSlideDraft[]) => void;
  onPptxOpenOtherAiTab?: (tab: 'summarize' | 'notes') => void;
  onClose: () => void;
}

const ACCENT_BY_MODE: Record<
  AiDrawerMode,
  { hex: string; hover: string; softBg: string; border: string; label: string }
> = {
  word: {
    hex: '#185ABD',
    hover: '#10489A',
    softBg: '#EEF4FF',
    border: 'border-blue-200',
    label: 'Word AI Assistant',
  },
  excel: {
    hex: '#107C41',
    hover: '#0B5E31',
    softBg: '#ECFDF5',
    border: 'border-emerald-200',
    label: 'Excel AI Assistant',
  },
  pptx: {
    hex: '#C43E1C',
    hover: '#9E3014',
    softBg: '#FFF4ED',
    border: 'border-orange-200',
    label: 'PowerPoint AI Assistant',
  },
};

const EXCEL_PROMPT_PRESETS = [
  'Tính tổng cột D nếu cột B là 12A5',
  'Tính điểm trung bình',
  'Đếm số dòng nếu cột C là Đạt',
  'Tìm giá trị lớn nhất cột D',
] as const;

const SLIDE_STYLE_OPTIONS = [
  'Chuyên nghiệp',
  'Giáo dục & Học thuật',
  'Báo cáo kinh doanh',
  'Sáng tạo / Pitch Deck',
] as const;

export function AIToolsDrawer({
  isOpen,
  mode,
  initialAction,
  isDark = false,
  isReadOnly = false,
  documentTitle,
  wordPlainText = '',
  wordSelectedText = '',
  onWordInsertContent,
  onWordReplaceSelection,
  excelSheetName = 'Sheet1',
  excelColumns = [],
  excelRows = [],
  excelActiveCellRef = 'A1',
  onExcelInsertFormula,
  onExcelOpenAudit,
  pptxSlideCount = 1,
  onPptxAppendSlides,
  onPptxOpenOtherAiTab,
  onClose,
}: AIToolsDrawerProps): JSX.Element | null {
  const accent = ACCENT_BY_MODE[mode];

  const [apiKeyInput, setApiKeyInput] = useState<string>(() =>
    getConfiguredGeminiApiKey()
  );
  const [showKeyBox, setShowKeyBox] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [statusNotice, setStatusNotice] = useState<string | null>(null);
  const [copiedFlag, setCopiedFlag] = useState<boolean>(false);

  // Word state
  const [wordAction, setWordAction] = useState<WordAiMode>(
    initialAction === 'grammar' || initialAction === 'expand'
      ? initialAction
      : 'summary'
  );
  const [wordPromptHint, setWordPromptHint] = useState<string>('');
  const [wordOutputText, setWordOutputText] = useState<string>('');
  const [wordOutputHtml, setWordOutputHtml] = useState<string>('');
  const [wordHighlights, setWordHighlights] = useState<string[]>([]);
  const [wordSource, setWordSource] = useState<AiExecutionSource>('local');

  // Excel state
  const [excelPrompt, setExcelPrompt] = useState<string>(
    'Tính tổng cột D nếu cột B là 12A5'
  );
  const [excelFormulaResult, setExcelFormulaResult] = useState<{
    formula: string;
    explanation: string;
    source: AiExecutionSource;
  } | null>(null);

  // PowerPoint state
  const [pptxTopic, setPptxTopic] = useState<string>(documentTitle || '');
  const [pptxCount, setPptxCount] = useState<number>(5);
  const [pptxStyle, setPptxStyle] = useState<string>(SLIDE_STYLE_OPTIONS[0]);
  const [pptxPreviewSlides, setPptxPreviewSlides] = useState<
    AiGeneratedSlideDraft[]
  >([]);
  const [pptxOutlineText, setPptxOutlineText] = useState<string>('');
  const [pptxSource, setPptxSource] = useState<AiExecutionSource>('local');

  useEffect(() => {
    if (!isOpen) return;
    setApiKeyInput(getConfiguredGeminiApiKey());
    if (
      mode === 'word' &&
      (initialAction === 'summary' ||
        initialAction === 'grammar' ||
        initialAction === 'expand')
    ) {
      setWordAction(initialAction);
    }
    if (mode === 'pptx' && !pptxTopic.trim() && documentTitle) {
      setPptxTopic(documentTitle);
    }
  }, [isOpen, mode, initialAction, documentTitle]);

  // Close drawer cleanly on Escape and Android Hardware Back Button
  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.stopPropagation();
        onClose();
      }
    };
    const handleHardwareBack = (event: Event) => {
      event.preventDefault();
      onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('lite:hardware-back', handleHardwareBack);
    return () => {
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('lite:hardware-back', handleHardwareBack);
    };
  }, [isOpen, onClose]);

  const excelQuickStats = useMemo(() => {
    if (mode !== 'excel') return null;
    let nonEmpty = 0;
    let numericCount = 0;
    let numericSum = 0;
    for (const row of excelRows) {
      for (const cell of row) {
        const s = String(cell ?? '').trim();
        if (s !== '') {
          nonEmpty += 1;
          const n = Number(s);
          if (!Number.isNaN(n) && Number.isFinite(n)) {
            numericCount += 1;
            numericSum += n;
          }
        }
      }
    }
    return {
      rows: excelRows.length,
      cols: excelColumns.length,
      nonEmpty,
      numericCount,
      avg: numericCount > 0 ? (numericSum / numericCount).toFixed(2) : '0',
    };
  }, [mode, excelRows, excelColumns.length]);

  if (!isOpen) {
    return null;
  }

  const activeGeminiReady = isAiEndpointAvailable(apiKeyInput);

  const handleSaveApiKey = () => {
    saveConfiguredGeminiApiKey(apiKeyInput);
    setShowKeyBox(false);
    setStatusNotice(
      apiKeyInput.trim()
        ? 'Đã lưu khóa Gemini API cho AI Tools.'
        : 'Đã chuyển AI Tools sang chế độ cục bộ (Offline).'
    );
  };

  const handleRunWordAction = async (targetMode: WordAiMode) => {
    setWordAction(targetMode);
    setIsLoading(true);
    setStatusNotice(null);
    try {
      const res = await runWordAiAssistant({
        mode: targetMode,
        text: wordPlainText,
        selectedText: wordSelectedText,
        title: documentTitle,
        promptHint: wordPromptHint,
        apiKey: apiKeyInput,
      });
      setWordOutputText(res.text);
      setWordOutputHtml(res.html);
      setWordHighlights(res.highlights);
      setWordSource(res.source);
      if (res.notice) {
        setStatusNotice(res.notice);
      }
    } finally {
      setIsLoading(false);
    }
  };

  const handleRunExcelFormula = async () => {
    if (!excelPrompt.trim()) return;
    setIsLoading(true);
    setStatusNotice(null);
    try {
      const res = await generateExcelFormulaWithAi({
        prompt: excelPrompt,
        sheetName: excelSheetName,
        columns: excelColumns,
        rows: excelRows,
        activeCellRef: excelActiveCellRef,
        apiKey: apiKeyInput,
      });
      setExcelFormulaResult(res);
      if (res.notice) {
        setStatusNotice(res.notice);
      }
    } finally {
      setIsLoading(false);
    }
  };

  const handleRunPptxOutline = async () => {
    const cleanTopic = pptxTopic.trim() || documentTitle || 'Chủ đề mới';
    setIsLoading(true);
    setStatusNotice(null);
    try {
      const res = await generateSlideDeckWithAi({
        topic: cleanTopic,
        slideCount: pptxCount,
        style: pptxStyle,
        currentTitle: documentTitle,
        apiKey: apiKeyInput,
      });
      setPptxPreviewSlides(res.slides);
      setPptxOutlineText(res.outlineText);
      setPptxSource(res.source);
      if (res.notice) {
        setStatusNotice(res.notice);
      }
    } finally {
      setIsLoading(false);
    }
  };

  const handleCopyText = (value: string) => {
    if (!value) return;
    void navigator.clipboard.writeText(value);
    setCopiedFlag(true);
    window.setTimeout(() => setCopiedFlag(false), 1800);
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={`Trợ lý AI LITE — ${accent.label}`}
      className="fixed inset-0 z-50 flex justify-end bg-slate-950/45 backdrop-blur-[1.5px]"
      onClick={onClose}
    >
      <aside
        onClick={(event) => event.stopPropagation()}
        className={`flex h-full w-full max-w-[440px] flex-col border-l shadow-2xl transition-transform ${
          isDark
            ? 'border-white/10 bg-[#141923] text-slate-100'
            : 'border-slate-200 bg-white text-slate-900'
        }`}
      >
        {/* Header */}
        <header
          className={`shrink-0 border-b px-4 py-3.5 flex items-center justify-between gap-2 ${
            isDark ? 'border-white/10 bg-[#1B2230]' : 'border-slate-200 bg-[#F8FAFC]'
          }`}
        >
          <div className="flex min-w-0 items-center gap-2.5">
            <span
              className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl text-white shadow-sm"
              style={{ backgroundColor: accent.hex }}
            >
              <Sparkles className="h-4.5 w-4.5" aria-hidden="true" />
            </span>
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h2 className="truncate text-sm font-bold">
                  Trợ lý AI LITE
                </h2>
                <span
                  className={`rounded-full px-2 py-0.5 text-[10px] font-semibold ${
                    activeGeminiReady
                      ? 'bg-emerald-100 text-emerald-800'
                      : isDark
                        ? 'bg-white/10 text-slate-300'
                        : 'bg-slate-200/80 text-slate-700'
                  }`}
                >
                  {activeGeminiReady ? 'Gemini API' : 'Cục bộ Offline'}
                </span>
              </div>
              <p
                className={`truncate text-[11px] ${
                  isDark ? 'text-slate-400' : 'text-slate-500'
                }`}
              >
                {accent.label}
              </p>
            </div>
          </div>

          <div className="flex shrink-0 items-center gap-1">
            <button
              type="button"
              onClick={() => setShowKeyBox((prev) => !prev)}
              title="Cấu hình khóa Gemini API"
              className={`inline-flex h-8 items-center gap-1 rounded-lg border px-2 text-[11px] font-semibold transition-colors ${
                isDark
                  ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                  : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-100'
              }`}
            >
              <KeyRound className="h-3.5 w-3.5" aria-hidden="true" />
              <span>API Key</span>
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng Trợ lý AI"
              className="flex h-8 w-8 items-center justify-center rounded-lg text-slate-400 hover:bg-slate-200/60 hover:text-slate-800"
            >
              <X className="h-4 w-4" />
            </button>
          </div>
        </header>

        {/* Inline API Key configuration */}
        {showKeyBox && (
          <div
            className={`shrink-0 border-b px-4 py-3 space-y-2 ${
              isDark
                ? 'border-white/10 bg-[#10151F]'
                : 'border-slate-200 bg-indigo-50/50'
            }`}
          >
            <label className="block text-[11px] font-semibold">
              Khóa Gemini API cá nhân (bỏ trống để chạy bộ xử lý cục bộ)
            </label>
            <div className="flex items-center gap-2">
              <input
                type="password"
                value={apiKeyInput}
                onChange={(e) => setApiKeyInput(e.target.value)}
                placeholder="AIzaSy..."
                className={`min-w-0 flex-1 rounded-xl border px-3 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/30 ${
                  isDark
                    ? 'border-white/15 bg-[#1B2230] text-white'
                    : 'border-slate-200 bg-white text-slate-900'
                }`}
              />
              <button
                type="button"
                onClick={handleSaveApiKey}
                className="rounded-xl px-3 py-1.5 text-xs font-semibold text-white"
                style={{ backgroundColor: accent.hex }}
              >
                Lưu
              </button>
            </div>
          </div>
        )}

        {/* Main Scrollable Body */}
        <div className="flex-1 min-h-0 overflow-y-auto p-4 space-y-4">
          {statusNotice && (
            <div className="rounded-xl border border-amber-200 bg-amber-50 px-3 py-2 text-[11.5px] text-amber-900">
              {statusNotice}
            </div>
          )}

          {/* ================= WORD MODE ================= */}
          {mode === 'word' && (
            <>
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
                {[
                  {
                    id: 'summary' as const,
                    label: 'Tóm tắt văn bản',
                    icon: Sparkles,
                  },
                  {
                    id: 'grammar' as const,
                    label: 'Sửa chính tả',
                    icon: SpellCheck2,
                  },
                  {
                    id: 'expand' as const,
                    label: 'Viết tiếp đoạn văn',
                    icon: PenLine,
                  },
                ].map((btn) => {
                  const Icon = btn.icon;
                  const active = wordAction === btn.id;
                  return (
                    <button
                      key={btn.id}
                      type="button"
                      disabled={isLoading}
                      onClick={() => void handleRunWordAction(btn.id)}
                      className={`flex items-center justify-center gap-1.5 rounded-xl border px-2.5 py-2 text-xs font-semibold transition-all disabled:opacity-50 ${
                        active
                          ? 'text-white shadow-sm'
                          : isDark
                            ? 'border-white/15 bg-white/5 text-slate-200 hover:bg-white/10'
                            : 'border-slate-200 bg-slate-50 text-slate-700 hover:bg-white'
                      }`}
                      style={
                        active
                          ? { backgroundColor: accent.hex, borderColor: accent.hex }
                          : undefined
                      }
                    >
                      <Icon className="h-3.5 w-3.5 shrink-0" />
                      <span className="truncate">{btn.label}</span>
                    </button>
                  );
                })}
              </div>

              {wordSelectedText.trim().length > 0 && (
                <div
                  className={`rounded-xl border px-3 py-2 text-xs ${
                    isDark
                      ? 'border-blue-500/30 bg-blue-500/10 text-blue-200'
                      : 'border-blue-200 bg-[#EEF4FF] text-[#185ABD]'
                  }`}
                >
                  <p className="font-semibold">Đang chọn đoạn văn bản:</p>
                  <p className="mt-0.5 line-clamp-2 opacity-90">
                    &ldquo;{wordSelectedText.trim()}&rdquo;
                  </p>
                </div>
              )}

              {wordAction === 'expand' && (
                <div className="space-y-1.5">
                  <label className="block text-[11px] font-semibold opacity-75">
                    Gợi ý hướng viết tiếp (tùy chọn)
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      value={wordPromptHint}
                      onChange={(e) => setWordPromptHint(e.target.value)}
                      placeholder="Ví dụ: nhấn mạnh lợi ích tiết kiệm chi phí..."
                      className={`min-w-0 flex-1 rounded-xl border px-3 py-2 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/25 ${
                        isDark
                          ? 'border-white/15 bg-[#1B2230] text-white'
                          : 'border-slate-200 bg-white text-slate-900'
                      }`}
                    />
                    <button
                      type="button"
                      disabled={isLoading}
                      onClick={() => void handleRunWordAction('expand')}
                      className="shrink-0 rounded-xl px-3 py-2 text-xs font-semibold text-white disabled:opacity-50"
                      style={{ backgroundColor: accent.hex }}
                    >
                      Tạo
                    </button>
                  </div>
                </div>
              )}

              <button
                type="button"
                disabled={isLoading}
                onClick={() => void handleRunWordAction(wordAction)}
                className="inline-flex w-full items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-xs font-semibold text-white shadow-sm transition-opacity hover:opacity-95 disabled:opacity-60"
                style={{ backgroundColor: accent.hex }}
              >
                {isLoading ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Đang xử lý văn bản...</span>
                  </>
                ) : (
                  <>
                    <Wand2 className="h-4 w-4" />
                    <span>
                      {wordAction === 'summary'
                        ? 'Chạy Tóm tắt văn bản'
                        : wordAction === 'grammar'
                          ? 'Chạy Sửa chính tả & chuẩn hóa câu'
                          : 'Chạy Viết tiếp đoạn văn'}
                    </span>
                  </>
                )}
              </button>

              {wordOutputText && (
                <div
                  className={`rounded-2xl border p-3.5 space-y-3 ${
                    isDark
                      ? 'border-white/15 bg-[#1B2230]'
                      : 'border-slate-200 bg-[#F8FAFC]'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[11px] font-bold uppercase tracking-wide opacity-70">
                      Kết quả AI ({wordSource === 'gemini' ? 'Gemini AI' : 'Cục bộ'})
                    </span>
                    <button
                      type="button"
                      onClick={() => handleCopyText(wordOutputText)}
                      className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-[11px] font-medium hover:bg-black/5"
                    >
                      {copiedFlag ? (
                        <Check className="h-3.5 w-3.5 text-emerald-600" />
                      ) : (
                        <Copy className="h-3.5 w-3.5" />
                      )}
                      <span>{copiedFlag ? 'Đã chép' : 'Sao chép'}</span>
                    </button>
                  </div>

                  {wordHighlights.length > 0 && (
                    <div className="flex flex-wrap gap-1.5">
                      {wordHighlights.map((note, idx) => (
                        <span
                          key={`${note}-${idx}`}
                          className="rounded-md bg-blue-50 px-2 py-0.5 text-[10.5px] font-medium text-[#185ABD]"
                        >
                          {note}
                        </span>
                      ))}
                    </div>
                  )}

                  <textarea
                    value={wordOutputText}
                    onChange={(e) => {
                      setWordOutputText(e.target.value);
                      setWordOutputHtml('');
                    }}
                    rows={6}
                    className={`w-full rounded-xl border p-2.5 text-xs leading-relaxed focus:outline-none focus:ring-2 focus:ring-blue-500/25 ${
                      isDark
                        ? 'border-white/15 bg-[#141923] text-slate-100'
                        : 'border-slate-200 bg-white text-slate-900'
                    }`}
                  />

                  <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                    <button
                      type="button"
                      disabled={isReadOnly}
                      onClick={() =>
                        onWordInsertContent?.({
                          text: wordOutputText,
                          html:
                            wordOutputHtml ||
                            `<p>${wordOutputText
                              .split(/\r?\n+/)
                              .map((l) => l.trim())
                              .filter(Boolean)
                              .join('</p><p>')}</p>`,
                        })
                      }
                      className="inline-flex min-h-[40px] items-center justify-center gap-1.5 rounded-xl px-3 py-2 text-xs font-semibold text-white shadow-xs transition-opacity hover:opacity-95 disabled:opacity-45"
                      style={{ backgroundColor: accent.hex }}
                    >
                      <FilePlus2 className="h-3.5 w-3.5" />
                      <span>Chèn vào tài liệu</span>
                    </button>

                    <button
                      type="button"
                      disabled={isReadOnly || wordSelectedText.trim().length === 0}
                      title={
                        wordSelectedText.trim().length === 0
                          ? 'Hãy bôi đen đoạn văn bản trong Word trước'
                          : 'Thay thế đoạn văn đang chọn'
                      }
                      onClick={() =>
                        onWordReplaceSelection?.({
                          text: wordOutputText,
                          html:
                            wordOutputHtml ||
                            `<p>${wordOutputText
                              .split(/\r?\n+/)
                              .map((l) => l.trim())
                              .filter(Boolean)
                              .join('</p><p>')}</p>`,
                        })
                      }
                      className={`inline-flex min-h-[40px] items-center justify-center gap-1.5 rounded-xl border px-3 py-2 text-xs font-semibold transition-colors disabled:opacity-45 ${
                        isDark
                          ? 'border-white/15 bg-white/5 text-slate-100 hover:bg-white/10'
                          : 'border-blue-200 bg-white text-[#185ABD] hover:bg-blue-50'
                      }`}
                    >
                      <Replace className="h-3.5 w-3.5" />
                      <span>Thay thế vùng đang chọn</span>
                    </button>
                  </div>
                </div>
              )}
            </>
          )}

          {/* ================= EXCEL MODE ================= */}
          {mode === 'excel' && (
            <>
              <div className="rounded-xl border border-emerald-200 bg-[#ECFDF5] px-3 py-2.5 flex items-center justify-between gap-2 text-xs">
                <div className="flex items-center gap-2 min-w-0">
                  <Table2 className="h-4 w-4 shrink-0 text-[#107C41]" />
                  <span className="truncate font-semibold text-emerald-950">
                    Sheet: {excelSheetName}
                  </span>
                </div>
                <span className="shrink-0 rounded-lg bg-white px-2 py-0.5 font-mono text-[11px] font-bold text-[#107C41] shadow-2xs">
                  Ô chọn: {excelActiveCellRef}
                </span>
              </div>

              {excelQuickStats && (
                <div className="grid grid-cols-3 gap-2 text-center">
                  <div className="rounded-xl border border-slate-200 bg-slate-50 p-2">
                    <p className="text-[10.5px] text-slate-500">Số dòng</p>
                    <p className="text-sm font-bold text-slate-900 tabular-nums">
                      {excelQuickStats.rows}
                    </p>
                  </div>
                  <div className="rounded-xl border border-slate-200 bg-slate-50 p-2">
                    <p className="text-[10.5px] text-slate-500">Ô dữ liệu</p>
                    <p className="text-sm font-bold text-[#107C41] tabular-nums">
                      {excelQuickStats.nonEmpty}
                    </p>
                  </div>
                  <div className="rounded-xl border border-slate-200 bg-slate-50 p-2">
                    <p className="text-[10.5px] text-slate-500">TB giá trị số</p>
                    <p className="text-sm font-bold text-slate-900 tabular-nums">
                      {excelQuickStats.avg}
                    </p>
                  </div>
                </div>
              )}

              <div className="space-y-2">
                <label className="block text-xs font-semibold text-slate-700">
                  Mô tả công thức bằng tiếng Việt
                </label>
                <textarea
                  value={excelPrompt}
                  onChange={(e) => setExcelPrompt(e.target.value)}
                  rows={3}
                  placeholder='Ví dụ: Tính tổng cột D nếu cột B là 12A5'
                  className="w-full rounded-xl border border-slate-200 bg-white p-3 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#107C41] focus:outline-none focus:ring-2 focus:ring-[#107C41]/20"
                />
                <div className="flex flex-wrap gap-1.5">
                  {EXCEL_PROMPT_PRESETS.map((sample) => (
                    <button
                      key={sample}
                      type="button"
                      onClick={() => setExcelPrompt(sample)}
                      className="rounded-lg border border-emerald-200 bg-emerald-50/70 px-2 py-1 text-[11px] font-medium text-[#107C41] hover:bg-emerald-100"
                    >
                      {sample}
                    </button>
                  ))}
                </div>
              </div>

              <button
                type="button"
                disabled={isLoading || !excelPrompt.trim()}
                onClick={() => void handleRunExcelFormula()}
                className="inline-flex w-full items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-xs font-semibold text-white shadow-sm transition-opacity hover:opacity-95 disabled:opacity-55"
                style={{ backgroundColor: accent.hex }}
              >
                {isLoading ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Đang tạo công thức Excel...</span>
                  </>
                ) : (
                  <>
                    <Wand2 className="h-4 w-4" />
                    <span>Tạo công thức Excel</span>
                  </>
                )}
              </button>

              {excelFormulaResult && (
                <div className="rounded-2xl border border-emerald-200 bg-[#F8FAFC] p-3.5 space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[11px] font-bold uppercase tracking-wide text-slate-500">
                      Công thức đề xuất (
                      {excelFormulaResult.source === 'gemini' ? 'Gemini AI' : 'Cục bộ'})
                    </span>
                    <button
                      type="button"
                      onClick={() => handleCopyText(excelFormulaResult.formula)}
                      className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-[11px] font-medium text-slate-600 hover:bg-slate-200/70"
                    >
                      {copiedFlag ? (
                        <Check className="h-3.5 w-3.5 text-[#107C41]" />
                      ) : (
                        <Copy className="h-3.5 w-3.5" />
                      )}
                      <span>{copiedFlag ? 'Đã chép' : 'Sao chép'}</span>
                    </button>
                  </div>

                  <div className="rounded-xl border border-emerald-300 bg-white px-3 py-2.5 font-mono text-sm font-bold text-[#107C41] break-all">
                    {excelFormulaResult.formula}
                  </div>
                  <p className="text-xs leading-relaxed text-slate-600">
                    {excelFormulaResult.explanation}
                  </p>

                  <button
                    type="button"
                    disabled={isReadOnly}
                    onClick={() =>
                      onExcelInsertFormula?.(excelFormulaResult.formula)
                    }
                    className="inline-flex min-h-[42px] w-full items-center justify-center gap-1.5 rounded-xl bg-[#107C41] px-4 py-2 text-xs font-semibold text-white shadow-sm hover:bg-[#0B5E31] disabled:opacity-45"
                  >
                    <ArrowRight className="h-4 w-4" />
                    <span>Điền thẳng vào ô {excelActiveCellRef}</span>
                  </button>
                </div>
              )}

              {onExcelOpenAudit && (
                <button
                  type="button"
                  onClick={onExcelOpenAudit}
                  className="inline-flex w-full items-center justify-center gap-1.5 rounded-xl border border-amber-300 bg-amber-50 px-3 py-2 text-xs font-semibold text-amber-800 hover:bg-amber-100"
                >
                  <AlertTriangle className="h-3.5 w-3.5" />
                  <span>Kiểm tra ô lỗi công thức &amp; dữ liệu trùng lặp</span>
                </button>
              )}
            </>
          )}

          {/* ================= POWERPOINT MODE ================= */}
          {mode === 'pptx' && (
            <>
              <div className="space-y-3">
                <div>
                  <label className="mb-1 block text-xs font-semibold text-slate-700">
                    Chủ đề bài thuyết trình
                  </label>
                  <input
                    type="text"
                    value={pptxTopic}
                    onChange={(e) => setPptxTopic(e.target.value)}
                    placeholder="Ví dụ: Chiến lược chuyển đổi số doanh nghiệp 2026"
                    className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#C43E1C] focus:outline-none focus:ring-2 focus:ring-[#C43E1C]/20"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2.5">
                  <div>
                    <label className="mb-1 block text-xs font-semibold text-slate-700">
                      Số lượng slide
                    </label>
                    <div className="flex items-center gap-1.5">
                      {[3, 5, 8, 10].map((num) => (
                        <button
                          key={num}
                          type="button"
                          onClick={() => setPptxCount(num)}
                          className={`flex-1 rounded-lg border py-1.5 text-xs font-bold transition-colors ${
                            pptxCount === num
                              ? 'border-[#C43E1C] bg-[#C43E1C] text-white'
                              : 'border-slate-200 bg-white text-slate-700 hover:bg-orange-50'
                          }`}
                        >
                          {num}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="mb-1 block text-xs font-semibold text-slate-700">
                      Tùy chỉnh số slide
                    </label>
                    <input
                      type="number"
                      min={1}
                      max={15}
                      value={pptxCount}
                      onChange={(e) =>
                        setPptxCount(
                          Math.max(1, Math.min(15, Number(e.target.value) || 5))
                        )
                      }
                      className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-semibold text-slate-900 focus:border-[#C43E1C] focus:outline-none"
                    />
                  </div>
                </div>

                <div>
                  <label className="mb-1 block text-xs font-semibold text-slate-700">
                    Phong cách trình bày
                  </label>
                  <select
                    value={pptxStyle}
                    onChange={(e) => setPptxStyle(e.target.value)}
                    className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-900 focus:border-[#C43E1C] focus:outline-none"
                  >
                    {SLIDE_STYLE_OPTIONS.map((opt) => (
                      <option key={opt} value={opt}>
                        {opt}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <button
                type="button"
                disabled={isLoading || !pptxTopic.trim()}
                onClick={() => void handleRunPptxOutline()}
                className="inline-flex w-full items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-xs font-semibold text-white shadow-sm transition-opacity hover:opacity-95 disabled:opacity-55"
                style={{ backgroundColor: accent.hex }}
              >
                {isLoading ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    <span>Đang tạo dàn ý {pptxCount} slide...</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="h-4 w-4" />
                    <span>Sinh dàn ý chi tiết ({pptxCount} slide)</span>
                  </>
                )}
              </button>

              {pptxPreviewSlides.length > 0 && (
                <div className="rounded-2xl border border-orange-200 bg-[#FFFDFB] p-3.5 space-y-3">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-[11px] font-bold uppercase tracking-wide text-[#C43E1C]">
                      Dàn ý {pptxPreviewSlides.length} slide (
                      {pptxSource === 'gemini' ? 'Gemini AI' : 'Cục bộ'})
                    </span>
                    <button
                      type="button"
                      onClick={() => handleCopyText(pptxOutlineText)}
                      className="inline-flex items-center gap-1 rounded-lg px-2 py-1 text-[11px] font-medium text-slate-600 hover:bg-orange-50"
                    >
                      {copiedFlag ? (
                        <Check className="h-3.5 w-3.5 text-emerald-600" />
                      ) : (
                        <Copy className="h-3.5 w-3.5" />
                      )}
                      <span>{copiedFlag ? 'Đã chép' : 'Chép dàn ý'}</span>
                    </button>
                  </div>

                  <div className="max-h-64 space-y-2 overflow-y-auto pr-1">
                    {pptxPreviewSlides.map((slide, idx) => (
                      <div
                        key={`ai-slide-preview-${idx}`}
                        className="rounded-xl border border-slate-200 bg-white p-2.5 text-xs"
                      >
                        <div className="flex items-center gap-1.5 font-bold text-slate-900">
                          <span className="inline-flex h-5 w-5 shrink-0 items-center justify-center rounded-md bg-[#FFF4ED] text-[10px] font-bold text-[#C43E1C]">
                            {idx + 1}
                          </span>
                          <span className="truncate">{slide.title}</span>
                        </div>
                        <ul className="mt-1.5 space-y-0.5 pl-6 text-[11px] text-slate-600 list-disc">
                          {slide.bullets.map((b, bIdx) => (
                            <li key={bIdx}>{b}</li>
                          ))}
                        </ul>
                      </div>
                    ))}
                  </div>

                  <button
                    type="button"
                    disabled={isReadOnly}
                    onClick={() => {
                      onPptxAppendSlides?.(pptxPreviewSlides);
                      onClose();
                    }}
                    className="inline-flex min-h-[42px] w-full items-center justify-center gap-1.5 rounded-xl bg-[#C43E1C] px-4 py-2 text-xs font-semibold text-white shadow-sm hover:bg-[#9E3014] disabled:opacity-45"
                  >
                    <Layers className="h-4 w-4" />
                    <span>
                      Tạo {pptxPreviewSlides.length} slide vào bài trình chiếu (hiện có {pptxSlideCount})
                    </span>
                  </button>
                </div>
              )}

              {onPptxOpenOtherAiTab && (
                <div className="grid grid-cols-2 gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onPptxOpenOtherAiTab('summarize');
                    }}
                    className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-[11.5px] font-semibold text-slate-700 hover:bg-orange-50 hover:text-[#C43E1C]"
                  >
                    Tóm tắt toàn bài
                  </button>
                  <button
                    type="button"
                    disabled={isReadOnly}
                    onClick={() => {
                      onClose();
                      onPptxOpenOtherAiTab('notes');
                    }}
                    className="rounded-xl border border-slate-200 bg-slate-50 px-3 py-2 text-[11.5px] font-semibold text-slate-700 hover:bg-orange-50 hover:text-[#C43E1C] disabled:opacity-45"
                  >
                    Tạo Speaker Notes
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </aside>
    </div>
  );
}

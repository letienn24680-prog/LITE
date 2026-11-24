export type AiExecutionSource = 'gemini' | 'local';

export type WordAiMode = 'summary' | 'grammar' | 'expand';

export interface WordAiResult {
  text: string;
  html: string;
  highlights: string[];
  source: AiExecutionSource;
  notice?: string;
}

export interface ExcelFormulaAiResult {
  formula: string;
  explanation: string;
  source: AiExecutionSource;
  notice?: string;
}

export interface AiGeneratedSlideDraft {
  title: string;
  subtitle?: string;
  bullets: string[];
  speakerNote: string;
}

export interface SlideDeckAiResult {
  outlineText: string;
  slides: AiGeneratedSlideDraft[];
  source: AiExecutionSource;
  notice?: string;
}

const SETTINGS_STORAGE_KEY = 'lite_user_settings_v1';
const AI_PROXY_URL = (import.meta.env.VITE_AI_API_URL ?? '').trim();

export function getConfiguredGeminiApiKey(overrideKey?: string): string {
  if (overrideKey && overrideKey.trim().length > 0) {
    return overrideKey.trim();
  }
  if (typeof window !== 'undefined') {
    try {
      const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as { geminiApiKey?: unknown };
        if (
          typeof parsed.geminiApiKey === 'string' &&
          parsed.geminiApiKey.trim().length > 0
        ) {
          return parsed.geminiApiKey.trim();
        }
      }
    } catch {
      // Ignore localStorage parse failures
    }
  }
  return (import.meta.env.VITE_GEMINI_API_KEY ?? '').trim();
}

export function saveConfiguredGeminiApiKey(apiKey: string): void {
  if (typeof window === 'undefined') {
    return;
  }
  try {
    const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY);
    const prev =
      raw && raw.trim()
        ? (JSON.parse(raw) as Record<string, unknown>)
        : { autoSyncOnEdit: true, defaultWordTheme: 'light' };
    const next = {
      ...prev,
      geminiApiKey: apiKey.trim(),
    };
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(next));
  } catch {
    // Ignore localStorage failures
  }
}

export function isAiEndpointAvailable(overrideKey?: string): boolean {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    return false;
  }
  return Boolean(getConfiguredGeminiApiKey(overrideKey) || AI_PROXY_URL);
}

async function callGeminiEndpoint(
  prompt: string,
  options?: {
    apiKey?: string;
    jsonMode?: boolean;
    temperature?: number;
  }
): Promise<string> {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    throw new Error('Thiết bị đang ngoại tuyến.');
  }

  const key = getConfiguredGeminiApiKey(options?.apiKey);
  if (!key && !AI_PROXY_URL) {
    throw new Error('Chưa cấu hình Gemini API Key.');
  }

  const requestBody = {
    contents: [
      {
        role: 'user',
        parts: [{ text: prompt }],
      },
    ],
    generationConfig: {
      temperature: options?.temperature ?? 0.35,
      ...(options?.jsonMode ? { responseMimeType: 'application/json' } : {}),
    },
  };

  if (AI_PROXY_URL && !key) {
    const proxyRes = await fetch(AI_PROXY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody),
    });
    if (!proxyRes.ok) {
      throw new Error(`AI proxy HTTP ${proxyRes.status}`);
    }
    const proxyJson = (await proxyRes.json()) as {
      text?: string;
      candidates?: Array<{
        content?: { parts?: Array<{ text?: string }> };
      }>;
    };
    const proxyText =
      proxyJson.text ??
      proxyJson.candidates?.[0]?.content?.parts?.[0]?.text ??
      '';
    if (!proxyText.trim()) {
      throw new Error('Phản hồi AI rỗng.');
    }
    return proxyText.trim();
  }

  const models = ['gemini-2.0-flash', 'gemini-1.5-flash'];
  let lastError: Error | null = null;

  for (const model of models) {
    try {
      const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(
        key
      )}`;
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody),
      });
      if (!res.ok) {
        const errBody = await res.text().catch(() => '');
        lastError = new Error(
          `Gemini HTTP ${res.status}${errBody ? `: ${errBody.slice(0, 140)}` : ''}`
        );
        if (res.status === 404) {
          continue;
        }
        throw lastError;
      }
      const data = (await res.json()) as {
        candidates?: Array<{
          content?: { parts?: Array<{ text?: string }> };
        }>;
      };
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? '';
      if (!text) {
        throw new Error('Gemini trả về nội dung rỗng.');
      }
      return text;
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));
    }
  }

  throw lastError ?? new Error('Không thể kết nối Gemini API.');
}

function escapeHtml(raw: string): string {
  return raw
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function plainTextToWordHtml(text: string): string {
  const paragraphs = text
    .split(/\n{2,}|\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (paragraphs.length === 0) {
    return '<p></p>';
  }
  return paragraphs
    .map((line) => {
      if (line.startsWith('- ') || line.startsWith('• ')) {
        return `<p style="margin:0 0 6pt 14pt;line-height:1.6;">• ${escapeHtml(
          line.slice(2).trim()
        )}</p>`;
      }
      return `<p style="margin:0 0 8pt 0;line-height:1.6;">${escapeHtml(line)}</p>`;
    })
    .join('');
}

const VIETNAMESE_COMMON_NORMALIZATIONS: Array<[RegExp, string]> = [
  [/\bko\b/gi, 'không'],
  [/\bkhong\b/gi, 'không'],
  [/\bđc\b/gi, 'được'],
  [/\bdc\b/gi, 'được'],
  [/\bntn\b/gi, 'như thế nào'],
  [/\bvs\b/gi, 'với'],
  [/\bbh\b/gi, 'bây giờ'],
  [/\bhđ\b/gi, 'hoạt động'],
  [/\bsd\b/gi, 'sử dụng'],
  [/\bng\b/gi, 'người'],
  [/\btrc\b/gi, 'trước'],
  [/\bđt\b/gi, 'điện thoại'],
  [/\bcv\b/gi, 'công việc'],
  [/\bsl\b/gi, 'số lượng'],
  [/\btb\b/gi, 'trung bình'],
];

function runLocalVietnamesePolish(input: string): {
  polished: string;
  changes: string[];
} {
  let result = input.replace(/\r\n/g, '\n');
  const changes: string[] = [];

  for (const [pattern, replacement] of VIETNAMESE_COMMON_NORMALIZATIONS) {
    if (pattern.test(result)) {
      result = result.replace(pattern, (matched) =>
        matched.charAt(0) === matched.charAt(0).toUpperCase() &&
        matched.charAt(0) !== matched.charAt(0).toLowerCase()
          ? replacement.charAt(0).toUpperCase() + replacement.slice(1)
          : replacement
      );
      changes.push(`Chuẩn hóa từ viết tắt sang "${replacement}"`);
    }
  }

  const beforeSpaceFix = result;
  result = result
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+([,.;:?!])/g, '$1')
    .replace(/([,.;:?!])([^\s\d.,;:?!"')\]”])/g, '$1 $2')
    .trim();
  if (result !== beforeSpaceFix) {
    changes.push('Chuẩn hóa khoảng trắng và dấu câu');
  }

  const beforeCapFix = result;
  result = result.replace(
    /(^|[.!?]\s+|\n+)([\p{Ll}])/gu,
    (_m, prefix: string, ch: string) => `${prefix}${ch.toLocaleUpperCase('vi-VN')}`
  );
  if (result !== beforeCapFix) {
    changes.push('Viết hoa chữ cái đầu câu / đầu đoạn');
  }

  if (result.length > 0 && !/[.!?…:]$/.test(result)) {
    result = `${result}.`;
    changes.push('Bổ sung dấu chấm kết thúc câu');
  }

  return {
    polished: result,
    changes: changes.length > 0 ? Array.from(new Set(changes)) : ['Câu văn đã đúng chuẩn chính tả và dấu câu.'],
  };
}

function runLocalWordSummary(text: string, title: string): {
  summaryText: string;
  summaryHtml: string;
  highlights: string[];
} {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (!clean) {
    const fallback = `Tài liệu "${title}" hiện chưa có nội dung văn bản để tóm tắt.`;
    return {
      summaryText: fallback,
      summaryHtml: `<p>${escapeHtml(fallback)}</p>`,
      highlights: ['Chưa có văn bản'],
    };
  }

  const sentences = clean
    .split(/(?<=[.!?])\s+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 12);
  const words = clean.split(/\s+/).filter(Boolean);
  const topSentences =
    sentences.length <= 3
      ? sentences
      : [sentences[0], sentences[Math.floor(sentences.length / 2)], sentences[sentences.length - 1]];

  const bullets = topSentences.map((s) => `• ${s}`);
  const summaryText = `Tóm tắt tài liệu "${title}" (${words.length} từ):\n${bullets.join('\n')}`;
  const summaryHtml = [
    `<div style="background:#EEF4FF;border-left:4px solid #185ABD;padding:10pt 12pt;margin:0 0 12pt 0;border-radius:4pt;">`,
    `<p style="margin:0 0 6pt 0;font-weight:700;color:#185ABD;">Tóm tắt nội dung (${words.length} từ)</p>`,
    ...topSentences.map(
      (s) => `<p style="margin:0 0 4pt 0;line-height:1.55;">• ${escapeHtml(s)}</p>`
    ),
    `</div>`,
  ].join('');

  return {
    summaryText,
    summaryHtml,
    highlights: [`Tổng số từ: ${words.length}`, `${sentences.length} câu chính`],
  };
}

function runLocalWordContinuation(
  text: string,
  title: string,
  promptHint?: string
): string {
  const clean = text.trim();
  const lastSentence =
    clean
      .split(/(?<=[.!?])\s+/)
      .filter(Boolean)
      .pop() ?? '';
  const focusTopic = promptHint?.trim() || title || 'nội dung trên';

  if (!clean) {
    return `Trong phạm vi "${focusTopic}", mục tiêu trọng tâm là hệ thống hóa các dữ liệu nền tảng, đánh giá hiện trạng một cách khách quan và đề xuất phương án triển khai khả thi. Các bước thực hiện tiếp theo cần đảm bảo tiến độ rõ ràng, phân công trách nhiệm cụ thể và đo lường kết quả định kỳ.`;
  }

  return `Từ nền tảng trên (${lastSentence.slice(0, 90)}${
    lastSentence.length > 90 ? '...' : ''
  }), có thể mở rộng thêm về "${focusTopic}" theo ba khía cạnh: (1) chuẩn hóa quy trình phối hợp giữa các bên liên quan, (2) rà soát các chỉ số trọng yếu nhằm tối ưu nguồn lực, và (3) thiết lập mốc kiểm tra định kỳ để kịp thời điều chỉnh khi phát sinh biến động thực tế.`;
}

export async function runWordAiAssistant(params: {
  mode: WordAiMode;
  text: string;
  selectedText?: string;
  title: string;
  promptHint?: string;
  apiKey?: string;
}): Promise<WordAiResult> {
  const targetInput =
    params.selectedText && params.selectedText.trim().length > 0
      ? params.selectedText.trim()
      : params.text.trim();

  if (isAiEndpointAvailable(params.apiKey)) {
    try {
      let prompt = '';
      if (params.mode === 'summary') {
        prompt = `Bạn là trợ lý biên tập tiếng Việt chuyên nghiệp. Hãy tóm tắt ngắn gọn, súc tích văn bản sau (tiêu đề: "${params.title}") gồm 1 câu tổng quan và 3-5 gạch đầu dòng ý chính bằng tiếng Việt chuẩn:\n\n"""\n${targetInput.slice(0, 8000)}\n"""`;
      } else if (params.mode === 'grammar') {
        prompt = `Hãy sửa lỗi chính tả, dấu câu, cách dùng từ và chuẩn hóa câu văn tiếng Việt tự nhiên, mạch lạc cho đoạn văn bản sau. Chỉ trả về phần văn bản đã được chuẩn hóa (không kèm lời mở đầu dư thừa):\n\n"""\n${targetInput.slice(0, 6000)}\n"""`;
      } else {
        const hintLine = params.promptHint?.trim()
          ? `Yêu cầu bổ sung của người dùng: ${params.promptHint.trim()}\n`
          : '';
        prompt = `Bạn là trợ lý soạn thảo văn bản tiếng Việt. Hãy viết tiếp 1-2 đoạn văn mạch lạc, cùng văn phong với tài liệu "${params.title}". ${hintLine}Chỉ trả về phần đoạn văn viết tiếp:\n\nNgữ cảnh hiện có:\n"""\n${targetInput.slice(-4500)}\n"""`;
      }

      const geminiText = await callGeminiEndpoint(prompt, {
        apiKey: params.apiKey,
        temperature: params.mode === 'grammar' ? 0.15 : 0.45,
      });

      return {
        text: geminiText,
        html: plainTextToWordHtml(geminiText),
        highlights: [
          params.mode === 'summary'
            ? 'Tóm tắt bằng Gemini AI'
            : params.mode === 'grammar'
              ? 'Đã chuẩn hóa chính tả & văn phong'
              : 'Đã sinh đoạn văn tiếp theo',
        ],
        source: 'gemini',
      };
    } catch (err) {
      const reason = err instanceof Error ? err.message : 'Lỗi kết nối Gemini';
      // Fall back smoothly to local processor
      return runLocalWordFallback(params, targetInput, `Đã chuyển sang xử lý cục bộ (${reason}).`);
    }
  }

  return runLocalWordFallback(params, targetInput);
}

function runLocalWordFallback(
  params: {
    mode: WordAiMode;
    title: string;
    promptHint?: string;
  },
  targetInput: string,
  notice?: string
): WordAiResult {
  if (params.mode === 'summary') {
    const localSum = runLocalWordSummary(targetInput, params.title);
    return {
      text: localSum.summaryText,
      html: localSum.summaryHtml,
      highlights: localSum.highlights,
      source: 'local',
      notice,
    };
  }
  if (params.mode === 'grammar') {
    const { polished, changes } = runLocalVietnamesePolish(targetInput);
    return {
      text: polished,
      html: plainTextToWordHtml(polished),
      highlights: changes,
      source: 'local',
      notice,
    };
  }
  const continuation = runLocalWordContinuation(
    targetInput,
    params.title,
    params.promptHint
  );
  return {
    text: continuation,
    html: plainTextToWordHtml(continuation),
    highlights: ['Mở rộng ngữ cảnh cục bộ'],
    source: 'local',
    notice,
  };
}

function colIndexToLetter(index: number): string {
  let num = index + 1;
  let label = '';
  while (num > 0) {
    const rem = (num - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    num = Math.floor((num - 1) / 26);
  }
  return label;
}

function runLocalExcelFormulaParser(params: {
  prompt: string;
  columns: string[];
  rows: (string | number)[][];
}): { formula: string; explanation: string } {
  const q = params.prompt.trim();
  const lower = q.toLowerCase();
  const dataLastRow = Math.max(2, params.rows.length + 1);

  const colMatches = Array.from(q.matchAll(/\bcột\s+([A-Z]{1,2})\b/gi)).map((m) =>
    m[1].toUpperCase()
  );
  const rangeMatch = q.match(/\b([A-Z]{1,2}\d+:[A-Z]{1,2}\d+)\b/i);

  // Detect conditional clause: "nếu cột B là 12A5" / "khi cột C = Đạt" / "nếu B >= 8"
  const condColMatch = q.match(
    /(?:nếu|khi|với|theo|điều kiện)\s+(?:cột\s+)?([A-Z]{1,2})\s*(?:là|bằng|=|>=|<=|>|<|chứa)?\s*["']?([^"',;.\n]+)["']?/i
  );

  // Infer numeric columns from sheet if user omitted column letter
  const numericColLetters: string[] = [];
  for (let c = 0; c < params.columns.length; c += 1) {
    let hasNum = false;
    for (let r = 0; r < Math.min(params.rows.length, 30); r += 1) {
      const val = params.rows[r]?.[c];
      if (typeof val === 'number' || (typeof val === 'string' && val.trim() !== '' && !Number.isNaN(Number(val)))) {
        hasNum = true;
        break;
      }
    }
    if (hasNum) {
      numericColLetters.push(colIndexToLetter(c));
    }
  }
  const fallbackNumCol = numericColLetters[0] || colIndexToLetter(Math.max(0, params.columns.length - 1));

  if ((lower.includes('nếu') || lower.includes('điều kiện') || lower.includes('sumif')) && condColMatch) {
    const condCol = condColMatch[1].toUpperCase();
    const rawCondVal = condColMatch[2].trim();
    const isOperatorExpr = /^(>=|<=|<>|>|<|=)/.test(rawCondVal);
    const isPureNum = /^\d+(\.\d+)?$/.test(rawCondVal);
    const condLiteral =
      isOperatorExpr || isPureNum ? `"${rawCondVal}"` : `"${rawCondVal}"`;

    if (lower.includes('đếm') || lower.includes('bao nhiêu') || lower.includes('số lượng')) {
      return {
        formula: `=COUNTIF(${condCol}2:${condCol}${dataLastRow},${condLiteral})`,
        explanation: `Đếm số dòng trong ${condCol}2:${condCol}${dataLastRow} thỏa điều kiện ${rawCondVal}.`,
      };
    }

    if (lower.includes('trung bình') || lower.includes('tb')) {
      const targetCol = colMatches.find((c) => c !== condCol) ?? fallbackNumCol;
      return {
        formula: `=AVERAGEIF(${condCol}2:${condCol}${dataLastRow},${condLiteral},${targetCol}2:${targetCol}${dataLastRow})`,
        explanation: `Tính trung bình cột ${targetCol} cho các dòng có cột ${condCol} là "${rawCondVal}".`,
      };
    }

    const sumCol = colMatches.find((c) => c !== condCol) ?? fallbackNumCol;
    return {
      formula: `=SUMIF(${condCol}2:${condCol}${dataLastRow},${condLiteral},${sumCol}2:${sumCol}${dataLastRow})`,
      explanation: `Tính tổng cột ${sumCol} cho các dòng có cột ${condCol} là "${rawCondVal}".`,
    };
  }

  const targetCol = colMatches[0] ?? fallbackNumCol;
  const targetRange = rangeMatch
    ? rangeMatch[1].toUpperCase()
    : `${targetCol}2:${targetCol}${dataLastRow}`;

  if (lower.includes('trung bình') || lower.includes('bình quân') || lower.includes('average')) {
    return {
      formula: `=AVERAGE(${targetRange})`,
      explanation: `Tính giá trị trung bình của vùng ${targetRange}.`,
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
  if (lower.includes('đếm') || lower.includes('số lượng') || lower.includes('count')) {
    return {
      formula: `=COUNTA(${targetRange})`,
      explanation: `Đếm số ô có dữ liệu trong vùng ${targetRange}.`,
    };
  }

  return {
    formula: `=SUM(${targetRange})`,
    explanation: `Tính tổng các giá trị số trong vùng ${targetRange}.`,
  };
}

export async function generateExcelFormulaWithAi(params: {
  prompt: string;
  sheetName: string;
  columns: string[];
  rows: (string | number)[][];
  activeCellRef: string;
  apiKey?: string;
}): Promise<ExcelFormulaAiResult> {
  if (isAiEndpointAvailable(params.apiKey)) {
    try {
      const headerSummary = params.columns
        .slice(0, 12)
        .map((h, i) => `${colIndexToLetter(i)}: ${h}`)
        .join(' | ');
      const lastRow = Math.max(2, params.rows.length + 1);
      const sampleRows = params.rows
        .slice(0, 4)
        .map((r, idx) => `Dòng ${idx + 2}: ${JSON.stringify(r.slice(0, 10))}`)
        .join('\n');

      const aiPrompt = `Bạn là chuyên gia công thức Excel.
Trang tính "${params.sheetName}" có ${params.rows.length} dòng dữ liệu (từ dòng 2 đến dòng ${lastRow}).
Các cột: ${headerSummary}
Dữ liệu mẫu:
${sampleRows}
Ô đang chọn để điền công thức: ${params.activeCellRef}
Yêu cầu tiếng Việt của người dùng: "${params.prompt}"

Hãy trả về duy nhất JSON hợp lệ dạng:
{"formula":"=...","explanation":"Giải thích ngắn gọn bằng tiếng Việt"}
Lưu ý dùng đúng dấu phẩy (,) giữa các tham số công thức Excel chuẩn.`;

      const rawText = await callGeminiEndpoint(aiPrompt, {
        apiKey: params.apiKey,
        jsonMode: true,
        temperature: 0.1,
      });
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      const parsed = JSON.parse(jsonMatch ? jsonMatch[0] : rawText) as {
        formula?: string;
        explanation?: string;
      };
      const formula = (parsed.formula ?? '').trim();
      if (formula.startsWith('=')) {
        return {
          formula,
          explanation:
            parsed.explanation?.trim() || 'Công thức được sinh bởi Gemini AI.',
          source: 'gemini',
        };
      }
    } catch {
      // Fallback to deterministic Vietnamese NLP formula engine
    }
  }

  const localResult = runLocalExcelFormulaParser(params);
  return {
    formula: localResult.formula,
    explanation: localResult.explanation,
    source: 'local',
  };
}

function runLocalSlideDeckGenerator(
  topic: string,
  slideCount: number,
  style: string
): { outlineText: string; slides: AiGeneratedSlideDraft[] } {
  const safeCount = Math.max(1, Math.min(15, Math.round(slideCount) || 5));
  const cleanTopic = topic.trim() || 'Chủ đề thuyết trình mới';

  const sectionTemplates: Array<{
    title: string;
    subtitle: string;
    bullets: string[];
    speakerNote: string;
  }> = [
    {
      title: cleanTopic,
      subtitle: `${style} · Tổng quan & Mục tiêu`,
      bullets: [
        `Bối cảnh thực tế và lý do triển khai "${cleanTopic}"`,
        'Mục tiêu trọng tâm và phạm vi trình bày',
        'Đối tượng hưởng lợi & giá trị cốt lõi',
      ],
      speakerNote: `Giới thiệu tổng quan chủ đề ${cleanTopic} và nêu rõ mục tiêu chính của buổi trình bày.`,
    },
    {
      title: 'Hiện trạng & Thách thức đặt ra',
      subtitle: 'Phân tích bối cảnh',
      bullets: [
        'Đánh giá bức tranh tổng thể và điểm mạnh hiện có',
        'Các nút thắt, hạn chế cần tháo gỡ ngay',
        'Yêu cầu đổi mới để bắt kịp xu hướng',
      ],
      speakerNote: 'Trình bày những vấn đề cốt lõi đang tồn tại trước khi đi vào giải pháp.',
    },
    {
      title: 'Giải pháp & Phương án đề xuất',
      subtitle: 'Trọng tâm triển khai',
      bullets: [
        `Chiến lược tiếp cận bài bản cho ${cleanTopic}`,
        'Ứng dụng công nghệ và tối ưu hóa quy trình vận hành',
        'Phân bổ nguồn lực & trách nhiệm rõ ràng',
      ],
      speakerNote: 'Nhấn mạnh tính khả thi và giá trị khác biệt của giải pháp đề xuất.',
    },
    {
      title: 'Lộ trình & Kế hoạch thực hiện',
      subtitle: 'Các giai đoạn chính',
      bullets: [
        'Giai đoạn 1: Chuẩn bị nền tảng, dữ liệu và nhân sự',
        'Giai đoạn 2: Thử nghiệm, đo lường và tinh chỉnh',
        'Giai đoạn 3: Nhân rộng và đánh giá hiệu quả định kỳ',
      ],
      speakerNote: 'Đi qua từng mốc thời gian và tiêu chí nghiệm thu của mỗi giai đoạn.',
    },
    {
      title: 'Tổng kết & Hành động tiếp theo',
      subtitle: 'Kết luận',
      bullets: [
        `Tóm tắt 3 điểm cần nhớ về "${cleanTopic}"`,
        'Đề xuất bước hành động ngay trong tuần tới',
        'Hỏi đáp & Thảo luận cùng người nghe',
      ],
      speakerNote: 'Chốt lại thông điệp quan trọng nhất và mở phần Q&A.',
    },
  ];

  const slides: AiGeneratedSlideDraft[] = [];
  for (let i = 0; i < safeCount; i += 1) {
    if (i < sectionTemplates.length - 1) {
      slides.push(sectionTemplates[i]);
    } else if (i === safeCount - 1 && safeCount > 1) {
      slides.push(sectionTemplates[sectionTemplates.length - 1]);
    } else {
      const partIdx = i + 1;
      slides.push({
        title: `Phân tích chuyên sâu #${partIdx}: ${cleanTopic}`,
        subtitle: `Nội dung mở rộng (${style})`,
        bullets: [
          `Chỉ số đo lường hiệu quả trọng tâm cho phần ${partIdx}`,
          'Ví dụ minh họa thực tế & bài học kinh nghiệm',
          'Lưu ý quản trị rủi ro khi triển khai',
        ],
        speakerNote: `Trình bày chi tiết phần mở rộng ${partIdx} cho chủ đề ${cleanTopic}.`,
      });
    }
  }

  const outlineText = slides
    .map(
      (s, idx) =>
        `Slide ${idx + 1}: ${s.title}\n${s.bullets.map((b) => `  - ${b}`).join('\n')}`
    )
    .join('\n\n');

  return { outlineText, slides };
}

export async function generateSlideDeckWithAi(params: {
  topic: string;
  slideCount: number;
  style: string;
  currentTitle: string;
  apiKey?: string;
}): Promise<SlideDeckAiResult> {
  const safeCount = Math.max(1, Math.min(15, Math.round(params.slideCount) || 5));

  if (isAiEndpointAvailable(params.apiKey)) {
    try {
      const prompt = `Bạn là chuyên gia thiết kế bài thuyết trình PowerPoint.
Hãy tạo dàn ý chi tiết và nội dung cho đúng ${safeCount} slide bằng tiếng Việt về chủ đề: "${params.topic}".
Phong cách trình bày: "${params.style}". Tài liệu hiện tại: "${params.currentTitle}".

Trả về duy nhất JSON hợp lệ theo cấu trúc:
{
  "slides": [
    {
      "title": "Tiêu đề ngắn gọn",
      "subtitle": "Phụ đề 3-6 từ",
      "bullets": ["Ý chính 1", "Ý chính 2", "Ý chính 3"],
      "speakerNote": "Ghi chú ngắn cho người thuyết trình"
    }
  ]
}`;

      const rawText = await callGeminiEndpoint(prompt, {
        apiKey: params.apiKey,
        jsonMode: true,
        temperature: 0.5,
      });
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      const parsed = JSON.parse(jsonMatch ? jsonMatch[0] : rawText) as {
        slides?: Array<{
          title?: string;
          subtitle?: string;
          bullets?: string[];
          speakerNote?: string;
        }>;
      };
      if (Array.isArray(parsed.slides) && parsed.slides.length > 0) {
        const slides: AiGeneratedSlideDraft[] = parsed.slides
          .slice(0, safeCount)
          .map((s, idx) => ({
            title: s.title?.trim() || `Slide ${idx + 1}: ${params.topic}`,
            subtitle: s.subtitle?.trim() || undefined,
            bullets: Array.isArray(s.bullets) && s.bullets.length > 0
              ? s.bullets.map((b) => String(b).trim()).filter(Boolean)
              : [`Nội dung chính về ${params.topic}`],
            speakerNote: s.speakerNote?.trim() || '',
          }));
        const outlineText = slides
          .map(
            (s, idx) =>
              `Slide ${idx + 1}: ${s.title}\n${s.bullets
                .map((b) => `  - ${b}`)
                .join('\n')}`
          )
          .join('\n\n');

        return {
          outlineText,
          slides,
          source: 'gemini',
        };
      }
    } catch (err) {
      const reason = err instanceof Error ? err.message : 'Lỗi kết nối Gemini';
      const local = runLocalSlideDeckGenerator(params.topic, safeCount, params.style);
      return {
        ...local,
        source: 'local',
        notice: `Đã dùng bộ sinh dàn ý cục bộ (${reason}).`,
      };
    }
  }

  const local = runLocalSlideDeckGenerator(params.topic, safeCount, params.style);
  return {
    ...local,
    source: 'local',
  };
}

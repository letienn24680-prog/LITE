export type DocumentType = 'pptx' | 'docx' | 'xlsx' | 'pdf';

export type SidebarNavKey =
  | 'home'
  | 'slide'
  | 'word'
  | 'excel'
  | 'pdf'
  | 'folders'
  | 'shared';

export type FilterKey = 'all' | 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'favorite';

export type StageMode = 'present' | 'reader';

export type ReaderTheme = 'light' | 'dark';

export type WordPaperTheme = 'light' | 'sepia' | 'dark';

export type ModalType = 'present' | 'qr' | 'fullscreen' | 'share' | null;

export type SyncStatus = 'local' | 'idle' | 'syncing' | 'synced' | 'error';

export type UtilityModalType =
  | 'auth'
  | 'settings'
  | 'help'
  | 'logoutConfirm'
  | 'migrateLocal'
  | null;

export interface CloudUserSummary {
  id: string;
  email: string;
  displayName: string;
}

export interface UserSettings {
  autoSyncOnEdit: boolean;
  defaultWordTheme: WordPaperTheme;
}

export interface SlideTextRun {
  text: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontFamily?: string;
}

export interface SlideRichParagraph {
  align?: "left" | "center" | "right";
  bullet?: string;
  level?: number;
  runs: SlideTextRun[];
}

export interface SlideTableCellStyle {
  fillColor?: string;
  textColor?: string;
  bold?: boolean;
  align?: "left" | "center" | "right";
}

export type SlideShapeKind =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'arrowRight'
  | 'arrowLeft'
  | 'arrowUp'
  | 'arrowDown'
  | 'plus'
  | 'arrowDouble'
  | 'star'
  | 'callout'
  | 'line';

export interface SlideChartSeries {
  name: string;
  values: number[];
  color?: string;
}

export interface SlideChartData {
  chartType: 'column' | 'line' | 'pie';
  title?: string;
  categories: string[];
  series: SlideChartSeries[];
}

export interface SlideAnimationConfig {
  type: 'none' | 'appear' | 'fadeIn' | 'flyIn' | 'zoomIn' | 'pulse';
  order?: number;
}

export interface SlideTransitionConfig {
  type: 'none' | 'fade' | 'push' | 'wipe' | 'split' | 'zoom';
  durationMs?: number;
}

export interface SlideTextBlock {
  id: string;
  text: string;
  fontSize: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
  highlightColor?: string;
  textScript?: 'normal' | 'super' | 'sub';
  lineHeight?: number;
  textColumns?: 1 | 2;
  align?: 'left' | 'center' | 'right' | 'justify';
  role?: 'title' | 'subtitle' | 'body';
  xPct?: number;
  yPct?: number;
  wPct?: number;
  hPct?: number;
  fillColor?: string;
  borderColor?: string;
  fontFamily?: string;
  valign?: 'top' | 'middle' | 'bottom';
  kind?: 'text' | 'table' | 'shape' | 'chart' | 'media';
  shapeKind?: SlideShapeKind;
  chartData?: SlideChartData;
  mediaSrc?: string;
  mediaKind?: 'video' | 'audio';
  linkUrl?: string;
  groupId?: string;
  zIndex?: number;
  animation?: SlideAnimationConfig;
  tableRows?: string[][];
  tableCellStyles?: SlideTableCellStyle[][];
  richParagraphs?: SlideRichParagraph[];
  fillGradient?: string;
  borderWidthPx?: number;
  borderRadius?: string;
  rotationDeg?: number;
  shadowCss?: string;
  paddingPct?: { top: number; right: number; bottom: number; left: number };
  fontScalePct?: number;
}

export interface SlideDrawingStroke {
  id: string;
  tool: 'pen' | 'highlighter';
  color: string;
  width: number;
  points: Array<{ x: number; y: number }>;
}

export interface SlideCommentItem {
  id: string;
  text: string;
  author: string;
  createdAt: string;
}

export interface SlideImageItem {
  id: string;
  src: string;
  alt?: string;
  xPct?: number;
  yPct?: number;
  wPct?: number;
  hPct?: number;
  rotationDeg?: number;
  borderRadius?: string;
  cropRect?: { l: number; t: number; r: number; b: number };
}

export interface SlidePreviewItem {
  id: string;
  index: number;
  title: string;
  subtitle?: string;
  bullets: string[];
  metricHighlight?: {
    label: string;
    value: string;
  };
  accentColor: string;
  speakerNote?: string;
  blocks?: SlideTextBlock[];
  images?: SlideImageItem[];
  drawings?: SlideDrawingStroke[];
  comments?: SlideCommentItem[];
  footerText?: string;
  showSlideNumber?: boolean;
  backgroundColor?: string;
  backgroundGradient?: string;
  backgroundImageSrc?: string;
  slideWidthRatio?: number;
  isMasterPositioned?: boolean;
  hidden?: boolean;
  transition?: SlideTransitionConfig;
}

export interface OutlineSection {
  id: string;
  romanIndex: string;
  title: string;
  summary: string;
  keyPoints: string[];
  pageNumber: number;
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontFamily?: string;
  textColor?: string;
  bgColor?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  valign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numFmt?: string;
  textRotation?: 'horizontal' | 'vertical' | 'up45' | 'down45';
  hyperlink?: string;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
}

export interface SpreadsheetMergeRange {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

export interface ExcelConditionalFormat {
  id: string;
  range: { minRow: number; maxRow: number; minCol: number; maxCol: number };
  rule: 'gt' | 'lt' | 'eq' | 'duplicate';
  value?: string;
  bgColor: string;
  textColor?: string;
}

export interface ExcelChartSeries {
  name: string;
  values: number[];
  color: string;
}

export interface ExcelChartItem {
  id: string;
  type: 'column' | 'line' | 'pie';
  title: string;
  categories: string[];
  series: ExcelChartSeries[];
}

export interface ExcelDrawingStroke {
  id: string;
  tool: 'pen' | 'highlighter';
  color: string;
  width: number;
  points: Array<{ x: number; y: number }>;
}

export interface ExcelFloatingObject {
  id: string;
  kind: 'image' | 'shape' | 'textbox';
  x: number;
  y: number;
  width: number;
  height: number;
  src?: string;
  text?: string;
  shapeKind?: 'rect' | 'roundRect' | 'ellipse' | 'arrowRight' | 'star' | 'callout';
  fillColor?: string;
  borderColor?: string;
}

export interface ExcelPageSetup {
  margins: 'normal' | 'narrow' | 'wide';
  orientation: 'portrait' | 'landscape';
  paperSize: 'A4' | 'A5' | 'Letter';
  printArea?: string;
  printGridlines?: boolean;
}

export interface SpreadsheetSheet {
  id: string;
  name: string;
  columns: string[];
  rows: (string | number)[][];
  cellStyles?: Record<string, CellStyle>;
  merges?: SpreadsheetMergeRange[];
  colWidths?: number[];
  rowHeights?: number[];
  frozenRows?: number;
  frozenCols?: number;
  tabColor?: string;
  hiddenRows?: number[];
  hiddenCols?: number[];
  autoFilterEnabled?: boolean;
  dataValidations?: Record<string, string[]>;
  conditionalFormats?: ExcelConditionalFormat[];
  charts?: ExcelChartItem[];
  cellComments?: Record<string, string>;
  drawings?: ExcelDrawingStroke[];
  floatingObjects?: ExcelFloatingObject[];
  pageSetup?: ExcelPageSetup;
  isProtected?: boolean;
  lockedRanges?: string[];
  showGridlines?: boolean;
}

export interface WordCommentItem {
  id: string;
  quote: string;
  text: string;
  author: string;
  createdAt: string;
}

export interface WordDrawingStroke {
  id: string;
  tool: 'pen' | 'highlighter';
  color: string;
  width: number;
  points: Array<{ x: number; y: number }>;
}

export interface PdfPageItem {
  id: string;
  pageNumber: number;
  sectionTitle: string;
  heading: string;
  paragraphs: string[];
  callout?: string;
  highlights?: string[];
}

export interface NormalizedPoint {
  x: number;
  y: number;
}

export interface AnnotationStroke {
  id: string;
  slideIndex: number;
  color: string;
  width: number;
  isHighlighter?: boolean;
  points: NormalizedPoint[];
}

export interface SpotlightState {
  active: boolean;
  x: number;
  y: number;
  radius: number;
}

export type RemoteCommandAction =
  | 'next'
  | 'prev'
  | 'goto'
  | 'laserMove'
  | 'laserOff'
  | 'spotlightMove'
  | 'spotlightOff'
  | 'clearAnnotations'
  | 'toggleTimer'
  | 'resetTimer'
  | 'requestState';

export interface RemoteChannelState {
  roomCode: string;
  documentId: string;
  documentTitle: string;
  slideIndex: number;
  totalSlides: number;
  slideTitle: string;
  speakerNote: string;
  elapsedSeconds: number;
  isTimerRunning: boolean;
  annotationsCount: number;
}

export type ShareAccessLevel = 'private' | 'link' | 'password';

export interface ShareConfigRecord {
  shareToken: string;
  documentId: string;
  accessLevel: ShareAccessLevel;
  allowDownload: boolean;
  r2ObjectKey?: string;
  storagePath?: string;
  viewCount: number;
  revokedAt?: string | null;
  updatedAt: string;
}

export type SharedViewStatus =
  | 'loading'
  | 'ready'
  | 'password_required'
  | 'private'
  | 'revoked'
  | 'not_found'
  | 'error';

export interface SharedDocumentResolution {
  status: SharedViewStatus;
  shareToken: string;
  title?: string;
  fileType?: DocumentType;
  fileName?: string;
  fileSize?: string;
  allowDownload?: boolean;
  viewCount?: number;
  r2ObjectKey?: string;
  storagePath?: string;
  downloadUrl?: string;
  document?: DocumentItem;
  errorMessage?: string;
}

export interface DocumentItem {
  id: string;
  title: string;
  type: DocumentType;
  updatedAt: string;
  isFavorite: boolean;
  thumbnail: string;
  slidesPreview: SlidePreviewItem[];
  outline: OutlineSection[];
  contentPreview: string;
  pageCount?: number;
  fileSize?: string;
  folder?: string;
  isShared?: boolean;
  sourceFile?: File | null;
  objectUrl?: string;
  sheets?: SpreadsheetSheet[];
  pdfPages?: PdfPageItem[];
  wordParagraphs?: Record<string, string[]>;
  fileUrl?: string;
  wordHtml?: string;
  isUploaded?: boolean;
  lastSavedAt?: string;
  storagePath?: string;
  originalFileName?: string;
  lastAccessedAt?: string;
  isSyncedToCloud?: boolean;
  userId?: string;
  r2ObjectKey?: string;
  shareToken?: string;
  shareAccess?: ShareAccessLevel;
  shareAllowDownload?: boolean;
  rawArrayBuffer?: ArrayBuffer;
  originalMimeType?: string;
  hasLocalEdits?: boolean;
  slideAspectRatio?: number;
  collabPermission?: 'owner' | 'edit' | 'view';
  sharedOwnerId?: string;
  wordComments?: WordCommentItem[];
  wordDrawings?: WordDrawingStroke[];
  wordProtected?: boolean;
  slideProtected?: boolean;
}

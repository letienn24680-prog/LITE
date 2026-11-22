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

export interface SlideTextBlock {
  id: string;
  text: string;
  fontSize: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: 'left' | 'center' | 'right';
  role?: 'title' | 'subtitle' | 'body';
  xPct?: number;
  yPct?: number;
  wPct?: number;
  hPct?: number;
  fillColor?: string;
  borderColor?: string;
  fontFamily?: string;
  valign?: 'top' | 'middle' | 'bottom';
  kind?: 'text' | 'table';
  tableRows?: string[][];
}

export interface SlideImageItem {
  id: string;
  src: string;
  alt?: string;
  xPct?: number;
  yPct?: number;
  wPct?: number;
  hPct?: number;
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
  backgroundColor?: string;
  backgroundGradient?: string;
  backgroundImageSrc?: string;
  slideWidthRatio?: number;
  isMasterPositioned?: boolean;
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
  align?: 'left' | 'center' | 'right';
}

export interface SpreadsheetSheet {
  id: string;
  name: string;
  columns: string[];
  rows: (string | number)[][];
  cellStyles?: Record<string, CellStyle>;
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
  slideAspectRatio?: number;
  collabPermission?: 'owner' | 'edit' | 'view';
  sharedOwnerId?: string;
}

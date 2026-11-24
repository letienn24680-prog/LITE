export interface LoadingWorkspaceSkeletonProps {
  label?: string;
}

export function LoadingWorkspaceSkeleton({
  label = 'Đang tải không gian làm việc...',
}: LoadingWorkspaceSkeletonProps) {
  return (
    <div
      role="status"
      aria-live="polite"
      className="h-[100dvh] w-full flex flex-col bg-[#F8FAFC] text-slate-800 overflow-hidden"
    >
      {/* Top Bar Skeleton */}
      <div className="h-13 shrink-0 border-b border-slate-200/80 bg-white px-4 flex items-center justify-between gap-3">
        <div className="flex items-center gap-3 min-w-0">
          <div className="h-8 w-8 rounded-xl bg-slate-200/80 animate-pulse" />
          <div className="h-4 w-44 max-w-[45vw] rounded-lg bg-slate-200/80 animate-pulse" />
        </div>
        <div className="flex items-center gap-2">
          <div className="h-8 w-20 rounded-xl bg-slate-200/80 animate-pulse hidden sm:block" />
          <div className="h-8 w-8 rounded-xl bg-indigo-100 animate-pulse" />
        </div>
      </div>

      {/* Ribbon Skeleton */}
      <div className="shrink-0 border-b border-slate-200/80 bg-white/90 px-4 py-2.5 space-y-2">
        <div className="flex items-center gap-3 overflow-hidden">
          <div className="h-3.5 w-14 rounded bg-slate-200/80 animate-pulse" />
          <div className="h-3.5 w-20 rounded bg-indigo-200/80 animate-pulse" />
          <div className="h-3.5 w-14 rounded bg-slate-200/70 animate-pulse" />
          <div className="h-3.5 w-16 rounded bg-slate-200/70 animate-pulse" />
        </div>
        <div className="h-10 w-full rounded-xl bg-slate-100 border border-slate-200/70 animate-pulse" />
      </div>

      {/* Main Canvas Skeleton */}
      <div className="flex-1 min-h-0 flex items-center justify-center p-4 sm:p-8 bg-[#F1F3F5]">
        <div className="w-full max-w-2xl rounded-2xl border border-slate-200/80 bg-white p-6 sm:p-10 shadow-sm space-y-5">
          <div className="flex items-center gap-3">
            <div className="h-5 w-5 rounded-full border-2 border-indigo-200 border-t-indigo-600 animate-spin shrink-0" />
            <span className="text-sm font-medium text-slate-600">{label}</span>
          </div>
          <div className="space-y-3">
            <div className="h-6 w-2/3 rounded-lg bg-slate-200/75 animate-pulse" />
            <div className="h-3.5 w-full rounded bg-slate-100 animate-pulse" />
            <div className="h-3.5 w-5/6 rounded bg-slate-100 animate-pulse" />
            <div className="h-3.5 w-4/5 rounded bg-slate-100 animate-pulse" />
          </div>
          <div className="h-44 w-full rounded-xl bg-slate-100/80 animate-pulse" />
        </div>
      </div>
    </div>
  );
}

export default LoadingWorkspaceSkeleton;

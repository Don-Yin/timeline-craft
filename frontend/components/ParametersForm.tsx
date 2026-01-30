'use client';

import { useState } from 'react';

export function ParametersForm({
  sidebarWidth,
  onSidebarWidthChange,
  itemHeight,
  onItemHeightChange,
  duration,
  onDurationChange,
  applyMorph,
  onApplyMorphChange,
  onDownload,
  isDownloading,
  canDownload,
  progress,
  progressMessage,
}: {
  sidebarWidth: number;
  onSidebarWidthChange: (value: number) => void;
  itemHeight: number;
  onItemHeightChange: (value: number) => void;
  duration: number;
  onDurationChange: (value: number) => void;
  applyMorph: boolean;
  onApplyMorphChange: (value: boolean) => void;
  onDownload: () => void;
  isDownloading: boolean;
  canDownload: boolean;
  progress: number;
  progressMessage: string;
}) {
  const [showAdvanced, setShowAdvanced] = useState(false);

  return (
    <section className="rounded-xl border p-5">
      <h2 className="mb-3 text-base font-medium">configure parameters</h2>

      <div className="flex flex-col gap-4 text-sm">
        <label className="flex flex-col gap-1">
          <span className="flex items-center justify-between text-zinc-700 dark:text-zinc-300">
            <span>sidebar width (%)</span>
            <span className="text-xs text-zinc-500">{sidebarWidth}%</span>
          </span>
          <input
            type="range"
            min={5}
            max={30}
            step={1}
            value={sidebarWidth}
            onChange={(e) => onSidebarWidthChange(Number(e.target.value))}
            className="w-full"
          />
        </label>
        <label className="flex flex-col gap-1">
          <span className="flex items-center justify-between text-zinc-700 dark:text-zinc-300">
            <span>item height (%)</span>
            <span className="text-xs text-zinc-500">{itemHeight}%</span>
          </span>
          <input
            type="range"
            min={5}
            max={30}
            step={1}
            value={itemHeight}
            onChange={(e) => onItemHeightChange(Number(e.target.value))}
            className="w-full"
          />
        </label>
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={applyMorph}
            onChange={(e) => onApplyMorphChange(e.target.checked)}
          />
          <span className="text-zinc-700 dark:text-zinc-300">apply morph transition</span>
        </label>
      </div>

      <details className="mt-4" open={showAdvanced} onToggle={(e) => setShowAdvanced((e.target as HTMLDetailsElement).open)}>
        <summary className="cursor-pointer text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300">
          advanced settings
        </summary>
        <div className="mt-3 flex flex-col gap-3 text-sm">
          <label className="flex flex-col gap-1">
            <span className="text-zinc-700 dark:text-zinc-300">transition duration (s)</span>
            <input
              type="number"
              min={0}
              max={2}
              step={0.1}
              value={duration}
              onChange={(e) => onDurationChange(Number(e.target.value))}
              className="rounded-md border px-3 py-2"
            />
          </label>
        </div>
      </details>

      <button
        className="mt-6 w-full rounded-full bg-emerald-600 text-white transition-all hover:bg-emerald-700 disabled:opacity-60 disabled:cursor-not-allowed relative overflow-hidden"
        onClick={onDownload}
        disabled={!canDownload || isDownloading}
        style={{ height: '48px' }}
      >
        {isDownloading ? (
          <div className="relative w-full h-full">
            <div
              className={`absolute inset-0 bg-emerald-500 transition-all duration-300 ease-out ${progress >= 100 ? 'animate-pulse' : ''}`}
              style={{ width: `${progress}%` }}
            />
            <div className="absolute inset-0 flex flex-col items-center justify-center">
              {progress >= 100 ? (
                <>
                  <div className="flex items-center gap-2">
                    <div className="h-3 w-3 animate-spin rounded-full border-2 border-white border-t-transparent" />
                    <span className="text-sm font-medium">finalizing</span>
                  </div>
                  <span className="text-xs opacity-80">{progressMessage}</span>
                </>
              ) : (
                <>
                  <span className="text-sm font-medium">{progress}%</span>
                  <span className="text-xs opacity-80 truncate max-w-[90%]">{progressMessage}</span>
                </>
              )}
            </div>
          </div>
        ) : (
          <span className="relative z-10">download .pptx</span>
        )}
      </button>
    </section>
  );
}

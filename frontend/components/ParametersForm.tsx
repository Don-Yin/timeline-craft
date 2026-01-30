'use client';

import { useState } from 'react';
import { ColorPickerWithPresets } from './ColorPickerWithPresets';

export function ParametersForm({
  sidebarWidth,
  onSidebarWidthChange,
  itemHeight,
  onItemHeightChange,
  duration,
  onDurationChange,
  applyMorph,
  onApplyMorphChange,
  sidebarColorHex,
  onSidebarColorHexChange,
  indicatorColorHex,
  onIndicatorColorHexChange,
  sidebarFontColorHex,
  onSidebarFontColorHexChange,
  sidebarTransparency,
  onSidebarTransparencyChange,
  fontSize,
  onFontSizeChange,
  verticallyCenter,
  onVerticallyCenterChange,
  roundedIndicator,
  onRoundedIndicatorChange,
  centerText,
  onCenterTextChange,
  compactIndicator,
  onCompactIndicatorChange,
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
  sidebarColorHex: string;
  onSidebarColorHexChange: (value: string) => void;
  indicatorColorHex: string;
  onIndicatorColorHexChange: (value: string) => void;
  sidebarFontColorHex: string;
  onSidebarFontColorHexChange: (value: string) => void;
  sidebarTransparency: number;
  onSidebarTransparencyChange: (value: number) => void;
  fontSize: number;
  onFontSizeChange: (value: number) => void;
  verticallyCenter: boolean;
  onVerticallyCenterChange: (value: boolean) => void;
  roundedIndicator: boolean;
  onRoundedIndicatorChange: (value: boolean) => void;
  centerText: boolean;
  onCenterTextChange: (value: boolean) => void;
  compactIndicator: boolean;
  onCompactIndicatorChange: (value: boolean) => void;
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
      </div>

      <details className="mt-4" open={showAdvanced} onToggle={(e) => setShowAdvanced((e.target as HTMLDetailsElement).open)}>
        <summary className="cursor-pointer text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300">
          advanced settings
        </summary>
        <div className="mt-3 flex flex-col gap-4 text-sm">
          {/* Toggles Section */}
          <div className="space-y-2.5 pb-3 border-b border-zinc-100 dark:border-zinc-800">
            <label className="flex items-center gap-2 cursor-pointer group">
              <input
                type="checkbox"
                checked={applyMorph}
                onChange={(e) => onApplyMorphChange(e.target.checked)}
                className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-zinc-700 dark:text-zinc-300 group-hover:text-zinc-900 dark:group-hover:text-zinc-100 transition-colors">morph transition</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer group">
              <input
                type="checkbox"
                checked={verticallyCenter}
                onChange={(e) => onVerticallyCenterChange(e.target.checked)}
                className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-zinc-700 dark:text-zinc-300 group-hover:text-zinc-900 dark:group-hover:text-zinc-100 transition-colors">center tags vertically</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer group">
              <input
                type="checkbox"
                checked={roundedIndicator}
                onChange={(e) => onRoundedIndicatorChange(e.target.checked)}
                className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-zinc-700 dark:text-zinc-300 group-hover:text-zinc-900 dark:group-hover:text-zinc-100 transition-colors">rounded indicator corners</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer group">
              <input
                type="checkbox"
                checked={centerText}
                onChange={(e) => onCenterTextChange(e.target.checked)}
                className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-zinc-700 dark:text-zinc-300 group-hover:text-zinc-900 dark:group-hover:text-zinc-100 transition-colors">center text in indicator</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer group">
              <input
                type="checkbox"
                checked={compactIndicator}
                onChange={(e) => onCompactIndicatorChange(e.target.checked)}
                className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
              />
              <span className="text-zinc-700 dark:text-zinc-300 group-hover:text-zinc-900 dark:group-hover:text-zinc-100 transition-colors">compact indicator (narrower than sidebar)</span>
            </label>
          </div>

          {/* Colors Section */}
          <div className="space-y-3 pb-3 border-b border-zinc-100 dark:border-zinc-800">
            <ColorPickerWithPresets
              label="sidebar color"
              value={sidebarColorHex}
              onChange={onSidebarColorHexChange}
            />
            <ColorPickerWithPresets
              label="indicator color"
              value={indicatorColorHex}
              onChange={onIndicatorColorHexChange}
            />
            <ColorPickerWithPresets
              label="font color"
              value={sidebarFontColorHex}
              onChange={onSidebarFontColorHexChange}
            />
          </div>

          {/* Sliders Section */}
          <div className="space-y-3">
            <label className="flex flex-col gap-1">
              <span className="flex items-center justify-between text-zinc-700 dark:text-zinc-300">
                <span>sidebar transparency</span>
                <span className="text-xs text-zinc-500">{sidebarTransparency}%</span>
              </span>
              <input
                type="range"
                min={0}
                max={100}
                step={5}
                value={sidebarTransparency}
                onChange={(e) => onSidebarTransparencyChange(Number(e.target.value))}
                className="w-full"
              />
            </label>
            <label className="flex flex-col gap-1">
              <span className="flex items-center justify-between text-zinc-700 dark:text-zinc-300">
                <span>font size</span>
                <span className="text-xs text-zinc-500">{fontSize}pt</span>
              </span>
              <input
                type="range"
                min={8}
                max={24}
                step={1}
                value={fontSize}
                onChange={(e) => onFontSizeChange(Number(e.target.value))}
                className="w-full"
              />
            </label>
            <label className="flex flex-col gap-1">
              <span className="flex items-center justify-between text-zinc-700 dark:text-zinc-300">
                <span>transition duration</span>
                <span className="text-xs text-zinc-500">{duration}s</span>
              </span>
              <input
                type="range"
                min={0}
                max={2}
                step={0.1}
                value={duration}
                onChange={(e) => onDurationChange(Number(e.target.value))}
                className="w-full"
              />
            </label>
          </div>
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

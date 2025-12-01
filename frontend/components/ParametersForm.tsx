'use client';

export function ParametersForm({
  sidebarWidth,
  onSidebarWidthChange,
  itemHeight,
  onItemHeightChange,
  duration,
  onDurationChange,
  applyMorph,
  onApplyMorphChange,
}: {
  sidebarWidth: number;
  onSidebarWidthChange: (value: number) => void;
  itemHeight: number;
  onItemHeightChange: (value: number) => void;
  duration: number;
  onDurationChange: (value: number) => void;
  applyMorph: boolean;
  onApplyMorphChange: (value: boolean) => void;
}) {
  return (
    <section className="rounded-xl border p-5">
      <h2 className="mb-3 text-base font-medium">configure parameters</h2>

      <div className="grid grid-cols-2 gap-4 text-sm">
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
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={applyMorph}
            onChange={(e) => onApplyMorphChange(e.target.checked)}
          />
          <span className="text-zinc-700 dark:text-zinc-300">apply morph transition</span>
        </label>
      </div>

      <button
        className="mt-6 w-full rounded-full bg-foreground px-5 py-3 text-background transition-colors hover:bg-[#383838] dark:hover:bg-[#ccc] disabled:opacity-60"
        disabled
        title="Processing pipeline not connected yet"
      >
        process (coming soon)
      </button>
      <p className="mt-2 text-xs text-zinc-500">
        the processing pipeline (api gateway, upload, worker) is not wired yet in this environment.
      </p>
    </section>
  );
}



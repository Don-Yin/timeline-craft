'use client';
import { cn } from "@/lib/utils";

export type Step = {
  id: string;
  label: string;
  completed: boolean;
};

export function StepTabs({
  steps,
  activeId,
  onChange,
}: {
  steps: Step[];
  activeId: string;
  onChange: (id: string) => void;
}) {
  return (
    <div className="flex gap-1 p-1 rounded-lg bg-zinc-100 dark:bg-zinc-900 w-fit">
      {steps.map((step, index) => {
        const isActive = activeId === step.id;
        return (
          <button
            key={step.id}
            onClick={() => onChange(step.id)}
            className={cn(
              "relative flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-md transition-all duration-200",
              isActive
                ? "bg-white dark:bg-zinc-800 text-zinc-900 dark:text-zinc-100 shadow-sm"
                : "text-zinc-500 dark:text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-300"
            )}
            aria-selected={isActive}
          >
            <span className={cn(
              "flex items-center justify-center w-5 h-5 rounded-full text-xs font-semibold transition-colors",
              isActive
                ? "bg-zinc-200 dark:bg-zinc-700 text-zinc-600 dark:text-zinc-300"
                : "bg-zinc-200 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500"
            )}>
              {index + 1}
            </span>
            <span>{step.label.replace(/^\d+\.\s*/, '')}</span>
          </button>
        );
      })}
    </div>
  );
}



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
    <section className="rounded-xl border p-2">
      <ul className="flex flex-wrap gap-2">
        {steps.map((step) => {
          const isActive = activeId === step.id;
          return (
            <li key={step.id}>
              <button
                onClick={() => onChange(step.id)}
                className={cn(
                  "inline-flex items-center gap-2 rounded-md border px-3 py-2 text-sm",
                  isActive ? "ring-2 ring-offset-1 ring-ring" : "",
                  step.completed
                    ? "border-green-700 bg-green-600 text-white"
                    : "hover:bg-black/5 dark:hover:bg-white/10"
                )}
                aria-selected={isActive}
              >
                {step.completed ? (
                  <svg viewBox="0 0 20 20" fill="none" className="h-4 w-4">
                    <path
                      d="M16.667 5.833 8.75 13.75 5 10"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                ) : (
                  <span className="inline-block h-2 w-2 rounded-full bg-zinc-400" />
                )}
                <span>{step.label}</span>
              </button>
            </li>
          );
        })}
      </ul>
    </section>
  );
}



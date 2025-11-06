'use client';
import { useEffect, useState } from "react";

type Theme = "light" | "dark";

function applyTheme(next: Theme) {
  const root = document.documentElement;
  if (next === "dark") {
    root.classList.add("dark");
  } else {
    root.classList.remove("dark");
  }
  try {
    localStorage.setItem("theme", next);
  } catch {
    // ignore
  }
}

export default function SettingsPage() {
  const [theme, setTheme] = useState<Theme>("dark");

  useEffect(() => {
    try {
      const stored = localStorage.getItem("theme");
      if (stored === "light" || stored === "dark") {
        setTheme(stored);
        applyTheme(stored);
        return;
      }
    } catch {
      // ignore
    }
    // default
    setTheme("dark");
    applyTheme("dark");
  }, []);

  function onChange(next: Theme) {
    setTheme(next);
    applyTheme(next);
  }

  return (
    <div className="flex min-h-screen w-full items-start justify-center bg-zinc-50 py-16 dark:bg-black">
      <main className="w-full max-w-xl rounded-xl border bg-white p-6 dark:bg-black">
        <h1 className="mb-6 text-xl font-semibold text-black dark:text-zinc-50">settings</h1>

        <section className="space-y-4">
          <div>
            <h2 className="mb-2 text-sm font-medium text-zinc-700 dark:text-zinc-300">
              theme
            </h2>
            <div className="flex gap-3">
              <label className="inline-flex items-center gap-2 rounded-md border px-3 py-2 text-sm">
                <input
                  type="radio"
                  name="theme"
                  value="dark"
                  checked={theme === "dark"}
                  onChange={() => onChange("dark")}
                />
                <span>dark</span>
              </label>
              <label className="inline-flex items-center gap-2 rounded-md border px-3 py-2 text-sm">
                <input
                  type="radio"
                  name="theme"
                  value="light"
                  checked={theme === "light"}
                  onChange={() => onChange("light")}
                />
                <span>light</span>
              </label>
            </div>
            <p className="mt-2 text-xs text-zinc-500">
              preference is saved in your browser and applied across pages.
            </p>
          </div>
        </section>
      </main>
    </div>
  );
}



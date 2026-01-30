'use client';
import { useTheme } from 'next-themes';
import { useEffect, useState } from 'react';

export default function SettingsPage() {
  const { theme, setTheme } = useTheme();
  const [mounted, setMounted] = useState(false);

  useEffect(() => setMounted(true), []);

  return (
    <div className="flex min-h-screen w-full items-start justify-center bg-zinc-50 py-16 dark:bg-black">
      <main className="w-full max-w-xl rounded-xl border bg-white p-6 dark:bg-black">
        <h1 className="mb-6 text-xl font-semibold text-black dark:text-zinc-50">settings</h1>

        <section className="space-y-4">
          <div>
            <h2 className="mb-2 text-sm font-medium text-zinc-700 dark:text-zinc-300">theme</h2>
            <div className="flex gap-3">
              <label className="inline-flex items-center gap-2 rounded-md border px-3 py-2 text-sm">
                <input
                  type="radio"
                  name="theme"
                  value="dark"
                  checked={mounted && theme === 'dark'}
                  onChange={() => setTheme('dark')}
                />
                <span>dark</span>
              </label>
              <label className="inline-flex items-center gap-2 rounded-md border px-3 py-2 text-sm">
                <input
                  type="radio"
                  name="theme"
                  value="light"
                  checked={mounted && theme === 'light'}
                  onChange={() => setTheme('light')}
                />
                <span>light</span>
              </label>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}

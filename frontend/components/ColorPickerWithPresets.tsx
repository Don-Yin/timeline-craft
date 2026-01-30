'use client';

import { useState, useRef, useEffect } from 'react';

const COLOR_PRESETS = [
  { name: 'Slate', hex: '#64748b' },
  { name: 'Gray', hex: '#6b7280' },
  { name: 'Zinc', hex: '#71717a' },
  { name: 'Stone', hex: '#78716c' },
  { name: 'Red', hex: '#ef4444' },
  { name: 'Orange', hex: '#f97316' },
  { name: 'Amber', hex: '#f59e0b' },
  { name: 'Yellow', hex: '#eab308' },
  { name: 'Lime', hex: '#84cc16' },
  { name: 'Green', hex: '#22c55e' },
  { name: 'Emerald', hex: '#10b981' },
  { name: 'Teal', hex: '#14b8a6' },
  { name: 'Cyan', hex: '#06b6d4' },
  { name: 'Sky', hex: '#0ea5e9' },
  { name: 'Blue', hex: '#3b82f6' },
  { name: 'Indigo', hex: '#6366f1' },
  { name: 'Violet', hex: '#8b5cf6' },
  { name: 'Purple', hex: '#a855f7' },
  { name: 'Fuchsia', hex: '#d946ef' },
  { name: 'Pink', hex: '#ec4899' },
  { name: 'Rose', hex: '#f43f5e' },
  { name: 'Black', hex: '#111111' },
  { name: 'White', hex: '#ffffff' },
  { name: 'Navy', hex: '#1e3a5f' },
];

interface ColorPickerWithPresetsProps {
  value: string;
  onChange: (color: string) => void;
  label: string;
}

export function ColorPickerWithPresets({ value, onChange, label }: ColorPickerWithPresetsProps) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className="flex items-center justify-between gap-3" ref={containerRef}>
      <span className="text-zinc-700 dark:text-zinc-300">{label}</span>
      <div className="relative">
        <button
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          className="flex items-center gap-2 rounded-lg border border-zinc-200 dark:border-zinc-700 px-2 py-1.5 transition-all hover:border-zinc-300 dark:hover:border-zinc-600 bg-white dark:bg-zinc-900"
        >
          <div
            className="w-6 h-6 rounded border border-zinc-300 dark:border-zinc-600 shadow-inner"
            style={{ backgroundColor: value }}
          />
          <span className="font-mono text-xs text-zinc-500 w-16">{value}</span>
          <svg
            className={`w-3 h-3 text-zinc-400 transition-transform ${isOpen ? 'rotate-180' : ''}`}
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
          >
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
          </svg>
        </button>

        {isOpen && (
          <div className="absolute right-0 top-full mt-2 z-50 w-[280px] rounded-xl border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 shadow-xl p-3 animate-in fade-in slide-in-from-top-2 duration-200">
            {/* Custom color input */}
            <div className="flex items-center gap-2 mb-3 pb-3 border-b border-zinc-100 dark:border-zinc-800">
              <input
                type="color"
                value={value}
                onChange={(e) => onChange(e.target.value)}
                className="w-10 h-10 cursor-pointer rounded-lg border-0 p-0"
              />
              <div className="flex-1">
                <label className="text-[10px] text-zinc-400 uppercase tracking-wider">Custom</label>
                <input
                  type="text"
                  value={value}
                  onChange={(e) => {
                    const val = e.target.value;
                    if (/^#[0-9A-Fa-f]{0,6}$/.test(val)) {
                      onChange(val);
                    }
                  }}
                  className="w-full bg-transparent font-mono text-sm border-none p-0 focus:outline-none focus:ring-0 text-zinc-700 dark:text-zinc-200"
                  placeholder="#000000"
                />
              </div>
            </div>

            {/* Preset grid */}
            <div className="grid grid-cols-8 gap-1.5">
              {COLOR_PRESETS.map((preset) => (
                <button
                  key={preset.hex}
                  type="button"
                  onClick={() => {
                    onChange(preset.hex);
                    setIsOpen(false);
                  }}
                  className={`group relative w-7 h-7 rounded-md border-2 transition-all hover:scale-110 hover:z-10 ${
                    value.toLowerCase() === preset.hex.toLowerCase()
                      ? 'border-emerald-500 ring-2 ring-emerald-500/30 scale-110 z-10'
                      : 'border-transparent hover:border-zinc-300 dark:hover:border-zinc-600'
                  }`}
                  style={{ backgroundColor: preset.hex }}
                  title={preset.name}
                >
                  {value.toLowerCase() === preset.hex.toLowerCase() && (
                    <span className="absolute inset-0 flex items-center justify-center">
                      <svg className={`w-3 h-3 ${preset.hex === '#ffffff' || preset.hex === '#f59e0b' || preset.hex === '#eab308' ? 'text-zinc-700' : 'text-white'}`} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={3}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                      </svg>
                    </span>
                  )}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}


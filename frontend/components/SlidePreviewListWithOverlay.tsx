'use client';

import { type IndexItem } from "@/lib/indexes";
import { type Slide } from "@/lib/slides";
import { useRef, useMemo } from "react";

/**
 * Renders slides with a client-side CSS overlay for the sidebar
 * This provides instant visual feedback without waiting for server-side PDF generation
 */
export function SlidePreviewListWithOverlay({
  slides,
  tags,
  slideTagMap,
  sidebarWidth = 12,
  itemHeight = 10,
  scrollRef,
}: {
  slides: Slide[];
  tags: IndexItem[];
  slideTagMap: Record<number, string | null>;
  sidebarWidth?: number;  // percentage 0-50
  itemHeight?: number;    // percentage 0-30
  scrollRef?: React.Ref<HTMLDivElement>;
}) {
  const innerRef = useRef<HTMLDivElement | null>(null);
  const setRef = (node: HTMLDivElement | null) => {
    innerRef.current = node;
    if (typeof scrollRef === "function") scrollRef(node);
    else if (scrollRef && "current" in (scrollRef as object)) {
      (scrollRef as { current: HTMLDivElement | null }).current = node;
    }
  };

  // Group consecutive slides by tag to build sidebar sections
  const tagGroups = useMemo(() => {
    const groups: { tagId: string; label: string; startIdx: number; count: number }[] = [];
    let lastTagId = '';
    const slideIds = Object.keys(slideTagMap).map(Number).sort((a, b) => a - b);
    
    for (const slideId of slideIds) {
      const tagId = slideTagMap[slideId] ?? '';
      if (tagId !== lastTagId || groups.length === 0) {
        const tag = tags.find(t => t.id === tagId);
        groups.push({ tagId, label: tag?.label ?? '', startIdx: slideId - 1, count: 1 });
        lastTagId = tagId;
      } else {
        groups[groups.length - 1].count++;
      }
    }
    return groups;
  }, [slideTagMap, tags]);

  // Find which group a slide belongs to
  const getGroupIndex = (slideIdx: number): number => {
    for (let i = 0; i < tagGroups.length; i++) {
      const group = tagGroups[i];
      if (slideIdx >= group.startIdx && slideIdx < group.startIdx + group.count) {
        return i;
      }
    }
    return 0;
  };

  return (
    <div className="rounded-lg border p-3">
      <div className="mb-2 flex items-center gap-2 text-xs text-zinc-500">
        <span className="inline-block h-2 w-2 rounded-full bg-emerald-500"></span>
        <span>live preview (instant client-side rendering)</span>
      </div>
      <div ref={setRef} className="max-h-[70vh] overflow-y-auto">
        <div className="flex flex-col gap-4">
          {slides.map((s) => {
            const slideIdx = s.id - 1;
            const currentGroupIdx = getGroupIndex(slideIdx);
            
            return (
              <div key={s.id} className="flex items-start gap-3">
                <span className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground">
                  {s.id}
                </span>
                <div className="relative w-full overflow-hidden rounded-md border bg-zinc-900">
                  {/* Sidebar overlay */}
                  <div 
                    className="absolute inset-y-0 left-0 z-10 flex flex-col overflow-hidden"
                    style={{ width: `${sidebarWidth}%` }}
                  >
                    {tagGroups.map((group, idx) => {
                      const isCurrent = idx === currentGroupIdx;
                      const groupHeight = group.count * itemHeight;
                      return (
                        <div
                          key={`${group.tagId}-${idx}`}
                          className={`flex items-center justify-center border-b border-zinc-700/30 text-center transition-colors ${
                            isCurrent 
                              ? 'bg-emerald-600 text-white font-medium' 
                              : 'bg-zinc-800/95 text-zinc-400'
                          }`}
                          style={{ 
                            height: `${groupHeight}%`,
                            minHeight: '18px',
                            fontSize: 'clamp(7px, 1.2vw, 10px)',
                          }}
                        >
                          <span className="truncate px-0.5 leading-tight">
                            {group.label || '—'}
                          </span>
                        </div>
                      );
                    })}
                    {/* Fill remaining space */}
                    <div className="flex-1 bg-zinc-900" />
                  </div>
                  
                  {/* Slide content - shifted right */}
                  <div 
                    style={{ 
                      marginLeft: `${sidebarWidth}%`,
                      width: `${100 - sidebarWidth}%` 
                    }}
                  >
                    <img
                      src={s.src}
                      alt={`Slide ${s.id}`}
                      className="h-auto w-full"
                    />
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}


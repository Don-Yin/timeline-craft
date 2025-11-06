'use client';
import { type IndexItem } from "@/lib/indexes";
import { useRef } from "react";

export function SlidePreviewList({
  slides,
  mode = "view",
  tags,
  slideTagMap,
  onSlideTagChange,
  scrollRef,
  showTagBadge = true,
}: {
  slides: number[];
  mode?: "view" | "associate";
  tags?: IndexItem[];
  slideTagMap?: Record<number, string | null>;
  onSlideTagChange?: (slideNum: number, tagId: string) => void;
  scrollRef?: React.Ref<HTMLDivElement>;
  showTagBadge?: boolean;
}) {
  const innerRef = useRef<HTMLDivElement | null>(null);
  const setRef = (node: HTMLDivElement | null) => {
    innerRef.current = node;
    if (typeof scrollRef === "function") scrollRef(node);
    else if (scrollRef && "current" in (scrollRef as any)) {
      (scrollRef as any).current = node;
    }
  };

  return (
    <div className="rounded-lg border p-3">
      <div ref={setRef} className="max-h-[70vh] overflow-y-auto">
        <div className="flex flex-col gap-4">
          {slides.map((n) => (
            <div key={n} className="flex items-start gap-3">
              <span className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground">
                {n}
              </span>
              <div className="relative w-full">
                <img
                  src={`https://picsum.photos/seed/slide-${n}/640/360`}
                  alt={`Slide ${n}`}
                  className="h-auto w-full rounded-md border"
                />
                {showTagBadge && tags && slideTagMap && (
                  <span className="absolute right-2 top-2 rounded bg-black/70 px-2 py-1 text-xs text-white dark:bg-white/20 dark:text-white">
                    {tags.find((t) => t.id === (slideTagMap[n] ?? ""))?.label ?? "no tag"}
                  </span>
                )}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}



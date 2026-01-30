'use client';
import { useState, useEffect, useCallback, useRef } from 'react';
import { type IndexItem } from '@/lib/indexes';
import { type Slide } from '@/lib/slides';
import { getFirstSlidePreview, type PreviewProgressEvent, type PreviewParams } from '@/lib/upload-client';
import { cn } from '@/lib/utils';

type Props = {
  fileId: string;
  slides: Slide[];
  tags: IndexItem[];
  slideTagMap: Record<number, string | null>;
  sidebarWidth: number; // percentage 0-100
  itemHeight: number;   // percentage 0-100
  scrollRef?: React.Ref<HTMLDivElement>;
  showTagBadge?: boolean;
};

export function PreviewWithProgress({
  fileId, slides, tags, slideTagMap, sidebarWidth, itemHeight, scrollRef, showTagBadge = true,
}: Props) {
  const [previewImages, setPreviewImages] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const abortControllerRef = useRef<AbortController | null>(null);

  const innerRef = useRef<HTMLDivElement | null>(null);
  const setRef = (node: HTMLDivElement | null) => {
    innerRef.current = node;
    if (typeof scrollRef === 'function') scrollRef(node);
    else if (scrollRef && 'current' in (scrollRef as object)) {
      (scrollRef as { current: HTMLDivElement | null }).current = node;
    }
  };

  const buildTagsArray = useCallback((): string[] => {
    return Array.from({ length: slides.length }, (_, idx) => {
      const slideNum = slides[idx].id;
      const tagId = slideTagMap[slideNum];
      const tag = tags.find((t) => t.id === tagId);
      return tag?.label ?? 'untitled';
    });
  }, [slides, slideTagMap, tags]);

  useEffect(() => {
    if (slides.length === 0 || !fileId) {
      setPreviewImages([]);
      return;
    }

    // Abort previous request
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }

    const controller = new AbortController();
    abortControllerRef.current = controller;

    async function fetchPreviews() {
      setIsLoading(true);
      setProgress(0);
      setProgressMessage('starting...');

      const params: PreviewParams = {
        tags: buildTagsArray(),
        sidebar_width: sidebarWidth / 100,
        sidebar_item_height: itemHeight / 100,
      };

      const result = await getFirstSlidePreview(
        fileId,
        params,
        (event: PreviewProgressEvent) => {
          setProgress(event.progress);
          setProgressMessage(event.message);
        },
        controller.signal
      );

      setPreviewImages(result.thumbnails.map((b64) => `data:image/${result.format};base64,${b64}`));
      setIsLoading(false);
    }

    fetchPreviews().catch((err) => {
      if (err.name !== 'AbortError') {
        console.error('Preview fetch error:', err);
        setIsLoading(false);
      }
    });

    return () => controller.abort();
  }, [fileId, slides.length, sidebarWidth, itemHeight, buildTagsArray]);

  return (
    <div className="rounded-lg border p-3">
      {/* Progress bar */}
      {isLoading && (
        <div className="mb-4">
          <div className="flex items-center justify-between text-sm text-zinc-500 mb-2">
            <span className="flex items-center gap-2">
              <div className="h-4 w-4 animate-spin rounded-full border-2 border-emerald-500 border-t-transparent" />
              {progressMessage}
            </span>
            <span>{progress}%</span>
          </div>
          <div className="h-2 w-full bg-zinc-200 rounded-full overflow-hidden dark:bg-zinc-700">
            <div
              className="h-full bg-emerald-500 transition-all duration-300 ease-out"
              style={{ width: `${progress}%` }}
            />
          </div>
        </div>
      )}

      {/* First slide preview only */}
      <div ref={setRef} className={cn('', isLoading && 'opacity-50')}>
        {slides.length > 0 && (() => {
          const firstSlide = slides[0];
          const currentTagId = slideTagMap[firstSlide.id];
          const currentTag = tags.find((t) => t.id === currentTagId);
          const previewSrc = previewImages[0];

          return (
            <div className="flex flex-col gap-2">
              <div className="flex items-center gap-2 text-sm text-zinc-500">
                <span className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground">
                  1
                </span>
                <span>preview of first slide</span>
              </div>
              <div className="relative w-full aspect-video overflow-hidden rounded-md border">
                {previewSrc ? (
                  <img src={previewSrc} alt="Slide 1 Preview" className="h-full w-full object-cover" />
                ) : (
                  <div className="flex h-full w-full items-center justify-center bg-zinc-100 text-zinc-400 text-xs dark:bg-zinc-800">
                    {isLoading ? 'rendering preview...' : 'no preview'}
                  </div>
                )}
                {showTagBadge && currentTag && (
                  <span className="absolute right-2 top-2 rounded bg-black/70 px-2 py-1 text-xs text-white">
                    {currentTag.label}
                  </span>
                )}
              </div>
              <p className="text-xs text-zinc-400 text-center">
                showing first slide only for faster preview
              </p>
            </div>
          );
        })()}
      </div>
    </div>
  );
}


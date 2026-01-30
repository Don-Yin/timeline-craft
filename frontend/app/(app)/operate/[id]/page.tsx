'use client';
import { useRef, useState, useEffect, useMemo, useCallback, use } from "react";
import DraggableIndexList from "@/components/DraggableIndexList";
import { createIndexItem, type IndexItem } from "@/lib/indexes";
import { generateDummySlides } from "@/lib/slides";
import { StepTabs, type Step } from "@/components/StepTabs";
import { ParametersForm } from "@/components/ParametersForm";
import { PreviewWithProgress } from "@/components/PreviewWithProgress";
import ResizableColumns from "@/components/ResizableColumns";
import { TagSlideManager } from "@/components/TagSlideManager";
import { getAllThumbnails, processWithProgress, type ProgressEvent } from "@/lib/api";

export default function Operate({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);

  const [slideImages, setSlideImages] = useState<string[]>([]);
  const [loadingThumbnails, setLoadingThumbnails] = useState(false);
  const [slideCount, setSlideCount] = useState<number>(0);
  const [sidebarWidth, setSidebarWidth] = useState<number>(12);
  const [itemHeight, setItemHeight] = useState<number>(10);
  const [duration, setDuration] = useState<number>(0.3);
  const [applyMorph, setApplyMorph] = useState<boolean>(true);
  const [sidebarColorHex, setSidebarColorHex] = useState<string>("#5A5A5A");
  const [indicatorColorHex, setIndicatorColorHex] = useState<string>("#111111");
  const [sidebarFontColorHex, setSidebarFontColorHex] = useState<string>("#FFFFFF");
  const [indexes, setIndexes] = useState<IndexItem[]>(
    ["intro", "methods", "results", "discussion", "conclusion"].map((l) => createIndexItem(l))
  );
  const [slideTagMap, setSlideTagMap] = useState<Record<number, string | null>>({});
  const [activeTab, setActiveTab] = useState<"tags" | "params">("tags");
  const [isDownloading, setIsDownloading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [progressMessage, setProgressMessage] = useState("");

  const previewScrollRef = useRef<HTMLDivElement | null>(null);

  const slides = useMemo(() => (
    generateDummySlides(slideCount).map((s, i) => ({ ...s, src: (slideImages[i] || s.src) as string }))
  ), [slideCount, slideImages]);

  const buildTagsArray = useCallback((): string[] => {
    return Array.from({ length: slideCount }, (_, idx) => {
      const slideNum = idx + 1;
      const tagId = slideTagMap[slideNum];
      const tag = indexes.find((t) => t.id === tagId);
      return tag?.label ?? "untitled";
    });
  }, [slideCount, slideTagMap, indexes]);

  const slideCounts = useMemo((): Record<string, number> => {
    const counts: Record<string, number> = {};
    Object.values(slideTagMap).forEach((tagId) => {
      if (tagId) counts[tagId] = (counts[tagId] || 0) + 1;
    });
    return counts;
  }, [slideTagMap]);

  const distributeSlides = useCallback((tags: IndexItem[]) => {
    if (slideCount === 0) {
      setSlideTagMap({});
      return;
    }
    if (tags.length === 0) {
      setSlideTagMap(Object.fromEntries(Array.from({ length: slideCount }, (_, idx) => [idx + 1, null])));
      return;
    }
    const map: Record<number, string | null> = {};
    const chunkSize = Math.max(1, Math.ceil(slideCount / tags.length));
    Array.from({ length: slideCount }, (_, idx) => idx + 1).forEach((slideId, index) => {
      const tagIndex = Math.floor(index / chunkSize);
      map[slideId] = tags[Math.min(tagIndex, tags.length - 1)]?.id ?? null;
    });
    setSlideTagMap(map);
  }, [slideCount]);

  useEffect(() => {
    async function init() {
      setLoadingThumbnails(true);
      const thumbnails = await getAllThumbnails(id);
      setSlideCount(thumbnails.length || 36);
      setSlideImages(thumbnails.map(b64 => b64 ? `data:image/png;base64,${b64}` : ''));
      setLoadingThumbnails(false);
    }
    if (id) init();
  }, [id]);

  useEffect(() => {
    distributeSlides(indexes);
  }, [indexes.length, distributeSlides]);

  function handleTagsChange(newIndexes: IndexItem[]) {
    setIndexes(newIndexes);
    if (newIndexes.length === indexes.length) {
      const orderChanged = newIndexes.some((item, i) => item.id !== indexes[i].id);
      if (orderChanged) distributeSlides(newIndexes);
    }
  }

  async function handleDownload() {
    setIsDownloading(true);
    setProgress(0);
    setProgressMessage("starting...");

    await processWithProgress(
      id,
      {
        tags: buildTagsArray(),
        sidebar_width: sidebarWidth / 100,
        sidebar_item_height: itemHeight / 100,
        transition_duration: duration,
        apply_morph_transition: applyMorph,
        sidebar_color_hex: sidebarColorHex,
        indicator_color_hex: indicatorColorHex,
        sidebar_item_font_color_hex: sidebarFontColorHex,
      },
      (event: ProgressEvent) => {
        setProgress(event.progress);
        setProgressMessage(event.message);
      }
    );

    setIsDownloading(false);
    setProgress(0);
    setProgressMessage("");
  }

  const tagsComplete = indexes.length > 0;
  const steps: Step[] = [
    { id: "tags", label: "1. set tags", completed: tagsComplete },
    { id: "params", label: "2. choose params", completed: true },
  ];

  return (
    <div className="flex min-h-screen w-full bg-zinc-50 font-sans dark:bg-black">
      <main className="flex min-h-screen w-full flex-col gap-8 bg-white py-8 px-6 dark:bg-black sm:px-12">
        <StepTabs steps={steps} activeId={activeTab} onChange={(id) => setActiveTab(id as typeof activeTab)} />
        <section className="rounded-xl border p-5">
          <ResizableColumns
            defaultLeftFraction={0.32}
            storageKey="operate.split"
            className="items-start"
            left={
              <div>
                {activeTab === "tags" && (
                  <DraggableIndexList title="arrange tags" items={indexes} onChange={handleTagsChange} slideCounts={slideCounts} />
                )}
                {activeTab === "params" && (
                  <ParametersForm
                    sidebarWidth={sidebarWidth}
                    onSidebarWidthChange={setSidebarWidth}
                    itemHeight={itemHeight}
                    onItemHeightChange={setItemHeight}
                    duration={duration}
                    onDurationChange={setDuration}
                    applyMorph={applyMorph}
                    onApplyMorphChange={setApplyMorph}
                    sidebarColorHex={sidebarColorHex}
                    onSidebarColorHexChange={setSidebarColorHex}
                    indicatorColorHex={indicatorColorHex}
                    onIndicatorColorHexChange={setIndicatorColorHex}
                    sidebarFontColorHex={sidebarFontColorHex}
                    onSidebarFontColorHexChange={setSidebarFontColorHex}
                    onDownload={handleDownload}
                    isDownloading={isDownloading}
                    canDownload={tagsComplete && slideCount > 0}
                    progress={progress}
                    progressMessage={progressMessage}
                  />
                )}
              </div>
            }
            right={
              activeTab === "tags" ? (
                loadingThumbnails ? (
                  <div className="flex h-[400px] items-center justify-center">
                    <div className="flex items-center gap-2 text-sm text-zinc-500">
                      <div className="h-4 w-4 animate-spin rounded-full border-2 border-emerald-500 border-t-transparent" />
                      loading slides...
                    </div>
                  </div>
                ) : (
                  <TagSlideManager slides={slides} tags={indexes} slideTagMap={slideTagMap} onSlideMove={(n, t) => setSlideTagMap((p) => ({ ...p, [n]: t }))} />
                )
              ) : (
                <PreviewWithProgress
                  fileId={id}
                  slides={slides}
                  tags={indexes}
                  slideTagMap={slideTagMap}
                  sidebarWidth={sidebarWidth}
                  itemHeight={itemHeight}
                  sidebarColorHex={sidebarColorHex}
                  indicatorColorHex={indicatorColorHex}
                  sidebarFontColorHex={sidebarFontColorHex}
                  scrollRef={previewScrollRef}
                />
              )
            }
          />
        </section>
        <p className="text-xs text-zinc-500">
          showing {slideCount || 36} {slideCount === 1 ? "slide" : "slides"} previews
        </p>
      </main>
    </div>
  );
}

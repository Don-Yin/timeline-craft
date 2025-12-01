'use client';
import { useRef, useState, useEffect, useMemo, useCallback, use } from "react";
import DraggableIndexList from "@/components/DraggableIndexList";
import { createIndexItem, type IndexItem } from "@/lib/indexes";
import { generateDummySlides } from "@/lib/slides";
import { StepTabs, type Step } from "@/components/StepTabs";
import { ParametersForm } from "@/components/ParametersForm";
import { SlidePreviewList } from "@/components/SlidePreviewList";
import ResizableColumns from "@/components/ResizableColumns";
import { TagSlideManager } from "@/components/TagSlideManager";
import { getThumbnail, processFile, getSlideCount } from "@/lib/upload-client";

export default function Operate({ params }: { params: Promise<{ id: string }> }) {
  // Unwrap params using React.use()
  const { id } = use(params);

  const [file, setFile] = useState<File | null>(null);
  // ... rest of state ...

  const [slideImages, setSlideImages] = useState<string[]>([]);
  const [loadingThumbnails, setLoadingThumbnails] = useState(false);
  const [slideCount, setSlideCount] = useState<number>(0);

  useEffect(() => {
    async function init() {
      setLoadingThumbnails(true);
      // 1. Get slide count
      const count = await getSlideCount(id);
      const actualCount = count > 0 ? count : 36; // fallback to 36 if 0 (e.g. fail or empty)
      setSlideCount(actualCount);

      // 2. Load thumbnails based on count
      const loadedImages: string[] = [];
      for (let i = 0; i < actualCount; i++) {
        try {
          const b64 = await getThumbnail(id, i);
          if (b64) {
            loadedImages.push(`data:image/png;base64,${b64}`);
          } else {
            loadedImages.push('');
          }
        } catch (e) {
          loadedImages.push('');
        }
      }
      setSlideImages(loadedImages);
      setLoadingThumbnails(false);
    }

    if (id) {
      init();
    }
  }, [id]);

  // ... rest of existing code ...

  const [sidebarWidth, setSidebarWidth] = useState<number>(12);
  const [itemHeight, setItemHeight] = useState<number>(10);
  const [duration, setDuration] = useState<number>(0.3);
  const [applyMorph, setApplyMorph] = useState<boolean>(true);
  const [indexes, setIndexes] = useState<IndexItem[]>(
    ["intro", "methods", "results", "discussion", "conclusion"].map((l) =>
      createIndexItem(l)
    )
  );

  // Merge loaded images with dummy slides structure
  // In reality, we should generate slides based on actual count.
  // For now, we override the 'src' of dummySlides if we have a real image.
  const slides = useMemo(() => (
    generateDummySlides(slideCount).map((s, i) => ({
      ...s,
      src: (slideImages[i] || s.src) as string // use loaded thumbnail or fallback to dummy
    }))
  ), [slideCount, slideImages]);

  const [slideTagMap, setSlideTagMap] = useState<Record<number, string | null>>(
    {}
  );
  const [activeTab, setActiveTab] = useState<"tags" | "params" | "preview">("tags");
  const previewScrollRef = useRef<HTMLDivElement | null>(null);

  // Helper to evenly distribute slides among tags
  const distributeSlides = useCallback((tags: IndexItem[]) => {
    if (slideCount === 0) {
      setSlideTagMap({});
      return;
    }

    if (tags.length === 0) {
      setSlideTagMap(
        Object.fromEntries(
          Array.from({ length: slideCount }, (_, idx) => idx + 1).map((id) => [id, null])
        )
      );
      return;
    }

    const map: Record<number, string | null> = {};
    const chunkSize = Math.max(1, Math.ceil(slideCount / tags.length));
    Array.from({ length: slideCount }, (_, idx) => idx + 1).forEach((slideId, index) => {
      const tagIndex = Math.floor(index / chunkSize);
      const tag = tags[Math.min(tagIndex, tags.length - 1)];
      map[slideId] = tag?.id ?? null;
    });
    setSlideTagMap(map);
  }, [slideCount]);

  // Initial distribution and whenever tags change (number or ids)
  // Note: We might want to preserve manual overrides, but requirement says "by default evenly assigned"
  // when tags are created. For simplicity, we re-distribute when the *count* of tags changes or if we have no map yet.
  useEffect(() => {
    // Re-distribute whenever tag count changes (e.g., tag added/removed) or slide count updates.
    distributeSlides(indexes);
  }, [indexes.length, distributeSlides]);
  // ^ Only re-distribute if number of tags changes. Renaming shouldn't trigger it.

  function handleTagsChange(newIndexes: IndexItem[]) {
    setIndexes(newIndexes);
    // If length is different, useEffect will handle it.
    // If just reordered, we might want to re-distribute to match new order?
    // "Default evenly assigned" implies the first tag gets the first chunk.
    // If I move "Intro" to the end, should "Intro" now get the last chunk?
    // Yes, probably.
    if (newIndexes.length === indexes.length) {
      // Check if order changed
      const orderChanged = newIndexes.some((item, i) => item.id !== indexes[i].id);
      if (orderChanged) {
        distributeSlides(newIndexes);
      }
    }
  }

  function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const selected = event.target.files?.[0] ?? null;
    setFile(selected);
  }

  function onSlideTagChange(slideNum: number, tagId: string | null) {
    setSlideTagMap((prev) => ({ ...prev, [slideNum]: tagId }));
  }

  const tagsComplete = indexes.length > 0;
  const paramsComplete =
    sidebarWidth !== 12 ||
    itemHeight !== 10 ||
    duration !== 0.3 ||
    applyMorph !== true;
  const previewComplete = tagsComplete && paramsComplete;

  const steps: Step[] = [
    { id: "tags", label: "1. set tags", completed: tagsComplete },
    { id: "params", label: "2. choose params", completed: paramsComplete },
    { id: "preview", label: "3. preview", completed: previewComplete },
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
                  <DraggableIndexList title="arrange tags" items={indexes} onChange={handleTagsChange} />
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
                  />
                )}
                {activeTab === "preview" && (
                  <div className="rounded-lg border p-4">
                    <h3 className="mb-3 text-sm font-medium">export</h3>
                    <div className="flex flex-wrap gap-3">
                      <button
                        className="rounded-md bg-foreground px-4 py-2 text-sm text-background transition-colors hover:opacity-95 disabled:opacity-60"
                        disabled
                        title="export not wired yet"
                      >
                        download powerpoint (.pptx)
                      </button>
                      <button
                        className="rounded-md border px-4 py-2 text-sm hover:bg-black/5 dark:hover:bg-white/10 disabled:opacity-60"
                        disabled
                        title="export not wired yet"
                      >
                        download pdf
                      </button>
                    </div>
                    <p className="mt-2 text-xs text-zinc-500">
                      exports are not available in this environment yet.
                    </p>
                  </div>
                )}
              </div>
            }
            right={
              activeTab === "tags" ? (
                <TagSlideManager
                  slides={slides}
                  tags={indexes}
                  slideTagMap={slideTagMap}
                  onSlideMove={onSlideTagChange}
                />
              ) : (
                <SlidePreviewList
                  slides={slides}
                  mode="view"
                  tags={indexes}
                  slideTagMap={slideTagMap}
                  // onSlideTagChange not needed in view mode usually
                  scrollRef={previewScrollRef}
                  showTagBadge={true}
                />
              )
            }
          />
        </section>
        <p className="text-xs text-zinc-500">
          showing {slideCount || 36} {slideCount === 1 ? "slide" : "slides"} as placeholders. actual previews will appear once the rendering pipeline is wired end-to-end.
        </p>
      </main>
    </div>
  );
}

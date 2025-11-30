'use client';
import { useRef, useState, useEffect } from "react";
import DraggableIndexList from "@/components/DraggableIndexList";
import { createIndexItem, type IndexItem } from "@/lib/indexes";
import { generateDummySlides } from "@/lib/slides";
import { StepTabs, type Step } from "@/components/StepTabs";
import { ParametersForm } from "@/components/ParametersForm";
import { SlidePreviewList } from "@/components/SlidePreviewList";
import ResizableColumns from "@/components/ResizableColumns";
import { TagSlideManager } from "@/components/TagSlideManager";

export default function Operate() {
  const [file, setFile] = useState<File | null>(null);
  const [sidebarWidth, setSidebarWidth] = useState<number>(12);
  const [itemHeight, setItemHeight] = useState<number>(10);
  const [duration, setDuration] = useState<number>(0.3);
  const [applyMorph, setApplyMorph] = useState<boolean>(true);
  const [indexes, setIndexes] = useState<IndexItem[]>(
    ["intro", "methods", "results", "discussion", "conclusion"].map((l) =>
      createIndexItem(l)
    )
  );
  const dummySlides = generateDummySlides(36);
  const [slideTagMap, setSlideTagMap] = useState<Record<number, string | null>>(
    {}
  );
  const [activeTab, setActiveTab] = useState<"tags" | "params" | "preview">("tags");
  const previewScrollRef = useRef<HTMLDivElement | null>(null);

  // Helper to evenly distribute slides among tags
  function distributeSlides(tags: IndexItem[]) {
    if (tags.length === 0) {
        setSlideTagMap(Object.fromEntries(dummySlides.map(s => [s, null])));
        return;
    }
    const map: Record<number, string | null> = {};
    const chunkSize = Math.ceil(dummySlides.length / tags.length);
    dummySlides.forEach((slide, index) => {
        const tagIndex = Math.floor(index / chunkSize);
        const tag = tags[Math.min(tagIndex, tags.length - 1)];
        map[slide] = tag.id;
    });
    setSlideTagMap(map);
  }

  // Initial distribution and whenever tags change (number or ids)
  // Note: We might want to preserve manual overrides, but requirement says "by default evenly assigned"
  // when tags are created. For simplicity, we re-distribute when the *count* of tags changes or if we have no map yet.
  useEffect(() => {
     const currentTaggedCount = Object.values(slideTagMap).filter(Boolean).length;
     // If we haven't initialized or tag count changed significantly (naive check)
     // A better check: if we added/removed a tag.
     // For now, let's re-distribute whenever the list of tags changes in length,
     // effectively resetting the distribution.
     // If the user just renames a tag, we shouldn't re-distribute.
     // If the user reorders tags, we probably shouldn't re-distribute unless we want to follow order?
     // Let's assume "evenly assigned" based on current order.
     distributeSlides(indexes);
  }, [indexes.length]); 
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
    file !== null ||
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
                    file={file}
                    onFileChange={setFile}
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
                        slides={dummySlides}
                        tags={indexes}
                        slideTagMap={slideTagMap}
                        onSlideMove={onSlideTagChange}
                    />
                ) : (
                  <SlidePreviewList
                    slides={dummySlides}
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
          showing 36 dummy images as placeholders. this will render generated previews once the pipeline is wired.
        </p>
      </main>
    </div>
  );
}

'use client';
import { useRef, useState } from "react";
import DraggableIndexList from "@/components/DraggableIndexList";
import { createIndexItem, type IndexItem } from "@/lib/indexes";
import { generateDummySlides } from "@/lib/slides";
import { StepTabs, type Step } from "@/components/StepTabs";
import { ParametersForm } from "@/components/ParametersForm";
import { SlidePreviewList } from "@/components/SlidePreviewList";
import ResizableColumns from "@/components/ResizableColumns";

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
    () =>
      Object.fromEntries(dummySlides.map((n) => [n, null])) as Record<
        number,
        string | null
      >
  );
  const [activeTab, setActiveTab] = useState<"tags" | "params" | "preview">("tags");
  const previewScrollRef = useRef<HTMLDivElement | null>(null);

  function onFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const selected = event.target.files?.[0] ?? null;
    setFile(selected);
  }

  function onSlideTagChange(slideNum: number, tagId: string) {
    setSlideTagMap((prev) => ({ ...prev, [slideNum]: tagId || null }));
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
                  <DraggableIndexList title="arrange tags" items={indexes} onChange={setIndexes} />
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
              <SlidePreviewList
                slides={dummySlides}
                mode={activeTab === "tags" ? "associate" : "view"}
                tags={indexes}
                slideTagMap={slideTagMap}
                onSlideTagChange={onSlideTagChange}
                scrollRef={previewScrollRef}
                showTagBadge={true}
              />
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



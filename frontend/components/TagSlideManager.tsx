'use client';
import React, { useState, useEffect } from "react";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragOverlay,
  DragEndEvent,
  DragOverEvent,
  defaultDropAnimationSideEffects,
  DropAnimation,
  useDraggable,
  useDroppable,
} from "@dnd-kit/core";
import { CSS } from "@dnd-kit/utilities";
import { type IndexItem } from "@/lib/indexes";
import { type Slide } from "@/lib/slides";

type TagSlideManagerProps = {
  slides: Slide[];
  tags: IndexItem[];
  slideTagMap: Record<number, string | null>;
  onSlideMove: (slideId: number, newTagId: string | null) => void;
};

function DraggableSlide({ id, slide }: { id: string; slide: Slide }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id,
    data: { slideNum: slide.id, src: slide.src },
  });

  return (
    <div
      ref={setNodeRef}
      style={{ transform: CSS.Translate.toString(transform), opacity: isDragging ? 0.4 : 1, zIndex: isDragging ? 50 : "auto" }}
      {...attributes}
      {...listeners}
      className="relative cursor-grab rounded-md border bg-zinc-100 dark:bg-zinc-800 transition-all hover:ring-2 hover:ring-primary/50"
    >
      <img src={slide.src} alt={`Slide ${slide.id}`} className="block w-full h-auto pointer-events-none rounded-md" />
      <span className="absolute left-1 top-1 flex h-5 w-5 items-center justify-center rounded bg-black/60 text-[10px] font-medium text-white">
        {slide.id}
      </span>
    </div>
  );
}

function DroppableTagContainer({ id, isEmpty, isOver, children }: { id: string; isEmpty: boolean; isOver: boolean; children: React.ReactNode }) {
  const { setNodeRef } = useDroppable({ id });

  return (
    <div
      ref={setNodeRef}
      className={`grid grid-cols-3 gap-2 items-start rounded-lg border-2 border-dashed p-2 min-h-[100px] transition-all duration-200 ${
        isOver ? "bg-emerald-50 border-emerald-500 dark:bg-emerald-950/30 dark:border-emerald-500" : "bg-zinc-50/50 border-zinc-200 dark:bg-zinc-900/50 dark:border-zinc-700"
      }`}
    >
      {children}
      {isEmpty && <div className="col-span-3 flex h-full items-center justify-center text-sm text-zinc-400 italic">drop slides here</div>}
    </div>
  );
}

function TagAccordion({ tag, slidesInTag, isOver, containerId, activeId, isExpanded, onToggle }: {
  tag: IndexItem;
  slidesInTag: Slide[];
  isOver: boolean;
  containerId: string;
  activeId: string | null;
  isExpanded: boolean;
  onToggle: () => void;
}) {
  return (
    <div className="rounded-lg border border-zinc-200 dark:border-zinc-700 overflow-hidden">
      <button
        onClick={onToggle}
        className={`w-full flex items-center justify-between px-4 py-3 transition-colors ${
          isExpanded ? "bg-zinc-100 dark:bg-zinc-800" : "bg-white dark:bg-zinc-900 hover:bg-zinc-50 dark:hover:bg-zinc-800/50"
        }`}
      >
        <div className="flex items-center gap-3">
          <span className={`transition-transform duration-200 ${isExpanded ? "rotate-90" : ""}`}>
            <svg className="w-4 h-4 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
            </svg>
          </span>
          <span className="font-medium text-zinc-800 dark:text-zinc-200">{tag.label}</span>
        </div>
        <span className={`rounded-full px-2.5 py-0.5 text-xs font-medium transition-colors ${
          slidesInTag.length > 0 ? "bg-emerald-100 text-emerald-700 dark:bg-emerald-900/50 dark:text-emerald-300" : "bg-zinc-200 text-zinc-500 dark:bg-zinc-700 dark:text-zinc-400"
        }`}>
          {slidesInTag.length}
        </span>
      </button>

      {isExpanded && (
        <div className="p-2">
          <DroppableTagContainer id={containerId} isEmpty={slidesInTag.length === 0 && !activeId} isOver={isOver}>
            {slidesInTag.map((s) => (
              <DraggableSlide key={`slide-${s.id}`} id={`slide-${s.id}`} slide={s} />
            ))}
          </DroppableTagContainer>
        </div>
      )}
    </div>
  );
}

export function TagSlideManager({ slides, tags, slideTagMap, onSlideMove }: TagSlideManagerProps) {
  const [activeId, setActiveId] = useState<string | null>(null);
  const [activeSrc, setActiveSrc] = useState<string | null>(null);
  const [overContainerId, setOverContainerId] = useState<string | null>(null);
  const [expandedTags, setExpandedTags] = useState<Set<string>>(new Set(tags.map(t => t.id)));

  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 5 } }), useSensor(KeyboardSensor));
  const activeSlideNum = activeId ? parseInt(activeId.replace("slide-", ""), 10) : null;

  useEffect(() => {
    setExpandedTags(new Set(tags.map(t => t.id)));
  }, [tags.length]);

  function getSlidesForTag(tagId: string) {
    return slides.filter(s => slideTagMap[s.id] === tagId);
  }

  function toggleTag(tagId: string) {
    setExpandedTags(prev => {
      const next = new Set(prev);
      if (next.has(tagId)) next.delete(tagId);
      else next.add(tagId);
      return next;
    });
  }

  function handleDragOver(event: DragOverEvent) {
    const { over } = event;
    if (over?.id) {
      const overId = over.id as string;
      if (overId.startsWith("container-")) {
        setOverContainerId(overId);
        const tagId = overId.replace("container-", "");
        if (!expandedTags.has(tagId)) {
          setExpandedTags(prev => new Set([...prev, tagId]));
        }
      }
    } else {
      setOverContainerId(null);
    }
  }

  function handleDragEnd(event: DragEndEvent) {
    const { active, over } = event;
    setActiveId(null);
    setActiveSrc(null);
    setOverContainerId(null);

    if (!over) return;

    const draggedSlideNum = active.data.current?.slideNum as number;
    if (!draggedSlideNum) return;

    const overId = over.id as string;
    let targetTagId: string | null = null;

    if (overId.startsWith("container-")) targetTagId = overId.replace("container-", "");
    else if (overId.startsWith("slide-")) {
      const targetSlideNum = parseInt(overId.replace("slide-", ""), 10);
      targetTagId = slideTagMap[targetSlideNum] ?? null;
    }

    if (!targetTagId) return;

    const sourceTagId = slideTagMap[draggedSlideNum];
    if (sourceTagId === targetTagId) return;

    const sourceTagIndex = tags.findIndex(t => t.id === sourceTagId);
    const targetTagIndex = tags.findIndex(t => t.id === targetTagId);

    if (sourceTagIndex === -1 || targetTagIndex === -1) return;

    const allSlidesSorted = [...slides].sort((a, b) => a.id - b.id);
    const draggedSlideIndex = allSlidesSorted.findIndex(s => s.id === draggedSlideNum);

    if (targetTagIndex > sourceTagIndex) {
      allSlidesSorted.slice(draggedSlideIndex).filter(s => {
        const sTagIndex = tags.findIndex(t => t.id === slideTagMap[s.id]);
        return sTagIndex >= sourceTagIndex && sTagIndex < targetTagIndex;
      }).forEach(s => onSlideMove(s.id, targetTagId));
    } else {
      allSlidesSorted.slice(0, draggedSlideIndex + 1).filter(s => {
        const sTagIndex = tags.findIndex(t => t.id === slideTagMap[s.id]);
        return sTagIndex > targetTagIndex && sTagIndex <= sourceTagIndex;
      }).forEach(s => onSlideMove(s.id, targetTagId));
    }
  }

  const dropAnimation: DropAnimation = { sideEffects: defaultDropAnimationSideEffects({ styles: { active: { opacity: "0.5" } } }) };

  return (
    <div className="flex flex-col h-full max-h-[calc(100vh-280px)] overflow-hidden">
      <div className="flex items-center justify-between mb-3 px-1 flex-shrink-0">
        <h3 className="text-sm font-medium text-zinc-600 dark:text-zinc-400">arrange slides by section</h3>
        <div className="flex gap-2">
          <button onClick={() => setExpandedTags(new Set(tags.map(t => t.id)))} className="text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 transition-colors">
            expand all
          </button>
          <span className="text-zinc-300 dark:text-zinc-600">|</span>
          <button onClick={() => setExpandedTags(new Set())} className="text-xs text-zinc-500 hover:text-zinc-700 dark:hover:text-zinc-300 transition-colors">
            collapse all
          </button>
        </div>
      </div>

      <DndContext
        sensors={sensors}
        collisionDetection={closestCenter}
        onDragStart={(e) => { setActiveId(e.active.id as string); setActiveSrc(e.active.data.current?.src); }}
        onDragOver={handleDragOver}
        onDragEnd={handleDragEnd}
      >
        <div className="flex-1 overflow-y-auto overflow-x-hidden space-y-2 pb-4">
          {tags.map((tag) => {
            const slidesInTag = getSlidesForTag(tag.id);
            const containerId = `container-${tag.id}`;

            return (
              <TagAccordion
                key={tag.id}
                tag={tag}
                slidesInTag={slidesInTag}
                isOver={overContainerId === containerId}
                containerId={containerId}
                activeId={activeId}
                isExpanded={expandedTags.has(tag.id)}
                onToggle={() => toggleTag(tag.id)}
              />
            );
          })}
        </div>

        <DragOverlay dropAnimation={dropAnimation}>
          {activeSlideNum && activeSrc ? (
            <div className="relative w-[180px] overflow-hidden rounded-lg shadow-2xl ring-2 ring-emerald-500 bg-white dark:bg-zinc-900">
              <img src={activeSrc} alt="" className="w-full h-auto" />
              <span className="absolute left-1 top-1 flex h-5 w-5 items-center justify-center rounded bg-emerald-600 text-[10px] font-bold text-white">
                {activeSlideNum}
              </span>
            </div>
          ) : null}
        </DragOverlay>
      </DndContext>
    </div>
  );
}

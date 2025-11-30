'use client';
import React, { useState } from "react";
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragOverlay,
  DragStartEvent,
  DragEndEvent,
  defaultDropAnimationSideEffects,
  DropAnimation,
  useDraggable,
  useDroppable,
} from "@dnd-kit/core";
import { CSS } from "@dnd-kit/utilities";
import { type IndexItem } from "@/lib/indexes";

type TagSlideManagerProps = {
  slides: number[];
  tags: IndexItem[];
  slideTagMap: Record<number, string | null>;
  onSlideMove: (slideId: number, newTagId: string | null) => void;
};

// Draggable slide item
function DraggableSlide({
  id,
  slideNum,
  tagId,
}: {
  id: string;
  slideNum: number;
  tagId: string | null;
}) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id,
    data: { slideNum, tagId, type: "slide" },
  });

  const style = {
    transform: CSS.Translate.toString(transform),
    opacity: isDragging ? 0.3 : 1,
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      {...attributes}
      {...listeners}
      className="relative aspect-video cursor-grab overflow-hidden rounded-md border bg-zinc-100 dark:bg-zinc-800 transition-opacity"
    >
      <img
        src={`https://picsum.photos/seed/slide-${slideNum}/320/180`}
        alt={`Slide ${slideNum}`}
        className="h-full w-full object-cover pointer-events-none"
      />
      <span className="absolute left-1 top-1 flex h-5 w-5 items-center justify-center rounded bg-black/50 text-[10px] font-medium text-white backdrop-blur-sm">
        {slideNum}
      </span>
    </div>
  );
}

// Droppable container for a tag
function DroppableTagContainer({
  id,
  tagId,
  items,
  children,
  activeId,
}: {
  id: string;
  tagId: string | null;
  items: number[];
  children: React.ReactNode;
  activeId: string | null;
}) {
  const { setNodeRef, isOver } = useDroppable({
    id,
    data: { tagId, type: "container" },
  });

  // Highlight if dragging over
  const isDragOver = isOver;

  return (
    <div
      ref={setNodeRef}
      className={`grid grid-cols-3 gap-2 rounded-md border border-dashed p-2 min-h-[100px] transition-colors ${
        isDragOver
          ? "bg-primary/10 border-primary dark:bg-primary/20"
          : "bg-zinc-50/50 border-zinc-200 dark:bg-zinc-900/50 dark:border-zinc-800"
      }`}
    >
      {children}
      {items.length === 0 && !activeId && (
        <div className="col-span-3 flex h-full items-center justify-center text-xs text-zinc-400 italic">
          drop slides here to start section
        </div>
      )}
    </div>
  );
}

export function TagSlideManager({
  slides,
  tags,
  slideTagMap,
  onSlideMove,
}: TagSlideManagerProps) {
  const [activeId, setActiveId] = useState<string | null>(null);

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 8,
      },
    }),
    useSensor(KeyboardSensor)
  );

  const activeSlideNum = activeId ? parseInt(activeId.replace("slide-", ""), 10) : null;

  // Helper to find the range of slides for a tag
  function getSlidesForTag(tagId: string | null) {
      return slides.filter(s => (slideTagMap[s] ?? null) === tagId);
  }

  // Logic to handle moving boundaries
  function handleDragEnd(event: DragEndEvent) {
    const { active, over } = event;
    setActiveId(null);

    if (!over) return;

    // Source slide info
    const activeSlideNum = active.data.current?.slideNum;
    const sourceTagId = active.data.current?.tagId;
    
    // Target tag info
    let targetTagId: string | null = null;

    if (over.data.current?.type === "container") {
        targetTagId = over.data.current.tagId;
    } else if (over.data.current?.type === "slide") {
        targetTagId = over.data.current.tagId;
    } else {
        // Fallback parsing ID
        const overId = over.id as string;
        if (overId.startsWith("container-")) {
             targetTagId = overId.replace("container-", "");
             if (targetTagId === "untagged") targetTagId = null;
        } else if (overId.startsWith("slide-")) {
             const sNum = parseInt(overId.replace("slide-", ""), 10);
             targetTagId = slideTagMap[sNum] ?? null;
        }
    }
    
    // If dropping on same tag, do nothing
    if (sourceTagId === targetTagId) return;

    // Identify indices of tags to check adjacency
    // Create an ordered list of "tag groups" including untagged if any?
    // Actually, we just care about the tags list order.
    const sourceTagIndex = tags.findIndex(t => t.id === sourceTagId);
    const targetTagIndex = tags.findIndex(t => t.id === targetTagId);

    // Only allow moving to adjacent tags
    const isNext = targetTagIndex === sourceTagIndex + 1;
    const isPrev = targetTagIndex === sourceTagIndex - 1;

    if (!isNext && !isPrev) return; // Enforce adjacency constraint

    // Logic:
    // If moving to NEXT tag:
    //   We are dragging a slide from Source to Target (which is After Source).
    //   This slide, and ALL subsequent slides in Source, should move to Target.
    //   Effectively, the "Start of Target" moves backwards to this slide.
    //
    // If moving to PREV tag:
    //   We are dragging a slide from Source to Target (which is Before Source).
    //   This slide, and ALL preceding slides in Source, should move to Target.
    //   Effectively, the "End of Target" moves forwards to include this slide.

    // Get all slides currently in Source Tag
    const sourceSlides = getSlidesForTag(sourceTagId);
    const slideIndexInSource = sourceSlides.indexOf(activeSlideNum);

    if (slideIndexInSource === -1) return;

    if (isNext) {
        // Move [slideIndexInSource ... end] to Target
        const slidesToMove = sourceSlides.slice(slideIndexInSource);
        slidesToMove.forEach(s => onSlideMove(s, targetTagId));
    } else if (isPrev) {
        // Move [0 ... slideIndexInSource] to Target
        const slidesToMove = sourceSlides.slice(0, slideIndexInSource + 1);
        slidesToMove.forEach(s => onSlideMove(s, targetTagId));
    }
  }

  const dropAnimation: DropAnimation = {
    sideEffects: defaultDropAnimationSideEffects({
      styles: {
        active: {
          opacity: "0.5",
        },
      },
    }),
  };

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      onDragStart={(e) => setActiveId(e.active.id as string)}
      onDragEnd={handleDragEnd}
    >
      <div className="flex flex-col gap-6 pb-20">
        {tags.map((tag) => {
           const slidesInTag = getSlidesForTag(tag.id);
           const containerId = `container-${tag.id}`;
           
           return (
             <div key={tag.id} className="flex flex-col gap-2">
               <div className="flex items-center gap-2 border-b pb-1">
                 <span className="font-medium text-sm text-zinc-700 dark:text-zinc-300">
                   {tag.label}
                 </span>
                 <span className="rounded-full bg-zinc-100 px-2 py-0.5 text-xs text-zinc-500 dark:bg-zinc-800">
                   {slidesInTag.length}
                 </span>
               </div>
               
               <DroppableTagContainer
                  id={containerId}
                  tagId={tag.id}
                  items={slidesInTag}
                  activeId={activeId}
               >
                   {slidesInTag.map((s) => (
                       <DraggableSlide
                          key={`slide-${s}`}
                          id={`slide-${s}`}
                          slideNum={s}
                          tagId={tag.id}
                       />
                   ))}
               </DroppableTagContainer>
             </div>
           );
        })}
      </div>

      <DragOverlay dropAnimation={dropAnimation}>
        {activeSlideNum ? (
            <div className="w-[200px] aspect-video overflow-hidden rounded-md shadow-2xl ring-2 ring-primary">
                 <img
                    src={`https://picsum.photos/seed/slide-${activeSlideNum}/320/180`}
                    alt=""
                    className="h-full w-full object-cover"
                  />
                 <div className="absolute inset-0 flex items-center justify-center bg-black/20 text-white font-medium text-sm">
                    adjusting boundary...
                 </div>
            </div>
        ) : null}
      </DragOverlay>
    </DndContext>
  );
}

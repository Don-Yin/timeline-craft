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
} from "@dnd-kit/core";
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import { GripVertical, X } from "lucide-react";
import { createIndexItem, type IndexItem } from "@/lib/indexes";

export type DraggableIndexListProps = {
  items: IndexItem[];
  onChange: (next: IndexItem[]) => void;
  title?: string;
};

function SortableItem({
  item,
  index,
  onRemove,
}: {
  item: IndexItem;
  index: number;
  onRemove?: (id: string) => void;
}) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: item.id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
    zIndex: isDragging ? 1 : "auto",
  };

  return (
    <li
      ref={setNodeRef}
      style={style}
      className="flex items-center justify-between rounded-md border bg-background px-3 py-2 text-sm transition-colors group"
    >
      <div className="flex items-center gap-3">
        <span className="inline-flex h-6 w-6 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground select-none">
          {index + 1}
        </span>
        <span className="text-zinc-800 dark:text-zinc-200 select-none">
          {item.label}
        </span>
      </div>
      <div className="flex items-center gap-2">
        {onRemove && (
          <button
            onClick={() => onRemove(item.id)}
            className="opacity-0 group-hover:opacity-100 transition-opacity p-1 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-md text-zinc-400 hover:text-red-500"
            title="Remove tag"
          >
            <X className="h-4 w-4" />
          </button>
        )}
        <span
          {...attributes}
          {...listeners}
          className="cursor-grab touch-none text-zinc-500 hover:text-zinc-700 dark:text-zinc-400 dark:hover:text-zinc-200 focus:outline-none"
          title="Drag to reorder"
        >
          <GripVertical className="h-4 w-4" />
        </span>
      </div>
    </li>
  );
}

function ItemOverlay({ item, index }: { item: IndexItem; index: number }) {
  return (
    <div className="flex items-center justify-between rounded-md border bg-background px-3 py-2 text-sm shadow-xl ring-1 ring-zinc-900/10 dark:ring-white/10">
      <div className="flex items-center gap-3">
        <span className="inline-flex h-6 w-6 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground select-none">
          {index + 1}
        </span>
        <span className="text-zinc-800 dark:text-zinc-200 select-none">
          {item.label}
        </span>
      </div>
      <div className="flex items-center gap-2">
        <span className="cursor-grabbing text-zinc-500 dark:text-zinc-400">
          <GripVertical className="h-4 w-4" />
        </span>
      </div>
    </div>
  );
}

export default function DraggableIndexList({
  items,
  onChange,
  title = "tags",
}: DraggableIndexListProps) {
  const [activeId, setActiveId] = useState<string | null>(null);
  const [newItem, setNewItem] = useState<string>("");

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  function handleDragStart(event: DragStartEvent) {
    setActiveId(event.active.id as string);
  }

  function handleDragEnd(event: DragEndEvent) {
    const { active, over } = event;

    if (over && active.id !== over.id) {
      const oldIndex = items.findIndex((item) => item.id === active.id);
      const newIndex = items.findIndex((item) => item.id === over.id);
      onChange(arrayMove(items, oldIndex, newIndex));
    }
    setActiveId(null);
  }

  function handleDragCancel() {
    setActiveId(null);
  }

  function addItem() {
    const value = newItem.trim();
    if (!value) return;
    onChange([...items, createIndexItem(value)]);
    setNewItem("");
  }

  function handleRemove(id: string) {
    onChange(items.filter((item) => item.id !== id));
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLInputElement>) {
    if (e.key === "Enter") {
      e.preventDefault();
      addItem();
    }
  }

  const activeItem = items.find((i) => i.id === activeId);
  const activeIndex = activeItem ? items.indexOf(activeItem) : -1;

  const dropAnimation: DropAnimation = {
    sideEffects: defaultDropAnimationSideEffects({
      styles: {
        active: {
          opacity: "0.4",
        },
      },
    }),
  };

  return (
    <div className="flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <h3 className="text-base font-medium">{title}</h3>
      </div>
      <div className="flex gap-2">
        <input
          type="text"
          placeholder="add tag (e.g., intro)"
          value={newItem}
          onChange={(e) => setNewItem(e.target.value)}
          onKeyDown={handleKeyDown}
          className="flex-1 rounded-md border px-3 py-2 text-sm bg-background text-foreground"
        />
        <button
          onClick={addItem}
          className="rounded-md bg-foreground px-3 py-2 text-sm text-background transition-colors hover:bg-[#383838] dark:hover:bg-[#ccc]"
        >
          add
        </button>
      </div>

      <DndContext
        sensors={sensors}
        collisionDetection={closestCenter}
        onDragStart={handleDragStart}
        onDragEnd={handleDragEnd}
        onDragCancel={handleDragCancel}
      >
        <SortableContext items={items} strategy={verticalListSortingStrategy}>
          <ul className="flex flex-col gap-2">
            {items.map((item, i) => (
              <SortableItem
                key={item.id}
                item={item}
                index={i}
                onRemove={handleRemove}
              />
            ))}
          </ul>
        </SortableContext>
        <DragOverlay dropAnimation={dropAnimation}>
          {activeId && activeItem ? (
            <ItemOverlay item={activeItem} index={activeIndex} />
          ) : null}
        </DragOverlay>
      </DndContext>

      <p className="text-xs text-zinc-500">
        drag items to reorder. this updates the in‑memory order only.
      </p>
    </div>
  );
}

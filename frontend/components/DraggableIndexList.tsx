'use client';
import React, { useState, useRef, useEffect } from "react";
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
import { GripVertical, X, Plus } from "lucide-react";
import { createIndexItem, type IndexItem } from "@/lib/indexes";

export type DraggableIndexListProps = {
  items: IndexItem[];
  onChange: (next: IndexItem[]) => void;
  title?: string;
};

function InsertZone({
  insertAt,
  onInsert,
  isDragging,
}: {
  insertAt: number;
  onInsert: (index: number, label: string) => void;
  isDragging: boolean;
}) {
  const [isHovered, setIsHovered] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [value, setValue] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
    }
  }, [isEditing]);

  function handleSubmit() {
    const trimmed = value.trim();
    if (trimmed) {
      onInsert(insertAt, trimmed);
      setValue("");
    }
    setIsEditing(false);
    setIsHovered(false);
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.key === "Enter") {
      e.preventDefault();
      handleSubmit();
    }
    if (e.key === "Escape") {
      setValue("");
      setIsEditing(false);
      setIsHovered(false);
    }
  }

  function handleBlur() {
    if (!value.trim()) {
      setIsEditing(false);
      setIsHovered(false);
    }
  }

  if (isDragging) {
    return <div className="h-1" />;
  }

  return (
    <div
      className="relative group"
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => !isEditing && setIsHovered(false)}
    >
      <div
        className={`
          transition-all duration-200 overflow-hidden
          ${isEditing ? "h-10 opacity-100" : isHovered ? "h-8 opacity-100" : "h-2 opacity-0 hover:opacity-100"}
        `}
      >
        {isEditing ? (
          <input
            ref={inputRef}
            type="text"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={handleKeyDown}
            onBlur={handleBlur}
            placeholder="type tag name, press Enter"
            className="w-full h-10 px-3 text-sm rounded-md border border-dashed border-emerald-500 bg-emerald-50 dark:bg-emerald-950/30 focus:outline-none focus:ring-2 focus:ring-emerald-500"
          />
        ) : (
          <button
            onClick={() => setIsEditing(true)}
            className="w-full h-full flex items-center justify-center gap-2 rounded-md border border-dashed border-zinc-300 dark:border-zinc-700 hover:border-emerald-500 hover:bg-emerald-50 dark:hover:bg-emerald-950/30 transition-colors text-zinc-400 hover:text-emerald-600"
          >
            <Plus className="h-3 w-3" />
            <span className="text-xs">add tag here</span>
          </button>
        )}
      </div>
    </div>
  );
}

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

  function handleInsert(index: number, label: string) {
    const newItems = [...items];
    newItems.splice(index, 0, createIndexItem(label));
    onChange(newItems);
  }

  function handleRemove(id: string) {
    onChange(items.filter((item) => item.id !== id));
  }

  const activeItem = items.find((i) => i.id === activeId);
  const activeIndex = activeItem ? items.indexOf(activeItem) : -1;
  const isDragging = activeId !== null;

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
    <div className="flex flex-col gap-1">
      <div className="flex items-center justify-between mb-2">
        <h3 className="text-base font-medium">{title}</h3>
        <span className="text-xs text-zinc-500">{items.length} tags</span>
      </div>

      <DndContext
        sensors={sensors}
        collisionDetection={closestCenter}
        onDragStart={handleDragStart}
        onDragEnd={handleDragEnd}
        onDragCancel={handleDragCancel}
      >
        <SortableContext items={items} strategy={verticalListSortingStrategy}>
          <ul className="flex flex-col">
            <InsertZone insertAt={0} onInsert={handleInsert} isDragging={isDragging} />
            {items.map((item, i) => (
              <React.Fragment key={item.id}>
                <SortableItem item={item} index={i} onRemove={handleRemove} />
                <InsertZone insertAt={i + 1} onInsert={handleInsert} isDragging={isDragging} />
              </React.Fragment>
            ))}
          </ul>
        </SortableContext>
        <DragOverlay dropAnimation={dropAnimation}>
          {activeId && activeItem ? (
            <ItemOverlay item={activeItem} index={activeIndex} />
          ) : null}
        </DragOverlay>
      </DndContext>

      <p className="text-xs text-zinc-500 mt-2">
        hover between items to add new tags. drag to reorder.
      </p>
    </div>
  );
}

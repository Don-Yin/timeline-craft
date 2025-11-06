'use client';
import { useState } from "react";
import { createIndexItem, type IndexItem, reorderArray } from "@/lib/indexes";

export type DraggableIndexListProps = {
  items: IndexItem[];
  onChange: (next: IndexItem[]) => void;
  title?: string;
};

export default function DraggableIndexList({
  items,
  onChange,
  title = "tags",
}: DraggableIndexListProps) {
  const [draggingIndex, setDraggingIndex] = useState<number | null>(null);
  const [overIndex, setOverIndex] = useState<number | null>(null);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [dragStartIndex, setDragStartIndex] = useState<number | null>(null);
  const [dragCurrentIndex, setDragCurrentIndex] = useState<number | null>(null);
  const [draftItems, setDraftItems] = useState<IndexItem[] | null>(null);
  const [newItem, setNewItem] = useState<string>("");

  function handleDragStart(e: React.DragEvent<HTMLLIElement>, index: number) {
    setDraggingIndex(index);
    setIsDragging(true);
    setDragStartIndex(index);
    setDragCurrentIndex(index);
    setDraftItems(items.slice());
    e.dataTransfer.setData("text/plain", String(index));
    e.dataTransfer.effectAllowed = "move";
  }

  function handleDragOver(e: React.DragEvent<HTMLLIElement>, index: number) {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    setOverIndex(index);
    if (!isDragging || dragCurrentIndex === null || !draftItems) return;
    if (index === dragCurrentIndex) return;
    const next = reorderArray(draftItems, dragCurrentIndex, index);
    setDraftItems(next);
    setDragCurrentIndex(index);
  }

  function handleDrop(e: React.DragEvent<HTMLLIElement>, targetIndex: number) {
    e.preventDefault();
    if (draftItems) {
      onChange(draftItems);
    }
    resetDrag();
  }

  function handleDragEnd() {
    resetDrag();
  }

  function resetDrag() {
    setDraggingIndex(null);
    setOverIndex(null);
    setIsDragging(false);
    setDragStartIndex(null);
    setDragCurrentIndex(null);
    setDraftItems(null);
  }

  function addItem() {
    const value = newItem.trim();
    if (!value) return;
    onChange([...items, createIndexItem(value)]);
    setNewItem("");
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLInputElement>) {
    if (e.key === "Enter") {
      e.preventDefault();
      addItem();
    }
  }

  const list = isDragging && draftItems ? draftItems : items;

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
          className="flex-1 rounded-md border px-3 py-2 text-sm"
        />
        <button
          onClick={addItem}
          className="rounded-md bg-foreground px-3 py-2 text-sm text-background transition-colors hover:bg-[#383838] dark:hover:bg-[#ccc]"
        >
          add
        </button>
      </div>
      <ul className="flex flex-col gap-2">
        {list.map((item: IndexItem, i: number) => {
          const isDragging = draggingIndex === i;
          const isOver = overIndex === i;
          return (
            <li
              key={item.id}
              draggable
              onDragStart={(e) => handleDragStart(e, i)}
              onDragOver={(e) => handleDragOver(e, i)}
              onDrop={(e) => handleDrop(e, i)}
              onDragEnd={handleDragEnd}
              className={[
                "flex items-center justify-between rounded-md border px-3 py-2 text-sm transition-colors",
                isDragging ? "opacity-60" : "",
                isOver ? "ring-2 ring-offset-1 ring-ring" : "",
              ].join(" ")}
            >
              <div className="flex items-center gap-3">
                <span className="inline-flex h-6 w-6 items-center justify-center rounded bg-secondary text-xs text-secondary-foreground">
                  {i + 1}
                </span>
                <span className="text-zinc-800 dark:text-zinc-200">{item.label}</span>
              </div>
              <span
                aria-hidden
                className="cursor-grab select-none text-zinc-500 dark:text-zinc-400"
                title="Drag to reorder"
              >
                ≡
              </span>
            </li>
          );
        })}
      </ul>
      <p className="text-xs text-zinc-500">
        drag items to reorder. this updates the in‑memory order only.
      </p>
    </div>
  );
}



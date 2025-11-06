export type IndexItem = {
  id: string;
  label: string;
};

export function createIndexItem(label: string): IndexItem {
  const safe = label.trim();
  const id =
    (typeof crypto !== "undefined" && "randomUUID" in crypto
      ? (crypto as any).randomUUID()
      : `idx_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`) as string;
  return { id, label: safe };
}

export function reorderArray<T>(items: T[], fromIndex: number, toIndex: number): T[] {
  if (fromIndex === toIndex) return items.slice();
  const copy = items.slice();
  const [moved] = copy.splice(fromIndex, 1);
  copy.splice(toIndex, 0, moved);
  return copy;
}



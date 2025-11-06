export function generateDummySlides(count: number = 36): number[] {
  return Array.from({ length: count }, (_, i) => i + 1);
}



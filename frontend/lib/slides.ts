export interface Slide {
  id: number;
  src: string;
}

export function generateDummySlides(count: number = 36): Slide[] {
  return Array.from({ length: count }, (_, i) => ({
    id: i + 1,
    src: "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", // 1x1 white pixel
  }));
}

import { WORKER_SERVICE_URL } from './config';
import type { AllThumbnailsResponse, PreviewParams, PreviewThumbnailsResult, SlideCountResponse, ThumbnailResponse } from './types';

export async function getSlideCount(fileId: string): Promise<number> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-slide-count/${fileId}`);
    if (!res.ok) {
        return 0;
    }
    const data: SlideCountResponse = await res.json();
    return data.slide_count;
}

export async function getThumbnail(fileId: string, slideIndex: number): Promise<string> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-thumbnail/${fileId}/${slideIndex}`);
    if (!res.ok) {
        return '';
    }
    const data: ThumbnailResponse = await res.json();
    return data.image_base64;
}

export async function getAllThumbnails(fileId: string): Promise<string[]> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-all-thumbnails/${fileId}`);
    if (!res.ok) {
        return [];
    }
    const data: AllThumbnailsResponse = await res.json();
    return data.thumbnails;
}

export async function getPreviewThumbnail(
    fileId: string,
    slideIndex: number,
    params: PreviewParams,
    signal?: AbortSignal
): Promise<string> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-preview-thumbnail/${fileId}/${slideIndex}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
        signal,
    });
    if (!res.ok) {
        return '';
    }
    const data: ThumbnailResponse = await res.json();
    return data.image_base64;
}

export async function getAllPreviewThumbnails(
    fileId: string,
    params: PreviewParams,
    signal?: AbortSignal
): Promise<PreviewThumbnailsResult> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-all-preview-thumbnails/${fileId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
        signal,
    });
    if (!res.ok) {
        return { thumbnails: [], format: 'png' };
    }
    const data: AllThumbnailsResponse = await res.json();
    return { thumbnails: data.thumbnails, format: data.format || 'png' };
}


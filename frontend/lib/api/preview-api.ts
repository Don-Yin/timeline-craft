import { PREVIEW_SERVICE_URL } from './config';
import type { PreviewParams, PreviewProgressEvent, PreviewThumbnailsResult } from './types';

export async function getFirstSlidePreview(
    fileId: string,
    params: PreviewParams,
    onProgress: (event: PreviewProgressEvent) => void,
    signal?: AbortSignal
): Promise<PreviewThumbnailsResult> {
    const res = await fetch(`${PREVIEW_SERVICE_URL}/render-first-slide/${fileId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
        signal,
    });

    if (!res.ok) {
        throw new Error('Failed to start preview rendering');
    }

    const reader = res.body?.getReader();
    if (!reader) {
        throw new Error('No response body');
    }

    const decoder = new TextDecoder();
    let buffer = '';
    let result: PreviewThumbnailsResult = { thumbnails: [], format: 'jpeg' };

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
            if (line.startsWith('data: ')) {
                const data = JSON.parse(line.slice(6)) as PreviewProgressEvent;
                onProgress(data);

                if (data.stage === 'done' && data.thumbnails) {
                    result = { thumbnails: data.thumbnails, format: data.format || 'jpeg' };
                }
            }
        }
    }

    return result;
}

export async function getPreviewsWithProgress(
    fileId: string,
    params: PreviewParams,
    onProgress: (event: PreviewProgressEvent) => void,
    signal?: AbortSignal
): Promise<PreviewThumbnailsResult> {
    const res = await fetch(`${PREVIEW_SERVICE_URL}/render-previews-with-sidebar/${fileId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
        signal,
    });

    if (!res.ok) {
        throw new Error('Failed to start preview rendering');
    }

    const reader = res.body?.getReader();
    if (!reader) {
        throw new Error('No response body');
    }

    const decoder = new TextDecoder();
    let buffer = '';
    let result: PreviewThumbnailsResult = { thumbnails: [], format: 'jpeg' };

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
            if (line.startsWith('data: ')) {
                const data = JSON.parse(line.slice(6)) as PreviewProgressEvent;
                onProgress(data);

                if (data.stage === 'done' && data.thumbnails) {
                    result = { thumbnails: data.thumbnails, format: data.format || 'jpeg' };
                }
            }
        }
    }

    return result;
}


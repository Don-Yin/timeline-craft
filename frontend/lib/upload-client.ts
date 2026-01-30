const UPLOAD_SERVICE_URL = process.env.NEXT_PUBLIC_UPLOAD_SERVICE_URL || 'http://localhost:8003';
const WORKER_SERVICE_URL = process.env.NEXT_PUBLIC_WORKER_SERVICE_URL || 'http://localhost:8002';
const PREVIEW_SERVICE_URL = process.env.NEXT_PUBLIC_PREVIEW_SERVICE_URL || 'http://localhost:8004';

export type FileMetadata = {
    id: string;
    filename: string;
    size: number;
    content_type: string | null;
    last_modified: string | null;
};

export type ListFilesResponse = {
    files: string[];
};

export type ThumbnailResponse = {
    slide_index: number;
    image_base64: string;
};

export type SlideCountResponse = {
    file_id: string;
    slide_count: number;
};

export async function listFiles(): Promise<string[]> {
    const res = await fetch(`${UPLOAD_SERVICE_URL}/list`);
    if (!res.ok) {
        throw new Error('Failed to list files');
    }
    const data: ListFilesResponse = await res.json();
    return data.files;
}

export async function getFileMetadata(id: string): Promise<FileMetadata> {
    const res = await fetch(`${UPLOAD_SERVICE_URL}/check-metadata/${id}`);
    if (!res.ok) {
        throw new Error('Failed to get file metadata');
    }
    return res.json();
}

export async function uploadFile(file: File): Promise<FileMetadata> {
    const formData = new FormData();
    formData.append('file', file);

    const res = await fetch(`${UPLOAD_SERVICE_URL}/upload`, {
        method: 'POST',
        body: formData,
    });

    if (!res.ok) {
        throw new Error('Failed to upload file');
    }

    const data = await res.json();
    return {
        id: data.id,
        filename: data.filename,
        size: file.size,
        content_type: file.type,
        last_modified: new Date().toISOString()
    };
}

export async function deleteFile(id: string): Promise<void> {
    const res = await fetch(`${UPLOAD_SERVICE_URL}/delete/${id}`, {
        method: 'DELETE',
    });

    if (!res.ok) {
        throw new Error('Failed to delete file');
    }
}

export async function getSlideCount(fileId: string): Promise<number> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-slide-count/${fileId}`);
    if (!res.ok) {
        // Return 0 or throw? Better return 0 to handle gracefully
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

export type AllThumbnailsResponse = {
    file_id: string;
    thumbnails: string[];
    format?: string; // 'png' or 'jpeg'
};

export async function getAllThumbnails(fileId: string): Promise<string[]> {
    const res = await fetch(`${WORKER_SERVICE_URL}/get-all-thumbnails/${fileId}`);
    if (!res.ok) {
        return [];
    }
    const data: AllThumbnailsResponse = await res.json();
    return data.thumbnails;
}

export type PreviewParams = {
    tags: string[];
    sidebar_width: number;
    sidebar_item_height: number;
};

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

export type PreviewThumbnailsResult = {
    thumbnails: string[];
    format: string;
};

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

export type PreviewProgressEvent = {
    stage: 'processing' | 'converting' | 'rendering' | 'done' | 'error';
    progress: number;
    message: string;
    current_slide?: number;
    total_slides?: number;
    thumbnails?: string[];
    format?: string;
};

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

export type ProcessParams = {
    tags: string[];
    sidebar_width: number;
    sidebar_item_height: number;
    transition_duration: number;
    apply_morph_transition: boolean;
};

export type ProgressEvent = {
    stage: string;
    progress: number;
    message: string;
    job_id?: string;
};

export async function processWithProgress(
    fileId: string,
    params: ProcessParams,
    onProgress: (event: ProgressEvent) => void
): Promise<void> {
    const res = await fetch(`${WORKER_SERVICE_URL}/process-with-progress/${fileId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(params),
    });

    if (!res.ok) {
        throw new Error('Failed to start processing');
    }

    const reader = res.body?.getReader();
    if (!reader) {
        throw new Error('No response body');
    }

    const decoder = new TextDecoder();
    let buffer = '';
    let jobId: string | null = null;

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
            if (line.startsWith('data: ')) {
                const data = JSON.parse(line.slice(6)) as ProgressEvent;
                onProgress(data);
                if (data.job_id) {
                    jobId = data.job_id;
                }
            }
        }
    }

    if (jobId) {
        onProgress({ stage: 'downloading', progress: 100, message: 'downloading file...' });

        const downloadRes = await fetch(`${WORKER_SERVICE_URL}/download-processed/${jobId}`);
        if (!downloadRes.ok) {
            throw new Error('Failed to download processed file');
        }

        onProgress({ stage: 'downloading', progress: 100, message: 'preparing download...' });
        const blob = await downloadRes.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `timeline-${fileId}.pptx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
}

export async function processAndDownload(fileId: string, params: ProcessParams): Promise<void> {
    const res = await fetch(`${WORKER_SERVICE_URL}/process-and-download/${fileId}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(params),
    });

    if (!res.ok) {
        throw new Error('Failed to process file');
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `timeline-${fileId}.pptx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

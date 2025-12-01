const UPLOAD_SERVICE_URL = process.env.NEXT_PUBLIC_UPLOAD_SERVICE_URL || 'http://localhost:8001';
const WORKER_SERVICE_URL = process.env.NEXT_PUBLIC_WORKER_SERVICE_URL || 'http://localhost:8002';

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
        // Fallback or throw
        // throw new Error('Failed to get thumbnail');
        // Return a placeholder or empty string to handle gracefully in UI
        return '';
    }
    const data: ThumbnailResponse = await res.json();
    return data.image_base64;
}

export async function processFile(fileId: string, params: any): Promise<void> {
    const res = await fetch(`${WORKER_SERVICE_URL}/process-file/${fileId}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(params),
    });

    if (!res.ok) {
        throw new Error('Failed to process file');
    }
}

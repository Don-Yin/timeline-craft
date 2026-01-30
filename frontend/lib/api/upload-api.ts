import { UPLOAD_SERVICE_URL } from './config';
import type { FileMetadata, ListFilesResponse } from './types';

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


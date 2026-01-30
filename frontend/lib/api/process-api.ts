import { WORKER_SERVICE_URL } from './config';
import type { ProcessParams, ProgressEvent } from './types';

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

        const downloadRes = await fetch(`${WORKER_SERVICE_URL}/download-processed/${jobId}?file_id=${encodeURIComponent(fileId)}`);
        if (!downloadRes.ok) {
            throw new Error('Failed to download processed file');
        }

        onProgress({ stage: 'downloading', progress: 100, message: 'preparing download...' });
        const blob = await downloadRes.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `timeline-${fileId}-${jobId}.pptx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
}

export async function processAndDownload(fileId: string, params: ProcessParams): Promise<void> {
    const res = await fetch(`${WORKER_SERVICE_URL}/process-and-download/${fileId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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


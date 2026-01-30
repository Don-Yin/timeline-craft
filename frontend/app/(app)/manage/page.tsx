'use client';

import { useState, useEffect } from 'react';
import { Upload, FileText, Download, Trash2 } from 'lucide-react';
import {
  listFiles,
  getFileMetadata,
  uploadFile,
  deleteFile,
  type FileMetadata,
} from '@/lib/api';

export default function Manage() {
  const [files, setFiles] = useState<FileMetadata[]>([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);

  useEffect(() => {
    loadFiles();
  }, []);

  async function loadFiles() {
    try {
      const fileIds = await listFiles();
      const metadataPromises = fileIds.map((id) => getFileMetadata(id));
      const filesData = await Promise.all(metadataPromises);
      setFiles(filesData);
    } catch (error) {
      console.error('Failed to load files:', error);
    } finally {
      setLoading(false);
    }
  }

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setUploading(true);
      try {
        const newFiles = Array.from(e.target.files);
        // Upload sequentially or parallel? Parallel for now
        const uploadPromises = newFiles.map((file) => uploadFile(file));
        const uploadedMetadata = await Promise.all(uploadPromises);

        // Refresh list or append
        // Appending is faster UI feedback
        setFiles((prev) => [...prev, ...uploadedMetadata]);
      } catch (error) {
        console.error('Upload failed:', error);
        alert('Failed to upload one or more files.');
      } finally {
        setUploading(false);
        // reset input
        e.target.value = '';
      }
    }
  };

  const removeFile = async (id: string) => {
    if (!confirm('Are you sure you want to delete this file?')) return;
    try {
      await deleteFile(id);
      setFiles((prev) => prev.filter((f) => f.id !== id));
    } catch (error) {
      console.error('Delete failed:', error);
      alert('Failed to delete file.');
    }
  };

  const downloadFile = (file: FileMetadata) => {
    // Direct link to minio or signed url?
    // For now, we don't have a direct download endpoint in the upload-service exposed as a GET stream
    // The requirement was just "download". 
    // If the upload service doesn't have a GET /download/{id}, we can't easily download.
    // But wait, the requirements didn't explicitly ask for a download endpoint implementation in the backend yet,
    // only that the UI should have a button.
    // Let's assume we can't download yet or we need to add it. 
    // Actually, for now let's alert not implemented or maybe add a TODO.
    alert('Download not implemented yet in backend');
  };

  if (loading) {
    return (
      <div className="flex min-h-screen w-full items-center justify-center bg-zinc-50 dark:bg-black">
        <div className="text-sm text-zinc-500">Loading files...</div>
      </div>
    );
  }

  return (
    <div className="flex min-h-screen w-full flex-col gap-8 bg-zinc-50 p-6 dark:bg-black sm:p-12">
      <div className="flex flex-col gap-2">
        <h1 className="text-2xl font-semibold">Manage PPTX</h1>
      </div>

      <section className="grid grid-cols-2 gap-4 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5">
        {/* Upload Slot */}
        <label className={`group relative flex aspect-[3/4] cursor-pointer flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed border-zinc-200 bg-white p-4 text-center transition-all hover:border-zinc-300 hover:bg-zinc-50 dark:border-zinc-800 dark:bg-zinc-900/50 dark:hover:border-zinc-700 dark:hover:bg-zinc-900 ${uploading ? 'pointer-events-none opacity-50' : ''}`}>
          <div className="rounded-full bg-zinc-100 p-4 transition-colors group-hover:bg-zinc-200 dark:bg-zinc-800 dark:group-hover:bg-zinc-700">
            <Upload className="h-6 w-6 text-zinc-600 dark:text-zinc-400" />
          </div>
          <div className="flex flex-col gap-1">
            <span className="text-sm font-medium">{uploading ? 'Uploading...' : 'Upload PPTX'}</span>
            <span className="text-xs text-zinc-500">{uploading ? 'Please wait' : 'Click to browse'}</span>
          </div>
          <input
            type="file"
            accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
            multiple
            className="hidden"
            onChange={handleFileChange}
            disabled={uploading}
          />
        </label>

        {/* File Slots */}
        {files.map((file) => (
          <div
            key={file.id}
            onClick={() => window.location.href = `/operate/${file.id}`}
            className="group relative flex aspect-[3/4] cursor-pointer flex-col justify-between rounded-xl border bg-white p-4 shadow-sm transition-all hover:shadow-md hover:ring-2 hover:ring-emerald-500/50 dark:border-zinc-800 dark:bg-zinc-900"
          >
            <div className="flex flex-1 flex-col items-center justify-center gap-3">
              <div className="rounded-lg bg-red-50 p-3 dark:bg-red-900/20">
                <FileText className="h-8 w-8 text-red-600 dark:text-red-400" />
              </div>
              <div className="w-full text-center">
                <p className="truncate text-sm font-medium" title={file.filename}>
                  {file.filename}
                </p>
                <p className="mt-1 text-xs text-zinc-500">
                  {(file.size / 1024).toFixed(1)} KB
                </p>
              </div>
            </div>

            <div className="flex items-center justify-center gap-1 pt-2 opacity-0 transition-opacity group-hover:opacity-100">
              <button
                onClick={(e) => { e.stopPropagation(); downloadFile(file); }}
                className="rounded-md p-1.5 text-zinc-500 hover:bg-zinc-100 hover:text-zinc-900 dark:hover:bg-zinc-800 dark:hover:text-zinc-100"
                title="Download"
              >
                <Download className="h-4 w-4" />
              </button>
              <button
                onClick={(e) => { e.stopPropagation(); removeFile(file.id); }}
                className="rounded-md p-1.5 text-zinc-500 hover:bg-red-50 hover:text-red-600 dark:hover:bg-red-900/20 dark:hover:text-red-400"
                title="Remove"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>
          </div>
        ))}
      </section>
    </div>
  );
}


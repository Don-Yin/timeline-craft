'use client';

import { useState } from 'react';
import { Upload, FileText, Download, Trash2 } from 'lucide-react';

export default function Manage() {
  const [files, setFiles] = useState<File[]>([]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files);
      setFiles((prev) => [...prev, ...newFiles]);
    }
  };

  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const downloadFile = (file: File) => {
    const url = URL.createObjectURL(file);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.name;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="flex min-h-screen w-full flex-col gap-8 bg-zinc-50 p-6 dark:bg-black sm:p-12">
      <div className="flex flex-col gap-2">
        <h1 className="text-2xl font-semibold">Manage PDFs</h1>
        <p className="text-sm text-zinc-500">
          Upload and manage your PDF documents.
        </p>
      </div>

      <section className="grid grid-cols-2 gap-4 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5">
        {/* Upload Slot */}
        <label className="group relative flex aspect-[3/4] cursor-pointer flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed border-zinc-200 bg-white p-4 text-center transition-all hover:border-zinc-300 hover:bg-zinc-50 dark:border-zinc-800 dark:bg-zinc-900/50 dark:hover:border-zinc-700 dark:hover:bg-zinc-900">
          <div className="rounded-full bg-zinc-100 p-4 transition-colors group-hover:bg-zinc-200 dark:bg-zinc-800 dark:group-hover:bg-zinc-700">
            <Upload className="h-6 w-6 text-zinc-600 dark:text-zinc-400" />
          </div>
          <div className="flex flex-col gap-1">
            <span className="text-sm font-medium">Upload PDF</span>
            <span className="text-xs text-zinc-500">Click to browse</span>
          </div>
          <input
            type="file"
            accept=".pdf"
            multiple
            className="hidden"
            onChange={handleFileChange}
          />
        </label>

        {/* File Slots */}
        {files.map((file, i) => (
          <div
            key={i}
            className="group relative flex aspect-[3/4] flex-col justify-between rounded-xl border bg-white p-4 shadow-sm transition-all hover:shadow-md dark:border-zinc-800 dark:bg-zinc-900"
          >
            <div className="flex flex-1 flex-col items-center justify-center gap-3">
              <div className="rounded-lg bg-red-50 p-3 dark:bg-red-900/20">
                <FileText className="h-8 w-8 text-red-600 dark:text-red-400" />
              </div>
              <div className="w-full text-center">
                <p className="truncate text-sm font-medium" title={file.name}>
                  {file.name}
                </p>
                <p className="mt-1 text-xs text-zinc-500">
                  {(file.size / 1024).toFixed(1)} KB
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-1 pt-2 opacity-0 transition-opacity group-hover:opacity-100">
              <button
                onClick={() => downloadFile(file)}
                className="rounded-md p-1.5 text-zinc-500 hover:bg-zinc-100 hover:text-zinc-900 dark:hover:bg-zinc-800 dark:hover:text-zinc-100"
                title="Download"
              >
                <Download className="h-4 w-4" />
              </button>
              <button
                onClick={() => removeFile(i)}
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


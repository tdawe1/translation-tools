'use client';

import { useState } from 'react';
import { File, CheckCircle, XCircle, Upload } from 'lucide-react';

export default function Home() {
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [currentJobId, setCurrentJobId] = useState<string | null>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setError(null);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await fetch('http://localhost:8000/upload', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Upload failed');
      }

      const data = await response.json();
      setUploadedFile(file);
      console.log('Upload successful:', data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Upload failed');
    } finally {
      setIsUploading(false);
    }
  };

  const handleTranslate = async () => {
    if (!uploadedFile) return;

    try {
      const response = await fetch('http://localhost:8000/translate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          input_path: uploadedFile.name,
          output_path: uploadedFile.name.replace(/\.(\w+)$/, '_translated.$1'),
          model: 'gpt-4o-mini',
        }),
      });

      if (!response.ok) {
        throw new Error('Translation failed');
      }

      const data = await response.json();
      setCurrentJobId(data.job_id);
      console.log('Translation started:', data);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Translation failed');
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-4xl mx-auto px-4 py-12">
        <div className="text-center mb-12">
          <h1 className="text-4xl font-bold text-gray-900 mb-4">
            Translation Pipeline
          </h1>
          <p className="text-xl text-gray-600">
            Translate PowerPoint presentations and PDF documents while preserving formatting
          </p>
        </div>

        <div className="bg-white rounded-lg shadow-sm border p-8 mb-8">
          {error && (
            <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-md">
              <div className="flex items-center">
                <XCircle className="h-5 w-5 text-red-400 mr-2" />
                <p className="text-sm text-red-700">{error}</p>
              </div>
            </div>
          )}

          <div className="mb-6">
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Upload Document
            </label>
            <div className="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-gray-300 border-dashed rounded-md">
              <div className="space-y-1 text-center">
                <Upload className="mx-auto h-12 w-12 text-gray-400" />
                <div className="flex text-sm text-gray-600">
                  <label className="relative cursor-pointer bg-white rounded-md font-medium text-indigo-600 hover:text-indigo-500">
                    <span>Upload a file</span>
                    <input
                      type="file"
                      className="sr-only"
                      accept=".pptx,.pdf"
                      onChange={handleFileUpload}
                      disabled={isUploading}
                    />
                  </label>
                  <p className="pl-1">or drag and drop</p>
                </div>
                <p className="text-xs text-gray-500">
                  PowerPoint or PDF up to 50MB
                </p>
              </div>
            </div>
          </div>

          {uploadedFile && (
            <div className="mb-6">
              <div className="flex items-center justify-between p-4 bg-gray-50 rounded-md">
                <div>
                  <p className="text-sm font-medium text-gray-900">{uploadedFile.name}</p>
                  <p className="text-sm text-gray-500">
                    {(uploadedFile.size / 1024 / 1024).toFixed(2)} MB
                  </p>
                </div>
                <button
                  onClick={handleTranslate}
                  disabled={isUploading || !!currentJobId}
                  className="px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50"
                >
                  {isUploading ? 'Processing...' : 'Start Translation'}
                </button>
              </div>
            </div>
          )}

          {currentJobId && (
            <div className="mt-6 p-4 bg-green-50 border border-green-200 rounded-md">
              <div className="flex items-center">
                <CheckCircle className="h-5 w-5 text-green-400 mr-2" />
                <p className="text-sm text-green-700">
                  Translation job started! Job ID: {currentJobId}
                </p>
              </div>
            </div>
          )}
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          <div className="bg-white rounded-lg border p-6">
            <div className="flex items-center justify-center h-12 w-12 rounded-md bg-indigo-100 text-indigo-600 mb-4 mx-auto">
              <File className="h-6 w-6" />
            </div>
            <h3 className="text-lg font-medium text-gray-900 mb-2 text-center">1. Upload</h3>
            <p className="text-gray-600 text-sm text-center">
              Upload your PowerPoint or PDF document
            </p>
          </div>

          <div className="bg-white rounded-lg border p-6">
            <div className="flex items-center justify-center h-12 w-12 rounded-md bg-indigo-100 text-indigo-600 mb-4 mx-auto">
              <svg className="h-6 w-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 8h10M7 12h4m1 8l-4-4H5a2 2 0 01-2-2V6a2 2 0 012-2h14a2 2 0 012 2v8a2 2 0 01-2 2h-3l-4 4z" />
              </svg>
            </div>
            <h3 className="text-lg font-medium text-gray-900 mb-2 text-center">2. Translate</h3>
            <p className="text-gray-600 text-sm text-center">
              AI-powered translation preserves formatting
            </p>
          </div>

          <div className="bg-white rounded-lg border p-6">
            <div className="flex items-center justify-center h-12 w-12 rounded-md bg-indigo-100 text-indigo-600 mb-4 mx-auto">
              <CheckCircle className="h-6 w-6" />
            </div>
            <h3 className="text-lg font-medium text-gray-900 mb-2 text-center">3. Download</h3>
            <p className="text-gray-600 text-sm text-center">
              Get your translated document
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
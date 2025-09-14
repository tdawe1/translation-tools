'use client';

import { useState } from 'react';
import { Play } from 'lucide-react';
import { startTranslation } from '@/lib/api';

interface TranslationFormProps {
  uploadedFile: {
    file_id: string;
    filename: string;
  };
  onTranslationStart: (jobId: string) => void;
  onError: (error: string) => void;
}

export default function TranslationForm({
  uploadedFile,
  onTranslationStart,
  onError
}: TranslationFormProps) {
  const [isStarting, setIsStarting] = useState(false);
  const [selectedModel, setSelectedModel] = useState('gpt-4o');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsStarting(true);

    try {
      const response = await startTranslation(
        uploadedFile.file_id,
        uploadedFile.filename,
        selectedModel
      );
      onTranslationStart(response.job_id);
    } catch (error: any) {
      onError(error.response?.data?.detail || 'Failed to start translation');
    } finally {
      setIsStarting(false);
    }
  };

  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6">
      <h3 className="text-lg font-medium text-gray-900 mb-4">
        Start Translation
      </h3>

      <form onSubmit={handleSubmit} className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Selected File
          </label>
          <p className="text-sm text-gray-600">{uploadedFile.filename}</p>
        </div>

        <div>
          <label htmlFor="model" className="block text-sm font-medium text-gray-700 mb-2">
            Translation Model
          </label>
          <select
            id="model"
            value={selectedModel}
            onChange={(e) => setSelectedModel(e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500"
            disabled={isStarting}
          >
            <option value="gpt-4o">GPT-4o (Recommended)</option>
            <option value="gpt-4o-mini">GPT-4o Mini (Faster)</option>
            <option value="gpt-5">GPT-5 (Best Quality)</option>
          </select>
        </div>

        <button
          type="submit"
          disabled={isStarting}
          className="w-full flex items-center justify-center space-x-2 px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50"
        >
          {isStarting ? (
            <>
              <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white"></div>
              <span>Starting...</span>
            </>
          ) : (
            <>
              <Play className="h-4 w-4" />
              <span>Start Translation</span>
            </>
          )}
        </button>
      </form>
    </div>
  );
}
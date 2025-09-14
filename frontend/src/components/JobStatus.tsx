'use client';

import { useState, useEffect } from 'react';
import { CheckCircle, XCircle, Clock, Download, Loader } from 'lucide-react';
import { TranslationJob, getJobStatus, downloadResult } from '@/lib/api';

interface JobStatusProps {
  jobId: string;
  onComplete?: () => void;
}

export default function JobStatus({ jobId, onComplete }: JobStatusProps) {
  const [job, setJob] = useState<TranslationJob | null>(null);
  const [isDownloading, setIsDownloading] = useState(false);

  useEffect(() => {
    const fetchJobStatus = async () => {
      try {
        const jobData = await getJobStatus(jobId);
        setJob(jobData);

        if (jobData.status === 'completed') {
          onComplete?.();
        } else if (jobData.status === 'pending' || jobData.status === 'running') {
          setTimeout(fetchJobStatus, 2000);
        }
      } catch (error) {
        console.error('Failed to fetch job status:', error);
      }
    };

    fetchJobStatus();
  }, [jobId, onComplete]);

  const handleDownload = async () => {
    if (!job) return;

    setIsDownloading(true);
    try {
      const blob = await downloadResult(jobId);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = job.filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (error) {
      console.error('Download failed:', error);
    } finally {
      setIsDownloading(false);
    }
  };

  const getStatusIcon = () => {
    if (!job) return <Loader className="h-5 w-5 animate-spin" />;

    switch (job.status) {
      case 'pending':
        return <Clock className="h-5 w-5 text-gray-500" />;
      case 'running':
        return <Loader className="h-5 w-5 animate-spin text-indigo-600" />;
      case 'completed':
        return <CheckCircle className="h-5 w-5 text-green-500" />;
      case 'failed':
        return <XCircle className="h-5 w-5 text-red-500" />;
    }
  };

  const getStatusText = () => {
    if (!job) return 'Loading...';

    switch (job.status) {
      case 'pending':
        return 'Waiting to start';
      case 'running':
        return `Translating (${job.progress}%)`;
      case 'completed':
        return 'Translation complete';
      case 'failed':
        return `Failed: ${job.error || 'Unknown error'}`;
    }
  };

  if (!job) {
    return (
      <div className="flex items-center space-x-2 text-gray-600">
        <Loader className="h-5 w-5 animate-spin" />
        <span>Loading job status...</span>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-lg border border-gray-200 p-6 space-y-4">
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-3">
          {getStatusIcon()}
          <div>
            <h3 className="font-medium text-gray-900">{job.filename}</h3>
            <p className="text-sm text-gray-600">{getStatusText()}</p>
          </div>
        </div>
        {job.status === 'completed' && (
          <button
            onClick={handleDownload}
            disabled={isDownloading}
            className="flex items-center space-x-2 px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 disabled:opacity-50"
          >
            {isDownloading ? (
              <Loader className="h-4 w-4 animate-spin" />
            ) : (
              <Download className="h-4 w-4" />
            )}
            <span>Download</span>
          </button>
        )}
      </div>

      {(job.status === 'pending' || job.status === 'running') && (
        <div className="space-y-2">
          <div className="w-full bg-gray-200 rounded-full h-2">
            <div
              className="bg-indigo-600 h-2 rounded-full transition-all duration-300"
              style={{ width: `${job.progress}%` }}
            />
          </div>
          <p className="text-xs text-gray-500">
            Model: {job.model}
          </p>
        </div>
      )}

      {job.status === 'failed' && job.error && (
        <div className="bg-red-50 border border-red-200 rounded-md p-3">
          <p className="text-sm text-red-700">{job.error}</p>
        </div>
      )}
    </div>
  );
}
'use client';

import { useState, useEffect } from 'react';
import { FileText, CheckCircle, XCircle, Clock, Download } from 'lucide-react';
import { TranslationJob, listJobs, downloadResult } from '@/lib/api';

export default function JobsList() {
  const [jobs, setJobs] = useState<TranslationJob[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchJobs();
    const interval = setInterval(fetchJobs, 5000);
    return () => clearInterval(interval);
  }, []);

  const fetchJobs = async () => {
    try {
      const jobsData = await listJobs();
      setJobs(jobsData);
    } catch (error) {
      console.error('Failed to fetch jobs:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async (jobId: string, filename: string) => {
    try {
      const blob = await downloadResult(jobId);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (error) {
      console.error('Download failed:', error);
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'pending':
        return <Clock className="h-4 w-4 text-gray-500" />;
      case 'running':
        return <div className="h-4 w-4 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin" />;
      case 'completed':
        return <CheckCircle className="h-4 w-4 text-green-500" />;
      case 'failed':
        return <XCircle className="h-4 w-4 text-red-500" />;
      default:
        return <Clock className="h-4 w-4 text-gray-500" />;
    }
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center py-8">
        <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div>
      </div>
    );
  }

  if (jobs.length === 0) {
    return (
      <div className="text-center py-8 text-gray-500">
        <FileText className="h-12 w-12 mx-auto mb-2 text-gray-400" />
        <p>No translation jobs yet</p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <h3 className="text-lg font-medium text-gray-900">Recent Jobs</h3>
      <div className="bg-white rounded-lg border border-gray-200 divide-y divide-gray-200">
        {jobs.map((job) => (
          <div key={job.id} className="p-4 hover:bg-gray-50">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3">
                {getStatusIcon(job.status)}
                <div>
                  <p className="text-sm font-medium text-gray-900">{job.filename}</p>
                  <p className="text-xs text-gray-500">
                    {job.model} • {new Date(job.created_at).toLocaleString()}
                  </p>
                </div>
              </div>
              <div className="flex items-center space-x-2">
                <span className={`text-xs px-2 py-1 rounded-full ${
                  job.status === 'completed' ? 'bg-green-100 text-green-800' :
                  job.status === 'failed' ? 'bg-red-100 text-red-800' :
                  job.status === 'running' ? 'bg-indigo-100 text-indigo-800' :
                  'bg-gray-100 text-gray-800'
                }`}>
                  {job.status}
                </span>
                {job.status === 'completed' && (
                  <button
                    onClick={() => handleDownload(job.id, job.filename)}
                    className="p-1 text-gray-400 hover:text-indigo-600 transition-colors"
                    title="Download"
                  >
                    <Download className="h-4 w-4" />
                  </button>
                )}
              </div>
            </div>
            {(job.status === 'pending' || job.status === 'running') && (
              <div className="mt-2 w-full bg-gray-200 rounded-full h-1">
                <div
                  className="bg-indigo-600 h-1 rounded-full transition-all duration-300"
                  style={{ width: `${job.progress}%` }}
                />
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { useAuth } from '@/contexts/AuthContext';
import { apiClient, type TranslationJob } from '@/lib/api';
import { JobQueueTable } from '@/components/jobs/JobQueueTable';
import { JobDetailModal } from '@/components/jobs/JobDetailModal';
import {
  FileText,
  Search,
  Filter,
  RefreshCw,
  Download,
  Eye,
  Play,
  X,
  CheckCircle,
  XCircle,
  Clock,
  AlertCircle,
  Trash2,
  Calendar,
  DollarSign,
  BarChart3,
  Upload,
  Plus,
} from 'lucide-react';

interface JobQueue {
  pending: number;
  processing: number;
  completed: number;
  failed: number;
  active_jobs: number;
  total_jobs: number;
}

interface JobStatistics {
  total_jobs: number;
  status_counts: Record<string, number>;
  average_duration_minutes: number;
  total_cost: number;
  daily_stats: any[];
  file_type_distribution: Record<string, number>;
  period_days: number;
}

export default function JobsPage() {
  const { isAuthenticated } = useAuth();
  const router = useRouter();
  const [jobs, setJobs] = useState<TranslationJob[]>([]);
  const [queueStats, setQueueStats] = useState<JobQueue | null>(null);
  const [jobStats, setJobStats] = useState<any>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [selectedJob, setSelectedJob] = useState<TranslationJob | null>(null);
  const [selectedJobs, setSelectedJobs] = useState<string[]>([]);
  const [isJobModalOpen, setIsJobModalOpen] = useState(false);
  const [sortField, setSortField] = useState<string>('createdAt');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('desc');
  const [actionLoading, setActionLoading] = useState<Record<string, boolean>>({});

  useEffect(() => {
    if (!isAuthenticated) {
      router.push('/login');
      return;
    }

    fetchJobs();
    fetchQueueStats();
    fetchJobStatistics();

    // Set up periodic refresh for active jobs
    const interval = setInterval(() => {
      const hasActiveJobs = jobs.some(job => job.status === 'processing' || job.status === 'pending');
      if (hasActiveJobs) {
        fetchJobs();
        fetchQueueStats();
      }
    }, 5000);

    return () => clearInterval(interval);
  }, [isAuthenticated, router, jobs]);

  const fetchJobs = async () => {
    try {
      const response = await apiClient.getJobs();
      if (response.success && response.data) {
        setJobs(response.data.jobs);
      }
    } catch (error) {
      console.error('Failed to fetch jobs:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const fetchQueueStats = async () => {
    try {
      const response = await apiClient.getJobQueue();
      if (response.success && response.data) {
        setQueueStats(response.data);
      }
    } catch (error) {
      console.error('Failed to fetch queue stats:', error);
    }
  };

  const fetchJobStatistics = async () => {
    try {
      const response = await apiClient.getJobStatistics(30);
      if (response.success && response.data) {
        setJobStats(response.data);
      }
    } catch (error) {
      console.error('Failed to fetch job statistics:', error);
    }
  };

  const handleRefresh = () => {
    setIsLoading(true);
    fetchJobs();
    fetchQueueStats();
    fetchJobStatistics();
    setSelectedJobs([]);
  };

  const handleSort = (field: string, direction: 'asc' | 'desc') => {
    setSortField(field);
    setSortDirection(direction);
  };

  const handleJobClick = (job: TranslationJob) => {
    setSelectedJob(job);
    setIsJobModalOpen(true);
  };

  const handleRetryJob = async (jobId: string) => {
    setActionLoading(prev => ({ ...prev, [jobId]: true }));
    try {
      const response = await apiClient.retryJob(jobId);
      if (response.success) {
        await fetchJobs();
        await fetchQueueStats();
      }
    } catch (error) {
      console.error('Failed to retry job:', error);
    } finally {
      setActionLoading(prev => ({ ...prev, [jobId]: false }));
    }
  };

  const handleCancelJob = async (jobId: string) => {
    setActionLoading(prev => ({ ...prev, [jobId]: true }));
    try {
      const response = await apiClient.cancelJob(jobId);
      if (response.success) {
        await fetchJobs();
        await fetchQueueStats();
      }
    } catch (error) {
      console.error('Failed to cancel job:', error);
    } finally {
      setActionLoading(prev => ({ ...prev, [jobId]: false }));
    }
  };

  const handleDeleteJob = async (jobId: string) => {
    if (!confirm('Are you sure you want to delete this job? This action cannot be undone.')) {
      return;
    }

    setActionLoading(prev => ({ ...prev, [jobId]: true }));
    try {
      const response = await apiClient.deleteJob(jobId);
      if (response.success) {
        await fetchJobs();
        await fetchQueueStats();
        await fetchJobStatistics();
      }
    } catch (error) {
      console.error('Failed to delete job:', error);
    } finally {
      setActionLoading(prev => ({ ...prev, [jobId]: false }));
    }
  };

  const handleDownloadJob = (jobId: string) => {
    const job = jobs.find(j => j.id === jobId);
    if (job?.downloadUrl) {
      const link = document.createElement('a');
      link.href = job.downloadUrl;
      const extension = job.fileName.split('.').pop();
      link.download = job.fileName.replace(/\.[^/.]+$/, '_translated') + (extension ? `.${extension}` : '');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const handleBulkCancel = async () => {
    if (selectedJobs.length === 0) return;

    if (!confirm(`Are you sure you want to cancel ${selectedJobs.length} jobs?`)) {
      return;
    }

    try {
      const response = await apiClient.cancelJobsBulk(selectedJobs);
      if (response.success) {
        await fetchJobs();
        await fetchQueueStats();
        setSelectedJobs([]);
      }
    } catch (error) {
      console.error('Failed to cancel jobs:', error);
    }
  };

  const handleBulkRetry = async () => {
    if (selectedJobs.length === 0) return;

    const failedJobs = selectedJobs.filter(jobId => {
      const job = jobs.find(j => j.id === jobId);
      return job?.status === 'failed';
    });

    if (failedJobs.length === 0) {
      alert('No failed jobs selected for retry');
      return;
    }

    try {
      const response = await apiClient.retryJobsBulk(failedJobs);
      if (response.success) {
        await fetchJobs();
        await fetchQueueStats();
        setSelectedJobs([]);
      }
    } catch (error) {
      console.error('Failed to retry jobs:', error);
    }
  };

  const handleExportJobs = async (format: 'csv' | 'json') => {
    try {
      const response = await apiClient.exportJobs(format);
      if (response.success && response.data) {
        const blob = new Blob([response.data.data], {
          type: response.data.media_type
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = response.data.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (error) {
      console.error('Failed to export jobs:', error);
    }
  };

  if (!isAuthenticated) {
    return null;
  }

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Translation Jobs</h1>
          <p className="text-gray-600">Manage and monitor your translation jobs</p>
        </div>
        <div className="flex items-center space-x-4">
          <Button
            variant="outline"
            onClick={() => router.push('/dashboard/translate')}
          >
            <Plus className="mr-2 h-4 w-4" />
            New Translation
          </Button>
          <Button variant="outline" onClick={handleRefresh} disabled={isLoading}>
            <RefreshCw className={`mr-2 h-4 w-4 ${isLoading ? 'animate-spin' : ''}`} />
            Refresh
          </Button>
        </div>
      </div>

      <main className="max-w-7xl mx-auto space-y-6">
        {/* Statistics Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          {queueStats && (
            <>
              <Card>
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Total Jobs</CardTitle>
                  <FileText className="h-4 w-4 text-muted-foreground" />
                </CardHeader>
                <CardContent>
                  <div className="text-2xl font-bold">{queueStats.total_jobs}</div>
                  <p className="text-xs text-muted-foreground">
                    {queueStats.active_jobs} active
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Completed</CardTitle>
                  <CheckCircle className="h-4 w-4 text-green-600" />
                </CardHeader>
                <CardContent>
                  <div className="text-2xl font-bold text-green-600">{queueStats.completed}</div>
                  <p className="text-xs text-muted-foreground">
                    {jobStats?.total_jobs ? Math.round((queueStats.completed / jobStats.total_jobs) * 100) : 0}% success rate
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Failed</CardTitle>
                  <XCircle className="h-4 w-4 text-red-600" />
                </CardHeader>
                <CardContent>
                  <div className="text-2xl font-bold text-red-600">{queueStats.failed}</div>
                  <p className="text-xs text-muted-foreground">
                    Needs attention
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Total Cost</CardTitle>
                  <DollarSign className="h-4 w-4 text-muted-foreground" />
                </CardHeader>
                <CardContent>
                  <div className="text-2xl font-bold">
                    ${jobStats?.total_cost?.toFixed(2) || '0.00'}
                  </div>
                  <p className="text-xs text-muted-foreground">
                    Last 30 days
                  </p>
                </CardContent>
              </Card>
            </>
          )}
        </div>

        {/* Bulk Actions */}
        {selectedJobs.length > 0 && (
          <Card>
            <CardContent className="pt-6">
              <div className="flex items-center justify-between">
                <div>
                  <p className="text-sm font-medium">
                    {selectedJobs.length} job{selectedJobs.length !== 1 ? 's' : ''} selected
                  </p>
                </div>
                <div className="flex items-center space-x-2">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={handleBulkRetry}
                    disabled={selectedJobs.length === 0}
                  >
                    <Play className="mr-2 h-4 w-4" />
                    Retry Failed
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={handleBulkCancel}
                    disabled={selectedJobs.length === 0}
                  >
                    <X className="mr-2 h-4 w-4" />
                    Cancel Selected
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => handleExportJobs('csv')}
                  >
                    <Download className="mr-2 h-4 w-4" />
                    Export CSV
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => setSelectedJobs([])}
                  >
                    Clear Selection
                  </Button>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Jobs Table */}
        <JobQueueTable
          jobs={jobs}
          selectedJobs={selectedJobs}
          onSelectionChange={setSelectedJobs}
          onJobClick={handleJobClick}
          onRetryJob={handleRetryJob}
          onCancelJob={handleCancelJob}
          onDeleteJob={handleDeleteJob}
          onDownloadJob={handleDownloadJob}
          onSort={handleSort}
          sortField={sortField}
          sortDirection={sortDirection}
          isLoading={isLoading}
        />

        {/* Quick Stats */}
        {jobStats && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <BarChart3 className="h-5 w-5" />
                <span>30-Day Summary</span>
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <p className="text-sm text-gray-500">Average Duration</p>
                  <p className="text-lg font-semibold">
                    {jobStats.average_duration_minutes.toFixed(1)} minutes
                  </p>
                </div>
                <div>
                  <p className="text-sm text-gray-500">Success Rate</p>
                  <p className="text-lg font-semibold">
                    {jobStats.total_jobs > 0
                      ? Math.round((jobStats.status_counts.completed || 0) / jobStats.total_jobs * 100)
                      : 0}%
                  </p>
                </div>
                <div>
                  <p className="text-sm text-gray-500">Most Used Format</p>
                  <p className="text-lg font-semibold capitalize">
                    {Object.entries(jobStats.file_type_distribution || {}).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A'}
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        )}
      </main>

      {/* Job Detail Modal */}
      <JobDetailModal
        job={selectedJob}
        open={isJobModalOpen}
        onOpenChange={setIsJobModalOpen}
        onRetry={handleRetryJob}
        onCancel={handleCancelJob}
        onDelete={handleDeleteJob}
      />
    </div>
  );
}
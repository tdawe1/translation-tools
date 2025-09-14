import React, { useState, useEffect } from 'react';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { ScrollArea } from '@/components/ui/scroll-area';
import { TranslationJob } from '@/lib/api';
import { apiClient } from '@/lib/api';
import { QualityMetrics } from '@/components/jobs/QualityMetrics';
import {
  FileText,
  Calendar,
  Clock,
  DollarSign,
  CheckCircle,
  XCircle,
  AlertCircle,
  Download,
  Play,
  X,
  Trash2,
  RefreshCw,
  Server,
  HardDrive,
  Zap,
} from 'lucide-react';

interface JobDetailModalProps {
  job: TranslationJob | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onRetry: (jobId: string) => void;
  onCancel: (jobId: string) => void;
  onDelete: (jobId: string) => void;
}

const LOG_LEVEL_COLORS = {
  INFO: 'text-gray-600',
  WARNING: 'text-yellow-600',
  ERROR: 'text-red-600',
} as const;

const LOG_LEVEL_ICONS = {
  INFO: Server,
  WARNING: AlertCircle,
  ERROR: XCircle,
} as const;

export function JobDetailModal({
  job,
  open,
  onOpenChange,
  onRetry,
  onCancel,
  onDelete,
}: JobDetailModalProps) {
  const [logs, setLogs] = useState<any[]>([]);
  const [isLoadingLogs, setIsLoadingLogs] = useState(false);

  useEffect(() => {
    if (job && open) {
      fetchJobLogs();
    }
  }, [job, open]);

  const fetchJobLogs = async () => {
    if (!job) return;

    setIsLoadingLogs(true);
    try {
      const response = await apiClient.getJobLogs(job.id);
      if (response.success && response.data) {
        setLogs(response.data);
      }
    } catch (error) {
      console.error('Failed to fetch job logs:', error);
    } finally {
      setIsLoadingLogs(false);
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'completed':
        return <CheckCircle className="h-5 w-5 text-green-500" />;
      case 'failed':
        return <XCircle className="h-5 w-5 text-red-500" />;
      case 'processing':
        return <RefreshCw className="h-5 w-5 text-blue-500 animate-spin" />;
      case 'cancelled':
        return <XCircle className="h-5 w-5 text-gray-500" />;
      default:
        return <Clock className="h-5 w-5 text-gray-500" />;
    }
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const formatDate = (dateString?: string) => {
    if (!dateString) return '-';
    return new Date(dateString).toLocaleString();
  };

  const getDuration = () => {
    if (!job?.startedAt) return '-';
    const end = job?.completedAt ? new Date(job.completedAt) : new Date();
    const start = new Date(job.startedAt);
    const duration = Math.floor((end.getTime() - start.getTime()) / 1000);

    if (duration < 60) return `${duration}s`;
    const minutes = Math.floor(duration / 60);
    const seconds = duration % 60;
    return `${minutes}m ${seconds}s`;
  };

  if (!job) return null;

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-4xl max-h-[90vh] overflow-hidden">
        <DialogHeader>
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              {getStatusIcon(job.status)}
              <div>
                <DialogTitle className="text-xl">{job.fileName}</DialogTitle>
                <DialogDescription>
                  Job ID: {job.id}
                </DialogDescription>
              </div>
            </div>
            <Badge variant={job.status === 'completed' ? 'default' : job.status === 'failed' ? 'destructive' : 'secondary'}>
              {job.status}
            </Badge>
          </div>
        </DialogHeader>

        <Tabs defaultValue="overview" className="h-full">
          <TabsList className="grid w-full grid-cols-4">
            <TabsTrigger value="overview">Overview</TabsTrigger>
            <TabsTrigger value="progress">Progress</TabsTrigger>
            {job.status === 'completed' && job.qualityMetrics && (
              <TabsTrigger value="quality">Quality</TabsTrigger>
            )}
            <TabsTrigger value="logs">Logs</TabsTrigger>
          </TabsList>

          <TabsContent value="overview" className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {/* Basic Info */}
              <Card>
                <CardHeader>
                  <CardTitle className="text-lg">Job Information</CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  <div className="flex items-center space-x-2">
                    <FileText className="h-4 w-4 text-gray-500" />
                    <div>
                      <p className="text-sm font-medium text-gray-500">File Type</p>
                      <p className="text-sm">{job.fileType.toUpperCase()}</p>
                    </div>
                  </div>
                  <div className="flex items-center space-x-2">
                    <Calendar className="h-4 w-4 text-gray-500" />
                    <div>
                      <p className="text-sm font-medium text-gray-500">Created</p>
                      <p className="text-sm">{formatDate(job.createdAt)}</p>
                    </div>
                  </div>
                  <div className="flex items-center space-x-2">
                    <Clock className="h-4 w-4 text-gray-500" />
                    <div>
                      <p className="text-sm font-medium text-gray-500">Duration</p>
                      <p className="text-sm">{getDuration()}</p>
                    </div>
                  </div>
                  {job.metadata?.file_size && (
                    <div className="flex items-center space-x-2">
                      <HardDrive className="h-4 w-4 text-gray-500" />
                      <div>
                        <p className="text-sm font-medium text-gray-500">File Size</p>
                        <p className="text-sm">{formatFileSize(job.metadata.file_size)}</p>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>

              {/* Translation Settings */}
              <Card>
                <CardHeader>
                  <CardTitle className="text-lg">Translation Settings</CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  <div>
                    <p className="text-sm font-medium text-gray-500">Model</p>
                    <p className="text-sm">{job.metadata?.model || 'gpt-4o-2024-08-06'}</p>
                  </div>
                  <div>
                    <p className="text-sm font-medium text-gray-500">Temperature</p>
                    <p className="text-sm">{job.metadata?.temperature || 0.6}</p>
                  </div>
                  {job.metadata?.pages && (
                    <div>
                      <p className="text-sm font-medium text-gray-500">Pages</p>
                      <p className="text-sm">{job.metadata.pages}</p>
                    </div>
                  )}
                  <div>
                    <p className="text-sm font-medium text-gray-500">Offline Mode</p>
                    <p className="text-sm">{job.metadata?.offline ? 'Yes' : 'No'}</p>
                  </div>
                </CardContent>
              </Card>

              {/* Cost Information */}
              <Card>
                <CardHeader>
                  <CardTitle className="text-lg flex items-center space-x-2">
                    <DollarSign className="h-5 w-5" />
                    <span>Cost Information</span>
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  {job.estimatedCost !== undefined && (
                    <div>
                      <p className="text-sm font-medium text-gray-500">Estimated Cost</p>
                      <p className="text-lg font-semibold">${job.estimatedCost.toFixed(4)}</p>
                    </div>
                  )}
                  {job.actualCost !== undefined && (
                    <div>
                      <p className="text-sm font-medium text-gray-500">Actual Cost</p>
                      <p className="text-lg font-semibold">${job.actualCost.toFixed(4)}</p>
                    </div>
                  )}
                </CardContent>
              </Card>

              {/* Output Information */}
              {job.status === 'completed' && job.metadata?.output_file_size && (
                <Card>
                  <CardHeader>
                    <CardTitle className="text-lg">Output Information</CardTitle>
                  </CardHeader>
                  <CardContent className="space-y-3">
                    <div className="flex items-center space-x-2">
                      <HardDrive className="h-4 w-4 text-gray-500" />
                      <div>
                        <p className="text-sm font-medium text-gray-500">Output Size</p>
                        <p className="text-sm">{formatFileSize(job.metadata.output_file_size)}</p>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              )}
            </div>

            {/* Error Message */}
            {job.errorMessage && (
              <Alert variant="destructive">
                <AlertCircle className="h-4 w-4" />
                <AlertDescription>{job.errorMessage}</AlertDescription>
              </Alert>
            )}

            {/* Actions */}
            <div className="flex justify-end space-x-2 pt-4">
              {job.downloadUrl && (
                <Button variant="outline" asChild>
                  <a href={job.downloadUrl} download>
                    <Download className="mr-2 h-4 w-4" />
                    Download
                  </a>
                </Button>
              )}
              {job.status === 'failed' && (
                <Button variant="outline" onClick={() => onRetry(job.id)}>
                  <Play className="mr-2 h-4 w-4" />
                  Retry Job
                </Button>
              )}
              {(job.status === 'pending' || job.status === 'processing') && (
                <Button variant="destructive" onClick={() => onCancel(job.id)}>
                  <X className="mr-2 h-4 w-4" />
                  Cancel Job
                </Button>
              )}
              {(job.status === 'completed' || job.status === 'failed' || job.status === 'cancelled') && (
                <Button
                  variant="outline"
                  onClick={() => onDelete(job.id)}
                  className="text-red-600 hover:text-red-700"
                >
                  <Trash2 className="mr-2 h-4 w-4" />
                  Delete
                </Button>
              )}
            </div>
          </TabsContent>

          <TabsContent value="progress" className="space-y-4">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center space-x-2">
                  <Zap className="h-5 w-5" />
                  <span>Job Progress</span>
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div>
                  <div className="flex justify-between mb-2">
                    <span className="text-sm font-medium">Overall Progress</span>
                    <span className="text-sm text-gray-500">{job.progress}%</span>
                  </div>
                  <Progress value={job.progress} className="h-3" />
                </div>

                {/* Progress Stages */}
                <div className="space-y-2">
                  {[
                    { stage: 'Queued', progress: 0, completed: job.status !== 'pending' || job.progress > 0 },
                    { stage: 'Initializing', progress: 5, completed: job.progress > 5 },
                    { stage: 'Processing', progress: 15, completed: job.progress > 15 },
                    { stage: 'Finalizing', progress: 95, completed: job.progress >= 95 },
                    { stage: 'Complete', progress: 100, completed: job.progress === 100 },
                  ].map((stage, index) => (
                    <div key={stage.stage} className="flex items-center space-x-3">
                      <div className={`w-8 h-8 rounded-full flex items-center justify-center ${
                        stage.completed
                          ? 'bg-green-500 text-white'
                          : job.progress >= stage.progress
                          ? 'bg-blue-500 text-white'
                          : 'bg-gray-200 text-gray-500'
                      }`}>
                        {stage.completed ? (
                          <CheckCircle className="h-4 w-4" />
                        ) : (
                          <span className="text-xs">{index + 1}</span>
                        )}
                      </div>
                      <span className={`text-sm ${
                        stage.completed ? 'text-green-600' : job.progress >= stage.progress ? 'text-blue-600' : 'text-gray-500'
                      }`}>
                        {stage.stage}
                      </span>
                    </div>
                  ))}
                </div>

                {/* Timeline */}
                <div className="mt-6 space-y-2">
                  <h4 className="text-sm font-medium">Timeline</h4>
                  <div className="space-y-1 text-sm text-gray-600">
                    <div>Created: {formatDate(job.createdAt)}</div>
                    {job.startedAt && <div>Started: {formatDate(job.startedAt)}</div>}
                    {job.completedAt && <div>Completed: {formatDate(job.completedAt)}</div>}
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {job.status === 'completed' && job.qualityMetrics && (
            <TabsContent value="quality" className="space-y-4">
              <QualityMetrics metrics={job.qualityMetrics} />
            </TabsContent>
          )}

          <TabsContent value="logs" className="space-y-4">
            <Card>
              <CardHeader>
                <div className="flex items-center justify-between">
                  <CardTitle>Job Logs</CardTitle>
                  <Button variant="outline" size="sm" onClick={fetchJobLogs} disabled={isLoadingLogs}>
                    <RefreshCw className={`h-4 w-4 mr-2 ${isLoadingLogs ? 'animate-spin' : ''}`} />
                    Refresh
                  </Button>
                </div>
              </CardHeader>
              <CardContent>
                <ScrollArea className="h-[400px] border rounded-md p-4">
                  {isLoadingLogs ? (
                    <div className="flex items-center justify-center h-full">
                      <RefreshCw className="h-6 w-6 animate-spin text-gray-400" />
                    </div>
                  ) : logs.length === 0 ? (
                    <p className="text-center text-gray-500">No logs available</p>
                  ) : (
                    <div className="space-y-2">
                      {logs.map((log, index) => {
                        const Icon = LOG_LEVEL_ICONS[log.level as keyof typeof LOG_LEVEL_ICONS] || Server;
                        const colorClass = LOG_LEVEL_COLORS[log.level as keyof typeof LOG_LEVEL_COLORS] || 'text-gray-600';

                        return (
                          <div key={index} className="flex items-start space-x-3 text-sm">
                            <Icon className={`h-4 w-4 mt-0.5 ${colorClass}`} />
                            <div className="flex-1">
                              <div className="flex items-center space-x-2">
                                <span className="font-medium text-gray-700">{log.level}</span>
                                <span className="text-gray-500">
                                  {new Date(log.timestamp).toLocaleTimeString()}
                                </span>
                              </div>
                              <p className="text-gray-600">{log.message}</p>
                              {log.data && (
                                <details className="mt-1">
                                  <summary className="text-xs text-gray-500 cursor-pointer hover:text-gray-700">
                                    Show details
                                  </summary>
                                  <pre className="mt-1 text-xs bg-gray-100 p-2 rounded overflow-x-auto">
                                    {JSON.stringify(log.data, null, 2)}
                                  </pre>
                                </details>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </ScrollArea>
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
      </DialogContent>
    </Dialog>
  );
}
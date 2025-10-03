'use client';

import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Alert, AlertDescription } from '@/components/ui/alert';
import {
  Download,
  Clock,
  DollarSign,
  FileText,
  Wifi,
  WifiOff,
  Loader2,
  CheckCircle,
  AlertCircle,
  Activity,
  BarChart3
} from 'lucide-react';
import { translationWebSocket, JobProgress } from '@/lib/websocket';
import { formatFileSize, formatDuration } from '@/lib/format';

interface RealTimeProgressProps {
  jobId: string;
  fileName: string;
  onDownload?: (url: string) => void;
  className?: string;
}

export function RealTimeProgress({
  jobId,
  fileName,
  onDownload,
  className = ''
}: RealTimeProgressProps) {
  const [progress, setProgress] = useState<JobProgress | null>(null);
  const [isConnected, setIsConnected] = useState(false);
  const [connectionStatus, setConnectionStatus] = useState<'connecting' | 'connected' | 'disconnected'>('disconnected');
  const [events, setEvents] = useState<{ time: string; message: string }[]>([]);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  // Add event to activity feed
  const addEvent = (message: string) => {
    const time = new Date().toLocaleTimeString();
    setEvents(prev => [{ time, message }, ...prev.slice(0, 9)]); // Keep last 10 events
  };

  // Setup WebSocket connection
  useEffect(() => {
    // Connection handlers
    const handleConnection = (connected: boolean) => {
      setIsConnected(connected);
      setConnectionStatus(connected ? 'connected' : 'disconnected');
      addEvent(connected ? 'Connected to server' : 'Disconnected from server');
    };

    // Progress handlers
    const handleJobStarted = (update: any) => {
      setProgress(update);
      addEvent(`Translation started for ${update.file_name}`);
    };

    const handleJobProgress = (update: any) => {
      setProgress(update);
    };

    const handleJobCompleted = (update: any) => {
      if (progress) {
        setProgress({ ...progress, status: 'completed', progress: 100 });
      }
      addEvent('Translation completed successfully!');
      if (update.downloadUrl) {
        setDownloadUrl(update.downloadUrl);
      }
    };

    const handleJobFailed = (update: any) => {
      if (progress) {
        setProgress({ ...progress, status: 'failed', error_message: update.error_message });
      }
      addEvent(`Translation failed: ${update.error_message}`);
    };

    // Register handlers
    translationWebSocket.onConnection(handleConnection);
    translationWebSocket.on('job_started', handleJobStarted);
    translationWebSocket.on('job_progress', handleJobProgress);
    translationWebSocket.on('job_completed', handleJobCompleted);
    translationWebSocket.on('job_failed', handleJobFailed);

    // Connect
    setConnectionStatus('connecting');
    translationWebSocket.connect();
    translationWebSocket.subscribe(jobId);

    // Cleanup
    return () => {
      translationWebSocket.offConnection(handleConnection);
      translationWebSocket.off('job_started', handleJobStarted);
      translationWebSocket.off('job_progress', handleJobProgress);
      translationWebSocket.off('job_completed', handleJobCompleted);
      translationWebSocket.off('job_failed', handleJobFailed);
      translationWebSocket.unsubscribe(jobId);
    };
  }, [jobId]);

  // Calculate stage progress
  const getStageProgress = () => {
    if (!progress) return 0;

    const stages = {
      extracting: 20,
      translating: 60,
      applying: 15,
      finalizing: 5
    };

    if (progress.status === 'completed') return 100;
    if (progress.status === 'failed') return 0;

    const currentStage = progress.stage as keyof typeof stages;
    return stages[currentStage] || 0;
  };

  // Get status badge
  const getStatusBadge = () => {
    if (!progress) return null;

    const statusConfig = {
      queued: { variant: 'secondary' as const, label: 'Queued' },
      extracting: { variant: 'default' as const, label: 'Extracting Text' },
      translating: { variant: 'default' as const, label: 'Translating' },
      applying: { variant: 'default' as const, label: 'Applying' },
      finalizing: { variant: 'default' as const, label: 'Finalizing' },
      completed: { variant: 'default' as const, label: 'Completed' },
      failed: { variant: 'destructive' as const, label: 'Failed' }
    };

    const config = statusConfig[progress.status as keyof typeof statusConfig];
    return (
      <Badge variant={config.variant} className="capitalize">
        {config.label}
      </Badge>
    );
  };

  return (
    <div className={`space-y-4 ${className}`}>
      {/* Connection Status */}
      <div className="flex items-center justify-between">
        <div className="flex items-center space-x-2">
          {isConnected ? (
            <Wifi className="h-4 w-4 text-green-500" />
          ) : (
            <WifiOff className="h-4 w-4 text-red-500" />
          )}
          <span className="text-sm text-gray-600">
            {connectionStatus === 'connecting' ? 'Connecting...' :
             isConnected ? 'Connected' : 'Disconnected'}
          </span>
        </div>
        <div className="flex items-center space-x-2">
          {progress?.current_batch && progress?.total_batches && (
            <span className="text-sm text-gray-600">
              Batch {progress.current_batch} of {progress.total_batches}
            </span>
          )}
        </div>
      </div>

      {/* Progress Overview */}
      <Card>
        <CardHeader>
          <div className="flex items-center justify-between">
            <div>
              <CardTitle className="flex items-center space-x-2">
                <FileText className="h-5 w-5" />
                <span>{fileName}</span>
              </CardTitle>
              <CardDescription>
                Real-time translation progress
              </CardDescription>
            </div>
            {getStatusBadge()}
          </div>
        </CardHeader>
        <CardContent className="space-y-6">
          {/* Main Progress Bar */}
          <div className="space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-sm font-medium">Overall Progress</span>
              <span className="text-sm text-gray-500">
                {progress?.progress ? Math.round(progress.progress) : 0}%
              </span>
            </div>
            <Progress value={progress?.progress || 0} />
          </div>

          {/* Stage Details */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* Stage Progress */}
            <Card>
              <CardContent className="p-4">
                <div className="flex items-center space-x-2">
                  <Activity className="h-4 w-4 text-blue-500" />
                  <span className="text-sm font-medium">Current Stage</span>
                </div>
                <p className="text-xs text-gray-500 mt-1">
                  {progress?.stage || 'Initializing'} - {getStageProgress()}%
                </p>
              </CardContent>
            </Card>

            {/* Cost Tracking */}
            <Card>
              <CardContent className="p-4">
                <div className="flex items-center space-x-2">
                  <DollarSign className="h-4 w-4 text-green-500" />
                  <span className="text-sm font-medium">Cost</span>
                </div>
                <p className="text-xs text-gray-500 mt-1">
                  ${progress?.current_cost?.toFixed(4) || '0.0000'} / ${progress?.estimated_cost?.toFixed(4) || '0.0000'}
                </p>
              </CardContent>
            </Card>

            {/* ETA */}
            {progress?.eta_seconds && (
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center space-x-2">
                    <Clock className="h-4 w-4 text-purple-500" />
                    <span className="text-sm font-medium">ETA</span>
                  </div>
                  <p className="text-xs text-gray-500 mt-1">
                    {formatDuration(progress.eta_seconds)}
                  </p>
                </CardContent>
              </Card>
            )}
          </div>

          {/* Token Progress */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <Card>
              <CardContent className="p-4">
                <div className="flex items-center justify-between">
                  <span className="text-sm font-medium">Tokens Processed</span>
                  <span className="text-sm text-gray-500">
                    {progress?.tokens_processed?.toLocaleString() || 0} / {progress?.total_tokens?.toLocaleString() || 0}
                  </span>
                </div>
                <Progress
                  value={progress?.total_tokens ? (progress.tokens_processed / progress.total_tokens) * 100 : 0}
                  className="mt-2"
                />
              </CardContent>
            </Card>

            {/* Quality Score */}
            {progress?.quality_score !== undefined && (
              <Card>
                <CardContent className="p-4">
                  <div className="flex items-center space-x-2">
                    <BarChart3 className="h-4 w-4 text-indigo-500" />
                    <span className="text-sm font-medium">Quality Score</span>
                  </div>
                  <div className="flex items-center space-x-2 mt-2">
                    <Progress value={progress.quality_score * 100} className="flex-1" />
                    <span className="text-sm font-medium">
                      {Math.round(progress.quality_score * 100)}%
                    </span>
                  </div>
                </CardContent>
              </Card>
            )}
          </div>

          {/* Error Message */}
          {progress?.error_message && (
            <Alert variant="destructive">
              <AlertCircle className="h-4 w-4" />
              <AlertDescription>{progress.error_message}</AlertDescription>
            </Alert>
          )}

          {/* Download Button */}
          {progress?.status === 'completed' && downloadUrl && (
            <div className="flex justify-center">
              <Button onClick={() => onDownload?.(downloadUrl)}>
                <Download className="mr-2 h-4 w-4" />
                Download Translated File
              </Button>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Activity Feed */}
      <Card>
        <CardHeader>
          <CardTitle className="text-base">Activity Feed</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="space-y-2 max-h-40 overflow-y-auto">
            {events.length === 0 ? (
              <p className="text-sm text-gray-500 text-center py-4">No activity yet</p>
            ) : (
              events.map((event, index) => (
                <div key={index} className="flex items-start space-x-2 text-sm">
                  <span className="text-gray-500 font-mono text-xs">{event.time}</span>
                  <span className="text-gray-700">{event.message}</span>
                </div>
              ))
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
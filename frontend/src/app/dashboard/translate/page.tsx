'use client';

import { useState, useRef, useCallback } from 'react';
import { useRouter } from 'next/navigation';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Textarea } from '@/components/ui/textarea';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { useAuth } from '@/contexts/AuthContext';
import { apiClient } from '@/lib/api';
import { RealTimeProgress } from '@/components/ui/RealTimeProgress';
import { ConnectionStatus } from '@/components/ui/ConnectionStatus';
import {
  Upload,
  FileText,
  X,
  CheckCircle,
  AlertCircle,
  Download,
  Eye,
  Clock,
  Loader2,
} from 'lucide-react';

interface FileWithPreview {
  file: File;
  preview: string;
  id: string;
}

export default function TranslatePage() {
  const { isAuthenticated } = useAuth();
  const router = useRouter();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<FileWithPreview[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [currentJob, setCurrentJob] = useState<any>(null);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [estimate, setEstimate] = useState<{ cost: number; tokens: number } | null>(null);
  const [realTimeJobId, setRealTimeJobId] = useState<string | null>(null);

  // Translation options
  const [model, setModel] = useState('gpt-4o-2024-08-06');
  const [tone, setTone] = useState('professional');
  const [style, setStyle] = useState('default');
  const [pages, setPages] = useState('');
  const [offline, setOffline] = useState(false);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []);
    const validFiles = selectedFiles.filter(file =>
      file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
      file.type === 'application/pdf'
    );

    if (validFiles.length !== selectedFiles.length) {
      setError('Only PowerPoint (.pptx) and PDF files are supported');
    }

    const newFiles = validFiles.map(file => ({
      file,
      preview: file.name,
      id: Math.random().toString(36).substring(7),
    }));

    setFiles(prev => [...prev, ...newFiles]);
  }, []);

  const removeFile = (id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  };

  const handleEstimate = async (file: File) => {
    try {
      const response = await apiClient.estimateCost(file, { model, pages });
      if (response.success && response.data) {
        setEstimate(response.data);
      }
    } catch (error) {
      console.error('Failed to estimate cost:', error);
    }
  };

  const handleUpload = async () => {
    if (files.length === 0) {
      setError('Please select at least one file');
      return;
    }

    setIsUploading(true);
    setError('');
    setSuccess('');

    for (const fileObj of files) {
      const formData = new FormData();
      formData.append('file', fileObj.file);
      formData.append('model', model);
      formData.append('tone', tone);
      formData.append('style', style);
      if (pages) formData.append('pages', pages);
      if (offline) formData.append('offline', offline.toString());

      try {
        const response = await apiClient.uploadFile(formData);

        if (response.success && response.data) {
          setCurrentJob(response.data);
          setRealTimeJobId(response.data.jobId);
          setSuccess(`File "${fileObj.file.name}" uploaded successfully!`);

          // The real-time progress will be handled by RealTimeProgress component
          // Keep minimal polling as fallback
          const pollStatus = setInterval(async () => {
            const statusResponse = await apiClient.getJobStatus(response.data.jobId);
            if (statusResponse.success && statusResponse.data) {
              setCurrentJob(statusResponse.data);

              if (statusResponse.data.status === 'completed' || statusResponse.data.status === 'failed') {
                clearInterval(pollStatus);
              }
            }
          }, 5000); // Less frequent polling

          // Stop polling after 10 minutes
          setTimeout(() => clearInterval(pollStatus), 600000);
        } else {
          setError(response.error || 'Upload failed');
        }
      } catch (error: any) {
        setError(error.response?.data?.message || error.message);
      }
    }

    setIsUploading(false);
    setFiles([]);
    setUploadProgress(0);
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const getFileType = (fileName: string) => {
    return fileName.toLowerCase().endsWith('.pptx') ? 'PowerPoint' : 'PDF';
  };

  if (!isAuthenticated) {
    router.push('/login');
    return null;
  }

  return (
    <div>
      <div className="mb-8">
        <h1 className="text-2xl font-bold text-gray-900">New Translation</h1>
        <p className="text-gray-600">Upload and translate your documents</p>
      </div>

      <main className="max-w-6xl mx-auto">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Left Column - Upload */}
          <div className="lg:col-span-2 space-y-6">
            {/* File Upload */}
            <Card>
              <CardHeader>
                <CardTitle>Upload Files</CardTitle>
                <CardDescription>
                  Select PowerPoint (.pptx) or PDF files for translation
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div
                  className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center cursor-pointer hover:border-gray-400 transition-colors"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <Upload className="mx-auto h-12 w-12 text-gray-400 mb-4" />
                  <p className="text-lg text-gray-600 mb-2">Drop files here or click to upload</p>
                  <p className="text-sm text-gray-500">
                    Supports .pptx and .pdf files up to 100MB
                  </p>
                  <input
                    ref={fileInputRef}
                    type="file"
                    multiple
                    accept=".pptx,.pdf"
                    onChange={handleFileSelect}
                    className="hidden"
                  />
                </div>

                {files.length > 0 && (
                  <div className="mt-4 space-y-2">
                    {files.map((fileObj) => (
                      <div
                        key={fileObj.id}
                        className="flex items-center justify-between p-3 bg-gray-50 rounded-lg"
                      >
                        <div className="flex items-center space-x-3">
                          <FileText className="h-8 w-8 text-blue-500" />
                          <div>
                            <p className="font-medium">{fileObj.file.name}</p>
                            <p className="text-sm text-gray-500">
                              {getFileType(fileObj.file.name)} • {formatFileSize(fileObj.file.size)}
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center space-x-2">
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => handleEstimate(fileObj.file)}
                          >
                            Estimate Cost
                          </Button>
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => removeFile(fileObj.id)}
                          >
                            <X className="h-4 w-4" />
                          </Button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                {estimate && (
                  <div className="mt-4 p-4 bg-blue-50 rounded-lg">
                    <p className="text-sm font-medium text-blue-900">
                      Estimated Cost: ${estimate.cost.toFixed(4)}
                    </p>
                    <p className="text-sm text-blue-700">
                      Tokens: {estimate.tokens.toLocaleString()}
                    </p>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Translation Options */}
            <Card>
              <CardHeader>
                <CardTitle>Translation Options</CardTitle>
                <CardDescription>
                  Configure translation settings and preferences
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label htmlFor="model">AI Model</Label>
                    <Select value={model} onValueChange={setModel}>
                      <SelectTrigger>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="gpt-4o-2024-08-06">GPT-4o (Recommended)</SelectItem>
                        <SelectItem value="gpt-4o-mini">GPT-4o Mini (Cost-Optimized)</SelectItem>
                        <SelectItem value="gpt-5">GPT-5 (Premium)</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="tone">Tone</Label>
                    <Select value={tone} onValueChange={setTone}>
                      <SelectTrigger>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="professional">Professional</SelectItem>
                        <SelectItem value="casual">Casual</SelectItem>
                        <SelectItem value="formal">Formal</SelectItem>
                        <SelectItem value="technical">Technical</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="style">Style</Label>
                    <Select value={style} onValueChange={setStyle}>
                      <SelectTrigger>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="default">Default</SelectItem>
                        <SelectItem value="minimal">Minimal</SelectItem>
                        <SelectItem value="detailed">Detailed</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="pages">Pages (PDF only)</Label>
                    <Input
                      id="pages"
                      placeholder="e.g., 1-10, 15, 20-25"
                      value={pages}
                      onChange={(e) => setPages(e.target.value)}
                    />
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  <input
                    type="checkbox"
                    id="offline"
                    checked={offline}
                    onChange={(e) => setOffline(e.target.checked)}
                    className="rounded border-gray-300"
                  />
                  <Label htmlFor="offline">Use offline translation (cache only)</Label>
                </div>
              </CardContent>
            </Card>

            {/* Upload Button */}
            <div className="flex justify-end">
              <Button
                onClick={handleUpload}
                disabled={files.length === 0 || isUploading}
                className="px-8"
              >
                {isUploading ? (
                  <>
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    Uploading...
                  </>
                ) : (
                  'Start Translation'
                )}
              </Button>
            </div>

            {/* Job Status */}
            {realTimeJobId && currentJob && (
              <RealTimeProgress
                jobId={realTimeJobId}
                fileName={currentJob.fileName || currentJob.file_name || files[0]?.file.name || 'Unknown'}
                onDownload={(url) => {
                  const link = document.createElement('a');
                  link.href = url;
                  link.download = url.split('/').pop() || 'translated_file';
                  document.body.appendChild(link);
                  link.click();
                  document.body.removeChild(link);
                }}
              />
            )}

            {/* Fallback Status for non-real-time updates */}
            {currentJob && !realTimeJobId && (
              <Card>
                <CardHeader>
                  <CardTitle>Translation Status</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <span className="text-sm font-medium">Progress</span>
                      <span className="text-sm text-gray-500">{currentJob.progress}%</span>
                    </div>
                    <Progress value={currentJob.progress} />
                    <div className="flex items-center space-x-2">
                      {currentJob.status === 'processing' && (
                        <Loader2 className="h-4 w-4 animate-spin" />
                      )}
                      <Badge
                        variant={
                          currentJob.status === 'completed'
                            ? 'default'
                            : currentJob.status === 'failed'
                            ? 'destructive'
                            : 'secondary'
                        }
                        className="capitalize"
                      >
                        {currentJob.status}
                      </Badge>
                    </div>
                    {currentJob.errorMessage && (
                      <Alert variant="destructive">
                        <AlertCircle className="h-4 w-4" />
                        <AlertDescription>{currentJob.errorMessage}</AlertDescription>
                      </Alert>
                    )}
                    {currentJob.downloadUrl && (
                      <Button variant="outline" asChild>
                        <a href={currentJob.downloadUrl} download>
                          <Download className="mr-2 h-4 w-4" />
                          Download Translated File
                        </a>
                      </Button>
                    )}
                  </div>
                </CardContent>
              </Card>
            )}
          </div>

          {/* Right Column - Instructions */}
          <div>
            {/* Connection Status */}
            <Card className="mb-6">
              <CardHeader>
                <CardTitle className="text-base">Connection Status</CardTitle>
              </CardHeader>
              <CardContent>
                <ConnectionStatus />
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle>How It Works</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-3">
                  <div className="flex items-start space-x-3">
                    <div className="flex-shrink-0 w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center">
                      <span className="text-xs font-medium text-blue-600">1</span>
                    </div>
                    <div>
                      <p className="text-sm font-medium">Upload Your File</p>
                      <p className="text-xs text-gray-500">Select .pptx or .pdf files</p>
                    </div>
                  </div>

                  <div className="flex items-start space-x-3">
                    <div className="flex-shrink-0 w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center">
                      <span className="text-xs font-medium text-blue-600">2</span>
                    </div>
                    <div>
                      <p className="text-sm font-medium">Configure Options</p>
                      <p className="text-xs text-gray-500">Choose model, tone, and style</p>
                    </div>
                  </div>

                  <div className="flex items-start space-x-3">
                    <div className="flex-shrink-0 w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center">
                      <span className="text-xs font-medium text-blue-600">3</span>
                    </div>
                    <div>
                      <p className="text-sm font-medium">Translation Process</p>
                      <p className="text-xs text-gray-500">
                        AI translates while preserving layout
                      </p>
                    </div>
                  </div>

                  <div className="flex items-start space-x-3">
                    <div className="flex-shrink-0 w-6 h-6 bg-blue-100 rounded-full flex items-center justify-center">
                      <span className="text-xs font-medium text-blue-600">4</span>
                    </div>
                    <div>
                      <p className="text-sm font-medium">Download Result</p>
                      <p className="text-xs text-gray-500">Get your translated file</p>
                    </div>
                  </div>
                </div>

                <div className="pt-4 border-t">
                  <p className="text-sm font-medium mb-2">Supported Formats</p>
                  <div className="space-y-1">
                    <p className="text-xs text-gray-600">
                      <strong>PowerPoint:</strong> .pptx files
                    </p>
                    <p className="text-xs text-gray-600">
                      <strong>PDF:</strong> .pdf files with selectable text
                    </p>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>

        {/* Error/Success Messages */}
        {error && (
          <Alert variant="destructive" className="mt-6">
            <AlertCircle className="h-4 w-4" />
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}

        {success && (
          <Alert className="mt-6">
            <CheckCircle className="h-4 w-4" />
            <AlertDescription>{success}</AlertDescription>
          </Alert>
        )}
      </main>
    </div>
  );
}
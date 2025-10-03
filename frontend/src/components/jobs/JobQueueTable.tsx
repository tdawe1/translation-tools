import React, { useState, useCallback, useMemo } from 'react';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Checkbox } from '@/components/ui/checkbox';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { Input } from '@/components/ui/input';
import { TranslationJob } from '@/lib/api';
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
  MoreVertical,
  ChevronUp,
  ChevronDown,
} from 'lucide-react';

interface JobQueueTableProps {
  jobs: TranslationJob[];
  selectedJobs: string[];
  onSelectionChange: (jobIds: string[]) => void;
  onJobClick: (job: TranslationJob) => void;
  onRetryJob: (jobId: string) => void;
  onCancelJob: (jobId: string) => void;
  onDeleteJob: (jobId: string) => void;
  onDownloadJob: (jobId: string) => void;
  onSort: (field: string, direction: 'asc' | 'desc') => void;
  sortField?: string;
  sortDirection?: 'asc' | 'desc';
  isLoading?: boolean;
}

const STATUS_COLORS = {
  pending: 'secondary',
  processing: 'default',
  completed: 'default',
  failed: 'destructive',
  cancelled: 'secondary',
} as const;

const STATUS_ICONS = {
  pending: Clock,
  processing: RefreshCw,
  completed: CheckCircle,
  failed: XCircle,
  cancelled: XCircle,
} as const;

export function JobQueueTable({
  jobs,
  selectedJobs,
  onSelectionChange,
  onJobClick,
  onRetryJob,
  onCancelJob,
  onDeleteJob,
  onDownloadJob,
  onSort,
  sortField,
  sortDirection,
  isLoading,
}: JobQueueTableProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [typeFilter, setTypeFilter] = useState<string>('all');

  // Filter and sort jobs
  const filteredJobs = useMemo(() => {
    return jobs.filter((job) => {
      const matchesSearch = job.fileName.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesStatus = statusFilter === 'all' || job.status === statusFilter;
      const matchesType = typeFilter === 'all' || job.fileType === typeFilter;
      return matchesSearch && matchesStatus && matchesType;
    });
  }, [jobs, searchTerm, statusFilter, typeFilter]);

  const handleSelectAll = useCallback((checked: boolean) => {
    if (checked) {
      onSelectionChange(filteredJobs.map((job) => job.id));
    } else {
      onSelectionChange([]);
    }
  }, [filteredJobs, onSelectionChange]);

  const handleSelectJob = useCallback((jobId: string, checked: boolean) => {
    if (checked) {
      onSelectionChange([...selectedJobs, jobId]);
    } else {
      onSelectionChange(selectedJobs.filter((id) => id !== jobId));
    }
  }, [selectedJobs, onSelectionChange]);

  const handleSort = useCallback((field: string) => {
    const direction = sortField === field && sortDirection === 'asc' ? 'desc' : 'asc';
    onSort(field, direction);
  }, [sortField, sortDirection, onSort]);

  const SortIcon = ({ field }: { field: string }) => {
    if (sortField !== field) return null;
    return sortDirection === 'asc' ? (
      <ChevronUp className="h-4 w-4" />
    ) : (
      <ChevronDown className="h-4 w-4" />
    );
  };

  const getStatusIcon = (status: string) => {
    const Icon = STATUS_ICONS[status as keyof typeof STATUS_ICONS] || Clock;
    const colorClass = {
      pending: 'text-gray-500',
      processing: 'text-blue-500',
      completed: 'text-green-500',
      failed: 'text-red-500',
      cancelled: 'text-gray-500',
    }[status];

    return <Icon className={`h-4 w-4 ${colorClass} ${status === 'processing' ? 'animate-spin' : ''}`} />;
  };

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleString();
  };

  const getDuration = (job: TranslationJob) => {
    if (!job.startedAt) return '-';
    const end = job.completedAt ? new Date(job.completedAt) : new Date();
    const start = new Date(job.startedAt);
    const duration = Math.floor((end.getTime() - start.getTime()) / 60000);
    return duration > 0 ? `${duration}m` : '< 1m';
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center py-8">
        <RefreshCw className="h-8 w-8 animate-spin text-gray-400" />
        <span className="ml-2 text-gray-600">Loading jobs...</span>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Filters */}
      <div className="flex flex-col sm:flex-row gap-4">
        <div className="flex-1">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-4 w-4" />
            <Input
              placeholder="Search jobs..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="pl-10"
            />
          </div>
        </div>
        <Select value={statusFilter} onValueChange={setStatusFilter}>
          <SelectTrigger className="w-[180px]">
            <SelectValue placeholder="Filter by status" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">All Statuses</SelectItem>
            <SelectItem value="pending">Pending</SelectItem>
            <SelectItem value="processing">Processing</SelectItem>
            <SelectItem value="completed">Completed</SelectItem>
            <SelectItem value="failed">Failed</SelectItem>
            <SelectItem value="cancelled">Cancelled</SelectItem>
          </SelectContent>
        </Select>
        <Select value={typeFilter} onValueChange={setTypeFilter}>
          <SelectTrigger className="w-[180px]">
            <SelectValue placeholder="Filter by type" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="all">All Types</SelectItem>
            <SelectItem value="pptx">PowerPoint</SelectItem>
            <SelectItem value="pdf">PDF</SelectItem>
          </SelectContent>
        </Select>
      </div>

      {/* Jobs Table */}
      <div className="border rounded-lg">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead className="w-12">
                <Checkbox
                  checked={selectedJobs.length === filteredJobs.length && filteredJobs.length > 0}
                  onCheckedChange={handleSelectAll}
                />
              </TableHead>
              <TableHead>
                <Button
                  variant="ghost"
                  className="h-auto p-0 font-semibold"
                  onClick={() => handleSort('fileName')}
                >
                  File Name
                  <SortIcon field="fileName" />
                </Button>
              </TableHead>
              <TableHead>Type</TableHead>
              <TableHead>
                <Button
                  variant="ghost"
                  className="h-auto p-0 font-semibold"
                  onClick={() => handleSort('status')}
                >
                  Status
                  <SortIcon field="status" />
                </Button>
              </TableHead>
              <TableHead>Progress</TableHead>
              <TableHead>
                <Button
                  variant="ghost"
                  className="h-auto p-0 font-semibold"
                  onClick={() => handleSort('createdAt')}
                >
                  Created
                  <SortIcon field="createdAt" />
                </Button>
              </TableHead>
              <TableHead>Duration</TableHead>
              <TableHead>Cost</TableHead>
              <TableHead className="w-12"></TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filteredJobs.length === 0 ? (
              <TableRow>
                <TableCell colSpan={9} className="text-center py-8 text-gray-500">
                  No jobs found matching your filters.
                </TableCell>
              </TableRow>
            ) : (
              filteredJobs.map((job) => (
                <TableRow
                  key={job.id}
                  className="cursor-pointer hover:bg-gray-50"
                  onClick={() => onJobClick(job)}
                >
                  <TableCell onClick={(e) => e.stopPropagation()}>
                    <Checkbox
                      checked={selectedJobs.includes(job.id)}
                      onCheckedChange={(checked) => handleSelectJob(job.id, checked as boolean)}
                    />
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center space-x-2">
                      <FileText className="h-4 w-4 text-gray-500" />
                      <span className="font-medium">{job.fileName}</span>
                    </div>
                  </TableCell>
                  <TableCell>
                    <Badge variant="outline">{job.fileType.toUpperCase()}</Badge>
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center space-x-2">
                      {getStatusIcon(job.status)}
                      <Badge variant={STATUS_COLORS[job.status as keyof typeof STATUS_COLORS]}>
                        {job.status}
                      </Badge>
                    </div>
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center space-x-2">
                      <Progress value={job.progress} className="w-20" />
                      <span className="text-sm text-gray-600">{job.progress}%</span>
                    </div>
                  </TableCell>
                  <TableCell>{formatDate(job.createdAt)}</TableCell>
                  <TableCell>
                    {job.startedAt && (
                      <span className="text-sm text-gray-500">{getDuration(job)}</span>
                    )}
                  </TableCell>
                  <TableCell>
                    {job.actualCost && (
                      <span className="text-sm">${job.actualCost.toFixed(4)}</span>
                    )}
                  </TableCell>
                  <TableCell onClick={(e) => e.stopPropagation()}>
                    <DropdownMenu>
                      <DropdownMenuTrigger asChild>
                        <Button variant="ghost" size="sm">
                          <MoreVertical className="h-4 w-4" />
                        </Button>
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="end">
                        <DropdownMenuItem onClick={() => onJobClick(job)}>
                          <Eye className="h-4 w-4 mr-2" />
                          View Details
                        </DropdownMenuItem>
                        {job.downloadUrl && (
                          <DropdownMenuItem onClick={() => onDownloadJob(job.id)}>
                            <Download className="h-4 w-4 mr-2" />
                            Download
                          </DropdownMenuItem>
                        )}
                        {job.status === 'failed' && (
                          <DropdownMenuItem onClick={() => onRetryJob(job.id)}>
                            <Play className="h-4 w-4 mr-2" />
                            Retry
                          </DropdownMenuItem>
                        )}
                        {(job.status === 'pending' || job.status === 'processing') && (
                          <DropdownMenuItem onClick={() => onCancelJob(job.id)}>
                            <X className="h-4 w-4 mr-2" />
                            Cancel
                          </DropdownMenuItem>
                        )}
                        {(job.status === 'completed' || job.status === 'failed' || job.status === 'cancelled') && (
                          <DropdownMenuItem
                            onClick={() => onDeleteJob(job.id)}
                            className="text-red-600"
                          >
                            <Trash2 className="h-4 w-4 mr-2" />
                            Delete
                          </DropdownMenuItem>
                        )}
                      </DropdownMenuContent>
                    </DropdownMenu>
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>

      {/* Summary */}
      <div className="text-sm text-gray-500">
        Showing {filteredJobs.length} of {jobs.length} jobs
        {selectedJobs.length > 0 && (
          <span className="ml-2">
            ({selectedJobs.length} selected)
          </span>
        )}
      </div>
    </div>
  );
}
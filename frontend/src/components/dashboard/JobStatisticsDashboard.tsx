import React, { useState, useEffect } from 'react';
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from '@/components/ui/card';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Button } from '@/components/ui/button';
import { apiClient } from '@/lib/api';
import {
  BarChart3,
  TrendingUp,
  Clock,
  DollarSign,
  FileText,
  CheckCircle,
  XCircle,
  RefreshCw,
  Download,
} from 'lucide-react';

interface JobStatistics {
  total_jobs: number;
  status_counts: Record<string, number>;
  average_duration_minutes: number;
  total_cost: number;
  daily_stats: Array<{
    date: string;
    total: number;
    completed: number;
    failed: number;
  }>;
  file_type_distribution: Record<string, number>;
  period_days: number;
}

interface JobStatisticsDashboardProps {
  className?: string;
}

export function JobStatisticsDashboard({ className }: JobStatisticsDashboardProps) {
  const [stats, setStats] = useState<JobStatistics | null>(null);
  const [selectedPeriod, setSelectedPeriod] = useState<number>(30);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    fetchStatistics(selectedPeriod);
  }, [selectedPeriod]);

  const fetchStatistics = async (days: number) => {
    setIsLoading(true);
    try {
      const response = await apiClient.getJobStatistics(days);
      if (response.success && response.data) {
        setStats(response.data);
      }
    } catch (error) {
      console.error('Failed to fetch job statistics:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleExport = async () => {
    try {
      const response = await apiClient.exportJobs('csv');
      if (response.success && response.data) {
        const blob = new Blob([response.data.data], {
          type: response.data.media_type,
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
      console.error('Failed to export data:', error);
    }
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD',
    }).format(value);
  };

  const formatDuration = (minutes: number) => {
    if (minutes < 1) return '< 1 min';
    if (minutes < 60) return `${Math.round(minutes)} min`;
    const hours = Math.floor(minutes / 60);
    const mins = Math.round(minutes % 60);
    return `${hours}h ${mins}m`;
  };

  const getSuccessRate = () => {
    if (!stats || stats.total_jobs === 0) return 0;
    const completed = stats.status_counts.completed || 0;
    return Math.round((completed / stats.total_jobs) * 100);
  };

  const renderChart = () => {
    if (!stats || stats.daily_stats.length === 0) {
      return (
        <div className="flex items-center justify-center h-64 text-gray-500">
          No data available for the selected period
        </div>
      );
    }

    const maxValue = Math.max(...stats.daily_stats.map(d => d.total));
    const chartHeight = 200;
    const barWidth = Math.max(20, 600 / stats.daily_stats.length);

    return (
      <div className="relative">
        <div className="flex items-end justify-between h-[200px] space-x-1">
          {stats.daily_stats.slice(-7).map((day, index) => (
            <div key={day.date} className="flex flex-col items-center flex-1">
              <div className="flex items-end justify-center w-full space-x-1">
                <div
                  className="bg-green-500 rounded-t transition-all duration-300 hover:bg-green-600"
                  style={{
                    height: `${(day.completed / maxValue) * chartHeight}px`,
                    width: `${barWidth * 0.4}px`,
                  }}
                  title={`Completed: ${day.completed}`}
                />
                <div
                  className="bg-red-500 rounded-t transition-all duration-300 hover:bg-red-600"
                  style={{
                    height: `${(day.failed / maxValue) * chartHeight}px`,
                    width: `${barWidth * 0.4}px`,
                  }}
                  title={`Failed: ${day.failed}`}
                />
              </div>
              <div className="text-xs text-gray-500 mt-2 rotate-45 origin-left">
                {new Date(day.date).toLocaleDateString('en-US', { weekday: 'short' })}
              </div>
            </div>
          ))}
        </div>
        <div className="flex justify-center space-x-4 mt-4">
          <div className="flex items-center space-x-2">
            <div className="w-3 h-3 bg-green-500 rounded"></div>
            <span className="text-sm">Completed</span>
          </div>
          <div className="flex items-center space-x-2">
            <div className="w-3 h-3 bg-red-500 rounded"></div>
            <span className="text-sm">Failed</span>
          </div>
        </div>
      </div>
    );
  };

  if (isLoading) {
    return (
      <Card className={className}>
        <CardContent className="flex items-center justify-center h-64">
          <RefreshCw className="h-8 w-8 animate-spin text-gray-400" />
        </CardContent>
      </Card>
    );
  }

  if (!stats) {
    return (
      <Card className={className}>
        <CardContent className="flex items-center justify-center h-64">
          <p className="text-gray-500">Failed to load statistics</p>
        </CardContent>
      </Card>
    );
  }

  return (
    <div className={`space-y-6 ${className}`}>
      {/* Header */}
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-2xl font-bold">Job Statistics</h2>
          <p className="text-gray-600">Overview of your translation jobs</p>
        </div>
        <div className="flex items-center space-x-4">
          <Select value={selectedPeriod.toString()} onValueChange={(v) => setSelectedPeriod(parseInt(v))}>
            <SelectTrigger className="w-[180px]">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="7">Last 7 days</SelectItem>
              <SelectItem value="30">Last 30 days</SelectItem>
              <SelectItem value="90">Last 90 days</SelectItem>
            </SelectContent>
          </Select>
          <Button variant="outline" onClick={handleExport}>
            <Download className="mr-2 h-4 w-4" />
            Export Data
          </Button>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium">Total Jobs</CardTitle>
            <FileText className="h-4 w-4 text-muted-foreground" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold">{stats.total_jobs}</div>
            <p className="text-xs text-muted-foreground">
              {selectedPeriod} days
            </p>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium">Success Rate</CardTitle>
            <TrendingUp className="h-4 w-4 text-muted-foreground" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold text-green-600">{getSuccessRate()}%</div>
            <div className="mt-2">
              <Progress value={getSuccessRate()} className="h-2" />
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium">Avg Duration</CardTitle>
            <Clock className="h-4 w-4 text-muted-foreground" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold">{formatDuration(stats.average_duration_minutes)}</div>
            <p className="text-xs text-muted-foreground">
              Per job
            </p>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
            <CardTitle className="text-sm font-medium">Total Cost</CardTitle>
            <DollarSign className="h-4 w-4 text-muted-foreground" />
          </CardHeader>
          <CardContent>
            <div className="text-2xl font-bold">{formatCurrency(stats.total_cost)}</div>
            <p className="text-xs text-muted-foreground">
              {selectedPeriod} days
            </p>
          </CardContent>
        </Card>
      </div>

      {/* Status Breakdown */}
      <Card>
        <CardHeader>
          <CardTitle>Job Status Distribution</CardTitle>
          <CardDescription>Breakdown of job statuses</CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
            {Object.entries(stats.status_counts).map(([status, count]) => {
              const percentage = stats.total_jobs > 0 ? Math.round((count / stats.total_jobs) * 100) : 0;
              const statusConfig = {
                completed: { color: 'bg-green-500', icon: CheckCircle, label: 'Completed' },
                failed: { color: 'bg-red-500', icon: XCircle, label: 'Failed' },
                processing: { color: 'bg-blue-500', icon: RefreshCw, label: 'Processing' },
                pending: { color: 'bg-gray-500', icon: Clock, label: 'Pending' },
                cancelled: { color: 'bg-yellow-500', icon: XCircle, label: 'Cancelled' },
              };
              const config = statusConfig[status as keyof typeof statusConfig] || statusConfig.pending;
              const Icon = config.icon;

              return (
                <div key={status} className="text-center">
                  <div className={`w-16 h-16 rounded-full ${config.color} flex items-center justify-center mx-auto mb-2`}>
                    <Icon className="h-8 w-8 text-white" />
                  </div>
                  <div className="text-2xl font-bold">{count}</div>
                  <div className="text-sm text-gray-600">{config.label}</div>
                  <div className="text-xs text-gray-500">{percentage}%</div>
                </div>
              );
            })}
          </div>
        </CardContent>
      </Card>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Daily Activity Chart */}
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <BarChart3 className="h-5 w-5" />
              <span>Daily Activity (Last 7 Days)</span>
            </CardTitle>
          </CardHeader>
          <CardContent>
            {renderChart()}
          </CardContent>
        </Card>

        {/* File Type Distribution */}
        <Card>
          <CardHeader>
            <CardTitle>File Type Distribution</CardTitle>
            <CardDescription>Most used file formats</CardDescription>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              {Object.entries(stats.file_type_distribution)
                .sort((a, b) => b[1] - a[1])
                .map(([type, count]) => {
                  const percentage = stats.total_jobs > 0 ? Math.round((count / stats.total_jobs) * 100) : 0;
                  return (
                    <div key={type}>
                      <div className="flex justify-between items-center mb-1">
                        <span className="text-sm font-medium capitalize">{type}</span>
                        <span className="text-sm text-gray-600">{count} ({percentage}%)</span>
                      </div>
                      <Progress value={percentage} className="h-2" />
                    </div>
                  );
                })}
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Detailed Stats */}
      <Tabs defaultValue="daily" className="w-full">
        <TabsList className="grid w-full grid-cols-2">
          <TabsTrigger value="daily">Daily Breakdown</TabsTrigger>
          <TabsTrigger value="insights">Insights</TabsTrigger>
        </TabsList>

        <TabsContent value="daily">
          <Card>
            <CardHeader>
              <CardTitle>Daily Job Summary</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-2">
                {stats.daily_stats.slice(-10).reverse().map((day) => {
                  const date = new Date(day.date);
                  const successRate = day.total > 0 ? Math.round((day.completed / day.total) * 100) : 0;

                  return (
                    <div key={day.date} className="flex items-center justify-between p-3 border rounded-lg">
                      <div className="flex items-center space-x-4">
                        <div className="text-sm font-medium">
                          {date.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' })}
                        </div>
                        <div className="flex space-x-2">
                          <Badge variant="secondary">{day.total} total</Badge>
                          <Badge variant="default" className="bg-green-100 text-green-800">
                            {day.completed} completed
                          </Badge>
                          {day.failed > 0 && (
                            <Badge variant="destructive">{day.failed} failed</Badge>
                          )}
                        </div>
                      </div>
                      <div className="flex items-center space-x-2">
                        <Progress value={successRate} className="w-20" />
                        <span className="text-sm text-gray-600">{successRate}%</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="insights">
          <Card>
            <CardHeader>
              <CardTitle>Key Insights</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="p-4 border rounded-lg">
                  <h4 className="font-semibold mb-2">Peak Performance</h4>
                  <p className="text-sm text-gray-600">
                    Your average job completion time is {formatDuration(stats.average_duration_minutes)}, which is{' '}
                    {stats.average_duration_minutes < 10 ? 'excellent' : stats.average_duration_minutes < 30 ? 'good' : 'above average'}.
                  </p>
                </div>
                <div className="p-4 border rounded-lg">
                  <h4 className="font-semibold mb-2">Cost Efficiency</h4>
                  <p className="text-sm text-gray-600">
                    Average cost per job: {stats.total_jobs > 0 ? formatCurrency(stats.total_cost / stats.total_jobs) : 'N/A'}
                  </p>
                </div>
                <div className="p-4 border rounded-lg">
                  <h4 className="font-semibold mb-2">Reliability</h4>
                  <p className="text-sm text-gray-600">
                    {getSuccessRate()}% success rate shows {getSuccessRate() > 95 ? 'excellent' : getSuccessRate() > 85 ? 'good' : 'room for improvement'} reliability.
                  </p>
                </div>
                <div className="p-4 border rounded-lg">
                  <h4 className="font-semibold mb-2">Preferred Format</h4>
                  <p className="text-sm text-gray-600">
                    {Object.entries(stats.file_type_distribution).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A'} is your most translated format.
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
}
import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import {
  Star,
  TrendingUp,
  AlertTriangle,
  CheckCircle,
  Clock,
  DollarSign,
  FileText,
  Zap,
} from 'lucide-react';

interface QualityMetricsProps {
  metrics?: Record<string, any>;
  className?: string;
}

export function QualityMetrics({ metrics, className }: QualityMetricsProps) {
  if (!metrics) {
    return (
      <Card className={className}>
        <CardContent className="flex items-center justify-center h-32">
          <p className="text-gray-500">Quality assessment not available</p>
        </CardContent>
      </Card>
    );
  }

  const getQualityColor = (grade: string) => {
    switch (grade) {
      case 'excellent':
        return 'text-green-600 bg-green-100 border-green-200';
      case 'good':
        return 'text-blue-600 bg-blue-100 border-blue-200';
      case 'fair':
        return 'text-yellow-600 bg-yellow-100 border-yellow-200';
      case 'poor':
        return 'text-red-600 bg-red-100 border-red-200';
      default:
        return 'text-gray-600 bg-gray-100 border-gray-200';
    }
  };

  const getQualityIcon = (grade: string) => {
    switch (grade) {
      case 'excellent':
        return <Star className="h-5 w-5 text-green-600" />;
      case 'good':
        return <CheckCircle className="h-5 w-5 text-blue-600" />;
      case 'fair':
        return <AlertTriangle className="h-5 w-5 text-yellow-600" />;
      case 'poor':
        return <AlertTriangle className="h-5 w-5 text-red-600" />;
      default:
        return <Clock className="h-5 w-5 text-gray-600" />;
    }
  };

  const formatMetric = (key: string, value: any) => {
    switch (key) {
      case 'overall_score':
      case 'completion_rate':
      case 'cost_efficiency':
      case 'file_integrity':
        return `${(value * 100).toFixed(1)}%`;
      case 'error_rate':
        return `${(value * 100).toFixed(1)}%`;
      case 'processing_time':
        return value < 1 ? '< 1 min' : value < 60 ? `${Math.round(value)} min` : `${Math.floor(value / 60)}h ${Math.round(value % 60)}m`;
      case 'size_ratio':
        return value.toFixed(2) + 'x';
      default:
        return value?.toString() || 'N/A';
    }
  };

  const recommendations = metrics.recommendations || [];

  return (
    <div className={className}>
      {/* Quality Grade Header */}
      <Card className="mb-4">
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              {getQualityIcon(metrics.quality_grade)}
              <div>
                <span className="text-lg font-semibold">Translation Quality</span>
                <div className="flex items-center space-x-2 mt-1">
                  <Badge className={getQualityColor(metrics.quality_grade)}>
                    {metrics.quality_grade?.toUpperCase() || 'UNKNOWN'}
                  </Badge>
                  <span className="text-sm text-gray-600">
                    Score: {(metrics.overall_score * 100).toFixed(1)}%
                  </span>
                </div>
              </div>
            </div>
            <div className="text-right">
              <div className="text-2xl font-bold">{(metrics.overall_score * 100).toFixed(0)}</div>
              <div className="text-xs text-gray-500">Quality Score</div>
            </div>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <Progress value={metrics.overall_score * 100} className="h-3" />
        </CardContent>
      </Card>

      {/* Detailed Metrics */}
      <Tabs defaultValue="overview" className="w-full">
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="overview">Overview</TabsTrigger>
          <TabsTrigger value="performance">Performance</TabsTrigger>
          <TabsTrigger value="details">Details</TabsTrigger>
        </TabsList>

        <TabsContent value="overview">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <Card>
              <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                <CardTitle className="text-sm font-medium">Completion Rate</CardTitle>
                <CheckCircle className="h-4 w-4 text-muted-foreground" />
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold">
                  {formatMetric('completion_rate', metrics.completion_rate)}
                </div>
                <Progress value={metrics.completion_rate * 100} className="mt-2 h-2" />
              </CardContent>
            </Card>

            <Card>
              <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                <CardTitle className="text-sm font-medium">Error Rate</CardTitle>
                <AlertTriangle className="h-4 w-4 text-muted-foreground" />
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold">
                  {formatMetric('error_rate', metrics.error_rate)}
                </div>
                <Progress value={(1 - metrics.error_rate) * 100} className="mt-2 h-2" />
              </CardContent>
            </Card>

            <Card>
              <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                <CardTitle className="text-sm font-medium">Processing Time</CardTitle>
                <Clock className="h-4 w-4 text-muted-foreground" />
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold">
                  {formatMetric('processing_time', metrics.processing_time)}
                </div>
                <p className="text-xs text-muted-foreground mt-1">
                  Total duration
                </p>
              </CardContent>
            </Card>

            <Card>
              <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                <CardTitle className="text-sm font-medium">Cost Efficiency</CardTitle>
                <DollarSign className="h-4 w-4 text-muted-foreground" />
              </CardHeader>
              <CardContent>
                <div className="text-2xl font-bold">
                  {formatMetric('cost_efficiency', metrics.cost_efficiency)}
                </div>
                <Progress value={metrics.cost_efficiency * 100} className="mt-2 h-2" />
              </CardContent>
            </Card>
          </div>
        </TabsContent>

        <TabsContent value="performance">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Zap className="h-5 w-5" />
                <span>Performance Metrics</span>
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-4">
                {metrics.file_integrity !== undefined && (
                  <div>
                    <div className="flex justify-between items-center mb-2">
                      <span className="text-sm font-medium">File Integrity</span>
                      <span className="text-sm text-gray-600">
                        {formatMetric('file_integrity', metrics.file_integrity)}
                      </span>
                    </div>
                    <Progress value={metrics.file_integrity * 100} className="h-2" />
                  </div>
                )}

                {metrics.size_ratio !== undefined && (
                  <div>
                    <div className="flex justify-between items-center mb-2">
                      <span className="text-sm font-medium">Size Ratio (Output/Input)</span>
                      <span className="text-sm text-gray-600">
                        {formatMetric('size_ratio', metrics.size_ratio)}
                      </span>
                    </div>
                    <div className="text-xs text-gray-500 mt-1">
                      {metrics.size_ratio < 0.5 && 'Output significantly smaller'}
                      {metrics.size_ratio >= 0.5 && metrics.size_ratio <= 1.5 && 'Normal size variation'}
                      {metrics.size_ratio > 1.5 && 'Output significantly larger'}
                    </div>
                  </div>
                )}

                {metrics.issues_found && metrics.issues_found.length > 0 && (
                  <Alert>
                    <AlertTriangle className="h-4 w-4" />
                    <AlertDescription>
                      <div className="font-medium mb-1">Issues Detected:</div>
                      <ul className="list-disc list-inside text-sm">
                        {metrics.issues_found.map((issue: string, index: number) => (
                          <li key={index}>{issue}</li>
                        ))}
                      </ul>
                    </AlertDescription>
                  </Alert>
                )}
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="details">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <TrendingUp className="h-5 w-5" />
                <span>All Metrics</span>
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {Object.entries(metrics).map(([key, value]) => {
                  if (['quality_grade', 'overall_score', 'recommendations', 'issues_found', 'assessment_error'].includes(key)) {
                    return null;
                  }
                  return (
                    <div key={key} className="flex justify-between items-center p-3 border rounded-lg">
                      <span className="text-sm font-medium capitalize">
                        {key.replace(/_/g, ' ')}
                      </span>
                      <span className="text-sm font-semibold">
                        {formatMetric(key, value)}
                      </span>
                    </div>
                  );
                })}
              </div>

              {recommendations.length > 0 && (
                <div className="mt-6">
                  <h4 className="font-semibold mb-3">Recommendations</h4>
                  <div className="space-y-2">
                    {recommendations.map((rec: string, index: number) => (
                      <Alert key={index}>
                        <AlertTriangle className="h-4 w-4" />
                        <AlertDescription>{rec}</AlertDescription>
                      </Alert>
                    ))}
                  </div>
                </div>
              )}

              {metrics.assessment_error && (
                <Alert className="mt-4" variant="destructive">
                  <AlertTriangle className="h-4 w-4" />
                  <AlertDescription>
                    Quality assessment error: {metrics.assessment_error}
                  </AlertDescription>
                </Alert>
              )}
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
}
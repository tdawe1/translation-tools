'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Textarea } from '@/components/ui/textarea';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Switch } from '@/components/ui/switch';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Badge } from '@/components/ui/badge';
import { useAuth } from '@/contexts/AuthContext';
import { apiClient, type Settings } from '@/lib/api';
import {
  Save,
  RotateCcw,
  Key,
  Settings as SettingsIcon,
  FileText,
  Database,
  Shield,
  CheckCircle,
  AlertCircle,
  Eye,
  EyeOff,
} from 'lucide-react';

export default function SettingsPage() {
  const { isAuthenticated } = useAuth();
  const router = useRouter();
  const [settings, setSettings] = useState<Settings>({});
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [showApiKeys, setShowApiKeys] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  useEffect(() => {
    if (!isAuthenticated) {
      router.push('/login');
      return;
    }

    fetchSettings();
  }, [isAuthenticated, router]);

  const fetchSettings = async () => {
    try {
      const response = await apiClient.getSettings();
      if (response.success && response.data) {
        setSettings(response.data);
      }
    } catch (error) {
      console.error('Failed to fetch settings:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSave = async () => {
    setIsSaving(true);
    setMessage(null);

    try {
      const response = await apiClient.updateSettings(settings);
      if (response.success) {
        setMessage({ type: 'success', text: 'Settings saved successfully!' });
      } else {
        setMessage({ type: 'error', text: response.error || 'Failed to save settings' });
      }
    } catch (error: any) {
      setMessage({ type: 'error', text: error.response?.data?.message || error.message });
    } finally {
      setIsSaving(false);
    }
  };

  const handleReset = () => {
    setSettings({});
    setMessage(null);
  };

  const updateSetting = (key: keyof Settings, value: any) => {
    setSettings(prev => ({ ...prev, [key]: value }));
  };

  if (!isAuthenticated) {
    return null;
  }

  return (
    <div>
      <div className="mb-8 flex justify-between items-center">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Settings</h1>
          <p className="text-gray-600">Configure your translation pipeline</p>
        </div>
        <div className="flex items-center space-x-4">
          <Button variant="outline" onClick={handleReset}>
            <RotateCcw className="mr-2 h-4 w-4" />
            Reset
          </Button>
          <Button onClick={handleSave} disabled={isSaving}>
            <Save className="mr-2 h-4 w-4" />
            {isSaving ? 'Saving...' : 'Save Changes'}
          </Button>
        </div>
      </div>

      <main className="max-w-4xl mx-auto">
        {message && (
          <Alert className={`mb-6 ${message.type === 'success' ? 'border-green-200 bg-green-50' : 'border-red-200 bg-red-50'}`}>
            {message.type === 'success' ? (
              <CheckCircle className="h-4 w-4 text-green-600" />
            ) : (
              <AlertCircle className="h-4 w-4 text-red-600" />
            )}
            <AlertDescription className={message.type === 'success' ? 'text-green-800' : 'text-red-800'}>
              {message.text}
            </AlertDescription>
          </Alert>
        )}

        <div className="space-y-6">
          {/* API Settings */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Key className="h-5 w-5" />
                <span>API Configuration</span>
              </CardTitle>
              <CardDescription>
                Configure your AI service provider settings
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex items-center justify-between mb-4">
                <Label>Show API Keys</Label>
                <Switch
                  checked={showApiKeys}
                  onCheckedChange={setShowApiKeys}
                />
              </div>

              <div className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="openaiApiKey">OpenAI API Key</Label>
                  <div className="relative">
                    <Input
                      id="openaiApiKey"
                      type={showApiKeys ? 'text' : 'password'}
                      value={settings.openaiApiKey || ''}
                      onChange={(e) => updateSetting('openaiApiKey', e.target.value)}
                      placeholder="sk-..."
                    />
                    <Button
                      variant="ghost"
                      size="sm"
                      className="absolute right-2 top-1/2 transform -translate-y-1/2"
                      onClick={() => setShowApiKeys(!showApiKeys)}
                    >
                      {showApiKeys ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                    </Button>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label htmlFor="openaiModel">Default Model</Label>
                    <Select
                      value={settings.openaiModel || 'gpt-4o-2024-08-06'}
                      onValueChange={(value) => updateSetting('openaiModel', value)}
                    >
                      <SelectTrigger>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="gpt-4o-2024-08-06">GPT-4o (Recommended)</SelectItem>
                        <SelectItem value="gpt-4o-mini">GPT-4o Mini</SelectItem>
                        <SelectItem value="gpt-5">GPT-5</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="openaiTemperature">Temperature</Label>
                    <Input
                      id="openaiTemperature"
                      type="number"
                      step="0.1"
                      min="0"
                      max="1"
                      value={settings.openaiTemperature || 0.6}
                      onChange={(e) => updateSetting('openaiTemperature', parseFloat(e.target.value))}
                    />
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Feature Flags */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <SettingsIcon className="h-5 w-5" />
                <span>Feature Flags</span>
              </CardTitle>
              <CardDescription>
                Enable or disable translation features
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex items-center justify-between">
                <div className="space-y-0.5">
                  <Label>Style Checking</Label>
                  <p className="text-sm text-muted-foreground">
                    Enable style consistency validation
                  </p>
                </div>
                <Switch
                  checked={settings.enableStyleChecking || false}
                  onCheckedChange={(checked) => updateSetting('enableStyleChecking', checked)}
                />
              </div>

              <div className="flex items-center justify-between">
                <div className="space-y-0.5">
                  <Label>Expansion Policy</Label>
                  <p className="text-sm text-muted-foreground">
                    Handle text expansion in translated content
                  </p>
                </div>
                <Switch
                  checked={settings.enableExpansionPolicy || false}
                  onCheckedChange={(checked) => updateSetting('enableExpansionPolicy', checked)}
                />
              </div>

              <div className="flex items-center justify-between">
                <div className="space-y-0.5">
                  <Label>Formatting Profile</Label>
                  <p className="text-sm text-muted-foreground">
                    Apply formatting optimization
                  </p>
                </div>
                <Switch
                  checked={settings.enableFormattingProfile || false}
                  onCheckedChange={(checked) => updateSetting('enableFormattingProfile', checked)}
                />
              </div>
            </CardContent>
          </Card>

          {/* Google Drive Integration */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Database className="h-5 w-5" />
                <span>Google Drive Integration</span>
                <Badge variant="outline">Optional</Badge>
              </CardTitle>
              <CardDescription>
                Connect Google Drive for file storage and retrieval
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="googleOauthClientId">OAuth Client ID</Label>
                  <Input
                    id="googleOauthClientId"
                    type={showApiKeys ? 'text' : 'password'}
                    value={settings.googleOauthClientId || ''}
                    onChange={(e) => updateSetting('googleOauthClientId', e.target.value)}
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="googleOauthClientSecret">OAuth Client Secret</Label>
                  <Input
                    id="googleOauthClientSecret"
                    type={showApiKeys ? 'text' : 'password'}
                    value={settings.googleOauthClientSecret || ''}
                    onChange={(e) => updateSetting('googleOauthClientSecret', e.target.value)}
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="googleOauthRefreshToken">Refresh Token</Label>
                  <Input
                    id="googleOauthRefreshToken"
                    type={showApiKeys ? 'text' : 'password'}
                    value={settings.googleOauthRefreshToken || ''}
                    onChange={(e) => updateSetting('googleOauthRefreshToken', e.target.value)}
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="gdriveSaJson">Service Account JSON</Label>
                  <Textarea
                    id="gdriveSaJson"
                    placeholder='{"type": "service_account", ...}'
                    value={settings.gdriveSaJson || ''}
                    onChange={(e) => updateSetting('gdriveSaJson', e.target.value)}
                    rows={3}
                  />
                </div>
              </div>

              <Alert>
                <Shield className="h-4 w-4" />
                <AlertDescription>
                  Your Google Drive credentials are encrypted and stored securely. Only provide credentials
                  from trusted sources.
                </AlertDescription>
              </Alert>
            </CardContent>
          </Card>

          {/* Advanced Settings */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <FileText className="h-5 w-5" />
                <span>Advanced Configuration</span>
              </CardTitle>
              <CardDescription>
                Additional settings for power users
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="space-y-4">
                <div className="p-4 bg-gray-50 rounded-lg">
                  <h3 className="font-medium mb-2">Environment Variables</h3>
                  <div className="text-sm text-gray-600 space-y-1">
                    <p><code>NEXT_PUBLIC_API_BASE_URL</code> - Backend API URL</p>
                    <p><code>NODE_ENV</code> - Environment mode</p>
                  </div>
                </div>

                <div className="p-4 bg-blue-50 rounded-lg">
                  <h3 className="font-medium mb-2">Current Configuration</h3>
                  <div className="text-sm text-gray-600 space-y-1">
                    <p>API Base URL: {process.env.NEXT_PUBLIC_API_BASE_URL || 'http://localhost:8080/api'}</p>
                    <p>Environment: {process.env.NODE_ENV || 'development'}</p>
                    <p>Default Model: {settings.openaiModel || 'gpt-4o-2024-08-06'}</p>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  );
}
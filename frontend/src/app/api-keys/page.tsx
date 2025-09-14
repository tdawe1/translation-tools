'use client';

import { useState, useEffect } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { toast } from 'react-hot-toast';

interface APIKey {
  id: string;
  name: string;
  prefix: string;
  created_at: string;
  last_used: string | null;
}

export default function APIKeysPage() {
  const { getAuthHeaders, isAuthenticated } = useAuth();
  const [apiKeys, setApiKeys] = useState<APIKey[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [showCreateForm, setShowCreateForm] = useState(false);
  const [newKeyName, setNewKeyName] = useState('');
  const [newApiKey, setNewApiKey] = useState<string | null>(null);
  const [isCreating, setIsCreating] = useState(false);

  const API_BASE_URL = process.env.NEXT_PUBLIC_API_URL || 'http://localhost:8000/api';

  useEffect(() => {
    if (isAuthenticated) {
      fetchAPIKeys();
    }
  }, [isAuthenticated]);

  const fetchAPIKeys = async () => {
    try {
      const response = await fetch(`${API_BASE_URL}/auth/api-keys`, {
        headers: getAuthHeaders()
      });

      if (response.ok) {
        const data = await response.json();
        setApiKeys(data.keys || []);
      } else {
        toast.error('Failed to fetch API keys');
      }
    } catch (error) {
      console.error('Error fetching API keys:', error);
      toast.error('Network error');
    } finally {
      setIsLoading(false);
    }
  };

  const createAPIKey = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newKeyName.trim()) {
      toast.error('Please enter a name for the API key');
      return;
    }

    setIsCreating(true);
    try {
      const response = await fetch(`${API_BASE_URL}/auth/api-keys`, {
        method: 'POST',
        headers: {
          ...getAuthHeaders(),
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ name: newKeyName })
      });

      if (response.ok) {
        const data = await response.json();
        setNewApiKey(data.api_key);
        setNewKeyName('');
        setShowCreateForm(false);
        // Refresh the list
        fetchAPIKeys();
        toast.success('API key created successfully');
      } else {
        const error = await response.json();
        toast.error(error.detail || 'Failed to create API key');
      }
    } catch (error) {
      console.error('Error creating API key:', error);
      toast.error('Network error');
    } finally {
      setIsCreating(false);
    }
  };

  const revokeAPIKey = async (keyId: string) => {
    if (!confirm('Are you sure you want to revoke this API key? This action cannot be undone.')) {
      return;
    }

    try {
      const response = await fetch(`${API_BASE_URL}/auth/api-keys/${keyId}`, {
        method: 'DELETE',
        headers: getAuthHeaders()
      });

      if (response.ok) {
        // Remove from the list
        setApiKeys(keys => keys.filter(key => key.id !== keyId));
        toast.success('API key revoked successfully');
      } else {
        toast.error('Failed to revoke API key');
      }
    } catch (error) {
      console.error('Error revoking API key:', error);
      toast.error('Network error');
    }
  };

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    toast.success('Copied to clipboard');
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <h1 className="text-2xl font-bold text-gray-900 mb-4">Please Log In</h1>
          <p className="text-gray-600">You need to be logged in to manage API keys.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="md:flex md:items-center md:justify-between mb-8">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">API Keys</h1>
            <p className="mt-2 text-sm text-gray-600">
              Manage your API keys for programmatic access to the translation pipeline
            </p>
          </div>
          <div className="mt-4 md:mt-0">
            <button
              onClick={() => setShowCreateForm(true)}
              className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md text-white bg-indigo-600 hover:bg-indigo-700"
            >
              Create New Key
            </button>
          </div>
        </div>

        {/* New API Key Modal */}
        {newApiKey && (
          <div className="fixed inset-0 bg-gray-500 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg p-6 max-w-md w-full mx-4">
              <h3 className="text-lg font-medium text-gray-900 mb-4">New API Key Created</h3>
              <p className="text-sm text-gray-600 mb-4">
                Copy this key now. You won't be able to see it again.
              </p>
              <div className="bg-gray-100 p-3 rounded-md font-mono text-sm mb-4 break-all">
                {newApiKey}
              </div>
              <div className="flex justify-end space-x-3">
                <button
                  onClick={() => copyToClipboard(newApiKey)}
                  className="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300"
                >
                  Copy
                </button>
                <button
                  onClick={() => setNewApiKey(null)}
                  className="px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700"
                >
                  I've Saved It
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Create Form Modal */}
        {showCreateForm && (
          <div className="fixed inset-0 bg-gray-500 bg-opacity-75 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg p-6 max-w-md w-full mx-4">
              <h3 className="text-lg font-medium text-gray-900 mb-4">Create New API Key</h3>
              <form onSubmit={createAPIKey}>
                <div className="mb-4">
                  <label htmlFor="key-name" className="block text-sm font-medium text-gray-700">
                    Key Name
                  </label>
                  <input
                    type="text"
                    id="key-name"
                    value={newKeyName}
                    onChange={(e) => setNewKeyName(e.target.value)}
                    className="mt-1 block w-full border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500"
                    placeholder="e.g., Production API Key"
                    required
                  />
                </div>
                <div className="flex justify-end space-x-3">
                  <button
                    type="button"
                    onClick={() => setShowCreateForm(false)}
                    className="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    disabled={isCreating}
                    className="px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:opacity-50"
                  >
                    {isCreating ? 'Creating...' : 'Create Key'}
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* API Keys List */}
        {isLoading ? (
          <div className="text-center py-12">
            <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-gray-900 mx-auto"></div>
            <p className="mt-4 text-gray-600">Loading API keys...</p>
          </div>
        ) : apiKeys.length === 0 ? (
          <div className="bg-white shadow rounded-lg p-6 text-center">
            <div className="text-gray-400 mb-4">
              <svg className="mx-auto h-12 w-12" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
              </svg>
            </div>
            <h3 className="text-lg font-medium text-gray-900 mb-2">No API Keys</h3>
            <p className="text-gray-600 mb-4">
              Create your first API key to start using the translation pipeline programmatically.
            </p>
            <button
              onClick={() => setShowCreateForm(true)}
              className="inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md text-white bg-indigo-600 hover:bg-indigo-700"
            >
              Create Your First Key
            </button>
          </div>
        ) : (
          <div className="bg-white shadow overflow-hidden sm:rounded-md">
            <ul className="divide-y divide-gray-200">
              {apiKeys.map((key) => (
                <li key={key.id} className="px-6 py-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="text-sm font-medium text-gray-900">{key.name}</h4>
                      <p className="text-sm text-gray-500 font-mono">{key.prefix}...</p>
                      <p className="text-xs text-gray-400 mt-1">
                        Created: {new Date(key.created_at).toLocaleDateString()}
                        {key.last_used && ` • Last used: ${new Date(key.last_used).toLocaleDateString()}`}
                      </p>
                    </div>
                    <div className="flex space-x-2">
                      <button
                        onClick={() => copyToClipboard(key.prefix + '...')}
                        className="text-sm text-gray-600 hover:text-gray-900"
                      >
                        Copy Prefix
                      </button>
                      <button
                        onClick={() => revokeAPIKey(key.id)}
                        className="text-sm text-red-600 hover:text-red-900"
                      >
                        Revoke
                      </button>
                    </div>
                  </div>
                </li>
              ))}
            </ul>
          </div>
        )}

        {/* Documentation */}
        <div className="mt-8 bg-blue-50 rounded-lg p-6">
          <h3 className="text-lg font-medium text-blue-900 mb-4">Using Your API Key</h3>
          <div className="space-y-4 text-sm text-blue-800">
            <div>
              <p className="font-medium mb-2">Include the API key in your requests:</p>
              <div className="bg-blue-100 p-3 rounded font-mono text-xs">
                <p>Authorization: Bearer YOUR_API_KEY</p>
                <p className="mt-2">or</p>
                <p className="mt-2">X-API-Key: YOUR_API_KEY</p>
              </div>
            </div>
            <div>
              <p className="font-medium mb-2">Example curl command:</p>
              <div className="bg-blue-100 p-3 rounded font-mono text-xs">
                curl -H "Authorization: Bearer YOUR_API_KEY" \<br/>
                &nbsp;&nbsp;&nbsp;&nbsp;-H "Content-Type: application/json" \<br/>
                &nbsp;&nbsp;&nbsp;&nbsp;-d '{"input": "...", "output": "..."}' \<br/>
                &nbsp;&nbsp;&nbsp;&nbsp;http://localhost:8000/api/translate
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
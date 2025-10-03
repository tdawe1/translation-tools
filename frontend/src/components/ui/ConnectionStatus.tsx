'use client';

import React, { useState, useEffect } from 'react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Alert, AlertDescription } from '@/components/ui/alert';
import {
  Wifi,
  WifiOff,
  RefreshCw,
  AlertTriangle,
  CheckCircle
} from 'lucide-react';
import { translationWebSocket } from '@/lib/websocket';

export function ConnectionStatus() {
  const [isConnected, setIsConnected] = useState(false);
  const [reconnectAttempts, setReconnectAttempts] = useState(0);
  const [lastError, setLastError] = useState<string | null>(null);

  useEffect(() => {
    const handleConnection = (connected: boolean) => {
      setIsConnected(connected);
      if (connected) {
        setLastError(null);
      }
    };

    const handleError = (error: Error) => {
      setLastError(error.message);
    };

    translationWebSocket.onConnection(handleConnection);
    translationWebSocket.onError(handleError);

    // Initial connection check
    setIsConnected(translationWebSocket.connected);

    return () => {
      translationWebSocket.offConnection(handleConnection);
      translationWebSocket.onError(handleError);
    };
  }, []);

  const handleReconnect = () => {
    translationWebSocket.disconnect();
    setTimeout(() => translationWebSocket.connect(), 100);
  };

  const getStatusColor = () => {
    if (isConnected) return 'bg-green-500';
    return 'bg-red-500';
  };

  const getStatusText = () => {
    if (isConnected) return 'Connected';
    return 'Disconnected';
  };

  return (
    <div className="space-y-2">
      <div className="flex items-center space-x-2">
        <div className={`w-2 h-2 rounded-full ${getStatusColor()}`} />
        <span className="text-sm font-medium">{getStatusText()}</span>
        {!isConnected && (
          <Button
            variant="outline"
            size="sm"
            onClick={handleReconnect}
            className="ml-auto"
          >
            <RefreshCw className="h-3 w-3 mr-1" />
            Reconnect
          </Button>
        )}
      </div>

      {lastError && (
        <Alert variant="destructive" className="mt-2">
          <AlertTriangle className="h-4 w-4" />
          <AlertDescription className="text-xs">
            Connection error: {lastError}
          </AlertDescription>
        </Alert>
      )}

      {isConnected && (
        <div className="flex items-center space-x-2">
          <CheckCircle className="h-3 w-3 text-green-500" />
          <span className="text-xs text-gray-600">Real-time updates active</span>
        </div>
      )}
    </div>
  );
}
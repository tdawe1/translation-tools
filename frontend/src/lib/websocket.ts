/**
 * WebSocket service for real-time translation progress updates
 * Handles connection management, reconnection, and fallback to SSE
 */

export interface JobProgress {
  job_id: string;
  status: 'queued' | 'extracting' | 'translating' | 'applying' | 'finalizing' | 'completed' | 'failed';
  progress: number;
  stage: string;
  tokens_processed: number;
  total_tokens: number;
  current_cost: number;
  estimated_cost: number;
  quality_score?: number;
  error_message?: string;
  file_name: string;
  file_size: number;
  eta_seconds?: number;
  current_batch: number;
  total_batches: number;
}

export interface ProgressUpdate {
  type: 'job_started' | 'job_progress' | 'job_completed' | 'job_failed' | 'connection_established' | 'heartbeat' | 'pong';
  job_id?: string;
  timestamp?: string;
  [key: string]: any;
}

interface WebSocketOptions {
  url?: string;
  reconnectAttempts?: number;
  reconnectInterval?: number;
  heartbeatInterval?: number;
  enableSSEFallback?: boolean;
}

type EventHandler = (update: ProgressUpdate) => void;
type ConnectionHandler = (connected: boolean) => void;
type ErrorHandler = (error: Error) => void;

export class TranslationWebSocket {
  private ws: WebSocket | null = null;
  private eventSource: EventSource | null = null;
  private isConnected = false;
  private reconnectAttempts = 0;
  private maxReconnectAttempts = 5;
  private reconnectInterval = 2000;
  private heartbeatInterval = 30000;
  private heartbeatTimer: NodeJS.Timeout | null = null;
  private eventHandlers: Map<string, EventHandler[]> = new Map();
  private connectionHandlers: ConnectionHandler[] = [];
  private errorHandlers: ErrorHandler[] = [];
  private options: WebSocketOptions;

  constructor(options: WebSocketOptions = {}) {
    this.options = {
      url: 'ws://localhost:8081',
      reconnectAttempts: 5,
      reconnectInterval: 2000,
      heartbeatInterval: 30000,
      enableSSEFallback: true,
      ...options
    };

    this.maxReconnectAttempts = this.options.reconnectAttempts!;
    this.reconnectInterval = this.options.reconnectInterval!;
    this.heartbeatInterval = this.options.heartbeatInterval!;
  }

  connect() {
    if (this.isConnected || this.ws?.readyState === WebSocket.CONNECTING) {
      return;
    }

    try {
      this.ws = new WebSocket(this.options.url!);

      this.ws.onopen = () => {
        this.isConnected = true;
        this.reconnectAttempts = 0;
        this.startHeartbeat();
        this.notifyConnectionHandlers(true);
        console.log('WebSocket connected');
      };

      this.ws.onmessage = (event) => {
        try {
          const update: ProgressUpdate = JSON.parse(event.data);
          this.handleMessage(update);
        } catch (error) {
          console.error('Failed to parse WebSocket message:', error);
        }
      };

      this.ws.onclose = (event) => {
        this.isConnected = false;
        this.stopHeartbeat();
        this.notifyConnectionHandlers(false);

        if (!event.wasClean && this.reconnectAttempts < this.maxReconnectAttempts) {
          this.reconnect();
        } else if (this.options.enableSSEFallback) {
          this.fallbackToSSE();
        }
      };

      this.ws.onerror = (error) => {
        console.error('WebSocket error:', error);
        this.notifyErrorHandlers(new Error('WebSocket connection failed'));
      };
    } catch (error) {
      console.error('Failed to create WebSocket:', error);
      if (this.options.enableSSEFallback) {
        this.fallbackToSSE();
      }
    }
  }

  disconnect() {
    this.stopHeartbeat();

    if (this.ws) {
      this.ws.close();
      this.ws = null;
    }

    if (this.eventSource) {
      this.eventSource.close();
      this.eventSource = null;
    }

    this.isConnected = false;
    this.notifyConnectionHandlers(false);
  }

  private reconnect() {
    this.reconnectAttempts++;
    const delay = this.reconnectInterval * Math.pow(2, this.reconnectAttempts - 1);

    console.log(`Attempting to reconnect in ${delay}ms (attempt ${this.reconnectAttempts}/${this.maxReconnectAttempts})`);

    setTimeout(() => {
      this.connect();
    }, delay);
  }

  private fallbackToSSE() {
    console.log('Falling back to Server-Sent Events');

    try {
      const sseUrl = this.options.url!.replace('ws://', 'http://').replace('wss://', 'https://') + '/sse';
      this.eventSource = new EventSource(sseUrl);

      this.eventSource.onopen = () => {
        this.isConnected = true;
        this.notifyConnectionHandlers(true);
        console.log('SSE connection established');
      };

      this.eventSource.onmessage = (event) => {
        try {
          const update: ProgressUpdate = JSON.parse(event.data);
          this.handleMessage(update);
        } catch (error) {
          console.error('Failed to parse SSE message:', error);
        }
      };

      this.eventSource.onerror = () => {
        this.isConnected = false;
        this.notifyConnectionHandlers(false);
        console.error('SSE connection error');
      };
    } catch (error) {
      console.error('Failed to establish SSE connection:', error);
    }
  }

  private handleMessage(update: ProgressUpdate) {
    // Handle heartbeat
    if (update.type === 'heartbeat' || update.type === 'pong') {
      return;
    }

    // Emit to specific handlers
    const handlers = this.eventHandlers.get(update.type) || [];
    handlers.forEach(handler => handler(update));

    // Emit to all handlers
    const allHandlers = this.eventHandlers.get('*') || [];
    allHandlers.forEach(handler => handler(update));
  }

  private startHeartbeat() {
    this.stopHeartbeat();

    this.heartbeatTimer = setInterval(() => {
      if (this.isConnected && this.ws) {
        try {
          this.ws.send(JSON.stringify({ type: 'ping', timestamp: new Date().toISOString() }));
        } catch (error) {
          console.error('Failed to send heartbeat:', error);
        }
      }
    }, this.heartbeatInterval);
  }

  private stopHeartbeat() {
    if (this.heartbeatTimer) {
      clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
  }

  private notifyConnectionHandlers(connected: boolean) {
    this.connectionHandlers.forEach(handler => handler(connected));
  }

  private notifyErrorHandlers(error: Error) {
    this.errorHandlers.forEach(handler => handler(error));
  }

  // Public API
  on(event: string, handler: EventHandler) {
    if (!this.eventHandlers.has(event)) {
      this.eventHandlers.set(event, []);
    }
    this.eventHandlers.get(event)!.push(handler);
  }

  off(event: string, handler: EventHandler) {
    const handlers = this.eventHandlers.get(event);
    if (handlers) {
      const index = handlers.indexOf(handler);
      if (index > -1) {
        handlers.splice(index, 1);
      }
    }
  }

  onConnection(handler: ConnectionHandler) {
    this.connectionHandlers.push(handler);
  }

  onError(handler: ErrorHandler) {
    this.errorHandlers.push(handler);
  }

  subscribe(jobId: string) {
    if (this.isConnected && this.ws) {
      try {
        this.ws.send(JSON.stringify({ type: 'subscribe', job_id: jobId }));
      } catch (error) {
        console.error('Failed to subscribe to job:', error);
      }
    }
  }

  unsubscribe(jobId: string) {
    if (this.isConnected && this.ws) {
      try {
        this.ws.send(JSON.stringify({ type: 'unsubscribe', job_id: jobId }));
      } catch (error) {
        console.error('Failed to unsubscribe from job:', error);
      }
    }
  }

  getJobStatus(jobId: string) {
    if (this.isConnected && this.ws) {
      try {
        this.ws.send(JSON.stringify({ type: 'get_job_status', job_id: jobId }));
      } catch (error) {
        console.error('Failed to get job status:', error);
      }
    }
  }

  get connected() {
    return this.isConnected;
  }
}

// Singleton instance
export const translationWebSocket = new TranslationWebSocket();
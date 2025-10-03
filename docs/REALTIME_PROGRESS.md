# Real-time Progress Tracking System

This document describes the real-time progress tracking system for the translation pipeline.

## Overview

The system provides real-time updates for translation jobs through WebSocket connections with Server-Sent Events (SSE) as a fallback. This allows users to see live progress, cost tracking, and quality metrics as their documents are being translated.

## Architecture

### Components

1. **WebSocket Server** (`scripts/translation_websocket_server.py`)
   - Handles multiple client connections
   - Manages job subscriptions
   - Broadcasts progress updates
   - Provides heartbeat for connection health

2. **SSE Fallback Server** (`scripts/sse_fallback_server.py`)
   - HTTP-based alternative to WebSockets
   - Uses Server-Sent Events for real-time updates
   - Automatically used when WebSocket fails

3. **Progress Tracker** (`scripts/progress_tracker.py`)
   - Library for translation scripts to emit progress
   - Handles connection management
   - Calculates progress percentages and ETA

4. **Frontend WebSocket Service** (`frontend/src/lib/websocket.ts`)
   - Manages WebSocket/SSE connections
   - Handles reconnection with exponential backoff
   - Provides simple API for components

5. **UI Components**
   - `RealTimeProgress`: Main progress display
   - `ConnectionStatus`: Shows connection state

## Setup

### 1. Install Dependencies

```bash
pip install websockets aiohttp aiohttp-cors
```

### 2. Start the Real-time Servers

```bash
# Start both WebSocket and SSE servers
python scripts/start_realtime_servers.py
```

The servers will start on:
- WebSocket: `ws://localhost:8081`
- SSE: `http://localhost:8082/sse`

### 3. Update Translation Scripts

Import and use the progress tracker in your translation scripts:

```python
from progress_tracker import get_tracker

async def translate_with_progress():
    tracker = await get_tracker()
    await tracker.connect()

    # Start job
    await tracker.start_job(job_id, file_name, file_size, estimated_tokens, estimated_cost)

    # Update progress
    await tracker.update_progress(progress=50, stage="translating")

    # Complete job
    await tracker.complete_job(success=True)
```

### 4. Frontend Integration

The frontend components are already integrated. Just ensure:

1. The WebSocket service is imported
2. Components use the `RealTimeProgress` component
3. Connection status is displayed

## Features

### Real-time Updates

- **Progress Percentage**: Live progress from 0-100%
- **Stage Tracking**: Shows current stage (extracting, translating, applying, finalizing)
- **Batch Progress**: For multi-batch translations
- **ETA Calculation**: Dynamic estimated time remaining

### Cost Tracking

- **Live Cost Counter**: Shows current cost as it accumulates
- **Estimated Cost**: Initial cost estimate
- **Cost per Token**: Calculated in real-time

### Quality Metrics

- **Quality Score**: Final translation quality assessment
- **Live Updates**: Quality indicators during translation

### Connection Management

- **Automatic Reconnection**: Exponential backoff retry
- **Fallback to SSE**: When WebSocket fails
- **Connection Status**: Visual indicator of connection state
- **Heartbeat**: Keeps connections alive

## API Reference

### ProgressTracker Class

```python
class ProgressTracker:
    async def connect(self)
    async def disconnect(self)
    async def start_job(job_id, file_name, file_size, estimated_tokens, estimated_cost)
    async def update_progress(**kwargs)
    async def update_stage(stage, status=None)
    async def update_tokens(tokens_processed, cost_increment=0.0)
    async def update_batch_progress(current_batch, total_batches)
    async def set_quality_score(score)
    async def complete_job(success=True, error_message=None)
```

### TranslationWebSocket Class (Frontend)

```typescript
class TranslationWebSocket {
  connect(): void
  disconnect(): void
  on(event: string, handler: EventHandler): void
  off(event: string, handler: EventHandler): void
  subscribe(jobId: string): void
  unsubscribe(jobId: string): void
  get connected: boolean
}
```

## Events

### Server to Client

- `connection_established`: Initial connection
- `job_started`: New translation job started
- `job_progress`: Progress update
- `job_completed`: Job finished successfully
- `job_failed`: Job failed with error
- `heartbeat`: Keep-alive ping

### Client to Server

- `subscribe`: Subscribe to job updates
- `unsubscribe`: Unsubscribe from job
- `ping`: Heartbeat response
- `get_job_status`: Request current status

## Demo

Run the demo to see the system in action:

```bash
# 1. Start servers
python scripts/start_realtime_servers.py

# 2. In another terminal, run demo
python scripts/demo_realtime_translation.py

# 3. Open frontend to see real-time updates
cd frontend && npm run dev
```

## Troubleshooting

### Connection Issues

1. **Check server status**: Ensure WebSocket server is running on port 8081
2. **Firewall**: Allow WebSocket connections (port 8081) and HTTP (port 8082)
3. **CORS**: Ensure servers allow cross-origin requests

### High CPU Usage

1. **Reduce update frequency**: Progress updates are throttled
2. **Batch updates**: Multiple progress updates are combined
3. **Client limits**: Each client receives only relevant updates

### Memory Issues

1. **Clean up**: Unsubscribe from completed jobs
2. **Limit history**: Activity feed keeps only recent events
3. **Connection pooling**: Reuse connections when possible

## Performance Considerations

- Update frequency is limited to prevent overwhelming clients
- Only subscribed clients receive updates
- Progress calculations are cached and reused
- Connection pooling reduces overhead

## Security

- All connections use WebSocket secure (wss://) in production
- CORS configured to allow only trusted origins
- Job IDs are validated before subscription
- No sensitive data transmitted in progress updates
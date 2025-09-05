import os, json, asyncio
from fastapi import FastAPI, Request, HTTPException
from openai import OpenAI

app = FastAPI()

# Required:
#   OPENAI_API_KEY=sk-...
#   OPENAI_WEBHOOK_SECRET=whsec_...   (from dashboard Webhooks page)
client = OpenAI()  # picks up both env vars

# In-memory idempotency store for demo; replace with Redis in prod
SEEN = set()

# Map response.id -> your context (persist this when you dispatch requests)
# In your translator, write to this map (or Redis) when calling responses.create(...)
RESPONSE_INDEX = {}  # e.g., {"resp_abc123": {"slide": 8, "shape": 3, "para": 0, "batch_id": "run_2025-09-05"}}

@app.post("/webhooks/openai")
async def openai_webhook(request: Request):
    raw = await request.body()
    try:
        # Verifies signature + returns a typed Event object
        event = client.webhooks.unwrap(raw.decode("utf-8"), request.headers)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"invalid signature: {e}")

    # Basic idempotency (OpenAI may retry)
    event_id = getattr(event, "id", None) or hash(raw)
    if event_id in SEEN:
        return {"ok": True, "deduped": True}
    SEEN.add(event_id)

    # Handle key events
    if event.type == "response.completed":
        rid = event.data.id
        ctx = RESPONSE_INDEX.get(rid, {})
        # TODO: increment progress, mark block completed, write to a CSV/Redis log, etc.
        print("✅ response.completed", {"response_id": rid, **ctx})
    elif event.type == "response.failed":
        rid = getattr(event.data, "id", None)
        ctx = RESPONSE_INDEX.get(rid or "", {})
        print("❌ response.failed", {"response_id": rid, **ctx, "error": event.data})
        # TODO: mark failure; optionally queue a retry with a fallback model
    else:
        print("ℹ️ unhandled event", event.type)

    # Respond fast; do heavy work async in a task/queue if needed
    return {"ok": True}
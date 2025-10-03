from fastapi import HTTPException, status, Request
from fastapi.security.utils import get_authorization_scheme_param
from typing import Dict, Optional
import time
import redis
from collections import defaultdict, deque

from ..core.config import settings

class RateLimiter:
    """Rate limiting middleware using Redis or in-memory storage"""

    def __init__(self, redis_url: Optional[str] = None):
        self.redis_url = redis_url
        self.redis_client = None
        self.in_memory_store: Dict[str, deque] = defaultdict(deque)

        # Initialize Redis if URL provided
        if redis_url:
            try:
                self.redis_client = redis.from_url(redis_url)
            except:
                # Fall back to in-memory storage
                self.redis_client = None

    async def is_rate_limited(self, key: str, limit: int, window: int) -> bool:
        """Check if the key has exceeded rate limit"""
        if self.redis_client:
            return await self._check_redis_rate_limit(key, limit, window)
        else:
            return self._check_memory_rate_limit(key, limit, window)

    async def _check_redis_rate_limit(self, key: str, limit: int, window: int) -> bool:
        """Check rate limit using Redis"""
        current_time = time.time()
        window_start = current_time - window

        # Remove old entries
        await self.redis_client.zremrangebyscore(key, 0, window_start)

        # Get current count
        current_count = await self.redis_client.zcard(key)

        if current_count >= limit:
            return True

        # Add new entry
        await self.redis_client.zadd(key, {str(current_time): current_time})
        await self.redis_client.expire(key, window)

        return False

    def _check_memory_rate_limit(self, key: str, limit: int, window: int) -> bool:
        """Check rate limit using in-memory storage"""
        current_time = time.time()
        window_start = current_time - window

        # Remove old entries
        while self.in_memory_store[key] and self.in_memory_store[key][0] < window_start:
            self.in_memory_store[key].popleft()

        # Check limit
        if len(self.in_memory_store[key]) >= limit:
            return True

        # Add new entry
        self.in_memory_store[key].append(current_time)

        return False

async def rate_limit_middleware(
    request: Request,
    call_next,
    limiter: RateLimiter = None
):
    """Rate limiting middleware"""
    if limiter is None:
        limiter = RateLimiter(settings.REDIS_URL)

    # Get client identifier
    client_ip = request.client.host
    forwarded_for = request.headers.get("X-Forwarded-For")
    if forwarded_for:
        client_ip = forwarded_for.split(",")[0].strip()

    # Get user identifier if authenticated
    user_id = None
    auth_header = request.headers.get("Authorization")
    if auth_header:
        try:
            scheme, credentials = get_authorization_scheme_param(auth_header)
            if scheme.lower() == "bearer":
                # In a real app, you'd decode the JWT to get user ID
                # For now, use the credentials as a unique identifier
                user_id = credentials
        except:
            pass

    # API key check
    api_key = request.headers.get("X-API-Key")
    if api_key:
        # Use API key as identifier
        rate_key = f"api_key:{api_key}"
    elif user_id:
        # Use user ID as identifier
        rate_key = f"user:{user_id}"
    else:
        # Use IP address as identifier
        rate_key = f"ip:{client_ip}"

    # Check rate limit
    is_limited = await limiter.is_rate_limited(
        rate_key,
        settings.RATE_LIMIT_REQUESTS,
        settings.RATE_LIMIT_WINDOW
    )

    if is_limited:
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail="Rate limit exceeded",
            headers={
                "X-RateLimit-Limit": str(settings.RATE_LIMIT_REQUESTS),
                "X-RateLimit-Window": str(settings.RATE_LIMIT_WINDOW),
                "X-RateLimit-Remaining": "0",
                "Retry-After": str(settings.RATE_LIMIT_WINDOW)
            }
        )

    response = await call_next(request)

    # Add rate limit headers
    response.headers["X-RateLimit-Limit"] = str(settings.RATE_LIMIT_REQUESTS)
    response.headers["X-RateLimit-Window"] = str(settings.RATE_LIMIT_WINDOW)

    return response
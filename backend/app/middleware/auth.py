from fastapi import HTTPException, status, Depends, Request
from fastapi.security import HTTPBearer, APIKeyHeader
from fastapi.security.utils import get_authorization_scheme_param
from typing import Optional
from sqlalchemy.orm import Session

from ..services.auth_service import AuthService
from ..database.session import get_db

security = HTTPBearer()
api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)
auth_service = AuthService()

async def get_current_user(
    request: Request,
    token: Optional[str] = Depends(security),
    api_key: Optional[str] = Depends(api_key_header),
    db: Session = Depends(get_db)
):
    """Get current authenticated user from JWT token or API key"""
    # Try API key first
    if api_key:
        user_id = auth_service.verify_api_key(db, api_key)
        if user_id:
            # Store authentication method in request state
            request.state.auth_method = "api_key"
            return user_id

    # Then try JWT token
    if token and token.credentials:
        try:
            user_id = auth_service.verify_token(token.credentials)
            request.state.auth_method = "jwt"
            return user_id
        except HTTPException:
            pass

    raise HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )

async def get_current_active_user(
    user_id: str = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """Get current active user"""
    user = auth_service.get_user_by_id(db, user_id)
    if not user or not user.is_active:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Inactive user"
        )
    return user_id

class CSRFMiddleware:
    """CSRF protection middleware"""

    def __init__(self, secret_key: str):
        self.secret_key = secret_key
        self.csrf_exempt_paths = {
            "/api/auth/login",
            "/api/auth/register",
            "/api/auth/google/callback",
            "/health",
            "/docs",
            "/openapi.json",
            "/redoc"
        }

    async def __call__(self, request: Request, call_next):
        # Skip CSRF check for exempt paths and safe methods
        if (
            request.url.path in self.csrf_exempt_paths or
            request.method in {"GET", "HEAD", "OPTIONS", "TRACE"}
        ):
            return await call_next(request)

        # Check CSRF token for state-changing requests
        csrf_token = request.headers.get("X-CSRF-Token")
        if not csrf_token:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="CSRF token missing"
            )

        # In a real implementation, you would validate the token
        # For now, we'll just check if it exists
        return await call_next(request)
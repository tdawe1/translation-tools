from fastapi import APIRouter, Depends, HTTPException, status, Query
from fastapi.security import HTTPBearer, APIKeyHeader
from datetime import timedelta
from sqlalchemy.orm import Session

from ..models.auth import (
    UserLogin, UserCreate, User, Token, APIKeyCreate, APIKey,
    GoogleAuthRequest, GoogleAuthResponse
)
from ..services.auth_service import AuthService
from ..services.oauth_service import OAuthService
from ..database.session import get_db
from ..core.config import settings

router = APIRouter()
security = HTTPBearer()
api_key_header = APIKeyHeader(name="X-API-Key", auto_error=False)
auth_service = AuthService()
oauth_service = OAuthService()

@router.post("/register", response_model=User)
async def register(
    user_data: UserCreate,
    db: Session = Depends(get_db)
):
    """Register a new user"""
    try:
        user = auth_service.create_user(db, user_data)
        return user
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to create user"
        )

@router.post("/login", response_model=Token)
async def login(
    user_credentials: UserLogin,
    db: Session = Depends(get_db)
):
    """Authenticate user and return access token"""
    user = auth_service.authenticate_user(
        db,
        user_credentials.email,
        user_credentials.password
    )

    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
            headers={"WWW-Authenticate": "Bearer"},
        )

    if not user.is_active:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Inactive user"
        )

    # Create access token
    access_token_expires = timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = auth_service.create_access_token(
        user_id=user.id,
        expires_delta=access_token_expires
    )

    # Create refresh token
    refresh_token = auth_service.create_refresh_token(db, user.id)

    return {
        "access_token": access_token,
        "token_type": "bearer",
        "expires_in": settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60,
        "refresh_token": refresh_token
    }

@router.post("/refresh", response_model=Token)
async def refresh_token(
    refresh_token: str,
    db: Session = Depends(get_db)
):
    """Refresh access token"""
    try:
        access_token = auth_service.refresh_access_token(db, refresh_token)
        return {
            "access_token": access_token,
            "token_type": "bearer",
            "expires_in": settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60
        }
    except HTTPException:
        raise

@router.post("/logout")
async def logout(
    refresh_token: str,
    db: Session = Depends(get_db)
):
    """Logout user by revoking refresh token"""
    auth_service.revoke_refresh_token(db, refresh_token)
    return {"message": "Successfully logged out"}

@router.get("/me", response_model=User)
async def get_current_user(
    token: str = Depends(security),
    db: Session = Depends(get_db)
):
    """Get current user information"""
    user_id = auth_service.verify_token(token.credentials)
    user = auth_service.get_user_by_id(db, user_id)

    if not user:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="User not found"
        )

    return user

@router.get("/google/auth-url")
async def get_google_auth_url(
    redirect_uri: str,
    state: Optional[str] = None
):
    """Get Google OAuth authorization URL"""
    auth_url = oauth_service.get_google_auth_url(redirect_uri, state)
    return {"auth_url": auth_url}

@router.post("/google/callback", response_model=GoogleAuthResponse)
async def google_callback(
    auth_data: GoogleAuthRequest,
    db: Session = Depends(get_db)
):
    """Handle Google OAuth callback"""
    try:
        # Exchange code for tokens
        token_data = await oauth_service.exchange_google_code(
            auth_data.code,
            auth_data.redirect_uri
        )

        # Get user info
        user_info = await oauth_service.get_google_user_info(token_data["access_token"])

        # Create or update user
        user = auth_service.create_or_update_google_user(
            db=db,
            google_id=user_info["id"],
            email=user_info["email"],
            full_name=user_info.get("name", ""),
            access_token=token_data["access_token"],
            refresh_token=token_data.get("refresh_token")
        )

        # Create our access token
        access_token_expires = timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = auth_service.create_access_token(
            user_id=user.id,
            expires_delta=access_token_expires
        )

        # Create refresh token
        refresh_token = auth_service.create_refresh_token(db, user.id)

        return GoogleAuthResponse(
            access_token=access_token,
            token_type="bearer",
            expires_in=settings.ACCESS_TOKEN_EXPIRE_MINUTES * 60,
            refresh_token=refresh_token,
            user=user
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to authenticate with Google"
        )

@router.post("/api-keys", response_model=APIKey)
async def create_api_key(
    key_data: APIKeyCreate,
    db: Session = Depends(get_db),
    token: str = Depends(security)
):
    """Create a new API key"""
    user_id = auth_service.verify_token(token.credentials)
    api_key = auth_service.create_api_key(
        db=db,
        user_id=user_id,
        name=key_data.name,
        expires_days=key_data.expires_days
    )

    # Return API key info
    db_key = db.query(APIKey.__fields__["__tablename__"]).filter(
        APIKey.__fields__["key"] == api_key
    ).first()
    return APIKey.from_orm(db_key)

@router.get("/api-keys", response_model=list[APIKey])
async def list_api_keys(
    db: Session = Depends(get_db),
    token: str = Depends(security)
):
    """List user's API keys"""
    user_id = auth_service.verify_token(token.credentials)
    keys = auth_service.list_user_api_keys(db, user_id)
    return [APIKey.from_orm(key) for key in keys]

@router.delete("/api-keys/{key_id}")
async def revoke_api_key(
    key_id: str,
    db: Session = Depends(get_db),
    token: str = Depends(security)
):
    """Revoke an API key"""
    user_id = auth_service.verify_token(token.credentials)
    success = auth_service.revoke_api_key(db, key_id, user_id)

    if not success:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="API key not found"
        )

    return {"message": "API key revoked successfully"}

async def get_api_user(
    api_key: str = Depends(api_key_header),
    db: Session = Depends(get_db)
):
    """Get user from API key"""
    if not api_key:
        return None

    user_id = auth_service.verify_api_key(db, api_key)
    if user_id:
        return user_id
    return None
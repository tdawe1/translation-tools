from pydantic import BaseModel, EmailStr, Field
from typing import Optional
from datetime import datetime

class UserBase(BaseModel):
    email: EmailStr = Field(..., description="User email address")
    full_name: Optional[str] = Field(None, description="User full name")

class UserCreate(UserBase):
    password: str = Field(..., min_length=8, description="User password")

class UserLogin(BaseModel):
    email: EmailStr = Field(..., description="User email")
    password: str = Field(..., description="User password")

class User(UserBase):
    id: str = Field(..., description="User ID")
    is_active: bool = Field(default=True, description="Is user active")
    created_at: datetime = Field(..., description="User creation timestamp")

    class Config:
        from_attributes = True

class Token(BaseModel):
    access_token: str = Field(..., description="JWT access token")
    token_type: str = Field(default="bearer", description="Token type")
    expires_in: int = Field(..., description="Token expiration time in seconds")
    refresh_token: Optional[str] = Field(None, description="Refresh token")

class TokenData(BaseModel):
    user_id: Optional[str] = Field(None, description="User ID from token")

class APIKeyCreate(BaseModel):
    name: str = Field(..., min_length=1, max_length=100, description="API key name")
    expires_days: Optional[int] = Field(None, gt=0, description="Days until expiration")

class APIKey(BaseModel):
    id: str = Field(..., description="API key ID")
    name: str = Field(..., description="API key name")
    key: str = Field(..., description="API key value")
    is_active: bool = Field(..., description="Is API key active")
    last_used: Optional[datetime] = Field(None, description="Last used timestamp")
    created_at: datetime = Field(..., description="Creation timestamp")
    expires_at: Optional[datetime] = Field(None, description="Expiration timestamp")

    class Config:
        from_attributes = True

class GoogleAuthRequest(BaseModel):
    code: str = Field(..., description="OAuth authorization code")
    redirect_uri: str = Field(..., description="Redirect URI used in authorization")

class GoogleAuthResponse(BaseModel):
    access_token: str = Field(..., description="JWT access token")
    token_type: str = Field(default="bearer", description="Token type")
    expires_in: int = Field(..., description="Token expiration time in seconds")
    refresh_token: Optional[str] = Field(None, description="Refresh token")
    user: User = Field(..., description="User information")
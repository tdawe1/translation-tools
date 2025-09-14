from datetime import datetime, timedelta
from typing import Optional
from jose import JWTError, jwt
from passlib.context import CryptContext
from fastapi import HTTPException, status, Depends
from sqlalchemy.orm import Session
import uuid
import secrets

from ..models.auth import User, UserCreate, TokenData
from ..database.session import get_db
from ..database.database import User as DBUser, RefreshToken, APIKey
from ..core.config import settings

# Password hashing
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

class AuthService:
    def __init__(self):
        pass

    def verify_password(self, plain_password: str, hashed_password: str) -> bool:
        """Verify a password against its hash"""
        return pwd_context.verify(plain_password, hashed_password)

    def get_password_hash(self, password: str) -> str:
        """Generate password hash"""
        return pwd_context.hash(password)

    def create_user(self, db: Session, user: UserCreate) -> User:
        """Create a new user"""
        # Check if user already exists
        existing_user = db.query(DBUser).filter(DBUser.email == user.email).first()
        if existing_user:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Email already registered"
            )

        # Create user
        hashed_password = self.get_password_hash(user.password)

        db_user = DBUser(
            email=user.email,
            full_name=user.full_name,
            hashed_password=hashed_password,
            is_active=True,
            created_at=datetime.utcnow()
        )

        db.add(db_user)
        db.commit()
        db.refresh(db_user)

        return User(
            id=db_user.id,
            email=db_user.email,
            full_name=db_user.full_name,
            is_active=db_user.is_active,
            created_at=db_user.created_at
        )

    def authenticate_user(self, db: Session, email: str, password: str) -> Optional[User]:
        """Authenticate a user"""
        user = db.query(DBUser).filter(DBUser.email == email).first()

        if not user or not self.verify_password(password, user.hashed_password):
            return None

        return User(
            id=user.id,
            email=user.email,
            full_name=user.full_name,
            is_active=user.is_active,
            created_at=user.created_at
        )

    def create_access_token(self, user_id: str, expires_delta: Optional[timedelta] = None) -> str:
        """Create JWT access token"""
        if expires_delta:
            expire = datetime.utcnow() + expires_delta
        else:
            expire = datetime.utcnow() + timedelta(minutes=settings.ACCESS_TOKEN_EXPIRE_MINUTES)

        to_encode = {
            "sub": user_id,
            "exp": expire,
            "type": "access"
        }
        encoded_jwt = jwt.encode(to_encode, settings.SECRET_KEY, algorithm="HS256")
        return encoded_jwt

    def create_refresh_token(self, db: Session, user_id: str) -> str:
        """Create refresh token"""
        refresh_token = secrets.token_urlsafe(32)
        expire = datetime.utcnow() + timedelta(days=settings.REFRESH_TOKEN_EXPIRE_DAYS)

        db_refresh_token = RefreshToken(
            token=refresh_token,
            user_id=user_id,
            expires_at=expire
        )

        db.add(db_refresh_token)
        db.commit()

        return refresh_token

    def verify_token(self, token: str) -> str:
        """Verify JWT token and return user ID"""
        try:
            payload = jwt.decode(token, settings.SECRET_KEY, algorithms=["HS256"])
            user_id: str = payload.get("sub")
            if user_id is None:
                raise HTTPException(
                    status_code=status.HTTP_401_UNAUTHORIZED,
                    detail="Invalid authentication credentials",
                    headers={"WWW-Authenticate": "Bearer"},
                )
            return user_id
        except JWTError:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid authentication credentials",
                headers={"WWW-Authenticate": "Bearer"},
            )

    def refresh_access_token(self, db: Session, refresh_token: str) -> str:
        """Refresh access token using refresh token"""
        db_token = db.query(RefreshToken).filter(
            RefreshToken.token == refresh_token,
            RefreshToken.is_revoked == False
        ).first()

        if not db_token:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid refresh token"
            )

        # Check if refresh token is expired
        if datetime.utcnow() > db_token.expires_at:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Refresh token expired"
            )

        # Create new access token
        access_token = self.create_access_token(db_token.user_id)

        return access_token

    def revoke_refresh_token(self, db: Session, refresh_token: str):
        """Revoke a refresh token"""
        db_token = db.query(RefreshToken).filter(
            RefreshToken.token == refresh_token
        ).first()

        if db_token:
            db_token.is_revoked = True
            db.commit()

    def get_user_by_id(self, db: Session, user_id: str) -> Optional[User]:
        """Get user by ID"""
        db_user = db.query(DBUser).filter(DBUser.id == user_id).first()
        if db_user:
            return User(
                id=db_user.id,
                email=db_user.email,
                full_name=db_user.full_name,
                is_active=db_user.is_active,
                created_at=db_user.created_at
            )
        return None

    def create_api_key(self, db: Session, user_id: str, name: str, expires_days: Optional[int] = None) -> str:
        """Create a new API key for a user"""
        # Generate secure API key
        api_key = settings.API_KEY_PREFIX + secrets.token_urlsafe(settings.API_KEY_LENGTH)

        # Calculate expiration
        expires_at = None
        if expires_days:
            expires_at = datetime.utcnow() + timedelta(days=expires_days)

        db_api_key = APIKey(
            key=api_key,
            name=name,
            user_id=user_id,
            expires_at=expires_at
        )

        db.add(db_api_key)
        db.commit()

        return api_key

    def verify_api_key(self, db: Session, api_key: str) -> Optional[str]:
        """Verify API key and return user ID"""
        if not api_key.startswith(settings.API_KEY_PREFIX):
            return None

        db_key = db.query(APIKey).filter(
            APIKey.key == api_key,
            APIKey.is_active == True
        ).first()

        if not db_key:
            return None

        # Check expiration
        if db_key.expires_at and datetime.utcnow() > db_key.expires_at:
            return None

        # Update last used
        db_key.last_used = datetime.utcnow()
        db.commit()

        return db_key.user_id

    def list_user_api_keys(self, db: Session, user_id: str) -> list:
        """List all API keys for a user"""
        return db.query(APIKey).filter(APIKey.user_id == user_id).all()

    def revoke_api_key(self, db: Session, api_key_id: str, user_id: str) -> bool:
        """Revoke an API key"""
        db_key = db.query(APIKey).filter(
            APIKey.id == api_key_id,
            APIKey.user_id == user_id
        ).first()

        if db_key:
            db_key.is_active = False
            db.commit()
            return True
        return False

    def create_or_update_google_user(self, db: Session, google_id: str, email: str, full_name: str,
                                   access_token: str, refresh_token: Optional[str] = None) -> User:
        """Create or update user from Google OAuth"""
        # Check if user exists by Google ID
        user = db.query(DBUser).filter(DBUser.google_id == google_id).first()

        if user:
            # Update existing user
            user.google_access_token = access_token
            if refresh_token:
                user.google_refresh_token = refresh_token
            user.is_verified = True
            db.commit()
        else:
            # Check if user exists by email
            user = db.query(DBUser).filter(DBUser.email == email).first()

            if user:
                # Link Google account to existing user
                user.google_id = google_id
                user.google_access_token = access_token
                if refresh_token:
                    user.google_refresh_token = refresh_token
                user.is_verified = True
                db.commit()
            else:
                # Create new user
                user = DBUser(
                    email=email,
                    full_name=full_name,
                    google_id=google_id,
                    google_access_token=access_token,
                    google_refresh_token=refresh_token,
                    is_verified=True,
                    is_active=True,
                    hashed_password="",  # No password for OAuth users
                )
                db.add(user)
                db.commit()
                db.refresh(user)

        return User(
            id=user.id,
            email=user.email,
            full_name=user.full_name,
            is_active=user.is_active,
            created_at=user.created_at
        )
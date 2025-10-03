import requests
from fastapi import HTTPException, status
from typing import Optional, Dict, Any
import json

from ..core.config import settings

class OAuthService:
    def __init__(self):
        self.google_client_id = settings.GOOGLE_CLIENT_ID
        self.google_client_secret = settings.GOOGLE_CLIENT_SECRET
        self.google_auth_url = "https://accounts.google.com/o/oauth2/v2/auth"
        self.google_token_url = "https://oauth2.googleapis.com/token"
        self.google_userinfo_url = "https://www.googleapis.com/oauth2/v2/userinfo"
        self.google_scopes = [
            "openid",
            "email",
            "profile",
            "https://www.googleapis.com/auth/drive.readonly"
        ]

    def get_google_auth_url(self, redirect_uri: str, state: Optional[str] = None) -> str:
        """Generate Google OAuth authorization URL"""
        params = {
            "client_id": self.google_client_id,
            "redirect_uri": redirect_uri,
            "scope": " ".join(self.scopes),
            "response_type": "code",
            "access_type": "offline",
            "prompt": "consent"
        }

        if state:
            params["state"] = state

        return f"{self.google_auth_url}?{'&'.join(f'{k}={v}' for k, v in params.items())}"

    async def exchange_google_code(self, code: str, redirect_uri: str) -> Dict[str, Any]:
        """Exchange authorization code for access token"""
        data = {
            "client_id": self.google_client_id,
            "client_secret": self.google_client_secret,
            "code": code,
            "redirect_uri": redirect_uri,
            "grant_type": "authorization_code"
        }

        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }

        response = requests.post(self.google_token_url, data=data, headers=headers)

        if response.status_code != 200:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Failed to exchange authorization code"
            )

        return response.json()

    async def get_google_user_info(self, access_token: str) -> Dict[str, Any]:
        """Get user information from Google"""
        headers = {
            "Authorization": f"Bearer {access_token}"
        }

        response = requests.get(self.google_userinfo_url, headers=headers)

        if response.status_code != 200:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Failed to get user information"
            )

        return response.json()

    async def get_google_drive_access_token(self, refresh_token: str) -> Optional[str]:
        """Refresh Google Drive access token if needed"""
        if not refresh_token:
            return None

        data = {
            "client_id": self.google_client_id,
            "client_secret": self.google_client_secret,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token"
        }

        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }

        response = requests.post(self.google_token_url, data=data, headers=headers)

        if response.status_code != 200:
            return None

        token_data = response.json()
        return token_data.get("access_token")
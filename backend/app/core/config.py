"""
Configuration settings for the FastAPI app.
"""

import os
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    """Application settings."""

    # App settings
    app_name: str = "DOCX Translation API"
    app_version: str = "0.1.0"
    debug: bool = False

    # API settings
    max_file_size: int = 50 * 1024 * 1024  # 50MB
    supported_formats: list = [".docx"]

    # Security
    secret_key: str = "your-secret-key-here"

    # OpenAI
    openai_api_key: str = ""

    class Config:
        env_file = ".env"

# Create settings instance
settings = Settings()
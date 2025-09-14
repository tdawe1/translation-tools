import os
from typing import List
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    # Application settings
    APP_NAME: str = "Translation Pipeline API"
    DEBUG: bool = False
    VERSION: str = "1.0.0"

    # API settings
    API_PREFIX: str = "/api"
    SECRET_KEY: str = "your-secret-key-change-this-in-production"

    # CORS settings
    ALLOWED_ORIGINS: List[str] = ["http://localhost:3000", "http://localhost:3001"]

    # File storage
    UPLOAD_DIR: str = "uploads"
    OUTPUT_DIR: str = "outputs"
    MAX_FILE_SIZE: int = 100 * 1024 * 1024  # 100MB

    # Job settings
    JOB_TIMEOUT: int = 3600  # 1 hour
    MAX_CONCURRENT_JOBS: int = 5

    # OpenAI settings
    OPENAI_API_KEY: str = ""
    OPENAI_MODEL: str = "gpt-4o-2024-08-06"
    OPENAI_TEMPERATURE: float = 0.6

    # Authentication
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    # OAuth2 Google
    GOOGLE_CLIENT_ID: str = ""
    GOOGLE_CLIENT_SECRET: str = ""

    # API Keys
    API_KEY_PREFIX: str = "tr_"
    API_KEY_LENGTH: int = 32

    # Rate limiting
    RATE_LIMIT_REQUESTS: int = 100
    RATE_LIMIT_WINDOW: int = 60  # seconds

    # Redis (for job queue)
    REDIS_URL: str = "redis://localhost:6379"

    # Paths to translation scripts
    SCRIPTS_DIR: str = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "scripts")

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()
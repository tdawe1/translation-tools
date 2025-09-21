import os
import warnings
from typing import List, Optional
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """
    Application configuration with environment variable support.

    This configuration class manages all application settings with proper
    fallback mechanisms for development and strict validation for production.

    Environment Variables:
    ------------------------
    Core Settings:
        DEBUG: bool (default: False)
            Enable debug mode with development defaults
        SECRET_KEY: str (required in production)
            Secret key for JWT tokens and CSRF protection (32+ characters)
        APP_NAME: str (default: "Translation Pipeline API")
            Application name
        VERSION: str (default: "1.0.0")
            Application version

    API Settings:
        API_PREFIX: str (default: "/api")
            URL prefix for all API endpoints
        ALLOWED_ORIGINS: List[str] (default: ["http://localhost:3000", "http://localhost:3001"])
            CORS allowed origins for frontend applications

    File Storage:
        UPLOAD_DIR: str (default: "uploads")
            Directory for uploaded files
        OUTPUT_DIR: str (default: "outputs")
            Directory for translated files
        MAX_FILE_SIZE: int (default: 104857600)
            Maximum file upload size in bytes (100MB)

    Job Processing:
        JOB_TIMEOUT: int (default: 3600)
            Job timeout in seconds (1 hour)
        MAX_CONCURRENT_JOBS: int (default: 5)
            Maximum concurrent translation jobs

    OpenAI Integration:
        OPENAI_API_KEY: str (required for translation)
            OpenAI API key for translation services
        OPENAI_MODEL: str (default: "gpt-4o-2024-08-06")
            Default model for translation
        OPENAI_TEMPERATURE: float (default: 0.6)
            Translation temperature (0.0-1.0)

    Authentication:
        ACCESS_TOKEN_EXPIRE_MINUTES: int (default: 30)
            JWT access token expiration in minutes
        REFRESH_TOKEN_EXPIRE_DAYS: int (default: 7)
            JWT refresh token expiration in days

    OAuth2 (Google):
        GOOGLE_CLIENT_ID: str (optional)
            Google OAuth2 client ID
        GOOGLE_CLIENT_SECRET: str (optional)
            Google OAuth2 client secret

    API Keys:
        API_KEY_PREFIX: str (default: "tr_")
            Prefix for generated API keys
        API_KEY_LENGTH: int (default: 32)
            Length of generated API keys

    Rate Limiting:
        RATE_LIMIT_REQUESTS: int (default: 100)
            Maximum requests per window
        RATE_LIMIT_WINDOW: int (default: 60)
            Rate limiting window in seconds

    Infrastructure:
        REDIS_URL: str (default: "redis://localhost:6379")
            Redis connection URL for job queue

    Development Defaults:
    -------------------
    In DEBUG mode (DEBUG=True), the following stable defaults are applied:
    - SECRET_KEY: "dev-secret-key-32-characters-long-for-local-dev"
    - OPENAI_API_KEY: "debug-placeholder-key-invalid-for-production"

    These defaults are stable across restarts and should NEVER be used in production.
    For testing environments (pytest), configuration is loaded from .env.test file.
    """
    # Application settings
    APP_NAME: str = "Translation Pipeline API"
    DEBUG: bool = False
    VERSION: str = "1.0.0"

    # API settings
    API_PREFIX: str = "/api"
    # SECRET_KEY: Required for production, auto-generated in DEBUG mode
    # Override with: SECRET_KEY=your-secure-32+character-key
    SECRET_KEY: str = ""

    # CORS settings
    # Override with: ALLOWED_ORIGINS=["https://yourdomain.com", "https://app.yourdomain.com"]
    ALLOWED_ORIGINS: List[str] = ["http://localhost:3000", "http://localhost:3001"]

    # File storage
    # Override with: UPLOAD_DIR=/path/to/uploads
    UPLOAD_DIR: str = "uploads"
    # Override with: OUTPUT_DIR=/path/to/outputs
    OUTPUT_DIR: str = "outputs"
    # Override with: MAX_FILE_SIZE=524288000 (500MB)
    MAX_FILE_SIZE: int = 100 * 1024 * 1024  # 100MB

    # Job settings
    # Override with: JOB_TIMEOUT=7200 (2 hours)
    JOB_TIMEOUT: int = 3600  # 1 hour
    # Override with: MAX_CONCURRENT_JOBS=10
    MAX_CONCURRENT_JOBS: int = 5

    # OpenAI settings
    # Required for production. Override with: OPENAI_API_KEY=sk-your-actual-key
    OPENAI_API_KEY: str = ""
    # Override with: OPENAI_MODEL=gpt-4o-mini
    OPENAI_MODEL: str = "gpt-4o-2024-08-06"
    # Override with: OPENAI_TEMPERATURE=0.3
    OPENAI_TEMPERATURE: float = 0.6

    # Authentication
    # Override with: ACCESS_TOKEN_EXPIRE_MINUTES=60
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    # Override with: REFRESH_TOKEN_EXPIRE_DAYS=30
    REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    # OAuth2 Google (optional)
    # Override with: GOOGLE_CLIENT_ID=your-client-id
    GOOGLE_CLIENT_ID: str = ""
    # Override with: GOOGLE_CLIENT_SECRET=your-client-secret
    GOOGLE_CLIENT_SECRET: str = ""

    # API Keys
    # Override with: API_KEY_PREFIX=custom_
    API_KEY_PREFIX: str = "tr_"
    # Override with: API_KEY_LENGTH=64
    API_KEY_LENGTH: int = 32

    # Rate limiting
    # Override with: RATE_LIMIT_REQUESTS=1000
    RATE_LIMIT_REQUESTS: int = 100
    # Override with: RATE_LIMIT_WINDOW=3600 (1 hour)
    RATE_LIMIT_WINDOW: int = 60  # seconds

    # Redis (for job queue)
    # Override with: REDIS_URL=redis://user:pass@host:port/db
    REDIS_URL: str = "redis://localhost:6379"

    # Database URL
    DATABASE_URL: str = "sqlite:///./translation_pipeline.db"

    # Feature flags
    ENABLE_STYLE_CHECKING: bool = True
    ENABLE_EXPANSION_POLICY: bool = True
    ENABLE_FORMATTING_PROFILE: bool = True

    # Logging level
    LOG_LEVEL: str = "DEBUG"

    # Paths to translation scripts
    SCRIPTS_DIR: str = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "scripts")

    class Config:
        env_file = ".env.test" if os.environ.get("PYTEST_RUNNING") else ".env"
        env_file_encoding = "utf-8"

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self._apply_debug_defaults()
        self._validate_required_settings()

    def _apply_debug_defaults(self):
        """
        Apply stable, safe defaults for DEBUG mode.

        These defaults are intentionally predictable and safe for local development.
        They will persist across application restarts, maintaining session consistency.
        """
        if self.DEBUG:
            # Use a stable secret key for DEBUG mode (not randomly generated)
            if not self.SECRET_KEY or self.SECRET_KEY == "":
                self.SECRET_KEY = "dev-secret-key-32-characters-long-for-local-dev"
                warnings.warn(
                    "DEBUG mode: Using stable development SECRET_KEY. "
                    "This key is NOT secure for production use. "
                    "Set SECRET_KEY environment variable for any real deployment.",
                    UserWarning,
                    stacklevel=2
                )

            # Use a placeholder API key that will fail gracefully
            if not self.OPENAI_API_KEY:
                self.OPENAI_API_KEY = "debug-placeholder-key-invalid-for-production"
                warnings.warn(
                    "DEBUG mode: Using placeholder OPENAI_API_KEY. "
                    "Translation features will not work without a real API key. "
                    "Set OPENAI_API_KEY environment variable to enable translation.",
                    UserWarning,
                    stacklevel=2
                )

    def _validate_required_settings(self):
        """
        Validate that required settings are properly configured.

        Production validation is strict - no defaults are allowed.
        Development mode allows defaults with clear warnings.
        """
        # SECRET_KEY validation
        if not self.SECRET_KEY or self.SECRET_KEY == "":
            raise ValueError(
                "SECRET_KEY is required and cannot be empty. "
                "Set SECRET_KEY environment variable to a secure value (32+ characters)."
            )

        # Warn about short secret keys in production
        if not self.DEBUG and len(self.SECRET_KEY) < 32:
            raise ValueError(
                "SECRET_KEY must be at least 32 characters long for security. "
                f"Current length: {len(self.SECRET_KEY)}"
            )

        # Warn about short secret keys in development
        if self.DEBUG and len(self.SECRET_KEY) < 32:
            warnings.warn(
                f"SECRET_KEY should be at least 32 characters long for security. "
                f"Current length: {len(self.SECRET_KEY)}",
                UserWarning,
                stacklevel=2
            )

        # OPENAI_API_KEY validation
        if not self.OPENAI_API_KEY:
            if self.DEBUG:
                # In debug mode, we allow the placeholder but warn
                if self.OPENAI_API_KEY == "debug-placeholder-key-invalid-for-production":
                    warnings.warn(
                        "DEBUG mode: OPENAI_API_KEY is using a placeholder. "
                        "Translation features will not work. "
                        "Set OPENAI_API_KEY environment variable to enable translation features.",
                        UserWarning,
                        stacklevel=2
                    )
            else:
                # In production, we require a real API key
                raise ValueError(
                    "OPENAI_API_KEY is required in production mode. "
                    "Set OPENAI_API_KEY environment variable to enable translation features."
                )

    def is_openai_configured(self) -> bool:
        """
        Check if OpenAI API is properly configured with a real API key.

        Returns:
            bool: True if OpenAI is configured with a non-placeholder key
        """
        return (
            bool(self.OPENAI_API_KEY) and
            self.OPENAI_API_KEY != "debug-placeholder-key-invalid-for-production"
        )

    def is_redis_configured(self) -> bool:
        """
        Check if Redis is configured with a non-default URL.

        Returns:
            bool: True if Redis URL is not the default localhost
        """
        return self.REDIS_URL != "redis://localhost:6379"

    def get_environment_info(self) -> dict:
        """
        Get information about the current configuration environment.

        Returns:
            dict: Environment information including warnings and overrides
        """
        info = {
            "debug": self.DEBUG,
            "environment": "development" if self.DEBUG else "production",
            "using_defaults": {},
            "warnings": []
        }

        if self.DEBUG:
            if self.SECRET_KEY == "dev-secret-key-32-characters-long-for-local-dev":
                info["using_defaults"]["secret_key"] = True
                info["warnings"].append("Using default development SECRET_KEY")

            if self.OPENAI_API_KEY == "debug-placeholder-key-invalid-for-production":
                info["using_defaults"]["openai_api_key"] = True
                info["warnings"].append("Using placeholder OPENAI_API_KEY")

        return info


settings = Settings()
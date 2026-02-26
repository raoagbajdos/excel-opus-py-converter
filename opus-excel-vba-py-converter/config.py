"""
Application Configuration Module

Centralizes all configuration settings, loaded from environment variables
with sensible defaults.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


class Config:
    """Base configuration."""

    # Application
    SECRET_KEY: str = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
    APP_ENV: str = os.getenv("APP_ENV", os.getenv("FLASK_ENV", "development"))
    DEBUG: bool = os.getenv("APP_DEBUG", os.getenv("FLASK_DEBUG", "1")) == "1"

    # File uploads
    UPLOAD_FOLDER: str = os.getenv("UPLOAD_FOLDER", "uploads")
    MAX_FILE_SIZE_MB: int = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
    MAX_CONTENT_LENGTH: int = MAX_FILE_SIZE_MB * 1024 * 1024

    # Allowed file extensions
    VBA_EXTENSIONS: set[str] = {"xlsm", "xls", "xlsb", "xla", "xlam"}
    DATA_EXTENSIONS: set[str] = {"xlsm", "xls", "xlsb", "xla", "xlam", "xlsx"}
    ALL_EXTENSIONS: set[str] = VBA_EXTENSIONS | DATA_EXTENSIONS

    # LLM provider settings
    LLM_PROVIDER: str = os.getenv("LLM_PROVIDER", "anthropic")
    LLM_MODEL: str = os.getenv("LLM_MODEL", "claude-sonnet-4-20250514")
    ANTHROPIC_API_KEY: str | None = os.getenv("ANTHROPIC_API_KEY")
    OPENAI_API_KEY: str | None = os.getenv("OPENAI_API_KEY")

    # LLM request settings
    LLM_MAX_TOKENS: int = int(os.getenv("LLM_MAX_TOKENS", "4096"))
    LLM_FORMULA_MAX_TOKENS: int = int(os.getenv("LLM_FORMULA_MAX_TOKENS", "2048"))
    LLM_MAX_RETRIES: int = int(os.getenv("LLM_MAX_RETRIES", "3"))
    LLM_RETRY_BASE_DELAY: float = float(os.getenv("LLM_RETRY_BASE_DELAY", "1.0"))

    # Logging
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")

    @classmethod
    def ensure_upload_folder(cls) -> None:
        """Create the upload folder if it doesn't exist."""
        Path(cls.UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)


class ProductionConfig(Config):
    """Production configuration overrides."""
    DEBUG = False
    APP_ENV = "production"
    LOG_LEVEL = "WARNING"


class DevelopmentConfig(Config):
    """Development configuration."""
    DEBUG = True
    APP_ENV = "development"
    LOG_LEVEL = "DEBUG"


def get_config() -> type[Config]:
    """Return the appropriate config class based on APP_ENV."""
    env = os.getenv("APP_ENV", os.getenv("FLASK_ENV", "development"))
    if env == "production":
        return ProductionConfig
    return DevelopmentConfig

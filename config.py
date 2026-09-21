import os
from dataclasses import dataclass
from dotenv import load_dotenv
from openai import AsyncOpenAI

# Load environment variables from .env file
load_dotenv()


class ConfigError(Exception):
    """Custom exception raised for missing or invalid configuration settings."""
    pass


@dataclass
class Settings:
    """Central configuration class for the LQA Tool."""
    
    # LLM Settings
    openrouter_api_key: str = os.getenv("OPENROUTER_API_KEY", "")
    openrouter_base_url: str = os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
    
    # Read the comma-separated string from .env
    openrouter_models_raw: str = os.getenv("OPENROUTER_MODELS", "qwen/qwen3.8-flash")
    
    batch_size: int = int(os.getenv("BATCH_SIZE", "10"))
    max_retries: int = int(os.getenv("MAX_RETRIES", "3"))
    temperature: float = float(os.getenv("TEMPERATURE", "0.0"))
    
    # Supabase Settings
    supabase_url: str = os.getenv("SUPABASE_URL", "")
    supabase_key: str = os.getenv("SUPABASE_KEY", "")
    glossary_page_size: int = int(os.getenv("GLOSSARY_PAGE_SIZE", "1000"))

    @property
    def openrouter_models(self) -> list[str]:
        """Converts the raw comma-separated string into a clean list of model names."""
        return [m.strip() for m in self.openrouter_models_raw.split(",") if m.strip()]

# Global client instances for reuse across module calls
_settings_instance: Settings | None = None
_async_openai_client: AsyncOpenAI | None = None
_supabase_client = None


def get_settings() -> Settings:
    """Retrieve or initialize the global Settings instance."""
    global _settings_instance
    if _settings_instance is None:
        _settings_instance = Settings()
    return _settings_instance


def get_async_openai(settings: Settings | None = None) -> AsyncOpenAI:
    """Retrieve or initialize the AsyncOpenAI client configured for OpenRouter."""
    global _async_openai_client
    if _async_openai_client is not None:
        return _async_openai_client

    if settings is None:
        settings = get_settings()

    if not settings.openrouter_api_key:
        raise ConfigError("OPENROUTER_API_KEY is missing. Please set it in your .env file.")

    _async_openai_client = AsyncOpenAI(
        base_url=settings.openrouter_base_url,
        api_key=settings.openrouter_api_key,
        default_headers={
            "HTTP-Referer": "https://github.com/game-lqa-tool",
            "X-Title": "Game LQA Tool",
        },
    )
    return _async_openai_client


def require_supabase(settings: Settings | None = None) -> None:
    """Validate that Supabase credentials are present."""
    if settings is None:
        settings = get_settings()

    if not settings.supabase_url or not settings.supabase_key:
        raise ConfigError("SUPABASE_URL or SUPABASE_KEY is missing. Please check your .env file.")


def get_supabase(settings: Settings | None = None):
    """Retrieve or initialize the Supabase client."""
    global _supabase_client
    if _supabase_client is not None:
        return _supabase_client

    if settings is None:
        settings = get_settings()

    require_supabase(settings)

    try:
        from supabase import create_client
    except ImportError as exc:
        raise ConfigError("Supabase library not installed. Please run `pip install supabase`.") from exc

    _supabase_client = create_client(settings.supabase_url, settings.supabase_key)
    return _supabase_client
# Copyright (c) Microsoft. All rights reserved.

"""
Configuration Module

Centralized configuration management for the A365 Agent Framework.
Loads settings from environment variables with validation.
"""

import logging
import os
from dataclasses import dataclass, field
from typing import Optional

from dotenv import load_dotenv

# Load environment variables on module import
load_dotenv()

logger = logging.getLogger(__name__)


@dataclass
class AzureOpenAIModelConfig:
    """Configuration for a single Azure OpenAI model deployment."""
    
    endpoint: str
    deployment: str
    api_key: str
    api_version: str = "2024-05-01-preview"
    name: str = ""  # Friendly name for logging
    
    def __post_init__(self):
        if not self.name:
            self.name = f"{self.deployment}@{self.endpoint.split('//')[1].split('.')[0]}"
    
    @property
    def is_valid(self) -> bool:
        """Check if settings are valid."""
        return bool(self.endpoint and self.deployment and self.api_key)


class AzureOpenAIModelPool:
    """
    Pool of Azure OpenAI models with round-robin load balancing and failover.
    
    Automatically rotates through available models and skips throttled ones.
    """
    
    def __init__(self):
        self.models: list[AzureOpenAIModelConfig] = []
        self._current_index = 0
        self._throttled_until: dict[int, float] = {}  # model_index -> timestamp
        self._load_models_from_env()
    
    def _load_models_from_env(self) -> None:
        """Load all models from environment variables."""
        api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-05-01-preview")
        
        # Load numbered models (AZURE_OPENAI_MODEL_1_*, AZURE_OPENAI_MODEL_2_*, etc.)
        for i in range(1, 20):  # Support up to 20 models
            endpoint = os.getenv(f"AZURE_OPENAI_MODEL_{i}_ENDPOINT")
            deployment = os.getenv(f"AZURE_OPENAI_MODEL_{i}_DEPLOYMENT")
            api_key = os.getenv(f"AZURE_OPENAI_MODEL_{i}_API_KEY")
            
            if endpoint and deployment and api_key:
                model = AzureOpenAIModelConfig(
                    endpoint=endpoint,
                    deployment=deployment,
                    api_key=api_key,
                    api_version=api_version,
                )
                self.models.append(model)
                logger.info(f"📦 Loaded model {i}: {model.name}")
        
        # Fallback to legacy single-model config if no numbered models found
        if not self.models:
            endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "")
            deployment = os.getenv("AZURE_OPENAI_DEPLOYMENT", "")
            api_key = os.getenv("AZURE_OPENAI_API_KEY", "")
            
            if endpoint and deployment and api_key:
                model = AzureOpenAIModelConfig(
                    endpoint=endpoint,
                    deployment=deployment,
                    api_key=api_key,
                    api_version=api_version,
                )
                self.models.append(model)
                logger.info(f"📦 Loaded legacy model: {model.name}")
        
        if self.models:
            logger.info(f"✅ Model pool initialized with {len(self.models)} models")
        else:
            logger.warning("⚠️ No Azure OpenAI models configured!")
    
    def get_next_model(self) -> AzureOpenAIModelConfig:
        """
        Get the next available model using round-robin.
        Skips models that are currently throttled.
        """
        import time
        
        if not self.models:
            raise ValueError("No Azure OpenAI models configured")
        
        now = time.time()
        attempts = 0
        
        while attempts < len(self.models):
            # Check if current model is throttled
            throttled_until = self._throttled_until.get(self._current_index, 0)
            
            if now >= throttled_until:
                # Model is available
                model = self.models[self._current_index]
                self._current_index = (self._current_index + 1) % len(self.models)
                return model
            
            # Model is throttled, try next
            self._current_index = (self._current_index + 1) % len(self.models)
            attempts += 1
        
        # All models are throttled, return the one with shortest wait
        min_wait_idx = min(self._throttled_until, key=self._throttled_until.get)
        wait_time = self._throttled_until[min_wait_idx] - now
        logger.warning(f"⏳ All models throttled! Shortest wait: {wait_time:.1f}s for {self.models[min_wait_idx].name}")
        return self.models[min_wait_idx]
    
    def mark_throttled(self, model: AzureOpenAIModelConfig, retry_after: float = 60.0) -> None:
        """Mark a model as throttled for a period of time."""
        import time
        
        try:
            idx = self.models.index(model)
            self._throttled_until[idx] = time.time() + retry_after
            logger.warning(f"🚫 Model {model.name} throttled for {retry_after:.0f}s")
        except ValueError:
            pass  # Model not in list
    
    def clear_throttle(self, model: AzureOpenAIModelConfig) -> None:
        """Clear throttle status for a model."""
        try:
            idx = self.models.index(model)
            if idx in self._throttled_until:
                del self._throttled_until[idx]
        except ValueError:
            pass
    
    @property
    def available_count(self) -> int:
        """Number of models not currently throttled."""
        import time
        now = time.time()
        return sum(1 for i in range(len(self.models)) 
                   if self._throttled_until.get(i, 0) <= now)
    
    def __len__(self) -> int:
        return len(self.models)


@dataclass
class AzureOpenAISettings:
    """Azure OpenAI configuration settings (legacy single-model support)."""
    
    endpoint: str = ""
    deployment: str = ""
    api_version: str = ""
    api_key: Optional[str] = None
    
    def __post_init__(self):
        """Load from environment if not provided."""
        self.endpoint = self.endpoint or os.getenv("AZURE_OPENAI_ENDPOINT", "")
        self.deployment = self.deployment or os.getenv("AZURE_OPENAI_DEPLOYMENT", "")
        self.api_version = self.api_version or os.getenv("AZURE_OPENAI_API_VERSION", "")
        self.api_key = self.api_key or os.getenv("AZURE_OPENAI_API_KEY")
    
    def validate(self) -> None:
        """Validate required settings are present."""
        if not self.endpoint:
            raise ValueError("AZURE_OPENAI_ENDPOINT is required")
        if not self.deployment:
            raise ValueError("AZURE_OPENAI_DEPLOYMENT is required")
        if not self.api_version:
            raise ValueError("AZURE_OPENAI_API_VERSION is required")
    
    @property
    def is_valid(self) -> bool:
        """Check if settings are valid."""
        return bool(self.endpoint and self.deployment and self.api_version)


@dataclass
class AgentAuthSettings:
    """Agent authentication configuration (Blueprint/Service Connection)."""
    
    client_id: str = ""
    client_secret: str = ""
    tenant_id: str = ""
    scopes: str = ""
    
    def __post_init__(self):
        """Load from environment if not provided."""
        # Try consolidated CONNECTIONS vars first (preferred), then legacy vars
        self.client_id = self.client_id or os.getenv(
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID",
            os.getenv("CLIENT_ID", "")
        )
        self.client_secret = self.client_secret or os.getenv(
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET",
            os.getenv("CLIENT_SECRET", "")
        )
        self.tenant_id = self.tenant_id or os.getenv(
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID",
            os.getenv("TENANT_ID", "")
        )
        self.scopes = self.scopes or os.getenv(
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__SCOPES",
            "5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default"
        )
    
    @property
    def is_valid(self) -> bool:
        """Check if client credentials are available."""
        return bool(self.client_id and self.client_secret and self.tenant_id)
    
    @property
    def scopes_list(self) -> list[str]:
        """Get scopes as a list."""
        return [self.scopes] if self.scopes else []


@dataclass
class ObservabilitySettings:
    """Observability configuration settings."""
    
    enabled: bool = False
    service_name: str = "agent-framework-sample"
    service_namespace: str = "agent-framework.samples"
    enable_a365_exporter: bool = False
    enable_otel: bool = False
    enable_sensitive_data: bool = False
    
    def __post_init__(self):
        """Load from environment if not provided."""
        self.enabled = os.getenv("ENABLE_OBSERVABILITY", "").lower() == "true"
        self.service_name = os.getenv("OBSERVABILITY_SERVICE_NAME", self.service_name)
        self.service_namespace = os.getenv("OBSERVABILITY_SERVICE_NAMESPACE", self.service_namespace)
        self.enable_a365_exporter = os.getenv("ENABLE_A365_OBSERVABILITY_EXPORTER", "").lower() == "true"
        self.enable_otel = os.getenv("ENABLE_OTEL", "").lower() == "true"
        self.enable_sensitive_data = os.getenv("ENABLE_SENSITIVE_DATA", "").lower() == "true"


@dataclass
class ServerSettings:
    """Server configuration settings."""
    
    port: int = 3978
    auth_handler_name: Optional[str] = None
    
    def __post_init__(self):
        """Load from environment if not provided."""
        self.port = int(os.getenv("PORT", str(self.port)))
        auth_handler = os.getenv("AUTH_HANDLER_NAME", "")
        self.auth_handler_name = auth_handler if auth_handler else None


@dataclass 
class MCPSettings:
    """MCP (Model Context Protocol) configuration settings."""
    
    server_host: str = ""
    platform_endpoint: str = "https://agent365.svc.cloud.microsoft"
    
    def __post_init__(self):
        """Load from environment if not provided."""
        self.server_host = os.getenv("MCP_SERVER_HOST", "")
        self.platform_endpoint = os.getenv("MCP_PLATFORM_ENDPOINT", self.platform_endpoint)


@dataclass
class Settings:
    """
    Master configuration class that aggregates all settings.
    
    Usage:
        settings = Settings()
        settings.azure_openai.validate()
        if settings.agent_auth.is_valid:
            # Use client credentials
    """
    
    azure_openai: AzureOpenAISettings = field(default_factory=AzureOpenAISettings)
    agent_auth: AgentAuthSettings = field(default_factory=AgentAuthSettings)
    observability: ObservabilitySettings = field(default_factory=ObservabilitySettings)
    server: ServerSettings = field(default_factory=ServerSettings)
    mcp: MCPSettings = field(default_factory=MCPSettings)
    
    # Model pool for load balancing (not a dataclass field - initialized separately)
    model_pool: Optional[AzureOpenAIModelPool] = field(default=None, repr=False)
    
    # Development settings
    bearer_token: str = ""
    use_agentic_auth: bool = True
    log_level: str = "INFO"
    
    def __post_init__(self):
        """Load development settings from environment."""
        self.bearer_token = os.getenv("BEARER_TOKEN", "")
        self.use_agentic_auth = os.getenv("USE_AGENTIC_AUTH", "true").lower() == "true"
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
        
        # Initialize the model pool for multi-model load balancing
        self.model_pool = AzureOpenAIModelPool()
    
    @classmethod
    def from_environment(cls) -> "Settings":
        """Create settings instance loaded from environment."""
        return cls()
    
    def configure_logging(self) -> None:
        """Configure logging based on settings."""
        # Use a cleaner format that shows what matters
        logging.basicConfig(
            level=getattr(logging, self.log_level),
            format="%(levelname)s:%(name)s:%(message)s"
        )
        
        # Control agent_framework log level (MCP tool calls) via env
        af_log_level = os.getenv("AGENT_FRAMEWORK_LOG_LEVEL", "INFO").upper()
        logging.getLogger("agent_framework").setLevel(getattr(logging, af_log_level, logging.INFO))
        
        # Suppress verbose Azure SDK logging
        logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
        logging.getLogger("azure.identity").setLevel(logging.ERROR)
        logging.getLogger("microsoft_agents_a365.observability").setLevel(logging.ERROR)
        
        # Suppress noisy Microsoft Agents SDK loggers (typing indicators, token attempts, etc.)
        logging.getLogger("microsoft_agents.hosting.core.connector.client.connector_client").setLevel(logging.WARNING)
        logging.getLogger("microsoft_agents.authentication.msal.msal_auth").setLevel(logging.WARNING)
        logging.getLogger("microsoft_agents.hosting.core.rest_channel_service_client_factory").setLevel(logging.WARNING)
        logging.getLogger("microsoft_agents.hosting.core.app.oauth._handlers.agentic_user_authorization").setLevel(logging.WARNING)
        
        # Suppress HTTP request logging (httpx, aiohttp)
        logging.getLogger("httpx").setLevel(logging.WARNING)
        logging.getLogger("aiohttp.access").setLevel(logging.WARNING)
        
        # Suppress MCP protocol noise (session IDs, stream reconnects)
        logging.getLogger("mcp.client.streamable_http").setLevel(logging.WARNING)
        
        # Suppress OpenAI SDK retry logging (we handle failover ourselves)
        logging.getLogger("openai._base_client").setLevel(logging.WARNING)


# Global singleton settings instance
_settings: Optional[Settings] = None


def get_settings() -> Settings:
    """Get the global settings instance (lazy loaded)."""
    global _settings
    if _settings is None:
        _settings = Settings.from_environment()
    return _settings

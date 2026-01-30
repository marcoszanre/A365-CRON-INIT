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
class AzureOpenAISettings:
    """Azure OpenAI configuration settings."""
    
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
    
    # Development settings
    bearer_token: str = ""
    use_agentic_auth: bool = True
    log_level: str = "INFO"
    
    def __post_init__(self):
        """Load development settings from environment."""
        self.bearer_token = os.getenv("BEARER_TOKEN", "")
        self.use_agentic_auth = os.getenv("USE_AGENTIC_AUTH", "true").lower() == "true"
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
    
    @classmethod
    def from_environment(cls) -> "Settings":
        """Create settings instance loaded from environment."""
        return cls()
    
    def configure_logging(self) -> None:
        """Configure logging based on settings."""
        logging.basicConfig(level=getattr(logging, self.log_level))
        
        # Suppress verbose Azure SDK logging
        logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
        logging.getLogger("azure.identity").setLevel(logging.ERROR)
        logging.getLogger("microsoft_agents_a365.observability").setLevel(logging.ERROR)


# Global singleton settings instance
_settings: Optional[Settings] = None


def get_settings() -> Settings:
    """Get the global settings instance (lazy loaded)."""
    global _settings
    if _settings is None:
        _settings = Settings.from_environment()
    return _settings

# Copyright (c) Microsoft. All rights reserved.

"""
Authentication Module

Token management, caching, and authentication utilities for A365 agents.
"""

import logging
from dataclasses import dataclass
from typing import Optional

from azure.identity import ClientSecretCredential

from a365_agent.config import get_settings

logger = logging.getLogger(__name__)


# =============================================================================
# TOKEN CACHE
# =============================================================================

class TokenCache:
    """
    Thread-safe token cache for Agent 365 Observability and other services.
    
    Tokens are keyed by tenant_id:agent_id combination.
    """
    
    def __init__(self):
        self._cache: dict[str, str] = {}
    
    def _make_key(self, tenant_id: str, agent_id: str) -> str:
        """Create a cache key from tenant and agent IDs."""
        return f"{tenant_id}:{agent_id}"
    
    def set(self, tenant_id: str, agent_id: str, token: str) -> None:
        """Cache a token for the given tenant and agent."""
        key = self._make_key(tenant_id, agent_id)
        self._cache[key] = token
        logger.debug(f"Cached token for {key}")
    
    def get(self, tenant_id: str, agent_id: str) -> Optional[str]:
        """Retrieve a cached token, or None if not found."""
        key = self._make_key(tenant_id, agent_id)
        token = self._cache.get(key)
        if token:
            logger.debug(f"Token cache hit for {key}")
        else:
            logger.debug(f"Token cache miss for {key}")
        return token
    
    def clear(self, tenant_id: Optional[str] = None, agent_id: Optional[str] = None) -> None:
        """Clear tokens from the cache.
        
        If both tenant_id and agent_id provided, clears that specific entry.
        Otherwise clears all tokens.
        """
        if tenant_id and agent_id:
            key = self._make_key(tenant_id, agent_id)
            self._cache.pop(key, None)
            logger.debug(f"Cleared token for {key}")
        else:
            self._cache.clear()
            logger.debug("Cleared all tokens from cache")


# Global token cache instance
_token_cache = TokenCache()


def cache_agentic_token(tenant_id: str, agent_id: str, token: str) -> None:
    """Cache an agentic token for observability."""
    _token_cache.set(tenant_id, agent_id, token)


def get_cached_agentic_token(tenant_id: str, agent_id: str) -> Optional[str]:
    """Get a cached agentic token."""
    return _token_cache.get(tenant_id, agent_id)


# =============================================================================
# LOCAL AUTHENTICATION OPTIONS
# =============================================================================

@dataclass
class LocalAuthOptions:
    """
    Authentication options for local/development scenarios.
    
    Used when running the agent locally without full agentic auth.
    """
    
    env_id: str = ""
    bearer_token: str = ""
    
    def __post_init__(self):
        """Ensure string types."""
        if not isinstance(self.env_id, str):
            self.env_id = str(self.env_id) if self.env_id else ""
        if not isinstance(self.bearer_token, str):
            self.bearer_token = str(self.bearer_token) if self.bearer_token else ""
    
    @property
    def is_valid(self) -> bool:
        """Check if auth options are valid."""
        return bool(self.env_id and self.bearer_token)
    
    @classmethod
    def from_environment(cls) -> "LocalAuthOptions":
        """Create from environment variables."""
        settings = get_settings()
        import os
        return cls(
            env_id=os.getenv("ENV_ID", ""),
            bearer_token=settings.bearer_token
        )


# =============================================================================
# CLIENT CREDENTIALS
# =============================================================================

def get_client_credential() -> Optional[ClientSecretCredential]:
    """
    Get a ClientSecretCredential for the agent's service connection.
    
    Returns None if client credentials are not configured.
    
    NOTE: Agentic applications cannot use client credentials to acquire
    tokens for MCP or Bot Framework resources (AADSTS82001 error).
    This is only useful for non-agentic scenarios.
    """
    settings = get_settings()
    
    if not settings.agent_auth.is_valid:
        logger.debug("Client credentials not configured")
        return None
    
    return ClientSecretCredential(
        tenant_id=settings.agent_auth.tenant_id,
        client_id=settings.agent_auth.client_id,
        client_secret=settings.agent_auth.client_secret,
    )


async def acquire_token_with_client_credentials(scopes: list[str]) -> Optional[str]:
    """
    Attempt to acquire a token using client credentials.
    
    Args:
        scopes: The scopes to request
        
    Returns:
        The access token, or None if acquisition failed
        
    NOTE: This will fail for agentic apps trying to access MCP or Bot Framework.
    """
    credential = get_client_credential()
    if not credential:
        return None
    
    try:
        token = credential.get_token(*scopes)
        logger.info("Successfully acquired token with client credentials")
        return token.token
    except Exception as e:
        logger.warning(f"Failed to acquire token with client credentials: {e}")
        return None

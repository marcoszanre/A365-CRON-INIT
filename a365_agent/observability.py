# Copyright (c) Microsoft. All rights reserved.

"""
Observability Module

Telemetry, tracing, and monitoring setup for A365 agents.
Integrates with Microsoft Agent 365 Observability SDK.
"""

import logging
from typing import Callable, Optional

from a365_agent.auth import get_cached_agentic_token
from a365_agent.config import get_settings

logger = logging.getLogger(__name__)

# Type alias for token resolver function
TokenResolver = Callable[[str, str], Optional[str]]


def default_token_resolver(agent_id: str, tenant_id: str) -> Optional[str]:
    """
    Default token resolver for Agent 365 Observability exporter.
    
    This function is called by the observability SDK to retrieve authentication
    tokens for exporting telemetry to the Agent 365 service.
    
    Args:
        agent_id: The unique identifier for the AI agent
        tenant_id: The tenant ID for the agent
        
    Returns:
        The cached agentic token, or None if not available
    """
    try:
        cached_token = get_cached_agentic_token(tenant_id, agent_id)
        if not cached_token:
            logger.debug(f"No cached observability token for agent {agent_id}, tenant {tenant_id}")
        return cached_token
    except Exception as e:
        logger.error(f"Error resolving observability token: {e}")
        return None


def configure_observability(
    token_resolver: Optional[TokenResolver] = None,
    service_name: Optional[str] = None,
    service_namespace: Optional[str] = None,
) -> None:
    """
    Configure Agent 365 Observability.
    
    This sets up telemetry export to the Agent 365 platform.
    
    Args:
        token_resolver: Function to resolve auth tokens. Defaults to using token cache.
        service_name: Service name for telemetry. Defaults to settings value.
        service_namespace: Service namespace for telemetry. Defaults to settings value.
    """
    settings = get_settings()
    
    if not settings.observability.enabled:
        logger.info("Observability is disabled")
        return
    
    try:
        from microsoft_agents_a365.observability.core.config import configure
        
        configure(
            service_name=service_name or settings.observability.service_name,
            service_namespace=service_namespace or settings.observability.service_namespace,
            token_resolver=token_resolver or default_token_resolver,
        )
        
        logger.info("✅ Observability configured")
        
    except ImportError:
        logger.warning("⚠️ Observability SDK not available")
    except Exception as e:
        logger.warning(f"⚠️ Failed to configure observability: {e}")


def enable_agentframework_instrumentation() -> None:
    """
    Enable automatic instrumentation for AgentFramework SDK.
    
    This instruments the AgentFramework to automatically capture spans
    for agent operations, tool calls, and LLM interactions.
    """
    try:
        from microsoft_agents_a365.observability.extensions.agentframework.trace_instrumentor import (
            AgentFrameworkInstrumentor,
        )
        
        AgentFrameworkInstrumentor().instrument()
        logger.info("✅ AgentFramework instrumentation enabled")
        
    except ImportError:
        logger.warning("⚠️ AgentFramework instrumentor not available")
    except Exception as e:
        logger.warning(f"⚠️ Failed to enable instrumentation: {e}")


class ObservabilityContext:
    """
    Context manager for observability baggage (correlation IDs, etc.).
    
    Usage:
        with ObservabilityContext(tenant_id, agent_id, correlation_id):
            # Operations here will have baggage attached
            response = await agent.run(message)
    """
    
    def __init__(
        self,
        tenant_id: str,
        agent_id: str,
        correlation_id: str,
    ):
        self.tenant_id = tenant_id
        self.agent_id = agent_id
        self.correlation_id = correlation_id
        self._baggage_context = None
    
    def __enter__(self):
        """Enter the observability context."""
        try:
            from microsoft_agents_a365.observability.core.middleware.baggage_builder import (
                BaggageBuilder,
            )
            
            self._baggage_context = (
                BaggageBuilder()
                .tenant_id(self.tenant_id)
                .agent_id(self.agent_id)
                .correlation_id(self.correlation_id)
                .build()
            )
            return self._baggage_context.__enter__()
            
        except ImportError:
            logger.debug("Baggage builder not available")
            return self
        except Exception as e:
            logger.debug(f"Failed to create baggage context: {e}")
            return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit the observability context."""
        if self._baggage_context:
            return self._baggage_context.__exit__(exc_type, exc_val, exc_tb)
        return False

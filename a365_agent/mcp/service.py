# Copyright (c) Microsoft. All rights reserved.

"""
MCP Service Module

Manages MCP (Model Context Protocol) server connections and tool registration.

Platform Limitations for Agentic Apps:
- Cannot acquire app-only tokens (AADSTS82001 error)
- Cannot pre-initialize MCP servers at startup (requires user token)
- MCP initialization must happen on first user request using agentic token exchange
"""

import asyncio
import logging
from typing import Any, Optional

from a365_agent.config import get_settings

logger = logging.getLogger(__name__)


class MCPService:
    """
    Service for managing MCP server connections and tool registration.
    
    This service handles:
    - Lazy initialization of MCP servers
    - Tool registration with the AgentFramework agent
    - Token management for MCP platform access
    
    Platform Limitations (Agentic Auth):
        Agentic applications CANNOT use client credentials to acquire MCP tokens.
        The AADSTS82001 error occurs because agentic apps are not permitted to
        request app-only tokens. As a result, MCP initialization must happen
        on the first user request, using the agentic token exchange flow.
    """
    
    def __init__(self):
        """Initialize the MCP service."""
        self._tool_service = None
        self._initialized = False
        self._init_error: Optional[str] = None
    
    @property
    def is_initialized(self) -> bool:
        """Check if MCP servers have been initialized."""
        return self._initialized
    
    @property
    def init_error(self) -> Optional[str]:
        """Get the initialization error message, if any."""
        return self._init_error
    
    def _get_tool_service(self):
        """Get or create the MCP tool registration service."""
        if self._tool_service is None:
            try:
                from microsoft_agents_a365.tooling.extensions.agentframework.services.mcp_tool_registration_service import (
                    McpToolRegistrationService,
                )
                self._tool_service = McpToolRegistrationService()
                logger.info("✅ MCP tool service created")
            except ImportError as e:
                logger.error(f"❌ MCP SDK not available: {e}")
                raise
            except Exception as e:
                logger.error(f"❌ Failed to create MCP tool service: {e}")
                raise
        return self._tool_service
    
    async def initialize_with_bearer_token(
        self,
        chat_client: Any,
        agent_instructions: str,
        bearer_token: str,
        auth: Any,
        auth_handler_name: Optional[str],
        turn_context: Any,
        initial_tools: Optional[list] = None,
    ) -> Any:
        """
        Initialize MCP servers using a bearer token (dev mode or proactive mode).
        
        This is the fast path for development when you have a pre-acquired token.
        For proactive scenarios, you can pass the Blueprint-acquired token directly.
        
        Args:
            chat_client: The Azure OpenAI chat client
            agent_instructions: The agent's system prompt
            bearer_token: Pre-acquired bearer token for MCP
            auth: Authorization handler (required by SDK)
            auth_handler_name: Auth handler name (required by SDK)
            turn_context: TurnContext (required by SDK)
            initial_tools: Optional list of initial tools
            
        Returns:
            The agent with MCP tools registered
        """
        if self._initialized:
            logger.info("✅ MCP servers already initialized")
            return None
        
        logger.info("🚀 Initializing MCP servers with bearer token...")
        init_start = asyncio.get_event_loop().time()
        
        try:
            tool_service = self._get_tool_service()
            
            # Build kwargs for SDK call - all params are now required
            sdk_kwargs = {
                "chat_client": chat_client,
                "agent_instructions": agent_instructions,
                "initial_tools": initial_tools or [],
                "auth_token": bearer_token,
                "auth": auth,
                "auth_handler_name": auth_handler_name,
                "turn_context": turn_context,
            }
            
            agent = await tool_service.add_tool_servers_to_agent(**sdk_kwargs)
            
            init_duration = asyncio.get_event_loop().time() - init_start
            self._initialized = True
            logger.info(f"✅ MCP initialization completed in {init_duration:.1f}s")
            
            return agent
            
        except Exception as e:
            self._init_error = str(e)
            logger.error(f"❌ MCP initialization failed: {e}")
            raise
    
    async def initialize_with_agentic_auth(
        self,
        chat_client: Any,
        agent_instructions: str,
        auth: Any,  # Authorization from microsoft_agents SDK
        auth_handler_name: Optional[str],
        turn_context: Any,  # TurnContext
        initial_tools: Optional[list] = None,
    ) -> Any:
        """
        Initialize MCP servers using agentic token exchange (production).
        
        This is required for production agentic apps because they cannot
        acquire tokens using client credentials. The SDK will perform
        the agentic token exchange using the user's request context.
        
        Args:
            chat_client: The Azure OpenAI chat client
            agent_instructions: The agent's system prompt
            auth: The Authorization handler from microsoft_agents SDK
            auth_handler_name: Name of the auth handler (e.g., "AGENTIC")
            turn_context: The TurnContext from the current request
            initial_tools: Optional list of initial tools
            
        Returns:
            The agent with MCP tools registered
        """
        if self._initialized:
            logger.info("✅ MCP servers already initialized")
            return None
        
        logger.info("🚀 Initializing MCP servers with agentic token exchange...")
        init_start = asyncio.get_event_loop().time()
        
        try:
            tool_service = self._get_tool_service()
            
            agent = await tool_service.add_tool_servers_to_agent(
                chat_client=chat_client,
                agent_instructions=agent_instructions,
                initial_tools=initial_tools or [],
                auth=auth,
                auth_handler_name=auth_handler_name,
                turn_context=turn_context,
                # No auth_token - SDK will do agentic token exchange
            )
            
            init_duration = asyncio.get_event_loop().time() - init_start
            self._initialized = True
            logger.info(f"✅ MCP initialization completed in {init_duration:.1f}s")
            
            return agent
            
        except Exception as e:
            self._init_error = str(e)
            logger.error(f"❌ MCP initialization failed: {e}")
            raise
    
    def ensure_ready(self) -> None:
        """
        Verify that MCP servers are initialized.
        
        Raises:
            RuntimeError: If MCP servers are not initialized
        """
        if not self._initialized:
            error_msg = self._init_error or "MCP servers not initialized"
            raise RuntimeError(f"MCP not ready: {error_msg}")
    
    async def cleanup(self) -> None:
        """Clean up MCP service resources."""
        if self._tool_service:
            try:
                await self._tool_service.cleanup()
                logger.info("✅ MCP service cleaned up")
            except Exception as e:
                logger.error(f"❌ MCP cleanup error: {e}")
        
        self._initialized = False
        self._tool_service = None

# Copyright (c) Microsoft. All rights reserved.

"""
Generic Agent Host Module

Provides the server infrastructure for hosting A365 agents.
Handles HTTP routing, authentication, and notification dispatch.

Platform Limitations (Agentic Auth):
    - Cannot acquire app-only tokens (AADSTS82001)
    - Cannot send proactive messages
    - Cannot pre-initialize MCP servers at startup
    - All processing must complete within HTTP request lifecycle
"""

import asyncio
import logging
import socket
import uuid
from os import environ
from typing import Optional, Type

from aiohttp.web import Application, Request, Response, json_response, run_app
from aiohttp.web_middlewares import middleware as web_middleware

from a365_agent.auth import cache_agentic_token
from a365_agent.base import AgentBase, check_agent_inheritance
from a365_agent.config import get_settings
from a365_agent.notifications import (
    NotificationHandlerMixin,
    safe_send_activity,
    safe_send_email_response,
)
from a365_agent.observability import ObservabilityContext, configure_observability

# Microsoft Agents SDK imports
from microsoft_agents.activity import load_configuration_from_env
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import (
    CloudAdapter,
    jwt_authorization_middleware,
    start_agent_process,
)
from microsoft_agents.hosting.core import (
    AgentApplication,
    AgentAuthConfiguration,
    AuthenticationConstants,
    Authorization,
    ClaimsIdentity,
    MemoryStorage,
    TurnContext,
    TurnState,
)

# Agent 365 SDK imports
from microsoft_agents_a365.notifications.agent_notification import (
    AgentNotification,
    AgentNotificationActivity,
    ChannelId,
)
from microsoft_agents_a365.runtime.environment_utils import (
    get_observability_authentication_scope,
)

logger = logging.getLogger(__name__)

# Load agents SDK configuration
_agents_sdk_config = load_configuration_from_env(environ)


def create_and_run_host(
    agent_class: Type[AgentBase],
    *agent_args,
    **agent_kwargs,
) -> None:
    """
    Create and run a generic agent host.
    
    This is the main entry point for hosting an A365 agent.
    
    Args:
        agent_class: The agent class to host (must inherit from AgentBase)
        *agent_args: Positional arguments to pass to the agent constructor
        **agent_kwargs: Keyword arguments to pass to the agent constructor
        
    Example:
        from a365_agent import create_and_run_host
        from agents.contoso_agent import ContosoAgent
        
        create_and_run_host(ContosoAgent)
    """
    if not check_agent_inheritance(agent_class):
        raise TypeError(
            f"Agent class {agent_class.__name__} must inherit from AgentBase"
        )
    
    # Configure observability
    configure_observability()
    
    # Create and start the host
    host = GenericAgentHost(agent_class, *agent_args, **agent_kwargs)
    auth_config = host.create_auth_configuration()
    host.start_server(auth_config)


class GenericAgentHost(NotificationHandlerMixin):
    """
    Generic host for agents implementing AgentBase.
    
    Provides:
    - HTTP server for Bot Framework messages
    - Notification routing (email, Word, Excel, PowerPoint, lifecycle)
    - Authentication handling
    - Observability integration
    
    Platform Limitations (Agentic Auth):
        - Cannot acquire app-only tokens (AADSTS82001)
        - Cannot send proactive messages
        - Cannot pre-initialize MCP servers at startup
        - All processing must complete within HTTP request lifecycle
    """
    
    # Processing timeout constants
    FIRST_REQUEST_TIMEOUT = 120  # 2 minutes for first request (MCP init)
    NORMAL_REQUEST_TIMEOUT = 90  # 90 seconds for normal requests
    
    def __init__(
        self,
        agent_class: Type[AgentBase],
        *agent_args,
        **agent_kwargs,
    ):
        """
        Initialize the agent host.
        
        Args:
            agent_class: The agent class to host
            *agent_args: Arguments for agent constructor
            **agent_kwargs: Keyword arguments for agent constructor
        """
        if not check_agent_inheritance(agent_class):
            raise TypeError(
                f"Agent class {agent_class.__name__} must inherit from AgentBase"
            )
        
        self.settings = get_settings()
        self.settings.configure_logging()
        
        # Auth handler configuration
        self.auth_handler_name = self.settings.server.auth_handler_name
        if self.auth_handler_name:
            logger.info(f"🔐 Using auth handler: {self.auth_handler_name}")
        else:
            logger.info("🔓 No auth handler configured")
        
        # Agent configuration
        self.agent_class = agent_class
        self.agent_args = agent_args
        self.agent_kwargs = agent_kwargs
        self.agent_instance: Optional[AgentBase] = None
        
        # SDK components
        self.storage = MemoryStorage()
        self.connection_manager = MsalConnectionManager(**_agents_sdk_config)
        self.adapter = CloudAdapter(connection_manager=self.connection_manager)
        self.authorization = Authorization(
            self.storage, self.connection_manager, **_agents_sdk_config
        )
        self.agent_app = AgentApplication[TurnState](
            storage=self.storage,
            adapter=self.adapter,
            authorization=self.authorization,
            **_agents_sdk_config,
        )
        
        # Notification dispatcher
        self.agent_notification = AgentNotification(self.agent_app)
        
        # Register all handlers
        self._setup_handlers()
        logger.info("✅ Notification handlers registered")
    
    # =========================================================================
    # OBSERVABILITY
    # =========================================================================
    
    async def _setup_observability_token(
        self,
        context: TurnContext,
        tenant_id: str,
        agent_id: str,
    ) -> None:
        """Exchange and cache token for observability."""
        if not self.auth_handler_name:
            return
        
        try:
            token_result = await self.agent_app.auth.exchange_token(
                context,
                scopes=get_observability_authentication_scope(),
                auth_handler_id=self.auth_handler_name,
            )
            
            if token_result and hasattr(token_result, 'token') and token_result.token:
                cache_agentic_token(tenant_id, agent_id, token_result.token)
                logger.info("✅ Observability token cached")
            else:
                logger.warning("⚠️ Token exchange returned no token")
                
        except Exception as e:
            logger.warning(f"⚠️ Failed to cache observability token: {e}")
    
    async def _validate_and_setup_context(
        self,
        context: TurnContext,
    ) -> Optional[tuple[str, str, str]]:
        """
        Validate agent instance and setup observability context.
        
        Returns:
            Tuple of (tenant_id, agent_id, correlation_id) or None if validation fails
        """
        tenant_id = context.activity.recipient.tenant_id
        agent_id = context.activity.recipient.agentic_app_id
        correlation_id = context.activity.id or str(uuid.uuid4())
        
        if not self.agent_instance:
            logger.error("Agent not available")
            await context.send_activity("❌ Sorry, the agent is not available.")
            return None
        
        await self._setup_observability_token(context, tenant_id, agent_id)
        return tenant_id, agent_id, correlation_id
    
    # =========================================================================
    # HANDLER REGISTRATION
    # =========================================================================
    
    def _setup_handlers(self) -> None:
        """Register all message and notification handlers."""
        handler_config = (
            {"auth_handlers": [self.auth_handler_name]}
            if self.auth_handler_name
            else {}
        )
        
        # Welcome/help handler
        self._register_help_handler(handler_config)
        
        # Notification handlers (must be registered BEFORE message handler)
        self._register_email_handler(handler_config)
        self._register_word_handler(handler_config)
        self._register_excel_handler(handler_config)
        self._register_powerpoint_handler(handler_config)
        self._register_lifecycle_handler(handler_config)
        self._register_generic_notification_handler(handler_config)
        
        # Message handler (registered last - fallback)
        self._register_message_handler(handler_config)
    
    def _register_help_handler(self, handler_config: dict) -> None:
        """Register welcome/help handler."""
        
        async def help_handler(context: TurnContext, _: TurnState):
            await context.send_activity(
                f"👋 **Hi there!** I'm **{self.agent_class.__name__}**, your AI assistant.\n\n"
                "How can I help you today?"
            )
        
        self.agent_app.conversation_update("membersAdded", **handler_config)(help_handler)
        self.agent_app.message("/help", **handler_config)(help_handler)
    
    def _register_email_handler(self, handler_config: dict) -> None:
        """Register email notification handler."""
        
        @self.agent_notification.on_email(**handler_config)
        async def on_email(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    logger.info("📧 EMAIL notification received")
                    
                    if not hasattr(self.agent_instance, "handle_email_notification"):
                        await safe_send_email_response(
                            context, "This agent doesn't support email notifications."
                        )
                        return
                    
                    try:
                        async with asyncio.timeout(self.EMAIL_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_email_notification(
                                notification_activity,
                                self.agent_app.auth,
                                self.auth_handler_name,
                                context,
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Email timeout after {self.EMAIL_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your email. I'm still processing and will follow up."
                    
                    # When the agent uses MCP to send emails, we don't need to reply through
                    # the notification channel. The MCP response IS the email. Just log and return.
                    # Only try notification channel reply if MCP wasn't used (legacy fallback).
                    if response and response.strip():
                        logger.info(f"📧 Email handled: {response[:80]}...")
                        # Skip notification channel reply - agent already sent via MCP
                        # The notification channel has a ~30s timeout and often 404s anyway
                    else:
                        logger.info("📧 No response needed (system notification ignored)")
                    
            except Exception as e:
                logger.error(f"❌ Email notification error: {e}")
                await safe_send_email_response(
                    context, "Thank you for your email. I encountered an issue but will review it."
                )
    
    def _register_word_handler(self, handler_config: dict) -> None:
        """Register Word notification handler."""
        
        @self.agent_notification.on_word(**handler_config)
        async def on_word(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    logger.info("📄 WORD notification received")
                    
                    if not hasattr(self.agent_instance, "handle_word_notification"):
                        await safe_send_activity(context, "This agent doesn't support Word notifications.")
                        return
                    
                    try:
                        async with asyncio.timeout(self.DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_word_notification(
                                notification_activity,
                                self.agent_app.auth,
                                self.auth_handler_name,
                                context,
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Word timeout after {self.DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing."
                    
                    await safe_send_activity(context, response)
                    
            except Exception as e:
                logger.error(f"❌ Word notification error: {e}")
                await safe_send_activity(context, "Thank you for your comment. I encountered an issue.")
    
    def _register_excel_handler(self, handler_config: dict) -> None:
        """Register Excel notification handler."""
        
        @self.agent_notification.on_excel(**handler_config)
        async def on_excel(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    logger.info("📊 EXCEL notification received")
                    
                    if not hasattr(self.agent_instance, "handle_excel_notification"):
                        await safe_send_activity(context, "This agent doesn't support Excel notifications.")
                        return
                    
                    try:
                        async with asyncio.timeout(self.DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_excel_notification(
                                notification_activity,
                                self.agent_app.auth,
                                self.auth_handler_name,
                                context,
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Excel timeout after {self.DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing."
                    
                    await safe_send_activity(context, response)
                    
            except Exception as e:
                logger.error(f"❌ Excel notification error: {e}")
                await safe_send_activity(context, "Thank you for your comment. I encountered an issue.")
    
    def _register_powerpoint_handler(self, handler_config: dict) -> None:
        """Register PowerPoint notification handler."""
        
        @self.agent_notification.on_powerpoint(**handler_config)
        async def on_powerpoint(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    logger.info("📽️ POWERPOINT notification received")
                    
                    if not hasattr(self.agent_instance, "handle_powerpoint_notification"):
                        await safe_send_activity(context, "This agent doesn't support PowerPoint notifications.")
                        return
                    
                    try:
                        async with asyncio.timeout(self.DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_powerpoint_notification(
                                notification_activity,
                                self.agent_app.auth,
                                self.auth_handler_name,
                                context,
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ PowerPoint timeout after {self.DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing."
                    
                    await safe_send_activity(context, response)
                    
            except Exception as e:
                logger.error(f"❌ PowerPoint notification error: {e}")
                await safe_send_activity(context, "Thank you for your comment. I encountered an issue.")
    
    def _register_lifecycle_handler(self, handler_config: dict) -> None:
        """Register lifecycle notification handler."""
        
        @self.agent_notification.on_agent_lifecycle_notification("*", **handler_config)
        async def on_lifecycle(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    logger.info("📋 LIFECYCLE notification received")
                    
                    if hasattr(self.agent_instance, "handle_lifecycle_notification"):
                        response = await self.agent_instance.handle_lifecycle_notification(
                            notification_activity,
                            self.agent_app.auth,
                            self.auth_handler_name,
                            context,
                        )
                        logger.info(f"📋 Lifecycle processed: {response}")
                    
                    # Lifecycle notifications don't send replies
                    
            except Exception as e:
                logger.error(f"❌ Lifecycle notification error: {e}")
    
    def _register_generic_notification_handler(self, handler_config: dict) -> None:
        """Register fallback handler for unhandled notification types."""
        
        @self.agent_notification.on_agent_notification(
            channel_id=ChannelId(channel="agents", sub_channel="*"),
            **handler_config,
        )
        async def on_generic(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    notification_type = notification_activity.notification_type
                    logger.info(f"📬 Generic notification: {notification_type}")
                    
                    notification_text = getattr(notification_activity, 'text', None)
                    if notification_text:
                        await context.send_activity(f"Notification received: {notification_text[:100]}...")
                    else:
                        await context.send_activity(f"Notification of type {notification_type} acknowledged.")
                        
            except Exception as e:
                logger.error(f"❌ Generic notification error: {e}")
    
    def _register_message_handler(self, handler_config: dict) -> None:
        """Register the main message handler."""
        
        @self.agent_app.activity("message", **handler_config)
        async def on_message(context: TurnContext, _: TurnState):
            try:
                result = await self._validate_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result
                
                with ObservabilityContext(tenant_id, agent_id, correlation_id):
                    user_message = context.activity.text or ""
                    if not user_message.strip() or user_message.strip() == "/help":
                        return
                    
                    # Skip Teams system messages
                    if user_message.strip().startswith("<") and any(
                        tag in user_message.lower()
                        for tag in ["<addmember>", "<removemember>", "<topicupdate>", "<historyupdate>"]
                    ):
                        logger.info("🔇 Ignoring Teams system message")
                        return
                    
                    logger.info(f"📨 {user_message}")
                    
                    # Check if MCP needs initialization
                    is_first_request = (
                        hasattr(self.agent_instance, 'mcp_servers_initialized')
                        and not self.agent_instance.mcp_servers_initialized
                    )
                    
                    if is_first_request:
                        logger.info("🔄 First request - MCP initialization required")
                        await context.send_activity(
                            "🔧 **Getting ready!** Connecting to Microsoft 365 services. "
                            "This may take 30-60 seconds..."
                        )
                    
                    # Process with appropriate timeout
                    timeout = self.FIRST_REQUEST_TIMEOUT if is_first_request else self.NORMAL_REQUEST_TIMEOUT
                    
                    try:
                        async with asyncio.timeout(timeout):
                            response = await self.agent_instance.process_user_message(
                                user_message,
                                self.agent_app.auth,
                                self.auth_handler_name,
                                context,
                            )
                            await context.send_activity(response)
                            
                            if is_first_request:
                                logger.info("✅ First request completed - MCP initialized")
                            else:
                                logger.info("✅ Response sent")
                                
                    except asyncio.TimeoutError:
                        logger.warning(f"⏳ Request timed out after {timeout}s")
                        await context.send_activity(
                            "⏳ I'm sorry, your request is taking too long. "
                            "Please try a simpler query."
                        )
                        
            except Exception as e:
                logger.error(f"❌ Error: {e}")
                await context.send_activity(f"Sorry, I encountered an error: {str(e)}")
    
    # =========================================================================
    # AGENT LIFECYCLE
    # =========================================================================
    
    async def initialize_agent(self) -> None:
        """Initialize the agent instance."""
        if self.agent_instance is None:
            logger.info(f"🤖 Initializing {self.agent_class.__name__}...")
            self.agent_instance = self.agent_class(*self.agent_args, **self.agent_kwargs)
            await self.agent_instance.initialize()
            logger.info(f"✅ {self.agent_class.__name__} initialized")
    
    async def cleanup(self) -> None:
        """Clean up resources."""
        if self.agent_instance:
            try:
                await self.agent_instance.cleanup()
            except Exception as e:
                logger.error(f"Cleanup error: {e}")
    
    # =========================================================================
    # AUTHENTICATION
    # =========================================================================
    
    def create_auth_configuration(self) -> Optional[AgentAuthConfiguration]:
        """Create authentication configuration from settings."""
        settings = self.settings.agent_auth
        
        if settings.is_valid:
            logger.info("🔒 Using Client Credentials authentication")
            return AgentAuthConfiguration(
                client_id=settings.client_id,
                tenant_id=settings.tenant_id,
                client_secret=settings.client_secret,
                scopes=settings.scopes_list,
            )
        
        if self.settings.bearer_token:
            logger.info("🔑 Anonymous dev mode (bearer token)")
        else:
            logger.warning("⚠️ No auth configured; running anonymous")
        
        return None
    
    # =========================================================================
    # SERVER
    # =========================================================================
    
    def start_server(self, auth_configuration: Optional[AgentAuthConfiguration] = None) -> None:
        """Start the HTTP server."""
        
        async def entry_point(req: Request) -> Response:
            return await start_agent_process(req, req.app["agent_app"], req.app["adapter"])
        
        async def health(_req: Request) -> Response:
            mcp_ready = False
            if self.agent_instance and hasattr(self.agent_instance, 'mcp_servers_initialized'):
                mcp_ready = self.agent_instance.mcp_servers_initialized
            
            return json_response({
                "status": "ok",
                "agent_type": self.agent_class.__name__,
                "agent_initialized": self.agent_instance is not None,
                "mcp_ready": mcp_ready,
            })
        
        # Setup middlewares
        middlewares = []
        if auth_configuration:
            middlewares.append(jwt_authorization_middleware)
        
        @web_middleware
        async def anonymous_claims(request, handler):
            if not auth_configuration:
                request["claims_identity"] = ClaimsIdentity(
                    {
                        AuthenticationConstants.AUDIENCE_CLAIM: "anonymous",
                        AuthenticationConstants.APP_ID_CLAIM: "anonymous-app",
                    },
                    False,
                    "Anonymous",
                )
            return await handler(request)
        
        middlewares.append(anonymous_claims)
        
        # Create application
        app = Application(middlewares=middlewares)
        app.router.add_post("/api/messages", entry_point)
        app.router.add_get("/api/messages", lambda _: Response(status=200))
        app.router.add_get("/api/health", health)
        
        app["agent_configuration"] = auth_configuration
        app["agent_app"] = self.agent_app
        app["adapter"] = self.agent_app.adapter
        
        app.on_startup.append(lambda app: self.initialize_agent())
        app.on_shutdown.append(lambda app: self.cleanup())
        
        # Find available port
        port = self.settings.server.port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            if s.connect_ex(("127.0.0.1", port)) == 0:
                port = port + 1
        
        # Print startup banner
        print("=" * 80)
        print(f"🏢 {self.agent_class.__name__}")
        print("=" * 80)
        print(f"🔒 Auth: {'Enabled' if auth_configuration else 'Anonymous'}")
        print(f"🚀 Server: localhost:{port}")
        print(f"📚 Endpoint: http://localhost:{port}/api/messages")
        print(f"❤️  Health: http://localhost:{port}/api/health\n")
        
        try:
            run_app(app, host="localhost", port=port, handle_signals=True)
        except KeyboardInterrupt:
            print("\n👋 Server stopped")

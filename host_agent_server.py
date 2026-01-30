# Copyright (c) Microsoft. All rights reserved.

"""Generic Agent Host Server - Hosts agents implementing AgentInterface"""

# --- Imports ---
import asyncio
import logging
import os
import socket
import uuid
from os import environ

from aiohttp.web import Application, Request, Response, json_response, run_app
from aiohttp.client_exceptions import ClientResponseError
from aiohttp.web_middlewares import middleware as web_middleware
from dotenv import load_dotenv
from azure.identity import ClientSecretCredential
from agent_interface import AgentInterface, check_agent_inheritance
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
from microsoft_agents_a365.notifications.agent_notification import (
    AgentNotification,
    NotificationTypes,
    AgentNotificationActivity,
    ChannelId,
)
from microsoft_agents_a365.notifications import EmailResponse

from microsoft_agents_a365.observability.core.config import configure
from microsoft_agents_a365.observability.core.middleware.baggage_builder import (
    BaggageBuilder,
)
from microsoft_agents_a365.runtime.environment_utils import (
    get_observability_authentication_scope,
)
from microsoft_agents_a365.tooling.utils.utility import get_mcp_platform_authentication_scope
from token_cache import cache_agentic_token, get_cached_agentic_token

# NOTE: BackgroundTaskManager removed because Agentic applications CANNOT send proactive messages.
# The adapter.continue_conversation() method requires app-only tokens (client credentials),
# but Agentic apps get AADSTS82001: "not permitted to request app-only tokens".
# 
# For non-agentic auth modes (e.g., BEARER_TOKEN), proactive messaging would work.
# If you need background task + proactive messaging, use a non-agentic auth mode.


# --- Configuration ---
ms_agents_logger = logging.getLogger("microsoft_agents")
ms_agents_logger.addHandler(logging.StreamHandler())
ms_agents_logger.setLevel(logging.INFO)

observability_logger = logging.getLogger("microsoft_agents_a365.observability")
observability_logger.setLevel(logging.ERROR)

# Suppress verbose Azure Identity HTTP logging during startup token acquisition
azure_http_logger = logging.getLogger("azure.core.pipeline.policies.http_logging_policy")
azure_http_logger.setLevel(logging.WARNING)

# Suppress verbose Azure Identity token warnings (we handle failures gracefully)
azure_identity_logger = logging.getLogger("azure.identity")
azure_identity_logger.setLevel(logging.ERROR)

logger = logging.getLogger(__name__)

load_dotenv()
agents_sdk_config = load_configuration_from_env(environ)


# --- Token Resolver for Observability ---
def observability_token_resolver(agent_id: str, tenant_id: str) -> str | None:
    """
    Token resolver for Agent 365 Observability exporter.
    
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


# --- Public API ---
def create_and_run_host(
    agent_class: type[AgentInterface], *agent_args, **agent_kwargs
):
    """Create and run a generic agent host"""
    if not check_agent_inheritance(agent_class):
        raise TypeError(
            f"Agent class {agent_class.__name__} must inherit from AgentInterface"
        )

    configure(
        service_name="AgentFrameworkTracingWithAzureOpenAI",
        service_namespace="AgentFrameworkTesting",
        token_resolver=observability_token_resolver,
    )

    host = GenericAgentHost(agent_class, *agent_args, **agent_kwargs)
    auth_config = host.create_auth_configuration()
    host.start_server(auth_config)


# --- Generic Agent Host ---
class GenericAgentHost:
    """Generic host for agents implementing AgentInterface.
    
    NOTE: Agentic applications have platform limitations:
    - Cannot acquire app-only tokens (AADSTS82001)
    - Cannot send proactive messages (requires app-only token)
    - Cannot pre-initialize MCP servers at startup (requires user token)
    
    As a result, all processing must complete within the HTTP request lifecycle.
    """

    # --- Initialization ---
    def __init__(self, agent_class: type[AgentInterface], *agent_args, **agent_kwargs):
        if not check_agent_inheritance(agent_class):
            raise TypeError(
                f"Agent class {agent_class.__name__} must inherit from AgentInterface"
            )

        # Auth handler name can be configured via environment
        # Defaults to empty (no auth handler) - set AUTH_HANDLER_NAME=AGENTIC for production agentic auth
        self.auth_handler_name = os.getenv("AUTH_HANDLER_NAME", "") or None
        if self.auth_handler_name:
            logger.info(f"🔐 Using auth handler: {self.auth_handler_name}")
        else:
            logger.info("🔓 No auth handler configured (AUTH_HANDLER_NAME not set)")

        self.agent_class = agent_class
        self.agent_args = agent_args
        self.agent_kwargs = agent_kwargs
        self.agent_instance = None

        self.storage = MemoryStorage()
        self.connection_manager = MsalConnectionManager(**agents_sdk_config)
        self.adapter = CloudAdapter(connection_manager=self.connection_manager)
        
        # Get agent app ID (Blueprint ID)
        self.agent_app_id = os.getenv("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID", "")
        
        self.authorization = Authorization(
            self.storage, self.connection_manager, **agents_sdk_config
        )
        self.agent_app = AgentApplication[TurnState](
            storage=self.storage,
            adapter=self.adapter,
            authorization=self.authorization,
            **agents_sdk_config,
        )
        self.agent_notification = AgentNotification(self.agent_app)
        self._setup_handlers()
        logger.info("✅ Notification handlers registered successfully")

    # --- Observability ---
    async def _setup_observability_token(
        self, context: TurnContext, tenant_id: str, agent_id: str
    ):
        # Only attempt token exchange when auth handler is configured
        if not self.auth_handler_name:
            logger.debug("Skipping observability token exchange (no auth handler)")
            return
            
        try:
            logger.info(
                f"🔐 Attempting token exchange for observability... "
                f"(tenant_id={tenant_id}, agent_id={agent_id})"
            )
            exaau_token = await self.agent_app.auth.exchange_token(
                context,
                scopes=get_observability_authentication_scope(),
                auth_handler_id=self.auth_handler_name,
            )
            
            # Validate that we actually got a token (SDK may return None or empty on consent errors)
            if exaau_token and hasattr(exaau_token, 'token') and exaau_token.token:
                cache_agentic_token(tenant_id, agent_id, exaau_token.token)
                logger.info(
                    f"✅ Token exchange successful "
                    f"(tenant_id={tenant_id}, agent_id={agent_id})"
                )
            else:
                logger.warning(
                    f"⚠️ Token exchange returned no token - observability may be limited "
                    f"(tenant_id={tenant_id}, agent_id={agent_id})"
                )
        except Exception as e:
            logger.warning(f"⚠️ Failed to cache observability token: {e}")

    async def _validate_agent_and_setup_context(self, context: TurnContext):
        """Validate agent instance and setup observability context.
        
        Returns:
            Tuple of (tenant_id, agent_id, correlation_id) or None if validation fails
        """
        logger.info("🔍 Validating agent and setting up context...")
        tenant_id = context.activity.recipient.tenant_id
        agent_id = context.activity.recipient.agentic_app_id
        
        # Generate correlation_id from activity.id or create a new UUID
        correlation_id = context.activity.id or str(uuid.uuid4())
        
        logger.info(f"🔍 tenant_id={tenant_id}, agent_id={agent_id}, correlation_id={correlation_id}")

        if not self.agent_instance:
            logger.error("Agent not available")
            await context.send_activity("❌ Sorry, the agent is not available.")
            return None

        await self._setup_observability_token(context, tenant_id, agent_id)
        return tenant_id, agent_id, correlation_id

    # --- Handlers (Messages & Notifications) ---
    def _setup_handlers(self):
        """Setup message and notification handlers"""
        # Configure auth handlers - only required when auth_handler_name is set
        handler_config = {"auth_handlers": [self.auth_handler_name]} if self.auth_handler_name else {}

        async def help_handler(context: TurnContext, _: TurnState):
            await context.send_activity(
                f"👋 **Hi there!** I'm **{self.agent_class.__name__}**, your AI assistant.\n\n"
                "How can I help you today?"
            )

        self.agent_app.conversation_update("membersAdded", **handler_config)(help_handler)
        self.agent_app.message("/help", **handler_config)(help_handler)

        # IMPORTANT: Register notification handlers BEFORE message handler!
        # The SDK uses "first match wins" routing, so notification handlers must come first
        # to catch activities with channelId="agents" before the generic message handler.

        # =====================================================================
        # EMAIL NOTIFICATION HANDLER
        # =====================================================================
        # Timeout for email notification processing (shorter than Teams due to email channel limits)
        EMAIL_NOTIFICATION_TIMEOUT = 25  # seconds - email channel typically times out at ~30s

        @self.agent_notification.on_email(**handler_config)
        async def on_email_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    logger.info("📧 EMAIL notification received")

                    if not hasattr(self.agent_instance, "handle_email_notification"):
                        logger.warning("⚠️ Agent doesn't support email notifications")
                        await self._safe_send_email_response(context, "This agent doesn't support email notifications yet.")
                        return

                    # Process with timeout to avoid notification channel timeout
                    try:
                        async with asyncio.timeout(EMAIL_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_email_notification(
                                notification_activity, self.agent_app.auth, self.auth_handler_name, context
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Email processing timeout after {EMAIL_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your email. I'm still processing your request and will follow up shortly."

                    # Email responses use special EmailResponse format
                    await self._safe_send_email_response(context, response)

            except Exception as e:
                logger.error(f"❌ Email notification error: {e}")
                await self._safe_send_email_response(context, "Thank you for your email. I encountered an issue but will review it.")

        async def _safe_send_email_response(context: TurnContext, response: str):
            """Safely send email response, handling 404 errors gracefully.
            
            404 errors occur when the notification channel times out before we can respond.
            In this case, we log the issue but don't crash - the email was still processed.
            """
            try:
                response_activity = EmailResponse.create_email_response_activity(response)
                await context.send_activity(response_activity)
                logger.info("✅ Email response sent successfully")
            except ClientResponseError as e:
                if e.status == 404:
                    logger.warning(
                        f"⚠️ Email reply window expired (404). Response was: {response[:100]}... "
                        f"The notification channel timed out, but the email was processed."
                    )
                else:
                    logger.error(f"❌ Failed to send email response: {e}")
                    raise
            except Exception as e:
                logger.error(f"❌ Unexpected error sending email response: {e}")

        # =====================================================================
        # WORD NOTIFICATION HANDLER
        # =====================================================================
        # Timeout for document notification processing
        DOC_NOTIFICATION_TIMEOUT = 25  # seconds
        
        @self.agent_notification.on_word(**handler_config)
        async def on_word_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    logger.info("📄 WORD comment notification received")

                    if not hasattr(self.agent_instance, "handle_word_notification"):
                        logger.warning("⚠️ Agent doesn't support Word notifications")
                        await _safe_send_activity(context, "This agent doesn't support Word comment notifications yet.")
                        return

                    try:
                        async with asyncio.timeout(DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_word_notification(
                                notification_activity, self.agent_app.auth, self.auth_handler_name, context
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Word notification timeout after {DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing and will respond shortly."

                    await _safe_send_activity(context, response)

            except Exception as e:
                logger.error(f"❌ Word notification error: {e}")
                await _safe_send_activity(context, "Thank you for your comment. I encountered an issue but will review it.")

        async def _safe_send_activity(context: TurnContext, message: str):
            """Safely send activity, handling 404 errors gracefully."""
            try:
                await context.send_activity(message)
                logger.info("✅ Activity sent successfully")
            except ClientResponseError as e:
                if e.status == 404:
                    logger.warning(f"⚠️ Reply window expired (404). Message was: {message[:100]}...")
                else:
                    logger.error(f"❌ Failed to send activity: {e}")
            except Exception as e:
                logger.error(f"❌ Unexpected error sending activity: {e}")

        # =====================================================================
        # EXCEL NOTIFICATION HANDLER
        # =====================================================================
        @self.agent_notification.on_excel(**handler_config)
        async def on_excel_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    logger.info("📊 EXCEL comment notification received")

                    if not hasattr(self.agent_instance, "handle_excel_notification"):
                        logger.warning("⚠️ Agent doesn't support Excel notifications")
                        await _safe_send_activity(context, "This agent doesn't support Excel comment notifications yet.")
                        return

                    try:
                        async with asyncio.timeout(DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_excel_notification(
                                notification_activity, self.agent_app.auth, self.auth_handler_name, context
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ Excel notification timeout after {DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing and will respond shortly."

                    await _safe_send_activity(context, response)

            except Exception as e:
                logger.error(f"❌ Excel notification error: {e}")
                await _safe_send_activity(context, "Thank you for your comment. I encountered an issue but will review it.")

        # =====================================================================
        # POWERPOINT NOTIFICATION HANDLER
        # =====================================================================
        @self.agent_notification.on_powerpoint(**handler_config)
        async def on_powerpoint_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    logger.info("📽️ POWERPOINT comment notification received")

                    if not hasattr(self.agent_instance, "handle_powerpoint_notification"):
                        logger.warning("⚠️ Agent doesn't support PowerPoint notifications")
                        await _safe_send_activity(context, "This agent doesn't support PowerPoint comment notifications yet.")
                        return

                    try:
                        async with asyncio.timeout(DOC_NOTIFICATION_TIMEOUT):
                            response = await self.agent_instance.handle_powerpoint_notification(
                                notification_activity, self.agent_app.auth, self.auth_handler_name, context
                            )
                    except asyncio.TimeoutError:
                        logger.warning(f"⚠️ PowerPoint notification timeout after {DOC_NOTIFICATION_TIMEOUT}s")
                        response = "Thank you for your comment. I'm still processing and will respond shortly."

                    await _safe_send_activity(context, response)

            except Exception as e:
                logger.error(f"❌ PowerPoint notification error: {e}")
                await _safe_send_activity(context, "Thank you for your comment. I encountered an issue but will review it.")

        # =====================================================================
        # LIFECYCLE NOTIFICATION HANDLER
        # =====================================================================
        @self.agent_notification.on_agent_lifecycle_notification("*", **handler_config)
        async def on_lifecycle_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    logger.info("📋 LIFECYCLE notification received")

                    if not hasattr(self.agent_instance, "handle_lifecycle_notification"):
                        logger.warning("⚠️ Agent doesn't support lifecycle notifications")
                        return  # Lifecycle notifications don't need a response

                    response = await self.agent_instance.handle_lifecycle_notification(
                        notification_activity, self.agent_app.auth, self.auth_handler_name, context
                    )

                    # Lifecycle notifications don't send replies (onboarding channel doesn't support them)
                    logger.info(f"📋 Lifecycle notification processed: {response}")

            except Exception as e:
                logger.error(f"❌ Lifecycle notification error: {e}")
                # Don't send activity for lifecycle errors - channel may not support it

        # =====================================================================
        # FALLBACK: GENERIC NOTIFICATION HANDLER (for any unhandled types)
        # =====================================================================
        @self.agent_notification.on_agent_notification(
            channel_id=ChannelId(channel="agents", sub_channel="*"),
            **handler_config,
        )
        async def on_generic_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    notification_type = notification_activity.notification_type
                    logger.info(f"📬 Generic notification received: {notification_type}")

                    # This is a fallback for any notification types not explicitly handled above
                    notification_text = getattr(notification_activity, 'text', None)
                    if notification_text:
                        await context.send_activity(f"Notification received: {notification_text[:100]}...")
                    else:
                        await context.send_activity(f"Notification of type {notification_type} acknowledged.")

            except Exception as e:
                logger.error(f"❌ Generic notification error: {e}")
                await context.send_activity(f"Sorry, I encountered an error processing the notification: {str(e)}")

        # Message handler comes AFTER notification handler
        # NOTE: Agentic apps CANNOT use proactive messaging (AADSTS82001 - can't get app-only tokens)
        # So we must wait for responses to complete within the HTTP request lifecycle.
        @self.agent_app.activity("message", **handler_config)
        async def on_message(context: TurnContext, _: TurnState):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id, correlation_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).correlation_id(correlation_id).build():
                    user_message = context.activity.text or ""
                    if not user_message.strip() or user_message.strip() == "/help":
                        return

                    # Skip Teams system messages (roster changes, etc.)
                    if user_message.strip().startswith("<") and any(
                        tag in user_message.lower() 
                        for tag in ["<addmember>", "<removemember>", "<topicupdate>", "<historyupdate>"]
                    ):
                        logger.info("🔇 Ignoring Teams system message")
                        return

                    logger.info(f"📨 {user_message}")
                    
                    # =============================================================
                    # CHECK IF MCP NEEDS FIRST-REQUEST INITIALIZATION
                    # =============================================================
                    # In production with agentic auth, MCPs can only init on first request
                    # because client credentials can't get MCP tokens (platform restriction)
                    is_first_request_init = (
                        hasattr(self.agent_instance, 'mcp_servers_initialized') and 
                        not self.agent_instance.mcp_servers_initialized
                    )
                    
                    if is_first_request_init:
                        logger.info("🔄 First request - MCP initialization required (agentic auth)")
                        await context.send_activity(
                            "🔧 **Getting ready!** Connecting to Microsoft 365 services for the first time. "
                            "This may take 30-60 seconds, but I'll be much faster after that!"
                        )
                    
                    # =============================================================
                    # AGENTIC AUTH LIMITATION: NO PROACTIVE MESSAGING
                    # =============================================================
                    # Agentic applications CANNOT send proactive messages because they
                    # cannot acquire app-only tokens (AADSTS82001 error).
                    # 
                    # Proactive messaging requires: adapter.continue_conversation()
                    # Which internally uses: Confidential Client Application (app-only token)
                    # Which is blocked for: Agentic apps
                    #
                    # SOLUTION: We must complete ALL processing within the HTTP request.
                    # Use a generous timeout and let Teams show "typing..." indicator.
                    # If we timeout, we apologize but cannot deliver the result later.
                    
                    # Generous timeout: 2 minutes for first request (MCP init), 90s for normal
                    processing_timeout = 120 if is_first_request_init else 90
                    
                    try:
                        async with asyncio.timeout(processing_timeout):
                            response = await self.agent_instance.process_user_message(
                                user_message, self.agent_app.auth, self.auth_handler_name, context
                            )
                            await context.send_activity(response)
                            
                            if is_first_request_init:
                                logger.info("✅ First request completed - MCP servers now initialized")
                            else:
                                logger.info("✅ Response sent")
                                
                    except asyncio.TimeoutError:
                        logger.warning(f"⏳ Request timed out after {processing_timeout}s")
                        await context.send_activity(
                            "⏳ I'm sorry, your request is taking longer than expected. "
                            "Please try again with a simpler query, or try again in a moment."
                        )

            except Exception as e:
                logger.error(f"❌ Error: {e}")
                await context.send_activity(f"Sorry, I encountered an error: {str(e)}")

    # --- Agent and MCP Initialization at Startup ---
    async def initialize_agent_and_mcp(self):
        """Initialize agent AND MCP servers at startup - everything ready before first request."""
        if self.agent_instance is None:
            logger.info(f"🤖 Initializing {self.agent_class.__name__}...")
            self.agent_instance = self.agent_class(*self.agent_args, **self.agent_kwargs)
            await self.agent_instance.initialize()
        
        # Now initialize MCP servers at startup
        await self._initialize_mcp_at_startup()
    
    async def _initialize_mcp_at_startup(self):
        """Pre-initialize MCP servers during server startup using client credentials."""
        logger.info("=" * 60)
        logger.info("🔧 INITIALIZING MCP SERVERS AT STARTUP")
        logger.info("=" * 60)
        
        try:
            # Get client credentials from environment
            client_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID") or environ.get("CLIENT_ID")
            client_secret = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET") or environ.get("CLIENT_SECRET")
            tenant_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID") or environ.get("TENANT_ID")
            
            if not all([client_id, client_secret, tenant_id]):
                logger.warning("⚠️ Missing client credentials - MCP init will happen on first request")
                return
            
            # Acquire token for MCP platform using client credentials
            logger.info("🔐 Acquiring token for MCP servers using client credentials...")
            startup_token = await self._acquire_mcp_token_with_client_credentials(
                client_id, client_secret, tenant_id
            )
            
            if not startup_token:
                logger.warning("⚠️ Could not acquire startup token - MCP init will happen on first request")
                return
            
            # Create a startup context for MCP initialization
            startup_context = await self._create_startup_context()
            
            if startup_context is None:
                logger.warning("⚠️ Could not create startup context - MCP init will happen on first request")
                return
            
            # Initialize MCP servers with the acquired token
            await self.agent_instance.startup_initialize_mcp(
                auth=self.agent_app.auth,
                auth_handler_name=self.auth_handler_name,
                context=startup_context,
                auth_token=startup_token,  # Pass the token we acquired
            )
            
            logger.info("=" * 60)
            logger.info("✅ MCP SERVERS READY - Agent fully operational!")
            logger.info("=" * 60)
            
        except Exception as e:
            logger.error(f"❌ MCP startup initialization failed: {e}")
            logger.warning("⚠️ Agent will start but MCP tools may not be available")
    
    async def _acquire_mcp_token_with_client_credentials(
        self, client_id: str, client_secret: str, tenant_id: str
    ) -> str | None:
        """Acquire a token for MCP platform using client credentials flow."""
        try:
            # Get the MCP platform scope
            mcp_scopes = get_mcp_platform_authentication_scope()
            logger.info(f"🔐 Requesting token for scopes: {mcp_scopes}")
            
            # Use Azure Identity ClientSecretCredential
            credential = ClientSecretCredential(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
            )
            
            # Get token - convert scope list to string format Azure Identity expects
            token = credential.get_token(*mcp_scopes)
            
            logger.info("✅ Successfully acquired MCP platform token")
            return token.token
            
        except Exception as e:
            logger.error(f"❌ Failed to acquire MCP token: {e}")
            return None
    
    async def _create_startup_context(self):
        """Create a minimal TurnContext for startup MCP initialization."""
        try:
            # Get agent identity from environment
            client_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID") or environ.get("CLIENT_ID")
            tenant_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID") or environ.get("TENANT_ID")
            
            if not client_id:
                logger.warning("⚠️ No CLIENT_ID available for startup context")
                return None
            
            # Create a minimal activity for the startup context
            startup_activity = Activity(
                type=ActivityTypes.event,
                channel_id="startup",
                service_url="https://api.botframework.com",
                recipient={
                    "id": client_id,
                    "agentic_app_id": client_id,
                    "tenant_id": tenant_id or "unknown",
                },
                from_property={
                    "id": "startup-init",
                    "name": "Startup Initialization",
                },
                conversation={
                    "id": "startup-init-conversation",
                },
            )
            
            # Create TurnContext using the adapter
            turn_context = TurnContext(self.adapter, startup_activity)
            
            logger.info(f"✅ Created startup context (agent_id={client_id})")
            return turn_context
            
        except Exception as e:
            logger.error(f"❌ Failed to create startup context: {e}")
            return None

    # Legacy method for compatibility
    async def initialize_agent(self):
        """Initialize agent only (legacy - use initialize_agent_and_mcp instead)."""
        if self.agent_instance is None:
            logger.info(f"🤖 Initializing {self.agent_class.__name__}...")
            self.agent_instance = self.agent_class(*self.agent_args, **self.agent_kwargs)
            await self.agent_instance.initialize()

    # --- Authentication ---
    def create_auth_configuration(self) -> AgentAuthConfiguration | None:
        # First try the consolidated CONNECTIONS__SERVICE_CONNECTION env vars (preferred)
        client_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID")
        tenant_id = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID")
        client_secret = environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET")
        
        # Fall back to legacy env vars if CONNECTIONS vars not set
        if not client_id:
            client_id = environ.get("CLIENT_ID")
        if not tenant_id:
            tenant_id = environ.get("TENANT_ID")
        if not client_secret:
            client_secret = environ.get("CLIENT_SECRET")

        if client_id and tenant_id and client_secret:
            logger.info("🔒 Using Client Credentials authentication")
            return AgentAuthConfiguration(
                client_id=client_id,
                tenant_id=tenant_id,
                client_secret=client_secret,
                scopes=["5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default"],
            )

        if environ.get("BEARER_TOKEN"):
            logger.info("🔑 Anonymous dev mode")
        else:
            logger.warning("⚠️ No auth env vars; running anonymous")
        return None

    # --- Server ---
    def start_server(self, auth_configuration: AgentAuthConfiguration | None = None):
        async def entry_point(req: Request) -> Response:
            return await start_agent_process(
                req, req.app["agent_app"], req.app["adapter"]
            )

        async def health(_req: Request) -> Response:
            # Include MCP initialization status in health check
            mcp_ready = False
            if self.agent_instance and hasattr(self.agent_instance, 'mcp_servers_initialized'):
                mcp_ready = self.agent_instance.mcp_servers_initialized
            
            return json_response(
                {
                    "status": "ok",
                    "agent_type": self.agent_class.__name__,
                    "agent_initialized": self.agent_instance is not None,
                    "mcp_ready": mcp_ready,
                    "background_tasks": self.background_tasks.active_tasks if self.background_tasks else 0,
                }
            )

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
        app = Application(middlewares=middlewares)

        app.router.add_post("/api/messages", entry_point)
        app.router.add_get("/api/messages", lambda _: Response(status=200))
        app.router.add_get("/api/health", health)

        app["agent_configuration"] = auth_configuration
        app["agent_app"] = self.agent_app
        app["adapter"] = self.agent_app.adapter

        # Initialize agent AND MCP servers at startup - everything ready before first request
        app.on_startup.append(lambda app: self.initialize_agent_and_mcp())
        app.on_shutdown.append(lambda app: self.cleanup())

        desired_port = int(environ.get("PORT", 3978))
        port = desired_port

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(0.5)
            if s.connect_ex(("127.0.0.1", desired_port)) == 0:
                port = desired_port + 1

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

    # --- Cleanup ---
    async def cleanup(self):
        # Cancel any running background tasks
        if self.background_tasks:
            await self.background_tasks.cleanup()
            
        if self.agent_instance:
            try:
                await self.agent_instance.cleanup()
            except Exception as e:
                logger.error(f"Cleanup error: {e}")




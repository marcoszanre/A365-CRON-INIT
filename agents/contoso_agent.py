# Copyright (c) Microsoft. All rights reserved.

"""
Contoso Agent

A full-featured AI agent for Contoso organization integrated with Microsoft 365.
Uses Azure OpenAI and MCP servers for extended functionality.
"""

import asyncio
import logging
from pathlib import Path
from typing import Optional

from a365_agent.auth import LocalAuthOptions
from a365_agent.base import AgentBase
from a365_agent.config import get_settings, AzureOpenAIModelConfig
from a365_agent.mcp import MCPService
from a365_agent.observability import enable_agentframework_instrumentation

from agent_framework import ChatAgent
from agent_framework.azure import AzureOpenAIChatClient
from microsoft_agents.hosting.core import Authorization, TurnContext

logger = logging.getLogger(__name__)


def _load_system_prompt(is_dev_mode: bool = False) -> str:
    """
    Load the system prompt from the separate markdown file.
    
    Args:
        is_dev_mode: If True, includes DEV_ONLY sections and excludes PROD_ONLY sections.
                     If False (default), includes PROD_ONLY sections and excludes DEV_ONLY sections.
    
    Returns:
        The processed system prompt with appropriate sections included/excluded.
    """
    prompt_path = Path(__file__).parent / "system_prompt.md"
    try:
        content = prompt_path.read_text(encoding="utf-8")
        
        # Process conditional sections based on mode
        import re
        
        if is_dev_mode:
            # Dev mode: Remove PROD_ONLY sections, keep DEV_ONLY content
            content = re.sub(
                r'\{\{PROD_ONLY_START\}\}.*?\{\{PROD_ONLY_END\}\}',
                '',
                content,
                flags=re.DOTALL
            )
            # Remove DEV_ONLY markers but keep content
            content = content.replace('{{DEV_ONLY_START}}', '')
            content = content.replace('{{DEV_ONLY_END}}', '')
            logger.info("📝 Loaded DEV mode system prompt (no chat history retrieval)")
        else:
            # Prod mode: Remove DEV_ONLY sections, keep PROD_ONLY content
            content = re.sub(
                r'\{\{DEV_ONLY_START\}\}.*?\{\{DEV_ONLY_END\}\}',
                '',
                content,
                flags=re.DOTALL
            )
            # Remove PROD_ONLY markers but keep content
            content = content.replace('{{PROD_ONLY_START}}', '')
            content = content.replace('{{PROD_ONLY_END}}', '')
            logger.info("📝 Loaded PROD mode system prompt (with chat history retrieval)")
        
        # Clean up extra blank lines
        content = re.sub(r'\n{3,}', '\n\n', content)
        
        return content.strip()
    except FileNotFoundError:
        logger.error(f"System prompt file not found: {prompt_path}")
        raise
    except Exception as e:
        logger.error(f"Error loading system prompt: {e}")
        raise


def _is_dev_mode() -> bool:
    """Check if running in dev mode based on AGENT_MODE env var."""
    import os
    agent_mode = os.environ.get("AGENT_MODE", "prod").lower()
    return agent_mode == "dev"


class ContosoAgent(AgentBase):
    """
    AI-powered colleague for the Contoso organization in Microsoft 365.
    
    Features:
    - Azure OpenAI integration for intelligent conversations
    - MCP server integration for M365 tool access (email, calendar, Teams, etc.)
    - Observability with Agent 365 telemetry
    - Notification handling (email, Word, Excel, PowerPoint, lifecycle)
    """
    
    # Agent system prompt - loaded from separate file for easier maintenance
    # Mode is determined by AGENT_MODE env var (dev or prod)
    AGENT_INSTRUCTIONS = _load_system_prompt(is_dev_mode=_is_dev_mode())

    # Processing timeout (seconds)
    PROCESSING_TIMEOUT = 120  # 2 minutes max for complex tasks with MCP
    EMAIL_PROCESSING_TIMEOUT = 60  # Email needs time for MCP tools
    
    def __init__(self):
        """Initialize the Contoso Agent."""
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # Load settings
        self.settings = get_settings()
        self.auth_options = LocalAuthOptions.from_environment()
        
        # Enable instrumentation
        enable_agentframework_instrumentation()
        
        # Current model tracking for failover
        self.current_model: Optional[AzureOpenAIModelConfig] = None
        
        # Initialize components
        self._create_chat_client()
        self._create_agent()
        
        # MCP service (lazy initialization)
        self.mcp_service = MCPService()
        
        # Track MCP initialization state
        self.mcp_servers_initialized = False
    
    def _create_chat_client(self, model_config: Optional[AzureOpenAIModelConfig] = None) -> None:
        """
        Create the Azure OpenAI chat client with retry configuration.
        
        Args:
            model_config: Optional specific model to use. If None, uses model pool.
        """
        if model_config:
            # Use specific model config (for failover)
            self.current_model = model_config
        elif self.settings.model_pool and len(self.settings.model_pool) > 0:
            # Use model pool for load balancing
            self.current_model = self.settings.model_pool.get_next_model()
        else:
            # Fallback to legacy single-model config
            settings = self.settings.azure_openai
            settings.validate()
            self.current_model = AzureOpenAIModelConfig(
                endpoint=settings.endpoint,
                deployment=settings.deployment,
                api_key=settings.api_key or "",
                api_version=settings.api_version,
            )
        
        # Create the chat client with API key authentication
        self.chat_client = AzureOpenAIChatClient(
            endpoint=self.current_model.endpoint,
            api_key=self.current_model.api_key,
            deployment_name=self.current_model.deployment,
            api_version=self.current_model.api_version,
        )
        logger.info(f"🤖 Using model: {self.current_model.name}")
    
    def _create_agent(self) -> None:
        """Create the AgentFramework agent."""
        self.agent = ChatAgent(
            chat_client=self.chat_client,
            instructions=self.AGENT_INSTRUCTIONS,
            tools=[],
        )
        logger.info("✅ ChatAgent created")
    
    async def initialize(self) -> None:
        """Initialize the agent (called at startup)."""
        logger.info("✅ ContosoAgent initialized")
    
    async def _ensure_mcp_initialized(
        self,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> None:
        """Ensure MCP servers are initialized (lazy init on first use)."""
        if self.mcp_servers_initialized:
            return
        
        logger.info("🔧 Initializing MCP servers...")
        
        # Try bearer token first (dev mode), then agentic auth (production)
        if self.auth_options.bearer_token:
            self.agent = await self.mcp_service.initialize_with_bearer_token(
                chat_client=self.chat_client,
                agent_instructions=self.AGENT_INSTRUCTIONS,
                bearer_token=self.auth_options.bearer_token,
                auth=auth,
                auth_handler_name=auth_handler_name,
                turn_context=context,
            ) or self.agent
        else:
            self.agent = await self.mcp_service.initialize_with_agentic_auth(
                chat_client=self.chat_client,
                agent_instructions=self.AGENT_INSTRUCTIONS,
                auth=auth,
                auth_handler_name=auth_handler_name,
                turn_context=context,
            ) or self.agent
        
        self.mcp_servers_initialized = True
        logger.info("✅ MCP servers ready")

    async def _reset_mcp_after_error(self) -> None:
        """Best-effort cleanup for MCP resources after a failure."""
        try:
            await self.mcp_service.cleanup()
        finally:
            self.mcp_servers_initialized = False
    
    def _extract_result(self, result) -> str:
        """Extract text content from agent result."""
        if not result:
            return ""
        if hasattr(result, "contents"):
            return str(result.contents)
        elif hasattr(result, "text"):
            return str(result.text)
        elif hasattr(result, "content"):
            return str(result.content)
        return str(result)

    async def _run_with_failover(self, message: str, max_retries: int = 3) -> str:
        """
        Run agent with automatic failover to other models on rate limiting (429).
        
        Args:
            message: The message to process
            max_retries: Maximum number of failover attempts
            
        Returns:
            The agent's response
        """
        last_error = None
        
        for attempt in range(max_retries):
            try:
                logger.info(f"🤖 Calling agent.run() (attempt {attempt + 1}/{max_retries})...")
                result = await self.agent.run(message)
                logger.info("✅ Agent response received")
                
                # Success - clear any throttle on current model
                if self.settings.model_pool and self.current_model:
                    self.settings.model_pool.clear_throttle(self.current_model)
                
                return self._extract_result(result) or "I couldn't process your request."
                
            except Exception as e:
                error_str = str(e).lower()
                last_error = e
                
                # Check if it's a rate limiting error (429)
                is_rate_limit = (
                    "429" in error_str or 
                    "rate limit" in error_str or 
                    "too many requests" in error_str or
                    "retry" in error_str
                )
                
                if is_rate_limit and self.settings.model_pool and len(self.settings.model_pool) > 1:
                    # Mark current model as throttled
                    if self.current_model:
                        # Extract retry-after if present, default to 60s
                        retry_after = 60.0
                        if "retry" in error_str:
                            # Try to extract seconds from error message
                            import re
                            match = re.search(r'(\d+\.?\d*)\s*second', error_str)
                            if match:
                                retry_after = float(match.group(1))
                        
                        self.settings.model_pool.mark_throttled(self.current_model, retry_after)
                    
                    # Get next available model
                    available = self.settings.model_pool.available_count
                    logger.warning(f"🔄 Rate limited! Failover attempt {attempt + 1}/{max_retries}. Available models: {available}/{len(self.settings.model_pool)}")
                    
                    # Switch to next model
                    self._create_chat_client()
                    self._create_agent()
                    
                    # Small delay before retry
                    await asyncio.sleep(0.5)
                else:
                    # Not a rate limit error, or no failover available
                    raise
        
        # All retries exhausted
        logger.error(f"All {max_retries} failover attempts failed")
        raise last_error or Exception("All models failed")
    
    # =========================================================================
    # MESSAGE PROCESSING
    # =========================================================================
    
    async def process_user_message(
        self,
        message: str,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Process a user message and return a response."""
        try:
            # Ensure MCP is initialized
            await self._ensure_mcp_initialized(auth, auth_handler_name, context)
            
            # In dev mode, skip chat history retrieval (Playground doesn't have real chats)
            is_dev_mode = _is_dev_mode()
            
            # Extract chat ID from context for conversation history retrieval
            chat_id = None
            if not is_dev_mode and context.activity and context.activity.conversation:
                chat_id = getattr(context.activity.conversation, "id", None)
            
            # Build the prompt with chat context for history retrieval (prod mode only)
            if chat_id:
                augmented_message = f"""CHAT CONTEXT:
- Chat ID: {chat_id}
- IMPORTANT: Before answering, call listMessages with this chat ID to get conversation history!

USER MESSAGE:
{message}"""
            else:
                augmented_message = message
            
            # Process with timeout and automatic failover on rate limits
            async with asyncio.timeout(self.PROCESSING_TIMEOUT):
                response = await self._run_with_failover(augmented_message)
            return response
            
        except asyncio.TimeoutError:
            logger.error(f"Processing timeout after {self.PROCESSING_TIMEOUT}s")
            await self._reset_mcp_after_error()
            return "Sorry, the request took too long. Please try again."
        except Exception as e:
            logger.error(f"Error processing message: {e}")
            await self._reset_mcp_after_error()
            return f"Sorry, I encountered an error: {str(e)}"
    
    # =========================================================================
    # NOTIFICATION HANDLERS
    # =========================================================================
    
    def _is_system_generated_email(self, context: TurnContext) -> bool:
        """
        Check if this email is a system-generated notification that should be IGNORED.
        
        System notifications include:
        - Site/document sharing notifications
        - Comment mention notifications (handled separately via Word/Excel/PowerPoint)
        - Calendar invites from the system
        - Any automated Microsoft 365 notification
        """
        subject = ""
        if context.activity.conversation:
            subject = getattr(context.activity.conversation, "topic", "") or ""
        subject_lower = subject.lower()

        text_content = getattr(context.activity, "text", "") or ""
        text_lower = text_content.lower()

        # Get HTML body for pattern matching
        html_body = ""
        entities = getattr(context.activity, "entities", []) or []
        for entity in entities:
            entity_type = getattr(entity, "type", "") if hasattr(entity, "type") else entity.get("type", "")
            if entity_type == "emailNotification":
                if hasattr(entity, "htmlBody"):
                    html_body = entity.htmlBody or ""
                elif isinstance(entity, dict):
                    html_body = entity.get("htmlBody", "") or ""
                break
        html_lower = html_body.lower()

        # Patterns that indicate system-generated notifications
        system_patterns = [
            # Sharing notifications
            "shared with you",
            "compartilhou com você",
            "convidou você para",
            "invited you to",
            "has shared",
            "gave you access",
            "deu acesso",

            # Comment mention notifications (duplicates - handled by Word/Excel/PPT handlers)
            "mentioned you in",
            "mencionou você",
            "go to comment",
            "ir para comentário",

            # Site/Team notifications
            "follow this site",
            "siga este site",
            "you've been added to",
            "você foi adicionado",
            "welcome to the team",

            # Document notifications
            "document is ready",
            "shared a file",
            "shared a folder",
            "compartilhou um arquivo",
            "compartilhou uma pasta",

            # Calendar system notifications (not actual invites from people)
            "your meeting was updated",
            "meeting canceled",
            "reunião foi atualizada",
            "reunião cancelada",
        ]

        # Check all text fields for system patterns
        all_text = f"{subject_lower} {text_lower} {html_lower}"
        for pattern in system_patterns:
            if pattern in all_text:
                return True

        # Check for SharePoint/OneDrive system URLs in HTML (indicates automated notification)
        if html_body:
            sharepoint_patterns = [
                "sharepoint.com/sites/",
                "sharepoint.com/personal/",
                "-my.sharepoint.com/",
                "FollowSite=1",  # SharePoint follow button
            ]
            for pattern in sharepoint_patterns:
                if pattern in html_body and "go to comment" not in html_lower:
                    # If it has SharePoint links but isn't a comment notification
                    # Check if it seems like a sharing/access notification
                    if any(x in all_text for x in ["shared", "compartilh", "access", "acesso", "convid", "invited"]):
                        return True

        return False

    async def handle_email_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Handle email notifications - IGNORE system-generated, process real emails."""
        try:
            logger.info("📧 Processing email notification")
            
            # Check if this is a system-generated notification (shares, mentions, etc.)
            if self._is_system_generated_email(context):
                subject = ""
                if context.activity.conversation:
                    subject = getattr(context.activity.conversation, "topic", "") or ""
                logger.info(f"📧 Ignoring system-generated email notification: '{subject[:50]}...'")
                return ""  # Return empty - don't send any reply
            
            # Extract email data for real emails
            sender_email = ""
            sender_name = ""
            subject = ""
            message_id = ""
            text_content = getattr(context.activity, "text", "") or ""
            html_body = ""
            
            # Get the message ID from activity.id (this is the email message ID)
            message_id = getattr(context.activity, "id", "") or ""
            
            if context.activity.from_property:
                sender_email = getattr(context.activity.from_property, "id", "") or ""
                sender_name = getattr(context.activity.from_property, "name", "") or ""
            
            if context.activity.conversation:
                subject = getattr(context.activity.conversation, "topic", "") or ""
            
            # Get htmlBody from emailNotification entity (and message ID if not already set)
            entities = getattr(context.activity, "entities", []) or []
            for entity in entities:
                entity_type = getattr(entity, "type", "") if hasattr(entity, "type") else entity.get("type", "")
                if entity_type == "emailNotification":
                    if hasattr(entity, "htmlBody"):
                        html_body = entity.htmlBody
                    elif isinstance(entity, dict):
                        html_body = entity.get("htmlBody", "")
                    # Also get message ID from entity if not already set
                    if not message_id:
                        if hasattr(entity, "id"):
                            message_id = entity.id
                        elif isinstance(entity, dict):
                            message_id = entity.get("id", "")
                    break
            
            # Use the best available content
            email_content = html_body[:3000] if html_body else text_content[:3000]
            
            logger.info(f"📧 Real email from {sender_name} ({sender_email}): '{subject[:50]}...'")
            logger.info(f"📧 Message ID: {message_id[:50]}..." if message_id else "📧 No message ID found")
            
            # Initialize MCP for full tool access
            await self._ensure_mcp_initialized(auth, auth_handler_name, context)
            
            # Let the AI decide what to do based on the email content
            message = f"""You received a real email from a person. Analyze it and respond appropriately.

FROM: {sender_name} <{sender_email}>
SUBJECT: {subject}
MESSAGE_ID: {message_id}

EMAIL CONTENT:
{email_content}

INSTRUCTIONS:
- This is a real email from a human, not a system notification
- Analyze what the sender is asking or telling you
- If they're asking a question, answer it directly
- If they're asking you to do something (send email, schedule meeting, look up info, etc.), USE YOUR TOOLS to do it
- To reply to this email:
  * Use **ReplyAllToMessageAsync** if the sender uses words like "us", "we", "team", or if there are CC recipients (default choice)
  * Use ReplyToMessageAsync only if the message is clearly personal/private to the sender alone
  * Pass the MESSAGE_ID above as the 'id' parameter
- If it's just informational with no action needed, acknowledge it briefly
- Be helpful and take action when appropriate

Respond and take any necessary actions."""
            
            try:
                async with asyncio.timeout(self.PROCESSING_TIMEOUT):
                    response = await self._run_with_failover(message)
                    
                logger.info(f"📧 Email processed: {response[:100] if response else 'No response'}...")
                return response or "Email processed."
                
            except asyncio.TimeoutError:
                logger.warning("Email processing timeout")
                return "Thank you for your email. I've received it and will review it shortly."
            
        except Exception as e:
            logger.error(f"Email notification error: {e}")
            return "Thank you for your email. I encountered an issue but will review it."
    
    async def handle_word_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Handle Word document comment notifications - let AI decide what to do."""
        try:
            logger.info("📄 Processing Word notification")
            
            # Initialize MCP for full tool access - the user might ask for anything!
            await self._ensure_mcp_initialized(auth, auth_handler_name, context)
            
            # Get the comment text and context
            comment_text = getattr(context.activity, "text", "") or ""
            comment_text = comment_text.replace("<at>", "").replace("</at>", "").strip()
            
            doc_name = getattr(context.activity.conversation, "topic", "") or "Document"
            sender_name = ""
            if context.activity.from_property:
                sender_name = getattr(context.activity.from_property, "name", "") or ""
            
            logger.info(f"📄 Word comment from {sender_name}: '{comment_text[:50]}...'")
            
            async with asyncio.timeout(self.PROCESSING_TIMEOUT):
                message = f"""Someone commented on a Word document and mentioned you.

DOCUMENT: {doc_name}
FROM: {sender_name}
COMMENT: "{comment_text}"

INSTRUCTIONS:
- Analyze what they're asking or saying
- If it's a question (like "what is geography?"), answer it directly and clearly
- If they're asking you to do something (send email, look up info, schedule meeting, etc.), USE YOUR TOOLS to do it
- If they reference the document content, help with that
- Your response will be posted as a reply to their comment
- Be helpful, concise, and take action when needed

Respond appropriately:"""
                
                response = await self._run_with_failover(message)
            
            return response or "I've reviewed your comment."
            
        except asyncio.TimeoutError:
            return "Sorry, the request took too long. Please try again."
        except Exception as e:
            logger.error(f"Word notification error: {e}")
            return f"Sorry, I encountered an error: {str(e)}"
    
    async def handle_excel_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Handle Excel document comment notifications - let AI decide what to do."""
        try:
            logger.info("📊 Processing Excel notification")
            
            # Initialize MCP for full tool access - the user might ask for anything!
            await self._ensure_mcp_initialized(auth, auth_handler_name, context)
            
            # Get the comment text and context
            comment_text = getattr(context.activity, "text", "") or ""
            # Excel uses @ mentions without <at> tags
            comment_text = comment_text.strip()
            
            doc_name = getattr(context.activity.conversation, "topic", "") or "Spreadsheet"
            sender_name = ""
            if context.activity.from_property:
                sender_name = getattr(context.activity.from_property, "name", "") or ""
            
            logger.info(f"📊 Excel comment from {sender_name}: '{comment_text[:50]}...'")
            
            async with asyncio.timeout(self.PROCESSING_TIMEOUT):
                message = f"""Someone commented on an Excel spreadsheet and mentioned you.

SPREADSHEET: {doc_name}
FROM: {sender_name}
COMMENT: "{comment_text}"

INSTRUCTIONS:
- Analyze what they're asking or saying
- If it's a question (like "what is geography?"), answer it directly and clearly
- If they're asking you to do something (send email, look up info, analyze data, etc.), USE YOUR TOOLS to do it
- If they reference the spreadsheet data, help with that
- Your response will be posted as a reply to their comment
- Be helpful, concise, and take action when needed

Respond appropriately:"""
                
                response = await self._run_with_failover(message)
            
            return response or "I've reviewed your comment."
            
        except asyncio.TimeoutError:
            return "Sorry, the request took too long. Please try again."
        except Exception as e:
            logger.error(f"Excel notification error: {e}")
            return f"Sorry, I encountered an error: {str(e)}"
    
    async def handle_powerpoint_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Handle PowerPoint document comment notifications - let AI decide what to do."""
        try:
            logger.info("📽️ Processing PowerPoint notification")
            
            # Initialize MCP for full tool access - the user might ask for anything!
            await self._ensure_mcp_initialized(auth, auth_handler_name, context)
            
            # Get the comment text and context
            comment_text = getattr(context.activity, "text", "") or ""
            comment_text = comment_text.replace("<at>", "").replace("</at>", "").strip()
            
            doc_name = getattr(context.activity.conversation, "topic", "") or "Presentation"
            sender_name = ""
            if context.activity.from_property:
                sender_name = getattr(context.activity.from_property, "name", "") or ""
            
            logger.info(f"📽️ PowerPoint comment from {sender_name}: '{comment_text[:50]}...'")
            
            async with asyncio.timeout(self.PROCESSING_TIMEOUT):
                message = f"""Someone commented on a PowerPoint presentation and mentioned you.

PRESENTATION: {doc_name}
FROM: {sender_name}
COMMENT: "{comment_text}"

INSTRUCTIONS:
- Analyze what they're asking or saying
- If it's a question (like "what is geography?"), answer it directly and clearly
- If they're asking you to do something (send email, look up info, etc.), USE YOUR TOOLS to do it
- If they reference the presentation content, help with that
- Your response will be posted as a reply to their comment
- Be helpful, concise, and take action when needed

Respond appropriately:"""
                
                response = await self._run_with_failover(message)
            
            return response or "I've reviewed your comment."
            
        except asyncio.TimeoutError:
            return "Sorry, the request took too long. Please try again."
        except Exception as e:
            logger.error(f"PowerPoint notification error: {e}")
            return f"Sorry, I encountered an error: {str(e)}"
    
    async def handle_lifecycle_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """Handle agent lifecycle notifications."""
        try:
            logger.info("📋 Processing lifecycle notification")
            
            # Extract event type
            event_type = None
            if hasattr(notification_activity, 'activity') and notification_activity.activity:
                event_type = getattr(notification_activity.activity, 'name', None)
            
            value_data = getattr(notification_activity, 'value', None)
            if isinstance(value_data, dict):
                event_type = value_data.get('lifecycle_event_type', event_type)
            
            if event_type == "agenticUserIdentityCreated":
                logger.info("✅ User identity created")
                return "User identity created - agent initialized."
            elif event_type == "agenticUserWorkloadOnboardingUpdated":
                logger.info("🔄 Workload onboarding updated")
                return "Workload onboarding updated."
            elif event_type == "agenticUserDeleted":
                logger.info("🗑️ User identity deleted")
                return "User identity deleted - cleanup completed."
            else:
                logger.info(f"📋 Lifecycle event: {event_type}")
                return f"Lifecycle event '{event_type}' acknowledged."
                
        except Exception as e:
            logger.error(f"Lifecycle notification error: {e}")
            return "Lifecycle event processed with warnings."
    
    # =========================================================================
    # CLEANUP
    # =========================================================================
    
    async def cleanup(self) -> None:
        """Clean up agent resources."""
        try:
            await self.mcp_service.cleanup()
            logger.info("✅ ContosoAgent cleanup completed")
        except Exception as e:
            logger.error(f"Cleanup error: {e}")

# Copyright (c) Microsoft. All rights reserved.

"""
Contoso Agent

A full-featured AI agent for Contoso organization integrated with Microsoft 365.
Uses Azure OpenAI and MCP servers for extended functionality.
"""

import asyncio
import aiohttp
import json
import logging
import os
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
        
        # =====================================================================
        # INITIALIZATION GATE STATE
        # =====================================================================
        # Tracks whether the agent has been initialized against the SharePoint
        # list. This is checked deterministically in code, not left to the LLM.
        self._init_gate_checked = False
        self._init_gate_passed = False  # True only if IsInstructionsComplete=true
        self._agent_manager_email = None  # The manager from the SP list
        self._agent_instructions = None  # Instructions from the SP list
        self._agent_user_id = None  # The agent's own user ID
    
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
                
                # Content filter errors are NOT retryable — raise immediately
                is_content_filter = (
                    "content_filter" in error_str or
                    "content management policy" in error_str or
                    "responsibleaipolicyviolation" in error_str
                )
                if is_content_filter:
                    logger.error("🚫 Content filter rejection — not retryable")
                    raise
                
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
    # DIRECT GRAPH API (fully deterministic, no LLM)
    # =========================================================================

    _GRAPH_BASE = "https://graph.microsoft.com/v1.0"
    _SP_SITE_URL = "https://m365cpi76377892.sharepoint.com/sites/Contoso"

    # Power Automate flow URL for SharePoint list operations.
    # Set via SP_FLOW_URL env var.  The flow handles GET (query) and
    # POST (create) operations for the AgentUsersInstructions list.
    _SP_FLOW_URL = os.environ.get("SP_FLOW_URL", "")

    async def _get_graph_token(
        self,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> Optional[str]:
        """Exchange the agentic token for a Microsoft Graph API token."""
        if not auth_handler_name:
            return None
        try:
            token_result = await auth.exchange_token(
                context,
                scopes=["https://graph.microsoft.com/.default"],
                auth_handler_id=auth_handler_name,
            )
            if token_result and hasattr(token_result, "token") and token_result.token:
                logger.info("✅ Graph API token obtained")
                return token_result.token
        except Exception as e:
            logger.warning(f"⚠️ Graph token exchange failed: {e}")
        return None

    async def _graph_request(
        self, method: str, path: str, token: str, body: Optional[dict] = None
    ) -> Optional[dict]:
        """Make a request to the Microsoft Graph API. Returns JSON or None."""
        url = f"{self._GRAPH_BASE}{path}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        try:
            async with aiohttp.ClientSession() as session:
                async with session.request(
                    method, url, json=body, headers=headers
                ) as resp:
                    if resp.status in (200, 201):
                        return await resp.json()
                    text = await resp.text()
                    logger.error(
                        f"Graph {method} {path}: {resp.status} — {text[:300]}"
                    )
        except Exception as e:
            logger.error(f"Graph {method} {path} exception: {e}")
        return None

    # =========================================================================
    # INITIALIZATION GATE (deterministic, code-enforced via Power Automate)
    # =========================================================================

    async def _ensure_init_gate(
        self,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> tuple[bool, str]:
        """
        Deterministic initialization gate. Returns (passed, message).

        Uses a Power Automate flow for all SharePoint list operations and
        Graph API only for read-only profile lookups (GET /me, GET /me/manager).
        """
        # Always get a Graph token (needed for sender identity resolution)
        graph_token = await self._get_graph_token(auth, auth_handler_name, context)

        # Cached: already passed — just verify sender
        if self._init_gate_checked and self._init_gate_passed:
            return await self._check_sender_is_manager(context, graph_token)

        # NOT passed yet (first time OR still pending) — always re-query
        # the flow so we pick up changes the manager made in SharePoint.

        if not graph_token:
            return False, (
                "Hi! I couldn't obtain an authentication token. "
                "Please try again or contact an administrator."
            )

        # First time — ensure MCP is ready (needed for Teams notification later)
        await self._ensure_mcp_initialized(auth, auth_handler_name, context)

        # Get agent UPN from Graph
        profile = await self._graph_request("GET", "/me", graph_token)
        if not profile or not profile.get("userPrincipalName"):
            return False, (
                "Hi! I couldn't retrieve my profile. "
                "Please try again or contact an administrator."
            )
        agent_upn = profile["userPrincipalName"]
        logger.info(f"🔍 Init gate: agent UPN = {agent_upn}")

        # Query SP list via flow, passing the Graph token for authorization
        return await self._init_gate_via_flow(agent_upn, graph_token, context)

    # ------------------------------------------------------------------
    # Power Automate flow path (fully deterministic, no LLM)
    # ------------------------------------------------------------------

    async def _flow_request(self, payload: dict, token: str = "") -> Optional[dict]:
        """Call the Power Automate flow. Returns parsed JSON or None.

        The flow trigger is set to "Anyone" — the URL itself contains a SAS
        signature and acts as the authentication secret.  The Graph token is
        passed inside the JSON body (not as an Authorization header) so the
        flow can use it for internal Graph/SharePoint calls if needed.
        """
        if not self._SP_FLOW_URL:
            logger.error(
                "❌ SP_FLOW_URL not set. "
                "Cannot perform SharePoint list operations without the flow URL."
            )
            return None

        # Include token in body so the flow can use it internally
        if token:
            payload = {**payload, "token": token}

        try:
            async with aiohttp.ClientSession() as session:
                async with session.post(
                    self._SP_FLOW_URL,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                    timeout=aiohttp.ClientTimeout(total=30),
                ) as resp:
                    body = await resp.text()
                    if resp.status in (200, 201, 202):
                        try:
                            return json.loads(body)
                        except json.JSONDecodeError:
                            logger.info(f"✅ Flow returned non-JSON: {body[:200]}")
                            return {"success": True, "raw": body}
                    else:
                        logger.error(
                            f"❌ Flow returned HTTP {resp.status}: {body[:300]}"
                        )
        except Exception as e:
            logger.error(f"❌ Flow call failed: {e}")
        return None

    async def _init_gate_via_flow(
        self,
        agent_upn: str,
        graph_token: str,
        context: TurnContext,
    ) -> tuple[bool, str]:
        """
        Query the SP list via Power Automate flow, evaluate the result,
        and create a new entry if the agent is not found.
        """
        # 1. Query list items filtered by agentUserId
        result = await self._flow_request({
            "action": "get",
            "agentUserId": agent_upn,
        }, token=graph_token)

        if result is None:
            return False, (
                "Hi! I couldn't check my setup status. "
                "Please try again or contact an administrator."
            )

        items = result.get("value", [])

        # 2. Find agent's own entry
        my_entry = None
        if isinstance(items, list):
            for item in items:
                fields = item if isinstance(item, dict) else {}
                uid = fields.get("agentUserId", "")
                if uid and uid.lower() == agent_upn.lower():
                    my_entry = fields
                    break

        # 3. Resolve manager from Graph (always reliable)
        manager_email = await self._resolve_manager_email(graph_token)

        # 4. Evaluate
        if my_entry:
            is_complete = my_entry.get("IsInstructionsComplete", False)
            has_instructions = bool(
                my_entry.get("Instructions", "").strip()
            )
            if is_complete or has_instructions:
                # READY ✅  (flag set OR instructions text present)
                self._init_gate_checked = True
                self._init_gate_passed = True
                self._agent_user_id = agent_upn
                self._agent_manager_email = manager_email
                self._agent_instructions = my_entry.get("Instructions", "")
                logger.info(
                    f"✅ Init gate PASSED. Manager: {self._agent_manager_email}"
                )
                return await self._check_sender_is_manager(context, graph_token)
            else:
                # PENDING ⏳  (no flag AND no instructions text)
                self._init_gate_checked = True
                self._init_gate_passed = False
                self._agent_manager_email = manager_email
                logger.info("⏳ Init gate: PENDING (instructions not complete)")
                return False, (
                    "Hi! My setup is still pending — my manager needs to complete "
                    "the instructions before I can assist anyone.\n\n"
                    "I've already notified them. Please check with my manager if you'd "
                    "like to expedite this."
                )
        else:
            # NOT FOUND — create entry via flow + notify manager
            return await self._init_gate_create_via_flow(
                agent_upn, graph_token, manager_email, context
            )

    async def _resolve_manager_email(self, token: str) -> str:
        """Resolve the agent's manager email via Graph API (GET /me/manager)."""
        manager_data = await self._graph_request("GET", "/me/manager", token)
        if not manager_data:
            return ""
        email = (
            manager_data.get("mail")
            or manager_data.get("userPrincipalName", "")
        )
        return (email or "").strip().lower()

    async def _init_gate_create_via_flow(
        self,
        agent_upn: str,
        graph_token: str,
        manager_email: str,
        context: TurnContext,
    ) -> tuple[bool, str]:
        """Create the SP list entry via flow + notify manager via Teams."""
        # Get manager display name from Graph
        manager_data = await self._graph_request("GET", "/me/manager", graph_token)
        manager_name = manager_data.get("displayName", "") if manager_data else ""
        logger.info(f"📝 Agent manager: {manager_name} ({manager_email})")

        # Create SP list item via flow
        create_result = await self._flow_request({
            "action": "create",
            "agentUserId": agent_upn,
            "isInstructionsComplete": False,
        }, token=graph_token)
        if create_result:
            logger.info(f"✅ SP list item created via flow for {agent_upn}")
        else:
            logger.error("❌ Failed to create SP list item via flow")

        self._init_gate_checked = True
        self._init_gate_passed = False
        self._agent_manager_email = manager_email

        # Check if sender is the manager — if so, give them a direct link
        # instead of sending a redundant Teams DM to themselves.
        sender_is_manager = False
        if manager_email and graph_token:
            passed, _ = await self._check_sender_is_manager(context, graph_token)
            sender_is_manager = passed

        setup_link = f"{self._SP_SITE_URL}/Lists/AgentUsersInstructions"

        if sender_is_manager:
            logger.info("✅ Sender is the manager — skipping Teams DM, giving direct link")
            return False, (
                f"Hi {manager_name or 'there'}! I've just been activated and my "
                f"profile has been created in the setup list.\n\n"
                f"Please go to {setup_link} and fill in the **Instructions** "
                f"for my entry, then send me another message and I'll be ready to help!"
            )

        # Sender is NOT the manager — notify manager via Teams
        if manager_email:
            await self._notify_manager_via_teams(agent_upn, manager_email, manager_name)

        manager_display = manager_name or manager_email or "your manager"
        return False, (
            f"Hi! I need to get set up before I can help. I've notified my "
            f"manager ({manager_display}) to complete the required setup instructions "
            f"for my profile.\n\n"
            f"Once my manager fills in the instructions, I'll be fully ready "
            f"to assist. Please reach out to them if you'd like to speed things up!"
        )

    async def _notify_manager_via_teams(
        self, agent_upn: str, manager_email: str, manager_name: str
    ) -> None:
        """Send a Teams message to the manager. Uses LLM since Teams MCP works fine."""
        greeting = f" {manager_name}" if manager_name else ""
        prompt = (
            f"Send a Teams message to {manager_email}. "
            f"Use createChat with their email, then postMessage with this text:\n\n"
            f'"Hi{greeting}! I\'m the Contoso Assistant agent ({agent_upn}). '
            f"I've just been activated but don't have instructions set up yet. "
            f"Could you please go to {self._SP_SITE_URL}/Lists/AgentUsersInstructions "
            f"and fill in my instructions? I won't be able to assist anyone until "
            f'that\'s complete. Thank you!"\n\n'
            f"After sending, return only: done"
        )
        try:
            async with asyncio.timeout(60):
                await self._run_with_failover(prompt)
                logger.info(f"✅ Teams notification sent to {manager_email}")
        except asyncio.TimeoutError:
            logger.warning(f"⚠️ Teams notification to {manager_email} timed out (message may still have been delivered)")
        except Exception as e:
            logger.error(f"❌ Failed to notify manager via Teams: {e}")

    # ------------------------------------------------------------------
    # Sender authorization check
    # ------------------------------------------------------------------

    async def _check_sender_is_manager(
        self, context: TurnContext, graph_token: str | None = None
    ) -> tuple[bool, str]:
        """
        Check if the person who sent the message is the agent's assigned manager.
        Resolves the sender's AAD object ID to an email via Graph if needed.
        Returns (passed, message).
        """
        if not self._agent_manager_email:
            # No manager set — can't verify, allow through
            logger.warning("⚠️ No manager email set, allowing request through")
            return True, ""

        # Extract all available sender identifiers
        sender_email = ""
        sender_aad_id = ""
        sender_raw_id = ""
        if context.activity.from_property:
            sender_raw_id = getattr(context.activity.from_property, "id", "") or ""
            sender_aad_id = (
                getattr(context.activity.from_property, "aad_object_id", "") or ""
            )

        # If the raw id looks like an email, use it directly
        if "@" in sender_raw_id:
            sender_email = sender_raw_id.strip().lower()
        elif sender_aad_id and graph_token:
            # Resolve AAD object ID → email via Graph
            user_data = await self._graph_request(
                "GET", f"/users/{sender_aad_id}", graph_token
            )
            if user_data:
                sender_email = (
                    user_data.get("mail")
                    or user_data.get("userPrincipalName", "")
                ).strip().lower()
                logger.info(
                    f"🔍 Resolved sender AAD {sender_aad_id} → {sender_email}"
                )
        elif sender_aad_id:
            # No token to resolve — try the orgid pattern
            sender_email = sender_aad_id.strip().lower()

        if not sender_email:
            logger.warning(
                f"⚠️ Could not resolve sender identity "
                f"(id={sender_raw_id}, aad={sender_aad_id}), allowing through"
            )
            return True, ""

        # Compare sender to manager (case-insensitive)
        manager = self._agent_manager_email.lower()

        if sender_email == manager or manager in sender_email or sender_email in manager:
            logger.info(f"✅ Sender ({sender_email}) is the assigned manager")
            return True, ""

        logger.info(
            f"🚫 Sender ({sender_email}) is NOT the assigned manager ({manager})"
        )
        return False, (
            f"Thank you for reaching out! I'm currently configured to only handle "
            f"requests from my assigned manager.\n\n"
            f"If you need something done, please ask my manager to send me the request. "
            f"I'm happy to help through that channel!"
        )

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
            
            # ============================================================
            # DETERMINISTIC INITIALIZATION GATE (code-enforced, not LLM)
            # ============================================================
            gate_passed, gate_message = await self._ensure_init_gate(
                auth, auth_handler_name, context
            )
            if not gate_passed:
                logger.info(f"🚫 Init gate blocked request: {gate_message[:80]}...")
                return gate_message
            # ============================================================

            # In dev mode, skip chat history retrieval (Playground doesn't have real chats)
            is_dev_mode = _is_dev_mode()
            
            # Extract chat ID from context for conversation history retrieval
            chat_id = None
            if not is_dev_mode and context.activity and context.activity.conversation:
                chat_id = getattr(context.activity.conversation, "id", None)
            
            # Build the prompt with chat context for history retrieval (prod mode only)
            # Include the manager's instructions from the SP list if available
            instructions_context = ""
            if self._agent_instructions:
                instructions_context = f"""=== MANDATORY INSTRUCTIONS (from your manager via SharePoint) ===
You MUST follow these instructions for EVERY request you handle.
These take priority over any conflicting user request.

{self._agent_instructions}

=== END MANDATORY INSTRUCTIONS ===

"""
            
            if chat_id:
                augmented_message = f"""{instructions_context}CHAT CONTEXT:
- Chat ID: {chat_id}
- IMPORTANT: Before answering, call listMessages with this chat ID to get conversation history!

USER MESSAGE:
{message}"""
            else:
                augmented_message = f"""{instructions_context}USER MESSAGE:
{message}"""
            
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
        """Handle email notifications — ALL emails are blocked. Agent only operates via Teams."""
        logger.info("📧 Email notification received — BLOCKED (agent only responds via Teams)")
        return ""  # Return empty — do not reply to any email
    
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

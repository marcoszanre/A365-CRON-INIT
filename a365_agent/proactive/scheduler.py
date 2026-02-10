# Copyright (c) Microsoft. All rights reserved.

"""
Proactive Scheduler

An asyncio-based cron scheduler that periodically:
    1. Queries PostgreSQL for all agents that have enabled scheduled tasks
    2. For each agent: acquires an MCP token via Agent User Impersonation
       (Blueprint creds from .env + per-agent identity from DB)
    3. Queries the agent's scheduled_tasks from DB
    4. Executes each task prompt with MCP tools
    5. Logs results back to PostgreSQL

The interval is controlled by the ``CRON_INTERVAL_SECONDS`` environment
variable (default: 3600 = 1 hour).

Usage (standalone - for testing):
    CRON_INTERVAL_SECONDS=30 uv run -m a365_agent.proactive.scheduler

Usage (integrated - started alongside the HTTP server):
    The ``create_and_run_host`` path in main.py launches the scheduler
    as a background asyncio task when ``CRON_ENABLED=true``.
"""

import asyncio
import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from a365_agent.config import get_settings
from a365_agent.mcp import MCPService
from a365_agent.proactive.auth import AgentCredentials, ProactiveTokenProvider
from a365_agent.proactive.mock_context import MockAuthorization, MockTurnContext

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_cron_system_prompt() -> str:
    """Load the cron system prompt from agents/cron_system_prompt.md."""
    path = Path(__file__).resolve().parent.parent.parent / "agents" / "cron_system_prompt.md"
    if not path.exists():
        raise FileNotFoundError(f"Required file not found: {path}")
    return path.read_text(encoding="utf-8").strip()


def _render_task_prompt(task_prompt: str, manager_email: str, agent_upn: str) -> str:
    """Substitute runtime variables into a task prompt from the DB."""
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    try:
        return task_prompt.format(
            manager_email=manager_email,
            target_email=manager_email,  # alias for backwards compat
            agent_upn=agent_upn,
            timestamp=now,
        )
    except KeyError:
        # If the prompt doesn't use all placeholders, just return as-is
        return task_prompt


# ---------------------------------------------------------------------------
# Scheduler
# ---------------------------------------------------------------------------

class ProactiveScheduler:
    """
    Multi-agent periodic scheduler for proactive tasks.

    On each tick:
        1. Query all agents with enabled scheduled_tasks from PostgreSQL
        2. For each agent, acquire an MCP token (Blueprint .env + agent identity from DB)
        3. Init MCP, load tasks, run agent for each task, log results
        4. Cleanup and sleep

    Lifecycle:
        scheduler = ProactiveScheduler()
        task = asyncio.create_task(scheduler.start())   # non-blocking
        ...
        await scheduler.stop()                           # graceful shutdown
    """

    def __init__(self, interval_seconds: Optional[int] = None):
        self.interval = interval_seconds or int(
            os.getenv("CRON_INTERVAL_SECONDS", "3600")
        )
        self._running = False
        self._task: Optional[asyncio.Task] = None
        self._token_provider = ProactiveTokenProvider()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def start(self) -> None:
        """Run the scheduler loop. Blocks until ``stop()`` is called."""
        if self._running:
            logger.warning("⏰ Scheduler is already running")
            return

        self._running = True
        logger.info(f"⏰ Proactive scheduler started — interval {self.interval}s")

        while self._running:
            try:
                await self._tick()
            except Exception:
                logger.exception("❌ Cron tick failed (will retry next interval)")

            # Sleep in small increments so we can bail quickly on stop()
            for _ in range(self.interval):
                if not self._running:
                    break
                await asyncio.sleep(1)

        logger.info("⏰ Proactive scheduler stopped")

    async def stop(self) -> None:
        """Signal the scheduler to stop after the current sleep."""
        self._running = False
        if self._task and not self._task.done():
            self._task.cancel()
            try:
                await self._task
            except asyncio.CancelledError:
                pass

    # ------------------------------------------------------------------
    # Single execution tick — loops through ALL agents
    # ------------------------------------------------------------------

    async def _tick(self) -> None:
        """Execute one cron cycle across all agents with scheduled tasks."""
        tick_start = datetime.now(timezone.utc)
        logger.info("🔄 Cron tick starting...")

        from a365_agent.storage import get_storage
        storage = await get_storage()

        # 1. Get all agents that have enabled scheduled tasks
        agents = await storage.get_all_agents_with_tasks()
        if not agents:
            logger.info("📋 No agents with enabled scheduled tasks — skipping tick")
            return

        logger.info(f"📋 Found {len(agents)} agent(s) with scheduled tasks")

        # 2. Process each agent
        for agent_row in agents:
            agent_upn = agent_row["agent_user_id"]
            try:
                await self._process_agent(storage, agent_row)
            except Exception:
                logger.exception(f"❌ Failed processing agent {agent_upn}")

        tick_end = datetime.now(timezone.utc)
        duration_ms = int((tick_end - tick_start).total_seconds() * 1000)
        logger.info(f"✅ Cron tick completed in {duration_ms}ms ({len(agents)} agents)")

    # ------------------------------------------------------------------
    # Per-agent processing
    # ------------------------------------------------------------------

    async def _process_agent(self, storage, agent_row: dict) -> None:
        """Acquire token for one agent, run all its scheduled tasks."""
        agent_upn = agent_row["agent_user_id"]
        manager_email = agent_row.get("manager_email", "")
        logger.info(f"👤 Processing agent: {agent_upn} (manager: {manager_email})")

        # Build per-agent credentials (Blueprint from .env + identity from DB)
        creds = AgentCredentials.from_agent_row(agent_row)
        missing = creds.validate()
        if missing:
            logger.error(
                f"⚠️ Skipping {agent_upn} — missing credentials: {', '.join(missing)}"
            )
            return

        # Acquire MCP token for this agent
        mcp_token = await self._token_provider.acquire_mcp_token(creds)
        logger.info(f"🔑 Token acquired for {agent_upn}")

        # Get scheduled tasks for this agent
        tasks = await storage.get_scheduled_tasks(agent_upn)
        if not tasks:
            logger.info(f"📋 No enabled tasks for {agent_upn} — skipping")
            return

        logger.info(f"📋 {len(tasks)} task(s) for {agent_upn}")

        # Init MCP servers once per agent (all tasks share the same session)
        settings = get_settings()
        if not settings.model_pool or len(settings.model_pool) == 0:
            raise RuntimeError("No Azure OpenAI models configured (model_pool is empty)")

        model = settings.model_pool.get_next_model()

        from agent_framework.azure import AzureOpenAIChatClient

        chat_client = AzureOpenAIChatClient(
            endpoint=model.endpoint,
            api_key=model.api_key,
            deployment_name=model.deployment,
            api_version=model.api_version,
        )

        system_prompt = _load_cron_system_prompt()
        mock_auth = MockAuthorization(mcp_token)
        mock_ctx = MockTurnContext(agent_upn)

        mcp_service = MCPService()
        agent = await mcp_service.initialize_with_bearer_token(
            chat_client=chat_client,
            agent_instructions=system_prompt,
            bearer_token=mcp_token,
            auth=mock_auth,
            auth_handler_name="PROACTIVE-CRON",
            turn_context=mock_ctx,
        )

        if agent is None:
            raise RuntimeError(f"MCP initialization returned None for {agent_upn}")

        # Optionally upgrade to planning model
        planning_pool = settings.planning_pool
        if planning_pool and len(planning_pool) > 0:
            planning_cfg = planning_pool.get_next_model()
            planning_client = AzureOpenAIChatClient(
                endpoint=planning_cfg.endpoint,
                api_key=planning_cfg.api_key,
                deployment_name=planning_cfg.deployment,
                api_version=planning_cfg.api_version,
            )
            agent.chat_client = planning_client
            logger.info(f"🧠 Upgraded to planning model: {planning_cfg.name}")

        # Execute each task
        for task_row in tasks:
            await self._execute_task(storage, agent, agent_upn, manager_email, task_row)

        # Cleanup MCP for this agent
        await mcp_service.cleanup()

    # ------------------------------------------------------------------
    # Single task execution
    # ------------------------------------------------------------------

    async def _execute_task(
        self, storage, agent, agent_upn: str, manager_email: str, task_row: dict
    ) -> None:
        """Run one scheduled task and log the result."""
        task_id = task_row["task_id"]
        task_name = task_row["task_name"]
        raw_prompt = task_row["task_prompt"]

        logger.info(f"🤖 Running task '{task_name}' for {agent_upn}...")

        # Render the prompt with runtime variables
        prompt = _render_task_prompt(raw_prompt, manager_email, agent_upn)

        task_start = datetime.now(timezone.utc)
        status = "success"
        response = ""

        try:
            result = await agent.run(prompt)

            # Extract response text
            if hasattr(result, "contents"):
                response = str(result.contents)
            elif hasattr(result, "text"):
                response = str(result.text)
            elif hasattr(result, "content"):
                response = str(result.content)
            else:
                response = str(result)

            logger.info(
                f"✅ Task '{task_name}' completed: "
                f"{response[:150]}{'...' if len(response) > 150 else ''}"
            )
        except Exception as e:
            status = "error"
            response = str(e)
            logger.error(f"❌ Task '{task_name}' failed: {e}")

        # Update task result in DB
        task_end = datetime.now(timezone.utc)
        duration_ms = int((task_end - task_start).total_seconds() * 1000)

        try:
            await storage.update_scheduled_task_result(
                task_id=str(task_id),
                status=status,
                result_text=response[:2000],
            )
            await storage.log_tool_execution(
                agent_id=agent_upn,
                tool_name=f"cron:{task_name}",
                conversation_id="proactive-cron",
                tool_input={"prompt": prompt[:500]},
                tool_output={"response": response[:500]},
                status=status,
                duration_ms=duration_ms,
            )
        except Exception:
            logger.debug("⚠️ Could not log task result to PostgreSQL (non-fatal)")


# ---------------------------------------------------------------------------
# Standalone entry-point (for testing)
# ---------------------------------------------------------------------------

async def _main() -> None:
    """Run the scheduler standalone (useful for testing)."""
    import warnings
    warnings.filterwarnings("ignore", category=RuntimeWarning, message=".*cancel scope.*")

    settings = get_settings()
    settings.configure_logging()

    scheduler = ProactiveScheduler()
    try:
        await scheduler.start()
    except KeyboardInterrupt:
        await scheduler.stop()


if __name__ == "__main__":
    asyncio.run(_main())

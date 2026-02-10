# Copyright (c) Microsoft. All rights reserved.

"""
PostgreSQL Storage Implementation

Async PostgreSQL storage backend for the multi-agent MCP server environment.
Replaces SharePoint list + Power Automate flow with direct database operations.

Connection string configurable via PG_DSN env var, defaults to:
    postgresql://mcpagent:mcpagent_dev@localhost:5432/mcp_agents
"""

import asyncio
import json
import logging
import os
import time
import uuid
from datetime import datetime, timezone
from typing import Any, Optional

import asyncpg

logger = logging.getLogger(__name__)

# Default connection string for local dev
_DEFAULT_DSN = "postgresql://mcpagent:mcpagent_dev@localhost:5432/mcp_agents"

# Singleton instance
_instance: Optional["PostgresStorage"] = None
_lock = asyncio.Lock()


async def get_storage() -> "PostgresStorage":
    """Get or create the singleton PostgresStorage instance."""
    global _instance
    if _instance is None:
        async with _lock:
            if _instance is None:
                _instance = PostgresStorage()
                await _instance.connect()
    return _instance


class PostgresStorage:
    """
    Async PostgreSQL storage for multi-agent MCP scenarios.

    Tables (all under `agent_storage` schema):
        - agent_registry:   replaces SharePoint AgentUsersInstructions list
        - conversations:    replaces MemoryStorage for cross-agent history
        - shared_state:     general key-value shared state across agents
        - tool_executions:  audit log for MCP tool calls
        - task_queue:       inter-agent task coordination / handoffs
    """

    SCHEMA = "agent_storage"

    def __init__(self, dsn: Optional[str] = None):
        self.dsn = dsn or os.environ.get("PG_DSN", _DEFAULT_DSN)
        self._pool: Optional[asyncpg.Pool] = None

    # ------------------------------------------------------------------
    # Connection lifecycle
    # ------------------------------------------------------------------

    async def connect(self, min_size: int = 2, max_size: int = 10) -> None:
        """Create the connection pool."""
        if self._pool is not None:
            return
        try:
            t0 = time.monotonic()
            self._pool = await asyncpg.create_pool(
                self.dsn, min_size=min_size, max_size=max_size
            )
            elapsed = (time.monotonic() - t0) * 1000
            logger.info(f"✅ PostgreSQL connection pool created ({elapsed:.0f}ms)")
        except Exception as e:
            logger.error(f"❌ PostgreSQL connection failed: {e}")
            raise

    async def close(self) -> None:
        """Close the connection pool."""
        if self._pool:
            await self._pool.close()
            self._pool = None
            logger.info("PostgreSQL connection pool closed")

    async def health_check(self) -> bool:
        """Quick connectivity check."""
        try:
            async with self._pool.acquire() as conn:
                await conn.fetchval("SELECT 1")
            return True
        except Exception:
            return False

    # ==================================================================
    # AGENT REGISTRY (replaces SharePoint AgentUsersInstructions list)
    # ==================================================================

    async def get_agent(self, agent_user_id: str) -> Optional[dict]:
        """
        Look up an agent by UPN. Returns dict with fields matching the
        old SharePoint list columns, or None if not found.
        """
        t0 = time.monotonic()
        logger.debug(f"🔍 PG get_agent: looking up {agent_user_id}")
        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                SELECT agent_user_id, instructions, is_instructions_complete,
                       manager_email, manager_name, created_at, updated_at
                FROM {self.SCHEMA}.agent_registry
                WHERE LOWER(agent_user_id) = LOWER($1)
                """,
                agent_user_id,
            )
        elapsed = (time.monotonic() - t0) * 1000
        if row is None:
            logger.info(f"🔍 PG get_agent: NOT FOUND ({elapsed:.0f}ms)")
            return None
        logger.info(
            f"🔍 PG get_agent: found {agent_user_id} "
            f"(complete={row['is_instructions_complete']}, {elapsed:.0f}ms)"
        )
        return dict(row)

    async def create_agent(
        self,
        agent_user_id: str,
        *,
        manager_email: str = "",
        manager_name: str = "",
        instructions: str = "",
        is_instructions_complete: bool = False,
    ) -> dict:
        """Create a new agent registry entry (replaces flow 'create' action)."""
        t0 = time.monotonic()
        logger.debug(f"📝 PG create_agent: upserting {agent_user_id}")
        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                INSERT INTO {self.SCHEMA}.agent_registry
                    (agent_user_id, instructions, is_instructions_complete,
                     manager_email, manager_name)
                VALUES ($1, $2, $3, $4, $5)
                ON CONFLICT (agent_user_id) DO UPDATE SET
                    instructions = EXCLUDED.instructions,
                    is_instructions_complete = EXCLUDED.is_instructions_complete,
                    manager_email = EXCLUDED.manager_email,
                    manager_name = EXCLUDED.manager_name,
                    updated_at = NOW()
                RETURNING *
                """,
                agent_user_id,
                instructions,
                is_instructions_complete,
                manager_email,
                manager_name,
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.info(f"✅ PG create_agent: upserted {agent_user_id} ({elapsed:.0f}ms)")
        return dict(row)

    async def update_agent(
        self, agent_user_id: str, **fields
    ) -> Optional[dict]:
        """Update specific fields on an agent registry entry."""
        allowed = {
            "instructions",
            "is_instructions_complete",
            "manager_email",
            "manager_name",
        }
        updates = {k: v for k, v in fields.items() if k in allowed}
        if not updates:
            return await self.get_agent(agent_user_id)

        t0 = time.monotonic()
        logger.debug(f"📝 PG update_agent: {agent_user_id} fields={list(updates.keys())}")

        set_clauses = ", ".join(
            f"{k} = ${i + 2}" for i, k in enumerate(updates)
        )
        values = [agent_user_id, *updates.values()]

        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                UPDATE {self.SCHEMA}.agent_registry
                SET {set_clauses}, updated_at = NOW()
                WHERE LOWER(agent_user_id) = LOWER($1)
                RETURNING *
                """,
                *values,
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.info(f"✅ PG update_agent: {agent_user_id} ({elapsed:.0f}ms)")
        return dict(row) if row else None

    # ==================================================================
    # CONVERSATIONS (replaces MemoryStorage for multi-agent history)
    # ==================================================================

    async def save_message(
        self,
        conversation_id: str,
        agent_id: str,
        role: str,
        content: str,
        *,
        user_id: str = "",
        metadata: Optional[dict] = None,
    ) -> int:
        """Persist a single conversation message. Returns the row ID."""
        t0 = time.monotonic()
        async with self._pool.acquire() as conn:
            row_id = await conn.fetchval(
                f"""
                INSERT INTO {self.SCHEMA}.conversations
                    (conversation_id, agent_id, user_id, role, content, metadata)
                VALUES ($1, $2, $3, $4, $5, $6::jsonb)
                RETURNING id
                """,
                conversation_id,
                agent_id,
                user_id,
                role,
                content,
                json.dumps(metadata or {}),
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.debug(f"💬 PG save_message: conv={conversation_id} role={role} ({elapsed:.0f}ms)")
        return row_id

    async def get_conversation(
        self,
        conversation_id: str,
        *,
        agent_id: Optional[str] = None,
        limit: int = 50,
    ) -> list[dict]:
        """
        Retrieve conversation history, optionally filtered by agent.
        Returns messages in chronological order.
        """
        t0 = time.monotonic()
        if agent_id:
            query = f"""
                SELECT id, conversation_id, agent_id, user_id, role, content,
                       metadata, created_at
                FROM {self.SCHEMA}.conversations
                WHERE conversation_id = $1 AND agent_id = $2
                ORDER BY created_at ASC
                LIMIT $3
            """
            args = (conversation_id, agent_id, limit)
        else:
            query = f"""
                SELECT id, conversation_id, agent_id, user_id, role, content,
                       metadata, created_at
                FROM {self.SCHEMA}.conversations
                WHERE conversation_id = $1
                ORDER BY created_at ASC
                LIMIT $2
            """
            args = (conversation_id, limit)

        async with self._pool.acquire() as conn:
            rows = await conn.fetch(query, *args)
        elapsed = (time.monotonic() - t0) * 1000
        logger.debug(
            f"💬 PG get_conversation: conv={conversation_id} "
            f"returned {len(rows)} msgs ({elapsed:.0f}ms)"
        )
        return [dict(r) for r in rows]

    # ==================================================================
    # SHARED STATE (cross-agent key-value store)
    # ==================================================================

    async def set_state(
        self,
        key: str,
        value: Any,
        *,
        owner_agent: str = "",
        ttl_seconds: Optional[int] = None,
    ) -> None:
        """Set a shared state value (upsert). Optional TTL in seconds."""
        expires = None
        if ttl_seconds:
            expires = datetime.now(timezone.utc).replace(
                second=datetime.now(timezone.utc).second + ttl_seconds
            )
            # Use proper timedelta instead
            from datetime import timedelta
            expires = datetime.now(timezone.utc) + timedelta(seconds=ttl_seconds)

        async with self._pool.acquire() as conn:
            await conn.execute(
                f"""
                INSERT INTO {self.SCHEMA}.shared_state
                    (key, value, owner_agent, expires_at)
                VALUES ($1, $2::jsonb, $3, $4)
                ON CONFLICT (key) DO UPDATE SET
                    value = EXCLUDED.value,
                    owner_agent = EXCLUDED.owner_agent,
                    expires_at = EXCLUDED.expires_at,
                    updated_at = NOW()
                """,
                key,
                json.dumps(value),
                owner_agent,
                expires,
            )

    async def get_state(self, key: str) -> Optional[Any]:
        """Get a shared state value. Returns None if expired or missing."""
        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                SELECT value FROM {self.SCHEMA}.shared_state
                WHERE key = $1
                  AND (expires_at IS NULL OR expires_at > NOW())
                """,
                key,
            )
        if row is None:
            return None
        return json.loads(row["value"])

    async def delete_state(self, key: str) -> bool:
        """Delete a shared state key. Returns True if something was deleted."""
        async with self._pool.acquire() as conn:
            result = await conn.execute(
                f"DELETE FROM {self.SCHEMA}.shared_state WHERE key = $1",
                key,
            )
        return result.endswith("1")

    # ==================================================================
    # TOOL EXECUTIONS (audit trail for MCP tool calls)
    # ==================================================================

    async def log_tool_execution(
        self,
        agent_id: str,
        tool_name: str,
        *,
        conversation_id: str = "",
        tool_input: Optional[dict] = None,
        tool_output: Optional[dict] = None,
        status: str = "success",
        duration_ms: int = 0,
    ) -> int:
        """Log an MCP tool execution. Returns the row ID."""
        async with self._pool.acquire() as conn:
            row_id = await conn.fetchval(
                f"""
                INSERT INTO {self.SCHEMA}.tool_executions
                    (agent_id, conversation_id, tool_name,
                     tool_input, tool_output, status, duration_ms)
                VALUES ($1, $2, $3, $4::jsonb, $5::jsonb, $6, $7)
                RETURNING id
                """,
                agent_id,
                conversation_id,
                tool_name,
                json.dumps(tool_input or {}),
                json.dumps(tool_output or {}),
                status,
                duration_ms,
            )
        return row_id

    # ==================================================================
    # TASK QUEUE (inter-agent coordination / handoffs)
    # ==================================================================

    async def enqueue_task(
        self,
        source_agent: str,
        target_agent: str,
        task_type: str,
        payload: Optional[dict] = None,
    ) -> str:
        """
        Enqueue a task for another agent. Returns the task_id (UUID).
        """
        task_id = str(uuid.uuid4())
        async with self._pool.acquire() as conn:
            await conn.execute(
                f"""
                INSERT INTO {self.SCHEMA}.task_queue
                    (task_id, source_agent, target_agent, task_type, payload)
                VALUES ($1, $2, $3, $4, $5::jsonb)
                """,
                task_id,
                source_agent,
                target_agent,
                task_type,
                json.dumps(payload or {}),
            )
        logger.info(
            f"📋 Task {task_id}: {source_agent} → {target_agent} ({task_type})"
        )
        return task_id

    async def dequeue_tasks(
        self, target_agent: str, *, limit: int = 10
    ) -> list[dict]:
        """
        Fetch and claim pending tasks for an agent (atomic: sets status to 'in_progress').
        """
        async with self._pool.acquire() as conn:
            rows = await conn.fetch(
                f"""
                UPDATE {self.SCHEMA}.task_queue
                SET status = 'in_progress', updated_at = NOW()
                WHERE id IN (
                    SELECT id FROM {self.SCHEMA}.task_queue
                    WHERE target_agent = $1 AND status = 'pending'
                    ORDER BY created_at ASC
                    LIMIT $2
                    FOR UPDATE SKIP LOCKED
                )
                RETURNING *
                """,
                target_agent,
                limit,
            )
        return [dict(r) for r in rows]

    async def complete_task(
        self,
        task_id: str,
        *,
        result: Optional[dict] = None,
        status: str = "completed",
    ) -> None:
        """Mark a task as completed (or failed)."""
        async with self._pool.acquire() as conn:
            await conn.execute(
                f"""
                UPDATE {self.SCHEMA}.task_queue
                SET status = $2, result = $3::jsonb, updated_at = NOW()
                WHERE task_id = $1
                """,
                task_id,
                status,
                json.dumps(result or {}),
            )

    # ==================================================================
    # SCHEDULED TASKS (cron job task definitions per agent)
    # ==================================================================

    async def get_all_agents_with_tasks(self) -> list[dict]:
        """
        Return all agents that have at least one enabled scheduled task.
        Joins agent_registry with scheduled_tasks for a single query.
        """
        t0 = time.monotonic()
        async with self._pool.acquire() as conn:
            rows = await conn.fetch(
                f"""
                SELECT DISTINCT
                    ar.agent_user_id,
                    ar.manager_email,
                    ar.manager_name,
                    ar.instructions,
                    ar.agent_identity_client_id,
                    ar.agent_user_object_id
                FROM {self.SCHEMA}.agent_registry ar
                INNER JOIN {self.SCHEMA}.scheduled_tasks st
                    ON LOWER(ar.agent_user_id) = LOWER(st.agent_user_id)
                WHERE st.is_enabled = TRUE
                  AND ar.is_instructions_complete = TRUE
                """
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.debug(f"📋 PG get_all_agents_with_tasks: {len(rows)} agents ({elapsed:.0f}ms)")
        return [dict(r) for r in rows]

    async def get_scheduled_tasks(self, agent_user_id: str) -> list[dict]:
        """
        Get all enabled scheduled tasks for a specific agent.
        """
        t0 = time.monotonic()
        async with self._pool.acquire() as conn:
            rows = await conn.fetch(
                f"""
                SELECT id, task_id, agent_user_id, task_name, task_prompt,
                       is_enabled, last_run_at, last_status, last_result,
                       created_at, updated_at
                FROM {self.SCHEMA}.scheduled_tasks
                WHERE LOWER(agent_user_id) = LOWER($1)
                  AND is_enabled = TRUE
                ORDER BY created_at ASC
                """,
                agent_user_id,
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.debug(
            f"📋 PG get_scheduled_tasks: {len(rows)} tasks for {agent_user_id} ({elapsed:.0f}ms)"
        )
        return [dict(r) for r in rows]

    async def update_scheduled_task_result(
        self,
        task_id: str,
        *,
        status: str = "success",
        result_text: str = "",
    ) -> None:
        """Update last_run_at, last_status and last_result for a scheduled task."""
        async with self._pool.acquire() as conn:
            await conn.execute(
                f"""
                UPDATE {self.SCHEMA}.scheduled_tasks
                SET last_run_at = NOW(),
                    last_status = $2,
                    last_result = $3,
                    updated_at = NOW()
                WHERE task_id = $1
                """,
                task_id,
                status,
                result_text[:2000],
            )

    async def create_scheduled_task(
        self,
        agent_user_id: str,
        task_name: str,
        task_prompt: str,
        *,
        is_enabled: bool = True,
    ) -> dict:
        """Create a new scheduled task for an agent."""
        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                INSERT INTO {self.SCHEMA}.scheduled_tasks
                    (agent_user_id, task_name, task_prompt, is_enabled)
                VALUES ($1, $2, $3, $4)
                RETURNING *
                """,
                agent_user_id,
                task_name,
                task_prompt,
                is_enabled,
            )
        logger.info(f"✅ PG create_scheduled_task: {task_name} for {agent_user_id}")
        return dict(row)

    async def get_all_tasks_for_agent(self, agent_user_id: str) -> list[dict]:
        """
        Get ALL scheduled tasks for an agent (enabled and disabled).
        Used by the agent's task management tools.
        """
        t0 = time.monotonic()
        async with self._pool.acquire() as conn:
            rows = await conn.fetch(
                f"""
                SELECT id, task_id, agent_user_id, task_name, task_prompt,
                       is_enabled, last_run_at, last_status, last_result,
                       created_at, updated_at
                FROM {self.SCHEMA}.scheduled_tasks
                WHERE LOWER(agent_user_id) = LOWER($1)
                ORDER BY created_at ASC
                """,
                agent_user_id,
            )
        elapsed = (time.monotonic() - t0) * 1000
        logger.debug(
            f"📋 PG get_all_tasks_for_agent: {len(rows)} tasks for {agent_user_id} ({elapsed:.0f}ms)"
        )
        return [dict(r) for r in rows]

    async def update_scheduled_task_fields(
        self,
        task_id: str,
        **fields,
    ) -> Optional[dict]:
        """Update specific fields on a scheduled task (task_name, task_prompt, is_enabled)."""
        allowed = {"task_name", "task_prompt", "is_enabled"}
        updates = {k: v for k, v in fields.items() if k in allowed}
        if not updates:
            return None

        t0 = time.monotonic()
        set_clauses = ", ".join(
            f"{k} = ${i + 2}" for i, k in enumerate(updates)
        )
        values = [task_id, *updates.values()]

        async with self._pool.acquire() as conn:
            row = await conn.fetchrow(
                f"""
                UPDATE {self.SCHEMA}.scheduled_tasks
                SET {set_clauses}, updated_at = NOW()
                WHERE task_id = $1
                RETURNING *
                """,
                *values,
            )
        elapsed = (time.monotonic() - t0) * 1000
        if row:
            logger.info(f"✅ PG update_scheduled_task_fields: {task_id} ({elapsed:.0f}ms)")
            return dict(row)
        logger.warning(f"⚠️ PG update_scheduled_task_fields: {task_id} not found")
        return None

    async def delete_scheduled_task(self, task_id: str) -> bool:
        """Delete a scheduled task by task_id. Returns True if deleted."""
        t0 = time.monotonic()
        async with self._pool.acquire() as conn:
            result = await conn.execute(
                f"DELETE FROM {self.SCHEMA}.scheduled_tasks WHERE task_id = $1",
                task_id,
            )
        elapsed = (time.monotonic() - t0) * 1000
        deleted = result.endswith("1")
        if deleted:
            logger.info(f"🗑️ PG delete_scheduled_task: {task_id} ({elapsed:.0f}ms)")
        else:
            logger.warning(f"⚠️ PG delete_scheduled_task: {task_id} not found")
        return deleted

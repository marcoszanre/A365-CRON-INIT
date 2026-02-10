# Copyright (c) Microsoft. All rights reserved.

"""
Scheduled Task Management Tools

Local FunctionTool definitions that let the agent manage its own scheduled
tasks in PostgreSQL. The cron scheduler picks up these tasks and executes
them autonomously on each tick.

These tools are scoped to the agent's own UPN — an agent can only see and
manage its own tasks. The tools are created as closures so the ``agent_upn``
is captured at registration time.

Usage:
    tools = create_task_tools(agent_upn="UPN.agent@tenant.onmicrosoft.com")
    # Then add to agent: agent.default_options.setdefault("tools", []).extend(tools)
"""

import json
import logging
from typing import Annotated, Optional

from agent_framework import tool

logger = logging.getLogger(__name__)


def create_task_tools(agent_upn: str) -> list:
    """
    Create task management FunctionTool instances scoped to ``agent_upn``.

    Returns a list of tools that can be added to a ChatAgent's tool list.
    The agent_upn is captured in closures so the tools are scoped to that
    specific agent.
    """

    # ------------------------------------------------------------------
    # list_my_scheduled_tasks
    # ------------------------------------------------------------------
    @tool(name="list_my_scheduled_tasks", approval_mode="never_require")
    async def list_my_scheduled_tasks() -> str:
        """List all scheduled tasks assigned to me in the PostgreSQL database.
        Returns task details including task_id, name, prompt, enabled status,
        last run time, and last status. Use this to check what cron jobs are
        registered for execution."""
        from a365_agent.storage import get_storage

        storage = await get_storage()
        tasks = await storage.get_all_tasks_for_agent(agent_upn)
        if not tasks:
            return "No scheduled tasks found. You can create one with create_scheduled_task."

        result_lines = [f"Found {len(tasks)} task(s):\n"]
        for t in tasks:
            enabled = "✅ enabled" if t.get("is_enabled") else "❌ disabled"
            last_run = t.get("last_run_at") or "never"
            last_status = t.get("last_status") or "n/a"
            result_lines.append(
                f"• **{t['task_name']}** ({enabled})\n"
                f"  task_id: {t['task_id']}\n"
                f"  prompt: {t.get('task_prompt', '')[:120]}...\n"
                f"  last_run: {last_run} | status: {last_status}"
            )
        return "\n".join(result_lines)

    # ------------------------------------------------------------------
    # create_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="create_scheduled_task", approval_mode="never_require")
    async def create_scheduled_task(
        task_name: Annotated[str, "Short name for the task (e.g. 'weekly_report', 'daily_inbox_check')"],
        task_prompt: Annotated[str, "The prompt that describes what the task should do. "
                     "Supports {manager_email}, {agent_upn}, {timestamp} placeholders."],
        is_recurrent: Annotated[bool, "True if the task should repeat on each cron tick, "
                      "False for a one-time task"] = True,
    ) -> str:
        """Create a new scheduled task in the PostgreSQL database. The cron
        scheduler will execute this task automatically at each interval. One-time
        tasks (is_recurrent=False) are disabled after their first execution."""
        from a365_agent.storage import get_storage

        storage = await get_storage()
        row = await storage.create_scheduled_task(
            agent_user_id=agent_upn,
            task_name=task_name,
            task_prompt=task_prompt,
            is_enabled=True,
        )
        task_id = row.get("task_id", "unknown")
        recurrence = "recurrent" if is_recurrent else "one-time"
        logger.info(f"📋 Task created: {task_name} ({task_id}) for {agent_upn} [{recurrence}]")
        return (
            f"Task created successfully!\n"
            f"• task_id: {task_id}\n"
            f"• name: {task_name}\n"
            f"• recurrence: {recurrence}\n"
            f"• status: enabled\n"
            f"The cron scheduler will execute this task on the next interval."
        )

    # ------------------------------------------------------------------
    # update_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="update_scheduled_task", approval_mode="never_require")
    async def update_scheduled_task(
        task_id: Annotated[str, "The UUID task_id of the task to update"],
        task_name: Annotated[Optional[str], "New name for the task (leave empty to keep current)"] = None,
        task_prompt: Annotated[Optional[str], "New prompt for the task (leave empty to keep current)"] = None,
        is_enabled: Annotated[Optional[bool], "Set to true to enable or false to disable the task"] = None,
    ) -> str:
        """Update an existing scheduled task's name, prompt, or enabled status.
        Use list_my_scheduled_tasks first to get the task_id."""
        from a365_agent.storage import get_storage

        fields = {}
        if task_name is not None:
            fields["task_name"] = task_name
        if task_prompt is not None:
            fields["task_prompt"] = task_prompt
        if is_enabled is not None:
            fields["is_enabled"] = is_enabled

        if not fields:
            return "No fields provided to update. Specify at least one of: task_name, task_prompt, is_enabled."

        storage = await get_storage()
        updated = await storage.update_scheduled_task_fields(task_id, **fields)
        if updated is None:
            return f"Task with task_id '{task_id}' was not found. Use list_my_scheduled_tasks to see your tasks."

        return (
            f"Task updated successfully!\n"
            f"• task_id: {task_id}\n"
            f"• name: {updated.get('task_name')}\n"
            f"• enabled: {updated.get('is_enabled')}\n"
            f"Changes will take effect on the next cron tick."
        )

    # ------------------------------------------------------------------
    # delete_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="delete_scheduled_task", approval_mode="never_require")
    async def delete_scheduled_task(
        task_id: Annotated[str, "The UUID task_id of the task to delete"],
    ) -> str:
        """Permanently delete a scheduled task from the database. This cannot
        be undone. Use list_my_scheduled_tasks first to get the task_id."""
        from a365_agent.storage import get_storage

        storage = await get_storage()
        deleted = await storage.delete_scheduled_task(task_id)
        if not deleted:
            return f"Task with task_id '{task_id}' was not found. Use list_my_scheduled_tasks to see your tasks."

        return f"Task '{task_id}' has been permanently deleted."

    return [
        list_my_scheduled_tasks,
        create_scheduled_task,
        update_scheduled_task,
        delete_scheduled_task,
    ]

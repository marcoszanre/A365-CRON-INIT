# Copyright (c) Microsoft. All rights reserved.

"""
Scheduled Task Management Tools

Local FunctionTool definitions for managing cron-scheduled tasks in PostgreSQL.
The agent uses these tools to create, list, update, and delete tasks that the
background cron scheduler picks up and executes autonomously.

Scoped per agent UPN — each agent can only manage its own tasks.

Usage:
    tools = create_task_tools(agent_upn="UPN.agent@tenant.onmicrosoft.com")
    # Then register on agent: agent.default_options["tools"].extend(tools)
"""

import logging
from typing import Annotated, Optional

from agent_framework import tool

logger = logging.getLogger(__name__)


def create_task_tools(agent_upn: str) -> list:
    """
    Create task-management FunctionTools scoped to *agent_upn*.

    Returns a list of tools to register on a ChatAgent.
    """

    # ------------------------------------------------------------------
    # list_my_scheduled_tasks
    # ------------------------------------------------------------------
    @tool(name="list_my_scheduled_tasks", approval_mode="never_require")
    async def list_my_scheduled_tasks() -> str:
        """Retrieve every scheduled task that belongs to this agent from the
        PostgreSQL database (both enabled and disabled).

        Returns a formatted list with each task's task_id, name, prompt,
        enabled/disabled status, last execution time, and last execution
        result.

        Call this tool when:
        - The user asks "what tasks do I have?" or "show my tasks"
        - You need to look up a task_id before updating or deleting a task
        - The user wants to verify a task was created

        This tool takes NO parameters — it automatically filters to the
        current agent's tasks."""
        from a365_agent.storage import get_storage

        storage = await get_storage()
        tasks = await storage.get_all_tasks_for_agent(agent_upn)
        if not tasks:
            return "You have no scheduled tasks yet."

        lines = []
        for i, t in enumerate(tasks, 1):
            status = "Enabled" if t.get("is_enabled") else "Disabled"
            last_run = str(t.get("last_run_at") or "Never run")
            last_status = t.get("last_status") or "-"
            prompt_preview = (t.get("task_prompt") or "")[:100]
            lines.append(
                f"{i}. {t['task_name']} ({status})\n"
                f"   ID: {t['task_id']}\n"
                f"   Prompt: {prompt_preview}\n"
                f"   Last run: {last_run} ({last_status})"
            )
        return f"Your scheduled tasks ({len(tasks)}):\n\n" + "\n\n".join(lines)

    # ------------------------------------------------------------------
    # create_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="create_scheduled_task", approval_mode="never_require")
    async def create_scheduled_task(
        task_name: Annotated[
            str,
            "A short snake_case name for the task, e.g. 'weekly_report' or 'happiness_quote'.",
        ],
        task_prompt: Annotated[
            str,
            "The full prompt the cron agent will execute. Write it as a complete "
            "instruction, e.g. 'Send a Teams message to {manager_email} with an "
            "inspiring quote about happiness.' "
            "Supported placeholders (auto-resolved at runtime): "
            "{manager_email} = the manager's email, "
            "{agent_upn} = this agent's UPN, "
            "{timestamp} = current UTC time. "
            "DO NOT look up emails before calling this tool — just use the "
            "placeholders and they will be filled in automatically.",
        ],
        is_recurrent: Annotated[
            bool,
            "True (default) = repeats every cron interval. "
            "False = runs once then auto-disables.",
        ] = True,
    ) -> str:
        """Create a new scheduled task in PostgreSQL so the background cron job
        will execute it automatically on the next interval.

        IMPORTANT — call this tool IMMEDIATELY when the user asks you to
        create, schedule, or register a task. Do NOT call getMyProfile,
        getUserProfile, or any other tool first. The task_prompt supports
        {manager_email} and other placeholders that are resolved at execution
        time, so you never need to look up emails beforehand.

        Examples of when to call this tool:
        - "Create a task to send me a quote about happiness"
        - "Schedule a weekly report"
        - "Add a recurring reminder to check my inbox"
        """
        try:
            from a365_agent.storage import get_storage

            logger.info(f"Creating task '{task_name}' for agent {agent_upn}")
            storage = await get_storage()
            row = await storage.create_scheduled_task(
                agent_user_id=agent_upn,
                task_name=task_name,
                task_prompt=task_prompt,
                is_enabled=True,
            )
            task_id = row.get("task_id", "unknown")
            recurrence = "Repeats every cycle" if is_recurrent else "Runs once"
            logger.info(f"Task created: {task_name} ({task_id}) for {agent_upn}")
            return (
                f"Task created: {task_name}\n"
                f"ID: {task_id}\n"
                f"Schedule: {recurrence}\n"
                f"Status: Enabled\n"
                f"It will run automatically on the next cron cycle."
            )
        except Exception as e:
            logger.error(f"create_scheduled_task failed: {e}")
            return f"Error creating task: {e}"

    # ------------------------------------------------------------------
    # update_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="update_scheduled_task", approval_mode="never_require")
    async def update_scheduled_task(
        task_id: Annotated[
            str,
            "The UUID task_id to update. Get it from list_my_scheduled_tasks.",
        ],
        task_name: Annotated[
            Optional[str],
            "New name for the task. Omit or pass null to keep the current name.",
        ] = None,
        task_prompt: Annotated[
            Optional[str],
            "New prompt for the task. Omit or pass null to keep the current prompt.",
        ] = None,
        is_enabled: Annotated[
            Optional[bool],
            "Set true to enable or false to disable the task. "
            "Omit or pass null to keep the current state.",
        ] = None,
    ) -> str:
        """Update an existing scheduled task's name, prompt, or enabled status
        in PostgreSQL.

        Call this tool when the user wants to:
        - Rename a task
        - Change what a task does (update the prompt)
        - Enable or disable a task without deleting it

        You must provide the task_id. If you don't have it, call
        list_my_scheduled_tasks first to look it up."""
        try:
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
                return f"Task {task_id} not found. Use list_my_scheduled_tasks to check your tasks."

            status = "Enabled" if updated.get("is_enabled") else "Disabled"
            return (
                f"Task updated: {updated.get('task_name')}\n"
                f"ID: {task_id}\n"
                f"Status: {status}\n"
                f"Changes take effect on the next cron cycle."
            )
        except Exception as e:
            logger.error(f"update_scheduled_task failed: {e}")
            return f"Error updating task: {e}"

    # ------------------------------------------------------------------
    # delete_scheduled_task
    # ------------------------------------------------------------------
    @tool(name="delete_scheduled_task", approval_mode="never_require")
    async def delete_scheduled_task(
        task_id: Annotated[
            str,
            "The UUID task_id to delete. Get it from list_my_scheduled_tasks.",
        ],
    ) -> str:
        """Permanently delete a scheduled task from PostgreSQL. This cannot be
        undone.

        Call this tool when the user asks to remove or delete a specific task.
        You must provide the task_id. If you don't have it, call
        list_my_scheduled_tasks first to look it up."""
        try:
            from a365_agent.storage import get_storage

            storage = await get_storage()
            deleted = await storage.delete_scheduled_task(task_id)
            if not deleted:
                return f"Task {task_id} not found. Use list_my_scheduled_tasks to check your tasks."

            return f"Task deleted. ID: {task_id}"
        except Exception as e:
            logger.error(f"delete_scheduled_task failed: {e}")
            return f"Error deleting task: {e}"

    return [
        list_my_scheduled_tasks,
        create_scheduled_task,
        update_scheduled_task,
        delete_scheduled_task,
    ]

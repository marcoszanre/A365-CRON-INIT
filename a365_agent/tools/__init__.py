# Copyright (c) Microsoft. All rights reserved.

"""
Local Tools Module

Provides native FunctionTool definitions that run locally (no MCP round-trip).
These tools are registered alongside MCP servers on the ChatAgent and are
automatically invoked by the agent framework's function-calling pipeline.
"""

from a365_agent.tools.task_tools import create_task_tools

__all__ = ["create_task_tools"]

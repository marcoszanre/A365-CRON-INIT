# Copyright (c) Microsoft. All rights reserved.

"""
PostgreSQL Storage Module

Replaces SharePoint list storage with PostgreSQL for multi-agent MCP server environments.
Provides async storage for agent registry, conversations, shared state, tool execution
logging, and inter-agent task coordination.
"""

from a365_agent.storage.pg_storage import (
    PostgresStorage,
    get_storage,
)

__all__ = ["PostgresStorage", "get_storage"]

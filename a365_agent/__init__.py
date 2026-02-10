# Copyright (c) Microsoft. All rights reserved.

"""
A365 Agent Framework Package

A modular Python framework for building AI agents integrated with Microsoft 365
using the Agent 365 SDK.

Modules:
    - config: Configuration and environment management
    - auth: Authentication utilities and token management
    - observability: Telemetry and tracing setup
    - notifications: Notification handlers (email, Word, Excel, PowerPoint, lifecycle)
    - mcp: MCP (Model Context Protocol) server integration
    - base: Abstract base class for agents
    - host: Generic agent host server

Usage:
    from a365_agent import AgentBase, create_and_run_host
    from a365_agent.config import Settings
"""

from a365_agent.base import AgentBase
from a365_agent.host import GenericAgentHost, create_and_run_host
from a365_agent.proactive import ProactiveScheduler, ProactiveTokenProvider

__all__ = [
    "AgentBase",
    "GenericAgentHost",
    "create_and_run_host",
    "ProactiveScheduler",
    "ProactiveTokenProvider",
]

__version__ = "0.1.0"

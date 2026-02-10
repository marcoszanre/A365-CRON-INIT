# Copyright (c) Microsoft. All rights reserved.

"""
Proactive Module

Provides autonomous (proactive) agent capabilities:
- Token acquisition via Agent User Impersonation (3-step flow)
- Mock auth/context objects for headless MCP initialization
- Cron-based scheduler for periodic agent tasks
"""

from a365_agent.proactive.auth import AgentCredentials, ProactiveTokenProvider
from a365_agent.proactive.mock_context import (
    MockAuthorization,
    MockTurnContext,
)
from a365_agent.proactive.scheduler import ProactiveScheduler

__all__ = [
    "AgentCredentials",
    "ProactiveTokenProvider",
    "MockAuthorization",
    "MockTurnContext",
    "ProactiveScheduler",
]

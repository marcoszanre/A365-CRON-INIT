# Copyright (c) Microsoft. All rights reserved.

"""
Mock Auth / Context for Proactive Scenarios

Provides lightweight stand-ins for the Microsoft Agents SDK
Authorization and TurnContext objects. These are needed by the
MCP tool registration service but have no real Bot Framework
connection behind them in proactive (headless) mode.
"""


class MockAuthorization:
    """Mock Authorization that returns a pre-acquired bearer token."""

    def __init__(self, bearer_token: str):
        self._token = bearer_token

    async def get_token_async(self, *args, **kwargs) -> str:
        return self._token


class MockActivity:
    """Minimal Activity stub."""

    def __init__(self, user_id: str):
        self.from_property = _MockFrom(user_id)
        self.conversation = _MockConversation()


class MockTurnContext:
    """Mock TurnContext for proactive scenarios (no real Bot Framework connection)."""

    def __init__(self, user_id: str = "proactive-cron"):
        self.activity = MockActivity(user_id)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

class _MockFrom:
    def __init__(self, user_id: str):
        self.id = user_id


class _MockConversation:
    def __init__(self):
        self.id = "proactive-cron-conversation"

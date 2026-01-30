# Copyright (c) Microsoft. All rights reserved.

"""
Agent Base Class
Defines the abstract base class that agents must inherit from to work with the generic host.
"""

from abc import ABC, abstractmethod
from typing import Optional

from microsoft_agents.hosting.core import Authorization, TurnContext


class AgentInterface(ABC):
    """
    Abstract base class that any hosted agent must inherit from.

    This ensures agents implement the required methods at class definition time,
    providing stronger guarantees than a Protocol.
    """

    @abstractmethod
    async def initialize(self) -> None:
        """Initialize the agent and any required resources."""
        pass

    @abstractmethod
    async def process_user_message(
        self, message: str, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Process a user message and return a response."""
        pass

    @abstractmethod
    async def cleanup(self) -> None:
        """Clean up any resources used by the agent."""
        pass

    # =========================================================================
    # NOTIFICATION HANDLERS
    # =========================================================================

    async def handle_email_notification(
        self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Handle email notification. Override to customize behavior."""
        return "Email notification received."

    async def handle_word_notification(
        self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Handle Word document comment notification. Override to customize behavior."""
        return "Word comment notification received."

    async def handle_excel_notification(
        self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Handle Excel document comment notification. Override to customize behavior."""
        return "Excel comment notification received."

    async def handle_powerpoint_notification(
        self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Handle PowerPoint document comment notification. Override to customize behavior."""
        return "PowerPoint comment notification received."

    async def handle_lifecycle_notification(
        self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
    ) -> str:
        """Handle agent lifecycle notification. Override to customize behavior."""
        return "Lifecycle notification received."


def check_agent_inheritance(agent_class) -> bool:
    """
    Check that an agent class inherits from AgentInterface.

    Args:
        agent_class: The agent class to check

    Returns:
        True if the agent inherits from AgentInterface, False otherwise
    """
    if not issubclass(agent_class, AgentInterface):
        print(f"❌ Agent {agent_class.__name__} does not inherit from AgentInterface")
        return False
    return True


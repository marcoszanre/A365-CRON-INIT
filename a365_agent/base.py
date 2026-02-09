# Copyright (c) Microsoft. All rights reserved.

"""
Agent Base Module

Defines the abstract base class that all agents must inherit from
to work with the GenericAgentHost.
"""

from abc import ABC, abstractmethod
from typing import Optional

from microsoft_agents.hosting.core import Authorization, TurnContext


class AgentBase(ABC):
    """
    Abstract base class for A365 agents.
    
    Any agent hosted by GenericAgentHost must inherit from this class
    and implement the required abstract methods.
    
    Required Methods:
        - initialize(): Set up agent resources
        - process_user_message(): Handle user messages
        - cleanup(): Clean up resources
        
    Optional Notification Handlers:
        Override these to handle specific notification types:
        - handle_email_notification()
        - handle_word_notification()
        - handle_excel_notification()
        - handle_powerpoint_notification()
        - handle_lifecycle_notification()
    """

    @abstractmethod
    async def initialize(self) -> None:
        """
        Initialize the agent and any required resources.
        
        Called once when the agent is first started.
        """
        pass

    @abstractmethod
    async def process_user_message(
        self,
        message: str,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Process a user message and return a response.
        
        Args:
            message: The user's message text
            auth: Authorization handler for token operations
            auth_handler_name: Name of the auth handler (e.g., "AGENTIC")
            context: The TurnContext for the current conversation
            
        Returns:
            The agent's response text
        """
        pass

    @abstractmethod
    async def cleanup(self) -> None:
        """
        Clean up any resources used by the agent.
        
        Called when the server is shutting down.
        """
        pass

    # =========================================================================
    # NOTIFICATION HANDLERS (Override to customize)
    # =========================================================================

    async def handle_email_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Handle email notification.
        
        Triggered when the agent is mentioned in an email.
        Override to provide custom email handling logic.
        
        Note: Email channels have strict timeouts (~30s). Keep processing fast.
        """
        return "Email notification received."

    async def handle_word_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Handle Word document comment notification.
        
        Triggered when the agent is mentioned in a Word document comment.
        Override to provide custom Word document handling logic.
        """
        return "Word comment notification received."

    async def handle_excel_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Handle Excel document comment notification.
        
        Triggered when the agent is mentioned in an Excel document comment.
        Override to provide custom Excel document handling logic.
        """
        return "Excel comment notification received."

    async def handle_powerpoint_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Handle PowerPoint document comment notification.
        
        Triggered when the agent is mentioned in a PowerPoint document comment.
        Override to provide custom PowerPoint document handling logic.
        """
        return "PowerPoint comment notification received."

    async def handle_lifecycle_notification(
        self,
        notification_activity,
        auth: Authorization,
        auth_handler_name: Optional[str],
        context: TurnContext,
    ) -> str:
        """
        Handle agent lifecycle notification.
        
        Lifecycle Event Types:
        - agenticUserIdentityCreated: User identity created
        - agenticUserWorkloadOnboardingUpdated: Workload onboarding updated
        - agenticUserDeleted: User identity deleted
        
        Override to perform initialization, cleanup, or state management.
        Note: Lifecycle notifications don't send replies.
        """
        return "Lifecycle notification received."


def check_agent_inheritance(agent_class: type) -> bool:
    """
    Verify that an agent class inherits from AgentBase.
    
    Args:
        agent_class: The agent class to check
        
    Returns:
        True if the class inherits from AgentBase, False otherwise
    """
    if not issubclass(agent_class, AgentBase):
        print(f"❌ Agent {agent_class.__name__} must inherit from AgentBase")
        return False
    return True

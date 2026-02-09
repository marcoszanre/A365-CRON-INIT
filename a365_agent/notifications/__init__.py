# Copyright (c) Microsoft. All rights reserved.

"""
Notifications Module

Handlers for various notification types from Microsoft 365:
- Email notifications
- Word document comment notifications
- Excel document comment notifications
- PowerPoint document comment notifications
- Agent lifecycle notifications
"""

from a365_agent.notifications.handlers import (
    NotificationHandlerMixin,
    safe_send_activity,
    safe_send_email_response,
)

__all__ = [
    "NotificationHandlerMixin",
    "safe_send_activity",
    "safe_send_email_response",
]

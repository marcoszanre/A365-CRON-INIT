# Copyright (c) Microsoft. All rights reserved.

"""
Notification Handlers

Safe handlers for sending responses to various notification channels.
These handle timeouts and errors gracefully.
"""

import logging
from typing import Optional

from aiohttp.client_exceptions import ClientResponseError
from microsoft_agents.hosting.core import TurnContext
from microsoft_agents_a365.notifications import EmailResponse

logger = logging.getLogger(__name__)


# =============================================================================
# SAFE RESPONSE HELPERS
# =============================================================================

async def safe_send_activity(context: TurnContext, message: str) -> bool:
    """
    Safely send an activity, handling 404 errors gracefully.
    
    404 errors occur when the notification channel times out before we can respond.
    In this case, we log the issue but don't crash.
    
    Args:
        context: The turn context
        message: The message to send
        
    Returns:
        True if sent successfully, False otherwise
    """
    try:
        await context.send_activity(message)
        logger.info("✅ Activity sent successfully")
        return True
    except ClientResponseError as e:
        if e.status == 404:
            logger.warning(f"⚠️ Reply window expired (404). Message was: {message[:100]}...")
            return False
        logger.error(f"❌ Failed to send activity: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ Unexpected error sending activity: {e}")
        return False


async def safe_send_email_response(context: TurnContext, response: str) -> bool:
    """
    Safely send an email response, handling 404 errors gracefully.
    
    Email responses use the special EmailResponse format.
    
    Args:
        context: The turn context
        response: The email response text
        
    Returns:
        True if sent successfully, False otherwise
    """
    try:
        response_activity = EmailResponse.create_email_response_activity(response)
        await context.send_activity(response_activity)
        logger.info("✅ Email response sent successfully")
        return True
    except ClientResponseError as e:
        if e.status == 404:
            logger.warning(
                f"⚠️ Email reply window expired (404). Response was: {response[:100]}... "
                f"The notification channel timed out, but the email was processed."
            )
            return False
        logger.error(f"❌ Failed to send email response: {e}")
        return False
    except Exception as e:
        logger.error(f"❌ Unexpected error sending email response: {e}")
        return False


# =============================================================================
# NOTIFICATION HANDLER MIXIN
# =============================================================================

class NotificationHandlerMixin:
    """
    Mixin class providing notification handling capabilities.
    
    Add this to your agent host to get standardized notification handling.
    Provides timeout constants and safe response methods.
    """
    
    # Timeout constants (seconds)
    EMAIL_NOTIFICATION_TIMEOUT = 25  # Email channel typically times out at ~30s
    DOC_NOTIFICATION_TIMEOUT = 25    # Word/Excel/PowerPoint
    
    async def _handle_notification_timeout(
        self,
        context: TurnContext,
        notification_type: str,
        is_email: bool = False,
    ) -> None:
        """Handle a notification that timed out during processing."""
        logger.warning(f"⚠️ {notification_type} processing timeout")
        
        message = "Thank you for your message. I'm still processing and will respond shortly."
        
        if is_email:
            await safe_send_email_response(context, message)
        else:
            await safe_send_activity(context, message)
    
    async def _handle_notification_error(
        self,
        context: TurnContext,
        notification_type: str,
        error: Exception,
        is_email: bool = False,
    ) -> None:
        """Handle an error that occurred during notification processing."""
        logger.error(f"❌ {notification_type} notification error: {error}")
        
        message = f"Thank you for your message. I encountered an issue but will review it."
        
        if is_email:
            await safe_send_email_response(context, message)
        else:
            await safe_send_activity(context, message)

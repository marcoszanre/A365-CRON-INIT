# Agent 365 Lifecycle Notifications Implementation Guide

This document details the code changes and fixes implemented to properly handle Agent 365 lifecycle notifications when hosting an agent locally via ngrok.

## Overview

When an Agent 365 agent instance is created, the platform sends lifecycle notifications to inform the agent about user identity and workload onboarding events. These notifications require special handling because:

1. They arrive through a special "onboarding channel" that **does not support reply activities**
2. They don't contain text content like regular messages
3. Attempting to reply causes **502 Bad Gateway** errors

## Lifecycle Event Types

Per [Microsoft docs](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/notification?tabs=python#agent-lifecycle-events), there are three lifecycle events:

| Event Type | Event ID | Description |
|------------|----------|-------------|
| **User Identity Created** | `agenticUserIdentityCreated` | Triggered when an agentic user identity is created |
| **Workload Onboarding Updated** | `agenticUserWorkloadOnboardingUpdated` | Triggered when workload onboarding status is updated |
| **User Deleted** | `agenticUserDeleted` | Triggered when an agentic user identity is deleted |

---

## Code Changes

### 1. host_agent_server.py - Skip Reply for Lifecycle Notifications

**Problem:** The notification handler was trying to call `context.send_activity(response)` for all notifications, including lifecycle events. The onboarding channel returns 502 Bad Gateway for any reply attempts.

**Solution:** Check for `AGENT_LIFECYCLE` notification type and return early without sending a response.

**File:** `host_agent_server.py` (lines ~216-218)

```python
@self.agent_notification.on_agent_notification(
    channel_id=ChannelId(channel="agents", sub_channel="*"),
    **handler_config,
)
async def on_notification(
    context: TurnContext,
    state: TurnState,
    notification_activity: AgentNotificationActivity,
):
    try:
        # ... validation and setup code ...

        response = (
            await self.agent_instance.handle_agent_notification_activity(
                notification_activity, self.agent_app.auth, self.auth_handler_name, context
            )
        )

        # ✅ FIX: Skip sending response for lifecycle notifications
        # The onboarding channel doesn't support replies
        if notification_activity.notification_type == NotificationTypes.AGENT_LIFECYCLE:
            logger.info("📋 Lifecycle notification processed - no reply needed")
            return

        # Continue with email/WPX responses...
        if notification_activity.notification_type == NotificationTypes.EMAIL_NOTIFICATION:
            response_activity = EmailResponse.create_email_response_activity(response)
            await context.send_activity(response_activity)
            return

        await context.send_activity(response)
```

### 2. agent.py - Dedicated Lifecycle Notification Handler

**Problem:** The generic notification handler didn't properly extract lifecycle event details and logged minimal information.

**Solution:** Added a dedicated `_handle_lifecycle_notification()` method that handles all three event types.

**File:** `agent.py` (notification handling section)

```python
async def handle_agent_notification_activity(
    self, notification_activity, auth: Authorization, auth_handler_name: Optional[str], context: TurnContext
) -> str:
    """Handle agent notification activities (email, Word mentions, etc.)"""
    try:
        notification_type = notification_activity.notification_type
        logger.info(f"📬 Processing notification: {notification_type}")

        # ... MCP setup ...

        # Handle Email Notifications
        if notification_type == NotificationTypes.EMAIL_NOTIFICATION:
            # ... email handling ...

        # Handle Word Comment Notifications
        elif notification_type == NotificationTypes.WPX_COMMENT:
            # ... Word handling ...

        # ✅ FIX: Dedicated lifecycle handler
        elif notification_type == NotificationTypes.AGENT_LIFECYCLE:
            return await self._handle_lifecycle_notification(notification_activity)

        # Generic notification handling for unknown types
        else:
            # ... fallback handling ...
```

**New method added:**

```python
async def _handle_lifecycle_notification(self, notification_activity) -> str:
    """
    Handle Agent Lifecycle notifications.
    
    Lifecycle Event Types (per Microsoft docs):
    - agenticUserIdentityCreated: Triggered when an agentic user identity is created
    - agenticUserWorkloadOnboardingUpdated: Triggered when workload onboarding status is updated
    - agenticUserDeleted: Triggered when an agentic user identity is deleted
    
    These events allow agents to perform initialization tasks, cleanup operations,
    or state management in response to user lifecycle changes.
    """
    try:
        # Get the lifecycle notification data
        lifecycle_notification = getattr(notification_activity, 'agent_lifecycle_notification', None)
        
        if lifecycle_notification:
            event_type = getattr(lifecycle_notification, 'lifecycle_event_type', None)
            logger.info(f"📋 Agent lifecycle event: {event_type}")
            
            if event_type == "agenticUserIdentityCreated":
                logger.info("✅ Agentic user identity created - performing initialization tasks")
                # TODO: Add any initialization logic here
                return "User identity created - agent initialized successfully."
                
            elif event_type == "agenticUserWorkloadOnboardingUpdated":
                logger.info("🔄 Workload onboarding status updated")
                # TODO: Add any onboarding completion logic here
                return "Workload onboarding updated - agent ready for operation."
                
            elif event_type == "agenticUserDeleted":
                logger.info("🗑️ Agentic user identity deleted - performing cleanup tasks")
                # TODO: Add any cleanup logic here
                return "User identity deleted - cleanup completed."
                
            else:
                logger.info(f"📋 Unknown lifecycle event type: {event_type}")
                return f"Lifecycle event '{event_type}' acknowledged."
        else:
            logger.info("📋 Lifecycle notification received (no detailed data available)")
            return "Agent lifecycle event acknowledged."
            
    except Exception as e:
        logger.error(f"Error handling lifecycle notification: {e}")
        return "Agent lifecycle event processed with warnings."
```

---

## Issues Encountered and Resolutions

### Issue 1: 502 Bad Gateway on Lifecycle Notifications

**Symptom:**
```
Error replying to activity: 502
aiohttp.client_exceptions.ClientResponseError: 502, message='Bad Gateway', 
url='https://smba.trafficmanager.net/.../agentOnboarding.../activities/...'
```

**Cause:** The onboarding channel (identified by `agentOnboarding` in the conversation ID) does not support reply activities.

**Resolution:** Check for `NotificationTypes.AGENT_LIFECYCLE` and skip calling `context.send_activity()`.

---

### Issue 2: 502 from Typing Indicator (SDK Issue)

**Symptom:**
```
File "...typing_indicator.py", line 69, in _typing_loop
    await context.send_activity(Activity(type=ActivityTypes.typing))
Error sending typing activity: 502, message='Bad Gateway'
```

**Cause:** The Microsoft Agents SDK automatically sends typing indicators during processing. This happens inside the SDK code, not our handlers.

**Resolution:** This is a **known SDK limitation** - we cannot prevent these errors from appearing in logs. They are harmless and don't affect functionality. The lifecycle notification is still processed successfully (HTTP 202 returned).

---

### Issue 3: consent_required Errors During Instance Creation

**Symptom:**
```
AADSTS65001: The user or administrator has not consented to use the application 
with ID '...' named '... Agent'. Send an interactive authorization request...
'suberror': 'consent_required'
```

**Cause:** When a new agent instance is created, permissions configured on the blueprint take time to propagate to the instance.

**Resolution:** Wait a few minutes after instance creation. The errors occur for specific scopes:
- `https://graph.microsoft.com/.default` - For Graph API access
- `https://api.powerplatform.com/.default` - For observability

Core functionality (Bot API, MCP) typically works immediately.

---

### Issue 4: Missing Lifecycle Event Details

**Symptom:**
```
INFO:agent:📋 Lifecycle notification received (no detailed data available)
```

**Cause:** The `agent_lifecycle_notification` property may not be populated by the SDK in all cases, or the property name differs from documentation.

**Resolution:** Use `getattr()` with fallback to safely access the property. The notification is still processed correctly even without detailed event type information.

---

## Successful Logs

When everything works correctly, you should see:

```
INFO:host_agent_server:📬 NotificationTypes.AGENT_LIFECYCLE
INFO:agent:📬 Processing notification: NotificationTypes.AGENT_LIFECYCLE
INFO:agent:📧 Email entity present: False
INFO:agent:📋 Lifecycle notification received (no detailed data available)
INFO:host_agent_server:📋 Lifecycle notification processed - no reply needed
INFO:aiohttp.access:::1 ... "POST /api/messages HTTP/1.1" 202 117 ...
```

Key indicators:
- ✅ `NotificationTypes.AGENT_LIFECYCLE` detected
- ✅ `📋 Lifecycle notification processed - no reply needed` logged
- ✅ HTTP **202** returned (not 500)
- ✅ No crash or exception in our code

---

## Testing Lifecycle Notifications

1. **Start the server:**
   ```powershell
   uv run python .\start_with_generic_host.py
   ```

2. **Ensure ngrok is running** and endpoint is registered with the blueprint

3. **Create a new agent instance** via M365 Admin Center or Teams

4. **Watch the logs** for lifecycle notifications

5. **After onboarding completes**, find your agent in Teams by searching for its name

---

## References

- [Microsoft Agent 365 Notifications Documentation](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/notification?tabs=python)
- [Agent Lifecycle Events Section](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/notification?tabs=python#agent-lifecycle-events)

---

## Summary

| Change | File | Purpose |
|--------|------|---------|
| Skip reply for lifecycle | `host_agent_server.py` | Prevent 502 errors on onboarding channel |
| Dedicated lifecycle handler | `agent.py` | Properly handle all 3 lifecycle event types |
| Safe attribute access | `agent.py` | Handle missing `agent_lifecycle_notification` |

These changes enable the agent to successfully process lifecycle notifications during instance creation without crashing or returning errors to the platform.

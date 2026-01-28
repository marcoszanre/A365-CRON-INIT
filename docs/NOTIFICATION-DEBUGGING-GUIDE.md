# Agent 365 Email Notification - SDK & Documentation Issues

This document details issues encountered when implementing email notification handling in a Python Agent 365 agent using the `microsoft_agents_a365` SDK. These issues stem from gaps in SDK documentation, unexpected API behavior, and tooling limitations.

## Overview

**Goal:** Receive email notifications via Agent 365 and trigger MCP tool invocations (e.g., `WordCreateNewDocument`) based on email content.

**Environment:**
- Python 3.13.11 with `microsoft_agents_a365` SDK
- MCP Server integration via Agent 365

---

## Issues Encountered

### Issue 1: Invalid NotificationTypes Enum Value

**Symptom:**
```
ValueError: 'agent/notification' is not a valid NotificationTypes
```

**Root Cause:**
The activity `name` field was set to `"agent/notification"` which doesn't match the SDK's `NotificationTypes` enum values.

**SDK Enum Definition** (from `microsoft_agents_a365.notifications.agent_notification`):
```python
class NotificationTypes(str, Enum):
    EMAIL_NOTIFICATION = "emailNotification"
    WPX_COMMENT = "wpxComment"
    AGENT_LIFECYCLE = "agentLifecycle"
```

**Fix:**
Use the exact enum string value in the activity `name`:
```json
{
  "name": "emailNotification"  // ✅ Correct
  // "name": "agent/notification"  // ❌ Wrong
}
```

> **⚠️ Documentation Gap:** The `NotificationTypes` enum values are not documented. Developers must inspect SDK source code to discover valid values.

---

### Issue 2: Activity Type Must Be "message"

**Symptom:**
Email entity not being parsed despite correct `name` value.

**Root Cause:**
Activity `type` was set to `"event"` but the SDK notification handler expects `"message"`.

**Fix:**
```json
{
  "type": "message",  // ✅ Correct
  // "type": "event"  // ❌ Wrong - notification handler won't route correctly
  "name": "emailNotification"
}
```

> **⚠️ Documentation Gap:** The required activity `type` for notifications is not documented.

---

### Issue 3: Email Data Location - Must Use `entities` Array

**Symptom:**
`'AgentNotificationActivity' object has no attribute 'text'` and email entity showing as `None`.

**Root Cause:**
SDK documentation does not specify where notification data should be placed. The SDK parses from the `entities` array.

**Incorrect Structure:**
```json
{
  "type": "message",
  "name": "emailNotification",
  "channelData": {
    "email": { ... }  // ❌ SDK doesn't look here
  }
}
```

**Correct Structure:**
```json
{
  "type": "message",
  "name": "emailNotification",
  "entities": [
    {
      "type": "emailNotification",
      "id": "...",
      "conversation_id": "...",
      "html_body": "..."
    }
  ]
}
```

> **⚠️ Documentation Gap:** The expected location of notification data (`entities` array) is not documented.

---

### Issue 4: Field Names Must Use snake_case (Critical!)

**Symptom:**
`📧 Email entity present: True` but `html_body` showing as `None` despite being in the payload.

**Root Cause:**
The SDK uses **Pydantic models** which expect **snake_case** field names, but typical JSON APIs use camelCase. This mismatch is not documented and fails silently.

**SDK Model Definition** (from `microsoft_agents_a365`):
```python
class EmailReference(BaseModel):
    id: str
    conversation_id: str  # snake_case!
    html_body: str        # snake_case!
```

**Incorrect Payload (camelCase):**
```json
{
  "type": "emailNotification",
  "id": "...",
  "conversationId": "...",    // ❌ Pydantic won't map this
  "htmlBody": "..."           // ❌ Pydantic won't map this
}
```

**Correct Payload (snake_case):**
```json
{
  "type": "emailNotification",
  "id": "email-123",
  "conversation_id": "conv-456",  // ✅ Matches Pydantic model
  "html_body": "<body>...</body>" // ✅ Matches Pydantic model
}
```

> **⚠️ SDK Issue:** Pydantic models should include `alias` configuration to accept both camelCase (standard JSON) and snake_case. The current implementation silently fails when camelCase is used.

---

### Issue 5: MCP Server Authorization Headers Not Passed

**Symptom:**
MCP server requests failing with 401 Unauthorized, even though bearer token was configured.

**Root Cause:**
The `mcp` Python library's `streamablehttp_client` function ignores a `headers=` parameter when passed directly. Headers must be passed via an `httpx.AsyncClient` instance.

**Incorrect Code:**
```python
async with streamablehttp_client(
    url=server_url,
    headers={"Authorization": f"Bearer {token}"}  # ❌ Silently ignored!
) as client:
    ...
```

**Correct Code:**
```python
import httpx

http_client = httpx.AsyncClient(
    headers={"Authorization": f"Bearer {token}"}
)

async with streamablehttp_client(
    url=server_url,
    http_client=http_client  # ✅ Headers properly passed
) as client:
    ...
```

**Location:** `microsoft_agents_a365/tooling/extensions/agentframework/services/mcp_tool_registration_service.py`

> **⚠️ Documentation Gap:** The `streamablehttp_client` API is misleading. If `headers=` parameter is not supported, it should either raise an error or be removed from the function signature. Silent failure causes significant debugging effort.

---

### Issue 6: Agents Playground Strips `channelId` from Custom Activities

**Symptom:**
Notification activities sent via Agents Playground fail to route correctly to the notification handler, even when the payload appears correct.

**Root Cause:**
Agents Playground **strips the `channelId` field** from custom activities before forwarding them to the agent. The notification routing in the SDK relies on `channelId` being present to properly identify and route notification activities.

**Impact:**
- Cannot test email notifications through Agents Playground
- Developers must create alternative testing methods to bypass Agents Playground
- Delays development and debugging cycles

**Workaround:**
Send notification activities directly to the agent endpoint (e.g., via `curl` or HTTP client) instead of through Agents Playground.

> **⚠️ Tooling Issue:** Agents Playground should preserve all activity fields, including `channelId`, when forwarding custom activities. This limitation prevents proper notification testing during development.

---

## Complete Working Notification Payload

For reference, here's a complete working email notification payload:

```json
{
  "type": "message",
  "name": "emailNotification",
  "id": "test-notification-12345",
  "timestamp": "2026-01-28T05:00:00.000Z",
  "channelId": "agent365",
  "from": {
    "id": "system",
    "name": "Agent365 Notification Service"
  },
  "conversation": {
    "id": "00000000-0000-0000-0000-000000000001"
  },
  "recipient": {
    "id": "d6a93b83-2341-4783-89db-1838bb91c7b1",
    "name": "MyAgent"
  },
  "serviceUrl": "http://localhost:3978",
  "entities": [
    {
      "type": "emailNotification",
      "id": "550e8400-e29b-41d4-a716-446655440000",
      "conversation_id": "00000000-0000-0000-0000-000000000001",
      "html_body": "<body><p>Please create a Word document summarizing Q1 sales.</p></body>"
    }
  ]
}
```

---

## Summary of Issues

| # | Component | Issue | Impact |
|---|-----------|-------|--------|
| 1 | SDK Documentation | `NotificationTypes` enum values not documented | Developers guess activity `name` values |
| 2 | SDK Documentation | Activity `type` requirement not documented | Must be `"message"`, not `"event"` |
| 3 | SDK Documentation | Entity data location not documented | Must use `entities[]`, not `channelData` |
| 4 | SDK Pydantic Models | Expects snake_case, JSON typically uses camelCase | `html_body` works, `htmlBody` silently fails |
| 5 | MCP SDK | `headers=` parameter silently ignored | Must use `http_client=httpx.AsyncClient(headers=...)` |
| 6 | Agents Playground | Strips `channelId` from custom activities | Cannot test notifications via Playground |

---

## Recommendations

1. **Document `NotificationTypes` enum values** and required activity structure in SDK documentation
2. **Add Pydantic aliases** to accept both camelCase and snake_case for notification entity fields
3. **Fix `streamablehttp_client`** to either support `headers=` parameter or remove it from signature
4. **Add SDK validation** with clear error messages when notification payloads are malformed
5. **Provide complete example payloads** in documentation showing working notification structures
6. **Fix Agents Playground** to preserve `channelId` and other fields for custom activities

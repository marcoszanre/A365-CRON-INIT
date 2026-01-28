# MCP Server Invocation Fix

This document describes the changes required to fix the MCP server invocation issues in the `microsoft_agents_a365` package.

## Problem Summary

The agent was not correctly invoking MCP servers due to two issues:

1. **Empty `MCP_PLATFORM_ENDPOINT`**: The `.env` file had `MCP_PLATFORM_ENDPOINT=` (empty), causing the base URL to be empty instead of the production URL.

2. **`headers` parameter ignored**: The `MCPStreamableHTTPTool` class in the `agent_framework` package no longer accepts a `headers` parameter directly. Instead, it requires an `httpx.AsyncClient` to be passed via the `http_client` parameter.

---

## Fix 1: Environment Configuration

### File: `.env`

**Before:**
```env
MCP_PLATFORM_ENDPOINT=
```

**After:**
```env
MCP_PLATFORM_ENDPOINT=https://agent365.svc.cloud.microsoft
```

---

## Fix 2: MCP Tool Registration Service

### File: `.venv/Lib/site-packages/microsoft_agents_a365/tooling/extensions/agentframework/services/mcp_tool_registration_service.py`

### Change 1: Add httpx import

**Before (lines 1-6):**
```python
# Copyright (c) Microsoft. All rights reserved.

from typing import Optional, List, Any, Union
import logging

from agent_framework import ChatAgent, MCPStreamableHTTPTool
```

**After:**
```python
# Copyright (c) Microsoft. All rights reserved.

from typing import Optional, List, Any, Union
import logging
import httpx

from agent_framework import ChatAgent, MCPStreamableHTTPTool
```

### Change 2: Use httpx.AsyncClient instead of headers parameter

**Before (around lines 102-117):**
```python
                    # Prepare auth headers
                    headers = {}
                    if auth_token:
                        headers[Constants.Headers.AUTHORIZATION] = (
                            f"{Constants.Headers.BEARER_PREFIX} {auth_token}"
                        )

                    server_name = getattr(config, "mcp_server_name", "Unknown")

                    # Create and configure MCPStreamableHTTPTool
                    mcp_tools = MCPStreamableHTTPTool(
                        name=server_name,
                        url=server_url,
                        headers=headers,
                        description=f"MCP tools from {server_name}",
                    )
```

**After:**
```python
                    # Prepare auth headers
                    headers = {}
                    if auth_token:
                        headers[Constants.Headers.AUTHORIZATION] = (
                            f"{Constants.Headers.BEARER_PREFIX} {auth_token}"
                        )

                    server_name = getattr(config, "mcp_server_name", "Unknown")

                    # Create httpx client with auth headers (MCPStreamableHTTPTool requires http_client parameter)
                    http_client = httpx.AsyncClient(headers=headers)

                    # Create and configure MCPStreamableHTTPTool
                    mcp_tools = MCPStreamableHTTPTool(
                        name=server_name,
                        url=server_url,
                        http_client=http_client,
                        description=f"MCP tools from {server_name}",
                    )
```

---

## Root Cause Analysis

### Issue 1: Empty Base URL

The `_get_mcp_platform_base_url()` function in `utility.py` checks:

```python
if os.getenv("MCP_PLATFORM_ENDPOINT") is not None:
    return os.getenv("MCP_PLATFORM_ENDPOINT")
return MCP_PLATFORM_PROD_BASE_URL
```

When `MCP_PLATFORM_ENDPOINT=` is set (empty string), `os.getenv()` returns `""` (not `None`), so it returns an empty string instead of the default production URL.

### Issue 2: Headers Not Being Sent

The `MCPStreamableHTTPTool` constructor signature changed. The `headers` parameter is now captured by `**kwargs` and silently ignored:

```python
def __init__(
    self,
    name: str,
    url: str,
    *,
    # ... other params ...
    http_client: httpx.AsyncClient | None = None,
    **kwargs: Any,  # <-- headers goes here and is ignored!
) -> None:
```

The correct way to pass headers is through an `httpx.AsyncClient`:

```python
http_client = httpx.AsyncClient(headers={"Authorization": f"Bearer {token}"})
mcp_tools = MCPStreamableHTTPTool(name=name, url=url, http_client=http_client)
```

---

## Testing

After applying the fixes, the MCP server should correctly:

1. Build the full URL: `https://agent365.svc.cloud.microsoft/agents/servers/mcp_WordServer`
2. Send the Authorization header with the bearer token
3. Include the required `Accept: application/json, text/event-stream` header (handled by the MCP client library)

You can verify the MCP server is accessible by running:

```powershell
$token = "YOUR_BEARER_TOKEN"
$headers = @{
    Authorization = "Bearer $token"
    "Content-Type" = "application/json"
    Accept = "application/json, text/event-stream"
}
Invoke-RestMethod -Uri "https://agent365.svc.cloud.microsoft/agents/servers/mcp_WordServer" `
    -Method POST -Headers $headers `
    -Body '{"jsonrpc": "2.0", "method": "tools/list", "id": 1}'
```

---

## Note

These changes are applied to the installed package in `.venv`. If you reinstall or update the `microsoft_agents_a365` package, these changes will be overwritten. Consider reporting this issue to the package maintainers.

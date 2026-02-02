# Proactive Messaging Errors Reference

This document catalogs errors encountered during proactive messaging scenarios using the Agent User Impersonation flow.

## Overview

Proactive messaging requires the **agentic user** (not a human) to have proper licensing and permissions. The Agent User Impersonation flow uses `user_fic` grant type to acquire tokens that represent the agentic user identity.

## Error: Agentic User Missing Teams License

### Symptom

When attempting to send a Teams message via MCP, the `createChat` operation fails with a license-related error.

### Agent Response

```
I attempted to create a chat and send the message to [target]@[domain].com, 
but it failed because the user does not have a valid Office365 license assigned. 
Therefore, I could not create the chat or send the message.
```

### Root Cause

The **agentic user** (`MZHEROPYTHON01fe1c57@...` in our case) does not have a Teams license assigned, even though:
- The token acquisition flow (T1 → T2 → MCP Token) succeeds
- The MCP token contains valid `McpServers.Teams.All` scope
- The target user (Lisa) has a valid license

### Why This Happens

1. **Auth flow works without license** - The `user_fic` grant type only validates:
   - Blueprint credentials
   - Agent Identity → Agent User relationship
   - OAuth2 permission grants

2. **Teams API requires license** - When the MCP Teams Server calls the Graph API to create a chat, Teams validates that the **caller** (the agentic user represented by the token) has:
   - `TeamspaceAPI` service plan enabled
   - Valid Microsoft 365 license with Teams

### Verification

Check the agentic user's Teams license status:

```powershell
az rest --method GET `
  --url "https://graph.microsoft.com/beta/users/{agent-user-object-id}?`$select=displayName,assignedPlans" `
  | ConvertFrom-Json `
  | Select-Object -ExpandProperty assignedPlans `
  | Where-Object { $_.service -like "*Team*" } `
  | Format-Table service, capabilityStatus -AutoSize
```

**Working state:**
```
service                   capabilityStatus
-------                   ----------------
TeamspaceAPI              Enabled
LearningAppServiceInTeams Enabled
```

**Broken state:**
```
service                   capabilityStatus
-------                   ----------------
TeamspaceAPI              Deleted
LearningAppServiceInTeams Enabled
```

### Resolution

Assign a Microsoft 365 license with Teams to the agentic user:

1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Users** → **Active users**
3. Find the agentic user (e.g., `MZHEROPYTHON_01 Agent`)
4. Click **Licenses and apps**
5. Assign a license that includes Teams (e.g., Microsoft 365 E5)
6. Ensure **Microsoft Teams** service is enabled

### Note on Licensing Requirements

While the auth flow doesn't require a license, the underlying Microsoft 365 services do. Each MCP server may have different licensing requirements:

| MCP Server | License Requirement |
|------------|---------------------|
| `mcp_TeamsServer` | Teams license (TeamspaceAPI) |
| `mcp_MailTools` | Exchange Online license |
| `mcp_CalendarTools` | Exchange Online license |
| `mcp_ODSPRemoteServer` | SharePoint Online license |
| `mcp_WordServer` | Office apps license |

The agentic user should have a comprehensive license (e.g., Microsoft 365 E3/E5) to use all MCP capabilities.

---

## Related Documentation

- [Agent User Impersonation Flow](https://learn.microsoft.com/en-us/entra/agent-id/identity-platform/agent-user-oauth-flow)
- [Request Agent User Tokens](https://learn.microsoft.com/en-us/entra/agent-id/identity-platform/autonomous-agent-request-agent-user-tokens)
- [Test Script](../devTools/test_proactive_teams.py)

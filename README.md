# Agent 365 Python Framework Sample

A production-ready Python agent built with the **Microsoft Agent 365 SDK** and **AgentFramework SDK**. This agent demonstrates enterprise-grade AI capabilities:

- 🔐 **Agentic Identity** - Entra-backed identity with its own mailbox and Teams presence
- 📧 **Email Notifications** - Respond to emails via MCP Mail tools
- 💬 **Teams Integration** - Chat conversations with conversation history
- 👤 **Directory Search** - Look up users, managers, org structure
- 🔧 **MCP Tooling** - Access Microsoft 365 via governed MCP servers
- 📊 **Observability** - OpenTelemetry-based tracing
- ⚡ **Multi-Model Failover** - Automatic load balancing across Azure OpenAI models

## Prerequisites

- **Python 3.11+**
- **[uv](https://docs.astral.sh/uv/)** - Fast Python package manager
- **[Agent 365 CLI](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/cli)** - For blueprint setup
- **[ngrok](https://ngrok.com/)** - For local development tunneling
- **Azure OpenAI** - Deployed GPT model (e.g., `gpt-4o`, `gpt-4.1-mini`)
- **Microsoft 365 Tenant** - With Agent 365 preview access

## Quick Start

### 1. Clone and Install

```bash
git clone <repo-url>
cd agent365-agentframework-python

# Install dependencies with uv
uv sync
```

### 2. Initialize Agent 365 Blueprint

```bash
# Login and create blueprint
a365 login
a365 init

# Setup permissions and authentication
a365 setup
```

### 3. Configure Environment

Copy the template and the CLI will populate most values:

```bash
cp .env.template .env
```

Then fill in your **Azure OpenAI** configuration:

```env
# Model 1 (PRIMARY)
AZURE_OPENAI_MODEL_1_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_MODEL_1_DEPLOYMENT=gpt-4.1-mini
AZURE_OPENAI_MODEL_1_API_KEY=your-api-key

# Model 2 (FALLBACK - optional)
AZURE_OPENAI_MODEL_2_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_MODEL_2_DEPLOYMENT=gpt-4o-mini
AZURE_OPENAI_MODEL_2_API_KEY=your-api-key
```

### 4. Add MCP Servers

```bash
# Add the MCP servers you need
a365 develop add-mcp-servers mcp_TeamsServer
a365 develop add-mcp-servers mcp_MailTools
a365 develop add-mcp-servers mcp_MeServer
```

### 5. Get Development Token

```bash
# Get a token for local MCP testing
a365 develop get-token
```

### 6. Run the Agent

```bash
# Start ngrok tunnel
ngrok http 3978

# In another terminal, start the agent
uv run main.py
```

### 7. Update Endpoint

```bash
# Update the agent endpoint to your ngrok URL
a365 develop update-endpoint --url https://your-ngrok-url.ngrok-free.app/api/messages
```

## Project Structure

```
├── main.py                     # Entry point
├── agents/
│   ├── contoso_agent.py        # Main agent implementation
│   └── system_prompt.md        # Agent system prompt (editable)
├── a365_agent/
│   ├── host.py                 # HTTP server and notification routing
│   ├── base.py                 # Base agent class
│   ├── config.py               # Configuration and model pool
│   ├── auth.py                 # Authentication helpers
│   ├── mcp/                    # MCP service integration
│   └── notifications/          # Notification handlers
├── devTools/
│   ├── test_proactive_teams.py # Test proactive Teams messaging
│   ├── test_proactive_email.py # Test proactive email sending
│   └── test_directory_search.py# Test directory/user lookup
├── manifest/                   # Agent manifest for publishing
├── ToolingManifest.json        # MCP server configuration
├── .env.template               # Production environment template
└── .env.dev.template           # Development environment template
```

## Development Tools

The `devTools/` folder contains test scripts for proactive scenarios:

### Test Proactive Teams Message
```bash
# With Agent User Impersonation (message FROM agentic user):
$env:BEARER_TOKEN=""; uv run devTools/test_proactive_teams.py

# With delegated token (message FROM human):
a365 develop get-token
uv run devTools/test_proactive_teams.py
```

### Test Proactive Email
```bash
# With Agent User Impersonation (email FROM agentic user):
$env:BEARER_TOKEN=""; uv run devTools/test_proactive_email.py
```

### Test Directory Search
```bash
# Search for a user by name (configured in .env DIRECTORY_SEARCH_NAME)
uv run devTools/test_directory_search.py
```

## Agent User Impersonation

For proactive scenarios (background processes), the agent can act as its own identity using the 3-step Agent User Impersonation flow:

1. **T1**: Blueprint → Agent Identity (client_credentials + fmi_path)
2. **T2**: Agent Identity → Agent User exchange (jwt-bearer)
3. **MCP Token**: user_fic grant with T1 + T2 + user_id

This allows the agent to send Teams messages and emails **as itself**, not delegating from a human user.

### Required Configuration

```env
# Agent Identity (from blueprint)
AGENT_IDENTITY_CLIENT_ID=your-agent-identity-id

# Agent User (the agentic user with mailbox/Teams)
AGENT_USER_UPN=agentname@tenant.onmicrosoft.com
AGENT_USER_OBJECT_ID=user-object-id

# Target for testing
TARGET_USER_EMAIL=target@tenant.onmicrosoft.com
```

## MCP Servers

This sample supports these MCP servers:

| Server | Description | Key Tools |
|--------|-------------|-----------|
| `mcp_TeamsServer` | Microsoft Teams | `createChat`, `postMessage`, `listChatMessages` |
| `mcp_MailTools` | Outlook Email | `SendEmailWithAttachmentsAsync`, `ReplyAllToMessageAsync` |
| `mcp_MeServer` | User/Directory | `listUsers`, `getUserProfile`, `getUsersManager` |

Add more servers via:
```bash
a365 develop add-mcp-servers <server-name>
```

## System Prompt

The agent's personality and instructions are defined in `agents/system_prompt.md`. Edit this file to customize:

- Available MCP tools and how to use them
- Conversation context handling (Teams chat history)
- Email reply behavior (Reply vs Reply All)
- User lookup instructions

## Multi-Model Failover

The agent supports automatic failover across multiple Azure OpenAI models:

```env
# Model 1 (PRIMARY)
AZURE_OPENAI_MODEL_1_ENDPOINT=https://endpoint1.openai.azure.com
AZURE_OPENAI_MODEL_1_DEPLOYMENT=gpt-4.1-mini

# Model 2 (FALLBACK)
AZURE_OPENAI_MODEL_2_ENDPOINT=https://endpoint2.openai.azure.com
AZURE_OPENAI_MODEL_2_DEPLOYMENT=gpt-4o-mini
```

On 429 (rate limit) errors, the agent automatically switches to the next available model.

## Publishing

To publish the agent to your tenant:

```bash
# Build and publish
a365 publish
```

## Logging

Control log verbosity via `.env`:

```env
# General log level
LOG_LEVEL=INFO

# MCP tool call logging (DEBUG shows parameters)
AGENT_FRAMEWORK_LOG_LEVEL=DEBUG
```

## License

MIT License - See LICENSE file for details.

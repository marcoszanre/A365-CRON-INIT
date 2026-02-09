# Agent 365 Python Framework Sample

<p>
  <a href="https://youtu.be/HBr3PicYtNw" target="_blank"><img src="https://img.shields.io/badge/▶_Watch_Demo-YouTube-FF0000?style=flat-square&logo=youtube&logoColor=white" alt="Watch Demo"></a>
  <a href="https://learn.microsoft.com/en-us/microsoft-agent-365/" target="_blank"><img src="https://img.shields.io/badge/📚_Docs-Microsoft_Learn-0078D4?style=flat-square&logo=microsoft&logoColor=white" alt="Official Docs"></a>
</p>

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
- **[Agent 365 CLI](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/agent-365-cli)** - For blueprint setup and configuration
- **[Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli)** - For Azure authentication (`az login`)
- **[ngrok](https://ngrok.com/)** - **Required** for local development (exposes your agent to receive Teams/Email notifications)
- **[Agents Playground](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/test-with-toolkit-project)** - Optional local testing UI (`winget install agentsplayground`)
- **Azure OpenAI** - Deployed GPT model (e.g., `gpt-4o`, `gpt-4.1-mini`)
- **Microsoft 365 Tenant** - With [Agent 365 Frontier preview](https://adoption.microsoft.com/copilot/frontier-program/) access

## Quick Start

This sample supports two workflows based on the [Agent 365 Development Lifecycle](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/a365-dev-lifecycle):

| Mode | Use Case | Authentication | Endpoint |
|------|----------|----------------|----------|
| **Dev** | Quick local testing with Agents Playground | Bearer token (`a365 develop get-token`) | localhost |
| **Prod** | Full Teams/Email integration with real agent instance | Agentic auth (blueprint credentials) | ngrok static domain |

---

## 🛠️ Development Mode (Quick Testing)

Use this for rapid iteration with Agents Playground. No blueprint setup required.

### 1. Clone and Install

```bash
git clone <repo-url>
cd python-a365-sample
uv sync
```

### 2. Initialize Config (Dev Mode)

```bash
az login
a365 config init
```

When prompted, set `needDeployment: false` (you're running locally, not deploying to Azure).

### 3. Configure Environment

```bash
cp .env.dev.template .env
# PowerShell:
# copy .env.dev.template .env
```

**Configure Azure OpenAI** in `.env`:

```env
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-4.1-mini
AZURE_OPENAI_API_KEY=your-api-key
```

> **Note:** Dev mode uses single-model config. The template already has `AGENT_MODE=dev` and `USE_AGENTIC_AUTH=false` set.

### 4. Add MCP Servers

```bash
a365 develop add-permissions
a365 develop add-mcp-servers mcp_TeamsServer
a365 develop add-mcp-servers mcp_MailTools
a365 develop add-mcp-servers mcp_MeServer
```

### 5. Get Bearer Token

```bash
# Get tokens for local MCP testing (expires ~1 hour)
a365 develop get-token
```

This updates `BEARER_TOKEN` in your `.env` file.

> ⚠️ **Troubleshooting MCP Issues:** If you encounter MCP-related errors during development, try clearing your local token cache:
> 
> **Windows:**
> ```powershell
> Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\TokenBroker\a365" -Recurse -Force
> ```
> 
> **macOS/Linux:**
> ```bash
> rm -rf ~/.local/share/Microsoft/TokenBroker/a365
> ```
> 
> Then run `a365 develop get-token` again to refresh your tokens.

### 6. Customize System Prompt

Edit `agents/system_prompt.md` to define your agent's personality and tool instructions.

### 7. Run and Test

```bash
# Terminal 1: Start the agent
uv run python main.py
# PowerShell:
# uv run python .\main.py

# Terminal 2: Start Agents Playground
agentsplayground
```

Send test messages in Agents Playground to verify your agent works.

> **Note:** In dev mode with bearer token, you're testing as **your user identity**, not the agent's identity. Tool calls use your permissions.

---

## 🚀 Production Mode (Full Deployment)

Use this for real Teams/Email integration with an agentic user identity.

### 1. Clone and Install

```bash
git clone <repo-url>
cd agent365-agentframework-python
uv sync
```

### 2. Setup ngrok Static Domain

Get a [free static domain from ngrok](https://dashboard.ngrok.com/domains) (e.g., `your-agent.ngrok-free.app`). This gives you a consistent URL that won't change.

```bash
# Start ngrok with your static domain
ngrok http 3978 --domain=your-agent.ngrok-free.app
```

### 3. Initialize Config (Prod Mode)

```bash
az login
a365 config init
```

During the wizard:
- Set `needDeployment: false` (running locally with ngrok)
- Set `messagingEndpoint: https://your-agent.ngrok-free.app/api/messages`

Your `a365.config.json` should look like:

```json
{
  "needDeployment": false,
  "messagingEndpoint": "https://francisco-unorational-manuela.ngrok-free.dev/api/messages",
  ...
}
```

### 4. Setup Agent Blueprint and Permissions

Since `needDeployment: false` was set during config init, no Azure infrastructure is created. Run the full setup:

```bash
a365 setup all
```

This automatically:
- Creates the agent blueprint (Entra ID app registration)
- Configures MCP server OAuth2 grants
- Configures Messaging Bot API permissions
- Registers the messaging endpoint

This creates `a365.generated.config.json` with your agent's credentials.

> **Note:** If you need more control, you can run steps individually:
> ```bash
> a365 setup blueprint                 # Create blueprint + register endpoint
> a365 setup permissions mcp           # Configure MCP server OAuth2 grants
> a365 setup permissions bot           # Configure Messaging Bot API permissions
> ```

### 5. Add MCP Servers

```bash
# Add MCP servers to your ToolingManifest.json
a365 develop add-mcp-servers mcp_TeamsServer
a365 develop add-mcp-servers mcp_MailTools
a365 develop add-mcp-servers mcp_MeServer
```

### 6. Grant MCP Permissions to Blueprint

After adding MCP servers, grant the permissions to your blueprint:

```bash
# Add delegated permissions for the MCP servers you added
a365 develop add-permissions
```

> **Note:** This grants the OAuth scopes from your `ToolingManifest.json` to your blueprint.

### 7. Configure Environment

Prod mode uses the multi-model template:

```bash
cp .env.template .env
```

**Get blueprint credentials:**

```bash
a365 config display -g
```

**Copy to `.env`:**

| From `a365 config display -g` | `.env` Variable |
|-------------------------------|-----------------|
| `agentBlueprintId` | `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID` |
| `agentBlueprintClientSecret` | `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET` |
| `tenantId` | `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID` |

**Configure Azure OpenAI** (multi-model with failover support):

```env
# Model 1 (PRIMARY)
AZURE_OPENAI_MODEL_1_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_MODEL_1_DEPLOYMENT=gpt-4.1-mini
AZURE_OPENAI_MODEL_1_API_KEY=your-api-key

# Model 2 (FALLBACK - optional)
AZURE_OPENAI_MODEL_2_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_MODEL_2_DEPLOYMENT=gpt-4o-mini
AZURE_OPENAI_MODEL_2_API_KEY=your-api-key

# Common API version for all models
AZURE_OPENAI_API_VERSION=2024-05-01-preview
```

> **Note:** The template already has `AGENT_MODE=prod` and `USE_AGENTIC_AUTH=true` set.

### 8. Customize System Prompt

Edit `agents/system_prompt.md` to define your agent's personality and tool instructions.

### 9. Publish to Microsoft 365 Admin Center

```bash
a365 publish
```

This makes your agent available for admins to approve and users to create instances.

### 10. Approve in Microsoft 365 Admin Center

1. Go to [Microsoft 365 Admin Center - Agents](https://admin.cloud.microsoft/#/agents/all)
2. Find your agent in the list
3. Approve for your tenant

### 11. Configure in Teams Developer Portal

1. Get your blueprint ID:
   ```bash
   a365 config display -g
   # Copy agentBlueprintId
   ```

2. Open: `https://dev.teams.microsoft.com/tools/agent-blueprint/<agentBlueprintId>/configuration`

3. Configure:
   - **Agent Type**: Bot Based
   - **Bot ID**: `<agentBlueprintId>`
   - Click **Save**

### 12. Create Agent Instance in Teams

1. Open **Teams** → **Apps**
2. Search for your agent name
3. Click **Request Instance** or **Add**
4. Wait for admin approval (if required)
5. Once approved, give your agent instance a name

### 13. Run the Agent

```bash
# Make sure ngrok is running with your static domain
ngrok http 3978 --domain=your-agent.ngrok-free.app

# Start the agent
uv run main.py
```

### 14. Test in Teams

1. Search for your agent user in Teams
2. Start a chat
3. Send a test message: `Hello!`
4. Test tool calls: `Send an email to yourself@company.com with subject "Test" and body "Hello from my agent!"`

See [Create agent instances](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/create-instance) and [Testing agents](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/testing?tabs=python) for full documentation.

---

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

**Important:** The agent's personality and instructions are defined in `agents/system_prompt.md`. **You must customize this file** before running your agent to define:

- Agent name and personality
- Available MCP tools and how to use them
- Conversation context handling (Teams chat history)
- Email reply behavior (Reply vs Reply All)
- User lookup instructions
- Any domain-specific instructions

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

To publish the agent to Microsoft 365 admin center. See [Publish agent](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/publish).

```bash
# Verify setup is complete
a365 config display -g  # Should show agentBlueprintId

# Deploy to Azure (if not already deployed)
a365 deploy

# Publish to Microsoft 365 admin center
a365 publish
```

After publishing, configure the agent blueprint in [Teams Developer Portal](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/create-instance#1-configure-agent-in-teams-developer-portal), then create agent instances from [Microsoft 365 admin center](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/create-instance).

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

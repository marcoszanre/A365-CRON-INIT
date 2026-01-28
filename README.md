# Python Agent 365 Sample

A sample Python agent built with the **Microsoft Agent 365 SDK** and **AgentFramework SDK**. This agent demonstrates how to build an enterprise-grade AI agent with:

- 🔐 **Agentic Identity** - Entra-backed identity with its own mailbox
- 📧 **Email & Document Notifications** - Respond to @mentions in Teams, Outlook, Word
- 🔧 **MCP Tooling** - Access Microsoft 365 data via governed MCP servers (Word, Mail)
- 📊 **Observability** - OpenTelemetry-based tracing and monitoring
- 🤖 **Azure OpenAI** - Powered by GPT models

## Prerequisites

- **Python 3.11+**
- **[uv](https://docs.astral.sh/uv/)** - Fast Python package manager
- **[Agent 365 CLI](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/cli)** - For blueprint setup and publishing
- **[ngrok](https://ngrok.com/)** - For local development tunneling
- **Azure OpenAI** - Deployed GPT model (e.g., `gpt-4o`, `gpt-4.1`)
- **Microsoft 365 Tenant** - With Agent 365 Frontier preview access

## Quick Start

### 1. Clone and Install

```bash
git clone https://github.com/marcoszanre/python-a365-sample.git
cd python-a365-sample

# Install dependencies with uv
uv sync
```

### 2. Configure Environment

Copy the template and fill in your values:

```bash
cp .env.template .env
```

Edit `.env` with your configuration:

```env
# Azure OpenAI
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-4o
AZURE_OPENAI_API_VERSION=2024-05-01-preview
AZURE_OPENAI_API_KEY=your-api-key

# Agent 365 Agentic Authentication (from a365 setup)
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=your-blueprint-id
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=your-secret
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=your-tenant-id

# Enable Agentic Auth
USE_AGENTIC_AUTH=true
AUTH_HANDLER_NAME=AGENTIC
```

## Local Development with ngrok

### 1. Start ngrok Tunnel

```bash
ngrok http 3978
```

Copy the HTTPS URL (e.g., `https://abc123.ngrok-free.app`)

### 2. Setup Agent 365 Blueprint

```bash
# Initialize blueprint with your ngrok endpoint
a365 setup blueprint --endpoint-only
# Enter your ngrok URL + /api/messages when prompted

# Setup MCP permissions (for Word and Mail tools)
a365 setup permissions mcp

# Publish to M365 Admin Center
a365 publish
```

### 3. Run the Agent

```bash
uv run python start_with_generic_host.py
```

### 4. Test in Teams

1. Open Microsoft Teams
2. Search for your agent by name in the search bar
3. Start a chat and send a message
4. Send an email to your agent's mailbox

## Project Structure

```
├── agent.py                    # Main agent implementation
├── agent_interface.py          # Agent interface definition
├── host_agent_server.py        # Generic agent host server
├── start_with_generic_host.py  # Entry point
├── local_authentication_options.py  # Auth configuration
├── token_cache.py              # Token caching utilities
├── ToolingManifest.json        # MCP server configuration
├── pyproject.toml              # Python dependencies
├── .env.template               # Environment template
└── docs/                       # Additional documentation
```

## MCP Tooling Configuration

The agent uses MCP servers for Microsoft 365 integration. Configure in `ToolingManifest.json`:

```json
{
  "mcpServers": [
    {
      "mcpServerName": "mcp_WordServer",
      "scope": "McpServers.Word.All"
    },
    {
      "mcpServerName": "mcp_MailTools",
      "scope": "McpServers.Mail.All"
    }
  ]
}
```

## Publishing to Azure

### 1. Setup Azure Resources

```bash
# Run full Azure setup (creates App Service, etc.)
a365 setup azure
```

### 2. Update Endpoint

Update `a365.config.json` with your Azure App Service URL:

```json
{
  "messagingEndpoint": "https://your-app.azurewebsites.net/api/messages",
  "needDeployment": true
}
```

### 3. Deploy

```bash
# Deploy to Azure App Service
a365 deploy
```

### 4. Update Environment Variables

In Azure Portal, add the same environment variables from your `.env` file to your App Service Configuration.

## Features

### Email Notifications

The agent automatically responds to emails sent to its mailbox. The SDK handles:
- Receiving email notifications
- Generating AI-powered responses
- Sending replies via `EmailResponse.create_email_response_activity()`

### Word Document Comments

When @mentioned in Word document comments, the agent:
- Retrieves the document content
- Reads the comment context
- Generates and posts a reply

### Teams Chat

Direct chat with the agent in Teams for:
- Creating Word documents
- Answering questions
- General assistance

## Troubleshooting

### Common Issues

1. **502 Bad Gateway on typing indicator**
   - Usually indicates auth/consent issues
   - Run `a365 cleanup instance` and recreate

2. **MCP tools returning 403**
   - Permission not propagated to instance
   - Run `a365 setup permissions mcp` then recreate instance

3. **Consent Required errors**
   - Instance needs new permissions
   - Delete instance with `a365 cleanup instance` and create new one

### Useful Commands

```bash
# Check blueprint permissions
a365 query-entra blueprint-scopes

# Cleanup agent instance
a365 cleanup instance

# View agent status
a365 status
```

## Documentation

- [Microsoft Agent 365 Developer Docs](https://learn.microsoft.com/en-us/microsoft-agent-365/developer)
- [Notifications Guide](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/notification)
- [MCP Tooling Servers](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview)
- [AgentFramework SDK](https://github.com/microsoft/agent-framework)

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA). For details, visit <https://cla.opensource.microsoft.com>.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).

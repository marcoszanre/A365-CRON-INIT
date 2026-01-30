# A365 Agent Framework (Python)

A modular Python framework for building AI agents integrated with Microsoft 365 using the Agent 365 SDK.

## 📁 Project Structure

```
agent365-agentframework-python/
├── a365_agent/                    # Core framework package
│   ├── __init__.py               # Package exports
│   ├── config.py                 # Configuration management
│   ├── auth.py                   # Authentication & token cache
│   ├── observability.py          # Telemetry & tracing
│   ├── base.py                   # AgentBase abstract class
│   ├── host.py                   # GenericAgentHost server
│   ├── mcp/                      # MCP server integration
│   │   ├── __init__.py
│   │   └── service.py
│   └── notifications/            # Notification handlers
│       ├── __init__.py
│       └── handlers.py
├── agents/                        # Agent implementations
│   ├── __init__.py
│   └── contoso_agent.py          # Contoso Assistant agent
├── main.py                        # Entry point
├── .env                           # Environment configuration
└── pyproject.toml                # Python project config
```

## 🚀 Quick Start

### 1. Install Dependencies

```bash
uv pip install -e .
```

### 2. Configure Environment

Copy `.env.example` to `.env` and configure:

```env
# Azure OpenAI
AZURE_OPENAI_ENDPOINT=https://your-endpoint.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-4
AZURE_OPENAI_API_KEY=your-key
AZURE_OPENAI_API_VERSION=2024-05-01-preview

# Agent 365 Authentication
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=your-blueprint-id
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=your-secret
CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=your-tenant-id

# Auth Handler (set to AGENTIC for production)
AUTH_HANDLER_NAME=AGENTIC
```

### 3. Run the Agent

```bash
python main.py
# Or with uv:
uv run main.py
```

## 📦 Modules

### `a365_agent.config`
Centralized configuration from environment variables.

```python
from a365_agent.config import get_settings

settings = get_settings()
print(settings.azure_openai.endpoint)
print(settings.agent_auth.client_id)
```

### `a365_agent.auth`
Token caching and authentication utilities.

```python
from a365_agent.auth import cache_agentic_token, get_cached_agentic_token

cache_agentic_token(tenant_id, agent_id, token)
token = get_cached_agentic_token(tenant_id, agent_id)
```

### `a365_agent.observability`
Telemetry and tracing with Agent 365 SDK.

```python
from a365_agent.observability import configure_observability, ObservabilityContext

configure_observability()

with ObservabilityContext(tenant_id, agent_id, correlation_id):
    # Operations with telemetry
    pass
```

### `a365_agent.mcp`
MCP server integration for M365 tools.

```python
from a365_agent.mcp import MCPService

mcp = MCPService()
agent = await mcp.initialize_with_agentic_auth(
    chat_client, instructions, auth, handler_name, context
)
```

### `a365_agent.base`
Base class for agents.

```python
from a365_agent.base import AgentBase

class MyAgent(AgentBase):
    async def initialize(self) -> None:
        pass
    
    async def process_user_message(self, message, auth, handler, context) -> str:
        return "Hello!"
    
    async def cleanup(self) -> None:
        pass
```

### `a365_agent.host`
Server hosting and notification routing.

```python
from a365_agent import create_and_run_host
from agents import ContosoAgent

create_and_run_host(ContosoAgent)
```

## 🔔 Notification Types

| Type | Handler Method | Channel |
|------|---------------|---------|
| Email | `handle_email_notification()` | `agents/email` |
| Word | `handle_word_notification()` | `agents/word` |
| Excel | `handle_excel_notification()` | `agents/excel` |
| PowerPoint | `handle_powerpoint_notification()` | `agents/powerpoint` |
| Lifecycle | `handle_lifecycle_notification()` | `agents/onboarding` |

## ⚠️ Platform Limitations (Agentic Auth)

Agentic applications have these limitations:

| Feature | Status | Reason |
|---------|--------|--------|
| App-only tokens | ❌ Blocked | AADSTS82001 |
| Proactive messaging | ❌ Blocked | Requires app-only token |
| Startup MCP init | ❌ Blocked | Requires user token |

**Solution:** All processing must complete within the HTTP request lifecycle using generous timeouts (90-120s).

## 📊 Health Check

```bash
curl http://localhost:3978/api/health
```

Response:
```json
{
  "status": "ok",
  "agent_type": "ContosoAgent",
  "agent_initialized": true,
  "mcp_ready": true
}
```

## 🧪 Creating a Custom Agent

```python
# agents/my_custom_agent.py
from a365_agent.base import AgentBase

class MyCustomAgent(AgentBase):
    async def initialize(self) -> None:
        # Setup resources
        pass
    
    async def process_user_message(self, message, auth, handler, context) -> str:
        # Process and respond
        return f"You said: {message}"
    
    async def cleanup(self) -> None:
        # Cleanup resources
        pass

# main.py
from a365_agent import create_and_run_host
from agents.my_custom_agent import MyCustomAgent

create_and_run_host(MyCustomAgent)
```

## 📝 License

Copyright (c) Microsoft. All rights reserved.

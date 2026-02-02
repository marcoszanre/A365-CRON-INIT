#!/usr/bin/env python3
# Copyright (c) Microsoft. All rights reserved.

"""
Test Directory Search via Blueprint Auth

This script tests directory search scenarios using Agent User Impersonation:
1. Uses delegated token (BEARER_TOKEN from a365 develop get-token) if available
2. Otherwise uses Agent User Impersonation flow (3-step: T1 -> T2 -> user_fic)

Uses mcp_MeServer to search for users in the directory.

Usage:
    # With delegated token:
    a365 develop get-token
    uv run devTools/test_directory_search.py
    
    # Without token (Agent User Impersonation):
    $env:BEARER_TOKEN=""; uv run devTools/test_directory_search.py
"""

import asyncio
import base64
import json
import logging
import os
import sys
from pathlib import Path

import aiohttp
from dotenv import load_dotenv

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from a365_agent.config import get_settings
from a365_agent.mcp import MCPService
from agent_framework.azure import AzureOpenAIChatClient

# Configure logging to stdout (not stderr, which shows as red in PowerShell)
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    stream=sys.stdout
)

# Enable DEBUG for agent_framework to see MCP tool call parameters
af_log_level = os.getenv("AGENT_FRAMEWORK_LOG_LEVEL", "DEBUG").upper()
logging.getLogger("agent_framework").setLevel(getattr(logging, af_log_level, logging.INFO))

logger = logging.getLogger(__name__)

# Load .env
load_dotenv(Path(__file__).parent.parent / ".env")


# =============================================================================
# MOCK AUTH CLASSES FOR PROACTIVE SCENARIOS
# =============================================================================

class MockAuthorization:
    """Mock Authorization for proactive scenarios."""
    
    def __init__(self, bearer_token: str):
        self._token = bearer_token
    
    async def get_token_async(self, *args, **kwargs) -> str:
        return self._token


class MockTurnContext:
    """Mock TurnContext for proactive scenarios."""
    
    def __init__(self, user_id: str = "proactive"):
        self.activity = MockActivity(user_id)


class MockActivity:
    def __init__(self, user_id: str):
        self.from_property = MockFrom(user_id)
        self.conversation = MockConversation()


class MockFrom:
    def __init__(self, user_id: str):
        self.id = user_id


class MockConversation:
    def __init__(self):
        self.id = "proactive-directory-conversation"


# =============================================================================
# CONFIGURATION
# =============================================================================

class Config:
    """Configuration from environment."""
    
    # Blueprint credentials
    blueprint_client_id = os.getenv("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID", "")
    blueprint_client_secret = os.getenv("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET", "")
    tenant_id = os.getenv("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID", "")
    
    # Agent Identity
    agent_identity_client_id = os.getenv("AGENT_IDENTITY_CLIENT_ID", "")
    agent_user_upn = os.getenv("AGENT_USER_UPN", "")
    agent_user_object_id = os.getenv("AGENT_USER_OBJECT_ID", "")
    
    # Search target
    search_name = os.getenv("DIRECTORY_SEARCH_NAME", "Lisa Taylor")
    
    # MCP Platform
    mcp_audience = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"
    
    @classmethod
    def validate(cls) -> list[str]:
        """Return list of missing config values."""
        missing = []
        if not cls.blueprint_client_id:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID")
        if not cls.blueprint_client_secret:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET")
        if not cls.tenant_id:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID")
        if not cls.agent_identity_client_id:
            missing.append("AGENT_IDENTITY_CLIENT_ID")
        if not cls.agent_user_upn:
            missing.append("AGENT_USER_UPN")
        if not cls.agent_user_object_id:
            missing.append("AGENT_USER_OBJECT_ID")
        return missing


def decode_token(token: str, label: str) -> dict:
    """Decode and log JWT token claims."""
    try:
        parts = token.split(".")
        payload_b64 = parts[1]
        padding = 4 - (len(payload_b64) % 4)
        if padding != 4:
            payload_b64 += "=" * padding
        payload = json.loads(base64.urlsafe_b64decode(payload_b64))
        
        logger.info(f"   {label} claims:")
        for key in ["aud", "sub", "oid", "upn", "idtyp", "scp", "roles"]:
            if key in payload:
                logger.info(f"      {key}: {payload.get(key)}")
        
        return payload
    except Exception as e:
        logger.warning(f"   Could not decode {label}: {e}")
        return {}


# =============================================================================
# AGENT USER IMPERSONATION FLOW (3-step)
# =============================================================================

async def get_t1(session: aiohttp.ClientSession) -> str:
    """Step 1: Blueprint requests exchange token T1."""
    logger.info("Step 1: Acquiring T1 (Blueprint -> Agent Identity exchange token)...")
    
    token_url = f"https://login.microsoftonline.com/{Config.tenant_id}/oauth2/v2.0/token"
    
    data = {
        "client_id": Config.blueprint_client_id,
        "scope": "api://AzureADTokenExchange/.default",
        "grant_type": "client_credentials",
        "client_secret": Config.blueprint_client_secret,
        "fmi_path": Config.agent_identity_client_id,
    }
    
    logger.info(f"   Blueprint:       {Config.blueprint_client_id}")
    logger.info(f"   Agent Identity:  {Config.agent_identity_client_id}")
    
    async with session.post(token_url, data=data) as response:
        result = await response.json()
        
        if response.status != 200:
            raise Exception(f"T1 failed: {result.get('error_description', result)}")
        
        token = result["access_token"]
        logger.info(f"   SUCCESS - T1 length: {len(token)} chars")
        decode_token(token, "T1")
        return token


async def get_t2(session: aiohttp.ClientSession, t1: str) -> str:
    """Step 2: Agent Identity requests exchange token T2 for Agent User impersonation."""
    logger.info("Step 2: Acquiring T2 (Agent Identity -> Agent User exchange token)...")
    
    token_url = f"https://login.microsoftonline.com/{Config.tenant_id}/oauth2/v2.0/token"
    
    data = {
        "client_id": Config.agent_identity_client_id,
        "scope": "api://AzureADTokenExchange/.default",
        "grant_type": "client_credentials",
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "client_assertion": t1,
    }
    
    logger.info(f"   Agent Identity: {Config.agent_identity_client_id}")
    
    async with session.post(token_url, data=data) as response:
        result = await response.json()
        
        if response.status != 200:
            raise Exception(f"T2 failed: {result.get('error_description', result)}")
        
        token = result["access_token"]
        logger.info(f"   SUCCESS - T2 length: {len(token)} chars")
        decode_token(token, "T2")
        return token


async def get_mcp_token_as_agent_user(session: aiohttp.ClientSession, t1: str, t2: str) -> str:
    """Step 3: Agent Identity requests MCP token via user_fic grant for Agent User."""
    logger.info("Step 3: Acquiring MCP Token (user_fic grant type)...")
    
    token_url = f"https://login.microsoftonline.com/{Config.tenant_id}/oauth2/v2.0/token"
    
    data = {
        "client_id": Config.agent_identity_client_id,
        "scope": f"{Config.mcp_audience}/.default",
        "grant_type": "user_fic",
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "client_assertion": t1,
        "user_id": Config.agent_user_object_id,
        "user_federated_identity_credential": t2,
    }
    
    logger.info(f"   Agent User Object ID: {Config.agent_user_object_id}")
    logger.info(f"   MCP Audience:         {Config.mcp_audience}")
    
    async with session.post(token_url, data=data) as response:
        result = await response.json()
        
        if response.status != 200:
            error_desc = result.get('error_description', result.get('error', str(result)))
            logger.error(f"   FAILED: {error_desc}")
            raise Exception(f"MCP Token failed: {error_desc}")
        
        token = result["access_token"]
        logger.info(f"   SUCCESS - MCP Token length: {len(token)} chars")
        decode_token(token, "MCP Token")
        
        return token


# =============================================================================
# AGENT INSTRUCTIONS
# =============================================================================

AGENT_INSTRUCTIONS = """You are a proactive assistant for Contoso. 
You have access to Microsoft Graph user/directory APIs via MCP Me Server tools.

When asked to search for a user or get user details:
1. Use the listUsers or search tools to find users by name
2. Retrieve detailed information about the user (email, job title, department, manager, etc.)
3. Present the information in a clear, formatted way

Be thorough and retrieve as much relevant information as available.
"""


# =============================================================================
# MAIN
# =============================================================================

async def main():
    """Main entry point."""
    print()
    print("=" * 70)
    print("  Directory Search Test (mcp_MeServer)")
    print("=" * 70)
    print()
    
    # Validate config
    missing = Config.validate()
    if missing:
        logger.error("Missing required .env values:")
        for m in missing:
            logger.error(f"   - {m}")
        return 1
    
    # Show config
    logger.info("=== Configuration ===")
    logger.info(f"Blueprint:       {Config.blueprint_client_id}")
    logger.info(f"Agent Identity:  {Config.agent_identity_client_id}")
    logger.info(f"Agent User UPN:  {Config.agent_user_upn}")
    logger.info(f"Search Name:     {Config.search_name}")
    print()
    
    # Check for BEARER_TOKEN first
    bearer_token = os.getenv("BEARER_TOKEN", "")
    
    if bearer_token:
        logger.info("Using BEARER_TOKEN from .env")
        logger.info("(This is a delegated token)")
        mcp_token = bearer_token
    else:
        logger.info("No BEARER_TOKEN found - trying Agent User Impersonation flow...")
        logger.info("(Search will be performed as the agentic user identity)")
        
        async with aiohttp.ClientSession() as session:
            try:
                # Step 1: Get T1
                t1 = await get_t1(session)
                print()
                
                # Step 2: Get T2
                t2 = await get_t2(session, t1)
                print()
                
                # Step 3: Get MCP token as Agent User
                mcp_token = await get_mcp_token_as_agent_user(session, t1, t2)
                print()
                
            except Exception as e:
                logger.error(f"Agent User Impersonation failed: {e}")
                logger.info("")
                logger.info("To use directory search, you have these options:")
                logger.info("1. Run 'a365 develop get-token' for a delegated token")
                logger.info("2. Use Agent User Impersonation flow")
                return 1
    
    logger.info(f"Token length: {len(mcp_token)} chars")
    print()
    
    try:
        # Initialize MCP with the token
        logger.info("Initializing MCP servers...")
        
        settings = get_settings()
        if settings.model_pool is None:
            logger.error("model_pool is not configured in settings")
            return 1
        model = settings.model_pool.get_next_model()
        
        chat_client = AzureOpenAIChatClient(
            endpoint=model.endpoint,
            api_key=model.api_key,
            deployment_name=model.deployment,
            api_version=model.api_version,
        )
        
        # Create mock auth objects
        mock_auth = MockAuthorization(mcp_token)
        mock_context = MockTurnContext(Config.agent_user_upn)
        
        mcp_service = MCPService()
        agent = await mcp_service.initialize_with_bearer_token(
            chat_client=chat_client,
            agent_instructions=AGENT_INSTRUCTIONS,
            bearer_token=mcp_token,
            auth=mock_auth,
            auth_handler_name="PROACTIVE-DIRECTORY",
            turn_context=mock_context,
        )
        
        logger.info("MCP servers initialized!")
        print()
        
        # Search for user
        search_name = Config.search_name
        logger.info(f"Searching for user: {search_name}...")
        
        prompt = f"""Search for a user named "{search_name}" in the organization directory.

Please retrieve and display their:
- Full name
- Email address
- Job title
- Department
- Office location
- Phone number (if available)
- Manager (if available)
- Any other relevant profile information

Format the results clearly."""
        
        logger.info(f"   Prompt: {prompt[:80]}...")
        
        result = await agent.run(prompt)
        
        # Extract text
        if hasattr(result, 'content'):
            response_text = str(result.content)
        elif hasattr(result, 'text'):
            response_text = str(result.text)
        else:
            response_text = str(result)
        
        print()
        logger.info("=" * 50)
        logger.info("AGENT RESPONSE:")
        print()
        print(response_text)
        print()
        logger.info("=" * 50)
        
        return 0
        
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    import warnings
    warnings.filterwarnings("ignore", category=RuntimeWarning, message=".*cancel scope.*")
    exit_code = asyncio.run(main())
    sys.exit(exit_code)

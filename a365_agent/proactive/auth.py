# Copyright (c) Microsoft. All rights reserved.

"""
Proactive Token Provider

Encapsulates the 3-step Agent User Impersonation flow used by all
proactive / autonomous scenarios:

    T1  – Blueprint → Agent Identity exchange token
    T2  – Agent Identity → Agent User exchange token
    MCP – user_fic grant → resource-scoped token (e.g. MCP platform)

The provider accepts per-agent credentials dynamically so the scheduler
can loop through multiple agents from the database.

Reference:
    https://learn.microsoft.com/en-us/entra/agent-id/identity-platform/agent-user-oauth-flow
"""

import base64
import json
import logging
import os
from dataclasses import dataclass
from typing import Optional

import aiohttp

from a365_agent.config import get_settings

logger = logging.getLogger(__name__)

# Default MCP platform audience (Agent 365 MCP)
_DEFAULT_MCP_AUDIENCE = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"


def _decode_jwt_payload(token: str) -> dict:
    """Decode the payload section of a JWT for logging (best-effort)."""
    try:
        payload_b64 = token.split(".")[1]
        padding = 4 - (len(payload_b64) % 4)
        if padding != 4:
            payload_b64 += "=" * padding
        return json.loads(base64.urlsafe_b64decode(payload_b64))
    except Exception:
        return {}


@dataclass
class AgentCredentials:
    """
    Per-agent credentials for the Agent User Impersonation flow.

    Built dynamically from the PostgreSQL agent_registry + shared
    Blueprint credentials from the environment.
    """
    # Agent-specific (from DB agent_registry)
    agent_user_id: str              # UPN
    agent_identity_client_id: str   # Service principal client ID
    agent_user_object_id: str       # Entra object ID of the agentic user

    # Shared Blueprint credentials (from env — same for all agents in a tenant)
    blueprint_client_id: str = ""
    blueprint_client_secret: str = ""
    tenant_id: str = ""
    mcp_audience: str = _DEFAULT_MCP_AUDIENCE

    @classmethod
    def from_agent_row(cls, agent_row: dict) -> "AgentCredentials":
        """
        Build credentials from a PostgreSQL agent_registry row,
        combined with shared Blueprint env vars.
        """
        return cls(
            agent_user_id=agent_row.get("agent_user_id", ""),
            agent_identity_client_id=agent_row.get("agent_identity_client_id", ""),
            agent_user_object_id=agent_row.get("agent_user_object_id", ""),
            blueprint_client_id=os.getenv(
                "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID", ""
            ),
            blueprint_client_secret=os.getenv(
                "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET", ""
            ),
            tenant_id=os.getenv(
                "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID", ""
            ),
            mcp_audience=os.getenv("MCP_AUDIENCE", _DEFAULT_MCP_AUDIENCE),
        )

    def validate(self) -> list[str]:
        """Return list of missing fields (empty = all good)."""
        missing: list[str] = []
        if not self.blueprint_client_id:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID (env)")
        if not self.blueprint_client_secret:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET (env)")
        if not self.tenant_id:
            missing.append("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID (env)")
        if not self.agent_identity_client_id:
            missing.append("agent_identity_client_id (DB)")
        if not self.agent_user_object_id:
            missing.append("agent_user_object_id (DB)")
        return missing


class ProactiveTokenProvider:
    """
    Acquires tokens for proactive (headless) agent scenarios using the
    Agent User Impersonation 3-step flow (T1 → T2 → user_fic).

    Accepts ``AgentCredentials`` so the same provider can serve
    multiple agents in a single scheduler loop.
    """

    async def acquire_mcp_token(self, creds: AgentCredentials) -> str:
        """
        Acquire an MCP-scoped token for the given agent credentials.

        Raises:
            RuntimeError: if credentials are incomplete or token acquisition fails.
        """
        missing = creds.validate()
        if missing:
            raise RuntimeError(
                f"Cannot acquire token for {creds.agent_user_id} – "
                f"missing: {', '.join(missing)}"
            )

        logger.info(f"🔑 Acquiring MCP token for {creds.agent_user_id}...")
        return await self._agent_user_impersonation_flow(creds)

    # ------------------------------------------------------------------
    # 3-step Agent User Impersonation flow
    # ------------------------------------------------------------------

    async def _agent_user_impersonation_flow(self, creds: AgentCredentials) -> str:
        """Execute the T1 → T2 → user_fic flow and return the MCP token."""
        async with aiohttp.ClientSession() as session:
            t1 = await self._get_t1(session, creds)
            t2 = await self._get_t2(session, creds, t1)
            mcp_token = await self._get_mcp_token(session, creds, t1, t2)
            return mcp_token

    async def _get_t1(self, session: aiohttp.ClientSession, creds: AgentCredentials) -> str:
        """Step 1: Blueprint → Agent Identity exchange token (T1)."""
        logger.info("   Step 1/3: Acquiring T1 (Blueprint → Agent Identity)...")
        token_url = f"https://login.microsoftonline.com/{creds.tenant_id}/oauth2/v2.0/token"

        data = {
            "client_id": creds.blueprint_client_id,
            "scope": "api://AzureADTokenExchange/.default",
            "grant_type": "client_credentials",
            "client_secret": creds.blueprint_client_secret,
            "fmi_path": creds.agent_identity_client_id,
        }

        async with session.post(token_url, data=data) as resp:
            result = await resp.json()
            if resp.status != 200:
                raise RuntimeError(f"T1 failed: {result.get('error_description', result)}")
            token = result["access_token"]
            logger.info(f"   T1 acquired ({len(token)} chars)")
            return token

    async def _get_t2(self, session: aiohttp.ClientSession, creds: AgentCredentials, t1: str) -> str:
        """Step 2: Agent Identity → Agent User exchange token (T2)."""
        logger.info("   Step 2/3: Acquiring T2 (Agent Identity → Agent User)...")
        token_url = f"https://login.microsoftonline.com/{creds.tenant_id}/oauth2/v2.0/token"

        data = {
            "client_id": creds.agent_identity_client_id,
            "scope": "api://AzureADTokenExchange/.default",
            "grant_type": "client_credentials",
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": t1,
        }

        async with session.post(token_url, data=data) as resp:
            result = await resp.json()
            if resp.status != 200:
                raise RuntimeError(f"T2 failed: {result.get('error_description', result)}")
            token = result["access_token"]
            logger.info(f"   T2 acquired ({len(token)} chars)")
            return token

    async def _get_mcp_token(
        self, session: aiohttp.ClientSession, creds: AgentCredentials, t1: str, t2: str
    ) -> str:
        """Step 3: user_fic grant → MCP-scoped token."""
        logger.info("   Step 3/3: Acquiring MCP token (user_fic grant)...")
        token_url = f"https://login.microsoftonline.com/{creds.tenant_id}/oauth2/v2.0/token"

        data = {
            "client_id": creds.agent_identity_client_id,
            "scope": f"{creds.mcp_audience}/.default",
            "grant_type": "user_fic",
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": t1,
            "user_id": creds.agent_user_object_id,
            "user_federated_identity_credential": t2,
        }

        async with session.post(token_url, data=data) as resp:
            result = await resp.json()
            if resp.status != 200:
                desc = result.get("error_description", result.get("error", str(result)))
                raise RuntimeError(f"MCP token failed: {desc}")
            token = result["access_token"]
            logger.info(f"   MCP token acquired ({len(token)} chars)")
            return token

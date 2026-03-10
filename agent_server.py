# Copyright (c) Microsoft. All rights reserved.
# Licensed under the MIT License.

"""
Server-Side Workplace Agent — OBO Auth + Agent 365 MCP Tools

A FastAPI web app demonstrating a reference sample for delegated auth:
  1. User logs in via Microsoft Entra ID (browser redirect)
  2. Server acquires tokens via Authorization Code flow
  3. Server uses OBO (On-Behalf-Of) to get Agent 365 tokens
  4. Agent uses MCPStreamableHTTPTool to call Work IQ MCP servers
  5. MSAL token cache handles silent refresh — no re-prompts

For the demo, this also supports a simpler "interactive" mode where
tokens come from MSAL PublicClientApplication (no client secret needed).

Prerequisites:
    pip install -r requirements.txt
    # Set .env with your Foundry + App Registration details

Usage (interactive mode — for testing):
    python agent_server.py --mode interactive

Usage (OBO mode — reference sample):
    python agent_server.py --mode obo
    # Then open http://localhost:8080/login
"""

import asyncio
import argparse
import os
import sys
import time
from collections.abc import Awaitable, Callable

import httpx
import msal
from dotenv import load_dotenv

load_dotenv()

# --- Config ---
A365_RESOURCE_ID = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"
APP_CLIENT_ID = os.environ.get("A365_CLIENT_ID", "")
APP_CLIENT_SECRET = os.environ.get("A365_CLIENT_SECRET", "")  # Only needed for OBO
TENANT_ID = os.environ.get("A365_TENANT_ID", "")

# Scopes for Agent 365 MCP servers (without resource prefix — MSAL adds it)
A365_SCOPES = [
    f"{A365_RESOURCE_ID}/McpServers.Teams.All",
    f"{A365_RESOURCE_ID}/McpServers.Mail.All",
    f"{A365_RESOURCE_ID}/McpServers.Calendar.All",
    f"{A365_RESOURCE_ID}/McpServers.Me.All",
    f"{A365_RESOURCE_ID}/McpServers.Planner.All",
    f"{A365_RESOURCE_ID}/McpServers.Word.All",
    f"{A365_RESOURCE_ID}/McpServers.Knowledge.All",
    f"{A365_RESOURCE_ID}/McpServers.OneDriveSharepoint.All",
]

# MCP Server endpoints
MCP_SERVERS = {
    "teams": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer",
    "mail": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools",
    "calendar": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_CalendarTools",
    "planner": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_PlannerServer",
}

AGENT_NAME = "WorkplaceHelper"
AGENT_INSTRUCTIONS = """\
You are a workplace assistant that helps users with their Microsoft 365 data.
You have access to Teams, Mail, Calendar, and Planner tools.

When answering questions:
- Be concise and direct
- Include relevant dates, names, and specifics
- If the user asks about meetings, include key discussion points
- If the user asks about emails, summarize the important content
- Always clarify if you need more context to answer accurately

Available tool categories:
- Teams: list chats, channels, messages, search Teams messages, post messages
- Mail: search emails, read, reply, send, manage attachments
- Calendar: list events, create/update/delete events, find meeting times
- Planner: list plans, create/update tasks
"""


# ---------- Token Management ----------

class TokenManager:
    """Manages MSAL token acquisition and caching for Agent 365 scopes."""

    def __init__(self, mode: str = "interactive"):
        self.mode = mode
        self._token: str | None = None
        self._expires_at: float = 0

        if mode == "interactive":
            self.app = msal.PublicClientApplication(
                APP_CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            )
        elif mode == "obo":
            if not APP_CLIENT_SECRET:
                print("ERROR: OBO mode requires A365_CLIENT_SECRET in .env")
                sys.exit(1)
            self.app = msal.ConfidentialClientApplication(
                APP_CLIENT_ID,
                client_credential=APP_CLIENT_SECRET,
                authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            )

    def get_token(self) -> str:
        """Get a valid access token, refreshing if needed."""
        # Check if current token is still valid (with 5 min buffer)
        if self._token and time.time() < self._expires_at - 300:
            return self._token

        if self.mode == "interactive":
            return self._acquire_interactive()
        elif self.mode == "obo":
            return self._acquire_obo()
        raise ValueError(f"Unknown mode: {self.mode}")

    def _acquire_interactive(self) -> str:
        """Acquire token via MSAL interactive flow (with silent refresh)."""
        # Try silent first (uses MSAL's built-in token cache)
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(A365_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._expires_at = time.time() + result.get("expires_in", 3600)
                print("  🔄 Token refreshed silently")
                return self._token

        # Fall back to interactive
        print("  🔐 Launching browser for sign-in...")
        result = self.app.acquire_token_interactive(scopes=A365_SCOPES)
        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "unknown"))
            raise RuntimeError(f"Token acquisition failed: {error}")

        self._token = result["access_token"]
        self._expires_at = time.time() + result.get("expires_in", 3600)
        print(f"  ✅ Signed in (token expires in {result.get('expires_in', '?')}s)")
        return self._token

    def _acquire_obo(self) -> str:
        """Acquire token via OBO flow (requires user assertion from prior login)."""
        # In a real web app, the user_assertion comes from the incoming request's
        # Authorization header (the user's ID token or access token from the frontend)
        # For this demo, we first do an interactive login to get the user assertion
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(A365_SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._token = result["access_token"]
                self._expires_at = time.time() + result.get("expires_in", 3600)
                print("  🔄 OBO token refreshed silently")
                return self._token

        raise RuntimeError("OBO: No cached user session. User must log in first.")

    def set_user_assertion(self, user_token: str):
        """Set the user's token for OBO exchange (called after user login)."""
        result = self.app.acquire_token_on_behalf_of(
            user_assertion=user_token,
            scopes=A365_SCOPES,
        )
        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "unknown"))
            raise RuntimeError(f"OBO exchange failed: {error}")
        self._token = result["access_token"]
        self._expires_at = time.time() + result.get("expires_in", 3600)
        print(f"  ✅ OBO token acquired (expires in {result.get('expires_in', '?')}s)")
        return self._token


# ---------- Agent Setup ----------

async def create_mcp_tools(token_manager: TokenManager):
    """Create MCPStreamableHTTPTool instances for each Work IQ server."""
    from agent_framework import MCPStreamableHTTPTool

    token = token_manager.get_token()
    tools = []

    for name, url in MCP_SERVERS.items():
        http_client = httpx.AsyncClient(
            headers={"Authorization": f"Bearer {token}"},
            timeout=httpx.Timeout(60.0),
        )
        tool = MCPStreamableHTTPTool(
            name=f"WorkIQ-{name.capitalize()}",
            url=url,
            http_client=http_client,
            description=f"{name.capitalize()} tools via Agent 365 MCP",
            approval_mode="never_require",
            load_prompts=False,
        )
        tools.append((tool, http_client))

    return tools


async def run_repl(token_manager: TokenManager):
    """Run the multi-turn REPL agent."""
    from agent_framework import Agent, FunctionInvocationContext
    from agent_framework.azure import AzureOpenAIResponsesClient

    # --- Function middleware ---
    async def log_tool_calls(
        context: FunctionInvocationContext,
        call_next: Callable[[], Awaitable[None]],
    ) -> None:
        name = context.function.name
        args = context.arguments if hasattr(context, "arguments") else {}
        print(f"\n  🔧 Calling tool: {name}")
        if args:
            for k, v in args.items():
                val = str(v)[:120] + ("..." if len(str(v)) > 120 else "")
                print(f"     {k}: {val}")
        start = time.time()
        await call_next()
        elapsed = time.time() - start
        print(f"  ✅ {name} returned ({elapsed:.1f}s)\n")

    # --- LLM Client ---
    project_endpoint = os.environ["AZURE_AI_PROJECT_ENDPOINT"]
    api_key = os.environ["AZURE_AI_API_KEY"]
    deployment = os.environ["AZURE_AI_MODEL_DEPLOYMENT_NAME"]
    base = project_endpoint.split("/api/projects/")[0]

    client = AzureOpenAIResponsesClient(
        base_url=f"{base}/openai/v1/",
        api_key=api_key,
        deployment_name=deployment,
        middleware=[log_tool_calls],
    )

    # --- Connect to MCP servers ---
    print(f"\nStarting {AGENT_NAME}...")
    print("Connecting to Agent 365 MCP servers...\n")

    tool_pairs = await create_mcp_tools(token_manager)
    tool_objects = [t for t, _ in tool_pairs]
    http_clients = [c for _, c in tool_pairs]

    # Connect all tools
    for tool in tool_objects:
        await tool.connect()
        print(f"  ✅ {tool.name}: {len(tool.functions)} tools")

    try:
        agent = Agent(
            client=client,
            name=AGENT_NAME,
            instructions=AGENT_INSTRUCTIONS,
            tools=tool_objects,
        )

        session = agent.create_session()

        print()
        print("=" * 60)
        print(f"  {AGENT_NAME} — Server-Side Agent w/ Agent 365 MCP")
        print(f"  Auth: {token_manager.mode} | Token auto-refreshes")
        print("  Type your question, or 'quit' to exit.")
        print("=" * 60)
        print()

        while True:
            try:
                user_input = input("You: ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not user_input:
                continue
            if user_input.lower() in ("quit", "exit", "q"):
                print("Goodbye!")
                break

            # Refresh token if needed before each request
            try:
                new_token = token_manager.get_token()
                # Update all HTTP clients with fresh token
                for hc in http_clients:
                    hc.headers["authorization"] = f"Bearer {new_token}"
            except Exception as e:
                print(f"  ⚠️ Token refresh failed: {e}")

            try:
                result = await agent.run(user_input, session=session)
                print(f"\nAgent: {result}\n")
            except Exception as e:
                print(f"\n❌ Error: {e}\n", file=sys.stderr)
    finally:
        # Cleanup
        for tool in tool_objects:
            try:
                await tool.close()
            except Exception:
                pass
        for hc in http_clients:
            await hc.aclose()


# ---------- Main ----------

def main():
    parser = argparse.ArgumentParser(description="Server-Side Workplace Agent")
    parser.add_argument(
        "--mode", choices=["interactive", "obo"], default="interactive",
        help="Auth mode: 'interactive' (MSAL browser login) or 'obo' (server-side OBO)"
    )
    args = parser.parse_args()

    token_manager = TokenManager(mode=args.mode)
    asyncio.run(run_repl(token_manager))


if __name__ == "__main__":
    main()

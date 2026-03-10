"""
Test connecting to Agent 365 Work IQ MCP servers from a custom agent.

Validates that MCPStreamableHTTPTool from agent-framework can consume
remote Agent 365 MCP servers (Planner, Teams, Mail, etc.) with a
delegated user OAuth token.

Proven working:
    - Planner MCP server via MCPStreamableHTTPTool ✅
    - Teams requires McpServers.Teams.All scope (needs App Registration)

Auth options:
    1. az cli token (easiest, but limited to scopes the az cli app has)
    2. MSAL interactive flow with a custom App Registration
    3. Manual Bearer token from Postman

Prerequisites:
    1. pip install agent-framework --pre msal python-dotenv httpx
    2. az login (for option 1)

Usage:
    python test_remote_mcp.py
"""

import asyncio
import os
import sys

from dotenv import load_dotenv

load_dotenv()

# --- Agent 365 MCP Server Endpoints ---
MCP_SERVERS = {
    "planner": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_PlannerServer",
    "teams": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer",
    "mail": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools",
    "calendar": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_CalendarTools",
}

# Agent 365 API resource ID
A365_RESOURCE_ID = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"

# Our custom App Registration (created by setup_a365.py)
OUR_APP_CLIENT_ID = os.environ.get("A365_CLIENT_ID", "")
TENANT_ID = os.environ.get("A365_TENANT_ID", "")

# All Work IQ scopes we want
MSAL_SCOPES = [
    f"{A365_RESOURCE_ID}/McpServers.Teams.All",
    f"{A365_RESOURCE_ID}/McpServers.Mail.All",
    f"{A365_RESOURCE_ID}/McpServers.Calendar.All",
    f"{A365_RESOURCE_ID}/McpServers.Me.All",
    f"{A365_RESOURCE_ID}/McpServers.Planner.All",
    f"{A365_RESOURCE_ID}/McpServers.Word.All",
    f"{A365_RESOURCE_ID}/McpServers.Knowledge.All",
    f"{A365_RESOURCE_ID}/McpServers.OneDriveSharepoint.All",
]


def get_token_msal_interactive() -> str | None:
    """Acquire token via MSAL interactive browser login (one-time per session)."""
    import msal
    app = msal.PublicClientApplication(
        OUR_APP_CLIENT_ID,
        authority="https://login.microsoftonline.com/common",  # multi-tenant: any org account
    )
    # Try silent first (cached tokens from a previous run)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(MSAL_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("  Using cached MSAL token (silent refresh)")
            return result["access_token"]

    print("  Launching browser for interactive sign-in...")
    result = app.acquire_token_interactive(
        scopes=MSAL_SCOPES,
    )
    if "access_token" in result:
        print(f"  Token acquired (expires in {result.get('expires_in', '?')}s)")
        return result["access_token"]
    else:
        print(f"  MSAL error: {result.get('error_description', result.get('error', 'unknown'))}")
        return None


async def get_token_az_cli() -> str | None:
    """Acquire token via az cli (limited scopes)."""
    import subprocess

    try:
        result = subprocess.run(
            ["az", "account", "get-access-token",
             "--resource", f"api://{A365_RESOURCE_ID}",
             "--query", "accessToken", "-o", "tsv"],
            capture_output=True, text=True, timeout=15,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except Exception as e:
        print(f"  az cli failed: {e}")
    return None


async def connect_and_list_tools(server_name: str, url: str, access_token: str):
    """Connect to a remote MCP server and list its tools."""
    import httpx
    from agent_framework import MCPStreamableHTTPTool

    http_client = httpx.AsyncClient(
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=httpx.Timeout(60.0),
    )

    mcp_tool = MCPStreamableHTTPTool(
        name=f"WorkIQ-{server_name}",
        url=url,
        http_client=http_client,
        description=f"{server_name} tools via Agent 365",
        approval_mode="never_require",
        load_prompts=False,
    )

    try:
        async with mcp_tool:
            print(f"  ✅ Connected to {server_name}!")
            print(f"  📋 Discovered {len(mcp_tool.functions)} tools:")
            for fn in mcp_tool.functions:
                desc = fn.description[:65] if fn.description else "(no description)"
                print(f"     - {fn.name}: {desc}")
            return mcp_tool.functions
    except Exception as e:
        error_msg = str(e)
        if "403" in error_msg:
            print(f"  ❌ {server_name}: 403 Forbidden — missing scope (McpServers.{server_name}.All)")
        elif "400" in error_msg:
            print(f"  ❌ {server_name}: 400 Bad Request — {error_msg[:100]}")
        else:
            print(f"  ❌ {server_name}: {type(e).__name__}: {error_msg[:150]}")
        return None
    finally:
        await http_client.aclose()


async def test_tool_call(server_name: str, url: str, access_token: str, tool_name: str, **kwargs):
    """Call a specific tool on the remote MCP server."""
    import httpx
    from agent_framework import MCPStreamableHTTPTool

    http_client = httpx.AsyncClient(
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=httpx.Timeout(60.0),
    )

    mcp_tool = MCPStreamableHTTPTool(
        name=f"WorkIQ-{server_name}",
        url=url,
        http_client=http_client,
        description=f"{server_name} tools via Agent 365",
        approval_mode="never_require",
        load_prompts=False,
    )

    try:
        async with mcp_tool:
            print(f"\n  🔧 Calling {tool_name}({kwargs})...")
            result = await mcp_tool.call_tool(tool_name, **kwargs)
            for item in result:
                text = str(item)
                print(f"  📤 Result: {text[:300]}")
            return result
    except Exception as e:
        print(f"  ❌ Tool call failed: {type(e).__name__}: {e}")
        return None
    finally:
        await http_client.aclose()


async def main():
    print("=" * 60)
    print("  Agent 365 Work IQ MCP — Remote Server Test")
    print("=" * 60)

    # Get token — prefer MSAL interactive (has all scopes), fall back to az cli
    token = os.environ.get("A365_BEARER_TOKEN", "")
    if token:
        print(f"✅ Using A365_BEARER_TOKEN from env ({len(token)} chars)")
    else:
        print("\n🔐 Acquiring token via MSAL interactive login...")
        token = get_token_msal_interactive()

    if not token:
        print("\n🔐 Falling back to az cli token (limited scopes)...")
        token = await get_token_az_cli()

    if not token:
        print("❌ No token available. Run: az login")
        return

    print(f"✅ Got Bearer token ({len(token)} chars)\n")

    # Test each MCP server
    for name, url in MCP_SERVERS.items():
        print(f"\n--- {name.upper()} MCP Server ---")
        print(f"  URL: {url}")
        tools = await connect_and_list_tools(name, url, token)

    # Demo: Call Planner QueryPlans
    print("\n\n--- DEMO: Planner Tool Call ---")
    await test_tool_call("planner", MCP_SERVERS["planner"], token, "QueryPlans")

    print("\n✅ Test complete!")


if __name__ == "__main__":
    asyncio.run(main())

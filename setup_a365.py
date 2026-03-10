"""Setup Agent 365 App Registration permissions and admin consent.

Prerequisite: Agent 365 SP must already exist (run: az ad sp create --id ea9ffc3e-8a23-4a7d-836d-234d7c7565c1)
"""
import os
import subprocess
import json
import sys
import requests

# Known IDs from SP creation output
A365_APP_ID = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"
# These IDs are specific to YOUR tenant. Update them after running the setup.
# A365_SP_ID: Run `az ad sp show --id ea9ffc3e-... --query id -o tsv` to get this.
# OUR_APP_OBJ_ID/CLIENT_ID: From your `az ad app create` output.
A365_SP_ID = os.environ.get("A365_SP_ID", "")
OUR_APP_OBJ_ID = os.environ.get("A365_APP_OBJ_ID", "")
OUR_APP_CLIENT_ID = os.environ.get("A365_CLIENT_ID", "")
GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

# Scope GUIDs from SP output
SCOPES = {
    "McpServers.Teams.All": "5efd4b9c-e459-40d4-a524-35db033b072f",
    "McpServers.Mail.All": "be685e8e-277f-43ec-aff6-087fdca57ca3",
    "McpServers.Calendar.All": "75c3a580-2c8f-4906-adc6-ffa8601d78dc",
    "McpServers.Me.All": "2ce6ce0f-4701-4b11-8087-5031f87ad5b9",
    "McpServers.Word.All": "6f7b3c3c-d822-4164-b9ec-8bf520399d24",
    "McpServers.Planner.All": "127adc5b-6aa1-4b75-924b-87145052e3c2",
    "McpServers.Knowledge.All": "798204a9-2b1d-4109-a2d6-3c641183c48c",
    "McpServers.OneDriveSharepoint.All": "45b74cfc-7a12-4589-8d26-781de38fbfcc",
    "McpServersMetadata.Read.All": "59ccebf0-00a5-4d33-8769-cef7d7acb59d",
}

# Graph basic scopes
GRAPH_SCOPES = {
    "openid": "37f7f235-527c-4136-accd-4a02d197296e",
    "profile": "14dad69e-099b-42c9-810b-d002981feec1",
    "offline_access": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182",
}


def get_graph_token():
    result = subprocess.run(
        ["az", "account", "get-access-token", "--resource", "https://graph.microsoft.com", "--query", "accessToken", "-o", "tsv"],
        capture_output=True, text=True, timeout=15, shell=True,
    )
    return result.stdout.strip()


def main():
    token = get_graph_token()
    if not token:
        print("ERROR: No Graph token. Run: az login")
        return
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Step 1: Add API permissions to our app registration
    print("Step 1: Add API permissions to WorkIQ-MCP-Agent")
    a365_access = [{"id": sid, "type": "Scope"} for sid in SCOPES.values()]
    graph_access = [{"id": sid, "type": "Scope"} for sid in GRAPH_SCOPES.values()]

    resp = requests.patch(
        f"https://graph.microsoft.com/v1.0/applications/{OUR_APP_OBJ_ID}",
        headers=headers,
        json={
            "requiredResourceAccess": [
                {"resourceAppId": A365_APP_ID, "resourceAccess": a365_access},
                {"resourceAppId": GRAPH_APP_ID, "resourceAccess": graph_access},
            ]
        },
    )
    if resp.status_code == 204:
        print(f"  OK: Added {len(a365_access)} Agent 365 scopes + {len(graph_access)} Graph scopes")
    else:
        print(f"  FAILED ({resp.status_code}): {resp.text[:300]}")
        return

    # Step 2: Create SP for our app (needed for consent grants)
    print("Step 2: Ensure our app has a service principal")
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/servicePrincipals",
        params={"$filter": f"appId eq '{OUR_APP_CLIENT_ID}'"},
        headers=headers,
    )
    our_sps = resp.json().get("value", [])
    if our_sps:
        our_sp_id = our_sps[0]["id"]
        print(f"  Exists: {our_sp_id}")
    else:
        resp = requests.post(
            "https://graph.microsoft.com/v1.0/servicePrincipals",
            headers=headers,
            json={"appId": OUR_APP_CLIENT_ID},
        )
        if resp.status_code == 201:
            our_sp_id = resp.json()["id"]
            print(f"  Created: {our_sp_id}")
        else:
            print(f"  FAILED: {resp.text[:200]}")
            return

    # Step 3: Grant admin consent for Agent 365 scopes
    print("Step 3: Grant admin consent for Agent 365")
    scope_string = " ".join(SCOPES.keys())
    resp = requests.post(
        "https://graph.microsoft.com/v1.0/oauth2PermissionGrants",
        headers=headers,
        json={
            "clientId": our_sp_id,
            "consentType": "AllPrincipals",
            "resourceId": A365_SP_ID,
            "scope": scope_string,
        },
    )
    if resp.status_code == 201:
        print(f"  GRANTED: {scope_string}")
    elif resp.status_code == 409:
        print("  Already exists, updating...")
        resp2 = requests.get(
            "https://graph.microsoft.com/v1.0/oauth2PermissionGrants",
            params={"$filter": f"clientId eq '{our_sp_id}' and resourceId eq '{A365_SP_ID}'"},
            headers=headers,
        )
        grants = resp2.json().get("value", [])
        if grants:
            resp3 = requests.patch(
                f"https://graph.microsoft.com/v1.0/oauth2PermissionGrants/{grants[0]['id']}",
                headers=headers,
                json={"scope": scope_string},
            )
            print(f"  Updated: {resp3.status_code}")
    else:
        print(f"  FAILED ({resp.status_code}): {resp.text[:300]}")

    # Step 4: Also grant consent for Graph basic scopes
    print("Step 4: Grant admin consent for Graph scopes")
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/servicePrincipals",
        params={"$filter": f"appId eq '{GRAPH_APP_ID}'"},
        headers=headers,
    )
    graph_sps = resp.json().get("value", [])
    if graph_sps:
        graph_sp_id = graph_sps[0]["id"]
        resp = requests.post(
            "https://graph.microsoft.com/v1.0/oauth2PermissionGrants",
            headers=headers,
            json={
                "clientId": our_sp_id,
                "consentType": "AllPrincipals",
                "resourceId": graph_sp_id,
                "scope": "openid profile offline_access",
            },
        )
        if resp.status_code == 201:
            print("  GRANTED: openid profile offline_access")
        elif resp.status_code == 409:
            print("  Already exists (OK)")
        else:
            print(f"  FAILED ({resp.status_code}): {resp.text[:200]}")

    print("\nSETUP COMPLETE")
    print(f"  App: WorkIQ-MCP-Agent ({OUR_APP_CLIENT_ID})")
    print(f"  Scopes: {scope_string}")
    print(f"\n  Next: python test_remote_mcp.py")


if __name__ == "__main__":
    main()

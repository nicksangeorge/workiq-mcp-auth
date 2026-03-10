# Copyright (c) Microsoft. All rights reserved.
# Licensed under the MIT License.

"""
OBO Web Agent -- FastAPI + Agent 365 MCP + Microsoft Agent Framework

Reference sample for delegated-user authentication and direct MCP connectivity
to Agent 365 Work IQ servers from a custom Python agent. Not a production-ready
web application. See README for production hardening guidance.

Flow:
  1. User visits /login, redirected to Microsoft Entra ID
  2. User signs in with their org account
  3. Server exchanges auth code via MSAL ConfidentialClient
  4. Server uses OBO to get Agent 365-scoped tokens
  5. User chats with the agent via /chat endpoint
  6. Per-session MSAL token cache handles silent refresh

Usage:
    python agent_obo.py
    # Open http://localhost:8080
"""

import asyncio
import json
import logging
import os
import sys
import time
import uuid
from collections.abc import Awaitable, Callable

import httpx
import msal
import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse

load_dotenv()
logger = logging.getLogger(__name__)

# --- Config ---
APP_CLIENT_ID = os.environ.get("A365_CLIENT_ID", "")
APP_CLIENT_SECRET = os.environ.get("A365_CLIENT_SECRET", "")
TENANT_ID = os.environ.get("A365_TENANT_ID", "")
A365_RESOURCE_ID = "ea9ffc3e-8a23-4a7d-836d-234d7c7565c1"
REDIRECT_URI = "http://localhost:8080/auth/callback"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SESSION_MAX_AGE = 3600  # 1 hour

# Scope the frontend requests (our own API)
FRONTEND_SCOPES = [f"api://{APP_CLIENT_ID}/user_impersonation"]

# Scopes for Agent 365 MCP servers (what OBO exchanges for)
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

# MCP endpoints
MCP_SERVERS = {
    "teams": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_TeamsServer",
    "mail": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools",
    "calendar": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_CalendarTools",
    "planner": "https://agent365.svc.cloud.microsoft/agents/servers/mcp_PlannerServer",
}

if not APP_CLIENT_ID or not TENANT_ID:
    print("ERROR: Set A365_CLIENT_ID and A365_TENANT_ID in .env")
    sys.exit(1)

if not APP_CLIENT_SECRET:
    print("ERROR: Set A365_CLIENT_SECRET in .env")
    print("  Run: az ad app credential reset --id <YOUR_APP_OBJECT_ID> --append")
    sys.exit(1)

# Global MSAL app used only for initiating auth code flows (stateless, no token
# cache attached). All token-acquiring operations (auth code redemption, OBO,
# silent refresh) go through per-session MSAL instances with isolated caches.
_login_msal_app = msal.ConfidentialClientApplication(
    APP_CLIENT_ID,
    client_credential=APP_CLIENT_SECRET,
    authority=AUTHORITY,
)

# In-memory session store. Use Redis or a database-backed store in production.
sessions: dict[str, dict] = {}

app = FastAPI(title="WorkIQ OBO Agent")


# ---------- Per-Session MSAL Cache ----------

def _build_msal_app_for_session(session: dict) -> msal.ConfidentialClientApplication:
    """Build a ConfidentialClientApplication with the session's token cache.

    Each session gets its own SerializableTokenCache so tokens from one
    user's session are never visible to another user's session.
    """
    cache = msal.SerializableTokenCache()
    cache_blob = session.get("token_cache")
    if cache_blob:
        cache.deserialize(cache_blob)

    app = msal.ConfidentialClientApplication(
        APP_CLIENT_ID,
        client_credential=APP_CLIENT_SECRET,
        authority=AUTHORITY,
        token_cache=cache,
    )
    return app


def _save_cache_to_session(session: dict, msal_app: msal.ConfidentialClientApplication):
    """Persist the MSAL token cache back to the session if it changed."""
    cache = msal_app.token_cache
    if hasattr(cache, "has_state_changed") and cache.has_state_changed:
        session["token_cache"] = cache.serialize()


def _is_secure_request(request: Request) -> bool:
    """Check if the request came over HTTPS (for Secure cookie flag)."""
    return request.url.scheme == "https"


# ---------- Auth Endpoints ----------

@app.get("/login")
async def login():
    """Start the auth code flow using MSAL helpers."""
    flow = _login_msal_app.initiate_auth_code_flow(
        scopes=FRONTEND_SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    # Store the flow keyed by its state value so /auth/callback can retrieve it.
    state = flow["state"]
    sessions[state] = {"status": "pending", "auth_flow": flow}
    return RedirectResponse(flow["auth_uri"])


@app.get("/auth/callback")
async def auth_callback(request: Request):
    """Complete the auth code flow and establish a user session."""
    state = request.query_params.get("state", "")

    if request.query_params.get("error"):
        logger.warning("Auth error from Entra: %s", request.query_params.get("error_description", ""))
        return HTMLResponse("<h2>Login failed</h2><p>Authentication was denied or cancelled.</p>", status_code=400)

    pending = sessions.pop(state, None)
    if not pending or pending.get("status") != "pending" or "auth_flow" not in pending:
        return HTMLResponse("<h2>Login failed</h2><p>Invalid or expired login state.</p>", status_code=400)

    # Complete the flow using a fresh, cache-isolated MSAL app so the redeemed
    # tokens land in the session-specific cache, not a shared process-wide cache.
    session_data: dict = {
        "status": "pending_obo",
        "token_cache": None,
    }
    session_msal = _build_msal_app_for_session(session_data)

    result = session_msal.acquire_token_by_auth_code_flow(
        auth_code_flow=pending["auth_flow"],
        auth_response=dict(request.query_params),
    )

    if "access_token" not in result:
        logger.error("Token exchange failed: %s", result.get("error_description", result.get("error", "")))
        return HTMLResponse("<h2>Login failed</h2><p>Token exchange failed. Try again.</p>", status_code=400)

    # Persist the cache after auth code redemption so OBO can find the tokens
    _save_cache_to_session(session_data, session_msal)

    # OBO exchange using the same session-scoped MSAL app
    session_id = str(uuid.uuid4())
    id_claims = result.get("id_token_claims", {})
    account_oid = id_claims.get("oid", "")

    session_data.update({
        "status": "authenticated",
        "user": id_claims.get("preferred_username", id_claims.get("name", "User")),
        "account_oid": account_oid,
        "chat_history": [],
        "created_at": time.time(),
    })

    obo_result = session_msal.acquire_token_on_behalf_of(
        user_assertion=result["access_token"],
        scopes=A365_SCOPES,
    )

    if "access_token" not in obo_result:
        logger.error("OBO exchange failed: %s", obo_result.get("error_description", obo_result.get("error", "")))
        return HTMLResponse("<h2>Login failed</h2><p>Could not obtain Work IQ access. Try again.</p>", status_code=400)

    # Save the cache (now contains the OBO tokens) into the session
    _save_cache_to_session(session_data, session_msal)
    sessions[session_id] = session_data

    response = RedirectResponse("/")
    response.set_cookie(
        "session_id",
        session_id,
        httponly=True,
        samesite="lax",
        secure=_is_secure_request(request),
        max_age=SESSION_MAX_AGE,
    )
    return response


@app.get("/logout")
async def logout(request: Request):
    """Clear the user session and redirect to login."""
    session_id = request.cookies.get("session_id", "")
    sessions.pop(session_id, None)
    response = RedirectResponse("/")
    response.delete_cookie("session_id")
    return response


def get_session(request: Request) -> dict | None:
    """Get the current user's session, or None if not authenticated."""
    session_id = request.cookies.get("session_id")
    if not session_id or session_id not in sessions:
        return None
    session = sessions[session_id]
    # Expire old sessions
    if time.time() - session.get("created_at", 0) > SESSION_MAX_AGE:
        sessions.pop(session_id, None)
        return None
    return session


def get_a365_token(session: dict) -> str:
    """Get a valid Agent 365 token from the session's MSAL cache.

    Uses acquire_token_silent with the session's per-user cache and filters
    to the specific account that signed in, not the first cached account.
    """
    session_msal = _build_msal_app_for_session(session)
    target_oid = session.get("account_oid")

    # Find the account that matches this session's signed-in user
    account = None
    for acct in session_msal.get_accounts():
        if acct.get("local_account_id") == target_oid:
            account = acct
            break

    if not account:
        raise RuntimeError("Session expired. Please log in again.")

    result = session_msal.acquire_token_silent(A365_SCOPES, account=account)
    if not result or "access_token" not in result:
        raise RuntimeError("Session expired. Please log in again.")

    _save_cache_to_session(session, session_msal)
    return result["access_token"]


# ---------- Chat Endpoint ----------

@app.post("/chat")
async def chat(request: Request):
    """Stream tool-call events and the final response via SSE."""
    session = get_session(request)
    if not session or session["status"] != "authenticated":
        return JSONResponse({"error": "Not authenticated. Visit /login first."}, status_code=401)

    body = await request.json()
    message = body.get("message", "").strip()
    if not message:
        return JSONResponse({"error": "Empty message"}, status_code=400)

    try:
        token = get_a365_token(session)
    except RuntimeError:
        return JSONResponse({"error": "Session expired. Please log in again."}, status_code=401)

    event_queue: asyncio.Queue = asyncio.Queue()

    async def stream():
        task = asyncio.create_task(
            run_agent_query(message, token, session.get("chat_history", []), event_queue)
        )
        while True:
            event = await event_queue.get()
            if event is None:
                break
            yield f"data: {json.dumps(event)}\n\n"
        try:
            response_text = await task
            session.setdefault("chat_history", []).append({"role": "user", "content": message})
            session["chat_history"].append({"role": "assistant", "content": response_text})
            yield f"data: {json.dumps({'type': 'response', 'text': response_text})}\n\n"
        except Exception:
            logger.exception("Agent query failed")
            yield f"data: {json.dumps({'type': 'error', 'text': 'Something went wrong processing your request.'})}\n\n"

    return StreamingResponse(stream(), media_type="text/event-stream")


async def run_agent_query(message: str, a365_token: str, history: list, event_queue: asyncio.Queue) -> str:
    """Run a single agent query, pushing tool-call events to the queue in real time."""
    from collections.abc import Awaitable, Callable
    from agent_framework import Agent, FunctionInvocationContext, MCPStreamableHTTPTool
    from agent_framework.azure import AzureOpenAIResponsesClient

    project_endpoint = os.environ["AZURE_AI_PROJECT_ENDPOINT"]
    api_key = os.environ["AZURE_AI_API_KEY"]
    deployment = os.environ["AZURE_AI_MODEL_DEPLOYMENT_NAME"]
    base = project_endpoint.split("/api/projects/")[0]

    async def log_tool_calls(
        context: FunctionInvocationContext,
        call_next: Callable[[], Awaitable[None]],
    ) -> None:
        name = context.function.name
        await event_queue.put({"type": "tool_start", "name": name})
        start = time.time()
        await call_next()
        elapsed = round(time.time() - start, 1)
        await event_queue.put({"type": "tool_done", "name": name, "duration_s": elapsed})
        logger.info("Tool call: %s (%.1fs)", name, elapsed)

    client = AzureOpenAIResponsesClient(
        base_url=f"{base}/openai/v1/",
        api_key=api_key,
        deployment_name=deployment,
        middleware=[log_tool_calls],
    )

    instructions = """\
You are a workplace assistant. You have access to Teams, Mail, Calendar, and Planner tools.
Be concise and direct. Include relevant dates, names, and specifics."""

    tools = []
    http_clients = []
    for name, url in MCP_SERVERS.items():
        hc = httpx.AsyncClient(
            headers={"Authorization": f"Bearer {a365_token}"},
            timeout=httpx.Timeout(60.0),
        )
        http_clients.append(hc)
        tool = MCPStreamableHTTPTool(
            name=f"WorkIQ-{name.capitalize()}",
            url=url,
            http_client=hc,
            approval_mode="never_require",
            load_prompts=False,
        )
        tools.append(tool)

    try:
        for tool in tools:
            await tool.connect()

        agent = Agent(
            client=client,
            name="WorkplaceHelper",
            instructions=instructions,
            tools=tools,
        )

        agent_session = agent.create_session()
        result = await agent.run(message, session=agent_session)
        return str(result)
    finally:
        await event_queue.put(None)  # Signal stream end
        for tool in tools:
            try:
                await tool.close()
            except Exception:
                pass
        for hc in http_clients:
            await hc.aclose()


# ---------- UI ----------

@app.get("/")
async def index(request: Request):
    """Simple chat UI."""
    session = get_session(request)
    if not session or session["status"] != "authenticated":
        return HTMLResponse(LOGIN_PAGE)
    return HTMLResponse(CHAT_PAGE.replace("{{USER}}", session.get("user", "User")))


LOGIN_PAGE = """
<!DOCTYPE html>
<html><head><title>WorkIQ Agent</title>
<style>body{font-family:system-ui;display:flex;justify-content:center;align-items:center;height:100vh;margin:0;background:#1a1a2e}
.card{background:#16213e;padding:40px;border-radius:12px;text-align:center;color:#e0e0e0}
a{display:inline-block;margin-top:20px;padding:12px 24px;background:#0078d4;color:white;text-decoration:none;border-radius:6px;font-size:16px}
a:hover{background:#106ebe}</style></head>
<body><div class="card">
<h1>WorkIQ Agent</h1>
<p>Server-side agent with Agent 365 MCP tools</p>
<p style="color:#888">Sign in with your Microsoft account to access Teams, Mail, Calendar & Planner</p>
<a href="/login">Login with Microsoft</a>
</div></body></html>
"""

CHAT_PAGE = """
<!DOCTYPE html>
<html><head><title>WorkIQ Agent</title>
<style>
*{box-sizing:border-box}
body{font-family:system-ui;margin:0;background:#0f0f23;color:#e0e0e0;display:flex;flex-direction:column;height:100vh}
header{background:#16213e;padding:12px 20px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #2a2a4a}
h1{margin:0;font-size:18px}
.user{color:#888;font-size:14px}
#messages{flex:1;overflow-y:auto;padding:20px;display:flex;flex-direction:column;gap:12px}
.msg{max-width:80%;padding:12px 16px;border-radius:12px;line-height:1.5;white-space:pre-wrap}
.msg.user{align-self:flex-end;background:#0078d4;color:white}
.msg.assistant{align-self:flex-start;background:#1e1e3f;border:1px solid #2a2a5a}
.msg.loading{color:#888;font-style:italic}
.msg.toolcall{align-self:flex-start;background:transparent;border:1px solid #333;color:#888;font-size:13px;padding:6px 12px}
footer{padding:12px 20px;background:#16213e;border-top:1px solid #2a2a4a;display:flex;gap:8px}
input{flex:1;padding:12px;border:1px solid #2a2a5a;background:#0f0f23;color:#e0e0e0;border-radius:8px;font-size:15px;outline:none}
input:focus{border-color:#0078d4}
button{padding:12px 24px;background:#0078d4;color:white;border:none;border-radius:8px;cursor:pointer;font-size:15px}
button:hover{background:#106ebe}
button:disabled{background:#333;cursor:not-allowed}
</style></head>
<body>
<header><h1>WorkIQ Agent</h1><div><span class="user">{{USER}}</span> <a href="/logout" style="color:#888;font-size:13px;margin-left:12px">Logout</a></div></header>
<div id="messages"></div>
<footer>
<input id="input" placeholder="Ask about Teams, email, calendar..." autofocus onkeydown="if(event.key==='Enter')send()">
<button onclick="send()">Send</button>
</footer>
<script>
const msgs = document.getElementById('messages');
const inp = document.getElementById('input');

function addMsg(text, role) {
  const d = document.createElement('div');
  d.className = 'msg ' + role;
  d.textContent = text;
  msgs.appendChild(d);
  msgs.scrollTop = msgs.scrollHeight;
  return d;
}

async function send() {
  const text = inp.value.trim();
  if (!text) return;
  inp.value = '';
  addMsg(text, 'user');
  const loading = addMsg('Thinking...', 'loading');
  inp.disabled = true;
  document.querySelector('button').disabled = true;

  try {
    const res = await fetch('/chat', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({message: text})
    });
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    while (true) {
      const {done, value} = await reader.read();
      if (done) break;
      buf += decoder.decode(value, {stream: true});
      const lines = buf.split('\\n');
      buf = lines.pop();
      for (const line of lines) {
        if (!line.startsWith('data: ')) continue;
        const evt = JSON.parse(line.slice(6));
        if (evt.type === 'tool_start') {
          addMsg('Calling ' + evt.name + '...', 'toolcall');
        } else if (evt.type === 'tool_done') {
          const toolMsgs = document.querySelectorAll('.msg.toolcall');
          for (const m of toolMsgs) {
            if (m.textContent === 'Calling ' + evt.name + '...') {
              m.textContent = evt.name + ' (' + evt.duration_s + 's)';
            }
          }
        } else if (evt.type === 'response') {
          loading.remove();
          addMsg(evt.text, 'assistant');
        } else if (evt.type === 'error') {
          loading.remove();
          addMsg('Error: ' + evt.text, 'assistant');
        }
      }
    }
    if (loading.parentNode) loading.remove();
  } catch (e) {
    loading.remove();
    addMsg('Network error: ' + e.message, 'assistant');
  }
  inp.disabled = false;
  document.querySelector('button').disabled = false;
  inp.focus();
}
</script>
</body></html>
"""


if __name__ == "__main__":
    print("=" * 50)
    print("  WorkIQ OBO Agent")
    print("=" * 50)
    print(f"  App: {APP_CLIENT_ID}")
    print(f"  Tenant: {TENANT_ID}")
    print(f"  MCP Servers: {', '.join(MCP_SERVERS.keys())}")
    print()
    print("  Open http://localhost:8080 in your browser")
    print("  Click 'Login with Microsoft' to authenticate")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8080)

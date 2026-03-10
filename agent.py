# Copyright (c) Microsoft. All rights reserved.
# Licensed under the MIT License.

"""
Workplace Helper Agent — Microsoft Agent Framework + WorkIQ MCP

A multi-turn REPL agent that connects to Microsoft Foundry for LLM
inference and WorkIQ MCP for querying Microsoft 365 data (emails,
calendar, meetings, documents, Teams messages).

WorkIQ handles its own authentication via MSAL cached tokens.
On first run, it will pop a browser for sign-in; subsequent runs
reuse cached tokens (~24h before re-prompt).

Prerequisites:
    1. pip install agent-framework --pre python-dotenv azure-identity
    2. npm install -g @microsoft/workiq   (or let npx handle it)
    3. npx -y @microsoft/workiq accept-eula
    4. npx -y @microsoft/workiq ask -q "test"   (triggers browser auth)
    5. Copy .env.example to .env and fill in your Foundry details

Usage:
    python agent.py
"""

import asyncio
import os
import sys

from dotenv import load_dotenv

load_dotenv()

AGENT_NAME = "WorkplaceHelper"
AGENT_INSTRUCTIONS = """\
You are a workplace assistant that helps users with their Microsoft 365 data.
You can query emails, calendar events, meetings, documents, Teams messages,
and information about people in the organization.

When answering questions:
- Be concise and direct
- Include relevant dates, names, and specifics
- If the user asks about meetings, include key discussion points
- If the user asks about emails, summarize the important content
- Always clarify if you need more context to answer accurately
"""


async def main() -> None:
    import time
    from collections.abc import Awaitable, Callable

    from agent_framework import Agent, FunctionInvocationContext, MCPStdioTool
    from agent_framework.azure import AzureOpenAIResponsesClient

    # --- Function middleware to show tool calls in real time ---
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

    # --- LLM Client (Microsoft Foundry) ---
    project_endpoint = os.environ["AZURE_AI_PROJECT_ENDPOINT"]
    api_key = os.environ["AZURE_AI_API_KEY"]
    deployment = os.environ["AZURE_AI_MODEL_DEPLOYMENT_NAME"]

    # Extract resource endpoint (everything before /api/projects/)
    base = project_endpoint.split("/api/projects/")[0]

    client = AzureOpenAIResponsesClient(
        base_url=f"{base}/openai/v1/",
        api_key=api_key,
        deployment_name=deployment,
        middleware=[log_tool_calls],
    )

    # --- WorkIQ MCP Tool (stdio) ---
    # WorkIQ MCP doesn't support prompts/list, so disable prompt loading
    workiq_tool = MCPStdioTool(
        name="WorkIQ",
        command="npx",
        args=["-y", "@microsoft/workiq", "mcp"],
        description="Query Microsoft 365 data — emails, meetings, documents, Teams messages, people",
        approval_mode="never_require",
        load_prompts=False,
    )

    print(f"Starting {AGENT_NAME}...")
    print("Connecting to WorkIQ MCP server...\n")

    async with workiq_tool:
        agent = Agent(
            client=client,
            name=AGENT_NAME,
            instructions=AGENT_INSTRUCTIONS,
            tools=workiq_tool,
        )

        # Create a session for multi-turn conversation
        session = agent.create_session()

        print("=" * 60)
        print(f"  {AGENT_NAME} — powered by Microsoft Agent Framework")
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

            try:
                result = await agent.run(user_input, session=session)
                print(f"\nAgent: {result}\n")
            except Exception as e:
                print(f"\n❌ Error: {e}\n", file=sys.stderr)


if __name__ == "__main__":
    asyncio.run(main())

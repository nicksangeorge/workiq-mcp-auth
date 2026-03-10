# Copyright (c) Microsoft. All rights reserved.
# Licensed under the MIT License.

"""
LLM Connectivity Smoke Test

Verifies that your Microsoft Foundry endpoint and model deployment
are correctly configured before adding MCP tools.

Usage:
    python test_llm.py
"""

import asyncio
import os

from dotenv import load_dotenv

load_dotenv()


async def main() -> None:
    from agent_framework import Agent
    from agent_framework.azure import AzureOpenAIResponsesClient

    # Microsoft Foundry — derive the OpenAI-compatible endpoint from
    # the project endpoint. The project URL looks like:
    #   https://<resource>.services.ai.azure.com/api/projects/<project>
    # The OpenAI-compatible base is the resource root:
    #   https://<resource>.services.ai.azure.com
    project_endpoint = os.environ["AZURE_AI_PROJECT_ENDPOINT"]
    api_key = os.environ["AZURE_AI_API_KEY"]
    deployment = os.environ["AZURE_AI_MODEL_DEPLOYMENT_NAME"]

    # Extract resource endpoint (everything before /api/projects/)
    base = project_endpoint.split("/api/projects/")[0]

    client = AzureOpenAIResponsesClient(
        base_url=f"{base}/openai/v1/",
        api_key=api_key,
        deployment_name=deployment,
    )

    agent = Agent(
        client=client,
        name="SmokeTest",
        instructions="You are a helpful assistant. Keep answers to one sentence.",
    )

    print("Testing LLM connectivity...")
    result = await agent.run("Say hello and confirm you are working.")
    print(f"Response: {result}")
    print("\n✅ LLM connectivity verified!")


if __name__ == "__main__":
    asyncio.run(main())

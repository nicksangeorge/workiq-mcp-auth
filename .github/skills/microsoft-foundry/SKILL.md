---
name: microsoft-foundry
description: "Deploy, evaluate, and manage Foundry agents end-to-end: Docker build, ACR push, hosted/prompt agent create, container start, batch eval, prompt optimization, agent.yaml, dataset curation from traces. USE FOR: deploy agent to Foundry, hosted agent, create agent, invoke agent, evaluate agent, run batch eval, optimize prompt, deploy model, Foundry project, RBAC, role assignment, permissions, quota, capacity, region, troubleshoot agent, deployment failure, create dataset from traces, dataset versioning, eval trending, create AI Services, Cognitive Services, create Foundry resource, provision resource, knowledge index, agent monitoring, customize deployment, onboard, availability, standard agent setup, capability host. DO NOT USE FOR: Azure Functions, App Service, general Azure deploy (use azure-deploy), general Azure prep (use azure-prepare)."
license: MIT
metadata:
  author: Microsoft
  version: "1.0.3"
---

# Microsoft Foundry Skill

> **MANDATORY:** Read this skill and the relevant sub-skill BEFORE calling any Foundry MCP tool.

## Sub-Skills

| Sub-Skill | When to Use | Reference |
|-----------|-------------|-----------|
| **deploy** | Containerize, build, push to ACR, create/update/start/stop/clone agent deployments | [deploy](foundry-agent/deploy/deploy.md) |
| **invoke** | Send messages to an agent, single or multi-turn conversations | [invoke](foundry-agent/invoke/invoke.md) |
| **observe** | Eval-driven optimization loop: evaluate → analyze → optimize → compare → iterate | [observe](foundry-agent/observe/observe.md) |
| **trace** | Query traces, analyze latency/failures, correlate eval results to specific responses via App Insights `customEvents` | [trace](foundry-agent/trace/trace.md) |
| **troubleshoot** | View container logs, query telemetry, diagnose failures | [troubleshoot](foundry-agent/troubleshoot/troubleshoot.md) |
| **create** | Create new hosted agent applications. Supports Microsoft Agent Framework, LangGraph, or custom frameworks in Python or C#. Downloads starter samples from foundry-samples repo. | [create](foundry-agent/create/create.md) |
| **eval-datasets** | Harvest production traces into evaluation datasets, manage dataset versions and splits, track evaluation metrics over time, detect regressions, and maintain full lineage from trace to deployment. | [eval-datasets](foundry-agent/eval-datasets/eval-datasets.md) |
| **project/create** | Creating a new Microsoft Foundry project for hosting agents and models. | [project/create/create-foundry-project.md](project/create/create-foundry-project.md) |
| **resource/create** | Creating Azure AI Services multi-service resource (Foundry resource) using Azure CLI. | [resource/create/create-foundry-resource.md](resource/create/create-foundry-resource.md) |
| **models/deploy-model** | Unified model deployment with intelligent routing. | [models/deploy-model/SKILL.md](models/deploy-model/SKILL.md) |
| **quota** | Managing quotas and capacity for Microsoft Foundry resources. | [quota/quota.md](quota/quota.md) |
| **rbac** | Managing RBAC permissions, role assignments, managed identities, and service principals. | [rbac/rbac.md](rbac/rbac.md) |

Onboarding flow: `project/create` → `deploy` → `invoke`

## Agent Lifecycle

| Intent | Workflow |
|--------|----------|
| New agent from scratch | create → deploy → invoke |
| Deploy existing code | deploy → invoke |
| Test/chat with agent | invoke |
| Troubleshoot | invoke → troubleshoot |
| Fix + redeploy | troubleshoot → fix → deploy → invoke |

## Project Context Resolution

Resolve only missing values. Extract from user message first, then azd, then ask.

1. Check for `azure.yaml`; if found, run `azd env get-values`
2. Map azd variables:

| azd Variable | Resolves To |
|-------------|-------------|
| `AZURE_AI_PROJECT_ENDPOINT` / `AZURE_AIPROJECT_ENDPOINT` | Project endpoint |
| `AZURE_CONTAINER_REGISTRY_NAME` / `AZURE_CONTAINER_REGISTRY_ENDPOINT` | ACR registry |
| `AZURE_SUBSCRIPTION_ID` | Subscription |

3. Ask user only for unresolved values (project endpoint, agent name)

## Validation

After each workflow step, validate before proceeding:
1. Run the operation
2. Check output for errors or unexpected results
3. If failed → diagnose using troubleshoot sub-skill → fix → retry
4. Only proceed to next step when validation passes

## Agent Types

| Type | Kind | Description |
|------|------|-------------|
| **Prompt** | `"prompt"` | LLM-based, backed by model deployment |
| **Hosted** | `"hosted"` | Container-based, running custom code |

## Agent: Setup Types

| Setup | Capability Host | Description |
|-------|----------------|-------------|
| **Basic** | None | Default. All resources Microsoft-managed. |
| **Standard** | Azure AI Services | Bring-your-own storage and search (public network). |
| **Standard + Private Network** | Azure AI Services | Standard setup with VNet isolation and private endpoints. |

## Tool Usage Conventions

- Use the `ask_user` or `askQuestions` tool whenever collecting information from the user
- Use the `task` or `runSubagent` tool to delegate long-running or independent sub-tasks
- Prefer Azure MCP tools over direct CLI commands when available
- Reference official Microsoft documentation URLs instead of embedding CLI command syntax

## References

- [Hosted Agents](https://learn.microsoft.com/azure/ai-foundry/agents/concepts/hosted-agents?view=foundry)
- [Runtime Components](https://learn.microsoft.com/azure/ai-foundry/agents/concepts/runtime-components?view=foundry)
- [Foundry Samples](https://github.com/azure-ai-foundry/foundry-samples)

## Dependencies

Scripts in sub-skills require: Azure CLI (`az`) ≥2.0, `jq` (for shell scripts). Install via `pip install azure-ai-projects azure-identity` for Python SDK usage.

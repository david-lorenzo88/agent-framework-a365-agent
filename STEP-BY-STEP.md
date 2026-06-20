# Step-by-Step: Deploy this Agent 365 agent (cross-tenant)

End-to-end guide to take this repo from a clone to a working agent that replies in
Microsoft Teams **and** exports telemetry to Agent 365 — in a **cross-tenant** setup.

> **The model.** The agent's *compute* runs in **Tenant 1** (an Azure subscription's
> Azure Container Apps). The agent's *identity, telemetry, and Teams presence* live in
> **Tenant 2** (the Microsoft 365 / Agent 365 tenant). The only links between them are
> (1) the messaging-endpoint URL registered in Tenant 2 points at the Tenant-1 Container
> App, and (2) the Tenant-2 blueprint credentials are deployed as Container App secrets.

Replace every `<<PLACEHOLDER>>` with your own value. Real values for an existing
deployment live in the git-ignored `deploy/.env.aca`.

---

## 0. Prerequisites

- **Tenant 2** (Microsoft 365): at least one **Microsoft 365 Copilot** or **Microsoft Agent 365**
  license assigned, and the tenant **enrolled in the Frontier preview program**.
  Roles: **Agent ID Developer** (blueprint) + **Global Administrator** (consent).
- **Tenant 1**: an Azure subscription where you can create resources (Contributor).
- **Azure OpenAI** resource with a chat model deployment (e.g. `gpt-4.1`).
- Local tools: **Python 3.11+**, **[uv](https://docs.astral.sh/uv/)**, **Azure CLI** (`az`),
  **.NET SDK** (for the A365 CLI), and optionally **Docker** (not required — images build in ACR).

---

## 1. Clone

```bash
git clone https://github.com/david-lorenzo88/agent-framework-a365-agent.git
cd agent-framework-a365-agent
```

---

## 2. Install tooling

### Azure CLI + Container Apps extension + resource providers (Tenant 1 prep)
```bash
az extension add --name containerapp
az provider register --namespace Microsoft.App
az provider register --namespace Microsoft.OperationalInsights
```

### Agent 365 CLI (.NET global tool)
```bash
dotnet tool install --global Microsoft.Agents.A365.DevTools.Cli --prerelease
```
Make sure `~/.dotnet/tools` is on your `PATH`.

> **macOS (Apple Silicon) only:** if `a365` fails with an `incompatible architecture`
> / `libhostfxr.dylib` error, point it at the arm64 runtime:
> ```bash
> echo 'export DOTNET_ROOT="/opt/homebrew/opt/dotnet/libexec"' >> ~/.zprofile
> source ~/.zprofile
> ```

Verify: `a365 --version`

---

## 3. Tenant 2 — register the agent (Entra ID + Agent 365)

Sign in to **Tenant 2**:
```bash
az login --tenant <<TENANT2_TENANT_ID>>
az account set --subscription <<TENANT2_SUBSCRIPTION_OR_DIRECTORY>>
```

### 3.1 Create the blueprint (Entra app registration)
```bash
a365 setup blueprint --agent-name <<AGENT_NAME>> --no-endpoint --verbose
```
This creates the blueprint Entra app + service principal and prints:
- **Blueprint App ID** → your `CLIENT_ID` / `CONN_CLIENTID`
- **Client secret** → your `CLIENT_SECRET` / `CONN_CLIENTSECRET` (shown once; retrieve later
  with `a365 setup blueprint --show-secret`)
- **Tenant ID** → your `TENANT_ID`

### 3.2 Configure inheritable permissions
```bash
a365 setup permissions mcp --agent-name <<AGENT_NAME>> --verbose   # MCP tool servers
a365 setup permissions bot --agent-name <<AGENT_NAME>> --verbose   # Bot API + Observability + Power Platform
```

### 3.3 Grant admin consent for the required permissions
`a365 setup permissions …` opens a browser for tenant-wide admin consent. **If your
Conditional Access policy blocks that page** (or the consent doesn't land), grant the
permissions directly via Microsoft Graph. These resource app IDs are Microsoft
first-party constants:

| Resource | App ID | Scope(s) this agent needs |
|---|---|---|
| Microsoft Graph | `00000003-0000-0000-c000-000000000000` | `Mail.ReadWrite Mail.Send Chat.ReadWrite Files.ReadWrite.All Sites.Read.All ChannelMessage.Read.All ChannelMessage.Send User.Read.All` |
| Messaging Bot API | `5a807f24-c9de-44ee-a3a7-329e88a00ffc` | `AgentData.ReadWrite` |
| Agent Tools (MCP) | `ea9ffc3e-8a23-4a7d-836d-234d7c7565c1` | `McpServersMetadata.Read.All McpServers.Mail.All` |
| Power Platform API | `8578e004-a5c6-46e7-913e-12f58912df43` | `Connectivity.Connections.Read` |
| Observability API | `9b975845-388f-4429-889e-eab1ef63949c` | `Agent365.Observability.OtelWrite` *(delegated — needed by the FIC telemetry flow)* |

Graph-direct consent (run as a Global Admin signed into Tenant 2):
```bash
BP_SP=$(az ad sp show --id <<CLIENT_ID>> --query id -o tsv)   # blueprint service principal

grant() {  # grant <resourceAppId> "<space-separated scopes>" as tenant-wide admin consent
  local RES_SP; RES_SP=$(az ad sp show --id "$1" --query id -o tsv)
  local EXISTING; EXISTING=$(az rest --method GET \
    --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?\$filter=clientId eq '$BP_SP' and resourceId eq '$RES_SP'" \
    --query "value[0].id" -o tsv 2>/dev/null)
  if [ -n "$EXISTING" ] && [ "$EXISTING" != "null" ]; then
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$EXISTING" \
      --headers "Content-Type=application/json" --body "{\"scope\":\"$2\"}"
  else
    az rest --method POST --url "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" \
      --headers "Content-Type=application/json" \
      --body "{\"clientId\":\"$BP_SP\",\"consentType\":\"AllPrincipals\",\"resourceId\":\"$RES_SP\",\"scope\":\"$2\"}"
  fi
}

grant 5a807f24-c9de-44ee-a3a7-329e88a00ffc "AgentData.ReadWrite"
grant ea9ffc3e-8a23-4a7d-836d-234d7c7565c1 "McpServersMetadata.Read.All McpServers.Mail.All"
grant 8578e004-a5c6-46e7-913e-12f58912df43 "Connectivity.Connections.Read"
grant 9b975845-388f-4429-889e-eab1ef63949c "Agent365.Observability.OtelWrite"
```

> **Why these:** `AgentData.ReadWrite` lets the agent get a channel token (otherwise Teams
> messages return HTTP 500). The MCP scopes power the Mail tools. The **delegated**
> `Agent365.Observability.OtelWrite` is required because the telemetry exporter uses the
> SDK's FIC (user-token) flow — the app-role alone is not enough.

### 3.4 Verify
```bash
a365 query-entra inheritance --agent-name <<AGENT_NAME>>
```
Expect **5 of 5 resources** with `Effective inheritance: OK`.

---

## 4. Azure OpenAI

Note your resource's **endpoint**, **API key**, and **deployment name** (e.g. `gpt-4.1`).
The Agent Framework client targets the Azure `/openai/v1/` surface, so the API version
**must be** `preview` (already set in the template).

---

## 5. First deploy to Azure Container Apps (Tenant 1)

Sign in to **Tenant 1** and select the hosting subscription:
```bash
az login --tenant <<TENANT1_TENANT_ID>>
az account set --subscription <<TENANT1_AZURE_SUBSCRIPTION_ID>>
```

Create your deployment config from the template:
```bash
cp deploy/.env.aca.template deploy/.env.aca
```
Fill in `deploy/.env.aca` with the Azure target + the **Tenant-2 blueprint** values from
step 3 + the Azure OpenAI values from step 4. **Leave `A365_AGENT_APP_INSTANCE_ID` and
`A365_AGENTIC_USER_ID` blank for now** — those don't exist until the agent is approved
(step 7). Then:

```bash
./deploy/deploy-aca.sh
```
This builds the image in ACR, creates the Container App with external ingress on port
`3978`, injects secrets, and prints your **messaging endpoint**:
`https://<<APP>>.<<REGION>>.azurecontainerapps.io/api/messages`

Sanity check (JWT-exempt health probe):
```bash
curl https://<<APP>>.<<REGION>>.azurecontainerapps.io/api/health
```

---

## 6. Tenant 2 — register the messaging endpoint

Create `a365.config.json` in the repo root (git-ignored):
```json
{
  "agentName": "<<AGENT_NAME>>",
  "agentIdentityDisplayName": "<<AGENT_DISPLAY_NAME>>",
  "tenantId": "<<TENANT2_TENANT_ID>>",
  "needDeployment": false,
  "messagingEndpoint": "https://<<APP>>.<<REGION>>.azurecontainerapps.io/api/messages"
}
```
> `needDeployment: false` tells the CLI **you** host the compute (don't deploy to Azure).

Switch back to Tenant 2 and register the endpoint on the blueprint:
```bash
az account set --subscription <<TENANT2_SUBSCRIPTION_OR_DIRECTORY>>
a365 setup blueprint --agent-name <<AGENT_NAME>> --endpoint-only --m365 \
  --messaging-endpoint "https://<<APP>>.<<REGION>>.azurecontainerapps.io/api/messages" --verbose
```

---

## 7. Tenant 2 — publish to Teams and approve

```bash
a365 publish --agent-name <<AGENT_NAME>> --aiteammate true --verbose
```
This generates `manifest/` and `manifest/manifest.zip`. Then, as an admin:

1. Go to **https://admin.microsoft.com → Agents → All agents → Upload custom agent**.
2. Upload **`manifest/manifest.zip`**.
3. **Approve** the agent. This provisions the **agent instance** and its **agentic user**.

---

## 8. Wire the telemetry identity and redeploy

After approval, the agent instance has two new ids you need for telemetry:

- **`A365_AGENT_APP_INSTANCE_ID`** — the agent instance app id
- **`A365_AGENTIC_USER_ID`** — the agentic user's object id

Find them by sending the agent one message in Teams and reading the container log line, or
from the Agent 365 admin center:
```bash
az containerapp logs show -n <<APP_NAME>> -g <<RESOURCE_GROUP>> --tail 100 --type console \
  | grep "tenant_id="
# -> "🔍 tenant_id=<tenant>, agent_id=<A365_AGENT_APP_INSTANCE_ID>"
```

Put both into `deploy/.env.aca`, **bump `IMAGE_TAG`**, and redeploy:
```bash
./deploy/deploy-aca.sh
```

---

## 9. Test and verify

- **Teams:** message the agent — it should reply, and (with MCP) answer mail questions.
- **Telemetry export (container logs):**
  ```bash
  az containerapp logs show -n <<APP_NAME>> -g <<RESOURCE_GROUP>> --tail 80 --type console \
    | grep -i "exporting\|HTTP 200 success"
  ```
  Expect `Exporting N spans …` followed by `HTTP 200 success` with `sinks … sent`.
- **Defender (fastest authoritative check):** Advanced Hunting →
  ```kql
  CloudAppEvents
  | where Timestamp > ago(1h)
  | where RawEventData has "<<A365_AGENT_APP_INSTANCE_ID>>"
  | project Timestamp, ActionType, RawEventData
  ```
  Look for `ActionType == "InvokeAgent"`.
- **Admin center:** the agent instance's **Activity** tab populates from `invoke_agent`
  rows (with ingestion delay).

---

## Notes / gotchas already handled in this repo

These were the non-obvious fixes baked into the code/config so you don't have to rediscover them:

- **Container bind:** the server binds `HOST=0.0.0.0` in-container (not `localhost`).
- **Azure OpenAI:** `AZURE_OPENAI_API_VERSION=preview` (the `/openai/v1/` surface rejects dated versions).
- **Telemetry token:** uses the SDK's built-in **FIC** resolver (blueprint secret → agent app →
  instance → user-FIC token) on the **non-S2S** endpoint (`A365_USE_S2S_ENDPOINT=false`).
  Requires the FIC env vars + the delegated `OtelWrite` consent (step 3.3).
- **`invoke_agent` span:** the A365 MCP tooling returns a `RawAgent` with no agent-level
  telemetry, so the host emits the `invoke_agent` root span (with the full A365 attribute
  set) itself — required for the admin center / Defender agent-activity views.
- **Conditional Access:** if it blocks the consent browser page, use the Graph-direct
  grants in step 3.3.

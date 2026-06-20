# Deploying to Azure Container Apps (cross-tenant)

This deploys the agent to **Azure Container Apps in tenant 1** (compute only) while
the agent is **registered, observed, and used in Teams in tenant 2**.

> The agent's identity lives entirely in **tenant 2**. Tenant 1 just runs the
> container and exposes a public HTTPS endpoint. The two are bridged by:
> 1. the messaging endpoint URL (tenant-2 registration → tenant-1 Container App FQDN), and
> 2. the tenant-2 credentials deployed as Container App secrets.

## Order of operations

### 1. Register the agent in tenant 2 (Agent 365 Admin Center)

This mints the agentic identity and gives you the values for `deploy/.env.aca`:
the agent **client ID / secret**, **tenant 2 ID**, and the **service connection**
(`CONN_*`) used for agentic auth. Telemetry to Agent 365 flows automatically once
`AUTH_HANDLER_NAME=AGENTIC` is set — it rides on the token exchanged by this identity.

Leave the messaging endpoint blank for now; you fill it in after step 2.

### 2. Deploy to Azure Container Apps in tenant 1

```bash
cp deploy/.env.aca.template deploy/.env.aca   # fill in tenant-2 values + tenant-1 Azure target
az login                                       # sign in (tenant 1 subscription)
az extension add --name containerapp           # one-time
./deploy/deploy-aca.sh
```

The script builds the image in ACR (no local Docker needed), creates the Container
App with `--ingress external` on port `3978`, injects secrets, and prints the FQDN.

### 3. Connect Teams (tenant 2) to the tenant-1 endpoint

Paste the printed URL into the tenant-2 registration's **messaging endpoint**:

```
https://<app>.<region>.azurecontainerapps.io/api/messages
```

Then publish the agent to Teams in tenant 2. Inbound activities are JWT-signed by
the Agents channel and validated in-container against the tenant-2 `CLIENT_ID`, so
the public endpoint is safe to expose.

## Notes

- `/api/health` is JWT-exempt and used as the readiness signal.
- `HOST=0.0.0.0` is required in-container (set by the Dockerfile/script); locally the
  server still defaults to `localhost`.
- `deploy/.env.aca` is git-ignored — never commit real secrets.
- `ENABLE_SENSITIVE_DATA` is `false` here (production default); the local `.env.template`
  sets it `true` for debugging.

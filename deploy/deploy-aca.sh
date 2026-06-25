#!/usr/bin/env bash
# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
#
# Deploy the Agent Framework sample to Azure Container Apps (tenant 1 = compute).
# All agent identity / telemetry config comes from the tenant 2 Agent 365
# registration — see deploy/.env.aca.template.
#
# Usage:
#   cp deploy/.env.aca.template deploy/.env.aca   # then fill it in
#   ./deploy/deploy-aca.sh
#
# Prereqs: az CLI logged in (az login), containerapp extension
# (az extension add --name containerapp), and Docker NOT required — the image
# is built remotely with `az acr build`.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SAMPLE_DIR="$(dirname "$SCRIPT_DIR")"
ENV_FILE="$SCRIPT_DIR/.env.aca"

if [[ ! -f "$ENV_FILE" ]]; then
  echo "❌ $ENV_FILE not found. Copy deploy/.env.aca.template to deploy/.env.aca and fill it in." >&2
  exit 1
fi

# shellcheck disable=SC1090
set -a; source "$ENV_FILE"; set +a

IMAGE="${ACR_NAME}.azurecr.io/${APP_NAME}:${IMAGE_TAG}"

echo "==> Targeting tenant-1 subscription: $SUBSCRIPTION_ID"
az account set --subscription "$SUBSCRIPTION_ID"

echo "==> Ensuring resource group: $RESOURCE_GROUP ($LOCATION)"
az group create --name "$RESOURCE_GROUP" --location "$LOCATION" --output none

echo "==> Ensuring container registry: $ACR_NAME"
az acr show --name "$ACR_NAME" --output none 2>/dev/null || \
  az acr create --resource-group "$RESOURCE_GROUP" --name "$ACR_NAME" \
    --sku Basic --admin-enabled true --output none

echo "==> Building image remotely in ACR: $IMAGE"
az acr build --registry "$ACR_NAME" --image "${APP_NAME}:${IMAGE_TAG}" "$SAMPLE_DIR"

echo "==> Ensuring Container Apps environment: $CONTAINERAPP_ENV"
az containerapp env show --name "$CONTAINERAPP_ENV" --resource-group "$RESOURCE_GROUP" --output none 2>/dev/null || \
  az containerapp env create --name "$CONTAINERAPP_ENV" --resource-group "$RESOURCE_GROUP" \
    --location "$LOCATION" --output none

ACR_PASSWORD="$(az acr credential show --name "$ACR_NAME" --query 'passwords[0].value' -o tsv)"

# Secrets (never passed as plain env-vars).
SECRETS=(
  "client-secret=${CLIENT_SECRET}"
  "conn-clientsecret=${CONN_CLIENTSECRET}"
  "azure-openai-key=${AZURE_OPENAI_API_KEY}"
)

# Environment variables. Identity values come from tenant 2; *-secret values are
# referenced from the secret store via secretref. The AGENTAPPLICATION__* and
# CONNECTIONSMAP_* keys mirror .env.template and drive the agentic auth handler
# that powers Agent 365 telemetry export.
ENV_VARS=(
  "HOST=0.0.0.0"
  "PORT=3978"
  "PYTHON_ENVIRONMENT=production"
  "LOG_LEVEL=INFO"

  # Inbound JWT validation (tenant 2 agent identity)
  "CLIENT_ID=${CLIENT_ID}"
  "TENANT_ID=${TENANT_ID}"
  "CLIENT_SECRET=secretref:client-secret"

  # Agentic auth — required for telemetry export to Agent 365 (tenant 2)
  "USE_AGENTIC_AUTH=true"
  "AUTH_HANDLER_NAME=AGENTIC"
  "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID=${CONN_CLIENTID}"
  "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET=secretref:conn-clientsecret"
  "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID=${CONN_TENANTID}"
  "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__SCOPES=${CONN_SCOPES}"
  "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__AGENTIC__SETTINGS__TYPE=AgenticUserAuthorization"
  "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__AGENTIC__SETTINGS__SCOPES=https://graph.microsoft.com/.default"
  "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__AGENTIC__SETTINGS__ALTERNATEBLUEPRINTCONNECTIONNAME=SERVICE_CONNECTION"
  "CONNECTIONSMAP_0_SERVICEURL=*"
  "CONNECTIONSMAP_0_CONNECTION=SERVICE_CONNECTION"

  # Observability — export traces/metrics/logs to Agent 365 (tenant 2)
  "ENABLE_OBSERVABILITY=true"
  "ENABLE_A365_OBSERVABILITY_EXPORTER=true"
  "ENABLE_OTEL=true"
  "ENABLE_SENSITIVE_DATA=true"
  "OBSERVABILITY_SERVICE_NAME=${OBSERVABILITY_SERVICE_NAME}"
  "OBSERVABILITY_SERVICE_NAMESPACE=${OBSERVABILITY_SERVICE_NAMESPACE}"
  # Telemetry auth via the SDK's built-in FIC (app-token) resolver
  "A365_AGENT_APP_INSTANCE_ID=${A365_AGENT_APP_INSTANCE_ID}"
  "A365_AGENTIC_USER_ID=${A365_AGENTIC_USER_ID}"
  "A365_USE_S2S_ENDPOINT=${A365_USE_S2S_ENDPOINT}"

  # LLM (Azure OpenAI — the agent uses this exclusively, see agent.py)
  "AZURE_OPENAI_ENDPOINT=${AZURE_OPENAI_ENDPOINT}"
  "AZURE_OPENAI_DEPLOYMENT=${AZURE_OPENAI_DEPLOYMENT}"
  "AZURE_OPENAI_API_VERSION=${AZURE_OPENAI_API_VERSION}"
  "AZURE_OPENAI_API_KEY=secretref:azure-openai-key"
)

echo "==> Deploying Container App: $APP_NAME"
if az containerapp show --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" --output none 2>/dev/null; then
  az containerapp secret set --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" \
    --secrets "${SECRETS[@]}" --output none
  az containerapp update --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" \
    --image "$IMAGE" --set-env-vars "${ENV_VARS[@]}" --output none
else
  az containerapp create --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" \
    --environment "$CONTAINERAPP_ENV" \
    --image "$IMAGE" \
    --registry-server "${ACR_NAME}.azurecr.io" \
    --registry-username "$ACR_NAME" \
    --registry-password "$ACR_PASSWORD" \
    --target-port 3978 \
    --ingress external \
    --min-replicas 1 \
    --secrets "${SECRETS[@]}" \
    --env-vars "${ENV_VARS[@]}" \
    --output none
fi

FQDN="$(az containerapp show --name "$APP_NAME" --resource-group "$RESOURCE_GROUP" \
  --query 'properties.configuration.ingress.fqdn' -o tsv)"

echo ""
echo "✅ Deployed. Messaging endpoint (paste into the tenant-2 Agent 365 registration):"
echo ""
echo "    https://${FQDN}/api/messages"
echo ""
echo "   Health: https://${FQDN}/api/health"

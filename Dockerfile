# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

FROM python:3.11-slim AS base

# uv for fast, prerelease-aware dependency resolution (matches pyproject.toml).
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    UV_PRERELEASE=allow \
    UV_LINK_MODE=copy

WORKDIR /app

# Install dependencies first (cached layer) from the manifest only.
# --no-install-project: install deps into /app/.venv but don't build this flat
# (non-package) app as a wheel — it runs directly from source below.
COPY pyproject.toml ./
RUN uv sync --no-dev --no-install-project

# Copy the application source.
COPY . .

# Container hosting: bind all interfaces so Azure Container Apps ingress can
# reach the server. PORT must match the Container App target port.
ENV HOST=0.0.0.0 \
    PORT=3978 \
    PATH="/app/.venv/bin:$PATH"

EXPOSE 3978

CMD ["python", "start_with_generic_host.py"]

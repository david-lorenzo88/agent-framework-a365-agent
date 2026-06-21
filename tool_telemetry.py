# Copyright (c) Microsoft. All rights reserved.

"""Per-tool-call telemetry for the Agent 365 observability pipeline.

Why this exists
---------------
The A365 MCP tooling extension builds the agent as a ``RawAgent`` (see
``McpToolRegistrationService.add_tool_servers_to_agent``). ``RawAgent`` is the
minimal Agent Framework agent — it has **no** ``AgentTelemetryLayer`` — so the
individual tool invocations the model makes are not traced out of the box. The
host already emits the per-turn ``invoke_agent`` root span explicitly
(``host_agent_server.py``) for the same reason; this module does the equivalent
for the tool calls that hang underneath it.

The Agent 365 observability exporter only ingests spans whose
``gen_ai.operation.name`` is in a known set (``invoke_agent``, ``execute_tool``,
``chat`` …). So to make MCP (and any other) tool calls show up in the M365 admin
center / Defender agent-activity views, we emit an ``execute_tool`` child span
for every tool invocation, nested under the active ``invoke_agent`` span. The
A365 enriching span processor stamps ``tenant_id`` / ``gen_ai.agent.id`` onto
the span from the per-turn baggage, so those are not set here.

How it attaches
---------------
``RawAgent`` ignores agent-level ``middleware``, but it *does* run middleware
that a context provider contributes via ``SessionContext.extend_middleware``
during ``before_run``. So we ship a tiny :class:`ToolTelemetryProvider` that
injects a :class:`ToolTelemetryMiddleware`; ``agent.py`` appends the provider to
the agent's ``context_providers`` after the SDK builds it.
"""

from __future__ import annotations

import logging
import os
from typing import Any, Awaitable, Callable, Mapping

from agent_framework import (
    ContextProvider,
    FunctionInvocationContext,
    FunctionMiddleware,
)

from microsoft_agents_a365.observability.core import (
    AgentDetails,
    ExecuteToolScope,
    Request,
    ToolCallDetails,
)

logger = logging.getLogger(__name__)


def _agent_details() -> AgentDetails:
    """Best-effort static agent identity for tool spans.

    ``tenant_id`` and ``gen_ai.agent.id`` are stamped by the A365 enriching span
    processor from the active baggage (set per turn in ``host_agent_server.py``),
    so we only supply the identity bits we know from the environment.
    """
    return AgentDetails(
        agent_id=os.environ.get("A365_AGENT_APP_INSTANCE_ID", "") or "",
        agent_name=os.environ.get("OBSERVABILITY_SERVICE_NAME") or "Agent",
        agent_blueprint_id=(
            os.environ.get("CLIENT_ID")
            or os.environ.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID")
        ),
        agentic_user_id=os.environ.get("A365_AGENTIC_USER_ID"),
        provider_name="microsoft.agent_framework",
    )


def _arguments_to_dict(arguments: Any) -> dict[str, Any] | None:
    """Normalize Agent Framework tool arguments (BaseModel or Mapping) to a dict."""
    if arguments is None:
        return None
    if isinstance(arguments, Mapping):
        return dict(arguments)
    model_dump = getattr(arguments, "model_dump", None)
    if callable(model_dump):
        try:
            return model_dump()
        except Exception:  # pragma: no cover - defensive
            return None
    return None


def _result_to_recordable(result: Any) -> dict[str, object] | str:
    """ExecuteToolScope.record_response expects a dict or a string."""
    if isinstance(result, (dict, str)):
        return result
    return str(result)


class ToolTelemetryMiddleware(FunctionMiddleware):
    """Emits one ``execute_tool`` span per tool invocation for Agent 365.

    Telemetry never breaks a tool call: if span setup fails the tool still runs,
    and any tool exception is recorded on the span and re-raised unchanged.
    """

    async def process(
        self,
        context: FunctionInvocationContext,
        call_next: Callable[[], Awaitable[None]],
    ) -> None:
        tool_name = getattr(context.function, "name", None) or "tool"

        scope: ExecuteToolScope | None = None
        try:
            details = ToolCallDetails(
                tool_name=tool_name,
                arguments=_arguments_to_dict(context.arguments),
                description=getattr(context.function, "description", None),
                tool_type="mcp",
            )
            scope = ExecuteToolScope.start(
                request=Request(),
                details=details,
                agent_details=_agent_details(),
            )
        except Exception as e:  # never let telemetry break a tool call
            logger.debug("execute_tool span setup failed for '%s': %s", tool_name, e)
            scope = None

        if scope is None:
            await call_next()
            return

        # Entering the scope makes the execute_tool span the active span and ends
        # it (recording any exception) on exit. It nests under whatever span is
        # current — the per-turn invoke_agent span emitted by the host.
        with scope:
            try:
                await call_next()
            except Exception as e:
                scope.record_error(e)
                raise
            try:
                if context.result is not None:
                    scope.record_response(_result_to_recordable(context.result))
            except Exception as e:  # pragma: no cover - defensive
                logger.debug("execute_tool result record failed for '%s': %s", tool_name, e)


class ToolTelemetryProvider(ContextProvider):
    """Injects :class:`ToolTelemetryMiddleware` into every agent run.

    ``RawAgent`` does not run agent-level middleware, but it does run middleware
    contributed by a context provider via ``SessionContext.extend_middleware``.
    """

    def __init__(self, source_id: str = "a365-tool-telemetry") -> None:
        super().__init__(source_id=source_id)
        self._middleware = ToolTelemetryMiddleware()

    async def before_run(self, *, agent, session, context, state) -> None:  # type: ignore[override]
        context.extend_middleware(self.source_id, [self._middleware])

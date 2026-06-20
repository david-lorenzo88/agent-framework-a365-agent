# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Generic Agent Host Server - Hosts agents implementing AgentInterface"""

# --- Imports ---
import asyncio
import json
import logging
import os
import socket
from os import environ

from aiohttp.web import Application, Request, Response, json_response, run_app
from aiohttp.web_middlewares import middleware as web_middleware
from dotenv import load_dotenv
from agent_interface import AgentInterface, check_agent_inheritance
from microsoft_agents.activity import load_configuration_from_env, Activity, ActivityTypes
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import (
    CloudAdapter,
    jwt_authorization_middleware,
    start_agent_process,
)
from microsoft_agents.hosting.core import (
    AgentApplication,
    AgentAuthConfiguration,
    AuthenticationConstants,
    Authorization,
    ClaimsIdentity,
    MemoryStorage,
    TurnContext,
    TurnState,
)
from microsoft_agents_a365.notifications.agent_notification import (
    AgentNotification,
    NotificationTypes,
    AgentNotificationActivity,
    ChannelId,
)
from microsoft_agents_a365.notifications import EmailResponse

from microsoft.opentelemetry import use_microsoft_opentelemetry
from microsoft_agents_a365.observability.core.middleware.baggage_builder import (
    BaggageBuilder,
)

# --- Configuration ---
ms_agents_logger = logging.getLogger("microsoft_agents")
ms_agents_logger.addHandler(logging.StreamHandler())
ms_agents_logger.setLevel(logging.INFO)

observability_logger = logging.getLogger("microsoft_agents_a365.observability")
observability_logger.setLevel(logging.ERROR)

logger = logging.getLogger(__name__)

# Tracer used to emit the `invoke_agent` root span for each turn. The A365 MCP
# tooling extension swaps the telemetry-capable `Agent` for a `RawAgent` (no
# AgentTelemetryLayer), so no `invoke_agent` span is emitted natively — and the
# Microsoft 365 admin center / Defender agent-activity views only ingest runs
# that have one. We emit it explicitly here, as the parent of the agent run, so
# the LLM `chat` span nests under it and the A365 enricher stamps tenant/agent id.
from opentelemetry import trace as _otel_trace

_invoke_agent_tracer = _otel_trace.get_tracer("microsoft.a365.sample")


def _set_invoke_agent_attributes(span, context, user_message: str) -> None:
    """Populate the mandatory Agent 365 attributes on the invoke_agent root span.

    The A365 enricher only stamps tenant + gen_ai.agent.id (from baggage); the
    remaining mandatory invoke_agent attributes must be set here, where we have
    the activity context. Without them the run either doesn't surface or shows
    blank fields in the M365 admin center / Defender agent-activity views.
    See: https://learn.microsoft.com/microsoft-agent-365/developer/observability-attribute-reference
    """
    act = context.activity
    recipient = getattr(act, "recipient", None)
    from_prop = getattr(act, "from_property", None)

    span.set_attribute("gen_ai.operation.name", "invoke_agent")

    conv = getattr(getattr(act, "conversation", None), "id", None)
    if conv:
        span.set_attribute("gen_ai.conversation.id", conv)

    # Blueprint appId — required for the admin center blueprint roll-up.
    blueprint_id = environ.get("CLIENT_ID") or environ.get(
        "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID"
    )
    if blueprint_id:
        span.set_attribute("microsoft.a365.agent.blueprint.id", blueprint_id)

    # Agent display name (else views show the raw GUID).
    span.set_attribute("gen_ai.agent.name", getattr(recipient, "name", None) or "Agent")

    # Channel — canonical lowercase token (e.g. "msteams").
    if getattr(act, "channel_id", None):
        span.set_attribute("microsoft.channel.name", act.channel_id)

    # Human caller identity ("who ran this agent").
    if getattr(from_prop, "aad_object_id", None):
        span.set_attribute("user.id", from_prop.aad_object_id)
    if getattr(from_prop, "name", None):
        span.set_attribute("user.name", from_prop.name)

    # AI teammate: the agent's own Entra user account (mandatory for embodied agents).
    agent_user_id = environ.get("A365_AGENTIC_USER_ID")
    if agent_user_id:
        span.set_attribute("microsoft.agent.user.id", agent_user_id)

    # Required transport attributes (placeholders accepted per the docs).
    span.set_attribute("client.address", "0.0.0.0")
    span.set_attribute(
        "server.address",
        f"{environ.get('CONTAINER_APP_NAME', 'agent')}."
        f"{environ.get('CONTAINER_APP_ENV_DNS_SUFFIX', 'local')}",
    )
    span.set_attribute("server.port", "443")

    # Request payload.
    span.set_attribute(
        "gen_ai.input.messages",
        json.dumps([{"role": "user", "content": user_message}]),
    )


load_dotenv()
agents_sdk_config = load_configuration_from_env(environ)


# --- Public API ---
def create_and_run_host(
    agent_class: type[AgentInterface], *agent_args, **agent_kwargs
):
    """Create and run a generic agent host"""
    if not check_agent_inheritance(agent_class):
        raise TypeError(
            f"Agent class {agent_class.__name__} must inherit from AgentInterface"
        )

    # Initialize Microsoft OpenTelemetry distro for observability.
    # Replaces the legacy configure() call with a single entrypoint that sets up
    # tracing, metrics, and logging pipelines including A365 telemetry export.
    # See: https://github.com/microsoft/opentelemetry-distro-python
    # Telemetry auth: use the distro's built-in DEFAULT token resolver instead of
    # overriding it. When the FIC env vars are present
    # (CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID/CLIENTSECRET/TENANTID +
    # A365_AGENT_APP_INSTANCE_ID + A365_AGENTIC_USER_ID), the default resolver runs
    # the S2S/FIC app-token flow — blueprint client secret -> agent app token via
    # fmi_path -> instance token -> user-FIC token for the Observability OtelWrite
    # scope — which the Agent 365 observability service accepts. The previous
    # agentic-user token (get_cached_agentic_token) was rejected with
    # EndpointInvalid/TenantIdInvalid by the ingestion service.
    use_microsoft_opentelemetry(
        enable_a365=True,
        enable_azure_monitor=False,
    )

    host = GenericAgentHost(agent_class, *agent_args, **agent_kwargs)
    auth_config = host.create_auth_configuration()
    host.start_server(auth_config)


# --- Generic Agent Host ---
class GenericAgentHost:
    """Generic host for agents implementing AgentInterface"""

    # --- Initialization ---
    def __init__(self, agent_class: type[AgentInterface], *agent_args, **agent_kwargs):
        if not check_agent_inheritance(agent_class):
            raise TypeError(
                f"Agent class {agent_class.__name__} must inherit from AgentInterface"
            )

        # Auth handler name can be configured via environment
        # Defaults to empty (no auth handler) - set AUTH_HANDLER_NAME=AGENTIC for production agentic auth
        self.auth_handler_name = os.getenv("AUTH_HANDLER_NAME", "") or None
        if self.auth_handler_name:
            logger.info(f"🔐 Using auth handler: {self.auth_handler_name}")
        else:
            logger.info("🔓 No auth handler configured (AUTH_HANDLER_NAME not set)")

        self.agent_class = agent_class
        self.agent_args = agent_args
        self.agent_kwargs = agent_kwargs
        self.agent_instance = None

        self.storage = MemoryStorage()
        self.connection_manager = MsalConnectionManager(**agents_sdk_config)
        self.adapter = CloudAdapter(connection_manager=self.connection_manager)
        self.authorization = Authorization(
            self.storage, self.connection_manager, **agents_sdk_config
        )
        self.agent_app = AgentApplication[TurnState](
            storage=self.storage,
            adapter=self.adapter,
            authorization=self.authorization,
            **agents_sdk_config,
        )
        self.agent_notification = AgentNotification(self.agent_app)
        self._setup_handlers()
        logger.info("✅ Notification handlers registered successfully")

    # --- Observability ---
    # Telemetry tokens for the Agent 365 exporter are acquired by the Microsoft
    # OpenTelemetry distro's built-in FIC resolver (see create_and_run_host). No
    # per-turn token exchange is needed here.

    async def _validate_agent_and_setup_context(self, context: TurnContext):
        logger.info("🔍 Validating agent and setting up context...")
        tenant_id = context.activity.recipient.tenant_id
        agent_id = context.activity.recipient.agentic_app_id
        logger.info(f"🔍 tenant_id={tenant_id}, agent_id={agent_id}")

        if not self.agent_instance:
            logger.error("Agent not available")
            await context.send_activity("❌ Sorry, the agent is not available.")
            return None

        return tenant_id, agent_id

    # --- Handlers (Messages & Notifications) ---
    def _setup_handlers(self):
        """Setup message and notification handlers"""
        # Configure auth handlers - only required when auth_handler_name is set
        handler_config = {"auth_handlers": [self.auth_handler_name]} if self.auth_handler_name else {}

        async def help_handler(context: TurnContext, _: TurnState):
            await context.send_activity(
                f"👋 **Hi there!** I'm **{self.agent_class.__name__}**, your AI assistant.\n\n"
                "How can I help you today?"
            )

        self.agent_app.conversation_update("membersAdded", **handler_config)(help_handler)
        self.agent_app.message("/help", **handler_config)(help_handler)

        # Handle agent install / uninstall events (agentInstanceCreated / InstallationUpdate)
        @self.agent_app.activity("installationUpdate")
        async def on_installation_update(context: TurnContext, _: TurnState):
            action = context.activity.action
            from_prop = context.activity.from_property
            logger.info(
                "InstallationUpdate received — Action: '%s', DisplayName: '%s', UserId: '%s'",
                action or "(none)",
                getattr(from_prop, "name", "(unknown)") if from_prop else "(unknown)",
                getattr(from_prop, "id", "(unknown)") if from_prop else "(unknown)",
            )
            if action == "add":
                await context.send_activity("Thank you for hiring me! Looking forward to assisting you in your professional journey!")
            elif action == "remove":
                await context.send_activity("Thank you for your time, I enjoyed working with you.")

        @self.agent_app.activity("message", **handler_config)
        async def on_message(context: TurnContext, _: TurnState):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).build():
                    user_message = context.activity.text or ""
                    if not user_message.strip() or user_message.strip() == "/help":
                        return

                    logger.info(f"📨 {user_message}")

                    # Multiple messages pattern: send an immediate acknowledgment before the LLM work begins.
                    # Each send_activity call produces a discrete Teams message.
                    # NOTE: For Teams agentic identities, streaming is buffered into a single message by the SDK;
                    #       use send_activity for any messages that must arrive immediately.
                    await context.send_activity("Got it — working on it…")
                    await context.send_activity(Activity(type="typing"))

                    # Typing indicator loop — refreshes the "..." animation every ~4s for long-running operations.
                    # Typing indicators time out after ~5s and must be re-sent. Only visible in 1:1 and small group chats.
                    async def _typing_loop():
                        try:
                            while True:
                                await asyncio.sleep(4)
                                await context.send_activity(Activity(type="typing"))
                        except asyncio.CancelledError:
                            pass  # Expected: loop is cancelled when processing completes.

                    typing_task = asyncio.create_task(_typing_loop())
                    try:
                        # Emit the invoke_agent root span for this turn (see tracer
                        # comment above). The LLM `chat` span nests under it.
                        with _invoke_agent_tracer.start_as_current_span("invoke_agent") as inv_span:
                            _set_invoke_agent_attributes(inv_span, context, user_message)
                            response = await self.agent_instance.process_user_message(
                                user_message, self.agent_app.auth, self.auth_handler_name, context
                            )
                            inv_span.set_attribute(
                                "gen_ai.output.messages",
                                json.dumps([{"role": "assistant", "content": str(response)}]),
                            )
                        await context.send_activity(response)
                    finally:
                        typing_task.cancel()
                        try:
                            await typing_task
                        except asyncio.CancelledError:
                            pass  # Expected on cancel.

            except Exception as e:
                logger.error(f"❌ Error: {e}")
                await context.send_activity(f"Sorry, I encountered an error: {str(e)}")

        @self.agent_notification.on_agent_notification(
            channel_id=ChannelId(channel="agents", sub_channel="*"),
            **handler_config,
        )
        async def on_notification(
            context: TurnContext,
            state: TurnState,
            notification_activity: AgentNotificationActivity,
        ):
            try:
                result = await self._validate_agent_and_setup_context(context)
                if result is None:
                    return
                tenant_id, agent_id = result

                with BaggageBuilder().tenant_id(tenant_id).agent_id(agent_id).build():
                    logger.info(f"📬 {notification_activity.notification_type}")

                    if not hasattr(
                        self.agent_instance, "handle_agent_notification_activity"
                    ):
                        logger.warning("⚠️ Agent doesn't support notifications")
                        await context.send_activity(
                            "This agent doesn't support notification handling yet."
                        )
                        return

                    response = (
                        await self.agent_instance.handle_agent_notification_activity(
                            notification_activity, self.agent_app.auth, self.auth_handler_name, context
                        )
                    )

                    if notification_activity.notification_type == NotificationTypes.EMAIL_NOTIFICATION:
                        response_activity = EmailResponse.create_email_response_activity(response)
                        await context.send_activity(response_activity)
                        return

                    await context.send_activity(response)

            except Exception as e:
                logger.error(f"❌ Notification error: {e}")
                await context.send_activity(
                    f"Sorry, I encountered an error processing the notification: {str(e)}"
                )

    # --- Agent Initialization ---
    async def initialize_agent(self):
        if self.agent_instance is None:
            logger.info(f"🤖 Initializing {self.agent_class.__name__}...")
            self.agent_instance = self.agent_class(*self.agent_args, **self.agent_kwargs)
            await self.agent_instance.initialize()

    # --- Authentication ---
    def create_auth_configuration(self) -> AgentAuthConfiguration | None:
        client_id = environ.get("CLIENT_ID")
        tenant_id = environ.get("TENANT_ID")
        client_secret = environ.get("CLIENT_SECRET")

        if client_id and tenant_id and client_secret:
            logger.info("🔒 Using Client Credentials authentication")
            return AgentAuthConfiguration(
                client_id=client_id,
                tenant_id=tenant_id,
                client_secret=client_secret,
                scopes=["5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default"],
            )

        if environ.get("BEARER_TOKEN"):
            logger.info("🔑 Anonymous dev mode")
        else:
            logger.warning("⚠️ No auth env vars; running anonymous")
        return None

    # --- Server ---
    def start_server(self, auth_configuration: AgentAuthConfiguration | None = None):
        async def entry_point(req: Request) -> Response:
            return await start_agent_process(
                req, req.app["agent_app"], req.app["adapter"]
            )

        async def health(_req: Request) -> Response:
            return json_response(
                {
                    "status": "ok",
                    "agent_type": self.agent_class.__name__,
                    "agent_initialized": self.agent_instance is not None,
                }
            )

        middlewares = []
        if auth_configuration:

            @web_middleware
            async def jwt_with_health_bypass(request, handler):
                # Skip JWT validation for health endpoint so that container
                # orchestrators (Azure Container Apps, Kubernetes, App Service)
                # can reach /api/health without a bearer token.
                if request.path == "/api/health":
                    return await handler(request)
                return await jwt_authorization_middleware(request, handler)

            middlewares.append(jwt_with_health_bypass)

        @web_middleware
        async def anonymous_claims(request, handler):
            if not auth_configuration:
                request["claims_identity"] = ClaimsIdentity(
                    {
                        AuthenticationConstants.AUDIENCE_CLAIM: "anonymous",
                        AuthenticationConstants.APP_ID_CLAIM: "anonymous-app",
                    },
                    False,
                    "Anonymous",
                )
            return await handler(request)

        middlewares.append(anonymous_claims)
        app = Application(middlewares=middlewares)

        app.router.add_post("/api/messages", entry_point)
        app.router.add_get("/api/messages", lambda _: Response(status=200))
        app.router.add_get("/api/health", health)

        app["agent_configuration"] = auth_configuration
        app["agent_app"] = self.agent_app
        app["adapter"] = self.agent_app.adapter

        app.on_startup.append(lambda app: self.initialize_agent())
        app.on_shutdown.append(lambda app: self.cleanup())

        # Bind host: defaults to "localhost" for local dev. In a container
        # (Azure Container Apps, Kubernetes) set HOST=0.0.0.0 so the platform
        # ingress can reach the server — localhost would only be reachable
        # from inside the container.
        host = environ.get("HOST", "localhost")
        desired_port = int(environ.get("PORT", 3978))
        port = desired_port

        # Local convenience only: if the desired port is already taken, fall back
        # to the next one. Skip this probe when binding 0.0.0.0 (container hosting),
        # where the platform expects the server on exactly PORT.
        if host == "localhost":
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                if s.connect_ex(("127.0.0.1", desired_port)) == 0:
                    port = desired_port + 1

        print("=" * 80)
        print(f"🏢 {self.agent_class.__name__}")
        print("=" * 80)
        print(f"🔒 Auth: {'Enabled' if auth_configuration else 'Anonymous'}")
        print(f"🚀 Server: {host}:{port}")
        print(f"📚 Endpoint: http://{host}:{port}/api/messages")
        print(f"❤️  Health: http://{host}:{port}/api/health\n")

        try:
            run_app(app, host=host, port=port, handle_signals=True)
        except KeyboardInterrupt:
            print("\n👋 Server stopped")

    # --- Cleanup ---
    async def cleanup(self):
        if self.agent_instance:
            try:
                await self.agent_instance.cleanup()
            except Exception as e:
                logger.error(f"Cleanup error: {e}")




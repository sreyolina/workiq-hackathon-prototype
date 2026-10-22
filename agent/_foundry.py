r"""Shared Azure AI Foundry chat-client factory.

Every Work IQ sub-agent (and the orchestrator) uses this to construct the
OpenAI-compatible client backed by the Foundry deployment. Auth uses
`AZURE_OPENAI_API_KEY` when present; otherwise it falls back to Entra ID via
Azure credentials.
"""
from __future__ import annotations

import os

from openai import AsyncOpenAI
from azure.identity.aio import AzureCliCredential, DefaultAzureCredential, get_bearer_token_provider

from agent_framework.openai import OpenAIChatClient


FOUNDRY_ENDPOINT = os.environ.get("AZURE_OPENAI_ENDPOINT") or os.environ.get(
    "AZURE_AI_FOUNDRY_ENDPOINT"
)
DEFAULT_DEPLOYMENT = os.environ.get("AZURE_AI_FOUNDRY_DEPLOYMENT", "gpt-4o-mini")
CREDENTIAL_MODE = os.environ.get("WORKIQ_AZURE_CREDENTIAL_MODE", "cli").strip().lower()
OPENAI_API_KEY = os.environ.get("AZURE_OPENAI_API_KEY") or os.environ.get("AZURE_AI_FOUNDRY_API_KEY")
TOKEN_SCOPE = (
    "https://cognitiveservices.azure.com/.default"
    if FOUNDRY_ENDPOINT and ".openai.azure.com" in FOUNDRY_ENDPOINT
    else "https://ai.azure.com/.default"
)


def _build_credential():
        """Return an async Azure credential based on runtime mode.

        Modes:
            - cli (default): deterministic local-dev auth via `az login`
            - default: cloud-friendly chain (Managed Identity, env creds, etc.)
        """
        if CREDENTIAL_MODE == "default":
                return DefaultAzureCredential()
        return AzureCliCredential()


def build_chat_client(deployment: str | None = None) -> OpenAIChatClient:
    """Return an `OpenAIChatClient` targeting the Foundry deployment.

    Uses `AZURE_OPENAI_API_KEY` when present. Otherwise uses Azure credentials ->
    bearer token against the Azure AI/OpenAI scope; tokens auto-refresh
    because the openai SDK accepts a callable for `api_key`.
    """
    if not FOUNDRY_ENDPOINT:
        raise RuntimeError(
            "AZURE_OPENAI_ENDPOINT (or AZURE_AI_FOUNDRY_ENDPOINT) is not set. "
            "Point it at your Foundry /openai/v1 endpoint."
        )
    if OPENAI_API_KEY:
        api_key = OPENAI_API_KEY
    else:
        credential = _build_credential()
        api_key = get_bearer_token_provider(credential, TOKEN_SCOPE)
    openai_client = AsyncOpenAI(
        base_url=FOUNDRY_ENDPOINT,
        api_key=api_key,  # type: ignore[arg-type]
    )
    return OpenAIChatClient(model=deployment or DEFAULT_DEPLOYMENT, async_client=openai_client)

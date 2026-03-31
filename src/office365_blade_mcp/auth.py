"""Authentication for Microsoft Graph API and HTTP transport.

Supports two auth modes:
- **device_code**: Interactive OAuth 2.0 device code flow (default)
- **client_credentials**: App-only auth for headless/CI environments

Bearer token middleware for remote/tunnel HTTP transport access.
"""

from __future__ import annotations

import json
import logging
import os
import secrets
from pathlib import Path

import msal
from starlette.types import ASGIApp, Receive, Scope, Send

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Token cache
# ---------------------------------------------------------------------------

_TOKEN_CACHE_DIR: Path | None = None


def _get_cache_dir() -> Path:
    """Get or create the token cache directory."""
    global _TOKEN_CACHE_DIR  # noqa: PLW0603
    if _TOKEN_CACHE_DIR is None:
        cache_dir = os.environ.get("O365_TOKEN_CACHE_DIR", "").strip()
        if cache_dir:
            _TOKEN_CACHE_DIR = Path(cache_dir)
        else:
            _TOKEN_CACHE_DIR = Path.home() / ".office365-blade-mcp"
        _TOKEN_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    return _TOKEN_CACHE_DIR


def _get_token_cache() -> msal.SerializableTokenCache:
    """Load or create a persistent MSAL token cache."""
    cache = msal.SerializableTokenCache()
    cache_file = _get_cache_dir() / "token_cache.json"
    if cache_file.exists():
        cache.deserialize(cache_file.read_text())
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """Persist the MSAL token cache to disk."""
    if cache.has_state_changed:
        cache_file = _get_cache_dir() / "token_cache.json"
        cache_file.write_text(cache.serialize())
        # Restrict permissions (owner-only read/write)
        cache_file.chmod(0o600)


# ---------------------------------------------------------------------------
# MSAL app construction
# ---------------------------------------------------------------------------


def _get_tenant_id() -> str:
    tenant = os.environ.get("O365_TENANT_ID", "").strip()
    if not tenant:
        raise ValueError("O365_TENANT_ID is not set")
    return tenant


def _get_client_id() -> str:
    client_id = os.environ.get("O365_CLIENT_ID", "").strip()
    if not client_id:
        raise ValueError("O365_CLIENT_ID is not set")
    return client_id


def _get_auth_mode() -> str:
    return os.environ.get("O365_AUTH_MODE", "device_code").strip().lower()


def _build_public_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    """Build an MSAL public client app for device code flow."""
    return msal.PublicClientApplication(
        client_id=_get_client_id(),
        authority=f"https://login.microsoftonline.com/{_get_tenant_id()}",
        token_cache=cache,
    )


def _build_confidential_app(cache: msal.SerializableTokenCache) -> msal.ConfidentialClientApplication:
    """Build an MSAL confidential client app for client_credentials flow."""
    client_secret = os.environ.get("O365_CLIENT_SECRET", "").strip()
    if not client_secret:
        raise ValueError("O365_CLIENT_SECRET is required for client_credentials auth mode")
    return msal.ConfidentialClientApplication(
        client_id=_get_client_id(),
        authority=f"https://login.microsoftonline.com/{_get_tenant_id()}",
        client_credential=client_secret,
        token_cache=cache,
    )


# ---------------------------------------------------------------------------
# Token acquisition
# ---------------------------------------------------------------------------


def acquire_token(scopes: list[str]) -> str:
    """Acquire a Graph API access token using the configured auth mode.

    Returns the access token string. Raises ValueError on auth failure.
    """
    cache = _get_token_cache()
    auth_mode = _get_auth_mode()

    try:
        if auth_mode == "client_credentials":
            return _acquire_client_credentials(cache, scopes)
        else:
            return _acquire_device_code(cache, scopes)
    finally:
        _save_token_cache(cache)


def _acquire_device_code(cache: msal.SerializableTokenCache, scopes: list[str]) -> str:
    """Acquire token via device code flow (interactive)."""
    app = _build_public_app(cache)

    # Try silent acquisition first (cached token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            logger.debug("Token acquired silently (cached)")
            return result["access_token"]

    # Initiate device code flow
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise ValueError(f"Device code flow failed: {flow.get('error_description', 'unknown error')}")

    logger.info("Device code auth: %s", flow["message"])
    # Print to stderr so the user sees it in stdio transport
    import sys

    print(flow["message"], file=sys.stderr, flush=True)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "unknown error"))
        raise ValueError(f"Device code auth failed: {_scrub_secrets(error)}")

    logger.info("Token acquired via device code flow")
    return result["access_token"]


def _acquire_client_credentials(cache: msal.SerializableTokenCache, scopes: list[str]) -> str:
    """Acquire token via client credentials flow (app-only)."""
    app = _build_confidential_app(cache)
    # Client credentials always uses .default scope
    cc_scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=cc_scopes)
    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "unknown error"))
        raise ValueError(f"Client credentials auth failed: {_scrub_secrets(error)}")

    logger.info("Token acquired via client credentials flow")
    return result["access_token"]


def _scrub_secrets(text: str) -> str:
    """Remove sensitive values from error messages."""
    client_secret = os.environ.get("O365_CLIENT_SECRET", "")
    if client_secret and client_secret in text:
        text = text.replace(client_secret, "***")
    client_id = os.environ.get("O365_CLIENT_ID", "")
    if client_id and client_id in text:
        text = text.replace(client_id, "[client_id]")
    return text


# ---------------------------------------------------------------------------
# Bearer token middleware for HTTP transport
# ---------------------------------------------------------------------------

_BEARER_TOKEN: str | None = None
_BEARER_CHECKED: bool = False


def get_bearer_token() -> str | None:
    """Return the bearer token from env, or None if not configured."""
    global _BEARER_TOKEN, _BEARER_CHECKED  # noqa: PLW0603
    if _BEARER_CHECKED:
        return _BEARER_TOKEN
    _BEARER_CHECKED = True
    token = os.environ.get("O365_MCP_API_TOKEN", "").strip()
    _BEARER_TOKEN = token if token else None
    return _BEARER_TOKEN


class BearerAuthMiddleware:
    """Starlette-compatible ASGI middleware for Bearer token auth.

    When ``O365_MCP_API_TOKEN`` is set, every request must carry a matching
    ``Authorization: Bearer <token>`` header.

    If the env var is unset or empty, this middleware is a transparent pass-through.
    """

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] not in ("http", "websocket"):
            await self.app(scope, receive, send)
            return

        expected = get_bearer_token()
        if expected is None:
            await self.app(scope, receive, send)
            return

        headers = dict(scope.get("headers", []))
        auth_value = headers.get(b"authorization", b"").decode("latin-1")

        provided = ""
        if auth_value.lower().startswith("bearer "):
            provided = auth_value[7:]

        if provided and secrets.compare_digest(provided, expected):
            await self.app(scope, receive, send)
            return

        body = json.dumps({"error": "Unauthorized"}).encode()
        await send(
            {
                "type": "http.response.start",
                "status": 401,
                "headers": [
                    [b"content-type", b"application/json"],
                    [b"content-length", str(len(body)).encode()],
                ],
            }
        )
        await send({"type": "http.response.body", "body": body})

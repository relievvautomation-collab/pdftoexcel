"""
ConvertAPI credentials: sandbox vs production.

Set tokens only via environment (or a local .env file — never commit real keys).

- CONVERTAPI_ENV=production (default) → CONVERTAPI_SECRET or CONVERTAPI_PRODUCTION_TOKEN
- CONVERTAPI_ENV=sandbox → CONVERTAPI_SECRET_SANDBOX or CONVERTAPI_SANDBOX_TOKEN

Dashboard: https://www.convertapi.com/a/authentication
"""

from __future__ import annotations

import os
from typing import Tuple


def _strip(name: str) -> str:
    return (os.environ.get(name) or "").strip()


def resolve_convertapi_secret() -> Tuple[str, str]:
    """
    Return (secret, mode) where mode is 'sandbox' or 'production'.
    Raises ValueError if the selected mode has no token.
    """
    raw = (_strip("CONVERTAPI_ENV") or "production").lower()
    use_sandbox = raw in ("sandbox", "dev", "development", "test")

    if use_sandbox:
        secret = _strip("CONVERTAPI_SECRET_SANDBOX") or _strip("CONVERTAPI_SANDBOX_TOKEN")
        if not secret:
            raise ValueError(
                "Sandbox mode (CONVERTAPI_ENV=sandbox): set CONVERTAPI_SECRET_SANDBOX. "
                "Copy the token from ConvertAPI → Integration → Authentication."
            )
        return secret, "sandbox"

    secret = _strip("CONVERTAPI_SECRET") or _strip("CONVERTAPI_PRODUCTION_TOKEN")
    if not secret:
        raise ValueError(
            "Production mode: set CONVERTAPI_SECRET (or CONVERTAPI_PRODUCTION_TOKEN). "
            "Copy the token from ConvertAPI → Integration → Authentication."
        )
    return secret, "production"


def convertapi_any_token_set() -> bool:
    """True if at least one known token env var is non-empty (enables ConvertAPI in UI)."""
    return bool(
        _strip("CONVERTAPI_SECRET")
        or _strip("CONVERTAPI_PRODUCTION_TOKEN")
        or _strip("CONVERTAPI_SECRET_SANDBOX")
        or _strip("CONVERTAPI_SANDBOX_TOKEN")
    )


def convertapi_env_label() -> str:
    """Current CONVERTAPI_ENV intent: sandbox | production."""
    raw = (_strip("CONVERTAPI_ENV") or "production").lower()
    if raw in ("sandbox", "dev", "development", "test"):
        return "sandbox"
    return "production"

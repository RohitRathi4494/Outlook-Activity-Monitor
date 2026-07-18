"""Thin Microsoft Graph HTTP client: pagination + 429 retry/backoff.

Every call here takes the CALLING USER's own access token and hits /me
endpoints, so results are naturally scoped to that user by Graph itself.
"""

import asyncio
import logging
from typing import Any, Optional

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

logger = logging.getLogger("graph_client")

MAX_RETRIES = 5
DEFAULT_TIMEOUT = 30.0


async def _get_with_retry(
    client: httpx.AsyncClient,
    url: str,
    headers: dict,
    params: Optional[dict] = None,
) -> dict:
    """GET a Graph URL, retrying on HTTP 429 with exponential backoff, honoring
    the Retry-After header when Graph sends one."""
    attempt = 0
    while True:
        response = await client.get(url, headers=headers, params=params)

        if response.status_code == 429:
            if attempt >= MAX_RETRIES:
                response.raise_for_status()
            retry_after = response.headers.get("Retry-After")
            wait_seconds = float(retry_after) if retry_after else float(2 ** attempt)
            logger.warning(
                "Graph throttled request (429). Waiting %.1fs before retry %d/%d.",
                wait_seconds,
                attempt + 1,
                MAX_RETRIES,
            )
            await asyncio.sleep(wait_seconds)
            attempt += 1
            continue

        response.raise_for_status()
        return response.json()


async def get_all_pages(
    access_token: str,
    url: str,
    params: Optional[dict] = None,
) -> list[dict[str, Any]]:
    """Fetch every page of a Graph collection endpoint, following
    @odata.nextLink until exhausted, and return the concatenated 'value' items.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    items: list[dict[str, Any]] = []

    async with httpx.AsyncClient(timeout=DEFAULT_TIMEOUT) as client:
        next_url: Optional[str] = url
        next_params = params

        while next_url:
            data = await _get_with_retry(client, next_url, headers, next_params)
            items.extend(data.get("value", []))
            # @odata.nextLink is already a complete URL with its own query string.
            next_url = data.get("@odata.nextLink")
            next_params = None

    return items

"""
A small, dependency-light HTTP helper shared by every module that talks
to a network API. Centralises:

- a persistent ``requests.Session`` with a descriptive User-Agent
- retry-with-backoff for transient failures (timeouts, 5xx, connection
  errors)
- a politeness delay between calls so we do not hammer free public
  services
"""

from __future__ import annotations

import logging
import time
from typing import Any, Dict, Optional

import requests
from requests import Response

import config

logger = logging.getLogger(__name__)


class HttpClient:
    """Thin wrapper around :class:`requests.Session` with retry logic."""

    def __init__(
        self,
        timeout: int = config.REQUEST_TIMEOUT,
        max_retries: int = config.MAX_RETRIES,
        backoff: float = config.RETRY_BACKOFF_SECONDS,
        rate_limit_delay: float = config.RATE_LIMIT_DELAY_SECONDS,
    ) -> None:
        self.timeout = timeout
        self.max_retries = max_retries
        self.backoff = backoff
        self.rate_limit_delay = rate_limit_delay
        self._session = requests.Session()
        self._session.headers.update({"User-Agent": config.USER_AGENT})
        self._last_call_time: float = 0.0

    def _throttle(self) -> None:
        """Sleep just enough to respect ``rate_limit_delay`` between calls."""
        elapsed = time.monotonic() - self._last_call_time
        remaining = self.rate_limit_delay - elapsed
        if remaining > 0:
            time.sleep(remaining)
        self._last_call_time = time.monotonic()

    def get(
        self,
        url: str,
        params: Optional[Dict[str, Any]] = None,
        stream: bool = False,
    ) -> Optional[Response]:
        """Perform a GET request with retries.

        Returns the ``Response`` on success (status < 400) or ``None`` if
        every retry attempt failed. Never raises for network-level
        failures — callers should treat ``None`` as "source unavailable"
        and continue with other sources / leave the cell blank.
        """
        last_error: Optional[Exception] = None
        for attempt in range(1, self.max_retries + 1):
            self._throttle()
            try:
                response = self._session.get(
                    url, params=params, timeout=self.timeout, stream=stream
                )
                if response.status_code >= 500:
                    raise requests.HTTPError(
                        f"Server error {response.status_code}"
                    )
                if response.status_code >= 400:
                    logger.warning(
                        "GET %s -> HTTP %s (not retried, client error)",
                        url,
                        response.status_code,
                    )
                    return None
                return response
            except (
                requests.ConnectionError,
                requests.Timeout,
                requests.HTTPError,
            ) as exc:
                last_error = exc
                wait = self.backoff * attempt
                logger.warning(
                    "GET %s failed (attempt %d/%d): %s — retrying in %.1fs",
                    url,
                    attempt,
                    self.max_retries,
                    exc,
                    wait,
                )
                time.sleep(wait)
        logger.error(
            "GET %s failed after %d attempts: %s", url, self.max_retries, last_error
        )
        return None


# Module-level singleton reused by every module (keeps one connection pool
# and one rate-limit clock for the whole run).
client = HttpClient()

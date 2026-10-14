"""HTTP client for the HUAWEI consumer site (e.g. consumer.huawei.com/my/)."""

from __future__ import annotations

import logging
import re
import time
from typing import Any

import requests


class ConsumerClientError(RuntimeError):
    """Base consumer-site client error."""


class ConsumerRateLimitError(ConsumerClientError):
    """Raised when the consumer site rate-limits the client."""


class ConsumerHttpError(ConsumerClientError):
    """Raised for non-retryable HTTP errors."""

    def __init__(self, status_code: int, url: str) -> None:
        super().__init__(f"HTTP {status_code} for {url}")
        self.status_code = status_code
        self.url = url


class ConsumerClient:
    """HTTP client for static pages on consumer.huawei.com.

    The consumer site is fully server-rendered, so plain GET requests are
    enough. The client only adds throttling, retries and a browser UA.
    """

    BASE_URL_TEMPLATE = "https://consumer.huawei.com/{region}/"

    def __init__(
        self,
        *,
        region: str = "my",
        timeout: int = 30,
        session: requests.Session | None = None,
        user_agent: str | None = None,
        min_request_interval: float = 0.8,
        max_retries: int = 4,
        backoff_factor: float = 1.5,
        logger: logging.Logger | None = None,
    ) -> None:
        self.region = region.strip("/")
        self.timeout = timeout
        self.session = session or requests.Session()
        self.session.headers.update(
            {
                "User-Agent": user_agent or self._default_user_agent(),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-MY,en;q=0.9",
            }
        )
        self.min_request_interval = min_request_interval
        self.max_retries = max_retries
        self.backoff_factor = backoff_factor
        self.logger = logger or logging.getLogger(__name__)
        self._last_request_ts = 0.0

    @property
    def base_url(self) -> str:
        return self.BASE_URL_TEMPLATE.format(region=self.region)

    def _default_user_agent(self) -> str:
        return (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/136.0.0.0 Safari/537.36"
        )

    def _throttle(self) -> None:
        now = time.monotonic()
        elapsed = now - self._last_request_ts
        if elapsed < self.min_request_interval:
            sleep_for = self.min_request_interval - elapsed
            self.logger.debug("Sleeping %.2fs to respect request interval", sleep_for)
            time.sleep(sleep_for)

    def fetch_text(self, url: str) -> str:
        """Fetch a URL and return the response text with retry/throttling."""
        return self.fetch_response(url).text

    def fetch_response(self, url: str) -> requests.Response:
        """Fetch a URL and return the final response with retry/throttling.

        Redirects are followed; inspect ``response.url`` for the final URL.
        """
        attempts = max(1, self.max_retries)
        for attempt in range(1, attempts + 1):
            self._throttle()
            try:
                response = self.session.get(url, timeout=self.timeout)
                self._last_request_ts = time.monotonic()
            except requests.RequestException as exc:
                if attempt >= attempts:
                    raise ConsumerClientError(f"Request failed: {url}: {exc}") from exc
                sleep_for = self.backoff_factor * attempt
                self.logger.warning(
                    "Request error on attempt %s/%s for %s: %s. Retrying in %.2fs",
                    attempt,
                    attempts,
                    url,
                    exc,
                    sleep_for,
                )
                time.sleep(sleep_for)
                continue

            if response.status_code == 404:
                raise ConsumerHttpError(404, url)
            if response.status_code in {403, 429, 500, 502, 503, 504}:
                if attempt >= attempts:
                    raise ConsumerRateLimitError(
                        f"HTTP {response.status_code} for {url} after {attempts} attempts"
                    )
                sleep_for = self.backoff_factor * attempt
                self.logger.warning(
                    "HTTP %s on attempt %s/%s for %s. Retrying in %.2fs",
                    response.status_code,
                    attempt,
                    attempts,
                    url,
                    sleep_for,
                )
                time.sleep(sleep_for)
                continue
            if response.status_code >= 400:
                raise ConsumerHttpError(response.status_code, url)
            return response

        raise ConsumerClientError(f"Request failed unexpectedly: {url}")

    def fetch_page(self, path: str) -> str:
        """Fetch a site-relative path such as ``phones/mate-x7/specs/``."""
        path = path.lstrip("/")
        return self.fetch_text(f"{self.base_url}{path}")

    def fetch_page_with_final_url(self, path: str) -> tuple[str, str]:
        """Fetch a path and return ``(final_url, text)`` after redirects."""
        path = path.lstrip("/")
        response = self.fetch_response(f"{self.base_url}{path}")
        return response.url, response.text

    def try_fetch_page(self, path: str) -> str | None:
        """Fetch a page, returning ``None`` when it does not exist (404)."""
        try:
            return self.fetch_page(path)
        except ConsumerHttpError as exc:
            if exc.status_code == 404:
                self.logger.info("Page not found (404), skipped: %s", path)
                return None
            raise


class ConsumerSitemap:
    """Fetch and parse the region sitemap for supplementary product URLs."""

    LOC_PATTERN = re.compile(r"<loc>([^<]+)</loc>")

    def __init__(self, client: ConsumerClient) -> None:
        self.client = client
        self.logger = client.logger

    def fetch_specs_urls(self, category_paths: set[str]) -> list[str]:
        """Return site-relative product paths discovered from the sitemap.

        Any sitemap URL shaped like ``{category}/{slug}/...`` is treated as a
        product candidate (``specs/``/``buy/`` sub-pages included); only URLs
        under the configured category paths are kept.
        """
        try:
            text = self.client.fetch_text(f"{self.client.base_url}sitemap.xml")
        except ConsumerClientError as exc:
            self.logger.warning("Unable to fetch sitemap, skipped: %s", exc)
            return []

        paths: list[str] = []
        prefix = self.client.base_url
        for match in self.LOC_PATTERN.finditer(text):
            url = match.group(1).strip()
            if not url.startswith(prefix):
                continue
            rel = url[len(prefix):]
            segments = [part for part in rel.split("/") if part]
            if len(segments) < 2:
                continue
            if segments[0] not in category_paths:
                continue
            paths.append(f"{segments[0]}/{segments[1]}/")
        self.logger.info("Sitemap yielded %s candidate product paths", len(paths))
        return paths

"""Client for interacting with the OpenAPI Company search endpoint."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, Iterator, List, Optional

import requests

PRODUCTION_BASE_URL = "https://company.openapi.com"
SANDBOX_BASE_URL = "https://test.company.openapi.com"


class OpenAPIError(RuntimeError):
    """Generic error raised when interacting with the OpenAPI platform."""


class AuthenticationError(OpenAPIError):
    """Raised when the provided token is invalid or missing permissions."""


class PaymentRequiredError(OpenAPIError):
    """Raised when the account has insufficient credits."""


class ValidationError(OpenAPIError):
    """Raised when input parameters are rejected by the API."""


def sanitize_ateco_code(code: str) -> str:
    """Return the ATECO code formatted as expected by the API (digits only).

    The API accepts the ATECO code without dots. Users often provide formats
    such as "10.71" or "1071"; this helper strips non-digit characters and
    ensures the remaining value is numeric.
    """

    if not code:
        raise ValueError("ATECO code must be provided")

    digits = "".join(ch for ch in code if ch.isdigit())
    if not digits:
        raise ValueError(f"ATECO code '{code}' does not contain any digits")

    return digits


@dataclass(slots=True)
class SearchParams:
    """Strongly typed search parameters for /IT-search."""

    province: str
    ateco_code: str
    data_enrichment: str = "Advanced"
    limit: int = 100
    activity_status: Optional[str] = None
    dry_run: bool = False

    def as_query_params(self, skip: int = 0) -> Dict[str, Any]:
        params: Dict[str, Any] = {
            "province": self.province.upper(),
            "atecoCode": self.ateco_code,
            "limit": self.limit,
            "skip": skip,
        }

        if self.data_enrichment:
            params["dataEnrichment"] = self.data_enrichment
        if self.activity_status:
            params["activityStatus"] = self.activity_status
        if self.dry_run:
            params["dryRun"] = 1

        return params


class OpenAPICompanyClient:
    """Thin wrapper around the OpenAPI Company API for Italian businesses."""

    def __init__(
        self,
        token: str,
        *,
        sandbox: bool = False,
        timeout: int = 30,
        session: Optional[requests.Session] = None,
    ) -> None:
        if not token:
            raise ValueError("A valid OpenAPI bearer token must be provided")

        self.base_url = SANDBOX_BASE_URL if sandbox else PRODUCTION_BASE_URL
        self.timeout = timeout
        self.session = session or requests.Session()
        self.session.headers.setdefault("Authorization", f"Bearer {token}")
        self.session.headers.setdefault(
            "User-Agent",
            "ateco-extractor/1.0 (+https://openapi.com)",
        )

    def search_companies(
        self,
        *,
        province: str,
        ateco_code: str,
        data_enrichment: str = "Advanced",
        limit: int = 100,
        max_records: Optional[int] = None,
        activity_status: Optional[str] = None,
    ) -> Iterator[Dict[str, Any]]:
        """Yield companies returned by /IT-search matching the given filters.

        Args:
            province: Italian province code (e.g., "VR" for Verona).
            ateco_code: ATECO code to filter by (digits only).
            data_enrichment: Optional enrichment to include in the results.
                Supported values include "Start", "Advanced", "Address",
                "Pec", "Shareholders", etc. Defaults to "Advanced".
            limit: Page size for pagination. OpenAPI currently allows up to 100.
            max_records: Maximum number of records to return across all pages.
            activity_status: Optional activity status filter (e.g., "ATTIVA").
        """

        if limit <= 0:
            raise ValueError("limit must be greater than zero")

        sanitized_code = sanitize_ateco_code(ateco_code)
        params = SearchParams(
            province=province,
            ateco_code=sanitized_code,
            data_enrichment=data_enrichment,
            limit=min(limit, 100),
            activity_status=activity_status,
        )

        fetched = 0
        skip = 0

        while True:
            payload = self._get("/IT-search", params.as_query_params(skip=skip))
            data = payload.get("data", [])
            if not isinstance(data, list):
                raise OpenAPIError("Unexpected response structure: 'data' is not a list")

            if not data:
                break

            for item in data:
                yield item
                fetched += 1
                if max_records is not None and fetched >= max_records:
                    return

            if len(data) < params.limit:
                break

            skip += params.limit

    def dry_run_count(
        self,
        *,
        province: str,
        ateco_code: str,
        activity_status: Optional[str] = None,
    ) -> int:
        """Return the number of records available without consuming credits."""

        sanitized_code = sanitize_ateco_code(ateco_code)
        params = SearchParams(
            province=province,
            ateco_code=sanitized_code,
            dry_run=True,
            data_enrichment="",
            limit=1,
            activity_status=activity_status,
        )

        payload = self._get("/IT-search", params.as_query_params())
        metadata = payload.get("metadata") or {}
        total = metadata.get("total")
        if isinstance(total, int):
            return total
        if isinstance(total, str) and total.isdigit():
            return int(total)
        raise OpenAPIError("Dry run response did not include 'metadata.total'")

    def _get(self, path: str, params: Dict[str, Any]) -> Dict[str, Any]:
        url = f"{self.base_url}{path}"
        response = self.session.get(url, params=params, timeout=self.timeout)

        if response.status_code == 401:
            raise AuthenticationError("Unauthorized: check your bearer token")
        if response.status_code == 402:
            raise PaymentRequiredError("Insufficient credits on OpenAPI account")
        if response.status_code == 422:
            raise ValidationError("Request rejected: invalid parameters")

        try:
            response.raise_for_status()
        except requests.HTTPError as exc:  # pragma: no cover - defensive
            message = response.text or str(exc)
            raise OpenAPIError(f"OpenAPI request failed: {message}") from exc

        payload = response.json()
        if not payload.get("success", False):
            message = payload.get("message") or "Unknown error"
            raise OpenAPIError(f"OpenAPI reported an error: {message}")

        return payload


__all__ = [
    "OpenAPICompanyClient",
    "OpenAPIError",
    "AuthenticationError",
    "PaymentRequiredError",
    "ValidationError",
    "sanitize_ateco_code",
]

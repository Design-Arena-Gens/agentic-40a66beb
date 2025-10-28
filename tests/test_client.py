from __future__ import annotations

import pytest

from ateco_extractor.client import (
    OpenAPICompanyClient,
    OpenAPIError,
    sanitize_ateco_code,
)


class DummyResponse:
    def __init__(self, status_code: int, payload: dict):
        self.status_code = status_code
        self._payload = payload
        self.text = ""

    def json(self) -> dict:
        return self._payload

    def raise_for_status(self) -> None:
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP error {self.status_code}")


class DummySession:
    def __init__(self, responses: list[DummyResponse]):
        self._responses = responses
        self.headers = {}
        self.calls = []

    def get(self, url: str, params: dict, timeout: int):
        self.calls.append((url, params, timeout))
        if not self._responses:
            raise AssertionError("No more dummy responses available")
        return self._responses.pop(0)


def make_client(*payloads: dict) -> OpenAPICompanyClient:
    responses = [DummyResponse(200, payload) for payload in payloads]
    session = DummySession(responses)
    client = OpenAPICompanyClient("test-token", sandbox=True, session=session)
    return client


def test_sanitize_ateco_code_removes_non_digits():
    assert sanitize_ateco_code("10.71") == "1071"
    assert sanitize_ateco_code(" 6201 ") == "6201"

    with pytest.raises(ValueError):
        sanitize_ateco_code("abc")


def test_search_companies_paginates_until_empty_page():
    first_page = {
        "success": True,
        "data": [{"id": "A"}, {"id": "B"}],
    }
    second_page = {
        "success": True,
        "data": [{"id": "C"}],
    }
    third_page = {
        "success": True,
        "data": [],
    }

    client = make_client(first_page, second_page, third_page)
    results = list(
        client.search_companies(province="VR", ateco_code="10.71", limit=2)
    )

    assert [item["id"] for item in results] == ["A", "B", "C"]


def test_search_companies_stops_at_max_records():
    page = {
        "success": True,
        "data": [{"id": str(i)} for i in range(5)],
    }
    # Provide enough responses so that the client could keep fetching
    client = make_client(page, page, {"success": True, "data": []})

    results = list(
        client.search_companies(
            province="VR",
            ateco_code="1071",
            limit=5,
            max_records=6,
        )
    )

    assert len(results) == 6


def test_search_companies_raises_on_api_error():
    error_payload = {
        "success": False,
        "message": "Invalid token",
    }
    client = make_client(error_payload)

    with pytest.raises(OpenAPIError):
        list(client.search_companies(province="VR", ateco_code="1071"))


def test_dry_run_count_returns_total():
    payload = {
        "success": True,
        "data": [],
        "metadata": {"total": 42},
    }
    client = make_client(payload)

    assert client.dry_run_count(province="VR", ateco_code="1071") == 42

from __future__ import annotations

import json
import time
from datetime import date
from pathlib import Path

import requests

BASE_URL = "https://app.pennylane.com/api/external/v2"
CACHE_DIR = Path(__file__).resolve().parent.parent / "data" / "cache"


class PennylaneClient:
    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers.update(
            {"accept": "application/json", "authorization": f"Bearer {token}"}
        )
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        self._supplier_cache = self._load_cache("suppliers.json")
        self._customer_cache = self._load_cache("customers.json")

    def _load_cache(self, filename: str) -> dict:
        path = CACHE_DIR / filename
        if path.exists():
            return json.loads(path.read_text())
        return {}

    def _save_cache(self, filename: str, cache: dict) -> None:
        (CACHE_DIR / filename).write_text(json.dumps(cache, ensure_ascii=False, indent=2))

    def _get(self, path: str, params: dict | None = None) -> dict:
        resp = self.session.get(f"{BASE_URL}{path}", params=params)
        if resp.status_code == 429:
            time.sleep(2)
            resp = self.session.get(f"{BASE_URL}{path}", params=params)
        resp.raise_for_status()
        return resp.json()

    def iter_transactions_since(self, since: date):
        """Yields transactions in reverse chronological order, stopping once
        a transaction older than `since` is reached (transactions endpoint is
        sorted by -id which tracks creation order / date)."""
        cursor = None
        while True:
            params = {"sort": "-id"}
            if cursor:
                params["cursor"] = cursor
            data = self._get("/transactions", params=params)
            for item in data["items"]:
                tx_date = date.fromisoformat(item["date"])
                if tx_date < since:
                    return
                yield item
            if not data.get("has_more"):
                return
            cursor = data["next_cursor"]

    def iter_new_transactions(self, last_synced_id: int | None):
        """Yields transactions more recent than `last_synced_id` (most recent
        first). Pass None to fetch everything currently available."""
        cursor = None
        while True:
            params = {"sort": "-id"}
            if cursor:
                params["cursor"] = cursor
            data = self._get("/transactions", params=params)
            for item in data["items"]:
                if last_synced_id is not None and item["id"] <= last_synced_id:
                    return
                yield item
            if not data.get("has_more"):
                return
            cursor = data["next_cursor"]

    def get_supplier_name(self, supplier_id: int) -> str:
        key = str(supplier_id)
        if key not in self._supplier_cache:
            data = self._get(f"/suppliers/{supplier_id}")
            self._supplier_cache[key] = data.get("name", "")
            self._save_cache("suppliers.json", self._supplier_cache)
        return self._supplier_cache[key]

    def get_customer_name(self, customer_id: int) -> str:
        key = str(customer_id)
        if key not in self._customer_cache:
            data = self._get(f"/customers/{customer_id}")
            name = data.get("name") or " ".join(
                filter(None, [data.get("first_name"), data.get("last_name")])
            )
            self._customer_cache[key] = name
            self._save_cache("customers.json", self._customer_cache)
        return self._customer_cache[key]

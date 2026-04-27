from __future__ import annotations


class LookupService:
    def load_lookup_data(self, api) -> None:
        return api._load_lookup_data_legacy()

    def lookup_zip_rainfall(self, api, zip_code: str) -> dict:
        return api.lookup_zip_rainfall_legacy(zip_code)

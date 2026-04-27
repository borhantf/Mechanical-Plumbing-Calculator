from __future__ import annotations


class WsfuService:
    def load_wsfu_data(self, api) -> None:
        return api._load_wsfu_data_legacy()

    def calculate_wsfu(self, api, payload) -> dict:
        return api.calculate_wsfu_legacy(payload)

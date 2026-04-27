from __future__ import annotations


class SolarService:
    def load_solar_zip_data(self, api) -> None:
        return api._load_solar_zip_data_legacy()

    def get_solar_catalog(self, api) -> dict:
        return api.get_solar_catalog_legacy()

    def lookup_solar_climate(self, api, zip_code: str) -> dict:
        return api.lookup_solar_climate_legacy(zip_code)

    def calculate_solar(self, api, payload: dict) -> dict:
        return api.calculate_solar_legacy(payload)

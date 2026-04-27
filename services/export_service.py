from __future__ import annotations


class ExportService:
    def export_per_area(self, api, snapshot: dict) -> dict:
        return api.export_per_area_legacy(snapshot)

    def export_section(self, api, snapshot: dict) -> dict:
        return api.export_section_legacy(snapshot)

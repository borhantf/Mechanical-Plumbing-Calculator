from __future__ import annotations

from datetime import datetime


class ProjectService:
    """Project payload normalization orchestrator with compatibility wrappers."""

    def normalize_settings_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("settings", {}))

    def normalize_storm_payload(self, normalized: dict) -> dict:
        return {
            "areas": list(normalized.get("areas", [])),
            "settings": dict(normalized.get("settings", {})),
        }

    def normalize_condensate_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("condensate", {}))

    def normalize_vent_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("vent", {}))

    def normalize_gas_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("gas", {}))

    def normalize_ventilation_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("ventilation", {}))

    def normalize_exhaust_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("exhaust", {}))

    def normalize_wsfu_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("wsfu", {}))

    def normalize_duct_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("duct", {}))

    def normalize_duct_static_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("ductStatic", {}))

    def normalize_solar_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("solar", {}))

    def normalize_refrigerant_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("refrigerant", {}))

    def normalize_project_context_payload(self, normalized: dict) -> dict:
        return dict(normalized.get("project", {}))

    def normalize_project_payload(self, payload: dict, legacy_normalizer, schema_version: int) -> dict:
        """
        Keep legacy behavior exactly, but orchestrate via section-specific functions
        for future maintainability.
        """
        normalized = legacy_normalizer(payload)
        return {
            "schema_version": normalized.get("schema_version", schema_version),
            "saved_at": normalized.get("saved_at", datetime.now().isoformat(timespec="seconds")),
            "activeModule": normalized.get("activeModule", "home"),
            "project": self.normalize_project_context_payload(normalized),
            "settings": self.normalize_settings_payload(normalized),
            "areas": self.normalize_storm_payload(normalized).get("areas", []),
            "condensate": self.normalize_condensate_payload(normalized),
            "vent": self.normalize_vent_payload(normalized),
            "gas": self.normalize_gas_payload(normalized),
            "ventilation": self.normalize_ventilation_payload(normalized),
            "exhaust": self.normalize_exhaust_payload(normalized),
            "wsfu": self.normalize_wsfu_payload(normalized),
            "solar": self.normalize_solar_payload(normalized),
            "refrigerant": self.normalize_refrigerant_payload(normalized),
            "duct": self.normalize_duct_payload(normalized),
            "ductStatic": self.normalize_duct_static_payload(normalized),
        }

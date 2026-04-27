from __future__ import annotations


class GasService:
    def upload_gas_table_pdfs(self, api) -> dict:
        return api.upload_gas_table_pdfs_legacy()

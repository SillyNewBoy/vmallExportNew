"""Shared JSON export helpers."""

from __future__ import annotations

import json
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any


class JsonCategoryExportMixin:
    """Shared file writing, pull-time, and new-product diff helpers."""

    logger: logging.Logger

    def _current_pull_time(self) -> str:
        return datetime.now().strftime("%y%m%d")

    def _write_category_payload(
        self,
        payload: dict[str, Any],
        *,
        output_dir: str | Path,
        diff_dir: str | Path,
    ) -> Path:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        category_key = str(payload.get("category_key") or "category")
        file_path = output_path / f"{category_key}.json"
        previous_payload = self._read_payload(file_path)
        diff_path = self._write_new_products_diff(
            previous_payload,
            payload,
            diff_dir=diff_dir,
        )
        if diff_path:
            self.logger.info(
                "Wrote new-products diff: category=%s file=%s",
                category_key,
                diff_path,
            )
        file_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return file_path

    def _read_payload(self, file_path: Path) -> dict[str, Any] | None:
        if not file_path.exists():
            return None
        try:
            data = json.loads(file_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            self.logger.warning("Unable to read previous export file %s: %s", file_path, exc)
            return None
        return data if isinstance(data, dict) else None

    def _write_new_products_diff(
        self,
        previous_payload: dict[str, Any] | None,
        current_payload: dict[str, Any],
        *,
        diff_dir: str | Path,
    ) -> Path | None:
        if not previous_payload:
            self.logger.info("No previous export found, skipping new-products diff")
            return None
        if previous_payload.get("category_key") != current_payload.get("category_key"):
            self.logger.warning(
                "Previous export category does not match current category, skipping diff: previous=%s current=%s",
                previous_payload.get("category_key"),
                current_payload.get("category_key"),
            )
            return None

        diff_payload = self._build_new_products_payload(previous_payload, current_payload)
        if diff_payload["new_product_count"] == 0:
            self.logger.info(
                "No new products found for category=%s",
                current_payload.get("category_key"),
            )
            return None

        output_dir = Path(diff_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        category_key = self._safe_file_part(str(current_payload.get("category_key") or "category"))
        previous_pull_time = self._safe_file_part(str(diff_payload["previous_pull_time"] or "unknown"))
        current_pull_time = self._safe_file_part(str(diff_payload["current_pull_time"] or "unknown"))
        file_path = output_dir / f"{category_key}_{previous_pull_time}_{current_pull_time}_new.json"
        file_path.write_text(
            json.dumps(diff_payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return file_path

    def _build_new_products_payload(
        self,
        previous_payload: dict[str, Any],
        current_payload: dict[str, Any],
    ) -> dict[str, Any]:
        previous_ids = self._collect_product_ids(previous_payload)
        diff_payload: dict[str, Any] = {
            "category_key": current_payload.get("category_key"),
            "category_name": current_payload.get("category_name"),
            "previous_pull_time": previous_payload.get("pull_time"),
            "current_pull_time": current_payload.get("pull_time"),
            "new_product_count": 0,
            "series": {},
        }

        for series_key, series in (current_payload.get("series") or {}).items():
            products = series.get("products") or {}
            new_products = {
                product_key: product
                for product_key, product in products.items()
                if self._product_identity(product_key, product) not in previous_ids
            }
            if not new_products:
                continue
            diff_payload["series"][series_key] = {
                "series_key": series.get("series_key"),
                "series_name": series.get("series_name"),
                "products": new_products,
            }
            diff_payload["new_product_count"] += len(new_products)

        return diff_payload

    def _collect_product_ids(self, payload: dict[str, Any]) -> set[str]:
        ids: set[str] = set()
        for series in (payload.get("series") or {}).values():
            for product_key, product in (series.get("products") or {}).items():
                ids.add(self._product_identity(product_key, product))
        return ids

    def _product_identity(self, product_key: str, product: dict[str, Any]) -> str:
        return str(product.get("product_id") or product.get("sku_code") or product_key)

    def _safe_file_part(self, value: str) -> str:
        return re.sub(r"[^0-9a-zA-Z_\-.]+", "_", value).strip("._") or "unknown"

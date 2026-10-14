"""Category-based export workflow."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

from .categories import CATEGORY_PRESETS
from .client import VmallClient
from .json_export import JsonCategoryExportMixin
from .models import CategoryConfig, SearchProduct, SearchQueryConfig


class VmallExporter(JsonCategoryExportMixin):
    """High-level export orchestration."""

    COLOR_SUMMARY_ATTR_NAMES = ("颜色", "款式")

    def __init__(self, client: VmallClient | None = None) -> None:
        self.client = client or VmallClient()
        self.logger = logging.getLogger(__name__)

    def list_categories(self) -> list[CategoryConfig]:
        return [
            CategoryConfig(
                key=key,
                name=value["name"],
                keyword=value["keyword"],
                category_id=value["category_id"],
                search_queries=value.get("search_queries", ()),
            )
            for key, value in CATEGORY_PRESETS.items()
        ]

    def get_category(self, key: str) -> CategoryConfig:
        try:
            value = CATEGORY_PRESETS[key]
        except KeyError as exc:
            raise KeyError(f"Unknown category: {key}") from exc
        return CategoryConfig(
            key=key,
            name=value["name"],
            keyword=value["keyword"],
            category_id=value["category_id"],
            search_queries=value.get("search_queries", ()),
        )

    def export_category(
        self,
        category: CategoryConfig,
        *,
        output_dir: str | Path,
        diff_dir: str | Path = "diff",
        limit: int | None = None,
        max_pages: int = 10,
        page_size: int = 20,
    ) -> tuple[Path, int]:
        self.logger.info("Starting export: category=%s query_count=%s", category.key, len(category.queries))
        products: list[SearchProduct] = []
        for query in category.queries:
            products.extend(
                self._search_products_for_query(
                    query,
                    page_size=page_size,
                    max_pages=max_pages,
                )
            )
        self.logger.info("Search returned %s raw products", len(products))
        selected_products = self._dedupe_products(products, limit=limit)
        self.logger.info("Selected %s unique products after dedupe", len(selected_products))
        items = [
            self._build_export_item(category, product)
            for product in selected_products
        ]
        payload = self._build_category_tree(
            category,
            items,
            pull_time=self._current_pull_time(),
        )

        file_path = self._write_category_payload(
            payload,
            output_dir=output_dir,
            diff_dir=diff_dir,
        )
        self.logger.info("Finished export: category=%s file=%s", category.key, file_path)
        return file_path, self._count_products(payload)

    def _search_products_for_query(
        self,
        query: SearchQueryConfig,
        *,
        page_size: int,
        max_pages: int,
    ) -> list[SearchProduct]:
        self.logger.info(
            "Searching query: keyword=%s category_id=%s allowed_category_codes=%s",
            query.keyword,
            query.category_id,
            ",".join(query.allowed_category_codes) or "*",
        )
        products = self.client.search_products(
            keyword=query.keyword,
            category_id=query.category_id,
            page_size=page_size,
            max_pages=max_pages,
        )
        if not query.allowed_category_codes:
            return products
        allowed_codes = set(query.allowed_category_codes)
        filtered = [
            product for product in products if (product.category_code or "") in allowed_codes
        ]
        self.logger.info(
            "Filtered query products from %s to %s by third_level_category_code",
            len(products),
            len(filtered),
        )
        return filtered

    def _dedupe_products(
        self,
        products: list[SearchProduct],
        *,
        limit: int | None,
    ) -> list[SearchProduct]:
        deduped: list[SearchProduct] = []
        seen: set[str] = set()

        for product in products:
            key = self._normalize_name(product.product_name)
            if key in seen:
                continue
            seen.add(key)
            deduped.append(product)
            if limit is not None and len(deduped) >= limit:
                break

        return deduped

    def _normalize_name(self, name: str) -> str:
        return " ".join(name.split()).strip().lower()

    def _build_export_item(
        self,
        category: CategoryConfig,
        product: SearchProduct,
    ) -> dict[str, Any]:
        self.logger.info(
            "Building export item: category=%s product_id=%s product_name=%s",
            category.key,
            product.product_id,
            product.product_name,
        )
        next_data = self.client.fetch_product_next_data(product.product_id)
        main_data = next_data.get("props", {}).get("pageProps", {}).get("mainData", {}).get("current", {})
        specs = self._extract_specifications(next_data, preferred_sku_code=product.sku_code)
        color_summary = self._extract_color_summary(main_data)
        series_name = self._extract_series_name(specs)
        product_name_short = self._extract_product_short_name(main_data, specs, fallback=product.product_name)
        return {
            "category_key": category.key,
            "category_name": category.name,
            "series_name": series_name,
            "series_key": self._normalize_hierarchy_key(series_name),
            "product_id": product.product_id,
            "product_name": product.product_name,
            "product_name_short": product_name_short,
            "product_key": self._normalize_hierarchy_key(product_name_short),
            "sku_code": product.sku_code,
            "sku_name": product.sku_name,
            "product_url": product.product_url,
            "third_level_category_code": product.category_code,
            "color_summary": color_summary,
            "specifications": specs,
        }

    def _build_category_tree(
        self,
        category: CategoryConfig,
        items: list[dict[str, Any]],
        *,
        pull_time: str,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "category_key": category.key,
            "category_name": category.name,
            "pull_time": pull_time,
            "series": {},
        }

        for item in items:
            series_bucket = payload["series"].setdefault(
                item["series_key"],
                {
                    "series_key": item["series_key"],
                    "series_name": item["series_name"],
                    "products": {},
                },
            )
            product_key = self._ensure_unique_product_key(
                series_bucket["products"],
                item["product_key"],
                item["product_id"],
            )
            color_summary = item.get("color_summary", [])
            series_bucket["products"][product_key] = {
                "product_key": product_key,
                "product_id": item["product_id"],
                "product_name": item["product_name"],
                "product_name_short": item["product_name_short"],
                "sku_code": item["sku_code"],
                "sku_name": item["sku_name"],
                "product_url": item["product_url"],
                "third_level_category_code": item["third_level_category_code"],
                "color_summary_count": len(color_summary),
                "color_summary": color_summary,
                "specifications": item["specifications"],
            }

        return payload

    def _count_products(self, payload: dict[str, Any]) -> int:
        series = payload.get("series", {})
        return sum(len(item.get("products", {})) for item in series.values())

    def _extract_color_summary(
        self,
        main_data: dict[str, Any],
    ) -> list[dict[str, Any]]:
        summary: dict[tuple[str, str], dict[str, Any]] = {}

        for item in self._as_list(main_data.get("colorSelectItem")):
            if not isinstance(item, dict):
                continue
            self._add_color_summary_item(
                summary,
                attr_name=item.get("attrName") or "颜色",
                attr_value=item.get("attrValue"),
                sku_code=item.get("sbomCode"),
                sku_codes=item.get("surrportSbomCodes") or item.get("supportSbomCodes"),
                image_url=item.get("url"),
                source="colorSelectItem",
            )

        product_options = self._as_dict(main_data.get("productOptions"))
        mappings = self._as_dict(product_options.get("gbomAttrMappings"))
        for attr_name, options in mappings.items():
            attr_name_text = str(attr_name or "")
            if not self._is_color_summary_attr(attr_name_text):
                continue
            for option in self._as_list(options):
                if not isinstance(option, dict):
                    continue
                self._add_color_summary_item(
                    summary,
                    attr_name=option.get("attrName") or attr_name_text,
                    attr_value=option.get("attrValue"),
                    sku_code=option.get("sbomCode"),
                    sku_codes=option.get("supportSbomCodes") or option.get("surrportSbomCodes"),
                    source="productOptions.gbomAttrMappings",
                )

        for sbom_map in (
            self._as_dict(main_data.get("base")),
            self._as_dict(product_options.get("sbomList")),
        ):
            for sbom_code, sbom in sbom_map.items():
                if not isinstance(sbom, dict):
                    continue
                for attr in self._as_list(sbom.get("gbomAttrList")):
                    if not isinstance(attr, dict):
                        continue
                    attr_name = str(attr.get("attrName") or "")
                    if not self._is_color_summary_attr(attr_name):
                        continue
                    self._add_color_summary_item(
                        summary,
                        attr_name=attr_name,
                        attr_value=attr.get("attrValue"),
                        sku_code=attr.get("sbomCode") or sbom.get("sbomCode") or sbom_code,
                        source="sbom.gbomAttrList",
                    )

        return list(summary.values())

    def _add_color_summary_item(
        self,
        summary: dict[tuple[str, str], dict[str, Any]],
        *,
        attr_name: Any,
        attr_value: Any,
        sku_code: Any = None,
        sku_codes: Any = None,
        image_url: Any = None,
        source: str = "",
    ) -> None:
        value = str(attr_value or "").strip()
        if not value:
            return
        name = str(attr_name or "颜色").strip() or "颜色"
        key = (name, value)
        if key not in summary:
            summary[key] = {
                "attr_name": name,
                "attr_value": value,
                "sku_codes": [],
            }
            if image_url:
                summary[key]["image_url"] = str(image_url)
            if source:
                summary[key]["source"] = source

        entry = summary[key]
        for code in self._as_list(sku_codes):
            self._append_unique_sku_code(entry["sku_codes"], code)
        self._append_unique_sku_code(entry["sku_codes"], sku_code)
        if image_url and "image_url" not in entry:
            entry["image_url"] = str(image_url)

    def _append_unique_sku_code(self, sku_codes: list[str], value: Any) -> None:
        code = str(value or "").strip()
        if code and code not in sku_codes:
            sku_codes.append(code)

    def _is_color_summary_attr(self, attr_name: str) -> bool:
        return any(name in attr_name for name in self.COLOR_SUMMARY_ATTR_NAMES)

    def _as_list(self, value: Any) -> list[Any]:
        return value if isinstance(value, list) else []

    def _as_dict(self, value: Any) -> dict[str, Any]:
        return value if isinstance(value, dict) else {}

    def _extract_specifications(
        self,
        next_data: dict[str, Any],
        *,
        preferred_sku_code: str,
    ) -> list[dict[str, Any]]:
        page_props = next_data.get("props", {}).get("pageProps", {})
        main_data = page_props.get("mainData", {}).get("current", {})
        base = main_data.get("base", {})
        blobs = (
            page_props.get("extData", {})
            .get("productRelDetailResp", {})
            .get("querySkuPicDetailResp", {})
            .get("sbomBlobsList", [])
        )
        blob = self._select_blob(blobs, preferred_sku_code=preferred_sku_code)
        if not blob:
            raise RuntimeError(f"No specification blob found for sku {preferred_sku_code}.")

        sbom_code = str(blob.get("sbomCode") or preferred_sku_code)
        product_base = base.get(sbom_code, {})
        spec_items = blob.get("specificationList", [])
        parent_names = {
            str(item.get("attrCode")): str(item.get("attrName"))
            for item in spec_items
            if item.get("attrCode") and item.get("attrName")
        }
        return [
            {
                "group_name": parent_names.get(str(item.get("parentAttrCode")), ""),
                "attr_code": item.get("attrCode"),
                "attr_name": item.get("attrName"),
                "attr_value": item.get("attrValue"),
                "attr_remark": item.get("attrRemark"),
                "parent_attr_code": item.get("parentAttrCode"),
                "sbom_code": sbom_code,
                "sbom_name": blob.get("sbomName") or product_base.get("sbomName"),
            }
            for item in spec_items
            if item.get("attrName") and not self._is_group_header(item)
        ]

    def _is_group_header(self, item: dict[str, Any]) -> bool:
        parent_attr_code = str(item.get("parentAttrCode") or "")
        attr_value = item.get("attrValue")
        return parent_attr_code in {"", "0"} and not attr_value

    def _select_blob(
        self,
        blobs: list[dict[str, Any]],
        *,
        preferred_sku_code: str,
    ) -> dict[str, Any] | None:
        for blob in blobs:
            if str(blob.get("sbomCode")) == preferred_sku_code:
                return blob
        return blobs[0] if blobs else None

    def _extract_series_name(self, specs: list[dict[str, Any]]) -> str:
        for item in specs:
            if item.get("attr_name") == "系列" and item.get("attr_value"):
                return str(item["attr_value"]).strip()
        return "未分类系列"

    def _extract_product_short_name(
        self,
        main_data: dict[str, Any],
        specs: list[dict[str, Any]],
        *,
        fallback: str,
    ) -> str:
        brief_name = str(main_data.get("briefName") or "").strip()
        if brief_name:
            return brief_name

        for item in specs:
            if item.get("attr_name") == "传播名" and item.get("attr_value"):
                value = str(item["attr_value"]).strip()
                return self._strip_brand_prefix(value)

        return self._strip_brand_prefix(fallback)

    def _strip_brand_prefix(self, value: str) -> str:
        value = value.strip()
        for prefix in ("HUAWEI ", "Huawei ", "华为 "):
            if value.startswith(prefix):
                return value[len(prefix) :].strip()
        return value

    def _normalize_hierarchy_key(self, value: str) -> str:
        cleaned = self._strip_brand_prefix(value).strip().lower()
        cleaned = re.sub(r"[()（）]+", "", cleaned)
        cleaned = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", cleaned)
        return cleaned or "unknown"

    def _ensure_unique_product_key(
        self,
        products: dict[str, Any],
        desired_key: str,
        product_id: str,
    ) -> str:
        if desired_key not in products:
            return desired_key

        suffix = 2
        while f"{desired_key}_{suffix}" in products:
            suffix += 1
        return f"{desired_key}_{suffix}"

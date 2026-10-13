"""Export full product specifications from the HUAWEI consumer site.

Target site example: https://consumer.huawei.com/my/

The consumer site is fully server-rendered:

- Category listing pages (``/my/phones/``) link to product pages
  (``/my/phones/mate-x7/``).
- Product specification pages live at ``/my/phones/mate-x7/specs/`` and carry
  the full parameter list in a static ``large-accordion`` HTML structure.

Export flow per category:

1. Fetch the category listing page and extract product paths.
2. Supplement the path list with ``specs/`` URLs from the region sitemap.
3. Fetch each product's ``specs/`` page and parse every parameter group.
4. Save the result as the shared category JSON tree (reusing the diff logic
   from :class:`JsonCategoryExportMixin`) and optionally as Excel/CSV.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup, Tag

from .consumer_categories import (
    CONSUMER_CATEGORY_PRESETS,
    CONSUMER_DEFAULT_SERIES_NAME,
    CONSUMER_SERIES_RULES,
)
from .consumer_client import ConsumerClient, ConsumerClientError, ConsumerSitemap
from .json_export import JsonCategoryExportMixin


class ConsumerHuaweiExporter(JsonCategoryExportMixin):
    """Export specifications from consumer.huawei.com region sites."""

    BRAND_PREFIXES = ("HUAWEI ", "Huawei ", "huawei ")

    # noise words used in <title> tags across locales
    TITLE_PREFIX_PATTERN = re.compile(
        r"^(especificaciones\s+(de\s+)?|specifications\s+(of\s+)?|specs\s+)",
        re.IGNORECASE,
    )
    TITLE_SUFFIX_PATTERN = re.compile(
        r"\s+(especificaciones|specifications|specs)$",
        re.IGNORECASE,
    )

    def _clean_title_name(self, text: str) -> str:
        """Turn a page <title> into a plain product name.

        Handles e.g. "HUAWEI Mate X7 Specifications - HUAWEI Malaysia" and
        "Especificaciones de HUAWEI MateBook D 14 Intel - HUAWEI México".
        """
        text = re.split(r"\s[-–|]\s", text)[0].strip()
        text = self.TITLE_PREFIX_PATTERN.sub("", text)
        text = self.TITLE_SUFFIX_PATTERN.sub("", text)
        return text.strip()

    def __init__(
        self,
        client: ConsumerClient | None = None,
        *,
        region: str = "my",
    ) -> None:
        self.client = client or ConsumerClient(region=region)
        self.sitemap = ConsumerSitemap(self.client)
        self.logger = logging.getLogger(__name__)

    # ------------------------------------------------------------------
    # category configuration
    # ------------------------------------------------------------------
    def list_categories(self) -> list[dict[str, str]]:
        return [
            {"key": key, "name": value["name"], "path": value["path"]}
            for key, value in CONSUMER_CATEGORY_PRESETS.items()
        ]

    def get_category(self, key: str) -> dict[str, str]:
        try:
            value = CONSUMER_CATEGORY_PRESETS[key]
        except KeyError as exc:
            raise KeyError(f"Unknown consumer category: {key}") from exc
        return {"key": key, "name": value["name"], "path": value["path"]}

    # ------------------------------------------------------------------
    # product discovery
    # ------------------------------------------------------------------
    def discover_product_paths(self, category: dict[str, str]) -> list[str]:
        """Return deduplicated site-relative product paths for a category."""
        path = category["path"]
        found: dict[str, None] = {}

        pattern = re.compile(
            rf"/{re.escape(self.client.region)}/{re.escape(path)}/([a-zA-Z0-9][a-zA-Z0-9_-]*)/"
        )
        try:
            html = self.client.fetch_page(f"{path}/")
        except ConsumerClientError as exc:
            self.logger.warning("Category listing unavailable: %s (%s)", path, exc)
            html = ""
        for match in pattern.finditer(html):
            slug = match.group(1).lower()
            if self._is_non_product_slug(slug):
                continue
            found[f"{path}/{slug}/"] = None
        self.logger.info(
            "Listing page discovery: category=%s products=%s", path, len(found)
        )

        for product_path in self.sitemap.fetch_specs_urls({path}):
            slug = product_path.strip("/").split("/")[-1].lower()
            if self._is_non_product_slug(slug):
                continue
            found.setdefault(product_path, None)
        self.logger.info(
            "Discovery finished: category=%s total_candidates=%s", path, len(found)
        )
        return list(found)

    def _is_non_product_slug(self, slug: str) -> bool:
        blocked = {
            "specs",
            "buy",
            "support",
            "campaign",
            "compare",
            "gallery",
            "review",
            "reviews",
        }
        return slug in blocked

    # ------------------------------------------------------------------
    # spec page parsing
    # ------------------------------------------------------------------
    def parse_specs_page(self, html: str, *, product_path: str) -> dict[str, Any]:
        """Parse a ``specs/`` page into product metadata and specifications."""
        soup = BeautifulSoup(html, "html.parser")

        product_name = self._extract_product_name(soup, product_path)
        colors = self._extract_colors(soup)

        specifications: list[dict[str, Any]] = []
        for item in soup.select("li.large-accordion__item"):
            group_name = self._tag_text(item.select_one("span.large-accordion__title"))
            content = item.select_one("div.large-accordion__content")
            if content is None:
                continue
            specifications.extend(self._parse_multi_variant_blocks(content, group_name))
            for inner in content.select("div.large-accordion__inner"):
                wraps = inner.select("div.large-accordion__wrap")
                if wraps:
                    # variant layout: one inner block holds several
                    # subtitle/value pairs wrapped in __wrap containers
                    for wrap in wraps:
                        attr_name = self._tag_text(
                            wrap.select_one(".large-accordion-subtitle")
                        )
                        if not attr_name:
                            continue
                        attr_value, attr_remark = self._extract_attr_value(wrap)
                        specifications.append(
                            {
                                "group_name": group_name,
                                "attr_name": attr_name,
                                "attr_value": attr_value,
                                "attr_remark": attr_remark,
                            }
                        )
                    continue
                attr_name = self._tag_text(inner.select_one(".large-accordion-subtitle"))
                if not attr_name:
                    # group-level footnote, e.g. <p class="large-accordion-text">
                    note = inner.select_one("p.large-accordion-text")
                    note_text = self._tag_text(note)
                    if note_text:
                        specifications.append(
                            {
                                "group_name": group_name,
                                "attr_name": "备注",
                                "attr_value": note_text,
                                "attr_remark": "",
                            }
                        )
                    continue
                attr_value, attr_remark = self._extract_attr_value(inner)
                specifications.append(
                    {
                        "group_name": group_name,
                        "attr_name": attr_name,
                        "attr_value": attr_value,
                        "attr_remark": attr_remark,
                    }
                )

        return {
            "product_name": product_name,
            "color_summary": colors,
            "specifications": specifications,
        }

    # class name patterns used by multi-variant (e.g. watch sizes) spec blocks
    MULTI_ITEM_PATTERN = re.compile(r"^multi-[\w-]+-item$")
    MULTI_TITLE_PATTERN = re.compile(r"^multi-[\w-]+-title$")
    MULTI_LIST_PATTERN = re.compile(r"^multi-[\w-]+-list$")
    MULTI_SUBTITLE_PATTERN = re.compile(r"^multi-[\w-]+-subtitle$")
    MULTI_TEXT_PATTERN = re.compile(r"^multi-[\w-]+-text$")
    MULTI_SUBTEXT_PATTERN = re.compile(r"^multi-[\w-]+-subtext$")

    def _parse_multi_variant_blocks(
        self,
        content: Tag,
        group_name: str,
    ) -> list[dict[str, Any]]:
        """Parse multi-variant spec blocks (``multi-{product}-*`` markup).

        Products with several variants (e.g. a watch in 46 mm and 41 mm) use
        per-product custom classes instead of ``large-accordion__inner``.
        Each variant item becomes one row per attribute, with the variant name
        appended to the attribute name.
        """
        results: list[dict[str, Any]] = []
        for variant_block in content.find_all(
            "div", class_=self.MULTI_ITEM_PATTERN
        ):
            title_tag = variant_block.find("div", class_=self.MULTI_TITLE_PATTERN)
            variant_name = self._tag_text(title_tag)
            for entry in variant_block.find_all("div", class_=self.MULTI_LIST_PATTERN):
                attr_name = self._tag_text(
                    entry.find("div", class_=self.MULTI_SUBTITLE_PATTERN)
                )
                value_tag = entry.find("div", class_=self.MULTI_TEXT_PATTERN)
                attr_value = (
                    value_tag.get_text("\n", strip=True) if value_tag else ""
                )
                remark_tag = entry.find(class_=self.MULTI_SUBTEXT_PATTERN)
                attr_remark = self._tag_text(remark_tag)
                if not attr_name and not attr_value:
                    continue
                if not attr_name:
                    attr_name = group_name
                if variant_name:
                    attr_name = f"{attr_name} ({variant_name})"
                results.append(
                    {
                        "group_name": group_name,
                        "attr_name": attr_name,
                        "attr_value": attr_value,
                        "attr_remark": attr_remark,
                    }
                )
        return results

    def _extract_product_name(self, soup: BeautifulSoup, product_path: str) -> str:
        banner = soup.select_one(".specs-banner-cnt .large-accordion-title")
        name = self._tag_text(banner)
        if name:
            return name
        title = soup.find("title")
        if title and title.get_text(strip=True):
            cleaned = self._clean_title_name(title.get_text(strip=True))
            if cleaned:
                return cleaned
        slug = product_path.strip("/").split("/")[-1]
        return slug.replace("-", " ").title()

    def _extract_colors(self, soup: BeautifulSoup) -> list[dict[str, Any]]:
        node = soup.select_one(".specs-banner-cnt .color-text")
        text = self._tag_text(node)
        colors: list[dict[str, Any]] = []
        if not text:
            return colors
        for color in re.split(r"[,;/，；]", text):
            color = color.strip()
            if color:
                colors.append(
                    {"attr_name": "颜色", "attr_value": color, "sku_codes": []}
                )
        return colors

    def _extract_attr_value(self, inner: Tag) -> tuple[str, str]:
        value_parts: list[str] = []
        remark_parts: list[str] = []
        for child in inner.find_all(recursive=False):
            classes = child.get("class") or []
            text = child.get_text("\n", strip=True) if child.name else ""
            if not text:
                continue
            if "large-accordion-subtext" in classes:
                remark_parts.append(text)
            elif "large-accordion-subtitle" in classes:
                continue
            else:
                value_parts.append(text)
        # collapse the newlines produced for <br> separators
        value = "\n".join(value_parts).strip()
        remark = "\n".join(remark_parts).strip()
        return value, remark

    def _tag_text(self, tag: Tag | None) -> str:
        if tag is None:
            return ""
        return " ".join(tag.get_text(" ", strip=True).split())

    # ------------------------------------------------------------------
    # export orchestration
    # ------------------------------------------------------------------
    def export_category(
        self,
        category: dict[str, str],
        *,
        output_dir: str | Path,
        diff_dir: str | Path = "diff",
        limit: int | None = None,
    ) -> tuple[Path, int, dict[str, Any]]:
        self.logger.info(
            "Starting consumer export: category=%s path=%s region=%s",
            category["key"],
            category["path"],
            self.client.region,
        )
        product_paths = self.discover_product_paths(category)
        if limit is not None:
            product_paths = product_paths[:limit]

        items: list[dict[str, Any]] = []
        failures: list[str] = []
        for product_path in product_paths:
            item = self._build_export_item(category, product_path, failures)
            if item is not None:
                items.append(item)

        payload = self._build_category_tree(
            category,
            items,
            pull_time=self._current_pull_time(),
        )
        if failures:
            payload["failed_products"] = failures
            self.logger.warning(
                "Category %s finished with %s failed products: %s",
                category["key"],
                len(failures),
                ", ".join(failures),
            )

        file_path = self._write_category_payload(
            payload,
            output_dir=output_dir,
            diff_dir=diff_dir,
        )
        count = self._count_products(payload)
        self.logger.info(
            "Finished consumer export: category=%s products=%s file=%s",
            category["key"],
            count,
            file_path,
        )
        return file_path, count, payload

    def _build_export_item(
        self,
        category: dict[str, str],
        product_path: str,
        failures: list[str],
    ) -> dict[str, Any] | None:
        specs_path = f"{product_path}specs/"
        self.logger.info("Fetching specs page: %s", specs_path)
        html: str | None = None
        try:
            final_url, html = self.client.fetch_page_with_final_url(specs_path)
            expected_suffix = f"{product_path}specs/"
            if not final_url.rstrip("/").endswith(expected_suffix.rstrip("/")):
                # e.g. the specs page redirects to the homepage for products
                # that do not publish specifications on this region site
                self.logger.info(
                    "Specs page redirected away, product kept without specs: %s -> %s",
                    specs_path,
                    final_url,
                )
                html = None
        except ConsumerClientError as exc:
            # products without a specs page (campaign pages etc.) are skipped
            self.logger.info("No specs page for %s: %s", product_path, exc)
            return None

        if html is None:
            return self._build_item_without_specs(category, product_path)

        try:
            parsed = self.parse_specs_page(html, product_path=product_path)
        except Exception as exc:  # noqa: BLE001 - keep the batch running
            self.logger.warning("Failed to parse %s: %s", specs_path, exc)
            failures.append(product_path)
            return None

        product_name = parsed["product_name"]
        product_name_short = self._strip_brand_prefix(product_name)
        series_name = self._derive_series_name(product_name)
        return {
            "category_key": category["key"],
            "category_name": category["name"],
            "series_name": series_name,
            "series_key": self._normalize_hierarchy_key(series_name),
            "product_id": product_path.strip("/").split("/")[-1],
            "product_name": product_name,
            "product_name_short": product_name_short,
            "product_key": self._normalize_hierarchy_key(product_name_short),
            "sku_code": "",
            "sku_name": "",
            "product_url": f"{self.client.base_url}{product_path}",
            "specs_url": f"{self.client.base_url}{specs_path}",
            "third_level_category_code": "",
            "color_summary": parsed["color_summary"],
            "specifications": parsed["specifications"],
        }

    def _build_item_without_specs(
        self,
        category: dict[str, str],
        product_path: str,
    ) -> dict[str, Any] | None:
        """Keep catalog products that do not publish a specs page.

        The product name falls back to the main product page title.
        """
        slug = product_path.strip("/").split("/")[-1]
        product_name = self._name_from_slug(slug)
        try:
            final_url, html = self.client.fetch_page_with_final_url(product_path)
            if final_url.rstrip("/").endswith(product_path.rstrip("/")):
                soup = BeautifulSoup(html, "html.parser")
                title = soup.find("title")
                if title and title.get_text(strip=True):
                    text = self._clean_title_name(title.get_text(strip=True))
                    if text and "malaysia" not in text.lower() and "méxico" not in text.lower():
                        product_name = text
        except ConsumerClientError:
            self.logger.info("Product page unavailable as well: %s", product_path)
            return None

        series_name = self._derive_series_name(product_name)
        product_name_short = self._strip_brand_prefix(product_name)
        return {
            "category_key": category["key"],
            "category_name": category["name"],
            "series_name": series_name,
            "series_key": self._normalize_hierarchy_key(series_name),
            "product_id": slug,
            "product_name": product_name,
            "product_name_short": product_name_short,
            "product_key": self._normalize_hierarchy_key(product_name_short),
            "sku_code": "",
            "sku_name": "",
            "product_url": f"{self.client.base_url}{product_path}",
            "specs_url": "",
            "third_level_category_code": "",
            "color_summary": [],
            "specifications": [],
        }

    def _build_category_tree(
        self,
        category: dict[str, str],
        items: list[dict[str, Any]],
        *,
        pull_time: str,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "category_key": category["key"],
            "category_name": category["name"],
            "source": self.client.base_url,
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
                "specs_url": item["specs_url"],
                "third_level_category_code": item["third_level_category_code"],
                "color_summary_count": len(color_summary),
                "color_summary": color_summary,
                "specifications": item["specifications"],
            }

        return payload

    def _count_products(self, payload: dict[str, Any]) -> int:
        return sum(
            len(series.get("products", {}))
            for series in (payload.get("series") or {}).values()
        )

    # ------------------------------------------------------------------
    # naming helpers
    # ------------------------------------------------------------------
    def _strip_brand_prefix(self, value: str) -> str:
        value = value.strip()
        for prefix in self.BRAND_PREFIXES:
            if value.startswith(prefix):
                return value[len(prefix):].strip()
        return value

    def _name_from_slug(self, slug: str) -> str:
        """Turn a URL slug into a readable product name, e.g. mate-x3 -> Mate X3."""
        words: list[str] = []
        for part in slug.split("-"):
            if part.isalpha() and len(part) <= 2:
                words.append(part.upper())
            elif re.fullmatch(r"[a-z]+\d+[a-z]*", part):
                words.append(part.upper())
            else:
                words.append(part.capitalize())
        return " ".join(words)

    def _derive_series_name(self, product_name: str) -> str:
        lowered = f" {product_name.strip().lower()} "
        for keyword, series_name in CONSUMER_SERIES_RULES:
            if keyword in lowered:
                return series_name
        return CONSUMER_DEFAULT_SERIES_NAME

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

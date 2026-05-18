"""HIMA car configuration export workflow."""

from __future__ import annotations

import hashlib
import logging
import re
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup, NavigableString, Tag

from .json_export import JsonCategoryExportMixin


HIMA_CATEGORY_KEY = "cars"
HIMA_CATEGORY_NAME = "汽车"

BASE_URL = "https://hima.auto/"
BRAND_URLS = (
    "https://hima.auto/wenjie/",
    "https://hima.auto/zhijie/",
    "https://hima.auto/xiangjie/",
    "https://hima.auto/zunjie/",
    "https://hima.auto/shangjie/",
)
BRAND_NAMES = {
    "wenjie": "问界",
    "zhijie": "智界",
    "xiangjie": "享界",
    "zunjie": "尊界",
    "shangjie": "尚界",
}
HEADERS = {"User-Agent": "Mozilla/5.0"}


def normalize_space(text: str) -> str:
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Cf")
    text = text.replace("\xa0", " ")
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    return text.strip()


def extract_text(node: Tag | None, line_joiner: str) -> str:
    if not node:
        return ""
    parts: list[str] = []
    for item in node.descendants:
        if isinstance(item, NavigableString):
            if item.parent and item.parent.name == "sup":
                continue
            parts.append(str(item))
        elif isinstance(item, Tag) and item.name == "br":
            parts.append("\n")
    text = "".join(parts).replace("\r", "")
    lines = [normalize_space(line) for line in text.split("\n")]
    lines = [line for line in lines if line]
    return normalize_space(line_joiner.join(lines))


def text_inline(node: Tag | None) -> str:
    return extract_text(node, " ")


def text_multiline(node: Tag | None) -> str:
    return extract_text(node, " / ")


def slugify(text: str) -> str:
    text = unicodedata.normalize("NFKC", text).lower()
    text = text.replace("｜", " ")
    text = text.replace("+", " plus ")
    text = re.sub(r"[（）()【】\[\]{}<>]", " ", text)
    text = re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "-", text)
    text = re.sub(r"-{2,}", "-", text)
    return text.strip("-") or "item"


def is_404_page(html: str) -> bool:
    return "<title>404" in html or 'pageName":"hima:cn:404"' in html


def first_title_text(soup: BeautifulSoup) -> str:
    if soup.title:
        return normalize_space(soup.title.get_text(" ", strip=True))
    return ""


def infer_brand_name(model_name: str) -> str:
    for prefix in ("问界", "智界", "享界", "尊界", "尚界"):
        if model_name.startswith(prefix):
            return prefix
    return model_name.split(" ")[0]


def find_page_title_block(root: Tag) -> str | None:
    sosuke = root.select_one(":scope > .asking-content.cp-sosuke h2")
    if sosuke:
        text = text_inline(sosuke)
        if text:
            return text
    direct_h2 = root.select_one(":scope > h2")
    if direct_h2:
        text = text_inline(direct_h2)
        if text:
            return text
    return None


def find_tab_labels(root: Tag) -> list[str]:
    for selector in (".car-select-tab > h2", ".m7-tab > h2", ".m9-tab > h2"):
        labels = [text_inline(tag) for tag in root.select(selector)]
        labels = [label for label in labels if label]
        if labels:
            return labels
    return []


def find_current_tab_label(root: Tag) -> str | None:
    for selector in (".car-select-tab > h2.active", ".m7-tab > h2.active", ".m9-tab > h2.active"):
        tag = root.select_one(selector)
        if tag:
            return text_inline(tag)
    for selector in (".car-select-tab > h2", ".m7-tab > h2", ".m9-tab > h2"):
        tags = root.select(selector)
        for tag in tags:
            if not tag.select_one("a"):
                label = text_inline(tag)
                if label:
                    return label
        if tags:
            return text_inline(tags[0])
    return None


def direct_children(root: Tag) -> list[Tag]:
    return [child for child in root.children if isinstance(child, Tag)]


def build_group_label_map(header_scope: Tag | None) -> dict[str, str]:
    if not header_scope:
        return {}
    title_container = header_scope.select_one(".mc-wrapper-1")
    if not title_container:
        return {}

    result: dict[str, str] = {}
    for child in direct_children(title_container):
        label = text_inline(child)
        if not label:
            continue
        for cls in child.get("class", []):
            if cls.endswith("-title"):
                result[cls[: -len("-title")]] = label
    return result


def infer_group_sizes(names: list[str], group_label_map: dict[str, str]) -> list[int]:
    labels = list(group_label_map.values())
    if len(labels) <= 1:
        return []

    sizes: list[int] = []
    idx = 0
    for label in labels:
        count = 0
        while idx < len(names) and label in names[idx]:
            count += 1
            idx += 1
        if count:
            sizes.append(count)
    if sum(sizes) == len(names) and len(sizes) > 1:
        return sizes
    return []


def extract_direct_item_texts(container: Tag | None, multiline: bool = False) -> list[str]:
    if not container:
        return []
    results: list[str] = []
    for child in direct_children(container):
        if child.name == "button":
            continue
        text = text_multiline(child) if multiline else text_inline(child)
        if text:
            results.append(text)
    return results


def extract_named_item_texts(container: Tag | None, group_label_map: dict[str, str]) -> list[str]:
    if not container:
        return []

    items: list[tuple[str, str]] = []
    for child in direct_children(container):
        if child.name == "button":
            continue
        text = text_inline(child)
        if not text:
            continue
        group_label = ""
        for cls in child.get("class", []):
            if cls.endswith("-car"):
                group_label = group_label_map.get(cls[: -len("-car")], "")
                if group_label:
                    break
        items.append((text, group_label))

    counts = Counter(text for text, _ in items)
    results: list[str] = []
    for text, group_label in items:
        if group_label and counts[text] > 1:
            results.append(f"{text} {group_label}")
        else:
            results.append(text)
    return results


def build_header_names(header_scope: Tag | None) -> list[list[str]]:
    if not header_scope:
        return []
    group_label_map = build_group_label_map(header_scope)
    name_groups = header_scope.select(":scope > .name-select-amount")
    if not name_groups:
        name_groups = header_scope.select(".mc-header-wrapper > .name-select-amount")
    if name_groups:
        results = []
        for group in name_groups:
            slider = group.select_one(":scope > .list.mc-wrapper.mc-slider") or group.select_one(".mc-wrapper.mc-slider")
            names = extract_named_item_texts(slider, group_label_map)
            results.append(names)
        return results

    header_line = header_scope.select_one(":scope > .line.line-header") or header_scope.select_one(".line.line-header")
    if not header_line:
        return []

    slider = header_line.select_one(".mc-header-wrapper > .list.mc-wrapper.mc-slider") or header_line.select_one(
        ".mc-header-wrapper .list.mc-wrapper.mc-slider"
    )
    if slider:
        return [extract_named_item_texts(slider, group_label_map)]

    list_container = header_line.select_one(":scope > .list")
    if list_container:
        return [extract_direct_item_texts(list_container)]
    return []


def build_price_rows(header_scope: Tag | None) -> list[list[str]]:
    if not header_scope:
        return []

    price_groups: list[Tag] = []
    for selector in (":scope > .price-select-amount", ":scope > .m9-select-amount"):
        price_groups = header_scope.select(selector)
        if price_groups:
            break

    if price_groups:
        results = []
        for group in price_groups:
            list_container = group.select_one(".mc-header-list") or group.select_one(":scope > .line > .list")
            results.append(extract_direct_item_texts(list_container))
        return results

    price_line = None
    for line in header_scope.select(":scope > .line"):
        head = line.select_one(".line-head")
        head_text = text_inline(head)
        if "售价" in head_text:
            price_line = line
            break
    if not price_line:
        return []

    list_container = price_line.select_one(".mc-header-list") or price_line.select_one(":scope > .list")
    return [extract_direct_item_texts(list_container)]


def split_active_panels(root: Tag) -> list[dict[str, Tag]]:
    children = direct_children(root)
    panel_classes = ("car-specs-select", "m9-select")
    panels = [child for child in children if any(cls in child.get("class", []) for cls in panel_classes)]
    if panels:
        return [{"container": panel} for panel in panels]
    return [{"container": root}]


def extract_lines_from_container(container: Tag) -> list[Tag]:
    return container.select(".line")


def is_section_heading(line: Tag) -> bool:
    classes = line.get("class", [])
    return "line-one" in classes and line.select_one(".line-head") is not None


def repeat_width_for_classes(classes: list[str]) -> int:
    width_map = {
        "three-col": 3,
        "four-col": 4,
        "five-col": 5,
        "six-col": 6,
        "seven-col": 7,
        "two-col": 2,
    }
    for cls in classes:
        if cls in width_map:
            return width_map[cls]
    return 1


def fit_widths_to_target(widths: list[int], target: int) -> list[int] | None:
    adjusted = widths[:]
    total = sum(adjusted)
    if total == target:
        return adjusted
    if total > target:
        diff = total - target
        for i, width in enumerate(adjusted):
            reducible = max(width - 1, 0)
            if reducible <= 0:
                continue
            take = min(reducible, diff)
            adjusted[i] -= take
            diff -= take
            if diff == 0:
                break
        if diff == 0 and sum(adjusted) == target:
            return adjusted
        return None
    if total < target and adjusted:
        adjusted[-1] += target - total
        if sum(adjusted) == target:
            return adjusted
    return None


def expand_values_for_group_sizes(
    values: list[str],
    classes_list: list[list[str]],
    group_sizes: list[int],
    target_count: int,
) -> list[str]:
    if not group_sizes or len(values) % len(group_sizes) != 0:
        return values

    chunk_size = len(values) // len(group_sizes)
    expanded: list[str] = []
    start = 0
    for group_size in group_sizes:
        group_values = values[start : start + chunk_size]
        group_classes = classes_list[start : start + chunk_size]
        widths = [repeat_width_for_classes(classes) for classes in group_classes]
        fitted = fit_widths_to_target(widths, group_size)
        if not fitted:
            return values
        for value, repeat in zip(group_values, fitted):
            expanded.extend([value] * repeat)
        start += chunk_size

    if len(expanded) == target_count:
        return expanded
    return values


def expand_values_for_classes(values: list[str], classes_list: list[list[str]], target_count: int) -> list[str]:
    if not values:
        return values
    if len(values) == target_count:
        return values

    expanded: list[str] = []
    for value, classes in zip(values, classes_list):
        repeat = repeat_width_for_classes(classes)
        expanded.extend([value] * repeat)

    if len(expanded) == target_count:
        return expanded
    return values


def extract_attr_values(line: Tag, expected_count: int = 0, group_sizes: list[int] | None = None) -> list[str]:
    values: list[str] = []
    classes_list: list[list[str]] = []
    list_container = line.select_one(":scope > .list")
    if not list_container:
        return values

    direct_divs = [child for child in list_container.children if isinstance(child, Tag) and child.name == "div"]
    if not direct_divs:
        return values

    for div in direct_divs:
        text = text_multiline(div)
        if text:
            values.append(text)
            classes_list.append(div.get("class", []))

    if expected_count:
        if group_sizes:
            expanded = expand_values_for_group_sizes(values, classes_list, group_sizes, expected_count)
            if len(expanded) == expected_count:
                return expanded
        expanded = expand_values_for_classes(values, classes_list, expected_count)
        if len(expanded) == expected_count:
            return expanded
    return values


def ensure_length(values: list[str], target: int) -> list[str]:
    if target <= 0:
        return values
    if len(values) == target:
        return values
    if len(values) == 1 and target > 1:
        return values * target
    if len(values) > target:
        return values[:target]
    return values + [""] * (target - len(values))


@dataclass
class ProductEntry:
    product_key: str
    product_id: str
    product_name: str
    product_name_short: str
    sku_code: str
    sku_name: str
    product_url: str
    third_level_category_code: str
    specifications: list[dict[str, str]]

    def to_dict(self) -> dict[str, Any]:
        return {
            "product_key": self.product_key,
            "product_id": self.product_id,
            "product_name": self.product_name,
            "product_name_short": self.product_name_short,
            "sku_code": self.sku_code,
            "sku_name": self.sku_name,
            "product_url": self.product_url,
            "third_level_category_code": self.third_level_category_code,
            "specifications": self.specifications,
        }


def product_id_from_name(source: str) -> str:
    digest = hashlib.md5(source.encode("utf-8")).hexdigest()
    return str(int(digest[:15], 16) % 10**12).zfill(12)


def build_product_entries(
    model_name: str,
    detail_url: str,
    config_url: str,
    tab_label: str | None,
    names: list[str],
    prices: list[str],
    section_specs: list[dict[str, Any]],
) -> list[ProductEntry]:
    model_slug = slugify(model_name)
    tab_suffix = slugify(tab_label) if tab_label else ""
    products: list[ProductEntry] = []
    prices = ensure_length(prices, len(names))

    for idx, raw_name in enumerate(names):
        short_name = normalize_space(raw_name)
        full_name = short_name
        if tab_label:
            full_name = f"{short_name} [{normalize_space(tab_label)}]"

        source_id = f"{config_url}::{tab_label or 'default'}::{short_name}"
        product_key_parts = [model_slug]
        if tab_suffix:
            product_key_parts.append(tab_suffix)
        product_key_parts.append(slugify(short_name))
        product_key = "-".join(part for part in product_key_parts if part)

        specs: list[dict[str, str]] = []
        brand_name = infer_brand_name(short_name)
        specs.append({"group_name": "主体", "attr_name": "品牌", "attr_value": brand_name})
        specs.append({"group_name": "主体", "attr_name": "车型", "attr_value": model_name})
        if tab_label:
            specs.append({"group_name": "主体", "attr_name": "版本页", "attr_value": normalize_space(tab_label)})
        price = prices[idx] if idx < len(prices) else ""
        if price:
            specs.append({"group_name": "价格", "attr_name": "建议零售价", "attr_value": price})

        for item in section_specs:
            values = ensure_length(item["values"], len(names))
            value = values[idx]
            if not value:
                continue
            specs.append(
                {
                    "group_name": item["group_name"],
                    "attr_name": item["attr_name"],
                    "attr_value": value,
                }
            )

        product_id = product_id_from_name(source_id)
        products.append(
            ProductEntry(
                product_key=product_key,
                product_id=product_id,
                product_name=full_name,
                product_name_short=short_name,
                sku_code=product_id,
                sku_name=full_name,
                product_url=detail_url,
                third_level_category_code=model_slug,
                specifications=specs,
            )
        )
    return products


class HimaCarExporter(JsonCategoryExportMixin):
    """Export car configuration specs from hima.auto."""

    def __init__(
        self,
        *,
        session: requests.Session | None = None,
        timeout: int = 30,
        request_interval: float = 0.8,
        logger: logging.Logger | None = None,
    ) -> None:
        self.session = session or requests.Session()
        self.timeout = timeout
        self.request_interval = request_interval
        self.logger = logger or logging.getLogger(__name__)
        self._last_request_ts = 0.0

    def export_category(
        self,
        *,
        output_dir: str | Path,
        diff_dir: str | Path,
        limit: int | None = None,
    ) -> tuple[Path, int]:
        payload = self.build_payload(limit=limit)
        file_path = self._write_category_payload(
            payload,
            output_dir=output_dir,
            diff_dir=diff_dir,
        )
        self.logger.info("Finished HIMA export: file=%s products=%s", file_path, self._count_products(payload))
        return file_path, self._count_products(payload)

    def build_payload(self, *, limit: int | None = None) -> dict[str, Any]:
        models = self.discover_model_pages()
        payload: dict[str, Any] = {
            "category_key": HIMA_CATEGORY_KEY,
            "category_name": HIMA_CATEGORY_NAME,
            "pull_time": self._current_pull_time(),
            "series": {},
        }
        total_count = 0

        for model in models:
            if limit is not None and total_count >= limit:
                break
            config_urls = self.discover_configuration_urls(model["detail_url"])
            if not config_urls:
                continue

            series_key = model["brand_key"]
            series_name = model["brand_name"]
            series_entry = payload["series"].setdefault(
                series_key,
                {
                    "series_key": series_key,
                    "series_name": series_name,
                    "products": {},
                },
            )
            product_map: dict[str, dict[str, Any]] = series_entry["products"]

            for config_url in config_urls:
                if limit is not None and total_count >= limit:
                    break
                parsed = self.parse_configuration_page(model, config_url)
                for products in parsed.values():
                    for product in products:
                        if limit is not None and total_count >= limit:
                            break
                        is_new = product.product_key not in product_map
                        product_map[product.product_key] = product.to_dict()
                        if is_new:
                            total_count += 1

        return payload

    def discover_model_pages(self) -> list[dict[str, str]]:
        items: list[dict[str, str]] = []
        seen = set()
        for brand_url in BRAND_URLS:
            brand_key = [part for part in urlparse(brand_url).path.split("/") if part][0]
            soup = BeautifulSoup(self.fetch_html(brand_url), "html.parser")
            for anchor in soup.select("a[href]"):
                href = anchor.get("href", "")
                if not href:
                    continue
                absolute = urljoin(BASE_URL, href)
                parsed = urlparse(absolute)
                if parsed.netloc != "hima.auto":
                    continue
                path = parsed.path
                if not path.endswith("/"):
                    path += "/"
                parts = [part for part in path.split("/") if part]
                if len(parts) != 2:
                    continue
                brand, model = parts
                if brand != brand_key or model == "service":
                    continue
                key = (brand, model)
                if key in seen:
                    continue
                seen.add(key)
                items.append(
                    {
                        "brand_key": brand,
                        "brand_name": BRAND_NAMES.get(brand, brand),
                        "model_key": model,
                        "detail_url": f"https://hima.auto/{brand}/{model}/",
                    }
                )
        items.sort(key=lambda x: (x["brand_key"], x["model_key"]))
        return items

    def discover_configuration_urls(self, detail_url: str) -> list[str]:
        base_config = urljoin(detail_url, "configuration/")
        urls = [base_config]
        seen = {base_config}

        html = self.fetch_html(base_config)
        if is_404_page(html):
            return []

        for href in re.findall(r'href="([^"]*configuration\d+/)"', html):
            absolute = urljoin(base_config, href)
            if absolute not in seen:
                seen.add(absolute)
                urls.append(absolute)
        return urls

    def parse_configuration_page(self, model_meta: dict[str, str], config_url: str) -> dict[str, list[ProductEntry]]:
        html = self.fetch_html(config_url)
        if is_404_page(html):
            return {}

        soup = BeautifulSoup(html, "html.parser")
        asking_blocks = [
            block for block in soup.select("div.asking-m5") if "table-mob-show" not in block.get("class", [])
        ]
        if not asking_blocks:
            return {}

        root = asking_blocks[0]
        title_text = find_page_title_block(root) or first_title_text(soup)
        model_name = re.sub(r"\s*参数配置表?.*$", "", title_text).strip(" -")
        if not model_name:
            model_name = model_meta["brand_name"]

        header_scope = root.select_one(":scope > .asking-content.cp-header")
        header_container = header_scope.select_one(".header-containr") if header_scope else None
        panel_groups = split_active_panels(root)
        header_names_groups = build_header_names(header_container or header_scope)
        price_groups = build_price_rows(header_container or header_scope)
        group_label_map = build_group_label_map(header_container or header_scope)
        tab_labels = find_tab_labels(root)
        current_tab_label = find_current_tab_label(root)

        result: dict[str, list[ProductEntry]] = {}
        total_groups = max(len(panel_groups), len(header_names_groups), len(price_groups), 1)

        if total_groups > 1:
            if not tab_labels:
                tab_labels = [current_tab_label or ""] * total_groups
            elif len(tab_labels) < total_groups:
                tab_labels = tab_labels + [tab_labels[-1]] * (total_groups - len(tab_labels))
        else:
            tab_labels = [current_tab_label or ""]

        if not header_names_groups and header_scope:
            header_names_groups = [[] for _ in range(total_groups)]
        elif len(header_names_groups) == 1 and total_groups > 1:
            header_names_groups = header_names_groups * total_groups

        if not price_groups:
            price_groups = [[] for _ in range(total_groups)]
        elif len(price_groups) == 1 and total_groups > 1:
            price_groups = price_groups * total_groups

        for idx in range(total_groups):
            panel = panel_groups[idx]["container"] if idx < len(panel_groups) else root
            names = header_names_groups[idx] if idx < len(header_names_groups) else []
            prices = price_groups[idx] if idx < len(price_groups) else []
            tab_label = tab_labels[idx] if idx < len(tab_labels) else None
            group_sizes = infer_group_sizes(names, group_label_map) if names else []

            lines = extract_lines_from_container(panel)
            current_group = ""
            section_specs: list[dict[str, Any]] = []
            for line in lines:
                line_head = line.select_one(":scope > .line-head")
                if not line_head:
                    continue
                head_text = text_inline(line_head)
                if not head_text or "配置项 / 配置名称" in head_text:
                    continue
                if "售价" in head_text:
                    continue
                if is_section_heading(line):
                    current_group = head_text
                    continue
                values = extract_attr_values(line, expected_count=len(names), group_sizes=group_sizes)
                if not values:
                    continue
                section_specs.append({"group_name": current_group or "未分组", "attr_name": head_text, "values": values})

            if not names:
                max_value_len = max((len(item["values"]) for item in section_specs), default=0)
                names = [f"{model_name} 版本{i + 1}" for i in range(max_value_len)]

            products = build_product_entries(
                model_name=model_name,
                detail_url=model_meta["detail_url"],
                config_url=config_url,
                tab_label=tab_label,
                names=names,
                prices=prices,
                section_specs=section_specs,
            )
            group_key = normalize_space(tab_label) if isinstance(tab_label, str) and tab_label else model_name
            result[group_key] = products
        return result

    def fetch_html(self, url: str) -> str:
        self._throttle()
        self.logger.info("Fetching HIMA page: %s", url)
        response = self.session.get(url, headers=HEADERS, timeout=self.timeout)
        self._last_request_ts = time.monotonic()
        response.raise_for_status()
        return response.text

    def _throttle(self) -> None:
        if self.request_interval <= 0:
            return
        now = time.monotonic()
        elapsed = now - self._last_request_ts
        if elapsed < self.request_interval:
            time.sleep(self.request_interval - elapsed)

    def _count_products(self, payload: dict[str, Any]) -> int:
        series = payload.get("series", {})
        return sum(len(item.get("products", {})) for item in series.values())

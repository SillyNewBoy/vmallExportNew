"""Typed models for VMALL export."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class SearchQueryConfig:
    keyword: str
    category_id: int | None = None
    allowed_category_codes: tuple[str, ...] = ()


@dataclass(frozen=True)
class CategoryConfig:
    key: str
    name: str
    keyword: str
    category_id: int | None = None
    search_queries: tuple[SearchQueryConfig, ...] = ()

    @property
    def queries(self) -> tuple[SearchQueryConfig, ...]:
        if self.search_queries:
            return self.search_queries
        return (
            SearchQueryConfig(
                keyword=self.keyword,
                category_id=self.category_id,
            ),
        )


@dataclass(frozen=True)
class SearchProduct:
    product_id: str
    product_name: str
    sku_code: str
    sku_name: str
    category_code: str | None
    raw: dict[str, Any]

    @property
    def product_url(self) -> str:
        return f"https://www.vmall.com/product/{self.product_id}.html"

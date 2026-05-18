from vmall_export_new.exporter import VmallExporter
from vmall_export_new.models import CategoryConfig, SearchProduct, SearchQueryConfig


def test_dedupe_products_keeps_first_unique_name():
    exporter = VmallExporter(client=None)
    products = [
        SearchProduct(
            product_id="1",
            product_name="HUAWEI Pura 90 Pro",
            sku_code="sku-1",
            sku_name="HUAWEI Pura 90 Pro 12GB+256GB 椰青白",
            category_code="260101",
            raw={},
        ),
        SearchProduct(
            product_id="2",
            product_name="HUAWEI Pura 90 Pro",
            sku_code="sku-2",
            sku_name="HUAWEI Pura 90 Pro 16GB+512GB 粉红芭乐",
            category_code="260101",
            raw={},
        ),
        SearchProduct(
            product_id="3",
            product_name="HUAWEI MatePad Edge",
            sku_code="sku-3",
            sku_name="HUAWEI MatePad Edge WiFi 16GB+512GB 皓月银",
            category_code="270101",
            raw={},
        ),
    ]

    deduped = exporter._dedupe_products(products, limit=None)

    assert [item.product_id for item in deduped] == ["1", "3"]


def test_count_products_uses_nested_product_total():
    exporter = VmallExporter(client=None)
    payload = {
        "series": {
            "a": {"products": {"p1": {}, "p2": {}}},
            "b": {"products": {"p3": {}}},
        }
    }

    assert exporter._count_products(payload) == 3


def test_category_queries_fallbacks_to_single_query():
    category = CategoryConfig(
        key="phones",
        name="手机",
        keyword="手机",
        category_id=100507,
    )

    assert category.queries == (
        SearchQueryConfig(keyword="手机", category_id=100507),
    )


def test_search_query_filter_keeps_allowed_category_codes_only():
    class StubClient:
        def search_products(self, **kwargs):
            return [
                SearchProduct(
                    product_id="1",
                    product_name="A",
                    sku_code="sku-1",
                    sku_name="A-1",
                    category_code="310205",
                    raw={},
                ),
                SearchProduct(
                    product_id="2",
                    product_name="B",
                    sku_code="sku-2",
                    sku_name="B-1",
                    category_code="310206",
                    raw={},
                ),
            ]

    exporter = VmallExporter(client=StubClient())

    items = exporter._search_products_for_query(
        SearchQueryConfig(
            keyword="耳机",
            category_id=100559,
            allowed_category_codes=("310205",),
        ),
        page_size=20,
        max_pages=1,
    )

    assert [item.product_id for item in items] == ["1"]

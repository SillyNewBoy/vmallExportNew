import json

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


def test_build_category_tree_includes_pull_time():
    exporter = VmallExporter(client=None)
    category = CategoryConfig(
        key="phones",
        name="鎵嬫満",
        keyword="鎵嬫満",
        category_id=100507,
    )

    payload = exporter._build_category_tree(category, [], pull_time="260518")

    assert payload["pull_time"] == "260518"


def test_build_new_products_payload_compares_product_ids():
    exporter = VmallExporter(client=None)
    previous_payload = {
        "category_key": "phones",
        "category_name": "鎵嬫満",
        "pull_time": "260517",
        "series": {
            "mate": {
                "series_key": "mate",
                "series_name": "Mate",
                "products": {
                    "mate70": {"product_id": "1001", "product_name": "Mate 70"},
                },
            },
        },
    }
    current_payload = {
        "category_key": "phones",
        "category_name": "鎵嬫満",
        "pull_time": "260518",
        "series": {
            "mate": {
                "series_key": "mate",
                "series_name": "Mate",
                "products": {
                    "mate70": {"product_id": "1001", "product_name": "Mate 70"},
                    "mate80": {"product_id": "1002", "product_name": "Mate 80"},
                },
            },
        },
    }

    diff_payload = exporter._build_new_products_payload(previous_payload, current_payload)

    assert diff_payload["previous_pull_time"] == "260517"
    assert diff_payload["current_pull_time"] == "260518"
    assert diff_payload["new_product_count"] == 1
    assert list(diff_payload["series"]["mate"]["products"]) == ["mate80"]


def test_write_new_products_diff_creates_file(tmp_path):
    exporter = VmallExporter(client=None)
    previous_payload = {
        "category_key": "phones",
        "category_name": "鎵嬫満",
        "pull_time": "260517",
        "series": {},
    }
    current_payload = {
        "category_key": "phones",
        "category_name": "鎵嬫満",
        "pull_time": "260518",
        "series": {
            "mate": {
                "series_key": "mate",
                "series_name": "Mate",
                "products": {
                    "mate80": {"product_id": "1002", "product_name": "Mate 80"},
                },
            },
        },
    }

    diff_path = exporter._write_new_products_diff(
        previous_payload,
        current_payload,
        diff_dir=tmp_path,
    )

    assert diff_path == tmp_path / "phones_260517_260518_new.json"
    diff_payload = json.loads(diff_path.read_text(encoding="utf-8"))
    assert diff_payload["new_product_count"] == 1
    assert diff_payload["series"]["mate"]["products"]["mate80"]["product_id"] == "1002"


def test_write_new_products_diff_skips_when_no_previous_payload(tmp_path):
    exporter = VmallExporter(client=None)

    diff_path = exporter._write_new_products_diff(
        None,
        {"category_key": "phones", "pull_time": "260518", "series": {}},
        diff_dir=tmp_path,
    )

    assert diff_path is None
    assert not any(tmp_path.iterdir())


def test_export_category_writes_pull_time_and_diff(tmp_path):
    class StubExporter(VmallExporter):
        def __init__(self):
            super().__init__(client=None)

        def _current_pull_time(self):
            return "260518"

        def _build_export_item(self, category, product):
            return {
                "category_key": category.key,
                "category_name": category.name,
                "series_name": "Mate",
                "series_key": "mate",
                "product_id": product.product_id,
                "product_name": product.product_name,
                "product_name_short": product.product_name,
                "product_key": product.product_name.lower().replace(" ", ""),
                "sku_code": product.sku_code,
                "sku_name": product.sku_name,
                "product_url": product.product_url,
                "third_level_category_code": product.category_code,
                "specifications": [],
            }

    class StubClient:
        def search_products(self, **kwargs):
            return [
                SearchProduct(
                    product_id="1001",
                    product_name="Mate 70",
                    sku_code="sku-1",
                    sku_name="Mate 70",
                    category_code="260101",
                    raw={},
                ),
                SearchProduct(
                    product_id="1002",
                    product_name="Mate 80",
                    sku_code="sku-2",
                    sku_name="Mate 80",
                    category_code="260101",
                    raw={},
                ),
            ]

    output_dir = tmp_path / "output"
    diff_dir = tmp_path / "diff"
    output_dir.mkdir()
    (output_dir / "phones.json").write_text(
        json.dumps(
            {
                "category_key": "phones",
                "category_name": "鎵嬫満",
                "pull_time": "260517",
                "series": {
                    "mate": {
                        "series_key": "mate",
                        "series_name": "Mate",
                        "products": {
                            "mate70": {
                                "product_id": "1001",
                                "product_name": "Mate 70",
                            },
                        },
                    },
                },
            }
        ),
        encoding="utf-8",
    )
    exporter = StubExporter()
    exporter.client = StubClient()
    category = CategoryConfig(
        key="phones",
        name="鎵嬫満",
        keyword="鎵嬫満",
        category_id=100507,
    )

    file_path, count = exporter.export_category(
        category,
        output_dir=output_dir,
        diff_dir=diff_dir,
    )

    payload = json.loads(file_path.read_text(encoding="utf-8"))
    diff_payload = json.loads(
        (diff_dir / "phones_260517_260518_new.json").read_text(encoding="utf-8")
    )
    assert count == 2
    assert payload["pull_time"] == "260518"
    assert diff_payload["new_product_count"] == 1
    assert diff_payload["series"]["mate"]["products"]["mate80"]["product_id"] == "1002"


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

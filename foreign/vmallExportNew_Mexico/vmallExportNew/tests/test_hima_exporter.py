import json

from vmall_export_new.hima_exporter import HIMA_CATEGORY_KEY, HimaCarExporter, slugify


class StubHimaExporter(HimaCarExporter):
    def __init__(self, pages):
        super().__init__(request_interval=0)
        self.pages = pages

    def _current_pull_time(self):
        return "260518"

    def fetch_html(self, url):
        return self.pages[url]


def test_slugify_keeps_chinese_and_normalizes_symbols():
    assert slugify("问界 M9｜Max+") == "问界-m9-max-plus"


def test_parse_configuration_page_extracts_car_products():
    config_url = "https://hima.auto/wenjie/m9/configuration/"
    html = """
    <html>
      <head><title>问界 M9 参数配置表</title></head>
      <body>
        <div class="asking-m5">
          <h2>问界 M9 参数配置表</h2>
          <div class="asking-content cp-header">
            <div class="header-containr">
              <div class="line line-header">
                <div class="mc-header-wrapper">
                  <div class="list mc-wrapper mc-slider">
                    <div>问界 M9 Max</div>
                    <div>问界 M9 Ultra</div>
                  </div>
                </div>
              </div>
              <div class="line">
                <div class="line-head">售价</div>
                <div class="list">
                  <div>46.98万元</div>
                  <div>56.98万元</div>
                </div>
              </div>
            </div>
          </div>
          <div class="line line-one"><div class="line-head">基本参数</div></div>
          <div class="line">
            <div class="line-head">续航里程</div>
            <div class="list">
              <div>630km</div>
              <div>700km</div>
            </div>
          </div>
        </div>
      </body>
    </html>
    """
    exporter = StubHimaExporter({config_url: html})

    parsed = exporter.parse_configuration_page(
        {"brand_key": "wenjie", "brand_name": "问界", "detail_url": "https://hima.auto/wenjie/m9/"},
        config_url,
    )

    products = parsed["问界 M9"]
    assert [item.product_name for item in products] == ["问界 M9 Max", "问界 M9 Ultra"]
    assert products[0].product_key == "问界-m9-问界-m9-max"
    assert products[0].specifications[:4] == [
        {"group_name": "主体", "attr_name": "品牌", "attr_value": "问界"},
        {"group_name": "主体", "attr_name": "车型", "attr_value": "问界 M9"},
        {"group_name": "价格", "attr_name": "建议零售价", "attr_value": "46.98万元"},
        {"group_name": "基本参数", "attr_name": "续航里程", "attr_value": "630km"},
    ]
    assert products[1].specifications[-1]["attr_value"] == "700km"


def test_export_category_writes_cars_payload_and_diff(tmp_path):
    class Exporter(StubHimaExporter):
        def discover_model_pages(self):
            return [
                {
                    "brand_key": "wenjie",
                    "brand_name": "问界",
                    "model_key": "m9",
                    "detail_url": "https://hima.auto/wenjie/m9/",
                }
            ]

        def discover_configuration_urls(self, detail_url):
            return ["https://hima.auto/wenjie/m9/configuration/"]

        def parse_configuration_page(self, model_meta, config_url):
            return {
                "问界 M9": [
                    self._product("问界-m9-max", "1001", "问界 M9 Max"),
                    self._product("问界-m9-ultra", "1002", "问界 M9 Ultra"),
                ]
            }

        def _product(self, product_key, product_id, product_name):
            from vmall_export_new.hima_exporter import ProductEntry

            return ProductEntry(
                product_key=product_key,
                product_id=product_id,
                product_name=product_name,
                product_name_short=product_name,
                sku_code=product_id,
                sku_name=product_name,
                product_url="https://hima.auto/wenjie/m9/",
                third_level_category_code="问界-m9",
                specifications=[],
            )

    output_dir = tmp_path / "out"
    diff_dir = tmp_path / "diff"
    output_dir.mkdir()
    (output_dir / "cars.json").write_text(
        json.dumps(
            {
                "category_key": HIMA_CATEGORY_KEY,
                "category_name": "汽车",
                "pull_time": "260517",
                "series": {
                    "wenjie": {
                        "series_key": "wenjie",
                        "series_name": "问界",
                        "products": {
                            "问界-m9-max": {"product_id": "1001", "product_name": "问界 M9 Max"},
                        },
                    },
                },
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    exporter = Exporter({})

    file_path, count = exporter.export_category(output_dir=output_dir, diff_dir=diff_dir)

    payload = json.loads(file_path.read_text(encoding="utf-8"))
    diff_payload = json.loads((diff_dir / "cars_260517_260518_new.json").read_text(encoding="utf-8"))
    assert count == 2
    assert payload["category_key"] == "cars"
    assert payload["pull_time"] == "260518"
    assert diff_payload["new_product_count"] == 1
    assert diff_payload["series"]["wenjie"]["products"]["问界-m9-ultra"]["product_id"] == "1002"

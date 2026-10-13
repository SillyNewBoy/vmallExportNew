"""Tests for the consumer.huawei.com specs page parser."""

from __future__ import annotations

import logging

from vmall_export_new.consumer_exporter import ConsumerHuaweiExporter

SAMPLE_HTML = """
<html><body>
<div class="specs-banner-cnt">
    <div class="large-accordion-title">HUAWEI Mate X7</div>
    <div class="color-text">Brocade White, Nebula Red, Black</div>
</div>
<ul class="large-accordion__list">
    <li class="large-accordion__item">
        <span class="large-accordion__title large-accordion-title">Dimensions</span>
        <div class="large-accordion__content">
            <div class="large-accordion-columns">
                <div class="large-accordion__inner">
                    <div class="large-accordion-subtitle">Length</div>
                    <p>156.8 mm</p>
                </div>
                <div class="large-accordion__inner">
                    <div class="large-accordion-subtitle">Width</div>
                    <p>Folded: 73.8 mm<br>Unfolded: 144.2 mm</p>
                    <p class="large-accordion-subtext">*Measured across.</p>
                </div>
            </div>
        </div>
    </li>
    <li class="large-accordion__item">
        <span class="large-accordion__title large-accordion-title">Dimensions</span>
        <div class="large-accordion__content">
            <div class="spec-size">
                <div class="product-size large-accordion-columns">
                    <div class="large-accordion__inner large-accordion__inner--short">
                        <div class="large-accordion__wrap">
                            <div class="large-accordion-subtitle size-height-name">Height</div>
                            <p>14.5 mm</p>
                        </div>
                        <div class="large-accordion__wrap">
                            <div class="large-accordion-subtitle size-width-name">Width</div>
                            <p>312.6 mm</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>
</body></html>
"""


def _make_exporter() -> ConsumerHuaweiExporter:
    exporter = ConsumerHuaweiExporter.__new__(ConsumerHuaweiExporter)
    exporter.logger = logging.getLogger("test")
    return exporter


def test_parse_specs_page_basic_groups():
    exporter = _make_exporter()
    result = exporter.parse_specs_page(SAMPLE_HTML, product_path="phones/mate-x7/")

    assert result["product_name"] == "HUAWEI Mate X7"
    assert [c["attr_value"] for c in result["color_summary"]] == [
        "Brocade White",
        "Nebula Red",
        "Black",
    ]

    specs = result["specifications"]
    by_name: dict[str, dict] = {}
    for spec in specs:
        by_name.setdefault(spec["attr_name"], spec)
    assert by_name["Length"]["group_name"] == "Dimensions"
    assert by_name["Length"]["attr_value"] == "156.8 mm"
    assert "Folded: 73.8 mm" in by_name["Width"]["attr_value"]
    assert "Unfolded: 144.2 mm" in by_name["Width"]["attr_value"]
    assert by_name["Width"]["attr_remark"] == "*Measured across."


def test_parse_specs_page_wrap_variant():
    exporter = _make_exporter()
    result = exporter.parse_specs_page(SAMPLE_HTML, product_path="laptops/matebook-14/")

    specs = result["specifications"]
    by_name = {spec["attr_name"]: spec for spec in specs}
    assert by_name["Height"]["attr_value"] == "14.5 mm"
    assert by_name["Width"]["attr_value"] == "312.6 mm"


def test_series_derivation_rules():
    exporter = _make_exporter()
    assert exporter._derive_series_name("HUAWEI Mate X7") == "Mate X 系列"
    assert exporter._derive_series_name("HUAWEI Mate 80 Pro") == "Mate 系列"
    assert exporter._derive_series_name("HUAWEI Pura 90s Pro Max") == "Pura 系列"
    assert exporter._derive_series_name("HUAWEI nova 14") == "nova 系列"
    assert exporter._derive_series_name("HUAWEI MateBook 14") == "MateBook 系列"
    assert exporter._derive_series_name("HUAWEI WATCH GT 6 Pro") == "WATCH GT 系列"
    assert exporter._derive_series_name("HUAWEI FreeBuds Pro 4") == "FreeBuds 系列"
    assert exporter._derive_series_name("Unknown Gadget") == "其他"


def test_non_product_slug_filter():
    exporter = _make_exporter()
    assert exporter._is_non_product_slug("specs")
    assert exporter._is_non_product_slug("buy")
    assert not exporter._is_non_product_slug("mate-x7")

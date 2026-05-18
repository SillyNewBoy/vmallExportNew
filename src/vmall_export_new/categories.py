"""Built-in VMALL category presets."""

from __future__ import annotations

from .models import SearchQueryConfig

CATEGORY_PRESETS = {
    "phones": {"name": "手机", "keyword": "手机", "category_id": 100507},
    "tablets": {"name": "平板电脑", "keyword": "平板", "category_id": 100515},
    "wearables": {"name": "手表", "keyword": "手表", "category_id": 100523},
    "audio": {
        "name": "耳机音箱",
        "keyword": "耳机音箱",
        "category_id": None,
        "search_queries": (
            SearchQueryConfig(
                keyword="耳机",
                category_id=100559,
                allowed_category_codes=("310205",),
            ),
            SearchQueryConfig(
                keyword="智能音箱",
                category_id=100561,
                allowed_category_codes=("310206",),
            ),
            SearchQueryConfig(
                keyword="智能眼镜",
                category_id=100620,
                allowed_category_codes=("310218",),
            ),
            SearchQueryConfig(
                keyword="有线耳机",
                category_id=100591,
                allowed_category_codes=("310213",),
            ),
        ),
    },
    "laptops": {"name": "笔记本电脑", "keyword": "笔记本", "category_id": 100519},
    "routers": {"name": "路由器", "keyword": "路由", "category_id": 100529},
    "smart_screens": {"name": "智慧屏", "keyword": "智慧屏", "category_id": 100636},
    "printers": {"name": "打印机", "keyword": "打印机", "category_id": 101086},
    "monitors": {"name": "显示器", "keyword": "显示器", "category_id": 100848},
}

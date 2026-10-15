"""Built-in category presets for the HUAWEI consumer site (consumer.huawei.com/my/).

Each entry maps a category key to its site-relative listing path. Products are
discovered from the listing page and supplemented by the region sitemap.
"""

from __future__ import annotations

# key -> {name, path}  path is the top-level section on the consumer site.
CONSUMER_CATEGORY_PRESETS = {
    "phones": {"name": "手机", "path": "phones"},
    "tablets": {"name": "平板电脑", "path": "tablets"},
    "wearables": {"name": "穿戴设备", "path": "wearables"},
    "laptops": {"name": "笔记本电脑", "path": "laptops"},
    "headphones": {"name": "耳机", "path": "headphones"},
    "speakers": {"name": "音箱", "path": "speakers"},
    "routers": {"name": "路由器", "path": "routers"},
    "monitors": {"name": "显示器", "path": "monitors"},
    "printers": {"name": "打印机", "path": "printers"},
    "visions": {"name": "智慧屏", "path": "visions"},
}

# Series derivation rules: checked in order against the product name.
# The first matching keyword (case-insensitive) decides the series name.
CONSUMER_SERIES_RULES: tuple[tuple[str, str], ...] = (
    ("mate xt", "Mate XT 系列"),
    ("mate xs", "Mate X 系列"),
    ("mate x", "Mate X 系列"),
    ("matebook", "MateBook 系列"),
    ("matepad", "MatePad 系列"),
    ("matestation", "MateStation 系列"),
    ("mateview", "MateView 系列"),
    ("mate ", "Mate 系列"),
    ("pura", "Pura 系列"),
    ("nova", "nova 系列"),
    ("enjoy", "畅享系列"),
    ("watch gt", "WATCH GT 系列"),
    ("watch fit", "WATCH FIT 系列"),
    ("watch d", "WATCH D 系列"),
    ("watch ultimate", "WATCH ULTIMATE 系列"),
    ("watch kids", "儿童手表系列"),
    ("watch", "WATCH 系列"),
    ("band", "手环系列"),
    ("freebuds", "FreeBuds 系列"),
    ("freeclip", "FreeClip 系列"),
    ("freearc", "FreeArc 系列"),
    ("freelace", "FreeLace 系列"),
    ("sound x", "Sound 系列"),
    ("sound", "Sound 系列"),
    ("vision", "智慧屏系列"),
    ("pixlab", "打印机系列"),
    ("wifi mesh", "路由系列"),
    ("wifi", "路由系列"),
)

CONSUMER_DEFAULT_SERIES_NAME = "其他"

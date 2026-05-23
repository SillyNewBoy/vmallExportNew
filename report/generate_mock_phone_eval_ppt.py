from __future__ import annotations

import argparse
import random
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "report"
RAW_XLSX = REPORT_DIR / "手机全量参数评测_模拟数据.xlsx"
SUMMARY_XLSX = REPORT_DIR / "手机全量参数评测统计_模拟数据.xlsx"
PPTX_PATH = ROOT / "小艺手机全量参数评测报告_模拟数据.pptx"

FONT = "Microsoft YaHei"
TOTAL_CASES = 5800
RANDOM_SEED = 20260523

COLORS = {
    "red": "C7000B",
    "red_dark": "8F0008",
    "dark": "1F2329",
    "gray": "6B7280",
    "light_gray": "EEF1F5",
    "mid_gray": "D5DAE1",
    "bg": "F7F8FA",
    "white": "FFFFFF",
    "green": "16845B",
    "blue": "2563EB",
    "orange": "D97706",
    "amber": "F59E0B",
    "purple": "7C3AED",
}

DIMENSIONS = ["相关性", "正确性", "信息量", "一致性"]
BASE_COLUMNS = [
    "case_id",
    "四级分类",
    "query所属设备",
    "五级分类",
    "六级分类",
    "query",
    "displayText",
    "备注",
    "人工综合评测结果",
    *DIMENSIONS,
]

L5_SPECS = {
    "主体": {
        "weight": 11,
        "dims": {"相关性": 0.010, "正确性": 0.055, "信息量": 0.055, "一致性": 0.010},
        "l6": ["品牌", "上市时间", "操作系统", "机身内存", "运行内存", "CPU型号", "双卡类型"],
    },
    "屏幕": {
        "weight": 13,
        "dims": {"相关性": 0.012, "正确性": 0.080, "信息量": 0.090, "一致性": 0.012},
        "l6": ["屏幕尺寸", "屏幕色彩", "屏幕类型", "分辨率", "刷新率", "触控采样率", "护眼认证"],
    },
    "摄像头": {
        "weight": 13,
        "dims": {"相关性": 0.015, "正确性": 0.105, "信息量": 0.125, "一致性": 0.018},
        "l6": ["后置摄像头", "前置摄像头", "变焦模式", "防抖模式", "拍摄功能", "视频拍摄", "闪光灯"],
    },
    "电池与充电": {
        "weight": 10,
        "dims": {"相关性": 0.012, "正确性": 0.100, "信息量": 0.105, "一致性": 0.014},
        "l6": ["电池容量", "有线充电", "无线充电", "反向充电", "快充协议", "充电器规格"],
    },
    "网络制式": {
        "weight": 9,
        "dims": {"相关性": 0.020, "正确性": 0.110, "信息量": 0.095, "一致性": 0.016},
        "l6": ["5G网络", "4G网络", "SIM卡类型", "卫星通信", "双卡双待", "频段"],
    },
    "传输功能": {
        "weight": 9,
        "dims": {"相关性": 0.025, "正确性": 0.090, "信息量": 0.155, "一致性": 0.018},
        "l6": ["WLAN", "蓝牙", "NFC", "USB接口", "OTG", "定位", "红外遥控"],
    },
    "机身外观": {
        "weight": 8,
        "dims": {"相关性": 0.010, "正确性": 0.060, "信息量": 0.070, "一致性": 0.010},
        "l6": ["机身尺寸", "机身重量", "机身材质", "配色", "防水等级", "折叠形态"],
    },
    "传感器": {
        "weight": 7,
        "dims": {"相关性": 0.018, "正确性": 0.075, "信息量": 0.120, "一致性": 0.012},
        "l6": ["重力传感器", "陀螺仪", "指纹识别", "人脸识别", "环境光传感器", "距离传感器"],
    },
    "音频": {
        "weight": 5,
        "dims": {"相关性": 0.018, "正确性": 0.070, "信息量": 0.110, "一致性": 0.010},
        "l6": ["扬声器", "麦克风", "音频格式", "耳机接口", "音效能力"],
    },
    "软件与服务": {
        "weight": 5,
        "dims": {"相关性": 0.018, "正确性": 0.075, "信息量": 0.095, "一致性": 0.012},
        "l6": ["AI功能", "隐私安全", "智慧互联", "系统版本", "预置应用"],
    },
    "包装清单": {
        "weight": 5,
        "dims": {"相关性": 0.012, "正确性": 0.060, "信息量": 0.090, "一致性": 0.008},
        "l6": ["主机", "充电器", "数据线", "保护壳", "说明书", "取卡针"],
    },
    "其他": {
        "weight": 5,
        "dims": {"相关性": 0.018, "正确性": 0.070, "信息量": 0.095, "一致性": 0.012},
        "l6": ["保修政策", "备注说明", "适用场景", "版本差异", "销售状态"],
    },
}

L6_OVERRIDES = {
    "OTG": {"信息量": 0.34, "正确性": 0.12},
    "卫星通信": {"正确性": 0.22, "信息量": 0.16},
    "反向充电": {"信息量": 0.25, "正确性": 0.16},
    "变焦模式": {"正确性": 0.18, "信息量": 0.18},
    "快充协议": {"正确性": 0.16, "信息量": 0.18},
    "防水等级": {"正确性": 0.15},
    "折叠形态": {"正确性": 0.14, "信息量": 0.13},
    "频段": {"正确性": 0.17, "信息量": 0.16},
    "触控采样率": {"信息量": 0.18},
    "护眼认证": {"信息量": 0.17},
    "AI功能": {"信息量": 0.18, "相关性": 0.04},
}

DEVICES = [
    {"device": "pura80", "display": "HUAWEI Pura 80", "weight": 9, "risk": 0.98},
    {"device": "pura80pro", "display": "HUAWEI Pura 80 Pro", "weight": 10, "risk": 1.08},
    {"device": "pura80ultra", "display": "HUAWEI Pura 80 Ultra", "weight": 8, "risk": 1.18},
    {"device": "pura90", "display": "HUAWEI Pura 90", "weight": 8, "risk": 1.16},
    {"device": "pura90pro", "display": "HUAWEI Pura 90 Pro", "weight": 9, "risk": 1.28},
    {"device": "mate70", "display": "HUAWEI Mate 70", "weight": 9, "risk": 0.95},
    {"device": "mate70pro", "display": "HUAWEI Mate 70 Pro", "weight": 10, "risk": 1.02},
    {"device": "mate70rs", "display": "HUAWEI Mate 70 RS", "weight": 5, "risk": 1.12},
    {"device": "matex6", "display": "HUAWEI Mate X6", "weight": 7, "risk": 1.22},
    {"device": "matext", "display": "HUAWEI Mate XT", "weight": 5, "risk": 1.35},
    {"device": "nova14", "display": "HUAWEI nova 14", "weight": 10, "risk": 0.86},
    {"device": "nova14pro", "display": "HUAWEI nova 14 Pro", "weight": 9, "risk": 0.92},
    {"device": "novaflip", "display": "HUAWEI nova Flip", "weight": 6, "risk": 1.05},
    {"device": "pocket2", "display": "HUAWEI Pocket 2", "weight": 5, "risk": 1.12},
    {"device": "changxiang70pro", "display": "HUAWEI 畅享 70 Pro", "weight": 6, "risk": 0.82},
    {"device": "maimang30", "display": "HUAWEI 麦芒 30", "weight": 4, "risk": 0.80},
]

QUERY_TEMPLATES = [
    "{model}的{param}是多少？",
    "{model}支持{param}吗？",
    "帮我查一下{model}的{param}",
    "{model}在{param}方面是什么规格？",
    "我想确认{model}的{param}参数",
]

GOOD_ANSWERS = [
    "根据当前商品规格，{model}的{param}为{value}。",
    "{model}的{param}参数是{value}，以华为商城规格页为准。",
    "查到{model}的{param}为{value}。",
]

VALUES = {
    "操作系统": ["HarmonyOS 4.2", "HarmonyOS 5.0", "HarmonyOS NEXT"],
    "OTG": ["支持，最大输出电流以官方规格为准"],
    "反向充电": ["支持有线反向充电"],
    "卫星通信": ["支持北斗卫星消息", "部分版本支持卫星通信"],
    "屏幕尺寸": ["6.7英寸", "6.8英寸", "7.85英寸"],
    "电池容量": ["5000mAh", "5200mAh", "5600mAh"],
    "有线充电": ["最高支持88W超级快充", "最高支持100W超级快充"],
    "无线充电": ["最高支持50W无线超级快充", "支持无线充电"],
    "后置摄像头": ["5000万像素主摄组合", "超聚光主摄组合"],
    "变焦模式": ["支持光学变焦和数字变焦"],
    "防水等级": ["IP68", "IPX8", "不支持防水"],
    "机身重量": ["约209g", "约226g", "约198g"],
}


def rgb(hex_color: str) -> RGBColor:
    hex_color = hex_color.strip("#")
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


def truncate(text: object, limit: int) -> str:
    value = "" if text is None else str(text)
    return value if len(value) <= limit else value[: limit - 1] + "…"


def choose_l5() -> str:
    names = list(L5_SPECS)
    weights = [L5_SPECS[name]["weight"] for name in names]
    return random.choices(names, weights=weights, k=1)[0]


def choose_device() -> dict[str, object]:
    return random.choices(DEVICES, weights=[int(item["weight"]) for item in DEVICES], k=1)[0]


def answer_value(param: str) -> str:
    if param in VALUES:
        return random.choice(VALUES[param])
    generic = ["支持", "不支持", "以具体版本为准", "官方规格页已标注", "标准配置"]
    return random.choice(generic)


def load_eval_rows(input_path: Path, sheet_name: str | None = None) -> list[dict[str, object]]:
    wb = load_workbook(input_path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    headers = [str(value).strip() if value is not None else "" for value in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    index = {header: idx for idx, header in enumerate(headers)}
    missing = [column for column in BASE_COLUMNS if column not in index]
    if missing:
        raise ValueError(f"评测表缺少必要字段：{', '.join(missing)}")

    rows: list[dict[str, object]] = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(value is not None for value in values):
            continue
        row = {column: values[index[column]] for column in BASE_COLUMNS}
        row["case_id"] = str(row["case_id"] or "").strip()
        row["四级分类"] = str(row["四级分类"] or "").strip()
        row["query所属设备"] = str(row["query所属设备"] or "未标注设备").strip()
        row["五级分类"] = str(row["五级分类"] or "未标注五级分类").strip()
        row["六级分类"] = str(row["六级分类"] or "未标注六级分类").strip()
        row["query"] = str(row["query"] or "").strip()
        row["displayText"] = str(row["displayText"] or "").strip()
        row["备注"] = str(row["备注"] or "").strip()
        for column in ["人工综合评测结果", *DIMENSIONS]:
            row[column] = parse_flag(row[column], column)
        rows.append(row)
    if not rows:
        raise ValueError("评测表没有可用数据行")
    return rows


def parse_flag(value: object, column: str) -> int:
    if value is None or value == "":
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return 1 if int(value) == 1 else 0
    text = str(value).strip()
    if text in {"1", "是", "有问题", "问题", "true", "TRUE"}:
        return 1
    if text in {"0", "否", "无问题", "通过", "false", "FALSE"}:
        return 0
    raise ValueError(f"{column} 字段只支持 0/1 或等价文本，发现：{value!r}")


def generate_rows(total: int = TOTAL_CASES) -> list[dict[str, object]]:
    random.seed(RANDOM_SEED)
    rows: list[dict[str, object]] = []
    for idx in range(1, total + 1):
        l5 = choose_l5()
        l6 = random.choice(L5_SPECS[l5]["l6"])
        device_info = choose_device()
        device = str(device_info["device"])
        model = str(device_info["display"])
        query = random.choice(QUERY_TEMPLATES).format(model=model, param=l6)

        dim_probs = dict(L5_SPECS[l5]["dims"])
        dim_probs.update(L6_OVERRIDES.get(l6, {}))
        risk = float(device_info["risk"])
        dim_probs = {dim: min(prob * risk, 0.55) for dim, prob in dim_probs.items()}
        dims = {dim: 1 if random.random() < prob else 0 for dim, prob in dim_probs.items()}

        # Keep the synthetic data realistic: unrelated answers usually also lack useful information.
        if dims["相关性"]:
            dims["信息量"] = 1
        if dims["一致性"] and random.random() < 0.45:
            dims["正确性"] = 1

        issue = 1 if any(dims.values()) else 0
        value = answer_value(l6)
        if not issue:
            display = random.choice(GOOD_ANSWERS).format(model=model, param=l6, value=value)
            remark = ""
        elif dims["相关性"]:
            display = f"你好，关于{model}的活动优惠可以在商城页面查看。"
            remark = "回复偏离规格参数查询意图"
        elif dims["正确性"] and dims["信息量"]:
            display = f"{model}的{l6}为{random.choice(['不支持', '8GB', '4200mAh', 'IP53'])}，其他信息暂未查询到。"
            remark = "核心参数存在事实错误，且补充信息不足"
        elif dims["正确性"]:
            display = f"根据规格，{model}的{l6}为{random.choice(['不支持', '6.1英寸', '66W', 'IP53'])}。"
            remark = "回答与官方参数不一致"
        elif dims["信息量"]:
            display = f"你好，{model}的{l6}官方暂未提供相关信息。"
            remark = "无有效信息提供或回答不完整"
        else:
            display = f"{model}的{l6}支持，同时也可能不支持，需以官方页面为准。"
            remark = "同一回复中存在前后不一致表述"

        rows.append(
            {
                "case_id": f"parameter_{idx:05d}",
                "四级分类": "规格查询",
                "query所属设备": device,
                "五级分类": l5,
                "六级分类": l6,
                "query": query,
                "displayText": display,
                "备注": remark,
                "人工综合评测结果": issue,
                "相关性": dims["相关性"],
                "正确性": dims["正确性"],
                "信息量": dims["信息量"],
                "一致性": dims["一致性"],
            }
        )
    return rows


def compute_summary(rows: list[dict[str, object]]) -> dict[str, object]:
    total = len(rows)
    issue_count = sum(int(row["人工综合评测结果"]) for row in rows)
    pass_count = total - issue_count
    dim_counts = {dim: sum(int(row[dim]) for row in rows) for dim in DIMENSIONS}

    by_l5: dict[str, dict[str, object]] = {}
    for l5 in sorted({str(row["五级分类"]) for row in rows}):
        subset = [row for row in rows if row["五级分类"] == l5]
        issue = sum(int(row["人工综合评测结果"]) for row in subset)
        by_l5[l5] = {
            "total": len(subset),
            "issue": issue,
            "pass": len(subset) - issue,
            "accuracy": (len(subset) - issue) / len(subset),
            "issue_rate": issue / len(subset),
            "dims": {dim: sum(int(row[dim]) for row in subset) for dim in DIMENSIONS},
        }

    by_l6: dict[tuple[str, str], dict[str, object]] = {}
    groups: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        groups[(str(row["五级分类"]), str(row["六级分类"]))].append(row)
    for key, subset in groups.items():
        issue = sum(int(row["人工综合评测结果"]) for row in subset)
        by_l6[key] = {
            "total": len(subset),
            "issue": issue,
            "pass": len(subset) - issue,
            "accuracy": (len(subset) - issue) / len(subset),
            "issue_rate": issue / len(subset),
            "dims": {dim: sum(int(row[dim]) for row in subset) for dim in DIMENSIONS},
        }

    by_l6_device: dict[tuple[str, str], list[tuple[str, int]]] = {}
    for key, subset in groups.items():
        device_counter: Counter[str] = Counter()
        for row in subset:
            if int(row["人工综合评测结果"]) == 1:
                device_counter[str(row["query所属设备"])] += 1
        by_l6_device[key] = device_counter.most_common(3)

    by_device: dict[str, dict[str, object]] = {}
    for device in sorted({str(row["query所属设备"]) for row in rows}):
        subset = [row for row in rows if row["query所属设备"] == device]
        issue = sum(int(row["人工综合评测结果"]) for row in subset)
        by_device[device] = {
            "total": len(subset),
            "issue": issue,
            "pass": len(subset) - issue,
            "accuracy": (len(subset) - issue) / len(subset),
            "issue_rate": issue / len(subset),
            "dims": {dim: sum(int(row[dim]) for row in subset) for dim in DIMENSIONS},
        }

    examples = []
    seen_l5: set[str] = set()
    issue_rows = [row for row in rows if int(row["人工综合评测结果"]) == 1]
    issue_rows.sort(
        key=lambda row: (
            int(row["正确性"]),
            int(row["信息量"]),
            int(row["相关性"]),
            int(row["一致性"]),
        ),
        reverse=True,
    )
    for row in issue_rows:
        l5 = str(row["五级分类"])
        if l5 in seen_l5 and len(examples) < 4:
            continue
        examples.append(row)
        seen_l5.add(l5)
        if len(examples) == 4:
            break

    top_l6 = sorted(by_l6.items(), key=lambda item: item[1]["issue"], reverse=True)
    device_rank = sorted(by_device.items(), key=lambda item: item[1]["accuracy"])
    top_dim = max(dim_counts.items(), key=lambda item: item[1])[0]
    worst_l5 = min(by_l5.items(), key=lambda item: item[1]["accuracy"])[0]
    worst_device = device_rank[0][0]

    return {
        "total": total,
        "issue_count": issue_count,
        "pass_count": pass_count,
        "accuracy": pass_count / total,
        "issue_rate": issue_count / total,
        "dim_counts": dim_counts,
        "by_l5": by_l5,
        "by_l6": by_l6,
        "by_l6_device": by_l6_device,
        "by_device": by_device,
        "device_rank": device_rank,
        "top_l6": top_l6,
        "examples": examples,
        "top_dim": top_dim,
        "worst_l5": worst_l5,
        "worst_device": worst_device,
    }


def write_raw_workbook(rows: list[dict[str, object]], output_path: Path = RAW_XLSX) -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "mock_eval"
    headers = BASE_COLUMNS
    ws.append(headers)
    for row in rows:
        ws.append([row[key] for key in headers])
    style_header(ws, len(headers))
    set_widths(
        ws,
        {
            "A": 18,
            "B": 12,
            "C": 18,
            "D": 14,
            "E": 16,
            "F": 42,
            "G": 56,
            "H": 34,
            "I": 18,
            "J": 10,
            "K": 10,
            "L": 10,
            "M": 10,
        },
    )
    wb.save(output_path)


def write_summary_workbook(summary: dict[str, object], output_path: Path = SUMMARY_XLSX) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "总览"
    ws.append(["指标", "值"])
    ws.append(["总评测数", summary["total"]])
    ws.append(["通过数", summary["pass_count"]])
    ws.append(["问题数", summary["issue_count"]])
    ws.append(["综合准确率", summary["accuracy"]])
    ws.append(["综合问题率", summary["issue_rate"]])
    style_header(ws, 2)
    set_widths(ws, {"A": 22, "B": 18})

    ws = wb.create_sheet("维度问题")
    ws.append(["评测维度", "问题数", "问题率"])
    for dim, count in summary["dim_counts"].items():
        ws.append([dim, count, count / int(summary["total"])])
    style_header(ws, 3)
    set_widths(ws, {"A": 16, "B": 14, "C": 14})

    ws = wb.create_sheet("五级分类")
    ws.append(["五级分类", "总数", "通过数", "问题数", "准确率", "问题率", *DIMENSIONS])
    for l5, stat in sorted(summary["by_l5"].items(), key=lambda item: item[1]["accuracy"]):
        ws.append(
            [
                l5,
                stat["total"],
                stat["pass"],
                stat["issue"],
                stat["accuracy"],
                stat["issue_rate"],
                *[stat["dims"][dim] for dim in DIMENSIONS],
            ]
        )
    style_header(ws, 10)
    set_widths(ws, {"A": 16, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12})

    ws = wb.create_sheet("设备维度")
    ws.append(["query所属设备", "总数", "通过数", "问题数", "准确率", "问题率", *DIMENSIONS])
    for device, stat in summary["device_rank"]:
        ws.append(
            [
                device,
                stat["total"],
                stat["pass"],
                stat["issue"],
                stat["accuracy"],
                stat["issue_rate"],
                *[stat["dims"][dim] for dim in DIMENSIONS],
            ]
        )
    style_header(ws, 10)
    set_widths(ws, {"A": 18, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12})

    ws = wb.create_sheet("五级维度矩阵")
    ws.append(["五级分类", *DIMENSIONS])
    for l5, stat in sorted(summary["by_l5"].items(), key=lambda item: item[1]["issue"], reverse=True):
        ws.append([l5, *[stat["dims"][dim] / stat["total"] for dim in DIMENSIONS]])
    style_header(ws, 5)
    set_widths(ws, {"A": 16, "B": 12, "C": 12, "D": 12, "E": 12})

    ws = wb.create_sheet("六级Top")
    ws.append(["五级分类", "六级分类", "总数", "问题数", "准确率", "主要问题维度"])
    for (l5, l6), stat in summary["top_l6"][:30]:
        top_dim = max(stat["dims"].items(), key=lambda item: item[1])[0]
        ws.append([l5, l6, stat["total"], stat["issue"], stat["accuracy"], top_dim])
    style_header(ws, 6)
    set_widths(ws, {"A": 16, "B": 18, "C": 12, "D": 12, "E": 12, "F": 16})

    ws = wb.create_sheet("典型案例")
    ws.append(["case_id", "query所属设备", "五级分类", "六级分类", "query", "displayText", "备注", *DIMENSIONS])
    for row in summary["examples"]:
        ws.append(
            [
                row["case_id"],
                row["query所属设备"],
                row["五级分类"],
                row["六级分类"],
                row["query"],
                row["displayText"],
                row["备注"],
                *[row[dim] for dim in DIMENSIONS],
            ]
        )
    style_header(ws, 11)
    set_widths(ws, {"A": 18, "B": 18, "C": 14, "D": 16, "E": 44, "F": 56, "G": 34})

    ws = wb.create_sheet("优先级")
    ws.append(["优先级", "参数", "重点设备", "问题数", "准确率", "主要问题", "建议动作"])
    for row in build_action_items(summary):
        ws.append(
            [
                row["priority"],
                row["param"],
                row["devices"],
                row["issue"],
                row["accuracy"],
                row["top_dim"],
                row["action"],
            ]
        )
    style_header(ws, 7)
    set_widths(ws, {"A": 12, "B": 22, "C": 24, "D": 12, "E": 12, "F": 14, "G": 54})

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if isinstance(cell.value, float):
                    cell.number_format = "0.0%"
    wb.save(output_path)


def style_header(ws, cols: int) -> None:
    fill = PatternFill("solid", fgColor=COLORS["red"])
    font = Font(color=COLORS["white"], bold=True)
    for cell in ws[1][:cols]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def set_widths(ws, widths: dict[str, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def build_action_items(summary: dict[str, object]) -> list[dict[str, object]]:
    actions = []
    for (l5, l6), stat in summary["top_l6"][:8]:
        top_dim = max(stat["dims"].items(), key=lambda item: item[1])[0]
        device_hits = summary["by_l6_device"].get((l5, l6), [])
        devices = "、".join(device for device, _ in device_hits) or "-"
        issue = int(stat["issue"])
        accuracy = float(stat["accuracy"])
        if issue >= 45 and (top_dim == "正确性" or accuracy < 0.80):
            priority = "P0"
        elif issue >= 35 or top_dim in {"正确性", "信息量"}:
            priority = "P1"
        else:
            priority = "P2"
        action = {
            "正确性": "补齐参数事实表，强化型号-参数精确匹配与单位校验",
            "信息量": "补充字段召回策略，约束答案必须给出完整规格值",
            "相关性": "优化规格查询意图识别，扩展参数别名和同义表达",
            "一致性": "增加答案前后一致性校验，统一支持/不支持表述",
        }[top_dim]
        actions.append(
            {
                "priority": priority,
                "param": f"{l5} / {l6}",
                "devices": devices,
                "issue": issue,
                "accuracy": accuracy,
                "top_dim": top_dim,
                "action": action,
            }
        )
    return actions


def add_text(
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    size: int = 16,
    color: str = "dark",
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
) -> object:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = Inches(0.03)
    tf.margin_right = Inches(0.03)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    for idx, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        p.font.name = FONT
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = rgb(COLORS[color])
    return box


def add_rect(slide, x: float, y: float, w: float, h: float, fill: str, line: str | None = None) -> object:
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(COLORS[fill])
    if line:
        shape.line.color.rgb = rgb(COLORS[line])
        shape.line.width = Pt(0.8)
    else:
        shape.line.fill.background()
    return shape


def add_slide_header(slide, title: str, subtitle: str | None = None, label: str = "模拟数据") -> None:
    add_text(slide, title, 0.55, 0.35, 8.2, 0.45, 21, "dark", True)
    if subtitle:
        add_text(slide, subtitle, 0.58, 0.82, 8.7, 0.28, 9, "gray")
    add_rect(slide, 0.55, 1.12, 1.15, 0.035, "red")
    add_text(slide, label, 11.45, 0.42, 1.2, 0.28, 9, "red", True, PP_ALIGN.RIGHT)


def add_footer(slide, page: int) -> None:
    add_text(slide, f"{page:02d}", 12.55, 7.05, 0.35, 0.2, 8, "gray", False, PP_ALIGN.RIGHT)


def add_kpi_card(slide, x: float, y: float, w: float, h: float, label: str, value: str, sub: str, accent: str) -> None:
    add_rect(slide, x, y, w, h, "white", "light_gray")
    add_rect(slide, x, y, 0.06, h, accent)
    add_text(slide, label, x + 0.22, y + 0.18, w - 0.35, 0.25, 9, "gray", True)
    add_text(slide, value, x + 0.22, y + 0.46, w - 0.35, 0.45, 25, accent, True)
    add_text(slide, sub, x + 0.22, y + 1.03, w - 0.35, 0.25, 8, "gray")


def add_manual_bar_list(
    slide,
    items: list[tuple[str, float, str]],
    x: float,
    y: float,
    w: float,
    row_h: float,
    max_value: float = 1.0,
    value_fmt: str = "{:.1%}",
    label_w: float = 2.0,
    color: str = "red",
    limit: int | None = None,
) -> None:
    visible = items[:limit] if limit else items
    bar_x = x + label_w
    bar_w = w - label_w - 0.85
    for idx, (label, value, tag) in enumerate(visible):
        yy = y + idx * row_h
        add_text(slide, truncate(label, 16), x, yy + 0.02, label_w - 0.05, row_h - 0.02, 8, "dark")
        add_rect(slide, bar_x, yy + 0.07, bar_w, row_h * 0.38, "light_gray")
        fill_w = max(0.02, bar_w * min(value / max_value, 1))
        add_rect(slide, bar_x, yy + 0.07, fill_w, row_h * 0.38, color)
        add_text(slide, value_fmt.format(value), bar_x + bar_w + 0.08, yy + 0.02, 0.75, row_h - 0.02, 8, "dark")
        if tag:
            add_text(slide, tag, bar_x + bar_w - 0.95, yy + 0.02, 0.9, row_h - 0.02, 7, "gray", False, PP_ALIGN.RIGHT)


def heat_color(rate: float, max_rate: float) -> str:
    if max_rate <= 0:
        return "FFFFFF"
    ratio = min(rate / max_rate, 1.0)
    if ratio < 0.25:
        return "FEE2E2"
    if ratio < 0.5:
        return "FCA5A5"
    if ratio < 0.75:
        return "F87171"
    return "C7000B"


def add_heatmap(slide, summary: dict[str, object], x: float, y: float, w: float, h: float) -> None:
    by_l5 = summary["by_l5"]
    l5_rows = sorted(by_l5.items(), key=lambda item: item[1]["issue_rate"], reverse=True)
    label_w = 2.1
    cell_w = (w - label_w) / len(DIMENSIONS)
    row_h = h / (len(l5_rows) + 1)
    add_rect(slide, x, y, w, row_h, "dark")
    add_text(slide, "五级分类", x + 0.06, y + 0.05, label_w - 0.12, row_h - 0.05, 8, "white", True)
    for col, dim in enumerate(DIMENSIONS):
        add_text(slide, dim, x + label_w + col * cell_w, y + 0.05, cell_w, row_h - 0.05, 8, "white", True, PP_ALIGN.CENTER)

    max_rate = max(stat["dims"][dim] / stat["total"] for _, stat in l5_rows for dim in DIMENSIONS)
    for row_idx, (l5, stat) in enumerate(l5_rows, start=1):
        yy = y + row_idx * row_h
        add_rect(slide, x, yy, label_w, row_h, "white", "light_gray")
        add_text(slide, truncate(l5, 12), x + 0.06, yy + 0.05, label_w - 0.12, row_h - 0.04, 8, "dark")
        for col, dim in enumerate(DIMENSIONS):
            rate = stat["dims"][dim] / stat["total"]
            color_hex = heat_color(rate, max_rate)
            shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(x + label_w + col * cell_w),
                Inches(yy),
                Inches(cell_w),
                Inches(row_h),
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(color_hex)
            shape.line.color.rgb = rgb(COLORS["white"])
            add_text(
                slide,
                f"{rate:.1%}",
                x + label_w + col * cell_w,
                yy + 0.05,
                cell_w,
                row_h - 0.04,
                8,
                "white" if rate / max_rate >= 0.75 else "dark",
                False,
                PP_ALIGN.CENTER,
            )


def create_ppt(
    summary: dict[str, object],
    *,
    output_path: Path = PPTX_PATH,
    product_name: str = "手机",
    subtitle: str | None = None,
    mock: bool = True,
) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    label = "模拟数据" if mock else "正式数据"

    # 1. Cover
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_rect(slide, 0, 0, 13.333, 0.18, "red")
    add_text(slide, f"小艺{product_name}全量参数\n知识覆盖评测报告", 0.85, 1.45, 7.2, 1.5, 34, "dark", True)
    cover_subtitle = subtitle or f"华为商城在售{product_name}全系列规格查询"
    if mock:
        cover_subtitle = f"{cover_subtitle} | PPT预览版（模拟数据）"
    add_text(slide, cover_subtitle, 0.9, 3.25, 7.2, 0.35, 14, "gray")
    add_kpi_card(slide, 8.8, 1.55, 3.25, 1.35, "样本规模", f"{summary['total']:,}", f"覆盖 {len(summary['by_device'])} 个设备", "red")
    add_kpi_card(slide, 8.8, 3.15, 3.25, 1.35, "综合准确率", f"{summary['accuracy']:.1%}", "人工综合评测结果=0", "green")
    add_kpi_card(slide, 8.8, 4.75, 3.25, 1.35, "主要问题", str(summary["top_dim"]), "四维度中命中最多", "orange")
    add_text(slide, f"生成日期：{date.today().isoformat()}", 0.9, 6.7, 3.0, 0.25, 9, "gray")
    add_footer(slide, 1)

    # 2. Summary dashboard
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "综合看板", "合并展示核心结论、整体准确率和四维度问题分布。", label)
    add_kpi_card(slide, 0.6, 1.28, 2.45, 1.08, "总评测数", f"{summary['total']:,}", f"设备 {len(summary['by_device'])} 个", "red")
    add_kpi_card(slide, 3.22, 1.28, 2.45, 1.08, "通过数", f"{summary['pass_count']:,}", "无问题 case", "green")
    add_kpi_card(slide, 5.84, 1.28, 2.45, 1.08, "问题数", f"{summary['issue_count']:,}", "至少一个维度命中", "orange")
    add_kpi_card(slide, 8.46, 1.28, 2.45, 1.08, "综合准确率", f"{summary['accuracy']:.1%}", "通过数 / 总评测数", "blue")
    add_kpi_card(slide, 11.08, 1.28, 1.6, 1.08, "最低设备", str(summary["worst_device"]), "按准确率", "red")

    chart_data = ChartData()
    chart_data.categories = ["通过", "问题"]
    chart_data.add_series("综合结果", (summary["pass_count"], summary["issue_count"]))
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.82), Inches(2.65), Inches(3.2), Inches(3.2), chart_data).chart
    chart.has_legend = False
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.number_format = "0%"
    chart.plots[0].data_labels.position = XL_LABEL_POSITION.BEST_FIT
    chart.series[0].points[0].format.fill.solid()
    chart.series[0].points[0].format.fill.fore_color.rgb = rgb(COLORS["green"])
    chart.series[0].points[1].format.fill.solid()
    chart.series[0].points[1].format.fill.fore_color.rgb = rgb(COLORS["red"])
    add_text(slide, f"{summary['accuracy']:.1%}", 1.48, 3.74, 1.85, 0.4, 23, "green", True, PP_ALIGN.CENTER)
    add_text(slide, "综合准确率", 1.58, 4.2, 1.65, 0.22, 8, "gray", False, PP_ALIGN.CENTER)

    dim_items = [
        (dim, summary["dim_counts"][dim] / summary["total"], f"{summary['dim_counts'][dim]:,}条")
        for dim in DIMENSIONS
    ]
    add_text(slide, "四维度问题率", 4.45, 2.72, 2.2, 0.25, 12, "dark", True)
    add_manual_bar_list(slide, dim_items, 4.45, 3.1, 4.1, 0.48, max(v for _, v, _ in dim_items), "{:.1%}", 1.05, "red")

    insights = [
        f"整体准确率 {summary['accuracy']:.1%}，问题 case {summary['issue_count']:,} 条。",
        f"主要短板集中在「{summary['top_dim']}」，应优先处理该维度的高频参数。",
        f"设备维度最低为「{summary['worst_device']}」，五级分类最低为「{summary['worst_l5']}」。",
    ]
    for idx, text in enumerate(insights):
        y = 3.0 + idx * 0.72
        add_rect(slide, 9.15, y, 3.35, 0.5, "white", "light_gray")
        add_rect(slide, 9.15, y, 0.055, 0.5, "red" if idx == 0 else "orange")
        add_text(slide, text, 9.35, y + 0.12, 2.85, 0.22, 9, "dark", idx == 0)
    add_text(slide, "判断方式：综合结果=1 表示四个维度中至少一个存在问题；四维度命中率可重叠，不做互斥占比。", 0.88, 6.38, 11.6, 0.25, 8, "gray")
    add_footer(slide, 2)

    # 3. Device and L5 ranking
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "设备与参数聚类表现", "按 query所属设备 定位具体机型的知识覆盖差异。", label)
    device_items = [
        (device, stat["accuracy"], f"n={stat['total']}")
        for device, stat in summary["device_rank"]
    ]
    add_text(slide, "设备准确率排行", 0.8, 1.28, 2.2, 0.25, 12, "dark", True)
    add_manual_bar_list(slide, device_items, 0.8, 1.68, 5.55, 0.29, 1.0, "{:.1%}", 1.55, "red", limit=16)
    l5_items = [
        (l5, stat["accuracy"], f"n={stat['total']}")
        for l5, stat in sorted(summary["by_l5"].items(), key=lambda item: item[1]["accuracy"])
    ]
    add_text(slide, "五级分类准确率排行", 7.0, 1.28, 2.6, 0.25, 12, "dark", True)
    add_manual_bar_list(slide, l5_items, 7.0, 1.68, 5.15, 0.38, 1.0, "{:.1%}", 1.45, "blue")
    worst = l5_items[0]
    worst_device = device_items[0]
    add_text(slide, f"最低设备：{worst_device[0]}（{worst_device[1]:.1%}）", 0.8, 6.65, 4.0, 0.24, 9, "red", True)
    add_text(slide, f"最低参数聚类：{worst[0]}（{worst[1]:.1%}）", 7.0, 6.65, 4.2, 0.24, 9, "red", True)
    add_footer(slide, 3)

    # 4. Heatmap
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "五级分类 × 问题维度矩阵", "颜色越深，表示该分类在该评测维度上的问题率越高。", label)
    add_heatmap(slide, summary, 0.8, 1.35, 11.65, 5.4)
    add_footer(slide, 4)

    # 5. Top L6
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "六级分类问题 Top 20", "具体参数级别的问题排序，用来形成开发优化队列。", label)
    top20 = summary["top_l6"][:20]
    total_issue = summary["issue_count"]
    cumulative = 0
    items = []
    max_issue = max(stat["issue"] for _, stat in top20)
    for (l5, l6), stat in top20:
        cumulative += stat["issue"]
        label = f"{l6}"
        tag = f"{cumulative / total_issue:.0%}"
        items.append((label, stat["issue"] / max_issue, tag))
    add_manual_bar_list(slide, items, 0.75, 1.28, 8.25, 0.255, 1.0, "{:.0%}", 1.55, "red")
    add_text(slide, "右侧百分比为累计覆盖问题占比", 6.25, 6.65, 2.5, 0.22, 8, "gray", False, PP_ALIGN.RIGHT)
    add_rect(slide, 9.45, 1.35, 3.05, 4.9, "white", "light_gray")
    top_coverage = sum(stat["issue"] for _, stat in top20) / total_issue
    add_text(slide, "Top 20 覆盖", 9.75, 1.75, 1.4, 0.3, 10, "gray", True)
    add_text(slide, f"{top_coverage:.1%}", 9.75, 2.12, 1.7, 0.45, 28, "red", True)
    add_text(slide, "的问题 case", 11.25, 2.28, 0.9, 0.25, 10, "gray")
    top_param = top20[0]
    add_text(slide, "最高频参数", 9.75, 3.05, 1.4, 0.3, 10, "gray", True)
    add_text(slide, f"{top_param[0][0]} / {top_param[0][1]}", 9.75, 3.42, 2.35, 0.45, 15, "dark", True)
    add_text(slide, f"问题数 {top_param[1]['issue']}，准确率 {top_param[1]['accuracy']:.1%}", 9.75, 4.02, 2.25, 0.25, 9, "gray")
    top_devices = "、".join(device for device, _ in summary["by_l6_device"].get(top_param[0], []))
    add_text(slide, f"高发设备：{top_devices or '-'}", 9.75, 4.38, 2.25, 0.25, 9, "gray")
    add_text(slide, "开发优先处理高频参数，能最快拉动整体准确率。", 9.75, 5.15, 2.25, 0.55, 11, "dark")
    add_footer(slide, 5)

    # 6. Examples
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "典型错误案例", "从高频问题池中抽样，帮助开发理解判错原因和修复方向。", label)
    positions = [(0.75, 1.35), (6.8, 1.35), (0.75, 4.05), (6.8, 4.05)]
    for idx, (row, (x, y)) in enumerate(zip(summary["examples"], positions), start=1):
        add_rect(slide, x, y, 5.65, 2.25, "white", "light_gray")
        dims = " / ".join([dim for dim in DIMENSIONS if row[dim]])
        add_text(slide, f"案例 {idx}｜{row['query所属设备']}｜{row['五级分类']} / {row['六级分类']}", x + 0.2, y + 0.18, 3.95, 0.25, 10, "red", True)
        add_text(slide, dims, x + 4.1, y + 0.18, 1.25, 0.25, 9, "orange", True, PP_ALIGN.RIGHT)
        add_text(slide, f"Q：{truncate(row['query'], 42)}", x + 0.2, y + 0.55, 5.15, 0.35, 10, "dark", True)
        add_text(slide, f"A：{truncate(row['displayText'], 58)}", x + 0.2, y + 0.98, 5.15, 0.45, 9, "gray")
        add_text(slide, f"备注：{truncate(row['备注'], 38)}", x + 0.2, y + 1.62, 5.15, 0.35, 10, "dark")
    add_footer(slide, 6)

    # 7. Actions
    slide = prs.slides.add_slide(blank)
    add_rect(slide, 0, 0, 13.333, 7.5, "bg")
    add_slide_header(slide, "优化优先级与行动清单", "把参数问题和重点设备一起给到开发，优先拉动整体准确率。", label)
    actions = build_action_items(summary)
    x = 0.65
    y = 1.35
    widths = [0.68, 2.0, 1.65, 0.72, 0.78, 0.85, 5.05]
    headers = ["优先级", "参数", "重点设备", "问题数", "准确率", "主问题", "建议动作"]
    add_rect(slide, x, y, sum(widths), 0.38, "dark")
    xx = x
    for width, header in zip(widths, headers):
        add_text(slide, header, xx + 0.05, y + 0.08, width - 0.1, 0.18, 8, "white", True, PP_ALIGN.CENTER)
        xx += width
    for idx, item in enumerate(actions):
        yy = y + 0.38 + idx * 0.54
        add_rect(slide, x, yy, sum(widths), 0.54, "white", "light_gray")
        xx = x
        values = [
            item["priority"],
            item["param"],
            item["devices"],
            str(item["issue"]),
            f"{item['accuracy']:.1%}",
            item["top_dim"],
            item["action"],
        ]
        colors = ["red" if item["priority"] == "P0" else "orange", "dark", "gray", "dark", "dark", "red", "dark"]
        for width, value, color in zip(widths, values, colors):
            add_text(slide, truncate(value, 33), xx + 0.06, yy + 0.12, width - 0.12, 0.23, 8, color, item["priority"] == "P0" and value == item["priority"])
            xx += width
    add_text(slide, "优先级规则：优先处理问题量高、正确性问题突出、准确率低的具体参数。", 0.75, 6.38, 8.2, 0.25, 9, "gray")
    add_footer(slide, 7)

    prs.save(output_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate Xiaoyi product-parameter evaluation PPT.")
    parser.add_argument("--input", type=Path, help="评测结果 Excel。缺省时生成模拟数据。")
    parser.add_argument("--sheet", help="Excel sheet 名。缺省读取第一个 sheet。")
    parser.add_argument("--product", default="手机", help="报告品类名，例如 手机、平板、耳机。")
    parser.add_argument("--subtitle", help="封面副标题。")
    parser.add_argument("--output", type=Path, help="PPT 输出路径。")
    parser.add_argument("--summary-output", type=Path, help="统计汇总 Excel 输出路径。")
    parser.add_argument("--mock-output", type=Path, help="模拟原始评测 Excel 输出路径；仅模拟模式使用。")
    parser.add_argument("--mock-count", type=int, default=TOTAL_CASES, help="模拟数据条数。")
    return parser


def default_output_path(product_name: str, mock: bool) -> Path:
    suffix = "_模拟数据" if mock else ""
    return ROOT / f"小艺{product_name}全量参数评测报告{suffix}.pptx"


def default_summary_path(product_name: str, mock: bool) -> Path:
    suffix = "_模拟数据" if mock else ""
    return REPORT_DIR / f"{product_name}全量参数评测统计{suffix}.xlsx"


def default_mock_raw_path(product_name: str) -> Path:
    return REPORT_DIR / f"{product_name}全量参数评测_模拟数据.xlsx"


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    mock = args.input is None
    if mock:
        rows = generate_rows(args.mock_count)
    else:
        rows = load_eval_rows(args.input, args.sheet)

    summary = compute_summary(rows)
    pptx_path = args.output or default_output_path(args.product, mock)
    summary_path = args.summary_output or default_summary_path(args.product, mock)
    write_summary_workbook(summary, summary_path)
    if mock:
        raw_path = args.mock_output or default_mock_raw_path(args.product)
        write_raw_workbook(rows, raw_path)
        print(f"raw={raw_path}")
    print(f"summary={summary_path}")
    create_ppt(
        summary,
        output_path=pptx_path,
        product_name=args.product,
        subtitle=args.subtitle,
        mock=mock,
    )
    print(f"pptx={pptx_path}")
    print(f"total={summary['total']} accuracy={summary['accuracy']:.2%} issue={summary['issue_count']}")


if __name__ == "__main__":
    main()

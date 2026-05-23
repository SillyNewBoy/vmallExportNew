# PPT 生成脚本使用说明

## 脚本说明

脚本路径：

```powershell
D:\01_code_workspace\vmallExportNew\report\generate_mock_phone_eval_ppt.py
```

该脚本用于根据小艺参数问答人工评测 Excel 自动生成 PPT 报告和统计汇总 Excel。

脚本不调用模型、不访问网络，执行过程完全在本地完成：

```text
评测 Excel -> 统计聚合 -> 生成统计 Excel -> 生成 PPT
```

## 依赖环境

需要本机 Python 环境已安装以下依赖：

```powershell
python -m pip install -e .[report]
```

也可以单独安装：

```powershell
python -m pip install openpyxl python-pptx
```

## 使用方式

脚本不接收命令行参数。使用时只需要修改脚本顶部的 `CONFIG` 配置区，然后直接执行脚本。

```powershell
python D:\01_code_workspace\vmallExportNew\report\generate_mock_phone_eval_ppt.py
```

## CONFIG 配置说明

打开 `generate_mock_phone_eval_ppt.py`，只需要修改这 3 个字段：

```python
CONFIG = {
    "PRODUCT_NAME": "手机",
    "INPUT_XLSX": REPORT_DIR / "手机全量参数评测.xlsx",
    "OUTPUT_PPTX": ROOT / "小艺手机全量参数评测报告.pptx",
}
```

| 配置项 | 说明 |
|---|---|
| `PRODUCT_NAME` | 报告品类名，例如手机、平板、耳机 |
| `INPUT_XLSX` | 真实评测 Excel 路径 |
| `OUTPUT_PPTX` | PPT 输出路径 |

统计汇总 Excel 不需要手动配置，会根据 `OUTPUT_PPTX` 自动生成在同一目录。例如：

```text
OUTPUT_PPTX = D:\01_code_workspace\vmallExportNew\小艺手机全量参数评测报告.pptx
统计表 = D:\01_code_workspace\vmallExportNew\小艺手机全量参数评测报告_统计.xlsx
```

## 示例

### 手机报告

```python
CONFIG = {
    "PRODUCT_NAME": "手机",
    "INPUT_XLSX": REPORT_DIR / "手机全量参数评测.xlsx",
    "OUTPUT_PPTX": ROOT / "小艺手机全量参数评测报告.pptx",
}
```

### 平板报告

只要字段保持一致，平板、耳机、汽车等品类也可以复用同一个脚本。

```python
CONFIG = {
    "PRODUCT_NAME": "平板",
    "INPUT_XLSX": REPORT_DIR / "平板全量参数评测.xlsx",
    "OUTPUT_PPTX": ROOT / "小艺平板全量参数评测报告.pptx",
}
```

## 输入表字段要求

正式评测 Excel 需要包含以下固定字段：

| 字段 | 说明 |
|---|---|
| case_id | query 编号 |
| 四级分类 | 测试场景分类，例如规格查询 |
| query所属设备 | query 对应设备/机型，例如 pura80、pura90、matex6 |
| 五级分类 | 参数聚类，例如主体、屏幕、传输功能 |
| 六级分类 | 具体参数，例如操作系统、OTG、屏幕尺寸 |
| query | 测试问题 |
| displayText | 小艺模型回复 |
| 备注 | 人工评测备注 |
| 人工综合评测结果 | 0 表示无问题，1 表示存在问题 |
| 相关性 | 0 表示无问题，1 表示存在问题 |
| 正确性 | 0 表示无问题，1 表示存在问题 |
| 信息量 | 0 表示无问题，1 表示存在问题 |
| 一致性 | 0 表示无问题，1 表示存在问题 |

四个评测维度可以同时命中，因此 PPT 中展示的是各维度问题率，不会把它们做成互斥占比。

## 输出内容

PPT 当前为 7 页：

1. 封面
2. 综合看板
3. 设备与参数聚类表现
4. 五级分类 × 问题维度矩阵
5. 六级分类问题 Top 20
6. 典型错误案例
7. 优化优先级与行动清单

统计汇总 Excel 会包含：

- 总览
- 维度问题
- 五级分类
- 设备维度
- 五级维度矩阵
- 六级 Top
- 典型案例
- 优先级

## 注意事项

- 如果 PPT 文件正在被 PowerPoint 打开，脚本覆盖同名文件时会报 `Permission denied`。关闭 PPT 后重新执行，或修改 `OUTPUT_PPTX` 为新的文件名。
- 脚本不会改动原始评测 Excel，只会读取数据并生成新的统计文件和 PPT。
- `query所属设备` 是设备维度分析的来源，建议保持设备命名稳定，例如 `pura80`、`pura90pro`、`matex6`。

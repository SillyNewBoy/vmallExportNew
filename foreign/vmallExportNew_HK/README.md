# vmallExportNew

基于华为商城 VMALL 官方接口和商品详情页内嵌 JSON 的规格参数导出工具。
同时支持从鸿蒙智行官网 `hima.auto` 抓取汽车配置参数，以及从华为消费者官网
`consumer.huawei.com`（如马来西亚站 `https://consumer.huawei.com/my/`）抓取全量产品规格参数。

项目目标：

- 按品类抓取商品列表
- 同一商品只保留一个代表配置，不重复抓取全部 SKU
- 以 `品类 -> 系列 -> 产品` 的层级保存 JSON
- 支持 VMALL 数码品类和鸿蒙智行汽车品类统一导出
- 支持 consumer.huawei.com 区域站点全量规格参数抓取，并额外导出 Excel/CSV
- 每次拉取会在品类 JSON 顶层记录 `pull_time`，格式为 `yymmdd`
- 自动对比同品类上一次拉取结果，发现新增商品时生成新增部分 JSON
- 支持 macOS 和 Windows
- 内置限速、重试、退避和日志输出

## 运行环境

必需条件：

- Python `3.9+`
- `pip`，建议升级到最新版本
- 能访问 `https://www.vmall.com/`、`https://openapi.vmall.com/` 和 `https://hima.auto/`

运行时依赖：

- `requests>=2.31.0`
- `beautifulsoup4>=4.12.0`
- `openpyxl>=3.1.0`

可选开发依赖：

- `pytest>=8.0`

## 项目结构

```text
vmallExportNew/
├── README.md
├── requirements.txt
├── run.py
├── pyproject.toml
├── report/
│   ├── generate_mock_phone_eval_ppt.py  # 小艺参数评测 PPT/统计表生成脚本
│   └── PPT生成脚本使用说明.md             # 评测报告生成说明
├── src/
│   ├── out/                 # 默认导出目录，运行时自动创建
│   ├── diff/                # 默认新增商品 diff 目录，运行时按需自动创建
│   └── vmall_export_new/
│       ├── __init__.py
│       ├── __main__.py
│       ├── categories.py
│       ├── cli.py
│       ├── client.py
│       ├── consumer_categories.py  # consumer.huawei.com 品类与系列配置
│       ├── consumer_client.py      # consumer.huawei.com HTTP 客户端
│       ├── consumer_exporter.py    # consumer.huawei.com 规格抓取与解析
│       ├── excel_export.py         # Excel/CSV 导出
│       ├── exporter.py
│       ├── hima_exporter.py
│       ├── json_export.py
│       └── models.py
└── tests/
    ├── test_cli.py
    └── test_exporter.py
```

## 推荐使用方式

推荐直接运行根目录的 `run.py`，这样不依赖平台上的可执行脚本名，macOS 和 Windows 都一致。

### 1. 创建虚拟环境

macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Windows PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Windows CMD:

```cmd
py -3 -m venv .venv
.\.venv\Scripts\activate.bat
```

### 2. 安装依赖

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 3. 运行

查看内置品类：

```bash
python run.py --list-categories
```

导出汽车配置参数：

```bash
python run.py --category cars
```

导出手机参数。未指定目录时，默认写入 `src/out`，如果发现新增商品，新增部分默认写入 `src/diff`：

```bash
python run.py --category phones
```

导出全部内置品类：

```bash
python run.py --all
```

限制每个品类最多抓取 10 个商品：

```bash
python run.py --category tablets --limit 10
```

手动指定输出目录和新增 diff 目录：

```bash
python run.py --category phones --output-dir output --diff-dir diff
```

开启调试日志并写入文件：

```bash
python run.py --category phones --log-level DEBUG --log-file vmall.log
```

## 可选安装方式

如果你希望安装成命令行工具，也可以：

```bash
python -m pip install .
```

安装后可以使用以下任一方式运行：

```bash
vmall-export --category phones
python -m vmall_export_new --category phones
```

## 常用参数

- `--site <vmall|consumer>`：数据源站点，默认 `vmall`；`consumer` 表示抓取华为消费者官网（`consumer.huawei.com`）
- `--region <code>`：consumer 站点区域代码，默认 `my`（马来西亚站），例如 `cn`、`en`
- `--excel`：同时导出 Excel（`.xlsx`）
- `--csv`：同时导出 CSV
- `--list-categories`：列出内置品类
- `--category <key>`：导出单个品类，例如 `phones`
- `--all`：导出全部内置品类
- `--output-dir <dir>`：完整品类 JSON 输出目录，默认是当前工程的 `src/out`
- `--diff-dir <dir>`：新增商品 JSON 输出目录，默认是当前工程的 `src/diff`
- `--limit <n>`：限制每个品类抓取的商品数量
- `--max-pages <n>`：限制搜索结果翻页数（仅 VMALL 站点）
- `--page-size <n>`：每页搜索结果数量（仅 VMALL 站点）
- `--log-level <level>`：日志级别，支持 `DEBUG/INFO/WARNING/ERROR`
- `--log-file <path>`：把日志同时写入文件
- `--request-interval <sec>`：请求间隔秒数，默认 `0.8`
- `--max-retries <n>`：请求失败最大重试次数，默认 `4`
- `--backoff-factor <sec>`：重试退避基数，默认 `1.5`

## 华为消费者官网（consumer.huawei.com）全量参数导出

通过 `--site consumer` 可以抓取消费者官网区域站点的**全量产品规格参数**，例如马来西亚站
`https://consumer.huawei.com/my/`。

工作原理：

1. 抓取品类列表页（如 `/my/phones/`）发现全部产品页面
2. 结合站点 sitemap 中的 `specs/` 链接补充发现，避免遗漏
3. 逐个抓取产品的 `specs/` 规格页（纯静态 HTML），解析 `large-accordion`
   结构中的全部参数分组、参数名、参数值和备注
4. 输出与 VMALL 相同的 `品类 -> 系列 -> 产品` JSON 层级，并**默认同时导出
   Excel 和 CSV**（每个品类一个文件，Excel 含“全部参数”汇总表 + 每个系列一个 sheet）

使用方式：

```bash
# 抓取马来西亚站全部品类的全量规格参数（JSON + Excel + CSV）
python run.py --site consumer --all

# 只抓取手机品类
python run.py --site consumer --category phones

# 抓取其他区域站点，例如墨西哥站、中国区
python run.py --site consumer --region mx --all
python run.py --site consumer --region cn --category phones

# 限制每个品类抓取数量（调试用）
python run.py --site consumer --category phones --limit 5

# 只输出 JSON + CSV，不输出 Excel
python run.py --site consumer --category phones --csv
```

consumer 模式输出目录按区域隔离，避免不同站点互相覆盖：

- 默认输出目录：`src/out_{region}/`，例如 `src/out_my/`、`src/out_mx/`
- 默认 diff 目录：`src/diff_{region}/`
- 显式传入 `--output-dir` / `--diff-dir` 时以传入值为准
- VMALL 模式默认值不变，仍为 `src/out/`、`src/diff/`

consumer 站点内置品类：

- `phones`：手机（phones）
- `tablets`：平板电脑（tablets）
- `wearables`：穿戴设备（wearables）
- `laptops`：笔记本电脑（laptops）
- `headphones`：耳机（headphones）
- `speakers`：音箱（speakers）
- `routers`：路由器（routers）
- `monitors`：显示器（monitors）
- `printers`：打印机（printers）
- `visions`：智慧屏（visions）

说明：

- consumer 站点没有 SKU 概念，`product_id` 使用产品 URL slug
- 产品系列按产品名规则自动归类（Mate / Pura / nova / MateBook / WATCH GT 等），
  规则见 `consumer_categories.py` 的 `CONSUMER_SERIES_RULES`
- 没有 `specs/` 页面的产品（如活动页）会自动跳过并记录日志
- 同样支持 `pull_time` 与新增商品 diff 机制
- Excel 和 CSV 为平铺结构，每行一个参数：`品类/系列/产品/参数分组/参数名/参数值/备注/链接`


## 当前内置品类

- `phones`：手机
- `tablets`：平板电脑
- `wearables`：手表
- `audio`：耳机音箱
- `laptops`：笔记本电脑
- `routers`：路由器
- `smart_screens`：智慧屏
- `printers`：打印机
- `monitors`：显示器
- `cars`：汽车，来源为鸿蒙智行官网 `hima.auto`

## 输出格式

每个品类会生成一个完整 JSON 文件。默认目录是 `src/out`，例如：

- `src/out/phones.json`
- `src/out/tablets.json`
- `src/out/cars.json`

如果你显式传入 `--output-dir output`，则会生成到：

- `output/phones.json`
- `output/tablets.json`

文件结构示例：

```json
{
  "category_key": "phones",
  "category_name": "手机",
  "pull_time": "260518",
  "series": {
    "pura系列": {
      "series_key": "pura系列",
      "series_name": "Pura系列",
      "products": {
        "pura90pro": {
          "product_key": "pura90pro",
          "product_id": "10086866372688",
          "product_name": "HUAWEI Pura 90 Pro",
          "product_name_short": "Pura 90 Pro",
          "sku_code": "2601010614418",
          "sku_name": "HUAWEI Pura 90 Pro 12GB+256GB 椰青白",
          "product_url": "https://www.vmall.com/product/10086866372688.html",
          "third_level_category_code": "260101",
          "specifications": [
            {
              "group_name": "主体",
              "attr_name": "品牌",
              "attr_value": "华为"
            }
          ]
        }
      }
    }
  }
}
```

说明：

- 顶层按品类保存
- `pull_time` 表示本次拉取时间，格式为 `yymmdd`
- 品类下按系列归类
- 系列下按产品归类
- `specifications` 为最终规格参数列表

## 新增商品识别

每次导出同一个品类时，工具会先读取上一次完整导出文件，例如 `src/out/phones.json`，再和本次拉取结果对比。

对比规则：

- 优先使用 `product_id` 判断是否为同一个商品
- 如果缺少 `product_id`，会依次回退到 `sku_code` 和产品键名
- 首次拉取时没有历史文件，只生成完整品类 JSON，不生成 diff 文件
- 如果本次没有新增商品，不生成 diff 文件
- 如果发现新增商品，会只把新增部分写入 `src/diff`

新增 diff 文件命名格式：

```text
{category_key}_{上次pull_time}_{本次pull_time}_new.json
```

示例：

```text
src/diff/phones_260517_260518_new.json
```

diff 文件结构示例：

```json
{
  "category_key": "phones",
  "category_name": "手机",
  "previous_pull_time": "260517",
  "current_pull_time": "260518",
  "new_product_count": 1,
  "series": {
    "mate系列": {
      "series_key": "mate系列",
      "series_name": "Mate系列",
      "products": {
        "mate80": {
          "product_id": "10000000000000",
          "product_name": "HUAWEI Mate 80",
          "specifications": []
        }
      }
    }
  }
}
```

常见使用流程：

```bash
python run.py --category phones
```

第一次运行：

- 生成 `src/out/phones.json`
- JSON 顶层包含本次 `pull_time`
- 因为没有历史文件，不生成 `src/diff`

第二次及以后运行：

- 先读取旧的 `src/out/phones.json`
- 拉取本次最新数据
- 覆盖写入新的 `src/out/phones.json`
- 如果有新增商品，生成 `src/diff/phones_旧时间_新时间_new.json`

## 汽车配置导出

汽车品类使用 `cars`：

```bash
python run.py --category cars
```

该导出器会从鸿蒙智行官网抓取以下品牌入口：

- 问界：`https://hima.auto/wenjie/`
- 智界：`https://hima.auto/zhijie/`
- 享界：`https://hima.auto/xiangjie/`
- 尊界：`https://hima.auto/zunjie/`
- 尚界：`https://hima.auto/shangjie/`

输出文件默认是：

```text
src/out/cars.json
```

汽车输出同样包含：

- 顶层 `pull_time`
- `category_key: "cars"`
- `category_name: "汽车"`
- `series` 下按品牌归类，例如 `wenjie`、`zhijie`
- 每个车型配置版本以产品形式写入 `products`
- `specifications` 中包含品牌、车型、售价、续航、车身、安全、智能化等页面可解析的配置项

汽车品类也会复用新增识别机制。第二次及以后运行，如果发现新增车型配置，会生成类似：

```text
src/diff/cars_260517_260518_new.json
```

## 耳机音箱类目说明

`audio` 不是单一搜索条件，而是聚合了多组华为商城子类目查询：

- 蓝牙耳机：`keyword=耳机`，`categoryId=100559`，三级物理类目码 `310205`
- 智能音箱：`keyword=智能音箱`，`categoryId=100561`，三级物理类目码 `310206`
- 智能眼镜：`keyword=智能眼镜`，`categoryId=100620`，三级物理类目码 `310218`
- 有线耳机：`keyword=有线耳机`，`categoryId=100591`，三级物理类目码 `310213`

如果后续华为商城继续把“耳机音箱”拆出新的子类目，优先在 [categories.py](src/vmall_export_new/categories.py) 的 `audio.search_queries` 里新增一组查询配置，而不是只改 `keyword`。

## 评测报告 PPT 生成

`report` 目录提供了小艺参数问答评测报告生成脚本，可根据人工评测 Excel 自动生成 PPT 报告和统计汇总 Excel。

脚本入口：

```text
report/generate_mock_phone_eval_ppt.py
```

脚本不调用模型、不访问网络，执行过程为：

```text
评测 Excel -> 统计聚合 -> 生成统计 Excel -> 生成 PPT
```

正式评测表字段固定时，可复用于手机、平板、耳机等品类。评测表需要包含 `query所属设备` 字段，用于区分 `pura80`、`pura90` 等不同机型或设备。

使用方式：

1. 打开 `report/generate_mock_phone_eval_ppt.py`
2. 修改脚本顶部 `CONFIG` 中的输入 Excel、输出 PPT、品类名等配置
3. 直接执行脚本

```powershell
python report\generate_mock_phone_eval_ppt.py
```

详细字段说明、依赖安装、参数说明和常见问题见 [PPT生成脚本使用说明.md](report/PPT生成脚本使用说明.md)。

## 测试

安装开发依赖：

```bash
python -m pip install -e .[dev]
```

运行测试：

```bash
python -m pytest
```

## 注意事项

- 当前搜索阶段使用官方接口 `openapi.vmall.com/mcp/v1/search/queryPrd`
- 当前规格提取阶段使用商品详情页官方 `__NEXT_DATA__`
- “全部品类”基于内置品类配置，不依赖商城导航树自动发现
- 华为商城页面结构如果后续变化，解析逻辑可能需要调整
- 当前已实现基础稳定性策略：
  - 自动获取并缓存 csrf token
  - csrf 失效后自动刷新并重试
  - 403/429/网络异常自动重试
  - 串行请求 + 最小请求间隔控制
  - 标准输出日志和可选日志文件

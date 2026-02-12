# Cost Calculator (电商订单成本计算器)

这是一个用于处理电商订单 Excel 文件的 Python 工具，能够自动根据订单备注计算各类成本，并生成包含详细成本明细和店铺统计的新 Excel 文件。

## 功能特点

*   **自动成本计算**：
    *   **基础成本**：根据 `size_material_price.json` 配置文件，通过关键词、尺寸和材质匹配自动计算商品基础成本。
    *   **代发成本**：能识别 "枕芯"、"yr"/"义乳"、"yt"/"义臀" 等关键词，并自动计算代发附加费。
    *   **义乳/义臀动销成本**：通过 `moving_and_selling_costs.json` 配置进行匹配计算。
    *   **枕芯成本**：通过 `pillow_cost.json` 配置进行匹配计算。
    *   **硅胶/电动其他成本**：通过 `others.json` 配置进行匹配计算。
*   **智能店铺统计**：
    *   自动从 "商家/店铺" 列提取并规范化店铺名称。
    *   生成独立的 "店铺统计" Sheet，按店铺汇总各项成本（支持 Excel 公式动态更新）。
*   **特殊处理**：
    *   支持多商品订单合并计算（通过 `。` 分割）。
    *   自动识别数量（通过逗号统计）。
    *   海外订单自动识别（"发海外"关键词），并预留空位供手动填写。
    *   错误检测与标记。
*   **输出格式**：
    *   保留原始数据，新增成本明细列。
    *   生成带有公式的统计表，方便后续手动调整数据后自动重算。

## 环境要求

*   Python 3.10+
*   依赖库：`pandas`, `openpyxl`
*   可选增强：`prompt_toolkit`（交互模式 Tab 路径补全 + 输入历史；已包含在 `requirements.txt`）

## 安装

1.  克隆或下载本项目。
2.  安装依赖：

```bash
pip install -r requirements.txt
# 或者
pip install .
```

> 说明：如果你不需要 Tab 补全，也可以不安装 `prompt_toolkit`，程序会自动回退到普通输入模式。

## 使用方法

### 方式一：命令行运行

```bash
python cost_calculator.py "path/to/your/order_file.xlsx"
```

也支持一次传入多个文件/通配符：

```bash
python cost_calculator.py "a.xlsx" "b.xlsx"
python cost_calculator.py "*.xlsx"
```

### 方式二：交互式运行

直接运行脚本，根据提示输入文件路径：

```bash
python cost_calculator.py
```

交互模式下支持：

* 连续处理多个文件：处理完不会自动退出
* 退出命令：`exit` / `quit` / `q`
* 常用命令：`help`、`ls`、`cd <目录>`
* 直接输入目录：自动处理该目录下所有 `.xlsx`/`.xls`（不递归）

### 方式三：打包为 EXE (Windows)

如果已经使用 PyInstaller 打包生成了 exe 文件，可以直接将 Excel 文件拖拽到 exe 图标上运行。

当前构建脚本默认生成 **单文件 EXE**（`--onefile`），并且会把以下配置内嵌到可执行文件中：

* `size_material_price.json`
* `moving_and_selling_costs.json`
* `pillow_cost.json`
* `others.json`

因此发布时 `dist` 目录通常只需要保留：

* `CostCalculator.exe`

## 配置文件说明

工具在运行时会读取同目录下的以下 JSON 配置文件：

1.  **`size_material_price.json`**
    *   定义商品的基础价格体系。
    *   包含：类别、关键词、优先级、不同尺寸对应的材质价格。
    *   示例结构：
        ```json
        [
          {
            "category": "枕套",
            "keywords": ["等身枕套", "枕套"],
            "priority": 3,
            "products": [
              {
                "尺寸": "50*150",
                "价格": { "桃皮绒": 20, "2way": 26 }
              }
            ]
          }
        ]
        ```

2.  **`moving_and_selling_costs.json`**
    *   定义义乳、义臀等动销商品的成本。
    *   匹配规则：仅当备注中包含 "yr" 或 "yt" 时触发。

3.  **`pillow_cost.json`**
    *   定义枕芯成本。
    *   匹配规则：仅当备注中包含 "枕芯" 时触发，匹配尺寸和类型。

4.  **`others.json`**
    *   定义其他硅胶或电动类商品的成本。

## 输入输出规范

*   **输入文件**：必须包含 **"卖家备注"** 和 **"商家/店铺"** 两列。支持 `.xlsx` 和 `.xls` 格式。
*   **输出文件**：将在原文件目录下生成 `{原文件名}-已处理.xlsx`。

## 注意事项

*   对于 `.xls` 格式的旧版 Excel 文件，处理后的样式可能无法完全保留，建议使用 `.xlsx`。
*   海外订单（备注包含 "发海外"）的成本列会被置空，需要人工核对填写。

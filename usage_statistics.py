"""将有效计费明细汇总到与配置价格矩阵一致的用量表。"""

from copy import copy
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

from config_workbook import (
    BASE_PRICE_SHEET,
    PILLOW_MATRIX_HEADERS,
    PILLOW_PRICE_SHEET,
    PILLOW_STYLE_PREFIXES,
)


BASE_USAGE_SHEET = "抱枕规格用量统计"
PILLOW_USAGE_SHEET = "枕芯规格用量统计"


def _pillow_spec(size, keyword, metadata):
    if (size, keyword) in metadata:
        entry = metadata[(size, keyword)]
        return entry["style"], entry["material"]
    # 配置加载器使用的关键词格式：类型前缀 + 材质 + 枕芯。
    body = keyword.removesuffix("枕芯")
    style = next(name for name, prefix in sorted(
        PILLOW_STYLE_PREFIXES.items(), key=lambda item: -len(item[1])
    ) if body.startswith(prefix))
    return style, body[len(PILLOW_STYLE_PREFIXES[style]):] or "PP棉"


def fallback_usage_template(price_data, pillow_cost_data, pillow_metadata):
    """兼容只传价格字典的旧调用，按字典顺序还原矩阵。"""
    workbook = Workbook()
    base = workbook.active
    base.title = BASE_PRICE_SHEET
    sizes = list(dict.fromkeys(
        product["尺寸"] for category in price_data for product in category["products"]
    ))
    base.append(["类别名称", "材质", *sizes])
    for category in price_data:
        materials = dict.fromkeys(
            material for product in category["products"] for material in product["价格"]
        )
        products = {product["尺寸"]: product["价格"] for product in category["products"]}
        for material in materials:
            base.append([category["category"], material, *[
                products.get(size, {}).get(material) for size in sizes
            ]])

    pillow = workbook.create_sheet(PILLOW_PRICE_SHEET)
    pillow.append(PILLOW_MATRIX_HEADERS)
    rows = {}
    for size, prices in pillow_cost_data.items():
        for keyword, price in prices.items():
            style, material = _pillow_spec(size, keyword, pillow_metadata)
            row = rows.setdefault((size, material), dict.fromkeys(PILLOW_STYLE_PREFIXES))
            row[style] = price
    for (size, material), prices in rows.items():
        pillow.append([size, material, *prices.values()])

    for sheet in (base, pillow):
        sheet.freeze_panes = "C2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="70AD47")
            cell.font = Font(color="FFFFFF", bold=True)
        for row in sheet:
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
        for column in range(1, sheet.max_column + 1):
            sheet.column_dimensions[sheet.cell(1, column).column_letter].width = 18
    return workbook


def _copy_layout(source, workbook, name, table_name):
    if name in workbook.sheetnames:
        del workbook[name]
    target = workbook.create_sheet(name)
    for row in source:
        for cell in row:
            output = target.cell(cell.row, cell.column, cell.value)
            output.font = copy(cell.font)
            output.fill = copy(cell.fill)
            output.border = copy(cell.border)
            output.alignment = copy(cell.alignment)
            output.protection = copy(cell.protection)
            output.number_format = cell.number_format
    for key, dimension in source.column_dimensions.items():
        target.column_dimensions[key] = copy(dimension)
    for key, dimension in source.row_dimensions.items():
        target.row_dimensions[key] = copy(dimension)
    target.freeze_panes = source.freeze_panes
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    # 表名在整个工作簿中必须唯一；避免与订单原有的表格重名。
    existing = {name.lower() for sheet in workbook for name in sheet.tables}
    while table_name.lower() in existing:
        table_name += "_"
    table = Table(displayName=table_name, ref=target.dimensions)
    source_table = next(iter(source.tables.values()), None)
    table.tableStyleInfo = copy(source_table.tableStyleInfo) if source_table else TableStyleInfo(
        name="TableStyleMedium4", showRowStripes=True
    )
    target.add_table(table)
    return target


def write_usage_sheets(workbook, quantities, price_data, pillow_cost_data,
                       pillow_metadata=None, usage_template=None):
    """写入静态数量；空单价仍留空，零单价也属于支持的规格。"""
    metadata = pillow_metadata or {}
    source = (load_workbook(BytesIO(usage_template)) if usage_template else
              fallback_usage_template(price_data, pillow_cost_data, metadata))
    try:
        base = _copy_layout(source[BASE_PRICE_SHEET], workbook, BASE_USAGE_SHEET, "BaseUsage")
        pillow = _copy_layout(source[PILLOW_PRICE_SHEET], workbook, PILLOW_USAGE_SHEET, "PillowUsage")
        for sheet, cost_type in ((base, "base_cost"), (pillow, "pillow_cost")):
            for row in sheet.iter_rows(min_row=2):
                for cell in row[2:]:
                    if cell.value is None or str(cell.value).strip() == "":
                        cell.value = None
                        continue
                    if cost_type == "base_cost":
                        category, material, size = row[0].value, row[1].value, sheet.cell(1, cell.column).value
                    else:
                        category, material, size = sheet.cell(1, cell.column).value, row[1].value, row[0].value
                    cell.value = quantities[(cost_type, str(category).strip().casefold(),
                                             str(material).strip().casefold(), str(size).strip())]
                    cell.number_format = "0"
    finally:
        source.close()


def collect_usage(quantities, details, pillow_metadata):
    """只累计最终纳入计费的基础商品和枕芯。"""
    for detail in details:
        if not detail.included or detail.cost_type not in ("base_cost", "pillow_cost"):
            continue
        category, material = detail.category, detail.material
        if detail.cost_type == "pillow_cost" and not material:
            category, material = _pillow_spec(detail.size, detail.keyword, pillow_metadata)
        quantities[(detail.cost_type, category.casefold(), material.casefold(), detail.size)] += detail.quantity

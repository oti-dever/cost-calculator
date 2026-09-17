"""
Excel文件处理脚本 - 成本计算与店铺统计
================================================================================

功能概述：
本脚本用于处理电商订单Excel文件，自动计算各类成本并按店铺汇总统计。
输入原始Excel文件，输出包含成本明细和店铺统计的新Excel文件。

核心功能：
1. 成本计算
   - 基础成本：根据「成本配置.xlsx」中的尺寸、材质、关键词匹配计算
   - 代发成本：识别"枕芯"、"yr"/"义乳"、"yt"/"义臀"等关键词
   - 义乳/义臀、枕芯、硅胶/电动成本均由统一 Excel 配置加载

2. 店铺识别
   - 自动从"商家/店铺"列提取店铺名称（去空白、规范化处理）
   - 不再依赖shop.json配置文件

3. 输出结构（两个Sheet）
   a) 成本明细Sheet
      - 保留原始数据的所有列
      - 新增列：成本、代发成本、总成本（=成本+代发成本）
      - 新增列（空列后）：义乳/义臀成本、枕芯成本、硅胶/电动成本、无法匹配原因说明
      - 底部添加合计行，使用SUM公式汇总各成本列

   b) 店铺统计Sheet
      - 按店铺汇总各类成本（成本、代发成本、义乳/义臀、枕芯、硅胶/电动、总成本）
      - 使用SUMIF公式动态引用成本明细Sheet数据
      - 支持手动修改成本明细后自动更新统计
      - 底部添加合计行

4. 特殊处理
   - 海外订单：包含"发海外"关键词的订单，所有成本列置空，需手动补全
   - 数量识别：通过统计逗号（半角/全角）数量来判断订单项数量
   - 错误标记：义乳/义臀或枕芯成本匹配失败时，对应列置空

输入要求：
- Excel文件必须包含"卖家备注"和"商家/店铺"列
- 支持.xlsx和.xls格式（.xls可能无法保留原始样式）

配置文件：「成本配置.xlsx」（首次运行从内嵌模板自动生成）

输出文件：
- 文件名：${原文件名}-已处理.xlsx
- 位置：与原文件相同目录

使用方法：
1. 命令行运行：python price_calculator.py <Excel文件路径>
2. 交互式运行：python price_calculator.py（然后输入文件路径）
3. 打包为exe后拖拽Excel文件到exe图标上运行

版本信息：
- 支持多Sheet输出（成本明细 + 店铺统计）
- 支持多种成本类型计算
- 使用Excel公式实现动态更新
================================================================================
"""

import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Dict, List, Optional, Sequence, Tuple
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import pandas as pd


@dataclass
class ProcessResult:
    """单个文件处理结果。"""

    input_path: str
    success: bool
    status: str
    message: str
    output_path: Optional[str] = None
    matched_count: Optional[int] = None
    total_count: Optional[int] = None
    overseas_count: int = 0
    unmatched_count: int = 0
    unmatched_details: List[Dict[str, object]] = field(default_factory=list)


@dataclass
class MatchResult:
    """卖家备注的成本匹配结果及字段级诊断。"""

    base_cost: Optional[float]
    dropship_cost: Optional[float]
    total_cost: Optional[float]
    moving_cost: Optional[float]
    pillow_cost: Optional[float]
    other_cost: Optional[float]
    matched_any: bool = False
    failure_reasons: List[str] = field(default_factory=list)

    def as_tuple(self):
        """返回与旧版 ``match_price`` 一致的六元组。"""

        return (
            self.base_cost,
            self.dropship_cost,
            self.total_cost,
            self.moving_cost,
            self.pillow_cost,
            self.other_cost,
        )


ORDER_ID_COLUMN = "订单编号"


def _format_cell_text(value) -> str:
    """将单元格值格式化为字符串。"""

    if pd.isna(value):
        return ""
    return str(value).strip()


def extract_order_id(row, seller_note: str) -> str:
    """提取订单编号：仅读取“订单编号”列。"""

    if ORDER_ID_COLUMN in row:
        cell_text = _format_cell_text(row[ORDER_ID_COLUMN])
        if cell_text:
            return cell_text

    return "无"


def build_output_file_path(
    input_file_path: str,
    output_dir: Optional[str] = None,
    overwrite: bool = False,
) -> str:
    """构建输出文件路径。

    参数:
        input_file_path (str): 输入文件路径
        output_dir (Optional[str]): 输出目录；为空时默认输出到输入文件目录
        overwrite (bool): 是否允许覆盖同名文件

    返回:
        str: 最终输出文件路径
    """

    source_dir = os.path.dirname(input_file_path)
    target_dir = output_dir if output_dir else source_dir
    file_name = os.path.basename(input_file_path)
    name_without_ext, _ext = os.path.splitext(file_name)
    base_name = f"{name_without_ext}-已处理"
    output_path = os.path.join(target_dir, f"{base_name}.xlsx")

    if overwrite or not os.path.exists(output_path):
        return output_path

    index = 1
    while True:
        candidate = os.path.join(target_dir, f"{base_name}({index}).xlsx")
        if not os.path.exists(candidate):
            return candidate
        index += 1


def create_session_log(base_path: str) -> Tuple[Optional[str], Callable[[str], None]]:
    """创建会话日志文件，并返回日志函数。"""

    try:
        logs_dir = os.path.join(base_path, "logs")
        os.makedirs(logs_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = os.path.join(logs_dir, f"cost_calculator_{timestamp}.log")

        def _write_log(message: str) -> None:
            time_prefix = datetime.now().strftime("%H:%M:%S")
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{time_prefix}] {message}\n")

        _write_log("日志已启动")
        return log_path, _write_log
    except Exception:
        return None, lambda _msg: None


def print_batch_summary(results: List[ProcessResult], write_log: Callable[[str], None]) -> None:
    """打印批处理汇总与失败清单。"""

    if not results:
        print("本批次没有可处理项。")
        write_log("本批次没有可处理项")
        return

    success_items = [r for r in results if r.status == "success"]
    failed_items = [r for r in results if r.status == "failed"]
    skipped_items = [r for r in results if r.status == "skipped"]

    total_files = len(results)
    total_overseas = sum(r.overseas_count for r in success_items)

    summary_line = (
        f"批处理汇总: 总计 {total_files} | 成功 {len(success_items)} | "
        f"失败 {len(failed_items)} | 跳过 {len(skipped_items)}"
    )
    print(summary_line)
    write_log(summary_line)

    if total_overseas > 0:
        overseas_line = f"海外订单累计: {total_overseas} 条"
        print(overseas_line)
        write_log(overseas_line)

    if failed_items:
        print("失败清单:")
        write_log("失败清单:")
        for idx, item in enumerate(failed_items, start=1):
            line = f"  {idx}. {item.input_path} | 原因: {item.message}"
            print(line)
            write_log(line)

    if skipped_items:
        print("跳过清单:")
        write_log("跳过清单:")
        for idx, item in enumerate(skipped_items, start=1):
            line = f"  {idx}. {item.input_path} | 原因: {item.message}"
            print(line)
            write_log(line)


def find_leftmost_match(text, candidates):
    """
    从左向右找第一个（最左边）匹配的候选项

    参数:
        text (str): 要搜索的文本（已转为小写）
        candidates (list): 候选项列表（字符串或元组）

    返回:
        tuple: (匹配的候选项, 位置)，如果没有匹配则返回 (None, float('inf'))
    """
    best_match = None
    best_pos = float("inf")

    for candidate in candidates:
        # 如果候选项是元组，取第一个元素作为搜索字符串
        search_str = candidate[0] if isinstance(candidate, tuple) else candidate
        pos = text.find(search_str.lower())
        if pos != -1 and pos < best_pos:
            best_match = candidate
            best_pos = pos

    return best_match, best_pos


def find_longest_match_at_leftmost(text, candidates):
    """
    从左向右找匹配,如果多个候选项在同一位置匹配,选择最长的
    专用于材质匹配,按长度从长到短优先

    参数:
        text (str): 要搜索的文本（已转为小写）
        candidates (list): 候选项列表（已按长度从长到短排序）

    返回:
        tuple: (匹配的候选项, 位置)，如果没有匹配则返回 (None, float('inf'))
    """
    best_match = None
    best_pos = float("inf")
    best_len = 0

    for candidate in candidates:
        # 如果候选项是元组，取第一个元素作为搜索字符串
        search_str = candidate[0] if isinstance(candidate, tuple) else candidate
        pos = text.find(search_str.lower())
        if pos != -1:
            # 优先级: 1. 位置更靠左 2. 同一位置时长度更长
            if pos < best_pos or (pos == best_pos and len(search_str) > best_len):
                best_match = candidate
                best_pos = pos
                best_len = len(search_str)

    return best_match, best_pos


DEFAULT_DROPSHIP_KEYWORDS = ("枕芯", "yr", "义乳", "yt", "义臀")


def _reason_fragment(text: str, max_length: int = 40) -> str:
    """生成适合写入 Excel 的短备注片段。"""

    normalized = re.sub(r"\s+", " ", str(text)).strip()
    if not normalized:
        return "(空)"
    if len(normalized) <= max_length:
        return normalized
    return normalized[: max_length - 3] + "..."


def _base_size_formats(size: str) -> List[str]:
    return [size.lower(), size.lower().replace("*", "x"), size.lower().replace("*", "")]


def _diagnose_base_failure(
    item_text_lower: str,
    quantity: int,
    keyword_seen: bool,
    keyword_size_seen: bool,
    price_data: list,
) -> Optional[str]:
    """判断基础商品匹配停在哪个字段。"""

    if keyword_seen:
        if not keyword_size_seen:
            return "缺少或不支持尺寸"
        return "缺少或不支持材质"

    known_sizes = {
        size_format
        for category in price_data
        for product in category.get("products", [])
        for size_format in _base_size_formats(str(product.get("尺寸", "")))
        if size_format
    }
    known_materials = {
        str(material).lower()
        for category in price_data
        for product in category.get("products", [])
        for material in product.get("价格", {})
    }
    size_seen = any(size and size in item_text_lower for size in known_sizes)
    material_seen = any(
        material and material in item_text_lower for material in known_materials
    )

    # 有逗号时，旧规则认为该分段必须存在基础商品；没有逗号时，仅在尺寸和材质
    # 都明显存在的情况下判断为缺少商品类型，避免把普通说明文字误报为枕套。
    if quantity > 0 or (size_seen and material_seen):
        return "未识别商品类型关键词"
    return None


def match_price_detailed(
    text,
    price_data,
    moving_costs_data=None,
    pillow_cost_data=None,
    others_cost_data=None,
    dropship_unit_cost: float = 5.5,
    dropship_keywords: Optional[Sequence[str]] = None,
) -> MatchResult:
    """匹配一条卖家备注，并返回成本及字段级失败原因。"""

    if pd.isna(text) or not str(text).strip():
        return MatchResult(
            None,
            None,
            None,
            None,
            None,
            None,
            matched_any=False,
            failure_reasons=["卖家备注为空"],
        )

    moving_costs_data = moving_costs_data or []
    pillow_cost_data = pillow_cost_data or {}
    others_cost_data = others_cost_data or []
    effective_dropship_keywords = tuple(
        dropship_keywords or DEFAULT_DROPSHIP_KEYWORDS
    )

    original_text = str(text)
    compact_text = re.sub(r"\s+", "", original_text)
    item_texts = compact_text.split("。")
    raw_items = original_text.split("。")

    total_base_cost = 0.0
    total_dropship_cost = 0.0
    total_moving_cost = 0.0
    total_pillow_cost = 0.0
    total_others_cost = 0.0

    moving_has_error = False
    pillow_has_error = False
    fatal_base_error = False
    matched_any = False
    failure_reasons: List[str] = []

    sorted_categories = sorted(price_data, key=lambda x: x.get("priority", 999))

    for item_index, item_text in enumerate(item_texts, start=1):
        if not item_text.strip():
            continue

        reasons_before_item = len(failure_reasons)
        raw_item = raw_items[item_index - 1] if item_index - 1 < len(raw_items) else item_text
        fragment = _reason_fragment(raw_item)
        item_text_lower = item_text.lower()
        quantity = item_text.count(",") + item_text.count("，")

        item_price = 0.0
        base_price_found = False
        keyword_seen = False
        keyword_size_seen = False

        for category_data in sorted_categories:
            keywords = category_data.get("keywords", [])
            products = category_data.get("products", [])

            matched_keyword, _keyword_pos = find_leftmost_match(
                item_text_lower, keywords
            )
            if matched_keyword is None:
                continue
            keyword_seen = True

            size_candidates = []
            for product in products:
                size = str(product.get("尺寸", ""))
                for size_format in _base_size_formats(size):
                    size_candidates.append((size_format, product))

            matched_size, _size_pos = find_leftmost_match(
                item_text_lower, size_candidates
            )
            if matched_size is None:
                continue
            keyword_size_seen = True
            matched_product = matched_size[1]

            material_candidates = [
                (material, price)
                for material, price in matched_product.get("价格", {}).items()
            ]
            material_candidates.sort(key=lambda x: len(x[0]), reverse=True)
            matched_material, _material_pos = find_longest_match_at_leftmost(
                item_text_lower, material_candidates
            )
            if matched_material:
                item_price = float(matched_material[1])
                base_price_found = True
                break

        if base_price_found:
            if quantity == 0:
                quantity = 1
            total_base_cost += item_price * quantity
            matched_any = True
        else:
            base_failure = _diagnose_base_failure(
                item_text_lower,
                quantity,
                keyword_seen,
                keyword_size_seen,
                price_data,
            )
            if base_failure:
                failure_reasons.append(
                    f"第{item_index}段基础成本：{base_failure}（片段：{fragment}）"
                )
            if quantity > 0:
                fatal_base_error = True

        extra_count = sum(
            item_text_lower.count(str(keyword).lower())
            for keyword in effective_dropship_keywords
            if str(keyword)
        )
        extra_cost = extra_count * float(dropship_unit_cost)
        total_dropship_cost += extra_cost
        if extra_count > 0:
            matched_any = True

        moving_triggered = bool(re.search(r"yr|yt", item_text_lower))
        if moving_costs_data:
            item_moving_cost = calculate_moving_cost_for_item(
                item_text, moving_costs_data
            )
            if item_moving_cost is None:
                moving_has_error = True
                failure_reasons.append(
                    f"第{item_index}段义乳/义臀成本：未匹配完整规格"
                    f"（片段：{fragment}）"
                )
            else:
                total_moving_cost += item_moving_cost
        if moving_triggered:
            matched_any = True

        pillow_triggered = "枕芯" in item_text_lower
        if pillow_cost_data:
            item_pillow_cost = calculate_pillow_cost_for_item(
                item_text, pillow_cost_data
            )
            if item_pillow_cost is None:
                pillow_has_error = True
                failure_reasons.append(
                    f"第{item_index}段枕芯成本：缺少或不支持“类型关键词+尺寸”"
                    f"（片段：{fragment}）"
                )
            else:
                total_pillow_cost += item_pillow_cost
        if pillow_triggered:
            matched_any = True

        other_triggered = any(
            str(item.get("remark", "")).lower() in item_text_lower
            for item in others_cost_data
            if str(item.get("remark", ""))
        )
        if others_cost_data:
            total_others_cost += calculate_others_cost_for_item(
                item_text, others_cost_data
            )
        if other_triggered:
            matched_any = True

        # 保留旧规则：「。」分隔的任一独立分段若无法计费，
        # 则整行成本返回空。普通说明文字与有效商品位于同一
        # 分段时，仍不会影响已匹配项目。
        if (
            quantity == 0
            and not base_price_found
            and extra_cost == 0
            and total_others_cost == 0
        ):
            fatal_base_error = True
            if len(failure_reasons) == reasons_before_item:
                failure_reasons.append(
                    f"第{item_index}段：未识别任何可计费项目"
                    f"（片段：{fragment}）"
                )

    if not matched_any and not failure_reasons:
        failure_reasons.append(
            f"第1段：未识别任何可计费项目（片段：{_reason_fragment(original_text)}）"
        )

    if fatal_base_error or not matched_any:
        return MatchResult(
            None,
            None,
            None,
            None,
            None,
            None,
            matched_any=matched_any,
            failure_reasons=failure_reasons,
        )

    final_moving_cost = None if moving_has_error else total_moving_cost
    final_pillow_cost = None if pillow_has_error else total_pillow_cost
    total_cost = total_base_cost + total_dropship_cost

    return MatchResult(
        total_base_cost,
        total_dropship_cost,
        total_cost,
        final_moving_cost,
        final_pillow_cost,
        total_others_cost,
        matched_any=matched_any,
        failure_reasons=failure_reasons,
    )


def match_price(
    text,
    price_data,
    moving_costs_data=None,
    pillow_cost_data=None,
    others_cost_data=None,
    dropship_unit_cost: float = 5.5,
    dropship_keywords: Optional[Sequence[str]] = None,
):
    """
    兼容旧调用的成本匹配接口。

    返回:
        tuple: (成本, 代发成本, 总成本, 义乳/义臀成本, 枕芯成本, 硅胶/电动成本)
    """

    return match_price_detailed(
        text,
        price_data,
        moving_costs_data,
        pillow_cost_data,
        others_cost_data,
        dropship_unit_cost=dropship_unit_cost,
        dropship_keywords=dropship_keywords,
    ).as_tuple()

def calculate_moving_cost_for_item(item_text, moving_costs_data):
    """
    为单个订单项计算义乳/义臀成本（严格匹配）

    规则：
    - 如果订单项不包含yr或yt关键词，返回0
    - 如果包含yr或yt，必须每个yr/yt都能在「义乳义臀价格」工作表中匹配到完整关键词
    - 搜索范围: 上一个yr/yt位置的末尾后(如有)到下一个yr/yt位置的开头前(如有)
    - 如果任何一个yr/yt匹配不到，返回None（表示备注错误）

    参数:
        item_text (str): 单个订单项文本（已去除空格）
        moving_costs_data (list): 动销成本数据列表

    返回:
        float or None: 该订单项的义乳/义臀成本，如果有匹配错误返回 None，无yr/yt关键词返回 0
    """
    if not item_text or not moving_costs_data:
        return 0

    item_lower = item_text.lower()

    # 检查是否包含yr或yt关键词
    if "yr" not in item_lower and "yt" not in item_lower:
        return 0

    # 构建remark字典（小写）
    remark_dict = {}
    for item in moving_costs_data:
        remark = str(item.get("remark", "")).lower()
        if remark:
            remark_dict[remark] = float(item.get("cost", 0))

    # 找出所有yr和yt的位置(起始位置+长度)
    positions = []  # [(start_pos, end_pos, keyword)]

    # 查找yr
    pos = 0
    while True:
        pos = item_lower.find("yr", pos)
        if pos == -1:
            break
        positions.append((pos, pos + 2, "yr"))
        pos += 1

    # 查找yt
    pos = 0
    while True:
        pos = item_lower.find("yt", pos)
        if pos == -1:
            break
        positions.append((pos, pos + 2, "yt"))
        pos += 1

    if not positions:
        return 0

    # 按起始位置排序
    positions.sort(key=lambda x: x[0])

    item_cost = 0.0

    # 对每个yr/yt位置，在限定范围内尝试匹配完整的remark
    for i, (start_pos, end_pos, keyword) in enumerate(positions):
        matched = False

        # 确定搜索范围: 上一个yr/yt末尾后 到 下一个yr/yt开头前
        search_start = positions[i - 1][1] if i > 0 else 0  # 上一个的末尾
        search_end = (
            positions[i + 1][0] if i < len(positions) - 1 else len(item_lower)
        )  # 下一个的开头
        search_text = item_lower[search_start:search_end]

        # 按remark长度从长到短排序，优先匹配更长的
        sorted_remarks = sorted(
            remark_dict.items(), key=lambda x: len(x[0]), reverse=True
        )

        for remark, cost in sorted_remarks:
            # 检查remark是否包含当前关键词
            if keyword not in remark:
                continue

            # 在搜索范围内查找remark
            if remark in search_text:
                item_cost += cost
                matched = True
                break

        # 如果这个yr/yt没有匹配到任何remark，返回None表示错误
        if not matched:
            return None

    return item_cost


def calculate_pillow_cost_for_item(item_text, pillow_cost_data):
    """
    为单个订单项计算枕芯成本（严格匹配）

    规则：
    - 如果订单项不包含"枕芯"关键词，返回0
    - 如果包含"枕芯"，必须每个枕芯都能匹配到 尺寸+枕芯类型关键词
    - 搜索范围: 上一个枕芯位置后(如有)到当前枕芯位置
    - 使用正则表达式匹配: {尺寸}{0-n个空格}{枕芯关键词}
    - 如果任何一个枕芯匹配不到，返回None（表示备注错误）

    参数:
        item_text (str): 单个订单项文本（已去除空格）
        pillow_cost_data (dict): 枕芯成本数据字典 {size: {keyword: cost}}

    返回:
        float or None: 该订单项的枕芯成本，如果有匹配错误返回 None，无枕芯关键词返回 0
    """
    if not item_text or not pillow_cost_data:
        return 0

    item_lower = item_text.lower()

    # 检查是否包含"枕芯"关键词
    if "枕芯" not in item_lower:
        return 0

    # 找出所有"枕芯"的位置
    pillow_positions = []
    pos = 0
    while True:
        pos = item_lower.find("枕芯", pos)
        if pos == -1:
            break
        pillow_positions.append(pos)
        pos += 2  # "枕芯"是2个字符，跳过以避免重复

    if not pillow_positions:
        return 0

    item_cost = 0.0

    # 对每个"枕芯"位置，在限定范围内尝试匹配尺寸和关键词
    for i, pillow_pos in enumerate(pillow_positions):
        matched = False

        # 确定搜索范围: 上一个枕芯位置后 到 当前枕芯位置(包含)
        search_start = pillow_positions[i - 1] + 2 if i > 0 else 0  # 上一个枕芯的末尾
        search_end = (
            pillow_positions[i + 1] + 2
            if i < len(pillow_positions) - 1
            else len(item_lower)
        )  # 当前枕芯的末尾
        search_text = item_lower[search_start:search_end]

        # 按尺寸从长到短排序
        sorted_sizes = sorted(pillow_cost_data.keys(), key=len, reverse=True)

        for size in sorted_sizes:
            size_lower = size.lower()
            # 准备三种尺寸格式
            size_formats = [
                size_lower,  # 120*40
                size_lower.replace("*", "x"),  # 120x40
                size_lower.replace("*", ""),  # 12040
                size_lower.split("*")[1] + "*" + size_lower.split("*")[0],  # 40*120
                size_lower.split("*")[1] + "x" + size_lower.split("*")[0],  # 40x120
                size_lower.split("*")[1] + size_lower.split("*")[0],  # 40120
            ]

            # 获取该尺寸下的所有关键词
            keywords = pillow_cost_data[size]
            sorted_keywords = sorted(keywords.keys(), key=len, reverse=True)

            # 尝试匹配每种尺寸格式 + 每个关键词的组合
            for size_format in size_formats:
                for keyword in sorted_keywords:
                    keyword_lower = keyword.lower()
                    # 构建正则表达式: 关键词(已包含"枕芯") + 0-n个空格 + 尺寸
                    # 注意：item_text已经去除空格，所以实际上不会有空格，但为了健壮性保留这个逻辑
                    pattern = re.escape(keyword_lower) + r"\s*" + re.escape(size_format)

                    if re.search(pattern, search_text):
                        item_cost += keywords[keyword]
                        matched = True
                        break

                if matched:
                    break

            if matched:
                break

        # 如果这个"枕芯"没有匹配到，返回None表示错误
        if not matched:
            return None

    return item_cost


def calculate_others_cost_for_item(item_text, others_cost_data):
    """
    为单个订单项计算硅胶/电动成本

    规则：
    - 匹配「其他成本」工作表中的关键词
    - 按最长匹配原则

    参数:
        item_text (str): 单个订单项文本（已去除空格）
        others_cost_data (list): 硅胶/电动成本数据列表

    返回:
        float: 该订单项的硅胶/电动成本，如果没有匹配则返回 0
    """
    if not item_text or not others_cost_data:
        return 0

    item_lower = item_text.lower()

    # 构建remark字典（小写）
    remark_dict = {}
    for item in others_cost_data:
        remark = str(item.get("remark", "")).lower()
        if remark:
            remark_dict[remark] = float(item.get("cost", 0))

    # 按remark长度从长到短排序（优先匹配更长的）
    sorted_remarks = sorted(remark_dict.items(), key=lambda x: len(x[0]), reverse=True)

    item_cost = 0.0

    # 匹配所有可能的remark
    for remark, cost in sorted_remarks:
        if remark in item_lower:
            item_cost += cost
            # 只匹配一次最长的
            break

    return item_cost


def process_shop_summary_sheet(
    workbook,
    detail_sheet_name,
    shop_names,
    data_start_row,
    data_end_row,
    shop_name_col,
    cost_col,
    moving_col,
    dropship_col,
    total_col,
    pillow_col,
    others_col,
):
    """
    创建独立的店铺统计sheet（使用公式引用成本明细sheet）

    参数:
        workbook: openpyxl 工作簿对象
        detail_sheet_name: 成本明细sheet的名称
        shop_names: 店铺名称集合
        data_start_row: 数据起始行（通常是2）
        data_end_row: 数据结束行
        shop_name_col: 店铺名称列的字母
        cost_col: 成本列的字母
        moving_col: 义乳/义臀成本列的字母
        dropship_col: 代发成本列的字母
        total_col: 总成本列的字母
        pillow_col: 枕芯成本列的字母
        others_col: 硅胶/电动成本列的字母
    """
    from openpyxl.styles import Alignment

    # 创建或获取店铺统计sheet
    if "店铺统计" in workbook.sheetnames:
        summary_sheet = workbook["店铺统计"]
        # 清空现有内容
        workbook.remove(summary_sheet)

    summary_sheet = workbook.create_sheet("店铺统计")

    # 写入表头
    header_row = 1
    summary_sheet.cell(row=header_row, column=1, value="店铺名称")
    summary_sheet.cell(row=header_row, column=2, value="成本")
    summary_sheet.cell(row=header_row, column=3, value="代发成本")
    summary_sheet.cell(row=header_row, column=4, value="义乳/义臀成本")
    summary_sheet.cell(row=header_row, column=5, value="枕芯成本")
    summary_sheet.cell(row=header_row, column=6, value="硅胶/电动成本")
    summary_sheet.cell(row=header_row, column=7, value="总成本")

    # 写入店铺数据
    current_row = header_row + 1
    for shop_name in sorted(shop_names):
        # 店铺名称
        summary_sheet.cell(row=current_row, column=1, value=shop_name)

        # 成本合计 - 使用SUMIF公式引用成本明细sheet
        cost_formula = f'=SUMIF({detail_sheet_name}!${shop_name_col}${data_start_row}:${shop_name_col}${data_end_row},"{shop_name}",{detail_sheet_name}!${cost_col}${data_start_row}:${cost_col}${data_end_row})'
        summary_sheet.cell(row=current_row, column=2, value=cost_formula)

        # 代发成本合计
        dropship_formula = f'=SUMIF({detail_sheet_name}!${shop_name_col}${data_start_row}:${shop_name_col}${data_end_row},"{shop_name}",{detail_sheet_name}!${dropship_col}${data_start_row}:${dropship_col}${data_end_row})'
        summary_sheet.cell(row=current_row, column=3, value=dropship_formula)

        # 义乳/义臀成本合计
        moving_formula = f'=SUMIF({detail_sheet_name}!${shop_name_col}${data_start_row}:${shop_name_col}${data_end_row},"{shop_name}",{detail_sheet_name}!${moving_col}${data_start_row}:${moving_col}${data_end_row})'
        summary_sheet.cell(row=current_row, column=4, value=moving_formula)

        # 枕芯成本合计
        pillow_formula = f'=SUMIF({detail_sheet_name}!${shop_name_col}${data_start_row}:${shop_name_col}${data_end_row},"{shop_name}",{detail_sheet_name}!${pillow_col}${data_start_row}:${pillow_col}${data_end_row})'
        summary_sheet.cell(row=current_row, column=5, value=pillow_formula)

        # 硅胶/电动成本合计
        others_formula = f'=SUMIF({detail_sheet_name}!${shop_name_col}${data_start_row}:${shop_name_col}${data_end_row},"{shop_name}",{detail_sheet_name}!${others_col}${data_start_row}:${others_col}${data_end_row})'
        summary_sheet.cell(row=current_row, column=6, value=others_formula)

        # 总成本 = 成本 + 代发成本 + 义乳/义臀成本 + 枕芯成本 + 硅胶/电动成本
        total_row_formula = f"=B{current_row}+C{current_row}+D{current_row}+E{current_row}+F{current_row}"
        summary_sheet.cell(row=current_row, column=7, value=total_row_formula)

        current_row += 1

    # 添加总计行
    summary_sheet.cell(row=current_row, column=1, value="合计")

    # 总计使用SUM公式汇总上面的统计
    sum_start_row = header_row + 1
    sum_end_row = current_row - 1

    # 总计行：分别汇总六个成本列
    summary_sheet.cell(
        row=current_row, column=2, value=f"=SUM(B{sum_start_row}:B{sum_end_row})"
    )
    summary_sheet.cell(
        row=current_row, column=3, value=f"=SUM(C{sum_start_row}:C{sum_end_row})"
    )
    summary_sheet.cell(
        row=current_row, column=4, value=f"=SUM(D{sum_start_row}:D{sum_end_row})"
    )
    summary_sheet.cell(
        row=current_row, column=5, value=f"=SUM(E{sum_start_row}:E{sum_end_row})"
    )
    summary_sheet.cell(
        row=current_row, column=6, value=f"=SUM(F{sum_start_row}:F{sum_end_row})"
    )
    summary_sheet.cell(
        row=current_row, column=7, value=f"=SUM(G{sum_start_row}:G{sum_end_row})"
    )

    print(f"已创建独立的店铺统计sheet（共 {len(shop_names)} 家店铺）")


def process_cost_detail_sheet(
    sheet,
    base_costs,
    dropship_costs,
    moving_costs,
    pillow_costs,
    others_costs,
    match_reasons,
    shop_names,
    shop_name_col_idx,
):
    """
    处理成本明细sheet,添加成本相关列和合计行

    参数:
        sheet: openpyxl工作表对象
        base_costs: 成本列表
        dropship_costs: 代发成本列表
        moving_costs: 义乳/义臀成本列表
        pillow_costs: 枕芯成本列表
        others_costs: 硅胶/电动成本列表
        match_reasons: 无法匹配原因说明列表
        shop_names: 店铺名称列表
        shop_name_col_idx: 店铺名称列索引

    返回:
        tuple: (cost_col_idx, dropship_col_idx, grand_total_col_idx, moving_col_idx, pillow_col_idx, others_col_idx)
    """
    # 查找或创建表头
    header_row = sheet[1]
    if header_row is None:
        print("警告: 无法读取表头")
        return None

    header = [cell.value for cell in header_row]

    # 按正确顺序创建列：成本、代发成本、总成本、空列、义乳/义臀成本、枕芯成本、硅胶/电动成本

    # 处理"成本"列
    if "成本" in header:
        cost_col_idx = header.index("成本") + 1
    else:
        cost_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=cost_col_idx, value="成本")

    # 处理"代发成本"列
    if "代发成本" in header:
        dropship_col_idx = header.index("代发成本") + 1
    else:
        dropship_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=dropship_col_idx, value="代发成本")

    # 处理"总成本"列（仅成本+代发成本）
    if "总成本" in header:
        grand_total_col_idx = header.index("总成本") + 1
    else:
        grand_total_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=grand_total_col_idx, value="总成本")

    # 插入一个空列（显示占位，不参与计算）
    empty_col_idx = sheet.max_column + 1
    sheet.cell(row=1, column=empty_col_idx, value="")

    # 处理"义乳/义臀成本"列（在空列之后）
    if "义乳/义臀成本" in header:
        moving_col_idx = header.index("义乳/义臀成本") + 1
    else:
        moving_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=moving_col_idx, value="义乳/义臀成本")

    # 处理"枕芯成本"列
    if "枕芯成本" in header:
        pillow_col_idx = header.index("枕芯成本") + 1
    else:
        pillow_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=pillow_col_idx, value="枕芯成本")

    # 处理"硅胶/电动成本"列
    if "硅胶/电动成本" in header:
        others_col_idx = header.index("硅胶/电动成本") + 1
    else:
        others_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=others_col_idx, value="硅胶/电动成本")

    # 原因列放在成本列组最后，已存在时直接复用。
    reason_header = "无法匹配原因说明"
    if reason_header in header:
        reason_col_idx = header.index(reason_header) + 1
    else:
        reason_col_idx = sheet.max_column + 1
        sheet.cell(row=1, column=reason_col_idx, value=reason_header)
    sheet.column_dimensions[get_column_letter(reason_col_idx)].width = 50

    # 写入数据（列顺序：成本、代发成本、总成本(成本+代发)、空列、义乳/义臀成本、枕芯成本、硅胶/电动成本）
    for i, (
        shop_name,
        base_cost,
        dropship_cost,
        moving_cost,
        pillow_cost,
        others_cost,
        match_reason,
    ) in enumerate(
        zip(
            shop_names,
            base_costs,
            dropship_costs,
            moving_costs,
            pillow_costs,
            others_costs,
            match_reasons,
        )
    ):
        row_index = i + 2
        if shop_name_col_idx:
            sheet.cell(row=row_index, column=shop_name_col_idx, value=shop_name)
        # 成本
        sheet.cell(row=row_index, column=cost_col_idx, value=base_cost)
        # 代发成本
        sheet.cell(row=row_index, column=dropship_col_idx, value=dropship_cost)
        # 明细总成本=成本+代发成本
        sheet.cell(
            row=row_index,
            column=grand_total_col_idx,
            value=f"=IFERROR({get_column_letter(cost_col_idx)}{row_index}+{get_column_letter(dropship_col_idx)}{row_index},0)",
        )
        # 空列留空
        sheet.cell(row=row_index, column=empty_col_idx, value="")
        # 义乳/义臀成本（空列之后）
        sheet.cell(row=row_index, column=moving_col_idx, value=moving_cost)
        # 枕芯成本
        sheet.cell(row=row_index, column=pillow_col_idx, value=pillow_cost)
        # 硅胶/电动成本
        sheet.cell(row=row_index, column=others_col_idx, value=others_cost)
        # 无法匹配原因说明
        reason_cell = sheet.cell(
            row=row_index, column=reason_col_idx, value=match_reason or ""
        )
        reason_cell.alignment = Alignment(vertical="top", wrap_text=True)

    # 在明细数据底部添加合计行
    data_start_row = 2
    data_end_row = len(base_costs) + 1
    summary_row = data_end_row + 1

    # 合计行标记在"商家/店铺"列
    if shop_name_col_idx:
        sheet.cell(row=summary_row, column=shop_name_col_idx, value="合计")

    sheet.cell(
        row=summary_row,
        column=cost_col_idx,
        value=f"=SUM({get_column_letter(cost_col_idx)}{data_start_row}:{get_column_letter(cost_col_idx)}{data_end_row})",
    )
    sheet.cell(
        row=summary_row,
        column=moving_col_idx,
        value=f"=SUM({get_column_letter(moving_col_idx)}{data_start_row}:{get_column_letter(moving_col_idx)}{data_end_row})",
    )
    sheet.cell(
        row=summary_row,
        column=dropship_col_idx,
        value=f"=SUM({get_column_letter(dropship_col_idx)}{data_start_row}:{get_column_letter(dropship_col_idx)}{data_end_row})",
    )
    # 明细底部合计：总成本为（成本+代发）逐行公式的求和
    sheet.cell(
        row=summary_row,
        column=grand_total_col_idx,
        value=f"=SUM({get_column_letter(grand_total_col_idx)}{data_start_row}:{get_column_letter(grand_total_col_idx)}{data_end_row})",
    )
    sheet.cell(
        row=summary_row,
        column=pillow_col_idx,
        value=f"=SUM({get_column_letter(pillow_col_idx)}{data_start_row}:{get_column_letter(pillow_col_idx)}{data_end_row})",
    )
    sheet.cell(
        row=summary_row,
        column=others_col_idx,
        value=f"=SUM({get_column_letter(others_col_idx)}{data_start_row}:{get_column_letter(others_col_idx)}{data_end_row})",
    )
    sheet.cell(row=summary_row, column=reason_col_idx, value="")

    return (
        cost_col_idx,
        dropship_col_idx,
        grand_total_col_idx,
        moving_col_idx,
        pillow_col_idx,
        others_col_idx,
    )


def process_excel_file(
    file_path,
    price_data,
    moving_costs_data,
    pillow_cost_data,
    others_cost_data,
    dropship_unit_cost=5.5,
    dropship_keywords=None,
    output_dir=None,
    overwrite=False,
):
    """处理单个Excel文件，输出到新文件。"""
    print(f"处理文件: {file_path}")

    # 生成输出文件路径（支持自定义目录与覆盖策略）
    output_file_path = build_output_file_path(file_path, output_dir, overwrite)

    try:
        # 对于 .xls 文件，样式可能无法保留
        if file_path.endswith(".xls"):
            print("警告: .xls 文件格式较旧，可能无法保留原始样式。")
            df = pd.read_excel(file_path)
        else:
            # 使用 openpyxl 引擎读取以支持后续的样式保留写入
            df = pd.read_excel(file_path, engine="openpyxl")

        # 检查必要的列是否存在
        required_columns = ["卖家备注", "商家/店铺"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            msg = f"文件缺少列 {missing_columns}"
            print(f"警告: {msg}，跳过处理")
            return ProcessResult(
                input_path=file_path,
                success=False,
                status="skipped",
                message=msg,
            )

        # 计算成本、代发成本、义乳/义臀成本、枕芯成本、硅胶/电动成本，同时识别店铺
        base_costs = []  # 成本
        dropship_costs = []  # 代发成本
        moving_costs = []  # 义乳/义臀成本
        pillow_costs = []  # 枕芯成本
        others_costs = []  # 硅胶/电动成本
        grand_total_costs = []  # 总成本
        match_reasons = []  # 无法匹配原因说明
        shop_names = []  # 店铺名称（直接使用"商家/店铺"列去空白）
        overseas_count = 0  # 海外订单计数
        unmatched_records = []  # 未匹配记录明细

        for row_number, (_idx, row) in enumerate(df.iterrows(), start=2):
            seller_note = str(row["卖家备注"]) if not pd.isna(row["卖家备注"]) else ""
            shop_column = str(row["商家/店铺"]) if not pd.isna(row["商家/店铺"]) else ""

            # 检查是否包含“发海外”关键词（忽略大小写）
            is_overseas = "发海外" in seller_note.lower()

            if is_overseas:
                # 如果是海外订单，所有字段置空
                base_costs.append("")
                dropship_costs.append("")
                moving_costs.append("")
                pillow_costs.append("")
                others_costs.append("")
                grand_total_costs.append("")
                match_reasons.append("发海外，成本需人工核对填写")
                shop_names.append("")
                overseas_count += 1
            else:
                # 成本匹配：仅从卖家备注中匹配，同时生成字段级诊断。
                match_result = match_price_detailed(
                    seller_note,
                    price_data,
                    moving_costs_data,
                    pillow_cost_data,
                    others_cost_data,
                    dropship_unit_cost=dropship_unit_cost,
                    dropship_keywords=dropship_keywords,
                )
                (
                    base_cost,
                    dropship_cost,
                    _,
                    moving_cost,
                    pillow_cost,
                    others_cost,
                ) = match_result.as_tuple()

                mismatch_reason = "；".join(match_result.failure_reasons)
                match_reasons.append(mismatch_reason)

                if mismatch_reason:
                    order_id = extract_order_id(row, seller_note)
                    normalized_note = re.sub(r"\s+", " ", seller_note).strip()
                    if len(normalized_note) > 80:
                        normalized_note = normalized_note[:77] + "..."
                    unmatched_records.append(
                        {
                            "row_number": row_number,
                            "order_id": order_id,
                            "reason": mismatch_reason,
                            "seller_note": normalized_note,
                        }
                    )

                base_costs.append(base_cost if base_cost is not None else "")
                dropship_costs.append(
                    dropship_cost if dropship_cost is not None else ""
                )

                # 义乳/义臀成本：None表示匹配错误，留空；0表示未匹配
                moving_costs.append(moving_cost if moving_cost is not None else "")

                # 枕芯成本：None表示匹配错误，留空；0表示未匹配
                pillow_costs.append(pillow_cost if pillow_cost is not None else "")

                # 硅胶/电动成本：None或0都可能出现
                others_costs.append(
                    others_cost if (others_cost is not None and others_cost > 0) else 0
                )

                # 计算总成本：成本 + 代发成本
                total_parts = []
                if base_cost is not None:
                    total_parts.append(base_cost)
                if dropship_cost is not None:
                    total_parts.append(dropship_cost)

                grand_total = sum(total_parts) if total_parts else None
                grand_total_costs.append(grand_total if grand_total is not None else "")

                # 店铺名称：直接使用“商家/店铺”原值（去除前后空白）
                shop_name = shop_column.strip()
                shop_names.append(
                    shop_name
                )  # 对于 .xlsx 文件,使用 openpyxl 写入以保留样式
        if file_path.endswith(".xlsx"):
            workbook = load_workbook(file_path)
            sheet = workbook.active

            # 检查工作表是否存在
            if sheet is None:
                msg = "无法获取活动工作表"
                print(f"警告: {msg},跳过处理")
                return ProcessResult(
                    input_path=file_path,
                    success=False,
                    status="failed",
                    message=msg,
                )

            # 检查工作表是否为空
            if sheet.max_row < 1:
                msg = "工作表为空"
                print(f"警告: {msg},跳过处理")
                return ProcessResult(
                    input_path=file_path,
                    success=False,
                    status="failed",
                    message=msg,
                )

            # 将当前sheet重命名为"成本明细"
            sheet.title = "成本明细"

            # 查找或创建表头
            header_row = sheet[1]
            if header_row is None:
                msg = "无法读取表头"
                print(f"警告: {msg},跳过处理")
                return ProcessResult(
                    input_path=file_path,
                    success=False,
                    status="failed",
                    message=msg,
                )

            header = [cell.value for cell in header_row]

            # 使用源数据中的“商家/店铺”列进行统计（不再新增店铺编号/店铺名称列）
            if "商家/店铺" in header:
                shop_name_col_idx = header.index("商家/店铺") + 1
            else:
                print("警告: 缺少‘商家/店铺’列，无法进行店铺统计")
                shop_name_col_idx = None

            # 处理成本明细sheet
            column_indices = process_cost_detail_sheet(
                sheet,
                base_costs,
                dropship_costs,
                moving_costs,
                pillow_costs,
                others_costs,
                match_reasons,
                shop_names,
                shop_name_col_idx,
            )

            if column_indices is None:
                msg = "成本明细列写入失败"
                return ProcessResult(
                    input_path=file_path,
                    success=False,
                    status="failed",
                    message=msg,
                )

            (
                cost_col_idx,
                dropship_col_idx,
                grand_total_col_idx,
                moving_col_idx,
                pillow_col_idx,
                others_col_idx,
            ) = column_indices

            # 创建独立的店铺统计sheet
            if shop_name_col_idx:
                unique_shop_names = sorted({name for name in shop_names if name})
                data_end_row = len(base_costs) + 1
                process_shop_summary_sheet(
                    workbook,
                    "成本明细",  # 成本明细sheet名称
                    unique_shop_names,
                    2,  # 数据起始行
                    data_end_row,  # 数据结束行
                    get_column_letter(
                        shop_name_col_idx
                    ),  # 店铺名称列字母（源列"商家/店铺"）
                    get_column_letter(cost_col_idx),  # 成本列字母
                    get_column_letter(moving_col_idx),  # 义乳/义臀成本列字母
                    get_column_letter(dropship_col_idx),  # 代发成本列字母
                    get_column_letter(grand_total_col_idx),  # 总成本列字母
                    get_column_letter(pillow_col_idx),  # 枕芯成本列字母
                    get_column_letter(others_col_idx),  # 硅胶/电动成本列字母
                )

            workbook.save(output_file_path)
            print(f"已保存处理后的文件: {output_file_path}")
        else:
            # 对于 .xls 文件，使用 pandas 写入，强制输出为 .xlsx
            # 不再新增店铺编号/店铺名称列，保留原“商家/店铺”列
            df["成本"] = base_costs
            df["代发成本"] = dropship_costs
            df["总成本"] = grand_total_costs
            df["义乳/义臀成本"] = moving_costs
            df["枕芯成本"] = pillow_costs
            df["硅胶/电动成本"] = others_costs
            df["无法匹配原因说明"] = match_reasons
            df.to_excel(output_file_path, index=False, engine="openpyxl")
            print(f"已保存处理后的文件: {output_file_path}")

        # 统计匹配情况
        total_count = len(base_costs)
        unmatched_count = len(unmatched_records)
        matched_count = total_count - overseas_count - unmatched_count
        print(f"匹配成功: {matched_count}/{total_count} 条记录")
        if overseas_count > 0:
            print(f"海外订单: {overseas_count} 条（已置空，需手动补全）")
        if unmatched_count > 0:
            print(f"未匹配记录: {unmatched_count} 条")
            for item in unmatched_records:
                print(
                    "  行号 {row_number} | 订单编号: {order_id} | 原因: {reason} | 备注: {seller_note}".format(
                        row_number=item["row_number"],
                        order_id=item["order_id"],
                        reason=item["reason"],
                        seller_note=item["seller_note"] or "(空)",
                    )
                )

        return ProcessResult(
            input_path=file_path,
            success=True,
            status="success",
            message=(
                f"处理成功（未匹配 {unmatched_count} 条）"
                if unmatched_count > 0
                else "处理成功"
            ),
            output_path=output_file_path,
            matched_count=matched_count,
            total_count=total_count,
            overseas_count=overseas_count,
            unmatched_count=unmatched_count,
            unmatched_details=unmatched_records,
        )

    except Exception as e:
        error_msg = f"处理文件时出错: {e}"
        print(error_msg)
        return ProcessResult(
            input_path=file_path,
            success=False,
            status="failed",
            message=str(e),
        )


def main():
    """主函数（默认启动 Textual 界面）。"""

    try:
        from textual_app import main as textual_main
    except Exception as exc:
        print(f"启动 Textual 界面失败: {exc}")
        print("请先安装依赖（例如 textual），然后重试。")
        return

    textual_main()

if __name__ == "__main__":
    main()

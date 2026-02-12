"""成本计算服务层。

说明：
- 本模块负责“目标路径解析 + 配置加载 + 处理调度”。
- 业务计算规则仍复用 `cost_calculator.py` 中既有函数，避免变更规则。
"""

from __future__ import annotations

import io
import os
import shlex
from contextlib import redirect_stdout
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence

from cost_calculator import (
    ProcessResult,
    load_moving_costs,
    load_others_cost,
    load_pillow_cost,
    load_price_data,
    load_shop_data,
    process_excel_file,
    resolve_resource_path,
)


LogFunc = Callable[[str], None]


@dataclass
class LoadedConfig:
    """已加载的配置数据。"""

    price_data: list
    shop_data: dict
    moving_costs_data: list
    pillow_cost_data: dict
    others_cost_data: list


class CostCalculatorService:
    """成本计算服务。"""

    def __init__(self, base_path: Optional[str] = None) -> None:
        self.base_path = base_path or os.getcwd()
        self.config: Optional[LoadedConfig] = None

    def load_configs(self, log: Optional[LogFunc] = None) -> LoadedConfig:
        """加载所有配置文件并缓存结果。"""

        logger = log or (lambda _msg: None)

        json_path = resolve_resource_path("size_material_price.json")
        shop_json_path = resolve_resource_path("shop.json")
        moving_costs_json_path = resolve_resource_path(
            "moving_and_selling_costs.json"
        )
        pillow_cost_json_path = resolve_resource_path("pillow_cost.json")
        others_json_path = resolve_resource_path("others.json")

        if not os.path.exists(json_path):
            raise FileNotFoundError(f"找不到价格配置文件: {json_path}")

        logger("加载价格数据...")
        price_data = load_price_data(json_path)
        logger(f"已加载 {len(price_data)} 个类别的价格数据")

        shop_data = load_shop_data(shop_json_path)
        moving_costs_data = load_moving_costs(moving_costs_json_path)
        if moving_costs_data:
            logger(f"已加载 {len(moving_costs_data)} 条动销成本数据")

        pillow_cost_data = load_pillow_cost(pillow_cost_json_path)
        if pillow_cost_data:
            logger(f"已加载 {len(pillow_cost_data)} 种尺寸的枕芯成本数据")

        others_cost_data = load_others_cost(others_json_path)
        if others_cost_data:
            logger(f"已加载 {len(others_cost_data)} 条硅胶/电动成本数据")

        self.config = LoadedConfig(
            price_data=price_data,
            shop_data=shop_data,
            moving_costs_data=moving_costs_data,
            pillow_cost_data=pillow_cost_data,
            others_cost_data=others_cost_data,
        )
        return self.config

    @staticmethod
    def parse_target_input(raw_input: str, cwd: Optional[str] = None) -> List[str]:
        """解析输入字符串为路径列表。

        支持：
        - 引号包裹路径
        - 多路径空格分隔
        - 相对路径基于 cwd
        """

        work_dir = cwd or os.getcwd()
        if not raw_input.strip():
            return []

        try:
            tokens = shlex.split(raw_input, posix=False)
        except Exception:
            tokens = [raw_input]

        results: List[str] = []
        for token in tokens:
            cleaned = token.strip().strip('"').strip("'")
            if not cleaned:
                continue

            expanded = os.path.expandvars(os.path.expanduser(cleaned))
            path = (
                os.path.normpath(expanded)
                if os.path.isabs(expanded)
                else os.path.normpath(os.path.join(work_dir, expanded))
            )
            results.append(path)

        # 去重保序
        deduped: List[str] = []
        seen = set()
        for item in results:
            if item not in seen:
                deduped.append(item)
                seen.add(item)
        return deduped

    @staticmethod
    def _list_excel_files(dir_path: str) -> List[str]:
        """列出目录下 Excel 文件（不递归）。"""

        path = Path(dir_path)
        if not path.exists() or not path.is_dir():
            return []

        files: List[str] = []
        for ext in ("*.xlsx", "*.xls"):
            files.extend(str(x) for x in path.glob(ext))

        files.sort(key=lambda x: x.lower())
        return files

    def _ensure_config(self) -> LoadedConfig:
        """确保配置已加载。"""

        if self.config is None:
            return self.load_configs()
        return self.config

    def process_targets(
        self,
        targets: Sequence[str],
        output_dir: Optional[str],
        overwrite: bool,
        log: Optional[LogFunc] = None,
    ) -> List[ProcessResult]:
        """批量处理文件或目录目标。"""

        logger = log or (lambda _msg: None)
        cfg = self._ensure_config()

        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        all_results: List[ProcessResult] = []
        for raw_target in targets:
            all_results.extend(
                self._process_one_target(
                    target=raw_target,
                    output_dir=output_dir,
                    overwrite=overwrite,
                    cfg=cfg,
                    log=logger,
                )
            )

        return all_results

    def _process_one_target(
        self,
        target: str,
        output_dir: Optional[str],
        overwrite: bool,
        cfg: LoadedConfig,
        log: LogFunc,
    ) -> List[ProcessResult]:
        """处理单个目标（文件或目录）。"""

        if not target:
            return []

        if not os.path.exists(target):
            return [
                ProcessResult(
                    input_path=target,
                    success=False,
                    status="failed",
                    message="路径不存在",
                )
            ]

        if os.path.isdir(target):
            excel_files = self._list_excel_files(target)
            if not excel_files:
                return [
                    ProcessResult(
                        input_path=target,
                        success=False,
                        status="skipped",
                        message="目录下未找到 Excel 文件",
                    )
                ]

            log(f"目录批处理: {target}（共 {len(excel_files)} 个文件）")
            dir_results: List[ProcessResult] = []
            for excel_path in excel_files:
                dir_results.extend(
                    self._process_one_target(
                        target=excel_path,
                        output_dir=output_dir,
                        overwrite=overwrite,
                        cfg=cfg,
                        log=log,
                    )
                )
            return dir_results

        if not os.path.isfile(target):
            return [
                ProcessResult(
                    input_path=target,
                    success=False,
                    status="failed",
                    message="不是文件",
                )
            ]

        lower_name = target.lower()
        if not (lower_name.endswith(".xlsx") or lower_name.endswith(".xls")):
            return [
                ProcessResult(
                    input_path=target,
                    success=False,
                    status="skipped",
                    message="非 Excel 文件",
                )
            ]

        result = self._process_excel_with_captured_output(
            target=target,
            output_dir=output_dir,
            overwrite=overwrite,
            cfg=cfg,
            log=log,
        )
        return [result]

    @staticmethod
    def _process_excel_with_captured_output(
        target: str,
        output_dir: Optional[str],
        overwrite: bool,
        cfg: LoadedConfig,
        log: LogFunc,
    ) -> ProcessResult:
        """调用核心处理函数并捕获标准输出写入日志。"""

        buffer = io.StringIO()
        with redirect_stdout(buffer):
            result = process_excel_file(
                target,
                cfg.price_data,
                cfg.shop_data,
                cfg.moving_costs_data,
                cfg.pillow_cost_data,
                cfg.others_cost_data,
                output_dir=output_dir,
                overwrite=overwrite,
            )

        captured = buffer.getvalue().strip()
        if captured:
            for line in captured.splitlines():
                line = line.strip()
                if line:
                    log(line)

        return result

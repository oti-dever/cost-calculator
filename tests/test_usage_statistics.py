from __future__ import annotations

import io
import shutil
import tempfile
import unittest
from contextlib import redirect_stdout
from copy import copy
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook

from config_workbook import load_config_workbook
from cost_calculator import process_excel_file, recognize_note
from usage_statistics import BASE_USAGE_SHEET, PILLOW_USAGE_SHEET


TEMPLATE = Path(__file__).resolve().parents[1] / "成本配置模板.xlsx"


class UsageStatisticsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = load_config_workbook(TEMPLATE, check_lock=False)

    def export(self, directory, notes, config=None, *, legacy=False, fallback=False):
        config = config or self.config
        source = Path(directory) / ("orders.xls" if legacy else "orders.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["卖家备注", "商家/店铺"])
        for index, note in enumerate(notes):
            sheet.append([note, "店A" if index % 2 else "店B"])
        for title in (BASE_USAGE_SHEET, PILLOW_USAGE_SHEET):
            workbook.create_sheet(title).append(["旧统计"])
        workbook.save(source)
        workbook.close()
        kwargs = {} if fallback else {
            "pillow_metadata": config.pillow_metadata,
            "usage_template": config.usage_template,
        }
        # 旧格式读取依赖可选的 xlrd；此处直接提供等价数据帧，验证转换后的写入分支。
        frame = pd.DataFrame({"卖家备注": notes, "商家/店铺": ["店A"] * len(notes)})
        with patch("cost_calculator.pd.read_excel", return_value=frame) if legacy else redirect_stdout(io.StringIO()):
            with redirect_stdout(io.StringIO()):
                result = process_excel_file(
                    str(source), config.price_data, config.moving_costs_data,
                    config.pillow_cost_data, config.others_cost_data,
                    config.dropship_unit_cost, config.dropship_keywords, **kwargs,
                )
        self.assertTrue(result.success, result.message)
        return load_workbook(result.output_path)

    def value(self, sheet, first, second, header):
        column = next(cell.column for cell in sheet[1] if cell.value == header)
        row = next(row for row in sheet.iter_rows(min_row=2)
                   if row[0].value == first and row[1].value == second)
        return row[column - 1].value

    def test_counts_layout_replacement_and_failure_policy(self):
        notes = [
            "A，B，桃皮绒50150ft不开孔枕套+枕芯50150+枕芯50150",
            "A，2way50160开胸枕套。B，桃皮绒50150开孔开胸枕套+加重羽绒棉枕芯40120",
            "桃皮绒50150开孔枕套",
            "A，桃皮绒50150枕套+枕芯50150。B，桃皮绒50150枕套+ft枕芯",
            "A，桃皮绒50150枕套+枕芯50150+yr9999",
            "A，桃皮绒50150枕套+枕芯50150。未知",
            "发海外+桃皮绒50150枕套+枕芯50150",
            "", "A，SAS50150枕套", "防浅粉50150",
        ]
        with tempfile.TemporaryDirectory() as directory:
            workbook = self.export(directory, notes)
            base, pillow = workbook[BASE_USAGE_SHEET], workbook[PILLOW_USAGE_SHEET]
            self.assertEqual(2, self.value(base, "分腿枕套", "桃皮绒", "50*150"))
            self.assertEqual(1, self.value(base, "开孔开胸枕套", "桃皮绒", "50*150"))
            self.assertEqual(1, self.value(base, "开孔枕套", "桃皮绒", "50*150"))
            self.assertEqual(1, self.value(base, "开胸枕套", "2way", "50*160"))
            self.assertEqual(3, self.value(base, "等身枕套", "桃皮绒", "50*150"))
            self.assertEqual(1, self.value(base, "遮羞套", "防浅粉", "50*150"))
            self.assertEqual(3, self.value(pillow, "150*50", "PP棉", "标准等身"))
            self.assertEqual(1, self.value(pillow, "120*40", "羽绒棉", "加重等身"))
            self.assertEqual(0, self.value(base, "等身枕套", "SSS", "40*60"))
            self.assertIsNone(self.value(base, "分腿枕套", "SSS", "40*60"))
            template = load_workbook(TEMPLATE)
            for actual, original in ((base, template["基础商品价格"]),
                                     (pillow, template["枕芯价格"])):
                self.assertEqual(original.max_row, actual.max_row)
                self.assertEqual(original.max_column, actual.max_column)
                self.assertEqual([c.value for c in original[1]], [c.value for c in actual[1]])
                self.assertEqual(original.freeze_panes, actual.freeze_panes)
                self.assertEqual(copy(original["A1"].fill), copy(actual["A1"].fill))
                self.assertEqual(copy(original["A2"].font), copy(actual["A2"].font))
                self.assertEqual(original.column_dimensions["A"].width,
                                 actual.column_dimensions["A"].width)
                self.assertEqual(actual.dimensions, next(iter(actual.tables.values())).ref)
            template.close()
            self.assertEqual(1, workbook.sheetnames.count(BASE_USAGE_SHEET))
            workbook.close()

    def test_zero_prices_custom_specs_and_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.xlsx"
            shutil.copyfile(TEMPLATE, path)
            workbook = load_workbook(path)
            base = workbook["基础商品价格"]
            base.cell(1, 8, "70*200")
            for row in base.iter_rows(min_row=2):
                if row[0].value == "开胸枕套" and row[1].value == "桃皮绒":
                    base.cell(row[0].row, 8, 0)
            workbook["枕芯价格"].append(["200*70", "测试棉", 0, None, None, None, None, None])
            workbook.save(path)
            workbook.close()
            config = load_config_workbook(path, check_lock=False)
            # 加载后移走源配置；导出必须只依赖本批次快照。
            path.unlink()
            output = self.export(directory, ["A，桃皮绒70200开胸枕套+测试棉枕芯20070"], config)
            self.assertEqual(1, self.value(output[BASE_USAGE_SHEET], "开胸枕套", "桃皮绒", "70*200"))
            self.assertEqual(1, self.value(output[PILLOW_USAGE_SHEET], "200*70", "测试棉", "标准等身"))
            self.assertIsNone(self.value(output[PILLOW_USAGE_SHEET], "200*70", "测试棉", "标准开孔"))
            output.close()

    def test_legacy_export_and_old_call_signature(self):
        for legacy in (False, True):
            with self.subTest(legacy=legacy), tempfile.TemporaryDirectory() as directory:
                output = self.export(directory, ["桃皮绒50150枕套+加重ft羽绒棉枕芯50160"],
                                     legacy=legacy, fallback=True)
                self.assertEqual(1, self.value(output[BASE_USAGE_SHEET], "等身枕套", "桃皮绒", "50*150"))
                self.assertEqual(1, self.value(output[PILLOW_USAGE_SHEET], "160*50", "羽绒棉", "加重分腿"))
                output.close()

    def test_split_prices_are_independent(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "config.xlsx"
            shutil.copyfile(TEMPLATE, path)
            workbook = load_workbook(path)
            for row in workbook["基础商品价格"].iter_rows(min_row=2):
                if row[0].value == "开孔枕套" and row[1].value == "桃皮绒":
                    row[3].value = 99
            workbook.save(path)
            workbook.close()
            cfg = load_config_workbook(path, check_lock=False)
            for keyword, expected in (("开孔", 99), ("开胸", 26), ("开孔开胸", 26), ("ft", 26)):
                result = recognize_note(f"桃皮绒50150{keyword}枕套", cfg.price_data)
                self.assertEqual(expected, result.base_cost)

    def test_each_output_starts_with_empty_counts(self):
        with tempfile.TemporaryDirectory() as directory:
            first = self.export(directory, ["桃皮绒50150枕套+枕芯50150"])
            first.close()
            second = self.export(directory, [])
            for title in (BASE_USAGE_SHEET, PILLOW_USAGE_SHEET):
                self.assertEqual(0, sum(cell.value or 0 for row in second[title].iter_rows(min_row=2)
                                        for cell in row[2:]))
            second.close()


if __name__ == "__main__":
    unittest.main()

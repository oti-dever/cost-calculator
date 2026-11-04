from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

import calculator_service
import config_workbook
from calculator_service import CostCalculatorService


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "成本配置模板.xlsx"


class ServiceConfigReloadTests(unittest.TestCase):
    def test_first_run_template_supports_recognition_and_batch_export(self) -> None:
        cases = (
            ("B-1，桃皮绒50150FT枕套+枕芯50150", 26, 5.5, 29),
            ("B-2，短毛绒50150开孔枕套", 26, 0, 0),
            ("B-3，2way50160开胸枕套", 35, 0, 0),
            ("B-4，new2way40120开孔开胸枕套", 38, 0, 0),
            ("B-5，桃皮绒50150枕套", 20, 0, 0),
            ("B-6，防粉色50150", 15, 0, 0),
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source = directory / "orders.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["订单编号", "卖家备注", "商家/店铺"])
            for index, (note, *_costs) in enumerate(cases, start=1):
                worksheet.append([str(index), note, "测试店铺"])
            workbook.save(source)
            workbook.close()

            service = CostCalculatorService()
            with patch.object(config_workbook, "_application_dir", return_value=directory):
                for note, base, dropship, pillow in cases:
                    with self.subTest(note=note):
                        result = service.recognize_note(note)
                        self.assertEqual([], result.failure_reasons)
                        self.assertEqual((base, dropship, base + dropship, 0, pillow, 0),
                                         result.as_tuple())

                self.assertTrue((directory / "成本配置.xlsx").is_file())
                outputs = service.process_targets([str(source)], str(directory / "output"), False)

            self.assertEqual(1, len(outputs))
            output = outputs[0]
            self.assertTrue(output.success, output.message)
            self.assertEqual(len(cases), output.matched_count)
            self.assertEqual(0, output.unmatched_count)
            workbook = load_workbook(output.output_path)
            try:
                self.assertIn("店铺统计", workbook.sheetnames)
                sheet = workbook["成本明细"]
                headers = {cell.value: cell.column for cell in sheet[1]}
                for row, (_note, base, dropship, pillow) in enumerate(cases, start=2):
                    with self.subTest(row=row):
                        self.assertEqual(base, sheet.cell(row, headers["成本"]).value)
                        self.assertEqual(dropship, sheet.cell(row, headers["代发成本"]).value)
                        self.assertEqual(pillow, sheet.cell(row, headers["枕芯成本"]).value)
                        self.assertIsNone(sheet.cell(row, headers["无法匹配原因说明"]).value)
            finally:
                workbook.close()

    def test_each_batch_reloads_saved_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "成本配置.xlsx"
            shutil.copyfile(TEMPLATE, config_path)
            service = CostCalculatorService()

            with patch.object(
                calculator_service,
                "ensure_editable_config",
                return_value=config_path,
            ):
                first = service.load_configs()
                self.assertEqual(19, first.pillow_cost_data["120*40"]["枕芯"])

                workbook = load_workbook(config_path)
                worksheet = workbook["枕芯价格"]
                for row in range(2, worksheet.max_row + 1):
                    if (
                        worksheet.cell(row, 1).value == "120*40"
                        and worksheet.cell(row, 2).value == "PP棉"
                    ):
                        worksheet.cell(row, 3, 20)
                        break
                workbook.save(config_path)
                workbook.close()

                self.assertEqual([], service.process_targets([], None, False))
                self.assertEqual(
                    20, service.config.pillow_cost_data["120*40"]["枕芯"]
                )


if __name__ == "__main__":
    unittest.main()

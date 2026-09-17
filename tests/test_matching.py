from __future__ import annotations

import unittest
from pathlib import Path

from config_workbook import load_config_workbook
from cost_calculator import match_price, match_price_detailed


ROOT = Path(__file__).resolve().parents[1]


class MatchingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.config = load_config_workbook(
            ROOT / "成本配置模板.xlsx", check_lock=False
        )

    def detailed(self, note: str):
        config = self.config
        return match_price_detailed(
            note,
            config.price_data,
            config.moving_costs_data,
            config.pillow_cost_data,
            config.others_cost_data,
            config.dropship_unit_cost,
            config.dropship_keywords,
        )

    def test_existing_cost_examples_are_unchanged(self) -> None:
        cases = {
            "B-14257，短毛绒50150枕套+枕芯50150": (20.0, 5.5, 26.0),
            "H067732B，SSS60180KK枕套+KK枕芯60180": (73.0, 5.5, 45.0),
            "C063756B，2way40120枕套＋加重羽绒棉枕芯40120＋香水": (
                23.0,
                5.5,
                30.0,
            ),
        }
        for note, expected in cases.items():
            with self.subTest(note=note):
                result = self.detailed(note)
                self.assertEqual(expected, (result.base_cost, result.dropship_cost, result.pillow_cost))
                self.assertEqual([], result.failure_reasons)

    def test_all_configured_pillows_match(self) -> None:
        for size, keyword_prices in self.config.pillow_cost_data.items():
            reverse = "*".join(reversed(size.split("*")))
            for keyword, expected in keyword_prices.items():
                for size_text in (size, size.replace("*", "x"), size.replace("*", ""), reverse):
                    with self.subTest(keyword=keyword, size=size_text):
                        result = self.detailed(f"{keyword}{size_text}")
                        self.assertEqual(float(expected), result.pillow_cost)
                        self.assertEqual([], result.failure_reasons)

    def test_legacy_tuple_interface_is_preserved(self) -> None:
        config = self.config
        result = match_price(
            "枕芯50150",
            config.price_data,
            config.moving_costs_data,
            config.pillow_cost_data,
            config.others_cost_data,
            config.dropship_unit_cost,
            config.dropship_keywords,
        )
        self.assertIsInstance(result, tuple)
        self.assertEqual(6, len(result))
        self.assertEqual(26.0, result[4])

    def test_field_level_reasons(self) -> None:
        cases = {
            "": "卖家备注为空",
            "起毛套50160": "未识别任何可计费项目",
            "B-1，SAS50150枕套": "缺少或不支持材质",
            "B-1，桃皮绒99999枕套": "缺少或不支持尺寸",
            "B-1，枕芯50150": "未识别商品类型关键词",
            "ft枕芯": "枕芯成本：缺少或不支持",
            "二yr9999": "义乳/义臀成本：未匹配完整规格",
        }
        for note, expected in cases.items():
            with self.subTest(note=note):
                result = self.detailed(note)
                self.assertTrue(any(expected in reason for reason in result.failure_reasons))

    def test_qimao_is_ignored_when_another_item_matches(self) -> None:
        result = self.detailed("三yr3250+起毛套50160")
        self.assertEqual(99.0, result.moving_cost)
        self.assertEqual([], result.failure_reasons)

    def test_reason_fragment_is_limited_to_40_characters(self) -> None:
        result = self.detailed("未知" * 30)
        self.assertEqual(1, len(result.failure_reasons))
        fragment = result.failure_reasons[0].split("片段：", 1)[1][:-1]
        self.assertLessEqual(len(fragment), 40)
        self.assertTrue(fragment.endswith("..."))


if __name__ == "__main__":
    unittest.main()

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

    def detailed(self, note: str, *, collect_details: bool = False):
        config = self.config
        return match_price_detailed(
            note,
            config.price_data,
            config.moving_costs_data,
            config.pillow_cost_data,
            config.others_cost_data,
            config.dropship_unit_cost,
            config.dropship_keywords,
            collect_details=collect_details,
        )

    def test_merged_category_keywords_use_shared_price_matrix(self) -> None:
        keywords = (
            "分腿等身枕套", "ft等身枕套", "分腿枕套", "ft枕套", "ft不开孔枕套",
            "开胸开孔枕套", "kxkk枕套", "开孔开胸枕套", "kkkx枕套",
            "功能版枕套", "功能枕套", "开孔等身枕套", "kk等身枕套", "开孔枕套",
            "kk枕套", "开胸等身枕套", "kx等身枕套", "开胸枕套", "kx枕套",
        )
        sizes = ("40*120", "50*150", "50*160", "60*180")
        prices = {
            "桃皮绒": (22, 26, 30, 35),
            "短毛绒": (22, 26, 30, 35),
            "2way": (28, 32, 35, 45),
            "new2way": (38, 41, 45, 57),
            "SSS": (50, 56, 57, 73),
        }
        for keyword in keywords:
            for material, costs in prices.items():
                for size, expected in zip(sizes, costs):
                    with self.subTest(keyword=keyword, material=material, size=size):
                        result = self.detailed(
                            f"B-1，{material}{size}{keyword}", collect_details=True
                        )
                        self.assertEqual([], result.failure_reasons)
                        self.assertEqual(float(expected), result.base_cost)
                        self.assertEqual(float(expected), result.total_cost)
                        detail, = result.details
                        self.assertEqual("开孔/开胸/开孔开胸/分腿枕套", detail.category)
                        self.assertEqual(keyword, detail.keyword)
                        self.assertEqual(material, detail.material)
                        self.assertTrue(detail.included)

    def test_merged_category_has_priority_over_generic_pillowcase(self) -> None:
        for keyword in ("ft枕套", "开孔枕套", "开胸枕套", "开孔开胸枕套"):
            with self.subTest(keyword=keyword):
                result = self.detailed(
                    f"B-1，桃皮绒50150枕套+{keyword}", collect_details=True
                )
                self.assertEqual([], result.failure_reasons)
                self.assertEqual(26.0, result.base_cost)
                self.assertEqual("开孔/开胸/开孔开胸/分腿枕套", result.details[0].category)

    def test_merged_category_quantities_and_segments(self) -> None:
        result = self.detailed(
            "B-1，B-2，短毛绒50150FT枕套。B-3，2WAY50X160开胸枕套",
            collect_details=True,
        )
        self.assertEqual([], result.failure_reasons)
        self.assertEqual(87.0, result.base_cost)
        self.assertEqual(87.0, result.total_cost)
        self.assertEqual([2, 1], [detail.quantity for detail in result.details])
        self.assertEqual([26.0, 35.0], [detail.unit_cost for detail in result.details])

    def test_cost_examples_use_updated_pillow_prices(self) -> None:
        cases = {
            "B-14257，短毛绒50150枕套+枕芯50150": (20.0, 5.5, 29.0),
            "H067732B，SSS60180KK枕套+KK枕芯60180": (73.0, 5.5, 49.0),
            "C063756B，2way40120枕套＋加重羽绒棉枕芯40120＋香水": (
                23.0,
                5.5,
                32.0,
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
        self.assertEqual(29.0, result[4])

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

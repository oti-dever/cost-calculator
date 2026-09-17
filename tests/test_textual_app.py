from __future__ import annotations

import asyncio
import unittest
from unittest.mock import patch

from calculator_service import CostCalculatorService
from textual_app import CostCalculatorTextualApp


class TextualAppTests(unittest.TestCase):
    def test_open_config_button_is_present(self) -> None:
        async def run() -> None:
            app = CostCalculatorTextualApp()
            with patch.object(CostCalculatorService, "load_configs", return_value=None):
                async with app.run_test(size=(120, 40)):
                    button = app.query_one("#open-config-btn")
                    self.assertEqual("打开配置", str(button.label))

        asyncio.run(run())


if __name__ == "__main__":
    unittest.main()

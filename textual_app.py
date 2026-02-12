"""Textual 终端界面入口。"""

from __future__ import annotations

from typing import List

from textual import on, work
from textual.app import App, ComposeResult
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import ModalScreen
from textual.widgets import (
    Button,
    Checkbox,
    DataTable,
    Footer,
    Header,
    Input,
    Label,
    RichLog,
    Static,
)

from calculator_service import CostCalculatorService
from cost_calculator import ProcessResult


def _truncate_text(text: str, limit: int = 32) -> str:
    """截断过长文本，保留摘要展示。"""

    text = text.strip()
    if len(text) <= limit:
        return text
    return f"{text[: limit - 1]}…"


class LogDetailScreen(ModalScreen[None]):
    """日志详情弹窗。"""

    CSS = """
    LogDetailScreen {
        align: center middle;
    }

    #log-detail-dialog {
        width: 88%;
        height: 80%;
        border: round $accent;
        background: $surface;
        padding: 1;
    }

    #log-detail-header {
        height: auto;
        margin-bottom: 1;
    }

    #log-detail-content {
        height: 1fr;
        border: round $panel;
        margin-bottom: 0;
        padding: 1;
    }

    #log-detail-actions {
        height: auto;
        margin-top: 1;
    }

    #log-detail-spacer {
        width: 1fr;
    }

    .file-title {
        margin-top: 1;
        margin-bottom: 1;
        text-style: bold;
        color: $accent;
    }

    .file-table {
        height: 10;
        margin-bottom: 1;
    }
    """

    def __init__(self, results: List[ProcessResult]) -> None:
        super().__init__()
        self.results = results

    def compose(self) -> ComposeResult:
        """构建弹窗。"""

        with Vertical(id="log-detail-dialog"):
            yield Static("日志详情（按文件分组表格）", id="log-detail-header")
            with VerticalScroll(id="log-detail-content"):
                for index, item in enumerate(self.results):
                    yield Static(
                        f"文件 {index + 1}: {item.input_path}",
                        classes="file-title",
                    )
                    yield DataTable(id=f"log-detail-table-{index}", classes="file-table")
            with Horizontal(id="log-detail-actions"):
                yield Static("", id="log-detail-spacer")
                yield Button("关闭", id="close-log-detail-btn", variant="primary")

    def on_mount(self) -> None:
        """写入日志内容。"""

        if not self.results:
            return

        for index, item in enumerate(self.results):
            table = self.query_one(f"#log-detail-table-{index}", DataTable)
            table.add_columns("行号", "订单编号", "原因", "备注")

            if not item.unmatched_details:
                table.add_row("-", "-", "无未匹配记录", "-")
                continue

            for detail in item.unmatched_details:
                table.add_row(
                    str(detail.get("row_number", "")),
                    str(detail.get("order_id", "无")),
                    str(detail.get("reason", "")),
                    str(detail.get("seller_note", "")),
                )

    @on(Button.Pressed, "#close-log-detail-btn")
    def close_detail(self) -> None:
        """关闭弹窗。"""

        self.dismiss()


class CostCalculatorTextualApp(App[None]):
    """成本计算器 Textual 界面。"""

    CSS = """
    Screen {
        layout: vertical;
    }

    #main {
        height: 1fr;
        padding: 1 2;
    }

    #controls {
        height: auto;
        border: round $accent;
        padding: 1;
        margin-bottom: 1;
    }

    .row {
        height: auto;
        margin-bottom: 1;
    }

    .label {
        width: 22;
        content-align: right middle;
    }

    .input {
        width: 1fr;
    }

    #actions {
        height: auto;
        margin-top: 1;
    }

    #result-table {
        height: 10;
        margin-top: 1;
    }

    #result-detail {
        height: 8;
        border: round $primary;
        padding: 0 1;
        margin-top: 1;
    }

    #log {
        height: 1fr;
        border: round $surface;
        padding: 0 1;
        margin-top: 1;
    }
    """

    BINDINGS = [("q", "quit", "退出")]

    def __init__(self) -> None:
        super().__init__()
        self.service = CostCalculatorService()
        self.processing = False
        self.latest_results: List[ProcessResult] = []
        self.log_lines: List[str] = []

    def compose(self) -> ComposeResult:
        """构建界面组件。"""

        yield Header(show_clock=True)

        with Vertical(id="main"):
            with Vertical(id="controls"):
                yield Static("成本计算器（Textual 版）", classes="title")

                with Horizontal(classes="row"):
                    yield Label("输入目标（文件/目录）:", classes="label")
                    yield Input(
                        placeholder='示例: "test_excel" 或 "a.xlsx b.xlsx"',
                        id="targets-input",
                        classes="input",
                    )

                with Horizontal(classes="row"):
                    yield Label("输出目录（可留空）:", classes="label")
                    yield Input(
                        placeholder="留空表示输出到源文件目录",
                        id="output-dir-input",
                        classes="input",
                    )

                with Horizontal(classes="row"):
                    yield Label("覆盖同名文件:", classes="label")
                    yield Checkbox("overwrite", id="overwrite-checkbox")

                with Horizontal(id="actions"):
                    yield Button("开始处理", id="run-btn", variant="success")
                    yield Button("清空日志", id="clear-log-btn", variant="primary")
                    yield Button("查看日志详情", id="show-log-detail-btn")

            yield DataTable(id="result-table")
            yield Static("结果详情：请选择一条结果记录", id="result-detail")
            yield RichLog(id="log", highlight=True, wrap=True, markup=False)

        yield Footer()

    def on_mount(self) -> None:
        """初始化表格与配置加载。"""

        table = self.query_one("#result-table", DataTable)
        table.add_columns("状态", "输入", "输出", "匹配", "说明")

        log = self.query_one("#log", RichLog)
        log.write("正在加载配置...")

        try:
            self.service.load_configs(log=self._write_log)
            log.write("配置加载完成，可开始处理。")
        except Exception as exc:
            log.write(f"配置加载失败: {exc}")

    @on(Button.Pressed, "#clear-log-btn")
    def clear_log(self) -> None:
        """清空日志与结果表。"""

        log = self.query_one("#log", RichLog)
        log.clear()

        table = self.query_one("#result-table", DataTable)
        table.clear()

        self.latest_results = []
        self.log_lines = []
        detail = self.query_one("#result-detail", Static)
        detail.update("结果详情：请选择一条结果记录")

    @on(Button.Pressed, "#show-log-detail-btn")
    def show_log_detail(self) -> None:
        """显示日志详情弹窗。"""

        self.push_screen(LogDetailScreen(self.latest_results))

    @on(DataTable.RowSelected, "#result-table")
    def on_result_row_selected(self, event: DataTable.RowSelected) -> None:
        """选择结果行时，更新详情面板。"""

        row_index = event.cursor_row
        if row_index < 0 or row_index >= len(self.latest_results):
            return

        item = self.latest_results[row_index]
        matched_text = ""
        if item.matched_count is not None and item.total_count is not None:
            matched_text = f"{item.matched_count}/{item.total_count}"

        detail_lines = [
            "结果详情",
            f"状态: {item.status}",
            f"输入: {item.input_path}",
            f"输出: {item.output_path or '无'}",
            f"匹配: {matched_text or '无'}",
            f"未匹配: {item.unmatched_count}",
            f"说明: {item.message}",
        ]

        detail = self.query_one("#result-detail", Static)
        detail.update("\n".join(detail_lines))

    @on(Button.Pressed, "#run-btn")
    def trigger_run(self) -> None:
        """触发处理任务。"""

        if self.processing:
            self._write_log("已有任务在运行，请稍候。")
            return

        raw_targets = self.query_one("#targets-input", Input).value.strip()
        if not raw_targets:
            self._write_log("请先输入目标路径。")
            return

        output_dir = self.query_one("#output-dir-input", Input).value.strip() or None
        overwrite = self.query_one("#overwrite-checkbox", Checkbox).value

        targets = self.service.parse_target_input(raw_targets)
        if not targets:
            self._write_log("未解析到有效路径，请检查输入。")
            return

        self.processing = True
        run_btn = self.query_one("#run-btn", Button)
        run_btn.disabled = True

        self._write_log(f"开始处理，共 {len(targets)} 个输入目标...")
        self.run_processing(targets, output_dir, overwrite)

    @work(thread=True)
    def run_processing(
        self,
        targets: List[str],
        output_dir: str | None,
        overwrite: bool,
    ) -> None:
        """后台线程执行批处理。"""

        try:
            results = self.service.process_targets(
                targets=targets,
                output_dir=output_dir,
                overwrite=overwrite,
                log=self._write_log_from_worker,
            )
            self.call_from_thread(self._update_result_table, results)
            self.call_from_thread(self._write_summary, results)
        except Exception as exc:
            self.call_from_thread(self._write_log, f"处理异常: {exc}")
        finally:
            self.call_from_thread(self._finish_processing)

    def _finish_processing(self) -> None:
        """恢复界面状态。"""

        self.processing = False
        run_btn = self.query_one("#run-btn", Button)
        run_btn.disabled = False

    def _write_log_from_worker(self, message: str) -> None:
        """后台线程写日志。"""

        self.call_from_thread(self._write_log, message)

    def _write_log(self, message: str) -> None:
        """向日志面板写入一行。"""

        self.log_lines.append(message)
        log = self.query_one("#log", RichLog)
        log.write(message)

    def _update_result_table(self, results: List[ProcessResult]) -> None:
        """刷新结果表格。"""

        self.latest_results = results
        table = self.query_one("#result-table", DataTable)
        table.clear()

        for item in results:
            status_text = {
                "success": "成功",
                "failed": "失败",
                "skipped": "跳过",
            }.get(item.status, item.status)

            matched_text = ""
            if item.matched_count is not None and item.total_count is not None:
                matched_text = f"{item.matched_count}/{item.total_count}"

            table.add_row(
                status_text,
                _truncate_text(item.input_path, 36),
                _truncate_text(item.output_path or "", 36),
                matched_text,
                _truncate_text(item.message, 28),
            )

        detail = self.query_one("#result-detail", Static)
        if results:
            detail.update("结果详情：请选择一条结果记录")
        else:
            detail.update("结果详情：暂无处理结果")

    def _write_summary(self, results: List[ProcessResult]) -> None:
        """输出批处理汇总。"""

        if not results:
            self._write_log("本批次没有可处理项。")
            return

        success_count = sum(1 for r in results if r.status == "success")
        failed_count = sum(1 for r in results if r.status == "failed")
        skipped_count = sum(1 for r in results if r.status == "skipped")
        overseas_count = sum(r.overseas_count for r in results if r.status == "success")

        self._write_log(
            f"批处理汇总: 总计 {len(results)} | 成功 {success_count} "
            f"| 失败 {failed_count} | 跳过 {skipped_count}"
        )

        if overseas_count > 0:
            self._write_log(f"海外订单累计: {overseas_count} 条")


def main() -> None:
    """运行 Textual 应用。"""

    CostCalculatorTextualApp().run()


if __name__ == "__main__":
    main()

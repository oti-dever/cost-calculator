"""Textual 终端界面入口。"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from typing import List, Literal

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
from config_workbook import ensure_editable_config
from cost_calculator import ProcessResult


PickerMode = Literal["file", "directory"]
EXCEL_SUFFIXES = {".xlsx", ".xls"}


def _truncate_text(text: str, limit: int = 32) -> str:
    """截断过长文本，保留摘要展示。"""

    text = text.strip()
    if len(text) <= limit:
        return text
    return f"{text[: limit - 1]}…"


def _format_target_path(path: str) -> str:
    """格式化系统选择结果，确保带空格路径可被正确解析。"""

    normalized = os.path.normpath(path)
    if any(character.isspace() for character in normalized):
        return f'"{normalized}"'
    return normalized


def _existing_start_directory(start_path: str | Path | None) -> Path:
    """将已有输入转换为系统选择窗口可用的起始目录。"""

    fallback = Path.cwd().resolve()
    if start_path is None:
        return fallback

    try:
        expanded = os.path.expandvars(os.path.expanduser(str(start_path)))
        candidate = Path(expanded)
        if not candidate.is_absolute():
            candidate = fallback / candidate
        candidate = candidate.resolve(strict=False)
        if candidate.is_file():
            return candidate.parent
        if candidate.is_dir():
            return candidate
    except OSError:
        pass

    return fallback


def _show_system_target_dialog(
    mode: PickerMode,
    initial_directory: str,
) -> str | None:
    """打开操作系统文件或文件夹选择窗口。"""

    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError("当前 Python 环境未提供 Tk 系统对话框组件") from exc

    root = tk.Tk()
    root.withdraw()
    try:
        try:
            root.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        root.update_idletasks()

        if mode == "file":
            selected_path = filedialog.askopenfilename(
                parent=root,
                title="选择输入 Excel 文件",
                initialdir=initial_directory,
                filetypes=[("Excel 文件", "*.xlsx *.xls")],
            )
        else:
            selected_path = filedialog.askdirectory(
                parent=root,
                title="选择输入文件夹",
                initialdir=initial_directory,
                mustexist=True,
            )
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass

    if not selected_path:
        return None
    return os.path.normpath(str(selected_path))


def _open_with_default_app(path: str | Path) -> None:
    """使用操作系统默认应用打开文件。"""

    resolved = str(Path(path).resolve())
    if hasattr(os, "startfile"):
        os.startfile(resolved)  # type: ignore[attr-defined]
        return
    if sys.platform == "darwin":
        subprocess.Popen(["open", resolved])
        return
    subprocess.Popen(["xdg-open", resolved])


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

    #targets-input {
        min-width: 16;
    }

    #select-file-btn {
        width: 12;
        min-width: 12;
        margin-left: 1;
    }

    #select-directory-btn {
        width: 14;
        min-width: 14;
        margin-left: 1;
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
        self.target_dialog_open = False
        self.latest_results: List[ProcessResult] = []
        self.log_lines: List[str] = []

    def compose(self) -> ComposeResult:
        """构建界面组件。"""

        yield Header(show_clock=True)

        with Vertical(id="main"):
            with Vertical(id="controls"):
                yield Static("成本计算器（Textual 版）", classes="title")

                with Horizontal(classes="row"):
                    yield Label("输入文件/文件夹:", classes="label")
                    yield Input(
                        placeholder="输入路径，或使用右侧按钮选择",
                        id="targets-input",
                        classes="input",
                    )
                    yield Button(
                        "选择文件",
                        id="select-file-btn",
                        variant="primary",
                    )
                    yield Button("选择文件夹", id="select-directory-btn")

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
                    yield Button("打开配置", id="open-config-btn", variant="primary")
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

    @on(Button.Pressed, "#open-config-btn")
    def open_config(self) -> None:
        """创建（如有需要）并打开可编辑成本配置。"""

        try:
            config_path = ensure_editable_config()
            _open_with_default_app(config_path)
            self._write_log(f"已打开成本配置：{config_path}")
            self._write_log("修改后请保存并关闭 Excel，再开始处理。")
        except Exception as exc:
            self._write_log(f"打开成本配置失败：{exc}")

    @on(Button.Pressed, "#show-log-detail-btn")
    def show_log_detail(self) -> None:
        """显示日志详情弹窗。"""

        self.push_screen(LogDetailScreen(self.latest_results))

    @on(Button.Pressed, "#select-file-btn")
    def choose_target_file(self) -> None:
        """打开系统 Excel 文件选择窗口。"""

        self._start_target_dialog("file")

    @on(Button.Pressed, "#select-directory-btn")
    def choose_target_directory(self) -> None:
        """打开系统文件夹选择窗口。"""

        self._start_target_dialog("directory")

    def _start_target_dialog(self, mode: PickerMode) -> None:
        """准备并启动后台系统选择窗口。"""

        if self.target_dialog_open:
            return

        raw_value = self.query_one("#targets-input", Input).value.strip()
        parsed_targets = self.service.parse_target_input(raw_value)
        start_path = parsed_targets[0] if parsed_targets else None
        initial_directory = _existing_start_directory(start_path)

        self.target_dialog_open = True
        self._set_target_dialog_buttons_disabled(True)
        target_kind = "文件" if mode == "file" else "文件夹"
        self._write_log(f"正在打开系统{target_kind}选择窗口...")

        try:
            selected_path = _show_system_target_dialog(
                mode,
                str(initial_directory),
            )
            if selected_path is None:
                self._handle_target_dialog_cancelled()
            else:
                self._replace_target_input(
                    selected_path,
                    mode,
                )
        except Exception as exc:
            self._handle_target_dialog_error(mode, exc)
        finally:
            self._finish_target_dialog()

    def _replace_target_input(
        self,
        selected_path: str,
        mode: PickerMode,
    ) -> None:
        """使用系统选择结果覆盖目标输入框。"""

        is_valid_file = (
            mode == "file"
            and os.path.isfile(selected_path)
            and Path(selected_path).suffix.lower() in EXCEL_SUFFIXES
        )
        is_valid_directory = mode == "directory" and os.path.isdir(selected_path)
        if not (is_valid_file or is_valid_directory):
            self._write_log(f"选择结果已失效或类型不受支持: {selected_path}")
            return

        targets_input = self.query_one("#targets-input", Input)
        formatted_path = _format_target_path(selected_path)
        targets_input.value = formatted_path
        targets_input.cursor_position = len(formatted_path)
        targets_input.focus()

        target_kind = "文件" if mode == "file" else "文件夹"
        self._write_log(f"已选择输入{target_kind}，原输入已覆盖: {selected_path}")

    def _handle_target_dialog_cancelled(self) -> None:
        """处理用户取消系统选择窗口。"""

        self._write_log("已取消选择，原输入保持不变。")

    def _handle_target_dialog_error(
        self,
        mode: PickerMode,
        error: Exception,
    ) -> None:
        """记录系统选择窗口启动失败。"""

        target_kind = "文件" if mode == "file" else "文件夹"
        self._write_log(f"无法打开系统{target_kind}选择窗口: {error}")

    def _set_target_dialog_buttons_disabled(self, disabled: bool) -> None:
        """统一设置两个系统选择按钮的禁用状态。"""

        self.query_one("#select-file-btn", Button).disabled = disabled
        self.query_one("#select-directory-btn", Button).disabled = disabled

    def _finish_target_dialog(self) -> None:
        """恢复系统选择按钮状态。"""

        self.target_dialog_open = False
        self._set_target_dialog_buttons_disabled(False)

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
        self.query_one("#open-config-btn", Button).disabled = True

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
        self.query_one("#open-config-btn", Button).disabled = False

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

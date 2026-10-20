"""单条备注识别测试工具，也可嵌入主程序。"""

from __future__ import annotations

from textual import on, work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, VerticalScroll
from textual.screen import ModalScreen
from textual.widgets import Button, DataTable, Footer, Static, TextArea

from calculator_service import CostCalculatorService
from config_workbook import ensure_editable_config, get_editable_config_path
from cost_calculator import MatchDetail, MatchResult
from ui_helpers import open_with_default_app


COST_LABELS = {
    "base_cost": "基础成本",
    "dropship_cost": "代发成本",
    "moving_cost": "义乳/义臀成本",
    "pillow_cost": "枕芯成本",
    "other_cost": "硅胶/电动成本",
    "total_cost": "总成本（基础＋代发）",
}


def recognition_status(result: MatchResult) -> str:
    """根据最终结果判断状态，不把关键词命中当作成功。"""

    if result.is_overseas:
        return "海外需人工核对"
    if result.failure_reasons:
        return "部分匹配" if any(value is not None for value in result.as_tuple()) else "识别失败"
    return "识别成功" if result.matched_any else "识别失败"


def excel_cost_values(result: MatchResult) -> dict[str, float | None]:
    """复现正式输出的空值口径，包括非海外行其他成本的零值。"""

    values = {name: getattr(result, name) for name in COST_LABELS}
    if not result.is_overseas:
        values["other_cost"] = result.other_cost if result.other_cost and result.other_cost > 0 else 0
    return values


def money(value: float | None) -> str:
    return "—（留空）" if value is None else f"{value:.2f}"


def detail_spec(detail: MatchDetail) -> str:
    return " / ".join(filter(None, (detail.category, detail.keyword, detail.size, detail.material)))


class RemarkTesterScreen(ModalScreen[None]):
    """主程序和独立程序共用的备注测试界面。"""

    CSS = """
    RemarkTesterScreen { background: $surface; }
    #remark-content { padding: 1 2; }
    #remark-title { text-style: bold; color: $accent; height: auto; }
    #remark-config { height: auto; margin-bottom: 1; }
    #remark-input { height: 6; min-height: 4; }
    #remark-actions { height: auto; margin: 1 0; }
    #remark-actions Button { margin-right: 1; min-width: 10; }
    #remark-status { height: auto; text-style: bold; margin-bottom: 1; }
    #remark-costs { height: auto; border: round $primary; padding: 0 1; }
    #remark-details { height: 10; margin-top: 1; }
    #remark-selected { height: auto; min-height: 3; border: round $panel; padding: 0 1; }
    #remark-reasons { height: auto; min-height: 3; margin-top: 1; }
    """

    BINDINGS = [
        Binding("ctrl+enter", "recognize", "识别", priority=True),
        Binding("escape", "close", "关闭", priority=True),
    ]

    def __init__(self, *, standalone: bool = False) -> None:
        super().__init__()
        self.standalone = standalone
        self.service = CostCalculatorService()
        self.processing = False
        self.last_result: MatchResult | None = None
        self.submitted_text: str | None = None

    def compose(self) -> ComposeResult:
        with VerticalScroll(id="remark-content"):
            yield Static("备注识别测试", id="remark-title")
            yield Static(f"共用配置：{get_editable_config_path()}", id="remark-config", markup=False)
            yield Static("输入一条完整卖家备注（换行仍属于同一条备注）：")
            yield TextArea(id="remark-input")
            with Horizontal(id="remark-actions"):
                yield Button("识别", id="remark-recognize", variant="success")
                yield Button("清空", id="remark-clear")
                yield Button("打开配置", id="remark-open-config", variant="primary")
                yield Button("关闭", id="remark-close")
            yield Static("请输入备注，点击识别或按 Ctrl+Enter。", id="remark-status", markup=False)
            yield Static("识别后显示各项成本；总成本＝基础成本＋代发成本。", id="remark-costs", markup=False)
            yield DataTable(id="remark-details", cursor_type="row")
            yield Static("选择明细行查看完整规格和备注分段。", id="remark-selected", markup=False)
            yield Static("失败原因：暂无", id="remark-reasons", markup=False)
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#remark-details", DataTable).add_columns(
            "分段", "成本项", "命中规格", "数量", "单价", "小计", "计入状态",
        )
        self.query_one("#remark-input", TextArea).focus()
        try:
            ensure_editable_config()
        except Exception as exc:
            self.query_one("#remark-status", Static).update(f"配置准备失败：{exc}")

    @on(Button.Pressed, "#remark-recognize")
    def action_recognize(self) -> None:
        if self.processing:
            return
        text = self.query_one("#remark-input", TextArea).text
        if not text.strip():
            self.reset_result()
            self.query_one("#remark-status", Static).update("请输入备注。")
            return
        self.reset_result()
        self.submitted_text = text
        self.set_processing(True)
        self.query_one("#remark-status", Static).update("正在加载配置并识别……")
        self.run_recognition(text)

    @work(thread=True, exclusive=True, exit_on_error=False)
    def run_recognition(self, text: str) -> None:
        try:
            result = self.service.recognize_note(text)
        except Exception as exc:
            self.app.call_from_thread(self.finish_recognition, None, str(exc))
        else:
            self.app.call_from_thread(self.finish_recognition, result, None)

    def set_processing(self, processing: bool) -> None:
        self.processing = processing
        for selector in ("#remark-input", "#remark-recognize", "#remark-clear", "#remark-open-config", "#remark-close"):
            self.query_one(selector).disabled = processing

    def finish_recognition(self, result: MatchResult | None, error: str | None) -> None:
        self.set_processing(False)
        self.last_result = result
        if result is None:
            self.query_one("#remark-status", Static).update("识别失败：配置或处理错误")
            self.query_one("#remark-reasons", Static).update(f"失败原因：{error}")
            return
        self.query_one("#remark-status", Static).update(recognition_status(result))
        self.query_one("#remark-costs", Static).update("\n".join(
            f"{COST_LABELS[name]}：{money(value)}" for name, value in excel_cost_values(result).items()
        ))
        table = self.query_one("#remark-details", DataTable)
        for index, detail in enumerate(result.details):
            table.add_row(
                str(detail.segment_index), COST_LABELS[detail.cost_type], detail_spec(detail),
                str(detail.quantity), money(detail.unit_cost), money(detail.subtotal),
                "已计入" if detail.included else "未计入", key=str(index),
            )
        self.query_one("#remark-reasons", Static).update(
            "\n".join(result.failure_reasons) if result.failure_reasons else "失败原因：无"
        )
        if result.details:
            self.render_detail(result.details[0])

    @on(TextArea.Changed, "#remark-input")
    def on_note_changed(self) -> None:
        if self.submitted_text is not None and self.query_one("#remark-input", TextArea).text != self.submitted_text:
            self.query_one("#remark-status", Static).update("输入已修改，请重新识别。")

    @on(DataTable.RowHighlighted, "#remark-details")
    @on(DataTable.RowSelected, "#remark-details")
    def show_detail(self, event: DataTable.RowHighlighted | DataTable.RowSelected) -> None:
        if self.last_result is None or event.row_key.value is None:
            return
        detail = self.last_result.details[int(event.row_key.value)]
        self.render_detail(detail)

    def render_detail(self, detail: MatchDetail) -> None:
        """完整文本在可换行区域显示，不依赖表格列宽。"""

        state = "已计入" if detail.included else "未计入（所属成本项或整条备注识别失败）"
        self.query_one("#remark-selected", Static).update(
            f"第{detail.segment_index}段 · {COST_LABELS[detail.cost_type]} · {state}\n"
            f"命中规格：{detail_spec(detail)}\n"
            f"数量 {detail.quantity} × 单价 {money(detail.unit_cost)} = {money(detail.subtotal)}\n"
            f"原始分段：{detail.segment_text}"
        )

    def reset_result(self) -> None:
        self.last_result = None
        self.submitted_text = None
        self.query_one("#remark-details", DataTable).clear()
        self.query_one("#remark-costs", Static).update("识别后显示各项成本；总成本＝基础成本＋代发成本。")
        self.query_one("#remark-selected", Static).update("选择明细行查看完整规格和备注分段。")
        self.query_one("#remark-reasons", Static).update("失败原因：暂无")

    @on(Button.Pressed, "#remark-clear")
    def clear_note(self) -> None:
        self.reset_result()
        note = self.query_one("#remark-input", TextArea)
        note.load_text("")
        note.focus()
        self.query_one("#remark-status", Static).update("请输入备注，点击识别或按 Ctrl+Enter。")

    @on(Button.Pressed, "#remark-open-config")
    def open_config(self) -> None:
        try:
            open_with_default_app(ensure_editable_config())
            self.query_one("#remark-status", Static).update("修改配置后请保存并关闭 Excel，再点击识别。")
        except Exception as exc:
            self.query_one("#remark-status", Static).update(f"打开配置失败：{exc}")

    @on(Button.Pressed, "#remark-close")
    def action_close(self) -> None:
        if self.processing:
            return
        if self.standalone:
            self.app.exit()
        else:
            self.dismiss()


class RemarkTesterApp(App[None]):
    TITLE = "备注识别测试"

    def on_mount(self) -> None:
        self.push_screen(RemarkTesterScreen(standalone=True))


def main() -> None:
    RemarkTesterApp().run()


if __name__ == "__main__":
    main()

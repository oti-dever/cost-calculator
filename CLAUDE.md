# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development commands

Run commands from the repository root. This is a flat Python application with a Textual terminal UI, pandas/openpyxl Excel processing, and Windows PyInstaller distribution. `.python-version` selects Python 3.13; `pyproject.toml` declares Python >=3.10, but the pinned dependencies may impose a newer minimum.

```bash
uv venv
uv pip install -r requirements.txt

# Main application (opens the Textual UI, not an argument-driven batch CLI)
uv run --no-sync python cost_calculator.py

# Standalone single-note recognition UI
uv run --no-sync python remark_tester.py

# All tests
uv run --no-sync python -m unittest discover -s tests -v

# One test module
uv run --no-sync python -m unittest tests.test_matching -v

# One test method
uv run --no-sync python -m unittest tests.test_matching.MatchingTests.test_legacy_tuple_interface_is_preserved -v
```

Runtime dependencies are pinned in `requirements.txt`, not declared in `pyproject.toml`; `uv.lock` contains only the local project. `uv sync` alone does not install the application dependencies. The commands above use `--no-sync` to run the explicitly provisioned environment. Installed console entry points are `cost-calculator`, `cost-calculator-tui`, and `remark-tester`. Root `main.py` is an unused hello-world scaffold, not the application entry point.

Tests use standard-library `unittest`, including asynchronous Textual `App.run_test()`/Pilot tests; pytest is not required. Tests load the checked-in template and create temporary order/configuration workbooks, so ignored `test_excel/` data is not required. No repository lint, formatter, or type-check command is configured.

### Windows executable build

From PowerShell or Command Prompt:

```text
uv run --no-sync cmd /c build.bat
```

From Git Bash, disable MSYS argument conversion for that invocation:

```bash
MSYS_NO_PATHCONV=1 uv run --no-sync cmd.exe /c build.bat
```

`build.bat` builds both `dist/CostCalculator.exe` and `dist/RemarkTester.exe` as console/onefile executables, embedding `成本配置模板.xlsx`. It prefers `.venv/Scripts/python.exe`, checks for Tk (needed by the system file pickers), pins `rich==13.9.4` for PyInstaller compatibility, and pauses interactively. It preserves existing `dist/` configuration/user data and copies `test_excel/` when present. Generated `.spec` files are ignored; the maintained build recipe is `build.bat`.

## Architecture and data flow

- **`config_workbook.py` — configuration boundary.** Validates the Excel configuration and translates its tables/matrices into a `CostConfig` snapshot used by the matching engine. It owns template discovery, editable-config creation, file-lock checks, and schema errors.
- **`calculator_service.py` — application orchestration.** `CostCalculatorService` parses target paths, expands directories nonrecursively, handles output naming/overwrite behavior, and calls the core. It reloads configuration before each batch and each single-note recognition. A batch uses one validated snapshot; invalid/locked configuration aborts before order output is generated, rather than falling back to stale prices.
- **`cost_calculator.py` — recognition and Excel adapters.** `recognize_note()` handles overseas orders and calls `match_price_detailed()`, which normalizes/splits remarks and combines the individual cost matchers. `MatchResult` holds amounts, failure reasons, and optional `MatchDetail` traces. `process_excel_file()` uses this same recognition path and returns a file-level `ProcessResult` with counts and unmatched-row details. File success does not imply every order matched. Legacy `match_price()` preserves a six-value tuple API, but does not apply the overseas guard; new full-note consumers should use `recognize_note()`.
- **`textual_app.py` / `remark_tester.py` — presentation.** The main app runs batch processing; `RemarkTesterScreen` works both as a modal in that app and inside standalone `RemarkTesterApp`. Both use the service rather than separate pricing implementations. Processing runs in Textual worker threads; UI updates return through `call_from_thread`. Native file pickers use Tk; `ui_helpers.py` opens configuration with the OS default application.

Tests cover the configuration schema, matching/compatibility behavior, Excel output, configuration reloads, and both UI entry points. Recognition-detail tests explicitly check that collecting traces does not change prices or failure reasons and that exported values agree with recognition results.

## Configuration lifecycle

`成本配置模板.xlsx` is the version-controlled default. `成本配置.xlsx` is an ignored, user-editable runtime copy created only when absent. In source mode it lives beside the source modules; in frozen mode it lives beside the executable, while the template is read from the bundle. This location is independent of the input workbook or service `base_path`. The two executables share configuration when placed together.

Updating the template or rebuilding does **not** migrate or overwrite an existing runtime configuration. Preserve user edits; testing new defaults requires a separate/fresh configuration or an explicitly backed-up migration. The old JSON configuration format is no longer supported.

The version-1 workbook requires eight named worksheets: `使用说明`, `全局参数`, `基础商品类别`, `基础商品关键词`, `基础商品价格`, `枕芯价格`, `义乳义臀价格`, and `其他成本`. Names/headers are schema, not presentation labels. Data formulas/merged cells, duplicate keys, invalid references, and negative/non-numeric prices are rejected. Blank matrix prices mean an unsupported combination, not zero. Close Excel before reloading: configuration loading rejects its lock file.

Base categories connect across the category, keyword, and price sheets; renaming/merging a category requires updating all three. The current template separates functional pillowcases into `开孔枕套`, `开胸枕套`, `开孔开胸枕套`, and `分腿枕套`, each with independent prices. Default priority is split-leg, combined opening, hole, chest, generic pillowcase, then cover; `功能枕套`/`功能版枕套` map to the combined category and `ft不开孔枕套` maps to split-leg. Prices and available specifications belong in the workbook, not hardcoded UI logic. Dropship unit price is configurable; its trigger keywords are fixed in code.

## Domain and Excel contracts

- Remarks are case-insensitive and have whitespace removed. Chinese `。` separates item segments; `+` is not a delimiter. Base quantity is the number of Chinese/ASCII commas in a segment, **not commas plus one**; a matched base item without commas has quantity one.
- Base matching selects one complete specification per segment in category-priority order. Pillow matching requires a configured keyword immediately followed by size; its size handling differs from base-product matching. Keep these matchers distinct rather than assuming one universal specification parser.
- Zero is a valid amount; `None` represents unavailable/failed cost. A pillow or moving-item failure can invalidate that entire cost component across segments; a fatal base failure can invalidate the whole row. Trace entries may therefore be recognized but marked `included=False`. Do not derive success solely from a nonempty trace or `matched_any`.
- Recognition/detail `总成本` is **基础成本 + 代发成本 only**. The other three components are displayed separately. The shop-summary total instead includes all five cost components; these are intentionally different existing output semantics.
- `发海外` bypasses pricing and leaves costs blank for manual review. Overseas records are counted separately from unmatched records. `起毛套` is not billed. The supported material keyword is `SSS`, not `SAS`.
- Input columns must include exactly named `卖家备注` and `商家/店铺`; diagnostic order IDs come only from optional `订单编号`.
- The `.xlsx` path preserves the workbook, writes `成本明细`, replaces `店铺统计`, and adds Excel formulas and unmatched-row highlighting/reasons. Python writes formulas but does not evaluate them. Input data is read from the first worksheet, while openpyxl modifies the active worksheet; keep this distinction in mind for multi-sheet inputs. The `.xls` path rebuilds a simpler `.xlsx` and does not preserve input styles or generate the same shop-summary/formula sheets. Both paths add `抱枕规格用量统计` and `枕芯规格用量统计`, using only included recognition details and the batch configuration matrix layout. Supported combinations without usage show zero; unsupported combinations stay blank. Usage values are static and per-file.
- Default output is `{stem}-已处理.xlsx`; collisions receive numeric suffixes unless overwrite is enabled. Directory input scans all Excel files nonrecursively, including previously processed outputs/configuration files, so use a directory containing only intended inputs.

`README.md` documents supported operation and workbook editing. `枕芯成本计算器备注规范说明.md` gives detailed remark examples, but fixed price/specification tables are snapshots; use the current template and executable matching logic as the authority.

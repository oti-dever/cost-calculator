@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ========================================
echo   成本计算器构建脚本
echo   Building CostCalculator.exe
echo ========================================
echo.

REM 选择 Python 解释器：优先当前目录 .venv，其次 python，再次 py
set "PY_CMD="
for %%I in ("%~dp0.venv\Scripts\python.exe") do (
    if exist "%%~fI" (
        set "PY_CMD=%%~fI"
    )
)

if not defined PY_CMD (
    python --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PY_CMD=python"
    )
)

if not defined PY_CMD (
    py --version >nul 2>&1
    if %errorlevel% equ 0 (
        set "PY_CMD=py"
    )
)

if not defined PY_CMD (
    echo [错误] 未检测到可用 Python（已尝试 .venv\Scripts\python.exe、python、py）
    pause
    exit /b 1
)

echo [信息] Python 版本:
"!PY_CMD!" --version

echo [信息] 使用解释器:
echo   !PY_CMD!

REM 检查 PyInstaller 是否安装
echo.
echo [检查] 检查 PyInstaller...
"!PY_CMD!" -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 未检测到 PyInstaller，正在安装...
    "!PY_CMD!" -m pip install pyinstaller
    if !errorlevel! neq 0 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
    echo [完成] PyInstaller 安装成功
) else (
    echo [完成] PyInstaller 已安装
)

REM 检查必需的 Python 包
echo.
echo [检查] 检查必需的依赖包...
"!PY_CMD!" -c "import openpyxl, pandas, textual, rich" >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 缺少必需的包，正在安装 openpyxl、pandas、textual、rich==13.9.4...
    "!PY_CMD!" -m pip install openpyxl pandas textual rich==13.9.4
    if !errorlevel! neq 0 (
        echo [错误] 依赖包安装失败
        pause
        exit /b 1
    )
    echo [完成] 依赖包安装成功
) else (
    echo [完成] 所有依赖包已安装
)

REM 固定 rich 版本，避免 PyInstaller + rich 14.x 在 onefile 下的 Unicode 子模块导入问题
"!PY_CMD!" -m pip install --disable-pip-version-check --quiet rich==13.9.4

REM 清理旧的构建文件
echo.
echo [清理] 删除旧的构建文件...
if exist "build" (
    rmdir /s /q "build"
    echo [完成] 已删除 build 目录
)
if exist "dist" (
    rmdir /s /q "dist"
    echo [完成] 已删除 dist 目录
)
if exist "CostCalculator.spec" (
    del /q "CostCalculator.spec"
    echo [完成] 已删除 CostCalculator.spec
)

echo.
echo ========================================
echo   开始构建可执行文件...
echo ========================================
echo.

REM 使用 PyInstaller 构建单文件可执行程序
"!PY_CMD!" -m PyInstaller --onefile ^
    --name CostCalculator ^
    --icon=NONE ^
    --clean ^
    --noconfirm ^
    --console ^
    --hidden-import textual ^
    --hidden-import rich ^
    --add-data "size_material_price.json;." ^
    --add-data "moving_and_selling_costs.json;." ^
    --add-data "pillow_cost.json;." ^
    --add-data "others.json;." ^
    textual_app.py

REM 检查构建是否成功
if %errorlevel% neq 0 (
    echo.
    echo [错误] PyInstaller 构建失败！请检查上面的错误信息。
    pause
    exit /b %errorlevel%
)

REM 复制测试数据目录到 dist，便于开箱验证
if exist "test_excel" (
    if exist "dist\test_excel" (
        rmdir /s /q "dist\test_excel"
    )
    xcopy /e /i /y "test_excel" "dist\test_excel" >nul
    echo [完成] 已复制 test_excel 到 dist\test_excel
) else (
    echo [提示] 未找到 test_excel 目录，跳过复制
)

echo.
echo ========================================
echo   构建完成！
echo ========================================
echo.
echo 输出目录: dist\
echo.
echo 可执行文件:
echo   - CostCalculator.exe
echo.
echo 配置文件:
echo   - 已内嵌到 CostCalculator.exe（无需额外 JSON 文件）
echo.
echo 功能说明:
echo   - 生成两个Sheet: 成本明细 + 店铺统计
echo   - 支持多种成本类型计算
echo   - 使用Excel公式实现动态更新
echo.
echo 使用方法:
echo   1. 将 Excel 文件拖放到 CostCalculator.exe 上
echo   2. Start exe then input file path in UI
echo   3. 处理完成后在结果表和日志区查看结果
echo.
echo 输入要求:
echo   - Excel文件必须包含 卖家备注 和 商家/店铺 列
echo   - file extensions: .xlsx / .xls
echo.
echo 输出结果:
echo   - 文件名: 原文件名-已处理.xlsx
echo   - Sheet1: 成本明细 (原始数据 + 成本列 + 合计行)
echo   - Sheet2: 店铺统计 (按店铺汇总 + 合计行)
echo.

pause

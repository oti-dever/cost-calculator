@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ========================================
echo   成本计算器构建脚本
echo   Building CostCalculator.exe
echo ========================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python
    pause
    exit /b 1
)

echo [信息] Python 版本:
python --version

REM 检查 PyInstaller 是否安装
echo.
echo [检查] 检查 PyInstaller...
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 未检测到 PyInstaller，正在安装...
    pip install pyinstaller
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
python -c "import openpyxl, pandas" >nul 2>&1
if %errorlevel% neq 0 (
    echo [警告] 缺少必需的包，正在安装 openpyxl 和 pandas...
    pip install openpyxl pandas
    if !errorlevel! neq 0 (
        echo [错误] 依赖包安装失败
        pause
        exit /b 1
    )
    echo [完成] 依赖包安装成功
) else (
    echo [完成] 所有依赖包已安装
)

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
pyinstaller --onefile ^
    --name CostCalculator ^
    --icon=NONE ^
    --clean ^
    --noconfirm ^
    --console ^
    cost_calculator.py

REM 检查构建是否成功
if %errorlevel% neq 0 (
    echo.
    echo [错误] PyInstaller 构建失败！请检查上面的错误信息。
    pause
    exit /b %errorlevel%
)

echo.
echo ========================================
echo   复制配置文件到输出目录...
echo ========================================
echo.

if exist "size_material_price.json" (
    xcopy /Y "size_material_price.json" "dist\" >nul
    echo [完成] size_material_price.json 已复制
) else (
    echo [错误] 未找到 size_material_price.json
)

if exist "moving_and_selling_costs.json" (
    xcopy /Y "moving_and_selling_costs.json" "dist\" >nul
    echo [完成] moving_and_selling_costs.json 已复制
) else (
    echo [提示] 未找到 moving_and_selling_costs.json
)

if exist "pillow_cost.json" (
    xcopy /Y "pillow_cost.json" "dist\" >nul
    echo [完成] pillow_cost.json 已复制
) else (
    echo [提示] 未找到 pillow_cost.json
)

if exist "others.json" (
    xcopy /Y "others.json" "dist\" >nul
    echo [完成] others.json 已复制
) else (
    echo [提示] 未找到 others.json
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
echo   - size_material_price.json (基础价格配置)
echo   - moving_and_selling_costs.json (义乳/义臀成本)
echo   - pillow_cost.json (枕芯成本)
echo   - others.json (硅胶/电动成本)
echo.
echo 功能说明:
echo   - 生成两个Sheet: 成本明细 + 店铺统计
echo   - 支持多种成本类型计算
echo   - 使用Excel公式实现动态更新
echo.
echo 使用方法:
echo   1. 将 Excel 文件拖放到 CostCalculator.exe 上
echo   2. 或在命令行运行: CostCalculator.exe "文件路径.xlsx"
echo   3. 或双击运行后手动输入文件路径
echo.
echo 输入要求:
echo   - Excel文件必须包含 "卖家备注" 和 "商家/店铺" 列
echo   - 支持 .xlsx 和 .xls 格式
echo.
echo 输出结果:
echo   - 文件名: 原文件名-已处理.xlsx
echo   - Sheet1: 成本明细 (原始数据 + 成本列 + 合计行)
echo   - Sheet2: 店铺统计 (按店铺汇总 + 合计行)
echo.

pause

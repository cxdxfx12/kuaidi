@echo off
REM ============================================
REM   大圣.快递物流派费结算系统V1.0 - 启动脚本
REM   杭州喵喵至家网络有限公司 · 大圣智慧软件
REM ============================================

REM 切换到脚本所在目录，避免双击后路径不对
cd /d "%~dp0"

echo.
echo ========================================
echo   大圣.快递物流派费结算系统V1.0
echo   正在启动，请稍候...
echo ========================================
echo.

REM ---------- 1. 检查 Python ----------
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo [错误] 未检测到 Python！
    echo 请先到 https://www.python.org 下载安装 Python 3.8 或更高版本
    echo.
    pause
    exit /b 1
)

REM 显示Python版本
for /f "tokens=2" %%I in ('python --version 2^>^&1') do set PY_VER=%%I
echo [信息] Python 版本: %PY_VER%

REM ---------- 2. 检查并安装依赖（首次运行） ----------
if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo [信息] 首次启动，正在初始化环境...
    REM 尝试用 python -m pip 安装依赖（比直接 pip 更稳）
    python -m pip install -r requirements.txt 2>nul
) else (
    REM 已有虚拟环境，检查关键依赖
    python -c "import PyQt5" 2>nul
    if errorlevel 1 (
        echo [信息] 依赖不完整，正在安装...
        python -m pip install -r requirements.txt
    )
)

REM ---------- 3. 启动程序（保留日志窗口） ----------
echo.
echo [信息] 正在启动程序...
echo.

REM 方式：新开窗口运行主程序，本窗口保留用于查看错误
start "大圣.快递物流派费结算系统V1.0" cmd /c "cd /d ""%~dp0"" & python main.py & echo. & echo 程序已退出，按任意键关闭 & pause"

exit /b 0

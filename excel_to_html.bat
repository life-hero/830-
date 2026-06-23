@echo off
chcp 65001 >nul
echo ========================================
echo 检测 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 未找到 Python，请先安装 Python 并添加到 PATH。
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b
)
echo.

echo 开始从现有 Excel 生成 HTML...
python main.py --from-excel
pause
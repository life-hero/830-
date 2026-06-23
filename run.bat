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
echo Python 已安装。

echo.
echo 正在安装所需依赖库（使用清华镜像）...
pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或手动执行 pip install -r requirements.txt
    pause
    exit /b
)

echo.
echo 依赖安装完成，开始执行统计...
python main.py
pause
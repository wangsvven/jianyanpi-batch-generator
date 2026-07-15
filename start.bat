@echo off
chcp 65001 >nul
title 检验批批量生成平台

echo ========================================
echo   检验批批量生成平台 v1.0
echo ========================================
echo.

cd /d "%~dp0"

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 创建虚拟环境（首次运行）
if not exist "venv\Scripts\activate.bat" (
    echo [1/3] 创建虚拟环境...
    python -m venv venv
)

:: 激活虚拟环境
call venv\Scripts\activate.bat

:: 安装依赖
echo [2/3] 检查并安装依赖...
pip install -r requirements.txt -q 2>nul

:: 启动服务
echo [3/3] 启动服务...
echo.
echo 请在浏览器中打开以下地址：
echo.
python -c "import socket; print(f'  本机访问: http://127.0.0.1:5000'); print(f'  局域网访问: http://{socket.gethostbyname(socket.gethostname())}:5000')"
echo.
python app.py

pause

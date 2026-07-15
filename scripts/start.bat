@echo off
chcp 65001 >nul 2>&1
title 检验批批量生成平台 v3.0

REM ============================================================
REM  Windows 启动脚本
REM  - 自动查找 Python（venv 优先 → 系统 Python）
REM  - 自动创建/激活 venv
REM  - 自动安装依赖
REM  - 显示局域网 IP
REM ============================================================

cd /d "%~dp0.."

echo.
echo ============================================================
echo   检验批批量生成平台 v3.0 - 启动中...
echo ============================================================
echo.

REM --- 1. 查找 Python ---
set "PYTHON="

REM 优先使用项目 venv
if exist "venv\Scripts\python.exe" (
    set "PYTHON=venv\Scripts\python.exe"
    echo [√] 使用项目虚拟环境 venv
) else (
    REM 查找系统 Python
    python --version >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=python"
        echo [√] 使用系统 Python

        REM 创建 venv
        echo [*] 正在创建虚拟环境...
        python -m venv venv
        if %errorlevel%==0 (
            set "PYTHON=venv\Scripts\python.exe"
            echo [√] 虚拟环境创建成功
        ) else (
            echo [!] 虚拟环境创建失败，使用系统 Python 继续
        )
    ) else (
        echo [×] 未找到 Python，请先安装 Python 3.8+
        echo     下载地址: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
)

REM --- 2. 检查并安装依赖 ---
echo.
echo [*] 检查依赖...
"%PYTHON%" -c "import flask, pandas, openpyxl, docxtpl, docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] 正在安装依赖包...
    "%PYTHON%" -m pip install -r requirements.txt -q
    if %errorlevel% neq 0 (
        echo [×] 依赖安装失败，请检查网络连接
        pause
        exit /b 1
    )
    echo [√] 依赖安装完成
) else (
    echo [√] 依赖已就绪
)

REM --- 3. 显示局域网 IP ---
echo.
echo ============================================================
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set "IP=%%a"
    setlocal enabledelayedexpansion
    set IP=!IP: =!
    echo   局域网访问: http://!IP!:5005
    endlocal
)
echo   本机访问:   http://127.0.0.1:5005
echo ============================================================
echo.
echo   浏览器将自动打开...
echo   按 Ctrl+C 停止服务
echo.

REM --- 4. 延迟打开浏览器 ---
start "" /b cmd /c "timeout /t 3 >nul && start http://127.0.0.1:5005"

REM --- 5. 启动服务 ---
cd src
"%PYTHON%" app.py
pause

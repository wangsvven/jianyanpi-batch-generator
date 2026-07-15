@echo off
chcp 65001 >nul 2>&1
title 检验批平台 - 一键安装（其他电脑）

REM ============================================================
REM  Windows 一键安装脚本（用于其他电脑首次部署）
REM  - 检测/安装 Python
REM  - 创建 venv
REM  - 安装依赖
REM  - 配置防火墙规则
REM  - 创建桌面快捷方式
REM ============================================================

cd /d "%~dp0.."

echo.
echo ============================================================
echo   检验批批量生成平台 - 一键安装
echo ============================================================
echo.

REM --- 1. 检测 Python ---
echo [1/5] 检测 Python 环境...

set "PYTHON="

REM 检查 venv 是否已存在
if exist "venv\Scripts\python.exe" (
    set "PYTHON=venv\Scripts\python.exe"
    echo       [√] 已有虚拟环境
    goto :INSTALL_DEPS
)

REM 检查系统 Python
python --version >nul 2>&1
if %errorlevel%==0 (
    echo       [√] 检测到系统 Python
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
    echo       版本: %PY_VER%
    set "PYTHON=python"
    goto :CREATE_VENV
)

REM 尝试 py launcher
py -3 --version >nul 2>&1
if %errorlevel%==0 (
    echo       [√] 检测到 Python (py launcher)
    set "PYTHON=py -3"
    goto :CREATE_VENV
)

echo       [×] 未检测到 Python
echo.
echo       请先安装 Python 3.8+：
echo       下载地址: https://www.python.org/downloads/
echo       安装时请勾选 "Add Python to PATH"
echo.
echo       安装完成后重新运行此脚本。
echo.
pause
exit /b 1

REM --- 2. 创建虚拟环境 ---
:CREATE_VENV
echo.
echo [2/5] 创建虚拟环境...
%PYTHON% -m venv venv
if %errorlevel%==0 (
    set "PYTHON=venv\Scripts\python.exe"
    echo       [√] 虚拟环境创建成功
) else (
    echo       [×] 虚拟环境创建失败
    echo       尝试使用系统 Python 继续...
    goto :INSTALL_DEPS
)

REM --- 3. 安装依赖 ---
:INSTALL_DEPS
echo.
echo [3/5] 安装依赖包...
echo       (首次安装可能需要几分钟，取决于网络速度)
echo.

"%PYTHON%" -m pip install --upgrade pip -q
"%PYTHON%" -m pip install -r requirements.txt -q

if %errorlevel%==0 (
    echo       [√] 依赖安装完成
) else (
    echo       [×] 依赖安装失败
    echo       请检查网络连接后重新运行此脚本
    pause
    exit /b 1
)

REM --- 4. 配置防火墙 ---
echo.
echo [4/5] 配置防火墙规则（允许局域网访问端口 5005）...

netsh advfirewall firewall delete rule name="检验批平台-5005" >nul 2>&1
netsh advfirewall firewall add rule name="检验批平台-5005" dir=in action=allow protocol=TCP localport=5005 >nul 2>&1

if %errorlevel%==0 (
    echo       [√] 防火墙规则已添加
) else (
    echo       [!] 防火墙配置失败（可能需要管理员权限）
    echo       局域网其他设备可能无法访问，请手动放行端口 5005
)

REM --- 5. 创建桌面快捷方式 ---
echo.
echo [5/5] 创建桌面快捷方式...

set "DESKTOP=%USERPROFILE%\Desktop"
set "SHORTCUT=%DESKTOP%\检验批平台.lnk"
set "TARGET=%~dp0start.bat"

REM 使用 PowerShell 创建快捷方式
powershell -NoProfile -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$sc = $ws.CreateShortcut('%SHORTCUT%'); " ^
    "$sc.TargetPath = '%TARGET%'; " ^
    "$sc.WorkingDirectory = '%~dp0'; " ^
    "$sc.IconLocation = '%PYTHON%,0'; " ^
    "$sc.Description = '检验批批量生成平台 v3.0'; " ^
    "$sc.Save()" >nul 2>&1

if %errorlevel%==0 (
    echo       [√] 桌面快捷方式已创建
) else (
    echo       [!] 快捷方式创建失败（不影响正常使用）
)

REM --- 完成 ---
echo.
echo ============================================================
echo   安装完成！
echo ============================================================
echo.
echo   启动方式：
echo     1. 双击桌面"检验批平台"快捷方式
echo     2. 或双击 scripts\start.bat
echo     3. 或双击项目根目录的"启动平台.bat"
echo.
echo   访问地址：http://127.0.0.1:5005
echo   局域网其他设备访问：http://本机IP:5005
echo.
pause

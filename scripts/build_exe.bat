@echo off
chcp 65001 >nul 2>&1
title PyInstaller 打包工具

REM ============================================================
REM  PyInstaller 打包脚本
REM  - 将项目打包为单个 EXE
REM  - 自动包含 web/ 静态资源
REM  - 输出到 dist/ 目录
REM ============================================================

cd /d "%~dp0.."

echo.
echo ============================================================
echo   PyInstaller 打包工具
echo ============================================================
echo.

REM --- 查找 Python ---
set "PYTHON="
if exist "venv\Scripts\python.exe" (
    set "PYTHON=venv\Scripts\python.exe"
) else (
    python --version >nul 2>&1
    if %errorlevel%==0 (
        set "PYTHON=python"
    ) else (
        echo [×] 未找到 Python，请先运行 start.bat 初始化环境
        pause
        exit /b 1
    )
)

REM --- 安装 PyInstaller ---
echo [*] 检查 PyInstaller...
"%PYTHON%" -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] 安装 PyInstaller...
    "%PYTHON%" -m pip install pyinstaller -q
)

REM --- 清理旧文件 ---
echo [*] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist\jianyanpi.exe del /q dist\jianyanpi.exe
if exist jianyanpi.spec del /q jianyanpi.spec

REM --- 执行打包 ---
echo.
echo [*] 开始打包（可能需要几分钟）...
echo.

"%PYTHON%" -m PyInstaller ^
    --name jianyanpi ^
    --noconfirm ^
    --clean ^
    --onedir ^
    --windowed ^
    --add-data "web;web" ^
    --add-data "src;src" ^
    --hidden-import flask ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import docxtpl ^
    --hidden-import docxcompose ^
    --hidden-import docx ^
    --collect-data docxtpl ^
    --workpath build ^
    --distpath dist ^
    src/app.py

if %errorlevel%==0 (
    echo.
    echo [√] 打包成功！
    echo     输出目录: dist\jianyanpi\
    echo     可执行文件: dist\jianyanpi\jianyanpi.exe
    echo.
    echo     将 dist\jianyanpi\ 整个文件夹复制到目标电脑即可使用。
) else (
    echo.
    echo [×] 打包失败，请检查错误信息
)

pause

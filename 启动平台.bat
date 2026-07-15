@echo off
chcp 65001 >nul 2>&1
title 检验批批量生成平台 v3.0

REM ============================================================
REM  根目录启动入口
REM  自动调用 scripts/start.bat
REM ============================================================

cd /d "%~dp0"
call scripts\start.bat

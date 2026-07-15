#!/bin/bash
# ============================================================
#  macOS / Linux 启动脚本
#  - 自动查找 Python3（venv 优先 → 系统 python3）
#  - 自动创建/激活 venv
#  - 自动安装依赖
#  - 显示局域网 IP
# ============================================================

set -e

# 切换到项目根目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR/.."

PROJECT_DIR="$(pwd)"
PORT=5005

echo ""
echo "============================================================"
echo "  检验批批量生成平台 v3.0 - 启动中..."
echo "============================================================"
echo ""

# --- 1. 查找 Python ---
PYTHON=""

if [ -f "$PROJECT_DIR/venv/bin/python" ]; then
    PYTHON="$PROJECT_DIR/venv/bin/python"
    echo "[√] 使用项目虚拟环境 venv"
else
    # 查找系统 python3
    if command -v python3 &>/dev/null; then
        SYS_PYTHON=$(command -v python3)
        echo "[√] 使用系统 Python3: $SYS_PYTHON"

        # 创建 venv
        echo "[*] 正在创建虚拟环境..."
        "$SYS_PYTHON" -m venv "$PROJECT_DIR/venv"
        if [ $? -eq 0 ]; then
            PYTHON="$PROJECT_DIR/venv/bin/python"
            echo "[√] 虚拟环境创建成功"
        else
            echo "[!] 虚拟环境创建失败，使用系统 Python 继续"
            PYTHON="$SYS_PYTHON"
        fi
    else
        echo "[×] 未找到 Python3，请先安装 Python 3.8+"
        echo "    macOS:  brew install python3"
        echo "    Ubuntu: sudo apt install python3 python3-venv"
        echo ""
        exit 1
    fi
fi

# --- 2. 检查并安装依赖 ---
echo ""
echo "[*] 检查依赖..."
"$PYTHON" -c "import flask, pandas, openpyxl, docxtpl, docx" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[*] 正在安装依赖包..."
    "$PYTHON" -m pip install -r requirements.txt -q
    if [ $? -ne 0 ]; then
        echo "[×] 依赖安装失败，请检查网络连接"
        exit 1
    fi
    echo "[√] 依赖安装完成"
else
    echo "[√] 依赖已就绪"
fi

# --- 3. 显示局域网 IP ---
echo ""
echo "============================================================"

# macOS 和 Linux 获取局域网 IP
LOCAL_IP=""
if [ "$(uname)" = "Darwin" ]; then
    LOCAL_IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "127.0.0.1")
else
    LOCAL_IP=$(hostname -I 2>/dev/null | awk '{print $1}' || echo "127.0.0.1")
fi

echo "  局域网访问: http://$LOCAL_IP:$PORT"
echo "  本机访问:   http://127.0.0.1:$PORT"
echo "============================================================"
echo ""
echo "  浏览器将自动打开..."
echo "  按 Ctrl+C 停止服务"
echo ""

# --- 4. 延迟打开浏览器 ---
(sleep 3 && (
    if [ "$(uname)" = "Darwin" ]; then
        open "http://127.0.0.1:$PORT"
    else
        xdg-open "http://127.0.0.1:$PORT" 2>/dev/null
    fi
)) &

# --- 5. 启动服务 ---
cd "$PROJECT_DIR/src"
"$PYTHON" app.py

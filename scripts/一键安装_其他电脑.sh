#!/bin/bash
# ============================================================
#  macOS / Linux 一键安装脚本（用于其他电脑首次部署）
#  - 检测/安装 Python3
#  - 创建 venv
#  - 安装依赖
#  - 创建桌面快捷方式（macOS）
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR/.."
PROJECT_DIR="$(pwd)"
PORT=5005

echo ""
echo "============================================================"
echo "  检验批批量生成平台 - 一键安装"
echo "============================================================"
echo ""

# --- 1. 检测 Python ---
echo "[1/4] 检测 Python3 环境..."

PYTHON=""

if [ -f "$PROJECT_DIR/venv/bin/python" ]; then
    PYTHON="$PROJECT_DIR/venv/bin/python"
    echo "      [√] 已有虚拟环境"
else
    if command -v python3 &>/dev/null; then
        SYS_PYTHON=$(command -v python3)
        PY_VERSION=$("$SYS_PYTHON" --version 2>&1)
        echo "      [√] 检测到 $PY_VERSION"
        PYTHON="$SYS_PYTHON"
    else
        echo "      [×] 未检测到 Python3"
        echo ""
        echo "      请先安装 Python 3.8+："
        if [ "$(uname)" = "Darwin" ]; then
            echo "      macOS:  brew install python3"
            echo "      或访问: https://www.python.org/downloads/"
        else
            echo "      Ubuntu/Debian: sudo apt install python3 python3-venv python3-pip"
            echo "      CentOS/RHEL:   sudo yum install python3"
        fi
        echo ""
        echo "      安装完成后重新运行此脚本。"
        exit 1
    fi
fi

# --- 2. 创建虚拟环境 ---
if [ ! -f "$PROJECT_DIR/venv/bin/python" ]; then
    echo ""
    echo "[2/4] 创建虚拟环境..."
    "$PYTHON" -m venv "$PROJECT_DIR/venv"
    if [ $? -eq 0 ]; then
        PYTHON="$PROJECT_DIR/venv/bin/python"
        echo "      [√] 虚拟环境创建成功"
    else
        echo "      [×] 虚拟环境创建失败"
        echo "      尝试使用系统 Python 继续..."
    fi
else
    echo ""
    echo "[2/4] 虚拟环境已存在，跳过"
fi

# --- 3. 安装依赖 ---
echo ""
echo "[3/4] 安装依赖包..."
echo "      (首次安装可能需要几分钟，取决于网络速度)"
echo ""

"$PYTHON" -m pip install --upgrade pip -q
"$PYTHON" -m pip install -r requirements.txt -q

if [ $? -eq 0 ]; then
    echo "      [√] 依赖安装完成"
else
    echo "      [×] 依赖安装失败"
    echo "      请检查网络连接后重新运行此脚本"
    exit 1
fi

# --- 4. 创建快捷方式 ---
echo ""
echo "[4/4] 创建快捷方式..."

if [ "$(uname)" = "Darwin" ]; then
    # macOS: 创建 .command 文件到桌面
    DESKTOP="$HOME/Desktop"
    SHORTCUT="$DESKTOP/检验批平台.command"

    cat > "$SHORTCUT" << 'EOF'
#!/bin/bash
cd "$(dirname "$0")"
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
# 尝试找到项目目录
if [ -f "$PROJECT_DIR/../../scripts/start.sh" ]; then
    "$PROJECT_DIR/../../scripts/start.sh"
elif [ -f "$PROJECT_DIR/scripts/start.sh" ]; then
    "$PROJECT_DIR/scripts/start.sh"
else
    echo "无法找到项目目录，请手动运行 scripts/start.sh"
fi
EOF
    chmod +x "$SHORTCUT"
    echo "      [√] 桌面快捷方式已创建: $SHORTCUT"
else
    # Linux: 创建 .desktop 文件
    DESKTOP_DIR="$HOME/Desktop"
    [ -d "$HOME/.local/share/applications" ] && DESKTOP_DIR="$HOME/.local/share/applications"

    cat > "$DESKTOP_DIR/检验批平台.desktop" << EOF
[Desktop Entry]
Version=3.0
Name=检验批平台
Comment=检验批批量生成平台
Exec=$PROJECT_DIR/scripts/start.sh
Terminal=true
Type=Application
Categories=Utility;
EOF
    chmod +x "$DESKTOP_DIR/检验批平台.desktop"
    echo "      [√] 快捷方式已创建: $DESKTOP_DIR/检验批平台.desktop"
fi

# --- 完成 ---
echo ""
echo "============================================================"
echo "  安装完成！"
echo "============================================================"
echo ""
echo "  启动方式："
echo "    1. 双击桌面快捷方式"
echo "    2. 或运行 scripts/start.sh"
echo ""
echo "  访问地址：http://127.0.0.1:$PORT"
echo "  局域网其他设备访问：http://本机IP:$PORT"
echo ""

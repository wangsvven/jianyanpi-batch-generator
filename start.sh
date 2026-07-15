#!/bin/bash
# 检验批批量生成平台 启动脚本 (macOS / Linux)

echo "========================================"
echo "  检验批批量生成平台 v1.0"
echo "========================================"
echo ""

cd "$(dirname "$0")"

# 检查 Python3
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未找到 Python3，请先安装 Python 3.9+"
    echo "下载地址：https://www.python.org/downloads/"
    exit 1
fi

# 创建虚拟环境（首次运行）
if [ ! -d "venv" ]; then
    echo "[1/3] 创建虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "[2/3] 检查并安装依赖..."
pip install -r requirements.txt -q 2>/dev/null

# 启动服务
echo "[3/3] 启动服务..."
echo ""
echo "请在浏览器中打开以下地址："
python3 -c "import socket; print(f'  本机访问: http://127.0.0.1:5000'); ip=socket.gethostbyname(socket.gethostname()); print(f'  局域网访问: http://{ip}:5000')"
echo ""
python3 app.py

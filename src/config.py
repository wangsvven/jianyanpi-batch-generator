"""
================================================================================
 全局配置模块
================================================================================
 集中管理端口、路径、常量，供 app.py 和 engine.py 共享。
 支持开发模式和 PyInstaller 冻结模式。
================================================================================
"""
import sys
from pathlib import Path

# ============================================================
#  端口配置
# ============================================================
PORT = 5005
HOST = '0.0.0.0'

# ============================================================
#  路径配置（自动适配开发 / 冻结模式）
# ============================================================
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后：资源在 _MEIPASS，运行时数据在 exe 同目录
    RESOURCE_DIR = Path(sys._MEIPASS)
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # 开发模式：config.py 在 src/ 下，回退到项目根目录
    BASE_DIR = Path(__file__).resolve().parent.parent
    RESOURCE_DIR = BASE_DIR / 'web'

UPLOAD_DIR = BASE_DIR / 'data' / 'uploads'
OUTPUT_DIR = BASE_DIR / 'data' / 'outputs'
PRESET_DIR = BASE_DIR / 'data' / 'presets'
TEMPLATE_DIR = RESOURCE_DIR / 'templates'
STATIC_DIR = RESOURCE_DIR / 'static'

# 确保运行时目录存在
for _d in [UPLOAD_DIR, OUTPUT_DIR, PRESET_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

# ============================================================
#  文件类型
# ============================================================
ALLOWED_EXTENSIONS = {
    'docx': ['.docx'],
    'xlsx': ['.xlsx', '.xls'],
}

MAX_UPLOAD_SIZE = 500 * 1024 * 1024  # 500MB

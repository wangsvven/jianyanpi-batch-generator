"""
================================================================================
 检验批批量生成平台 — Flask Web 后端
================================================================================

 用途
   将「Word 模板 + Excel 数据 → 批量生成文档」的工作流搬上浏览器，
   支持本机和局域网内多设备同时访问。

 架构
   app.py   — Flask 路由层（文件上传、任务调度、下载、预设管理）
   engine.py — 核心生成引擎（模板填充、多任务编排、文档合并）
   前端      — 原生 HTML/CSS/JS，位于 templates/ 和 static/

 核心 API 路由
   /api/session/init          — 初始化会话，获取 session_id
   /api/upload                — 上传 Word 模板 / Excel 文件
   /api/files                 — 列出已上传文件
   /api/files/delete          — 删除文件
   /api/excel/sheets          — 获取 Excel 工作表列表
   /api/excel/preview         — 预览 Excel 数据
   /api/template/placeholders — 提取 Word 模板中的 {{占位符}}
   /api/match-fields          — 对比模板占位符与 Excel 列名
   /api/generate              — 执行批量生成（后台线程）
   /api/generate/progress/id  — 查询生成进度
   /api/download              — 下载单个文件
   /api/download/zip          — 打包下载（全部或子目录）
   /api/outputs               — 列出输出目录结构
   /api/presets               — 预设方案 CRUD

 启动方式
   Windows:  双击 start.bat
   macOS/Linux: bash start.sh
   手动:     python app.py

 依赖
   pip install -r requirements.txt
================================================================================
"""
import os
import io
import json
import uuid
import shutil
import zipfile
import threading
import pandas as pd
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory

from engine import (
    generate_single_task,
    generate_tasks,
    merge_documents,
    get_excel_sheets,
    get_excel_preview,
    get_excel_engine,
    extract_template_placeholders,
)

# ============================================================
#  应用配置
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / 'uploads'
OUTPUT_DIR = BASE_DIR / 'outputs'
PRESET_DIR = BASE_DIR / 'presets'
TEMPLATE_DIR = BASE_DIR / 'templates'
STATIC_DIR = BASE_DIR / 'static'

# 确保目录存在
for d in [UPLOAD_DIR, OUTPUT_DIR, PRESET_DIR]:
    d.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {
    'docx': ['.docx'],
    'xlsx': ['.xlsx', '.xls'],
}

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB 上传限制
app.config['SECRET_KEY'] = str(uuid.uuid4())

# 跟踪运行中的任务
running_tasks = {}
task_lock = threading.Lock()


def allowed_file(filename, file_type):
    """检查文件扩展名是否允许"""
    if '.' not in filename:
        return False
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_EXTENSIONS.get(file_type, [])


def safe_filename(filename):
    """
    清理文件名，保留中文等非 ASCII 字符，
    仅移除 Windows 文件系统不允许的字符
    """
    # Windows 非法字符
    illegal = '<>:"/\\|?*\n\r\t\0'
    for ch in illegal:
        filename = filename.replace(ch, '_')
    # 移除控制字符，保留可打印字符（含中文）
    filename = ''.join(c for c in filename if c.isprintable() or c == ' ')
    # 去掉首尾空格和点号（Windows 不允许目录/文件名以点号结尾）
    return filename.strip().rstrip('.') or 'unnamed'


def get_unique_session_dir(session_id):
    """获取会话专属的目录"""
    upload_dir = UPLOAD_DIR / session_id
    output_dir = OUTPUT_DIR / session_id
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return upload_dir, output_dir


# ============================================================
#  页面路由
# ============================================================

@app.route('/')
def index():
    return render_template('index.html')


# ============================================================
#  文件上传
# ============================================================

@app.route('/api/session/init', methods=['POST'])
def init_session():
    """初始化会话，返回 session_id"""
    session_id = uuid.uuid4().hex[:12]
    get_unique_session_dir(session_id)
    return jsonify({'session_id': session_id, 'status': 'ok'})


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传文件（模板或 Excel）"""
    session_id = request.form.get('session_id', 'default')
    file_type = request.form.get('type', 'docx')  # 'docx' 或 'xlsx'

    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    if not allowed_file(file.filename, file_type):
        return jsonify({'error': f'不支持的文件类型，请上传 {" 或 ".join(ALLOWED_EXTENSIONS[file_type])}'}), 400

    upload_dir, _ = get_unique_session_dir(session_id)
    original_name = file.filename
    safe_name = safe_filename(original_name)
    # 添加时间戳避免重名
    timestamp = datetime.now().strftime('%H%M%S')
    save_name = f"{timestamp}_{safe_name}"
    save_path = upload_dir / save_name

    file.save(str(save_path))

    return jsonify({
        'status': 'ok',
        'filename': original_name,
        'saved_as': save_name,
        'path': str(save_path),
        'size': save_path.stat().st_size
    })


@app.route('/api/files', methods=['GET'])
def list_files():
    """列出会话中已上传的文件"""
    session_id = request.args.get('session_id', 'default')
    upload_dir, _ = get_unique_session_dir(session_id)

    files = []
    for f in upload_dir.iterdir():
        if f.is_file():
            ext = f.suffix.lower()
            ftype = 'word' if ext == '.docx' else 'excel' if ext in ('.xlsx', '.xls') else 'other'
            # 还原原始文件名：去掉 "HHMMSS_" 时间戳前缀
            name = f.name
            if '_' in name:
                # 时间戳格式为 6 位数字 + 下划线
                parts = name.split('_', 1)
                if parts[0].isdigit() and len(parts[0]) == 6:
                    original = parts[1]
                else:
                    original = name
            else:
                original = name
            files.append({
                'name': name,
                'original_name': original,
                'type': ftype,
                'size': f.stat().st_size,
                'path': str(f),
                'modified': datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })

    # 按修改时间倒序
    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify({'files': files})


@app.route('/api/files/delete', methods=['POST'])
def delete_file():
    """删除文件"""
    data = request.get_json()
    session_id = data.get('session_id', 'default')
    filename = data.get('filename', '')

    upload_dir, _ = get_unique_session_dir(session_id)
    file_path = upload_dir / filename

    if file_path.exists():
        file_path.unlink()
        return jsonify({'status': 'ok'})
    return jsonify({'error': '文件不存在'}), 404


# ============================================================
#  Excel 操作
# ============================================================

@app.route('/api/excel/sheets', methods=['POST'])
def excel_sheets():
    """获取 Excel 工作表列表"""
    data = request.get_json()
    excel_path = data.get('path', '')

    if not excel_path or not Path(excel_path).exists():
        return jsonify({'error': 'Excel 文件不存在'}), 404

    try:
        sheets = get_excel_sheets(excel_path)
        return jsonify({'sheets': sheets})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/excel/preview', methods=['POST'])
def excel_preview():
    """预览 Excel 数据"""
    data = request.get_json()
    excel_path = data.get('path', '')
    sheet_name = data.get('sheet_name', 'Sheet1')
    rows = data.get('rows', 5)

    if not excel_path or not Path(excel_path).exists():
        return jsonify({'error': 'Excel 文件不存在'}), 404

    try:
        result = get_excel_preview(excel_path, sheet_name, rows)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ============================================================
#  Word 模板占位符提取
# ============================================================

@app.route('/api/template/placeholders', methods=['POST'])
def template_placeholders():
    """提取 Word 模板中的 {{占位符}}"""
    data = request.get_json()
    template_path = data.get('path', '')

    if not template_path or not Path(template_path).exists():
        return jsonify({'error': '模板文件不存在'}), 404

    try:
        placeholders = extract_template_placeholders(template_path)
        return jsonify({'placeholders': placeholders})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/match-fields', methods=['POST'])
def match_fields():
    """
    对比模板占位符与 Excel 列名，返回匹配结果
    帮助用户确认数据列是否齐全
    """
    data = request.get_json()
    template_path = data.get('template_path', '')
    excel_path = data.get('excel_path', '')
    sheet_name = data.get('sheet_name', 'Sheet1')

    result = {
        'placeholders': [],
        'excel_columns': [],
        'matched': [],
        'missing_in_excel': [],
        'extra_in_excel': []
    }

    # 获取模板占位符
    if template_path and Path(template_path).exists():
        try:
            result['placeholders'] = extract_template_placeholders(template_path)
        except Exception:
            pass

    # 获取 Excel 列名
    if excel_path and Path(excel_path).exists():
        try:
            engine = get_excel_engine(excel_path)
            df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=0, engine=engine)
            result['excel_columns'] = list(df.columns)
        except Exception:
            pass

    # 对比
    ph_set = set(result['placeholders'])
    col_set = set(str(c) for c in result['excel_columns'])

    result['matched'] = sorted(ph_set & col_set)
    result['missing_in_excel'] = sorted(ph_set - col_set)
    result['extra_in_excel'] = sorted(col_set - ph_set)

    return jsonify(result)

@app.route('/api/generate', methods=['POST'])
def run_generate():
    """执行批量生成任务"""
    data = request.get_json()
    session_id = data.get('session_id', 'default')
    tasks_config = data.get('tasks', [])
    merge_config = data.get('merge_tasks', [])

    if not tasks_config:
        return jsonify({'error': '没有配置任何任务'}), 400

    task_id = uuid.uuid4().hex[:8]
    _, output_dir = get_unique_session_dir(session_id)

    # 初始化任务状态
    with task_lock:
        running_tasks[task_id] = {
            'status': 'running',
            'progress': {'current': 0, 'total': 0, 'filename': '', 'status': ''},
            'result': None,
            'started_at': datetime.now().isoformat()
        }

    def progress_callback(current, total, filename, status):
        with task_lock:
            running_tasks[task_id]['progress'] = {
                'current': current,
                'total': total,
                'filename': filename,
                'status': status
            }

    def run():
        try:
            result = generate_tasks(
                tasks_config=tasks_config,
                base_output_dir=str(output_dir),
                progress_callback=progress_callback
            )

            # 执行合并任务
            merge_results = []
            if merge_config:
                for mc in merge_config:
                    try:
                        input_dir = output_dir / mc.get('input_subdir', '')
                        output_file = output_dir / mc.get('output_file', 'merged.docx')
                        sort_mode = mc.get('sort_mode', 3)
                        date_kw = mc.get('date_sort_keyword', '')
                        count = merge_documents(str(input_dir), str(output_file), sort_mode, date_kw)
                        merge_results.append({
                            'input_subdir': mc.get('input_subdir', ''),
                            'output_file': str(output_file),
                            'count': count,
                            'status': 'success'
                        })
                    except Exception as e:
                        merge_results.append({
                            'input_subdir': mc.get('input_subdir', ''),
                            'error': str(e),
                            'status': 'error'
                        })

            result['merge_results'] = merge_results

            with task_lock:
                running_tasks[task_id]['status'] = 'completed'
                running_tasks[task_id]['result'] = result
        except Exception as e:
            with task_lock:
                running_tasks[task_id]['status'] = 'error'
                running_tasks[task_id]['result'] = {'error': str(e)}

    threading.Thread(target=run, daemon=True).start()

    return jsonify({'task_id': task_id, 'status': 'started'})


@app.route('/api/generate/progress/<task_id>', methods=['GET'])
def get_progress(task_id):
    """查询任务进度"""
    with task_lock:
        task = running_tasks.get(task_id)
    if not task:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(task)


# ============================================================
#  文件下载
# ============================================================

@app.route('/api/download', methods=['GET'])
def download_file():
    """下载生成的文件"""
    file_path = request.args.get('path', '')
    if not file_path:
        return jsonify({'error': '未指定文件路径'}), 400

    path = Path(file_path)
    if not path.exists():
        return jsonify({'error': '文件不存在'}), 404

    return send_file(str(path), as_attachment=True, download_name=path.name)


@app.route('/api/download/zip', methods=['GET'])
def download_zip():
    """打包下载输出文件（全部或指定子目录）"""
    session_id = request.args.get('session_id', 'default')
    subdir = request.args.get('subdir', '')

    _, output_dir = get_unique_session_dir(session_id)

    if subdir:
        target_dir = output_dir / subdir
    else:
        target_dir = output_dir

    if not target_dir.exists():
        return jsonify({'error': '目录不存在'}), 404

    # 创建内存中的 ZIP
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(str(target_dir)):
            for fname in files:
                if fname.startswith('~$') or not fname.endswith('.docx'):
                    continue
                fpath = Path(root) / fname
                arcname = fpath.relative_to(target_dir)
                zf.write(str(fpath), str(arcname))

    memory_file.seek(0)

    if subdir:
        clean_sub = subdir.replace('/', '_').replace('\\', '_')
        zip_name = f"{clean_sub}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    else:
        zip_name = f"all_outputs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

    return send_file(memory_file, mimetype='application/zip',
                     as_attachment=True, download_name=zip_name)


@app.route('/api/outputs', methods=['GET'])
def list_outputs():
    """列出会话输出目录中的文件"""
    session_id = request.args.get('session_id', 'default')
    _, output_dir = get_unique_session_dir(session_id)

    def scan_dir(directory, base_path):
        items = []
        if not directory.exists():
            return items
        for entry in sorted(directory.iterdir()):
            rel = str(entry.relative_to(base_path))
            if entry.is_dir():
                children = scan_dir(entry, base_path)
                if children:
                    items.append({
                        'name': entry.name,
                        'type': 'folder',
                        'path': str(entry),
                        'rel_path': rel,
                        'children': children
                    })
            elif entry.is_file() and entry.suffix.lower() == '.docx' and not entry.name.startswith('~$'):
                items.append({
                    'name': entry.name,
                    'type': 'file',
                    'path': str(entry),
                    'rel_path': rel,
                    'size': entry.stat().st_size
                })
        return items

    files = scan_dir(output_dir, output_dir)
    return jsonify({'files': files, 'base_dir': str(output_dir)})


# ============================================================
#  预设管理
# ============================================================

@app.route('/api/presets', methods=['GET'])
def list_presets():
    """列出所有预设"""
    presets = []
    for f in sorted(PRESET_DIR.glob('*.json')):
        try:
            with open(f, 'r', encoding='utf-8') as fp:
                data = json.load(fp)
            presets.append({
                'id': f.stem,
                'name': data.get('name', f.stem),
                'created': datetime.fromtimestamp(f.stat().st_ctime).isoformat(),
                'task_count': len(data.get('tasks', []))
            })
        except Exception:
            pass
    return jsonify({'presets': presets})


@app.route('/api/presets/save', methods=['POST'])
def save_preset():
    """保存预设"""
    data = request.get_json()
    name = data.get('name', '')
    tasks = data.get('tasks', [])
    merge_tasks = data.get('merge_tasks', [])

    if not name:
        return jsonify({'error': '请输入预设名称'}), 400

    safe_name = safe_filename(name) or 'preset'
    preset_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe_name}"

    preset_data = {
        'name': name,
        'tasks': tasks,
        'merge_tasks': merge_tasks,
        'created_at': datetime.now().isoformat()
    }

    preset_path = PRESET_DIR / f"{preset_id}.json"
    with open(preset_path, 'w', encoding='utf-8') as f:
        json.dump(preset_data, f, ensure_ascii=False, indent=2)

    return jsonify({'status': 'ok', 'id': preset_id, 'name': name})


@app.route('/api/presets/load/<preset_id>', methods=['GET'])
def load_preset(preset_id):
    """加载预设"""
    preset_path = PRESET_DIR / f"{preset_id}.json"
    if not preset_path.exists():
        return jsonify({'error': '预设不存在'}), 404

    with open(preset_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return jsonify(data)


@app.route('/api/presets/delete/<preset_id>', methods=['DELETE'])
def delete_preset(preset_id):
    """删除预设"""
    preset_path = PRESET_DIR / f"{preset_id}.json"
    if preset_path.exists():
        preset_path.unlink()
        return jsonify({'status': 'ok'})
    return jsonify({'error': '预设不存在'}), 404


# ============================================================
#  启动
# ============================================================

if __name__ == '__main__':
    import socket

    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)

    print("=" * 60)
    print("  检验批批量生成平台 v1.0")
    print("=" * 60)
    print(f"  本机访问: http://127.0.0.1:5000")
    print(f"  局域网访问: http://{local_ip}:5000")
    print("=" * 60)

    app.run(host='0.0.0.0', port=5000, debug=False)

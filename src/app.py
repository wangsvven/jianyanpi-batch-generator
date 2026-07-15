"""
================================================================================
 检验批批量生成平台 — Flask Web 后端
================================================================================
 将「Word 模板 + Excel 数据 -> 批量生成文档」搬上浏览器。
 支持本机和局域网多设备同时访问。

 API 路由:
   /api/session/init          — 初始化会话
   /api/upload                — 上传文件
   /api/files                 — 列出已上传文件
   /api/files/delete          — 删除文件
   /api/excel/sheets          — 获取 Excel 工作表
   /api/excel/preview         — 预览 Excel 数据
   /api/template/placeholders — 提取 Word 模板占位符
   /api/template/export-excel — 导出 Excel 模板
   /api/match-fields          — 字段匹配检查
   /api/generate              — 执行批量生成
   /api/generate/progress/<id>— 查询生成进度
   /api/merge/preview         — 合并排序预览
   /api/download              — 下载单个文件
   /api/download/zip          — 打包下载
   /api/outputs               — 列出输出目录
   /api/presets               — 预设管理 CRUD
   /api/cleanup/session       — 清空当前会话
   /api/cleanup/all           — 清空全部
   /api/open-folder           — 打开文件夹
================================================================================
"""
import os
import io
import json
import uuid
import shutil
import zipfile
import threading
import traceback
from datetime import datetime
from pathlib import Path

# ---- 依赖导入（冻结模式下捕获错误写入日志） ----
_frozen = getattr(os, 'frozen', False) or getattr(__import__('sys'), 'frozen', False)
if _frozen:
    import sys as _sys
    try:
        _exe_dir = Path(_sys.executable).resolve().parent
    except Exception:
        _exe_dir = Path.cwd()
    _error_log_path = _exe_dir / '启动错误.log'
    try:
        _error_log_path.unlink()
    except Exception:
        pass

try:
    import pandas as pd
    from flask import Flask, render_template, request, jsonify, send_file
    from config import (
        PORT, HOST, BASE_DIR, RESOURCE_DIR,
        UPLOAD_DIR, OUTPUT_DIR, PRESET_DIR, TEMPLATE_DIR, STATIC_DIR,
        ALLOWED_EXTENSIONS, MAX_UPLOAD_SIZE
    )
    from engine import (
        generate_single_task,
        generate_tasks,
        merge_documents,
        preview_merge_order,
        get_excel_sheets,
        get_excel_preview,
        get_excel_engine,
        extract_template_placeholders,
        generate_excel_template,
    )
except Exception as _import_err:
    import sys as _sys
    _msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 导入失败:\n{traceback.format_exc()}\n"
    print(_msg)
    if _frozen:
        try:
            with open(_error_log_path, 'w', encoding='utf-8') as _f:
                _f.write(_msg)
            print(f"错误日志: {_error_log_path}")
        except Exception:
            pass
        print("按任意键退出...")
        try:
            input()
        except Exception:
            pass
    raise


# ============================================================
#  Flask 应用初始化
# ============================================================

app = Flask(
    __name__,
    template_folder=str(TEMPLATE_DIR),
    static_folder=str(STATIC_DIR)
)
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_SIZE
app.config['SECRET_KEY'] = str(uuid.uuid4())

# 任务跟踪
running_tasks = {}
task_lock = threading.Lock()


# ============================================================
#  辅助函数
# ============================================================

def allowed_file(filename, file_type):
    """检查文件扩展名"""
    if '.' not in filename:
        return False
    ext = Path(filename).suffix.lower()
    return ext in ALLOWED_EXTENSIONS.get(file_type, [])


def safe_filename(filename):
    """清理文件名，保留中文，仅移除系统非法字符"""
    illegal = '<>:"/\\|?*\n\r\t\0'
    for ch in illegal:
        filename = filename.replace(ch, '_')
    filename = ''.join(c for c in filename if c.isprintable() or c == ' ')
    return filename.strip().rstrip('.') or 'unnamed'


def get_session_dirs(session_id):
    """获取会话专属的上传/输出目录"""
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
#  会话 & 文件上传
# ============================================================

@app.route('/api/session/init', methods=['POST'])
def init_session():
    session_id = uuid.uuid4().hex[:12]
    get_session_dirs(session_id)
    return jsonify({'session_id': session_id, 'status': 'ok'})


@app.route('/api/upload', methods=['POST'])
def upload_file():
    session_id = request.form.get('session_id', 'default')
    file_type = request.form.get('type', 'docx')

    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    if not allowed_file(file.filename, file_type):
        exts = ' 或 '.join(ALLOWED_EXTENSIONS.get(file_type, []))
        return jsonify({'error': f'不支持的文件类型，请上传 {exts}'}), 400

    upload_dir, _ = get_session_dirs(session_id)
    safe_name = safe_filename(file.filename)
    timestamp = datetime.now().strftime('%H%M%S')
    save_name = f"{timestamp}_{safe_name}"
    save_path = upload_dir / save_name
    file.save(str(save_path))

    return jsonify({
        'status': 'ok',
        'filename': file.filename,
        'saved_as': save_name,
        'path': str(save_path),
        'size': save_path.stat().st_size
    })


@app.route('/api/files', methods=['GET'])
def list_files():
    session_id = request.args.get('session_id', 'default')
    upload_dir, _ = get_session_dirs(session_id)

    files = []
    for f in upload_dir.iterdir():
        if not f.is_file():
            continue
        ext = f.suffix.lower()
        ftype = 'word' if ext == '.docx' else 'excel' if ext in ('.xlsx', '.xls') else 'other'
        name = f.name
        # 还原原始文件名（去掉 HHMMSS_ 时间戳前缀）
        if '_' in name:
            parts = name.split('_', 1)
            original = parts[1] if parts[0].isdigit() and len(parts[0]) == 6 else name
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

    files.sort(key=lambda x: x['modified'], reverse=True)
    return jsonify({'files': files})


@app.route('/api/files/delete', methods=['POST'])
def delete_file():
    data = request.get_json()
    session_id = data.get('session_id', 'default')
    filename = data.get('filename', '')

    upload_dir, _ = get_session_dirs(session_id)
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
#  Word 模板操作
# ============================================================

@app.route('/api/template/placeholders', methods=['POST'])
def template_placeholders():
    data = request.get_json()
    template_path = data.get('path', '')

    if not template_path or not Path(template_path).exists():
        return jsonify({'error': '模板文件不存在'}), 404

    try:
        placeholders = extract_template_placeholders(template_path)
        return jsonify({'placeholders': placeholders})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/template/export-excel', methods=['POST'])
def export_excel_template():
    """根据 Word 模板占位符生成 Excel 数据模板"""
    import tempfile

    data = request.get_json() or {}
    template_path = data.get('template_path', '')
    placeholders = data.get('placeholders', None)
    aliases = data.get('aliases', {})
    include_fields = data.get('include_fields', None)

    if placeholders is None:
        if not template_path or not Path(template_path).exists():
            return jsonify({'error': '模板文件不存在，请先上传模板'}), 404
        try:
            placeholders = extract_template_placeholders(template_path)
        except Exception as e:
            return jsonify({'error': f'提取占位符失败: {e}'}), 500

    if not placeholders:
        return jsonify({'error': '模板中没有找到任何占位符'}), 400

    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', prefix='excel_template_')
        os.close(fd)

        output_path = generate_excel_template(
            placeholders=placeholders,
            output_path=tmp_path,
            aliases=aliases,
            include_fields=include_fields or placeholders
        )

        template_name = Path(template_path).stem if template_path else 'template'
        download_name = f"{safe_filename(template_name)}_数据模板.xlsx"

        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=download_name
        )
    except Exception as e:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return jsonify({'error': str(e)}), 500


@app.route('/api/match-fields', methods=['POST'])
def match_fields():
    """对比模板占位符与 Excel 列名"""
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

    if template_path and Path(template_path).exists():
        try:
            result['placeholders'] = extract_template_placeholders(template_path)
        except Exception:
            pass

    if excel_path and Path(excel_path).exists():
        try:
            engine = get_excel_engine(excel_path)
            df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=0, engine=engine)
            result['excel_columns'] = list(df.columns)
        except Exception:
            pass

    ph_set = set(result['placeholders'])
    col_set = set(str(c) for c in result['excel_columns'])

    result['matched'] = sorted(ph_set & col_set)
    result['missing_in_excel'] = sorted(ph_set - col_set)
    result['extra_in_excel'] = sorted(col_set - ph_set)

    return jsonify(result)


# ============================================================
#  批量生成
# ============================================================

@app.route('/api/generate', methods=['POST'])
def run_generate():
    data = request.get_json()
    session_id = data.get('session_id', 'default')
    tasks_config = data.get('tasks', [])
    merge_config = data.get('merge_tasks', [])

    if not tasks_config:
        return jsonify({'error': '没有配置任何任务'}), 400

    task_id = uuid.uuid4().hex[:8]
    _, output_dir = get_session_dirs(session_id)

    with task_lock:
        running_tasks[task_id] = {
            'status': 'running',
            'progress': {'current': 0, 'total': 0, 'filename': '', 'status': ''},
            'progress_log': [],
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
            log = running_tasks[task_id]['progress_log']
            if log and log[-1]['current'] == current and log[-1]['filename'] == filename:
                return
            log.append({
                'current': current,
                'total': total,
                'filename': filename,
                'status': status,
                'time': datetime.now().isoformat()
            })
            if len(log) > 200:
                running_tasks[task_id]['progress_log'] = log[-200:]

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
                for i, mc in enumerate(merge_config):
                    subdir = (mc.get('input_subdir') or '').strip()
                    is_root = (subdir == '__root__')

                    if not subdir:
                        merge_results.append({
                            'input_subdir': '',
                            'error': '未指定输入目录',
                            'status': 'error'
                        })
                        continue

                    output_name = (mc.get('output_file') or '').strip() or '合并文档.docx'
                    sort_mode = mc.get('sort_mode', 3)
                    date_kw = mc.get('date_sort_keyword', '')
                    input_dir = output_dir if is_root else output_dir / subdir

                    if not input_dir.exists():
                        merge_results.append({
                            'input_subdir': subdir,
                            'error': f'目录不存在: {subdir}（先生成文档后再合并）',
                            'status': 'error'
                        })
                        continue

                    output_file = output_dir / output_name
                    try:
                        count = merge_documents(str(input_dir), str(output_file), sort_mode, date_kw)
                        merge_results.append({
                            'input_subdir': subdir,
                            'output_file': str(output_file),
                            'count': count,
                            'status': 'success'
                        })
                    except Exception as e:
                        merge_results.append({
                            'input_subdir': subdir,
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
    with task_lock:
        task = running_tasks.get(task_id)
    if not task:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(task)


# ============================================================
#  合并排序预览
# ============================================================

@app.route('/api/merge/preview', methods=['POST'])
def merge_preview():
    data = request.get_json() or {}
    session_id = data.get('session_id', '')
    input_subdir = data.get('input_subdir', '')
    sort_mode = data.get('sort_mode', 3)
    date_sort_keyword = data.get('date_sort_keyword', '')

    if not session_id or not input_subdir:
        return jsonify({'error': '缺少 session_id 或 input_subdir'}), 400

    is_root = (input_subdir == '__root__')
    input_dir = OUTPUT_DIR / session_id if is_root else OUTPUT_DIR / session_id / input_subdir

    try:
        result = preview_merge_order(str(input_dir), sort_mode, date_sort_keyword)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e), 'total': 0, 'files': []}), 500


# ============================================================
#  文件下载
# ============================================================

@app.route('/api/download', methods=['GET'])
def download_file():
    file_path = request.args.get('path', '')
    if not file_path:
        return jsonify({'error': '未指定文件路径'}), 400

    path = Path(file_path)
    if not path.exists():
        return jsonify({'error': '文件不存在'}), 404

    return send_file(str(path), as_attachment=True, download_name=path.name)


@app.route('/api/download/zip', methods=['GET'])
def download_zip():
    session_id = request.args.get('session_id', 'default')
    subdir = request.args.get('subdir', '')

    _, output_dir = get_session_dirs(session_id)
    target_dir = output_dir / subdir if subdir else output_dir

    if not target_dir.exists():
        return jsonify({'error': '目录不存在'}), 404

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
    session_id = request.args.get('session_id', 'default')
    _, output_dir = get_session_dirs(session_id)

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
    preset_path = PRESET_DIR / f"{preset_id}.json"
    if not preset_path.exists():
        return jsonify({'error': '预设不存在'}), 404

    with open(preset_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return jsonify(data)


@app.route('/api/presets/delete/<preset_id>', methods=['DELETE'])
def delete_preset(preset_id):
    preset_path = PRESET_DIR / f"{preset_id}.json"
    if preset_path.exists():
        preset_path.unlink()
        return jsonify({'status': 'ok'})
    return jsonify({'error': '预设不存在'}), 404


# ============================================================
#  文件清理
# ============================================================

def _safe_rmdir(dir_path):
    """安全删除目录（仅限 uploads/outputs 子目录）"""
    p = Path(dir_path)
    allowed_parents = (str(UPLOAD_DIR.resolve()), str(OUTPUT_DIR.resolve()))
    if str(p.parent.resolve()) not in allowed_parents:
        return 0, False
    count = sum(1 for _ in p.rglob('*') if _.is_file())
    shutil.rmtree(str(p), ignore_errors=True)
    return count, True


@app.route('/api/cleanup/session', methods=['POST'])
def cleanup_session():
    data = request.get_json() or {}
    session_id = data.get('session_id', '')
    if not session_id:
        return jsonify({'error': '缺少 session_id'}), 400

    up_count, _ = _safe_rmdir(UPLOAD_DIR / session_id)
    out_count, _ = _safe_rmdir(OUTPUT_DIR / session_id)

    return jsonify({'status': 'ok', 'deleted_files': up_count + out_count})


@app.route('/api/cleanup/all', methods=['POST'])
def cleanup_all():
    def count_and_remove(base_dir):
        total = 0
        if base_dir.exists():
            for child in list(base_dir.iterdir()):
                if child.is_dir():
                    total += sum(1 for _ in child.rglob('*') if _.is_file())
            shutil.rmtree(str(base_dir), ignore_errors=True)
            base_dir.mkdir(parents=True, exist_ok=True)
        return total

    up_count = count_and_remove(UPLOAD_DIR)
    out_count = count_and_remove(OUTPUT_DIR)

    return jsonify({'status': 'ok', 'deleted_files': up_count + out_count})


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    data = request.get_json() or {}
    folder_type = data.get('type', 'outputs')
    session_id = data.get('session_id', '')
    subdir = data.get('subdir', '')

    if folder_type == 'outputs':
        base = OUTPUT_DIR / session_id
    else:
        base = UPLOAD_DIR / session_id

    if subdir:
        base = base / subdir

    if not base.exists():
        return jsonify({'error': '目录不存在，请先生成文件'}), 404

    import subprocess
    subprocess.Popen(['explorer', str(base.resolve())], shell=True)
    return jsonify({'status': 'ok', 'path': str(base.resolve())})


# ============================================================
#  启动
# ============================================================

if __name__ == '__main__':
    import sys
    import socket
    import time

    frozen = getattr(sys, 'frozen', False)
    ERROR_LOG = BASE_DIR / '启动错误.log'
    if frozen and ERROR_LOG.exists():
        try:
            ERROR_LOG.unlink()
        except Exception:
            pass

    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except Exception:
        local_ip = '127.0.0.1'

    def _banner():
        print("=" * 60)
        print("  检验批批量生成平台 v3.0")
        print("=" * 60)
        print(f"  本机访问:   http://127.0.0.1:{PORT}")
        print(f"  局域网访问: http://{local_ip}:{PORT}")
        print("=" * 60)

    if frozen:
        print("正在启动服务，首次加载可能需要 10-30 秒...")
        while True:
            _banner()
            try:
                app.run(host=HOST, port=PORT, debug=False, use_reloader=False)
            except KeyboardInterrupt:
                print("收到退出信号，正在关闭...")
                break
            except Exception as e:
                msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 服务异常: {e}\n{traceback.format_exc()}"
                print(msg)
                try:
                    with open(ERROR_LOG, 'a', encoding='utf-8') as f:
                        f.write(msg + '\n')
                except Exception:
                    pass
                print("服务已停止，3秒后重启...")
                time.sleep(3)
    else:
        try:
            _banner()
            app.run(host=HOST, port=PORT, debug=False)
        except Exception as e:
            print(f"\n[启动失败] {e}")
            traceback.print_exc()
            print("\n按任意键退出...")
            try:
                input()
            except Exception:
                pass

/**
 * ================================================================================
 * 检验批批量生成平台 — 前端交互逻辑 v3.0
 * ================================================================================
 *
 * 工作流程（四步）:
 *   Step 1: 上传文件 — Word 模板 + Excel 数据
 *   Step 2: 任务配置 — 选模板、选 Excel、勾 Sheet、选命名字段
 *   Step 3: 执行生成 — 后台线程 + 前端轮询
 *   Step 4: 结果下载 — 单文件 / ZIP 打包
 *
 * 排序规则:
 *   生成 — 始终按 Excel 原始行顺序，文件名加 001_ 序号前缀
 *   合并 — 3 种模式 (1=编号+续号, 2=日期, 3=Excel原序)
 * ================================================================================
 */

// ============================================================
//  全局状态
// ============================================================
const STATE = {
    sessionId: null,
    files: [],
    tasks: [],
    mergeTasks: [],
    currentTaskId: null,
    templatePlaceholders: {},
    excelColumns: {},
    excelSheets: {}
};

// ============================================================
//  初始化
// ============================================================
document.addEventListener('DOMContentLoaded', async () => {
    await initSession();
    setupDropZones();
    setupFileInputs();
    refreshFiles();
    refreshPresets();
    showServerInfo();
});

async function initSession() {
    try {
        const res = await fetch('/api/session/init', { method: 'POST' });
        const data = await res.json();
        STATE.sessionId = data.session_id;
    } catch (e) {
        showToast('初始化失败，请刷新页面重试', 'error');
    }
}

function showServerInfo() {
    const el = document.getElementById('serverInfo');
    const host = window.location.hostname;
    const port = window.location.port || '5005';
    el.textContent = `局域网 ${host}:${port}`;
}

// ============================================================
//  Toast
// ============================================================
function showToast(msg, type = 'info') {
    let container = document.querySelector('.toast-container');
    if (!container) {
        container = document.createElement('div');
        container.className = 'toast-container';
        document.body.appendChild(container);
    }
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.3s';
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ============================================================
//  文件清理
// ============================================================
async function cleanupAllFiles() {
    if (!confirm('确定要清空所有历史文件吗？\n\n此操作不可恢复。')) return;
    try {
        const res = await fetch('/api/cleanup/all', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
        const data = await res.json();
        if (data.status === 'ok') showToast(`已清空（共 ${data.deleted_files} 个文件）`, 'success');
        else showToast(data.error || '清空失败', 'error');
    } catch (e) { showToast('清空失败: ' + e.message, 'error'); }
}

async function cleanupCurrentSession() {
    if (!confirm('确定要清空本次生成的所有文件吗？\n\n请先下载需要的文件。')) return;
    try {
        const res = await fetch('/api/cleanup/session', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: STATE.sessionId })
        });
        const data = await res.json();
        if (data.status === 'ok') {
            showToast(`已清空（共 ${data.deleted_files} 个文件）`, 'success');
            document.getElementById('outputTree').innerHTML = '<p class="text-muted">已清空</p>';
        } else showToast(data.error || '清空失败', 'error');
    } catch (e) { showToast('清空失败: ' + e.message, 'error'); }
}

async function openOutputFolder() {
    try {
        const res = await fetch('/api/open-folder', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: STATE.sessionId, type: 'outputs' })
        });
        const data = await res.json();
        if (data.status !== 'ok') showToast(data.error || '打开失败', 'error');
    } catch (e) { showToast('打开失败: ' + e.message, 'error'); }
}

// ============================================================
//  文件上传
// ============================================================
function setupDropZones() {
    document.querySelectorAll('.upload-zone').forEach(zone => {
        const type = zone.dataset.type;
        const input = zone.querySelector('.upload-input');
        zone.addEventListener('click', (e) => { if (e.target !== input) input.click(); });
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => { zone.classList.remove('drag-over'); });
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            for (const f of e.dataTransfer.files) uploadFile(f, type);
        });
    });
}

function setupFileInputs() {
    document.getElementById('input-docx').addEventListener('change', (e) => {
        for (const f of e.target.files) uploadFile(f, 'docx');
        e.target.value = '';
    });
    document.getElementById('input-xlsx').addEventListener('change', (e) => {
        for (const f of e.target.files) uploadFile(f, 'xlsx');
        e.target.value = '';
    });
}

async function uploadFile(file, type) {
    if (!STATE.sessionId) { showToast('会话未初始化，请刷新', 'error'); return; }

    // 前端文件大小校验 (500MB)
    const MAX_SIZE = 500 * 1024 * 1024;
    if (file.size > MAX_SIZE) {
        showToast(`文件过大: ${formatSize(file.size)}，最大允许 ${formatSize(MAX_SIZE)}`, 'error');
        return;
    }

    const form = new FormData();
    form.append('file', file);
    form.append('session_id', STATE.sessionId);
    form.append('type', type);
    try {
        const res = await fetch('/api/upload', { method: 'POST', body: form });
        const data = await res.json();
        if (data.status === 'ok') {
            showToast(`已上传: ${file.name}`, 'success');
            await refreshFiles();
        } else showToast(data.error || '上传失败', 'error');
    } catch (e) { showToast('上传失败: ' + e.message, 'error'); }
}

async function refreshFiles() {
    if (!STATE.sessionId) return;
    try {
        const res = await fetch(`/api/files?session_id=${STATE.sessionId}`);
        const data = await res.json();
        STATE.files = data.files;

        // 清除已删文件的任务引用
        const currentPaths = new Set(STATE.files.map(f => f.path));
        let tasksChanged = false;
        STATE.tasks.forEach((task, i) => {
            if (task.template_path && !currentPaths.has(task.template_path)) {
                task.template_path = '';
                delete STATE.templatePlaceholders[i];
                tasksChanged = true;
            }
            if (task.excel_path && !currentPaths.has(task.excel_path)) {
                task.excel_path = '';
                task.sheet_names = [];
                task.sheet_name = '';
                task.filename_column = '';
                delete STATE.excelSheets[i];
                delete STATE.excelColumns[i];
                tasksChanged = true;
            }
        });
        if (tasksChanged) showToast('已自动清除被删文件关联的任务配置', 'info');

        renderFileList();
        if (STATE.tasks.length > 0) renderTasks();
    } catch (e) { console.error('刷新文件列表失败', e); }
}

function renderFileList() {
    renderPanelFileList('word', STATE.files.filter(f => f.type === 'word'));
    renderPanelFileList('excel', STATE.files.filter(f => f.type === 'excel'));
}

function renderPanelFileList(type, files) {
    const listEl = document.getElementById(type === 'word' ? 'wordFileList' : 'excelFileList');
    const itemsEl = document.getElementById(type === 'word' ? 'wordFileItems' : 'excelFileItems');
    const countEl = document.getElementById(type === 'word' ? 'wordFileCount' : 'excelFileCount');

    if (!files.length) { listEl.style.display = 'none'; return; }
    listEl.style.display = 'block';
    countEl.textContent = files.length;
    itemsEl.innerHTML = files.map(f => `
        <div class="panel-file-item">
            <span class="file-type-icon">${type === 'word' ? '📄' : '📊'}</span>
            <div class="file-info">
                <div class="file-name" title="${f.original_name}">${f.original_name}</div>
                <div class="file-meta">${formatSize(f.size)}</div>
            </div>
            <div class="file-item-actions">
                <button class="btn btn-xs btn-ghost" onclick="previewFile('${f.type}', '${f.name}', '${f.original_name}', '${f.path}')" title="预览内容" style="font-size:12px;padding:2px 6px;">👁</button>
                <button class="file-del" onclick="deleteFile('${f.name}')" title="删除">&times;</button>
            </div>
        </div>
    `).join('') + `<div style="margin-top:8px;text-align:right;"><button class="btn btn-xs btn-ghost" onclick="deleteAllFiles('${type}')" style="color:var(--danger);font-size:11px;">🗑 全部删除</button></div>`;
}

async function deleteFile(filename) {
    if (!confirm('确定删除？')) return;
    try {
        const res = await fetch('/api/files/delete', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: STATE.sessionId, filename })
        });
        if (!res.ok) { showToast('删除失败', 'error'); return; }

        // 本地移除文件，不触发全量刷新
        const deletedFile = STATE.files.find(f => f.name === filename);
        const deletedPath = deletedFile ? deletedFile.path : null;
        STATE.files = STATE.files.filter(f => f.name !== filename);

        // 清除任务中对该文件的引用
        let tasksChanged = false;
        if (deletedPath) {
            STATE.tasks.forEach((task, i) => {
                if (task.template_path === deletedPath) {
                    task.template_path = '';
                    delete STATE.templatePlaceholders[i];
                    tasksChanged = true;
                }
                if (task.excel_path === deletedPath) {
                    task.excel_path = '';
                    task.sheet_names = [];
                    task.sheet_name = '';
                    task.filename_column = '';
                    delete STATE.excelColumns[i];
                    delete STATE.excelSheets[i];
                    delete task.row_filter;
                    delete task.row_filter_total;
                    tasksChanged = true;
                }
            });
        }

        renderFileList();
        if (tasksChanged) {
            showToast('文件已删除，关联的任务配置已自动清除', 'info');
            if (STATE.tasks.length > 0) renderTasks();
        } else {
            showToast('文件已删除', 'info');
        }
    } catch (e) { showToast('删除失败: ' + e.message, 'error'); }
}

// 批量删除某类型文件
async function deleteAllFiles(type) {
    const files = STATE.files.filter(f => f.type === type);
    if (files.length === 0) return;
    const typeLabel = type === 'word' ? 'Word 模板' : 'Excel 数据';
    if (!confirm(`确定删除全部 ${files.length} 个${typeLabel}文件？`)) return;

    let ok = 0, fail = 0;
    const deletedPaths = new Set();
    for (const f of files) {
        try {
            const res = await fetch('/api/files/delete', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: STATE.sessionId, filename: f.name })
            });
            if (res.ok) { ok++; deletedPaths.add(f.path); } else { fail++; }
        } catch (e) { fail++; }
    }

    // 本地移除已删文件
    STATE.files = STATE.files.filter(f => !deletedPaths.has(f.path));

    // 清除任务中引用
    let tasksChanged = false;
    STATE.tasks.forEach((task, i) => {
        if (task.template_path && deletedPaths.has(task.template_path)) {
            task.template_path = '';
            delete STATE.templatePlaceholders[i];
            tasksChanged = true;
        }
        if (task.excel_path && deletedPaths.has(task.excel_path)) {
            task.excel_path = '';
            task.sheet_names = [];
            task.sheet_name = '';
            task.filename_column = '';
            delete STATE.excelSheets[i];
            delete STATE.excelColumns[i];
            delete task.row_filter;
            delete task.row_filter_total;
            tasksChanged = true;
        }
    });

    renderFileList();
    if (tasksChanged && STATE.tasks.length > 0) renderTasks();
    showToast(`已删除 ${ok} 个文件${fail > 0 ? `，${fail} 个失败` : ''}`, fail > 0 ? 'error' : 'success');
}

function formatSize(bytes) {
    if (!bytes) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    let i = 0, size = bytes;
    while (size >= 1024 && i < units.length - 1) { size /= 1024; i++; }
    return size.toFixed(i === 0 ? 0 : 1) + ' ' + units[i];
}

// ============================================================
//  文件预览 — Word / Excel 内容查看
// ============================================================
async function previewFile(fileType, fileName, originalName, filePath) {
    const inner = document.getElementById('loadPresetModalInner');
    inner.classList.add('modal-xwide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = `👁 预览: ${originalName}`;
    document.getElementById('loadPresetBody').innerHTML = '<p style="text-align:center;padding:40px;color:var(--text-muted);">正在读取文件内容...</p>';
    document.getElementById('loadPresetModal').style.display = 'flex';

    if (fileType === 'word') {
        await previewDocx(filePath);
    } else if (fileType === 'excel') {
        await previewExcelFile(filePath);
    }
}

async function previewDocx(filePath) {
    try {
        const res = await fetch('/api/preview/docx', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: filePath })
        });
        const data = await res.json();
        if (data.error) {
            document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">${data.error}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
            return;
        }

        const paras = data.paragraphs || [];
        const tables = data.tables || [];
        let html = `<div style="margin-bottom:12px;padding:8px 12px;background:var(--bg-card);border-radius:6px;font-size:12px;color:var(--text-muted);">
            📄 共 ${data.total_paragraphs} 段文本 · ${data.total_tables} 个表格 · ${data.total_chars} 字符
        </div>`;

        // 渲染段落
        if (paras.length > 0) {
            html += '<div style="max-height:50vh;overflow-y:auto;border:1px solid var(--border);border-radius:6px;padding:12px 16px;background:var(--bg-card);">';
            for (const p of paras) {
                const isHeading = (p.style || '').toLowerCase().includes('heading') || (p.style || '').toLowerCase().includes('title');
                if (isHeading) {
                    html += `<p style="font-weight:700;font-size:14px;margin:12px 0 6px;color:var(--accent);">${escapeHtml(p.text)}</p>`;
                } else {
                    html += `<p style="margin:4px 0;line-height:1.8;">${escapeHtml(p.text)}</p>`;
                }
            }
            html += '</div>';
        }

        // 渲染表格
        if (tables.length > 0) {
            for (let ti = 0; ti < tables.length; ti++) {
                const table = tables[ti];
                html += `<p style="margin:16px 0 6px;font-weight:600;font-size:13px;">📋 表格 ${ti + 1} (${table.length} 行 × ${table[0]?.length || 0} 列)</p>`;
                html += '<div style="overflow-x:auto;border:1px solid var(--border);border-radius:6px;"><table class="preview-table" style="margin:0;">';
                html += '<tbody>';
                table.forEach((row, ri) => {
                    const isHeader = ri === 0;
                    html += `<tr${isHeader ? ' style="background:var(--bg-card);font-weight:600;"' : ''}>`;
                    row.forEach(cell => {
                        html += isHeader ? `<th>${escapeHtml(cell)}</th>` : `<td>${escapeHtml(cell)}</td>`;
                    });
                    html += '</tr>';
                });
                html += '</tbody></table></div>';
            }
        }

        if (paras.length === 0 && tables.length === 0) {
            html += '<p style="text-align:center;padding:30px;color:var(--text-muted);">文档内容为空</p>';
        }

        html += `<div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
        document.getElementById('loadPresetBody').innerHTML = html;
    } catch (e) {
        document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">预览失败: ${e.message}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
    }
}

async function previewExcelFile(filePath) {
    try {
        // 先获取 Sheet 列表
        const sheetsRes = await fetch('/api/excel/sheets', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: filePath })
        });
        const sheetsData = await sheetsRes.json();
        if (sheetsData.error) {
            document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">${sheetsData.error}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
            return;
        }
        const sheets = sheetsData.sheets || [];
        if (sheets.length === 0) {
            document.getElementById('loadPresetBody').innerHTML = `<p style="text-align:center;padding:20px;color:var(--text-muted);">Excel 中没有工作表</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
            return;
        }

        // 默认预览第一个 Sheet
        window._excelPreviewState = { filePath, sheets, currentSheet: sheets[0] };
        await renderExcelPreviewSheet(sheets[0]);
    } catch (e) {
        document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">预览失败: ${e.message}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
    }
}

async function renderExcelPreviewSheet(sheetName) {
    const state = window._excelPreviewState;
    if (!state) return;
    state.currentSheet = sheetName;

    document.getElementById('loadPresetBody').innerHTML = '<p style="text-align:center;padding:30px;color:var(--text-muted);">正在读取数据...</p>';

    try {
        const res = await fetch('/api/excel/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: state.filePath, sheet_name: sheetName, rows: 50 })
        });
        const data = await res.json();
        if (data.error) {
            document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">${data.error}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
            return;
        }

        const cols = data.columns || [];
        const rows = data.preview || [];
        const totalRows = data.total_rows || 0;

        // Sheet 切换按钮
        let html = `<div style="margin-bottom:12px;display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
            <span style="font-size:12px;color:var(--text-muted);">工作表:</span>`;
        state.sheets.forEach(s => {
            const active = s === sheetName;
            html += `<button class="btn btn-xs ${active ? 'btn-primary' : 'btn-outline'}" onclick="renderExcelPreviewSheet('${s}')" style="font-size:11px;padding:2px 10px;">${s}</button>`;
        });
        html += `</div>`;

        // 统计信息
        html += `<div style="margin-bottom:8px;font-size:12px;color:var(--text-muted);">📊 ${sheetName} — 共 ${totalRows} 行 · 显示前 ${rows.length} 行 · ${cols.length} 列</div>`;

        // 数据表格
        if (cols.length > 0 && rows.length > 0) {
            html += '<div style="overflow:auto;max-height:50vh;border:1px solid var(--border);border-radius:6px;"><table class="preview-table" style="margin:0;"><thead><tr>';
            html += '<th style="width:40px;text-align:center;color:var(--text-muted);">#</th>';
            cols.forEach(c => { html += `<th>${escapeHtml(String(c))}</th>`; });
            html += '</tr></thead><tbody>';
            rows.forEach(r => {
                html += `<tr><td style="text-align:center;color:var(--text-muted);font-size:11px;">${(r._row_index ?? 0) + 1}</td>`;
                cols.forEach(c => {
                    const val = r[c];
                    html += `<td>${val !== undefined && val !== null ? escapeHtml(String(val).substring(0, 100)) : ''}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table></div>';
        } else {
            html += '<p style="text-align:center;padding:30px;color:var(--text-muted);">此工作表没有数据</p>';
        }

        html += `<div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
        document.getElementById('loadPresetBody').innerHTML = html;
    } catch (e) {
        document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">读取失败: ${e.message}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
    }
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============================================================
//  任务配置
// ============================================================
function addTaskRow(presetData = null) {
    const sheetNames = presetData?.sheet_names || (presetData?.sheet_name ? [presetData.sheet_name] : []);
    STATE.tasks.push({
        template_path: presetData?.template_path || '',
        excel_path: presetData?.excel_path || '',
        sheet_name: sheetNames[0] || '',
        sheet_names: sheetNames,
        filename_column: presetData?.filename_column || '',
        output_subdir: presetData?.output_subdir || '',
        add_index_prefix: presetData?.add_index_prefix !== false,
        int_columns: presetData?.int_columns || [],
        date_format: presetData?.date_format || '%Y年%m月%d日'
    });
    renderTasks();
}

function removeTaskRow(index) {
    STATE.tasks.splice(index, 1);
    delete STATE.templatePlaceholders[index];
    delete STATE.excelColumns[index];
    delete STATE.excelSheets[index];
    renderTasks();
}

function updateTask(index, field, value) {
    STATE.tasks[index][field] = value;
}

function renderTasks() {
    const list = document.getElementById('taskList');
    const empty = document.getElementById('taskEmpty');
    const settings = document.getElementById('globalSettings');

    if (STATE.tasks.length === 0) {
        list.innerHTML = '';
        empty.style.display = 'block';
        settings.style.display = 'none';
        return;
    }

    empty.style.display = 'none';
    settings.style.display = 'block';

    const wordFiles = STATE.files.filter(f => f.type === 'word');
    const excelFiles = STATE.files.filter(f => f.type === 'excel');

    list.innerHTML = STATE.tasks.map((task, i) => {
        const placeholders = STATE.templatePlaceholders[i] || [];
        const columns = STATE.excelColumns[i] || [];
        const sheets = STATE.excelSheets[i] || [];
        const selectedSheets = task.sheet_names || (task.sheet_name ? [task.sheet_name] : []);

        return `
        <div class="task-card">
            <div class="task-card-header">
                <span class="task-card-title">
                    <span class="task-badge">${i + 1}</span>
                    任务 ${i + 1}
                </span>
                <div class="task-card-actions">
                    <button class="btn btn-xs btn-outline" onclick="checkFieldMatch(${i})">字段匹配</button>
                    <button class="btn btn-xs btn-primary" onclick="showTemplateWorkbench(${i})" ${!task.template_path ? 'disabled' : ''} style="${!task.template_path ? '' : 'background:var(--accent);border-color:var(--accent);'}">📋 模板工作台</button>
                    <button class="btn btn-xs btn-outline" onclick="showDataPreview(${i})">数据预览</button>
                    <button class="btn btn-xs btn-outline" onclick="showRowSelector(${i})">选择生成行</button>
                    <button class="btn btn-xs btn-ghost" onclick="removeTaskRow(${i})" style="color:var(--danger);">删除</button>
                </div>
            </div>
            <div class="task-card-body">
                <div class="form-group">
                    <label class="form-label">Word 模板 <span style="color:var(--danger);">*</span></label>
                    <select class="form-select" onchange="onTemplateChange(${i}, this.value)">
                        <option value="">-- 选择模板 --</option>
                        ${wordFiles.map(f => `<option value="${f.path}" ${task.template_path === f.path ? 'selected' : ''}>${f.original_name}</option>`).join('')}
                    </select>
                </div>
                <div class="form-group">
                    <label class="form-label">Excel 数据文件 <span style="color:var(--danger);">*</span></label>
                    <select class="form-select" onchange="onExcelChange(${i}, this.value)">
                        <option value="">-- 选择 Excel --</option>
                        ${excelFiles.map(f => `<option value="${f.path}" ${task.excel_path === f.path ? 'selected' : ''}>${f.original_name}</option>`).join('')}
                    </select>
                </div>
                <div class="form-group form-group-full">
                    <label class="form-label">工作表 (Sheet) <span style="color:var(--danger);">*</span> <span style="font-size:11px;color:var(--text-muted);font-weight:normal;">可多选</span></label>
                    <div class="sheet-checkbox-list" id="sheetList${i}">
                        ${sheets.length > 0
                            ? `<div style="margin-bottom:6px;display:flex;gap:6px;">
                                <button class="btn btn-xs btn-outline" onclick="sheetSelectAll(${i})" style="font-size:11px;padding:2px 8px;">全选</button>
                                <button class="btn btn-xs btn-outline" onclick="sheetDeselectAll(${i})" style="font-size:11px;padding:2px 8px;">取消</button>
                              </div>` + sheets.map(s => `<label class="checkbox-label sheet-checkbox"><input type="checkbox" value="${s}" ${selectedSheets.includes(s) ? 'checked' : ''} onchange="onSheetCheckboxChange(${i})"><span>${s}</span></label>`).join('')
                            : `<span class="text-muted" style="font-size:12px;">请先选择 Excel 文件</span>`
                        }
                    </div>
                </div>
                <div class="form-group">
                    <label class="form-label">文件命名字段 <span style="color:var(--danger);">*</span></label>
                    <select class="form-select" id="filenameSelect${i}" onchange="updateTask(${i}, 'filename_column', this.value)">
                        ${columns.length > 0
                            ? `<option value="">-- 选择列 --</option>` + columns.map(c => `<option value="${c}" ${task.filename_column === c ? 'selected' : ''}>${c}</option>`).join('')
                            : `<option value="${task.filename_column || ''}">${task.filename_column || '(请先选择Excel)'}</option>`
                        }
                    </select>
                </div>
                <div class="form-group">
                    <label class="form-label">输出子目录名</label>
                    <input class="form-input" value="${task.output_subdir}" placeholder="如：AB、隐蔽工程..." onchange="updateTask(${i}, 'output_subdir', this.value)">
                </div>
                <div class="form-group">
                    <label class="form-label">日期格式</label>
                    <input class="form-input" value="${task.date_format}" placeholder="%Y年%m月%d日" onchange="updateTask(${i}, 'date_format', this.value)">
                </div>
                <div class="form-group">
                    <label class="form-label">强制整数列 (逗号分隔)</label>
                    <input class="form-input" value="${(task.int_columns || []).join(', ')}" placeholder="如：根设AB, 间距" onchange="updateTask(${i}, 'int_columns', this.value.split(',').map(s => s.trim()).filter(Boolean))">
                </div>
                <div class="form-group" style="flex-direction:row;align-items:center;">
                    <label class="checkbox-label">
                        <input type="checkbox" ${task.add_index_prefix ? 'checked' : ''} onchange="updateTask(${i}, 'add_index_prefix', this.checked)">
                        <span>文件名加序号前缀 (001_)</span>
                    </label>
                </div>
            </div>
            ${placeholders.length > 0 ? `
                <div class="placeholder-section">
                    <div class="placeholder-header">
                        <span class="placeholder-title">模板占位符 (${placeholders.length} 个)</span>
                        <button class="btn btn-sm btn-ghost" onclick="togglePlaceholders(${i})" id="phToggle${i}">收起</button>
                    </div>
                    <div class="placeholder-tags" id="phTags${i}">
                        ${placeholders.map(p => {
                            const inExcel = columns.includes(p);
                            return `<span class="ph-tag ${inExcel ? 'ph-matched' : 'ph-missing'}" title="${inExcel ? 'Excel中已有此列' : 'Excel中未找到此列'}">${p}${inExcel ? ' ✓' : ' ⚠'}</span>`;
                        }).join('')}
                    </div>
                </div>
            ` : ''}
            ${(() => {
                const rf = task.row_filter;
                if (rf === undefined || rf === null) return '';
                const rowInfo = getRowFilterSummary(rf);
                return `<div class="task-card-footer"><div class="row-filter-status"><span style="color:var(--accent);">📋 ${rowInfo.text}</span><button class="btn btn-xs btn-ghost" onclick="clearRowFilter(${i})" style="color:var(--text-muted);">清除选择</button></div></div>`;
            })()}
        </div>`;
    }).join('');

    renderMergeTasks();
}

// ============================================================
//  模板占位符 & Excel 列名联动
// ============================================================
async function onTemplateChange(taskIndex, templatePath) {
    updateTask(taskIndex, 'template_path', templatePath);
    if (!templatePath) { delete STATE.templatePlaceholders[taskIndex]; renderTasks(); return; }
    try {
        const res = await fetch('/api/template/placeholders', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: templatePath })
        });
        const data = await res.json();
        if (data.placeholders) {
            STATE.templatePlaceholders[taskIndex] = data.placeholders;
            showToast(`模板中发现 ${data.placeholders.length} 个占位符`, 'success');
        }
    } catch (e) { console.error('获取占位符失败', e); }
    renderTasks();
}

async function onExcelChange(taskIndex, excelPath) {
    updateTask(taskIndex, 'excel_path', excelPath);
    delete STATE.tasks[taskIndex].row_filter;
    delete STATE.tasks[taskIndex].row_filter_total;
    if (!excelPath) {
        delete STATE.excelSheets[taskIndex];
        delete STATE.excelColumns[taskIndex];
        STATE.tasks[taskIndex].sheet_names = [];
        STATE.tasks[taskIndex].sheet_name = '';
        renderTasks();
        return;
    }
    try {
        const res = await fetch('/api/excel/sheets', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: excelPath })
        });
        const data = await res.json();
        if (data.sheets && data.sheets.length > 0) {
            STATE.excelSheets[taskIndex] = data.sheets;
            STATE.tasks[taskIndex].sheet_names = [data.sheets[0]];
            STATE.tasks[taskIndex].sheet_name = data.sheets[0];
            showToast(`Excel 中发现 ${data.sheets.length} 个工作表`, 'success');
            await fetchColumns(taskIndex);
        } else if (data.error) {
            showToast(`读取 Excel 失败: ${data.error}`, 'error');
        }
    } catch (e) { showToast('获取工作表失败: ' + e.message, 'error'); }
    renderTasks();
}

async function onSheetCheckboxChange(taskIndex) {
    const checkboxes = document.querySelectorAll(`#sheetList${taskIndex} input[type="checkbox"]:checked`);
    const selectedSheets = Array.from(checkboxes).map(cb => cb.value);
    STATE.tasks[taskIndex].sheet_names = selectedSheets;
    STATE.tasks[taskIndex].sheet_name = selectedSheets[0] || '';
    delete STATE.tasks[taskIndex].row_filter;
    delete STATE.tasks[taskIndex].row_filter_total;
    if (selectedSheets.length > 0) {
        await fetchColumns(taskIndex);
        renderTasks();
    } else {
        delete STATE.excelColumns[taskIndex];
        renderTasks();
    }
}

function sheetSelectAll(taskIndex) {
    const checkboxes = document.querySelectorAll(`#sheetList${taskIndex} input[type="checkbox"]`);
    checkboxes.forEach(cb => cb.checked = true);
    onSheetCheckboxChange(taskIndex);
}

function sheetDeselectAll(taskIndex) {
    const checkboxes = document.querySelectorAll(`#sheetList${taskIndex} input[type="checkbox"]`);
    checkboxes.forEach(cb => cb.checked = false);
    onSheetCheckboxChange(taskIndex);
}

async function fetchColumns(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.excel_path || !task.sheet_name) return;
    try {
        const res = await fetch('/api/excel/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: task.excel_path, sheet_name: task.sheet_name, rows: 1 })
        });
        const data = await res.json();
        if (data.columns) STATE.excelColumns[taskIndex] = data.columns.map(String);
    } catch (e) { console.error('获取列名失败', e); }
}

function togglePlaceholders(taskIndex) {
    const tags = document.getElementById(`phTags${taskIndex}`);
    const btn = document.getElementById(`phToggle${taskIndex}`);
    if (tags.style.display === 'none') { tags.style.display = 'flex'; btn.textContent = '收起'; }
    else { tags.style.display = 'none'; btn.textContent = '展开'; }
}

// ============================================================
//  字段匹配检查
// ============================================================
async function checkFieldMatch(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.template_path || !task.excel_path) { showToast('请先选择模板和Excel', 'error'); return; }
    try {
        const res = await fetch('/api/match-fields', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ template_path: task.template_path, excel_path: task.excel_path, sheet_name: task.sheet_name })
        });
        const data = await res.json();
        const inner = document.getElementById('loadPresetModalInner');
        inner.classList.add('modal-wide');
        document.querySelector('#loadPresetModal .modal-header h3').textContent = `字段匹配检查 — 任务 ${taskIndex + 1}`;
        document.getElementById('loadPresetBody').innerHTML = `
            <div class="field-match-result">
                <div class="match-section match-ok">
                    <h4>已匹配 (${data.matched.length})</h4>
                    <div class="match-tags">${data.matched.length > 0 ? data.matched.map(m => `<span class="ph-tag ph-matched">${m}</span>`).join('') : '<span class="text-muted">无</span>'}</div>
                </div>
                <div class="match-section match-missing">
                    <h4>模板有但Excel缺 (${data.missing_in_excel.length})</h4>
                    <div class="match-tags">${data.missing_in_excel.length > 0 ? data.missing_in_excel.map(m => `<span class="ph-tag ph-missing">${m}</span>`).join('') : '<span class="text-muted">无缺失，全部匹配</span>'}</div>
                    <p class="match-hint">这些占位符在Excel中找不到对应列，生成时将为空值</p>
                </div>
                <div class="match-section match-extra">
                    <h4>Excel有但模板没用 (${data.extra_in_excel.length})</h4>
                    <div class="match-tags">${data.extra_in_excel.length > 0 ? data.extra_in_excel.map(m => `<span class="ph-tag ph-extra">${m}</span>`).join('') : '<span class="text-muted">无</span>'}</div>
                    <p class="match-hint">这些列不会被使用（不影响生成）</p>
                </div>
            </div>
            <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>
        `;
        document.getElementById('loadPresetModal').style.display = 'flex';
    } catch (e) { showToast('字段匹配检查失败', 'error'); }
}

// ============================================================
//  数据预览
// ============================================================
async function showDataPreview(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.excel_path) { showToast('请先选择Excel', 'error'); return; }
    try {
        const res = await fetch('/api/excel/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: task.excel_path, sheet_name: task.sheet_name, rows: 8 })
        });
        const data = await res.json();
        const inner = document.getElementById('loadPresetModalInner');
        inner.classList.add('modal-wide');
        document.querySelector('#loadPresetModal .modal-header h3').textContent = `数据预览 — 任务 ${taskIndex + 1} (${data.total_rows} 行)`;
        const cols = data.columns || [];
        const rows = data.preview || [];
        document.getElementById('loadPresetBody').innerHTML = `
            <div style="overflow-x:auto;">
                <table class="preview-table"><thead><tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr></thead>
                <tbody>${rows.map(r => `<tr>${cols.map(c => `<td>${r[c] !== undefined && r[c] !== null ? r[c] : ''}</td>`).join('')}</tr>`).join('')}</tbody></table>
            </div>
            <p style="margin-top:8px;font-size:12px;color:var(--text-muted);">仅显示前 ${rows.length} 行，共 ${data.total_rows} 行</p>
            <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>
        `;
        document.getElementById('loadPresetModal').style.display = 'flex';
    } catch (e) { showToast('数据预览失败', 'error'); }
}

// ============================================================
//  行选择器 — 按需生成
// ============================================================
function getRowFilterSummary(rowFilter) {
    if (rowFilter === undefined || rowFilter === null) return { text: '将生成全部数据行', count: null };
    if (rowFilter.length === 0) return { text: '未选择任何行（不会生成）', count: 0 };
    return { text: `已选择 ${rowFilter.length} 行`, count: rowFilter.length };
}

function clearRowFilter(taskIndex) {
    STATE.tasks[taskIndex].row_filter = null;
    STATE.tasks[taskIndex].row_filter_total = null;
    renderTasks();
}

async function showRowSelector(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.excel_path) { showToast('请先选择 Excel', 'error'); return; }
    const sheets = task.sheet_names || (task.sheet_name ? [task.sheet_name] : []);
    if (sheets.length === 0) { showToast('请先选择工作表', 'error'); return; }
    if (sheets.length > 1) showToast('提示：多 Sheet 任务每个 Sheet 独立选择行', 'info');

    const sheetName = sheets[0];
    const currentFilter = task.row_filter || null;

    const inner = document.getElementById('loadPresetModalInner');
    inner.classList.add('modal-xwide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = `选择生成行 — 任务 ${taskIndex + 1} (${sheetName})`;
    document.getElementById('loadPresetBody').innerHTML = '<p style="text-align:center;padding:30px;color:var(--text-muted);">正在读取数据...</p>';
    document.getElementById('loadPresetModal').style.display = 'flex';

    try {
        const res = await fetch('/api/excel/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: task.excel_path, sheet_name: sheetName, rows: 0 })
        });
        const data = await res.json();
        if (data.error) { showToast(data.error, 'error'); closeLoadPreset(); return; }

        const cols = data.columns || [];
        const rows = data.preview || [];
        const totalRows = data.total_rows;
        const selectedSet = new Set();
        if (currentFilter) currentFilter.forEach(idx => selectedSet.add(idx));
        else rows.forEach(r => selectedSet.add(r._row_index));

        const renderTable = () => {
            const checkedCount = rows.filter(r => selectedSet.has(r._row_index)).length;
            const allChecked = checkedCount === rows.length;
            let tableHtml = `
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap;">
                <button class="btn btn-sm btn-outline" onclick="rowSelectorSelectAll()" style="font-size:12px;">✅ 全选</button>
                <button class="btn btn-sm btn-outline" onclick="rowSelectorDeselectAll()" style="font-size:12px;">❎ 取消全选</button>
                <button class="btn btn-sm btn-outline" onclick="rowSelectorInvert()" style="font-size:12px;">🔄 反选</button>
                <span style="margin-left:8px;font-weight:600;color:var(--accent);">已选 ${checkedCount} / ${rows.length} 行</span>
            </div>
            <div style="overflow:auto;max-height:50vh;border:1px solid var(--border);border-radius:6px;">
                <table class="preview-table" style="margin:0;">
                    <thead><tr>
                        <th style="width:40px;text-align:center;"><input type="checkbox" ${allChecked ? 'checked' : ''} onchange="rowSelectorToggleAll(this.checked)" id="rowSelectAll"></th>
                        <th style="width:50px;">#</th>
                        ${cols.map(c => `<th>${c}</th>`).join('')}
                    </tr></thead>
                    <tbody>
                        ${rows.map(r => {
                            const idx = r._row_index;
                            const checked = selectedSet.has(idx);
                            return `<tr class="${checked ? 'row-selected' : ''}" style="${checked ? 'background:#e8f5e9;' : ''}">
                                <td style="text-align:center;"><input type="checkbox" ${checked ? 'checked' : ''} onchange="rowSelectorToggleOne(${idx}, this.checked)"></td>
                                <td style="color:var(--text-muted);font-size:11px;">${idx + 1}</td>
                                ${cols.map(c => `<td>${r[c] !== undefined && r[c] !== null ? String(r[c]).substring(0, 80) : ''}</td>`).join('')}
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>`;
            document.getElementById('loadPresetBody').innerHTML = tableHtml + `
            <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;">
                <button class="btn btn-outline" onclick="closeLoadPreset()">取消</button>
                <button class="btn btn-primary" onclick="rowSelectorConfirm()">✅ 确认选择（${checkedCount} 行）</button>
            </div>`;
        };

        window._rowSelectorData = { taskIndex, rows, selectedSet, sheetName, renderTable };
        renderTable();
    } catch (e) {
        document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;">读取失败: ${e.message}</p>`;
    }
}

function rowSelectorToggleOne(rowIdx, checked) {
    const d = window._rowSelectorData; if (!d) return;
    if (checked) d.selectedSet.add(rowIdx); else d.selectedSet.delete(rowIdx);
    d.renderTable();
}
function rowSelectorToggleAll(checked) {
    const d = window._rowSelectorData; if (!d) return;
    if (checked) d.rows.forEach(r => d.selectedSet.add(r._row_index)); else d.selectedSet.clear();
    d.renderTable();
}
function rowSelectorSelectAll() {
    const d = window._rowSelectorData; if (!d) return;
    d.rows.forEach(r => d.selectedSet.add(r._row_index)); d.renderTable();
}
function rowSelectorDeselectAll() {
    const d = window._rowSelectorData; if (!d) return;
    d.selectedSet.clear(); d.renderTable();
}
function rowSelectorInvert() {
    const d = window._rowSelectorData; if (!d) return;
    d.rows.forEach(r => { if (d.selectedSet.has(r._row_index)) d.selectedSet.delete(r._row_index); else d.selectedSet.add(r._row_index); });
    d.renderTable();
}
function rowSelectorConfirm() {
    const d = window._rowSelectorData; if (!d) return;
    const allIndices = d.rows.map(r => r._row_index);
    const selected = allIndices.filter(idx => d.selectedSet.has(idx));
    if (selected.length === 0) { showToast('请至少选择一行', 'error'); return; }
    if (selected.length === allIndices.length) {
        STATE.tasks[d.taskIndex].row_filter = null;
        STATE.tasks[d.taskIndex].row_filter_total = allIndices.length;
    } else {
        selected.sort((a, b) => a - b);
        STATE.tasks[d.taskIndex].row_filter = selected;
        STATE.tasks[d.taskIndex].row_filter_total = allIndices.length;
    }
    window._rowSelectorData = null;
    closeLoadPreset();
    renderTasks();
    showToast(`已选择 ${selected.length} / ${allIndices.length} 行`, 'success');
}

// ============================================================
//  合并任务
// ============================================================
const MERGE_MODE_META = {
    1: { name: '按文档内编号+续号', desc: '提取文档中的编号模式（如 4-1-(5-1)），按编号层级排序。', fields: [], hint: '自动识别 数字-数字-数字 格式及 (续N) 标记' },
    2: { name: '按文档内日期', desc: '在文档中查找指定关键词附近的日期，按时间排序。', fields: ['date_sort_keyword'], hint: '可指定关键词如"施工日期"，或留空自动识别' },
    3: { name: '按Excel原始顺序', desc: '读取文件名 001_ 序号前缀，按数字升序排列（默认）。', fields: [], hint: '依赖生成时的「序号前缀」开关' },
};

function toggleMergeSettings() {
    const enabled = document.getElementById('enableMerge').checked;
    document.getElementById('mergeSettings').style.display = enabled ? 'block' : 'none';
    if (enabled && STATE.mergeTasks.length === 0) addMergeRow();
}

function addMergeRow() {
    STATE.mergeTasks.push({ input_subdir: '', output_file: '合并文档.docx', sort_mode: 3, date_sort_keyword: '' });
    renderMergeTasks();
}

function removeMergeRow(index) { STATE.mergeTasks.splice(index, 1); renderMergeTasks(); }

function updateMergeTask(index, field, value) {
    STATE.mergeTasks[index][field] = value;
    if (field === 'sort_mode') renderMergeTasks();
}

function getAvailableMergeDirs() {
    const dirs = [];
    const seen = new Set();
    const hasRootTask = STATE.tasks.some(t => !(t.output_subdir || '').trim());
    if (hasRootTask) { dirs.push({ value: '__root__', label: '（根目录 — 所有直接生成的文件）' }); seen.add('__root__'); }
    STATE.tasks.forEach((t, i) => {
        const subdir = (t.output_subdir || '').trim();
        if (!subdir) return;
        const sheets = t.sheet_names || (t.sheet_name ? [t.sheet_name] : []);
        if (sheets.length > 1) {
            sheets.forEach(s => {
                const val = `${subdir}/${s}`;
                if (!seen.has(val)) { seen.add(val); dirs.push({ value: val, label: `${subdir} / ${s} (任务 ${i + 1})` }); }
            });
        } else {
            if (!seen.has(subdir)) { seen.add(subdir); dirs.push({ value: subdir, label: `${subdir} (任务 ${i + 1})` }); }
        }
    });
    return dirs;
}

function renderMergeTasks() {
    const list = document.getElementById('mergeList');
    if (!list) return;
    const dirs = getAvailableMergeDirs();
    const hint = dirs.length === 0 ? '<p class="text-muted" style="font-size:12px;margin-bottom:8px;">提示：当前无可用输出目录，请先设置「输出子目录名」</p>' : '';

    list.innerHTML = hint + STATE.mergeTasks.map((mt, i) => {
        const mode = mt.sort_mode || 3;
        const meta = MERGE_MODE_META[mode];
        const hasDir = !!(mt.input_subdir || '').trim();
        let configHtml = mode === 2
            ? `<span class="mode-hint">📅 ${meta.hint}</span><input class="form-input" value="${mt.date_sort_keyword || ''}" placeholder="日期关键词（如：施工日期）" onchange="updateMergeTask(${i}, 'date_sort_keyword', this.value)" style="max-width:200px;">`
            : `<span class="mode-hint">💡 ${meta.hint}</span>`;

        return `
        <div class="merge-row" style="background:var(--bg-card);border:1px solid var(--border-light);border-radius:var(--radius-sm);padding:14px 16px;display:flex;flex-direction:column;gap:10px;">
            <div class="merge-row-top">
                <span style="font-size:13px;font-weight:600;color:var(--accent);white-space:nowrap;">合并 #${i + 1}</span>
                <select class="form-select" onchange="onMergeSubdirChange(${i}, this.value)" style="max-width:240px;">
                    <option value="">-- 选择输出目录 --</option>
                    ${dirs.map(d => `<option value="${d.value}" ${mt.input_subdir === d.value ? 'selected' : ''}>${d.label}</option>`).join('')}
                </select>
                <input class="form-input" value="${mt.output_file}" placeholder="合并后文件名" onchange="updateMergeTask(${i}, 'output_file', this.value)" style="max-width:220px;">
                <button class="btn btn-sm btn-ghost" onclick="removeMergeRow(${i})" style="color:var(--danger);font-size:16px;" title="删除">✕</button>
            </div>
            <div class="merge-row-mid">
                <label style="font-size:12px;color:var(--text-muted);white-space:nowrap;">排序模式：</label>
                <select class="form-select" onchange="updateMergeTask(${i}, 'sort_mode', parseInt(this.value))" style="max-width:200px;">
                    ${[1,2,3].map(v => `<option value="${v}" ${mode === v ? 'selected' : ''}>${v}. ${MERGE_MODE_META[v].name}</option>`).join('')}
                </select>
            </div>
            <div class="merge-mode-config">${configHtml}</div>
            <div class="merge-row-actions">
                <button class="btn btn-sm btn-outline" onclick="previewMergeOrder(${i})" ${!hasDir ? 'disabled' : ''} style="font-size:12px;">🔍 预览排序</button>
            </div>
        </div>`;
    }).join('');
}

function onMergeSubdirChange(index, value) {
    STATE.mergeTasks[index].input_subdir = value;
    if (value && (!STATE.mergeTasks[index].output_file || STATE.mergeTasks[index].output_file === '合并文档.docx')) {
        const baseName = value === '__root__' ? 'output' : value.split('/').pop();
        STATE.mergeTasks[index].output_file = `${baseName}_合并.docx`;
    }
    renderMergeTasks();
}

// ============================================================
//  合并排序预览
// ============================================================
async function previewMergeOrder(mergeIndex) {
    const mt = STATE.mergeTasks[mergeIndex];
    if (!mt || !mt.input_subdir) { showToast('请先选择输出目录', 'error'); return; }
    const mode = mt.sort_mode || 3;
    const meta = MERGE_MODE_META[mode];

    const inner = document.getElementById('loadPresetModalInner');
    inner.classList.add('modal-wide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = `排序预览 — ${meta.name}`;
    document.getElementById('loadPresetBody').innerHTML = '<p style="text-align:center;padding:30px;color:var(--text-muted);">正在分析...</p>';
    document.getElementById('loadPresetModal').style.display = 'flex';

    try {
        const res = await fetch('/api/merge/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: STATE.sessionId, input_subdir: mt.input_subdir, sort_mode: mode, date_sort_keyword: mt.date_sort_keyword || '' })
        });
        const data = await res.json();

        if (data.error) {
            document.getElementById('loadPresetBody').innerHTML = `<div style="padding:20px;text-align:center;color:var(--danger);"><p>${data.error}</p></div><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
            return;
        }

        const files = data.files || [];
        document.getElementById('loadPresetBody').innerHTML = `
            <div style="margin-bottom:8px;font-size:12px;color:var(--text-secondary);">排序模式：<strong>${meta.name}</strong>${mode === 2 && mt.date_sort_keyword ? ` · 关键词："${mt.date_sort_keyword}"` : ''}</div>
            ${files.length === 0 ? '<p class="text-muted">没有 docx 文件</p>' : `
            <div style="max-height:50vh;overflow-y:auto;">
                <table class="sort-preview-table"><thead><tr><th style="width:50px;">#</th><th>文件名</th><th>排序依据</th><th>详情</th></tr></thead>
                <tbody>${files.map(f => `<tr><td class="sort-rank">${f.sort_rank}</td><td class="sort-filename">${f.filename}</td><td class="${f.sort_key_display === '未识别到编号' || f.sort_key_display === '未识别到日期' || f.sort_key_display === '—' ? 'sort-key-missing' : 'sort-key'}">${f.sort_key_display}</td><td style="font-size:11px;color:var(--text-muted);">${f.sort_key_detail}</td></tr>`).join('')}</tbody></table>
            </div>
            <div class="preview-summary"><span>📁 ${mt.input_subdir}</span><span>共 <strong>${data.total}</strong> 份</span></div>`}
            <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:12px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>
        `;
    } catch (e) {
        document.getElementById('loadPresetBody').innerHTML = `<p style="color:var(--danger);text-align:center;padding:20px;">预览失败：${e.message}</p><div style="display:flex;justify-content:flex-end;gap:8px;"><button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button></div>`;
    }
}

// ============================================================
//  执行生成
// ============================================================
async function startGenerate() {
    if (STATE.tasks.length === 0) { showToast('请先添加任务', 'error'); return; }

    for (let i = 0; i < STATE.tasks.length; i++) {
        const t = STATE.tasks[i];
        if (!t.template_path) { showToast(`任务 ${i + 1}: 请选择 Word 模板`, 'error'); return; }
        if (!t.excel_path) { showToast(`任务 ${i + 1}: 请选择 Excel`, 'error'); return; }
        if (!t.sheet_names || t.sheet_names.length === 0) { showToast(`任务 ${i + 1}: 请选择工作表`, 'error'); return; }
        if (!t.filename_column) { showToast(`任务 ${i + 1}: 请选择命名字段`, 'error'); return; }
        if (t.row_filter !== undefined && t.row_filter !== null && t.row_filter.length === 0) {
            showToast(`任务 ${i + 1}: 已选择 0 行，请至少选择一行`, 'error'); return;
        }
    }

    const enableMerge = document.getElementById('enableMerge').checked;
    const payload = { session_id: STATE.sessionId, tasks: STATE.tasks, merge_tasks: enableMerge ? STATE.mergeTasks : [] };

    const btn = document.getElementById('btnRun');
    btn.disabled = true;
    btn.innerHTML = '生成中...';
    document.getElementById('progressContainer').style.display = 'block';
    document.getElementById('resultSummary').style.display = 'none';
    document.getElementById('progressLog').innerHTML = '';
    document.getElementById('progressTitle').textContent = '正在准备...';
    document.getElementById('progressText').textContent = '';
    document.getElementById('progressFill').style.width = '0%';

    try {
        const res = await fetch('/api/generate', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        const data = await res.json();
        if (data.error) { showToast(data.error, 'error'); resetRunButton(btn); return; }
        STATE.currentTaskId = data.task_id;
        pollProgress(data.task_id, btn);
    } catch (e) { showToast('请求失败: ' + e.message, 'error'); resetRunButton(btn); }
}

function resetRunButton(btn) {
    btn.disabled = false;
    btn.innerHTML = `<svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg>开始生成`;
}

function pollProgress(taskId, btn) {
    let pollCount = 0;
    let consecutiveErrors = 0;
    const MAX_ERRORS = 10;  // 连续 10 次失败后停止
    const interval = setInterval(async () => {
        pollCount++;
        try {
            const res = await fetch(`/api/generate/progress/${taskId}`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            consecutiveErrors = 0;
            const data = await res.json();
            const p = data.progress || {};
            const total = p.total || 0;
            const current = p.current || 0;
            const pct = total > 0 ? Math.round((current / total) * 100) : 0;

            document.getElementById('progressTitle').textContent = data.status === 'running' ? '生成中...' : (data.status === 'completed' ? '已完成' : '出错');
            document.getElementById('progressText').textContent = `${current} / ${total}`;
            document.getElementById('progressFill').style.width = `${pct}%`;

            const log = document.getElementById('progressLog');
            const logEntries = data.progress_log || [];
            if (logEntries.length > 0) {
                log.innerHTML = logEntries.map(entry => {
                    const isError = (entry.status || '').includes('error');
                    return `<div class="${isError ? 'log-error' : 'log-success'}">[${entry.current}/${entry.total}] ${isError ? '✗' : '✓'} ${entry.filename}</div>`;
                }).join('');
                log.scrollTop = log.scrollHeight;
            }

            if (data.status === 'completed' || data.status === 'error') {
                clearInterval(interval);
                const summary = document.getElementById('resultSummary');
                summary.style.display = 'block';

                if (data.status === 'completed' && data.result) {
                    const r = data.result;
                    summary.className = 'result-summary';
                    let mergeHtml = '';
                    if (r.merge_results && r.merge_results.length > 0) {
                        const successCount = r.merge_results.filter(m => m.status === 'success').length;
                        const errorCount = r.merge_results.filter(m => m.status === 'error').length;
                        mergeHtml = `<div style="margin-top:12px;padding:10px;background:var(--bg-card);border-radius:6px;text-align:left;">
                            <p style="font-weight:600;margin-bottom:8px;">📦 合并：成功 ${successCount} / ${r.merge_results.length} ${errorCount > 0 ? `（${errorCount} 失败）` : ''}</p>
                            ${r.merge_results.map((m, idx) => `<div style="font-size:12px;padding:4px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;"><span>#${idx + 1} ${m.input_subdir || '(未指定)'}</span><span style="color:${m.status === 'success' ? 'var(--success)' : 'var(--danger)'};">${m.status === 'success' ? `✅ 已合并 ${m.count} 份 → ${(m.output_file || '').split(/[\\\\/]/).pop()}` : `❌ ${m.error || '失败'}`}</span></div>`).join('')}
                        </div>`;
                    }
                    summary.innerHTML = `<h3>✅ 生成完成</h3><p>成功 ${r.total_success} / ${r.total_count} 条，耗时 ${r.elapsed} 秒</p>${mergeHtml}`;
                } else {
                    summary.className = 'result-summary error';
                    summary.innerHTML = `<h3>❌ 生成出错</h3><p>${(data.result && data.result.error) || '未知错误'}</p><div style="margin-top:12px;"><button class="btn btn-primary" onclick="startGenerate()">🔄 重试</button></div>`;
                }
                await refreshOutputs();
                resetRunButton(btn);
            }
        } catch (e) {
            consecutiveErrors++;
            if (consecutiveErrors >= MAX_ERRORS) {
                clearInterval(interval);
                showToast('与服务器的连接已断开，请检查服务是否运行', 'error');
                document.getElementById('progressTitle').textContent = '连接断开';
                resetRunButton(btn);
            }
        }
    }, 500);
}

// ============================================================
//  结果下载
// ============================================================
async function refreshOutputs() {
    if (!STATE.sessionId) return;
    try {
        const res = await fetch(`/api/outputs?session_id=${STATE.sessionId}`);
        const data = await res.json();
        renderOutputTree(data.files);
    } catch (e) { console.error('获取输出失败', e); }
}

function renderOutputTree(files) {
    const section = document.getElementById('section-download');
    const tree = document.getElementById('outputTree');
    if (!files || files.length === 0) { section.style.display = 'none'; return; }
    section.style.display = 'block';

    const zipAllBtn = `<div style="margin-bottom:12px;"><button class="btn btn-primary btn-sm" onclick="downloadAllZip()">📦 打包下载全部</button></div>`;

    function renderItems(items) {
        return items.map(item => {
            if (item.type === 'folder') {
                return `<div class="output-folder">
                    <div class="output-folder-name" onclick="toggleFolder(this)">
                        <span class="folder-icon">📁</span> ${item.name}
                        <span style="font-size:11px;color:var(--text-muted);margin-left:4px;">(${item.children ? item.children.length : 0} 个文件)</span>
                        <button class="btn btn-sm btn-outline" onclick="event.stopPropagation();downloadFolderZip('${item.rel_path}')" style="margin-left:8px;padding:2px 8px;">📦 打包</button>
                    </div>
                    <div class="output-files" style="display:block;">${item.children ? renderItems(item.children) : ''}</div>
                </div>`;
            } else {
                return `<div class="output-file">
                    <a class="output-file-name" href="/api/download?path=${encodeURIComponent(item.path)}" download><span class="file-icon">📄</span><span>${item.name}</span></a>
                    <span class="output-file-size">${formatSize(item.size)}</span>
                </div>`;
            }
        }).join('');
    }
    tree.innerHTML = zipAllBtn + renderItems(files);
}

function downloadAllZip() { window.location.href = `/api/download/zip?session_id=${STATE.sessionId}`; }
function downloadFolderZip(relPath) { window.location.href = `/api/download/zip?session_id=${STATE.sessionId}&subdir=${encodeURIComponent(relPath)}`; }
function toggleFolder(el) {
    const filesDiv = el.nextElementSibling;
    if (filesDiv) filesDiv.style.display = filesDiv.style.display === 'none' ? 'block' : 'none';
}

// ============================================================
//  预设管理
// ============================================================
async function refreshPresets() {
    try {
        const res = await fetch('/api/presets');
        const data = await res.json();
        renderPresetList(data.presets);
    } catch (e) { console.error('获取预设失败', e); }
}

function renderPresetList(presets) {
    const list = document.getElementById('presetList');
    if (!presets || presets.length === 0) { list.innerHTML = '<p class="text-muted">暂无保存的预设方案</p>'; return; }
    list.innerHTML = presets.map(p => `
        <div class="preset-item">
            <div class="preset-info">
                <h4>${p.name}</h4>
                <p>${p.task_count} 个任务 · ${new Date(p.created).toLocaleString('zh-CN')}</p>
            </div>
            <div class="preset-actions">
                <button class="btn btn-sm btn-outline" onclick="applyPreset('${p.id}')">应用</button>
                <button class="btn btn-sm btn-outline" onclick="deletePresetConfirm('${p.id}')" style="color:var(--danger);">删除</button>
            </div>
        </div>
    `).join('') + `<div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border-light);display:flex;gap:8px;">
        <button class="btn btn-primary" onclick="savePresetFromManager()">💾 保存当前配置</button>
        <button class="btn btn-outline" onclick="closePresetManager()">关闭</button>
    </div>`;
}

async function savePreset() {
    const name = prompt('请输入预设方案名称:');
    if (!name) return;
    try {
        const res = await fetch('/api/presets/save', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, tasks: STATE.tasks, merge_tasks: STATE.mergeTasks })
        });
        const data = await res.json();
        if (data.status === 'ok') { showToast(`预设「${name}」已保存`, 'success'); await refreshPresets(); }
        else showToast(data.error, 'error');
    } catch (e) { showToast('保存失败', 'error'); }
}

async function applyPreset(presetId) {
    try {
        const res = await fetch(`/api/presets/load/${presetId}`);
        const data = await res.json();
        STATE.tasks = data.tasks || [];
        STATE.mergeTasks = data.merge_tasks || [];
        STATE.templatePlaceholders = {};
        STATE.excelColumns = {};
        STATE.excelSheets = {};
        STATE.tasks.forEach((t, i) => {
            if (!t.sheet_names) t.sheet_names = t.sheet_name ? [t.sheet_name] : [];
        });
        document.getElementById('enableMerge').checked = STATE.mergeTasks.length > 0;
        toggleMergeSettings();

        for (let i = 0; i < STATE.tasks.length; i++) {
            if (STATE.tasks[i].template_path) {
                try {
                    const r = await fetch('/api/template/placeholders', {
                        method: 'POST', headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ path: STATE.tasks[i].template_path })
                    });
                    const d = await r.json();
                    if (d.placeholders) STATE.templatePlaceholders[i] = d.placeholders;
                } catch (e) {}
            }
            if (STATE.tasks[i].excel_path) {
                try {
                    const r = await fetch('/api/excel/sheets', {
                        method: 'POST', headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ path: STATE.tasks[i].excel_path })
                    });
                    const d = await r.json();
                    if (d.sheets) STATE.excelSheets[i] = d.sheets;
                } catch (e) {}
                await fetchColumns(i);
            }
        }
        renderTasks();
        closeLoadPreset();
        showToast(`已加载预设「${data.name}」`, 'success');
    } catch (e) { showToast('加载失败', 'error'); }
}

async function deletePresetConfirm(presetId) {
    if (!confirm('确定删除？')) return;
    try {
        await fetch(`/api/presets/delete/${presetId}`, { method: 'DELETE' });
        showToast('预设已删除', 'info');
        await refreshPresets();
    } catch (e) { showToast('删除失败', 'error'); }
}

function showPresetManager() { document.getElementById('presetModal').style.display = 'flex'; refreshPresets(); }
function closePresetManager() { document.getElementById('presetModal').style.display = 'none'; }

function loadPresetPrompt() {
    document.getElementById('loadPresetModalInner').classList.remove('modal-wide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = '加载预设方案';
    document.getElementById('loadPresetModal').style.display = 'flex';
    loadPresetOptions();
}

function closeLoadPreset() {
    document.getElementById('loadPresetModal').style.display = 'none';
    const inner = document.getElementById('loadPresetModalInner');
    inner.classList.remove('modal-wide');
    inner.classList.remove('modal-xwide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = '加载预设方案';
    window._rowSelectorData = null;
    window._excelPreviewState = null;
}

async function loadPresetOptions() {
    try {
        const res = await fetch('/api/presets');
        const data = await res.json();
        const body = document.getElementById('loadPresetBody');
        if (!data.presets || data.presets.length === 0) { body.innerHTML = '<p class="text-muted">暂无预设</p>'; return; }
        body.innerHTML = `<p style="margin-bottom:12px;color:var(--text-secondary);">选择一个预设方案：</p>${data.presets.map(p => `
            <div class="preset-item" style="cursor:pointer;" onclick="applyPreset('${p.id}')">
                <div class="preset-info"><h4>${p.name}</h4><p>${p.task_count} 个任务 · ${new Date(p.created).toLocaleString('zh-CN')}</p></div>
                <button class="btn btn-sm btn-primary">加载</button>
            </div>`).join('')}`;
    } catch (e) { console.error('加载预设失败', e); }
}

function savePresetFromManager() { closePresetManager(); savePreset(); }

// 点击弹窗外部关闭
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal-overlay')) {
        e.target.style.display = 'none';
        if (e.target.id === 'templateWorkbenchModal') saveWorkbenchState();
    }
});

// Esc 键关闭所有弹窗
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        document.querySelectorAll('.modal-overlay').forEach(m => {
            if (m.style.display !== 'none') {
                m.style.display = 'none';
                if (m.id === 'templateWorkbenchModal') saveWorkbenchState();
                if (m.id === 'loadPresetModal') { window._rowSelectorData = null; }
            }
        });
    }
});

// ============================================================
//  模板工作台
// ============================================================
const WORKBENCH_ALIASES_KEY = 'workbench_aliases';

function getWorkbenchAliases() {
    try { return JSON.parse(localStorage.getItem(WORKBENCH_ALIASES_KEY) || '{}'); } catch (e) { return {}; }
}
function saveWorkbenchAliases(aliases) { localStorage.setItem(WORKBENCH_ALIASES_KEY, JSON.stringify(aliases)); }

async function showTemplateWorkbench(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.template_path) { showToast('请先选择 Word 模板', 'error'); return; }

    let placeholders = STATE.templatePlaceholders[taskIndex];
    if (!placeholders) {
        try {
            const res = await fetch('/api/template/placeholders', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ path: task.template_path })
            });
            const data = await res.json();
            if (data.placeholders) { STATE.templatePlaceholders[taskIndex] = data.placeholders; placeholders = data.placeholders; }
            else if (data.error) { showToast(data.error, 'error'); return; }
        } catch (e) { showToast('获取占位符失败', 'error'); return; }
    }

    if (!placeholders || placeholders.length === 0) { showToast('模板中没有找到占位符', 'error'); return; }

    const savedAliases = getWorkbenchAliases();
    window._workbenchData = { taskIndex, placeholders, selectedFields: new Set(placeholders), aliases: {} };
    placeholders.forEach(p => { if (savedAliases[p]) window._workbenchData.aliases[p] = savedAliases[p]; });

    try {
        const savedSelected = JSON.parse(localStorage.getItem(`workbench_selected_${taskIndex}`) || 'null');
        if (savedSelected && Array.isArray(savedSelected)) {
            window._workbenchData.selectedFields = new Set(savedSelected.filter(p => placeholders.includes(p)));
        }
    } catch (e) {}

    renderWorkbenchBody();
    document.getElementById('templateWorkbenchModal').style.display = 'flex';
}

function renderWorkbenchBody() {
    const d = window._workbenchData;
    if (!d) return;
    const checkedCount = d.selectedFields.size;
    document.getElementById('templateWorkbenchBody').innerHTML = `
        <div class="workbench-toolbar">
            <button class="btn btn-sm btn-outline" onclick="workbenchSelectAll()">✅ 全选</button>
            <button class="btn btn-sm btn-outline" onclick="workbenchDeselectAll()">❎ 取消全选</button>
            <button class="btn btn-sm btn-outline" onclick="workbenchInvert()">🔄 反选</button>
            <span class="toolbar-info">已选 <strong>${checkedCount}</strong> / ${d.placeholders.length} 个字段</span>
        </div>
        <div class="workbench-grid">
            ${d.placeholders.map(p => {
                const selected = d.selectedFields.has(p);
                const alias = d.aliases[p] || '';
                return `<div class="workbench-field ${selected ? 'selected' : ''}" id="wbf_${p}">
                    <input type="checkbox" ${selected ? 'checked' : ''} onchange="workbenchToggleField('${p}', this.checked)">
                    <div class="field-info">
                        <div class="field-name">{{ ${p} }}</div>
                        <input class="field-alias-input" value="${alias}" placeholder="别名（用于Excel列名）" onchange="workbenchSetAlias('${p}', this.value)">
                    </div>
                </div>`;
            }).join('')}
        </div>
        <div class="workbench-download-section">
            <p class="download-hint">💡 选择需要的字段，设置别名（可选），下载 Excel 模板。</p>
            <button class="btn btn-primary" onclick="downloadExcelTemplate()" ${checkedCount === 0 ? 'disabled' : ''}>📥 下载 Excel 模板</button>
        </div>
    `;
}

function workbenchToggleField(fieldName, checked) {
    const d = window._workbenchData; if (!d) return;
    if (checked) { d.selectedFields.add(fieldName); document.getElementById(`wbf_${fieldName}`).classList.add('selected'); }
    else { d.selectedFields.delete(fieldName); document.getElementById(`wbf_${fieldName}`).classList.remove('selected'); }
    renderWorkbenchBody();
}
function workbenchSetAlias(fieldName, alias) {
    const d = window._workbenchData; if (!d) return;
    if (alias.trim()) d.aliases[fieldName] = alias.trim(); else delete d.aliases[fieldName];
}
function workbenchSelectAll() { const d = window._workbenchData; if (!d) return; d.placeholders.forEach(p => d.selectedFields.add(p)); renderWorkbenchBody(); }
function workbenchDeselectAll() { const d = window._workbenchData; if (!d) return; d.selectedFields.clear(); renderWorkbenchBody(); }
function workbenchInvert() {
    const d = window._workbenchData; if (!d) return;
    d.placeholders.forEach(p => { if (d.selectedFields.has(p)) d.selectedFields.delete(p); else d.selectedFields.add(p); });
    renderWorkbenchBody();
}
function saveWorkbenchState() {
    const d = window._workbenchData; if (!d) return;
    localStorage.setItem(`workbench_selected_${d.taskIndex}`, JSON.stringify(Array.from(d.selectedFields)));
    const globalAliases = getWorkbenchAliases();
    Object.entries(d.aliases).forEach(([k, v]) => { if (v) globalAliases[k] = v; });
    saveWorkbenchAliases(globalAliases);
}
function closeTemplateWorkbench() { saveWorkbenchState(); document.getElementById('templateWorkbenchModal').style.display = 'none'; window._workbenchData = null; }

async function downloadExcelTemplate() {
    const d = window._workbenchData; if (!d) return;
    const selectedFields = Array.from(d.selectedFields);
    if (selectedFields.length === 0) { showToast('请至少选择一个字段', 'error'); return; }
    const task = STATE.tasks[d.taskIndex];
    const aliases = {};
    selectedFields.forEach(f => { if (d.aliases[f]) aliases[f] = d.aliases[f]; });

    try {
        const res = await fetch('/api/template/export-excel', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ template_path: task.template_path, placeholders: d.placeholders, aliases: aliases, include_fields: selectedFields })
        });
        if (!res.ok) { const err = await res.json(); showToast(err.error || '导出失败', 'error'); return; }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const disposition = res.headers.get('Content-Disposition') || '';
        const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
        a.download = match ? match[1].replace(/['"]/g, '') : 'excel_template.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        saveWorkbenchState();
        showToast(`Excel 模板已下载（${selectedFields.length} 个字段）`, 'success');
    } catch (e) { showToast('下载失败: ' + e.message, 'error'); }
}

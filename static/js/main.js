/**
 * ================================================================================
 * 检验批批量生成平台 — 前端交互逻辑 (通用版 v1.0)
 * ================================================================================
 *
 * 工作流程（四步）:
 *   Step 1: 上传文件   — 拖拽/选择 Word 模板（.docx）和 Excel 数据（.xlsx/.xls）
 *   Step 2: 任务配置   — 选模板、选 Excel、勾 Sheet、选命名字段、设整数列
 *   Step 3: 执行生成   — 一键运行，后台线程 + 前端轮询，实时显示进度和日志
 *   Step 4: 结果下载   — 浏览输出文件夹树，单文件下载 / ZIP 打包下载
 *
 * 核心设计:
 *   - STATE 全局对象管理所有状态：文件列表、任务配置、缓存数据
 *   - 渲染函数以 render* 命名：renderFileList, renderTasks, renderOutputTree...
 *   - 所有 API 调用通过 fetch()，返回 JSON
 *   - socket.io 不需要，用轮询 /api/generate/progress/<task_id> 实现进度
 *
 * 文件上传区: 支持拖拽、点击选择，上传后自动刷新列表
 * 任务配置区: 每行一个任务，可增删，sheet 和列名选完文件后自动填充
 * 合并设置区: 下拉选择输出目录，多组合并互不干扰
 * 结果下载区: 树形结构展示，设文件夹折叠、文件下载、打包下载按钮
 *
 * ================================================================================
 */

// ============================================================
//  全局状态 (STATE)
//  ────────────────────────────────────
//  sessionId              — 会话 ID，用于文件隔离
//  files[]                — 已上传文件列表 {name, type, path, ...}
//  tasks[]                — 生成任务配置 [{template_path, excel_path, sheet_names, ...}]
//  mergeTasks[]           — 合并任务 [{input_subdir, output_file, sort_mode, ...}]
//  currentTaskId          — 正在运行的任务 ID
//  templatePlaceholders{} — 缓存: {taskIndex: [占位符列表]}
//  excelColumns{}         — 缓存: {taskIndex: [列名列表]}
//  excelSheets{}          — 缓存: {taskIndex: [Sheet名列表]}
// ============================================================

const STATE = {
    sessionId: null,
    files: [],
    tasks: [],
    mergeTasks: [],
    currentTaskId: null,
    // 缓存：每个任务的模板占位符、Excel 列名、Sheet 列表
    templatePlaceholders: {},  // { taskIndex: ['占位符1', '占位符2', ...] }
    excelColumns: {},          // { taskIndex: ['列名1', '列名2', ...] }
    excelSheets: {}            // { taskIndex: ['Sheet1', 'Sheet2', ...] }
};

// ============================================================
//  Step 1 相关: 文件上传
//  ────────────────────────────────────
//  支持拖拽上传和点击选择，上传后自动刷新文件列表
//  文件列表渲染在左侧面板，用色块区分 Word(蓝) / Excel(绿)
// ============================================================

// ============================================================
//  初始化 & 会话管理
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
    const port = window.location.port || '5000';
    el.textContent = `局域网 ${host}:${port}`;
}

// ============================================================
//  Toast 消息
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
    if (!STATE.sessionId) { showToast('会话未初始化，请刷新页面', 'error'); return; }
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
        } else {
            showToast(data.error || '上传失败', 'error');
        }
    } catch (e) {
        showToast('上传失败: ' + e.message, 'error');
    }
}

async function refreshFiles() {
    if (!STATE.sessionId) return;
    try {
        const res = await fetch(`/api/files?session_id=${STATE.sessionId}`);
        const data = await res.json();
        STATE.files = data.files;

        // 检测被删除的文件，自动清除任务中的引用
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
        if (tasksChanged) {
            showToast('已自动清除被删文件关联的任务配置', 'info');
        }

        renderFileList();
        // 文件列表变化后，刷新任务配置中的下拉选项
        if (STATE.tasks.length > 0) renderTasks();
    } catch (e) {
        console.error('刷新文件列表失败', e);
    }
}

function renderFileList() {
    const container = document.getElementById('fileList');
    const grid = document.getElementById('fileListGrid');
    if (STATE.files.length === 0) { container.style.display = 'none'; return; }
    container.style.display = 'block';
    grid.innerHTML = STATE.files.map(f => `
        <div class="file-item">
            <div class="file-item-icon ${f.type}">${f.type === 'word' ? '📄' : '📊'}</div>
            <div class="file-item-info">
                <div class="file-item-name" title="${f.original_name}">${f.original_name}</div>
                <div class="file-item-size">${formatSize(f.size)}</div>
            </div>
            <button class="file-item-del" onclick="deleteFile('${f.name}')" title="删除">&times;</button>
        </div>
    `).join('');
}

async function deleteFile(filename) {
    if (!confirm('确定删除这个文件吗？')) return;
    try {
        await fetch('/api/files/delete', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_id: STATE.sessionId, filename })
        });
        await refreshFiles();
        showToast('文件已删除', 'info');
    } catch (e) { showToast('删除失败', 'error'); }
}

function formatSize(bytes) {
    if (!bytes) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    let i = 0, size = bytes;
    while (size >= 1024 && i < units.length - 1) { size /= 1024; i++; }
    return size.toFixed(i === 0 ? 0 : 1) + ' ' + units[i];
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
        <div class="task-row">
            <div class="task-row-header">
                <span class="task-row-title">
                    <span class="task-row-num">${i + 1}</span>
                    任务 ${i + 1}
                </span>
                <div style="display:flex;gap:6px;">
                    <button class="btn btn-sm btn-outline" onclick="checkFieldMatch(${i})" title="对比模板占位符与Excel列名">字段匹配</button>
                    <button class="btn btn-sm btn-outline" onclick="showDataPreview(${i})" title="预览Excel前5行数据">数据预览</button>
                    <button class="btn btn-sm btn-ghost" onclick="removeTaskRow(${i})" style="color:var(--danger);">删除</button>
                </div>
            </div>
            <div class="task-row-body">
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
                            ? sheets.map(s => `
                                <label class="checkbox-label sheet-checkbox">
                                    <input type="checkbox" value="${s}" ${selectedSheets.includes(s) ? 'checked' : ''}
                                        onchange="onSheetCheckboxChange(${i})">
                                    <span>${s}</span>
                                </label>
                            `).join('')
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
                    <input class="form-input" value="${task.output_subdir}" placeholder="如：AB、隐蔽工程..."
                        onchange="updateTask(${i}, 'output_subdir', this.value)">
                </div>
                <div class="form-group">
                    <label class="form-label">日期格式</label>
                    <input class="form-input" value="${task.date_format}" placeholder="%Y年%m月%d日"
                        onchange="updateTask(${i}, 'date_format', this.value)">
                </div>
                <div class="form-group">
                    <label class="form-label">强制整数列 (逗号分隔)</label>
                    <input class="form-input" value="${(task.int_columns || []).join(', ')}" placeholder="如：根设AB, 间距"
                        onchange="updateTask(${i}, 'int_columns', this.value.split(',').map(s => s.trim()).filter(Boolean))">
                </div>
                <div class="form-group">
                    <label class="checkbox-label" style="margin-top:18px;">
                        <input type="checkbox" ${task.add_index_prefix ? 'checked' : ''}
                            onchange="updateTask(${i}, 'add_index_prefix', this.checked)">
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
        </div>`;
    }).join('');

    renderMergeTasks();
}

// ============================================================
//  模板占位符 & Excel 列名 联动
// ============================================================

async function onTemplateChange(taskIndex, templatePath) {
    updateTask(taskIndex, 'template_path', templatePath);
    if (!templatePath) {
        delete STATE.templatePlaceholders[taskIndex];
        renderTasks();
        return;
    }
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
            // 缓存 sheet 列表
            STATE.excelSheets[taskIndex] = data.sheets;
            // 默认选中第一个 sheet
            const task = STATE.tasks[taskIndex];
            task.sheet_names = [data.sheets[0]];
            task.sheet_name = data.sheets[0];
            showToast(`Excel 中发现 ${data.sheets.length} 个工作表`, 'success');
            // 获取列名
            await fetchColumns(taskIndex);
        } else if (data.error) {
            showToast(`读取 Excel 失败: ${data.error}`, 'error');
        }
    } catch (e) {
        showToast('获取工作表失败: ' + e.message, 'error');
    }
    renderTasks();
}

async function onSheetCheckboxChange(taskIndex) {
    const checkboxes = document.querySelectorAll(`#sheetList${taskIndex} input[type="checkbox"]:checked`);
    const selectedSheets = Array.from(checkboxes).map(cb => cb.value);
    STATE.tasks[taskIndex].sheet_names = selectedSheets;
    STATE.tasks[taskIndex].sheet_name = selectedSheets[0] || '';
    // 列名可能因 sheet 不同而变化，重新获取并刷新界面
    if (selectedSheets.length > 0) {
        await fetchColumns(taskIndex);
        renderTasks();
    } else {
        delete STATE.excelColumns[taskIndex];
        renderTasks();
    }
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
        if (data.columns) {
            STATE.excelColumns[taskIndex] = data.columns.map(String);
        }
    } catch (e) { console.error('获取列名失败', e); }
}

function togglePlaceholders(taskIndex) {
    const tags = document.getElementById(`phTags${taskIndex}`);
    const btn = document.getElementById(`phToggle${taskIndex}`);
    if (tags.style.display === 'none') {
        tags.style.display = 'flex';
        btn.textContent = '收起';
    } else {
        tags.style.display = 'none';
        btn.textContent = '展开';
    }
}

// ============================================================
//  字段匹配检查
// ============================================================

async function checkFieldMatch(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.template_path || !task.excel_path) {
        showToast('请先选择模板和Excel文件', 'error');
        return;
    }

    try {
        const res = await fetch('/api/match-fields', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                template_path: task.template_path,
                excel_path: task.excel_path,
                sheet_name: task.sheet_name
            })
        });
        const data = await res.json();

        const inner = document.getElementById('loadPresetModalInner');
        inner.classList.add('modal-wide');
        document.querySelector('#loadPresetModal .modal-header h3').textContent = `字段匹配检查 — 任务 ${taskIndex + 1}`;
        const body = document.getElementById('loadPresetBody');
        body.innerHTML = `
            <div class="field-match-result">
                <div class="match-section match-ok">
                    <h4>已匹配 (${data.matched.length})</h4>
                    <div class="match-tags">
                        ${data.matched.length > 0
                            ? data.matched.map(m => `<span class="ph-tag ph-matched">${m}</span>`).join('')
                            : '<span class="text-muted">无</span>'}
                    </div>
                </div>
                <div class="match-section match-missing">
                    <h4>模板有但Excel缺 (${data.missing_in_excel.length})</h4>
                    <div class="match-tags">
                        ${data.missing_in_excel.length > 0
                            ? data.missing_in_excel.map(m => `<span class="ph-tag ph-missing">${m}</span>`).join('')
                            : '<span class="text-muted">无缺失，全部匹配</span>'}
                    </div>
                    <p class="match-hint">这些占位符在Excel中找不到对应列，生成时将为空值</p>
                </div>
                <div class="match-section match-extra">
                    <h4>Excel有但模板没用 (${data.extra_in_excel.length})</h4>
                    <div class="match-tags">
                        ${data.extra_in_excel.length > 0
                            ? data.extra_in_excel.map(m => `<span class="ph-tag ph-extra">${m}</span>`).join('')
                            : '<span class="text-muted">无</span>'}
                    </div>
                    <p class="match-hint">这些列不会被使用（不影响生成）</p>
                </div>
            </div>
            <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;">
                <button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button>
            </div>
        `;
        document.getElementById('loadPresetModal').style.display = 'flex';
    } catch (e) {
        showToast('字段匹配检查失败', 'error');
    }
}

// ============================================================
//  数据预览
// ============================================================

async function showDataPreview(taskIndex) {
    const task = STATE.tasks[taskIndex];
    if (!task.excel_path) {
        showToast('请先选择Excel文件', 'error');
        return;
    }

    try {
        const res = await fetch('/api/excel/preview', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ path: task.excel_path, sheet_name: task.sheet_name, rows: 8 })
        });
        const data = await res.json();

        const inner = document.getElementById('loadPresetModalInner');
        inner.classList.add('modal-wide');
        document.querySelector('#loadPresetModal .modal-header h3').textContent = `数据预览 — 任务 ${taskIndex + 1} (${data.total_rows} 行)`;
        const body = document.getElementById('loadPresetBody');

        const cols = data.columns || [];
        const rows = data.preview || [];

        body.innerHTML = `
            <div style="overflow-x:auto;">
                <table class="preview-table">
                    <thead>
                        <tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>
                    </thead>
                    <tbody>
                        ${rows.map(r => `<tr>${cols.map(c => `<td>${r[c] !== undefined && r[c] !== null ? r[c] : ''}</td>`).join('')}</tr>`).join('')}
                    </tbody>
                </table>
            </div>
            <p style="margin-top:8px;font-size:12px;color:var(--text-muted);">仅显示前 ${rows.length} 行，共 ${data.total_rows} 行数据</p>
            <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;">
                <button class="btn btn-outline" onclick="closeLoadPreset()">关闭</button>
            </div>
        `;
        document.getElementById('loadPresetModal').style.display = 'flex';
    } catch (e) {
        showToast('数据预览失败', 'error');
    }
}

// ============================================================
//  合并任务
// ============================================================

function toggleMergeSettings() {
    const enabled = document.getElementById('enableMerge').checked;
    document.getElementById('mergeSettings').style.display = enabled ? 'block' : 'none';
    if (enabled && STATE.mergeTasks.length === 0) addMergeRow();
}

function addMergeRow() {
    STATE.mergeTasks.push({
        input_subdir: '',
        output_file: 'merged.docx',
        sort_mode: 3,
        date_sort_keyword: ''
    });
    renderMergeTasks();
}

function removeMergeRow(index) {
    STATE.mergeTasks.splice(index, 1);
    renderMergeTasks();
}

function updateMergeTask(index, field, value) {
    STATE.mergeTasks[index][field] = value;
    if (field === 'sort_mode') renderMergeTasks();
}

function getAvailableMergeDirs() {
    const dirs = [];
    const seen = new Set();
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
    const hint = dirs.length === 0 && STATE.mergeTasks.length > 0
        ? '<p class="text-muted" style="font-size:12px;margin-bottom:8px;">提示：请先在任务配置中填写「输出子目录名」，合并时才能选择对应目录</p>'
        : '';
    list.innerHTML = hint + STATE.mergeTasks.map((mt, i) => `
        <div class="merge-row">
            <span style="font-size:12px;color:var(--text-muted);white-space:nowrap;">合并 #${i + 1}</span>
            <select class="form-select" onchange="onMergeSubdirChange(${i}, this.value)" style="max-width:220px;">
                <option value="">-- 选择输出目录 --</option>
                ${dirs.map(d => `<option value="${d.value}" ${mt.input_subdir === d.value ? 'selected' : ''}>${d.label}</option>`).join('')}
            </select>
            <input class="form-input" value="${mt.output_file}" placeholder="合并后文件名"
                onchange="updateMergeTask(${i}, 'output_file', this.value)" style="max-width:180px;">
            <select class="form-select" onchange="updateMergeTask(${i}, 'sort_mode', parseInt(this.value))" style="max-width:150px;">
                <option value="3" ${mt.sort_mode === 3 ? 'selected' : ''}>按Excel原顺序</option>
                <option value="1" ${mt.sort_mode === 1 ? 'selected' : ''}>按编号排序</option>
                <option value="2" ${mt.sort_mode === 2 ? 'selected' : ''}>按日期排序</option>
            </select>
            ${mt.sort_mode === 2 ? `
                <input class="form-input" value="${mt.date_sort_keyword || ''}" placeholder="日期关键词（如：施工日期）"
                    onchange="updateMergeTask(${i}, 'date_sort_keyword', this.value)" style="max-width:180px;">
            ` : ''}
            <button class="btn btn-sm btn-ghost" onclick="removeMergeRow(${i})" style="color:var(--danger);">✕</button>
        </div>
    `).join('');
}

function onMergeSubdirChange(index, value) {
    STATE.mergeTasks[index].input_subdir = value;
    // 自动生成默认合并文件名
    if (value && (!STATE.mergeTasks[index].output_file || STATE.mergeTasks[index].output_file === 'merged.docx')) {
        const baseName = value.split('/').pop();
        STATE.mergeTasks[index].output_file = `${baseName}_合并.docx`;
    }
    renderMergeTasks();
}

// ============================================================
//  执行生成
// ============================================================

async function startGenerate() {
    if (STATE.tasks.length === 0) { showToast('请先添加至少一个生成任务', 'error'); return; }

    for (let i = 0; i < STATE.tasks.length; i++) {
        const t = STATE.tasks[i];
        if (!t.template_path) { showToast(`任务 ${i + 1}: 请选择 Word 模板`, 'error'); return; }
        if (!t.excel_path) { showToast(`任务 ${i + 1}: 请选择 Excel 数据文件`, 'error'); return; }
        if (!t.sheet_names || t.sheet_names.length === 0) { showToast(`任务 ${i + 1}: 请至少选择一个工作表`, 'error'); return; }
        if (!t.filename_column) { showToast(`任务 ${i + 1}: 请选择文件命名字段`, 'error'); return; }
    }

    const enableMerge = document.getElementById('enableMerge').checked;
    const payload = {
        session_id: STATE.sessionId,
        tasks: STATE.tasks,
        merge_tasks: enableMerge ? STATE.mergeTasks : []
    };

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
        if (data.error) {
            showToast(data.error, 'error');
            resetRunButton(btn);
            return;
        }
        STATE.currentTaskId = data.task_id;
        pollProgress(data.task_id, btn);
    } catch (e) {
        showToast('请求失败: ' + e.message, 'error');
        resetRunButton(btn);
    }
}

function resetRunButton(btn) {
    btn.disabled = false;
    btn.innerHTML = `<svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg>开始生成`;
}

function pollProgress(taskId, btn) {
    const interval = setInterval(async () => {
        try {
            const res = await fetch(`/api/generate/progress/${taskId}`);
            const data = await res.json();
            const p = data.progress || {};
            const total = p.total || 0;
            const current = p.current || 0;
            const pct = total > 0 ? Math.round((current / total) * 100) : 0;

            document.getElementById('progressTitle').textContent =
                data.status === 'running' ? '生成中...' : (data.status === 'completed' ? '已完成' : '出错');
            document.getElementById('progressText').textContent = `${current} / ${total}`;
            document.getElementById('progressFill').style.width = `${pct}%`;

            if (p.filename) {
                const log = document.getElementById('progressLog');
                const div = document.createElement('div');
                div.className = p.status === 'success' ? 'log-success' : 'log-error';
                div.textContent = `[${current}/${total}] ${p.status === 'success' ? '✓' : '✗'} ${p.filename}`;
                log.appendChild(div);
                log.scrollTop = log.scrollHeight;
            }

            if (data.status === 'completed' || data.status === 'error') {
                clearInterval(interval);
                const summary = document.getElementById('resultSummary');
                summary.style.display = 'block';

                if (data.status === 'completed' && data.result) {
                    const r = data.result;
                    summary.className = 'result-summary';
                    summary.innerHTML = `
                        <h3>✅ 生成完成</h3>
                        <p>成功 ${r.total_success} / ${r.total_count} 条，耗时 ${r.elapsed} 秒</p>
                        ${r.merge_results && r.merge_results.length > 0
                            ? `<p style="margin-top:4px;">已完成 ${r.merge_results.filter(m => m.status === 'success').length} 个合并任务</p>` : ''}
                    `;
                } else {
                    summary.className = 'result-summary error';
                    summary.innerHTML = `<h3>❌ 生成出错</h3><p>${(data.result && data.result.error) || '未知错误'}</p>`;
                }
                await refreshOutputs();
                resetRunButton(btn);
            }
        } catch (e) {
            clearInterval(interval);
            resetRunButton(btn);
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
    } catch (e) { console.error('获取输出文件失败', e); }
}

function renderOutputTree(files) {
    const section = document.getElementById('section-download');
    const tree = document.getElementById('outputTree');
    if (!files || files.length === 0) { section.style.display = 'none'; return; }
    section.style.display = 'block';

    // 顶部「打包下载全部」按钮
    const zipAllBtn = `
        <div style="margin-bottom:12px;">
            <button class="btn btn-primary btn-sm" onclick="downloadAllZip()">📦 打包下载全部</button>
        </div>`;

    function renderItems(items) {
        return items.map(item => {
            if (item.type === 'folder') {
                return `
                    <div class="output-folder">
                        <div class="output-folder-name" onclick="toggleFolder(this)">
                            <span class="folder-icon">📁</span> ${item.name}
                            <span style="font-size:11px;color:var(--text-muted);margin-left:4px;">(${item.children ? item.children.length : 0} 个文件)</span>
                            <button class="btn btn-sm btn-outline" onclick="event.stopPropagation();downloadFolderZip('${item.rel_path}')" style="margin-left:8px;padding:2px 8px;">📦 打包</button>
                        </div>
                        <div class="output-files" style="display:block;">
                            ${item.children ? renderItems(item.children) : ''}
                        </div>
                    </div>`;
            } else {
                return `
                    <div class="output-file">
                        <a class="output-file-name" href="/api/download?path=${encodeURIComponent(item.path)}" download>
                            <span class="file-icon">📄</span><span>${item.name}</span>
                        </a>
                        <span class="output-file-size">${formatSize(item.size)}</span>
                    </div>`;
            }
        }).join('');
    }
    tree.innerHTML = zipAllBtn + renderItems(files);
}

function downloadAllZip() {
    window.location.href = `/api/download/zip?session_id=${STATE.sessionId}`;
}

function downloadFolderZip(relPath) {
    window.location.href = `/api/download/zip?session_id=${STATE.sessionId}&subdir=${encodeURIComponent(relPath)}`;
}

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
    if (!presets || presets.length === 0) {
        list.innerHTML = '<p class="text-muted">暂无保存的预设方案</p>';
        return;
    }
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
    `).join('') + `
        <div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border-light);display:flex;gap:8px;">
            <button class="btn btn-primary" onclick="savePresetFromManager()">💾 保存当前配置为预设</button>
            <button class="btn btn-outline" onclick="closePresetManager()">关闭</button>
        </div>
    `;
}

async function savePreset() {
    const name = prompt('请输入预设方案名称（如：攀枝花米易项目-隐蔽接地）:');
    if (!name) return;
    try {
        const res = await fetch('/api/presets/save', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name, tasks: STATE.tasks, merge_tasks: STATE.mergeTasks })
        });
        const data = await res.json();
        if (data.status === 'ok') {
            showToast(`预设「${name}」已保存`, 'success');
            await refreshPresets();
        } else { showToast(data.error, 'error'); }
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
        // 兼容旧预设：确保每个 task 有 sheet_names
        STATE.tasks.forEach((t, i) => {
            if (!t.sheet_names) {
                t.sheet_names = t.sheet_name ? [t.sheet_name] : [];
            }
        });
        document.getElementById('enableMerge').checked = STATE.mergeTasks.length > 0;
        toggleMergeSettings();
        // 加载预设后，自动拉取已选模板和Excel的信息
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
                // 拉取 sheet 列表并缓存
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
    } catch (e) { showToast('加载预设失败', 'error'); }
}

async function deletePresetConfirm(presetId) {
    if (!confirm('确定删除这个预设方案吗？')) return;
    try {
        await fetch(`/api/presets/delete/${presetId}`, { method: 'DELETE' });
        showToast('预设已删除', 'info');
        await refreshPresets();
    } catch (e) { showToast('删除失败', 'error'); }
}

function showPresetManager() {
    document.getElementById('presetModal').style.display = 'flex';
    refreshPresets();
}

function closePresetManager() {
    document.getElementById('presetModal').style.display = 'none';
}

function loadPresetPrompt() {
    document.getElementById('loadPresetModalInner').classList.remove('modal-wide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = '加载预设方案';
    document.getElementById('loadPresetModal').style.display = 'flex';
    loadPresetOptions();
}

function closeLoadPreset() {
    document.getElementById('loadPresetModal').style.display = 'none';
    document.getElementById('loadPresetModalInner').classList.remove('modal-wide');
    document.querySelector('#loadPresetModal .modal-header h3').textContent = '加载预设方案';
}

async function loadPresetOptions() {
    try {
        const res = await fetch('/api/presets');
        const data = await res.json();
        const body = document.getElementById('loadPresetBody');
        if (!data.presets || data.presets.length === 0) {
            body.innerHTML = '<p class="text-muted">暂无保存的预设方案</p>';
            return;
        }
        body.innerHTML = `
            <p style="margin-bottom:12px;color:var(--text-secondary);">选择一个预设方案来加载：</p>
            ${data.presets.map(p => `
                <div class="preset-item" style="cursor:pointer;" onclick="applyPreset('${p.id}')">
                    <div class="preset-info">
                        <h4>${p.name}</h4>
                        <p>${p.task_count} 个任务 · ${new Date(p.created).toLocaleString('zh-CN')}</p>
                    </div>
                    <button class="btn btn-sm btn-primary">加载</button>
                </div>
            `).join('')}
        `;
    } catch (e) { console.error('加载预设选项失败', e); }
}

function savePresetFromManager() {
    closePresetManager();
    savePreset();
}

// 点击弹窗外部关闭
document.addEventListener('click', function(e) {
    if (e.target.classList.contains('modal-overlay')) {
        e.target.style.display = 'none';
    }
});

# 检验批批量生成平台 v2.1

将 Python 脚本（docxtpl + Excel → Word）改造为浏览器操作的 Web 服务，支持局域网多设备访问。

## 功能特性

- **Word 模板填充**：通过 `{{占位符}}` 语法，将 Excel 数据批量填充到 Word 模板
- **多任务编排**：支持同时配置多个生成任务（不同模板 + 不同 Excel）
- **三种合并排序**：编号+续号 / 日期关键词 / Excel 原始顺序
- **按需生成**：可选择 Excel 中的部分行进行生成，非全量
- **合并预览**：合并前可预览排序结果，避免错误合并
- **打包下载**：生成的文件可单独下载或打包为 ZIP
- **预设管理**：保存常用任务配置，下次直接加载
- **模板工作台**：提取模板占位符、生成 Excel 填写模板
- **会话隔离**：每个浏览器会话独立的上传/输出目录，互不干扰
- **跨平台**：Windows / macOS / Linux 均可运行
- **局域网访问**：其他电脑/手机通过浏览器即可使用

## 快速开始

### Windows

1. 双击 `启动平台.bat`（或 `scripts/start.bat`）
2. 首次运行会自动创建虚拟环境并安装依赖
3. 浏览器自动打开 `http://127.0.0.1:5005`

### macOS / Linux

```bash
chmod +x scripts/start.sh
./scripts/start.sh
```

### 其他电脑首次安装

- **Windows**：复制整个项目文件夹，双击 `scripts/一键安装_其他电脑.bat`
- **macOS/Linux**：运行 `scripts/一键安装_其他电脑.sh`

## 使用流程

```
上传文件 → 配置任务 → 批量生成 → 下载结果
```

1. **上传文件**：拖拽或点击上传 Word 模板（.docx）和 Excel 数据（.xlsx/.xls）
2. **配置任务**：选择模板、Excel、Sheet、文件名列，匹配字段
3. **批量生成**：点击生成，实时查看进度，支持按需选择行
4. **下载结果**：单独下载或打包 ZIP，支持多文档合并

## 项目结构

```
.
├── src/                    # 后端源码
│   ├── config.py           # 全局配置（端口、路径、常量）
│   ├── engine.py           # 核心生成引擎（模板渲染、合并、排序）
│   └── app.py              # Flask Web 服务（API 路由）
├── web/                    # 前端文件
│   ├── templates/
│   │   └── index.html      # 主页面
│   └── static/
│       ├── css/style.css   # 样式
│       └── js/main.js      # 交互逻辑
├── scripts/                # 脚本
│   ├── start.bat           # Windows 启动
│   ├── start.sh            # macOS/Linux 启动
│   ├── build_exe.bat       # PyInstaller 打包
│   ├── 一键安装_其他电脑.bat  # Windows 一键安装
│   └── 一键安装_其他电脑.sh   # macOS/Linux 一键安装
├── data/                   # 运行时数据（自动创建）
│   ├── uploads/            # 上传文件（按会话隔离）
│   ├── outputs/            # 生成结果（按会话隔离）
│   └── presets/            # 预设配置
├── 启动平台.bat             # Windows 快捷入口
├── requirements.txt        # Python 依赖
├── .gitignore
└── README.md
```

## 技术栈

| 层级 | 技术 |
|------|------|
| 后端 | Flask 3.0+ |
| 前端 | 原生 HTML/CSS/JS（无框架） |
| 模板引擎 | docxtpl (Jinja2) |
| Excel 读取 | pandas + openpyxl + xlrd |
| Word 合并 | docxcompose + python-docx |
| 打包 | PyInstaller |

## 合并排序说明

| 模式 | 说明 | 示例 |
|------|------|------|
| 模式 1 | 按文档内编号+续号 | `4-1-(5-1)` → `4-1-(5-1)(续2)` |
| 模式 2 | 按文档内日期（关键词识别） | 自动识别"施工日期"后的日期排序 |
| 模式 3 | 按 Excel 原始顺序 | 提取文件名 `001_` 前缀数字排序（默认） |

## 配置

端口、路径等配置集中在 `src/config.py`：

```python
PORT = 5005        # 服务端口
HOST = '0.0.0.0'   # 监听地址（0.0.0.0 = 所有网卡）
```

## 打包为 EXE

```bash
scripts/build_exe.bat
```

打包后 `dist/jianyanpi/` 文件夹可复制到任何 Windows 电脑直接运行，无需安装 Python。

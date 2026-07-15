# 检验批批量生成平台

基于 Python Flask 的 Web 平台，将「Word 模板 + Excel 数据 → 批量生成文档」的工作流搬到浏览器中操作。

## 功能

- **拖拽上传** Word 模板（.docx）和 Excel 数据（.xlsx / .xls）
- **智能配置** 自动提取模板占位符，下拉选择 Excel 列名作为命名字段
- **多 Sheet 支持** 一个 Excel 多个数据表可同时勾选批量生成
- **字段匹配检查** 模板占位符与 Excel 列名一一对照，缺漏一目了然
- **多任务并行** 添加多个生成任务，不同模板 + 不同 Excel 同时跑
- **文档合并** 按目录分组合并，互不干扰
- **ZIP 打包下载** 一键打包全部或单个文件夹
- **预设方案** 常用配置一键保存/加载，切换项目不用重新配
- **局域网共享** 绑定 `0.0.0.0:5000`，同网络下其他电脑也能访问
- **跨平台** Windows / macOS / Linux 均可用

## 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 启动
**Windows**：双击 `start.bat`
**macOS / Linux**：
```bash
bash start.sh
```

### 3. 使用
浏览器打开 `http://127.0.0.1:5000`（本机），或终端显示的局域网地址（其他电脑）。

## 依赖

- Flask ≥ 3.0
- pandas ≥ 2.0
- openpyxl ≥ 3.1
- xlrd ≥ 2.0
- docxtpl ≥ 0.16
- python-docx ≥ 1.0
- docxcompose ≥ 1.4

## 项目结构

```
├── app.py           # Flask 后端 API
├── engine.py        # 核心生成引擎
├── start.bat        # Windows 启动脚本
├── start.sh         # macOS/Linux 启动脚本
├── requirements.txt # Python 依赖
├── templates/
│   └── index.html   # 网页界面
├── static/
│   ├── css/style.css
│   └── js/main.js
├── uploads/         # 上传文件（自动创建）
├── outputs/         # 生成结果（自动创建）
└── presets/         # 预设方案（自动创建）
```

## License

MIT

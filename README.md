# URL 文件大小计算器

批量获取 URL 对应文件的大小，无需实际下载文件。支持从 Excel / TXT 文件中读取 URL 列表，快速检测并汇总文件总大小，结果可导出为 Excel 报告。

## 功能特点

- 📁 **多格式支持**：支持 `.xlsx`、`.xls`、`.txt` 文件导入 URL 列表
- ⚡ **并发检测**：20 线程并发请求，快速获取文件大小
- 📊 **结果汇总**：自动统计总文件数、成功/失败数量、文件总大小
- 📥 **Excel 导出**：一键下载带格式的检测结果报告
- 🔄 **智能实例管理**：自动检测已运行实例，避免重复启动
- 🌐 **Web 界面**：基于 Flask 的本地 Web 应用，支持拖拽上传

## 截图预览

启动后会自动打开浏览器，界面如下：

- 上传区域：拖拽或点击上传包含 URL 的文件
- 结果展示：表格展示每个 URL 的文件大小和检测状态
- 汇总统计：文件总大小、成功/失败数量一目了然

## 快速开始

### 环境要求

- Python 3.10+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行

```bash
python app.py
```

启动后会自动打开浏览器访问 `http://127.0.0.1:5001`。

## 使用方法

1. 准备一个包含 URL 列表的文件（`.xlsx`、`.xls` 或 `.txt`）
   - Excel 文件：URL 放在第一列
   - TXT 文件：每行一个 URL
   - URL 需以 `http://` 或 `https://` 开头
2. 在页面中上传文件（支持拖拽）
3. 点击 **"开始检测"**
4. 检测完成后查看结果，可点击 **"下载 Excel 报告"** 导出

## 打包为可执行文件

使用 PyInstaller 打包为独立可执行文件（无需 Python 环境即可运行）：

### macOS

```bash
pip install pyinstaller
pyinstaller --noconfirm --clean --console --name URLSizeChecker \
  --add-data "templates:templates" \
  --hidden-import=requests --hidden-import=flask \
  --hidden-import=openpyxl --hidden-import=xlrd \
  app.py
```

### Windows

```bash
pip install pyinstaller
pyinstaller --noconfirm --clean --console --name URLSizeChecker ^
  --add-data "templates;templates" ^
  --hidden-import=requests --hidden-import=flask ^
  --hidden-import=openpyxl --hidden-import=xlrd ^
  app.py
```

> ⚠️ 注意：macOS 使用冒号 `:` 分隔路径，Windows 使用分号 `;`。

打包后的可执行文件位于 `dist/URLSizeChecker/` 目录下。

## 项目结构

```
├── app.py                 # 主应用（Flask 后端 + 路由）
├── file_size_checker.py   # 旧版独立脚本（参考）
├── templates/
│   └── index.html         # Web 前端页面
├── requirements.txt       # Python 依赖
├── SPEC.md                # 项目规格文档
└── README.md              # 本文件
```

## 技术栈

- **后端**：Python + Flask
- **前端**：原生 HTML/CSS/JavaScript
- **HTTP 请求**：requests（带连接池复用）
- **Excel 处理**：openpyxl（.xlsx）、xlrd（.xls）
- **打包**：PyInstaller

## License

MIT

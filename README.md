# TemplateForge — Word 模板批量生成工具

根据 Word 模板和数据文件，批量生成个性化 Word 文档。

## 功能特性

- 📄 支持 Word 模板（.docx），使用 `【占位符】` 标记变量位置
- 📊 支持 Excel (.xlsx) 和 CSV (.csv) 数据源
- 📎 支持附录 Excel 文件，按条件筛选数据自动生成表格插入文档
- 🖥️ 图形化界面，无需命令行操作
- 🌍 跨平台支持 macOS 和 Windows
- 📦 可打包为独立可执行文件，无需安装 Python

## 快速开始

### 方式一：直接运行（需要 Python 3.10+）

```bash
pip install -r requirements.txt
python main.py
```

### 方式二：打包为可执行文件

**macOS:**
```bash
./build.sh
# 产物在 dist/TemplateForge
```

**Windows:**
```cmd
build.bat
REM 产物在 dist\TemplateForge.exe
```

## 使用方法

1. **选择模板文件** — Word 文档，包含 `【占位符】` 标记
2. **选择数据文件** — Excel 或 CSV，列名对应占位符名称
3. **选择附录文件**（可选） — Excel 文件，按指定列筛选后生成表格
4. **配置附录筛选**（可选） — 选择筛选列和匹配列
5. **点击生成** — 批量生成文档到输出目录

## 占位符格式

在 Word 模板中使用中文方括号标记变量：

```
甲方：欧加隆（上海）医药贸易有限公司
乙方：【经销商名称】

协议编号：【协议编号】
```

数据文件中对应的列名 `经销商名称`、`协议编号` 会自动替换为对应值。

## 项目结构

```
TemplateForge/
├── main.py              # GUI 入口
├── core/
│   ├── engine.py        # 模板引擎（占位符替换 + 附录表格插入）
│   └── data_loader.py   # 数据加载（Excel/CSV）
├── requirements.txt     # 依赖清单
├── build.sh             # macOS 打包脚本
└── build.bat            # Windows 打包脚本
```

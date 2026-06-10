# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

TemplateForge 是一个基于 Python tkinter 的桌面 GUI 工具，用于根据 Word 模板和数据文件批量生成个性化 Word 文档。

## 常用命令

```bash
# 安装依赖并运行（开发）
pip install -r requirements.txt
python main.py

# macOS 打包为独立可执行文件
./build.sh
# 产物在 dist/TemplateForge

# Windows 打包
build.bat
# 产物在 dist\TemplateForge.exe
```

项目无测试框架、无 lint 工具，纯手动验证。

## 架构

### 模块职责

- **`main.py`** — GUI 入口。使用 tkinter 构建界面，处理文件选择、数据预览、进度展示。生成操作在独立线程中执行（`_do_generate`），通过 `root.after` 回写 UI。
- **`core/engine.py`** — 模板引擎。负责占位符扫描、替换、附录表格插入、批量生成。
- **`core/data_loader.py`** — 数据加载。支持 Excel (.xlsx) 和 CSV (.csv)，自动检测编码。

### 核心流程

1. **扫描占位符**：`find_placeholders()` 读取模板，提取所有 `【xxx】` 格式的占位符名称。
2. **生成单份文档**：`generate_single()` 先插入附录表格（如果有），再替换所有文本占位符。顺序不可颠倒——若先替换文本，`【附录A】` 会被清空导致无法定位插入点。
3. **批量生成**：`generate_documents()` 遍历数据行，按 `filename_template` 规则生成文件名，调用 `generate_single()`。

### 占位符格式

模板中使用中文方括号标记变量：`【经销商名称】`、`【协议编号】`。
数据文件中的列名需与占位符名称一致（不含括号）。

### 附录表格的特殊处理

`insert_appendix_table()` 专为欧加隆分销协议设计，包含以下硬编码逻辑：

- 扫描表头中的 **"内舒拿"** 和 **"妈富隆"** 列，在表格第一行生成合并表头 **"Q1是否参与分销项目"**，并横向合并这两列之间的单元格。
- 插入表格后，在附录 A 标题前插入纵向（portrait）分节符，在附录 B 标题前插入横向（landscape）分节符，使附录 A 区域为横向排版。
- 附录表格紧接在附录 A 标题段落后，占位符段落本身会被删除。

修改此函数时需同时考虑 Word XML 结构（`w:sectPr`、`w:vMerge`、`w:gridSpan` 等）。

### 数据格式化

`data_loader.py` 中的 `_format_cell()` 按 Excel 单元格数字格式输出：
- 含 `%` 的格式 → 百分比（如 `0.06` → `6%`）
- 含 `,` 的格式 → 千分位（如 `4051` → `4,051.00`）

### 分析脚本

根目录下有几个独立的分析脚本，用于辅助调试 Word 文件结构：

- `analyze_word_files.py` — 提取占位符、对比模板与结果文件差异
- `simple_analysis.py` — 简化版占位符分析
- `detailed_analysis.py` — 详细文档结构分析（段落样式、表格、分节等）

这些脚本直接读取 `分销协议制作/` 目录下的硬编码路径，仅用于开发调试，不参与主程序运行。

### 打包注意事项

- 使用 PyInstaller `--onefile` 模式打包。
- 通过 `--add-data "core:core"` 将 `core/` 包纳入可执行文件。
- macOS 使用 `build.sh`，Windows 使用 `build.bat`，两者的 `--add-data` 分隔符不同（`:` vs `;`）。
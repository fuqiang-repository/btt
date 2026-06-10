#!/bin/bash
# macOS 打包脚本
# 使用方法: ./build.sh

set -e

echo "🔨 TemplateForge macOS 打包..."
echo ""

# 检查依赖
pip install -r requirements.txt pyinstaller 2>/dev/null || pip3 install -r requirements.txt pyinstaller

# 清理旧构建
rm -rf build dist

# 打包
pyinstaller --name "TemplateForge" \
    --noconfirm \
    --noconsole \
    --onefile \
    --clean \
    --add-data "core:core" \
    main.py

echo ""
echo "✅ 打包完成！可执行文件位于: dist/TemplateForge"
open dist/

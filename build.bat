@echo off
REM Windows 打包脚本
REM 使用方法: build.bat

echo 🔨 TemplateForge Windows 打包...
echo.

pip install -r requirements.txt pyinstaller

REM 清理旧构建
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller --name "TemplateForge" ^
    --noconfirm ^
    --noconsole ^
    --onefile ^
    --clean ^
    --add-data "core;core" ^
    main.py

echo.
echo ✅ 打包完成！可执行文件位于: dist\TemplateForge.exe
explorer dist\

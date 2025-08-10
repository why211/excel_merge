@echo off
chcp 65001 >nul
echo ===============================================
echo 🎯 Excel工具集 - 快速打包脚本
echo ===============================================
echo.

echo 📦 开始打包...
python build_exe.py

echo.
echo ✅ 打包完成！
echo 📁 输出文件位于 dist 目录
echo.
pause
@echo off
chcp 65001 >nul
echo ===============================================
echo 🎯 Excel工具集 - 一键部署脚本
echo ===============================================
echo.

echo 🚀 开始自动化部署...
echo.

python deploy.py

echo.
echo ✅ 部署完成！
echo.
pause
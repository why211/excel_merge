#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集打包脚本
使用PyInstaller将Python程序打包成exe文件
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_build_dirs():
    """清理之前的构建目录"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"🗑️  清理目录: {dir_name}")
            shutil.rmtree(dir_name)

def install_dependencies():
    """安装依赖包"""
    print("📦 安装依赖包...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"], 
                      check=True, capture_output=True, text=True)
        print("✅ 依赖包安装完成")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 依赖包安装失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False

def build_console_exe_only():
    """只构建控制台版本exe"""
    print("🔨 开始构建控制台版exe文件...")
    
    # PyInstaller命令参数
    cmd = [
        "pyinstaller",
        "--onefile",                    # 打包成单个exe文件
        "--console",                    # 保留控制台窗口
        "--name=Excel合并同步工具V1.0",   # 输出文件名
        "--icon=excel.jpg",             # 图标文件
        "--add-data=excel_merger.py;.", # 包含合并模块
        "--add-data=excel_processor.py;.", # 包含处理模块
        "--distpath=dist",              # 输出目录
        "--workpath=build",             # 工作目录
        "--specpath=.",                 # spec文件位置
        "--clean",                      # 清理临时文件
        "excel_tool.py"                 # 主程序文件
    ]
    
    # 如果没有图标文件，移除图标参数
    if not os.path.exists("excel.jpg"):
        cmd.remove("--icon=excel.jpg")
        print("⚠️  未找到图标文件 excel.jpg，将使用默认图标")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ 控制台版exe文件构建完成")
        print(f"📁 输出目录: {os.path.abspath('dist')}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 控制台版exe构建失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False



def create_readme():
    """创建使用说明文件"""
    readme_content = """# Excel合并同步工具V1.0

## 功能介绍
- Excel文件合并：将多个Excel文件合并成一个文件，支持智能去重
- Excel数据同步：将一个或多个Excel文件的数据同步到另一个文件
- 智能列名匹配和字段补充功能
- 冲突数据处理：支持用户选择处理方式
- 新记录插入：自动检测并询问是否插入新记录

## 使用方法
1. 双击运行 `Excel合并同步工具V1.0.exe`（控制台版本）
2. 根据提示选择相应功能
3. 按照程序引导完成操作

## 注意事项
- 请确保Excel文件没有被其他程序占用
- 建议在处理前备份重要数据
- 程序会自动创建备份文件
- 支持.xlsx、.xls等Excel格式

## 版本信息
- 版本：v1.0
- 更新日期：2025年8月10日
- 新增功能：多源同步、冲突处理、新记录插入

## 技术支持
如有问题请联系开发者
"""
    
    with open("dist/README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print("📝 创建使用说明文件")

def main():
    """主函数"""
    print("=" * 60)
    print("🎯 Excel合并同步工具V1.0 - 打包脚本")
    print("=" * 60)
    
    # 检查主程序文件
    if not os.path.exists("excel_tool.py"):
        print("❌ 找不到主程序文件: excel_tool.py")
        return False
    
    # 检查依赖模块
    required_files = ["excel_merger.py", "excel_processor.py"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"❌ 找不到依赖模块: {file}")
            return False
    
    # 清理构建目录
    clean_build_dirs()
    
    # 安装依赖
    if not install_dependencies():
        return False
    
    # 构建exe文件（仅控制台版本）
    success = build_console_exe_only()
    
    if success:
        # 创建说明文件
        create_readme()
        
        print("\n" + "=" * 60)
        print("🎉 打包完成！")
        print("=" * 60)
        print("📁 输出文件:")
        
        dist_path = Path("dist")
        if dist_path.exists():
            for file in dist_path.iterdir():
                if file.is_file():
                    size = file.stat().st_size / (1024 * 1024)  # MB
                    print(f"  📄 {file.name} ({size:.1f} MB)")
        
        print(f"\n📂 完整路径: {os.path.abspath('dist')}")
        print("\n💡 使用建议:")
        print("  • 双击运行 Excel合并同步工具V1.0.exe (控制台版本)")
        print("  • 根据提示选择所需功能")
        print("  • 首次运行可能需要一些时间加载")
        
        return True
    else:
        print("\n❌ 打包过程中出现错误")
        return False

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️  打包过程被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 打包过程出现异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    input("\n按回车键退出...")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集自动化部署脚本
包含打包、版本管理、Git提交和推送功能
"""

import os
import sys
import json
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

class ExcelToolDeployer:
    def __init__(self):
        self.version = "1.0.0"
        self.project_name = "Excel工具集"
        self.build_date = datetime.now().strftime("%Y%m%d_%H%M%S")
        
    def print_header(self, title):
        """打印标题"""
        print("\n" + "=" * 60)
        print(f"🎯 {title}")
        print("=" * 60)
    
    def run_command(self, cmd, description="", check=True):
        """执行命令"""
        if description:
            print(f"🔄 {description}")
        
        try:
            result = subprocess.run(cmd, shell=True, check=check, 
                                  capture_output=True, text=True, encoding='utf-8')
            if result.stdout.strip():
                print(f"✅ {result.stdout.strip()}")
            return True, result.stdout
        except subprocess.CalledProcessError as e:
            print(f"❌ 命令执行失败: {e}")
            if e.stderr:
                print(f"错误信息: {e.stderr}")
            return False, e.stderr
    
    def clean_build_dirs(self):
        """清理构建目录"""
        self.print_header("清理构建环境")
        
        dirs_to_clean = ['build', 'dist', '__pycache__']
        for dir_name in dirs_to_clean:
            if os.path.exists(dir_name):
                shutil.rmtree(dir_name)
                print(f"🗑️  已清理: {dir_name}")
    
    def install_dependencies(self):
        """安装依赖包"""
        self.print_header("安装依赖包")
        
        success, output = self.run_command(
            f"{sys.executable} -m pip install -r requirements.txt",
            "安装Python依赖包"
        )
        
        if success:
            print("✅ 所有依赖包安装完成")
        return success
    
    def build_executables(self):
        """构建exe文件"""
        self.print_header("构建可执行文件")
        
        # 构建GUI版本
        gui_cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            f"--name={self.project_name}",
            "--add-data=excel_merger.py;.",
            "--add-data=excel_processor.py;.",
            "--distpath=dist",
            "--workpath=build",
            "--clean",
            "excel_tool.py"
        ]
        
        success1, _ = self.run_command(
            " ".join(gui_cmd),
            "构建GUI版本"
        )
        
        # 构建控制台版本
        console_cmd = [
            "pyinstaller",
            "--onefile",
            "--console",
            f"--name={self.project_name}_console",
            "--add-data=excel_merger.py;.",
            "--add-data=excel_processor.py;.",
            "--distpath=dist",
            "--workpath=build",
            "--clean",
            "excel_tool.py"
        ]
        
        success2, _ = self.run_command(
            " ".join(console_cmd),
            "构建控制台版本"
        )
        
        return success1 and success2
    
    def create_release_package(self):
        """创建发布包"""
        self.print_header("创建发布包")
        
        # 创建发布目录
        release_dir = f"release_{self.version}_{self.build_date}"
        os.makedirs(release_dir, exist_ok=True)
        
        # 复制exe文件
        dist_path = Path("dist")
        if dist_path.exists():
            for exe_file in dist_path.glob("*.exe"):
                shutil.copy2(exe_file, release_dir)
                print(f"📄 复制文件: {exe_file.name}")
        
        # 创建使用说明
        readme_content = f"""# {self.project_name} v{self.version}

## 📋 功能介绍
- **Excel文件合并**: 将多个Excel文件合并成一个文件
- **Excel数据同步**: 将一个或多个Excel文件的数据同步到另一个文件
- **智能列名匹配**: 自动识别相似的列名
- **字段补充功能**: 自动补充缺失的字段
- **去重处理**: 支持基于学号+姓名的智能去重

## 🚀 使用方法
1. **推荐**: 双击运行 `{self.project_name}.exe` (GUI版本)
2. **调试**: 双击运行 `{self.project_name}_console.exe` (控制台版本)
3. 根据程序提示选择相应功能
4. 按照引导完成Excel文件处理

## ⚠️ 注意事项
- 请确保Excel文件没有被其他程序占用
- 建议在处理前备份重要数据
- 程序会自动创建备份文件
- 首次运行可能需要一些时间加载

## 📊 去重说明
- **学号+姓名完全相同**: 自动合并，静默处理
- **学号相同但姓名不同**: 根据选择的模式处理
  - 自动模式：保留第一条记录
  - 交互式模式：询问用户如何处理

## 📝 版本信息
- **版本**: v{self.version}
- **构建日期**: {datetime.now().strftime("%Y年%m月%d日")}
- **作者**: 小王

## 🔧 技术支持
如有问题请联系开发者

## 📈 更新日志
### v{self.version}
- ✅ 优化去重处理逻辑
- ✅ 减少冗余输出信息
- ✅ 提升用户体验
- ✅ 修复已知问题
"""
        
        with open(f"{release_dir}/README.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        # 创建版本信息文件
        version_info = {
            "version": self.version,
            "build_date": self.build_date,
            "build_time": datetime.now().isoformat(),
            "files": [f.name for f in Path(release_dir).glob("*.exe")]
        }
        
        with open(f"{release_dir}/version.json", "w", encoding="utf-8") as f:
            json.dump(version_info, f, indent=2, ensure_ascii=False)
        
        print(f"📦 发布包创建完成: {release_dir}")
        return release_dir
    
    def git_operations(self, release_dir):
        """Git操作：提交和推送"""
        self.print_header("Git版本控制")
        
        # 检查是否是Git仓库
        if not os.path.exists(".git"):
            print("🔧 初始化Git仓库...")
            self.run_command("git init", "初始化Git仓库")
        
        # 添加文件
        self.run_command("git add .", "添加文件到Git")
        
        # 提交
        commit_message = f"🚀 发布 {self.project_name} v{self.version} - {self.build_date}"
        self.run_command(f'git commit -m "{commit_message}"', "提交更改")
        
        # 创建标签
        tag_name = f"v{self.version}"
        self.run_command(f'git tag -a {tag_name} -m "Release {tag_name}"', f"创建标签 {tag_name}")
        
        # 检查远程仓库
        success, output = self.run_command("git remote -v", check=False)
        
        if "origin" not in output:
            print("\n⚠️  未配置远程仓库")
            print("请手动添加远程仓库:")
            print("git remote add origin <your-repo-url>")
            return False
        
        # 推送到远程仓库
        self.run_command("git push origin main", "推送到远程仓库")
        self.run_command("git push origin --tags", "推送标签")
        
        return True
    
    def show_summary(self, release_dir):
        """显示构建摘要"""
        self.print_header("构建完成")
        
        print(f"🎉 {self.project_name} v{self.version} 构建完成！")
        print(f"📅 构建时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📁 发布目录: {os.path.abspath(release_dir)}")
        
        # 显示文件信息
        release_path = Path(release_dir)
        if release_path.exists():
            print(f"\n📄 发布文件:")
            for file in release_path.iterdir():
                if file.is_file():
                    size = file.stat().st_size / (1024 * 1024)  # MB
                    print(f"  • {file.name} ({size:.1f} MB)")
        
        print(f"\n💡 使用说明:")
        print(f"  • 推荐使用: {self.project_name}.exe")
        print(f"  • 调试版本: {self.project_name}_console.exe")
        print(f"  • 详细说明: README.txt")
    
    def deploy(self):
        """执行完整的部署流程"""
        try:
            print("🎯 Excel工具集自动化部署开始")
            
            # 1. 清理环境
            self.clean_build_dirs()
            
            # 2. 安装依赖
            if not self.install_dependencies():
                return False
            
            # 3. 构建exe
            if not self.build_executables():
                return False
            
            # 4. 创建发布包
            release_dir = self.create_release_package()
            
            # 5. Git操作
            git_success = self.git_operations(release_dir)
            if git_success:
                print("✅ Git操作完成")
            else:
                print("⚠️  Git操作跳过，请手动处理")
            
            # 6. 显示摘要
            self.show_summary(release_dir)
            
            return True
            
        except Exception as e:
            print(f"❌ 部署过程出现异常: {e}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """主函数"""
    deployer = ExcelToolDeployer()
    
    try:
        success = deployer.deploy()
        if success:
            print("\n🎉 部署完成！")
        else:
            print("\n❌ 部署失败！")
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️  部署过程被用户中断")
        sys.exit(1)

if __name__ == "__main__":
    main()
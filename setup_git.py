#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Git仓库配置脚本
帮助配置Git仓库和远程推送
"""

import os
import subprocess
import sys

def run_command(cmd, description=""):
    """执行命令"""
    if description:
        print(f"🔄 {description}")
    
    try:
        result = subprocess.run(cmd, shell=True, check=True, 
                              capture_output=True, text=True, encoding='utf-8')
        if result.stdout.strip():
            print(f"✅ {result.stdout.strip()}")
        return True, result.stdout
    except subprocess.CalledProcessError as e:
        print(f"❌ 命令执行失败: {e}")
        if e.stderr:
            print(f"错误信息: {e.stderr}")
        return False, e.stderr

def check_git():
    """检查Git是否安装"""
    success, output = run_command("git --version", "检查Git版本")
    return success

def init_git_repo():
    """初始化Git仓库"""
    if os.path.exists(".git"):
        print("✅ Git仓库已存在")
        return True
    
    success, _ = run_command("git init", "初始化Git仓库")
    return success

def setup_gitignore():
    """创建.gitignore文件"""
    gitignore_content = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
share/python-wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# PyInstaller
*.manifest
*.spec

# Unit test / coverage reports
htmlcov/
.tox/
.nox/
.coverage
.coverage.*
.cache
nosetests.xml
coverage.xml
*.cover
*.py,cover
.hypothesis/
.pytest_cache/
cover/

# Virtual environments
.env
.venv
env/
venv/
ENV/
env.bak/
venv.bak/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
.DS_Store?
._*
.Spotlight-V100
.Trashes
ehthumbs.db
Thumbs.db

# Project specific
backup*/
release_*/
temp/
test_output/
*.log

# Excel files (temporary)
~$*.xlsx
~$*.xls
"""
    
    with open(".gitignore", "w", encoding="utf-8") as f:
        f.write(gitignore_content)
    
    print("📝 创建.gitignore文件")

def setup_git_config():
    """配置Git用户信息"""
    print("\n📋 配置Git用户信息:")
    
    # 检查现有配置
    success, current_name = run_command("git config user.name", check=False)
    success, current_email = run_command("git config user.email", check=False)
    
    if current_name.strip():
        print(f"当前用户名: {current_name.strip()}")
        use_current = input("是否使用当前用户名？(y/n，默认y): ").strip().lower()
        if use_current not in ['n', 'no', '否']:
            name = current_name.strip()
        else:
            name = input("请输入Git用户名: ").strip()
    else:
        name = input("请输入Git用户名: ").strip()
    
    if current_email.strip():
        print(f"当前邮箱: {current_email.strip()}")
        use_current = input("是否使用当前邮箱？(y/n，默认y): ").strip().lower()
        if use_current not in ['n', 'no', '否']:
            email = current_email.strip()
        else:
            email = input("请输入Git邮箱: ").strip()
    else:
        email = input("请输入Git邮箱: ").strip()
    
    if name:
        run_command(f'git config user.name "{name}"', "设置用户名")
    if email:
        run_command(f'git config user.email "{email}"', "设置邮箱")

def add_remote_repo():
    """添加远程仓库"""
    print("\n🌐 配置远程仓库:")
    
    # 检查现有远程仓库
    success, output = run_command("git remote -v", check=False)
    
    if "origin" in output:
        print("✅ 已配置远程仓库:")
        print(output)
        
        change_remote = input("是否更改远程仓库地址？(y/n，默认n): ").strip().lower()
        if change_remote not in ['y', 'yes', '是']:
            return True
        
        # 移除现有远程仓库
        run_command("git remote remove origin", "移除现有远程仓库")
    
    print("\n请选择远程仓库类型:")
    print("1. GitHub")
    print("2. Gitee (码云)")
    print("3. 其他Git仓库")
    
    choice = input("请选择 (1-3): ").strip()
    
    if choice == "1":
        print("\n📋 GitHub仓库配置:")
        print("格式: https://github.com/用户名/仓库名.git")
        repo_url = input("请输入GitHub仓库地址: ").strip()
    elif choice == "2":
        print("\n📋 Gitee仓库配置:")
        print("格式: https://gitee.com/用户名/仓库名.git")
        repo_url = input("请输入Gitee仓库地址: ").strip()
    else:
        repo_url = input("请输入Git仓库地址: ").strip()
    
    if repo_url:
        success, _ = run_command(f'git remote add origin "{repo_url}"', "添加远程仓库")
        if success:
            print(f"✅ 远程仓库添加成功: {repo_url}")
            return True
    
    return False

def initial_commit():
    """创建初始提交"""
    print("\n📝 创建初始提交:")
    
    # 添加所有文件
    run_command("git add .", "添加文件")
    
    # 检查是否有文件需要提交
    success, output = run_command("git status --porcelain", check=False)
    
    if not output.strip():
        print("✅ 没有需要提交的文件")
        return True
    
    # 创建初始提交
    commit_message = "🎉 初始提交: Excel工具集项目"
    success, _ = run_command(f'git commit -m "{commit_message}"', "创建初始提交")
    
    return success

def push_to_remote():
    """推送到远程仓库"""
    print("\n🚀 推送到远程仓库:")
    
    # 检查是否有远程仓库
    success, output = run_command("git remote -v", check=False)
    if "origin" not in output:
        print("❌ 未配置远程仓库，跳过推送")
        return False
    
    # 设置上游分支并推送
    success, _ = run_command("git push -u origin main", "推送到远程仓库")
    
    if not success:
        # 尝试master分支
        print("🔄 尝试推送到master分支...")
        success, _ = run_command("git push -u origin master", "推送到master分支")
    
    return success

def main():
    """主函数"""
    print("=" * 60)
    print("🎯 Excel工具集 - Git配置脚本")
    print("=" * 60)
    
    # 检查Git
    if not check_git():
        print("❌ Git未安装，请先安装Git")
        print("下载地址: https://git-scm.com/")
        return False
    
    # 初始化仓库
    if not init_git_repo():
        return False
    
    # 创建.gitignore
    setup_gitignore()
    
    # 配置Git用户信息
    setup_git_config()
    
    # 添加远程仓库
    remote_added = add_remote_repo()
    
    # 创建初始提交
    if not initial_commit():
        return False
    
    # 推送到远程仓库
    if remote_added:
        push_success = push_to_remote()
        if push_success:
            print("\n🎉 Git配置和推送完成！")
        else:
            print("\n⚠️  Git配置完成，但推送失败")
            print("请检查网络连接和仓库权限")
    else:
        print("\n✅ Git本地配置完成")
        print("请手动配置远程仓库后再推送")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n⚠️  配置过程被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 配置过程出现异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    input("\n按回车键退出...")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集 - 整合版本
整合了Excel合并和同步功能

作者: 小王
版本: v1.0
日期: 2025年8月8日
"""

import sys
import os
import traceback
import time

def is_console_available():
    """检查是否有可用的控制台"""
    try:
        sys.stdout.write("")
        return True
    except (OSError, AttributeError):
        return False

def show_message_box(title, message, msg_type="info"):
    """显示消息框（用于GUI模式）"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        if msg_type == "error":
            messagebox.showerror(title, message)
        elif msg_type == "warning":
            messagebox.showwarning(title, message)
        else:
            messagebox.showinfo(title, message)
        
        root.destroy()
    except ImportError:
        # 如果tkinter不可用，尝试使用Windows消息框
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, message, title, 0)
        except:
            pass

def safe_input(prompt="", default=""):
    """安全的输入函数，处理在打包环境中可能出现的输入流问题"""
    try:
        if not is_console_available():
            # GUI模式下，显示消息框并返回默认值
            show_message_box("Excel合并同步工具", f"{prompt}\n\n将使用默认值: {default}")
            return default
        return input(prompt)
    except (EOFError, RuntimeError):
        # 在打包环境中如果无法获取输入，使用默认值
        if is_console_available():
            print(f"输入流不可用，使用默认值: {default}")
        else:
            show_message_box("输入错误", f"输入流不可用，使用默认值: {default}")
        return default

def show_welcome():
    """显示欢迎界面"""
    print("=" * 80)
    print("🎯 Excel工具集 v1.0")
    print("📋 功能：Excel文件合并 + Excel数据同步")
    print("=" * 80)
    print("📝 说明：")
    print("  • 本工具整合了Excel合并和同步两个功能")
    print("  • 合并功能：将多个Excel文件合并成一个文件")
    print("  • 同步功能：支持单个或多个Excel文件作为数据源")
    print("  • 同步功能：将一个或者多个Excel文件的数据同步到另一个文件")
    print("  • 所有功能都支持智能列名匹配和字段补充")
    print("=" * 80)

def show_menu():
    """显示主菜单"""
    print("\n📋 请选择要使用的功能：")
    print("1. 📊 Excel文件合并")
    print("   - 将多个Excel文件合并成一个文件")
    print("   - 支持智能列名匹配、字段补充、去重等功能")
    print("   - 适合处理多个格式相似的数据文件")
    print()
    print("2. 🔄 Excel数据同步")
    print("   - 支持单个或多个Excel文件作为数据源")
    print("   - 将一个或者多个Excel文件的数据同步到另一个文件")
    print("   - 支持基于关联字段的数据更新")
    print("   - 适合更新现有文件中的部分数据")
    print()
    print("3. ❌ 退出程序")
    print()
    print("💡 提示：")
    print("  • 合并功能适合处理多个数据源文件")
    print("  • 同步功能适合更新现有文件的数据")
    print("  • 两个功能都支持智能列名识别")

def run_merge_function():
    """运行Excel合并功能"""
    print("\n" + "=" * 60)
    print("📊 启动Excel文件合并功能")
    print("=" * 60)
    
    try:
        # 导入合并模块
        from excel_merger import ExcelProcessor
        
        # 创建处理器实例并运行
        processor = ExcelProcessor()
        processor.run()
        
    except ImportError as e:
        print(f"❌ 导入合并模块失败: {str(e)}")
        print("请确保 excel_merger.py 文件存在且可访问")
        return False
    except Exception as e:
        print(f"❌ 运行合并功能时出错: {str(e)}")
        print("详细错误信息:")
        traceback.print_exc()
        return False
    
    return True

def run_sync_function():
    """运行Excel同步功能"""
    print("\n" + "=" * 60)
    print("🔄 启动Excel数据同步功能")
    print("=" * 60)
    
    try:
        # 导入同步模块
        from excel_processor import ExcelProcessor
        
        # 创建处理器实例
        processor = ExcelProcessor()
        
        # 运行同步功能，只显示同步相关选项
        processor.run_sync_only()
        
    except ImportError as e:
        print(f"❌ 导入同步模块失败: {str(e)}")
        print("请确保 excel_processor.py 文件存在且可访问")
        return False
    except Exception as e:
        print(f"❌ 运行同步功能时出错: {str(e)}")
        print("详细错误信息:")
        traceback.print_exc()
        return False
    
    return True



def main():
    """主函数"""
    try:
        # 显示欢迎界面
        show_welcome()
        
        # 主程序循环
        while True:
            try:
                # 显示菜单
                show_menu()
                
                # 获取用户选择
                choice = safe_input("\n请输入选择 (1/2/3): ", "3").strip()
                
                if choice == '1':
                    # 运行合并功能
                    success = run_merge_function()
                    if success:
                        print("\n✅ 合并功能执行完成")
                    else:
                        print("\n❌ 合并功能执行失败")
                    
                    # 询问是否继续
                    continue_choice = safe_input("\n是否返回主菜单？(y/n，默认y): ", "y").strip().lower()
                    if continue_choice in ['n', 'no', '否']:
                        print("👋 程序退出")
                        break
                
                elif choice == '2':
                    # 运行同步功能
                    success = run_sync_function()
                    if success:
                        print("\n✅ 同步功能执行完成")
                    else:
                        print("\n❌ 同步功能执行失败")
                    
                    # 询问是否继续
                    continue_choice = safe_input("\n是否返回主菜单？(y/n，默认y): ", "y").strip().lower()
                    if continue_choice in ['n', 'no', '否']:
                        print("👋 程序退出")
                        break
                
                elif choice == '3':
                    print("👋 程序退出")
                    break
                
                else:
                    print("❌ 无效选择，请输入 1、2 或 3")
                    continue
                
            except KeyboardInterrupt:
                print("\n\n⚠️  程序被用户中断")
                break
            except Exception as e:
                print(f"\n❌ 程序执行出错: {str(e)}")
                print("详细错误信息:")
                traceback.print_exc()
                
                # 询问是否继续
                continue_choice = safe_input("\n是否返回主菜单？(y/n，默认y): ", "y").strip().lower()
                if continue_choice in ['n', 'no', '否']:
                    print("👋 程序退出")
                    break
    
    except Exception as e:
        error_msg = f"程序启动失败: {str(e)}"
        if is_console_available():
            print(f"\n❌ {error_msg}")
            print("详细错误信息:")
            traceback.print_exc()
        else:
            # GUI模式下显示错误消息框
            show_message_box("Excel合并同步工具 - 错误", error_msg, "error")
    
    finally:
        if is_console_available():
            print("\n感谢使用Excel工具集！")
        try:
            safe_input("按回车键退出...")
        except Exception:
            # 处理在打包环境中可能出现的输入流问题
            if is_console_available():
                print("程序将在3秒后自动退出...")
            time.sleep(3)

if __name__ == "__main__":
    main() 
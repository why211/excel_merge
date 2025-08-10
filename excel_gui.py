#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具集 - GUI版本
提供图形界面的Excel合并和同步工具
"""

import sys
import os
import traceback
import tkinter as tk
from tkinter import messagebox, ttk

class ExcelToolGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel合并同步工具V1.0")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 设置窗口图标（如果有的话）
        try:
            if os.path.exists("excel.ico"):
                self.root.iconbitmap("excel.ico")
        except:
            pass
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 标题
        title_frame = tk.Frame(self.root)
        title_frame.pack(pady=20)
        
        title_label = tk.Label(title_frame, text="Excel合并同步工具V1.0", 
                              font=("Arial", 16, "bold"))
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, text="功能：Excel文件合并 + Excel数据同步", 
                                 font=("Arial", 10))
        subtitle_label.pack(pady=5)
        
        # 功能说明
        desc_frame = tk.Frame(self.root)
        desc_frame.pack(pady=10, padx=20, fill="x")
        
        desc_text = """说明：
• 本工具整合了Excel合并和同步两个功能
• 合并功能：将多个Excel文件合并成一个文件
• 同步功能：支持单个或多个Excel文件作为数据源
• 同步功能：将一个或者多个Excel文件的数据同步到另一个文件
• 所有功能都支持智能列名匹配和字段补充"""
        
        desc_label = tk.Label(desc_frame, text=desc_text, justify="left", 
                             font=("Arial", 9), wraplength=550)
        desc_label.pack()
        
        # 功能按钮
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        # 合并功能按钮
        merge_btn = tk.Button(button_frame, text="📊 Excel文件合并", 
                             font=("Arial", 12), width=20, height=2,
                             command=self.run_merge_function)
        merge_btn.pack(pady=10)
        
        merge_desc = tk.Label(button_frame, 
                             text="将多个Excel文件合并成一个文件\n支持智能列名匹配、字段补充、去重等功能", 
                             font=("Arial", 8), fg="gray")
        merge_desc.pack()
        
        # 同步功能按钮
        sync_btn = tk.Button(button_frame, text="🔄 Excel数据同步", 
                            font=("Arial", 12), width=20, height=2,
                            command=self.run_sync_function)
        sync_btn.pack(pady=(20, 10))
        
        sync_desc = tk.Label(button_frame, 
                            text="将一个或多个Excel文件的数据同步到另一个文件\n支持基于关联字段的数据更新", 
                            font=("Arial", 8), fg="gray")
        sync_desc.pack()
        
        # 退出按钮
        exit_btn = tk.Button(button_frame, text="❌ 退出程序", 
                            font=("Arial", 12), width=20, height=1,
                            command=self.exit_program)
        exit_btn.pack(pady=(20, 10))
        
    def run_merge_function(self):
        """运行Excel合并功能"""
        try:
            # 隐藏主窗口
            self.root.withdraw()
            
            # 导入合并模块
            from excel_merger import ExcelProcessor
            
            # 创建处理器实例并运行
            processor = ExcelProcessor()
            processor.run()
            
            # 显示完成消息
            messagebox.showinfo("完成", "Excel文件合并功能执行完成！")
            
        except ImportError as e:
            messagebox.showerror("错误", f"导入合并模块失败: {str(e)}\n请确保 excel_merger.py 文件存在且可访问")
        except Exception as e:
            messagebox.showerror("错误", f"运行合并功能时出错: {str(e)}")
        finally:
            # 显示主窗口
            self.root.deiconify()
    
    def run_sync_function(self):
        """运行Excel同步功能"""
        try:
            # 隐藏主窗口
            self.root.withdraw()
            
            # 导入同步模块
            from excel_processor import ExcelProcessor
            
            # 创建处理器实例
            processor = ExcelProcessor()
            
            # 运行同步功能，只显示同步相关选项
            processor.run_sync_only()
            
            # 显示完成消息
            messagebox.showinfo("完成", "Excel数据同步功能执行完成！")
            
        except ImportError as e:
            messagebox.showerror("错误", f"导入同步模块失败: {str(e)}\n请确保 excel_processor.py 文件存在且可访问")
        except Exception as e:
            messagebox.showerror("错误", f"运行同步功能时出错: {str(e)}")
        finally:
            # 显示主窗口
            self.root.deiconify()
    
    def exit_program(self):
        """退出程序"""
        if messagebox.askyesno("确认退出", "确定要退出Excel合并同步工具吗？"):
            self.root.quit()
            self.root.destroy()
    
    def run(self):
        """运行GUI程序"""
        try:
            self.root.mainloop()
        except Exception as e:
            messagebox.showerror("错误", f"程序运行出错: {str(e)}")

def main():
    """主函数"""
    try:
        # 检查是否在GUI模式下运行
        if len(sys.argv) > 1 and sys.argv[1] == "--console":
            # 控制台模式
            from excel_tool import main as console_main
            console_main()
        else:
            # GUI模式
            app = ExcelToolGUI()
            app.run()
    except Exception as e:
        try:
            messagebox.showerror("启动错误", f"程序启动失败: {str(e)}")
        except:
            print(f"程序启动失败: {str(e)}")

if __name__ == "__main__":
    main()
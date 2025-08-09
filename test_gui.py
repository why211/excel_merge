#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GUI消息框功能
"""

import tkinter as tk
from tkinter import messagebox

def test_message_box():
    """测试消息框"""
    try:
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        result = messagebox.askyesno("Excel合并同步工具", "程序正在运行。\n\n是否要继续？")
        
        if result:
            messagebox.showinfo("成功", "GUI消息框功能正常！")
        else:
            messagebox.showinfo("取消", "用户取消操作。")
        
        root.destroy()
        print("GUI测试完成")
    except Exception as e:
        print(f"GUI测试失败: {e}")

if __name__ == "__main__":
    test_message_box()
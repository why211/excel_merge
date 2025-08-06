#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
功能测试脚本

用于测试新增的功能是否正常工作
"""

import os
import sys
from excel_processor import analyze_excel_fields, select_files_by_fields, process_excel_files

def test_analyze_fields():
    """测试字段分析功能"""
    print("=== 测试字段分析功能 ===")
    
    folder_path = "excel"
    if not os.path.exists(folder_path):
        print(f"错误：测试文件夹 '{folder_path}' 不存在")
        return False
    
    try:
        field_stats = analyze_excel_fields(folder_path)
        if field_stats is not None:
            print("✓ 字段分析功能测试通过")
            return True
        else:
            print("✗ 字段分析功能测试失败")
            return False
    except Exception as e:
        print(f"✗ 字段分析功能测试出错: {str(e)}")
        return False

def test_select_files():
    """测试文件选择功能"""
    print("\n=== 测试文件选择功能 ===")
    
    folder_path = "excel"
    required_fields = ["学号", "*学生姓名"]
    
    if not os.path.exists(folder_path):
        print(f"错误：测试文件夹 '{folder_path}' 不存在")
        return False
    
    try:
        selected_files = select_files_by_fields(folder_path, required_fields)
        if isinstance(selected_files, list):
            print(f"✓ 文件选择功能测试通过，选择了 {len(selected_files)} 个文件")
            return True
        else:
            print("✗ 文件选择功能测试失败")
            return False
    except Exception as e:
        print(f"✗ 文件选择功能测试出错: {str(e)}")
        return False

def test_process_with_dedup():
    """测试处理功能（去重）"""
    print("\n=== 测试处理功能（去重） ===")
    
    folder_path = "excel"
    output_filename = "test_result_dedup.xlsx"
    required_fields = ["学号", "*学生姓名"]
    deduplicate = True
    
    if not os.path.exists(folder_path):
        print(f"错误：测试文件夹 '{folder_path}' 不存在")
        return False
    
    try:
        process_excel_files(folder_path, output_filename, required_fields, deduplicate)
        
        # 检查输出文件是否生成
        output_path = os.path.join(folder_path, output_filename)
        if os.path.exists(output_path):
            print("✓ 处理功能（去重）测试通过")
            # 清理测试文件
            try:
                os.remove(output_path)
                print("  已清理测试文件")
            except:
                pass
            return True
        else:
            print("✗ 处理功能（去重）测试失败：输出文件未生成")
            return False
    except Exception as e:
        print(f"✗ 处理功能（去重）测试出错: {str(e)}")
        return False

def test_process_without_dedup():
    """测试处理功能（不去重）"""
    print("\n=== 测试处理功能（不去重） ===")
    
    folder_path = "excel"
    output_filename = "test_result_no_dedup.xlsx"
    required_fields = ["学号", "*学生姓名"]
    deduplicate = False
    
    if not os.path.exists(folder_path):
        print(f"错误：测试文件夹 '{folder_path}' 不存在")
        return False
    
    try:
        process_excel_files(folder_path, output_filename, required_fields, deduplicate)
        
        # 检查输出文件是否生成
        output_path = os.path.join(folder_path, output_filename)
        if os.path.exists(output_path):
            print("✓ 处理功能（不去重）测试通过")
            # 清理测试文件
            try:
                os.remove(output_path)
                print("  已清理测试文件")
            except:
                pass
            return True
        else:
            print("✗ 处理功能（不去重）测试失败：输出文件未生成")
            return False
    except Exception as e:
        print(f"✗ 处理功能（不去重）测试出错: {str(e)}")
        return False

def test_custom_fields():
    """测试自定义字段功能"""
    print("\n=== 测试自定义字段功能 ===")
    
    folder_path = "excel"
    output_filename = "test_result_custom.xlsx"
    
    # 先分析字段，然后选择一些常见的字段进行测试
    field_stats = analyze_excel_fields(folder_path)
    if field_stats is None:
        print("✗ 无法分析字段，跳过自定义字段测试")
        return False
    
    # 选择出现频率较高的字段进行测试
    sorted_fields = sorted(field_stats.items(), key=lambda x: len(x[1]), reverse=True)
    if len(sorted_fields) >= 2:
        test_fields = [sorted_fields[0][0], sorted_fields[1][0]]
        print(f"使用字段进行测试: {test_fields}")
        
        try:
            process_excel_files(folder_path, output_filename, test_fields, True)
            
            # 检查输出文件是否生成
            output_path = os.path.join(folder_path, output_filename)
            if os.path.exists(output_path):
                print("✓ 自定义字段功能测试通过")
                # 清理测试文件
                try:
                    os.remove(output_path)
                    print("  已清理测试文件")
                except:
                    pass
                return True
            else:
                print("✗ 自定义字段功能测试失败：输出文件未生成")
                return False
        except Exception as e:
            print(f"✗ 自定义字段功能测试出错: {str(e)}")
            return False
    else:
        print("✗ 可用字段不足，跳过自定义字段测试")
        return False

def main():
    """主测试函数"""
    print("Excel文件处理工具 - 功能测试")
    print("=" * 50)
    
    # 检查测试环境
    if not os.path.exists("excel"):
        print("错误：找不到 'excel' 文件夹，请确保测试环境正确")
        print("请将Excel文件放在 'excel' 文件夹中，然后重新运行测试")
        return
    
    # 运行所有测试
    tests = [
        test_analyze_fields,
        test_select_files,
        test_process_with_dedup,
        test_process_without_dedup,
        test_custom_fields
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        try:
            if test():
                passed += 1
        except Exception as e:
            print(f"测试 {test.__name__} 出现异常: {str(e)}")
    
    print("\n" + "=" * 50)
    print(f"测试结果: {passed}/{total} 通过")
    
    if passed == total:
        print("🎉 所有测试通过！功能正常工作。")
    else:
        print("⚠️  部分测试失败，请检查相关功能。")
    
    print("\n测试完成！")

if __name__ == "__main__":
    main() 
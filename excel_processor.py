import pandas as pd
import os
from pathlib import Path
from collections import defaultdict

def analyze_excel_fields(folder_path):
    """
    分析文件夹中所有Excel文件的字段
    
    Args:
        folder_path (str): 包含Excel文件的文件夹路径
    
    Returns:
        dict: 字段统计信息
    """
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return None
    
    # 存储字段统计信息
    field_stats = defaultdict(list)
    all_fields = set()
    
    # 遍历文件夹中的所有文件
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_files.append(file)
    
    if not excel_files:
        print(f"在文件夹 '{folder_path}' 中没有找到Excel文件")
        return None
    
    print(f"正在分析 {len(excel_files)} 个Excel文件的字段...")
    
    # 分析每个Excel文件
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 记录每个字段在哪些文件中出现
            for field in df.columns:
                field_stats[field].append(file)
                all_fields.add(field)
                
        except Exception as e:
            print(f"  错误：分析文件 '{file}' 时出错: {str(e)}")
            continue
    
    # 生成字段统计报告
    print("\n=== 字段统计报告 ===")
    print(f"总共发现 {len(all_fields)} 个不同的字段")
    print("\n字段出现情况：")
    
    # 按出现频率排序
    sorted_fields = sorted(field_stats.items(), key=lambda x: len(x[1]), reverse=True)
    
    for field, files in sorted_fields:
        print(f"\n字段: {field}")
        print(f"  出现次数: {len(files)}")
        print(f"  出现文件: {', '.join(files)}")
    
    return field_stats

def select_files_by_fields(folder_path, required_fields):
    """
    根据字段选择文件
    
    Args:
        folder_path (str): 包含Excel文件的文件夹路径
        required_fields (list): 必需的字段列表
    
    Returns:
        list: 包含所有必需字段的文件列表
    """
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return []
    
    # 遍历文件夹中的所有文件
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_files.append(file)
    
    if not excel_files:
        print(f"在文件夹 '{folder_path}' 中没有找到Excel文件")
        return []
    
    # 筛选包含所有必需字段的文件
    selected_files = []
    
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 检查是否包含所有必需字段
            missing_fields = [field for field in required_fields if field not in df.columns]
            
            if not missing_fields:
                selected_files.append(file)
                print(f"✓ 文件 '{file}' 包含所有必需字段")
            else:
                print(f"✗ 文件 '{file}' 缺少字段: {missing_fields}")
                
        except Exception as e:
            print(f"  错误：检查文件 '{file}' 时出错: {str(e)}")
            continue
    
    print(f"\n总共选择了 {len(selected_files)} 个文件进行处理")
    return selected_files

def process_excel_files(folder_path, output_filename="result.xlsx", required_fields=None, deduplicate=True):
    """
    遍历指定文件夹中的所有Excel文件，读取指定字段，合并数据并保存
    
    Args:
        folder_path (str): 包含Excel文件的文件夹路径
        output_filename (str): 输出文件名，默认为"result.xlsx"
        required_fields (list): 必需的字段列表，默认为["学号", "*学生姓名"]
        deduplicate (bool): 是否去重，默认为True
    """
    
    # 设置默认字段
    if required_fields is None:
        required_fields = ["学号", "*学生姓名"]
    
    # 存储所有数据的列表
    all_data = []
    
    # 确保文件夹路径存在
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return
    
    # 选择包含所有必需字段的文件
    selected_files = select_files_by_fields(folder_path, required_fields)
    
    if not selected_files:
        print("没有找到包含所有必需字段的文件")
        return
    
    # 处理每个选中的Excel文件
    for file in selected_files:
        file_path = os.path.join(folder_path, file)
        print(f"正在处理文件: {file}")
        
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 提取所需的列
            selected_data = df[required_fields].copy()
            
            # 添加文件名列以便追踪数据来源
            selected_data['来源文件'] = file
            
            # 将数据添加到总列表中
            all_data.append(selected_data)
            print(f"  成功读取 {len(selected_data)} 行数据")
            
        except Exception as e:
            print(f"  错误：处理文件 '{file}' 时出错: {str(e)}")
            continue
    
    if not all_data:
        print("没有成功读取任何数据")
        return
    
    # 合并所有数据
    print("正在合并数据...")
    combined_df = pd.concat(all_data, ignore_index=True)
    
    print(f"合并前总行数: {len(combined_df)}")
    
    # 根据用户选择决定是否去重
    if deduplicate and '学号' in required_fields:
        print("正在按学号去重...")
        combined_df = combined_df.drop_duplicates(subset=['学号'], keep='first')
        print(f"去重后总行数: {len(combined_df)}")
    elif deduplicate:
        print("警告：未找到'学号'字段，无法进行去重操作")
    else:
        print("用户选择不去重，保留所有记录")
    
    # 保存为新的Excel文件
    output_path = os.path.join(folder_path, output_filename)
    
    # 检查文件是否被占用
    if os.path.exists(output_path):
        try:
            # 尝试删除已存在的文件
            os.remove(output_path)
            print(f"已删除已存在的文件: {output_path}")
        except PermissionError:
            # 如果无法删除，尝试使用不同的文件名
            base_name = os.path.splitext(output_filename)[0]
            extension = os.path.splitext(output_filename)[1]
            counter = 1
            while True:
                new_filename = f"{base_name}_{counter}{extension}"
                new_output_path = os.path.join(folder_path, new_filename)
                if not os.path.exists(new_output_path):
                    output_path = new_output_path
                    output_filename = new_filename
                    print(f"文件被占用，使用新文件名: {new_filename}")
                    break
                counter += 1
    
    try:
        combined_df.to_excel(output_path, index=False)
        print(f"数据已成功保存到: {output_path}")
        print(f"总共处理了 {len(combined_df)} 条记录")
    except PermissionError:
        print(f"权限错误：无法保存到 {output_path}")
        print("请确保文件没有被其他程序打开，或者尝试保存到其他位置")
    except Exception as e:
        print(f"保存文件时出错: {str(e)}")

def process_with_deduplication_option():
    """
    带去重选项的处理函数
    """
    try:
        # 获取用户输入
        folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用当前目录）: ").strip()
        if not folder_path:
            folder_path = "."
        
        output_filename = input("请输入输出文件名（或按回车使用默认名称 'result.xlsx'）: ").strip()
        if not output_filename:
            output_filename = "result.xlsx"
        
        # 询问是否需要去重
        deduplicate_input = input("是否需要去重？(y/n，默认y): ").strip().lower()
        deduplicate = deduplicate_input != 'n'
        
        # 询问是否需要分析字段
        analyze_fields = input("是否需要分析字段？(y/n，默认n): ").strip().lower() == 'y'
        
        if analyze_fields:
            analyze_excel_fields(folder_path)
        
        # 询问需要合并哪些字段
        print("\n请输入需要合并的字段名（用逗号分隔，按回车使用默认字段'学号,*学生姓名'）:")
        fields_input = input().strip()
        if fields_input:
            required_fields = [field.strip() for field in fields_input.split(',')]
        else:
            required_fields = ["学号", "*学生姓名"]
        
    except EOFError:
        # 非交互式环境，使用默认值
        folder_path = "."
        output_filename = "result.xlsx"
        deduplicate = True
        required_fields = ["学号", "*学生姓名"]
        print(f"使用默认设置：文件夹路径='{folder_path}', 输出文件名='{output_filename}', 去重={deduplicate}, 字段={required_fields}")
    
    # 执行处理
    process_excel_files(folder_path, output_filename, required_fields, deduplicate)

def main():
    """主函数，设置参数并执行处理"""
    
    print("=== Excel文件处理工具 ===")
    print("1. 分析字段")
    print("2. 处理文件")
    print("3. 分析字段并处理文件")
    
    try:
        choice = input("请选择操作 (1/2/3，默认2): ").strip()
        if not choice:
            choice = "2"
        
        if choice == "1":
            folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用当前目录）: ").strip()
            if not folder_path:
                folder_path = "."
            analyze_excel_fields(folder_path)
        elif choice == "2":
            process_with_deduplication_option()
        elif choice == "3":
            folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用当前目录）: ").strip()
            if not folder_path:
                folder_path = "."
            analyze_excel_fields(folder_path)
            print("\n" + "="*50 + "\n")
            process_with_deduplication_option()
        else:
            print("无效选择，使用默认操作")
            process_with_deduplication_option()
            
    except EOFError:
        # 非交互式环境，使用默认操作
        print("使用默认操作：处理文件")
        process_with_deduplication_option()

if __name__ == "__main__":
    main() 
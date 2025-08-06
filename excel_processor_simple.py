import pandas as pd
import os

def process_excel_files(folder_path, output_filename="result.xlsx"):
    """
    遍历指定文件夹中的所有Excel文件，读取"学号"和"*学生姓名"列，合并数据并保存
    
    Args:
        folder_path (str): 包含Excel文件的文件夹路径
        output_filename (str): 输出文件名，默认为"result.xlsx"
    """
    
    # 存储所有数据的列表
    all_data = []
    
    # 确保文件夹路径存在
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return
    
    # 遍历文件夹中的所有文件
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_files.append(file)
    
    if not excel_files:
        print(f"在文件夹 '{folder_path}' 中没有找到Excel文件")
        return
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    # 处理每个Excel文件
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        print(f"正在处理文件: {file}")
        
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)
            
            # 检查是否包含所需的列
            required_columns = ["学号", "*学生姓名"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                print(f"  警告：文件 '{file}' 缺少以下列: {missing_columns}，跳过此文件")
                continue
            
            # 提取所需的列
            selected_data = df[required_columns].copy()
            
            # 添加文件名列以便追踪数据来源（可选）
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
    
    # 移除重复的学号（如果需要）
    print(f"合并前总行数: {len(combined_df)}")
    combined_df = combined_df.drop_duplicates(subset=['学号'], keep='first')
    print(f"去重后总行数: {len(combined_df)}")
    
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

if __name__ == "__main__":
    # 在这里设置你的参数
    folder_path = "excel"  # 修改为excel文件夹路径
    output_filename = "学生名单.xlsx"  # 输出文件名
    
    # 执行处理
    process_excel_files(folder_path, output_filename) 
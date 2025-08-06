import pandas as pd
import os
import glob
from typing import List, Tuple

class ExcelProcessor:
    """Excel文件处理工具"""
    
    def __init__(self):
        self.selected_files = []
        self.all_fields = []
        self.selected_fields = []
        self.deduplicate = False
        self.dedup_fields = []
        self.output_filename = "result.xlsx"
    
    def select_files(self, folder_path: str = ".") -> List[str]:
        """
        文件选择功能
        
        Args:
            folder_path: 文件夹路径，默认为当前目录
            
        Returns:
            选中的文件列表
        """
        print(f"\n=== 步骤1: 文件选择 ===")
        print(f"正在扫描文件夹: {folder_path}")
        
        # 查找所有Excel文件
        excel_patterns = ['*.xlsx', '*.xls']
        excel_files = []
        
        for pattern in excel_patterns:
            excel_files.extend(glob.glob(os.path.join(folder_path, pattern)))
        
        if not excel_files:
            print(f"❌ 在文件夹 '{folder_path}' 中没有找到Excel文件")
            return []
        
        # 显示找到的文件
        print(f"\n✅ 找到 {len(excel_files)} 个Excel文件:")
        for i, file in enumerate(excel_files, 1):
            filename = os.path.basename(file)
            file_size = os.path.getsize(file) / 1024  # KB
            print(f"{i:2d}. {filename:<30} ({file_size:.1f} KB)")
        
        # 用户选择文件
        print(f"\n请选择要处理的文件:")
        print("�� 输入文件编号（用逗号分隔，如：1,2,3）")
        print("�� 输入 'all' 选择所有文件")
        print("📝 输入 'q' 退出程序")
        
        try:
            choice = input("\n请选择: ").strip().lower()
            
            if choice == 'q':
                print("程序退出")
                return []
            elif choice == 'all':
                self.selected_files = excel_files
                print(f"✅ 已选择所有 {len(excel_files)} 个文件")
            else:
                # 解析用户选择的文件编号
                indices = [int(x.strip()) - 1 for x in choice.split(',')]
                self.selected_files = [excel_files[i] for i in indices if 0 <= i < len(excel_files)]
                
                if not self.selected_files:
                    print("❌ 未选择任何有效文件，请重新选择")
                    return self.select_files(folder_path)
                
                print(f"✅ 已选择 {len(self.selected_files)} 个文件:")
                for file in self.selected_files:
                    print(f"  📄 {os.path.basename(file)}")
                
            return self.selected_files
            
        except (ValueError, IndexError) as e:
            print(f"❌ 输入格式错误: {str(e)}，请重新选择")
            return self.select_files(folder_path)
    
    def get_field_list(self, files: List[str]) -> List[str]:
        """
        获取所有文件的字段列表
        
        Args:
            files: 文件列表
            
        Returns:
            所有字段的列表
        """
        print(f"\n=== 步骤2: 字段分析 ===")
        all_fields = set()
        file_field_info = {}
        
        for file in files:
            try:
                df = pd.read_excel(file)
                file_fields = list(df.columns)
                all_fields.update(file_fields)
                file_field_info[os.path.basename(file)] = {
                    'field_count': len(file_fields),
                    'fields': file_fields
                }
                print(f"📊 文件 '{os.path.basename(file)}' 包含 {len(file_fields)} 个字段")
                
            except Exception as e:
                print(f"❌ 读取文件 '{os.path.basename(file)}' 时出错: {str(e)}")
        
        # 计算每个字段的出现次数并排序
        field_occurrence = {}
        for field in all_fields:
            files_with_field = [f for f, info in file_field_info.items() if field in info['fields']]
            field_occurrence[field] = len(files_with_field)
        
        # 按出现次数从高到低排序
        sorted_fields = sorted(field_occurrence.items(), key=lambda x: x[1], reverse=True)
        self.all_fields = [field for field, count in sorted_fields]
        
        print(f"\n✅ 总共发现 {len(self.all_fields)} 个不同字段")
        
        return self.all_fields
    
    def get_file_fields(self, file_path: str) -> List[str]:
        """
        获取单个文件的字段列表
        
        Args:
            file_path: 文件路径
            
        Returns:
            字段列表
        """
        try:
            df = pd.read_excel(file_path)
            return list(df.columns)
        except Exception as e:
            return []
    
    def select_fields(self, all_fields: List[str]) -> List[str]:
        """
        字段选择功能
        
        Args:
            all_fields: 所有可用字段列表
            
        Returns:
            选中的字段列表
        """
        print(f"\n=== 步骤3: 字段选择 ===")
        print("📋 可用字段列表（按出现次数排序）:")
        
        # 分页显示字段
        page_size = 10
        total_pages = (len(all_fields) + page_size - 1) // page_size
        
        for page in range(total_pages):
            start_idx = page * page_size
            end_idx = min(start_idx + page_size, len(all_fields))
            
            print(f"\n--- 第 {page + 1}/{total_pages} 页 ---")
            for i in range(start_idx, end_idx):
                field = all_fields[i]
                # 计算该字段的出现次数
                occurrence_count = sum(1 for f in self.selected_files if field in self.get_file_fields(f))
                print(f"{i + 1:2d}. {field:<25} (出现在 {occurrence_count} 个文件中)")
        
        print(f"\n请选择要导入的字段:")
        print("�� 输入字段编号（用逗号分隔，如：1,2,3）")
        print("�� 输入 'all' 选择所有字段")
        print("�� 输入 'page 1' 查看第1页（可替换页码）")
        
        try:
            choice = input("\n请选择: ").strip().lower()
            
            if choice.startswith('page '):
                try:
                    page_num = int(choice.split()[1]) - 1
                    if 0 <= page_num < total_pages:
                        print(f"\n--- 第 {page_num + 1}/{total_pages} 页 ---")
                        start_idx = page_num * page_size
                        end_idx = min(start_idx + page_size, len(all_fields))
                        for i in range(start_idx, end_idx):
                            field = all_fields[i]
                            print(f"{i + 1:2d}. {field}")
                        return self.select_fields(all_fields)
                    else:
                        print("❌ 页码超出范围")
                        return self.select_fields(all_fields)
                except:
                    print("❌ 页码格式错误")
                    return self.select_fields(all_fields)
            
            elif choice == 'all':
                self.selected_fields = all_fields
                print(f"✅ 已选择所有 {len(all_fields)} 个字段")
            else:
                # 解析用户选择的字段编号
                indices = [int(x.strip()) - 1 for x in choice.split(',')]
                self.selected_fields = [all_fields[i] for i in indices if 0 <= i < len(all_fields)]
                
                if not self.selected_fields:
                    print("❌ 未选择任何有效字段，请重新选择")
                    return self.select_fields(all_fields)
                
                print(f"✅ 已选择 {len(self.selected_fields)} 个字段:")
                for field in self.selected_fields:
                    print(f"  📋 {field}")
                
            return self.selected_fields
            
        except (ValueError, IndexError) as e:
            print(f"❌ 输入格式错误: {str(e)}，请重新选择")
            return self.select_fields(all_fields)
    
    def configure_deduplication(self) -> Tuple[bool, List[str]]:
        """
        去重配置：返回(是否去重, 去重字段列表)
        
        Returns:
            (是否去重, 去重字段列表)
        """
        print(f"\n=== 步骤4: 去重配置 ===")
        
        # 询问是否需要去重
        print("🤔 是否需要去重？")
        print("📝 去重将删除重复的记录，保留第一条")
        dedup_choice = input("请选择 (y/n，默认n): ").strip().lower()
        self.deduplicate = dedup_choice in ['y', 'yes', '是']
        
        if not self.deduplicate:
            print("✅ 已选择不去重，将保留所有记录")
            return False, []
        
        # 如果去重，选择去重字段
        print(f"\n📋 请选择去重字段（基于这些字段的组合来判断重复）:")
        print("可用字段列表:")
        for i, field in enumerate(self.selected_fields, 1):
            # 计算该字段的出现次数
            occurrence_count = sum(1 for f in self.selected_files if field in self.get_file_fields(f))
            print(f"{i:2d}. {field:<25} (出现在 {occurrence_count} 个文件中)")
        
        print(f"\n�� 输入字段编号（用逗号分隔，如：1,2）")
        print(f"📝 输入 'all' 使用所有选中字段进行去重")
        print(f"�� 输入 'single 1' 只使用第1个字段去重")
        
        try:
            choice = input("\n请选择去重字段: ").strip().lower()
            
            if choice == 'all':
                self.dedup_fields = self.selected_fields.copy()
                print(f"✅ 已选择所有 {len(self.dedup_fields)} 个字段进行去重")
            elif choice.startswith('single '):
                try:
                    field_idx = int(choice.split()[1]) - 1
                    if 0 <= field_idx < len(self.selected_fields):
                        self.dedup_fields = [self.selected_fields[field_idx]]
                        print(f"✅ 已选择单个字段进行去重: {self.dedup_fields[0]}")
                    else:
                        print("❌ 字段编号超出范围")
                        return self.configure_deduplication()
                except:
                    print("❌ 字段编号格式错误")
                    return self.configure_deduplication()
            else:
                # 解析用户选择的字段编号
                indices = [int(x.strip()) - 1 for x in choice.split(',')]
                self.dedup_fields = [self.selected_fields[i] for i in indices if 0 <= i < len(self.selected_fields)]
                
                if not self.dedup_fields:
                    print("❌ 未选择任何有效字段，请重新选择")
                    return self.configure_deduplication()
                
                print(f"✅ 已选择 {len(self.dedup_fields)} 个字段进行去重:")
                for field in self.dedup_fields:
                    print(f"  🔍 {field}")
                
            return True, self.dedup_fields
            
        except (ValueError, IndexError) as e:
            print(f"❌ 输入格式错误: {str(e)}，请重新选择")
            return self.configure_deduplication()
    
    def process_data(self, files: List[str], selected_fields: List[str], 
                    deduplicate: bool, dedup_fields: List[str]) -> pd.DataFrame:
        """
        数据处理主函数
        
        Args:
            files: 文件列表
            selected_fields: 选中的字段
            deduplicate: 是否去重
            dedup_fields: 去重字段列表
            
        Returns:
            处理后的数据框
        """
        print(f"\n=== 步骤5: 数据处理 ===")
        all_data = []
        total_rows = 0
        
        print("🔄 开始处理文件...")
        
        for i, file in enumerate(files, 1):
            try:
                print(f"\n📄 处理文件 {i}/{len(files)}: {os.path.basename(file)}")
                df = pd.read_excel(file)
                
                # 检查文件是否包含所有选中字段
                missing_fields = [field for field in selected_fields if field not in df.columns]
                if missing_fields:
                    print(f"⚠️  警告：文件缺少字段 {missing_fields}，跳过此文件")
                    continue
                
                # 提取选中的字段
                selected_data = df[selected_fields].copy()
                
                all_data.append(selected_data)
                file_rows = len(selected_data)
                total_rows += file_rows
                print(f"✅ 成功读取 {file_rows} 行数据")
                
            except Exception as e:
                print(f"❌ 错误：处理文件 '{os.path.basename(file)}' 时出错: {str(e)}")
                continue
        
        if not all_data:
            print("❌ 没有成功读取任何数据")
            return pd.DataFrame()
        
        # 合并所有数据
        print(f"\n🔄 正在合并数据...")
        combined_df = pd.concat(all_data, ignore_index=True)
        print(f"✅ 合并完成，总行数: {len(combined_df)}")
        
        # 去重处理
        if deduplicate and dedup_fields:
            print(f"\n🔄 正在按字段 {dedup_fields} 去重...")
            before_count = len(combined_df)
            combined_df = combined_df.drop_duplicates(subset=dedup_fields, keep='first')
            after_count = len(combined_df)
            removed_count = before_count - after_count
            
            print(f"✅ 去重完成:")
            print(f"  📊 去重前行数: {before_count}")
            print(f"  📊 去重后行数: {after_count}")
            print(f"  ��️  删除重复记录: {removed_count}")
            
            if removed_count > 0:
                print(f"  📈 去重率: {removed_count/before_count*100:.1f}%")
        
        return combined_df
    
    def export_to_excel(self, df: pd.DataFrame, output_filename: str = None):
        """
        导出到Excel
        
        Args:
            df: 数据框
            output_filename: 输出文件名
        """
        if output_filename is None:
            output_filename = self.output_filename
        
        print(f"\n=== 步骤6: 导出结果 ===")
        
        # 设置输出路径
        output_path = output_filename
        if not os.path.dirname(output_path):
            output_path = os.path.join(".", output_path)
        
        # 检查文件是否已存在
        if os.path.exists(output_path):
            print(f"⚠️  文件 '{output_filename}' 已存在")
            overwrite = input("是否覆盖？(y/n，默认n): ").strip().lower()
            if overwrite not in ['y', 'yes', '是']:
                # 生成新文件名
                base_name = os.path.splitext(output_filename)[0]
                extension = os.path.splitext(output_filename)[1]
                counter = 1
                while True:
                    new_filename = f"{base_name}_{counter}{extension}"
                    new_output_path = os.path.join(".", new_filename)
                    if not os.path.exists(new_output_path):
                        output_path = new_output_path
                        output_filename = new_filename
                        print(f"📝 使用新文件名: {new_filename}")
                        break
                    counter += 1
            else:
                # 尝试删除已存在的文件
                try:
                    os.remove(output_path)
                    print(f"✅ 已删除已存在的文件: {output_filename}")
                except PermissionError:
                    print(f"❌ 无法删除文件 '{output_filename}'，文件可能被其他程序占用")
                    print("请关闭Excel或其他可能打开该文件的程序，然后重试")
                    return None
                except Exception as e:
                    print(f"❌ 删除文件时出错: {str(e)}")
                    return None
        
        try:
            # 创建Excel写入器，支持多个工作表
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 主数据表
                df.to_excel(writer, sheet_name='合并数据', index=False)
                
                # 统计信息表
                stats_data = {
                    '统计项目': [
                        '总记录数',
                        '处理文件数',
                        '选择字段数',
                        '是否去重',
                        '去重字段数',
                        '删除重复记录数',
                        '处理时间'
                    ],
                    '数值': [
                        len(df),
                        len(self.selected_files),
                        len(self.selected_fields),
                        '是' if self.deduplicate else '否',
                        len(self.dedup_fields) if self.deduplicate else 0,
                        len(df) - len(df.drop_duplicates(subset=self.dedup_fields)) if self.deduplicate and self.dedup_fields else 0,
                        pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='处理统计', index=False)
                
                # 字段信息表
                field_info = {
                    '字段名称': self.selected_fields,
                    '字段类型': [str(df[field].dtype) for field in self.selected_fields],
                    '非空值数量': [df[field].notna().sum() for field in self.selected_fields],
                    '空值数量': [df[field].isna().sum() for field in self.selected_fields]
                }
                field_df = pd.DataFrame(field_info)
                field_df.to_excel(writer, sheet_name='字段信息', index=False)
            
            print(f"✅ 数据已成功导出到: {output_path}")
            print(f"�� 总共导出 {len(df)} 条记录")
            print(f"📋 包含工作表: 合并数据、处理统计、字段信息")
            
            return output_path
            
        except PermissionError:
            print(f"❌ 权限错误：无法保存到 {output_path}")
            print("请确保文件没有被其他程序（如Excel）打开")
            print("建议：")
            print("1. 关闭可能打开该文件的Excel程序")
            print("2. 使用不同的文件名")
            print("3. 检查文件是否被设置为只读")
            return None
        except Exception as e:
            print(f"❌ 导出文件时出错: {str(e)}")
            return None
    
    def set_output_filename(self):
        """设置输出文件名"""
        print(f"\n=== 步骤4.5: 输出设置 ===")
        print(f"📝 当前输出文件名: {self.output_filename}")
        filename = input("请输入新的输出文件名（或按回车保持默认）: ").strip()
        if filename:
            # 确保文件扩展名正确
            if not filename.endswith(('.xlsx', '.xls')):
                filename += '.xlsx'
            self.output_filename = filename
        print(f"✅ 输出文件名: {self.output_filename}")
    

    
    def run(self):
        """运行主程序"""
        print("=" * 60)
        print("🎯 Excel文件处理工具 v2.0")
        print("�� 功能：多文件数据合并、字段选择、去重处理")
        print("=" * 60)
        
        try:
            # 1. 文件选择
            folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用当前目录）: ").strip()
            if not folder_path:
                folder_path = "."
            
            files = self.select_files(folder_path)
            if not files:
                print("❌ 未选择任何文件，程序退出")
                return
            
            # 2. 字段分析
            all_fields = self.get_field_list(files)
            if not all_fields:
                print("❌ 未找到任何字段，程序退出")
                return
            
            # 3. 字段选择
            selected_fields = self.select_fields(all_fields)
            if not selected_fields:
                print("❌ 未选择任何字段，程序退出")
                return
            
            # 4. 去重配置
            deduplicate, dedup_fields = self.configure_deduplication()
            
            # 4.5. 输出设置
            self.set_output_filename()
            
            # 5. 数据处理
            result_df = self.process_data(files, selected_fields, deduplicate, dedup_fields)
            if result_df.empty:
                print("❌ 没有数据可处理，程序退出")
                return
            
            # 6. 导出结果
            output_path = self.export_to_excel(result_df)
            
            if output_path:
                print(f"\n" + "=" * 60)
                print("🎉 处理完成！")
                print("=" * 60)
                print(f"�� 结果文件: {output_path}")
                print(f"�� 处理记录数: {len(result_df)}")
                print(f"📁 处理文件数: {len(files)}")
                print(f"📋 选择字段数: {len(selected_fields)}")
                if deduplicate and dedup_fields:
                    print(f"�� 去重字段: {', '.join(dedup_fields)}")
                

            
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n❌ 程序执行出错: {str(e)}")

def main():
    """主函数"""
    processor = ExcelProcessor()
    processor.run()

if __name__ == "__main__":
    main()
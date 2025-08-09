import pandas as pd
import os
import glob
import re
from typing import List, Tuple, Dict, Optional

class ExcelProcessor:
    """Excel文件处理工具"""
    
    def __init__(self):
        self.selected_files = []
        self.all_fields = []
        self.selected_fields = []
        self.deduplicate = False
        self.dedup_fields = []
        self.output_filename = "result.xlsx"
        
        # 学生姓名补充功能相关属性（旧版本，保留兼容性）
        self.enable_name_supplement = False
        self.student_name_mapping = {}  # 学号到学生姓名的映射
        self.default_student_name = "未知学生"
        self.supplement_stats = {
            'total_supplemented': 0,
            'successful_matches': 0,
            'default_value_used': 0
        }
        



        # 同步模式相关属性
        self.operation_mode = "merge"  # "merge" or "sync"
        self.source_file = ""  # 源文件路径变量
        self.target_file = ""  # 目标文件路径变量
        # 保持有意义的默认关联字段，避免覆盖为空
        self.update_fields = []  # 更新字段列表变量
        self.output_directory = ""  # 输出目录变量
        self.output_file_path = None  # 完整的输出文件路径（如果用户指定）
        self.unmatched_handling = "empty"  # 未匹配记录处理方式: "empty" 或 "default"
        
        # 替换确认相关属性
        self.replacement_mode = "ask"  # "ask", "replace_all", "skip_all"
        self.replacement_decisions = {}  # 存储用户对特定字段的决策
        
        # 默认值相关属性
        self.default_values = {}  # 存储每个字段的默认值

        self.sync_stats = {
            'source_records': 0,
            'target_records': 0,
            'updated_records': 0,
            'failed_records': 0,
            'sync_success_rate': 0.0
        }
        # 自定义字段别名映射：用于将高相似度的字段视为同一逻辑字段
        # 形如 { canonical_field: [alias_field_1, alias_field_2, ...] }
        self.custom_field_aliases: Dict[str, List[str]] = {}
        
        # 多源同步字段映射：记录目标字段到各源文件字段的映射关系
        # 形如 { source_file_name: { target_field: source_field } }
        self.source_field_mapping: Dict[str, Dict[str, str]] = {}
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度
        
        Args:
            str1: 字符串1
            str2: 字符串2
            
        Returns:
            相似度 (0-1)
        """
        from difflib import SequenceMatcher
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
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
        print("- 输入文件编号（用逗号分隔，如：1,2,3）")
        print("- 输入 'all' 选择所有文件")
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
    
    def analyze_student_name_situation(self, files: List[str]) -> Dict:
        """
        分析学生姓名补充情况
        
        Args:
            files: 文件列表
            
        Returns:
            分析结果字典
        """
        analysis_result = {
            'files_with_both': [],  # 同时包含学号和姓名的文件
            'files_missing_name': [],  # 包含学号但缺少姓名的文件
            'files_without_student_id': [],  # 不包含学号的文件
            'total_files': len(files)
        }
        
        print(f"\n🔍 分析学生姓名补充情况...")
        
        for file in files:
            try:
                df = pd.read_excel(file)
                file_fields = list(df.columns)
                filename = os.path.basename(file)
                
                # 支持多种学号字段名称
                has_student_id = any(id_field in file_fields for id_field in ['学号', '*学号'])
                # 支持多种学生姓名字段名称
                has_student_name = any(name in file_fields for name in ['学生姓名', '*学生姓名'])
                
                if has_student_id and has_student_name:
                    analysis_result['files_with_both'].append(file)
                    print(f"✅ {filename}: 包含学号和姓名")
                elif has_student_id and not has_student_name:
                    analysis_result['files_missing_name'].append(file)
                    print(f"⚠️  {filename}: 包含学号但缺少姓名")
                else:
                    analysis_result['files_without_student_id'].append(file)
                    print(f"ℹ️  {filename}: 不包含学号")
                    
            except Exception as e:
                print(f"❌ 分析文件 '{os.path.basename(file)}' 时出错: {str(e)}")
                analysis_result['files_without_student_id'].append(file)
        
        return analysis_result
    
    def build_student_name_mapping(self, files_with_both: List[str]) -> Dict[str, str]:
        """
        构建学号到学生姓名的映射
        
        Args:
            files_with_both: 同时包含学号和姓名的文件列表
            
        Returns:
            学号到学生姓名的映射字典
        """
        if not files_with_both:
            return {}
        
        print(f"\n🔄 构建学号到学生姓名的映射...")
        mapping = {}
        total_mappings = 0
        
        for file in files_with_both:
            try:
                df = pd.read_excel(file)
                filename = os.path.basename(file)
                
                # 确定学号字段名称
                student_id_field = None
                for id_field in ['学号', '*学号']:
                    if id_field in df.columns:
                        student_id_field = id_field
                        break
                
                if not student_id_field:
                    print(f"⚠️  文件 '{filename}' 缺少学号字段，跳过")
                    continue
                
                # 确定学生姓名字段名称
                student_name_field = None
                for name_field in ['学生姓名', '*学生姓名']:
                    if name_field in df.columns:
                        student_name_field = name_field
                        break
                
                if not student_name_field:
                    print(f"⚠️  文件 '{filename}' 缺少学生姓名字段，跳过")
                    continue
                
                # 构建映射关系
                file_mappings = 0
                for _, row in df.iterrows():
                    student_id = str(row[student_id_field]).strip()
                    student_name = str(row[student_name_field]).strip()
                    
                    # 跳过空值
                    if pd.isna(student_id) or pd.isna(student_name) or student_id == '' or student_name == '':
                        continue
                    
                    # 如果学号已存在，优先使用第一个匹配的姓名
                    if student_id not in mapping:
                        mapping[student_id] = student_name
                        file_mappings += 1
                
                total_mappings += file_mappings
                print(f"📊 {filename}: 添加了 {file_mappings} 个映射关系")
                
            except Exception as e:
                print(f"❌ 处理文件 '{os.path.basename(file)}' 时出错: {str(e)}")
                continue
        
        print(f"✅ 总共构建了 {total_mappings} 个学号-姓名映射关系")
        return mapping
    
    def configure_name_supplement(self, analysis_result: Dict) -> Tuple[bool, str]:
        """
        配置学生姓名补充功能
        
        Args:
            analysis_result: 分析结果
            
        Returns:
            (是否启用补充功能, 默认学生姓名)
        """
        files_missing_name = analysis_result['files_missing_name']
        files_with_both = analysis_result['files_with_both']
        
        if not files_missing_name:
            print(f"\n✅ 所有文件都包含学生姓名字段，无需补充")
            return False, ""
        
        if not files_with_both:
            print(f"\n⚠️  没有找到包含学号和姓名的文件，无法构建映射关系")
            print(f"📝 建议：至少需要一个包含学号和姓名的文件来构建映射关系")
            return False, ""
        
        print(f"\n=== 学生姓名补充配置 ===")
        print(f"📊 分析结果:")
        print(f"  • 包含学号和姓名的文件: {len(files_with_both)} 个")
        print(f"  • 缺少学生姓名的文件: {len(files_missing_name)} 个")
        print(f"  • 不包含学号的文件: {len(analysis_result['files_without_student_id'])} 个")
        
        print(f"\n🤔 检测到部分文件缺少学生姓名字段，是否启用学生姓名补充功能？")
        print(f"📝 补充功能将从其他文件中根据学号匹配获取学生姓名")
        
        choice = input("请选择 (y/n，默认y): ").strip().lower()
        enable_supplement = choice not in ['n', 'no', '否']
        
        if not enable_supplement:
            print(f"✅ 已选择不启用学生姓名补充功能")
            return False, ""
        
        # 设置默认学生姓名
        print(f"\n📝 请输入未找到匹配学生姓名时使用的默认值")
        default_name = input(f"默认值（默认：{self.default_student_name}）: ").strip()
        if not default_name:
            default_name = self.default_student_name
        
        print(f"✅ 已设置默认学生姓名: {default_name}")
        return True, default_name
    
    def supplement_student_names(self, df: pd.DataFrame, mapping: Dict[str, str], 
                               default_name: str) -> pd.DataFrame:
        """
        为数据框补充学生姓名
        
        Args:
            df: 数据框
            mapping: 学号到学生姓名的映射
            default_name: 默认学生姓名
            
        Returns:
            补充后的数据框
        """
        # 确定学号字段名称
        student_id_field = None
        for id_field in ['学号', '*学号']:
            if id_field in df.columns:
                student_id_field = id_field
                break
        
        if not student_id_field:
            print(f"⚠️  数据框不包含学号字段，无法补充学生姓名")
            return df
        
        # 确定学生姓名字段名称
        student_name_field = None
        for name_field in ['学生姓名', '*学生姓名']:
            if name_field in df.columns:
                student_name_field = name_field
                break
        
        # 如果已经有学生姓名字段，先检查是否需要补充
        if student_name_field:
            # 检查是否有空的学生姓名
            missing_names = df[student_name_field].isna() | (df[student_name_field].astype(str).str.strip() == '')
            if not missing_names.any():
                print(f"✅ 学生姓名字段已完整，无需补充")
                return df
        
        # 创建学生姓名字段（如果不存在）
        if not student_name_field:
            student_name_field = '学生姓名'  # 默认使用标准名称
            df[student_name_field] = default_name
            print(f"📝 创建学生姓名字段")
        
        # 补充学生姓名
        supplemented_count = 0
        successful_matches = 0
        default_used = 0
        
        # 过滤掉学号为空的记录
        before_filter = len(df)
        df = df.dropna(subset=[student_id_field])
        after_filter = len(df)
        if before_filter > after_filter:
            print(f"⚠️  过滤掉 {before_filter - after_filter} 条学号为空的记录")
        
        for idx, row in df.iterrows():
            student_id = str(row[student_id_field]).strip()
            
            # 跳过空学号（双重检查）
            if pd.isna(student_id) or student_id == '':
                continue
            
            # 检查当前学生姓名是否为空
            current_name = str(row[student_name_field]).strip()
            if pd.isna(current_name) or current_name == '' or current_name == default_name:
                # 尝试从映射中获取学生姓名（精确匹配）
                if student_id in mapping:
                    df.at[idx, student_name_field] = mapping[student_id]
                    successful_matches += 1
                else:
                    # 尝试正则匹配（支持一位字符的模糊匹配）
                    matched_name = None
                    for map_id, map_name in mapping.items():
                        # 如果学号长度相同，尝试一位字符的模糊匹配
                        if len(student_id) == len(map_id):
                            # 计算不同字符的数量
                            diff_count = sum(1 for a, b in zip(student_id, map_id) if a != b)
                            if diff_count <= 1:  # 允许一位字符的差异
                                matched_name = map_name
                                break
                    
                    if matched_name:
                        df.at[idx, student_name_field] = matched_name
                        successful_matches += 1
                    else:
                        df.at[idx, student_name_field] = default_name
                        default_used += 1
                    supplemented_count += 1
        
        # 更新统计信息
        self.supplement_stats['total_supplemented'] += supplemented_count
        self.supplement_stats['successful_matches'] += successful_matches
        self.supplement_stats['default_value_used'] += default_used
        
        if supplemented_count > 0:
            print(f"📊 补充统计: 成功匹配 {successful_matches} 个，使用默认值 {default_used} 个")
        
        return df
    
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
        print("- 输入字段编号（用逗号分隔，如：1,2,3）")
        print("- 输入 'all' 选择所有字段")
        print("- 输入 'page 1' 查看第1页（可替换页码）")
        
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
        dedup_choice = input("请选择 (y/n，默认y): ").strip().lower()
        self.deduplicate = dedup_choice not in ['n', 'no', '否']
        
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
        
        print(f"\n- 输入字段编号（用逗号分隔，如：1,2）")
        print(f"- 输入 'all' 使用所有选中字段进行去重")
        print(f"- 输入 'single 1' 只使用第1个字段去重")
        
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
                
                # 检查文件是否包含所有选中字段，支持学号和学生姓名字段的变体
                missing_fields = []
                for field in selected_fields:
                    if field not in df.columns:
                        # 如果是学号字段，检查是否有变体
                        if field == '学号' and '*学号' in df.columns:
                            continue  # 有*学号变体，不算缺失
                        elif field == '*学号' and '学号' in df.columns:
                            continue  # 有学号变体，不算缺失
                        # 如果是学生姓名字段，检查是否有变体
                        elif field == '学生姓名' and '*学生姓名' in df.columns:
                            continue  # 有*学生姓名变体，不算缺失
                        elif field == '*学生姓名' and '学生姓名' in df.columns:
                            continue  # 有学生姓名变体，不算缺失
                        missing_fields.append(field)
                
                if missing_fields:
                    print(f"⚠️  警告：文件缺少字段 {missing_fields}，跳过此文件")
                    continue
                
                # 提取选中的字段，处理学号和学生姓名字段的变体
                df_temp = df.copy()
                actual_fields = []
                
                for field in selected_fields:
                    if field in df.columns:
                        actual_fields.append(field)
                    elif field == '学号' and '*学号' in df.columns:
                        # 将*学号重命名为学号
                        df_temp['学号'] = df_temp['*学号']
                        actual_fields.append('学号')
                    elif field == '*学号' and '学号' in df.columns:
                        # 将学号重命名为*学号
                        df_temp['*学号'] = df_temp['学号']
                        actual_fields.append('*学号')
                    elif field == '学生姓名' and '*学生姓名' in df.columns:
                        # 将*学生姓名重命名为学生姓名
                        df_temp['学生姓名'] = df_temp['*学生姓名']
                        actual_fields.append('学生姓名')
                    elif field == '*学生姓名' and '学生姓名' in df.columns:
                        # 将学生姓名重命名为*学生姓名
                        df_temp['*学生姓名'] = df_temp['学生姓名']
                        actual_fields.append('*学生姓名')
                    else:
                        actual_fields.append(field)
                
                selected_data = df_temp[actual_fields].copy()
                
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
        
        # 过滤掉学号为空的记录
        student_id_fields = [col for col in combined_df.columns if '学号' in col]
        if student_id_fields:
            before_filter = len(combined_df)
            combined_df = combined_df.dropna(subset=student_id_fields)
            after_filter = len(combined_df)
            if before_filter > after_filter:
                print(f"⚠️  过滤掉 {before_filter - after_filter} 条学号为空的记录")
                print(f"✅ 过滤后总行数: {len(combined_df)}")
        

        
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
            print(f"  删除重复记录: {removed_count}")
            
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
                stats_items = [
                    '总记录数',
                    '处理文件数',
                    '选择字段数',
                    '是否去重',
                    '去重字段数',
                    '删除重复记录数'
                ]
                stats_values = [
                    len(df),
                    len(self.selected_files),
                    len(self.selected_fields),
                    '是' if self.deduplicate else '否',
                    len(self.dedup_fields) if self.deduplicate else 0,
                    len(df) - len(df.drop_duplicates(subset=self.dedup_fields)) if self.deduplicate and self.dedup_fields else 0
                ]
                

                
                stats_items.append('处理时间')
                stats_values.append(pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S'))
                
                stats_data = {
                    '统计项目': stats_items,
                    '数值': stats_values
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
            print(f"总共导出 {len(df)} 条记录")
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
        filename = input("请输入新的输出文件名列如G:\\wang\\excel（默认格式为xlsx）: ").strip()
        if filename:
            # 确保文件扩展名正确
            if not filename.endswith(('.xlsx', '.xls')):
                filename += '.xlsx'
            self.output_filename = filename
        print(f"✅ 输出文件名: {self.output_filename}")
    

    
    def run(self):
        """运行主程序"""
        print("=" * 60)
        print("🎯 Excel文件处理工具 v2.4")
        print("📋 功能：多文件数据合并、字段选择、去重处理、学生姓名补充、单源同步、多源同步")
        print("=" * 60)
        
        # 选择操作模式
        mode = self.select_operation_mode()
        
        if mode == "merge":
            self.run_merge_mode()
        elif mode == "multi_sync":
            self.run_multi_sync_mode()
        else:
            print("👋 程序退出")
    
    def run_sync_only(self):
        """运行同步功能（专门用于excel_tool.py调用）"""
        print("=" * 60)
        print("🎯 Excel数据同步工具")
        print("📋 功能：将一个或者多个Excel文件的数据同步到另一个文件")
        print("💡 提示：支持单个或多个源文件同步到目标文件")
        print("=" * 60)
        
        # 直接使用多源同步模式（支持单个或多个源文件）
        self.run_multi_sync_mode()
    
    def select_operation_mode(self) -> str:
        """
        选择操作模式
        
        Returns:
            str: 操作模式 ("merge", "sync", 或 "multi_sync")
        """
        print("\n请选择操作模式：")
        print("1. 合并到空Excel（创建新的合并文件）")
        print("2. 同步到有数据的Excel（更新现有文件）")
        print("   - 2.1 源文件到目标文件（单个源文件同步）")
        print("   - 2.2 多个源文件到目标文件（多个源文件同步）")
        
        while True:
            choice = input("\n请选择 (1/2): ").strip()
            if choice == "1":
                print("✅ 已选择：合并模式")
                return "merge"
            elif choice == "2":
                print("✅ 已选择：Excel数据同步")
                print("💡 提示：支持单个或多个源文件同步到目标文件")
                return "multi_sync"
            else:
                print("❌ 无效选择，请输入 1 或 2")
    

    def run_merge_mode(self):
        """运行合并模式"""
        print("\n🔄 启动合并模式...")
        
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
                print(f"📄 结果文件: {output_path}")
                print(f"📊 处理记录数: {len(result_df)}")
                print(f"📁 处理文件数: {len(files)}")
                print(f"📋 选择字段数: {len(selected_fields)}")
                if deduplicate and dedup_fields:
                    print(f"🔍 去重字段: {', '.join(dedup_fields)}")

            
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n❌ 程序执行出错: {str(e)}")
    
    def run_sync_mode(self):
        """运行同步模式"""
        print("\n🔄 启动同步模式...")
        
        try:
            # 1. 文件角色选择
            self.select_file_roles()
            
            # 1.5. 文件备份
            if not self.backup_sync_files():
                print("❌ 备份失败，程序退出")
                return
            
            # 2. 关联字段选择
            self.select_link_field()
            
            # 3. 更新字段选择
            self.select_update_fields()
            
            # 3.5. 输出目录设置
            self.set_output_directory()
            
            # 3.6. 未匹配记录处理配置
            self.configure_unmatched_handling()
            
            # 4. 执行同步
            self.execute_sync()
            
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n❌ 程序执行出错: {str(e)}")
    
    def select_file_roles(self):
        """文件角色选择"""
        print(f"\n=== 步骤1: 文件角色选择 ===")
        
        # 选择文件夹
        folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用默认目录G:\\wang\\excel）: ").strip()
        if not folder_path:
            folder_path = "G:\\wang\\excel"
        
        # 扫描Excel文件
        excel_patterns = ['*.xlsx', '*.xls']
        excel_files = []
        
        for pattern in excel_patterns:
            excel_files.extend(glob.glob(os.path.join(folder_path, pattern)))
        
        if not excel_files:
            print(f"❌ 在文件夹 '{folder_path}' 中没有找到Excel文件")
            return
        
        # 显示文件列表
        print(f"\n✅ 找到 {len(excel_files)} 个Excel文件:")
        for i, file in enumerate(excel_files, 1):
            filename = os.path.basename(file)
            file_size = os.path.getsize(file) / 1024  # KB
            print(f"{i:2d}. {filename:<30} ({file_size:.1f} KB)")
        
        # 选择源文件
        print(f"\n📋 请选择源文件（提供数据的文件）:")
        while True:
            try:
                source_choice = input("请输入源文件编号: ").strip()
                source_index = int(source_choice) - 1
                if 0 <= source_index < len(excel_files):
                    self.source_file = excel_files[source_index]
                    print(f"✅ 源文件: {os.path.basename(self.source_file)}")
                    break
                else:
                    print("❌ 文件编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字")
        
        # 选择目标文件
        print(f"\n📋 请选择目标文件（需要更新的文件）:")
        while True:
            try:
                target_choice = input("请输入目标文件编号: ").strip()
                target_index = int(target_choice) - 1
                if 0 <= target_index < len(excel_files):
                    if target_index == source_index:
                        print("❌ 目标文件不能与源文件相同，请重新选择")
                        continue
                    self.target_file = excel_files[target_index]
                    print(f"✅ 目标文件: {os.path.basename(self.target_file)}")
                    break
                else:
                    print("❌ 文件编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字")
    
    def select_link_field(self):
        """关联字段选择"""
        print(f"\n=== 步骤2: 关联字段选择 ===")
        
        try:
            # 读取源文件和目标文件
            source_df = pd.read_excel(self.source_file)
            target_df = pd.read_excel(self.target_file)
            
            # 获取两个文件的列名
            source_columns = list(source_df.columns)
            target_columns = list(target_df.columns)
            
            # 找出共有的字段
            common_fields = list(set(source_columns) & set(target_columns))
            
            if not common_fields:
                print("❌ 源文件和目标文件没有共同的字段，无法进行同步")
                return
            
            print(f"📋 源文件和目标文件共有的字段:")
            for i, field in enumerate(common_fields, 1):
                print(f"{i:2d}. {field}")
            
            # 选择关联字段
            print(f"\n🔗 请选择用于关联记录的字段（如ID、姓名等唯一标识字段）:")
            while True:
                try:
                    link_choice = input("请输入关联字段编号: ").strip()
                    link_index = int(link_choice) - 1
                    if 0 <= link_index < len(common_fields):
                        self.link_field = common_fields[link_index]
                        print(f"✅ 关联字段: {self.link_field}")
                        break
                    else:
                        print("❌ 字段编号超出范围，请重新选择")
                except ValueError:
                    print("❌ 请输入有效的数字")
                    
        except Exception as e:
            print(f"❌ 读取文件时出错: {str(e)}")
    
    def set_output_directory(self):
        """设置输出目录"""
        print(f"\n=== 步骤3.5: 输出目录设置 ===")
        
        # 获取目标文件所在目录作为默认目录
        default_dir = os.path.dirname(self.target_file)
        print(f"📁 当前目标文件目录: {default_dir}")
        
        output_dir = input("请输入输出目录路径（或按回车使用目标文件所在目录）: ").strip()
        if not output_dir:
            output_dir = default_dir
        
        # 检查目录是否存在，如果不存在则创建
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"✅ 已创建输出目录: {output_dir}")
            except Exception as e:
                print(f"❌ 创建目录失败: {str(e)}")
                print(f"📁 使用默认目录: {default_dir}")
                output_dir = default_dir
        else:
            print(f"✅ 输出目录: {output_dir}")
        
        self.output_directory = output_dir
    
    def ask_for_replacement(self, field_name: str, link_value: str, current_value, new_value) -> bool:
        """
        询问用户是否要替换已有数据
        
        Args:
            field_name: 字段名称
            link_value: 关联字段的值
            current_value: 当前值
            new_value: 新值
            
        Returns:
            是否要替换
        """
        if self.replacement_mode == "replace_all":
            return True
        elif self.replacement_mode == "skip_all":
            return False
        elif self.replacement_mode == "ask":
            # 检查是否已经对此字段做过决策
            decision_key = f"{field_name}"
            if decision_key in self.replacement_decisions:
                return self.replacement_decisions[decision_key]
            
            print(f"\n{'🔄' + '='*60}")
            print(f"⚠️  发现数据冲突！")
            print(f"{'🔄' + '='*60}")
            print(f"🔑 关联字段值: {link_value}")
            print(f"📝 字段名称: {field_name}")
            print(f"📄 当前值: {current_value}")
            print(f"🆕 新值: {new_value}")
            
            print(f"\n🤔 请选择处理方式:")
            print(f"  1. 替换（使用新值）")
            print(f"  2. 跳过（保留当前值）")
            print(f"  3. 对此字段总是替换")
            print(f"  4. 对此字段总是跳过")
            print(f"  5. 对所有字段总是替换")
            print(f"  6. 对所有字段总是跳过")
            
            while True:
                try:
                    choice = input("\n请选择 (1-6): ").strip()
                    
                    if choice == "1":
                        return True
                    elif choice == "2":
                        return False
                    elif choice == "3":
                        self.replacement_decisions[decision_key] = True
                        print(f"✅ 已设置：字段 '{field_name}' 总是替换")
                        return True
                    elif choice == "4":
                        self.replacement_decisions[decision_key] = False
                        print(f"✅ 已设置：字段 '{field_name}' 总是跳过")
                        return False
                    elif choice == "5":
                        self.replacement_mode = "replace_all"
                        print(f"✅ 已设置：所有字段总是替换")
                        return True
                    elif choice == "6":
                        self.replacement_mode = "skip_all"
                        print(f"✅ 已设置：所有字段总是跳过")
                        return False
                    else:
                        print("❌ 请输入 1-6 之间的数字")
                        
                except KeyboardInterrupt:
                    print("\n⚠️  用户中断，跳过此次替换")
                    return False
        
        return False

    def configure_unmatched_handling(self):
        """配置未匹配记录的处理方式"""
        print(f"\n=== 步骤3.6: 未匹配记录处理配置 ===")
        
        print("🤔 对于匹配不上的记录，您希望如何处理？")
        print("1. 设置为空值（保持原有数据不变）")
        print("2. 使用默认值（为每个字段设置默认值）")
        
        while True:
            choice = input("请选择处理方式 (1/2): ").strip()
            if choice == "1":
                self.unmatched_handling = "empty"
                print("✅ 已选择：未匹配记录设置为空值")
                break
            elif choice == "2":
                self.unmatched_handling = "default"
                print("✅ 已选择：未匹配记录使用默认值")
                
                # 配置每个更新字段的默认值
                print(f"\n📝 请为每个更新字段设置默认值:")
                for field in self.update_fields:
                    while True:
                        default_value = input(f"请输入字段 '{field}' 的默认值（直接回车表示空值）: ").strip()
                        # 允许空值作为默认值
                        self.default_values[field] = default_value if default_value else ""
                        print(f"✅ 字段 '{field}' 的默认值已设置为: '{self.default_values[field]}'")
                        break
                
                print(f"\n📋 默认值配置完成:")
                for field, value in self.default_values.items():
                    display_value = value if value else "<空值>"
                    print(f"  • {field}: {display_value}")
                break
            else:
                print("❌ 无效选择，请输入 1 或 2")
        
        # 配置替换模式
        print(f"\n=== 步骤3.7: 数据替换策略配置 ===")
        print("🤔 当目标Excel中已有数据时，您希望如何处理？")
        print("1. 每次询问是否替换（推荐）")
        print("2. 自动替换所有数据")
        print("3. 跳过所有已有数据")
        
        while True:
            choice = input("请选择替换策略 (1/2/3): ").strip()
            if choice == "1":
                self.replacement_mode = "ask"
                print("✅ 已选择：遇到已有数据时询问是否替换")
                break
            elif choice == "2":
                self.replacement_mode = "replace_all"
                print("✅ 已选择：自动替换所有已有数据")
                break
            elif choice == "3":
                self.replacement_mode = "skip_all"
                print("✅ 已选择：跳过所有已有数据")
                break
            else:
                print("❌ 无效选择，请输入 1、2 或 3")
    

    
    def select_update_fields(self):
        """更新字段选择"""
        print(f"\n=== 步骤3: 更新字段选择 ===")
        
        try:
            # 读取源文件和目标文件
            source_df = pd.read_excel(self.source_file)
            target_df = pd.read_excel(self.target_file)
            
            # 获取两个文件的列名
            source_columns = list(source_df.columns)
            target_columns = list(target_df.columns)
            
            # 显示源文件的所有字段供用户选择
            print(f"📋 源文件中的所有字段:")
            for i, field in enumerate(source_columns, 1):
                # 标记哪些字段在目标文件中已存在
                status = "（目标文件中已存在）" if field in target_columns else "（目标文件中不存在）"
                print(f"{i:2d}. {field} {status}")
            
            # 选择更新字段
            print(f"\n📝 请选择需要从源文件同步到目标文件的字段:")
            print("💡 输入字段编号（用逗号分隔，如：1,2,3）")
            print("💡 输入 'all' 选择所有字段")
            print("💡 注意：如果字段在目标文件中已存在，将会覆盖原有数据")
            
            while True:
                choice = input("请选择: ").strip().lower()
                
                if choice == 'all':
                    self.update_fields = source_columns
                    print(f"✅ 已选择所有 {len(source_columns)} 个字段")
                    break
                else:
                    try:
                        indices = [int(x.strip()) - 1 for x in choice.split(',')]
                        selected_fields = [source_columns[i] for i in indices if 0 <= i < len(source_columns)]
                        
                        if not selected_fields:
                            print("❌ 未选择任何有效字段，请重新选择")
                            continue
                        
                        self.update_fields = selected_fields
                        print(f"✅ 已选择 {len(selected_fields)} 个字段:")
                        for field in selected_fields:
                            status = "（将覆盖目标文件中的现有数据）" if field in target_columns else "（将添加到目标文件中）"
                            print(f"  📋 {field} {status}")
                        break
                        
                    except (ValueError, IndexError):
                        print("❌ 输入格式错误，请重新选择")
                        
        except Exception as e:
            print(f"❌ 读取文件时出错: {str(e)}")
    
    def execute_sync(self):
        """执行同步操作"""
        print(f"\n=== 步骤4: 执行同步 ===")
        
        try:
            # 读取源文件和目标文件
            source_df = pd.read_excel(self.source_file)
            target_df = pd.read_excel(self.target_file)
            
            # 统计记录数
            self.sync_stats['source_records'] = len(source_df)
            self.sync_stats['target_records'] = len(target_df)
            
            print(f"📊 源文件记录数: {self.sync_stats['source_records']}")
            print(f"📊 目标文件记录数: {self.sync_stats['target_records']}")
            print(f"🔗 关联字段: {self.link_field}")
            print(f"📝 更新字段: {', '.join(self.update_fields)}")
            
            # 确认执行
            confirm = input(f"\n是否确认执行同步操作？(y/n): ").strip().lower()
            if confirm not in ['y', 'yes', '是']:
                print("❌ 用户取消操作")
                return
            
            # 执行同步
            updated_df = self.perform_sync(source_df, target_df)
            
            # 保存更新后的文件
            self.save_updated_file(updated_df)
            
            # 显示同步结果
            self.show_sync_results()
            
        except Exception as e:
            print(f"❌ 同步执行出错: {str(e)}")
    
    def find_actual_field_name(self, df: pd.DataFrame, target_field: str) -> str:
        """
        查找数据框中实际的字段名称，支持智能匹配
        
        Args:
            df: 数据框
            target_field: 目标字段名
            
        Returns:
            实际的字段名称，如果找不到返回None
        """
        if not hasattr(df, 'columns'):
            return None
            
        columns = list(df.columns)
        
        # 1. 精确匹配
        if target_field in columns:
            return target_field
        
        # 2. 大小写不敏感匹配
        for col in columns:
            if str(col).lower() == str(target_field).lower():
                return col
        
        # 3. 去除空格后匹配
        target_clean = str(target_field).strip()
        for col in columns:
            if str(col).strip() == target_clean:
                return col
        
        # 4. 包含匹配（目标字段包含在列名中）
        for col in columns:
            if target_clean in str(col) or str(col) in target_clean:
                return col
        
        # 5. 相似度匹配（简单版本）
        from difflib import SequenceMatcher
        best_match = None
        best_ratio = 0.8  # 相似度阈值
        
        for col in columns:
            ratio = SequenceMatcher(None, str(target_field).lower(), str(col).lower()).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = col
        
        return best_match

    def backup_sync_files(self) -> bool:
        """
        备份同步相关的文件（源文件和目标文件）
        
        Returns:
            备份是否成功
        """
        print(f"\n=== 文件备份 ===")
        
        # 询问是否要备份
        backup_choice = input("🤔 是否要备份相关Excel文件？(y/n，默认y): ").strip().lower()
        if backup_choice in ['n', 'no', '否']:
            print("✅ 跳过备份，直接执行同步")
            return True
        
        # 创建备份目录
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = f"backup_{timestamp}"
        
        try:
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            print(f"📁 创建备份目录: {backup_dir}")
            
            # 收集要备份的文件
            files_to_backup = []
            
            # 添加源文件
            if hasattr(self, 'source_file') and self.source_file:
                files_to_backup.append(('源文件', self.source_file))
            
            # 添加目标文件
            if hasattr(self, 'target_file') and self.target_file:
                files_to_backup.append(('目标文件', self.target_file))
            
            # 如果是多源同步，添加所有源文件
            if hasattr(self, 'source_files') and self.source_files:
                for i, source_file in enumerate(self.source_files, 1):
                    files_to_backup.append((f'源文件{i}', source_file))
            
            if not files_to_backup:
                print("⚠️  没有找到需要备份的文件")
                return True
            
            # 备份文件
            backup_success = 0
            backup_failed = 0
            
            for file_type, file_path in files_to_backup:
                try:
                    if not os.path.exists(file_path):
                        print(f"⚠️  {file_type}不存在: {file_path}")
                        continue
                    
                    filename = os.path.basename(file_path)
                    backup_path = os.path.join(backup_dir, f"{file_type}_{filename}")
                    
                    # 如果备份目录中已有同名文件，添加序号
                    counter = 1
                    original_backup_path = backup_path
                    while os.path.exists(backup_path):
                        name, ext = os.path.splitext(original_backup_path)
                        backup_path = f"{name}_{counter}{ext}"
                        counter += 1
                    
                    # 复制文件
                    import shutil
                    shutil.copy2(file_path, backup_path)
                    print(f"✅ 已备份{file_type}: {filename} -> {os.path.basename(backup_path)}")
                    backup_success += 1
                    
                except Exception as e:
                    print(f"❌ 备份{file_type}失败: {os.path.basename(file_path)} - {str(e)}")
                    backup_failed += 1
            
            print(f"\n📊 备份结果:")
            print(f"  ✅ 成功备份: {backup_success} 个文件")
            if backup_failed > 0:
                print(f"  ❌ 备份失败: {backup_failed} 个文件")
            print(f"  📁 备份位置: {os.path.abspath(backup_dir)}")
            
            if backup_failed > 0:
                continue_choice = input("\n⚠️  部分文件备份失败，是否继续同步？(y/n，默认y): ").strip().lower()
                if continue_choice in ['n', 'no', '否']:
                    print("❌ 用户选择退出")
                    return False
            
            return True
            
        except Exception as e:
            print(f"❌ 创建备份目录失败: {str(e)}")
            continue_choice = input("⚠️  备份失败，是否继续同步？(y/n，默认n): ").strip().lower()
            return continue_choice in ['y', 'yes', '是']
    
    def perform_sync(self, source_df: pd.DataFrame, target_df: pd.DataFrame) -> pd.DataFrame:
        """执行同步操作"""
        print(f"\n🔄 正在执行同步...")
        
        # 创建目标文件的副本
        updated_df = target_df.copy()
        
        # 解析关联字段的实际名称
        actual_link_field = self.find_actual_field_name(updated_df, self.link_field) if hasattr(self, 'find_actual_field_name') else self.link_field
        if not actual_link_field or actual_link_field not in updated_df.columns:
            # 回退为原字段名尝试
            actual_link_field = self.link_field
            if actual_link_field not in updated_df.columns:
                print(f"⚠️  目标文件缺少关联字段 '{self.link_field}'，同步将仅尝试创建更新列")
        
        # 为每个更新字段添加新列（如果不存在），并解析目标/源的实际列名
        added_fields = []
        existing_fields = []
        update_pairs = []  # (target_field_in_df, source_field_in_source_df)
        for field in self.update_fields:
            actual_target_field = self.find_actual_field_name(updated_df, field) if hasattr(self, 'find_actual_field_name') else field
            if actual_target_field and actual_target_field in updated_df.columns:
                existing_fields.append(actual_target_field)
            else:
                # 目标中不存在则创建以目标命名
                actual_target_field = field
                updated_df[actual_target_field] = None
                added_fields.append(actual_target_field)

            # 源字段解析（基于原始更新字段名寻址源列变体）
            actual_source_field = self.find_actual_field_name(source_df, field) if hasattr(self, 'find_actual_field_name') else field
            update_pairs.append((actual_target_field, actual_source_field))
        
        if added_fields:
            print(f"📝 将添加新字段到目标文件: {', '.join(added_fields)}")
        if existing_fields:
            print(f"📝 将覆盖目标文件中的现有字段: {', '.join(existing_fields)}")
        
        # 构建源文件的映射关系
        source_mapping = {}
        # 解析源文件关联字段
        actual_source_link_field = self.find_actual_field_name(source_df, self.link_field) if hasattr(self, 'find_actual_field_name') else self.link_field
        if not actual_source_link_field or actual_source_link_field not in source_df.columns:
            actual_source_link_field = self.link_field

        for _, row in source_df.iterrows():
            link_value = str(row.get(actual_source_link_field, '')).strip()
            if link_value and link_value != 'nan':
                source_mapping[link_value] = row
        
        print(f"📊 源文件映射关系数量: {len(source_mapping)}")
        
        # 更新目标文件
        updated_count = 0
        failed_count = 0
        unmatched_count = 0
        
        for idx, row in updated_df.iterrows():
            link_value = str(row.get(actual_link_field, '')).strip()
            
            if link_value and link_value != 'nan' and link_value in source_mapping:
                # 找到匹配的记录，更新字段
                source_row = source_mapping[link_value]
                for target_field, source_field in update_pairs:
                    try:
                        # 处理数据类型转换，避免类型不匹配警告
                        value = source_row.get(source_field)
                        if pd.isna(value) or str(value).strip() == '':
                            continue
                        
                        # 检查目标字段是否已有值
                        current_value = updated_df.at[idx, target_field]
                        has_current_value = (
                            not pd.isna(current_value) and 
                            str(current_value).strip() != '' and 
                            str(current_value).strip().lower() != 'nan'
                        )
                        
                        # 如果目标字段已有值，检查是否需要替换
                        if has_current_value:
                            # 如果当前值和新值相同，直接跳过
                            if str(current_value).strip() == str(value).strip():
                                continue  # 值相同，无需更新
                            
                            should_replace = self.ask_for_replacement(
                                target_field, 
                                link_value, 
                                current_value, 
                                value
                            )
                            if not should_replace:
                                continue  # 跳过此字段的更新
                        
                        # 确保目标列是对象类型，以保持字符串格式
                        if updated_df[target_field].dtype in ['int64', 'float64']:
                            updated_df[target_field] = updated_df[target_field].astype('object')
                        
                        # 执行更新，保持原始字符串格式
                        updated_df.at[idx, target_field] = str(value)
                        
                    except Exception as e:
                        print(f"⚠️  更新字段 {target_field} 时出错: {str(e)}")
                        print(f"   📍 关联字段值: {link_value}")
                        print(f"   📄 当前值: {current_value}")
                        print(f"   🆕 新值: {value}")
                        print(f"   🔧 字段类型: {type(value)}")
                        failed_count += 1
                        continue
                updated_count += 1
            else:
                # 处理未匹配的记录
                if self.unmatched_handling == "default":
                    # 使用默认值
                    for target_field, _ in update_pairs:
                        try:
                            # 确保目标列是对象类型
                            if updated_df[target_field].dtype in ['int64', 'float64']:
                                updated_df[target_field] = updated_df[target_field].astype('object')
                            
                            # 使用用户配置的默认值
                            # 从字段映射中获取原始字段名
                            original_field = None
                            for orig_field, mapped_field in update_pairs:
                                if mapped_field == target_field:
                                    original_field = orig_field
                                    break
                            
                            # 获取用户设置的默认值
                            default_value = self.default_values.get(original_field, "")
                            # 确保数据类型兼容，先将列转换为object类型
                            if updated_df[target_field].dtype != 'object':
                                updated_df[target_field] = updated_df[target_field].astype('object')
                            
                            if pd.isna(default_value) or default_value == '':
                                updated_df.at[idx, target_field] = None
                            else:
                                updated_df.at[idx, target_field] = str(default_value)
                        except Exception as e:
                            print(f"⚠️  设置字段 {target_field} 默认值时出错: {str(e)}")
                            continue
                    unmatched_count += 1
                else:
                    # 设置为空值（保持原有数据不变）
                    failed_count += 1
        
        # 更新统计信息
        self.sync_stats['updated_records'] = updated_count
        self.sync_stats['failed_records'] = failed_count
        self.sync_stats['unmatched_records'] = unmatched_count
        
        target_denominator = self.sync_stats.get('target_records', 0) or len(updated_df)
        if target_denominator > 0:
            self.sync_stats['sync_success_rate'] = (updated_count / target_denominator) * 100
        
        print(f"✅ 同步完成:")
        print(f"  更新记录: {updated_count} 个")
        print(f"  未匹配记录: {unmatched_count} 个")
        print(f"  失败记录: {failed_count} 个")
        print(f"  成功率: {self.sync_stats['sync_success_rate']:.1f}%")
        
        if failed_count > 0:
            print(f"\n💡 失败记录说明:")
            print(f"  • 失败记录通常由以下原因造成:")
            print(f"    - 数据类型不匹配")
            print(f"    - 字段值包含特殊字符")
            print(f"    - 源数据格式异常")
            print(f"  • 请查看上方的详细错误信息进行排查")
            print(f"  • 成功更新的 {updated_count} 条记录已正常处理")
        
        return updated_df
    
    def save_updated_file(self, updated_df: pd.DataFrame):
        """保存更新后的文件"""
        try:
            # 检查文件是否被占用
            if os.path.exists(self.target_file):
                try:
                    # 尝试以写入模式打开文件，检查是否被占用
                    with open(self.target_file, 'r+b') as f:
                        pass
                except PermissionError:
                    print(f"❌ 目标文件被其他程序占用，无法保存")
                    print("请关闭Excel或其他可能打开该文件的程序，然后重试")
                    
                    # 询问是否保存到新文件
                    save_as_new = input("是否保存到新文件？(y/n): ").strip().lower()
                    if save_as_new in ['y', 'yes', '是']:
                        # 生成新文件名
                        filename = os.path.basename(self.target_file)
                        name, ext = os.path.splitext(filename)
                        timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                        new_filename = f"{name}_updated_{timestamp}{ext}"
                        new_path = os.path.join(self.output_directory, new_filename)
                        
                        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                            updated_df.to_excel(writer, index=False)
                        
                        print(f"✅ 已保存到新文件: {new_filename}")
                        return
                    else:
                        print("❌ 用户取消保存")
                        return
            
            # 保存到原文件
            with pd.ExcelWriter(self.target_file, engine='openpyxl') as writer:
                updated_df.to_excel(writer, index=False)
            
            print(f"✅ 目标文件已更新: {os.path.basename(self.target_file)}")
            
        except PermissionError:
            print(f"❌ 无法保存文件，文件可能被其他程序占用")
            print("自动保存到新文件...")
            
            # 生成新文件名
            filename = os.path.basename(self.target_file)
            name, ext = os.path.splitext(filename)
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            new_filename = f"{name}_updated_{timestamp}{ext}"
            new_path = os.path.join(self.output_directory, new_filename)
            
            try:
                with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                    updated_df.to_excel(writer, index=False)
                
                print(f"✅ 已保存到新文件: {new_filename}")
                # 更新目标文件路径为新的文件路径
                self.target_file = new_path
            except Exception as e2:
                print(f"❌ 保存到新文件也失败: {str(e2)}")
        except Exception as e:
            print(f"❌ 保存文件时出错: {str(e)}")
            print("尝试保存到新文件...")
            
            try:
                # 生成新文件名
                filename = os.path.basename(self.target_file)
                name, ext = os.path.splitext(filename)
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                new_filename = f"{name}_updated_{timestamp}{ext}"
                new_path = os.path.join(self.output_directory, new_filename)
                
                with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                    updated_df.to_excel(writer, index=False)
                
                print(f"✅ 已保存到新文件: {new_filename}")
            except Exception as e2:
                print(f"❌ 保存到新文件也失败: {str(e2)}")
    
    def show_sync_results(self):
        """显示同步结果"""
        print(f"\n" + "=" * 60)
        print("🎉 同步处理完成！")
        print("=" * 60)
        print(f"📊 同步统计信息：")
        print(f"源文件: {os.path.basename(self.source_file)}")
        print(f"目标文件: {os.path.basename(self.target_file)}")
        print(f"源文件记录数: {self.sync_stats['source_records']} 个")
        print(f"目标文件记录数: {self.sync_stats['target_records']} 个")
        print(f"成功更新记录: {self.sync_stats['updated_records']} 个")
        print(f"未匹配记录: {self.sync_stats.get('unmatched_records', 0)} 个")
        print(f"失败记录: {self.sync_stats['failed_records']} 个")
        print(f"同步成功率: {self.sync_stats['sync_success_rate']:.1f}%")
        print(f"关联字段: {self.link_field}")
        print(f"更新字段: {', '.join(self.update_fields)}")
        print(f"处理时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 移除同步报告保存
        # self.save_sync_report()
    
    def save_sync_report(self):
        """保存同步报告"""
        try:
            # 生成报告文件名
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f"同步处理报告_{timestamp}.xlsx"
            
            # 创建报告数据
            report_data = {
                '统计项目': [
                    '源文件',
                    '目标文件',
                    '源文件记录数',
                    '目标文件记录数',
                    '成功更新记录',
                    '失败记录',
                    '同步成功率',
                    '关联字段',
                    '更新字段',
                    '处理时间'
                ],
                '数值': [
                    os.path.basename(self.source_file),
                    os.path.basename(self.target_file),
                    f"{self.sync_stats['source_records']} 个",
                    f"{self.sync_stats['target_records']} 个",
                    f"{self.sync_stats['updated_records']} 个",
                    f"{self.sync_stats['failed_records']} 个",
                    f"{self.sync_stats['sync_success_rate']:.1f}%",
                    self.link_field,
                    ', '.join(self.update_fields),
                    pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            
            # 保存到Excel文件
            report_df = pd.DataFrame(report_data)
            report_df.to_excel(report_filename, index=False)
            
            # 移除同步报告输出信息
            # print(f"📄 同步报告已保存到: {report_filename}")
            
        except Exception as e:
            print(f"⚠️  保存同步报告时出错: {str(e)}")


    

    


    def run_multi_sync_mode(self):
        """运行多源同步模式"""
        print("\n🔄 启动多源同步模式...")
        
        try:
            # 1. 文件选择
            self.select_multi_sync_files()
            
            # 1.5. 文件备份
            if not self.backup_sync_files():
                print("❌ 备份失败，程序退出")
                return
            
            # 2. 关联字段选择
            self.select_multi_sync_link_field()
            
            # 3. 更新字段选择
            self.select_multi_sync_update_fields()
            
            # 4. 冲突处理配置
            self.configure_conflict_handling()
            
            # 5. 输出设置
            self.set_multi_sync_output()
            
            # 6. 执行多源同步
            self.execute_multi_sync()
            
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n❌ 程序执行出错: {str(e)}")
    
    def select_multi_sync_files(self):
        """多源同步文件选择"""
        print(f"\n=== 步骤1: 多源同步文件选择 ===")
        
        # 选择文件夹
        folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用默认目录G:\\wang\\excel）: ").strip()
        if not folder_path:
            folder_path = "G:\\wang\\excel"
        
        # 扫描Excel文件
        excel_patterns = ['*.xlsx', '*.xls']
        excel_files = []
        
        for pattern in excel_patterns:
            excel_files.extend(glob.glob(os.path.join(folder_path, pattern)))
        
        if not excel_files:
            print(f"❌ 在文件夹 '{folder_path}' 中没有找到Excel文件")
            return
        
        # 显示文件列表
        print(f"\n✅ 找到 {len(excel_files)} 个Excel文件:")
        for i, file in enumerate(excel_files, 1):
            filename = os.path.basename(file)
            file_size = os.path.getsize(file) / 1024  # KB
            print(f"{i:2d}. {filename:<30} ({file_size:.1f} KB)")
        
        # 选择源文件（多个）
        print(f"\n📋 请选择源文件（提供数据的文件，可多选）:")
        print("📝 输入文件编号，用逗号分隔（如：1,2,3）")
        print("📝 输入 'all' 选择所有文件作为源文件")
        
        while True:
            try:
                source_choice = input("请输入源文件编号: ").strip()
                if source_choice.lower() == 'all':
                    self.source_files = excel_files
                    print(f"✅ 已选择所有 {len(excel_files)} 个文件作为源文件")
                    break
                else:
                    source_indices = [int(x.strip()) - 1 for x in source_choice.split(',')]
                    self.source_files = [excel_files[i] for i in source_indices if 0 <= i < len(excel_files)]
                    
                    if not self.source_files:
                        print("❌ 未选择任何有效文件，请重新选择")
                        continue
                    
                    print(f"✅ 已选择 {len(self.source_files)} 个源文件:")
                    for file in self.source_files:
                        print(f"  📄 {os.path.basename(file)}")
                    break
            except ValueError:
                print("❌ 请输入有效的数字")
        
        # 选择目标文件
        print(f"\n📋 请选择目标文件（需要更新的文件）:")
        while True:
            try:
                target_choice = input("请输入目标文件编号: ").strip()
                target_index = int(target_choice) - 1
                if 0 <= target_index < len(excel_files):
                    self.target_file = excel_files[target_index]
                    print(f"✅ 目标文件: {os.path.basename(self.target_file)}")
                    break
                else:
                    print("❌ 文件编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字")
    
    def select_multi_sync_link_field(self):
        """多源同步关联字段选择 - 以目标文件为主导的新逻辑"""
        print(f"\n=== 步骤2: 关联字段选择 ===")
        
        try:
            # 读取目标文件
            target_df = pd.read_excel(self.target_file)
            target_columns = list(target_df.columns)
            
            if not target_columns:
                print("❌ 目标文件没有任何字段")
                return
            
            # 第一步：从目标文件选择关联字段
            print(f"🎯 步骤1: 从目标文件选择关联字段")
            print(f"📋 目标文件 '{os.path.basename(self.target_file)}' 的所有字段:")
            
            # 智能推荐关联字段
            priority_fields = ['学号', '教工号', '工号', '编号', 'ID', 'id', 'student_id', 'teacher_id']
            recommended_field = None
            
            for priority_field in priority_fields:
                for field in target_columns:
                    if priority_field in field or field in priority_field:
                        recommended_field = field
                        break
                if recommended_field:
                    break
            
            # 显示目标文件字段
            for i, field in enumerate(target_columns, 1):
                if recommended_field and field == recommended_field:
                    print(f"{i:2d}. {field} (推荐)")
                else:
                    print(f"{i:2d}. {field}")
            
            # 让用户选择目标关联字段
            print(f"\n📝 请选择目标文件的关联字段:")
            while True:
                try:
                    link_choice = input("请输入关联字段编号: ").strip()
                    link_index = int(link_choice) - 1
                    if 0 <= link_index < len(target_columns):
                        self.link_field = target_columns[link_index]
                        print(f"✅ 目标关联字段: {self.link_field}")
                        break
                    else:
                        print("❌ 字段编号超出范围，请重新选择")
                except ValueError:
                    print("❌ 请输入有效的数字")
            
            # 第二步：为每个源文件建立字段映射
            print(f"\n🔄 步骤2: 为每个源文件建立关联字段映射")
            self.source_field_mapping = {}
            
            for source_file in self.source_files:
                source_filename = os.path.basename(source_file)
                print(f"\n📄 处理源文件: {source_filename}")
                
                try:
                    source_df = pd.read_excel(source_file)
                    source_columns = list(source_df.columns)
                    
                    if not source_columns:
                        print(f"⚠️  源文件 '{source_filename}' 没有任何字段，跳过")
                        continue
                    
                    # 尝试模糊匹配目标关联字段
                    similar_fields = []
                    for source_field in source_columns:
                        similarity = self.calculate_similarity(self.link_field, source_field)
                        if similarity >= 0.8:
                            similar_fields.append((source_field, similarity))
                    
                    # 按相似度排序
                    similar_fields.sort(key=lambda x: x[1], reverse=True)
                    
                    selected_source_field = None
                    
                    if similar_fields:
                        print(f"🔍 找到与 '{self.link_field}' 相似的字段:")
                        for i, (field, sim) in enumerate(similar_fields, 1):
                            print(f"  {i}. {field} (相似度: {sim:.2f})")
                        
                        print(f"\n🤔 请选择:")
                        print(f"  1. 使用相似字段 (输入编号，默认选择1)")
                        print(f"  2. 手动选择其他字段 (输入 'm')")
                        print(f"  3. 跳过此源文件 (输入 's')")
                        
                        while True:
                            choice = input("请选择 (默认1): ").strip().lower()
                            if not choice:  # 用户按回车，默认选择第一个相似字段
                                choice = "1"
                            
                            if choice == 's':
                                print(f"⏭️  跳过源文件 '{source_filename}'")
                                break
                            elif choice == 'm':
                                # 手动选择
                                selected_source_field = self._manual_select_source_field(source_columns, source_filename)
                                break
                            else:
                                try:
                                    choice_idx = int(choice) - 1
                                    if 0 <= choice_idx < len(similar_fields):
                                        selected_source_field = similar_fields[choice_idx][0]
                                        print(f"✅ 选择了相似字段: {selected_source_field}")
                                        break
                                    else:
                                        print("❌ 编号超出范围，请重新选择")
                                except ValueError:
                                    print("❌ 请输入有效的编号、'm' 或 's'")
                    else:
                        print(f"❌ 未找到与 '{self.link_field}' 相似的字段")
                        print(f"🤔 请选择:")
                        print(f"  1. 手动选择字段 (输入 'm')")
                        print(f"  2. 跳过此源文件 (输入 's')")
                        
                        while True:
                            choice = input("请选择: ").strip().lower()
                            if choice == 's':
                                print(f"⏭️  跳过源文件 '{source_filename}'")
                                break
                            elif choice == 'm':
                                selected_source_field = self._manual_select_source_field(source_columns, source_filename)
                                break
                            else:
                                print("❌ 请输入 'm' 或 's'")
                    
                    # 保存映射关系
                    if selected_source_field:
                        self.source_field_mapping[source_filename] = {
                            self.link_field: selected_source_field
                        }
                        print(f"📝 建立映射: {self.link_field} -> {selected_source_field}")
                
                except Exception as e:
                    print(f"❌ 处理源文件 '{source_filename}' 时出错: {str(e)}")
                    continue
            
            # 显示最终的映射结果
            print(f"\n📋 关联字段映射结果:")
            print(f"🎯 目标字段: {self.link_field}")
            for source_file, mapping in self.source_field_mapping.items():
                source_field = mapping.get(self.link_field, "未映射")
                print(f"  📄 {source_file}: {source_field}")
            
            if not self.source_field_mapping:
                print("❌ 没有建立任何源文件的字段映射，无法进行同步")
                return
                    
        except Exception as e:
            print(f"❌ 读取文件时出错: {str(e)}")
    
    def _manual_select_source_field(self, source_columns: List[str], source_filename: str) -> str:
        """手动选择源文件字段"""
        print(f"\n📋 源文件 '{source_filename}' 的所有字段:")
        for i, field in enumerate(source_columns, 1):
            print(f"  {i:2d}. {field}")
        
        print(f"\n📝 请选择要用作关联字段的源文件字段:")
        while True:
            try:
                choice = input("请输入字段编号: ").strip()
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(source_columns):
                    selected_field = source_columns[choice_idx]
                    print(f"✅ 选择了字段: {selected_field}")
                    return selected_field
                else:
                    print("❌ 字段编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字")
    
    def select_multi_sync_update_fields(self):
        """多源同步更新字段选择 - 以目标文件为主导的新逻辑"""
        print(f"\n=== 步骤3: 更新字段选择 ===")
        
        try:
            # 读取目标文件
            target_df = pd.read_excel(self.target_file)
            target_columns = list(target_df.columns)
            
            # 排除关联字段，显示目标文件的可更新字段
            updateable_target_fields = [field for field in target_columns if field != self.link_field]
            
            if not updateable_target_fields:
                print("❌ 目标文件除了关联字段外没有其他字段可更新")
                return
            
            print(f"🎯 从目标文件选择要更新的字段")
            print(f"📋 目标文件 '{os.path.basename(self.target_file)}' 的可更新字段（排除关联字段 '{self.link_field}'）:")
            
            # 显示目标文件的可更新字段
            for i, field in enumerate(updateable_target_fields, 1):
                print(f"{i:2d}. {field}")
            
            # 让用户选择要更新的字段
            print(f"\n📝 请选择要更新的字段:")
            print("📝 输入字段编号，用逗号分隔（如：1,2,3）")
            print("📝 输入 'all' 选择所有可更新字段")
            
            selected_target_fields = []
            while True:
                try:
                    update_choice = input("请输入要更新的字段编号: ").strip()
                    
                    if update_choice.lower() == 'all':
                        selected_target_fields = updateable_target_fields.copy()
                        print(f"✅ 已选择所有 {len(selected_target_fields)} 个目标字段进行更新")
                        break
                    else:
                        # 解析用户选择的字段编号
                        indices = [int(x.strip()) - 1 for x in update_choice.split(',')]
                        selected_target_fields = []
                        
                        for index in indices:
                            if 0 <= index < len(updateable_target_fields):
                                selected_target_fields.append(updateable_target_fields[index])
                            else:
                                print(f"⚠️  字段编号 {index + 1} 超出范围，跳过")
                        
                        if selected_target_fields:
                            print(f"✅ 已选择 {len(selected_target_fields)} 个目标字段进行更新:")
                            for field in selected_target_fields:
                                print(f"  📝 {field}")
                            break
                        else:
                            print("❌ 未选择任何有效字段，请重新选择")
                            
                except ValueError:
                    print("❌ 请输入有效的数字或 'all'")
            
            # 为每个选择的目标字段建立与源文件的映射关系
            print(f"\n🔄 为每个更新字段建立源文件映射")
            self.update_fields = []
            
            for target_field in selected_target_fields:
                print(f"\n📝 处理目标字段: {target_field}")
                field_has_mapping = False
                
                # 为每个源文件查找对应的字段
                for source_file in self.source_files:
                    source_filename = os.path.basename(source_file)
                    
                    # 跳过没有建立关联字段映射的源文件
                    if source_filename not in self.source_field_mapping:
                        continue
                    
                    try:
                        source_df = pd.read_excel(source_file)
                        source_columns = list(source_df.columns)
                        
                        # 尝试模糊匹配目标更新字段
                        similar_fields = []
                        for source_field in source_columns:
                            if source_field != self.source_field_mapping[source_filename].get(self.link_field):  # 排除关联字段
                                similarity = self.calculate_similarity(target_field, source_field)
                                if similarity >= 0.8:
                                    similar_fields.append((source_field, similarity))
                        
                        # 按相似度排序
                        similar_fields.sort(key=lambda x: x[1], reverse=True)
                        
                        selected_source_field = None
                        
                        if similar_fields:
                            print(f"  📄 源文件 '{source_filename}' - 找到与 '{target_field}' 相似的字段:")
                            for i, (field, sim) in enumerate(similar_fields, 1):
                                print(f"    {i}. {field} (相似度: {sim:.2f})")
                            
                            print(f"  🤔 请选择:")
                            print(f"    1. 使用相似字段 (输入编号，默认选择1)")
                            print(f"    2. 手动选择其他字段 (输入 'm')")
                            print(f"    3. 跳过此源文件 (输入 's')")
                            
                            while True:
                                choice = input(f"  对于源文件 '{source_filename}' 请选择 (默认1): ").strip().lower()
                                if not choice:  # 用户按回车，默认选择第一个相似字段
                                    choice = "1"
                                
                                if choice == 's':
                                    print(f"  ⏭️  跳过源文件 '{source_filename}'")
                                    break
                                elif choice == 'm':
                                    selected_source_field = self._manual_select_update_field(source_columns, source_filename, target_field)
                                    break
                                else:
                                    try:
                                        choice_idx = int(choice) - 1
                                        if 0 <= choice_idx < len(similar_fields):
                                            selected_source_field = similar_fields[choice_idx][0]
                                            print(f"  ✅ 选择了相似字段: {selected_source_field}")
                                            field_has_mapping = True
                                            break
                                        else:
                                            print("  ❌ 编号超出范围，请重新选择")
                                    except ValueError:
                                        print("  ❌ 请输入有效的编号、'm' 或 's'")
                        else:
                            print(f"  📄 源文件 '{source_filename}' - 未找到与 '{target_field}' 相似的字段")
                            print(f"  🤔 请选择:")
                            print(f"    1. 手动选择字段 (输入 'm')")
                            print(f"    2. 跳过此源文件 (输入 's')")
                            
                            while True:
                                choice = input(f"  对于源文件 '{source_filename}' 请选择: ").strip().lower()
                                if choice == 's':
                                    print(f"  ⏭️  跳过源文件 '{source_filename}'")
                                    break
                                elif choice == 'm':
                                    selected_source_field = self._manual_select_update_field(source_columns, source_filename, target_field)
                                    if selected_source_field:
                                        field_has_mapping = True
                                    break
                                else:
                                    print("  ❌ 请输入 'm' 或 's'")
                        
                        # 保存更新字段映射关系
                        if selected_source_field:
                            if source_filename not in self.source_field_mapping:
                                self.source_field_mapping[source_filename] = {}
                            self.source_field_mapping[source_filename][target_field] = selected_source_field
                            print(f"  📝 建立映射: {target_field} -> {selected_source_field}")
                            field_has_mapping = True
                    
                    except Exception as e:
                        print(f"  ❌ 处理源文件 '{source_filename}' 时出错: {str(e)}")
                        continue
                
                # 如果至少有一个源文件建立了映射，则添加到更新字段列表
                if field_has_mapping:
                    self.update_fields.append(target_field)
                    print(f"✅ 目标字段 '{target_field}' 已添加到更新列表")
                else:
                    print(f"⚠️  目标字段 '{target_field}' 没有在任何源文件中找到对应字段，跳过")
            
            if not self.update_fields:
                print("❌ 没有建立任何更新字段的映射关系")
                return
            
            # 显示最终的更新字段映射结果
            print(f"\n📋 更新字段映射结果:")
            for target_field in self.update_fields:
                print(f"🎯 目标字段: {target_field}")
                for source_file, mapping in self.source_field_mapping.items():
                    if target_field in mapping:
                        print(f"  📄 {source_file}: {mapping[target_field]}")
                    
        except Exception as e:
            print(f"❌ 读取文件时出错: {str(e)}")
    
    def _manual_select_update_field(self, source_columns: List[str], source_filename: str, target_field: str) -> str:
        """手动选择源文件的更新字段"""
        # 排除已经用作关联字段的字段
        available_fields = []
        link_field_in_source = self.source_field_mapping.get(source_filename, {}).get(self.link_field)
        
        for field in source_columns:
            if field != link_field_in_source:  # 排除关联字段
                available_fields.append(field)
        
        if not available_fields:
            print(f"  ⚠️  源文件 '{source_filename}' 没有可用的更新字段")
            return None
        
        print(f"\n  📋 源文件 '{source_filename}' 的可用字段（排除关联字段）:")
        for i, field in enumerate(available_fields, 1):
            print(f"    {i:2d}. {field}")
        
        print(f"\n  📝 请选择要映射到目标字段 '{target_field}' 的源文件字段:")
        while True:
            try:
                choice = input("  请输入字段编号: ").strip()
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(available_fields):
                    selected_field = available_fields[choice_idx]
                    print(f"  ✅ 选择了字段: {selected_field}")
                    return selected_field
                else:
                    print("  ❌ 字段编号超出范围，请重新选择")
            except ValueError:
                print("  ❌ 请输入有效的数字")
    
    def configure_conflict_handling(self):
        """配置冲突处理方式"""
        print(f"\n=== 步骤4: 冲突处理配置 ===")
        
        print("🤔 当多个源文件对同一记录提供不同数据时，如何处理冲突？")
        print("1. 询问用户选择（推荐）")
        print("2. 使用第一个源文件的数据")
        print("3. 使用最后一个源文件的数据")
        print("4. 跳过冲突的记录")
        
        while True:
            choice = input("\n请选择冲突处理方式 (1/2/3/4): ").strip()
            if choice == "1":
                self.conflict_handling = "ask"
                print("✅ 已选择：询问用户选择")
                break
            elif choice == "2":
                self.conflict_handling = "first"
                print("✅ 已选择：使用第一个源文件的数据")
                break
            elif choice == "3":
                self.conflict_handling = "last"
                print("✅ 已选择：使用最后一个源文件的数据")
                break
            elif choice == "4":
                self.conflict_handling = "skip"
                print("✅ 已选择：跳过冲突的记录")
                break
            else:
                print("❌ 无效选择，请输入 1、2、3 或 4")
    
    def set_multi_sync_output(self):
        """设置多源同步输出"""
        print(f"\n=== 步骤5: 输出设置 ===")
        print("💡 您可以输入:")
        print("   • 完整的文件路径（如：G:\\wang\\excel\\result.xlsx）")
        print("   • 目录路径（将自动生成文件名）")
        
        # 获取用户输入
        user_input = input("请输入输出路径（或按回车使用默认目录G:\\wang\\excel）: ").strip()
        if not user_input:
            user_input = "G:\\wang\\excel"
        
        # 判断用户输入的是文件路径还是目录路径
        if user_input.lower().endswith(('.xlsx', '.xls')):
            # 用户输入的是完整文件路径
            self.output_file_path = os.path.abspath(user_input)
            self.output_directory = os.path.dirname(self.output_file_path)
            output_filename = os.path.basename(self.output_file_path)
            
            # 检查文件是否已存在
            if os.path.exists(self.output_file_path):
                print(f"⚠️  文件已存在: {self.output_file_path}")
                overwrite = input("是否要替换现有文件？(y/n，默认n): ").strip().lower()
                if overwrite not in ['y', 'yes', '是']:
                    # 生成新文件名
                    base_name = os.path.splitext(output_filename)[0]
                    extension = os.path.splitext(output_filename)[1]
                    counter = 1
                    while True:
                        new_filename = f"{base_name}_{counter}{extension}"
                        new_file_path = os.path.join(self.output_directory, new_filename)
                        if not os.path.exists(new_file_path):
                            self.output_file_path = new_file_path
                            output_filename = new_filename
                            print(f"📝 使用新文件名: {output_filename}")
                            break
                        counter += 1
                else:
                    print("✅ 将替换现有文件")
            
            # 确保输出目录存在
            if not os.path.exists(self.output_directory):
                try:
                    os.makedirs(self.output_directory)
                    print(f"✅ 已创建输出目录: {self.output_directory}")
                except Exception as e:
                    print(f"❌ 创建输出目录失败: {str(e)}")
                    return
            
            print(f"✅ 输出文件: {self.output_file_path}")
            
        else:
            # 用户输入的是目录路径
            self.output_directory = os.path.abspath(user_input)
            self.output_file_path = None  # 将在保存时自动生成文件名
            
            # 确保输出目录存在
            if not os.path.exists(self.output_directory):
                try:
                    os.makedirs(self.output_directory)
                    print(f"✅ 已创建输出目录: {self.output_directory}")
                except Exception as e:
                    print(f"❌ 创建输出目录失败: {str(e)}")
                    return
            
            print(f"✅ 输出目录: {self.output_directory}")
            print("💡 文件名将自动生成")
    
    def execute_multi_sync(self):
        """执行多源同步"""
        print(f"\n=== 步骤6: 执行多源同步 ===")
        
        try:
            # 读取目标文件
            target_df = pd.read_excel(self.target_file)
            print(f"📊 目标文件包含 {len(target_df)} 条记录")
            
            # 读取所有源文件
            source_data = {}
            for source_file in self.source_files:
                source_df = pd.read_excel(source_file)
                source_data[os.path.basename(source_file)] = source_df
                print(f"📊 源文件 '{os.path.basename(source_file)}' 包含 {len(source_df)} 条记录")
            
            # 执行多源同步
            updated_df = self.perform_multi_sync(target_df, source_data)
            
            # 保存更新后的文件
            self.save_multi_sync_file(updated_df)
            
            # 显示同步结果
            self.show_multi_sync_results(target_df, updated_df)
            
        except Exception as e:
            print(f"❌ 执行多源同步时出错: {str(e)}")
    
    def perform_multi_sync(self, target_df: pd.DataFrame, source_data: Dict[str, pd.DataFrame]) -> pd.DataFrame:
        """
        执行多源同步
        
        Args:
            target_df: 目标数据框
            source_data: 源数据字典 {文件名: 数据框}
            
        Returns:
            更新后的数据框
        """
        print(f"\n🔄 开始执行多源同步...")
        
        # 确定关联字段的实际名称
        actual_link_field = self.find_actual_field_name(target_df, self.link_field)
        if not actual_link_field:
            print(f"❌ 目标文件中找不到关联字段 '{self.link_field}'")
            return target_df
        
        # 创建更新后的数据框副本
        updated_df = target_df.copy()
        
        # 统计信息
        total_updates = 0
        conflicts_resolved = 0
        conflicts_skipped = 0
        
        # 为每个目标记录查找源数据
        for target_idx, target_row in updated_df.iterrows():
            link_value = str(target_row[actual_link_field]).strip()
            
            # 跳过空关联值
            if pd.isna(link_value) or link_value == '':
                continue
            
            # 在所有源文件中查找匹配的记录
            matching_data = {}
            for source_name, source_df in source_data.items():
                # 使用建立的字段映射关系确定源文件中的关联字段名称
                source_link_field = None
                if source_name in self.source_field_mapping:
                    source_link_field = self.source_field_mapping[source_name].get(self.link_field)
                
                # 如果没有建立映射关系，尝试使用原始逻辑作为后备
                if not source_link_field:
                    source_link_field = self.find_actual_field_name(source_df, self.link_field)
                
                if not source_link_field or source_link_field not in source_df.columns:
                    continue
                
                # 查找匹配的记录
                matching_rows = source_df[source_df[source_link_field].astype(str).str.strip() == link_value]
                if not matching_rows.empty:
                    matching_data[source_name] = matching_rows.iloc[0]  # 取第一条匹配记录
            
            if not matching_data:
                continue
            
            # 处理每个更新字段
            for update_field in self.update_fields:
                # 确定目标字段的实际名称
                actual_update_field = self.find_actual_field_name(updated_df, update_field)
                if not actual_update_field:
                    continue
                
                # 收集所有源文件中的值
                field_values = {}
                for source_name, source_row in matching_data.items():
                    # 使用建立的字段映射关系确定源文件中的字段名称
                    source_field = None
                    if source_name in self.source_field_mapping:
                        source_field = self.source_field_mapping[source_name].get(update_field)
                    
                    # 如果没有建立映射关系，尝试使用原始逻辑作为后备
                    if not source_field:
                        source_field = self.find_actual_field_name(source_data[source_name], update_field)
                    
                    if source_field and source_field in source_row.index and not pd.isna(source_row[source_field]):
                        field_values[source_name] = str(source_row[source_field]).strip()
                
                if not field_values:
                    continue
                
                # 检查是否有冲突（多个不同的值）
                unique_values = set(field_values.values())
                if len(unique_values) == 1:
                    # 没有冲突，检查是否需要更新
                    value = list(unique_values)[0]
                    current_value = updated_df.at[target_idx, actual_update_field]
                    
                    # 如果当前值和新值相同，跳过更新
                    if (not pd.isna(current_value) and 
                        str(current_value).strip() == str(value).strip()):
                        continue  # 值相同，无需更新
                    
                    # 确保数据类型兼容，先将列转换为object类型
                    if updated_df[actual_update_field].dtype != 'object':
                        updated_df[actual_update_field] = updated_df[actual_update_field].astype('object')
                    
                    if pd.isna(value) or value == '':
                        updated_df.at[target_idx, actual_update_field] = None
                    else:
                        # 将值转换为字符串以避免类型不兼容警告
                        updated_df.at[target_idx, actual_update_field] = str(value)
                    total_updates += 1
                else:
                    # 有冲突，根据配置处理
                    if self.conflict_handling == "ask":
                        # 询问用户选择
                        choice = self.ask_user_for_conflict_resolution(link_value, update_field, field_values)
                        if choice:
                            current_value = updated_df.at[target_idx, actual_update_field]
                            
                            # 检查是否需要更新
                            if (pd.isna(current_value) or 
                                str(current_value).strip() != str(choice).strip()):
                                # 确保数据类型兼容，先将列转换为object类型
                                if updated_df[actual_update_field].dtype != 'object':
                                    updated_df[actual_update_field] = updated_df[actual_update_field].astype('object')
                                
                                if pd.isna(choice) or choice == '':
                                    updated_df.at[target_idx, actual_update_field] = None
                                else:
                                    updated_df.at[target_idx, actual_update_field] = str(choice)
                                total_updates += 1
                            conflicts_resolved += 1
                    elif self.conflict_handling == "first":
                        # 使用第一个源文件的数据
                        first_source = list(field_values.keys())[0]
                        value = field_values[first_source]
                        current_value = updated_df.at[target_idx, actual_update_field]
                        
                        # 检查是否需要更新
                        if (pd.isna(current_value) or 
                            str(current_value).strip() != str(value).strip()):
                            # 确保数据类型兼容，先将列转换为object类型
                            if updated_df[actual_update_field].dtype != 'object':
                                updated_df[actual_update_field] = updated_df[actual_update_field].astype('object')
                            
                            if pd.isna(value) or value == '':
                                updated_df.at[target_idx, actual_update_field] = None
                            else:
                                updated_df.at[target_idx, actual_update_field] = str(value)
                            total_updates += 1
                        conflicts_resolved += 1
                    elif self.conflict_handling == "last":
                        # 使用最后一个源文件的数据
                        last_source = list(field_values.keys())[-1]
                        value = field_values[last_source]
                        current_value = updated_df.at[target_idx, actual_update_field]
                        
                        # 检查是否需要更新
                        if (pd.isna(current_value) or 
                            str(current_value).strip() != str(value).strip()):
                            # 确保数据类型兼容，先将列转换为object类型
                            if updated_df[actual_update_field].dtype != 'object':
                                updated_df[actual_update_field] = updated_df[actual_update_field].astype('object')
                            
                            if pd.isna(value) or value == '':
                                updated_df.at[target_idx, actual_update_field] = None
                            else:
                                updated_df.at[target_idx, actual_update_field] = str(value)
                            total_updates += 1
                        conflicts_resolved += 1
                    elif self.conflict_handling == "skip":
                        # 跳过冲突的记录
                        conflicts_skipped += 1
                        continue
        
        print(f"✅ 多源同步完成:")
        print(f"  📊 总更新数: {total_updates}")
        print(f"  🔄 冲突解决数: {conflicts_resolved}")
        print(f"  ⏭️  冲突跳过数: {conflicts_skipped}")
        
        return updated_df
    
    def ask_user_for_conflict_resolution(self, link_value: str, field_name: str, field_values: Dict[str, str]) -> str:
        """
        询问用户解决冲突
        
        Args:
            link_value: 关联值
            field_name: 字段名
            field_values: 字段值字典 {源文件名: 值}
            
        Returns:
            用户选择的值或None
        """
        print(f"\n⚠️  发现数据冲突:")
        print(f"  🔗 关联值: {link_value}")
        print(f"  📋 字段: {field_name}")
        print(f"  📄 不同源文件提供的值:")
        
        for i, (source_name, value) in enumerate(field_values.items(), 1):
            print(f"    {i}. {source_name}: {value}")
        
        print(f"  📝 请选择要使用的值（输入编号）:")
        print(f"  📝 输入 'skip' 跳过此字段的更新")
        
        while True:
            choice = input("请选择: ").strip()
            if choice.lower() == 'skip':
                return None
            try:
                choice_index = int(choice) - 1
                if 0 <= choice_index < len(field_values):
                    selected_source = list(field_values.keys())[choice_index]
                    selected_value = field_values[selected_source]
                    print(f"✅ 已选择: {selected_source} 的值 '{selected_value}'")
                    return selected_value
                else:
                    print("❌ 选择编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字或 'skip'")
    
    def save_multi_sync_file(self, updated_df: pd.DataFrame):
        """保存多源同步文件"""
        try:
            # 确定输出文件路径
            if hasattr(self, 'output_file_path') and self.output_file_path:
                # 用户指定了完整的文件路径
                output_path = self.output_file_path
            else:
                # 用户只指定了目录，自动生成文件名
                target_basename = os.path.splitext(os.path.basename(self.target_file))[0]
                timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
                output_filename = f"{target_basename}_多源同步_{timestamp}.xlsx"
                output_path = os.path.join(self.output_directory, output_filename)
            
            # 保存文件
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 主数据表
                updated_df.to_excel(writer, sheet_name='同步后数据', index=False)
                
                # 同步统计表
                stats_data = {
                    '统计项目': [
                        '源文件数量',
                        '目标文件',
                        '关联字段',
                        '更新字段数',
                        '冲突处理方式',
                        '同步时间'
                    ],
                    '数值': [
                        len(self.source_files),
                        os.path.basename(self.target_file),
                        self.link_field,
                        len(self.update_fields),
                        self.conflict_handling,
                        pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='同步统计', index=False)
            
            print(f"✅ 同步结果已保存到: {output_path}")
            self.output_file = output_path
            
        except Exception as e:
            print(f"❌ 保存文件时出错: {str(e)}")
    
    def show_multi_sync_results(self, original_df: pd.DataFrame, updated_df: pd.DataFrame):
        """显示多源同步结果"""
        print(f"\n=== 多源同步结果 ===")
        print(f"📊 原始记录数: {len(original_df)}")
        print(f"📊 更新后记录数: {len(updated_df)}")
        print(f"📄 输出文件: {self.output_file}")
        print(f"📋 更新的字段: {', '.join(self.update_fields)}")
        print(f"🔗 关联字段: {self.link_field}")

def main():
    """主函数"""
    processor = ExcelProcessor()
    processor.run()

if __name__ == "__main__":
    main()
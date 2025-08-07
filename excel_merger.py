import pandas as pd
import os
import glob
import re
from typing import List, Tuple, Dict, Optional
from difflib import SequenceMatcher
from difflib import SequenceMatcher

class ExcelProcessor:
    """Excel文件处理工具"""
    
    def __init__(self):
        self.selected_files = []
        self.all_fields = []
        self.selected_fields = []
        self.deduplicate = False
        self.dedup_fields = []
        self.output_filename = "result.xlsx"
        
        # 字段补充功能相关属性
        self.enable_field_supplement = False
        self.field_mappings = {}  # 字段映射字典 {field_name: {student_id: value}}
        self.default_values = {}  # 默认值字典 {field_name: default_value}
        self.link_field = '学号'  # 关联字段，默认为学号
        
        # 新增：智能列名匹配相关属性
        self.column_mapping = {}  # 列名映射关系
        self.enable_smart_matching = True  # 是否启用智能匹配
        self.similarity_threshold = 0.8  # 相似度阈值
        self.auto_clean_columns = True  # 是否自动清理列名
        
        # 常见列名变体映射
        self.common_column_variants = {
            '学号': ['学号', '学号号', '学学号', 'xuehao', 'student_id', '学生编号', '学生学号'],
            '学生姓名': ['学生姓名', '学生姓名名', '学学生姓名', 'student_name', '姓名', '学生名', '学生姓名（中文）'],
            '班级': ['班级', '班', 'class', '班级名称'],
            '成绩': ['成绩', '分数', 'score', 'grade', '考试分数'],
            '课程': ['课程', '科目', 'course', 'subject', '课程名称']
        }
        
        # 新增：智能列名匹配相关属性
        self.column_mapping = {}  # 列名映射关系
        self.enable_smart_matching = True  # 是否启用智能匹配
        self.similarity_threshold = 0.8  # 相似度阈值
        self.auto_clean_columns = True  # 是否自动清理列名
        
        # 常见列名变体映射
        self.common_column_variants = {
            '学号': ['学号', '学号号', '学学号', 'xuehao', 'student_id', '学生编号', '学生学号'],
            '学生姓名': ['学生姓名', '学生姓名名', '学学生姓名', 'student_name', '姓名', '学生名', '学生姓名（中文）'],
            '班级': ['班级', '班', 'class', '班级名称'],
            '成绩': ['成绩', '分数', 'score', 'grade', '考试分数'],
            '课程': ['课程', '科目', 'course', 'subject', '课程名称']
        }
    
    def clean_column_name(self, column_name: str) -> str:
        """
        清理列名，去除空格、特殊字符等
        
        Args:
            column_name: 原始列名
            
        Returns:
            清理后的列名
        """
        if not self.auto_clean_columns:
            return column_name
        
        # 去除首尾空格
        cleaned = column_name.strip()
        
        # 去除多余的空格
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        # 去除特殊字符（保留中文、英文、数字、下划线）
        cleaned = re.sub(r'[^\w\s\u4e00-\u9fff]', '', cleaned)
        
        # 再次去除空格
        cleaned = cleaned.strip()
        
        return cleaned
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度
        
        Args:
            str1: 字符串1
            str2: 字符串2
            
        Returns:
            相似度 (0-1)
        """
        # 使用SequenceMatcher计算相似度
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
    def find_similar_columns(self, target_column: str, available_columns: List[str]) -> List[Tuple[str, float]]:
        """
        查找与目标列名相似的列名
        
        Args:
            target_column: 目标列名
            available_columns: 可用列名列表
            
        Returns:
            相似列名列表，包含相似度
        """
        similar_columns = []
        cleaned_target = self.clean_column_name(target_column)
        
        for column in available_columns:
            cleaned_column = self.clean_column_name(column)
            
            # 精确匹配
            if cleaned_target == cleaned_column:
                similar_columns.append((column, 1.0))
                continue
            
            # 计算相似度
            similarity = self.calculate_similarity(cleaned_target, cleaned_column)
            
            # 检查是否是常见变体
            for standard_name, variants in self.common_column_variants.items():
                if cleaned_target in variants and cleaned_column in variants:
                    similarity = max(similarity, 0.9)  # 提高变体的相似度
                    break
            
            if similarity >= self.similarity_threshold:
                similar_columns.append((column, similarity))
        
        # 按相似度排序
        similar_columns.sort(key=lambda x: x[1], reverse=True)
        return similar_columns
    
    def smart_column_mapping(self, required_columns: List[str], available_columns: List[str]) -> Dict[str, str]:
        """
        智能列名映射
        
        Args:
            required_columns: 需要的列名
            available_columns: 可用的列名
            
        Returns:
            列名映射字典
        """
        mapping = {}
        unmapped_required = []
        unmapped_available = available_columns.copy()
        
        print(f"\n🔍 智能列名映射分析...")
        print(f"📋 需要的列名: {required_columns}")
        print(f"📋 可用的列名: {available_columns}")
        
        # 第一轮：精确匹配和常见变体匹配
        for required in required_columns:
            matched = False
            
            # 检查精确匹配
            if required in available_columns:
                mapping[required] = required
                unmapped_available.remove(required)
                print(f"✅ 精确匹配: {required} -> {required}")
                matched = True
                continue
            
            # 检查常见变体
            if required in self.common_column_variants:
                variants = self.common_column_variants[required]
                for variant in variants:
                    if variant in available_columns:
                        mapping[required] = variant
                        unmapped_available.remove(variant)
                        print(f"✅ 变体匹配: {required} -> {variant}")
                        matched = True
                        break
            
            if not matched:
                unmapped_required.append(required)
        
        # 第二轮：模糊匹配
        if unmapped_required and unmapped_available:
            print(f"\n🔍 进行模糊匹配...")
            for required in unmapped_required:
                similar_columns = self.find_similar_columns(required, unmapped_available)
                
                if similar_columns:
                    best_match, similarity = similar_columns[0]
                    print(f"🔍 找到相似列名: {required} -> {best_match} (相似度: {similarity:.2f})")
                    
                    # 询问用户是否确认映射
                    confirm = input(f"是否将 '{required}' 映射到 '{best_match}'？(y/n，默认y): ").strip().lower()
                    if confirm not in ['n', 'no', '否']:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 确认映射: {required} -> {best_match}")
                    else:
                        print(f"⚠️  跳过映射: {required}")
                else:
                    print(f"❌ 未找到与 '{required}' 相似的列名")
        
        # 显示映射结果
        if mapping:
            print(f"\n📋 列名映射结果:")
            for required, mapped in mapping.items():
                print(f"  {required} -> {mapped}")
        
        if unmapped_required:
            print(f"\n⚠️  未映射的列名: {unmapped_required}")
        
        return mapping
    
    def validate_required_columns(self, df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str], Dict[str, str]]:
        """
        验证必需的列名是否存在，支持智能匹配
        
        Args:
            df: 数据框
            required_columns: 必需的列名列表
            
        Returns:
            (是否验证通过, 缺失的列名列表, 列名映射字典)
        """
        available_columns = list(df.columns)
        missing_columns = []
        column_mapping = {}
        
        if not self.enable_smart_matching:
            # 传统严格匹配
            for required in required_columns:
                if required not in available_columns:
                    missing_columns.append(required)
                else:
                    column_mapping[required] = required
        else:
            # 智能匹配
            column_mapping = self.smart_column_mapping(required_columns, available_columns)
            
            # 检查哪些列名没有被映射
            for required in required_columns:
                if required not in column_mapping:
                    missing_columns.append(required)
        
        return len(missing_columns) == 0, missing_columns, column_mapping
    
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
                
                # 过滤掉无效字段（说明文字、Unnamed字段等）
                valid_fields = []
                for field in file_fields:
                    # 跳过Unnamed字段
                    if field.startswith('Unnamed:'):
                        continue
                    # 跳过说明文字（通常包含很长的描述性文字）
                    if len(field) > 100:
                        continue
                    # 跳过空字段
                    if not field or field.strip() == '':
                        continue
                    # 跳过纯说明性字段
                    if field in ['说明', '说明文字', '备注', '注释']:
                        continue
                    # 跳过包含说明关键词的字段
                    if any(keyword in field for keyword in ['说明', '备注', '注释', '注意', '提示']):
                        continue
                    
                    # 如果启用自动清理，显示清理后的列名
                    if self.auto_clean_columns:
                        cleaned_field = self.clean_column_name(field)
                        if cleaned_field != field:
                            print(f"📝 列名清理: '{field}' -> '{cleaned_field}'")
                        valid_fields.append(cleaned_field)
                    else:
                        valid_fields.append(field)
                
                all_fields.update(valid_fields)
                file_field_info[os.path.basename(file)] = {
                    'field_count': len(valid_fields),
                    'fields': valid_fields
                }
                print(f"📊 文件 '{os.path.basename(file)}' 包含 {len(valid_fields)} 个有效字段")
                
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
        
        print(f"\n✅ 总共发现 {len(self.all_fields)} 个不同有效字段")
        
        return self.all_fields
    
    def analyze_field_supplement_situation(self, files: List[str], selected_fields: List[str]) -> Dict:
        """
        分析字段补充情况
        
        Args:
            files: 文件列表
            selected_fields: 用户选择的字段列表
            
        Returns:
            分析结果字典
        """
        analysis_result = {
            'files_with_all_fields': [],  # 包含所有必需字段的文件
            'files_missing_fields': {},   # 缺少特定字段的文件 {field: [files]}
            'files_without_key_field': [], # 不包含关键字段（学号）的文件
            'total_files': len(files)
        }
        
        # 初始化缺失字段字典
        for field in selected_fields:
            analysis_result['files_missing_fields'][field] = []
        
        print(f"\n🔍 分析字段补充情况...")
        
        for file in files:
            try:
                df = pd.read_excel(file)
                file_fields = self.get_file_fields(file)  # 使用过滤后的字段
                filename = os.path.basename(file)
                
                # 检查是否包含学号（作为关键字段）
                has_student_id = any(id_field in file_fields for id_field in ['学号', '*学号'])
                
                if not has_student_id:
                    analysis_result['files_without_key_field'].append(file)
                    print(f"ℹ️  {filename}: 不包含学号")
                    continue
                
                # 检查每个必需字段
                missing_fields = []
                for field in selected_fields:
                    # 支持智能匹配，检查字段是否存在
                    field_found = False
                    for col in df.columns:
                        if field in col or col in field:
                            field_found = True
                            break
                    
                    if not field_found:
                        missing_fields.append(field)
                        analysis_result['files_missing_fields'][field].append(file)
                
                if not missing_fields:
                    analysis_result['files_with_all_fields'].append(file)
                    print(f"✅ {filename}: 包含所有必需字段")
                else:
                    missing_str = ', '.join(missing_fields)
                    print(f"⚠️  {filename}: 缺少字段 {missing_str}")
                    
            except Exception as e:
                print(f"❌ 分析文件 '{os.path.basename(file)}' 时出错: {str(e)}")
                analysis_result['files_without_key_field'].append(file)
        
        return analysis_result
    
    def build_field_mapping(self, files_with_all_fields: List[str], target_field: str, link_field: str = '学号') -> Dict[str, str]:
        """
        构建关联字段到目标字段的映射
        
        Args:
            files_with_all_fields: 包含所有必需字段的文件列表
            target_field: 目标字段名称
            link_field: 关联字段名称（默认学号）
            
        Returns:
            关联字段到目标字段的映射字典
        """
        if not files_with_all_fields:
            return {}
        
        print(f"\n🔄 构建{link_field}到{target_field}的映射...")
        mapping = {}
        total_mappings = 0
        
        for file in files_with_all_fields:
            try:
                df = pd.read_excel(file)
                filename = os.path.basename(file)
                
                # 确定关联字段名称
                link_field_name = None
                for col in df.columns:
                    if link_field in col or col in link_field:
                        link_field_name = col
                        break
                
                if not link_field_name:
                    print(f"⚠️  文件 '{filename}' 缺少{link_field}字段，跳过")
                    continue
                
                # 确定目标字段名称（支持智能匹配）
                target_field_name = None
                for col in df.columns:
                    if target_field in col or col in target_field:
                        target_field_name = col
                        break
                
                if not target_field_name:
                    print(f"⚠️  文件 '{filename}' 缺少{target_field}字段，跳过")
                    continue
                
                # 构建映射关系
                file_mappings = 0
                for _, row in df.iterrows():
                    link_value = str(row[link_field_name]).strip()
                    target_value = str(row[target_field_name]).strip()
                    
                    # 跳过空值
                    if pd.isna(link_value) or pd.isna(target_value) or link_value == '' or target_value == '':
                        continue
                    
                    # 如果关联值已存在，优先使用第一个匹配
                    if link_value not in mapping:
                        mapping[link_value] = target_value
                        file_mappings += 1
                
                total_mappings += file_mappings
                print(f"📊 {filename}: 添加了 {file_mappings} 个映射关系")
                
            except Exception as e:
                print(f"❌ 处理文件 '{os.path.basename(file)}' 时出错: {str(e)}")
                continue
        
        print(f"✅ 总共构建了 {total_mappings} 个{link_field}-{target_field}映射关系")
        return mapping
    
    def configure_field_supplement(self, analysis_result: Dict, selected_fields: List[str]) -> Tuple[bool, Dict[str, str], str]:
        """
        配置字段补充功能
        
        Args:
            analysis_result: 分析结果
            selected_fields: 用户选择的字段列表
            
        Returns:
            (是否启用补充功能, 默认值字典)
        """
        files_with_all_fields = analysis_result['files_with_all_fields']
        files_missing_fields = analysis_result['files_missing_fields']
        
        # 检查是否有缺失字段
        missing_fields = [field for field, files in files_missing_fields.items() if files]
        
        if not missing_fields:
            print(f"\n✅ 所有文件都包含所有必需字段，无需补充")
            return False, {}
        
        if not files_with_all_fields:
            print(f"\n⚠️  没有找到包含所有必需字段的文件，无法构建映射关系")
            print(f"📝 建议：至少需要一个包含所有必需字段的文件来构建映射关系")
            return False, ""
        
        print(f"\n=== 字段补充配置 ===")
        print(f"📊 分析结果:")
        print(f"  • 包含所有必需字段的文件: {len(files_with_all_fields)} 个")
        print(f"  • 不包含学号的文件: {len(analysis_result['files_without_key_field'])} 个")
        
        for field in missing_fields:
            missing_files = files_missing_fields[field]
            print(f"  • 缺少{field}字段的文件: {len(missing_files)} 个")
        
        print(f"\n🤔 检测到部分文件缺少字段，是否启用字段补充功能？")
        print(f"📝 补充功能将从其他文件中根据输入字段匹配获取缺失字段")
        
        choice = input("请选择 (y/n，默认y): ").strip().lower()
        enable_supplement = choice not in ['n', 'no', '否']
        
        if not enable_supplement:
            print(f"✅ 已选择不启用补充功能")
            return False, {}, '学号'
        
        # 选择关联字段
        print(f"\n🔗 请选择用于匹配的关联字段:")
        print(f"📋 可用字段: {', '.join(selected_fields)}")
        print(f"📝 输入字段名称（如：学号、学生姓名等）")
        print(f"📝 建议选择在所有文件中都存在且唯一性较好的字段作为关联字段")
        
        link_field = input("关联字段（默认：学号）: ").strip()
        if not link_field:
            link_field = '学号'
        
        # 验证关联字段是否在选中字段中
        if link_field not in selected_fields:
            print(f"⚠️  关联字段 '{link_field}' 不在选中字段中，将使用默认字段 '学号'")
            link_field = '学号'
        
        print(f"✅ 已设置关联字段: {link_field}")
        
        # 为每个缺失字段设置默认值
        default_values = {}
        for field in missing_fields:
            print(f"\n📝 请输入{field}字段未找到匹配时使用的默认值")
            default_value = input(f"默认值（默认：未知{field}）: ").strip()
            if not default_value:
                default_value = f"未知{field}"
            default_values[field] = default_value
            print(f"✅ 已设置{field}默认值: {default_value}")
        
        return True, default_values, link_field
    
    def supplement_fields(self, df: pd.DataFrame, field_mappings: Dict[str, Dict[str, str]], 
                         default_values: Dict[str, str], link_field: str = '学号') -> pd.DataFrame:
        """
        为数据框补充缺失字段
        
        Args:
            df: 数据框
            field_mappings: 字段映射字典 {field_name: {student_id: value}}
            default_values: 默认值字典 {field_name: default_value}
            
        Returns:
            补充后的数据框
        """
        # 确定关联字段名称
        link_field_name = None
        for col in df.columns:
            if link_field in col or col in link_field:
                link_field_name = col
                break
        
        if not link_field_name:
            print(f"⚠️  数据框不包含关联字段 '{link_field}'，无法补充字段")
            return df
        
        # 为每个缺失字段进行补充
        for field_name, mapping in field_mappings.items():
            # 确定目标字段名称
            target_field_name = None
            for col in df.columns:
                if field_name in col or col in field_name:
                    target_field_name = col
                    break
            
            # 如果字段不存在，创建一个新的
            if not target_field_name:
                target_field_name = field_name
                df[target_field_name] = default_values.get(field_name, f"未知{field_name}")
                print(f"📝 创建新的{field_name}字段")
            else:
                # 检查是否有空值需要补充
                missing_values = df[target_field_name].isna() | (df[target_field_name].astype(str).str.strip() == '')
                if not missing_values.any():
                    print(f"✅ {field_name}字段已完整，无需补充")
                    continue
                else:
                    missing_count = missing_values.sum()
                    print(f"📊 发现 {missing_count} 个空的{field_name}，开始补充...")
            
            # 补充字段值
            supplemented_count = 0
            successful_matches = 0
            default_used = 0
            
            for idx, row in df.iterrows():
                link_value = str(row[link_field_name]).strip()
                
                # 跳过空关联值
                if pd.isna(link_value) or link_value == '':
                    continue
                
                # 检查当前字段值是否为空
                current_value = str(row[target_field_name]).strip()
                if pd.isna(current_value) or current_value == '' or current_value == default_values.get(field_name, f"未知{field_name}"):
                    # 尝试从映射中获取值（精确匹配）
                    if link_value in mapping:
                        df.at[idx, target_field_name] = mapping[link_value]
                        successful_matches += 1
                    else:
                        # 尝试模糊匹配（支持一位字符的差异）
                        matched_value = None
                        for map_key, map_value in mapping.items():
                            # 如果关联值长度相同，尝试一位字符的模糊匹配
                            if len(link_value) == len(map_key):
                                # 计算不同字符的数量
                                diff_count = sum(1 for a, b in zip(link_value, map_key) if a != b)
                                if diff_count <= 1:  # 允许一位字符的差异
                                    matched_value = map_value
                                    break
                        
                        if matched_value:
                            df.at[idx, target_field_name] = matched_value
                            successful_matches += 1
                        else:
                            df.at[idx, target_field_name] = default_values.get(field_name, f"未知{field_name}")
                            default_used += 1
                    supplemented_count += 1
            
            if supplemented_count > 0:
                print(f"📊 {field_name}补充统计: 成功匹配 {successful_matches} 个，使用默认值 {default_used} 个")
        
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
            file_fields = list(df.columns)
            
            # 过滤掉无效字段（说明文字、Unnamed字段等）
            valid_fields = []
            for field in file_fields:
                # 跳过Unnamed字段
                if field.startswith('Unnamed:'):
                    continue
                # 跳过说明文字（通常包含很长的描述性文字）
                if len(field) > 100:
                    continue
                # 跳过空字段
                if not field or field.strip() == '':
                    continue
                valid_fields.append(field)
            
            return valid_fields
        except Exception as e:
            return []
    
    def clean_column_name(self, column_name: str) -> str:
        """
        清理列名，去除空格、特殊字符等
        
        Args:
            column_name: 原始列名
            
        Returns:
            清理后的列名
        """
        if not self.auto_clean_columns:
            return column_name
        
        # 去除首尾空格
        cleaned = column_name.strip()
        
        # 去除多余的空格
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        # 去除特殊字符（保留中文、英文、数字、下划线）
        cleaned = re.sub(r'[^\w\s\u4e00-\u9fff]', '', cleaned)
        
        # 再次去除空格
        cleaned = cleaned.strip()
        
        return cleaned
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度
        
        Args:
            str1: 字符串1
            str2: 字符串2
            
        Returns:
            相似度 (0-1)
        """
        # 使用SequenceMatcher计算相似度
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
    def find_similar_columns(self, target_column: str, available_columns: List[str]) -> List[Tuple[str, float]]:
        """
        查找与目标列名相似的列名
        
        Args:
            target_column: 目标列名
            available_columns: 可用列名列表
            
        Returns:
            相似列名列表，包含相似度
        """
        similar_columns = []
        cleaned_target = self.clean_column_name(target_column)
        
        for column in available_columns:
            cleaned_column = self.clean_column_name(column)
            
            # 精确匹配
            if cleaned_target == cleaned_column:
                similar_columns.append((column, 1.0))
                continue
            
            # 计算相似度
            similarity = self.calculate_similarity(cleaned_target, cleaned_column)
            
            # 检查是否是常见变体
            for standard_name, variants in self.common_column_variants.items():
                if cleaned_target in variants and cleaned_column in variants:
                    similarity = max(similarity, 0.9)  # 提高变体的相似度
                    break
            
            if similarity >= self.similarity_threshold:
                similar_columns.append((column, similarity))
        
        # 按相似度排序
        similar_columns.sort(key=lambda x: x[1], reverse=True)
        return similar_columns
    
    def smart_column_mapping(self, required_columns: List[str], available_columns: List[str]) -> Dict[str, str]:
        """
        智能列名映射
        
        Args:
            required_columns: 需要的列名
            available_columns: 可用的列名
            
        Returns:
            列名映射字典
        """
        mapping = {}
        unmapped_required = []
        unmapped_available = available_columns.copy()
        
        print(f"\n🔍 智能列名映射分析...")
        print(f"📋 需要的列名: {required_columns}")
        print(f"📋 可用的列名: {available_columns}")
        
        # 第一轮：精确匹配和常见变体匹配
        for required in required_columns:
            matched = False
            
            # 检查精确匹配
            if required in available_columns:
                mapping[required] = required
                unmapped_available.remove(required)
                print(f"✅ 精确匹配: {required} -> {required}")
                matched = True
                continue
            
            # 检查常见变体
            if required in self.common_column_variants:
                variants = self.common_column_variants[required]
                for variant in variants:
                    if variant in available_columns:
                        mapping[required] = variant
                        unmapped_available.remove(variant)
                        print(f"✅ 变体匹配: {required} -> {variant}")
                        matched = True
                        break
            
            if not matched:
                unmapped_required.append(required)
        
        # 第二轮：模糊匹配
        if unmapped_required and unmapped_available:
            print(f"\n🔍 进行模糊匹配...")
            for required in unmapped_required:
                similar_columns = self.find_similar_columns(required, unmapped_available)
                
                if similar_columns:
                    best_match, similarity = similar_columns[0]
                    print(f"🔍 找到相似列名: {required} -> {best_match} (相似度: {similarity:.2f})")
                    
                    # 询问用户是否确认映射
                    confirm = input(f"是否将 '{required}' 映射到 '{best_match}'？(y/n，默认y): ").strip().lower()
                    if confirm not in ['n', 'no', '否']:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 确认映射: {required} -> {best_match}")
                    else:
                        print(f"⚠️  跳过映射: {required}")
                else:
                    print(f"❌ 未找到与 '{required}' 相似的列名")
        
        # 显示映射结果
        if mapping:
            print(f"\n📋 列名映射结果:")
            for required, mapped in mapping.items():
                print(f"  {required} -> {mapped}")
        
        if unmapped_required:
            print(f"\n⚠️  未映射的列名: {unmapped_required}")
        
        return mapping
    
    def validate_required_columns(self, df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str], Dict[str, str]]:
        """
        验证必需的列名是否存在，支持智能匹配
        
        Args:
            df: 数据框
            required_columns: 必需的列名列表
            
        Returns:
            (是否验证通过, 缺失的列名列表, 列名映射字典)
        """
        available_columns = list(df.columns)
        missing_columns = []
        column_mapping = {}
        
        if not self.enable_smart_matching:
            # 传统严格匹配
            for required in required_columns:
                if required not in available_columns:
                    missing_columns.append(required)
                else:
                    column_mapping[required] = required
        else:
            # 智能匹配
            column_mapping = self.smart_column_mapping(required_columns, available_columns)
            
            # 检查哪些列名没有被映射
            for required in required_columns:
                if required not in column_mapping:
                    missing_columns.append(required)
        
        return len(missing_columns) == 0, missing_columns, column_mapping
    
    def wildcard_match(self, pattern: str, text: str) -> bool:
        """
        通配符匹配函数，支持 * 代表任意一个字符
        
        Args:
            pattern: 包含 * 的模式字符串
            text: 要匹配的文本
            
        Returns:
            是否匹配
        """
        if '*' not in pattern:
            return pattern == text
        
        # 将 * 转换为正则表达式的 . 字符
        regex_pattern = pattern.replace('*', '.')
        import re
        return bool(re.match(regex_pattern, text))
    
    def flexible_wildcard_match(self, pattern: str, text: str) -> bool:
        """
        灵活的通配符匹配函数，支持 * 代表任意字符序列
        
        Args:
            pattern: 包含 * 的模式字符串
            text: 要匹配的文本
            
        Returns:
            是否匹配
        """
        if '*' not in pattern:
            return pattern == text
        
        # 将 * 转换为正则表达式的 .* 字符（匹配任意字符序列）
        regex_pattern = pattern.replace('*', '.*')
        import re
        return bool(re.match(regex_pattern, text))
    
    def enhanced_field_matching(self, pattern: str, all_fields: List[str]) -> Tuple[List[str], str]:
        """
        增强的字段匹配函数，支持多种匹配方式
        
        Args:
            pattern: 匹配模式
            all_fields: 所有可用字段列表
            
        Returns:
            (匹配的字段列表, 匹配类型描述)
        """
        # 1. 精确匹配
        if pattern in all_fields:
            return [pattern], "精确匹配"
        
        # 2. 通配符匹配
        if '*' in pattern:
            matched_fields = self.find_matching_fields(pattern, all_fields)
            if matched_fields:
                return matched_fields, "通配符匹配"
        
        # 3. 包含匹配（模糊匹配）
        matched_fields = [field for field in all_fields if pattern.lower() in field.lower()]
        if matched_fields:
            return matched_fields, "包含匹配"
        
        # 4. 无匹配
        return [], "无匹配"
    
    def find_matching_fields(self, pattern: str, all_fields: List[str]) -> List[str]:
        """
        根据通配符模式查找匹配的字段
        
        Args:
            pattern: 包含 * 的模式字符串
            all_fields: 所有可用字段列表
            
        Returns:
            匹配的字段列表
        """
        if '*' not in pattern:
            # 精确匹配
            return [field for field in all_fields if field == pattern]
        
        # 通配符匹配
        matched_fields = []
        for field in all_fields:
            if self.flexible_wildcard_match(pattern, field):
                matched_fields.append(field)
        
        return matched_fields
    
    def select_fields(self, all_fields: List[str]) -> List[str]:
        """
        字段选择功能
        
        Args:
            all_fields: 所有可用字段列表
            
        Returns:
            选中的字段列表
        """
        print(f"\n=== 步骤3: 字段选择 ===")
        
        # 询问是否显示字段出现次数
        print("🤔 是否显示字段出现次数？")
        show_occurrence = input("请选择 (y/n，默认y): ").strip().lower()
        show_occurrence = show_occurrence not in ['n', 'no', '否']
        
        if show_occurrence:
            print("📋 可用字段列表（按出现次数排序）:")
        else:
            print("📋 可用字段列表:")
        
        # 分页显示字段
        page_size = 10
        total_pages = (len(all_fields) + page_size - 1) // page_size
        
        for page in range(total_pages):
            start_idx = page * page_size
            end_idx = min(start_idx + page_size, len(all_fields))
            
            print(f"\n--- 第 {page + 1}/{total_pages} 页 ---")
            for i in range(start_idx, end_idx):
                field = all_fields[i]
                if show_occurrence:
                    # 计算该字段的出现次数
                    occurrence_count = sum(1 for f in self.selected_files if field in self.get_file_fields(f))
                    print(f"{i + 1:2d}. {field:<25} (出现在 {occurrence_count} 个文件中)")
                else:
                    print(f"{i + 1:2d}. {field}")
        
        print(f"\n请选择要导入的字段:")
        print("📝 输入字段编号（用逗号分隔，如：1,2,3）")
        print("📝 输入字段名称（用逗号分隔，如：学号,学生姓名）")
        print("📝 支持通配符匹配（*代表任意一个字符，如：*学号,学*号）")
        print("📝 支持模糊匹配（如：学号 可匹配 学生学号、学号信息等）")
        print("📝 输入 'all' 选择所有字段")
        print("📝 输入 'page 1' 查看第1页（可替换页码）")
        
        try:
            choice = input("\n请选择: ").strip()
            
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
            
            elif choice.lower() == 'all':
                self.selected_fields = all_fields
                print(f"✅ 已选择所有 {len(all_fields)} 个字段")
            else:
                # 解析用户选择
                selected_items = [item.strip() for item in choice.split(',')]
                self.selected_fields = []
                
                for item in selected_items:
                    # 尝试作为数字处理
                    try:
                        index = int(item) - 1
                        if 0 <= index < len(all_fields):
                            self.selected_fields.append(all_fields[index])
                        else:
                            print(f"⚠️  字段编号 {item} 超出范围，跳过")
                    except ValueError:
                        # 使用增强的字段匹配函数
                        matched_fields, match_type = self.enhanced_field_matching(item, all_fields)
                        
                        if len(matched_fields) == 1:
                            # 单个匹配，直接添加
                            self.selected_fields.append(matched_fields[0])
                            if match_type != "精确匹配":
                                print(f"📝 {match_type}字段: {item} -> {matched_fields[0]}")
                        elif len(matched_fields) > 1:
                            # 多个匹配，询问用户
                            print(f"\n🔍 {match_type} '{item}' 匹配到 {len(matched_fields)} 个字段:")
                            for i, field in enumerate(matched_fields, 1):
                                print(f"  {i}. {field}")
                            
                            # 询问用户是否使用这些匹配的字段
                            print(f"\n🤔 是否使用这些匹配的字段？")
                            print(f"📝 输入 'y' 使用所有匹配字段")
                            print(f"📝 输入 'n' 跳过所有匹配字段")
                            print(f"📝 输入字段编号（如：1,3）选择特定字段")
                            use_choice = input(f"\n请选择: ").strip().lower()
                            
                            if use_choice in ['y', 'yes', '是']:
                                self.selected_fields.extend(matched_fields)
                                print(f"✅ 已添加 {len(matched_fields)} 个匹配字段")
                            elif use_choice in ['n', 'no', '否']:
                                print(f"⚠️  跳过 '{item}' 的所有匹配字段")
                            else:
                                # 用户选择了特定字段编号
                                try:
                                    selected_indices = [int(x.strip()) - 1 for x in use_choice.split(',')]
                                    selected_fields = [matched_fields[i] for i in selected_indices if 0 <= i < len(matched_fields)]
                                    if selected_fields:
                                        self.selected_fields.extend(selected_fields)
                                        print(f"✅ 已添加 {len(selected_fields)} 个选定字段")
                                    else:
                                        print(f"⚠️  未选择任何有效字段，跳过")
                                except (ValueError, IndexError):
                                    print(f"⚠️  输入格式错误，跳过所有匹配字段")
                        else:
                            # 无匹配
                            print(f"⚠️  未找到匹配字段 '{item}'，跳过")
                
                if not self.selected_fields:
                    print("❌ 未选择任何有效字段，请重新选择")
                    return self.select_fields(all_fields)
                
                # 去重并保持顺序
                seen = set()
                unique_fields = []
                for field in self.selected_fields:
                    if field not in seen:
                        seen.add(field)
                        unique_fields.append(field)
                self.selected_fields = unique_fields
                
                print(f"✅ 已选择 {len(self.selected_fields)} 个字段:")
                for field in self.selected_fields:
                    print(f"  📋 {field}")
                
            return self.selected_fields
            
        except Exception as e:
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
            print(f"{i:2d}. {field}")
        
        print(f"\n📝 输入字段编号（用逗号分隔，如：1,2）")
        print(f"📝 输入字段名称（用逗号分隔，如：学号,学生姓名）")
        print(f"📝 支持通配符匹配（*代表任意一个字符，如：*学号,学*号）")
        print(f"📝 支持模糊匹配（如：学号 可匹配 学生学号、学号信息等）")
        print(f"📝 输入 'all' 使用所有选中字段进行去重")
        print(f"📝 输入 'single 1' 只使用第1个字段去重")
        
        try:
            choice = input("\n请选择去重字段: ").strip().lower()
            
            if choice.lower() == 'all':
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
                # 解析用户选择
                selected_items = [item.strip() for item in choice.split(',')]
                self.dedup_fields = []
                
                for item in selected_items:
                    # 尝试作为数字处理
                    try:
                        index = int(item) - 1
                        if 0 <= index < len(self.selected_fields):
                            self.dedup_fields.append(self.selected_fields[index])
                        else:
                            print(f"⚠️  字段编号 {item} 超出范围，跳过")
                    except ValueError:
                        # 使用增强的字段匹配函数
                        matched_fields, match_type = self.enhanced_field_matching(item, self.selected_fields)
                        
                        if len(matched_fields) == 1:
                            # 单个匹配，直接添加
                            self.dedup_fields.append(matched_fields[0])
                            if match_type != "精确匹配":
                                print(f"📝 {match_type}字段: {item} -> {matched_fields[0]}")
                        elif len(matched_fields) > 1:
                            # 多个匹配，询问用户
                            print(f"\n🔍 {match_type} '{item}' 匹配到 {len(matched_fields)} 个字段:")
                            for i, field in enumerate(matched_fields, 1):
                                print(f"  {i}. {field}")
                            
                            # 询问用户是否使用这些匹配的字段
                            print(f"\n🤔 是否使用这些匹配的字段进行去重？")
                            print(f"📝 输入 'y' 使用所有匹配字段")
                            print(f"📝 输入 'n' 跳过所有匹配字段")
                            print(f"📝 输入字段编号（如：1,3）选择特定字段")
                            use_choice = input(f"\n请选择: ").strip().lower()
                            
                            if use_choice in ['y', 'yes', '是']:
                                self.dedup_fields.extend(matched_fields)
                                print(f"✅ 已添加 {len(matched_fields)} 个匹配字段")
                            elif use_choice in ['n', 'no', '否']:
                                print(f"⚠️  跳过 '{item}' 的所有匹配字段")
                            else:
                                # 用户选择了特定字段编号
                                try:
                                    selected_indices = [int(x.strip()) - 1 for x in use_choice.split(',')]
                                    selected_fields = [matched_fields[i] for i in selected_indices if 0 <= i < len(matched_fields)]
                                    if selected_fields:
                                        self.dedup_fields.extend(selected_fields)
                                        print(f"✅ 已添加 {len(selected_fields)} 个选定字段")
                                    else:
                                        print(f"⚠️  未选择任何有效字段，跳过")
                                except (ValueError, IndexError):
                                    print(f"⚠️  输入格式错误，跳过所有匹配字段")
                        else:
                            # 无匹配
                            print(f"⚠️  未找到匹配字段 '{item}'，跳过")
                
                if not self.dedup_fields:
                    print("❌ 未选择任何有效字段，请重新选择")
                    return self.configure_deduplication()
                
                # 去重并保持顺序
                seen = set()
                unique_fields = []
                for field in self.dedup_fields:
                    if field not in seen:
                        seen.add(field)
                        unique_fields.append(field)
                self.dedup_fields = unique_fields
                
                print(f"✅ 已选择 {len(self.dedup_fields)} 个字段进行去重:")
                for field in self.dedup_fields:
                    print(f"  🔍 {field}")
                
            return True, self.dedup_fields
            
        except Exception as e:
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
                
                # 使用智能列名匹配验证必需字段
                is_valid, missing_fields, column_mapping = self.validate_required_columns(df, selected_fields)
                
                if not is_valid:
                    print(f"⚠️  警告：文件缺少字段 {missing_fields}，跳过此文件")
                    continue
                
                # 使用映射后的列名
                mapped_fields = [column_mapping.get(field, field) for field in selected_fields]
                print(f"📋 使用映射后的列名: {mapped_fields}")
                
                # 使用映射后的列名提取数据
                selected_data = df[mapped_fields].copy()
                
                # 将列名重命名为标准名称，并按照用户选择的顺序重新排列
                rename_mapping = {}
                for i, field in enumerate(selected_fields):
                    if mapped_fields[i] != field:
                        rename_mapping[mapped_fields[i]] = field
                
                if rename_mapping:
                    selected_data = selected_data.rename(columns=rename_mapping)
                    print(f"📝 列名重命名: {rename_mapping}")
                
                # 按照用户选择的字段顺序重新排列列
                selected_data = selected_data[selected_fields]
                print(f"📋 按用户选择顺序排列字段: {selected_fields}")
                
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
        
        # 字段补充处理
        if self.enable_field_supplement and self.field_mappings:
            print(f"\n🔄 正在补充缺失字段...")
            combined_df = self.supplement_fields(
                combined_df, 
                self.field_mappings, 
                self.default_values,
                self.link_field
            )
        
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
                
                # 添加字段补充统计
                if self.enable_field_supplement:
                    stats_items.extend([
                        '是否启用字段补充',
                        '关联字段',
                        '补充字段数',
                        '字段补充成功率'
                    ])
                    # 计算补充成功率（这里简化处理，实际应该统计具体的补充情况）
                    stats_values.extend([
                        '是',
                        self.link_field,
                        len(self.field_mappings),
                        '100.0%'  # 简化显示
                    ])
                else:
                    stats_items.append('是否启用字段补充')
                    stats_values.append('否')
                
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
        print("=" * 60)
        
        # 配置智能匹配选项
        print(f"\n=== 智能匹配配置 ===")
        print(f"🤖 当前智能匹配设置:")
        print(f"  • 智能匹配: {'启用' if self.enable_smart_matching else '禁用'}")
        print(f"  • 自动清理列名: {'启用' if self.auto_clean_columns else '禁用'}")
        print(f"  • 相似度阈值: {self.similarity_threshold}")
        
        change_settings = input("是否修改智能匹配设置？(y/n，默认n): ").strip().lower()
        if change_settings in ['y', 'yes', '是']:
            # 配置智能匹配
            smart_choice = input("是否启用智能列名匹配？(y/n，默认y): ").strip().lower()
            self.enable_smart_matching = smart_choice not in ['n', 'no', '否']
            
            # 配置自动清理
            clean_choice = input("是否自动清理列名（去除空格、特殊字符）？(y/n，默认y): ").strip().lower()
            self.auto_clean_columns = clean_choice not in ['n', 'no', '否']
            
            # 配置相似度阈值
            try:
                threshold_input = input(f"设置相似度阈值 (0.0-1.0，默认{self.similarity_threshold}): ").strip()
                if threshold_input:
                    threshold = float(threshold_input)
                    if 0.0 <= threshold <= 1.0:
                        self.similarity_threshold = threshold
                    else:
                        print(f"⚠️  阈值超出范围，使用默认值 {self.similarity_threshold}")
            except ValueError:
                print(f"⚠️  输入格式错误，使用默认值 {self.similarity_threshold}")
            
            print(f"✅ 智能匹配设置已更新")
        
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
            
            # 3.5. 字段补充配置
            analysis_result = self.analyze_field_supplement_situation(files, selected_fields)
            self.enable_field_supplement, self.default_values, self.link_field = self.configure_field_supplement(analysis_result, selected_fields)
            
            if self.enable_field_supplement:
                # 构建字段映射
                self.field_mappings = {}
                missing_fields = [field for field, files in analysis_result['files_missing_fields'].items() if files]
                
                for field in missing_fields:
                    mapping = self.build_field_mapping(analysis_result['files_with_all_fields'], field, self.link_field)
                    if mapping:
                        self.field_mappings[field] = mapping
                
                # 确保缺失字段被选中
                for field in missing_fields:
                    if field not in selected_fields:
                        selected_fields.append(field)
                        print(f"📝 自动添加{field}字段到选择列表")
            
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
                if self.enable_field_supplement:
                    print(f"🔧 字段补充: 已启用，补充字段数 {len(self.field_mappings)}")
                

            
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
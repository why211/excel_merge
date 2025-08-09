import pandas as pd
import os
import glob
import re
from typing import List, Tuple, Dict, Optional
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
        
        # 重复记录相关属性
        self.duplicate_records = pd.DataFrame()  # 存储发现的重复记录
        self.duplicate_count = 0  # 重复记录数量
        self.enable_interactive_dedup = True  # 是否启用交互式去重
        self.conflict_resolution_choices = {}  # 存储用户的冲突解决选择

        
        # 新增：智能列名匹配相关属性
        self.column_mapping = {}  # 列名映射关系
        self.enable_smart_matching = True  # 是否启用智能匹配
        self.similarity_threshold = 0.8  # 相似度阈值
        self.auto_clean_columns = True  # 是否自动清理列名
        
        # 常见列名变体映射（去重）
        self.common_column_variants = {
            '学号': ['学号', '学号号', '学学号', 'xuehao', 'student_id', '学生编号', '学生学号'],
            '学生姓名': ['学生姓名', '学生姓名名', '学学生姓名', 'student_name', '姓名', '学生名', '学生姓名（中文）'],
            '班级': ['班级', '班', 'class', '班级名称', 'class_name'],
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
                        print(f"✅ 变体匹配: {variant} -> {required}")
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
                    print(f"🔍 找到相似列名: {best_match} -> {required} (相似度: {similarity:.2f})")
                    
                    # 询问用户是否确认映射
                    confirm = input(f"是否将文件列名 '{best_match}' 映射到标准字段 '{required}'？(y/n，默认y): ").strip().lower()
                    if confirm not in ['n', 'no', '否']:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 确认映射: {best_match} -> {required}")
                    else:
                        print(f"⚠️  跳过映射: {required}")
                else:
                    print(f"❌ 未找到与 '{required}' 相似的列名")
                    print(f"🤔 请选择:")
                    print(f"  1. 手动选择列名 (输入 'm')")
                    print(f"  2. 跳过此字段 (输入 's')")
                    
                    while True:
                        choice = input(f"对于字段 '{required}' 请选择: ").strip().lower()
                        if choice == 's':
                            print(f"⚠️  跳过映射: {required}")
                            break
                        elif choice == 'm':
                            selected_column = self._manual_select_column(required, unmapped_available)
                            if selected_column:
                                mapping[required] = selected_column
                                unmapped_available.remove(selected_column)
                                unmapped_required.remove(required)
                                print(f"✅ 手动映射: {selected_column} -> {required}")
                            break
                        else:
                            print("❌ 请输入 'm' 或 's'")
        
        # 显示映射结果
        if mapping:
            print(f"\n📋 列名映射结果:")
            for required, mapped in mapping.items():
                print(f"  {mapped} -> {required}")
        
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
                        print(f"✅ 变体匹配: {variant} -> {required}")
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
                    print(f"🔍 找到相似列名: {best_match} -> {required} (相似度: {similarity:.2f})")
                    
                    # 询问用户是否确认映射
                    confirm = input(f"是否将文件列名 '{best_match}' 映射到标准字段 '{required}'？(y/n，默认y): ").strip().lower()
                    if confirm not in ['n', 'no', '否']:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 确认映射: {best_match} -> {required}")
                    else:
                        print(f"⚠️  跳过映射: {required}")
                else:
                    print(f"❌ 未找到与 '{required}' 相似的列名")
                    print(f"🤔 请选择:")
                    print(f"  1. 手动选择列名 (输入 'm')")
                    print(f"  2. 跳过此字段 (输入 's')")
                    
                    while True:
                        choice = input(f"对于字段 '{required}' 请选择: ").strip().lower()
                        if choice == 's':
                            print(f"⚠️  跳过映射: {required}")
                            break
                        elif choice == 'm':
                            selected_column = self._manual_select_column(required, unmapped_available)
                            if selected_column:
                                mapping[required] = selected_column
                                unmapped_available.remove(selected_column)
                                unmapped_required.remove(required)
                                print(f"✅ 手动映射: {selected_column} -> {required}")
                            break
                        else:
                            print("❌ 请输入 'm' 或 's'")
        
        # 显示映射结果
        if mapping:
            print(f"\n📋 列名映射结果:")
            for required, mapped in mapping.items():
                print(f"  {mapped} -> {required}")
        
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
        print("⚠️  注意：选择 y 可能会导致程序出现卡顿，特别是在处理大量文件时")
        show_occurrence = input("请选择 (y/n，默认n): ").strip().lower()
        show_occurrence = show_occurrence in ['y', 'yes', '是']
        
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
        dedup_choice = input("请选择 (y/n，默认y): ").strip().lower()
        self.deduplicate = dedup_choice not in ['n', 'no', '否']
        
        if not self.deduplicate:
            print("✅ 已选择不去重，将保留所有记录")
            return False, []
        
        # 询问是否启用交互式去重
        print(f"\n🤖 去重模式选择:")
        print(f"📝 自动去重: 自动保留每组的第一条记录")
        print(f"🎯 交互式去重: 当发现字段值冲突时，让您选择如何处理")
        interactive_choice = input("是否启用交互式去重？(y/n，默认y): ").strip().lower()
        self.enable_interactive_dedup = interactive_choice not in ['n', 'no', '否']
        
        if self.enable_interactive_dedup:
            print("✅ 已启用交互式去重，遇到冲突时会询问您的处理方式")
        else:
            print("✅ 使用自动去重模式，将自动保留第一条记录")
        
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
        

        
        # 去重处理
        if deduplicate and dedup_fields:
            print(f"\n🔄 正在按字段 {dedup_fields} 去重...")
            before_count = len(combined_df)
            
            # 查找重复记录
            duplicated_mask = combined_df.duplicated(subset=dedup_fields, keep=False)
            duplicated_records = combined_df[duplicated_mask]
            
            # 保存重复记录到实例变量
            self.duplicate_records = duplicated_records.copy()
            self.duplicate_count = len(duplicated_records)
            
            if len(duplicated_records) > 0:
                print(f"\n" + "🔍" + "="*58)
                print(f"📋 发现重复记录详情")
                print(f"🔍" + "="*58)
                print(f"📊 重复记录总数: {len(duplicated_records)} 条")
                print(f"📊 重复组数量: {duplicated_records.groupby(dedup_fields).ngroups} 组")
                print(f"🔑 去重依据字段: {', '.join(dedup_fields)}")
                
                # 按去重字段分组显示重复记录
                duplicate_groups = duplicated_records.groupby(dedup_fields)
                group_count = 0
                
                for group_key, group_df in duplicate_groups:
                    group_count += 1
                    if group_count <= 10:  # 最多显示前10组重复记录
                        print(f"\n  {'='*50}")
                        print(f"  📝 重复组 {group_count} (共 {len(group_df)} 条重复记录)")
                        print(f"  {'='*50}")
                        
                        # 显示重复字段的值
                        if isinstance(group_key, tuple):
                            for i, field in enumerate(dedup_fields):
                                print(f"  🔑 {field}: {group_key[i]}")
                        else:
                            print(f"  🔑 {dedup_fields[0]}: {group_key}")
                        
                        print(f"  {'-'*40}")
                        
                        # 显示重复记录的详细信息（最多显示3条）
                        display_count = min(3, len(group_df))
                        for idx, (_, row) in enumerate(group_df.head(display_count).iterrows()):
                            print(f"  📄 第 {idx + 1} 条记录:")
                            
                            # 将字段值格式化为表格形式
                            field_values = []
                            for field in combined_df.columns:
                                value = row[field]
                                if pd.notna(value) and str(value).strip():
                                    # 处理过长的值
                                    str_value = str(value)
                                    if len(str_value) > 20:
                                        str_value = str_value[:17] + "..."
                                    field_values.append(f"{field}: {str_value}")
                                else:
                                    field_values.append(f"{field}: <空值>")
                            
                            # 每行显示2个字段，美化显示
                            for i in range(0, len(field_values), 2):
                                line_fields = field_values[i:i+2]
                                if len(line_fields) == 2:
                                    print(f"     {line_fields[0]:<30} | {line_fields[1]}")
                                else:
                                    print(f"     {line_fields[0]}")
                            
                            if idx < display_count - 1:
                                print(f"     {'-'*35}")
                        
                        if len(group_df) > display_count:
                            print(f"  ⚠️  还有 {len(group_df) - display_count} 条重复记录未显示")
                    else:
                        break
                
                if group_count > 10:
                    print(f"\n  ⚠️  还有 {len(duplicate_groups) - 10} 组重复记录未显示")
                
                print(f"\n" + "🔧" + "-"*58)
                if self.enable_interactive_dedup:
                    print(f"💡 交互式去重策略说明:")
                    print(f"   • 对于有字段值冲突的重复组，将询问您的处理方式")
                    print(f"   • 您可以选择保留特定值或创建多条记录")
                    print(f"   • 所有原始重复记录将保存到Excel的'重复记录'工作表中")
                else:
                    print(f"💡 自动去重策略说明:")
                    print(f"   • 保留每组的第一条记录")
                    print(f"   • 删除后续所有重复记录") 
                    print(f"   • 所有重复记录将保存到Excel的'重复记录'工作表中")
                print(f"🔧" + "-"*58)
            
            # 执行去重处理
            if self.enable_interactive_dedup and len(duplicated_records) > 0:
                print(f"\n🎯 开始交互式去重处理...")
                processed_records = []
                duplicate_groups = duplicated_records.groupby(dedup_fields)
                
                for group_key, group_df in duplicate_groups:
                    resolved_records = self.resolve_field_conflicts(group_key, group_df, dedup_fields)
                    if not resolved_records.empty:
                        processed_records.append(resolved_records)
                
                if processed_records:
                    # 重新构建数据框：非重复记录 + 处理后的重复记录
                    non_duplicated_records = combined_df[~duplicated_mask]
                    processed_duplicates = pd.concat(processed_records, ignore_index=True)
                    combined_df = pd.concat([non_duplicated_records, processed_duplicates], ignore_index=True)
                else:
                    # 如果所有重复组都被跳过，只保留非重复记录
                    combined_df = combined_df[~duplicated_mask]
                
                after_count = len(combined_df)
                removed_count = before_count - after_count
            else:
                # 传统自动去重
                combined_df = combined_df.drop_duplicates(subset=dedup_fields, keep='first')
                after_count = len(combined_df)
                removed_count = before_count - after_count
            
            print(f"\n✅ 去重完成:")
            print(f"  📊 去重前行数: {before_count}")
            print(f"  📊 去重后行数: {after_count}")
            print(f"  🗑️  删除重复记录: {removed_count}")
            
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
                    self.duplicate_count - len(df) if self.deduplicate and self.duplicate_count > 0 else 0
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
                
                # 重复记录表（如果有重复记录）
                sheet_names = ['合并数据', '处理统计', '字段信息']
                if not self.duplicate_records.empty:
                    # 添加重复标记列
                    duplicate_export = self.duplicate_records.copy()
                    
                    # 为重复记录添加分组信息
                    if self.dedup_fields:
                        duplicate_groups = duplicate_export.groupby(self.dedup_fields)
                        group_ids = []
                        group_sizes = []
                        
                        for group_id, (group_key, group_df) in enumerate(duplicate_groups, 1):
                            for _ in range(len(group_df)):
                                group_ids.append(group_id)
                                group_sizes.append(len(group_df))
                        
                        duplicate_export.insert(0, '重复组ID', group_ids)
                        duplicate_export.insert(1, '组内重复数', group_sizes)
                    
                    duplicate_export.to_excel(writer, sheet_name='重复记录', index=False)
                    sheet_names.append('重复记录')
                    print(f"📋 重复记录已保存到 '重复记录' 工作表，共 {len(self.duplicate_records)} 条记录")
            
            print(f"✅ 数据已成功导出到: {output_path}")
            print(f"总共导出 {len(df)} 条记录")
            print(f"📋 包含工作表: {', '.join(sheet_names)}")
            
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
        
        # 显示智能匹配配置（默认启用，不询问用户）
        print(f"\n=== 智能匹配配置 ===")
        print(f"🤖 智能匹配设置（已启用）:")
        print(f"  • 智能匹配: {'启用' if self.enable_smart_matching else '禁用'}")
        print(f"  • 自动清理列名: {'启用' if self.auto_clean_columns else '禁用'}")
        print(f"  • 相似度阈值: {self.similarity_threshold}")
        print(f"✅ 使用默认智能匹配设置，提升处理效率")
        
        try:
            # 1. 文件选择
            folder_path = input("请输入包含Excel文件的文件夹路径（或按回车使用当前目录）: ").strip()
            if not folder_path:
                folder_path = "."
            
            files = self.select_files(folder_path)
            if not files:
                print("❌ 未选择任何文件，程序退出")
                return
            
            # 1.5. 文件备份
            if not self.backup_files(files):
                print("❌ 备份失败，程序退出")
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
                    if self.duplicate_count > 0:
                        removed_count = self.duplicate_count - len(result_df)
                        print(f"📊 发现重复记录: {self.duplicate_count} 条")
                        print(f"🗑️  删除重复记录: {removed_count} 条")
                        print(f"💾 重复记录已保存到 '重复记录' 工作表")
                    else:
                        print(f"✅ 未发现重复记录")

                

            
        except KeyboardInterrupt:
            print("\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n❌ 程序执行出错: {str(e)}")
    
    def resolve_field_conflicts(self, group_key, group_df: pd.DataFrame, dedup_fields: List[str]) -> pd.DataFrame:
        """
        解决字段值冲突，让用户选择如何处理不同的字段值
        
        Args:
            group_key: 重复组的键值
            group_df: 重复组的数据框
            dedup_fields: 去重字段列表
            
        Returns:
            处理后的数据框
        """
        if not self.enable_interactive_dedup or len(group_df) <= 1:
            return group_df.head(1)  # 默认保留第一条
        
        # 检查非去重字段是否有冲突
        non_dedup_fields = [field for field in group_df.columns if field not in dedup_fields]
        conflicts = {}
        
        for field in non_dedup_fields:
            # 获取唯一值，保持出现顺序
            seen = set()
            unique_values = []
            for value in group_df[field].dropna():
                if value not in seen:
                    seen.add(value)
                    unique_values.append(value)
            
            if len(unique_values) > 1:
                conflicts[field] = unique_values
        
        if not conflicts:
            return group_df.head(1)  # 没有冲突，保留第一条
        
        print(f"\n{'🔧' + '='*60}")
        print(f"⚠️  发现字段值冲突！")
        print(f"{'🔧' + '='*60}")
        
        # 显示重复组信息
        if isinstance(group_key, tuple):
            for i, field in enumerate(dedup_fields):
                print(f"🔑 {field}: {group_key[i]}")
        else:
            print(f"🔑 {dedup_fields[0]}: {group_key}")
        
        print(f"\n📊 该组有 {len(group_df)} 条记录，以下字段存在不同值:")
        
        # 显示冲突的字段和值
        for field, values in conflicts.items():
            print(f"\n📝 字段 '{field}' 的不同值:")
            for i, value in enumerate(values, 1):
                if pd.isna(value):
                    print(f"  {i}. <空值>")
                else:
                    print(f"  {i}. {value}")
        
        print(f"\n🤔 请选择处理方式:")
        print(f"  1. 保留第一条记录 (默认)")
        print(f"  2. 手动选择每个冲突字段的值")
        print(f"  3. 为每个不同值创建单独记录")
        print(f"  4. 跳过此组，不做处理")
        
        while True:
            try:
                choice = input("\n请选择处理方式 (1-4，默认1): ").strip()
                if not choice:
                    choice = "1"
                
                if choice == "1":
                    print("✅ 保留第一条记录")
                    return group_df.head(1)
                
                elif choice == "2":
                    return self._manual_resolve_conflicts(group_df, conflicts, dedup_fields)
                
                elif choice == "3":
                    print("✅ 为每个不同值创建单独记录")
                    return self._create_separate_records(group_df, conflicts, dedup_fields)
                
                elif choice == "4":
                    print("⚠️  跳过此组")
                    return pd.DataFrame()  # 返回空数据框
                
                else:
                    print("❌ 请输入 1-4 之间的数字")
                    
            except KeyboardInterrupt:
                print("\n⚠️  用户中断，保留第一条记录")
                return group_df.head(1)
    
    def _manual_resolve_conflicts(self, group_df: pd.DataFrame, conflicts: Dict, dedup_fields: List[str]) -> pd.DataFrame:
        """手动解决冲突"""
        result_record = group_df.iloc[0].copy()  # 基于第一条记录
        
        print(f"\n🔧 开始手动解决冲突...")
        print(f"📄 基础记录（第一条）: {dict(result_record)}")
        
        for field, values in conflicts.items():
            print(f"\n📝 请选择字段 '{field}' 的值:")
            print(f"🔍 当前值: {result_record[field]}")
            print(f"📋 可选值:")
            
            for i, value in enumerate(values, 1):
                if pd.isna(value):
                    print(f"  {i}. <空值>")
                else:
                    print(f"  {i}. {value}")
            
            while True:
                try:
                    choice = input(f"请选择 (1-{len(values)}): ").strip()
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(values):
                        selected_value = values[choice_idx]
                        old_value = result_record[field]
                        result_record[field] = selected_value
                        
                        if pd.isna(selected_value):
                            print(f"✅ 已选择: <空值>")
                        else:
                            print(f"✅ 已选择: {selected_value}")
                        
                        print(f"🔄 字段 '{field}' 更新: {old_value} → {selected_value}")
                        break
                    else:
                        print("❌ 编号超出范围，请重新选择")
                except ValueError:
                    print("❌ 请输入有效的数字")
        
        print(f"\n✅ 冲突解决完成！")
        print(f"📄 最终记录: {dict(result_record)}")
        return pd.DataFrame([result_record])
    
    def _create_separate_records(self, group_df: pd.DataFrame, conflicts: Dict, dedup_fields: List[str]) -> pd.DataFrame:
        """为不同值创建单独记录"""
        # 找到最多值的字段作为主字段
        main_field = max(conflicts.keys(), key=lambda f: len(conflicts[f]))
        main_values = conflicts[main_field]
        
        print(f"📝 以字段 '{main_field}' 为主字段创建 {len(main_values)} 条记录")
        
        result_records = []
        base_record = group_df.iloc[0].copy()
        
        for i, main_value in enumerate(main_values):
            new_record = base_record.copy()
            new_record[main_field] = main_value
            
            # 为其他冲突字段选择对应的值
            for field in conflicts:
                if field != main_field:
                    # 找到与main_value对应的记录中该字段的值
                    matching_records = group_df[group_df[main_field] == main_value]
                    if not matching_records.empty:
                        new_record[field] = matching_records.iloc[0][field]
                    # 如果没有完全匹配的记录，保持原值
            
            result_records.append(new_record)
            print(f"  📄 记录 {i+1}: {main_field}={main_value}")
        
        return pd.DataFrame(result_records)

    def backup_files(self, files: List[str]) -> bool:
        """
        备份选中的Excel文件
        
        Args:
            files: 要备份的文件列表
            
        Returns:
            备份是否成功
        """
        print(f"\n=== 文件备份 ===")
        
        # 询问是否要备份
        backup_choice = input("🤔 是否要备份选中的Excel文件？(y/n，默认y): ").strip().lower()
        if backup_choice in ['n', 'no', '否']:
            print("✅ 跳过备份，直接处理文件")
            return True
        
        # 创建备份目录
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = f"backup_{timestamp}"
        
        try:
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            print(f"📁 创建备份目录: {backup_dir}")
            
            # 备份每个文件
            backup_success = 0
            backup_failed = 0
            
            for file_path in files:
                try:
                    filename = os.path.basename(file_path)
                    backup_path = os.path.join(backup_dir, filename)
                    
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
                    print(f"✅ 已备份: {filename} -> {os.path.basename(backup_path)}")
                    backup_success += 1
                    
                except Exception as e:
                    print(f"❌ 备份失败: {os.path.basename(file_path)} - {str(e)}")
                    backup_failed += 1
            
            print(f"\n📊 备份结果:")
            print(f"  ✅ 成功备份: {backup_success} 个文件")
            if backup_failed > 0:
                print(f"  ❌ 备份失败: {backup_failed} 个文件")
            print(f"  📁 备份位置: {os.path.abspath(backup_dir)}")
            
            if backup_failed > 0:
                continue_choice = input("\n⚠️  部分文件备份失败，是否继续处理？(y/n，默认y): ").strip().lower()
                if continue_choice in ['n', 'no', '否']:
                    print("❌ 用户选择退出")
                    return False
            
            return True
            
        except Exception as e:
            print(f"❌ 创建备份目录失败: {str(e)}")
            continue_choice = input("⚠️  备份失败，是否继续处理？(y/n，默认n): ").strip().lower()
            return continue_choice in ['y', 'yes', '是']

    def _manual_select_column(self, required_field: str, available_columns: List[str]) -> str:
        """手动选择列名"""
        if not available_columns:
            print(f"  ⚠️  没有可用的列名可选择")
            return None
        
        print(f"\n  📋 可用的列名:")
        for i, column in enumerate(available_columns, 1):
            print(f"    {i:2d}. {column}")
        
        print(f"\n  📝 请选择要映射到字段 '{required_field}' 的列名:")
        while True:
            try:
                choice = input("  请输入列名编号: ").strip()
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(available_columns):
                    selected_column = available_columns[choice_idx]
                    print(f"  ✅ 选择了列名: {selected_column}")
                    return selected_column
                else:
                    print("  ❌ 编号超出范围，请重新选择")
            except ValueError:
                print("  ❌ 请输入有效的数字")

def main():
    """主函数"""
    processor = ExcelProcessor()
    processor.run()

if __name__ == "__main__":
    main()
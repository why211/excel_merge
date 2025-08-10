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
        

        
        # 第一轮：精确匹配和常见变体匹配
        for required in required_columns:
            matched = False
            
            # 检查精确匹配
            if required in available_columns:
                mapping[required] = required
                unmapped_available.remove(required)

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
                    
                    # 如果相似度为1.00，自动确认映射
                    if similarity >= 1.0:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 自动映射 (完全匹配): {best_match} -> {required}")
                    else:
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
                    
                    # 如果相似度为1.00，自动确认映射
                    if similarity >= 1.0:
                        mapping[required] = best_match
                        unmapped_available.remove(best_match)
                        print(f"✅ 自动映射 (完全匹配): {best_match} -> {required}")
                    else:
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
        print(f"📝 自动去重: 学号+姓名相同的记录自动合并，学号相同但姓名不同的保留第一条")
        print(f"🎯 交互式去重: 学号+姓名相同的记录自动合并，学号相同但姓名不同时询问处理方式")
        interactive_choice = input("是否启用交互式去重？(y/n，默认y): ").strip().lower()
        self.enable_interactive_dedup = interactive_choice not in ['n', 'no', '否']
        
        if self.enable_interactive_dedup:
            print("✅ 已启用交互式去重，学号相同但姓名不同时会询问您的处理方式")
        else:
            print("✅ 使用自动去重模式，学号相同但姓名不同时将自动保留第一条记录")
        
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
                    print(f"⚠️  警告：文件缺少字段 {missing_fields}")
                    
                    # 询问用户是否要处理此文件
                    print(f"🤔 是否要处理此文件？")
                    print(f"  1. 是，为缺失字段填充默认值")
                    print(f"  2. 否，跳过此文件")
                    
                    while True:
                        try:
                            choice = input("请选择 (1-2，默认1): ").strip()
                            if not choice:
                                choice = "1"
                            
                            if choice == "1":
                                print("✅ 继续处理，为缺失字段填充默认值")
                                break
                            elif choice == "2":
                                print("⏭️  跳过此文件")
                                continue
                            else:
                                print("❌ 无效选择，请输入 1 或 2")
                        except (EOFError, KeyboardInterrupt):
                            print("✅ 使用默认选择：继续处理")
                            break
                    
                    # 为缺失字段填充默认值
                    for field in missing_fields:
                        if field not in column_mapping:
                            # 根据字段类型填充合适的默认值
                            if self._is_money_field(field):
                                default_value = 0
                            elif "名称" in field or "姓名" in field:
                                default_value = "<空值>"
                            elif "编号" in field or "ID" in field:
                                default_value = "<空值>"
                            else:
                                default_value = "<空值>"
                            
                            # 在数据框中添加缺失字段，填充默认值
                            df[field] = default_value
                            print(f"📝 为缺失字段 '{field}' 填充默认值: {default_value}")
                    
                    # 重新验证字段
                    is_valid, missing_fields, column_mapping = self.validate_required_columns(df, selected_fields)
                    if not is_valid:
                        print(f"❌ 字段验证仍然失败，跳过此文件")
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
                
                # 添加文件来源信息
                selected_data['数据来源文件'] = os.path.basename(file)
                selected_data['数据来源路径'] = os.path.abspath(file)
                
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
            
            # 智能识别学号和姓名字段
            student_id_field = None
            student_name_field = None
            
            # 智能识别学号字段
            student_id_field = self._identify_student_id_field(dedup_fields, combined_df.columns)
            
            # 智能识别姓名字段
            student_name_field = self._identify_name_field(combined_df.columns)
            
            # 如果没有找到学号字段，使用第一个去重字段作为主键字段
            if not student_id_field and dedup_fields:
                student_id_field = dedup_fields[0]
            
            print(f"📋 检测到的字段:")
            if student_id_field:
                field_icon = self._get_field_icon(student_id_field)
                print(f"  {field_icon} 主键字段: {student_id_field}")
            else:
                print(f"  🔑 主键字段: {dedup_fields[0] if dedup_fields else 'None'}")
            
            if student_name_field:
                field_icon = self._get_field_icon(student_name_field)
                print(f"  {field_icon} 姓名字段: {student_name_field}")
            else:
                print(f"  👤 姓名字段: None")
            
            # 查找重复记录（基于去重字段）
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
                conflict_group_count = 0  # 有冲突的组数量
                
                for group_key, group_df in duplicate_groups:
                    # 检查这个组是否有真正的冲突（学号相同但姓名不同）
                    has_conflict = self._group_has_student_name_conflict(group_df, dedup_fields, student_name_field)
                    
                    if has_conflict:
                        conflict_group_count += 1
                        if conflict_group_count <= 10:  # 最多显示前10组有冲突的重复记录
                            print(f"\n  {'='*50}")
                            print(f"  📝 冲突重复组 {conflict_group_count} (共 {len(group_df)} 条重复记录)")
                            print(f"  {'='*50}")
                            
                            # 显示重复字段的值
                            if isinstance(group_key, tuple):
                                for i, field in enumerate(dedup_fields):
                                    display_value = self._format_display_value(group_key[i])
                                    print(f"  🔑 {field}: {display_value}")
                            else:
                                display_value = self._format_display_value(group_key)
                                print(f"  🔑 {dedup_fields[0]}: {display_value}")
                            
                            # 定义非去重字段列表（在使用前定义）
                            non_dedup_fields = [field for field in group_df.columns if field not in dedup_fields]
                            
                            # 显示涉及的文件
                            if '数据来源文件' in group_df.columns:
                                # 基于文件名+路径去重
                                file_refs = group_df[['数据来源文件', '数据来源路径']].drop_duplicates()
                                # 先做逐文件校验，得到可用文件清单
                                verified_by_path: Dict[str, bool] = {}
                                for _, ref in file_refs.iterrows():
                                    full_path = str(ref['数据来源路径'])
                                    try:
                                        ok = self._verify_group_key_in_file(full_path, dedup_fields, group_key)
                                    except Exception:
                                        ok = False
                                    verified_by_path[full_path] = ok

                                # 仅展示校验通过的文件
                                verified_files = [str(ref['数据来源文件']) for _, ref in file_refs.iterrows() if verified_by_path.get(str(ref['数据来源路径']), False)]
                                skipped_files = [str(ref['数据来源文件']) for _, ref in file_refs.iterrows() if not verified_by_path.get(str(ref['数据来源路径']), False)]

                                if verified_files:
                                    print(f"  📁 涉及文件: {', '.join(verified_files)}")
                                if skipped_files:
                                    print(f"  ⚠️ 已忽略未在源文件找到的文件: {', '.join(skipped_files)}")
                                
                                # 调试信息：显示每个文件的记录数和具体内容（并校验是否真实存在）
                                print(f"  🔍 详细分布:")
                                for _, ref in file_refs.iterrows():
                                    base_name = str(ref['数据来源文件'])
                                    full_path = str(ref['数据来源路径'])
                                    file_records = group_df[group_df['数据来源路径'] == full_path]
                                    exists_in_src = verified_by_path.get(full_path, False)
                                    # 只显示校验通过的文件详情
                                    if not exists_in_src:
                                        continue
                                    print(f"     • {base_name}: {len(file_records)} 条记录")
                                    print(f"       校验: ✅ 已在源文件找到")
                                    
                                    # 显示该文件中的具体记录内容（显示所有字段用于调试）
                                    for idx, (_, record) in enumerate(file_records.iterrows()):
                                        if idx >= 2:  # 最多显示2条记录
                                            if len(file_records) > 2:
                                                print(f"       ... 还有 {len(file_records) - 2} 条记录")
                                            break
                                        
                                        record_info = []
                                        # 显示所有字段（包括去重字段）用于调试
                                        for field in group_df.columns:
                                            if field in ('数据来源文件', '数据来源路径'):
                                                continue
                                            value = record[field]
                                            if pd.notna(value) and str(value).strip():
                                                display_value = self._format_display_value(value)
                                                record_info.append(f"{field}={display_value}")
                                            else:
                                                record_info.append(f"{field}=<空值>")
                                        
                                        print(f"       [{idx+1}] {', '.join(record_info)}")
                            
                            print(f"  {'-'*40}")
                            
                            # 调试：显示数据框的完整结构信息
                            print(f"  🔧 调试信息:")
                            # 只基于校验通过的行统计
                            if '数据来源路径' in group_df.columns:
                                verified_mask = group_df['数据来源路径'].map(lambda p: verified_by_path.get(str(p), False))
                                group_df_verified = group_df[verified_mask] if verified_mask.any() else group_df.iloc[0:0]
                            else:
                                group_df_verified = group_df

                            print(f"     • 数据框形状: {group_df_verified.shape}")
                            print(f"     • 所有字段: {list(group_df.columns)}")
                            print(f"     • 去重字段: {dedup_fields}")
                            print(f"     • 非去重字段: {non_dedup_fields}")
                            
                            # 分析并显示冲突的具体情况
                            conflict_summary = {}
                            
                            # 找出每个字段的不同值（排除文件来源字段）
                            for field in non_dedup_fields:
                                if field in ('数据来源文件', '数据来源路径'):  # 跳过文件来源字段
                                    continue
                                unique_vals = []
                                seen = set()
                                for value in group_df_verified[field] if not group_df_verified.empty else []:
                                    if pd.isna(value):
                                        str_val = "<空值>"
                                    else:
                                        str_val = str(value).strip()
                                    if str_val not in seen:
                                        seen.add(str_val)
                                        unique_vals.append(str_val)
                                
                                if len([v for v in unique_vals if v != "<空值>"]) > 1:
                                    conflict_summary[field] = unique_vals
                            
                            # 显示冲突字段的不同值（汇总：每个取值的数量与来源文件）
                            if conflict_summary:
                                print(f"  🔍 冲突字段及其不同值（按取值统计）:")
                                for field in conflict_summary:
                                    # 为该字段统计不同取值的数量与来源文件
                                    value_to_count: Dict[str, int] = {}
                                    value_to_files: Dict[str, set] = {}

                                    for _, row in group_df_verified.iterrows() if not group_df_verified.empty else []:
                                        raw_val = row[field]
                                        if pd.isna(raw_val) or (isinstance(raw_val, str) and raw_val.strip() == ""):
                                            disp_val = "<空值>"
                                        else:
                                            disp_val = self._format_display_value(raw_val).strip()

                                        value_to_count[disp_val] = value_to_count.get(disp_val, 0) + 1
                                        if '数据来源文件' in group_df.columns:
                                            src_file = row['数据来源文件']
                                            value_to_files.setdefault(disp_val, set()).add(str(src_file))

                                    # 仅保留非空值用于冲突展示
                                    non_empty_items = [(v, c) for v, c in value_to_count.items() if v != "<空值>"]
                                    # 按数量降序
                                    non_empty_items.sort(key=lambda x: x[1], reverse=True)

                                    print(f"     • {field}: 共 {len(non_empty_items)} 种不同取值")
                                    for val, cnt in non_empty_items:
                                        files_list = sorted(list(value_to_files.get(val, [])))
                                        files_str = ", ".join(files_list) if files_list else "-"
                                        print(f"       - {val}: {cnt} 条 (来源: {files_str})")

                                print(f"  {'-'*40}")

                            # 统计说明（不再展示样本记录，避免重复与误解）
                            total_shown = len(group_df_verified) if not group_df_verified.empty else 0
                            print(f"  💡 已基于校验通过的 {total_shown} 条记录进行统计展示。")

                            # 显示统计信息
                            if len(group_df_verified) > 0:
                                remaining = 0  # 已以汇总方式展示，不再单独显示样本与剩余条目
                                if remaining > 0:
                                    print(f"  💡 还有 {remaining} 条记录与上述取值重复")
                
                # 更新统计信息显示
                total_duplicate_groups = duplicated_records.groupby(dedup_fields).ngroups
                if conflict_group_count > 0:
                    print(f"\n📊 统计信息:")
                    print(f"  📋 总重复组数: {total_duplicate_groups}")
                    print(f"  ⚠️  有冲突的重复组: {conflict_group_count}")
                    print(f"  ✅ 完全相同的重复组: {total_duplicate_groups - conflict_group_count}")
                    if conflict_group_count > 10:
                        print(f"  💡 只显示了前10组有冲突的重复记录")
                else:
                    print(f"\n✅ 所有重复记录都是完全相同的，将自动去除，无需用户处理")
                
                # 去重策略说明已移除
            
            # 执行去重处理
            if len(duplicated_records) > 0:
                processed_records = []
                duplicate_groups = duplicated_records.groupby(dedup_fields)
                conflicts_found = 0
                
                for group_key, group_df in duplicate_groups:
                    resolved_records, had_conflict = self.resolve_student_conflicts(group_key, group_df, dedup_fields, student_name_field, student_id_field)
                    if not resolved_records.empty:
                        processed_records.append(resolved_records)
                    if had_conflict:
                        conflicts_found += 1
                
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
                
                # 显示处理结果
                if conflicts_found > 0:
                    print(f"\n🔄 去重处理完成:")
                    if student_id_field and student_name_field:
                        id_icon = self._get_field_icon(student_id_field)
                        name_icon = self._get_field_icon(student_name_field)
                        print(f"  📊 发现{name_icon}冲突的{id_icon}: {conflicts_found} 个")
                    else:
                        print(f"  📊 发现字段冲突的重复组: {conflicts_found} 个")
                    print(f"  ✅ 自动合并的重复记录: {len(duplicate_groups) - conflicts_found} 组")
                else:
                    if student_id_field and student_name_field:
                        id_icon = self._get_field_icon(student_id_field)
                        name_icon = self._get_field_icon(student_name_field)
                        print(f"\n✅ 去重处理完成: 所有重复记录都是{id_icon}+{name_icon}完全相同，已自动合并")
                    else:
                        print(f"\n✅ 去重处理完成: 所有重复记录都是完全相同的，已自动合并")
                
                # 更新重复记录统计，避免导出时长度不匹配
                # 重新计算实际被处理的重复记录
                if processed_records:
                    # 保存原始的重复记录用于导出
                    self.duplicate_records = duplicated_records.copy()
                    # 更新重复记录数量为实际处理的数量
                    self.duplicate_count = len(duplicated_records)
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
                        try:
                            duplicate_groups = duplicate_export.groupby(self.dedup_fields)
                            group_ids = []
                            group_sizes = []
                            
                            for group_id, (group_key, group_df) in enumerate(duplicate_groups, 1):
                                for _ in range(len(group_df)):
                                    group_ids.append(group_id)
                                    group_sizes.append(len(group_df))
                            
                            # 确保长度匹配
                            if len(group_ids) == len(duplicate_export):
                                duplicate_export.insert(0, '重复组ID', group_ids)
                                duplicate_export.insert(1, '组内重复数', group_sizes)
                            else:
                                print(f"⚠️  重复记录分组信息长度不匹配，跳过分组标记")
                                print(f"   记录数: {len(duplicate_export)}, 分组标记数: {len(group_ids)}")
                        except Exception as e:
                            print(f"⚠️  处理重复记录分组信息时出错: {str(e)}")
                            print(f"   将导出原始重复记录，不包含分组信息")
                    
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
        if len(group_df) <= 1:
            return group_df.head(1)  # 只有一条记录，直接返回
        
        if not self.enable_interactive_dedup and not self.enable_smart_dedup:
            return group_df.head(1)  # 默认保留第一条
        
        # 检查是否所有记录完全相同（排除数据来源文件字段）
        all_fields = [field for field in group_df.columns if field != '数据来源文件']
        first_record = group_df.iloc[0]
        
        # 检查是否所有记录都与第一条记录完全相同
        all_identical = True
        for _, row in group_df.iterrows():
            for field in all_fields:
                # 处理NaN值的比较
                first_val = first_record[field]
                current_val = row[field]
                
                # 如果两个值都是NaN，认为相同
                if pd.isna(first_val) and pd.isna(current_val):
                    continue
                # 如果一个是NaN另一个不是，认为不同
                elif pd.isna(first_val) or pd.isna(current_val):
                    all_identical = False
                    break
                # 如果两个值都不是NaN，比较字符串形式
                elif str(first_val).strip() != str(current_val).strip():
                    all_identical = False
                    break
            
            if not all_identical:
                break
        
        if all_identical:
            # 所有记录完全相同，这是真正的重复，直接保留第一条
            return group_df.head(1)
        
        # 检查非去重字段是否有冲突（排除文件来源字段）
        non_dedup_fields = [field for field in group_df.columns if field not in dedup_fields and field != '数据来源文件']
        conflicts = {}
        
        for field in non_dedup_fields:
            # 获取唯一值，保持出现顺序，包括NaN值的处理
            seen = set()
            unique_values = []
            
            for value in group_df[field]:
                # 处理NaN值
                if pd.isna(value):
                    str_value = "<NaN>"
                else:
                    str_value = str(value).strip()
                
                if str_value not in seen:
                    seen.add(str_value)
                    unique_values.append(value)
            
            # 只有当确实有不同的非NaN值时才认为是冲突
            non_nan_values = [v for v in unique_values if not pd.isna(v)]
            if len(non_nan_values) > 1:
                conflicts[field] = unique_values
        
        if not conflicts:
            return group_df.head(1)  # 没有冲突，保留第一条
        
        # 这个函数现在已经被 resolve_student_conflicts 替代
        # 直接返回第一条记录作为后备方案
        print("⚠️  使用后备处理方案：保留第一条记录")
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
    
    def _keep_most_frequent_values(self, group_df: pd.DataFrame, conflicts: Dict, dedup_fields: List[str]) -> pd.DataFrame:
        """保留出现次数最多的值"""
        result_record = group_df.iloc[0].copy()  # 基于第一条记录
        
        print(f"\n🔧 开始按出现次数最多的值解决冲突...")
        
        for field, values in conflicts.items():
            # 统计每个值的出现次数
            value_counts = {}
            for _, row in group_df.iterrows():
                value = row[field]
                # 归一化值用于比较
                if pd.isna(value):
                    normalized_value = "<空值>"
                else:
                    normalized_value = str(value).strip()
                
                value_counts[normalized_value] = value_counts.get(normalized_value, 0) + 1
            
            # 找到出现次数最多的值
            most_frequent_normalized = max(value_counts.keys(), key=lambda k: value_counts[k])
            most_frequent_count = value_counts[most_frequent_normalized]
            
            # 找到对应的原始值
            if most_frequent_normalized == "<空值>":
                most_frequent_original = None
            else:
                # 在原始数据中找到第一个匹配的值
                most_frequent_original = None
                for _, row in group_df.iterrows():
                    value = row[field]
                    if not pd.isna(value) and str(value).strip() == most_frequent_normalized:
                        most_frequent_original = value
                        break
                if most_frequent_original is None:
                    most_frequent_original = most_frequent_normalized
            
            # 更新结果记录
            old_value = result_record[field]
            result_record[field] = most_frequent_original
            
            print(f"📊 字段 '{field}': 选择出现次数最多的值")
            print(f"   • 选择的值: {self._format_display_value(most_frequent_original)} (出现 {most_frequent_count} 次)")
            print(f"   • 其他值的统计:")
            for norm_val, count in sorted(value_counts.items(), key=lambda x: x[1], reverse=True)[1:]:
                print(f"     - {norm_val}: {count} 次")
            print(f"🔄 字段 '{field}' 更新: {self._format_display_value(old_value)} → {self._format_display_value(most_frequent_original)}")
        
        print(f"\n✅ 冲突解决完成！已选择出现次数最多的值")
        return pd.DataFrame([result_record])
    
    def resolve_student_conflicts(self, group_key, group_df: pd.DataFrame, dedup_fields: List[str], student_name_field: str, student_id_field: str = None) -> tuple:
        """
        解决学生记录冲突：学号相同但姓名不同的情况
        
        Args:
            group_key: 重复组的键值
            group_df: 重复组的数据框
            dedup_fields: 去重字段列表
            student_name_field: 学生姓名字段名
            student_id_field: 学生学号字段名（可选）
            
        Returns:
            (处理后的数据框, 是否有冲突)
        """
        if len(group_df) <= 1:
            return group_df, False  # 只有一条记录，直接返回
        
        # 检查是否有姓名冲突
        has_name_conflict = self._group_has_student_name_conflict(group_df, dedup_fields, student_name_field)
        
        if not has_name_conflict:
            # 没有姓名冲突，学号+姓名完全相同，静默合并（保留第一条）
            return group_df.head(1), False
        
        # 有冲突，需要处理
        print(f"\n{'⚠️' + '='*60}")
        
        # 智能判断冲突类型
        if student_id_field and student_name_field:
            id_icon = self._get_field_icon(student_id_field)
            name_icon = self._get_field_icon(student_name_field)
            print(f"发现{id_icon}相同但{name_icon}不同的记录！")
        else:
            print(f"发现重复记录存在字段冲突！")
        
        print(f"{'⚠️' + '='*60}")
        
        # 显示主键信息
        if isinstance(group_key, tuple):
            for i, field in enumerate(dedup_fields):
                display_value = self._format_display_value(group_key[i])
                print(f"🔑 {field}: {display_value}")
        else:
            display_value = self._format_display_value(group_key)
            print(f"🔑 {dedup_fields[0]}: {display_value}")
        
        # 显示冲突的字段信息
        conflict_info = {}
        exclude_fields = set(['数据来源文件', '数据来源路径'] + dedup_fields)
        
        for field in group_df.columns:
            if field in exclude_fields:
                continue
            
            unique_values = set()
            for value in group_df[field]:
                # 修改：包含空值，因为空值也是一种有效的值，需要用户选择
                if pd.notna(value):
                    normalized_value = str(value).strip() if str(value).strip() else "<空值>"
                    unique_values.add(normalized_value)
                else:
                    unique_values.add("<空值>")
            
            if len(unique_values) > 1:
                conflict_info[field] = unique_values
        
        if conflict_info:
            print(f"\n📋 发现冲突的字段:")
            for field, values in conflict_info.items():
                # 使用辅助函数智能选择图标
                field_icon = self._get_field_icon(field)
                
                print(f"  {field_icon} {field}: {len(values)} 个不同值")
                for i, value in enumerate(sorted(values), 1):
                    print(f"    {i}. {value}")
        
        # 如果有姓名字段，显示姓名冲突详情
        if student_name_field and student_name_field in group_df.columns:
            unique_names = {}
            for _, row in group_df.iterrows():
                name = row[student_name_field]
                # 修改：包含空值，因为空值也是一种有效的值，需要用户选择
                if pd.notna(name):
                    normalized_name = str(name).strip() if str(name).strip() else "<空值>"
                else:
                    normalized_name = "<空值>"
                
                if normalized_name not in unique_names:
                    unique_names[normalized_name] = []
                unique_names[normalized_name].append(row)
            
            if len(unique_names) > 1:
                # 使用辅助函数智能选择图标
                field_icon = self._get_field_icon(student_name_field)
                print(f"\n{field_icon} 发现 {len(unique_names)} 个不同的值:")
                
                for i, (name, records) in enumerate(unique_names.items(), 1):
                    # 统计该姓名出现的文件
                    files = set()
                    for record in records:
                        if '数据来源文件' in record:
                            files.add(str(record['数据来源文件']))
                    
                    print(f"  {i}. {name} (出现在 {len(records)} 条记录中)")
                    if files:
                        print(f"     来源文件: {', '.join(sorted(files))}")
        
        if not self.enable_interactive_dedup:
            # 自动模式：保留第一条记录
            print(f"\n✅ 自动模式：保留第一条记录")
            return group_df.head(1), True
        
        # 交互式模式：询问用户如何处理
        print(f"\n🤔 请选择处理方式:")
        print(f"  1. 保留第一条记录 (默认)")
        
        # 智能判断字段类型并显示相应选项
        if student_name_field:
            field_icon = self._get_field_icon(student_name_field)
            print(f"  2. 手动选择要保留的记录")
            print(f"  3. 为每个不同值创建单独记录")
        else:
            print(f"  2. 手动选择要保留的记录")
            print(f"  3. 为每个不同值创建单独记录")
        
        print(f"  4. 跳过此组，不做处理")
        
        while True:
            try:
                choice = input("\n请选择处理方式 (1-4，默认1): ").strip()
                if not choice:
                    choice = "1"
                
                if choice == "1":
                    print("✅ 保留第一条记录")
                    
                    # 检查是否还有其他冲突字段需要处理
                    if student_name_field:
                        # 如果有姓名字段，检查其他冲突字段
                        conflict_info = self._get_remaining_conflicts(group_df, [group_df.iloc[0]], student_name_field)
                        
                        if conflict_info:
                            print(f"\n⚠️  发现其他冲突字段，需要进一步处理:")
                            for field, values in conflict_info.items():
                                field_icon = self._get_field_icon(field)
                                print(f"  {field_icon} {field}: {len(values)} 个不同值")
                                for i, value in enumerate(sorted(values), 1):
                                    print(f"    {i}. {value}")
                            
                            # 询问用户是否要处理其他冲突字段
                            print(f"\n🤔 是否要处理其他冲突字段？")
                            print(f"  1. 是，手动选择每个字段的值")
                            print(f"  2. 否，使用第一条记录的值")
                            
                            conflict_choice = input("请选择 (1-2，默认2): ").strip()
                            if conflict_choice == "1":
                                # 手动处理其他冲突字段
                                result_record = self._manual_resolve_remaining_conflicts(group_df.iloc[0], conflict_info)
                                return pd.DataFrame([result_record]), True
                            else:
                                # 使用第一条记录
                                print("✅ 使用第一条记录的值")
                                return group_df.head(1), True
                        else:
                            # 没有其他冲突字段，直接返回第一条记录
                            return group_df.head(1), True
                    else:
                        # 没有姓名字段，直接返回第一条记录
                        return group_df.head(1), True
                
                elif choice == "2":
                    if student_name_field:
                        result = self._manual_select_student_name(group_df, unique_names, student_name_field)
                        # 检查是否还有其他冲突字段需要处理
                        if hasattr(result, 'iloc') and len(result) > 0:
                            remaining_conflicts = self._get_remaining_conflicts(group_df, [result.iloc[0]], student_name_field)
                            if remaining_conflicts:
                                print(f"\n⚠️  发现其他冲突字段，需要进一步处理:")
                                for field, values in remaining_conflicts.items():
                                    field_icon = self._get_field_icon(field)
                                    print(f"  {field_icon} {field}: {len(values)} 个不同值")
                                    for i, value in enumerate(sorted(values), 1):
                                        print(f"    {i}. {value}")
                                
                                # 询问用户是否要处理其他冲突字段
                                print(f"\n🤔 是否要处理其他冲突字段？")
                                print(f"  1. 是，手动选择每个字段的值")
                                print(f"  2. 否，使用已选择记录的值")
                                
                                conflict_choice = input("请选择 (1-2，默认2): ").strip()
                                if conflict_choice == "1":
                                    # 手动处理其他冲突字段
                                    result_record = self._manual_resolve_remaining_conflicts(result.iloc[0], remaining_conflicts)
                                    return pd.DataFrame([result_record]), True
                                else:
                                    # 使用已选择的记录
                                    print("✅ 使用已选择记录的值")
                                    return result, True
                    else:
                        result = self._manual_select_record(group_df, conflict_info)
                    return result, True
                
                elif choice == "3":
                    if student_name_field:
                        print("✅ 为每个不同值创建单独记录")
                        result = self._create_records_by_name(group_df, unique_names, student_name_field)
                        # 检查是否还有其他冲突字段需要处理
                        if len(result) > 0:
                            # 为每个记录检查其他冲突字段
                            final_records = []
                            for _, record in result.iterrows():
                                remaining_conflicts = self._get_remaining_conflicts(group_df, [record], student_name_field)
                                if remaining_conflicts:
                                    print(f"\n⚠️  记录 '{record[student_name_field]}' 发现其他冲突字段:")
                                    for field, values in remaining_conflicts.items():
                                        field_icon = self._get_field_icon(field)
                                        print(f"  {field_icon} {field}: {len(values)} 个不同值")
                                    
                                    # 询问用户是否要处理其他冲突字段
                                    print(f"\n🤔 是否要处理记录 '{record[student_name_field]}' 的其他冲突字段？")
                                    print(f"  1. 是，手动选择每个字段的值")
                                    print(f"  2. 否，使用当前记录的值")
                                    
                                    conflict_choice = input("请选择 (1-2，默认2): ").strip()
                                    if conflict_choice == "1":
                                        # 手动处理其他冲突字段
                                        resolved_record = self._manual_resolve_remaining_conflicts(record, remaining_conflicts)
                                        final_records.append(resolved_record)
                                    else:
                                        # 使用当前记录
                                        print("✅ 使用当前记录的值")
                                        final_records.append(record)
                                else:
                                    final_records.append(record)
                            
                            if final_records:
                                result = pd.DataFrame(final_records)
                    else:
                        print("✅ 为每个不同值创建单独记录")
                        result = self._create_records_by_conflict_fields(group_df, conflict_info)
                    return result, True
                
                elif choice == "4":
                    print("⚠️  跳过此组")
                    return pd.DataFrame(), True  # 返回空数据框
                
                else:
                    print("❌ 请输入 1-4 之间的数字")
                    
            except KeyboardInterrupt:
                print("\n⚠️  用户中断，保留第一条记录")
                return group_df.head(1), True
    
    def _manual_select_student_name(self, group_df: pd.DataFrame, unique_names: dict, student_name_field: str) -> pd.DataFrame:
        """手动选择要保留的学生姓名，并处理其他冲突字段"""
        print(f"\n📝 请选择要保留的姓名:")
        name_list = list(unique_names.keys())
        for i, name in enumerate(name_list, 1):
            records_count = len(unique_names[name])
            print(f"  {i}. {name} ({records_count} 条记录)")
        
        while True:
            try:
                choice = input(f"请选择姓名编号 (1-{len(name_list)}): ").strip()
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(name_list):
                    selected_name = name_list[choice_idx]
                    selected_records = unique_names[selected_name]
                    
                    print(f"✅ 已选择姓名: {selected_name}")
                    
                    # 检查是否还有其他冲突字段需要处理
                    conflict_info = self._get_remaining_conflicts(group_df, selected_records, student_name_field)
                    
                    if conflict_info:
                        print(f"\n⚠️  发现其他冲突字段，需要进一步处理:")
                        for field, values in conflict_info.items():
                            field_icon = self._get_field_icon(field)
                            print(f"  {field_icon} {field}: {len(values)} 个不同值")
                            for i, value in enumerate(sorted(values), 1):
                                print(f"    {i}. {value}")
                        
                        # 询问用户是否要处理其他冲突字段
                        print(f"\n🤔 是否要处理其他冲突字段？")
                        print(f"  1. 是，手动选择每个字段的值")
                        print(f"  2. 否，使用第一条记录的值")
                        
                        conflict_choice = input("请选择 (1-2，默认2): ").strip()
                        if conflict_choice == "1":
                            # 手动处理其他冲突字段
                            result_record = self._manual_resolve_remaining_conflicts(selected_records[0], conflict_info)
                            return pd.DataFrame([result_record])
                        else:
                            # 使用第一条记录
                            print("✅ 使用第一条记录的值")
                            return pd.DataFrame([selected_records[0]])
                    else:
                        # 没有其他冲突字段，直接返回第一条匹配的记录
                        return pd.DataFrame([selected_records[0]])
                else:
                    print("❌ 编号超出范围，请重新选择")
            except ValueError:
                print("❌ 请输入有效的数字")
    
    def _get_remaining_conflicts(self, group_df: pd.DataFrame, selected_records: list, student_name_field: str) -> dict:
        """获取除了姓名字段之外的其他冲突字段"""
        conflict_info = {}
        exclude_fields = set(['数据来源文件', '数据来源路径', student_name_field])
        
        # 检查整个 group_df 中是否还有其他冲突字段
        for field in group_df.columns:
            if field in exclude_fields:
                continue
            
            # 检查该字段在整个组中是否有冲突
            unique_values = set()
            for _, record in group_df.iterrows():
                value = record[field]
                # 修改：包含空值，因为空值也是一种有效的值，需要用户选择
                if pd.notna(value):
                    normalized_value = str(value).strip() if str(value).strip() else "<空值>"
                    unique_values.add(normalized_value)
                else:
                    unique_values.add("<空值>")
            
            if len(unique_values) > 1:
                conflict_info[field] = unique_values
        
        return conflict_info
    
    def _manual_resolve_remaining_conflicts(self, base_record: pd.Series, conflict_info: dict) -> pd.Series:
        """手动解决剩余冲突字段"""
        result_record = base_record.copy()
        
        print(f"\n🔧 开始处理其他冲突字段...")
        print(f"📄 基础记录: {dict(result_record)}")
        
        for field, values in conflict_info.items():
            print(f"\n📝 请选择字段 '{field}' 的值:")
            print(f"🔍 当前值: {result_record[field]}")
            print(f"📋 可选值:")
            
            # 将 set 转换为 list 以便索引访问
            values_list = list(values)
            
            for i, value in enumerate(values_list, 1):
                if pd.isna(value):
                    print(f"  {i}. <空值>")
                else:
                    print(f"  {i}. {value}")
            
            while True:
                try:
                    choice = input(f"请选择 (1-{len(values_list)}): ").strip()
                    choice_idx = int(choice) - 1
                    if 0 <= choice_idx < len(values_list):
                        selected_value = values_list[choice_idx]
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
        
        print(f"\n✅ 所有冲突字段处理完成！")
        print(f"📄 最终记录: {dict(result_record)}")
        return result_record
    
    def _create_records_by_name(self, group_df: pd.DataFrame, unique_names: dict, student_name_field: str) -> pd.DataFrame:
        """为每个不同姓名创建单独记录"""
        result_records = []
        
        print(f"\n📝 为每个不同姓名创建记录:")
        for i, (name, records) in enumerate(unique_names.items(), 1):
            # 使用该姓名的第一条记录
            record = records[0]
            result_records.append(record)
            print(f"  {i}. 创建记录: 姓名={name}")
        
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

    def _is_money_field(self, field_name: str) -> bool:
        """判断字段是否为金钱字段"""
        field_lower = field_name.lower()
        money_keywords = ['金额', '价格', 'price', 'amount', '费用', '成本', 'money', 'money', '元', '￥', '$', '¥']
        return any(keyword in field_lower for keyword in money_keywords)
    
    def _is_money_value_equal(self, val1, val2) -> bool:
        """
        比较两个金钱值是否相等
        
        Args:
            val1: 第一个值
            val2: 第二个值
            
        Returns:
            bool: 如果金钱值相等返回True，否则返回False
        """
        # 如果两个值都是NaN，认为相等
        if pd.isna(val1) and pd.isna(val2):
            return True
        
        # 如果一个是NaN另一个不是，认为不相等
        if pd.isna(val1) or pd.isna(val2):
            return False
        
        try:
            # 尝试转换为数值进行比较
            num1 = float(str(val1).replace(',', '').replace('￥', '').replace('$', '').replace('¥', '').replace('元', ''))
            num2 = float(str(val2).replace(',', '').replace('￥', '').replace('$', '').replace('¥', '').replace('元', ''))
            
            # 使用小的容差值比较浮点数
            return abs(num1 - num2) < 0.01
        except (ValueError, TypeError):
            # 如果无法转换为数值，回退到字符串比较
            return str(val1).strip() == str(val2).strip()
    
    def _get_field_icon(self, field_name: str) -> str:
        """根据字段名称智能选择图标"""
        field_lower = field_name.lower()
        
        # 姓名相关字段
        if any(keyword in field_lower for keyword in ['姓名', '名字', 'name', '姓', '名']):
            return "👤"
        # 名称/标题相关字段
        elif any(keyword in field_lower for keyword in ['名称', '标题', 'title', '名称']):
            return "🏷️"
        # 地址相关字段
        elif any(keyword in field_name.lower() for keyword in ['地址', '住址', 'address', '位置']):
            return "📍"
        # 电话相关字段
        elif any(keyword in field_name.lower() for keyword in ['电话', '手机', 'phone', 'tel', '号码']):
            return "📞"
        # 邮箱相关字段
        elif any(keyword in field_name.lower() for keyword in ['邮箱', '邮件', 'email', '信箱']):
            return "📧"
        # 日期时间相关字段
        elif any(keyword in field_name.lower() for keyword in ['日期', '时间', 'date', 'time', '年', '月', '日']):
            return "📅"
        # 数量金额相关字段
        elif any(keyword in field_name.lower() for keyword in ['数量', '金额', '价格', 'price', 'amount', '费用', '成本']):
            return "💰"
        # 默认图标
        else:
            return "🔍"

    def _format_display_value(self, value) -> str:
        """
        格式化显示值，处理数值类型的显示格式
        
        Args:
            value: 要格式化的值
            
        Returns:
            格式化后的字符串
        """
        if pd.isna(value):
            return "<空值>"
        
        # 如果是浮点数且小数部分为0，显示为整数
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        
        # 其他情况直接转换为字符串
        return str(value)

    def _has_field_conflicts(self, group_df: pd.DataFrame) -> bool:
        """
        检查重复组是否有字段冲突（不是所有记录都完全相同）
        
        Args:
            group_df: 重复组的数据框
            
        Returns:
            bool: 如果有冲突返回True，如果所有记录完全相同返回False
        """
        if len(group_df) <= 1:
            return False
        
        # 检查是否所有记录完全相同（排除数据来源文件字段）
        all_fields = [field for field in group_df.columns if field != '数据来源文件']
        first_record = group_df.iloc[0]
        
        # 检查是否所有记录都与第一条记录完全相同
        for _, row in group_df.iterrows():
            for field in all_fields:
                # 处理NaN值的比较
                first_val = first_record[field]
                current_val = row[field]
                
                # 如果两个值都是NaN，认为相同
                if pd.isna(first_val) and pd.isna(current_val):
                    continue
                # 如果一个是NaN另一个不是，认为不同
                elif pd.isna(first_val) or pd.isna(current_val):
                    return True  # 有冲突
                
                # 特殊处理金钱字段
                if self._is_money_field(field):
                    if not self._is_money_value_equal(first_val, current_val):
                        return True  # 金钱值不同，有冲突
                else:
                    # 非金钱字段，比较字符串形式
                    if str(first_val).strip() != str(current_val).strip():
                        return True  # 有冲突
        
        return False  # 所有记录完全相同，无冲突

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

    def _normalize_for_compare(self, value) -> str:
        """
        归一化比较值：
        - NaN -> ""
        - 浮点整数 -> 去掉 .0
        - 其他 -> 去首尾空格的字符串
        """
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _find_actual_field_name_silent(self, df: pd.DataFrame, target_field: str) -> str:
        """
        静默列名匹配：不打印、不交互。
        匹配顺序：精确 -> 不区分大小写 -> 清洗后的列名匹配 -> 常见变体 -> 相似度最高（>=阈值）
        """
        available = list(df.columns)
        # 1) 精确
        if target_field in available:
            return target_field

        # 2) 不区分大小写
        lower_map = {c.lower(): c for c in available}
        if target_field.lower() in lower_map:
            return lower_map[target_field.lower()]

        # 3) 清洗后的列名
        cleaned_target = self.clean_column_name(target_field)
        cleaned_map = {self.clean_column_name(c): c for c in available}
        if cleaned_target in cleaned_map:
            return cleaned_map[cleaned_target]

        # 4) 常见变体
        if hasattr(self, 'common_column_variants') and target_field in self.common_column_variants:
            for variant in self.common_column_variants[target_field]:
                # 先精确
                if variant in available:
                    return variant
                # 大小写
                if variant.lower() in lower_map:
                    return lower_map[variant.lower()]
                # 清洗后
                cv = self.clean_column_name(variant)
                if cv in cleaned_map:
                    return cleaned_map[cv]

        # 5) 相似度
        best_col = None
        best_sim = 0.0
        for col in available:
            sim = SequenceMatcher(None, cleaned_target.lower(), self.clean_column_name(col).lower()).ratio()
            if sim > best_sim:
                best_sim = sim
                best_col = col
        if best_col and best_sim >= getattr(self, 'similarity_threshold', 0.8):
            return best_col

        return None

    def _verify_group_key_in_file(self, file_path: str, dedup_fields: List[str], group_key) -> bool:
        """
        校验：在指定的源文件中，是否存在与当前重复组键一致的记录。
        静默匹配列名，避免打印和交互，且进行值归一化比较。
        """
        try:
            df_src = pd.read_excel(file_path)
        except Exception:
            return False

        # 定位实际列名（静默）
        actual_cols = []
        for field in dedup_fields:
            actual = self._find_actual_field_name_silent(df_src, field)
            if not actual:
                return False
            actual_cols.append(actual)

        # 组装组键值
        if isinstance(group_key, tuple):
            key_values = list(group_key)
        else:
            key_values = [group_key]
        if len(key_values) != len(actual_cols):
            return False

        # 构建掩码进行比较（统一归一化）
        mask = pd.Series([True] * len(df_src))
        for actual_col, key_value in zip(actual_cols, key_values):
            series_obj = df_src[actual_col]
            # 归一化列
            series_norm = series_obj.apply(self._normalize_for_compare)
            cmp_val = self._normalize_for_compare(key_value)
            mask = mask & (series_norm == cmp_val)

        return bool(mask.any())

    def _group_has_student_name_conflict(self, group_df: pd.DataFrame, dedup_fields: List[str], student_name_field: str) -> bool:
        """
        检查重复组是否存在冲突（学号相同但姓名不同，或其他字段不同）
        
        Args:
            group_df: 重复组的数据框
            dedup_fields: 去重字段列表
            student_name_field: 学生姓名字段名
            
        Returns:
            bool: 如果存在冲突返回True，否则返回False
        """
        if len(group_df) <= 1:
            return False
        
        # 如果有姓名字段，检查姓名冲突
        if student_name_field and student_name_field in group_df.columns:
            unique_names = set()
            for name in group_df[student_name_field]:
                if pd.notna(name) and str(name).strip():
                    normalized_name = str(name).strip()
                    unique_names.add(normalized_name)
            
            # 如果有超过1个不同的姓名，则认为有冲突
            if len(unique_names) > 1:
                return True
        
        # 检查其他非去重字段是否存在冲突
        exclude_fields = set(['数据来源文件', '数据来源路径'] + dedup_fields)
        for field in group_df.columns:
            if field in exclude_fields:
                continue
            
            # 检查该字段是否有不同的值
            unique_values = set()
            for value in group_df[field]:
                if pd.notna(value) and str(value).strip():
                    normalized_value = str(value).strip()
                    unique_values.add(normalized_value)
            
            # 如果有超过1个不同的值，则认为有冲突
            if len(unique_values) > 1:
                return True
        
        return False
    
    def _group_has_conflict_normalized(self, group_df: pd.DataFrame, dedup_fields: List[str]) -> bool:
        """
        使用归一化后的取值来判断是否存在真实冲突：
        - 仅检查非去重字段，且排除来源字段
        - 忽略空值
        - 同值不同类型（如 2020062959.0 与 '2020062959'）视为相同
        """
        exclude_fields = set(['数据来源文件', '数据来源路径'])
        for field in group_df.columns:
            if field in dedup_fields or field in exclude_fields:
                continue
            normalized_values = set()
            for v in group_df[field].tolist():
                nv = self._normalize_for_compare(v)
                if nv != "":
                    normalized_values.add(nv)
            if len(normalized_values) > 1:
                return True
        return False
    
    def _manual_select_record(self, group_df: pd.DataFrame, conflict_info: Dict[str, set]) -> pd.DataFrame:
        """
        手动选择要保留的记录（适用于非姓名字段冲突）
        
        Args:
            group_df: 重复组的数据框
            conflict_info: 冲突字段信息字典
            
        Returns:
            选择保留的记录
        """
        print(f"\n📋 请选择要保留的记录:")
        
        # 显示每条记录的详细信息
        for i, (_, record) in enumerate(group_df.iterrows(), 1):
            print(f"\n  📝 记录 {i}:")
            for field, value in record.items():
                if field in ['数据来源文件', '数据来源路径']:
                    continue
                display_value = self._format_display_value(value)
                if field in conflict_info:
                    print(f"    🔍 {field}: {display_value} (冲突字段)")
                else:
                    print(f"    📊 {field}: {display_value}")
        
        while True:
            try:
                choice = input(f"\n请选择要保留的记录 (1-{len(group_df)}): ").strip()
                if not choice:
                    choice = "1"
                
                choice_num = int(choice)
                if 1 <= choice_num <= len(group_df):
                    selected_record = group_df.iloc[choice_num - 1:choice_num]
                    print(f"✅ 已选择记录 {choice_num}")
                    return selected_record
                else:
                    print(f"❌ 请输入 1-{len(group_df)} 之间的数字")
                    
            except ValueError:
                print("❌ 请输入有效的数字")
            except KeyboardInterrupt:
                print("\n⚠️  用户中断，保留第一条记录")
                return group_df.head(1)
    
    def _create_records_by_conflict_fields(self, group_df: pd.DataFrame, conflict_info: Dict[str, set]) -> pd.DataFrame:
        """
        为每个不同值创建单独记录（适用于非姓名字段冲突）
        
        Args:
            group_df: 重复组的数据框
            conflict_info: 冲突字段信息字典
            
        Returns:
            处理后的记录
        """
        result_records = []
        
        # 按冲突字段分组
        for field, unique_values in conflict_info.items():
            for value in unique_values:
                # 找到该值的所有记录
                field_records = group_df[group_df[field] == value]
                if not field_records.empty:
                    # 保留第一条记录
                    result_records.append(field_records.head(1))
        
        if result_records:
            return pd.concat(result_records, ignore_index=True)
        else:
            return group_df.head(1)
    
    def _identify_student_id_field(self, dedup_fields: List[str], all_columns: List[str]) -> str:
        """
        智能识别学号字段
        
        Args:
            dedup_fields: 去重字段列表
            all_columns: 所有可用字段列表
            
        Returns:
            识别出的学号字段名，如果没有找到返回None
        """
        # 学号字段的常见关键词和模式
        id_keywords = [
            '学号', '学生号', '学籍号', '编号', 'ID', 'id', 'Id',
            '工号', '员工号', '职工号', '编号', '号码',
            '单位号', '部门号', '机构号', '组织号',
            '账号', '用户号', '会员号', '客户号',
            '订单号', '流水号', '序列号', '编码'
        ]
        
        # 优先在去重字段中查找
        for field in dedup_fields:
            field_lower = field.lower()
            for keyword in id_keywords:
                if keyword in field_lower:
                    return field
        
        # 在去重字段中查找包含数字的字段
        for field in dedup_fields:
            if any(char.isdigit() for char in field):
                return field
        
        # 在所有字段中查找学号相关字段
        for field in all_columns:
            field_lower = field.lower()
            for keyword in id_keywords:
                if keyword in field_lower:
                    return field
        
        return None
    
    def _identify_name_field(self, all_columns: List[str]) -> str:
        """
        智能识别姓名字段
        
        Args:
            all_columns: 所有可用字段列表
            
        Returns:
            识别出的姓名字段名，如果没有找到返回None
        """
        # 姓名字段的常见关键词和模式
        name_keywords = [
            '姓名', '名字', '名称', '全名', '中文名', '英文名',
            '姓', '名', '名字', '称谓',
            '单位名称', '部门名称', '机构名称', '组织名称',
            '产品名称', '商品名称', '项目名称', '标题',
            '名称', '名字', '标题', '描述'
        ]
        
        # 在所有字段中查找姓名相关字段
        for field in all_columns:
            field_lower = field.lower()
            for keyword in name_keywords:
                if keyword in field_lower:
                    return field
        
        return None

def main():
    """主函数"""
    processor = ExcelProcessor()
    processor.run()

if __name__ == "__main__":
    main()